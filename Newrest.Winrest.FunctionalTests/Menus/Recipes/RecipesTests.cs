using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Menus;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Recipes;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Production;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Item;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using static Newrest.Winrest.FunctionalTests.PageObjects.Menus.Recipes.RecipeMassiveDeletePopup;

namespace Newrest.Winrest.FunctionalTests.Menus
{

    [TestClass]
    public class RecipesTests : TestBase
    {

        private const int _timeout = 600000;
        private const string RECIPES_EXCEL_SHEET_NAME = "Sheet 1";
        private readonly string recipeNameToday = "Rec-" + DateUtils.Now.ToString("dd/MM/yyyy");

        /// <summary>
        /// 
        /// Mise en place du paramétrage pour la configuration Winrest 4.0 
        /// 
        /// </summary>
        [TestMethod]
        [Priority(0)]
        [Timeout(_timeout)]
        public void RE_RECI_SetConfigWinrest4_0()
        {
            string itemName = "testRecipeAllergen";
            var site = TestContext.Properties["Site"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            ClearCache();
            homePage.CloseNewsPopup();

            // New version search
            homePage.SetNewVersionSearchValue(true);

            // New VersionItemDetails
            //homePage.SetVersionItemDetailValue(true);

            // Vérifier que c'est activé
            var itemPage = homePage.GoToPurchasing_ItemPage();

            try
            {
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
            }
            catch
            {
                throw new Exception("La recherche n'a pas pu être effectuée, le NewSearchMode est inactif.");
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
        }

        [Priority(1)]
        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_CreateVariant()
        {
            // Prepare
            string variant = TestContext.Properties["RecipeVariant"].ToString();
            string meal = TestContext.Properties["RecipeMeal"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string siteBis = TestContext.Properties["SiteBis"].ToString();

            string guestName = variant.Substring(0, variant.IndexOf("-")).Trim();

            //Arrange
            var homePage = LogInAsAdmin();

            var productionPage = homePage.GoToParameters_ProductionPage();

            productionPage.GoToTab_Guest();

            if (!productionPage.IsGuestPresent(guestName))
            {
                productionPage.AddNewGuest(guestName, false, "1000");

            }
            WebDriver.Navigate().Refresh();
            productionPage.GoToTab_Guest();
            Assert.IsTrue(productionPage.IsGuestPresent(guestName), $"Le guest {guestName} n'est pas présent.");

            // Variant 1 pour MAD
            productionPage.GoToTab_Variant();
            productionPage.FilterSite(site);
            productionPage.FilterGuestType(guestName);
            productionPage.FilterMealType(meal);

            if (!productionPage.IsVariantPresent(guestName, meal))
            {
                productionPage.AddNewVariant(guestName, meal, site);
            }

            // Variant 2 pour IBZ
            WebDriver.Navigate().Refresh();
            productionPage.GoToTab_Variant();
            productionPage.GoToTab_Variant();
            productionPage.FilterSite(siteBis);
            productionPage.FilterGuestType(guestName);
            productionPage.FilterMealType(meal);

            if (!productionPage.IsVariantPresent(guestName, meal))
            {
                productionPage.AddNewVariant(guestName, meal, siteBis);
            }

            // Vérification
            WebDriver.Navigate().Refresh();
            productionPage.GoToTab_Variant();
            productionPage.FilterSite(site);
            productionPage.FilterGuestType(guestName);
            productionPage.FilterMealType(meal);
            Assert.IsTrue(productionPage.IsVariantPresent(guestName, meal), $"La variante {variant} n'est pas présente pour le site {site}.");

            WebDriver.Navigate().Refresh();
            productionPage.GoToTab_Variant();
            productionPage.FilterSite(siteBis);
            productionPage.FilterGuestType(guestName);
            productionPage.FilterMealType(meal);
            Assert.IsTrue(productionPage.IsVariantPresent(guestName, meal), $"La variante {variant} n'est pas présente pour le site {siteBis}.");
        }

        [Priority(2)]
        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_VerifyIngredientPackaging()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string siteBis = TestContext.Properties["SiteBis"].ToString();
            string siteACE = TestContext.Properties["SiteLP"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, ingredient);

            var itemGeneralInformationPage = new ItemGeneralInformationPage(WebDriver, TestContext);

            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(ingredient, group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, "10", "KG", "5", supplier);

                itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(siteBis, packagingName, "10", "KG", "5", supplier);

                itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(siteACE, packagingName, "10", "KG", "5", supplier);
            }
            else
            {
                itemGeneralInformationPage = itemPage.ClickOnFirstItem();

                if (!itemGeneralInformationPage.VerifyPackagingSite(site))
                {
                    var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                    itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, "10", "KG", "5", supplier);
                }

                if (!itemGeneralInformationPage.VerifyPackagingSite(siteBis))
                {
                    var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                    itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(siteBis, packagingName, "10", "KG", "5", supplier);
                }

                if (!itemGeneralInformationPage.VerifyPackagingSite(siteACE))
                {
                    var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                    itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(siteACE, packagingName, "10", "KG", "5", supplier);
                }
            }

            Assert.IsTrue(itemGeneralInformationPage.VerifyPackagingSite(site), $"L'item ne possède pas de packaging pour le site {site}.");
            Assert.IsTrue(itemGeneralInformationPage.VerifyPackagingSite(siteBis), $"L'item ne possède pas de packaging pour le site {siteBis}.");
            Assert.IsTrue(itemGeneralInformationPage.VerifyPackagingSite(siteBis), $"L'item ne possède pas de packaging pour le site {siteACE}.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Filter_Search_Recipe()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());
            recipesCreateModalPage.WaitPageLoading();
            recipeGeneralInfosPage.AddVariantWithSite(site, variant);
            var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

            recipeVariantPage.AddIngredient(ingredient);
            recipeVariantPage.WaitPageLoading();
            recipesPage = recipeVariantPage.BackToList();

            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);
            recipesPage.WaitPageLoading();
            //Assert
            var firstRecipeName = recipesPage.GetFirstRecipeName();
            Assert.AreEqual(recipeName, firstRecipeName, string.Format(MessageErreur.FILTRE_ERRONE, "Search"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Filter_SortByName()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariant"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            int totalNumber = recipesPage.CheckTotalNumber();
            try
            {
                if (totalNumber < 20)
                {
                    var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                    var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

                    recipeGeneralInfosPage.AddVariantWithSite(site, variant);
                    var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                    recipeVariantPage.AddIngredient(ingredient);
                    recipesPage = recipeVariantPage.BackToList();
                    recipesPage.ResetFilter();
                }

                if (!recipesPage.isPageSizeEqualsTo100())
                {
                    recipesPage.PageSize("8");
                    recipesPage.PageSize("100");
                }

                recipesPage.Filter(RecipesPage.FilterType.SortBy, "Name");
                var isSortedByName = recipesPage.IsSortedByName();

                //Assert
                Assert.IsTrue(isSortedByName, String.Format(MessageErreur.FILTRE_ERRONE, "'Sort by name'"));
            }
            finally
            {
                if (totalNumber < 20)
                {
                    recipesPage = homePage.GoToMenus_Recipes();
                    recipesPage.MassiveDeleteRecipes(recipeName, site, recipeType);
                }
            }
            
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Filter_SortByPortion()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariant"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            int totalNumber = recipesPage.CheckTotalNumber();
            try
            {
                if (totalNumber < 20)
                {
                    var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                    var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

                    recipeGeneralInfosPage.AddVariantWithSite(site, variant);
                    var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                    recipeVariantPage.AddIngredient(ingredient);
                    recipesPage = recipeVariantPage.BackToList();
                    recipesPage.ResetFilter();
                }

                recipesPage.Filter(RecipesPage.FilterType.SortBy, "Portions");

                if (!recipesPage.isPageSizeEqualsTo100())
                {
                    recipesPage.PageSize("8");
                    recipesPage.PageSize("100");
                }

                var isSortedByPortions = recipesPage.IsSortedByPortions();
                //Assert
                Assert.IsTrue(isSortedByPortions, String.Format(MessageErreur.FILTRE_ERRONE, "'Sort by portions'"));
            }
            finally
            {
                if (totalNumber < 20)
                {
                    recipesPage = homePage.GoToMenus_Recipes();
                    recipesPage.MassiveDeleteRecipes(recipeName, site, recipeType);
                }
            }

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Filter_ShowRecipesWithoutVariants()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();

            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.ShowWithoutVariants, true);

            if (recipesPage.CheckTotalNumber() < 20)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());
                recipesPage = recipeGeneralInfosPage.BackToList();
                recipesPage.ResetFilter();
                recipesPage.Filter(RecipesPage.FilterType.ShowWithoutVariants, true);
            }

            if (!recipesPage.isPageSizeEqualsTo100())
            {
                recipesPage.PageSize("8");
                recipesPage.PageSize("100");
            }

            //Assert
            Assert.IsTrue(recipesPage.IsWithoutVariant(), String.Format(MessageErreur.FILTRE_ERRONE, "'Show recipes without variants'"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Filter_ShowRecipesWithVariants()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();

            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);

            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.ShowWithoutVariants, false);

            if (recipesPage.CheckTotalNumber() < 20)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());
                recipesPage = recipeGeneralInfosPage.BackToList();
                recipesPage.ResetFilter();
                recipesPage.Filter(RecipesPage.FilterType.ShowWithoutVariants, false);
            }

            if (!recipesPage.isPageSizeEqualsTo100())
            {
                recipesPage.PageSize("8");
                recipesPage.PageSize("100");
            }

            //Assert
            Assert.IsTrue(recipesPage.IsWithVariant(), String.Format(MessageErreur.FILTRE_ERRONE, "'Show recipes without variants'"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Filter_ShowActive()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariant"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.ShowActive, true);

            if (recipesPage.CheckTotalNumber() < 20)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

                recipeGeneralInfosPage.AddVariantWithSite(site, variant);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(ingredient);
                recipesPage = recipeVariantPage.BackToList();
                recipesPage.ResetFilter();
                recipesPage.Filter(RecipesPage.FilterType.ShowActive, true);
            }

            if (!recipesPage.isPageSizeEqualsTo100())
            {
                recipesPage.PageSize("8");
                recipesPage.PageSize("100");
            }
            recipesPage.Filter(RecipesPage.FilterType.ShowAll, true);
            var showallNumber = recipesPage.CheckTotalNumber();
            recipesPage.Filter(RecipesPage.FilterType.ShowInactive, true);
            var showInactiveNumber = recipesPage.CheckTotalNumber();
            recipesPage.Filter(RecipesPage.FilterType.ShowActive, true);
            var showActiveNumber = recipesPage.CheckTotalNumber();

            //Assert
            var statusChecked = recipesPage.CheckStatus(true);
            Assert.IsTrue(recipesPage.CheckStatus(true), String.Format(MessageErreur.FILTRE_ERRONE, "'Show only active'"));
            Assert.AreEqual(showActiveNumber, showallNumber - showInactiveNumber, "le filtre ne fonctionne pas ");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Filter_ShowInactive()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariant"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.ShowInactive, true);

            if (recipesPage.CheckTotalNumber() < 20)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString(), false);

                recipeGeneralInfosPage.AddVariantWithSite(site, variant);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(ingredient);
                recipesPage = recipeVariantPage.BackToList();
                recipesPage.ResetFilter();
                recipesPage.Filter(RecipesPage.FilterType.ShowInactive, true);
            }

            if (!recipesPage.isPageSizeEqualsTo100())
            {
                recipesPage.PageSize("8");
                recipesPage.PageSize("100");
            }
            recipesPage.Filter(RecipesPage.FilterType.ShowAll, true);
            var showallNumber = recipesPage.CheckTotalNumber();
            recipesPage.Filter(RecipesPage.FilterType.ShowActive, true);
            var showActiveNumber = recipesPage.CheckTotalNumber();
            recipesPage.Filter(RecipesPage.FilterType.ShowInactive, true);
            var showInactiveNumber = recipesPage.CheckTotalNumber();



            var statusChecked = recipesPage.CheckStatus(false);
            //Assert
            Assert.IsFalse(statusChecked, String.Format(MessageErreur.FILTRE_ERRONE, "'Show only inactive'"));
            Assert.AreEqual(showInactiveNumber, showallNumber - showActiveNumber, "le filtre ne marche pas correctement");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Filter_ShowAll()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariant"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.ShowAll, true);

            if (recipesPage.CheckTotalNumber() < 20)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

                recipeGeneralInfosPage.AddVariantWithSite(site, variant);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(ingredient);
                recipesPage = recipeVariantPage.BackToList();
                recipesPage.ResetFilter();
            }

            recipesPage.Filter(RecipesPage.FilterType.ShowActive, true);
            int nbActifs = recipesPage.CheckTotalNumber();

            recipesPage.Filter(RecipesPage.FilterType.ShowInactive, true);
            int nbInactifs = recipesPage.CheckTotalNumber();

            recipesPage.Filter(RecipesPage.FilterType.ShowAll, true);
            int nbTotal = recipesPage.CheckTotalNumber();

            //Assert
            Assert.AreEqual(nbTotal, nbActifs + nbInactifs, String.Format(MessageErreur.FILTRE_ERRONE, "'Show all'"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Filter_TypeOfGuest()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariant"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();

            var homePage = LogInAsAdmin();
            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            try
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());
                recipeGeneralInfosPage.AddVariantWithSite(site, variant);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
                recipeVariantPage.AddIngredient(ingredient);
                recipesPage = recipeVariantPage.BackToList();
                recipesPage.ResetFilter();
                recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);
                recipesPage.Filter(RecipesPage.FilterType.TypeOfGuest, guestType);
                recipesPage.UnfoldAll();
                var Firstvariant = recipesPage.GetFirstVariant();
                var FirstGuestType = Firstvariant.Substring(0, variant.IndexOf("-")).Trim();
                var nbrRecette = recipesPage.CheckTotalNumber();
                Assert.AreEqual(nbrRecette,1, "La recette n'a pas été créée");
                Assert.AreEqual(FirstGuestType, guestType, "les résultats ne s'accordent pas bien au filtre appliqué");
            }
            finally
            {
                recipesPage = homePage.GoToMenus_Recipes();
                recipesPage.MassiveDeleteRecipes(recipeName, site, recipeType);
            }


        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Filter_TypeOfMeal()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariant"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);
            string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();
            string typeOfmeal = variant.Substring(variant.IndexOf("-") + 1).Trim();


            var homePage = LogInAsAdmin();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            try
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

                recipeGeneralInfosPage.AddVariantWithSite(site, variant);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(ingredient);
                recipesPage = recipeVariantPage.BackToList();
                recipesPage.ResetFilter();
                recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);
                recipesPage.Filter(RecipesPage.FilterType.TypeOfMeal, typeOfmeal);
                recipesPage.Filter(RecipesPage.FilterType.TypeOfGuest, guestType);
                var numberTotal = recipesPage.CheckTotalNumber();

                Assert.AreEqual(numberTotal, 1, " La recette n'a pas été créée ou meal Type ne correspond pas à cette recette.");
            }

            finally
            {
                recipesPage = homePage.GoToMenus_Recipes();
                recipesPage.MassiveDeleteRecipes(recipeName, site, recipeType);
            }

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Filter_Intolerance()
        {
            string recipeType = TestContext.Properties["RecipeTypeBis"].ToString();
            string variant = TestContext.Properties["RecipeVariant"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, ingredient);
            ItemGeneralInformationPage itemGeneralInformationPage = itemPage.ClickFirstItem();
            ItemIntolerancePage itemIntolerancePage = itemGeneralInformationPage.ClickOnIntolerancePage();
            var listAllergenItem = itemIntolerancePage.GetListAllergensSelected();
            string FirstItemAllergen = listAllergenItem[0];

            var recipesPage = homePage.GoToMenus_Recipes();
            try
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

                recipeGeneralInfosPage.AddVariantWithSite(site, variant);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
                recipeVariantPage.AddIngredient(ingredient);
                recipesPage = recipeVariantPage.BackToList();
                recipesPage.ResetFilter();
                recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);
                recipesPage.Filter(RecipesPage.FilterType.Intolerance, FirstItemAllergen);
                var numberTotal = recipesPage.CheckTotalNumber();

                Assert.AreEqual(numberTotal, 1, " La recette n'a pas été créée ou l'allergène ne correspond pas à cette recette.");

                RecipeGeneralInformationPage recipeGeneralInformationPage = recipesPage.SelectFirstRecipe();
                recipeVariantPage = recipeGeneralInformationPage.ClickOnFirstVariant();
                RecipeIntolerancePage recipeIntolerancePage = recipeVariantPage.ClickOnIntoleranceTab();
                var listAllergenRecipe = recipeIntolerancePage.GetListAllergensSelected();
                string FirstRecipeAllergen = listAllergenRecipe[0];

                Assert.AreEqual(FirstRecipeAllergen, FirstItemAllergen, " les résultats ne s'accordent pas bien au filtre appliqué");
            }
            finally
            {
                recipesPage = homePage.GoToMenus_Recipes();
                recipesPage.MassiveDeleteRecipes(recipeName, site, recipeType);
            }



        }



        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Filter_RecipeTypes()
        {
            //Prepare
            string recipeType = "ENSALADA";// TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariant"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);

            bool newVersionPrint = true;

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();

            recipesPage.ClearDownloads();

            recipesPage.Filter(RecipesPage.FilterType.RecipeType, recipeType);

            if (recipesPage.CheckTotalNumber() < 20)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

                recipeGeneralInfosPage.AddVariantWithSite(site, variant);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(ingredient);
                recipesPage = recipeVariantPage.BackToList();
                recipesPage.ResetFilter();
                recipesPage.Filter(RecipesPage.FilterType.RecipeType, recipeType);
            }
            recipesPage.WaitPageLoading();
            recipesPage.Export(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = recipesPage.GetExportExcelFile(taskFiles, false, false);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            int resultNumber = OpenXmlExcel.GetExportResultNumber(RECIPES_EXCEL_SHEET_NAME, filePath);
            bool result = OpenXmlExcel.ReadAllDataInColumn("Recipe type", RECIPES_EXCEL_SHEET_NAME, filePath, recipeType);

            //Assert
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.IsTrue(result, String.Format(MessageErreur.FILTRE_ERRONE, "'Recipe Types'"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Filter_DietaryTypes()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariant"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string dietaryType = TestContext.Properties["RecipeDietaryType"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);

            bool newVersionPrint = true;

            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();

            recipesPage.ClearDownloads();

            recipesPage.Filter(RecipesPage.FilterType.Dietarytype, dietaryType);

            if (recipesPage.CheckTotalNumber() < 20)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString(), true, dietaryType);

                recipeGeneralInfosPage.AddVariantWithSite(site, variant);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(ingredient);
                recipesPage = recipeVariantPage.BackToList();
                recipesPage.ResetFilter();
                recipesPage.Filter(RecipesPage.FilterType.Dietarytype, dietaryType);
            }

            recipesPage.Export(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = recipesPage.GetExportExcelFile(taskFiles, false, false);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            int resultNumber = OpenXmlExcel.GetExportResultNumber(RECIPES_EXCEL_SHEET_NAME, filePath);
            bool result = OpenXmlExcel.ReadAllDataInColumn("Dietary type", RECIPES_EXCEL_SHEET_NAME, filePath, dietaryType);

            //Assert
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.IsTrue(result, String.Format(MessageErreur.FILTRE_ERRONE, "'Dietary Types'"));
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Create_New_Recipe()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["MenuVariantACE3"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();

            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            try
            {
                var recipesPage = homePage.GoToMenus_Recipes();
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

                recipeGeneralInfosPage.AddVariantWithSite(site, variant);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(ingredient);
                recipesPage = recipeVariantPage.BackToList();

                recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);

                //Assert
                Assert.AreEqual(recipeName, recipesPage.GetFirstRecipeName(), "La recette n'a pas été ajoutée à la liste.");
            }
            finally
            {
                var recipesPage = homePage.GoToMenus_Recipes();
                recipesPage.MassiveDeleteRecipes(recipeName, site, recipeType);

            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Create_New_Recipe_Already_Existing()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

            recipesPage = recipeGeneralInfosPage.BackToList();

            // Creation of another recipe with the same name
            recipesCreateModalPage = recipesPage.CreateNewRecipe();
            recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

            //Assert
            var errorMessageName = recipesCreateModalPage.ErrorMessageNameAlreadyExists();
            Assert.IsTrue(errorMessageName, "La recette n'est pas dupliquée.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Duplicate_Recipe_Variant()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["SiteBis"].ToString();
            string siteDestination = TestContext.Properties["SiteLP"].ToString();

            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);


            //Arrange
            var homePage = LogInAsAdmin();
            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();
            ParametersProduction parametersProduction = homePage.GoToParameters_ProductionPage();
            parametersProduction.GoToTab_Variant();
            parametersProduction.FilterSite(site);
            var variant = parametersProduction.GetFirstVariant();
            parametersProduction.FilterSite(siteDestination);
            var variantBis = parametersProduction.GetFirstVariant();
            var recipesPage = homePage.GoToMenus_Recipes();

            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

            recipeGeneralInfosPage.AddVariantWithSite(site, variant);
            var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

            recipeVariantPage.AddIngredient(ingredient);
            recipesPage = recipeVariantPage.BackToList();

            try
            {
                var duplicateRecipesVariantModalPage = recipesPage.DuplicateRecipesVariant();
                duplicateRecipesVariantModalPage.FillField_DuplicateRecipesVariant(site, siteDestination, variant, variantBis);

                recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);
                recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();

                int nbVariants = recipeGeneralInfosPage.GetVariantTotal(decimalSeparatorValue);

                //Assert
                Assert.AreNotEqual(nbVariants, 1, "Le variant n'a pas été dupliqué.");
                Assert.IsTrue(recipeGeneralInfosPage.VerifySiteVariant(nbVariants, site), "Le variant d'origine est bien présent.");
                Assert.IsTrue(recipeGeneralInfosPage.VerifySiteVariant(nbVariants, siteDestination), "Le variant dupliqué est bien présent.");
            }
            finally
            {
                recipesPage = homePage.GoToMenus_Recipes();
                RecipeMassiveDeletePopup recipeMassiveDeletePopup = recipesPage.OpenMassiveDeletePopup();
                recipeMassiveDeletePopup.SetRecipeName(recipeName);
                recipeMassiveDeletePopup.ClickOnSearch();
                recipeMassiveDeletePopup.ClickOnSelectedAll();
                recipeMassiveDeletePopup.ClickDelete();
                recipeMassiveDeletePopup.ClickToConfirmDelete();


            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_RecipeSummary()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());
            recipeGeneralInfosPage.AddVariantWithSite(site, variant);
            var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
            recipeVariantPage.AddIngredient(ingredient);
            recipeGeneralInfosPage = recipeVariantPage.ClickOnGeneralInformation();
            recipeGeneralInfosPage.ClickOnSummary();
            recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();
            var isIngredientDisplayed = recipeVariantPage.IsIngredientDisplayed();

            //Assert
            Assert.IsTrue(isIngredientDisplayed, "Aucune variante n'est affichée.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_EditRecipeInformations()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["MenuVariantACE3"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();

            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            string updateRecipeName = "Rec-Updated" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);
            int updateNbPortions = nbPortions + 2;

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            try
            {
                var recipesPage = homePage.GoToMenus_Recipes();
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

                recipeGeneralInfosPage.AddVariantWithSite(site, variant);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(ingredient);

                recipeGeneralInfosPage = recipeVariantPage.ClickOnGeneralInformation();
                recipeGeneralInfosPage.ClickOnEditInformation();
                recipeGeneralInfosPage.UpdateRecipe(updateRecipeName, updateNbPortions.ToString());
                recipesPage = recipeGeneralInfosPage.BackToList();

                recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, updateRecipeName);
                recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
                recipeGeneralInfosPage.ClickOnEditInformation();

                string newRecipeName = recipeGeneralInfosPage.GetRecipeName();
                int newNbPortions = recipeGeneralInfosPage.GetNumberOfPortions();

                //Assert
                Assert.AreEqual(updateRecipeName, newRecipeName, "L'update du nom de la recette a échoué.");
                Assert.AreEqual(updateNbPortions, newNbPortions, "L'update du nombre de portions de la recette a échoué.");
            }
            finally
            {
                var recipesPage = homePage.GoToMenus_Recipes();
                recipesPage.MassiveDeleteRecipes(updateRecipeName, site, recipeType);

            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Edit_InactiveRecipe()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());
            recipeGeneralInfosPage.AddVariantWithSite(site, variant);

            var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
            recipeVariantPage.AddIngredient(ingredient);

            recipeGeneralInfosPage = recipeVariantPage.ClickOnGeneralInformation();
            recipeGeneralInfosPage.ClickOnEditInformation();
            recipeGeneralInfosPage.ClickOnActivatedCheckbox();
            recipesPage = recipeGeneralInfosPage.BackToList();

            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);
            recipesPage.Filter(RecipesPage.FilterType.ShowActive, true);
            Assert.AreNotEqual(recipeName, recipesPage.GetFirstRecipeName(), "La recette est visible dans la liste des recettes actives.");

            recipesPage.Filter(RecipesPage.FilterType.ShowInactive, true);
            Assert.AreEqual(recipeName, recipesPage.GetFirstRecipeName(), "La recette est visible dans la liste des recettes inactives.");
            Assert.IsFalse(recipesPage.CheckStatus(false), "La recette n'est pas inactive.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_AddVariantsInRecipe()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant1 = TestContext.Properties["MenuVariant"].ToString();
            string variant2 = TestContext.Properties["RecipeVariantBis"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            List<String> variants = new List<string>
            {
                variant1,
                variant2
            };
            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

            recipeGeneralInfosPage.AddMultipleVariants(site, variants);

            Assert.IsTrue(recipeGeneralInfosPage.IsVariantInAsideMenu(variants), "Les variantes n'ont pas été ajoutées.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_AddIngredientWithAllergenToRecipe()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string allergen = TestContext.Properties["RecipeAllergen"].ToString();
            string meal = TestContext.Properties["RecipeMeal"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            string guestName = "testAllergen";
            string ingredientName = "itemIntoleanceRecipekk";
            string variant = guestName + " - " + meal;
            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);

            bool isVariantCreated = false;

            //Arrange
            var homePage = LogInAsAdmin();
            //act
            var productionPage = homePage.GoToParameters_ProductionPage();

            try
            {
                // 1 - Création d'un guest avec allergie et d'un variant pour le guest               
                productionPage.GoToTab_Guest();
                if (productionPage.IsGuestPresent(guestName))
                {
                    productionPage.GoToTab_Variant();
                    productionPage.DeleteVariant(guestName);
                }
                else
                {
                    productionPage.AddNewGuest(guestName, true, "0");
                    productionPage.SetAllergy(allergen);
                }

                productionPage.GoToTab_Variant();
                isVariantCreated = productionPage.AddNewVariant(guestName, meal, site);

                // 2 -  Création d'un item pour contenant l'allergène désiré
                var itemName = CreateIngredientWithAllergen(homePage, site, allergen);

                // 3 - Création de la recette pour la variante
                var recipesPage = homePage.GoToMenus_Recipes();
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());
                recipeGeneralInfosPage.AddVariantWithSite(site, variant);

                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
                bool isIngredientAdded = recipeVariantPage.AddIngredientError(ingredientName);

                Assert.IsFalse(isIngredientAdded, "L'ingrédient a été ajouté à la recette malgré qu'il contienne un allergène.");
            }
            finally
            {
                if (isVariantCreated)
                {
                    productionPage = homePage.GoToParameters_ProductionPage();
                    productionPage.GoToTab_Variant();
                    productionPage.DeleteVariant(guestName);
                }
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_AddAllergenToIngredientInRecipe()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string allergen = TestContext.Properties["RecipeAllergen"].ToString();
            string ingredient = "ingredientForTestAllergen";
            string meal = TestContext.Properties["RecipeMeal"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            string guestName = "testAllergen";

            string variant = guestName + " - " + meal;
            string recipeName = recipeNameToday + "--" + new Random().Next().ToString();
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

            //Arrange
            var homePage = LogInAsAdmin();

            var productionPage = homePage.GoToParameters_ProductionPage();

            try
            {
                // 1 - Création d'un guest avec allergie et d'un variant pour le guest               
                productionPage.GoToTab_Guest();
                if (productionPage.IsGuestPresent(guestName))
                {
                    productionPage.UncheckAllAllergen(guestName);
                    productionPage.AddAllergen(allergen);
                    productionPage.GoToTab_Variant();
                    productionPage.DeleteVariant(guestName);
                }
                else
                {
                    productionPage.AddNewGuest(guestName, true, "0");
                    productionPage.SetAllergy(allergen);
                }

                productionPage.GoToTab_Variant();
                isVariantCreated = productionPage.AddNewVariant(guestName, meal, site);

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
                itemPage.Filter(ItemPage.FilterType.Search, ingredient);
                ItemGeneralInformationPage ItemGeneralInformationPage = itemPage.ClickOnFirstItems();
                ItemIntolerancePage itemIntolerancePage = ItemGeneralInformationPage.ClickOnIntolerancePage();
                itemIntolerancePage.UncheckAllAllergen();

                // 2 - Création de la recette pour la variante
                var recipesPage = homePage.GoToMenus_Recipes();
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());
                recipeGeneralInfosPage.AddVariantWithSite(site, variant);

                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
                recipeVariantPage.AddIngredient(ingredient);

                Assert.IsTrue(recipeVariantPage.IsIngredientDisplayed(), "L'ingrédient n'a pas été ajouté à la recette.");

                var itemGeneralInfosPage = recipeVariantPage.EditFirstIngredient(false);
                itemIntolerancePage = itemGeneralInfosPage.ClickOnIntolerancePage();
                itemIntolerancePage.AddAllergen(allergen);
                itemIntolerancePage.WaitPageLoading();
                if(itemIntolerancePage.IsDev())
                Assert.AreEqual(itemIntolerancePage.GetAllergenError(allergen), "Allergen prohibited for the guest " + guestName);

                var isAllergenSaved = itemIntolerancePage.SaveAllergenError(allergen, true);
                Assert.IsTrue(isAllergenSaved, "La donnée a pu être sauvegardée malgré l'allergène.");
            }
            finally
            {
                if (isVariantCreated)
                {
                    RecipesPage recipesPage = homePage.GoToMenus_Recipes();
                    recipesPage.MassiveDeleteRecipes(recipeName, site, recipeType);

                    ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
                    itemPage.Filter(ItemPage.FilterType.Search, ingredient);
                    ItemGeneralInformationPage itemGeneralInformationPage = itemPage.ClickOnFirstItem();
                    ItemIntolerancePage itemIntolerance = itemGeneralInformationPage.ClickOnIntolerancePage();
                    itemIntolerance.AddAllergen(allergen);

                    productionPage = homePage.GoToParameters_ProductionPage();
                    productionPage.GoToTab_Variant();
                    productionPage.DeleteVariant(guestName);

                }
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_AddIngredientToVariant()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant1 = TestContext.Properties["MenuVariant"].ToString();
            string variant2 = TestContext.Properties["RecipeVariantBis"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);


            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            List<String> variants = new List<string>
            {
                variant1,
                variant2
            };
            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

            recipeGeneralInfosPage.AddMultipleVariants(site, variants);
            Assert.IsTrue(recipeGeneralInfosPage.IsVariantInAsideMenu(variants), "Les variantes n'ont pas été ajoutées à la recette.");

            // Add an ingredient in the first variant only
            var recipeVariantPage = recipeGeneralInfosPage.SelectVariant(variant1);
            recipeVariantPage.AddIngredient(ingredient);
            Assert.IsTrue(recipeVariantPage.IsIngredientDisplayed(), "L'ingrédient n'a pas été ajouté à la variante 1.");

            // Verify that the ingredient is not present in the second variant
            recipeGeneralInfosPage = recipeVariantPage.ClickOnGeneralInformation();
            recipeVariantPage = recipeGeneralInfosPage.SelectVariant(variant2);
            Assert.IsFalse(recipeVariantPage.IsIngredientDisplayed(), "L'ingrédient est présent dans la variante 2 alors qu'il n'a pas été ajouté.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_AjoutSousRecetteError()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            string recipeName = "Recette" + new Random().Next().ToString();
            string recipeNameBis = recipeName + "Bis";
            int nbPortions = new Random().Next(1, 30);

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

            recipeGeneralInfosPage.AddVariantWithSite(site, variant);

            if (!recipeGeneralInfosPage.IsFirstVariantIsVisible())
            {
                recipeGeneralInfosPage.Validate();
            }
            var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

            recipeVariantPage.AddIngredient(ingredient);
            Assert.IsTrue(recipeVariantPage.IsIngredientDisplayed(), "L'ingrédient n'a pas été ajouté à la recette.");
            recipesPage = recipeVariantPage.BackToList();

            // Création d'une seconde recette
            recipesCreateModalPage = recipesPage.CreateNewRecipe();
            recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeNameBis, recipeType, nbPortions.ToString());

            recipeGeneralInfosPage.AddVariantWithSite(site, variant);
            recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

            // On ajoute la sous recette sans avoir décoché la checkbox (idem ajout ingrédient)
            Assert.IsFalse(recipeVariantPage.AddIngredientError(recipeName), "La sous-recette a pu être ajoutée malgré le fait que seuls les ingrédients peuvent être ajoutés.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_AjoutSousRecetteVariantDifferent()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string variantBis = TestContext.Properties["RecipeVariantBis"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            string recipeName = "Recette" + new Random().Next().ToString();
            string recipeNameBis = recipeName + "Bis";
            int nbPortions = new Random().Next(1, 30);

            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

            recipeGeneralInfosPage.AddVariantWithSite(site, variant);
            if (!recipeGeneralInfosPage.IsFirstVariantIsVisible())
            {
                recipeGeneralInfosPage.Validate();
            }
            var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

            recipeVariantPage.AddIngredient(ingredient);
            Assert.IsTrue(recipeVariantPage.IsIngredientDisplayed(), "L'ingrédient n'a pas été ajouté à la recette.");
            recipesPage = recipeVariantPage.BackToList();

            // Création d'une seconde recette
            recipesCreateModalPage = recipesPage.CreateNewRecipe();
            recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeNameBis, recipeType, nbPortions.ToString());

            recipeGeneralInfosPage.AddVariantWithSite(site, variantBis);
            if (!recipeGeneralInfosPage.IsFirstVariantIsVisible())
            {
                recipeGeneralInfosPage.Validate();
            }
            recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
            bool isSubRecipeAded = recipeVariantPage.AddSubRecipeError(recipeName);
            // On ajoute la sous recette sans avoir décoché la checkbox (idem ajout ingrédient)
            Assert.IsFalse(isSubRecipeAded, "La sous-recette a pu être ajoutée malgré le fait qu'elle n'ait pas la même variante que la recette.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_AjoutSousRecetteOK()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            string recipeName = "Recette" + new Random().Next().ToString();
            string recipeNameBis = recipeName + "Bis";
            int nbPortions = new Random().Next(1, 30);

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

            recipeGeneralInfosPage.AddVariantWithSite(site, variant);
            var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

            recipeVariantPage.AddIngredient(ingredient);
            Assert.IsTrue(recipeVariantPage.IsIngredientDisplayed(), "L'ingrédient n'a pas été ajouté à la recette.");
            recipesPage = recipeVariantPage.BackToList();

            // Création d'une seconde recette
            recipesCreateModalPage = recipesPage.CreateNewRecipe();
            recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeNameBis, recipeType, nbPortions.ToString());

            recipeGeneralInfosPage.AddVariantWithSite(site, variant);
            recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

            // On ajoute la sous recette
            recipeVariantPage.AddSubRecipe(recipeName);

            // Assert
            Assert.IsTrue(recipeVariantPage.IsSubRecipeDisplayed(), "La sous-recette n'a pas été ajouté à la recette.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_AddAllergenToGuestWithRecipe()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string allergen = TestContext.Properties["RecipeAllergen"].ToString();
            string meal = TestContext.Properties["RecipeMeal"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string guestName = "testAjoutAllergen";
            string variant = guestName + " - " + meal;
            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);

            bool isVariantCreated = false;

            //Arrange
            
            var homePage = LogInAsAdmin();

            var productionPage = homePage.GoToParameters_ProductionPage();

            try
            {
                // 1 - Création d'un guest avec allergie et d'un variant pour le guest               
                productionPage.GoToTab_Guest();
                if (productionPage.IsGuestPresent(guestName))
                {
                    productionPage.GoToTab_Variant();
                    productionPage.DeleteVariant(guestName);

                }
                else
                {
                    productionPage.AddNewGuest(guestName, false, "500");
                }

                productionPage.GoToTab_Variant();
                isVariantCreated = productionPage.AddNewVariant(guestName, meal, site);

                // 2 -  Création d'un item pour contenant l'allergène désiré
                string itemName = CreateIngredientWithAllergen(homePage, site, allergen);

                // 3 - Création de la recette pour la variante
                var recipesPage = homePage.GoToMenus_Recipes();
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());
                recipeGeneralInfosPage.AddVariantWithSite(site, variant);

                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
                recipeVariantPage.AddIngredient(itemName);

                Assert.IsTrue(recipeVariantPage.IsIngredientDisplayed(), "L'ingrédient n'a pas été ajouté à la recette.");

                // Ajout de l'allergène au convive
                homePage.Navigate();
                productionPage = homePage.GoToParameters_ProductionPage();
                productionPage.GoToTab_Guest();
                productionPage.EditGuest(guestName);
                productionPage.AddAllergenToGuest(allergen);

                // Assert
                Assert.IsTrue(productionPage.IsErrorMessageAllergen(), "L'allergène a pu être ajouté au convive malgré les recettes créées.");
            }
            finally
            {
                if (isVariantCreated)
                {
                    productionPage = homePage.GoToParameters_ProductionPage();
                    productionPage.GoToTab_Variant();
                    productionPage.DeleteVariant(guestName);
                }
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_DeleteIngredientFromRecipe()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

            recipeGeneralInfosPage.AddVariantWithSite(site, variant);
            var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

            recipeVariantPage.AddIngredient(ingredient);
            Assert.IsTrue(recipeVariantPage.IsIngredientDisplayed(), "L'ingrédient n'a pas été ajouté à la recette.");

            recipeVariantPage.DeleteFirstIngredient(false);
            Assert.IsFalse(recipeVariantPage.IsIngredientDisplayed(), "L'ingrédient n'a pas été supprimé de la recette.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_GetIngredientDetails()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            try
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

                recipeGeneralInfosPage.AddVariantWithSite(site, variant);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(ingredient);
                Assert.IsTrue(recipeVariantPage.IsIngredientDisplayed(), "L'ingrédient n'a pas été ajouté à la recette.");

                var itemGeneralInformationPage = recipeVariantPage.EditFirstIngredient(false);
                var itemName = itemGeneralInformationPage.GetItemName();
                Assert.AreEqual(itemName, ingredient, "Les détails de l'item sélectionné ne sont pas affichés.");
            }
            finally
            {
                recipesPage = homePage.GoToMenus_Recipes();
                recipesPage.MassiveDeleteRecipes(recipeName, site, recipeType);
            }


        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_AddCommentToIngredient()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);
            string comment = "This is a comment.";

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());
            recipeGeneralInfosPage.AddVariantWithSite(site, variant);
            var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
            recipeVariantPage.AddIngredient(ingredient);

            //Assert
            var isIngredientDisplayed = recipeVariantPage.IsIngredientDisplayed();
            Assert.IsTrue(isIngredientDisplayed, "L'ingrédient n'a pas été ajouté à la recette.");

            recipeVariantPage.AddCommentToIngredient(comment, false);
            recipesPage = recipeVariantPage.BackToList();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);
            recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
            recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();
            string newComment = recipeVariantPage.GetComment(false);

            //Assert
            Assert.AreEqual(comment, newComment, "Le commentaire n'a pas été ajouté à l'ingrédient.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_UpdateNetWeightValue()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);
            double newNetWeight = 10;

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

            recipeGeneralInfosPage.AddVariantWithSite(site, variant);
            var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

            recipeVariantPage.AddIngredient(ingredient);
            Assert.IsTrue(recipeVariantPage.IsIngredientDisplayed(), "L'ingrédient n'a pas été ajouté à la recette.");

            // Récupération des valeurs initiales pour chaque élément 
            recipeVariantPage.ClickOnFirstIngredient();
            double initNetWeight = recipeVariantPage.GetNetWeight(decimalSeparatorValue);
            double initNetQty = recipeVariantPage.GetNetQuantity(decimalSeparatorValue);
            double initGrossQty = recipeVariantPage.GetGrossQty(decimalSeparatorValue);
            double initTotalPortion = recipeVariantPage.GetTotalPortion(decimalSeparatorValue);
            string initPrice = recipeVariantPage.GetPrice();
            string initTotal1Portion = recipeVariantPage.GetTotalFor1Portion();
            string initPricePortion = recipeVariantPage.GetPricePortion();
            string initPrice1Portion = recipeVariantPage.GetPriceFor1Portion();

            // Modification de la valeur de NetWeight
            recipeVariantPage.SetNetWeight(newNetWeight.ToString());

            recipesPage = recipeVariantPage.BackToList();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);
            recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
            recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();
            recipeVariantPage.ClickOnFirstIngredient();

            Assert.AreEqual(newNetWeight, recipeVariantPage.GetNetWeight(decimalSeparatorValue), "La valeur 'NetWeight' n'a pas été modifiée.");
            Assert.AreNotEqual(newNetWeight, initNetWeight, "La valeur 'NetWeight' n'a pas été modifiée.");
            Assert.AreNotEqual(initNetQty, recipeVariantPage.GetNetQuantity(decimalSeparatorValue), "La valeur 'NetQty' n'a pas été modifiée.");
            Assert.AreNotEqual(initGrossQty, recipeVariantPage.GetGrossQty(decimalSeparatorValue), "La valeur 'GrossQty' n'a pas été modifiée.");
            Assert.AreNotEqual(initTotalPortion, recipeVariantPage.GetTotalPortion(decimalSeparatorValue), "La valeur 'TotalPortion' n'a pas été modifiée.");
            Assert.AreNotEqual(initPrice, recipeVariantPage.GetPrice(), "La valeur 'Price' n'a pas été modifiée.");
            Assert.AreNotEqual(initTotal1Portion, recipeVariantPage.GetTotalFor1Portion(), "La valeur 'TotalFor1Portion' n'a pas été modifiée.");
            Assert.AreNotEqual(initPricePortion, recipeVariantPage.GetPricePortion(), "La valeur 'PricePortion' n'a pas été modifiée.");
            Assert.AreNotEqual(initPrice1Portion, recipeVariantPage.GetPriceFor1Portion(), "La valeur 'PriceFor1Portion' n'a pas été modifiée.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_UpdateNetQuantityValue()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);
            double newNetQty = 20;

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

            recipeGeneralInfosPage.AddVariantWithSite(site, variant);
            if (!recipeGeneralInfosPage.IsFirstVariantIsVisible())
            {
                recipeGeneralInfosPage.Validate();
            }
            var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

            recipeVariantPage.AddIngredient(ingredient);
            Assert.IsTrue(recipeVariantPage.IsIngredientDisplayed(), "L'ingrédient n'a pas été ajouté à la recette.");

            // Récupération des valeurs initiales pour chaque élément 
            recipeVariantPage.ClickOnFirstIngredient();
            double initNetQty = recipeVariantPage.GetNetQuantity(decimalSeparatorValue);
            double initNetWeight = recipeVariantPage.GetNetWeight(decimalSeparatorValue);
            double initGrossQty = recipeVariantPage.GetGrossQty(decimalSeparatorValue);
            double initTotalPortion = recipeVariantPage.GetTotalPortion(decimalSeparatorValue);
            string initPrice = recipeVariantPage.GetPrice();
            string initTotal1Portion = recipeVariantPage.GetTotalFor1Portion();
            string initPricePortion = recipeVariantPage.GetPricePortion();
            string initPrice1Portion = recipeVariantPage.GetPriceFor1Portion();

            // Modification de la valeur de NetQuantity
            recipeVariantPage.SetNetQuantity(newNetQty.ToString());

            recipesPage = recipeVariantPage.BackToList();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);
            recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
            recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();
            recipeVariantPage.ClickOnFirstIngredient();

            Assert.AreEqual(newNetQty, recipeVariantPage.GetNetQuantity(decimalSeparatorValue), "La valeur 'NetQuantity' n'a pas été modifiée.");
            Assert.AreNotEqual(newNetQty, initNetQty, "La valeur 'NetQuantity' n'a pas été modifiée.");
            Assert.AreNotEqual(initNetWeight, recipeVariantPage.GetNetWeight(decimalSeparatorValue), "La valeur 'NetWeight' n'a pas été modifiée.");
            Assert.AreNotEqual(initGrossQty, recipeVariantPage.GetGrossQty(decimalSeparatorValue), "La valeur 'GrossQty' n'a pas été modifiée.");
            Assert.AreNotEqual(initTotalPortion, recipeVariantPage.GetTotalPortion(decimalSeparatorValue), "La valeur 'TotalPortion' n'a pas été modifiée.");
            Assert.AreNotEqual(initPrice, recipeVariantPage.GetPrice(), "La valeur 'Price' n'a pas été modifiée.");
            Assert.AreNotEqual(initTotal1Portion, recipeVariantPage.GetTotalFor1Portion(), "La valeur 'TotalFor1Portion' n'a pas été modifiée.");
            Assert.AreNotEqual(initPricePortion, recipeVariantPage.GetPricePortion(), "La valeur 'PricePortion' n'a pas été modifiée.");
            Assert.AreNotEqual(initPrice1Portion, recipeVariantPage.GetPriceFor1Portion(), "La valeur 'PriceFor1Portion' n'a pas été modifiée.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_UpdateYieldValue()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);
            double newYieldQty = new Random().Next(1, 30);

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

            recipeGeneralInfosPage.AddVariantWithSite(site, variant);
            var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

            recipeVariantPage.AddIngredient(ingredient);
            Assert.IsTrue(recipeVariantPage.IsIngredientDisplayed(), "L'ingrédient n'a pas été ajouté à la recette.");

            // Récupération des valeurs initiales pour chaque élément 
            recipeVariantPage.ClickOnFirstIngredient();
            double initYieldQty = recipeVariantPage.GetYield(decimalSeparatorValue);
            double initGrossQty = recipeVariantPage.GetGrossQty(decimalSeparatorValue);
            string initPrice = recipeVariantPage.GetPrice();
            string initPricePortion = recipeVariantPage.GetPricePortion();
            string initPrice1Portion = recipeVariantPage.GetPriceFor1Portion();

            // Modification de la valeur de Yield
            recipeVariantPage.SetYield(newYieldQty.ToString());

            recipesPage = recipeVariantPage.BackToList();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);
            recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
            recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();
            recipeVariantPage.ClickOnFirstIngredient();

            Assert.AreEqual(newYieldQty, recipeVariantPage.GetYield(decimalSeparatorValue), "La valeur 'Yield' n'a pas été modifiée.");
            Assert.AreNotEqual(newYieldQty, initYieldQty, "La valeur 'Yield' n'a pas été modifiée.");
            Assert.AreNotEqual(initGrossQty, recipeVariantPage.GetGrossQty(decimalSeparatorValue), "La valeur 'GrossQty' n'a pas été modifiée.");
            var newPrice = recipeVariantPage.GetPrice();
            Assert.AreNotEqual(initPrice, newPrice, "La valeur 'Price' n'a pas été modifiée.");
            Assert.AreNotEqual(initPricePortion, recipeVariantPage.GetPricePortion(), "La valeur 'PricePortion' n'a pas été modifiée.");
            Assert.AreNotEqual(initPrice1Portion, recipeVariantPage.GetPriceFor1Portion(), "La valeur 'PriceFor1Portion' n'a pas été modifiée.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_UpdateGrossQtyValue()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);
            double newGrossQty = 10;

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

            recipeGeneralInfosPage.AddVariantWithSite(site, variant);
            var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

            recipeVariantPage.AddIngredient(ingredient);
            Assert.IsTrue(recipeVariantPage.IsIngredientDisplayed(), "L'ingrédient n'a pas été ajouté à la recette.");

            // Récupération des valeurs initiales pour chaque élément 
            recipeVariantPage.ClickOnFirstIngredient();
            double initGrossQty = recipeVariantPage.GetGrossQty(decimalSeparatorValue);
            double initNetQty = recipeVariantPage.GetNetQuantity(decimalSeparatorValue);
            double initNetWeight = recipeVariantPage.GetNetWeight(decimalSeparatorValue);
            double initTotalPortion = recipeVariantPage.GetTotalPortion(decimalSeparatorValue);
            string initPrice = recipeVariantPage.GetPrice();
            string initTotal1Portion = recipeVariantPage.GetTotalFor1Portion();
            string initPricePortion = recipeVariantPage.GetPricePortion();
            string initPrice1Portion = recipeVariantPage.GetPriceFor1Portion();

            // Modification de la valeur de NetQuantity
            recipeVariantPage.SetGrossQty(newGrossQty.ToString());

            recipesPage = recipeVariantPage.BackToList();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);
            recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
            recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();
            recipeVariantPage.ClickOnFirstIngredient();

            Assert.AreEqual(newGrossQty, recipeVariantPage.GetGrossQty(decimalSeparatorValue), "La valeur 'GrossQty' n'a pas été modifiée.");
            Assert.AreNotEqual(newGrossQty, initGrossQty, "La valeur 'GrossQty' n'a pas été modifiée.");
            Assert.AreNotEqual(initNetWeight, recipeVariantPage.GetNetWeight(decimalSeparatorValue), "La valeur 'NetWeight' n'a pas été modifiée.");
            Assert.AreNotEqual(initNetQty, recipeVariantPage.GetNetQuantity(decimalSeparatorValue), "La valeur 'NetQuantity' n'a pas été modifiée.");
            Assert.AreNotEqual(initTotalPortion, recipeVariantPage.GetTotalPortion(decimalSeparatorValue), "La valeur 'TotalPortion' n'a pas été modifiée.");
            Assert.AreNotEqual(initPrice, recipeVariantPage.GetPrice(), "La valeur 'Price' n'a pas été modifiée.");
            Assert.AreNotEqual(initTotal1Portion, recipeVariantPage.GetTotalFor1Portion(), "La valeur 'TotalFor1Portion' n'a pas été modifiée.");
            Assert.AreNotEqual(initPricePortion, recipeVariantPage.GetPricePortion(), "La valeur 'PricePortion' n'a pas été modifiée.");
            Assert.AreNotEqual(initPrice1Portion, recipeVariantPage.GetPriceFor1Portion(), "La valeur 'PriceFor1Portion' n'a pas été modifiée.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_UpdateTotalPortionValue()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);
            double newTotalPortion = 10;

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

            recipeGeneralInfosPage.AddVariantWithSite(site, variant);
            var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

            recipeVariantPage.AddIngredient(ingredient);
            Assert.IsTrue(recipeVariantPage.IsIngredientDisplayed(), "L'ingrédient n'a pas été ajouté à la recette.");

            // Récupération des valeurs initiales pour chaque élément 
            recipeVariantPage.ClickOnFirstIngredient();
            double initTotalPortion = recipeVariantPage.GetTotalPortion(decimalSeparatorValue);
            double initGrossQty = recipeVariantPage.GetGrossQty(decimalSeparatorValue);
            double initNetQty = recipeVariantPage.GetNetQuantity(decimalSeparatorValue);
            double initNetWeight = recipeVariantPage.GetNetWeight(decimalSeparatorValue);
            string initPrice = recipeVariantPage.GetPrice();
            string initTotal1Portion = recipeVariantPage.GetTotalFor1Portion();
            string initPricePortion = recipeVariantPage.GetPricePortion();
            string initPrice1Portion = recipeVariantPage.GetPriceFor1Portion();

            // Modification de la valeur de NetQuantity
            recipeVariantPage.SetTotalPortion(newTotalPortion.ToString());

            recipesPage = recipeVariantPage.BackToList();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);
            recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
            recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();
            recipeVariantPage.ClickOnFirstIngredient();

            Assert.AreEqual(newTotalPortion, recipeVariantPage.GetTotalPortion(decimalSeparatorValue), "La valeur 'TotalPortion' n'a pas été modifiée.");
            Assert.AreNotEqual(newTotalPortion, initTotalPortion, "La valeur 'TotalPortion' n'a pas été modifiée.");
            Assert.AreNotEqual(initNetWeight, recipeVariantPage.GetNetWeight(decimalSeparatorValue), "La valeur 'NetWeight' n'a pas été modifiée.");
            Assert.AreNotEqual(initNetQty, recipeVariantPage.GetNetQuantity(decimalSeparatorValue), "La valeur 'NetQuantity' n'a pas été modifiée.");
            Assert.AreNotEqual(initGrossQty, recipeVariantPage.GetGrossQty(decimalSeparatorValue), "La valeur 'GrossQty' n'a pas été modifiée.");
            Assert.AreNotEqual(initPrice, recipeVariantPage.GetPrice(), "La valeur 'Price' n'a pas été modifiée.");
            Assert.AreNotEqual(initTotal1Portion, recipeVariantPage.GetTotalFor1Portion(), "La valeur 'TotalFor1Portion' n'a pas été modifiée.");
            Assert.AreNotEqual(initPricePortion, recipeVariantPage.GetPricePortion(), "La valeur 'PricePortion' n'a pas été modifiée.");
            Assert.AreNotEqual(initPrice1Portion, recipeVariantPage.GetPriceFor1Portion(), "La valeur 'PriceFor1Portion' n'a pas été modifiée.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_UpdateSalesPriceValue()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);
            double newSalesPrice = 10;

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

            recipeGeneralInfosPage.AddVariantWithSite(site, variant);
            var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

            recipeVariantPage.AddIngredient(ingredient);
            Assert.IsTrue(recipeVariantPage.IsIngredientDisplayed(), "L'ingrédient n'a pas été ajouté à la recette.");

            // Récupération des valeurs initiales pour chaque élément 
            recipeVariantPage.ClickOnFirstIngredient();
            double initSalesPrice = recipeVariantPage.GetSalesPrice(decimalSeparatorValue);

            // Modification de la valeur de NetQuantity
            recipeVariantPage.SetSalesPrice(newSalesPrice.ToString());

            recipesPage = recipeVariantPage.BackToList();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);
            recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
            recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();
            recipeVariantPage.ClickOnFirstIngredient();

            Assert.AreEqual(newSalesPrice, recipeVariantPage.GetSalesPrice(decimalSeparatorValue), "La valeur 'SalesPrice' n'a pas été modifiée.");
            Assert.AreNotEqual(newSalesPrice, initSalesPrice, "La valeur 'SalesPrice' n'a pas été modifiée.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_ImportIngredientWithEraseTargetRecipe()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string ingredientBis = TestContext.Properties["RecipeIngredientBis"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            string recipeName = "Recette" + new Random().Next().ToString();
            string recipeNameBis = recipeName + "Bis";
            int nbPortions = new Random().Next(1, 30);

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

            recipeGeneralInfosPage.AddVariantWithSite(site, variant);
            var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

            recipeVariantPage.AddIngredient(ingredient);
            Assert.IsTrue(recipeVariantPage.IsIngredientDisplayed(), "L'ingrédient n'a pas été ajouté à la recette.");
            Assert.AreEqual(recipeVariantPage.GetIngredientName(), ingredient, "L'ingrédient ajouté à la recette1 n'est pas celui attendu.");
            recipesPage = recipeVariantPage.BackToList();

            // Création d'une seconde recette
            recipesCreateModalPage = recipesPage.CreateNewRecipe();
            recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeNameBis, recipeType, nbPortions.ToString());

            recipeGeneralInfosPage.AddVariantWithSite(site, variant);
            recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
            recipeVariantPage.AddIngredient(ingredientBis);
            Assert.AreEqual(recipeVariantPage.GetIngredientName(), ingredientBis, "L'ingrédient ajouté à la recette2 n'est pas celui attendu.");

            var importIngredientPage = recipeVariantPage.ImportIngredient();
            importIngredientPage.FillFields_ImportIngredients(recipeName, variant, true);
            Assert.IsTrue(importIngredientPage.IsImportOK(), "L'import n'a pas été effectué.");
            recipeVariantPage = importIngredientPage.ConfirmImport();

            Assert.AreEqual(1, recipeVariantPage.GetNumberOfIngredient(), "Le nombre d'ingrédients de la recette est incorrect.");
            Assert.AreEqual(recipeVariantPage.GetIngredientName(), ingredient, "L'import d'ingrédient a échoué.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_ImportIngredientWithoutEraseTargetRecipe()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string ingredientBis = TestContext.Properties["RecipeIngredientBis"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string recipeName = "Recette" + new Random().Next().ToString();
            string recipeNameBis = recipeName + "Bis";
            int nbPortions = new Random().Next(1, 30);

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());
            recipeGeneralInfosPage.AddVariantWithSite(site, variant);
            var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
            recipeVariantPage.AddIngredient(ingredient);

            //Assert
            var ingredientDisplayed = recipeVariantPage.IsIngredientDisplayed();
            var ingredientName = recipeVariantPage.GetIngredientName();
            Assert.IsTrue(ingredientDisplayed, "L'ingrédient n'a pas été ajouté à la recette.");
            Assert.AreEqual(ingredientName, ingredient, "L'ingrédient ajouté à la recette1 n'est pas celui attendu.");
            recipesPage = recipeVariantPage.BackToList();

            // Création d'une seconde recette
            recipesCreateModalPage = recipesPage.CreateNewRecipe();
            recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeNameBis, recipeType, nbPortions.ToString());

            recipeGeneralInfosPage.AddVariantWithSite(site, variant);
            recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
            recipeVariantPage.AddIngredient(ingredientBis);
            ingredientName = recipeVariantPage.GetIngredientName();

            //Assert
            Assert.AreEqual(ingredientName, ingredientBis, "L'ingrédient ajouté à la recette2 n'est pas celui attendu.");

            var importIngredientPage = recipeVariantPage.ImportIngredient();
            importIngredientPage.FillFields_ImportIngredients(recipeName, variant, false);

            //Assert
            var isImportOK = importIngredientPage.IsImportOK();
            Assert.IsTrue(isImportOK, "L'import n'a pas été effectué.");
            recipeVariantPage = importIngredientPage.ConfirmImport();

            //Assert
            var numberOfIngredient = recipeVariantPage.GetNumberOfIngredient();
            Assert.AreEqual(2, numberOfIngredient, "Le nombre d'ingrédients de la recette est incorrect.");
            ingredientName = recipeVariantPage.GetIngredientName();
            Assert.AreEqual(recipeVariantPage.GetIngredientName(), ingredientBis, "L'ingrédient de la recette initiale a été supprimé.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_ImportIngredientWithEraseTargetRecipeAndSubRecipe()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string ingredientBis = TestContext.Properties["RecipeIngredientBis"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string recipeName = "Recette" + new Random().Next().ToString();
            string recipeNameBis = recipeName + "Bis";
            string recipeNameTer = recipeName + "Ter";
            int nbPortions = new Random().Next(1, 30);
            //Arrange
            HomePage homePage = LogInAsAdmin();
            //Act
            RecipesPage recipesPage = homePage.GoToMenus_Recipes();
            try
            {
                RecipesCreateModalPage recipesCreateModalPage = recipesPage.CreateNewRecipe();
                RecipeGeneralInformationPage recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());
                recipeGeneralInfosPage.AddVariantWithSite(site, variant);
                RecipesVariantPage recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
                recipeVariantPage.AddIngredient(ingredient);
                Assert.IsTrue(recipeVariantPage.IsIngredientDisplayed(), "L'ingrédient n'a pas été ajouté à la recette.");
                Assert.AreEqual(recipeVariantPage.GetIngredientName(), ingredient, "L'ingrédient ajouté à la recette1 n'est pas celui attendu.");
                recipesPage = recipeVariantPage.BackToList();
                // Création d'une seconde recette
                recipesCreateModalPage = recipesPage.CreateNewRecipe();
                recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeNameBis, recipeType, nbPortions.ToString());
                recipeGeneralInfosPage.AddVariantWithSite(site, variant);
                recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
                recipeVariantPage.AddIngredient(ingredientBis);
                Assert.AreEqual(recipeVariantPage.GetIngredientName(), ingredientBis, "L'ingrédient ajouté à la recette2 n'est pas celui attendu.");
                recipesPage = recipeVariantPage.BackToList();
                // Création d'une troisième recette
                recipesCreateModalPage = recipesPage.CreateNewRecipe();
                recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeNameTer, recipeType, nbPortions.ToString());
                recipeGeneralInfosPage.AddVariantWithSite(site, variant);
                recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
                recipeVariantPage.AddIngredient(ingredient);
                recipeVariantPage.AddSubRecipe(recipeNameBis);
                recipeVariantPage.AddSubRecipe(recipeName);
                Assert.IsTrue(recipeVariantPage.IsSubRecipeDisplayed(), "La sous-recette n'a pas été ajoutée à la recette3.");
                RecipeImportIngredientPage importIngredientPage = recipeVariantPage.ImportIngredient();
                importIngredientPage.FillFields_ImportIngredients(recipeName, variant, true);
                Assert.IsTrue(importIngredientPage.IsImportOK(), "L'import n'a pas été effectué.");
                recipeVariantPage = importIngredientPage.ConfirmImport();
                Assert.AreEqual(1, recipeVariantPage.GetNumberOfIngredient(), "Le nombre d'ingrédients de la recette est incorrect.");
                Assert.AreEqual(recipeVariantPage.GetIngredientName(), ingredient, "L'ingrédient de la recette initiale a été supprimé.");
                Assert.IsFalse(recipeVariantPage.IsSubRecipeDisplayed(), "La sous-recette n'a pas été supprimée.");
            }
            finally
            {
                recipesPage = homePage.GoToMenus_Recipes();
                recipesPage.MassiveDeleteRecipes(recipeNameTer,site,recipeType);
                recipesPage.MassiveDeleteRecipes(recipeNameBis, site,recipeType);
                recipesPage.MassiveDeleteRecipes(recipeName, site,recipeType);
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_ImportIngredientOtherSite()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = "RENFE - ALMUERZO";
            string variantBis = TestContext.Properties["RecipeVariantBis"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string ingredientBis = TestContext.Properties["RecipeIngredientBis"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string siteBis = TestContext.Properties["SiteToFlight"].ToString();
            string recipeName = "Recette" + new Random().Next().ToString() + site;
            string recipeNameBis = recipeName + siteBis;
            int nbPortions = new Random().Next(1, 30);

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());
            recipeGeneralInfosPage.AddVariantWithSite(site, variant);
            var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
            recipeVariantPage.AddIngredient(ingredient);

            //Assert
            var isIngredientDisplayed = recipeVariantPage.IsIngredientDisplayed();
            var ingredientName = recipeVariantPage.GetIngredientName();
            Assert.IsTrue(isIngredientDisplayed, "L'ingrédient n'a pas été ajouté à la recette.");
            Assert.AreEqual(ingredientName, ingredient, "L'ingrédient ajouté à la recette1 n'est pas celui attendu.");
            recipesPage = recipeVariantPage.BackToList();

            // Création d'une seconde recette
            recipesCreateModalPage = recipesPage.CreateNewRecipe();
            recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeNameBis, recipeType, nbPortions.ToString());
            recipeGeneralInfosPage.AddVariantWithSite(siteBis, variantBis);
            recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
            recipeVariantPage.AddIngredient(ingredientBis);

            //Assert
            Assert.AreEqual(recipeVariantPage.GetIngredientName(), ingredientBis, "L'ingrédient ajouté à la recette2 n'est pas celui attendu.");

            var importIngredientPage = recipeVariantPage.ImportIngredient();
            bool isImported = importIngredientPage.FillFields_ImportIngredientsError(recipeName);

            //Assert
            Assert.IsFalse(isImported, "L'import est possible alors que la recette à importer n'est pas du même site.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_CopyValuesForAllSites()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["Variant"].ToString();
            string site = TestContext.Properties["SiteLP"].ToString();
            string siteBis = TestContext.Properties["SiteBis"].ToString();

            string ingredient = "itemRecipeCopy";
            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);

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
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.Filter(ItemPage.FilterType.Search, ingredient);

            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(ingredient.ToString(), group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier, "ref");
                itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemCreatePackagingPage.FillField_CreateNewPackaging(siteBis, packagingName, storageQty, storageUnit, qty, supplier, "ref");
            }

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

            recipeGeneralInfosPage.AddVariantWithSite(site, variant);
            recipeGeneralInfosPage.AddVariantWithSite(siteBis, variant);

            // Add an ingredient in the first variant only
            var recipeVariantPage = recipeGeneralInfosPage.SelectVariantWithSite(site);
            recipeVariantPage.AddIngredient(ingredient);
            Assert.IsTrue(recipeVariantPage.IsIngredientDisplayed(), "L'ingrédient n'a pas été ajouté à la variante de site " + site + ".");

            // Verify that the ingredient is present in the second variant
            recipeGeneralInfosPage = recipeVariantPage.ClickOnGeneralInformation();
            recipeVariantPage = recipeGeneralInfosPage.SelectVariantWithSite(siteBis);
            Assert.IsTrue(recipeVariantPage.IsIngredientDisplayed(), "L'ingrédient n'a pas été ajouté à la variante de site " + site + ".");

            // Modify the quantity of the first variant
            recipeGeneralInfosPage = recipeVariantPage.ClickOnGeneralInformation();
            recipeVariantPage = recipeGeneralInfosPage.SelectVariantWithSite(site);
            recipeVariantPage.ClickOnFirstIngredient();
            recipeVariantPage.SetYield("60");
            var priceVariant1 = recipeVariantPage.GetPricePortion();

            WebDriver.Navigate().Refresh();
            // Verify the price on the second variant
            recipeGeneralInfosPage = recipeVariantPage.ClickOnGeneralInformation();
            recipeVariantPage = recipeGeneralInfosPage.SelectVariantWithSite(siteBis);
            var priceVariant2 = recipeVariantPage.GetPricePortion();

            Assert.AreNotEqual(priceVariant1, priceVariant2, "Le prix a été reporté sur l'autre variante sans activation de la fonctionnalité.");

            // Return on variant 1
            recipeGeneralInfosPage = recipeVariantPage.ClickOnGeneralInformation();
            recipeVariantPage = recipeGeneralInfosPage.SelectVariantWithSite(site);
            recipeVariantPage.CopyValuesForAllSites();

            // Verify the price on the second variant
            recipeGeneralInfosPage = recipeVariantPage.ClickOnGeneralInformation();
            recipeVariantPage = recipeGeneralInfosPage.SelectVariantWithSite(siteBis);
            var priceVariant2Bis = recipeVariantPage.GetPricePortion();

            Assert.AreEqual(priceVariant1, priceVariant2Bis, "Le prix n'a pas été reporté sur l'autre variante après activation de la fonctionnalité.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_DuplicateVariant()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string variantBis = TestContext.Properties["RecipeVariantBis"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            string recipeName = "Recette" + new Random().Next().ToString() + site;
            int nbPortions = new Random().Next(1, 30);

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

            recipeGeneralInfosPage.AddVariantWithSite(site, variant);
            var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

            recipeVariantPage.AddIngredient(ingredient);
            Assert.IsTrue(recipeVariantPage.IsIngredientDisplayed(), "L'ingrédient n'a pas été ajouté à la recette.");
            Assert.AreEqual(recipeVariantPage.GetIngredientName(), ingredient, "L'ingrédient ajouté à la recette1 n'est pas celui attendu.");

            var duplicateVariantModalPage = recipeVariantPage.DuplicateVariant();
            duplicateVariantModalPage.FillFields_DuplicateVariant(site, variantBis);
            Assert.IsTrue(duplicateVariantModalPage.IsDuplicateSaved(), "La duplication n'a pas été faite.");
            recipeVariantPage = duplicateVariantModalPage.CloseDuplicate();

            recipeGeneralInfosPage = recipeVariantPage.ClickOnGeneralInformation();
            recipeVariantPage = recipeGeneralInfosPage.SelectVariant(variantBis);

            Assert.IsTrue(recipeVariantPage.IsIngredientDisplayed(), "L'ingrédient n'a pas été ajouté à la variante.");
            Assert.AreEqual(recipeVariantPage.GetIngredientName(), ingredient, "L'ingrédient ajouté à la variante 2 n'est pas celui attendu.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_DuplicateVariantAlreadyExisting()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string variantBis = TestContext.Properties["RecipeVariantBis"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string ingredientBis = TestContext.Properties["RecipeIngredientBis"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            string recipeName = "Recette" + new Random().Next().ToString() + site;
            int nbPortions = new Random().Next(1, 30);

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            List<String> variants = new List<string>
            {
                variant,
                variantBis
            };
            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

            recipeGeneralInfosPage.AddMultipleVariants(site, variants);
            if (!recipeGeneralInfosPage.IsFirstVariantIsVisible())
            {
                recipeGeneralInfosPage.Validate();
            }
            var recipeVariantPage = recipeGeneralInfosPage.SelectVariant(variant);
            recipeVariantPage.AddIngredient(ingredient);
            Assert.IsTrue(recipeVariantPage.IsIngredientDisplayed(), "L'ingrédient n'a pas été ajouté à la recette.");
            Assert.AreEqual(recipeVariantPage.GetIngredientName(), ingredient, "L'ingrédient ajouté à la variante " + variant + " n'est pas celui attendu.");

            recipeGeneralInfosPage = recipeVariantPage.ClickOnGeneralInformation();
            recipeVariantPage = recipeGeneralInfosPage.SelectVariant(variantBis);
            recipeVariantPage.AddIngredient(ingredientBis);
            Assert.IsTrue(recipeVariantPage.IsIngredientDisplayed(), "L'ingrédient n'a pas été ajouté à la recette.");
            Assert.AreEqual(recipeVariantPage.GetIngredientName(), ingredientBis, "L'ingrédient ajouté à la variante " + variantBis + " n'est pas celui attendu.");

            recipeGeneralInfosPage = recipeVariantPage.ClickOnGeneralInformation();
            recipeVariantPage = recipeGeneralInfosPage.SelectVariant(variant);

            var duplicateVariantModalPage = recipeVariantPage.DuplicateVariant();
            duplicateVariantModalPage.FillFields_DuplicateVariant(site, variantBis);
            Assert.IsTrue(duplicateVariantModalPage.IsDuplicateSaved(), "La duplication n'a pas été faite.");
            recipeVariantPage = duplicateVariantModalPage.CloseDuplicate();

            recipeGeneralInfosPage = recipeVariantPage.ClickOnGeneralInformation();
            recipeVariantPage = recipeGeneralInfosPage.SelectVariant(variantBis);

            Assert.IsTrue(recipeVariantPage.IsIngredientDisplayed(), "Il n'y a pas d'ingrédients dans la variante dupliquée.");
            Assert.AreEqual(recipeVariantPage.GetIngredientName(), ingredient, "Le contenu de variante dupliquée n'a pas été remplacé.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_DuplicateVariantWithAllergen()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string allergen = TestContext.Properties["RecipeAllergen"].ToString();
            string meal = TestContext.Properties["RecipeMeal"].ToString();
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            string recipeName = "Recette" + new Random().Next().ToString() + site;
            int nbPortions = new Random().Next(1, 30);

            string guestName = "testAllergen";


            string variantBis = guestName + " - " + meal;

            bool isVariantCreated = false;

            //Arrange
            var homePage = LogInAsAdmin();

            var productionPage = homePage.GoToParameters_ProductionPage();
            try
            {
                // 1 - Création d'un guest avec allergie et d'un variant pour le guest               
                productionPage.GoToTab_Guest();
                if (productionPage.IsGuestPresent(guestName))
                {
                    productionPage.GoToTab_Variant();
                    productionPage.DeleteVariant(guestName);
                }
                else
                {
                    productionPage.AddNewGuest(guestName, true, "0");
                    productionPage.SetAllergy(allergen);
                }

                productionPage.GoToTab_Variant();
                isVariantCreated = productionPage.AddNewVariant(guestName, meal, site);

                // 2 -  Création d'un item pour contenant l'allergène désiré
                string itemName = CreateIngredientWithAllergen(homePage, site, allergen);

                // 3 - Création de la recette pour la variante
                var recipesPage = homePage.GoToMenus_Recipes();
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());
                recipeGeneralInfosPage.AddVariantWithSite(site, variant);

                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
                recipeVariantPage.AddIngredient(itemName);
                Assert.IsTrue(recipeVariantPage.IsIngredientDisplayed(), "L'ingrédient n'a pas été ajouté à la recette.");
                Assert.AreEqual(recipeVariantPage.GetIngredientName(), itemName, "L'ingrédient ajouté à la variante " + variant + " n'est pas celui attendu.");

                var duplicateVariantModalPage = recipeVariantPage.DuplicateVariant();
                duplicateVariantModalPage.FillFields_DuplicateVariant(site, variantBis);
                Assert.IsTrue(duplicateVariantModalPage.IsDuplicateInError(), "La duplication a été réalisée malgré la présence de l'allergène dans la recette source.");
            }
            finally
            {
                if (isVariantCreated)
                {
                    productionPage = homePage.GoToParameters_ProductionPage();
                    productionPage.GoToTab_Variant();
                    productionPage.DeleteVariant(guestName);
                }
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_AddProcedure()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);

            string procedure = "ajout procedure";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

            recipeGeneralInfosPage.AddVariantWithSite(site, variant);
            var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

            recipeVariantPage.AddIngredient(ingredient);
            Assert.IsTrue(recipeVariantPage.IsIngredientDisplayed(), "L'ingrédient n'a pas été ajouté à la recette.");

            // Récupération des valeurs initiales pour chaque élément 
            recipeVariantPage.ClickOnProcedureTab();
            string initProcedure = recipeVariantPage.GetProcedure();
            recipeVariantPage.SetProcedure(procedure);

            recipeVariantPage.ClickOnItemTab();
            Assert.IsTrue(recipeVariantPage.IsIngredientDisplayed());

            recipeVariantPage.ClickOnProcedureTab();
            string newProcedure = recipeVariantPage.GetProcedure();

            Assert.AreNotEqual(initProcedure, newProcedure, "La procédure a été modifiée.");
            Assert.AreEqual(newProcedure, procedure, "La procédure a été enregistrée.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_AddComposition()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);

            string composition = "ajout composition";

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

            recipeGeneralInfosPage.AddVariantWithSite(site, variant);
            var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

            recipeVariantPage.AddIngredient(ingredient);
            Assert.IsTrue(recipeVariantPage.IsIngredientDisplayed(), "L'ingrédient n'a pas été ajouté à la recette.");

            // Récupération des valeurs initiales pour chaque élément 
            recipeVariantPage.ClickOnCompositionTab();
            string initComposition = recipeVariantPage.GetComposition();
            recipeVariantPage.SetComposition(composition);

            recipeVariantPage.ClickOnItemTab();
            var isIngredientDisplayed = recipeVariantPage.IsIngredientDisplayed();
            Assert.IsTrue(isIngredientDisplayed);

            recipeVariantPage.ClickOnCompositionTab();
            string newComposition = recipeVariantPage.GetComposition();

            Assert.AreNotEqual(initComposition, newComposition, "La composition a été modifiée.");
            Assert.AreEqual(newComposition, composition, "La composition a été enregistrée.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_AddPicture()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);

            var path = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase);
            path = path.Substring(6);
            var imagePath = path + "\\PageObjects\\Menus\\Recipes\\test.jpg";

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

            recipeGeneralInfosPage.AddVariantWithSite(site, variant);
            var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

            recipeVariantPage.AddIngredient(ingredient);
            bool isIngredientDisplayed = recipeVariantPage.IsIngredientDisplayed();
            Assert.IsTrue(isIngredientDisplayed, "L'ingrédient n'a pas été ajouté à la recette.");

            // Récupération des valeurs initiales pour chaque élément 
            recipeVariantPage.ClickOnPictureTab();
            recipeVariantPage.AddPicture(imagePath);

            recipeVariantPage.ClickOnItemTab();

            isIngredientDisplayed = recipeVariantPage.IsIngredientDisplayed();
            Assert.IsTrue(isIngredientDisplayed, "Ingredient n'est pas affiché.");

            recipeVariantPage.ClickOnPictureTab();
            bool isPictureAdded = recipeVariantPage.IsPictureAdded();
            Assert.IsTrue(isPictureAdded, "L'image n'a pas été ajoutée.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Print_RecipeVariant_NewVersion()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);
            bool versionPrint = true;

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.ClearDownloads();

            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());
            recipeGeneralInfosPage.AddVariantWithSite(site, variant);
            var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
            recipeVariantPage.AddIngredient(ingredient);

            var reportPage = recipeVariantPage.Print(versionPrint);
            var isGenerated = reportPage.IsReportGenerated();
            reportPage.Close();

            //Assert
            Assert.IsTrue(isGenerated, "La recette n'a pas été imprimée.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Fold_Unfold_Recipes()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariant"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            if (recipesPage.CheckTotalNumber() < 5)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

                recipeGeneralInfosPage.AddVariantWithSite(site, variant);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(ingredient);
                recipesPage = recipeVariantPage.BackToList();
            }

            //Assert - unfold
            recipesPage.UnfoldAll();
            Assert.IsTrue(recipesPage.IsUnfoldAll(), "Le détail des items n'est pas affiché.");

            //Assert - fold
            recipesPage.FoldAll();
            Assert.IsTrue(recipesPage.IsFoldAll(), "Le détail des items n'est pas masqué.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Export_Recipe_SearchFilter_NewVersion()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);

            bool newVersionPrint = true;

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ClearDownloads();

            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

            recipeGeneralInfosPage.AddVariantWithSite(site, variant);
            var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

            recipeVariantPage.AddIngredient(ingredient);
            recipesPage = recipeVariantPage.BackToList();

            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);

            recipesPage.Export(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = recipesPage.GetExportExcelFile(taskFiles, false, false);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            int resultNumber = OpenXmlExcel.GetExportResultNumber(RECIPES_EXCEL_SHEET_NAME, filePath);
            bool result = OpenXmlExcel.ReadAllDataInColumn("Recipe", RECIPES_EXCEL_SHEET_NAME, filePath, recipeName);

            //Assert
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.IsTrue(result, MessageErreur.EXCEL_DONNEES_KO);
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Export_Recipe_RecipeTypeFilter_NewVersion()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeTypeBis"].ToString();
            string variant = TestContext.Properties["RecipeVariant"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);
            bool newVersionPrint = true;

            var homePage = LogInAsAdmin();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();

            recipesPage.Filter(RecipesPage.FilterType.RecipeType, recipeType);

            if (recipesPage.CheckTotalNumber() < 10)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

                recipeGeneralInfosPage.AddVariantWithSite(site, variant);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(ingredient);
                recipesPage = recipeVariantPage.BackToList();
                recipesPage.ResetFilter();
                recipesPage.Filter(RecipesPage.FilterType.RecipeType, recipeType);
            }
            DeleteAllFileDownload();
            recipesPage.ClearDownloads();
            recipesPage.Export(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = recipesPage.GetExportExcelFile(taskFiles, false, false);
            Assert.IsNotNull(correctDownloadedFile);
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            int resultNumber = OpenXmlExcel.GetExportResultNumber(RECIPES_EXCEL_SHEET_NAME, filePath);
            bool result = OpenXmlExcel.ReadAllDataInColumn("Recipe type", RECIPES_EXCEL_SHEET_NAME, filePath, recipeType);
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.IsTrue(result, MessageErreur.EXCEL_DONNEES_KO);
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_ExportSalesPrice()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeTypeBis"].ToString();
            string variant = TestContext.Properties["RecipeVariant"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);
            bool newVersionPrint = true;

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            try
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

                recipeGeneralInfosPage.AddVariantWithSite(site, variant);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(ingredient);
                recipesPage = recipeVariantPage.BackToList();
                recipesPage.ResetFilter();
                recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);

                DeleteAllFileDownload();
                recipesPage.ClearDownloads();
                recipesPage.ExportSalesPrice(newVersionPrint);
                DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
                FileInfo[] taskFiles = taskDirectory.GetFiles();
                var correctDownloadedFile = recipesPage.GetExportExcelFile(taskFiles, true, false);
                Assert.IsNotNull(correctDownloadedFile);

                var fileName = correctDownloadedFile.Name;
                var filePath = Path.Combine(downloadsPath, fileName);
                recipesPage.WaitPageLoading();
                int resultNumber = OpenXmlExcel.GetExportResultNumber("Recipes Sales Price", filePath);
                bool resultVriant = OpenXmlExcel.ReadAllDataInColumn("Variant", "Recipes Sales Price", filePath, "RENFE ALMUERZO");
                bool resultRecipeName = OpenXmlExcel.ReadAllDataInColumn("Recipe", "Recipes Sales Price", filePath, recipeName);
                bool resultSite = OpenXmlExcel.ReadAllDataInColumn("Site", "Recipes Sales Price", filePath, site);
                bool finalResult = resultVriant && resultRecipeName && resultSite;


                Assert.AreEqual(resultNumber, 1, MessageErreur.EXCEL_PAS_DE_DONNEES);
                Assert.IsTrue(finalResult, MessageErreur.EXCEL_DONNEES_KO);
            }
            finally
            {
                recipesPage = homePage.GoToMenus_Recipes();
                recipesPage.MassiveDeleteRecipes(recipeName, site, recipeType);
            }


        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_ExportRobot()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeTypeBis"].ToString();
            string typeOfGuest = "FLIGHT";
            string typeOfMeal = "FLIGHT";
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            bool newVersionPrint = true;

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();


            recipesPage.Filter(RecipesPage.FilterType.RecipeType, recipeType);
            recipesPage.Filter(RecipesPage.FilterType.TypeOfGuest, typeOfGuest);
            recipesPage.Filter(RecipesPage.FilterType.TypeOfMeal, typeOfMeal);

            var checkCount = recipesPage.CheckTotalNumber();

            DeleteAllFileDownload();
            recipesPage.ClearDownloads();
            recipesPage.ExportRobot(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = recipesPage.GetExportExcelFile(taskFiles, false, true);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            int resultNumber = OpenXmlExcel.GetExportResultNumber("Recipe", filePath);
            bool result = OpenXmlExcel.ReadAllDataInColumn("Recipe type", "Recipe", filePath, recipeType);

            //Assert
            Assert.AreEqual(resultNumber, checkCount, MessageErreur.EXCEL_PAS_DE_DONNEES);

            Assert.IsTrue(result, MessageErreur.EXCEL_DONNEES_KO);
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_ImportSalesPrice()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeTypeBis"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            bool newVersionPrint = true;

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();


            recipesPage.Filter(RecipesPage.FilterType.RecipeType, recipeType);

            var checkCount = recipesPage.CheckTotalNumber();

            DeleteAllFileDownload();
            recipesPage.ClearDownloads();
            recipesPage.ExportSalesPrice(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = recipesPage.GetExportExcelFile(taskFiles, true, false);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            recipesPage.UpdateFileExported(filePath, "Recipes Sales Price", correctDownloadedFile);
            var isRecipeImportd = recipesPage.ImportSalesPrice(filePath);
            Assert.IsTrue(isRecipeImportd, MessageErreur.EXCEL_DONNEES_KO);
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_DuplicateVariant_RecipeWithSubRecipe()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string variantBis = TestContext.Properties["RecipeVariant"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string siteBis = TestContext.Properties["SiteBis"].ToString();
            string siteDestination = TestContext.Properties["SiteLP"].ToString();

            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            string recipeNameBis = recipeName + "Bis";

            int nbPortions = new Random().Next(1, 30);

            //Arrange
            var homePage = LogInAsAdmin();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();

            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

            recipeGeneralInfosPage.AddVariantWithSite(siteBis, variant);
            var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

            recipeVariantPage.AddIngredient(ingredient);
            recipesPage = recipeVariantPage.BackToList();

            // Création d'une seconde recette
            recipesCreateModalPage = recipesPage.CreateNewRecipe();
            recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeNameBis, recipeType, nbPortions.ToString());

            recipeGeneralInfosPage.AddVariantWithSite(siteBis, variant);
            recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

            // On ajoute la sous recette
            recipeVariantPage.AddSubRecipe(recipeName);
            Assert.IsTrue(recipeVariantPage.IsSubRecipeDisplayed(), "La sous-recette n'a pas été ajouté à la recette.");

            try
            {
                //Duplicate variant
                var duplicateVariantModalPage = recipeVariantPage.DuplicateVariant();
                duplicateVariantModalPage.FillFields_DuplicateVariant(siteDestination, variantBis);
                Assert.IsTrue(duplicateVariantModalPage.IsDuplicateSaved(), "La duplication n'a pas été faite.");
                recipeVariantPage = duplicateVariantModalPage.CloseDuplicate();
                recipesPage = recipeVariantPage.BackToList();

                recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeNameBis);
                recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();

                int nbVariants = recipeGeneralInfosPage.GetVariantTotal(decimalSeparatorValue);

                //Assert
                Assert.AreNotEqual(nbVariants, 1, "Le variant n'a pas été dupliqué.");
                Assert.IsTrue(recipeGeneralInfosPage.VerifySiteVariant(nbVariants, siteBis), "Le variant d'origine est bien présent.");
                Assert.IsTrue(recipeGeneralInfosPage.VerifySiteVariant(nbVariants, siteDestination), "Le variant dupliqué est bien présent.");

                // Check if sous recette dans nouveau variant
                recipeGeneralInfosPage = recipeVariantPage.ClickOnGeneralInformation();
                recipeVariantPage = recipeGeneralInfosPage.SelectVariant(variant);
                Assert.IsTrue(recipeVariantPage.IsSubRecipeDisplayed(), "La sous-recette n'est pas présente dans le variant dupliqué.");

                // Edit sous recette et vérifier nom
                recipeVariantPage.EditFirstSubRecipe(false);
                Assert.AreNotEqual(recipeVariantPage.GetRecipeName(), recipeName, "La sous-recette peut-être éditée.");
                recipeVariantPage.Close();

                // Print Duplicate Variant-Recipe
                var reportPage = recipeVariantPage.Print(true);
                var isGenerated = reportPage.IsReportGenerated();
                reportPage.Close();
                Assert.IsTrue(isGenerated, "La recette n'a pas été imprimée.");
            }
            finally
            {
                recipesPage = homePage.GoToMenus_Recipes();
                recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);
                recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
                recipeGeneralInfosPage.ClickOnEditInformation();
                recipeGeneralInfosPage.ClickOnActivatedCheckbox();
                recipeGeneralInfosPage.BackToList();

                recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeNameBis);
                recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
                recipeGeneralInfosPage.ClickOnEditInformation();
                recipeGeneralInfosPage.ClickOnActivatedCheckbox();
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_DuplicateRecipesVariant_RecipeWithSubRecipe()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["SiteBis"].ToString();
            string siteDestination = TestContext.Properties["SiteLP"].ToString();

            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            string recipeNameBis = recipeName + "Bis";

            int nbPortions = new Random().Next(1, 30);


            //Arrange

            var homePage = LogInAsAdmin();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();
            ParametersProduction parametersProduction = homePage.GoToParameters_ProductionPage();
            parametersProduction.GoToTab_Variant();
            parametersProduction.FilterSite(site);
            var variant = parametersProduction.GetFirstVariant();
            parametersProduction.FilterSite(siteDestination);
            var variantBis = parametersProduction.GetFirstVariant();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();

            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

            recipeGeneralInfosPage.AddVariantWithSite(site, variant);
            var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

            recipeVariantPage.AddIngredient(ingredient);
            recipesPage = recipeVariantPage.BackToList();

            // Création d'une seconde recette
            recipesCreateModalPage = recipesPage.CreateNewRecipe();
            recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeNameBis, recipeType, nbPortions.ToString());

            recipeGeneralInfosPage.AddVariantWithSite(site, variant);
            recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

            // On ajoute la sous recette
            recipeVariantPage.AddSubRecipe(recipeName);

            try
            {
                // Assert
                Assert.IsTrue(recipeVariantPage.IsSubRecipeDisplayed(), "La sous-recette n'a pas été ajouté à la recette.");
                recipesPage = recipeVariantPage.BackToList();

                recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeNameBis);

                var duplicateRecipesVariantModalPage = recipesPage.DuplicateRecipesVariant();
                duplicateRecipesVariantModalPage.FillField_DuplicateRecipesVariant(site, siteDestination, variant, variantBis);

                recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();

                int nbVariants = recipeGeneralInfosPage.GetVariantTotal(decimalSeparatorValue);

                //Assert
                Assert.AreNotEqual(nbVariants, 1, "Le variant n'a pas été dupliqué.");
                Assert.IsTrue(recipeGeneralInfosPage.VerifySiteVariant(nbVariants, site), "Le variant d'origine est bien présent.");
                Assert.IsTrue(recipeGeneralInfosPage.VerifySiteVariant(nbVariants, siteDestination), "Le variant dupliqué est bien présent.");

                // Check if sous recette dans nouveau variant
                recipeGeneralInfosPage = recipeVariantPage.ClickOnGeneralInformation();
                recipeVariantPage = recipeGeneralInfosPage.SelectVariant(variantBis);
                Assert.IsTrue(recipeVariantPage.IsSubRecipeDisplayed(), "La sous-recette n'est pas présente dans le variant dupliqué.");

                // Edit sous recette et vérifier nom
                recipeVariantPage.EditFirstSubRecipe(false);
                Assert.AreNotEqual(recipeVariantPage.GetRecipeName(), recipeName, "La sous-recette peut-être éditée.");
                recipeVariantPage.Close();

                // Print Duplicate Variant-Recipe
                recipeVariantPage.ClearDownloads();
                var reportPage = recipeVariantPage.Print(true);
                var isGenerated = reportPage.IsReportGenerated();
                reportPage.Close();
                Assert.IsTrue(isGenerated, "La recette n'a pas été imprimée.");
            }
            finally
            {
                recipesPage = homePage.GoToMenus_Recipes();
                RecipeMassiveDeletePopup recipeMassiveDeletePopup = recipesPage.OpenMassiveDeletePopup();
                recipeMassiveDeletePopup.SetRecipeName(recipeNameBis);
                recipeMassiveDeletePopup.ClickOnSearch();
                recipeMassiveDeletePopup.ClickOnSelectedAll();
                recipeMassiveDeletePopup.ClickDelete();
                recipeMassiveDeletePopup.ClickToConfirmDelete();

                recipeMassiveDeletePopup = recipesPage.OpenMassiveDeletePopup();
                recipeMassiveDeletePopup.SetRecipeName(recipeName);
                recipeMassiveDeletePopup.ClickOnSearch();
                recipeMassiveDeletePopup.ClickOnSelectedAll();
                recipeMassiveDeletePopup.ClickDelete();
                recipeMassiveDeletePopup.ClickToConfirmDelete();
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_UseCase_ReplaceByAnotherItem()
        {
            //Prepare
            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            string menuName = "Menu_" + DateUtils.Now.ToString("dd/MM/yyyy") + "-" + new Random().Next();
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string serviceName = TestContext.Properties["MenuService"].ToString();
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            int nbPortions = new Random().Next(1, 30);

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);

            //Act
            try
            {
                var recipesPage = homePage.GoToMenus_Recipes();
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());
                recipeGeneralInfosPage.AddVariantWithSite(site, variant);

                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
                recipeVariantPage.AddIngredient("a");

                Assert.IsTrue(recipeVariantPage.IsIngredientDisplayed(), "L'ingrédient n'a pas été ajouté à la recette.");
                // 0. créer un menu
                MenusPage menuIndex = homePage.GoToMenus_Menus();
                MenusCreateModalPage create = menuIndex.MenuCreatePage();
                MenusDayViewPage menuPage = create.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now, site, variant, serviceName);
                menuPage.AddRecipe(recipeName);

                // 1. Etre sur une recette / Onglet Uses Case
                recipesPage = homePage.GoToMenus_Recipes();
                // récupérer le recipe (*ARROZ CON CONEJO PMI)
                recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);
                string recipe = recipesPage.GetFirstRecipeName();
                RecipeGeneralInformationPage generalInfo = recipesPage.SelectFirstRecipe();
                Assert.AreEqual(recipeName, generalInfo.GetRecipeName(), "Bad Recipe Name");
                RecipeUseCaseTab useCase = generalInfo.ClickOnUseCaseTab();
                // FIXME KO sur DEV ticket 12981
                if (useCase.isElementVisible(By.XPath("//*/b[text()='Some error(s) have occured :']")))
                {
                    Assert.Fail("popup Some error(s) have occured");
                }

                // 2. Cocher une Recipe
                useCase.Filter(RecipeUseCaseTab.FilterType.SearchMenu, "Menu_");
                useCase.WaitPageLoading();
                Assert.IsTrue(useCase.CheckTotalNumber() > 0);
                // récupérer le menu (Menu-14/11/2022-1134223950)
                string menu = useCase.GetFirstMenuName();
                useCase.SelectFirstMenuCheckBox();

                // 2. Cliquer sur "Replace By Another Recipe
                // 3. Une pop up s'ouvre et remplir et cliquer sur replace with...
                useCase.ReplaceByAnotherRecipe("RecipeForTestMenu");

                // on rétablit
                recipesPage = useCase.BackToList();
                recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, "RecipeForTestMenu");
                generalInfo = recipesPage.SelectFirstRecipe();
                useCase = generalInfo.ClickOnUseCaseTab();

                useCase.Filter(RecipeUseCaseTab.FilterType.DateFrom, DateUtils.Now.AddYears(-1));
                useCase.Filter(RecipeUseCaseTab.FilterType.SearchMenu, menu);
                Assert.IsTrue(useCase.CheckTotalNumber() > 0);

                useCase.SelectMenuCheckBox(menu);

                useCase.ReplaceByAnotherRecipe(recipe);
            }
            finally
            {

                var menusPage = homePage.GoToMenus_Menus();
                menusPage.MassiveDeleteMenus(menuName, site, variant);

                var recipesPage = homePage.GoToMenus_Recipes();
                recipesPage.MassiveDeleteRecipes(recipeName, site, recipeType);

            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_MassiveDelete()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

            recipeGeneralInfosPage.AddVariantWithSite(site, variant);
            var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

            recipeVariantPage.AddIngredient(ingredient);
            recipesPage = recipeVariantPage.BackToList();

            recipesPage.MassiveDeleteRecipes(recipeName, site, recipeType);
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);

            //Assert
            Assert.AreEqual(0, recipesPage.CheckTotalNumber(), "La massive delete ne fonctionne pas.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Massivedelete_namesearch()
        {
            string recipeName = "ABADEJO A LA PROVENZAL ALC";

            //Home page
            var homePage = LogInAsAdmin();

            var recipePage = homePage.GoToMenus_Recipes();
            var recipeMassiveDelete = recipePage.OpenMassiveDeletePopup();

            recipeMassiveDelete.SetRecipeName(recipeName);
            recipeMassiveDelete.ClickOnSearch();

            int rowNumber = 0;
            RecipeMassiveDeleteRowResult recipeRowResult = recipeMassiveDelete.GetRowResultInfo(rowNumber);

            Assert.AreEqual(recipeName, recipeRowResult.RecipeName, $"Le nom du recipe n'est pas égal à {recipeName} pour la ligne de résultat {rowNumber}");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Massivedelete_sitesearch()
        {
            string siteName = TestContext.Properties["SiteACE"].ToString();

            //Home page
            var homePage = LogInAsAdmin();

            var recipePage = homePage.GoToMenus_Recipes();
            var recipeMassiveDelete = recipePage.OpenMassiveDeletePopup();

            recipeMassiveDelete.SelectSiteByName(siteName, false);
            recipeMassiveDelete.ClickOnSearch();

            int rowNumber = 0;
            RecipeMassiveDeleteRowResult recipeRowResult = recipeMassiveDelete.GetRowResultInfo(rowNumber);

            while (recipeRowResult != null)
            {
                Assert.AreEqual(siteName, recipeRowResult.SiteName, $"Le nom du site n'est pas égal à {siteName} pour la ligne de résultat {rowNumber}");
                rowNumber++;
                recipeRowResult = recipeMassiveDelete.GetRowResultInfo(rowNumber);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Massivedelete_inactivesitessearch()
        {
            //Home page
            var homePage = LogInAsAdmin();

            var recipePage = homePage.GoToMenus_Recipes();
            var recipeMassiveDelete = recipePage.OpenMassiveDeletePopup();

            recipeMassiveDelete.ClickOnInactiveSiteCheck();
            recipeMassiveDelete.SelectAllInactiveSites();
            recipeMassiveDelete.ClickOnSearch();

            int rowNumber = 0;
            RecipeMassiveDeleteRowResult recipeRowResult = recipeMassiveDelete.GetRowResultInfo(rowNumber);

            while (recipeRowResult != null)
            {
                Assert.IsTrue(recipeRowResult.IsSiteInactive, $"Le site de la ligne de résultat {rowNumber} est active alors que le filtre est sur 'site inactif'.");
                rowNumber++;
                recipeRowResult = recipeMassiveDelete.GetRowResultInfo(rowNumber);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Massivedelete_Variantsearch()
        {
            string variantName = TestContext.Properties["MenuVariantACE1"].ToString();

            //Home page
            var homePage = LogInAsAdmin();
            var recipePage = homePage.GoToMenus_Recipes();
            var recipeMassiveDelete = recipePage.OpenMassiveDeletePopup();

            recipeMassiveDelete.SelectVariantByName(variantName, false);
            recipeMassiveDelete.ClickOnSearch();

            int rowNumber = 0;
            RecipeMassiveDeleteRowResult recipeRowResult = recipeMassiveDelete.GetRowResultInfo(rowNumber);

            while (recipeRowResult != null)
            {
                Assert.AreEqual(variantName, recipeRowResult.VariantName, $"Le nom du variant n'est pas égal à {variantName} pour la ligne de résultat {rowNumber}");
                rowNumber++;
                recipeRowResult = recipeMassiveDelete.GetRowResultInfo(rowNumber);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Massivedelete_RecipeType()
        {
            string recipeType = TestContext.Properties["RecipeType"].ToString();

            //Home page
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var recipePage = homePage.GoToMenus_Recipes();
            var recipeMassiveDelete = recipePage.OpenMassiveDeletePopup();

            recipeMassiveDelete.SelectRecipeTypeByName(recipeType, false);
            recipeMassiveDelete.ClickOnSearch();
            RecipeGeneralInformationPage rGIPage = recipeMassiveDelete.ClickOnRecipeLinkFromRow(1);

            rGIPage.ClickOnEditInformation();
            string linkRecipeType = rGIPage.GetRecipeType();

            Assert.AreEqual(linkRecipeType, recipeType, $"Le type de recette n'est pas égal à {recipeType} pour la ligne de résultat {1}");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Massivedelete_status_activeRecipe()
        {
            //Home page
            var homePage = LogInAsAdmin();

            var recipePage = homePage.GoToMenus_Recipes();
            var recipeMassiveDelete = recipePage.OpenMassiveDeletePopup();

            recipeMassiveDelete.SelectStatus(MassiveDeleteStatus.ActiveRecipes.ToString());
            recipeMassiveDelete.ClickOnSearch();

            int rowNumber = 0;
            RecipeMassiveDeleteRowResult recipeRowResult = recipeMassiveDelete.GetRowResultInfo(rowNumber);

            while (recipeRowResult != null)
            {
                Assert.IsFalse(recipeRowResult.IsRecipeInactive, $"La recette est inactive pour la ligne de résultat {rowNumber} alors que le filtre est sur 'active recipes'.");
                rowNumber++;
                recipeRowResult = recipeMassiveDelete.GetRowResultInfo(rowNumber);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Massivedelete_status_inactiveRecipes()
        {
            //Home page
            var homePage = LogInAsAdmin();

            var recipePage = homePage.GoToMenus_Recipes();
            var recipeMassiveDelete = recipePage.OpenMassiveDeletePopup();

            recipeMassiveDelete.SelectStatus(MassiveDeleteStatus.InactiveRecipes.ToString());
            recipeMassiveDelete.ClickOnSearch();

            int rowNumber = 0;
            RecipeMassiveDeleteRowResult recipeRowResult = recipeMassiveDelete.GetRowResultInfo(rowNumber);

            while (recipeRowResult != null)
            {
                Assert.IsTrue(recipeRowResult.IsRecipeInactive, $"La recette est active pour la ligne de résultat {rowNumber} alors que le filtre est sur 'inactive recipes'.");
                rowNumber++;
                recipeRowResult = recipeMassiveDelete.GetRowResultInfo(rowNumber);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Massivedelete_status_Combo()
        {
            //Home page
            var homePage = LogInAsAdmin();

            var recipePage = homePage.GoToMenus_Recipes();
            var recipeMassiveDelete = recipePage.OpenMassiveDeletePopup();

            recipeMassiveDelete.SelectStatus(MassiveDeleteStatus.ActiveRecipes.ToString());
            recipeMassiveDelete.SelectStatus(MassiveDeleteStatus.Used.ToString());
            recipeMassiveDelete.SelectStatus(MassiveDeleteStatus.EmptyRecipe.ToString());
            recipeMassiveDelete.ClickOnSearch();

            int rowNumber = 0;
            RecipeMassiveDeleteRowResult recipeRowResult = recipeMassiveDelete.GetRowResultInfo(rowNumber);

            while (recipeRowResult != null)
            {
                bool isOK = (recipeRowResult.IsRecipeInactive == false) && (recipeRowResult.UseCase > 0) && (recipeRowResult.Weight == 0.0);
                Assert.IsTrue(isOK, $"La recette pour la ligne de résultat {rowNumber} ne répond pas à plusieurs filtres en combo.");
                rowNumber++;
                recipeRowResult = recipeMassiveDelete.GetRowResultInfo(rowNumber);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Massivedelete_status_Empty()
        {
            //Home page
            var homePage = LogInAsAdmin();

            var recipePage = homePage.GoToMenus_Recipes();
            var recipeMassiveDelete = recipePage.OpenMassiveDeletePopup();

            recipeMassiveDelete.SelectStatus(MassiveDeleteStatus.EmptyRecipe.ToString());
            recipeMassiveDelete.ClickOnSearch();

            int rowNumber = 0;
            RecipeMassiveDeleteRowResult recipeRowResult = recipeMassiveDelete.GetRowResultInfo(rowNumber);

            while (recipeRowResult != null)
            {
                Assert.AreEqual(0.0, recipeRowResult.Weight, $"La recette contient un item pour la ligne de résultat {rowNumber} alors que le filtre est sur 'empty recipe'.");
                rowNumber++;
                recipeRowResult = recipeMassiveDelete.GetRowResultInfo(rowNumber);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Massivedelete_status_FlightTypeDatasheet()
        {
            //Home page
            var homePage = LogInAsAdmin();
            var recipePage = homePage.GoToMenus_Recipes();
            var recipeMassiveDelete = recipePage.OpenMassiveDeletePopup();

            recipeMassiveDelete.SelectStatus(MassiveDeleteStatus.FlightTypeUnrelatedDS.ToString());
            recipeMassiveDelete.ClickOnSearch();

            int rowNumber = 0;
            RecipeMassiveDeleteRowResult recipeRowResult = recipeMassiveDelete.GetRowResultInfo(rowNumber);

            while (recipeRowResult != null)
            {
                Assert.AreEqual(0, recipeRowResult.UseCase, $"La recette est utilisée pour la ligne de résultat {rowNumber} alors que le filtre est sur 'flight type and not related to datasheet'.");
                rowNumber++;
                recipeRowResult = recipeMassiveDelete.GetRowResultInfo(rowNumber);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Massivedelete_status_inactivesites()
        {
            //Home page
            var homePage = LogInAsAdmin();

            var recipePage = homePage.GoToMenus_Recipes();
            var recipeMassiveDelete = recipePage.OpenMassiveDeletePopup();

            recipeMassiveDelete.SelectStatus(MassiveDeleteStatus.OnlyInactiveSites.ToString());
            recipeMassiveDelete.ClickOnSearch();

            int rowNumber = 0;
            RecipeMassiveDeleteRowResult recipeRowResult = recipeMassiveDelete.GetRowResultInfo(rowNumber);

            while (recipeRowResult != null)
            {
                Assert.IsTrue(recipeRowResult.IsSiteInactive, $"La recette pour la ligne de résultat {rowNumber} a un site actif active alors que le filtre est sur 'inactive sites'.");
                rowNumber++;
                recipeRowResult = recipeMassiveDelete.GetRowResultInfo(rowNumber);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Massivedelete_status_NoVariant()
        {
            //Home page
            var homePage = LogInAsAdmin();

            var recipePage = homePage.GoToMenus_Recipes();
            var recipeMassiveDelete = recipePage.OpenMassiveDeletePopup();

            recipeMassiveDelete.SelectStatus(MassiveDeleteStatus.RecipeWithoutVariant.ToString());
            recipeMassiveDelete.ClickOnSearch();

            int rowNumber = 0;
            RecipeMassiveDeleteRowResult recipeRowResult = recipeMassiveDelete.GetRowResultInfo(rowNumber);

            while (recipeRowResult != null)
            {
                bool isOK = string.IsNullOrEmpty(recipeRowResult.SiteName) && string.IsNullOrEmpty(recipeRowResult.VariantName);
                Assert.IsTrue(isOK, $"La recette pour la ligne de résultat {rowNumber} a un variant alors que le filtre est sur 'no variant'.");
                rowNumber++;
                recipeRowResult = recipeMassiveDelete.GetRowResultInfo(rowNumber);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Massivedelete_status_unused()
        {
            //Home page
            var homePage = LogInAsAdmin();

            var recipePage = homePage.GoToMenus_Recipes();
            var recipeMassiveDelete = recipePage.OpenMassiveDeletePopup();

            recipeMassiveDelete.SelectStatus(MassiveDeleteStatus.Unused.ToString());
            recipeMassiveDelete.ClickOnSearch();

            int rowNumber = 0;
            RecipeMassiveDeleteRowResult recipeRowResult = recipeMassiveDelete.GetRowResultInfo(rowNumber);

            while (recipeRowResult != null)
            {
                Assert.AreEqual(0, recipeRowResult.UseCase, $"La recette est utilisée pour la ligne de résultat {rowNumber} alors que le filtre est sur 'unused'.");
                rowNumber++;
                recipeRowResult = recipeMassiveDelete.GetRowResultInfo(rowNumber);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Massivedelete_status_used()
        {
            //Home page
            var homePage = LogInAsAdmin();

            var recipePage = homePage.GoToMenus_Recipes();
            var recipeMassiveDelete = recipePage.OpenMassiveDeletePopup();

            recipeMassiveDelete.SelectStatus(MassiveDeleteStatus.Used.ToString());
            recipeMassiveDelete.ClickOnSearch();

            int rowNumber = 0;
            RecipeMassiveDeleteRowResult recipeRowResult = recipeMassiveDelete.GetRowResultInfo(rowNumber);

            while (recipeRowResult != null)
            {
                Assert.AreNotEqual(0, recipeRowResult.UseCase, $"La recette est utilisée pour la ligne de résultat {rowNumber} alors que le filtre est sur 'unused'.");
                rowNumber++;
                recipeRowResult = recipeMassiveDelete.GetRowResultInfo(rowNumber);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Massivedelete_organizebyRecipeName()
        {
            //Home page
            var homePage = LogInAsAdmin();

            var recipePage = homePage.GoToMenus_Recipes();
            var recipeMassiveDelete = recipePage.OpenMassiveDeletePopup();

            recipeMassiveDelete.ClickOnSearch();
            recipeMassiveDelete.SetPageSize("100");

            //Default search is by recipe name ascending
            bool isOK = recipeMassiveDelete.VerifySortByRecipeName(RecipeMassiveDeletePopup.SortType.Ascending);
            Assert.IsTrue(isOK, "The ordering of recipe names in ascending order is incorrect.");

            recipeMassiveDelete.ClickOnRecipeNameHeader();
            isOK = recipeMassiveDelete.VerifySortByRecipeName(RecipeMassiveDeletePopup.SortType.Descending);
            Assert.IsTrue(isOK, "The ordering of recipe names in descending order is incorrect.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Massivedelete_organizebysite()
        {
            //Home page
            var homePage = LogInAsAdmin();

            var recipePage = homePage.GoToMenus_Recipes();
            var recipeMassiveDelete = recipePage.OpenMassiveDeletePopup();

            recipeMassiveDelete.ClickOnSearch();
            recipeMassiveDelete.SetPageSize("100");

            recipeMassiveDelete.ClickOnSiteNameHeader();
            bool isOK = recipeMassiveDelete.VerifySortByRecipeSite(RecipeMassiveDeletePopup.SortType.Ascending);
            Assert.IsTrue(isOK, "The ordering of site names in ascending order is incorrect.");

            recipeMassiveDelete.ClickOnSiteNameHeader();
            isOK = recipeMassiveDelete.VerifySortByRecipeSite(RecipeMassiveDeletePopup.SortType.Descending);
            Assert.IsTrue(isOK, "The ordering of site names in descending order is incorrect.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Massivedelete_OrganizeByVariant()
        {
            //Home page
            var homePage = LogInAsAdmin();

            var recipePage = homePage.GoToMenus_Recipes();
            var recipeMassiveDelete = recipePage.OpenMassiveDeletePopup();

            recipeMassiveDelete.ClickOnSearch();
            recipeMassiveDelete.SetPageSize("100");

            recipeMassiveDelete.ClickOnVariantNameHeader();
            bool isOK = recipeMassiveDelete.VerifySortByVariant(RecipeMassiveDeletePopup.SortType.Ascending);
            Assert.IsTrue(isOK, "The ordering of variant names in ascending order is incorrect.");

            recipeMassiveDelete.ClickOnVariantNameHeader();
            isOK = recipeMassiveDelete.VerifySortByVariant(RecipeMassiveDeletePopup.SortType.Descending);
            Assert.IsTrue(isOK, "The ordering of variant names in descending order is incorrect.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_MassivePrint()
        {
            //Prepare
            string recipeType = "PAN";
            string typeOfGuest = "FLIGHT";
            string typeOfMeal = "FLIGHT";
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();


            recipesPage.Filter(RecipesPage.FilterType.RecipeType, recipeType);
            recipesPage.Filter(RecipesPage.FilterType.TypeOfGuest, typeOfGuest);
            recipesPage.Filter(RecipesPage.FilterType.TypeOfMeal, typeOfMeal);
            recipesPage.ClearDownloads();
            recipesPage.MassivePrint();

            //Assert
            Assert.IsTrue(recipesPage.IsFileExistInPrintList(), "La massive print ne fonctionne pas.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Massivedelete_link()
        {

            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var recipePage = homePage.GoToMenus_Recipes();
            var recipeMassiveDelete = recipePage.OpenMassiveDeletePopup();

            recipeMassiveDelete.SelectAllSites();
            recipeMassiveDelete.SelectAllVariant();
            recipeMassiveDelete.SelectAllRecipeType();
            recipeMassiveDelete.ClickOnSearch();
            var RecipeName = recipeMassiveDelete.GetFirstRecipeName();
            var RecipeGeneralInformation = recipeMassiveDelete.ClickOnRecipeLinkFromRow(1);

            // Assert
            Assert.AreEqual(RecipeName, RecipeGeneralInformation.GetRecipeName().Trim(), "La page detail recipe n'est pas visible.");
        }
        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Massivedelete_pagination()
        {
            HomePage homePage = LogInAsAdmin();

            RecipesPage recipePage = homePage.GoToMenus_Recipes();
            RecipeMassiveDeletePopup recipeMassiveDelete = recipePage.OpenMassiveDeletePopup();

            recipeMassiveDelete.ClickOnSearch();
            recipeMassiveDelete.ClickSelectAllButton();

            int tot = recipeMassiveDelete.CheckTotalSelectCount();

            recipeMassiveDelete.SetPageSize("16");
            Assert.IsTrue(recipeMassiveDelete.IsPageSizeEqualsTo("16"), "Pagination pour 16 ne fonctionne pas.");
            Assert.AreEqual(recipeMassiveDelete.GetTotalRowsForPagination(), tot >= 16 ? 16 : tot, "Pagination pour 16 ne fonctionne pas.");

            recipeMassiveDelete.SetPageSize("30");
            Assert.IsTrue(recipeMassiveDelete.IsPageSizeEqualsTo("30"), "Pagination pour 30 ne fonctionne pas.");
            Assert.AreEqual(recipeMassiveDelete.GetTotalRowsForPagination(), tot >= 30 ? 30 : tot, "Pagination pour 30 ne fonctionne pas.");

            recipeMassiveDelete.SetPageSize("50");
            Assert.IsTrue(recipeMassiveDelete.IsPageSizeEqualsTo("50"), "Pagination pour 50 ne fonctionne pas.");
            Assert.AreEqual(recipeMassiveDelete.GetTotalRowsForPagination(), tot >= 50 ? 50 : tot, "Pagination pour 50 ne fonctionne pas.");

            recipeMassiveDelete.SetPageSize("100");
            Assert.IsTrue(recipeMassiveDelete.IsPageSizeEqualsTo("100"), "Pagination pour 100 ne fonctionne pas.");
            Assert.AreEqual(recipeMassiveDelete.GetTotalRowsForPagination(), tot >= 100 ? 100 : tot, "Pagination pour 100 ne fonctionne pas.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Massivedelete_select()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);
            //string recipeNameSearch
            LogInAsAdmin();
            HomePage homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            RecipesPage recipesPage = homePage.GoToMenus_Recipes();
            for (int i = 0; i < 2; i++)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var newRecipeName = $"{recipeName}{i}";
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(newRecipeName, recipeType, nbPortions.ToString());
                recipeGeneralInfosPage.AddVariantWithSite(site, variant);
                if (!recipeGeneralInfosPage.IsFirstVariantIsVisible())
                {
                    recipeGeneralInfosPage.Validate();
                }
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
                recipeVariantPage.AddIngredient(ingredient);
                recipesPage = recipeVariantPage.BackToList();
            }
            RecipesPage recipePage = homePage.GoToMenus_Recipes();
            RecipeMassiveDeletePopup recipeMassiveDelete = recipePage.OpenMassiveDeletePopup();
            recipeMassiveDelete.SetRecipeName(recipeNameToday);
            recipeMassiveDelete.ClickOnSearch();
            string recipesNameDeleted = recipeMassiveDelete.GetFirstRecipeName();
            recipeMassiveDelete.CheckRowByRecipeName(recipesNameDeleted);
            recipeMassiveDelete.ClickDelete();
            recipeMassiveDelete.ClickToConfirmDelete();
            RecipeMassiveDeletePopup recipeMassiveDeletePopup = recipesPage.OpenMassiveDeletePopup();
            recipeMassiveDeletePopup.SetRecipeName(recipesNameDeleted);
            recipeMassiveDeletePopup.ClickOnSearch();
            string resultSearch = recipeMassiveDeletePopup.GetFirstRecipeName();
            Assert.AreNotEqual(recipesNameDeleted, resultSearch, "le Recipe sélectionné n'a pas été supprimé");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Massivedeletes()
        {
            //Prepare
            string variant = TestContext.Properties["MenuVariant"].ToString();
            string site = TestContext.Properties["SiteLP"].ToString();

            //string recipeNameSearch
            LogInAsAdmin();
            HomePage homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            RecipesPage recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            RecipeMassiveDeletePopup recipeMassiveDelete = recipesPage.OpenMassiveDeletePopup();
            recipeMassiveDelete.SelectSiteByName(site, false);
            recipeMassiveDelete.ClickOnSearch();
            var firstRow = recipeMassiveDelete.GetRowResultInfo(0);
            recipeMassiveDelete.CheckRowByRecipeName(firstRow.RecipeName);
            recipeMassiveDelete.ClickDelete();
            recipeMassiveDelete.ClickToConfirmDelete();
            recipesPage.OpenMassiveDeletePopup();
            recipeMassiveDelete.SelectSiteByName(site, false);
            recipeMassiveDelete.ClickOnSearch();
            var firstRowAfterDelete = recipeMassiveDelete.GetRowResultInfo(0);
            Assert.IsFalse(firstRowAfterDelete.RecipeName.Equals(firstRow.RecipeName) && firstRowAfterDelete.SiteName.Equals(firstRow.SiteName) && firstRowAfterDelete.VariantName.Equals(firstRow.VariantName), "le Recipe sélectionné n'a pas été supprimé");

            recipeMassiveDelete.CloseRecipeMassiveDeletePopup();
            recipesPage.OpenMassiveDeletePopup();
            recipeMassiveDelete.SelectVariantByName(variant, false);
            recipeMassiveDelete.ClickOnSearch();
            var firstRowSearchByVariant = recipeMassiveDelete.GetRowResultInfo(0);
            recipeMassiveDelete.CheckRowByRecipeName(firstRowSearchByVariant.RecipeName);
            recipeMassiveDelete.ClickDelete();
            recipeMassiveDelete.ClickToConfirmDelete();
            recipesPage.OpenMassiveDeletePopup();
            recipeMassiveDelete.SelectVariantByName(variant, false);
            recipeMassiveDelete.ClickOnSearch();
            var firstRowSearchByVariantAfterDelete = recipeMassiveDelete.GetRowResultInfo(0);
            Assert.IsFalse(firstRowSearchByVariantAfterDelete.RecipeName.Equals(firstRowSearchByVariant.RecipeName) && firstRowSearchByVariantAfterDelete.SiteName.Equals(firstRowSearchByVariant.SiteName) && firstRowSearchByVariantAfterDelete.VariantName.Equals(firstRowSearchByVariant.VariantName), "le Recipe sélectionné n'a pas été supprimé");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_CheckItemKeywordInRecipe()
        {

            //Prepare
            string variant = TestContext.Properties["Variant"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            string itemName = "Test" + DateTime.Now.ToString();
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();
            string keyword = "keyword2";
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
            var itemKeywordPage = itemGeneralInformationPage.ClickOnKeywordItem();
            itemKeywordPage.AddKeyword(keyword);
            var recipePage = homePage.GoToMenus_Recipes();
            var recipesCreateModalPage = recipePage.CreateNewRecipe();
            var recipeGeneralInformationPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());
            recipeGeneralInformationPage.AddVariantWithSite(site, variant);
            var firstVariant = recipeGeneralInformationPage.SelectFirstVariantFromList();
            firstVariant.AddIngredient(itemName);
            RecipeKeywordPage recipeKeyword = firstVariant.ClickOnKeyWordTab();
            var keywordRecipe = recipeKeyword.GetFirstKeyword();
            //Assert
            bool equalKeyword = keywordRecipe.Equals(keyword);
            Assert.IsTrue(equalKeyword, "La recette ne contient pas de keyword");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Massivedelete_UseCase()
        {
            string recipeName = "ABADEJO A LA PROVENZAL ALC";
            string siteName = "ALC";
            //Home page
            var homePage = LogInAsAdmin();

            var recipePage = homePage.GoToMenus_Recipes();
            var recipeMassiveDelete = recipePage.OpenMassiveDeletePopup();
            recipeMassiveDelete.SelectSiteByName(siteName, false);
            recipeMassiveDelete.SetRecipeName(recipeName);
            recipeMassiveDelete.ClickOnSearch();
            int tot = recipeMassiveDelete.GetTotalRowsForPagination();
            if (tot >= 1)
            {
                for (int i = 0; i < tot; i++)
                {
                    var recipeRowResult = recipeMassiveDelete.GetRowResultInfo(i);
                    RecipeGeneralInformationPage rGIPage1 = recipeMassiveDelete.ClickOnRecipeLinkFromRow(i + 1);
                    rGIPage1.ClickOnUseCaseTab();
                    Assert.AreEqual(recipeRowResult.UseCase, rGIPage1.CheckTotalRecipesNumber(), "Les UseCase affichés ne sont pas corrects");
                    rGIPage1.Go_To_New_Navigate();
                }
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Massivedelete_selectall()
        {
            //Home page
            var homePage = LogInAsAdmin();

            var recipePage = homePage.GoToMenus_Recipes();
            var recipeMassiveDelete = recipePage.OpenMassiveDeletePopup();
            recipeMassiveDelete.ClickOnSearch();

            var totalRecipeVariantsBeforeSelected = recipeMassiveDelete.CheckTotalSelectCount();
            recipeMassiveDelete.ClickSelectAllButton();
            var totalRecipeVariantsAfterSelected = recipeMassiveDelete.CheckTotalSelectCount();
            var isAllcheckedPage1 = recipeMassiveDelete.IsAllResultChecked();
            var firstPageIsAllSelected = recipeMassiveDelete.CheckResultIsAllSelected();
            recipeMassiveDelete.GoToPage("2");
            var secondPageIsAllSelected = recipeMassiveDelete.CheckResultIsAllSelected();
            var isAllcheckedPage2 = recipeMassiveDelete.IsAllResultChecked();
            Assert.IsTrue(totalRecipeVariantsBeforeSelected != totalRecipeVariantsAfterSelected, "L'affichage du nombre de Selected Recipe ne se met pas à jour");
            Assert.IsTrue(firstPageIsAllSelected && secondPageIsAllSelected && isAllcheckedPage1 && isAllcheckedPage2, "Le bouton 'Select All' ne fonctionne pas correctement ");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Massivedelete_OrganizeByUseCase()
        {
            string variantName = TestContext.Properties["MenuVariantACE1"].ToString();
            string siteName = TestContext.Properties["SiteToFlightBob"].ToString();
            List<int> listUseCase = new List<int>();
            //Home page
            var homePage = LogInAsAdmin();

            var recipePage = homePage.GoToMenus_Recipes();
            var recipeMassiveDelete = recipePage.OpenMassiveDeletePopup();
            recipeMassiveDelete.SelectSiteByName(siteName, false);
            recipeMassiveDelete.ClickOnSearch();
            recipeMassiveDelete.ClickOnSelectedAll();
            var nbrecipevariant = recipeMassiveDelete.NombreRecipeVariant();
            recipeMassiveDelete.ClickOnUnselectAll();
            var totalPages = (int)Math.Ceiling((double)(recipeMassiveDelete.ConvertStringToInt(nbrecipevariant)) / 8);
            recipeMassiveDelete.ClickUseCase();
            if (totalPages >= 1)
            {
                recipeMassiveDelete.PageSize("8");
                for (int i = 1; i <= totalPages; i++)
                {
                    recipeMassiveDelete.GetPageResults(i);
                    var tot = recipeMassiveDelete.GetTotalRowsForPagination();
                    for (int j = 1; j <= tot; j++)
                        listUseCase.Add(recipeMassiveDelete.ConvertStringToInt(recipeMassiveDelete.ValueUseCase(j)));
                }
                Assert.IsTrue(recipeMassiveDelete.IsSortingDescending(listUseCase), "Apres click  les résultats ne sont pas organiser par ordre croissant/décroissant");
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Massivedelete_OrganizeByWeight()
        {
            string variantName = TestContext.Properties["MenuVariantACE1"].ToString();
            string siteName = TestContext.Properties["SiteToFlightBob"].ToString();
            List<double> listUseCase = new List<double>();
            //Home page
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var recipePage = homePage.GoToMenus_Recipes();
            var recipeMassiveDelete = recipePage.OpenMassiveDeletePopup();
            recipeMassiveDelete.SelectSiteByName(siteName, false);
            recipeMassiveDelete.SelectVariantByName(variantName, false);
            recipeMassiveDelete.ClickOnSearch();
            recipeMassiveDelete.ClickOnSelectedAll();
            var nbrecipevariant = recipeMassiveDelete.NombreRecipeVariant();
            recipeMassiveDelete.ClickOnUnselectAll();
            var totalPages = (int)Math.Ceiling((double)(recipeMassiveDelete.ConvertStringToInt(nbrecipevariant)) / 8);
            recipeMassiveDelete.ClickWeight();
            recipeMassiveDelete.PageSize("100");
            List<double> itemsResults = recipeMassiveDelete.GetRecipeVariantWeight();
            var test = itemsResults.SequenceEqual(itemsResults.OrderBy(weight => weight));
            Assert.IsTrue(itemsResults.SequenceEqual(itemsResults.OrderBy(weight => weight)) || itemsResults.SequenceEqual(itemsResults.OrderByDescending(weight => weight)), "La recipe par weight n'est pas en ordre croissant/decroissant");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Massivedelete_status_unpurchasableItem()
        {
            //Arrange
            string variantName = TestContext.Properties["MenuVariantACE1"].ToString();
            string siteName = TestContext.Properties["SiteToFlightBob"].ToString();

            //Home page
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var recipePage = homePage.GoToMenus_Recipes();
            var recipeMassiveDelete = recipePage.OpenMassiveDeletePopup();
            recipeMassiveDelete.SelectSiteByName(siteName, false);
            recipeMassiveDelete.SelectVariantByName(variantName, false);
            recipeMassiveDelete.ClickOnSearch();
            recipeMassiveDelete.ClickOnSelectedAll();
            int rowCountWithoutFilter = int.Parse(recipeMassiveDelete.NombreRecipeVariant());

            recipeMassiveDelete.SelectStatus(MassiveDeleteStatus.WithUnpurchasableItems.ToString());
            recipeMassiveDelete.ClickOnSearch();

            recipeMassiveDelete.ClickOnUnselectAll();

            var rowCountWithFilter = int.Parse(recipeMassiveDelete.NombreRecipeVariant());


            Assert.AreNotEqual(rowCountWithoutFilter, rowCountWithFilter, $"Le filtre WithUnpurchasableItems ne fonctionne pas");

        }

        // __________________________________________ Utilitaire r

        private string CreateIngredientWithAllergen(HomePage homePage, string site, string allergens)
        {
            //Prepare
            string itemName = "itemIntoleanceRecipe";

            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();

            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string supplierRef = itemName + "_SupplierRef";
            string storageQty = 10.ToString();
            string qty = 10.ToString();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);

            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier, supplierRef);
                var itemIntolerancePage = itemGeneralInformationPage.ClickOnIntolerancePage();
                itemIntolerancePage.AddAllergen(allergens);
                itemIntolerancePage.SaveAllergen();
            }
            else
            {
                var itemDetailsPage = itemPage.ClickOnFirstItem();
                var itemIntolerancePage = itemDetailsPage.ClickOnIntolerancePage();
                itemIntolerancePage.AddAllergen(allergens);
            }

            return itemName;
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_ResetFilter()
        {
            //Prepare
            string mealType = "FLIGHT";
            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            var recipePage = homePage.GoToMenus_Recipes();
            recipePage.ResetFilter();
            var numberBeforeFilter = recipePage.CheckTotalNumber();
            recipePage.Filter(RecipesPage.FilterType.SearchRecipe, "RE_RECI_ResetFilter");
            recipePage.Filter(RecipesPage.FilterType.ShowAll, true);
            recipePage.Filter(RecipesPage.FilterType.TypeOfMeal, mealType);
            var numberAfterFilter = recipePage.CheckTotalNumber();
            recipePage.ResetFilter();
            var numberAfterResetFilter = recipePage.CheckTotalNumber();
            //Assert 
            Assert.AreNotEqual(numberBeforeFilter, numberAfterFilter, "Le filtre n'est pas fonctionnel");
            Assert.AreEqual(numberBeforeFilter, numberAfterResetFilter, "La réinitialisation du filtre n'est pas fonctionnelle");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Keyword_Item_Recipe()
        {
            string keyword = "keyword_for_test";
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            RecipesPage recipesPage = homePage.GoToMenus_Recipes();
            RecipeGeneralInformationPage generalInfo = recipesPage.SelectFirstRecipe();
            RecipesVariantPage variantPage = generalInfo.ClickOnFirstVariant();
            var ingredient = variantPage.GetFirstIngredient();
            RecipeKeywordPage recipeKeyword = variantPage.ClickOnKeyWordTab();

            try
            {
                var itemPage = homePage.GoToPurchasing_ItemPage();
                itemPage.Filter(ItemPage.FilterType.Search, ingredient);
                ItemGeneralInformationPage itemgeneralInfo = itemPage.ClickOnFirstItem();
                ItemKeywordPage itemKeywordPage = itemgeneralInfo.ClickOnKeywordItem();
                itemKeywordPage.Filter(ItemKeywordPage.FilterType.Keywords, keyword);
                recipesPage = homePage.GoToMenus_Recipes();
                recipesPage.SelectFirstRecipe();
                generalInfo.ClickOnFirstVariant();
                recipeKeyword = variantPage.ClickOnKeyWordTab();
                var isKeywordsExist = recipeKeyword.CheckKeyword(keyword);
                Assert.IsTrue(isKeywordsExist, "Le Keyword Recherché n'existe Pas");
            }
            finally
            {
                var itemPage = homePage.GoToPurchasing_ItemPage();
                itemPage.Filter(ItemPage.FilterType.Search, ingredient);
                ItemGeneralInformationPage itemgeneralInfo = itemPage.ClickOnFirstItem();
                ItemKeywordPage itemKeywordPage = itemgeneralInfo.ClickOnKeywordItem();
                itemKeywordPage.ClickOnuncheckKeyword();


            }
        }
        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_FilterTypeGuestMassivePrint()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string variant = TestContext.Properties["RecipeVariant"].ToString();
            //string guestType = variant.Substring(0, variant.IndexOf("-")).Trim();
            string guestType = "HOTELES";
            string Intolerance = "Apio/Celery";
            string RecipeType = "ENSALADA";
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            DeleteAllFileDownload();
            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.TypeOfGuest, guestType);
            recipesPage.Filter(RecipesPage.FilterType.Intolerance, Intolerance);
            recipesPage.Filter(RecipesPage.FilterType.RecipeType, RecipeType);

            recipesPage.ClearDownloads();
            recipesPage.MassivePrint();
            PrintReportPage printReport = recipesPage.ClickPrinterIcon();
            printReport.ClickRefreshUntilFinished();
            printReport.Purge(downloadsPath, "Recipe", "All_files_");
            string pdfFoundPath = printReport.PrintAllZipAllPdf(downloadsPath, "All_files_", false);
            PdfKeywordSearcher searcher = new PdfKeywordSearcher();
            bool pdfsFound = searcher.FindAllPdfsWithKeyword(downloadsPath, guestType);
            // Assert 
            Assert.IsTrue(pdfsFound, $"No PDF containing '{guestType}' was found in the directory: {downloadsPath}");


        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_Replace_Recipes()
        {
            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            string replaceRecipe = recipeNameToday + "-" + "Replace" + new Random().Next().ToString();
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["MenuVariantACE1"].ToString();
            string ingredient = "iemForDietaryTtpes";
            string site = TestContext.Properties["SiteACE"].ToString();
            int nbPortions = new Random().Next(1, 30);
            string storageUnit = "KG";
            string menuName = "Menu-" + DateUtils.Now.ToString("dd/MM /yyyy") + " - " + new Random().Next().ToString();
            string menuNameForReplace = "Menu-" + DateUtils.Now.ToString("dd/MM/yyyy") + " - " + "Replace" + new Random().Next().ToString();



            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            try
            {
                var recipesPage = homePage.GoToMenus_Recipes();
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());
                recipeGeneralInfosPage.AddVariantWithSite(site, variant);

                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
                recipeVariantPage.AddIngredient("a");

                Assert.IsTrue(recipeVariantPage.IsIngredientDisplayed(), "L'ingrédient n'a pas été ajouté à la recette.");

                recipesPage = homePage.GoToMenus_Recipes();
                recipesCreateModalPage = recipesPage.CreateNewRecipe();
                recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(replaceRecipe, recipeType, nbPortions.ToString());
                recipeGeneralInfosPage.AddVariantWithSite(site, variant);

                recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
                recipeVariantPage.AddIngredient("a");

                Assert.IsTrue(recipeVariantPage.IsIngredientDisplayed(), "L'ingrédient n'a pas été ajouté à la recette.");

                var menusPage = homePage.GoToMenus_Menus();
                var menusCreateModalPage = menusPage.MenuCreatePage();
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant);
                menuDayViewPage.AddRecipe(replaceRecipe);

                menusPage = homePage.GoToMenus_Menus();
                menusCreateModalPage = menusPage.MenuCreatePage();
                menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuNameForReplace, DateUtils.Now, DateUtils.Now.AddMonths(1), site, variant);
                menuDayViewPage.AddRecipe(recipeName);
                RecipesVariantPage RecipesVariantPage = menuDayViewPage.EditFirstRecipe();
                RecipeGeneralInformationPage generalInfo = RecipesVariantPage.ClickOnGeneralInformation();
                RecipeUseCaseTab useCase = generalInfo.ClickOnUseCaseTab();
                useCase.Filter(RecipeUseCaseTab.FilterType.SearchMenu, "Menu-");
                useCase.SelectAll();

                useCase.ReplaceByAnotherRecipe(replaceRecipe);
                useCase.BackToList();
                recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, replaceRecipe);
                recipesPage.SelectFirstRecipe();
                generalInfo.ClickOnUseCaseTab();
                useCase.Filter(RecipeUseCaseTab.FilterType.SearchMenu, "Menu-");

                Assert.IsTrue(useCase.CheckTotalNumber() > 0);
            }
            finally
            {

                var menusPage = homePage.GoToMenus_Menus();
                menusPage.MassiveDeleteMenus(menuName, site, variant);

                menusPage = homePage.GoToMenus_Menus();
                menusPage.MassiveDeleteMenus(menuNameForReplace, site, variant);

                var recipesPage = homePage.GoToMenus_Recipes();
                recipesPage.MassiveDeleteRecipes(recipeName, site, recipeType);

                recipesPage = homePage.GoToMenus_Recipes();
                recipesPage.MassiveDeleteRecipes(replaceRecipe, site, recipeType);

            }
        }
        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_ItemUnpurchasableColor()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string allergen = TestContext.Properties["RecipeAllergen"].ToString();
            string meal = TestContext.Properties["RecipeMeal"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string guestName = "testAllergen";
            string variant = guestName + " - " + meal;
            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);
            string ingredient = "Test" + new Random().Next().ToString();
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
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(ingredient, group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier, "2");
            itemGeneralInformationPage.ClickOnItem();
            itemGeneralInformationPage.UncheckIsPurchasable();
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());
            recipeGeneralInfosPage.AddSite(site);
            recipeGeneralInfosPage.SelectFirstVariantFromList();
            recipeGeneralInfosPage.AddIngredient(ingredient);
            bool isColorOrange = recipeGeneralInfosPage.IsIngredientTextColorOrange();

            // Assert if the color is correct
            Assert.IsTrue(isColorOrange, "The ingredient text color is not orange.");
        }
        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_DuplicateVariant_NetQuantity()
        {
            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["MenuVariantACE1"].ToString();
            string variant3 = TestContext.Properties["MenuVariantACE3"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();

            int nbPortions = new Random().Next(1, 30);

            var homePage = LogInAsAdmin();
            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            try
            {
                var recipesPage = homePage.GoToMenus_Recipes();
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());
                recipeGeneralInfosPage.AddVariantWithSite(site, variant);
                if (!recipeGeneralInfosPage.IsFirstVariantIsVisible())
                {
                    recipeGeneralInfosPage.Validate();
                }
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
                recipeVariantPage.AddIngredient(ingredient);
                recipeVariantPage.ClickOnFirstIngredient();
                var OriginalVariantNetQty = recipeVariantPage.GetNetQuantity(decimalSeparatorValue);
                var recipeDuplicateVariantPage = recipeVariantPage.DuplicateVariant();
                recipeDuplicateVariantPage.FillFields_DuplicateVariant(site, variant3);
                recipeDuplicateVariantPage.CloseDuplicate();
                recipeGeneralInfosPage.SelectVariant(variant3);
                recipeVariantPage.ClickOnFirstIngredient();
                var lVariantPriceDuplicated = recipeVariantPage.GetNetQuantity(decimalSeparatorValue);
                Assert.AreEqual(OriginalVariantNetQty, lVariantPriceDuplicated, "La quantité nette n'est pas la même après la duplication");
            }
            finally
            {
                var recipesPage = homePage.GoToMenus_Recipes();
                recipesPage.MassiveDeleteRecipes(recipeName, site, recipeType);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_SuppresionIngredient()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string meal = TestContext.Properties["RecipeMeal"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string guestName = "testAllergen";
            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);
            string ingredient = "Test" + new Random().Next().ToString();
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();

            var homePage = LogInAsAdmin();
            try
            {
                var itemPage = homePage.GoToPurchasing_ItemPage();
                itemPage.ResetFilter();
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(ingredient, group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier, "2");
                itemGeneralInformationPage.ClickOnItem();
                itemGeneralInformationPage.UncheckIsPurchasable();
                var recipesPage = homePage.GoToMenus_Recipes();
                recipesPage.ResetFilter();
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());
                recipeGeneralInfosPage.AddSite(site);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
                recipeVariantPage.AddIngredient(ingredient);

                // Récupération des valeurs initiales pour chaque élément 
                recipeVariantPage.ClickOnCompositionTab();
                string newComposition = recipeVariantPage.GetComposition();
                recipeVariantPage.ClickOnItemTab();
                bool isVerifiedSameItemComposition = recipeVariantPage.VerifiedSameItemComposition(newComposition);
                //Assert
                Assert.IsTrue(isVerifiedSameItemComposition, "Item dans composition n'existe pas.");

                //delete Ingredient 
                recipeVariantPage.ClickOnDeleteIngredient();
                recipeVariantPage.ClickOnCompositionTab();
                bool isVerifiedDeleteIngredient = recipeVariantPage.VerifiedDeleteIngredient(newComposition);
                //Assert
                Assert.IsTrue(isVerifiedDeleteIngredient, "En supprimant un ingrédient, il faut qu'il ne s'affiche plus sur l'onglet Composition.");

            }
            finally
            {
                var recipesPage = homePage.GoToMenus_Recipes();
                recipesPage.MassiveDeleteRecipes(recipeName, site, recipeType);

            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_ImportSalesPricePrix()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeTypeBis"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            bool newVersionPrint = true;

            //Prepare
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string salePriceModif = "10";
            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);

            //Arrange
            var homePage = LogInAsAdmin();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();
            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();

            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

            recipeGeneralInfosPage.AddVariantWithSite(site, variant);
            var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
            double initSalesPrice = recipeVariantPage.GetSalesPrice(decimalSeparatorValue);
            recipesPage = recipeVariantPage.BackToList();
            recipesPage.ResetFilter();
            recipesPage.ClearDownloads();
            DeleteAllFileDownload();

            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);
            recipesPage.ExportSalesPrice(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = recipesPage.GetExportExcelFile(taskFiles, true, false);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string sheetName, string fileName, string value
            OpenXmlExcel.WriteDataInColumn("Action", "Recipes Sales Price", filePath, "M", CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Sales Price", "Recipes Sales Price", filePath, salePriceModif, CellValues.SharedString);

            recipesPage.ImportSalesPrice(filePath);
            recipesPage.ImportFile();
            recipesPage.CloseAfterImport();

            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);

            recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
            recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();
            double SalesPriceAfterUpdate = recipeVariantPage.GetSalesPrice(decimalSeparatorValue);

            Assert.AreNotEqual(initSalesPrice, SalesPriceAfterUpdate, "La modification au niveau de Excel dans le champ Sales Price n'est pas effectué !");
            Assert.AreEqual(int.Parse(salePriceModif), SalesPriceAfterUpdate, "La modification au niveau de Excel dans le champ Sales Price n'est pas effectué !");

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_KeywordsDoublon()
        {
            //Prepare
            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["MenuVariantACE1"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            int nbPortions = new Random().Next(1, 30);
            string keyword2 = "keyword2";

            var homePage = LogInAsAdmin();

            // Act
            try
            {
                var recipesPage = homePage.GoToMenus_Recipes();
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());
                recipeGeneralInfosPage.AddVariantWithSite(site, variant);
                recipeGeneralInfosPage.Validate();
                recipeGeneralInfosPage.ClickOnSummary();
                homePage.GoToMenus_Recipes();
                recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);
                recipesPage.SelectFirstRecipe();
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
                recipeVariantPage.AddIngredient("a");
                string originalWindow = WebDriver.CurrentWindowHandle;
                var itemGeneralInformationPage = recipeVariantPage.EditFirstIngredient(false);
                var itemKeywordPage = itemGeneralInformationPage.ClickOnKeywordItem();
                itemKeywordPage.AddKeyword(keyword2);
                itemKeywordPage.CloseWindow(originalWindow);
                WebDriver.Navigate().Refresh();
                recipeGeneralInfosPage.SelectFirstVariantFromList();
                recipeVariantPage.AddIngredient(ingredient);
                recipeVariantPage.EditIngredient(ingredient);
                itemGeneralInformationPage.ClickOnKeywordItem();
                itemKeywordPage.AddKeyword(keyword2);
                homePage.GoToMenus_Recipes();
                recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);
                var RecipeGeneralInformationPage = recipesPage.SelectFirstRecipe();
                var RecipesVariantPage = RecipeGeneralInformationPage.SelectFirstVariantFromList();
                RecipesVariantPage.ClickOnKeyWordTab();
                bool VerifDuplication = RecipesVariantPage.CheckKeywordDuplicate(keyword2);

                //Assert
                Assert.IsTrue(VerifDuplication, "Le mot-clé est dupliqué");
            }
            finally
            {

                var recipesPage = homePage.GoToMenus_Recipes();
                recipesPage.MassiveDeleteRecipes(recipeName, site, recipeType);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_AffichageWorkshop()
        {
            string recipeName = "RecipeForAffichageWorkshop";
            //string recipeName = "AffichageWorkshop" + "-" + new Random().Next().ToString();
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["MenuVariantACE1"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            int nbPortions = new Random().Next(1, 30);
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();

            var homePage = LogInAsAdmin();

            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());
                recipeGeneralInfosPage.AddVariantWithSite(site, variant);
                recipeGeneralInfosPage.Validate();
            }
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, ingredient);

            var itemGeneralInformationPage = new ItemGeneralInformationPage(WebDriver, TestContext);

            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(ingredient, group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, "10", "KG", "5", supplier);
            }
            homePage.GoToMenus_Recipes();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);
            var recipeGeneralInfos = recipesPage.SelectFirstRecipe();
            var recipeVariantPage = recipeGeneralInfos.SelectFirstVariantFromList();
            recipeVariantPage.AddIngredient(ingredient);
            var verifyWorkshop = recipeVariantPage.VerifyWorkshop();
            Assert.IsTrue(verifyWorkshop, "La taille de la colonne workshop ne  s'adapte pas  au contenu.");




        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_RECI_AffichgeKeywords()
        {

            //Prepare
            string variant = TestContext.Properties["Variant"].ToString();
            string site = TestContext.Properties["SiteLP"].ToString();
            string itemName = "ItemKeyword" + DateTime.Now.ToString();
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();
            string keyword = "keyword2";
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);

            //Arrange
            var homePage = LogInAsAdmin();

            var itemPage = homePage.GoToPurchasing_ItemPage();
            try
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                var itemKeywordPage = itemGeneralInformationPage.ClickOnKeywordItem();
                itemKeywordPage.AddKeyword(keyword);
                var recipePage = homePage.GoToMenus_Recipes();
                var recipesCreateModalPage = recipePage.CreateNewRecipe();
                var recipeGeneralInformationPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());
                recipeGeneralInformationPage.AddVariantWithSite(site, variant);
                RecipesVariantPage recipesVariantPage = recipeGeneralInformationPage.SelectFirstVariantFromList();
                recipesVariantPage.AddIngredient(itemName);
                RecipeKeywordPage recipeKeywordPage = recipesVariantPage.ClickOnKeyWordTab();
                var keywordRecipe = recipeKeywordPage.GetFirstKeyword();

                bool equalKeyword = keywordRecipe.Equals(keyword);
                var isAligned = recipeKeywordPage.IsKeyWordsAlignedInTheTable();
                Assert.IsTrue(equalKeyword, "La recette ne contient pas de keyword");
                Assert.IsTrue(isAligned, "affichage n'est pas correcte (avec style) et les keyword ne sont pas alligné ");
            }
            finally
            {
                var recipePage = homePage.GoToMenus_Recipes();
                recipePage.MassiveDeleteRecipes(recipeName, site, recipeType);
                itemPage = homePage.GoToPurchasing_ItemPage();
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                ItemGeneralInformationPage itemgeneralInfo = itemPage.ClickOnFirstItem();
                ItemKeywordPage itemKeywordPage = itemgeneralInfo.ClickOnKeywordItem();
                itemKeywordPage.ClickOnuncheckKeyword();
                itemgeneralInfo = itemKeywordPage.ClickOnGeneralInfo();
                itemgeneralInfo.ClickOnDeleteItem();
                itemgeneralInfo.ConfirmDelete();
                itemPage = itemgeneralInfo.BackToList();
                MassiveDeleteModal massiveDeleteModal = itemPage.MenuMassiveDelete();
                massiveDeleteModal.Filter(MassiveDeleteModal.FilterType.SearchByName, itemName);
                massiveDeleteModal.ClickSearch();
                massiveDeleteModal.DeleteFirstService();


            }

        }

    }
}