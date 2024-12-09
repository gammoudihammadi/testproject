using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Accounting.SupplierInvoices;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Customer;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Datasheet;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Recipes;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using static Newrest.Winrest.FunctionalTests.PageObjects.Menus.Datasheet.DatasheetMassiveDeletePopup;
using Page = UglyToad.PdfPig.Content.Page;

namespace Newrest.Winrest.FunctionalTests.Menus
{
    [TestClass]
    public class DatasheetTests : TestBase
    {
        private const int _timeout = 600000;

        private readonly string datasheetNameToday = "Datasheet-" + DateUtils.Now.ToString("dd/MM/yyyy");
        private readonly string itemNameToday = "Item-" + DateUtils.Now.ToString("dd/MM/yyyy");
        private readonly string allergenNameToday = "Allergen-" + DateUtils.Now.ToString("dd/MM/yyyy");
        private readonly string NewallergenName = "NewAllergen-" + DateUtils.Now.ToString("dd/MM/yyyy");

        /// <summary>
        /// 
        /// Mise en place du paramétrage pour la configuration Winrest 4.0 
        /// 
        /// </summary>
        [TestMethod]
        [Priority(0)]
        [Timeout(_timeout)]
        public void RE_DATA_SetConfigWinrest4_0()
        {
            var site = TestContext.Properties["Site"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();
            ClearCache();

            // New version search
            homePage.SetNewVersionSearchValue(true);

            // New VersionItemDetails
            //homePage.SetVersionItemDetailValue(true);

            // Vérifier que c'est activé
            var itemPage = homePage.GoToPurchasing_ItemPage();

            try
            {
                var firstItem = itemPage.GetFirstItemName();
                itemPage.Filter(ItemPage.FilterType.Search, firstItem);
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
        public void RE_DATA_CreateRecipeForDatasheet()
        {
            //Prepare recipe
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();

            int nbPortions = new Random().Next(1, 30);

            //Prepare ingredient
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();

            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string supplierRef = ingredient + "_SupplierRef";
            string storageQty = 10.ToString();
            string qty = 10.ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, ingredient);

            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(ingredient.ToString(), group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier, supplierRef);
                itemPage = itemGeneralInformationPage.BackToList();
                itemPage.ResetFilter();
                itemPage.Filter(ItemPage.FilterType.Search, ingredient);
            }

            Assert.AreEqual(ingredient, itemPage.GetFirstItemName(), "L'ingredient n'est pas présent dans la liste.");

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

                recipeGeneralInfosPage.AddVariantWithSite(site, variant);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(ingredient);
                recipesPage = recipeVariantPage.BackToList();
                recipesPage.ResetFilter();
                recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);
            }
            else
            {
                var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
                var recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();
                if (!recipeVariantPage.IsIngredientDisplayed())
                {
                    recipeVariantPage.AddIngredient(ingredient);
                }
                recipesPage = recipeVariantPage.BackToList();
                recipesPage.ResetFilter();
                recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);
            }

            //Assert
            Assert.AreEqual(recipeName, recipesPage.GetFirstRecipeName(), "La recette n'a pas été ajoutée à la liste.");
        }

        [Priority(2)]
        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_CreateRecipeForDatasheetOtherSite()
        {
            //Prepare
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string site = TestContext.Properties["SiteLP"].ToString();

            string ingredient = "itemRecipeCopy";
            string recipeName = "OtherSiteDatasheetRecipe";
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
            var homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.Filter(ItemPage.FilterType.Search, ingredient);

            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(ingredient.ToString(), group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier, "ref");
            }

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

                recipeGeneralInfosPage.AddVariantWithSite(site, variant);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
                recipeVariantPage.AddIngredient(ingredient);
                Assert.IsTrue(recipeVariantPage.IsIngredientDisplayed(), "L'ingrédient n'a pas été ajouté à la variante de site " + site + ".");
                recipesPage = recipeVariantPage.BackToList();
                recipesPage.ResetFilter();
                recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);
            }

            //Assert
            Assert.AreEqual(recipeName, recipesPage.GetFirstRecipeName(), "La recette n'a pas été ajoutée à la liste.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_CreateNewDatasheet()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();

            string datasheetName = datasheetNameToday + "-" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);
            datasheetPage = datasheetDetailPage.BackToList();

            //Assert
            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
            // pas de csutomer, ni de recette liés donc on uncheck all
            datasheetPage.Filter(DatasheetPage.FilterType.Customers, null);
            datasheetPage.Filter(DatasheetPage.FilterType.Categories, null);
            datasheetPage.Filter(DatasheetPage.FilterType.Workshops, null);
            datasheetPage.Filter(DatasheetPage.FilterType.CookingModes, null);

            string firstDatasheet = datasheetPage.GetFirstDatasheetName();

            Assert.AreEqual(datasheetName, firstDatasheet, "La datasheet n'a pas été créée.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_CreateNewDatasheetWithoutSite()
        {
            //Prepare
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();

            string datasheetName = datasheetNameToday + "-" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            datasheetCreateModalPage.FillField_CreateNewDatasheetWithoutSite(datasheetName, guestType);

            bool isErrorMessageSiteDisplayed = datasheetCreateModalPage.IsErrorMessageSiteDisplayed();
            //Assert
            Assert.IsTrue(isErrorMessageSiteDisplayed, "La datasheet a été créée alors qu'aucun site n'a été spécifiée.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_CreateNewDatasheetAlreadyExisting()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();

            string datasheetName = datasheetNameToday + "-" + new Random().Next().ToString();

            DatasheetCreateModalPage datasheetCreateModalPage;
            DatasheetDetailsPage datasheetDetailPage;

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetPage.ResetFilter();

            if (datasheetPage.CheckTotalNumber() == 0)
            {
                datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
                datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);
                datasheetPage = datasheetDetailPage.BackToList();
                datasheetPage.ResetFilter();
            }
            else
            {
                datasheetName = datasheetPage.GetFirstDatasheetName();
            }

            datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            bool isErrorMessageNameDisplayed = datasheetCreateModalPage.IsErrorMessageNameDisplayed();
            //Assert
            Assert.IsTrue(isErrorMessageNameDisplayed, "La datasheet a été créée alors qu'une autre avec le même nom existe déjà.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_UpdateGeneralInformation()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();

            string datasheetName = datasheetNameToday + "-" + new Random().Next().ToString();

            string newName = datasheetName + "Bis";
            string newCommercialName = "test";
            string newCommercialName2 = "test2";
            string newCustomerCode = new Random().Next(1, 100).ToString();
            string newGuestType = TestContext.Properties["Guest"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);
            var datasheetGeneralInformationPage = datasheetDetailPage.ClickOnGeneralInformation();

            // Récupération des valeurs initiales
            var initName = datasheetGeneralInformationPage.GetName();
            var initCommercialName = datasheetGeneralInformationPage.GetCommercialName();
            var initCommercialName2 = datasheetGeneralInformationPage.GetCommercialName2();
            var initCustomerCode = datasheetGeneralInformationPage.GetCustomerCode();
            var initGuestType = datasheetGeneralInformationPage.GetGuestType();

            Assert.AreNotEqual(newName, initName, "Le nouveau nom de la datasheet est identique à l'ancien.");
            Assert.AreNotEqual(newCommercialName, initCommercialName, "Le nouveau commercialName de la datasheet est identique à l'ancien.");
            Assert.AreNotEqual(newCommercialName2, initCommercialName2, "Le nouveau commercialName2 de la datasheet est identique à l'ancien.");
            Assert.AreNotEqual(newCustomerCode, initCustomerCode, "Le nouveau customerCode de la datasheet est identique à l'ancien.");
            Assert.AreNotEqual(newGuestType, initGuestType, "Le nouveau guestType de la datasheet est identique à l'ancien.");

            // Modification des valeurs
            datasheetGeneralInformationPage.SetName(newName);
            datasheetGeneralInformationPage.SetCommercialName(newCommercialName);
            datasheetGeneralInformationPage.SetCommercialName2(newCommercialName2);
            datasheetGeneralInformationPage.SetCustomerCode(newCustomerCode);
            datasheetGeneralInformationPage.SetGuestType(newGuestType);

            datasheetDetailPage = datasheetGeneralInformationPage.ClickOnDetailsTab();
            datasheetGeneralInformationPage = datasheetDetailPage.ClickOnGeneralInformation();

            // Assert
            Assert.AreEqual(newName, datasheetGeneralInformationPage.GetName(), "Le nom de la datasheet n'a pas été modifié.");
            Assert.AreEqual(newCommercialName, datasheetGeneralInformationPage.GetCommercialName(), "Le commercialName de la datasheet n'a pas été modifié.");
            Assert.AreEqual(newCommercialName2, datasheetGeneralInformationPage.GetCommercialName2(), "Le commercialName2 de la datasheet n'a pas été modifié.");
            Assert.AreEqual(newCustomerCode, datasheetGeneralInformationPage.GetCustomerCode(), "Le customerCode de la datasheet n'a pas été modifié.");
            Assert.AreEqual(newGuestType, datasheetGeneralInformationPage.GetGuestType(), "Le guestType de la datasheet n'a pas été modifié.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_AddRecipeToDatasheet()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();

            string datasheetName = datasheetNameToday + "-" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            datasheetDetailPage.AddRecipe(recipeName);
            datasheetDetailPage.WaitLoading();
            WebDriver.Navigate().Refresh();

            //Assert
            datasheetDetailPage.WaitPageLoading();
            bool isRecipeAdded = datasheetDetailPage.IsRecipeAdded();
            Assert.IsTrue(isRecipeAdded, "La recette n'a pas été rajoutée à la datasheet.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_AddInactiveRecipeToDatasheet()
        {
            //Prepare datasheet
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string datasheetName = datasheetNameToday + "-" + new Random().Next().ToString();

            // Prépare recipe
            Random rnd = new Random();
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string recipeName = "InactiveRecipe-" + DateUtils.Now.ToString("dd/MM/yyyy") + rnd.Next(1, 100);
            int nbPortions = rnd.Next(1, 30);

            //Arrange
            var homePage = LogInAsAdmin();

            // Créer une recette inactive
            var recipesPage = homePage.GoToMenus_Recipes();
            try
            {
                recipesPage.ResetFilter();

                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString(), false);

                recipeGeneralInfosPage.AddVariantWithSite(site, variant);
                recipeGeneralInfosPage.ClickOnEditInformation();
                var isActivated = recipeGeneralInfosPage.IsRecipeActivated();
                Assert.IsFalse(isActivated, " la recette est activé");
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(ingredient);

                // Créer un datasheet
                var datasheetPage = homePage.GoToMenus_Datasheet();
                var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
                var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);
                var isAdded = datasheetDetailPage.AddRecipeError(recipeName);
                //Assert
                Assert.IsFalse(isAdded, "La recette a été rajoutée à la datasheet malgré le fait qu'elle soit inactive.");
            }
            finally
            {
                var datasheetPage = homePage.GoToMenus_Datasheet();
                var dsMassiveDelete = datasheetPage.OpenMassiveDeletePopup();
                dsMassiveDelete.SetDatasheetName(datasheetName);
                dsMassiveDelete.ClickOnSearch();
                dsMassiveDelete.SelectAllSites();
                dsMassiveDelete.ClickDeleteButton();

                recipesPage = homePage.GoToMenus_Recipes();
                recipesPage.MassiveDeleteRecipes(recipeName, site, recipeType);
            }

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_AddRecipeFromOtherSiteToDatasheet()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = "OtherSiteDatasheetRecipe";

            string datasheetName = datasheetNameToday + "-" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            var isAdded = datasheetDetailPage.AddRecipeError(recipeName);

            //Assert
            Assert.IsFalse(isAdded, "La recette a été rajoutée à la datasheet malgré le fait qu'elle ne soit pas associée au même site.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_AddRecipeFromOtherDatasheet()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();

            string datasheetName = datasheetNameToday + "-" + new Random().Next().ToString();
            string datasheetNameBis = datasheetName + "Bis";

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            datasheetDetailPage.AddRecipe(recipeName);
            datasheetDetailPage.WaitLoading();
            WebDriver.Navigate().Refresh();

            bool isRecipeAdded1 = datasheetDetailPage.IsRecipeAdded();
            Assert.IsTrue(isRecipeAdded1, "La recette n'a pas été rajoutée à la datasheet.");

            datasheetPage = datasheetDetailPage.BackToList();
            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
            // pas de customer lié donc on uncheck all

            /*datasheetPage.Filter(DatasheetPage.FilterType.Customers, null);
            datasheetPage.Filter(DatasheetPage.FilterType.Categories, null);*/

            string firstDtasheetName = datasheetPage.GetFirstDatasheetName();
            Assert.AreEqual(firstDtasheetName, datasheetName, "La datasheet n'a pas été créée.");

            datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetNameBis, guestType, site);

            bool isRecipeAdded2 = datasheetDetailPage.IsRecipeAdded();
            Assert.IsFalse(isRecipeAdded2, "La datasheet possède déjà une recette.");

            datasheetDetailPage.AddRecipeFromDatasheet(datasheetName);
            bool isRecipeAdded = datasheetDetailPage.IsRecipeAdded();
            string firstRecipeName = datasheetDetailPage.GetFirstRecipeName();
            Assert.IsTrue(isRecipeAdded, "La recette de la datasheet " + datasheetName + " n'a pas été rajoutée à la datasheet " + datasheetNameBis);
            Assert.AreEqual(firstRecipeName, recipeName, "La recette n'a pas été ajoutée correctement.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_AddRecipeFromOtherDatasheetOtherSite()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string siteBis = TestContext.Properties["SiteLP"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();

            Random rnd = new Random();
            string datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();
            string datasheetNameBis = datasheetNameToday + "-" + rnd.Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            datasheetDetailPage.AddRecipe(recipeName);
            datasheetDetailPage.WaitLoading();
            WebDriver.Navigate().Refresh();

            bool isRecipeAdded = datasheetDetailPage.IsRecipeAdded();
            Assert.IsTrue(isRecipeAdded, "La recette n'a pas été rajoutée à la datasheet.");

            datasheetPage = datasheetDetailPage.BackToList();
            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
            // pas de csutomer lié donc on uncheck all
            datasheetPage.Filter(DatasheetPage.FilterType.Customers, null);
            datasheetPage.Filter(DatasheetPage.FilterType.Categories, null);

            string firstDatasheetName = datasheetPage.GetFirstDatasheetName();
            Assert.AreEqual(firstDatasheetName, datasheetName, "La datasheet n'a pas été créée.");

            datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetNameBis, guestType, siteBis);

            bool isRecipeAdded2 = datasheetDetailPage.IsRecipeAdded();
            Assert.IsFalse(isRecipeAdded2, "La datasheet possède déjà une recette.");

            bool isAdded = datasheetDetailPage.AddRecipeFromDatasheetError(datasheetName);
            Assert.IsFalse(isAdded, "La datasheet a été importée malgré le fait qu'elle ne soit pas associée au même site que la datasheet cible.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_CreateNewRecipeInDatasheet()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();

            Random rnd = new Random();
            string datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();
            string recipeName = "RecipeTest-" + rnd.Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            bool isRecipeAdded_1 = datasheetDetailPage.IsRecipeAdded();
            Assert.IsFalse(isRecipeAdded_1, "La datasheet possède déjà une recette.");

            var datasheetCreateNewRecipePage = datasheetDetailPage.CreateNewRecipe();
            datasheetDetailPage = datasheetCreateNewRecipePage.FillFields_AddNewRecipeToDatasheet(recipeName, ingredient);

            bool isRecipeAdded_2 = datasheetDetailPage.IsRecipeAdded();
            Assert.IsTrue(isRecipeAdded_2, "La recette n'a pas été rajoutée à la datasheet.");

            string firstRecipeName = datasheetDetailPage.GetFirstRecipeName();
            Assert.AreEqual(firstRecipeName, recipeName, "La recette n'a pas été ajoutée correctement.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_PrintDatasheetDetails_NewVersion()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();

            string datasheetName = datasheetNameToday + "-" + new Random().Next().ToString();

            bool versionPrint = true;

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();

            datasheetPage.ClearDownloads();

            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            datasheetDetailPage.AddRecipe(recipeName);
            datasheetDetailPage.WaitLoading();
            WebDriver.Navigate().Refresh();
            bool isRecipeAdded = datasheetDetailPage.IsRecipeAdded();
            Assert.IsTrue(isRecipeAdded, "La recette n'a pas été rajoutée à la datasheet.");

            var reportPage = datasheetDetailPage.Print(versionPrint);
            var isGenerated = reportPage.IsReportGenerated();
            reportPage.Close();

            //Assert
            Assert.IsTrue(isGenerated, "La datasheet n'a pas été imprimée.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_PrintDatasheetList_NewVersion()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();

            string datasheetName = datasheetNameToday + "-" + new Random().Next().ToString();

            bool versionPrint = true;

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();

            datasheetPage.ClearDownloads();

            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            datasheetDetailPage.AddRecipe(recipeName);
            WebDriver.Navigate().Refresh();
            bool isRecipeAdded = datasheetDetailPage.IsRecipeAdded();
            Assert.IsTrue(isRecipeAdded, "La recette n'a pas été rajoutée à la datasheet.");

            datasheetPage = datasheetDetailPage.BackToList();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);

            var reportPage = datasheetPage.PrintDatasheet(versionPrint);
            var isGenerated = reportPage.IsReportGenerated();
            reportPage.Close();

            //Assert
            Assert.IsTrue(isGenerated, "La datasheet n'a pas été imprimée depuis la page Index.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_ValidateIntolerancesForDatasheet()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();

            string datasheetName = datasheetNameToday + "-" + new Random().Next().ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            datasheetDetailPage.AddRecipe(recipeName);
            datasheetDetailPage.WaitLoading();
            WebDriver.Navigate().Refresh();

            Assert.IsTrue(datasheetDetailPage.IsRecipeAdded(), "La recette n'a pas été rajoutée à la datasheet.");

            datasheetDetailPage.ClickOnIntoleranceTab();
            datasheetDetailPage.ValidateAllergen();
            Assert.IsTrue(datasheetDetailPage.IsAllergenValidated(), "Les intolerances associées à la datasheet n'ont pas été validées.");

            datasheetPage = datasheetDetailPage.BackToList();
            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
            // pas de customer lié donc on uncheck all
            datasheetPage.Filter(DatasheetPage.FilterType.Customers, null);
            datasheetPage.Filter(DatasheetPage.FilterType.Categories, null);

            Assert.IsTrue(datasheetPage.CheckAllergenValidation(true), "L'icone de validation des allergènes de la datasheet n'est pas à jour.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_DevalidateIntolerancesForDatasheet()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();
            string datasheetName = datasheetNameToday + "-" + new Random().Next().ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();
            try
            {
                //Act
                var datasheetPage = homePage.GoToMenus_Datasheet();
                var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
                var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);
                datasheetDetailPage.AddRecipe(recipeName);
                datasheetDetailPage.WaitLoading();
                WebDriver.Navigate().Refresh();
                //Assert
                var isRecipeAddedOk = datasheetDetailPage.IsRecipeAdded();
                Assert.IsTrue(isRecipeAddedOk, "La recette n'a pas été rajoutée à la datasheet.");

                datasheetDetailPage.ClickOnIntoleranceTab();
                datasheetDetailPage.ValidateAllergen();
                //Assert
                var IsAllergenValidated = datasheetDetailPage.IsAllergenValidated();
                Assert.IsTrue(IsAllergenValidated, "Les intolerances associées à la datasheet n'ont pas été validées.");

                datasheetPage = datasheetDetailPage.BackToList();
                datasheetPage.ResetFilter();
                datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
                // pas de csutomer lié donc on uncheck all
                datasheetPage.Filter(DatasheetPage.FilterType.Customers, null);
                datasheetPage.Filter(DatasheetPage.FilterType.Categories, null);
                //Assert
                var isCheckAllergenValidationOk = datasheetPage.CheckAllergenValidation(true);
                Assert.IsTrue(isCheckAllergenValidationOk, "L'icone de validation des allergènes de la datasheet n'est pas à jour.");

                datasheetDetailPage = datasheetPage.SelectFirstDatasheet();
                datasheetDetailPage.ClickOnIntoleranceTab();
                datasheetDetailPage.DevalidateAllergen();
                //Assert
                var isIsAllergenValidatedOk = datasheetDetailPage.IsAllergenValidated();
                Assert.IsFalse(isIsAllergenValidatedOk, "Les intolerances associées à la datasheet n'ont pas été invalidées.");

                datasheetPage = datasheetDetailPage.BackToList();
                datasheetPage.ResetFilter();
                datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
                // pas de csutomer lié donc on uncheck all
                datasheetPage.Filter(DatasheetPage.FilterType.Customers, null);
                datasheetPage.Filter(DatasheetPage.FilterType.Categories, null);
                //Assert
                var isCheckAllergenValidation = datasheetPage.CheckAllergenValidation(false);
                Assert.IsFalse(isCheckAllergenValidation, "L'icone d'invalidation des allergènes de la datasheet n'est pas à jour.");
            }
            finally
            {
                var datasheetPage = homePage.GoToMenus_Datasheet();
                datasheetPage.ResetFilter();
                datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
                DatasheetDetailsPage details = datasheetPage.SelectFirstDatasheet();
                DatasheetGeneralInformationPage generalInfo = details.ClickOnGeneralInformation();
                generalInfo.SetActive(false);
                generalInfo.DeleteDatasheet();
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_AddPictureToDatasheet()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();

            string datasheetName = datasheetNameToday + "-" + new Random().Next().ToString();

            var path = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase);
            path = path.Substring(6);
            var imagePath = path + "\\PageObjects\\Menus\\Datasheet\\test.jpg";

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            datasheetDetailPage.AddRecipe(recipeName);
            datasheetDetailPage.WaitLoading();
            WebDriver.Navigate().Refresh();
            Assert.IsTrue(datasheetDetailPage.IsRecipeAdded(), "La recette n'a pas été rajoutée à la datasheet.");

            datasheetDetailPage.ClickOnPictureTab();
            datasheetDetailPage.AddPicture(imagePath);
            Assert.IsTrue(datasheetDetailPage.IsPictureAdded(), "L'image n'a pas été ajoutée à la datasheet.");

            datasheetPage = datasheetDetailPage.BackToList();
            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
            // pas de csutomer lié donc on uncheck all
            datasheetPage.Filter(DatasheetPage.FilterType.Customers, null);
            datasheetPage.Filter(DatasheetPage.FilterType.Categories, null);
            Assert.IsTrue(datasheetPage.IsPictureLinked(), "L'icone d'ajout d'une image à la datasheet n'est pas à jour.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_DeletePictureToDatasheet()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();
            string datasheetName = datasheetNameToday + "-" + new Random().Next().ToString();
            var path = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase);
            path = path.Substring(6);
            var imagePath = path + "\\PageObjects\\Menus\\Datasheet\\test.jpg";

            //Arrange
            HomePage homePage = LogInAsAdmin();
            try
            {
                //Act
                var datasheetPage = homePage.GoToMenus_Datasheet();
                var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
                var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);
                datasheetDetailPage.AddRecipe(recipeName);
                datasheetDetailPage.WaitLoading();
                WebDriver.Navigate().Refresh();
                //Assert
                var IsRecipeAddedOk = datasheetDetailPage.IsRecipeAdded();
                Assert.IsTrue(IsRecipeAddedOk, "La recette n'a pas été rajoutée à la datasheet.");

                datasheetDetailPage.ClickOnPictureTab();
                datasheetDetailPage.AddPicture(imagePath);
                //Assert
                var IsPictureAddedOk = datasheetDetailPage.IsPictureAdded();
                Assert.IsTrue(IsPictureAddedOk, "L'image n'a pas été ajoutée à la datasheet.");

                datasheetPage = datasheetDetailPage.BackToList();
                datasheetPage.ResetFilter();
                datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
                // pas de csutomer lié donc on uncheck all
                datasheetPage.Filter(DatasheetPage.FilterType.Customers, null);
                datasheetPage.Filter(DatasheetPage.FilterType.Categories, null);
                //Assert
                var IsPictureLinkedOk = datasheetPage.IsPictureLinked();
                Assert.IsTrue(IsPictureLinkedOk, "L'icone d'ajout d'une image à la datasheet n'est pas à jour.");

                datasheetDetailPage = datasheetPage.SelectFirstDatasheet();
                datasheetDetailPage.ClickOnPictureTab();
                datasheetDetailPage.DeletePicture();
                //Assert
                Assert.IsFalse(datasheetDetailPage.IsPictureAdded(), "L'image n'a pas été supprimée de la datasheet.");

                datasheetPage = datasheetDetailPage.BackToList();
                datasheetPage.ResetFilter();
                datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
                // pas de csutomer lié donc on uncheck all
                datasheetPage.Filter(DatasheetPage.FilterType.Customers, null);
                datasheetPage.Filter(DatasheetPage.FilterType.Categories, null);
                //Assert
                var IsPictureLinked = datasheetPage.IsPictureLinked();
                Assert.IsFalse(IsPictureLinked, "L'icone de suppression d'une image de la datasheet n'est pas à jour.");
            }
            finally
            {
                var datasheetPage = homePage.GoToMenus_Datasheet();
                datasheetPage.ResetFilter();
                datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
                DatasheetDetailsPage details = datasheetPage.SelectFirstDatasheet();
                DatasheetGeneralInformationPage generalInfo = details.ClickOnGeneralInformation();
                generalInfo.SetActive(false);
                generalInfo.DeleteDatasheet();

            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_UnfoldDetails()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();

            string datasheetName = datasheetNameToday + "-" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            datasheetDetailPage.AddRecipe(recipeName);
            datasheetDetailPage.WaitLoading();
            WebDriver.Navigate().Refresh();
            bool isReceipeAdded = datasheetDetailPage.IsRecipeAdded();
            Assert.IsTrue(isReceipeAdded, "La recette n'a pas été rajoutée à la datasheet.");

            datasheetDetailPage.ShowDetails();
            bool isReceipeDetailsVisible = datasheetDetailPage.IsRecipeDetailsVisible();
            Assert.IsTrue(isReceipeDetailsVisible, "Le détail des recettes n'est pas visible.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_FoldDetails()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();

            string datasheetName = datasheetNameToday + "-" + new Random().Next().ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            DatasheetPage datasheetPage = homePage.GoToMenus_Datasheet();
            DatasheetCreateModalPage datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            datasheetDetailPage.AddRecipe(recipeName);
            datasheetDetailPage.WaitLoading();
            WebDriver.Navigate().Refresh();
            bool isRecipeAdded = datasheetDetailPage.IsRecipeAdded();
            Assert.IsTrue(isRecipeAdded, "La recette n'a pas été rajoutée à la datasheet.");

            datasheetDetailPage.ShowDetails();
            bool isVisible = datasheetDetailPage.IsRecipeDetailsVisible();
            Assert.IsTrue(isVisible, "Le détail des recettes n'est pas visible.");

            datasheetDetailPage.HideDetails();
            isVisible = datasheetDetailPage.IsRecipeDetailsVisible();
            Assert.IsFalse(isVisible, "Le détail des recettes n'est pas masqué.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_AddSubrecipe()
        {
            //Prepare
            var site = TestContext.Properties["Site"].ToString();
            var guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            var recipeName = TestContext.Properties["DatasheetRecipe"].ToString();
            var subRecipeIngredient = TestContext.Properties["RecipeIngredientBis"].ToString();
            var rnd = new Random();
            var datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();
            var subrecipeName = "RecipeTest-" + rnd.Next().ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);
            Assert.IsFalse(datasheetDetailPage.IsRecipeAdded(), "La datasheet possède déjà une recette.");

            datasheetDetailPage.AddRecipe(recipeName);
            Assert.IsTrue(datasheetDetailPage.IsRecipeAdded(), "La recette n'a pas été rajoutée à la datasheet.");

            var datasheetCreateNewRecipePage = datasheetDetailPage.AddSubrecipeForFirstRecipe();
            datasheetDetailPage = datasheetCreateNewRecipePage.FillFields_AddNewRecipeToDatasheet(subrecipeName, subRecipeIngredient);
            if (!datasheetDetailPage.isPopupClosed())
            {
                datasheetDetailPage.CloseDetailModal();
            }
            datasheetDetailPage.ShowDetails();
            Assert.IsTrue(datasheetDetailPage.IsRecipeDetailsVisible(), "Le détail des recettes n'est pas visible.");
            Assert.IsTrue(datasheetDetailPage.IsNewSubRecipeAdded(subrecipeName), "La sous-recette n'a pas été ajoutée à la recette.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_DeleteRecipe()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();

            Random rnd = new Random();
            string datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            datasheetDetailPage.AddRecipe(recipeName);
            datasheetDetailPage.WaitLoading();
            WebDriver.Navigate().Refresh();
            Assert.IsTrue(datasheetDetailPage.IsRecipeAdded(), "La recette n'a pas été rajoutée à la datasheet.");

            datasheetDetailPage.DeleteRecipe();
            datasheetDetailPage.WaitLoading();
            Assert.IsFalse(datasheetDetailPage.IsRecipeAdded(), "La recette n'a pas été supprimée de la datasheet.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_EditRecipe()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();

            Random rnd = new Random();
            string datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();

            string newValue = "10";

            bool isNewWindow = false;

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            DatasheetDetailsPage datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            datasheetDetailPage.AddRecipe(recipeName);
            Assert.IsTrue(datasheetDetailPage.IsRecipeAdded(), "La recette n'a pas été rajoutée à la datasheet.");

            datasheetDetailPage.ClickOnFirstRecipe();
            string initNetWeight = datasheetDetailPage.GetNetWeight(decimalSeparatorValue).ToString();

            // unselect first line
            datasheetDetailPage.ClickOnUseCaseTab();
            datasheetDetailPage.ClickOnDetailsTab();

            var recipeVariantPage = datasheetDetailPage.EditRecipe();
            isNewWindow = true;
            try
            {
                recipeVariantPage.ClickOnFirstIngredientDatasheet();
                recipeVariantPage.SetNetWeight(newValue);
                var newNetWeight = recipeVariantPage.GetTotalPortion(decimalSeparatorValue).ToString();
                datasheetDetailPage = recipeVariantPage.CloseWindow();
                isNewWindow = false;

                // unselect first line
                datasheetDetailPage.ClickOnUseCaseTab();
                datasheetDetailPage.ClickOnDetailsTab();

                datasheetDetailPage.ClickOnFirstRecipe();

                Assert.AreNotEqual(initNetWeight, datasheetDetailPage.GetNetWeight(decimalSeparatorValue).ToString(), "La modification de la recette n'a pas été prise en compte.");
                Assert.AreEqual(newNetWeight, datasheetDetailPage.GetNetWeight(decimalSeparatorValue).ToString(), "La nouvelle valeur du NetWeight de la recette n'est pas correct.");
            }
            finally
            {
                if (isNewWindow)
                {
                    recipeVariantPage.ClickOnFirstIngredient();
                    recipeVariantPage.SetNetWeight(initNetWeight);
                    recipeVariantPage.CloseWindow();
                }
                else
                {
                    WebDriver.Navigate().Refresh();
                    recipeVariantPage = datasheetDetailPage.EditRecipe();
                    recipeVariantPage.ClickOnFirstIngredientDatasheet();
                    recipeVariantPage.SetNetWeight(newValue);
                    recipeVariantPage.CloseWindow();
                }
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_UpdateRecipeMethod_Standard()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();

            Random rnd = new Random();
            string datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();

            string newMethod = "Std";
            string newMethod2 = "%";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            datasheetDetailPage.AddRecipe(recipeName);
            datasheetDetailPage.WaitLoading();
            WebDriver.Navigate().Refresh();
            Assert.IsTrue(datasheetDetailPage.IsRecipeAdded(), "La recette n'a pas été rajoutée à la datasheet.");

            datasheetDetailPage.ClickOnFirstRecipe();
            string initMethod = datasheetDetailPage.GetMethod();

            if (newMethod.Equals(initMethod))
            {
                datasheetDetailPage.SetMethod(newMethod2);
                datasheetDetailPage.ClickOnFirstRecipe();
                initMethod = datasheetDetailPage.GetMethod();
            }

            datasheetDetailPage.SetMethod(newMethod);
            WebDriver.Navigate().Refresh();

            datasheetDetailPage.ClickOnFirstRecipe();
            string finalMethod = datasheetDetailPage.GetMethod();
            string finalCoef = datasheetDetailPage.GetCoef();

            Assert.AreNotEqual(initMethod, finalMethod, "La valeur de la méthode de la recette n'a pas été modifiée.");
            Assert.AreEqual("1", finalCoef, "La valeur du coef de la recette n'a pas été modifié.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_UpdateRecipeMethod_Pourcentage()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();

            Random rnd = new Random();
            string datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();

            string newMethod = "%";
            string newMethod2 = "Std";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            datasheetDetailPage.AddRecipe(recipeName);
            datasheetDetailPage.WaitLoading();
            WebDriver.Navigate().Refresh();
            Assert.IsTrue(datasheetDetailPage.IsRecipeAdded(), "La recette n'a pas été rajoutée à la datasheet.");

            datasheetDetailPage.ClickOnFirstRecipe();
            string initMethod = datasheetDetailPage.GetMethod();

            if (newMethod.Equals(initMethod))
            {
                datasheetDetailPage.SetMethod(newMethod2);
                initMethod = datasheetDetailPage.GetMethod();
            }

            datasheetDetailPage.SetMethod(newMethod);
            WebDriver.Navigate().Refresh();

            datasheetDetailPage.ClickOnFirstRecipe();
            string finalMethod = datasheetDetailPage.GetMethod();
            string finalCoef = datasheetDetailPage.GetCoef();

            Assert.AreNotEqual(initMethod, finalMethod, "La valeur de la méthode de la recette n'a pas été modifiée.");
            Assert.AreEqual("100", finalCoef, "La valeur du coef de la recette n'a pas été modifié.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_UpdateRecipeMethod_Fraction()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();

            Random rnd = new Random();
            string datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            string newMethod = "1/x";
            string newMethod2 = "Std";
            newMethod = "1/X";

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            datasheetDetailPage.AddRecipe(recipeName);
            datasheetDetailPage.WaitLoading();
            WebDriver.Navigate().Refresh();
            Assert.IsTrue(datasheetDetailPage.IsRecipeAdded(), "La recette n'a pas été rajoutée à la datasheet.");

            datasheetDetailPage.ClickOnFirstRecipe();
            string initMethod = datasheetDetailPage.GetMethod();

            if (newMethod.Equals(initMethod))
            {
                datasheetDetailPage.SetMethod(newMethod2);
                initMethod = datasheetDetailPage.GetMethod();
            }

            datasheetDetailPage.SetMethod(newMethod);
            WebDriver.Navigate().Refresh();

            datasheetDetailPage.ClickOnFirstRecipe();
            string finalMethod = datasheetDetailPage.GetMethod();
            string finalCoef = datasheetDetailPage.GetCoef();

            Assert.AreNotEqual(initMethod, finalMethod, "La valeur de la méthode de la recette n'a pas été modifiée.");
            Assert.AreEqual("1", finalCoef, "La valeur du coef de la recette n'a pas été modifié.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_UpdateRecipeMethod_Scale()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();

            Random rnd = new Random();
            string datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();

            string newMethod = "Scale";
            string newMethod2 = "Std";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            datasheetDetailPage.AddRecipe(recipeName);
            datasheetDetailPage.WaitLoading();
            WebDriver.Navigate().Refresh();
            Assert.IsTrue(datasheetDetailPage.IsRecipeAdded(), "La recette n'a pas été rajoutée à la datasheet.");

            datasheetDetailPage.ClickOnFirstRecipe();
            string initMethod = datasheetDetailPage.GetMethod();

            if (newMethod.Equals(initMethod))
            {
                datasheetDetailPage.SetMethod(newMethod2);
                initMethod = datasheetDetailPage.GetMethod();
            }

            datasheetDetailPage.SetMethod(newMethod);
            datasheetDetailPage.FillPopUpScale();

            string finalMethod = datasheetDetailPage.GetMethod();
            string finalCoef = datasheetDetailPage.GetCoef();

            Assert.AreNotEqual(initMethod, finalMethod, "La valeur de la méthode de la recette n'a pas été modifiée.");
            Assert.AreEqual("1", finalCoef, "La valeur du coef de la recette n'a pas été modifié.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_UpdateRecipeCoef_Standard()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();

            Random rnd = new Random();
            string datasheetName = "Datasheet-" + DateUtils.Now.ToString("dd/MM/yyyy") + "-" + rnd.Next().ToString();

            string newMethod = "Std";
            string newCoef = "2";

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            datasheetDetailPage.AddRecipe(recipeName);
            datasheetDetailPage.WaitLoading();
            WebDriver.Navigate().Refresh();
            bool isReceipeAdded = datasheetDetailPage.IsRecipeAdded();
            Assert.IsTrue(isReceipeAdded, "La recette n'a pas été rajoutée à la datasheet.");

            datasheetDetailPage.ClickOnFirstRecipe();
            datasheetDetailPage.SetMethod(newMethod);
            string initCoef = datasheetDetailPage.GetCoef();
            string initPAX = datasheetDetailPage.GetPAXRecipe();

            datasheetDetailPage.SetCoef(newCoef);
            string finalCoef = datasheetDetailPage.GetCoef();
            string finalPAX = datasheetDetailPage.GetPAXRecipe();

            Assert.AreNotEqual(initCoef, finalCoef, "La valeur du coef de la recette n'a pas été modifié.");
            Assert.AreNotEqual(initPAX, finalPAX, "La valeur du PAX de la recette n'a pas été modifié avec le coef.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_UpdateRecipeCoef_Pourcentage()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();

            Random rnd = new Random();
            string datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();

            string newMethod = "%";
            string newCoef = "200";

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            datasheetDetailPage.AddRecipe(recipeName);
            datasheetDetailPage.WaitLoading();
            WebDriver.Navigate().Refresh();
            bool isReceipeAdded = datasheetDetailPage.IsRecipeAdded();
            Assert.IsTrue(isReceipeAdded, "La recette n'a pas été rajoutée à la datasheet.");

            datasheetDetailPage.ClickOnFirstRecipe();
            datasheetDetailPage.SetMethod(newMethod);
            string initCoef = datasheetDetailPage.GetCoef();
            string initPAX = datasheetDetailPage.GetPAXRecipe();

            datasheetDetailPage.SetCoef(newCoef);
            string finalCoef = datasheetDetailPage.GetCoef();
            string finalPAX = datasheetDetailPage.GetPAXRecipe();

            Assert.AreNotEqual(initCoef, finalCoef, "La valeur du coef de la recette n'a pas été modifié.");
            Assert.AreNotEqual(initPAX, finalPAX, "La valeur du PAX de la recette n'a pas été modifié avec le coef.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_UpdateRecipeCoef_Fraction()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();

            Random rnd = new Random();
            string datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();

            string newMethod = "1/x";
            string newCoef = "2";

            //Arrange
            var homePage = LogInAsAdmin();
            newMethod = "1/X";
            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            datasheetDetailPage.AddRecipe(recipeName);
            datasheetDetailPage.WaitLoading();
            WebDriver.Navigate().Refresh();
            Assert.IsTrue(datasheetDetailPage.IsRecipeAdded(), "La recette n'a pas été rajoutée à la datasheet.");

            datasheetDetailPage.ClickOnFirstRecipe();
            datasheetDetailPage.SetMethod(newMethod);
            string initCoef = datasheetDetailPage.GetCoef();
            string initPAX = datasheetDetailPage.GetPAXRecipe();

            datasheetDetailPage.SetCoef(newCoef);
            string finalCoef = datasheetDetailPage.GetCoef();
            string finalPAX = datasheetDetailPage.GetPAXRecipe();

            Assert.AreNotEqual(initCoef, finalCoef, "La valeur du coef de la recette n'a pas été modifié.");
            Assert.AreNotEqual(initPAX, finalPAX, "La valeur du PAX de la recette n'a pas été modifié avec le coef.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_UpdateRecipeCoef_Scale()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();

            Random rnd = new Random();
            string datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();

            string newMethod = "Scale";
            string newCoef = "2";

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            datasheetDetailPage.AddRecipe(recipeName);
            datasheetDetailPage.WaitLoading();
            WebDriver.Navigate().Refresh();
            bool isReceipeAdded = datasheetDetailPage.IsRecipeAdded();
            Assert.IsTrue(isReceipeAdded, "La recette n'a pas été rajoutée à la datasheet.");

            datasheetDetailPage.ClickOnFirstRecipe();
            datasheetDetailPage.SetMethod(newMethod);
            datasheetDetailPage.FillPopUpScale();

            string initCoef = datasheetDetailPage.GetCoef();

            Assert.AreEqual("1", initCoef, "La valeur du coef en mode 'Scale' est différente de 1.");
            Assert.IsFalse(datasheetDetailPage.SetCoef(newCoef), "La modification du coef a pu être réalisée en mode Scale.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_AddPictureToRecipe()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();

            Random rnd = new Random();
            string datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();

            var path = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase);
            path = path.Substring(6);
            var imagePath = path + "\\PageObjects\\Menus\\Datasheet\\test.jpg";

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            datasheetDetailPage.AddRecipe(recipeName);
            datasheetDetailPage.WaitPageLoading();
            Assert.IsTrue(datasheetDetailPage.IsRecipeAdded(), "La recette n'a pas été rajoutée à la datasheet (cas 1)");
            WebDriver.Navigate().Refresh();
            Assert.IsTrue(datasheetDetailPage.IsRecipeAdded(), "La recette n'a pas été rajoutée à la datasheet (cas 2)");

            var recipeVariantPage = datasheetDetailPage.AddPictureToRecipe();
            recipeVariantPage.AddPicture(imagePath);
            datasheetDetailPage = recipeVariantPage.CloseWindow();

            bool isIconGreen = recipeVariantPage.IsGreenIconDisplayedForRecipe(recipeName);

            bool isPictureAd = datasheetDetailPage.IsPictureAddedToRecipe();

            Assert.IsTrue(isPictureAd, "L'image n'a pas été ajoutée à la recette.");
            Assert.IsTrue(isIconGreen, "Le picto vert n'apparait pas devant le nom de la recette.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_ModifyPictureToRecipe()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();

            Random rnd = new Random();
            string datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();

            var path = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase);
            path = path.Substring(6);
            var imagePath = path + "\\PageObjects\\Menus\\Datasheet\\test.jpg";
            var imagePath2 = path + "\\PageObjects\\Menus\\Datasheet\\test2.jpg";

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            datasheetDetailPage.AddRecipe(recipeName);
            datasheetDetailPage.WaitLoading();
            WebDriver.Navigate().Refresh();
            Assert.IsTrue(datasheetDetailPage.IsRecipeAdded(), "La recette n'a pas été rajoutée à la datasheet.");

            var recipeVariantPage = datasheetDetailPage.AddPictureToRecipe();
            recipeVariantPage.AddPicture(imagePath);
            datasheetDetailPage = recipeVariantPage.CloseWindow();
            datasheetDetailPage.ClickOnRecipesTab();
            Assert.IsTrue(datasheetDetailPage.IsPictureAddedToRecipe(), "L'image n'a pas été ajoutée à la recette.");

            recipeVariantPage = datasheetDetailPage.AddPictureToRecipe();
            string url1 = recipeVariantPage.GetUrlPicture();
            recipeVariantPage.AddPicture(imagePath2);
            datasheetDetailPage = recipeVariantPage.CloseWindow();

            bool isIconGreen = recipeVariantPage.IsGreenIconDisplayedForRecipe(recipeName);
            Assert.IsTrue(isIconGreen, "Le picto vert n'apparait pas devant le nom de la recette.");

            datasheetDetailPage.ClickOnRecipesTab();
            Assert.IsTrue(datasheetDetailPage.IsPictureAddedToRecipe(), "L'image n'est plus présente dans la recette.");


            recipeVariantPage = datasheetDetailPage.AddPictureToRecipe();
            string url2 = recipeVariantPage.GetUrlPicture();

            Assert.AreNotEqual(url1, url2, "L'image n'a pas été mise à jour pour la recette.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_DeletePictureToRecipe()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();

            Random rnd = new Random();
            string datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();

            var path = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase);
            path = path.Substring(6);
            var imagePath = path + "\\PageObjects\\Menus\\Datasheet\\test.jpg";

            //Arrange
            HomePage homePage = LogInAsAdmin();
            try
            {
                //Act
                var datasheetPage = homePage.GoToMenus_Datasheet();
                var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
                var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

                datasheetDetailPage.AddRecipe(recipeName);
                datasheetDetailPage.WaitLoading();
                WebDriver.Navigate().Refresh();
                //Assert
                var isRecipeAddedOk = datasheetDetailPage.IsRecipeAdded();
                Assert.IsTrue(isRecipeAddedOk, "La recette n'a pas été rajoutée à la datasheet.");

                var recipeVariantPage = datasheetDetailPage.AddPictureToRecipe();
                recipeVariantPage.AddPicture(imagePath);
                datasheetDetailPage = recipeVariantPage.CloseWindow();
                //Assert
                var isPictureAddedToRecipe = datasheetDetailPage.IsPictureAddedToRecipe();
                Assert.IsTrue(isPictureAddedToRecipe, "L'image n'a pas été ajoutée à la recette.");

                recipeVariantPage = datasheetDetailPage.AddPictureToRecipe();
                recipeVariantPage.DeletePicture();
                datasheetDetailPage = recipeVariantPage.CloseWindow();
                datasheetDetailPage.ClickOnRecipesTab();
                //Assert
                var IsPictureAddedToRecipeOk = datasheetDetailPage.IsPictureAddedToRecipe();
                Assert.IsFalse(IsPictureAddedToRecipeOk, "L'image n'a pas été supprimée de la recette.");
            }
            finally
            {
                var datasheetPage = homePage.GoToMenus_Datasheet();
                datasheetPage.ResetFilter();
                datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
                DatasheetDetailsPage details = datasheetPage.SelectFirstDatasheet();
                DatasheetGeneralInformationPage generalInfo = details.ClickOnGeneralInformation();
                generalInfo.SetActive(false);
                generalInfo.DeleteDatasheet();
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_UpdateRecipeDietetic()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();

            Random rnd = new Random();
            string datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            datasheetDetailPage.AddRecipe(recipeName);
            datasheetDetailPage.WaitLoading();
            WebDriver.Navigate().Refresh();
            Assert.IsTrue(datasheetDetailPage.IsRecipeAdded(), "La recette n'a pas été rajoutée à la datasheet.");

            var recipeVariantPage = datasheetDetailPage.EditRecipe();
            recipeVariantPage.ClickOnDieteticTab();

            //Passage en computed, si on est en manuel
            if (!recipeVariantPage.IsDieteticComputed())
            {
                recipeVariantPage.ChangeDieteticComputedManual();
            }

            //Passage en manuel
            recipeVariantPage.ChangeDieteticComputedManual();
            Assert.IsFalse(recipeVariantPage.IsDieteticComputed(), "Le mode de calcul de la valeur diététique n'a pas été modifié en manuel.");
            Assert.IsTrue(recipeVariantPage.IsMalualDataVisible(), "Le nom de l'utilisateur et l'horaire de validation ne sont pas visibles.");

            //Passage en computed
            recipeVariantPage.ChangeDieteticComputedManual();
            Assert.IsTrue(recipeVariantPage.IsDieteticComputed(), "Le mode de calcul de la valeur diététique n'a pas été modifié en automatique.");
            Assert.IsFalse(recipeVariantPage.IsMalualDataVisible(), "Le nom de l'utilisateur et l'horaire de validation sont visibles.");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_DuplicateRecipeError()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();

            string datasheetName = datasheetNameToday + "-" + new Random().Next().ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            DatasheetPage datasheetPage = homePage.GoToMenus_Datasheet();
            DatasheetCreateModalPage datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            DatasheetDetailsPage datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            datasheetDetailPage.AddRecipe(recipeName);
            datasheetDetailPage.WaitLoading();
            WebDriver.Navigate().Refresh();
            bool isAdded = datasheetDetailPage.IsRecipeAdded();
            Assert.IsTrue(isAdded, "La recette n'a pas été rajoutée à la datasheet.");

            datasheetDetailPage.DuplicateRecipeError();
            isAdded = datasheetDetailPage.IsPopUpDuplicateErrorVisible();
            Assert.IsTrue(isAdded, "La duplication d'une recette existante dans la liste a été effectuée.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_DuplicateRecipe()
        {
            //Prepare
            var site = TestContext.Properties["Site"].ToString();
            var guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            var ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            var rnd = new Random();
            var datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();
            var recipeName = "RecipeTest-" + DateUtils.Now.ToString("dd/MM/yyyy") + "-" + rnd.Next().ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();
            try
            {
                //Act
                var recipesPage = homePage.GoToMenus_Recipes();
                recipesPage.ResetFilter();
                recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);
                //Assert
                var nbRecipe = recipesPage.CheckTotalNumber(); 
                Assert.AreEqual(0, nbRecipe, "La recette existe déjà dans la liste des recettes.");

                //Act
                var datasheetPage = homePage.GoToMenus_Datasheet();
                var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
                var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

                var datasheetCreateNewRecipePage = datasheetDetailPage.CreateNewRecipe();
                datasheetDetailPage = datasheetCreateNewRecipePage.FillFields_AddNewRecipeToDatasheet(recipeName, ingredient);
                //Assert
                var isRecipeAddedOk = datasheetDetailPage.IsRecipeAdded();
                Assert.IsTrue(isRecipeAddedOk, "La recette n'a pas été rajoutée à la datasheet.");

                var recipesVariantPage = datasheetDetailPage.DuplicateRecipe();
                recipesVariantPage.Close();

                recipesPage = homePage.GoToMenus_Recipes();
                recipesPage.ResetFilter();
                recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);
                //Assert
                var nbNewRecipe = recipesPage.CheckTotalNumber(); 
                Assert.AreEqual(1, nbNewRecipe, "La recette n'a pas été dupliquée dans la liste des recettes.");
            }
            finally
            {
                var datasheetPage = homePage.GoToMenus_Datasheet();
                datasheetPage.ResetFilter();
                datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
                DatasheetDetailsPage details = datasheetPage.SelectFirstDatasheet();
                DatasheetGeneralInformationPage generalInfo = details.ClickOnGeneralInformation();
                generalInfo.SetActive(false);
                generalInfo.DeleteDatasheet();
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_ConsultUseCaseForDatasheet()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string service = TestContext.Properties["DatasheetService"].ToString();

            string datasheetName = datasheetNameToday + "-" + new Random().Next().ToString();
            string serviceName = service + DateUtils.Now.ToString("dd/MM/yyyy") + "-" + new Random().Next().ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            AddServiceForDatasheet(homePage, serviceName, guestType, DateUtils.Now, DateUtils.Now.AddDays(10), datasheetName);

            datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
            var datasheetDetailPage = datasheetPage.SelectFirstDatasheet();

            datasheetDetailPage.ClickOnUseCaseTab();
            var servicePricePage = datasheetDetailPage.SelectService(serviceName);
            bool isNewWindow = true;

            try
            {
                var serviceGeneralInfo = servicePricePage.ClickOnGeneralInformationTab();
                Assert.AreEqual(serviceName, serviceGeneralInfo.GetServiceName(), "La page d'édition du service n'a pas été trouvée --> le service n'a pas été ajouté correctement à la datasheet.");
                serviceGeneralInfo.Close();
                isNewWindow = false;
            }
            finally
            {
                if (isNewWindow)
                {
                    servicePricePage.Close();
                }
            }
        }

        // _______________________________________________ Filtres _____________________________________________________________

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Filter_ByDatasheetName()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();

            string datasheetName = datasheetNameToday + "-" + new Random().Next().ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);
            datasheetPage = datasheetDetailPage.BackToList();

            //Assert
            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
            // pas de csutomer lié donc on uncheck all
            datasheetPage.Filter(DatasheetPage.FilterType.Customers, null);
            datasheetPage.Filter(DatasheetPage.FilterType.Categories, null);
            Assert.AreEqual(datasheetName, datasheetPage.GetFirstDatasheetName(), MessageErreur.FILTRE_ERRONE, "DatasheetName");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Filter_ByCustomerNameCode()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string service = TestContext.Properties["DatasheetService"].ToString();
            string customerCode = TestContext.Properties["MenuCustomerCode"].ToString();

            string datasheetName = datasheetNameToday + "-" + new Random().Next().ToString();
            string serviceName = service + DateUtils.Now.ToString("dd/MM/yyyy") + "-" + new Random().Next().ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site, customerCode);

            AddServiceForDatasheet(homePage, serviceName, guestType, DateUtils.Now, DateUtils.Now.AddDays(10), datasheetName);

            datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetPage.ResetFilter();

            datasheetPage.Filter(DatasheetPage.FilterType.CustomerCode, customerCode);
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
            Assert.AreEqual(datasheetName, datasheetPage.GetFirstDatasheetName(), MessageErreur.FILTRE_ERRONE, "CustomerCode");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Filter_ByServiceName()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string service = TestContext.Properties["DatasheetService"].ToString();

            string datasheetName = datasheetNameToday + "-" + new Random().Next().ToString();
            string serviceName = service + DateUtils.Now.ToString("dd/MM/yyyy") + "-" + new Random().Next().ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            AddServiceForDatasheet(homePage, serviceName, guestType, DateUtils.Now, DateUtils.Now.AddDays(10), datasheetName);

            datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetPage.Filter(DatasheetPage.FilterType.ServiceName, serviceName);

            Assert.AreEqual(datasheetName, datasheetPage.GetFirstDatasheetName(), MessageErreur.FILTRE_ERRONE, "ServiceName");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Filter_SortById()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();

            string datasheetName = datasheetNameToday + "-" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetPage.ResetFilter();
            int totalNumber = datasheetPage.CheckTotalNumber();

            try
            {
                if (totalNumber < 20)
                {
                    var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
                    var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);
                    datasheetPage = datasheetDetailPage.BackToList();
                    datasheetPage.ResetFilter();
                }

                if (!datasheetPage.isPageSizeEqualsTo100())
                {
                    datasheetPage.PageSize("8");
                    datasheetPage.PageSize("100");
                }

                datasheetPage.Filter(DatasheetPage.FilterType.SortBy, "Id");
                var isSortedById = datasheetPage.IsSortedByID();

                //Assert
                Assert.IsTrue(isSortedById, String.Format(MessageErreur.FILTRE_ERRONE, "'Sort by Id'"));
            }
            finally
            {
                if (totalNumber < 20)
                {

                    datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
                    datasheetPage.MassiveDelete(site, guestType);
                }
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Filter_SortByName()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();

            string datasheetName = datasheetNameToday + "-" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetPage.ResetFilter();
            int totalNumber = datasheetPage.CheckTotalNumber();

            try
            {
                if (totalNumber < 20)
                {
                    var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
                    var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);
                    datasheetPage = datasheetDetailPage.BackToList();
                    datasheetPage.ResetFilter();
                }

                if (!datasheetPage.isPageSizeEqualsTo100())
                {
                    datasheetPage.PageSize("8");
                    datasheetPage.PageSize("100");
                }

                datasheetPage.Filter(DatasheetPage.FilterType.SortBy, "Name");
                // pas de customer liés donc on uncheck all
                datasheetPage.Filter(DatasheetPage.FilterType.Customers, null);
                datasheetPage.Filter(DatasheetPage.FilterType.Categories, null);
                var isSortedByName = datasheetPage.IsSortedByName();

                //Assert
                Assert.IsTrue(isSortedByName, String.Format(MessageErreur.FILTRE_ERRONE, "'Sort by name'"));
            }
            finally
            {
                if (totalNumber < 20)
                {

                    datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
                    datasheetPage.MassiveDelete(site, guestType);
                }
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Filter_DateFrom()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string service = TestContext.Properties["DatasheetService"].ToString();

            string datasheetName = datasheetNameToday + "-" + new Random().Next().ToString();
            string serviceName = service + DateUtils.Now.ToString("dd/MM/yyyy") + "-" + new Random().Next().ToString();

            DateTime startDate = DateUtils.Now;
            DateTime endDate = DateUtils.Now.AddDays(10);

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetPage.ResetFilter();

            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            AddServiceForDatasheet(homePage, serviceName, guestType, startDate, endDate, datasheetName);

            datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
            datasheetPage.Filter(DatasheetPage.FilterType.DateFrom, endDate.AddDays(10));

            Assert.AreEqual(0, datasheetPage.CheckTotalNumber(), MessageErreur.FILTRE_ERRONE, "'DateFrom'");

            datasheetPage.Filter(DatasheetPage.FilterType.DateFrom, startDate.AddDays(-10));

            Assert.AreEqual(0, datasheetPage.CheckTotalNumber(), MessageErreur.FILTRE_ERRONE, "'DateFrom'");

            datasheetPage.Filter(DatasheetPage.FilterType.DateFrom, startDate);

            Assert.AreEqual(1, datasheetPage.CheckTotalNumber(), MessageErreur.FILTRE_ERRONE, "'DateFrom'");

            datasheetPage.Filter(DatasheetPage.FilterType.DateFrom, endDate);

            Assert.AreEqual(1, datasheetPage.CheckTotalNumber(), MessageErreur.FILTRE_ERRONE, "'DateFrom'");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Filter_DateTo()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string service = TestContext.Properties["DatasheetService"].ToString();

            string datasheetName = datasheetNameToday + "-" + new Random().Next().ToString();
            string serviceName = service + DateUtils.Now.ToString("dd/MM/yyyy") + "-" + new Random().Next().ToString();

            DateTime startDate = DateUtils.Now;
            DateTime endDate = DateUtils.Now.AddDays(10);

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetPage.ResetFilter();

            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            AddServiceForDatasheet(homePage, serviceName, guestType, startDate, endDate, datasheetName);

            datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
            datasheetPage.Filter(DatasheetPage.FilterType.DateTo, endDate.AddDays(10));

            Assert.AreEqual(0, datasheetPage.CheckTotalNumber(), MessageErreur.FILTRE_ERRONE, "'DateTo'");

            datasheetPage.Filter(DatasheetPage.FilterType.DateTo, startDate.AddDays(-10));

            Assert.AreEqual(0, datasheetPage.CheckTotalNumber(), MessageErreur.FILTRE_ERRONE, "'DateTo'");

            datasheetPage.Filter(DatasheetPage.FilterType.DateTo, startDate);

            Assert.AreEqual(1, datasheetPage.CheckTotalNumber(), MessageErreur.FILTRE_ERRONE, "'DateTo'");

            datasheetPage.Filter(DatasheetPage.FilterType.DateTo, endDate);

            Assert.AreEqual(1, datasheetPage.CheckTotalNumber(), MessageErreur.FILTRE_ERRONE, "'DateTo'");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Filter_ShowActiveOnly()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();

            string datasheetName = datasheetNameToday + "-" + new Random().Next().ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.ShowActive, true);

            if (datasheetPage.CheckTotalNumber() < 20)
            {
                var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
                var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);
                datasheetDetailPage.AddRecipe(recipeName);
                datasheetPage = datasheetDetailPage.BackToList();
                datasheetPage.ResetFilter();
                datasheetPage.Filter(DatasheetPage.FilterType.ShowActive, true);
            }

            if (!datasheetPage.isPageSizeEqualsTo100())
            {
                datasheetPage.PageSize("8");
                datasheetPage.PageSize("100");
            }

            //Assert
            Assert.IsTrue(datasheetPage.CheckStatus(true), String.Format(MessageErreur.FILTRE_ERRONE, "'Show active only'"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Filter_ShowInactiveOnly()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();

            string datasheetName = datasheetNameToday + "-" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.ShowInactive, true);
            // pas de csutomer lié donc on uncheck all
            datasheetPage.Filter(DatasheetPage.FilterType.Customers, null);
            datasheetPage.Filter(DatasheetPage.FilterType.Categories, null);

            if (datasheetPage.CheckTotalNumber() < 20)
            {
                var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
                var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);
                datasheetDetailPage.AddRecipe(recipeName);
                var datasheetGeneralInfo = datasheetDetailPage.ClickOnGeneralInformation();
                datasheetGeneralInfo.SetActive(false);
                datasheetDetailPage = datasheetGeneralInfo.ClickOnDetailsTab();
                datasheetPage = datasheetDetailPage.BackToList();
                datasheetPage.ResetFilter();
                datasheetPage.Filter(DatasheetPage.FilterType.ShowInactive, true);
                // pas de csutomer lié donc on uncheck all
                datasheetPage.Filter(DatasheetPage.FilterType.Customers, null);
                datasheetPage.Filter(DatasheetPage.FilterType.Categories, null);
            }

            if (!datasheetPage.isPageSizeEqualsTo100())
            {
                datasheetPage.PageSize("8");
                datasheetPage.PageSize("100");
            }

            //Assert
            Assert.IsFalse(datasheetPage.CheckStatus(false), String.Format(MessageErreur.FILTRE_ERRONE, "'Show inactive only'"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Filter_ShowAll()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();

            string datasheetName = datasheetNameToday + "-" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetPage.ResetFilter();

            if (datasheetPage.CheckTotalNumber() < 20)
            {
                var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
                var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);
                datasheetPage = datasheetDetailPage.BackToList();
                datasheetPage.ResetFilter();
            }
            // pas de csutomer lié donc on uncheck all
            datasheetPage.Filter(DatasheetPage.FilterType.Customers, null);
            datasheetPage.Filter(DatasheetPage.FilterType.Categories, null);

            datasheetPage.Filter(DatasheetPage.FilterType.ShowActive, true);
            int nbActive = datasheetPage.CheckTotalNumber();

            datasheetPage.Filter(DatasheetPage.FilterType.ShowInactive, true);
            int nbInactive = datasheetPage.CheckTotalNumber();

            datasheetPage.Filter(DatasheetPage.FilterType.ShowAll, true);
            int nbTotal = datasheetPage.CheckTotalNumber();

            //Assert
            Assert.AreEqual((nbActive + nbInactive), nbTotal, String.Format(MessageErreur.FILTRE_ERRONE, "'Show all'"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Filter_UseCaseAffectedOnly()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string service = TestContext.Properties["DatasheetService"].ToString();

            string datasheetName = datasheetNameToday + "-" + new Random().Next().ToString();
            string serviceName = service + DateUtils.Now.ToString("dd/MM/yyyy") + "-" + new Random().Next().ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.UseCaseAffected, true);

            if (datasheetPage.CheckTotalNumber() < 20)
            {
                var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
                datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

                AddServiceForDatasheet(homePage, serviceName, guestType, DateUtils.Now, DateUtils.Now.AddDays(10), datasheetName);

                datasheetPage = homePage.GoToMenus_Datasheet();
                datasheetPage.ResetFilter();
                datasheetPage.Filter(DatasheetPage.FilterType.UseCaseAffected, true);
            }

            if (!datasheetPage.isPageSizeEqualsTo100())
            {
                datasheetPage.PageSize("8");
                datasheetPage.PageSize("100");
            }

            //Assert
            Assert.IsTrue(datasheetPage.IsUseCaseAffected(), String.Format(MessageErreur.FILTRE_ERRONE, "'Use case affected only'"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Filter_UseCaseNotAffectedOnly()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string datasheetName = datasheetNameToday + "-" + new Random().Next().ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.UseCaseNotAffected, true);

            if (datasheetPage.CheckTotalNumber() < 20)
            {
                var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
                var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);
                datasheetPage = datasheetDetailPage.BackToList();
                datasheetPage.ResetFilter();
                datasheetPage.Filter(DatasheetPage.FilterType.UseCaseNotAffected, true);
            }

            if (!datasheetPage.isPageSizeEqualsTo100())
            {
                datasheetPage.PageSize("8");
                datasheetPage.PageSize("100");
            }

            //Assert
            Assert.IsFalse(datasheetPage.IsUseCaseAffected(), String.Format(MessageErreur.FILTRE_ERRONE, "'Use case not affected only'"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Filter_UseCaseExpiredOnly()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string service = TestContext.Properties["DatasheetService"].ToString();

            string datasheetName = datasheetNameToday + "-" + new Random().Next().ToString();
            string serviceName = service + DateUtils.Now.ToString("dd/MM/yyyy") + "-" + new Random().Next().ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetPage.ResetFilter();

            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            // On ajoute un service expiré à la datasheet
            AddServiceForDatasheet(homePage, serviceName, guestType, DateUtils.Now.AddMonths(-2), DateUtils.Now.AddMonths(-1), datasheetName);

            datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
            datasheetPage.Filter(DatasheetPage.FilterType.UseCaseAffected, true);

            Assert.AreEqual(0, datasheetPage.CheckTotalNumber(), "La datasheet avec service expiré apparaît dans la liste des datasheet 'affectées'");

            WebDriver.Navigate().Refresh();
            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
            datasheetPage.Filter(DatasheetPage.FilterType.UseCaseExpired, true);

            //Assert
            Assert.AreEqual(1, datasheetPage.CheckTotalNumber(), String.Format(MessageErreur.FILTRE_ERRONE, "'Use case expired only'"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Filter_UseCaseShowAll()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string service = TestContext.Properties["DatasheetService"].ToString();

            string datasheetName = datasheetNameToday + "-" + new Random().Next().ToString();
            string serviceName = service + DateUtils.Now.ToString("dd/MM/yyyy") + "-" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetPage.ResetFilter();

            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);
            datasheetPage = datasheetDetailPage.BackToList();
            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);

            datasheetPage.Filter(DatasheetPage.FilterType.UseCaseAffected, true);
            // pas de customer lié, ni de recette, donc on uncheck all
            datasheetPage.Filter(DatasheetPage.FilterType.Customers, null);
            datasheetPage.Filter(DatasheetPage.FilterType.Categories, null);
            datasheetPage.Filter(DatasheetPage.FilterType.Workshops, null);
            datasheetPage.Filter(DatasheetPage.FilterType.CookingModes, null);
            Assert.AreEqual(0, datasheetPage.CheckTotalNumber(), "Un datasheet sans service apparaît dans la liste des datasheet avec service affecté.");

            datasheetPage.Filter(DatasheetPage.FilterType.UseCaseNotAffected, true);
            Assert.AreEqual(1, datasheetPage.CheckTotalNumber(), "Un datasheet sans service n'apparaît pas dans la liste des datasheet sans service affecté.");

            datasheetPage.Filter(DatasheetPage.FilterType.UseCaseShowAll, true);
            Assert.AreEqual(1, datasheetPage.CheckTotalNumber(), MessageErreur.FILTRE_ERRONE, "'Use case show all'");

            // On affecte un service à la datasheet
            AddServiceForDatasheet(homePage, serviceName, guestType, DateUtils.Now, DateUtils.Now.AddDays(10), datasheetName);

            datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
            // pas de customer lié, ni de recette, donc on uncheck all
            datasheetPage.Filter(DatasheetPage.FilterType.Customers, null);
            datasheetPage.Filter(DatasheetPage.FilterType.Categories, null);
            datasheetPage.Filter(DatasheetPage.FilterType.Workshops, null);
            datasheetPage.Filter(DatasheetPage.FilterType.CookingModes, null);

            datasheetPage.Filter(DatasheetPage.FilterType.UseCaseAffected, true);
            Assert.AreEqual(1, datasheetPage.CheckTotalNumber(), "Un datasheet avec service n'apparaît pas dans la liste des datasheet avec service affecté.");

            datasheetPage.Filter(DatasheetPage.FilterType.UseCaseNotAffected, true);
            Assert.AreEqual(0, datasheetPage.CheckTotalNumber(), "Un datasheet avec service apparaît dans la liste des datasheet sans service affecté.");

            datasheetPage.Filter(DatasheetPage.FilterType.UseCaseShowAll, true);
            Assert.AreEqual(1, datasheetPage.CheckTotalNumber(), MessageErreur.FILTRE_ERRONE, "'Use case show all'");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Filter_ShowAllergenValidated()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();
            string datasheetName = datasheetNameToday + "-" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetPage.ResetFilter();

            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);
            datasheetDetailPage.AddRecipe(recipeName);
            datasheetDetailPage.ClickOnIntoleranceTab();
            datasheetDetailPage.ValidateAllergen();
            Assert.IsTrue(datasheetDetailPage.IsAllergenValidated(), "Les intolerances associées à la datasheet n'ont pas été validées.");

            datasheetPage = datasheetDetailPage.BackToList();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
            datasheetPage.Filter(DatasheetPage.FilterType.AllergenValidated, true);
            // pas de csutomer lié donc on uncheck all
            datasheetPage.Filter(DatasheetPage.FilterType.Customers, null);
            datasheetPage.Filter(DatasheetPage.FilterType.Categories, null);
            //Assert
            Assert.IsTrue(datasheetPage.CheckAllergenValidation(true), String.Format(MessageErreur.FILTRE_ERRONE, "'Show allergen validated'"));


            WebDriver.Navigate().Refresh();
            datasheetPage.ResetFilter();
            // pas de csutomer lié donc on uncheck all
            datasheetPage.Filter(DatasheetPage.FilterType.Customers, null);
            datasheetPage.Filter(DatasheetPage.FilterType.Categories, null);
            datasheetPage.Filter(DatasheetPage.FilterType.AllergenValidated, true);

            if (!datasheetPage.isPageSizeEqualsTo100())
            {
                datasheetPage.PageSize("8");
                datasheetPage.PageSize("100");
            }

            //Assert
            Assert.IsTrue(datasheetPage.CheckAllergenValidation(true), String.Format(MessageErreur.FILTRE_ERRONE, "'Show allergen validated'"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Filter_ShowAllergenNotValidated()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();
            string datasheetName = datasheetNameToday + "-" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetPage.ResetFilter();

            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);
            datasheetDetailPage.AddRecipe(recipeName);

            datasheetPage = datasheetDetailPage.BackToList();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
            datasheetPage.Filter(DatasheetPage.FilterType.AllergenNotValidated, true);

            //Assert
            Assert.IsFalse(datasheetPage.CheckAllergenValidation(false), String.Format(MessageErreur.FILTRE_ERRONE, "'Show allergen not validated'"));

            WebDriver.Navigate().Refresh();
            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.AllergenNotValidated, true);

            if (!datasheetPage.isPageSizeEqualsTo100())
            {
                datasheetPage.PageSize("8");
                datasheetPage.PageSize("100");
            }

            //Assert
            Assert.IsFalse(datasheetPage.CheckAllergenValidation(false), String.Format(MessageErreur.FILTRE_ERRONE, "'Show allergen not validated'"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Filter_ShowAllAllergen()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();
            string datasheetName = datasheetNameToday + "-" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.AllergenShowAll, true);

            if (datasheetPage.CheckTotalNumber() < 20)
            {
                var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
                var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);
                datasheetDetailPage.AddRecipe(recipeName);
                datasheetPage = datasheetDetailPage.BackToList();
                datasheetPage.ResetFilter();
            }

            datasheetPage.Filter(DatasheetPage.FilterType.AllergenValidated, true);
            int nbValidated = datasheetPage.CheckTotalNumber();

            datasheetPage.Filter(DatasheetPage.FilterType.AllergenNotValidated, true);
            int nbNotValidated = datasheetPage.CheckTotalNumber();

            datasheetPage.Filter(DatasheetPage.FilterType.AllergenShowAll, true);
            int nbTotal = datasheetPage.CheckTotalNumber();

            //Assert
            Assert.AreEqual((nbValidated + nbNotValidated), nbTotal, String.Format(MessageErreur.FILTRE_ERRONE, "'Show all allergen'"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Filter_Customers()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string service = TestContext.Properties["DatasheetService"].ToString();
            string customer = TestContext.Properties["MenuServiceCustomer"].ToString();
            string autreCustomer = TestContext.Properties["CustomerDelivery"].ToString();

            string datasheetName = datasheetNameToday + "-" + new Random().Next().ToString();
            string serviceName = service + DateUtils.Now.ToString("dd/MM/yyyy") + "-" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            var customerPage = homePage.GoToCustomers_CustomerPage();
            customerPage.ResetFilters();
            customerPage.Filter(PageObjects.Customers.Customer.CustomerPage.FilterType.Search, customer);
            string customerCode = customerPage.GetFirstCustomerIcao();

            customerPage.ResetFilters();
            customerPage.Filter(PageObjects.Customers.Customer.CustomerPage.FilterType.Search, autreCustomer);
            string autreCustomerCode = customerPage.GetFirstCustomerIcao();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetPage.ResetFilter();

            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site, customerCode);

            // On ajoute un service expiré à la datasheet
            AddServiceForDatasheet(homePage, serviceName, guestType, DateUtils.Now, DateUtils.Now.AddMonths(2), datasheetName);

            datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
            datasheetPage.Filter(DatasheetPage.FilterType.CustomerCode, autreCustomerCode);

            Assert.AreEqual(0, datasheetPage.CheckTotalNumber(), "La datasheet de customer " + customer + " apparaît dans la liste des datasheet du customer " + autreCustomer + ".");

            WebDriver.Navigate().Refresh();
            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
            datasheetPage.Filter(DatasheetPage.FilterType.CustomerCode, customerCode);

            //Assert
            Assert.AreEqual(1, datasheetPage.CheckTotalNumber(), String.Format(MessageErreur.FILTRE_ERRONE, "'Customers'"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Filter_Categories()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string service = TestContext.Properties["DatasheetService"].ToString();
            string serviceCategory = TestContext.Properties["CategoryServiceDefault"].ToString();
            string datasheetName = datasheetNameToday + "-" + new Random().Next().ToString();
            string serviceName = service + DateUtils.Now.ToString("dd/MM/yyyy") + "-" + new Random().Next().ToString();

            string autreServiceCategory = TestContext.Properties["CategoryService"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetPage.ResetFilter();

            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            // On ajoute un service à la datasheet
            AddServiceForDatasheet(homePage, serviceName, guestType, DateUtils.Now, DateUtils.Now.AddMonths(2), datasheetName);

            datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
            datasheetPage.Filter(DatasheetPage.FilterType.Categories, autreServiceCategory);

            Assert.AreEqual(0, datasheetPage.CheckTotalNumber(), "La datasheet de catégorie " + serviceCategory + " apparaît dans la liste des datasheet de catégorie " + autreServiceCategory + ".");

            WebDriver.Navigate().Refresh();
            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
            datasheetPage.Filter(DatasheetPage.FilterType.Categories, serviceCategory);

            //Assert
            Assert.AreEqual(1, datasheetPage.CheckTotalNumber(), String.Format(MessageErreur.FILTRE_ERRONE, "'Categories'"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Filter_Sites()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();

            string autreSite = TestContext.Properties["SiteLP"].ToString();
            string datasheetName = datasheetNameToday + "-" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetPage.ResetFilter();

            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);
            datasheetDetailPage.AddRecipe(recipeName);
            datasheetPage = datasheetDetailPage.BackToList();

            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
            // pas de csutomer lié donc on uncheck all
            datasheetPage.Filter(DatasheetPage.FilterType.Customers, null);
            datasheetPage.Filter(DatasheetPage.FilterType.Categories, null);
            datasheetPage.Filter(DatasheetPage.FilterType.Sites, autreSite);

            Assert.AreEqual(0, datasheetPage.CheckTotalNumber(), "La datasheet de site " + site + " apparaît dans la liste des datasheet du site " + autreSite + ".");

            WebDriver.Navigate().Refresh();
            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
            // pas de csutomer lié donc on uncheck all
            datasheetPage.Filter(DatasheetPage.FilterType.Customers, null);
            datasheetPage.Filter(DatasheetPage.FilterType.Categories, null);
            datasheetPage.Filter(DatasheetPage.FilterType.Sites, site);

            //Assert
            Assert.AreEqual(1, datasheetPage.CheckTotalNumber(), String.Format(MessageErreur.FILTRE_ERRONE, "'Sites'"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Filter_GuestType()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();

            string autreGuestType = TestContext.Properties["Guest"].ToString();
            string datasheetName = datasheetNameToday + "-" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetPage.ResetFilter();

            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);
            datasheetDetailPage.AddRecipe(recipeName);
            datasheetPage = datasheetDetailPage.BackToList();
            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
            datasheetPage.Filter(DatasheetPage.FilterType.GuestTypes, autreGuestType);

            Assert.AreEqual(0, datasheetPage.CheckTotalNumber(), "La datasheet de guest type " + guestType + " apparaît dans la liste des datasheet de guestType " + autreGuestType + ".");

            WebDriver.Navigate().Refresh();
            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
            datasheetPage.Filter(DatasheetPage.FilterType.GuestTypes, guestType);
            // pas de csutomer lié donc on uncheck all
            datasheetPage.Filter(DatasheetPage.FilterType.Customers, null);
            datasheetPage.Filter(DatasheetPage.FilterType.Categories, null);

            //Assert
            Assert.AreEqual(1, datasheetPage.CheckTotalNumber(), String.Format(MessageErreur.FILTRE_ERRONE, "'Guest Types'"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Filter_RecipesWorkshop()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();

            Random rnd = new Random();
            string datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();
            string recipeName = "RecipeTest-" + rnd.Next().ToString();
            string workshop = "Cocina Caliente";
            string autreWorkshop = "Corte";


            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetPage.ResetFilter();

            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);
            var datasheetCreateNewRecipePage = datasheetDetailPage.CreateNewRecipe();
            datasheetDetailPage = datasheetCreateNewRecipePage.FillFields_AddNewRecipeToDatasheet(recipeName, ingredient, null, workshop);
            datasheetPage = datasheetDetailPage.BackToList();

            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
            // pas de csutomer lié donc on uncheck all
            datasheetPage.Filter(DatasheetPage.FilterType.Customers, null);
            datasheetPage.Filter(DatasheetPage.FilterType.Categories, null);
            datasheetPage.Filter(DatasheetPage.FilterType.Workshops, autreWorkshop);

            Assert.AreEqual(0, datasheetPage.CheckTotalNumber(), "La datasheet de workshop " + workshop + " apparaît dans la liste des datasheet de workshop " + autreWorkshop + ".");

            WebDriver.Navigate().Refresh();
            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
            // pas de csutomer lié donc on uncheck all
            datasheetPage.Filter(DatasheetPage.FilterType.Customers, null);
            datasheetPage.Filter(DatasheetPage.FilterType.Categories, null);
            datasheetPage.Filter(DatasheetPage.FilterType.Workshops, workshop);

            //Assert
            Assert.AreEqual(1, datasheetPage.CheckTotalNumber(), String.Format(MessageErreur.FILTRE_ERRONE, "'Recipe workshop'"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Filter_RecipesCookingModes()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();

            Random rnd = new Random();
            string datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();
            string recipeName = "RecipeTest-" + rnd.Next().ToString();
            string cookingMode = "Cocina Caliente";
            string autreCookingMode = "Vacío";


            //Arrange
            HomePage homePage = LogInAsAdmin();


            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetPage.ResetFilter();

            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);
            var datasheetCreateNewRecipePage = datasheetDetailPage.CreateNewRecipe();
            datasheetDetailPage = datasheetCreateNewRecipePage.FillFields_AddNewRecipeToDatasheet(recipeName, ingredient, cookingMode);
            datasheetPage = datasheetDetailPage.BackToList();

            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
            // pas de csutomer lié donc on uncheck all
            datasheetPage.Filter(DatasheetPage.FilterType.Customers, null);
            datasheetPage.Filter(DatasheetPage.FilterType.Categories, null);

            datasheetPage.Filter(DatasheetPage.FilterType.CookingModes, autreCookingMode);

            Assert.AreEqual(0, datasheetPage.CheckTotalNumber(), "La datasheet de cooking mode " + cookingMode + " apparaît dans la liste des datasheet de cooking mode " + autreCookingMode + ".");

            WebDriver.Navigate().Refresh();
            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
            // pas de csutomer lié donc on uncheck all
            datasheetPage.Filter(DatasheetPage.FilterType.Customers, null);
            datasheetPage.Filter(DatasheetPage.FilterType.Categories, null);
            datasheetPage.Filter(DatasheetPage.FilterType.CookingModes, cookingMode);

            //Assert
            Assert.AreEqual(1, datasheetPage.CheckTotalNumber(), String.Format(MessageErreur.FILTRE_ERRONE, "'Cooking modes'"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_UpdateSubRecipeGeneralInformation()
        {
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();
            string cookingModeName = TestContext.Properties["Production_CookingMode1"].ToString();
            string recipeTypeName = TestContext.Properties["Prodman_Needs_RecipeType2"].ToString();
            string dietaryTypeName = TestContext.Properties["RecipeDietaryType"].ToString();
            string workshopName = TestContext.Properties["Workshop"].ToString();

            Random rnd = new Random();
            string datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();

            string commercialName1 = "commName1-" + rnd.Next().ToString();
            string commercialName2 = "commName2-" + rnd.Next().ToString();
            string conserveMode = "conserveMode-" + rnd.Next().ToString();
            string conserveTemp = "conserveTemp-" + rnd.Next().ToString();
            string comment = "comment-" + rnd.Next().ToString();

            //Home page
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            datasheetDetailPage.AddRecipe(recipeName);
            datasheetDetailPage.WaitLoading();
            WebDriver.Navigate().Refresh();
            Assert.IsTrue(datasheetDetailPage.IsRecipeAdded(), "La recette n'a pas été rajoutée à la datasheet.");

            var recipeVariantPage = datasheetDetailPage.EditRecipe();
            var recipeGIPage = recipeVariantPage.ClickOnGeneralInformationFromRecipeDatasheet();

            //Set values
            recipeGIPage.SetCommercialName1(commercialName1);
            recipeGIPage.SetCommercialName2(commercialName2);
            recipeGIPage.SetCookingMode(cookingModeName);
            recipeGIPage.SetRecipeType(recipeTypeName);
            recipeGIPage.SetDietaryType(dietaryTypeName);
            recipeGIPage.SetConserveMode(conserveMode);
            recipeGIPage.SetConserveTemperature(conserveTemp);
            recipeGIPage.SetComment(comment);
            recipeGIPage.SetWorkshop(workshopName);
            recipeGIPage.WaitLoading();

            //Reload page and check that everything was saved
            recipeVariantPage.ClickOnProcedureTab();
            recipeGIPage = recipeVariantPage.ClickOnGeneralInformationFromRecipeDatasheet();

            Assert.AreEqual(recipeGIPage.GetCommercialName1(), commercialName1, "Commercial name 1 was different from expected.");
            Assert.AreEqual(recipeGIPage.GetCommercialName2(), commercialName2, "Commercial name 1 was different from expected.");
            Assert.AreEqual(recipeGIPage.GetCookingMode(), cookingModeName, "Cooking mode was different from expected.");
            Assert.AreEqual(recipeGIPage.GetRecipeType(), recipeTypeName, "Recipe type was different from expected.");
            Assert.AreEqual(recipeGIPage.GetDietaryType(), dietaryTypeName, "Dietary name was different from expected.");
            Assert.AreEqual(recipeGIPage.GetConserveMode(), conserveMode, "Conserve mode name was different from expected.");
            Assert.AreEqual(recipeGIPage.GetConserveTemperature(), conserveTemp, "Conserve temp was different from expected.");
            Assert.AreEqual(recipeGIPage.GetComment(), comment, "Comment was different from expected.");
            Assert.AreEqual(recipeGIPage.GetWorkshop(), workshopName, "Workshop name was different from expected.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_DeleteDatasheet()
        {
            //Prepare
            var site = TestContext.Properties["Site"].ToString();
            var guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            var ingredient = TestContext.Properties["RecipeIngredient"].ToString();

            var rnd = new Random();
            var datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();
            var recipeName = "RecipeTest-" + rnd.Next().ToString();
            var cookingMode = "Cocina Caliente";


            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetPage.ResetFilter();

            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);
            var datasheetCreateNewRecipePage = datasheetDetailPage.CreateNewRecipe();
            datasheetDetailPage = datasheetCreateNewRecipePage.FillFields_AddNewRecipeToDatasheet(recipeName, ingredient, cookingMode);
            datasheetPage = datasheetDetailPage.BackToList();

            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
            // pas de csutomer lié donc on uncheck all
            datasheetPage.Filter(DatasheetPage.FilterType.Customers, null);
            datasheetPage.Filter(DatasheetPage.FilterType.Categories, null);
            Assert.AreEqual(1, datasheetPage.CheckTotalNumber(), "Datasheet non visible dans l'index");

            DatasheetDetailsPage details = datasheetPage.SelectFirstDatasheet();
            DatasheetGeneralInformationPage generalInfo = details.ClickOnGeneralInformation();
            generalInfo.SetActive(false);
            generalInfo.DeleteDatasheet();
            datasheetPage = generalInfo.BackToList();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
            // pas de customer lié donc on uncheck all
            datasheetPage.Filter(DatasheetPage.FilterType.Customers, null);
            datasheetPage.Filter(DatasheetPage.FilterType.Categories, null);
            Assert.AreEqual(0, datasheetPage.CheckTotalNumber(), "Datasheet non effacé");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_DeleteSubRecipe()
        {
            //Prepare
            //first datasheet , recipe and ingredient 
            var site = TestContext.Properties["Site"].ToString();
            var guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            var ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            var ingredientBis = TestContext.Properties["RecipeIngredientBis"].ToString();
            var rnd = new Random();
            var datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();
            var recipeName = "RecipeTest-" + rnd.Next().ToString();
            var cookingMode = "Cocina Caliente";
            var rnd2 = new Random();
            var datasheetName2 = datasheetNameToday + "-" + rnd2.Next().ToString() + "2";

            //arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var datasheetPage = homePage.GoToMenus_Datasheet();

            //act
            datasheetPage.ResetFilter();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);
            var datasheetCreateNewRecipePage = datasheetDetailPage.CreateNewRecipe();
            datasheetDetailPage = datasheetCreateNewRecipePage.FillFields_AddNewRecipeToDatasheet(recipeName, ingredient, cookingMode);
            datasheetPage = datasheetDetailPage.BackToList();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
            var datasheetDetails = datasheetPage.SelectFirstDatasheet();
            datasheetDetails.DeleteRecipe();
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.Filter(ItemPage.FilterType.Search, ingredient);
            var itemDetails = itemPage.ClickOnFirstItem();
            var itemUseCase = itemDetails.ClickOnUseCasePage();
            var recipeExist = itemUseCase.VerifierRecipeExist(recipeName);
            Assert.IsFalse(recipeExist, "Recipe existe");

            //create new datasheet

            datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName2, guestType, site);
            datasheetCreateNewRecipePage = datasheetDetailPage.CreateNewRecipe();
            datasheetDetailPage = datasheetCreateNewRecipePage.FillFields_AddNewRecipeToDatasheet(recipeName, ingredient, cookingMode);

            //add subrecipe 
            datasheetCreateNewRecipePage = datasheetDetailPage.AddSubrecipeForFirstRecipe();
            datasheetCreateNewRecipePage.FillFields_AddNewRecipeToDatasheet(recipeName, ingredientBis, cookingMode);

            //delete sub recipe 
            datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName2);
            datasheetPage.Filter(DatasheetPage.FilterType.Customers, null);
            datasheetPage.Filter(DatasheetPage.FilterType.Categories, null);
            datasheetPage.Filter(DatasheetPage.FilterType.Workshops, null);
            datasheetPage.Filter(DatasheetPage.FilterType.CookingModes, null);
            datasheetDetails = datasheetPage.SelectFirstDatasheet();
            var recipeVariantPage = datasheetDetails.EditRecipe();
            recipeVariantPage.DeleteSubRecipe(recipeName);
            itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.Filter(ItemPage.FilterType.Search, ingredientBis);
            itemDetails = itemPage.ClickOnFirstItem();
            itemUseCase = itemDetails.ClickOnUseCasePage();
            recipeExist = itemUseCase.VerifierRecipeExist(recipeName);
            Assert.IsFalse(recipeExist, "Recipe existe encore");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_MassiveDelete()
        {
            var site = "AGP";
            var customerName = $"Customer{new Random()}";
            var customerICAO = $"Cust{new Random()}";
            var customerType = "Ash";

            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();

            //Login
            HomePage homePage = LogInAsAdmin();

            //Create Customer 1
            var customerPage = homePage.GoToCustomers_CustomerPage();
            var customerCreatepage = customerPage.CustomerCreatePage();
            customerCreatepage.FillFields_CreateCustomerModalPage(customerName, customerICAO, customerType);
            customerCreatepage.Create();
            //Create 3 datasheets 
            for (int i = 0; i < 3; i++)
            {
                Random rnd = new Random();
                string datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();
                var dataSheetPage = homePage.GoToMenus_Datasheet();
                var createDataSheetPage = dataSheetPage.CreateNewDatasheet();
                createDataSheetPage.FillField_CreateNewDatasheet(datasheetName, guestType, site, customerICAO);
            }

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetPage.MassiveDelete(site, guestType);
            var allDeleted = datasheetPage.VerifierMassiveDelete(site, guestType);
            Assert.IsTrue(allDeleted, "erreur de massive delete");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Massivedelete_sitesearch()
        {
            string site = TestContext.Properties["SiteACE"].ToString();

            //Home page
            var homePage = LogInAsAdmin();

            var dataSheetPage = homePage.GoToMenus_Datasheet();
            var dsMassiveDelete = dataSheetPage.OpenMassiveDeletePopup();

            dsMassiveDelete.SelectSiteByName(site, true);
            dsMassiveDelete.ClickOnSearch();

            int rowNumber = 0;
            DatasheetMassiveDeleteRowResult dsRowResult = dsMassiveDelete.GetRowResultInfo(rowNumber);

            while (dsRowResult != null)
            {
                Assert.AreEqual(site, dsRowResult.SiteName, $"Le nom du site n'est pas égal à {site} pour la ligne de résultat {rowNumber}");
                rowNumber++;
                dsRowResult = dsMassiveDelete.GetRowResultInfo(rowNumber);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Massivedelete_customersearch()
        {
            string customerMD = "FRED OLSEN, S.A.";

            //Home page
            HomePage homePage = LogInAsAdmin();

            DatasheetPage dataSheetPage = homePage.GoToMenus_Datasheet();
            DatasheetMassiveDeletePopup dsMassiveDelete = dataSheetPage.OpenMassiveDeletePopup();

            dsMassiveDelete.SelectCustomerByName(customerMD, true);
            dsMassiveDelete.SelectStatus(MassiveDeleteStatus.Used.ToString());
            dsMassiveDelete.ClickOnSearch();

            int rowNumber = 0;
            DatasheetMassiveDeleteRowResult dsRowResult = dsMassiveDelete.GetRowResultInfo(rowNumber);

            while (dsRowResult != null)
            {
                Assert.AreEqual(customerMD, dsRowResult.CustomerName, $"Le nom du client n'est pas égal à {customerMD} pour la ligne de résultat {rowNumber}");
                rowNumber++;
                dsRowResult = dsMassiveDelete.GetRowResultInfo(rowNumber);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Massivedelete_pagination()
        {
            //connect as admin + go to homePage
            var homePage = LogInAsAdmin();

            DatasheetPage dataSheetPage = homePage.GoToMenus_Datasheet();
            DatasheetMassiveDeletePopup dsMassiveDelete = dataSheetPage.OpenMassiveDeletePopup();
            dsMassiveDelete.ClickOnSearch();

            int tot = dsMassiveDelete.CheckTotalNumberDataSheet();

            dsMassiveDelete.SetPageSize("16");

            bool isPageSizeEqualsTo16 = dsMassiveDelete.IsPageSizeEqualsTo16();
            Assert.IsTrue(isPageSizeEqualsTo16, "Paggination ne fonctionne pas.");
            Assert.AreEqual(dsMassiveDelete.GetTotalRowsForPagination(), tot >= 16 ? 16 : tot, "Paggination ne fonctionne pas.");

            dsMassiveDelete.SetPageSize("30");

            bool isPageSizeEqualsTo30 = dsMassiveDelete.IsPageSizeEqualsTo30();
            Assert.IsTrue(isPageSizeEqualsTo30, "Paggination ne fonctionne pas.");
            Assert.AreEqual(dsMassiveDelete.GetTotalRowsForPagination(), tot >= 30 ? 30 : tot, "Paggination ne fonctionne pas.");

            dsMassiveDelete.SetPageSize("50");

            bool isPageSizeEqualsTo50 = dsMassiveDelete.IsPageSizeEqualsTo50();
            Assert.IsTrue(isPageSizeEqualsTo50, "Paggination ne fonctionne pas.");
            Assert.AreEqual(dsMassiveDelete.GetTotalRowsForPagination(), tot >= 50 ? 50 : tot, "Paggination ne fonctionne pas.");

            dsMassiveDelete.SetPageSize("100");

            bool isPageSizeEqualsTo100 = dsMassiveDelete.IsPageSizeEqualsTo100();
            Assert.IsTrue(isPageSizeEqualsTo100, "Paggination ne fonctionne pas.");
            Assert.AreEqual(dsMassiveDelete.GetTotalRowsForPagination(), tot >= 100 ? 100 : tot, "Paggination ne fonctionne pas.");

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Massivedelete_organizebyname()
        {
            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var dataSheetPage = homePage.GoToMenus_Datasheet();
            var dsMassiveDelete = dataSheetPage.OpenMassiveDeletePopup();
            dsMassiveDelete.ClickOnSearch();


            dsMassiveDelete.ClickDatasheetNameButton();
            dsMassiveDelete.ClickDatasheetNameButton();

            bool isSortByDatasheetName = dsMassiveDelete.VerifySortByDatasheetName();
            // Assert
            Assert.IsTrue(isSortByDatasheetName, " Les delivery ne sont pas ordonnés par nom");

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Massivedelete_organizebysite()
        {
            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var dataSheetPage = homePage.GoToMenus_Datasheet();
            var dsMassiveDelete = dataSheetPage.OpenMassiveDeletePopup();

            dsMassiveDelete.ClickOnSearch();
            dsMassiveDelete.ClickSiteButton();

            bool isSortByDatasheetSite = dsMassiveDelete.VerifySortByDatasheetSite();
            // Assert
            Assert.IsTrue(isSortByDatasheetSite, " Les delivery ne sont pas ordonnés par site");

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Massivedelete_organizebycustomer()
        {
            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            DatasheetPage datasheetPage = homePage.GoToMenus_Datasheet();
            DatasheetMassiveDeletePopup dsMassiveDelete = datasheetPage.OpenMassiveDeletePopup();

            dsMassiveDelete.ClickOnSearch();

            dsMassiveDelete.ClickOnCustomerHeader();
            var firstComparison = dsMassiveDelete.VerifyCustomerCodeSort(SortType.Ascending);

            dsMassiveDelete.ClickOnCustomerHeader();
            var secondComparison = dsMassiveDelete.VerifyCustomerCodeSort(SortType.Descending);

            // Assert
            Assert.IsTrue(firstComparison, " Les Customers ne sont pas ordonnés ( de A à Z )");
            Assert.IsTrue(secondComparison, " Les Customers ne sont pas ordonnés ( de Z à A )");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Massivedelete_organizebyusecase()
        {
            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            DatasheetPage datasheetPage = homePage.GoToMenus_Datasheet();
            DatasheetMassiveDeletePopup dsMassiveDelete = datasheetPage.OpenMassiveDeletePopup();

            dsMassiveDelete.ClickOnSearch();

            dsMassiveDelete.ClickOnUseCaseHeader();
            var firstComparison = dsMassiveDelete.VerifyUseCaseSort(SortType.Ascending);

            dsMassiveDelete.ClickOnUseCaseHeader();
            var secondComparison = dsMassiveDelete.VerifyUseCaseSort(SortType.Descending);

            // Assert
            Assert.IsTrue(firstComparison, "Les usecases ne sont pas ordonnés par ordre croissant.");
            Assert.IsTrue(secondComparison, "Les usecases ne sont pas ordonnés par ordre décroissant.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Massivedelete_organizebyguesttype()
        {
            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            DatasheetPage datasheetPage = homePage.GoToMenus_Datasheet();
            DatasheetMassiveDeletePopup dsMassiveDelete = datasheetPage.OpenMassiveDeletePopup();

            dsMassiveDelete.ClickOnSearch();

            dsMassiveDelete.ClickOnGuestTypeHeader();
            bool firstComparison = dsMassiveDelete.VerifyGuestTypeSort(SortType.Ascending);

            dsMassiveDelete.ClickOnGuestTypeHeader();
            bool secondComparison = dsMassiveDelete.VerifyGuestTypeSort(SortType.Descending);

            // Assert
            Assert.IsTrue(firstComparison, " Les GuestTypes ne sont pas ordonnés ( de A à Z )");
            Assert.IsTrue(secondComparison, " Les GuestTypes ne sont pas ordonnés ( de Z à A )");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Massivedelete_selectall()
        {
            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            DatasheetPage datasheetPage = homePage.GoToMenus_Datasheet();
            DatasheetMassiveDeletePopup dsMassiveDelete = datasheetPage.OpenMassiveDeletePopup();
            dsMassiveDelete.ClickOnSearch();

            int dsCount = dsMassiveDelete.CheckTotalNumberDataSheet();
            bool areAllRowsSelected = dsMassiveDelete.VerifyRowsAreSelected();

            //Assert
            Assert.IsTrue(dsCount > 0, "L'affichage du nombre de Selected datasheet ne s'est pas mis à jour");
            Assert.IsTrue(areAllRowsSelected, "Tous les résultats ne sont pas cochés");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Massivedelete_link()
        {
            //Home page
            HomePage homePage = LogInAsAdmin();

            DatasheetPage dataSheetPage = homePage.GoToMenus_Datasheet();
            DatasheetMassiveDeletePopup dsMassiveDelete = dataSheetPage.OpenMassiveDeletePopup();
            dsMassiveDelete.ClickOnSearch();

            DatasheetMassiveDeleteRowResult dsRowResult = dsMassiveDelete.GetRowResultInfo(0);
            Assert.IsNotNull(dsRowResult, "Aucun résultat trouvé pour une recherche simple de massive delete.");

            bool linkWasClicked = dsMassiveDelete.ClickOnResultLinkAndCheckDatasheetName(0);
            Assert.IsTrue(linkWasClicked, "Lien du résultat vers le datasheet non opérationnel.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Massivedelete_namesearch()
        {
            string dsName = "NESTEA 33 CL";

            //Home page
            HomePage homePage = LogInAsAdmin();

            DatasheetPage dataSheetPage = homePage.GoToMenus_Datasheet();
            DatasheetMassiveDeletePopup dsMassiveDelete = dataSheetPage.OpenMassiveDeletePopup();

            dsMassiveDelete.SetDatasheetName(dsName);
            dsMassiveDelete.ClickOnSearch();

            int rowNumber = 0;
            DatasheetMassiveDeleteRowResult dsRowResult = dsMassiveDelete.GetRowResultInfo(rowNumber);

            Assert.AreEqual(dsName, dsRowResult.DatasheetName, $"Le nom du datasheet n'est pas égal à {dsName} pour la ligne de résultat {rowNumber}");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Massivedelete_guesttype()
        {
            string gtName = TestContext.Properties["DatasheetGuestType"].ToString();

            //Home page
            var homePage = LogInAsAdmin();

            var dataSheetPage = homePage.GoToMenus_Datasheet();
            var dsMassiveDelete = dataSheetPage.OpenMassiveDeletePopup();

            dsMassiveDelete.SelectGuestTypeByName(gtName, true);
            dsMassiveDelete.ClickOnSearch();

            int rowNumber = 0;
            DatasheetMassiveDeleteRowResult dsRowResult = dsMassiveDelete.GetRowResultInfo(rowNumber);

            while (dsRowResult != null)
            {
                Assert.AreEqual(gtName, dsRowResult.GuestTypeName, $"Le nom du guest type n'est pas égal à {gtName} pour la ligne de résultat {rowNumber}");
                rowNumber++;
                dsRowResult = dsMassiveDelete.GetRowResultInfo(rowNumber);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Massivedelete_inactivesitessearch()
        {
            //Home page
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var dataSheetPage = homePage.GoToMenus_Datasheet();
            var dsMassiveDelete = dataSheetPage.OpenMassiveDeletePopup();

            dsMassiveDelete.ClickOnInactiveSiteCheck();
            dsMassiveDelete.SelectAllInactiveSites();
            dsMassiveDelete.ClickOnSearch();

            int rowNumber = 0;
            DatasheetMassiveDeleteRowResult dsRowResult = dsMassiveDelete.GetRowResultInfo(rowNumber);

            while (dsRowResult != null)
            {
                Assert.IsTrue(dsRowResult.IsCustomerInactive, $"Le datasheet {dsRowResult.DatasheetName} est avec un site actif alors que le filtre est sur 'site inactive'.");
                rowNumber++;
                dsRowResult = dsMassiveDelete.GetRowResultInfo(rowNumber);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Massivedelete_inactivecustomersearch()
        {
            //Home page
            HomePage homePage = LogInAsAdmin();

            DatasheetPage dataSheetPage = homePage.GoToMenus_Datasheet();
            DatasheetMassiveDeletePopup dsMassiveDelete = dataSheetPage.OpenMassiveDeletePopup();

            dsMassiveDelete.ClickOnInactiveCustomerCheck();
            dsMassiveDelete.SelectAllInactiveCustomers();
            dsMassiveDelete.ClickOnSearch();

            int rowNumber = 0;
            DatasheetMassiveDeleteRowResult dsRowResult = dsMassiveDelete.GetRowResultInfo(rowNumber);

            while (dsRowResult != null)
            {
                Assert.IsTrue(dsRowResult.IsCustomerInactive, $"Le datasheet {dsRowResult.DatasheetName} est avec un customer actif alors que le filtre est sur 'customer inactive'.");
                rowNumber++;
                dsRowResult = dsMassiveDelete.GetRowResultInfo(rowNumber);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Massivedelete_status_inactivesites()
        {
            //Home page
            var homePage = LogInAsAdmin();

            var dataSheetPage = homePage.GoToMenus_Datasheet();
            var dsMassiveDelete = dataSheetPage.OpenMassiveDeletePopup();

            dsMassiveDelete.SelectStatus(MassiveDeleteStatus.OnlyInactiveSites.ToString());
            dsMassiveDelete.ClickOnSearch();
            dsMassiveDelete.PageSizeEqualsTo100();

            int rowNumber = 0;
            DatasheetMassiveDeleteRowResult dsRowResult = dsMassiveDelete.GetRowResultInfo(rowNumber);

            while (dsRowResult != null)
            {
                Assert.IsTrue(dsRowResult.IsSiteInactive, $"Le datasheet {dsRowResult.DatasheetName} est avec un site actif alors que le filtre est sur 'site inactive'.");
                rowNumber++;
                dsRowResult = dsMassiveDelete.GetRowResultInfo(rowNumber);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Massivedelete_status_inactivecustomers()
        {
            //Home page
            var homePage = LogInAsAdmin();

            var dataSheetPage = homePage.GoToMenus_Datasheet();
            var dsMassiveDelete = dataSheetPage.OpenMassiveDeletePopup();

            dsMassiveDelete.SelectStatus(MassiveDeleteStatus.OnlyInactiveCustomers.ToString());
            dsMassiveDelete.ClickOnSearch();
            dsMassiveDelete.PageSizeEqualsTo100();

            int rowNumber = 0;
            DatasheetMassiveDeleteRowResult dsRowResult = dsMassiveDelete.GetRowResultInfo(rowNumber);

            while (dsRowResult != null)
            {
                Assert.IsTrue(dsRowResult.IsCustomerInactive, $"Le datasheet {dsRowResult.DatasheetName} est avec un customer actif alors que le filtre est sur 'customer inactive'.");
                rowNumber++;
                dsRowResult = dsMassiveDelete.GetRowResultInfo(rowNumber);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Massivedelete_status_activedatasheet()
        {
            //Home page
            var homePage = LogInAsAdmin();

            var dataSheetPage = homePage.GoToMenus_Datasheet();
            var dsMassiveDelete = dataSheetPage.OpenMassiveDeletePopup();

            dsMassiveDelete.SelectStatus(MassiveDeleteStatus.ActiveDatasheets.ToString());
            dsMassiveDelete.ClickOnSearch();

            int rowNumber = 0;
            DatasheetMassiveDeleteRowResult dsRowResult = dsMassiveDelete.GetRowResultInfo(rowNumber);

            while (dsRowResult != null)
            {
                Assert.IsFalse(dsRowResult.IsDatasheetInactive, $"Le datasheet {dsRowResult.DatasheetName} apparait inactif alors que le filtre est sur 'active datasheets'.");
                rowNumber++;
                dsRowResult = dsMassiveDelete.GetRowResultInfo(rowNumber);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Massivedelete_status_inactivedatasheet()
        {
            //Home page
            var homePage = LogInAsAdmin();

            var dataSheetPage = homePage.GoToMenus_Datasheet();
            var dsMassiveDelete = dataSheetPage.OpenMassiveDeletePopup();

            dsMassiveDelete.SelectStatus(MassiveDeleteStatus.InactiveDatasheets.ToString());
            dsMassiveDelete.ClickOnSearch();

            int rowNumber = 0;
            DatasheetMassiveDeleteRowResult dsRowResult = dsMassiveDelete.GetRowResultInfo(rowNumber);

            while (dsRowResult != null)
            {
                Assert.IsTrue(dsRowResult.IsDatasheetInactive, $"Le datasheet {dsRowResult.DatasheetName} apparait actif alors que le filtre est sur 'datasheet inactive'.");
                rowNumber++;
                dsRowResult = dsMassiveDelete.GetRowResultInfo(rowNumber);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Massivedelete_status_unused()
        {
            //Home page
            var homePage = LogInAsAdmin();

            var dataSheetPage = homePage.GoToMenus_Datasheet();
            var dsMassiveDelete = dataSheetPage.OpenMassiveDeletePopup();

            dsMassiveDelete.SelectStatus(MassiveDeleteStatus.Unused.ToString());
            dsMassiveDelete.ClickOnSearch();

            int rowNumber = 0;
            DatasheetMassiveDeleteRowResult dsRowResult = dsMassiveDelete.GetRowResultInfo(rowNumber);

            while (dsRowResult != null)
            {
                bool isUnused = string.IsNullOrEmpty(dsRowResult.CustomerName) && dsRowResult.UseCase == 0;
                Assert.IsTrue(isUnused, $"Le datasheet {dsRowResult.DatasheetName} apparait comme utilisé alors que le filtre est sur 'Inutilisé'.");
                rowNumber++;
                dsRowResult = dsMassiveDelete.GetRowResultInfo(rowNumber);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Massivedelete_status_used()
        {
            //Home page
            var homePage = LogInAsAdmin();

            var dataSheetPage = homePage.GoToMenus_Datasheet();
            var dsMassiveDelete = dataSheetPage.OpenMassiveDeletePopup();

            dsMassiveDelete.SelectStatus(MassiveDeleteStatus.Used.ToString());
            dsMassiveDelete.ClickOnSearch();

            int rowNumber = 0;
            DatasheetMassiveDeleteRowResult dsRowResult = dsMassiveDelete.GetRowResultInfo(rowNumber);

            while (dsRowResult != null)
            {
                bool isUnused = string.IsNullOrEmpty(dsRowResult.CustomerName) == false && dsRowResult.UseCase != 0;
                Assert.IsTrue(isUnused, $"Le datasheet {dsRowResult.DatasheetName} apparait comme inutilisé alors que le filtre est sur 'Utilisé'.");
                rowNumber++;
                dsRowResult = dsMassiveDelete.GetRowResultInfo(rowNumber);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Massivedelete_status_unpurchasable()
        {
            //Home page
            var homePage = LogInAsAdmin();

            var dataSheetPage = homePage.GoToMenus_Datasheet();
            var dsMassiveDelete = dataSheetPage.OpenMassiveDeletePopup();

            dsMassiveDelete.SelectStatus(MassiveDeleteStatus.WithUnpurchasableItems.ToString());
            dsMassiveDelete.ClickOnSearch();

            int rowNumber = 0;
            DatasheetMassiveDeleteRowResult dsRowResult = dsMassiveDelete.GetRowResultInfo(rowNumber);

            Assert.IsTrue(dsRowResult != null, $"Aucun résultat trouvé avec le filtre 'With items unpurchasable'.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_RecalculAllergen_AddOrDeleteAllergen()
        {
            //Prepare recipe


            //prepare items 
            string itemName1 = "itemIntoleanceRecipe";
            int nbreAllItem1 = 0;
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();

            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            Random rnd = new Random();
            string datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();
            string subrecipeName = "RecipeTest-" + rnd.Next().ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString() + "-" + rnd.Next().ToString();
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // test 1: ajout des allergènes à un item, recalculer les allergènes des recettes, des sous recettes et des datasheets

            //add 1 item (a 3 allergens)
            var itemPage = homePage.GoToPurchasing_ItemPage();
            // ----------------------------------------- Item 1 ---------------------------------------------------------------
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName1);

            if (itemPage.CheckTotalNumber() == 0)
            {
                ItemCreateModalPage itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage1 = itemCreateModalPage.FillField_CreateNewItem(itemName1, group, workshop, taxType, prodUnit);
                //add 3 allergens
                ItemIntolerancePage intolerancePage = itemGeneralInformationPage1.ClickOnIntolerancePage();
                intolerancePage.AddAllergen("Z-Halal");
                intolerancePage.AddAllergen("Gluten/Gluten");
                intolerancePage.AddAllergen("Altramuces/Lupin");
                nbreAllItem1 = intolerancePage.CalculateAllergens();
                itemPage = itemGeneralInformationPage1.BackToList();
                itemPage.Filter(ItemPage.FilterType.Search, itemName1);
            }
            else
            {
                string item1 = itemPage.GetFirstItemName();
                ItemGeneralInformationPage itemGeneralInformationPage1 = itemPage.ClickOnFirstItem();
                ItemIntolerancePage intolerancePage = itemGeneralInformationPage1.ClickOnIntolerancePage();
                intolerancePage.UncheckAllAllergen();
                intolerancePage.AddAllergen("Z-Halal");
                intolerancePage.AddAllergen("Gluten/Gluten");
                intolerancePage.AddAllergen("Altramuces/Lupin");
                nbreAllItem1 = intolerancePage.CalculateAllergens();
                itemGeneralInformationPage1.BackToList();
            }

            //go to create datasheet
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            DatasheetDetailsPage datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);
            Assert.IsFalse(datasheetDetailPage.IsRecipeAdded(), "La datasheet possède déjà une recette.");

            //add recipe avec un ingredient 
            var recipesCreateModalPage = datasheetDetailPage.CreateNewRecipe();
            recipesCreateModalPage.FillFields_AddNewRecipeToDatasheet(recipeName, itemName1);
            datasheetDetailPage.ClickOnRecipesTab();
            Assert.IsTrue(datasheetDetailPage.IsRecipeAdded(), "La recette n'a pas été rajoutée à la datasheet.");
            //calculer allergens sur la recipe et verifier l'égalité
            RecipesVariantPage recipesVariantPage = datasheetDetailPage.EditRecipe();
            RecipeIntolerancePage recipeIntolerancePage = recipesVariantPage.ClickOnIntoleranceTab();
            var nbreAllergensRecipeBefore = recipeIntolerancePage.CalculateAllergens();
            Assert.AreEqual(nbreAllItem1, nbreAllergensRecipeBefore, "les allergènes ne sont pas identiques entre le item " + itemName1 + " et la recette " + recipeName);
            recipesVariantPage.CloseWindow();
            //calculer allergens sur la datasheet et verifier l'égalité
            DatasheetIntolerancePage datasheetIntolerancePage = datasheetDetailPage.ClickOnIntoleranceTab();
            var nbreAllergensDatasheetBefore = datasheetIntolerancePage.CalculateAllergens();
            Assert.AreEqual(nbreAllItem1, nbreAllergensDatasheetBefore, "les allergènes ne sont pas identiques entre le item " + itemName1 + " et la datasheet " + datasheetName);
            //add sub-recipe avec un ingredient 
            datasheetDetailPage.ClickOnRecipesTab();
            var datasheetCreateNewRecipePage = datasheetDetailPage.AddSubrecipeForFirstRecipe();
            datasheetDetailPage = datasheetCreateNewRecipePage.FillFields_AddNewRecipeToDatasheet(subrecipeName, itemName1);
            if (!datasheetDetailPage.isPopupClosed())
            {
                datasheetDetailPage.CloseDetailModal();
            }
            datasheetDetailPage.ClickOnRecipesTab();
            datasheetDetailPage.ShowDetails();
            Assert.IsTrue(datasheetDetailPage.IsRecipeDetailsVisible(), "Le détail des recettes n'est pas visible.");
            Assert.IsTrue(datasheetDetailPage.IsSubrecipeAdded(), "La sous-recette n'a pas été ajoutée à la recette.");

            //calculer allergens sur la sub-recipe et verifier l'égalité 
            recipesVariantPage = datasheetDetailPage.EditRecipe();
            recipesVariantPage.EditFirstSubRecipeDatasheet();
            RecipeIntolerancePage recipeIntolerancePage1 = recipesVariantPage.ClickOnIntoleranceTab();
            var nbreAllergensSub_RecipeBefore = recipeIntolerancePage1.CalculateAllergens();
            Assert.AreEqual(nbreAllItem1, nbreAllergensSub_RecipeBefore, "les allergènes ne sont pas identiques entre le item " + itemName1 + " et la sous-recette " + subrecipeName);

            //ajouter un allergène dans item1
            itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName1);
            var itemGeneralInformationPage = itemPage.ClickOnFirstItem();
            var itemIntolerancePage = itemGeneralInformationPage.ClickOnIntolerancePage();
            if (!itemIntolerancePage.IsAllergenChecked("Apio/Celery"))
            {
                itemIntolerancePage.AddAllergen("Apio/Celery");
            }
            if (!itemIntolerancePage.IsAllergenChecked("Leche/Milk"))
            {
                itemIntolerancePage.AddAllergen("Leche/Milk");
            }
            var nbreAllergensItem1After = itemIntolerancePage.CalculateAllergens();
            itemGeneralInformationPage.BackToList();

            //verify la recalculate dans la datasheet
            datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
            datasheetDetailPage = datasheetPage.SelectFirstDatasheet();
            datasheetIntolerancePage = datasheetDetailPage.ClickOnIntoleranceTab();
            var nbreAllergensDatasheetAfter = datasheetIntolerancePage.CalculateAllergens();
            Assert.AreEqual(nbreAllergensItem1After, nbreAllergensDatasheetAfter, "les allergènes ne sont pas identiques entre la datasheet et le item après l'ajout.");
            Assert.AreNotEqual(nbreAllergensDatasheetBefore, nbreAllergensDatasheetAfter, "il faut avoir les allergènes pas identiques dans la datasheet après l'ajout d'un allergen.");

            //verify la recalculate dans la recipe 
            datasheetDetailPage.ClickOnRecipesTab();
            recipesVariantPage = datasheetDetailPage.EditRecipe();
            recipeIntolerancePage = recipesVariantPage.ClickOnIntoleranceTab();
            var nbreAllergensRecipeAfter = recipeIntolerancePage.CalculateAllergens();
            Assert.AreEqual(nbreAllergensItem1After, nbreAllergensRecipeAfter, "les allergènes ne sont pas identiques entre la recipe et le item après l'ajout.");
            Assert.AreNotEqual(nbreAllergensRecipeBefore, nbreAllergensRecipeAfter, "il faut avoir les allergènes pas identiques dans la recipe après l'ajout d'un allergen.");

            //verify la recalculate dans la sub-recipe 
            recipesVariantPage.ClickOnItemTab();
            recipesVariantPage.EditFirstSubRecipeDatasheet();
            recipeIntolerancePage1 = recipesVariantPage.ClickOnIntoleranceTab();
            var nbreAllergensSub_RecipeAfter = recipeIntolerancePage1.CalculateAllergens();
            Assert.AreEqual(nbreAllergensItem1After, nbreAllergensSub_RecipeAfter, "les allergènes ne sont pas identiques entre la sub-recipe et le item après l'ajout.");
            Assert.AreNotEqual(nbreAllergensRecipeBefore, nbreAllergensSub_RecipeAfter, "il faut avoir les allergènes pas identiques dans la sub-recipe après l'ajout d'un allergen.");

            //test2 : suppression des allergènes à un item, recalculer les allergènes des recettes, des sous recettes et des datasheets
            //supprimer un allergène dans item1
            itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName1);
            itemGeneralInformationPage = itemPage.ClickOnFirstItem();
            itemIntolerancePage = itemGeneralInformationPage.ClickOnIntolerancePage();
            if (itemIntolerancePage.IsAllergenChecked("Apio/Celery"))
            {
                itemIntolerancePage.UncheckAllergen("Apio/Celery");
            }
            var nbreAllergensItem1AfterDelete = itemIntolerancePage.CalculateAllergens();
            itemGeneralInformationPage.BackToList();

            datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
            datasheetDetailPage = datasheetPage.SelectFirstDatasheet();
            datasheetIntolerancePage = datasheetDetailPage.ClickOnIntoleranceTab();
            var nbreAllergensDatasheetAfterDelete = datasheetIntolerancePage.CalculateAllergens();
            //verify la recalculate dans la datasheet
            Assert.AreEqual(nbreAllergensItem1AfterDelete, nbreAllergensDatasheetAfterDelete, "les allergènes ne sont pas identiques entre la datasheet et le item la suppression.");
            Assert.AreNotEqual(nbreAllergensDatasheetAfter, nbreAllergensDatasheetAfterDelete, "il faut avoir les allergènes pas identiques dans la datasheet après la suppression d'un allergen.");

            datasheetDetailPage.ClickOnRecipesTab();
            recipesVariantPage = datasheetDetailPage.EditRecipe();
            recipeIntolerancePage = recipesVariantPage.ClickOnIntoleranceTab();
            var nbreAllergensRecipeAfterDelete = recipeIntolerancePage.CalculateAllergens();
            //verify la recalculate dans la recipe 
            Assert.AreEqual(nbreAllergensItem1AfterDelete, nbreAllergensRecipeAfterDelete, "les allergènes ne sont pas identiques entre la recipe et le item après la suppression.");
            Assert.AreNotEqual(nbreAllergensRecipeAfter, nbreAllergensRecipeAfterDelete, "il faut avoir les allergènes pas identiques dans la recipe après la suppression d'un allergen.");

            //verify recalculate in sub-recipe 
            recipesVariantPage.ClickOnItemTab();
            recipesVariantPage.EditFirstSubRecipeDatasheet();
            recipeIntolerancePage1 = recipesVariantPage.ClickOnIntoleranceTab();
            var nbreAllergensSub_RecipeAfterDelete = recipeIntolerancePage1.CalculateAllergens();
            Assert.AreEqual(nbreAllergensItem1AfterDelete, nbreAllergensSub_RecipeAfterDelete, "les allergènes ne sont pas identiques entre la sub-recipe et le item après la suppression.");
            Assert.AreNotEqual(nbreAllergensSub_RecipeAfter, nbreAllergensSub_RecipeAfterDelete, "il faut avoir les allergènes pas identiques dans la sub-recipe après la suppression d'un allergen.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_RecalculAllergen_AddOrDeleteItem()
        {
            //Prepare recipe
            string recipeName = "RecipeForItem1Datasheet";

            //prepare items 
            string itemName1 = "itemIntoleanceRecipe";
            string itemName2 = "itemProdman1";
            int nbreAllItem1 = 0;
            int nbreAllItem2 = 0;
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            List<string> listAllergensSelected1 = new List<string>();
            List<string> listAllergensSelected2 = new List<string>();
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            Random rnd = new Random();
            string datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();
            string subrecipeName = "RecipeTest-" + rnd.Next().ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // test 1: ajouter un item, recalculer les allergènes des recettes, des sous recettes et des datasheets

            //add 2 items (ont 3 allergens)
            var itemPage = homePage.GoToPurchasing_ItemPage();
            // ----------------------------------------- Item 1 ---------------------------------------------------------------
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName1);

            if (itemPage.CheckTotalNumber() == 0)
            {
                ItemCreateModalPage itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage1 = itemCreateModalPage.FillField_CreateNewItem(itemName1, group, workshop, taxType, prodUnit);
                //add 3 allergens
                var intolerancePage = itemGeneralInformationPage1.ClickOnIntolerancePage();
                intolerancePage.AddAllergen("Z-Halal");
                intolerancePage.AddAllergen("Gluten/Gluten");
                intolerancePage.AddAllergen("Altramuces/Lupin");
                nbreAllItem1 = intolerancePage.CalculateAllergens();
                listAllergensSelected1 = intolerancePage.GetListAllergensSelected();
                itemPage = itemGeneralInformationPage1.BackToList();
                itemPage.Filter(ItemPage.FilterType.Search, itemName1);
            }
            else
            {
                string item1 = itemPage.GetFirstItemName();
                ItemGeneralInformationPage itemGeneralInformationPage1 = itemPage.ClickOnFirstItem();
                ItemIntolerancePage intolerancePage = itemGeneralInformationPage1.ClickOnIntolerancePage();
                intolerancePage.UncheckAllAllergen();
                intolerancePage.AddAllergen("Z-Halal");
                intolerancePage.AddAllergen("Gluten/Gluten");
                intolerancePage.AddAllergen("Altramuces/Lupin");
                nbreAllItem1 = intolerancePage.CalculateAllergens();
                listAllergensSelected1 = intolerancePage.GetListAllergensSelected();
                itemGeneralInformationPage1.BackToList();
            }

            // ----------------------------------------- Item 2 ---------------------------------------------------------------
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName2);

            if (itemPage.CheckTotalNumber() == 0)
            {
                ItemCreateModalPage itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage2 = itemCreateModalPage.FillField_CreateNewItem(itemName2, group, workshop, taxType, prodUnit);
                //add 3 allergens
                var intolerancePage = itemGeneralInformationPage2.ClickOnIntolerancePage();
                intolerancePage.AddAllergen("Z-Halal");
                intolerancePage.AddAllergen("Gluten/Gluten");
                intolerancePage.AddAllergen("Altramuces/Lupin");
                nbreAllItem2 = intolerancePage.CalculateAllergens();
                listAllergensSelected2 = intolerancePage.GetListAllergensSelected();
                itemPage = itemGeneralInformationPage2.BackToList();
                itemPage.Filter(ItemPage.FilterType.Search, itemName2);
            }
            else
            {
                string item2 = itemPage.GetFirstItemName();
                var itemGeneralInformationPage2 = itemPage.ClickOnFirstItem();
                ItemIntolerancePage itemIntolerancePage2 = itemGeneralInformationPage2.ClickOnIntolerancePage();
                if (!itemIntolerancePage2.IsAllergenChecked("Frutos de cáscara/Nuts"))
                {
                    itemIntolerancePage2.AddAllergen("Frutos de cáscara/Nuts");
                }
                if (!itemIntolerancePage2.IsAllergenChecked("Huevos/Eggs"))
                {
                    itemIntolerancePage2.AddAllergen("Huevos/Eggs");
                }
                nbreAllItem2 = itemIntolerancePage2.CalculateAllergens();
                listAllergensSelected2 = itemIntolerancePage2.GetListAllergensSelected();
                itemGeneralInformationPage2.BackToList();
            }
            //calculer les allergènes communs entre 2 items
            var nbreAllergen = itemPage.SUMAllergensCommunes2Items(listAllergensSelected1, listAllergensSelected2);

            //go to create datasheet
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            DatasheetDetailsPage datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);
            Assert.IsFalse(datasheetDetailPage.IsRecipeAdded(), "La datasheet possède déjà une recette.");

            //add recipe avec un ingredient "item1"
            if (!datasheetDetailPage.AddRecipeError(recipeName))
            {
                var recipesCreateModalPage = datasheetDetailPage.CreateNewRecipe();
                recipesCreateModalPage.FillFields_AddNewRecipeToDatasheet(recipeName, itemName1);
            }
            else
            {
                datasheetDetailPage.AddRecipe(recipeName);
            }
            datasheetDetailPage.ClickOnRecipesTab();
            Assert.IsTrue(datasheetDetailPage.IsRecipeAdded(), "La recette n'a pas été rajoutée à la datasheet.");
            //calculer allergens sur la recipe et verifier l'égalité
            RecipesVariantPage recipesVariantPage = datasheetDetailPage.EditRecipe();
            RecipeIntolerancePage recipeIntolerancePage = recipesVariantPage.ClickOnIntoleranceTab();
            var nbreAllergensRecipeBefore = recipeIntolerancePage.CalculateAllergens();
            Assert.AreEqual(nbreAllItem1, nbreAllergensRecipeBefore, "les allergènes ne sont pas identiques entre le item " + itemName1 + " et la recette " + recipeName);
            recipesVariantPage.CloseWindow();
            //calculer allergens sur la datasheet et verifier l'égalité
            DatasheetIntolerancePage datasheetIntolerancePage = datasheetDetailPage.ClickOnIntoleranceTab();
            var nbreAllergensDatasheetBefore = datasheetIntolerancePage.CalculateAllergens();
            Assert.AreEqual(nbreAllItem1, nbreAllergensDatasheetBefore, "les allergènes ne sont pas identiques entre le item " + itemName1 + " et la datasheet " + datasheetName);
            //add sub-recipe avec un ingredient "item1"
            datasheetDetailPage.ClickOnRecipesTab();
            var datasheetCreateNewRecipePage = datasheetDetailPage.AddSubrecipeForFirstRecipe();
            datasheetDetailPage = datasheetCreateNewRecipePage.FillFields_AddNewRecipeToDatasheet(subrecipeName, itemName1);
            if (!datasheetDetailPage.isPopupClosed())
            {
                datasheetDetailPage.CloseDetailModal();
            }
            datasheetDetailPage.ClickOnRecipesTab();
            datasheetDetailPage.ShowDetails();
            Assert.IsTrue(datasheetDetailPage.IsRecipeDetailsVisible(), "Le détail des recettes n'est pas visible.");
            Assert.IsTrue(datasheetDetailPage.IsSubrecipeAdded(), "La sous-recette n'a pas été ajoutée à la recette.");

            //calculer allergens sur la sub-recipe et verifier l'égalité 
            datasheetDetailPage.EditRecipe();
            recipesVariantPage.EditFirstSubRecipeDatasheet();
            RecipeIntolerancePage recipeIntolerancePage1 = recipesVariantPage.ClickOnIntoleranceTab();
            var nbreAllergensSub_RecipeBefore = recipeIntolerancePage1.CalculateAllergens();
            Assert.AreEqual(nbreAllItem1, nbreAllergensSub_RecipeBefore, "les allergènes ne sont pas identiques entre le item " + itemName1 + " et la sous-recette " + subrecipeName);


            //ajouter un deuxième item et recalculer
            recipesVariantPage.CloseSubRecipePopUp();
            recipesVariantPage.AjouterIngredient(itemName2);
            recipeIntolerancePage1 = recipesVariantPage.ClickOnIntoleranceTab();
            var nbreAllergensSub_RecipeAfter = recipeIntolerancePage1.CalculateAllergens();
            Assert.AreEqual(nbreAllergen, nbreAllergensSub_RecipeAfter, "les allergènes ne sont pas identiques entre la sub-recipe et la somme des 2 item après l'ajout d'un item.");
            Assert.AreNotEqual(nbreAllergensSub_RecipeBefore, nbreAllergensSub_RecipeAfter, "il faut avoir les allergènes pas identiques dans la sub-recipe après l'ajout d'un item.");

            datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
            datasheetDetailPage = datasheetPage.SelectFirstDatasheet();
            datasheetIntolerancePage = datasheetDetailPage.ClickOnIntoleranceTab();
            var nbreAllergensDatasheetAfter = datasheetIntolerancePage.CalculateAllergens();
            //verify la recalculate dans la datasheet
            Assert.AreEqual(nbreAllergen, nbreAllergensDatasheetAfter, "les allergènes ne sont pas identiques entre la datasheet et la somme des 2 item après l'ajout d'un item.");
            Assert.AreNotEqual(nbreAllergensDatasheetBefore, nbreAllergensDatasheetAfter, "il faut avoir les allergènes pas identiques dans la datasheet après l'ajout d'un item.");

            datasheetDetailPage.ClickOnRecipesTab();
            recipesVariantPage = datasheetDetailPage.EditRecipe();
            recipeIntolerancePage = recipesVariantPage.ClickOnIntoleranceTab();
            var nbreAllergensRecipeAfter = recipeIntolerancePage.CalculateAllergens();
            //verify la recalculate dans la recipe 
            Assert.AreEqual(nbreAllergen, nbreAllergensRecipeAfter, "les allergènes ne sont pas identiques entre la recipe et la somme des 2 item après l'ajout d'un item.");
            Assert.AreNotEqual(nbreAllergensRecipeBefore, nbreAllergensRecipeAfter, "il faut avoir les allergènes pas identiques dans la recipe après l'ajout d'un item.");

            // test 2: supprimer un item, recalculer les allergènes des recettes, des sous recettes et des datasheets
            recipesVariantPage.ClickOnItemTab();
            recipesVariantPage.DeleteIngredient(itemName2);
            recipeIntolerancePage = recipesVariantPage.ClickOnIntoleranceTab();
            var nbreAllergensRecipeAfterDelete = recipeIntolerancePage.CalculateAllergens();
            //verify la recalculate dans la recipe 
            Assert.AreEqual(nbreAllItem1, nbreAllergensRecipeAfterDelete, "les allergènes ne sont pas identiques entre la recipe et la somme des 2 item après la suppression d'un item.");
            Assert.AreNotEqual(nbreAllergensRecipeAfter, nbreAllergensRecipeAfterDelete, "il faut avoir les allergènes pas identiques dans la recipe après la suppression d'un item.");

            recipesVariantPage.ClickOnItemTab();
            recipesVariantPage.EditFirstSubRecipeDatasheet();
            recipeIntolerancePage = recipesVariantPage.ClickOnIntoleranceTab();
            var nbreAllergensSub_RecipeAfterDelete = recipeIntolerancePage.CalculateAllergens();
            Assert.AreEqual(nbreAllItem1, nbreAllergensSub_RecipeAfterDelete, "les allergènes ne sont pas identiques entre la sub-recipe et la somme des 2 item après la suppression d'un item.");
            Assert.AreNotEqual(nbreAllergensSub_RecipeAfter, nbreAllergensSub_RecipeAfterDelete, "il faut avoir les allergènes pas identiques dans la sub-recipe après la suppression d'un item.");

            datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
            datasheetDetailPage = datasheetPage.SelectFirstDatasheet();
            datasheetIntolerancePage = datasheetDetailPage.ClickOnIntoleranceTab();
            var nbreAllergensDatasheetAfterDelete = datasheetIntolerancePage.CalculateAllergens();
            //verify la recalculate dans la datasheet
            Assert.AreEqual(nbreAllItem1, nbreAllergensDatasheetAfterDelete, "les allergènes ne sont pas identiques entre la datasheet et la somme des 2 item après la suppression d'un item.");
            Assert.AreNotEqual(nbreAllergensDatasheetAfter, nbreAllergensDatasheetAfterDelete, "il faut avoir les allergènes pas identiques dans la datasheet après la suppression d'un item.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_ReplicateSubRecipe()
        {
            //Prepare recipe
            string recipeName = "RecipeForItemDatasheet";

            //prepare items 
            string netWeight = "20";
            //Prepare
            string site1 = TestContext.Properties["SiteACE"].ToString();
            string site2 = TestContext.Properties["SiteBCN"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            Random rnd = new Random();
            string datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();
            string subrecipeName = "RecipeTest-" + rnd.Next().ToString();
            string itemName1 = "Chicken";
            string currency = TestContext.Properties["Currency"].ToString();
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();
            //go to create datasheet
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            DatasheetDetailsPage datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheetWidh2Sites(datasheetName, guestType, site1, site2);
            var isRecipeAdded = datasheetDetailPage.IsRecipeAdded();
            Assert.IsFalse(isRecipeAdded, "La datasheet possède déjà une recette.");

            //add recipe with ingredient 

            var recipesCreateModalPage = datasheetDetailPage.CreateNewRecipe();
            recipesCreateModalPage.FillFields_AddNewRecipeToDatasheet(recipeName);
            datasheetDetailPage.ClickOnRecipesTab();
            isRecipeAdded = datasheetDetailPage.IsRecipeAdded();
            Assert.IsTrue(isRecipeAdded, "La recette n'a pas été rajoutée à la datasheet.");

            //add sub-recipe with ingredient
            var datasheetCreateNewRecipePage = datasheetDetailPage.AddSubrecipeForFirstRecipe();
            datasheetCreateNewRecipePage.FillFields_AddNewRecipeToDatasheet(subrecipeName, itemName1);
            datasheetDetailPage = datasheetCreateNewRecipePage.CloseWindow();

            datasheetDetailPage.SelectSite(site1);
            RecipesVariantPage recipesVariantPage = datasheetDetailPage.EditRecipe();
            recipesVariantPage.ClickOnFirstIngredientDatasheet();
            recipesVariantPage.SetNetWeight(netWeight);

            var priceVarianSubRecipeSite1 = recipesVariantPage.GetPrice(decimalSeparatorValue, currency).ToString("N4", new CultureInfo("fr-FR"));
            //verify price variants sur les 2 sites pour la sub-recipe
            recipesVariantPage.CloseWindow();
            datasheetDetailPage.SelectSite(site2);
            recipesVariantPage = datasheetDetailPage.EditRecipe();
            recipesVariantPage.ClickOnFirstIngredientDatasheet();
            var priceVarianSubRecipeSite2 = recipesVariantPage.GetPrice(decimalSeparatorValue, currency).ToString("N4", new CultureInfo("fr-FR"));
            recipesVariantPage.WaitLoading();
            var netWeightSite2 = recipesVariantPage.GetNetWeight(",");
            Assert.AreEqual(priceVarianSubRecipeSite1, priceVarianSubRecipeSite2, "Les données ne sont pas compatibles dans les deux sites.");
            Assert.AreEqual(netWeight, netWeightSite2.ToString(), "Les données ne sont pas compatibles dans les deux sites.");

            datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
            datasheetDetailPage = datasheetPage.SelectFirstDatasheet();
            //verify price variants sur les 2 sites pour la recipe
            datasheetDetailPage.ClickOnRecipesTab();
            datasheetDetailPage.SelectSite(site1);
            recipesVariantPage = datasheetDetailPage.EditRecipe();
            var verifyPriceIngredient = recipesVariantPage.verifyPricesIngredients(priceVarianSubRecipeSite1);
            var verifyWeightIngeredient = recipesVariantPage.verifyNetWeightIngredients(netWeight);
            Assert.IsTrue(verifyPriceIngredient, "Les données ne sont pas compatibles.");
            Assert.IsTrue(verifyWeightIngeredient, "Les données ne sont pas compatibles.");
            datasheetDetailPage = recipesVariantPage.CloseWindow();

            datasheetDetailPage.ClickOnRecipesTab();
            datasheetDetailPage.SelectSite(site2);
            recipesVariantPage = datasheetDetailPage.EditRecipe();
            verifyPriceIngredient = recipesVariantPage.verifyPricesIngredients(priceVarianSubRecipeSite1);
            verifyWeightIngeredient = recipesVariantPage.verifyNetWeightIngredients(netWeight);
            Assert.IsTrue(verifyPriceIngredient, "Les données ne sont pas compatibles.");
            Assert.IsTrue(verifyWeightIngeredient, "Les données ne sont pas compatibles.");
            datasheetDetailPage = recipesVariantPage.CloseWindow();

            datasheetDetailPage.SelectSite(site1);
            datasheetDetailPage.ClickOnFirstRecipe();

            double paxRecette = double.Parse(priceVarianSubRecipeSite1);
            double weightRecette = netWeightSite2;
            double pax = double.Parse(datasheetDetailPage.GetPAXRecipe());
            double weight = double.Parse(datasheetDetailPage.GetNetWeight(decimalSeparatorValue).ToString());
            // € 5,2124 versus 26,0619
            // weight = 4
            // netWeightSite2 = 20

            Assert.AreEqual(pax, weight * paxRecette / weightRecette, "Les données ne sont pas compatibles dans les deux sites.");
            Assert.AreEqual(weight, weightRecette, "Les données ne sont pas compatibles dans les deux sites.");

            datasheetDetailPage.SelectSite(site2);
            datasheetDetailPage.ClickOnFirstRecipe();
            var pax2 = datasheetDetailPage.GetPAXRecipe();
            var weight2 = datasheetDetailPage.GetNetWeight(decimalSeparatorValue).ToString();
            Assert.IsTrue(("€ " + pax2).Contains(priceVarianSubRecipeSite1), "Les données ne sont pas compatibles dans les deux sites.");
            Assert.AreEqual(weight2, netWeight, "Les données ne sont pas compatibles dans les deux sites.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_ReplicateRecipe()
        {
            //Prepare recipe
            string recipeName = "RecipeForItemDatasheet";

            //prepare items 
            string netWeight = "20";
            string yield = "100";
            //Prepare
            string site1 = TestContext.Properties["SiteACE"].ToString();
            string site2 = TestContext.Properties["SiteBCN"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            Random rnd = new Random();
            string datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();
            string currency = TestContext.Properties["Currency"].ToString();

            //Home page
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            //get first available item
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Site, site1);
            itemPage.Filter(ItemPage.FilterType.Site, site2, false);
            string itemName1 = itemPage.GetFirstItemName();

            itemPage.Filter(ItemPage.FilterType.Search, "ORANGE", true, true);
            string itemName2 = itemPage.GetFirstItemName();

            //Create new datasheet
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            DatasheetDetailsPage datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheetWidh2Sites(datasheetName, guestType, site1, site2);
            Assert.IsFalse(datasheetDetailPage.IsRecipeAdded(), "La datasheet possède déjà une recette.");

            //add recipe with ingredient 
            if (!datasheetDetailPage.AddRecipeError(recipeName))
            {
                var recipesCreateModalPage = datasheetDetailPage.CreateNewRecipe();
                recipesCreateModalPage.FillFields_AddNewRecipeToDatasheet(recipeName, itemName2);
            }
            else
            {
                datasheetDetailPage.AddRecipe(recipeName);
            }
            datasheetDetailPage.ClickOnRecipesTab();
            Assert.IsTrue(datasheetDetailPage.IsRecipeAdded(), "La recette n'a pas été rajoutée à la datasheet.");
            // test pour modifier tous les champs
            datasheetDetailPage.ClickOnRecipesTab();
            datasheetDetailPage.SelectSite(site1);
            RecipesVariantPage recipesVariantPage = datasheetDetailPage.EditRecipe();
            recipesVariantPage.ClickOnFirstIngredientDatasheet();
            recipesVariantPage.SetYield(yield);
            recipesVariantPage.SetNetWeight(netWeight);

            var priceVarianSubRecipeSite1 = recipesVariantPage.GetModalPrice(decimalSeparatorValue, currency);
            recipesVariantPage.CloseWindow();
            datasheetDetailPage.SelectSite(site1);
            recipesVariantPage = datasheetDetailPage.EditRecipe();
            var verifyPrice = recipesVariantPage.verifyPricesIngredient(priceVarianSubRecipeSite1, currency);
            var verifyNetWeight = recipesVariantPage.verifyNetWeightIngredients(netWeight);
            var verifyYield = recipesVariantPage.verifyYieldIngredients(yield);
            Assert.IsTrue(verifyPrice, "Les données ne sont pas compatibles.");
            Assert.IsTrue(verifyNetWeight, "Les données ne sont pas compatibles.");
            Assert.IsTrue(verifyYield, "Les données ne sont pas compatibles.");
            datasheetDetailPage = recipesVariantPage.CloseWindow();

            datasheetDetailPage.SelectSite(site2);
            recipesVariantPage = datasheetDetailPage.EditRecipe();
            verifyPrice = recipesVariantPage.verifyPricesIngredient(priceVarianSubRecipeSite1, currency);
            verifyNetWeight = recipesVariantPage.verifyNetWeightIngredients(netWeight);
            verifyYield = recipesVariantPage.verifyYieldIngredients(yield);
            Assert.IsTrue(verifyPrice, "Les données ne sont pas compatibles.");
            Assert.IsTrue(verifyNetWeight, "Les données ne sont pas compatibles.");
            Assert.IsTrue(verifyYield, "Les données ne sont pas compatibles.");
            datasheetDetailPage = recipesVariantPage.CloseWindow();

            // test pour ajouter un item
            datasheetDetailPage.SelectSite(site1);
            recipesVariantPage = datasheetDetailPage.EditRecipe();
            recipesVariantPage.AjouterIngredient(itemName1);
            var verifyIsExistFirstIngredient = recipesVariantPage.verifyIsExistIngredients(itemName1);
            var verifyIsExistSecondIngredient = recipesVariantPage.verifyIsExistIngredients(itemName2);
            Assert.IsTrue(verifyIsExistFirstIngredient, "Les données ne sont pas compatibles.");
            Assert.IsTrue(verifyIsExistSecondIngredient, "Les données ne sont pas compatibles.");
            recipesVariantPage.CloseWindow();

            datasheetDetailPage.SelectSite(site2);
            recipesVariantPage = datasheetDetailPage.EditRecipe();
            verifyIsExistFirstIngredient = recipesVariantPage.verifyIsExistIngredients(itemName1);
            verifyIsExistSecondIngredient = recipesVariantPage.verifyIsExistIngredients(itemName2);
            Assert.IsTrue(verifyIsExistFirstIngredient, "Les données ne sont pas compatibles.");
            Assert.IsTrue(verifyIsExistSecondIngredient, "Les données ne sont pas compatibles.");
            recipesVariantPage.CloseWindow();

            // test pour supprimer un item
            datasheetDetailPage.SelectSite(site1);
            recipesVariantPage = datasheetDetailPage.EditRecipe();
            recipesVariantPage.DeleteIngredient(itemName1);
            verifyIsExistFirstIngredient = recipesVariantPage.verifyIsExistIngredients(itemName1);
            verifyIsExistSecondIngredient = recipesVariantPage.verifyIsExistIngredients(itemName2);
            Assert.IsFalse(verifyIsExistFirstIngredient, "Les données ne sont pas compatibles.");
            Assert.IsTrue(verifyIsExistSecondIngredient, "Les données ne sont pas compatibles.");
            recipesVariantPage.CloseWindow();

            datasheetDetailPage.SelectSite(site2);
            recipesVariantPage = datasheetDetailPage.EditRecipe();
            verifyIsExistFirstIngredient = recipesVariantPage.verifyIsExistIngredients(itemName1);
            verifyIsExistSecondIngredient = recipesVariantPage.verifyIsExistIngredients(itemName2);
            Assert.IsFalse(verifyIsExistFirstIngredient, "Les données ne sont pas compatibles.");
            Assert.IsTrue(verifyIsExistSecondIngredient, "Les données ne sont pas compatibles.");
            recipesVariantPage.CloseWindow();

            // test pour modifier tous les champs de l'item
            datasheetDetailPage.SelectSite(site1);
            recipesVariantPage = datasheetDetailPage.EditRecipe();
            recipesVariantPage.ClickOnFirstIngredientDatasheet();
            recipesVariantPage.SetNetWeight("11");
            recipesVariantPage.SetYield("96");
            var priceItemSite1 = recipesVariantPage.GetModalPrice(decimalSeparatorValue, currency).ToString();
            recipesVariantPage.CloseWindow();

            datasheetDetailPage.SelectSite(site2);
            recipesVariantPage = datasheetDetailPage.EditRecipe();
            recipesVariantPage.ClickOnFirstIngredientDatasheet();
            var priceItemSite2 = recipesVariantPage.GetModalPrice(decimalSeparatorValue, currency).ToString();
            var NetWeight = recipesVariantPage.GetNetWeight(",");
            var Yeild = recipesVariantPage.GetYield(",");
            Assert.AreEqual(priceItemSite1, priceItemSite2, "Les données ne sont pas compatibles dans les deux sites.");
            Assert.AreEqual(NetWeight, 11, "Les données ne sont pas compatibles dans les deux sites.");
            Assert.AreEqual(Yeild, 96, "Les données ne sont pas compatibles dans les deux sites.");

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_ExportDatasheet()
        {
            //Prepare
            DateTime startdateInput = DateUtils.Now.AddDays(-61);
            DateTime enddateInput = DateUtils.Now.AddDays(-30);
            string downloadPath = TestContext.Properties["DownloadsPath"].ToString();
            //Arrange
            HomePage homePage = LogInAsAdmin();
            homePage.ClearDownloads();
            //Etre sur l'index des Datasheets
            DatasheetPage datasheetPage = homePage.GoToMenus_Datasheet();
            //1. Survoler les ...
            //2.Cliquer sur Export
            datasheetPage.Filter(DatasheetPage.FilterType.DateFrom, startdateInput);
            datasheetPage.Filter(DatasheetPage.FilterType.DateTo, enddateInput);
            datasheetPage.ExportDatasheet();
            //Le fichier Excel est téléchargé
            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();
            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = datasheetPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");
            // Récupération du nom du fichier et construction de l'URL du fichier Excel à ouvrir   
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadPath, fileName);
            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber("Datasheets", filePath);
            List<string> result = OpenXmlExcel.GetValuesInList("Name", "Datasheets", filePath);
            //Assert
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.IsFalse(result.Contains(""), MessageErreur.EXCEL_DONNEES_KO);
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_ImportDatasheet()
        {
            //Prepare
            string downloadPath = TestContext.Properties["DownloadsPath"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            homePage.ClearDownloads();

            //Etre sur l'index des Datasheets
            var datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetPage.Filter(DatasheetPage.FilterType.Sites, "AGP");
            //1. Survoler les ...
            //2.Cliquer sur Export
            datasheetPage.ExportDatasheet();

            //Le fichier Excel est téléchargé
            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = datasheetPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

            // Récupération du nom du fichier et construction de l'URL du fichier Excel à ouvrir   
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadPath, fileName);

            datasheetPage.ClosePrintButton();

            // Exploitation du fichier Excel

            // On récupère les 3 ids dans le fichier Excel
            var ids = OpenXmlExcel.GetValuesInList("Datasheet ID", "Datasheets", correctDownloadedFile.FullName);
            var id1 = ids[0];
            var id2 = ids[1];
            var id3 = ids[2];

            //On modifie l'Action pour les 3 datasheets
            OpenXmlExcel.WriteDataInCell("Action", "Datasheet ID", id1, "Datasheets", filePath, "M", CellValues.SharedString);
            OpenXmlExcel.WriteDataInCell("Action", "Datasheet ID", id2, "Datasheets", filePath, "X", CellValues.SharedString);
            OpenXmlExcel.WriteDataInCell("Action", "Datasheet ID", id3, "Datasheets", filePath, "M", CellValues.SharedString);

            datasheetPage.ImportDatasheet(DatasheetPage.ImportType.ImportDatasheet, filePath);
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_ChecktotaCosting()
        {
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();

            Random rnd = new Random();
            string datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();
            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            datasheetDetailPage.AddRecipe(recipeName);
            datasheetDetailPage.WaitLoading();
            WebDriver.Navigate().Refresh();
            Assert.IsTrue(datasheetDetailPage.IsRecipeAdded(), "La recette n'a pas été rajoutée à la datasheet.");

            var recipePrice = datasheetDetailPage.GetRecipePrice(decimalSeparatorValue);
            var totalCosting = datasheetDetailPage.GetTotalCosting(decimalSeparatorValue);

            Assert.AreEqual(recipePrice, totalCosting, "Le coût total n'est pas égal au coût de la recette.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_UpdateRecipeNetQty()
        {
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();
            string currency = TestContext.Properties["Currency"].ToString();

            Random rnd = new Random();
            string datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();
            double editedNetQty = rnd.Next(1, 10);

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            datasheetDetailPage.AddRecipe(recipeName);
            datasheetDetailPage.WaitLoading();
            WebDriver.Navigate().Refresh();
            Assert.IsTrue(datasheetDetailPage.IsRecipeAdded(), "La recette n'a pas été rajoutée à la datasheet.");

            var recipeVariantPage = datasheetDetailPage.EditRecipe();
            // Récupération du poids, avant sélection de la ligne
            var weight = recipeVariantPage.GetWeight(decimalSeparatorValue);
            recipeVariantPage.ClickOnFirstIngredientDatasheet();
            // Récupération des valeurs initiales
            var initPrice = recipeVariantPage.GetPrice(decimalSeparatorValue, currency);
            var yield = recipeVariantPage.GetYield(decimalSeparatorValue);

            // Modification de la valeur de NetQuantity
            recipeVariantPage.SetNetQuantity(editedNetQty.ToString());

            datasheetDetailPage = recipeVariantPage.CloseWindow();
            recipeVariantPage = datasheetDetailPage.EditRecipe();
            recipeVariantPage.ClickOnFirstIngredientDatasheet();

            // Récupération des valeurs modifiées
            var newNetQty = recipeVariantPage.GetNetQuantity(decimalSeparatorValue);
            var newNetWeight = recipeVariantPage.GetNetWeight(decimalSeparatorValue);
            var newGrossQty = recipeVariantPage.GetGrossQty(decimalSeparatorValue);
            var newPrice = recipeVariantPage.GetPrice(decimalSeparatorValue, currency);
            var newTotalPortion = recipeVariantPage.GetTotalPortion(decimalSeparatorValue);

            // Vérification de la modification
            Assert.AreEqual(editedNetQty, newNetQty, "La valeur de Net Quantity n'a pas été modifiée.");
            Assert.AreEqual(newNetWeight, Math.Round(editedNetQty * weight, 3), "La valeur de Net Weight est incorrecte.");
            Assert.AreEqual(editedNetQty, Math.Round(newGrossQty * yield / 100, 3), "La valeur de Gross Quantity est incorrecte.");
            Assert.AreNotEqual(newPrice, initPrice, "Le prix n'a pas été recalculé.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_UpdateRecipeNetWeight()
        {
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();
            string currency = TestContext.Properties["Currency"].ToString();

            Random rnd = new Random();
            string datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();
            double editedNetWeight = rnd.Next(1, 10);

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            datasheetDetailPage.AddRecipe(recipeName);
            //WebDriver.Navigate().Refresh();
            Assert.IsTrue(datasheetDetailPage.IsRecipeAdded(), "La recette n'a pas été rajoutée à la datasheet.");

            var recipeVariantPage = datasheetDetailPage.EditRecipe();
            // Récupération du poids, avant sélection de la ligne
            var weight = recipeVariantPage.GetWeight(decimalSeparatorValue);
            recipeVariantPage.ClickOnFirstIngredientDatasheet();
            // Récupération des valeurs initiales
            var initPrice = recipeVariantPage.GetPrice(decimalSeparatorValue, currency);
            var yield = recipeVariantPage.GetYield(decimalSeparatorValue);

            // Modification de la valeur de NetWeight
            recipeVariantPage.SetNetWeight(editedNetWeight.ToString());

            datasheetDetailPage = recipeVariantPage.CloseWindow();
            recipeVariantPage = datasheetDetailPage.EditRecipe();
            recipeVariantPage.ClickOnFirstIngredientDatasheet();

            // Récupération des valeurs modifiées
            var newNetQty = recipeVariantPage.GetNetQuantity(decimalSeparatorValue);
            var newNetWeight = recipeVariantPage.GetNetWeight(decimalSeparatorValue);
            var newGrossQty = recipeVariantPage.GetGrossQty(decimalSeparatorValue);
            var newPrice = recipeVariantPage.GetPrice(decimalSeparatorValue, currency);
            var newTotalPortion = recipeVariantPage.GetTotalPortion(decimalSeparatorValue);

            // Vérification de la modification
            Assert.AreEqual(editedNetWeight, newNetWeight, "La valeur de Net Weight n'a pas été modifiée.");
            Assert.AreEqual(newNetQty, Math.Round(editedNetWeight / weight, 3), "La valeur de Net Quantity est incorrecte.");
            Assert.AreEqual(newNetQty, Math.Round(newGrossQty * yield / 100, 3), "La valeur de Gross Quantity est incorrecte.");
            Assert.AreNotEqual(newPrice, initPrice, "Le prix n'a pas été recalculé.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_UpdateRecipePortion()
        {
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();
            string currency = TestContext.Properties["Currency"].ToString();

            Random rnd = new Random();
            string datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();
            double editedRecipePortion = rnd.Next(1, 10);

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            datasheetDetailPage.AddRecipe(recipeName);
            datasheetDetailPage.WaitLoading();
            WebDriver.Navigate().Refresh();
            Assert.IsTrue(datasheetDetailPage.IsRecipeAdded(), "La recette n'a pas été rajoutée à la datasheet.");

            var recipeVariantPage = datasheetDetailPage.EditRecipe();
            recipeVariantPage.ClickOnFirstIngredientDatasheet();
            // Récupération des valeurs initiales
            var initNetQty = recipeVariantPage.GetNetQuantity(decimalSeparatorValue);
            var initNetWeight = recipeVariantPage.GetNetWeight(decimalSeparatorValue);
            var initGrossQty = recipeVariantPage.GetGrossQty(decimalSeparatorValue);
            var initPrice = recipeVariantPage.GetPrice(decimalSeparatorValue, currency);

            // Modification de la valeur de Recipe Portion
            recipeVariantPage.SetTotalPortion(editedRecipePortion.ToString());

            datasheetDetailPage = recipeVariantPage.CloseWindow();
            recipeVariantPage = datasheetDetailPage.EditRecipe();
            recipeVariantPage.ClickOnFirstIngredientDatasheet();

            // Récupération des valeurs modifiées
            var newNetQty = recipeVariantPage.GetNetQuantity(decimalSeparatorValue);
            var newNetWeight = recipeVariantPage.GetNetWeight(decimalSeparatorValue);
            var newGrossQty = recipeVariantPage.GetGrossQty(decimalSeparatorValue);
            var newPrice = recipeVariantPage.GetPrice(decimalSeparatorValue, currency);
            var newTotalPortion = recipeVariantPage.GetTotalPortion(decimalSeparatorValue);

            // Vérification de la modification
            Assert.AreEqual(editedRecipePortion, newTotalPortion, "La valeur de Recipe Portion n'a pas été modifiée.");
            Assert.AreNotEqual(newNetQty, initNetQty, "La valeur de Net Quantity n'a pas été actualisée");
            Assert.AreNotEqual(newNetWeight, initNetWeight, "La valeur de Net Weight n'a pas été actualisée");
            Assert.AreNotEqual(newGrossQty, initGrossQty, "La valeur de Gross Quantity n'a pas été actualisée");
            Assert.AreNotEqual(newPrice, initPrice, "Le prix n'a pas été recalculé.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_UpdateRecipeYield()
        {
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();
            string currency = TestContext.Properties["Currency"].ToString();

            Random rnd = new Random();
            string datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();
            double editedYield = rnd.Next(1, 100);

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            datasheetDetailPage.AddRecipe(recipeName);
            datasheetDetailPage.WaitLoading();
            WebDriver.Navigate().Refresh();
            Assert.IsTrue(datasheetDetailPage.IsRecipeAdded(), "La recette n'a pas été rajoutée à la datasheet.");

            var recipeVariantPage = datasheetDetailPage.EditRecipe();
            // Récupération des valeurs initiales
            var initPrice = recipeVariantPage.GetPrice(decimalSeparatorValue, currency);

            // Modification de la valeur de Yield
            recipeVariantPage.ClickOnFirstIngredientDatasheet();
            recipeVariantPage.SetYield(editedYield.ToString());

            datasheetDetailPage = recipeVariantPage.CloseWindow();
            recipeVariantPage = datasheetDetailPage.EditRecipe();
            recipeVariantPage.ClickOnFirstIngredientDatasheet();

            // Récupération des valeurs modifiées
            var newNetQty = recipeVariantPage.GetNetQuantity(decimalSeparatorValue);
            var newGrossQty = recipeVariantPage.GetGrossQty(decimalSeparatorValue);
            var newYield = recipeVariantPage.GetYield(decimalSeparatorValue);
            var newPrice = recipeVariantPage.GetModalPrice(decimalSeparatorValue, currency);

            // Vérification de la modification
            Assert.AreEqual(editedYield, newYield, "La valeur de Yield n'a pas été modifiée.");
            Assert.AreEqual(newNetQty, Math.Round(newGrossQty * editedYield / 100, 3), "La valeur de Gross Quantity est incorrecte.");
            Assert.AreNotEqual(newPrice, initPrice, "Le prix n'a pas été recalculé.");
        }

        // _____________________________________________________________________________________________________________________

        public void AddServiceForDatasheet(HomePage homePage, string serviceName, string guestType, DateTime dateFrom, DateTime dateTo, string datasheetName)
        {
            // Prepare
            string customer = TestContext.Properties["MenuServiceCustomer"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string serviceCategory = TestContext.Properties["CategoryServiceDefault"].ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);

            // Act
            var customerPage = homePage.GoToCustomers_CustomerPage();
            customerPage.ResetFilters();
            customerPage.Filter(PageObjects.Customers.Customer.CustomerPage.FilterType.Search, customer);
            string customerCode = customerPage.GetFirstCustomerIcao();

            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();

            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction, serviceCategory, guestType);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            pricePage = priceModalPage.FillFields_CustomerPrice(site, customerCode + " - " + customer, dateFrom, dateTo, "2", datasheetName);
            servicePage = pricePage.BackToList();
            //Assert
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.WaitLoading();

            Assert.IsTrue(servicePage.GetFirstServiceName().Contains(serviceName), "Le service " + serviceName + " n'existe pas.");
        }
        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_UpdateSubRecipeAddIngredient()
        {
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();

            Random rnd = new Random();
            string datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            datasheetDetailPage.AddRecipe(recipeName);
            datasheetDetailPage.WaitLoading();
            WebDriver.Navigate().Refresh();
            var recipeVariantPage = datasheetDetailPage.EditRecipe();
            var nubIngradA = recipeVariantPage.CalculIngredientLignes();
            recipeVariantPage.AjouterIngredient("a");
            var nubIngradb = recipeVariantPage.CalculIngredientLignes();

            Assert.AreNotEqual(nubIngradb, nubIngradA, "l’ingrédient n'est pas ajouté dans la sous-recette");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_UpdateSubRecipeDieteticComputed()
        {
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();

            Random rnd = new Random();
            string datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            datasheetDetailPage.AddRecipe(recipeName);
            datasheetDetailPage.WaitLoading();
            WebDriver.Navigate().Refresh();
            var recipeVariantPage = datasheetDetailPage.EditRecipe();
            recipeVariantPage.ClickOnDieteticTab();
            if (!recipeVariantPage.IsDieteticComputed())
            {
                recipeVariantPage.ChangeDieteticComputedManual();
            }
            else
            {
                recipeVariantPage.ChangeDieteticComputedManual();
                recipeVariantPage.ChangeDieteticComputedManual();

            }

            Assert.IsFalse(recipeVariantPage.IsMalualDataVisible(), "Le nom de l'utilisateur et l'horaire de validation sont visibles.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_UpdateSubRecipeDieteticmanual()
        {
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();

            Random rnd = new Random();
            string datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            datasheetDetailPage.AddRecipe(recipeName);
            datasheetDetailPage.WaitLoading();
            WebDriver.Navigate().Refresh();
            var recipeVariantPage = datasheetDetailPage.EditRecipe();
            recipeVariantPage.ClickOnDieteticTab();
            if (recipeVariantPage.IsDieteticComputed())
            {
                recipeVariantPage.ChangeDieteticComputedManual();
            }
            else
            {
                recipeVariantPage.ChangeDieteticComputedManual();
                recipeVariantPage.ChangeDieteticComputedManual();

            }

            Assert.IsTrue(recipeVariantPage.IsMalualDataVisible(), "Le nom de l'utilisateur et l'horaire de validation ne sont pas visibles.");
        }
        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_UpdateSubRecipeNetQty()
        {
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();
            string currency = TestContext.Properties["Currency"].ToString();

            Random rnd = new Random();
            string datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();
            double editedNetQty = rnd.Next(1, 10);

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            datasheetDetailPage.AddRecipe(recipeName);
            datasheetDetailPage.WaitLoading();
            WebDriver.Navigate().Refresh();
            Assert.IsTrue(datasheetDetailPage.IsRecipeAdded(), "La recette n'a pas été rajoutée à la datasheet.");

            var recipeVariantPage = datasheetDetailPage.EditRecipe();
            // Récupération du poids, avant sélection de la ligne
            var weight = recipeVariantPage.GetWeight(decimalSeparatorValue);
            recipeVariantPage.ClickOnFirstIngredientDatasheet();
            // Récupération des valeurs initiales
            var initPrice = recipeVariantPage.GetPrice(decimalSeparatorValue, currency);
            var yield = recipeVariantPage.GetYield(decimalSeparatorValue);

            // Modification de la valeur de NetQuantity
            recipeVariantPage.SetNetQuantity(editedNetQty.ToString());
            recipeVariantPage.WaitLoading();

            // Récupération des valeurs modifiées
            var newNetQty = recipeVariantPage.GetNetQuantity(decimalSeparatorValue);
            var newNetWeight = recipeVariantPage.GetNetWeight(decimalSeparatorValue);
            var newGrossQty = recipeVariantPage.GetGrossQty(decimalSeparatorValue);
            var newPrice = recipeVariantPage.GetPrice(decimalSeparatorValue, currency);
            var newTotalPortion = recipeVariantPage.GetTotalPortion(decimalSeparatorValue);

            // Vérification de la modification
            Assert.AreEqual(editedNetQty, newNetQty, "La valeur de Net Quantity n'a pas été modifiée.");
            Assert.AreEqual(newNetWeight, Math.Round(editedNetQty * weight, 3), "La valeur de Net Weight est incorrecte.");
            Assert.AreEqual(editedNetQty, Math.Round(newGrossQty * yield / 100, 3), "La valeur de Gross Quantity est incorrecte.");
            Assert.AreNotEqual(newPrice, initPrice, "Le prix n'a pas été recalculé.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_UpdateSubRecipeNetWeight()
        {
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();
            string currency = TestContext.Properties["Currency"].ToString();

            Random rnd = new Random();
            string datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();
            double NetWeightEdited = rnd.Next(1, 10);
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            datasheetDetailPage.AddRecipe(recipeName);
            datasheetDetailPage.WaitLoading();
            WebDriver.Navigate().Refresh();
            Assert.IsTrue(datasheetDetailPage.IsRecipeAdded(), "La recette n'a pas été rajoutée à la datasheet.");

            var recipeVariantPage = datasheetDetailPage.EditRecipe();
            // Récupération du poids, avant sélection de la ligne
            var weight = recipeVariantPage.GetWeight(decimalSeparatorValue);
            recipeVariantPage.ClickOnFirstIngredientDatasheet();
            // Récupération des valeurs initiales
            var initPrice = recipeVariantPage.GetPrice(decimalSeparatorValue, currency);
            var yield = recipeVariantPage.GetYield(decimalSeparatorValue);

            // Modification de la valeur de NetQuantity
            recipeVariantPage.SetNetWeight(NetWeightEdited.ToString());
            recipeVariantPage.WaitLoading();

            // Récupération des valeurs modifiées
            var newNetQty = recipeVariantPage.GetNetQuantity(decimalSeparatorValue);
            var newNetWeight = recipeVariantPage.GetNetWeight(decimalSeparatorValue);
            var newGrossQty = recipeVariantPage.GetGrossQty(decimalSeparatorValue);
            var newPrice = recipeVariantPage.GetPrice(decimalSeparatorValue, currency);
            var newTotalPortion = recipeVariantPage.GetTotalPortion(decimalSeparatorValue);

            // Vérification de la modification

            Assert.AreEqual(NetWeightEdited, newNetWeight, "La valeur de Net Weight n'a pas été modifiée.");
            Assert.AreEqual(newNetQty, Math.Round(NetWeightEdited / weight, 3), "La valeur de Net Quantity est incorrecte.");
            Assert.AreEqual(newNetQty, Math.Round(newGrossQty * yield / 100, 3), "La valeur de Gross Quantity est incorrecte.");
            Assert.AreNotEqual(newPrice, initPrice, "Le prix n'a pas été recalculé.");
        }
        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_UpdateSubRecipeNotShowItemOnly()
        {
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();
            string recipeName1 = "Recette" + new Random().Next().ToString();
            Random rnd = new Random();
            string datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);
            var datasheetCreateNewRecipePage = datasheetDetailPage.CreateNewRecipe();
            datasheetDetailPage = datasheetCreateNewRecipePage.FillFields_AddNewRecipeToDatasheet(recipeName1, ingredient);
            datasheetDetailPage.WaitLoading();
            WebDriver.Navigate().Refresh();
            var recipeVariantPage = datasheetDetailPage.EditRecipe();
            var nubIngradA = recipeVariantPage.GetNumberOfIngredient();
            recipeVariantPage.AddSubRecipeDatasheet(recipeName);
            var nubIngradb = recipeVariantPage.GetNumberOfIngredient();

            Assert.AreNotEqual(nubIngradb, nubIngradA, "La sous-recette n'a pas été ajouté dans la sous-recette");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_UpdateSubRecipeImportIngredient()
        {
            //Prepare
            Random rnd = new Random();
            string datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = "FLIGHT - FLIGHT";
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string ingredientBis = TestContext.Properties["RecipeIngredientBis"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName1 = "Recette1" + new Random().Next().ToString();
            string recipeNameBis = recipeName1 + "Bis";
            int nbPortions = new Random().Next(1, 30);
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act Add recipe
            var recipesPage = homePage.GoToMenus_Recipes();
            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName1, recipeType, nbPortions.ToString());
            recipeGeneralInfosPage.AddVariantWithSite(site, variant);
            var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
            recipeVariantPage.AddIngredient(ingredient);
            Assert.IsTrue(recipeVariantPage.IsIngredientDisplayed(), "L'ingrédient n'a pas été ajouté à la recette.");
            //Act go to Datasheet
            homePage.Navigate();
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);
            var datasheetCreateNewRecipePage = datasheetDetailPage.CreateNewRecipe();
            datasheetDetailPage = datasheetCreateNewRecipePage.FillFields_AddNewRecipeToDatasheet(recipeNameBis, ingredientBis);
            datasheetDetailPage.WaitLoading();
            WebDriver.Navigate().Refresh();
            //recipeVariantPage = datasheetDetailPage.EditRecipe();
            Thread.Sleep(2000);
            //var recipeGeneralInformationPage = recipeVariantPage.ClickOnGeneralInformationFromDatasheet();
            //recipeNameBis = recipeGeneralInformationPage.GetRecipeNameFlight();
            recipeNameBis = datasheetDetailPage.GetFirstRecipeName();// = recipeGeneralInformationPage.ClickOnDetailsFromDatasheet();
            datasheetDetailPage.WaitLoading();
            datasheetDetailPage.AddRecipe(recipeName1);
            datasheetDetailPage.WaitLoading();
            WebDriver.Navigate().Refresh();
            Assert.IsTrue(datasheetDetailPage.IsRecipeAdded(), "La recette n'a pas été rajoutée à la datasheet.");
            recipeVariantPage = datasheetDetailPage.EditRecipe(recipeName1);
            var importIngredientPage = recipeVariantPage.ImportIngredientSubRecipe();
            importIngredientPage.FillFields_ImportIngredientSub(recipeNameBis, variant, true);
            Assert.IsTrue(importIngredientPage.IsImportOK(), "L'import n'a pas été effectué.");
            recipeVariantPage = importIngredientPage.ConfirmImport();

            Assert.AreEqual(1, recipeVariantPage.GetNumberOfIngredient(), "Le nombre d'ingrédients de la recette est incorrect.");
            Assert.AreEqual(recipeVariantPage.GetIngredientName(), ingredientBis, "L'import d'ingrédient a échoué.");
        }
        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_UpdateSubRecipeYield()
        {
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();
            string currency = TestContext.Properties["Currency"].ToString();

            Random rnd = new Random();
            string datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();
            double editedYield = rnd.Next(1, 100);

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            datasheetDetailPage.AddRecipe(recipeName);
            datasheetDetailPage.WaitLoading();
            WebDriver.Navigate().Refresh();
            Assert.IsTrue(datasheetDetailPage.IsRecipeAdded(), "La recette n'a pas été rajoutée à la datasheet.");
            var recipeVariantPage = datasheetDetailPage.EditRecipe();
            recipeVariantPage.ClickOnFirstIngredientDatasheet();
            // Récupération des valeurs initiales
            var initPrice = recipeVariantPage.GetPrice(decimalSeparatorValue, currency);

            // Modification de la valeur de Yield
            recipeVariantPage.SetYieldData(editedYield.ToString());

            // Récupération des valeurs modifiées

            var newNetQty = recipeVariantPage.GetNetQuantity(decimalSeparatorValue);
            var newGrossQty = recipeVariantPage.GetGrossQty(decimalSeparatorValue);
            var newYield = recipeVariantPage.GetYield(decimalSeparatorValue);
            var newPrice = recipeVariantPage.GetPrice(decimalSeparatorValue, currency);

            // Vérification de la modification
            Assert.AreEqual(editedYield, newYield, "La valeur de Yield n'a pas été modifiée.");
            Assert.AreEqual(newNetQty, Math.Round(newGrossQty * editedYield / 100, 3), "La valeur de Gross Quantity est incorrecte.");
            Assert.AreNotEqual(newPrice, initPrice, "Le prix n'a pas été recalculé.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_ChecktotaRecipe()
        {
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();
            string currency = TestContext.Properties["Currency"].ToString();

            Random rnd = new Random();
            string datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();
            double editedNetWeight = rnd.Next(1, 100);

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            datasheetDetailPage.AddRecipe(recipeName);
            Assert.IsTrue(datasheetDetailPage.IsRecipeAdded(), "La recette n'a pas été rajoutée à la datasheet.");
            var recipeVariantPage = datasheetDetailPage.EditRecipe();
            recipeVariantPage.ClickOnFirstIngredientDatasheet();

            // Modification de la valeur de Net Weight
            recipeVariantPage.SetNetWeight(editedNetWeight.ToString());
            var initPrice = recipeVariantPage.GetTotalPriceFromAllRows(decimalSeparatorValue, currency);
            var netWeight = recipeVariantPage.GetTotalPortion(decimalSeparatorValue).ToString();
            if (!datasheetDetailPage.isPopupClosed())
            {
                datasheetDetailPage.CloseDetailModal();
            }
            datasheetDetailPage.ClickOnFirstRecipe();

            //Assert
            var netWeightRecipe = datasheetDetailPage.GetNetWeight(decimalSeparatorValue).ToString();
            Assert.AreEqual(netWeight, netWeightRecipe, "La nouvelle valeur du NetWeight de la recette n'est pas correct.");
            Assert.AreEqual(initPrice.ToString().Replace(".", ","), datasheetDetailPage.GetPAXRecipe(), "La nouvelle valeur du Pax  n'est pas correct.");
            Assert.AreEqual(initPrice.ToString().Replace(".", ","), datasheetDetailPage.GetPortion(), "La nouvelle valeur du portion  n'est pas correct.");
            Assert.AreEqual(initPrice.ToString(), datasheetDetailPage.GetTotalCosting(decimalSeparatorValue).ToString(), "La nouvelle valeur du total costing  n'est pas correct.");

        }
        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_Massivedelete_usecase()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            DatasheetPage datasheetPage = homePage.GoToMenus_Datasheet();
            var firstDatasheetName = datasheetPage.GetFirstDatasheetName();
            DatasheetMassiveDeletePopup dsMassiveDelete = datasheetPage.OpenMassiveDeletePopup();
            dsMassiveDelete.SetDatasheetName(firstDatasheetName);
            dsMassiveDelete.ClickOnSearch();
            dsMassiveDelete.SetPageSize("100");
            int totalUseCases = dsMassiveDelete.GetUseCases();
            DatasheetDetailsPage dsPage = dsMassiveDelete.GoToDatasheetPage();
            dsPage.ClickOnUseCaseTab();
            int dsUseCases = dsPage.CountUseCases();

            Assert.AreEqual(dsUseCases, totalUseCases, "Les usecase affichées dans les résultats ne sont pas correctes");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_RecipePicture_Duplication_Datasheets()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();

            string datasheetName = datasheetNameToday + "-" + new Random().Next().ToString();
            string datasheetNameBis = datasheetName + "Bis";

            var path = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase);
            path = path.Substring(6);
            var imagePath = path + "\\PageObjects\\Menus\\Datasheet\\test.jpg";

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            datasheetDetailPage.AddRecipe(recipeName);
            datasheetDetailPage.WaitLoading();
            //add picture to recipe 
            var recipeVariantPage = datasheetDetailPage.AddPictureToRecipe();
            recipeVariantPage.AddPicture(imagePath);
            datasheetDetailPage = recipeVariantPage.CloseWindow();

            bool isPictureAdded = datasheetDetailPage.IsPictureAddedToRecipe();
            Assert.IsTrue(isPictureAdded, "L'image n'a pas été ajoutée à la recette.");

            WebDriver.Navigate().Refresh();

            bool isRecipeAdded = datasheetDetailPage.IsRecipeAdded();
            Assert.IsTrue(isRecipeAdded, "La recette n'a pas été rajoutée à la datasheet.");

            datasheetPage = datasheetDetailPage.BackToList();
            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);

            string firstDatasheetName = datasheetPage.GetFirstDatasheetName();
            Assert.AreEqual(firstDatasheetName, datasheetName, "La datasheet n'a pas été créée.");

            datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetNameBis, guestType, site);

            bool isRecipeAdded2 = datasheetDetailPage.IsRecipeAdded();
            Assert.IsFalse(isRecipeAdded2, "La datasheet possède déjà une recette.");

            datasheetDetailPage.AddRecipeFromDatasheet(datasheetName);

            bool isRecipeAdded3 = datasheetDetailPage.IsRecipeAdded();
            Assert.IsTrue(isRecipeAdded3, "La recette de la datasheet " + datasheetName + " n'a pas été rajoutée à la datasheet " + datasheetNameBis);

            string firstRecipeName = datasheetDetailPage.GetFirstRecipeName();
            Assert.AreEqual(firstRecipeName, recipeName, "La recette n'a pas été ajoutée correctement.");
            datasheetDetailPage.AddPictureToRecipe();

            bool isPictureExistOK = datasheetDetailPage.IsPictureExist();
            Assert.IsTrue(isPictureExistOK, "la duplication d'une recette existante depuis une datasheet vers une autre doit copier la photo de la recette aussi.");

        }

        [TestMethod]
        public void MENU_DATA_MassiveDelete_DatasheetName_Exist()
        {

            //Home page
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var dataSheetPage = homePage.GoToMenus_Datasheet();
            string datasheetName = dataSheetPage.GetFirstDatasheetName();
            var dsMassiveDelete = dataSheetPage.OpenMassiveDeletePopup();
            dataSheetPage.EnterDatasheetName(datasheetName);

            dsMassiveDelete.ClickOnSearch();
            dsMassiveDelete.SelectAllSites();
            dsMassiveDelete.ClickDeleteButton();
            // Revenir sur la page de création de datasheet
            var datasheetCreateModalPage = dataSheetPage.CreateNewDatasheet();

            // Remplir les champs avec le même nom de datasheet (datasheetName)
            datasheetCreateModalPage.FillName_NewDatasheet(datasheetName);

            // Sélectionner tous les sites pour la nouvelle datasheet
            datasheetCreateModalPage.SelectAllSites();

            datasheetCreateModalPage.ClickCreateButton();

            // Vérifier si la création a réussi en passant le nom de la datasheet
            bool isCreated = datasheetCreateModalPage.IsDatasheetCreatedSuccessfully(datasheetName);

            //Le test réussit si la création a été effectuée, sinon le test échoue
            Assert.IsTrue(isCreated, "La suppression de la datasheet n'a pas été effectuée correctement, car la création avec le même nom a échoué.");


        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_AccesErrorSite()
        {
            //Prepare recipe
            string recipeName1 = "RecipeForItemDatasheet";

            //Prepare
            string site1 = TestContext.Properties["SiteACE"].ToString();
            string site2 = TestContext.Properties["SiteBCN"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            Random rnd = new Random();
            string datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();
            string recipeName2 = "RecipeTest-" + rnd.Next().ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //go to create datasheet
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            DatasheetDetailsPage datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheetWidh2Sites(datasheetName, guestType, site1, site2);


            //add recipe with ingredient 
            datasheetDetailPage.SelectSite(site1);
            var recipesCreateModalPage = datasheetDetailPage.CreateNewRecipe();
            recipesCreateModalPage.FillFields_AddNewRecipeToDatasheet(recipeName1);
            datasheetDetailPage.ClickOnRecipesTab();
            var isRecipe1Added = datasheetDetailPage.IsRecipeAdded();
            Assert.IsTrue(isRecipe1Added, "La recette n'a pas été rajoutée à la datasheet.");

            //add sub-recipe with ingredient
            datasheetDetailPage.SelectSite(site2);
            var datasheetCreateNewRecipePage = datasheetDetailPage.CreateNewRecipe();
            datasheetCreateNewRecipePage.FillFields_AddNewRecipeToDatasheet(recipeName2);
            datasheetDetailPage.ClickOnRecipesTab();
            var isRecipeAdded2 = datasheetDetailPage.IsRecipeAdded();
            Assert.IsTrue(isRecipeAdded2, "La recette n'a pas été rajoutée à la datasheet.");

            datasheetDetailPage.SelectSite(site1);

            var checkRecipe1Exist = datasheetDetailPage.CheckRecipeExist(recipeName1);
            Assert.IsTrue(checkRecipe1Exist, "datasheet sur la liste n'a appartient pas à plus d'un site");
            datasheetDetailPage.SelectSite(site2);

            var checkRecipe2Exist = datasheetDetailPage.CheckRecipeExist(recipeName2);
            Assert.IsTrue(checkRecipe2Exist, "datasheet sur la liste n'a appartient pas à plus d'un site");

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_AddIngredient_Duplication()
        {
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();

            Random rnd = new Random();
            string datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);

            datasheetDetailPage.AddRecipe(recipeName);
            datasheetDetailPage.ClickOnFirstRecipe();

            datasheetDetailPage.WaitLoading();
            WebDriver.Navigate().Refresh();
            var recipeVariantPage = datasheetDetailPage.EditRecipe();
            var itemName = recipeVariantPage.GetFirstItemNameSousRecipe();
            var nubIngradA = recipeVariantPage.CalculIngredientLignes();
            string Ingredient = recipeVariantPage.AjouterIngredient2(itemName);

            //Test
            var nubIngradb = recipeVariantPage.CalculIngredientLignes();
            Assert.AreNotEqual(nubIngradb, nubIngradA + 1, "l’ingrédient a ajouté dans la sous-recette, l'item ne doit s'afficher qu'une seul fois.");
        }
        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_PrintDatasheet()
        {
            //Prepare
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            var customerICAO = "Cust198";
            bool versionPrint = true;
            string site = TestContext.Properties["SiteMAD"].ToString();
            string customer = TestContext.Properties["CustomerLPFlight"].ToString();
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNamePdfBegin = "DataSheets Report";
            string DocFileNameZipBegin = "All_files_";
            List<string> listdatasheetname = new List<string>();
            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            for (int i = 0; i < 3; i++)
            {
                Random rnd = new Random();
                string datasheetName = "datasheetforprint" + "-" + i + DateUtils.Now.ToString("dd/MM/yyyy") + "-" + rnd.Next().ToString();
                listdatasheetname.Add(datasheetName);
                datasheetPage = homePage.GoToMenus_Datasheet();
                var createDataSheetPage = datasheetPage.CreateNewDatasheet();
                createDataSheetPage.FillField_CreateNewDatasheet(datasheetName, guestType, site, customerICAO);
            }
            datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, "datasheetforprint");
            try
            {
                if (datasheetPage.CheckTotalNumber() > 0)
                {
                    var listNameDataSheet = datasheetPage.GetAllNameDatasheet();
                    datasheetPage.ClearDownloads();
                    var reportPage = datasheetPage.PrintDatasheet(versionPrint);
                    var isGenerated = reportPage.IsReportGenerated();
                    reportPage.Close();
                    Assert.IsTrue(isGenerated, "Le document PDF n'a pas pu être généré par l'application.");
                    reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
                    datasheetPage.WaitPageLoading();
                    string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin, false);
                    FileInfo filePdf = new FileInfo(trouve);
                    filePdf.Refresh();
                    Assert.IsTrue(filePdf.Exists, trouve + " non généré");
                    PdfDocument document = PdfDocument.Open(filePdf.FullName);
                    List<string> mots = new List<string>();
                    foreach (Page page in document.GetPages())
                    {
                        mots.AddRange(page.GetWords().Select(m => m.Text));
                    }
                    bool allDataSheetInResult = listNameDataSheet.All(datsheet => mots.Contains(datsheet));
                    Assert.IsTrue(allDataSheetInResult, "le print de toutes les datasheets n'a pas pu être généré.");
                }
            }
            finally
            {
                datasheetPage = homePage.GoToMenus_Datasheet();
                var dsMassiveDelete = datasheetPage.OpenMassiveDeletePopup();
                dsMassiveDelete.SetDatasheetName("datasheetforprint");
                dsMassiveDelete.ClickOnSearch();
                dsMassiveDelete.SelectAllSites();
                dsMassiveDelete.ClickDeleteButton();
            }
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void MENU_DATA_MAJ_TotalPRICE_TotalQTY_SousRecipes()
        {

            //Prepare recipe
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string recipeName = "RECETTE" + new Random().Next();
            string SousrecipeName = "SOUS RECETTE" + new Random().Next();
            string currency = TestContext.Properties["Currency"].ToString();


            //Prepare ingredient
            var guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            var ingredientBis = TestContext.Properties["RecipeIngredientBis"].ToString();
            var rnd = new Random();
            var datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            DatasheetPage datasheetPage = homePage.GoToMenus_Datasheet();
            DatasheetCreateModalPage datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            DatasheetDetailsPage datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);
            DatasheetCreateNewRecipePage datasheetCreateNewRecipePage = datasheetDetailPage.CreateNewRecipe();
            datasheetDetailPage = datasheetCreateNewRecipePage.FillFields_AddNewRecipeToDatasheet(recipeName, ingredientBis);

            //add subrecipe 
            datasheetCreateNewRecipePage = datasheetDetailPage.AddSubrecipeForFirstRecipe();
            datasheetCreateNewRecipePage.FillFields_AddNewRecipeToDatasheet(SousrecipeName, ingredientBis);
            datasheetCreateNewRecipePage.AddIngredient(ingredient);
            var weightBeforeedit = datasheetCreateNewRecipePage.GetTotalPortion(decimalSeparatorValue);
            datasheetCreateNewRecipePage.Closebtn();
            RecipesVariantPage recipesVariantPage = datasheetDetailPage.EditRecipe();
            double pricebeforeedit = recipesVariantPage.GetPriceSouRecette(decimalSeparatorValue, currency);
            recipesVariantPage.DeleteSubRecipe(ingredient);
            recipesVariantPage.ClickFirstSubRecipeDatasheet();
            double weightBeforeeditAfter = recipesVariantPage.GetNetWeight(decimalSeparatorValue);
            double pricebeforeeditAfter = recipesVariantPage.GetPrice(decimalSeparatorValue, currency);
            Assert.AreNotEqual(weightBeforeedit, weightBeforeeditAfter);
            Assert.AreNotEqual(pricebeforeedit, pricebeforeeditAfter);
        }

        [TestMethod]
        public void MENU_DATA_ModifYield()
        {
            string itemName = itemNameToday + "-" + new Random().Next().ToString();
            string datasheetNameToday = "Datasheet-" + DateUtils.Now.ToString("dd/MM/yyyy") + "-" + new Random().Next().ToString();
            string recipeNameToday = "recipe-" + DateUtils.Now.ToString("dd/MM/yyyy") + "-" + new Random().Next().ToString();

            //Prepare recipe
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            int nbPortions = new Random().Next(1, 30);

            //Prepare ingredient
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();

            //Arrange
            LogInAsAdmin();

            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();

            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
            itemPage = itemGeneralInformationPage.BackToList();
            itemPage.ResetFilter();

            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();

            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeNameToday, recipeType, nbPortions.ToString());

            recipeGeneralInfosPage.AddVariantWithSite(site, variant);
            var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

            recipeVariantPage.AddIngredient(itemName);

            recipesPage = recipeVariantPage.BackToList();
            recipesPage.ResetFilter();
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetNameToday, guestType, site);
            datasheetDetailPage.AddRecipe(recipeNameToday);

            recipesPage = recipeVariantPage.BackToList();
            recipesPage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
            ItemUseCasePage useCase = generalInfo.ClickOnUseCasePage();
            useCase.SelectUseCase(recipeNameToday);

            //Modifier le champ YIELD dans item
            useCase.MultipleUpdate("Net weight", "14");
            var netWeight = useCase.GetColumnFirstValue(ItemUseCasePage.ColumnName.NetWeight);
            Assert.AreEqual("14", netWeight, "multi update Net Weight KO");

            //retourner à Datasheet pour verifier la modification de champ YIELD
            recipesPage.GoToMenus_Datasheet();
            recipesPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetNameToday);
            datasheetDetailPage = datasheetPage.SelectFirstDatasheet();
            datasheetDetailPage.ClickOnFirstRecipe();
            string initNetWeight = datasheetDetailPage.GetNetWeight(decimalSeparatorValue).ToString();

            Assert.AreEqual(netWeight, initNetWeight, "La modification du Yield dans la vue Item n'est pas correctement reflétée.");


        }
        [TestMethod]
        public void MENU_DATA_Print_Image_Recette()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();
            string datasheetName = $"{datasheetNameToday}-{new Random().Next()}";
            string imagePath = Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase).Substring(6), "PageObjects", "Menus", "Datasheet", "test.jpg");
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNameZipBegin = "All_files_";
            string DocFileNamePdfBegin = "Datasheet-";

            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);
            datasheetDetailPage.AddRecipe(recipeName);
            datasheetDetailPage.WaitLoading();

            var recipeVariantPage = datasheetDetailPage.AddPictureToRecipe();
            recipeVariantPage.AddPicture(imagePath);
            datasheetDetailPage = recipeVariantPage.CloseWindow();
            DatasheetIntolerancePage datasheetIntolerancePage = datasheetDetailPage.ClickOnIntoleranceTab();
            var totalAlergen = datasheetIntolerancePage.CalculateAllergens();
            string originalWindow = WebDriver.CurrentWindowHandle;
            // 
            datasheetDetailPage.ClearDownloads();
            foreach (var file in new DirectoryInfo(downloadsPath).GetFiles())
            {
                file.Delete();
            }

            PrintReportPage reportPage = datasheetDetailPage.Print(true);
            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            datasheetDetailPage.CloseWindow(originalWindow);
            datasheetDetailPage.WaitPageLoading();
            string foundFile = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin, false);
            FileInfo fileInfo = new FileInfo(foundFile);
            fileInfo.Refresh();
            Assert.IsTrue(fileInfo.Exists, $"{foundFile} non généré");

            // Verify the image on the PDF
            PdfDocument document = PdfDocument.Open(fileInfo.FullName);
            Page page1 = document.GetPage(1);
            List<IPdfImage> images = page1.GetImages().ToList<IPdfImage>();
            var pdfImages = images.Count();
            var TotalNumberInPDF = totalAlergen + 2;
            Assert.AreEqual(TotalNumberInPDF, pdfImages, "Le nombre d'images dans le PDF n'est pas égal à 1.");
        }
        [TestMethod]
        public void MENU_DATA_SelectAll_AddRecipe()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();
            string datasheetName = datasheetNameToday + "-" + new Random().Next().ToString();
            string datasheetName1 = datasheetNameToday + "-_" + new Random().Next().ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //act

            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);
            datasheetDetailPage.AddRecipe(recipeName);
            datasheetDetailPage.WaitLoading();
            WebDriver.Navigate().Refresh();
            datasheetPage = datasheetDetailPage.BackToList();
            //Assert
            datasheetPage.ResetFilter();
            try
            {
                datasheetPage.CreateNewDatasheet();
                datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName1, guestType, site);
                datasheetDetailPage.AddRecipeFromDatasheetToshowRecipe(datasheetName);
                datasheetDetailPage.WaitLoading();
                var text = datasheetDetailPage.IsTextEqualNone();
                Assert.IsTrue(text, "Le texte de l'élément n'est pas 'none'");
            }
            finally
            {
                homePage.GoToMenus_Datasheet();
                var dsMassiveDelete = datasheetPage.OpenMassiveDeletePopup();

                dsMassiveDelete.SetDatasheetName(datasheetName);

                dsMassiveDelete.ClickOnSearch();
                dsMassiveDelete.ClickOnSelectAll();
                dsMassiveDelete.ClickDeleteButton();
                datasheetPage.OpenMassiveDeletePopup();
                dsMassiveDelete.SetDatasheetName(datasheetName1);
                dsMassiveDelete.ClickOnSearch();
                dsMassiveDelete.ClickOnSelectAll();

                dsMassiveDelete.ClickDeleteButton();
            }
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void MENU_DATA_PrixIngredient()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();
            string datasheetName = datasheetNameToday + "-" + new Random().Next().ToString();
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();

            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string supplierRef = ingredient + "_SupplierRef";
            string storageQty = 10.ToString();
            string qty = 10.ToString();
            string cookingMode = "Cocina Caliente";
            string netWeight = "3";

            //Arrange
            var homePage = LogInAsAdmin();

            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, ingredient);

            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(ingredient.ToString(), group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier, supplierRef);
                itemPage = itemGeneralInformationPage.BackToList();
                itemPage.ResetFilter();
                itemPage.Filter(ItemPage.FilterType.Search, ingredient);
            }
            var firstItemName = itemPage.GetFirstItemName();
            Assert.AreEqual(ingredient, firstItemName, "L'ingredient n'est pas présent dans la liste.");
            var datasheetPage = homePage.GoToMenus_Datasheet();
            try
            {
                var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
                var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);
                var datasheetCreateNewRecipePage = datasheetDetailPage.CreateNewRecipe();
                datasheetDetailPage = datasheetCreateNewRecipePage.FillFields_AddNewRecipeToDatasheet(recipeName, ingredient, cookingMode);
                var priceBeforeUpdate = datasheetDetailPage.GetPrice();
                var recipesVariantPage = datasheetDetailPage.EditRecipe();
                recipesVariantPage.WaitPageLoading();
                recipesVariantPage.EditNetWeight(netWeight);
                recipesVariantPage.CloseWindow();
                datasheetDetailPage.WaitPageLoading();
                var priceAfterUpdate = datasheetDetailPage.GetPrice();
                Assert.AreNotEqual(priceBeforeUpdate, priceAfterUpdate, "le prix n'est pas modifié");
            }
            finally
            {
                homePage.GoToMenus_Datasheet();
                var dsMassiveDelete = datasheetPage.OpenMassiveDeletePopup();

                dsMassiveDelete.SetDatasheetName(datasheetName);

                dsMassiveDelete.ClickOnSearch();
                dsMassiveDelete.ClickOnSelectAll();
                dsMassiveDelete.ClickDeleteButton();
            }
        }
        [TestMethod]
        public void MENU_DATA_GrammageSubRecipe()
        {
            Random rnd = new Random();
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string site = "ACE";
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeName = "RecipeForTest" + rnd.Next(1, 999).ToString();
            string subRecipeName = "supRecipeForTest" + rnd.Next(1, 100).ToString();
            int nbPortions = new Random().Next(1, 100);

            string datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();
            string variant = "FLIGHT - FLIGHT";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            RecipesPage recipesPage = homePage.GoToMenus_Recipes();
            RecipesCreateModalPage recipesModalPage = recipesPage.CreateNewRecipe();
            RecipeGeneralInformationPage recipeGeneralInfosPage = recipesModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());
            recipeGeneralInfosPage.SearchVariant(variant, site);
            recipeGeneralInfosPage.SelectFirstVariant();
            recipeGeneralInfosPage.Validate();
            DatasheetPage dataSheetPage = homePage.GoToMenus_Datasheet();
            DatasheetCreateModalPage datasheetCreateModalPage = dataSheetPage.CreateNewDatasheet();
            DatasheetDetailsPage datasheetDetailsPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);
            datasheetDetailsPage.AddRecipe(recipeName);

            Assert.IsTrue(datasheetDetailsPage.IsRecipeAdded(), "la recette n'a pas été rajoutée à la datasheet.");

            DatasheetCreateNewRecipePage recipeVariantPage = datasheetDetailsPage.AddSubrecipeForFirstRecipe();
            recipeVariantPage.FillFields_AddNewRecipeToDatasheet(subRecipeName);
            recipeVariantPage.AddIngredient("a");
            recipeVariantPage.AddIngredient("b");
            recipeVariantPage.Closebtn();
            recipeVariantPage.FIRSTLINECLICK();
            recipeVariantPage.FIRSTLINECLICK();
            recipeVariantPage.Editfirstline();
            float total_NetWeight = recipeVariantPage.GetTotalNetWeight();
            float first_NetWeight = recipeVariantPage.GetFirstNetWeight();
            float second_NetWeight = recipeVariantPage.GetSecondNetWeight();
            recipeVariantPage.ChangeFirstNetWeight("5");

            float total_NetWeight_after_update = recipeVariantPage.GetTotalNetWeight();

            Assert.AreNotEqual(total_NetWeight, total_NetWeight_after_update, "the total is not updated");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void RE_DATA_PrintEspacePictureRecipeName()
        {
            Random rnd = new Random();
            //Prepare recipe
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string recipeName = "recipeprintpicture";
            int nbPortions = new Random().Next(1, 30);

            //Prepare
            //string recipeName = TestContext.Properties["DatasheetRecipe"].ToString();
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            var customerICAO = "Cust198";
            bool versionPrint = true;
            //string site = TestContext.Properties["SiteMAD"].ToString();            
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNamePdfBegin = "DataSheets Report";
            string DocFileNameZipBegin = "All_files_";

            string dataSheetNameImage = "datasheetforprintimage";
            string datasheetName = dataSheetNameImage + DateUtils.Now.ToString("dd/MM/yyyy") + "-" + rnd.Next().ToString();
            var path = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase);
            path = path.Substring(6);
            var imagePath = path + "\\PageObjects\\Menus\\Datasheet\\test.jpg";
            DatasheetCreateModalPage datasheetCreateModalPage;
            DatasheetDetailsPage datasheetDetailPage;

            //Arrange
            var homePage = LogInAsAdmin();
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);
            if (recipesPage.GetTotalRcipes() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());
                recipeGeneralInfosPage.AddVariantWithSite(site, variant);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
                recipeVariantPage.AddIngredient(ingredient);
            }
            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site, customerICAO);
            datasheetDetailPage.AddRecipe(recipeName);
            datasheetDetailPage.WaitLoading();
            WebDriver.Navigate().Refresh();
            bool isRecipeAdded = datasheetDetailPage.IsRecipeAdded();
            Assert.IsTrue(isRecipeAdded, "La recette n'a pas été rajoutée à la datasheet.");
            datasheetDetailPage.ClickOnPictureTab();
            datasheetDetailPage.AddPicture(imagePath);
            bool isPictureAdded = datasheetDetailPage.IsPictureAdded();
            Assert.IsTrue(isPictureAdded, "L'image n'a pas été ajoutée à la datasheet.");
            datasheetPage = datasheetDetailPage.BackToList();
            datasheetPage.ResetFilter();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, dataSheetNameImage);
            var nombre_datasheet = datasheetPage.GetTotalDataSheets();
            Assert.IsFalse(nombre_datasheet == 0, "La datasheet avec l image ne sont pas créée.");
            try
            {
                if (nombre_datasheet > 0)
                {
                    datasheetPage.ClearDownloads();
                    foreach (var file in new DirectoryInfo(downloadsPath).GetFiles())
                    {
                        file.Delete();
                    }
                    var reportPage = datasheetPage.PrintDatasheet(versionPrint);
                    var isGenerated = reportPage.IsReportGenerated();
                    reportPage.Close();
                    Assert.IsTrue(isGenerated, "Le document PDF n'a pas pu être généré par l'application.");
                    reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
                    datasheetPage.WaitPageLoading();
                    string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin, false);
                    FileInfo filePdf = new FileInfo(trouve);
                    filePdf.Refresh();
                    Assert.IsTrue(filePdf.Exists, trouve + " non généré");

                    using (var pdfHelper = new PDFHelper(filePdf.FullName))
                    {
                        var imageDatasheet = pdfHelper.GetImages(1).FirstOrDefault();
                        var textRecipe = pdfHelper.GetTextByContent(recipeName.ToUpper(), 1);
                        if (imageDatasheet != null && textRecipe != null)
                        {
                            double distance = pdfHelper.CalculateDistance(imageDatasheet, textRecipe);
                            Assert.IsTrue(distance <= 200, "La distance entre l'image et le texte dépasse 200.");
                        }
                    }
                }
            }
            finally
            {
                datasheetPage = homePage.GoToMenus_Datasheet();
                var dsMassiveDelete = datasheetPage.OpenMassiveDeletePopup();
                dsMassiveDelete.SetDatasheetName(dataSheetNameImage);
                dsMassiveDelete.ClickOnSearch();
                dsMassiveDelete.SelectAllSites();
                dsMassiveDelete.ClickDeleteButton();
            }
        }

        [TestMethod]
        public void MENU_DATA_SuppressionRecipe_SubRecipe()
        {
            var site = TestContext.Properties["Site"].ToString();
            var guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            var recipeName = TestContext.Properties["DatasheetRecipe"].ToString();
            var subRecipeIngredient = TestContext.Properties["RecipeIngredientBis"].ToString();
            var rnd = new Random();
            var datasheetName = datasheetNameToday + "-" + rnd.Next().ToString();
            var subrecipeName = "RecipeTest-" + rnd.Next().ToString();

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var datasheetPage = homePage.GoToMenus_Datasheet();
            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);
            Assert.IsFalse(datasheetDetailPage.IsRecipeAdded(), "La datasheet possède déjà une recette.");

            datasheetDetailPage.AddRecipe(recipeName);
            Assert.IsTrue(datasheetDetailPage.IsRecipeAdded(), "La recette n'a pas été rajoutée à la datasheet.");

            var datasheetCreateNewRecipePage = datasheetDetailPage.AddSubrecipeForFirstRecipe();
            datasheetDetailPage = datasheetCreateNewRecipePage.FillFields_AddNewRecipeToDatasheet(subrecipeName, subRecipeIngredient);
            if (!datasheetDetailPage.isPopupClosed())
            {
                datasheetDetailPage.CloseDetailModal();
            }
            datasheetDetailPage.DeleteRecipe();
            datasheetDetailPage.WaitLoading();
            Assert.IsFalse(datasheetDetailPage.IsRecipeAdded(), "La recette n'a pas été supprimée de la datasheet.");

        }
    }
}