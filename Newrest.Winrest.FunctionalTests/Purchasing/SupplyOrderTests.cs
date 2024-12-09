using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Customer;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Recipes;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Production;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.ProductionManagement;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Item;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.SupplyOrder;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.Inventory;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.OutputForm;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Policy;
using System.Threading;
using System.Web.UI;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using static Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.SupplyOrderItem;
using static Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.SupplyOrderPage;
namespace Newrest.Winrest.FunctionalTests.Purchasing
{
    [TestClass]
    public class SupplyOrderTests : TestBase
    {
        private const int _timeout = 600000;
        private const string SUPPLY_ORDERS_EXCEL_SHEET_NAME = "Supply Orders";

        /// <summary>
        /// 
        /// Mise en place du paramétrage pour la configuration Winrest 4.0 
        /// 
        /// </summary>
        /// 
        [TestMethod]
        [Priority(0)]
        [Timeout(_timeout)]
        public void PU_SO_SetConfigWinrest4_0()
        {
            // Prepare
            string keyword = TestContext.Properties["Item_Keyword"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();

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

            // IsCompareActive et NumberOfItems
            homePage.SetIsCompareActiveValue(true);
            homePage.SetNumberItemToCompareValue("3");
            string numberItems = homePage.GetNumberItemToCompareValue();

            // New group display
            homePage.SetNewGroupDisplayValue(true);

            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();

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


            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
            createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
            var supplyOrderItemPage = createSupplyOrderPage.Submit();

            // vérifier new group display
            supplyOrderItemPage.ResetFilter();
            Assert.IsTrue(supplyOrderItemPage.IsGroupDisplayActive(), "Le paramètre 'NewGroupDisplay' n'est pas activé.");

            // vérifier isCompareActive et NumberOfItemsToCompare
            supplyOrderItemPage.ResetFilter();
            var itemName = supplyOrderItemPage.GetFirstItemName();
            Assert.IsTrue(supplyOrderItemPage.IsCompareActivated(itemName), "La checkbox 'Compare' n'a pas été trouvée, le IsCompareActive est inactif.");
            Assert.AreEqual("3", numberItems, "Le nombre d'items à comparer n'est pas égal à 3, le NumberOfItemsToCompare est inactif.");

        }

        [Priority(2)]
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_AddNewRecipeForSO()
        {
            // Recipe
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariant"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string recipeName = TestContext.Properties["RecipeName"].ToString();


            int nbPortions = 5;

            // Ingredient
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
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            ClearCache();

            // Vérifier que l'ingrédient existe
            var itemPage = homePage.GoToPurchasing_ItemPage();

            itemPage.Filter(ItemPage.FilterType.Search, ingredient);

            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(ingredient.ToString(), group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier, supplierRef);
            }

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

                recipeGeneralInfosPage.AddVariantWithSite(site, variant);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(ingredient);
                recipeVariantPage.BackToList();
                recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);
            }

            Assert.AreEqual(recipesPage.GetFirstRecipeName(), recipeName, "La recette n'a pas été créée.");
        }

        [Priority(3)]
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_AddNewRecipeForSOBis()
        {
            // Recipe
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariant"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string recipeName = TestContext.Properties["RecipeName1"].ToString();

            int nbPortions = 1;

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);

            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

                recipeGeneralInfosPage.AddVariantWithSite(site, variant);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(ingredient);
                recipeVariantPage.BackToList();
                recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);
            }

            Assert.AreEqual(recipesPage.GetFirstRecipeName(), recipeName, "La recette n'a pas été créée.");
        }

        // ___________________________________________________ Filtres _____________________________________________________
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SO_Filter_SearchByNumber()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.ResetFilter();

            var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
            createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
            var id = createSupplyOrderPage.GetSupplyOrderNumber();
            var supplyOrderItemPage = createSupplyOrderPage.Submit();

            var itemName = supplyOrderItemPage.GetFirstItemName();
            supplyOrderItemPage.SelectFirstItem();
            supplyOrderItemPage.AddQuantity(itemName, "10");

            supplyOrderPage = supplyOrderItemPage.BackToList();

            //Filter on number and select the first item on the list
            supplyOrderPage.ResetFilter();
            supplyOrderPage.Filter(FilterType.ByNumber, id);

            //Assert
            Assert.AreEqual(id, supplyOrderPage.GetFirstSONumber(), String.Format(MessageErreur.FILTRE_ERRONE, "'Search by Number'"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SO_Filter_SortByNumber()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();
            string number1 = "";
            string number2 = "";
            bool isCreated = false;
            string sortByNumber = "NUMBER";
            DateTime from = DateUtils.Now;
            DateTime to = DateUtils.Now.AddDays(30);
            DateTime deleveryDate = DateUtils.Now;
            // Arrange
            var homePage = LogInAsAdmin();
            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.ResetFilter();
            try
            {
                if (supplyOrderPage.CheckTotalNumber() < 20)
                {
                    // create supplier oreder 1
                    var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
                    createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, from, to, deleveryDate, deliveryLocation);
                    number1 = createSupplyOrderPage.GetSupplyOrderNumber();
                    SupplyOrderItem item = createSupplyOrderPage.Submit();
                    supplyOrderPage = item.BackToList();
                    // create supplier oreder 2
                    createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
                    createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, from, to, deleveryDate, deliveryLocation);
                    number2 = createSupplyOrderPage.GetSupplyOrderNumber();
                    item = createSupplyOrderPage.Submit();
                    supplyOrderPage = item.BackToList();
                    isCreated = true;
                }

                if (!supplyOrderPage.isPageSizeEqualsTo100())
                {
                    supplyOrderPage.PageSize("8");
                    supplyOrderPage.PageSize("100");
                }

                supplyOrderPage.Filter(FilterType.SortBy, sortByNumber);
                var isSortedByNumber = supplyOrderPage.IsSortedByNumber();
                //Assert
                Assert.IsTrue(isSortedByNumber, String.Format(MessageErreur.FILTRE_ERRONE, "Sort by Number"));
            }
            finally
            {
                if(isCreated)
                {
                    //delete supplier oreder 1
                    supplyOrderPage.Filter(FilterType.ByNumber, number1);
                    supplyOrderPage.DeleteFirstSupplyOrder();
                    //delete supplier oreder 2
                    supplyOrderPage.Filter(FilterType.ByNumber, number2);
                    supplyOrderPage.DeleteFirstSupplyOrder();
                }
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SO_Filter_SortByStartDate()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();
            string number1 = "";
            string number2 = "";
            bool isCreated = false;
            string sortByStartDate = "DATE";
            DateTime from1 = DateUtils.Now;
            DateTime from2 = DateUtils.Now.AddDays(1);
            DateTime to = DateUtils.Now.AddDays(30);
            DateTime deleveryDate = DateUtils.Now;
            // Arrange
            var homePage = LogInAsAdmin();
            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.ResetFilter();
            try
            {
                if (supplyOrderPage.CheckTotalNumber() < 20)
                {
                    // create supplier oreder 1
                    var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
                    createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, from1, to, deleveryDate, deliveryLocation);
                    number1 = createSupplyOrderPage.GetSupplyOrderNumber();
                    SupplyOrderItem item = createSupplyOrderPage.Submit();
                    supplyOrderPage = item.BackToList();
                    // create supplier oreder 2
                    createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
                    createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, from2, to, deleveryDate, deliveryLocation);
                    number2 = createSupplyOrderPage.GetSupplyOrderNumber();
                    item = createSupplyOrderPage.Submit();
                    supplyOrderPage = item.BackToList();
                    isCreated = true;
                }

                if (!supplyOrderPage.isPageSizeEqualsTo100())
                {
                    supplyOrderPage.PageSize("8");
                    supplyOrderPage.PageSize("100");
                }

                supplyOrderPage.Filter(FilterType.SortBy, sortByStartDate);
                var isSortedByStartDate = supplyOrderPage.IsSortedByStartDate();
                //Assert
                Assert.IsTrue(isSortedByStartDate, String.Format(MessageErreur.FILTRE_ERRONE, "Sort by start date"));
            }
            finally
            {
                if (isCreated)
                {
                    //delete supplier oreder 1
                    supplyOrderPage.Filter(FilterType.ByNumber, number1);
                    supplyOrderPage.DeleteFirstSupplyOrder();
                    //delete supplier oreder 2
                    supplyOrderPage.Filter(FilterType.ByNumber, number2);
                    supplyOrderPage.DeleteFirstSupplyOrder();
                }
            }
        }


        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SO_Filter_SortByEndDate()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();
            string number1 = "";
            string number2 = "";
            bool isCreated = false;
            string sortByEndDate = "ENDDATE";
            DateTime from = DateUtils.Now;
            DateTime to1 = DateUtils.Now.AddDays(30);
            DateTime to2 = DateUtils.Now.AddDays(30);
            DateTime deleveryDate = DateUtils.Now;
            // Arrange
            var homePage = LogInAsAdmin();
            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.ResetFilter();
            try
            {
                if (supplyOrderPage.CheckTotalNumber() < 20)
                {
                    // create supplier oreder 1
                    var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
                    createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, from, to1, deleveryDate, deliveryLocation);
                    number1 = createSupplyOrderPage.GetSupplyOrderNumber();
                    SupplyOrderItem item = createSupplyOrderPage.Submit();
                    supplyOrderPage = item.BackToList();
                    // create supplier oreder 2
                    createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
                    createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, from, to2, deleveryDate, deliveryLocation);
                    number2 = createSupplyOrderPage.GetSupplyOrderNumber();
                    item = createSupplyOrderPage.Submit();
                    supplyOrderPage = item.BackToList();
                    isCreated = true;
                }

                if (!supplyOrderPage.isPageSizeEqualsTo100())
                {
                    supplyOrderPage.PageSize("8");
                    supplyOrderPage.PageSize("100");
                }

                supplyOrderPage.Filter(FilterType.SortBy, sortByEndDate);
                var isSortedByEndDate = supplyOrderPage.IsSortedByEndDate();
                //Assert
                Assert.IsTrue(isSortedByEndDate, String.Format(MessageErreur.FILTRE_ERRONE, "Sort by end Date"));
            }
            finally
            {
                if (isCreated)
                {
                    //delete supplier oreder 1
                    supplyOrderPage.Filter(FilterType.ByNumber, number1);
                    supplyOrderPage.DeleteFirstSupplyOrder();
                    //delete supplier oreder 2
                    supplyOrderPage.Filter(FilterType.ByNumber, number2);
                    supplyOrderPage.DeleteFirstSupplyOrder();
                }
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SO_Filter_ShowItemsNotValidated()
        {
            //Prepare          
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.ResetFilter();
            supplyOrderPage.Filter(FilterType.ShowNotValidated, true);

            if (supplyOrderPage.CheckTotalNumber() < 20)
            {
                var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
                createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation, true);
                var supplyOrderItemPage = createSupplyOrderPage.Submit();

                var itemName = supplyOrderItemPage.GetFirstItemName();
                supplyOrderItemPage.SelectFirstItem();
                supplyOrderItemPage.AddQuantity(itemName, "10");

                supplyOrderPage = supplyOrderItemPage.BackToList();

                supplyOrderPage.ResetFilter();
                supplyOrderPage.Filter(FilterType.ShowNotValidated, true);
            }

            if (!supplyOrderPage.isPageSizeEqualsTo100())
            {
                supplyOrderPage.PageSize("8");
                supplyOrderPage.PageSize("100");
            }

            //Assert
            bool isItemNotValidated = supplyOrderPage.CheckValidation(false);
            Assert.IsFalse(isItemNotValidated, String.Format(MessageErreur.FILTRE_ERRONE, "'Show items not validated'"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SO_Filter_ShowItemsValidated()
        {
            //Prepare          
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.ResetFilter();
            supplyOrderPage.Filter(FilterType.ShowNotValidated, false);

            if (supplyOrderPage.CheckTotalNumber() < 20)
            {
                var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
                createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation, true);
                var supplyOrderItemPage = createSupplyOrderPage.Submit();

                var itemName = supplyOrderItemPage.GetFirstItemName();
                supplyOrderItemPage.SelectFirstItem();
                supplyOrderItemPage.AddQuantity(itemName, "10");

                supplyOrderItemPage.Validate();

                supplyOrderPage = supplyOrderItemPage.BackToList();

                supplyOrderPage.ResetFilter();
                supplyOrderPage.Filter(FilterType.ShowNotValidated, false);
            }

            if (!supplyOrderPage.isPageSizeEqualsTo100())
            {
                supplyOrderPage.PageSize("8");
                supplyOrderPage.PageSize("100");
            }

            //Assert
            bool isValidatedItemsOK = supplyOrderPage.CheckValidation(true);
            Assert.IsTrue(isValidatedItemsOK, String.Format(MessageErreur.FILTRE_ERRONE, "'Show items validated'"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SO_Filter_ShowAll()
        {
            //Prepare          
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.ResetFilter();

            if (supplyOrderPage.CheckTotalNumber() < 20)
            {
                //Create
                var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
                createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation, false);
                var supplyOrderItemPage = createSupplyOrderPage.Submit();

                var itemName = supplyOrderItemPage.GetFirstItemName();
                supplyOrderItemPage.SelectFirstItem();
                supplyOrderItemPage.AddQuantity(itemName, "10");

                supplyOrderPage = supplyOrderItemPage.BackToList();
                supplyOrderPage.ResetFilter();
            }

            supplyOrderPage.Filter(FilterType.ShowActive, true);
            int nbActive = supplyOrderPage.CheckTotalNumber();

            supplyOrderPage.Filter(FilterType.ShowNotActive, true);
            int nbInactive = supplyOrderPage.CheckTotalNumber();

            supplyOrderPage.Filter(FilterType.ShowAll, true);
            int nbTotal = supplyOrderPage.CheckTotalNumber();

            //Assert
            Assert.AreEqual((nbActive + nbInactive), nbTotal, String.Format(MessageErreur.FILTRE_ERRONE, "'Show all'"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SO_Filter_ShowActive()
        {
            //Prepare          
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.ResetFilter();
            supplyOrderPage.Filter(FilterType.ShowActive, true);

            if (supplyOrderPage.CheckTotalNumber() < 20)
            {
                //Create
                var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
                createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation, false);
                var supplyOrderItemPage = createSupplyOrderPage.Submit();

                var itemName = supplyOrderItemPage.GetFirstItemName();
                supplyOrderItemPage.SelectFirstItem();
                supplyOrderItemPage.AddQuantity(itemName, "10");

                supplyOrderPage = supplyOrderItemPage.BackToList();

                supplyOrderPage.ResetFilter();
                supplyOrderPage.Filter(FilterType.ShowActive, true);
            }

            if (!supplyOrderPage.isPageSizeEqualsTo100())
            {
                supplyOrderPage.PageSize("8");
                supplyOrderPage.PageSize("100");
            }

            //Assert
            bool showItemActiveOK = supplyOrderPage.CheckStatus(true);
            Assert.IsTrue(showItemActiveOK, String.Format(MessageErreur.FILTRE_ERRONE, "'Show only active'"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SO_Filter_ShowInactive()
        {
            //Prepare          
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.ResetFilter();
            supplyOrderPage.Filter(FilterType.ShowNotActive, true);

            if (supplyOrderPage.CheckTotalNumber() < 20)
            {
                //Create
                var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
                createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation, false);
                var supplyOrderItemPage = createSupplyOrderPage.Submit();

                var itemName = supplyOrderItemPage.GetFirstItemName();
                supplyOrderItemPage.SelectFirstItem();
                supplyOrderItemPage.AddQuantity(itemName, "10");

                supplyOrderPage = supplyOrderItemPage.BackToList();

                supplyOrderPage.ResetFilter();
                supplyOrderPage.Filter(FilterType.ShowNotActive, true);
            }

            if (!supplyOrderPage.isPageSizeEqualsTo100())
            {
                supplyOrderPage.PageSize("8");
                supplyOrderPage.PageSize("100");
            }

            //Assert
            Assert.IsFalse(supplyOrderPage.CheckStatus(false), String.Format(MessageErreur.FILTRE_ERRONE, "'Show only inactive'"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SO_Filter_DateFrom()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();


            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.ResetFilter();
            DateTime fromDate = DateTime.Now.Date;
            DateTime fromDateInf = fromDate.AddDays(-1);
            DateTime fromDateSup = fromDate.AddDays(2);
            string ID1 = "";
            string ID2 = "";
            string ID3 = "";

            try
            {
                //Create Supply Order 1
                var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
                createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, fromDate , DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation, true);
                ID1 = createSupplyOrderPage.GetFirstItemId();
                var supplyOrderItemPage = createSupplyOrderPage.Submit();
                var itemName = supplyOrderItemPage.GetFirstItemName();
                supplyOrderItemPage.SelectFirstItem();
                supplyOrderItemPage.AddQuantity(itemName, "10");
                supplyOrderPage = supplyOrderItemPage.BackToList();
                supplyOrderPage.ResetFilter();


                //Create Supply Order 2
                createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
                createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, fromDateInf, DateUtils.Now.AddDays(30), fromDateInf , deliveryLocation, true);
                ID2 = createSupplyOrderPage.GetFirstItemId();
                supplyOrderItemPage = createSupplyOrderPage.Submit();
                var itemName2 = supplyOrderItemPage.GetFirstItemName();
                supplyOrderItemPage.SelectFirstItem();
                supplyOrderItemPage.AddQuantity(itemName2, "20");
                supplyOrderPage = supplyOrderItemPage.BackToList();
                supplyOrderPage.ResetFilter();
             

                //Create Supply Order 3
                createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
                createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, fromDateSup, DateUtils.Now.AddDays(30), fromDateSup, deliveryLocation, true);
                ID3 = createSupplyOrderPage.GetFirstItemId();
                supplyOrderItemPage = createSupplyOrderPage.Submit();
                var itemName3 = supplyOrderItemPage.GetFirstItemName();
                supplyOrderItemPage.SelectFirstItem();
                supplyOrderItemPage.AddQuantity(itemName3, "30");
                supplyOrderPage = supplyOrderItemPage.BackToList();
                supplyOrderPage.ResetFilter();
                

                //Filter and Assert 
               supplyOrderPage.Filter(FilterType.DateFrom, fromDate);
                supplyOrderPage.Filter(FilterType.ByNumber, ID1);
                Assert.AreEqual(supplyOrderPage.CheckTotalNumber(), 1, "Le filtre Date From ne fonctionne pas correctement .");
                supplyOrderPage.Filter(FilterType.ByNumber, ID2);
                Assert.AreEqual(supplyOrderPage.CheckTotalNumber(), 0, "Le filtre Date From ne fonctionne pas correctement .");
                supplyOrderPage.Filter(FilterType.ByNumber, ID3);
                Assert.AreEqual(supplyOrderPage.CheckTotalNumber(), 1, "Le filtre Date From ne fonctionne pas correctement .");
                supplyOrderPage.ResetFilter();
            }
            finally
            {
                //Delete
                supplyOrderPage.Filter(FilterType.ByNumber, ID1);
                supplyOrderPage.DeleteFirstSupplyOrder();

                supplyOrderPage.Filter(FilterType.ByNumber, ID2);
                supplyOrderPage.DeleteFirstSupplyOrder();

                supplyOrderPage.Filter(FilterType.ByNumber, ID3);
                supplyOrderPage.DeleteFirstSupplyOrder();
            }

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SO_Filter_DateTo()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();


            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.ResetFilter();
            DateTime DateTo = DateTime.Now.Date;
            DateTime fromDateInf = DateTo.AddDays(-1);
            DateTime fromDateSup = DateTo.AddDays(2);
            string ID1 = "";
            string ID2 = "";
            string ID3 = "";

            try
            {
                //Create Supply Order 1
                var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
                createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateTo, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation, true);
                ID1 = createSupplyOrderPage.GetFirstItemId();
                var supplyOrderItemPage = createSupplyOrderPage.Submit();
                var itemName = supplyOrderItemPage.GetFirstItemName();
                supplyOrderItemPage.SelectFirstItem();
                supplyOrderItemPage.AddQuantity(itemName, "10");
                supplyOrderPage = supplyOrderItemPage.BackToList();
                supplyOrderPage.ResetFilter();


                //Create Supply Order 2
                createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
                createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, fromDateInf, DateUtils.Now.AddDays(30), fromDateInf, deliveryLocation, true);
                ID2 = createSupplyOrderPage.GetFirstItemId();
                supplyOrderItemPage = createSupplyOrderPage.Submit();
                var itemName2 = supplyOrderItemPage.GetFirstItemName();
                supplyOrderItemPage.SelectFirstItem();
                supplyOrderItemPage.AddQuantity(itemName2, "20");
                supplyOrderPage = supplyOrderItemPage.BackToList();
                supplyOrderPage.ResetFilter();


                //Create Supply Order 3
                createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
                createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, fromDateSup, DateUtils.Now.AddDays(30), fromDateSup, deliveryLocation, true);
                ID3 = createSupplyOrderPage.GetFirstItemId();
                supplyOrderItemPage = createSupplyOrderPage.Submit();
                var itemName3 = supplyOrderItemPage.GetFirstItemName();
                supplyOrderItemPage.SelectFirstItem();
                supplyOrderItemPage.AddQuantity(itemName3, "30");
                supplyOrderPage = supplyOrderItemPage.BackToList();
                supplyOrderPage.ResetFilter();


                //Filter and Assert 
                supplyOrderPage.Filter(FilterType.DateTo, DateTo);
                supplyOrderPage.Filter(FilterType.ByNumber, ID1);
                Assert.AreEqual(supplyOrderPage.CheckTotalNumber(), 1, "Le filtre Date To ne fonctionne pas correctement .");
                supplyOrderPage.Filter(FilterType.ByNumber, ID2);
                Assert.AreEqual(supplyOrderPage.CheckTotalNumber(), 1, "Le filtre Date To ne fonctionne pas correctement .");
                supplyOrderPage.Filter(FilterType.ByNumber, ID3);
                Assert.AreEqual(supplyOrderPage.CheckTotalNumber(), 0, "Le filtre Date To ne fonctionne pas correctement .");
                supplyOrderPage.ResetFilter();
            }
            finally
            {
                //Delete
                supplyOrderPage.Filter(FilterType.ByNumber, ID1);
                supplyOrderPage.DeleteFirstSupplyOrder();

                supplyOrderPage.Filter(FilterType.ByNumber, ID2);
                supplyOrderPage.DeleteFirstSupplyOrder();

                supplyOrderPage.Filter(FilterType.ByNumber, ID3);
                supplyOrderPage.DeleteFirstSupplyOrder();
            }

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SO_Filter_ByValidationDate()
        {
            // Prepare
            string site = TestContext.Properties["SiteLP"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.ResetFilter();
            supplyOrderPage.Filter(FilterType.DateFrom, DateUtils.Now.AddDays(-15));
            supplyOrderPage.Filter(FilterType.DateTo, DateUtils.Now.AddDays(+10));
            supplyOrderPage.Filter(FilterType.ByValidationDate, true);

            if (supplyOrderPage.CheckTotalNumber() < 20)
            {
                // Create
                var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
                createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
                var supplyOrderItemPage = createSupplyOrderPage.Submit();

                string itemName = supplyOrderItemPage.GetFirstItemName();
                supplyOrderItemPage.SelectFirstItem();
                supplyOrderItemPage.AddQuantity(itemName, "10");

                supplyOrderItemPage.Validate();

                supplyOrderPage = supplyOrderItemPage.BackToList();

                supplyOrderPage.ResetFilter();
                supplyOrderPage.Filter(FilterType.DateFrom, DateUtils.Now.AddDays(-15));
                supplyOrderPage.Filter(FilterType.DateTo, DateUtils.Now.AddDays(+10));
                supplyOrderPage.Filter(FilterType.ByValidationDate, true);
            }

            if (!supplyOrderPage.isPageSizeEqualsTo100())
            {
                supplyOrderPage.PageSize("8");
                supplyOrderPage.PageSize("100");
            }

            bool isValidationCorrect = supplyOrderPage.CheckValidation(true);
            bool isFromDateRespected = supplyOrderPage.IsFromDateRespected(DateUtils.Now.AddDays(-15));
            bool isToDateRespected = supplyOrderPage.IsToDateRespected(DateUtils.Now.AddDays(+10));

            Assert.IsTrue(isValidationCorrect, MessageErreur.FILTRE_ERRONE, "By validation date");
            Assert.IsTrue(isFromDateRespected, String.Format(MessageErreur.FILTRE_ERRONE, "From"));
            Assert.IsTrue(isToDateRespected, String.Format(MessageErreur.FILTRE_ERRONE, "To"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SO_Filter_Site()
        {
            //Prepare          
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.ResetFilter();
            supplyOrderPage.Filter(FilterType.Site, site);

            if (supplyOrderPage.CheckTotalNumber() < 20)
            {
                //Create
                var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
                createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation, true);
                var supplyOrderItemPage = createSupplyOrderPage.Submit();

                var itemName = supplyOrderItemPage.GetFirstItemName();
                supplyOrderItemPage.SelectFirstItem();
                supplyOrderItemPage.AddQuantity(itemName, "10");

                supplyOrderPage = supplyOrderItemPage.BackToList();

                supplyOrderPage.ResetFilter();
                supplyOrderPage.Filter(FilterType.Site, site);
            }

            if (!supplyOrderPage.isPageSizeEqualsTo100())
            {
                supplyOrderPage.PageSize("8");
                supplyOrderPage.PageSize("100");
            }

            //Assert
            bool isFiltreSiteOK = supplyOrderPage.VerifySite(site);
            Assert.IsTrue(isFiltreSiteOK, String.Format(MessageErreur.FILTRE_ERRONE, "Les résultats ne s'accordent pas bien au filtre appliqué su le site."));
        }

        // ___________________________________________________ Fin Filtres _________________________________________________

        [TestMethod]
        [Priority(1)]
        [Timeout(_timeout)]
        public void PU_SO_CreateNewSupplyOrder()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();

            // Create
            var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
            createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
            var id = createSupplyOrderPage.GetSupplyOrderNumber();
            var supplyOrderItemPage = createSupplyOrderPage.Submit();

            var itemName = supplyOrderItemPage.GetFirstItemName();
            supplyOrderItemPage.SelectFirstItem();
            supplyOrderItemPage.AddQuantity(itemName, "10");
            supplyOrderPage = supplyOrderItemPage.BackToList();
            supplyOrderPage.ResetFilter();
            supplyOrderPage.Filter(FilterType.ByNumber, id);

            //Assert
            Assert.AreEqual(id, supplyOrderPage.GetFirstSONumber(), "Le supply order créé n'apparaît pas dans la liste des supply orders.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SO_UpdateGeneralInformation()
        {
            //Prepare 
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();

            // Create
            var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
            createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
            var id = createSupplyOrderPage.GetSupplyOrderNumber();
            var supplyOrderItemPage = createSupplyOrderPage.Submit();

            // Update fields
            var supplyOrderGeneralInformation = supplyOrderItemPage.ClickOnGeneralInformation();
            supplyOrderGeneralInformation.DeliveryDateUpdate(DateUtils.Now.AddDays(3));
            supplyOrderGeneralInformation.FromDateUpdate(DateUtils.Now.AddDays(-7));
            supplyOrderGeneralInformation.ToDateUpdate(DateUtils.Now.AddDays(14));
            //update checkboxes
            //based on main supplier set false

            supplyOrderGeneralInformation.SetBasedOnMainSupplier(false);
            //round prefill quantities when needed set true
            supplyOrderGeneralInformation.SetRoundPrefilledQuantities(true);
            //wait for the update
            Thread.Sleep(2000);

            supplyOrderPage = supplyOrderGeneralInformation.BackToList();
            supplyOrderPage.ResetFilter();
            supplyOrderPage.Filter(FilterType.ByNumber, id);
            supplyOrderItemPage = supplyOrderPage.SelectFirstItem();
            supplyOrderGeneralInformation = supplyOrderItemPage.ClickOnGeneralInformation();

            //Assert
            Assert.AreEqual(false, supplyOrderGeneralInformation.GetBasedOnMainSupplierValue(), "le checkbox n'a pas été mis a jour");
            Assert.AreEqual(true, supplyOrderGeneralInformation.GetRoundPrefiledQuantitiesValue(), "le checkbox n'a pas été mis a jour");
            Assert.AreEqual(DateUtils.Now.AddDays(3).ToShortDateString(), supplyOrderGeneralInformation.GetDeliveryDateValue(), "La delivery date n'a pas été modifiée.");
            Assert.AreEqual(DateUtils.Now.AddDays(-7).ToShortDateString(), supplyOrderGeneralInformation.GetFromDateValue(), "La date de début n'a pas été modifiée.");
            Assert.AreEqual(DateUtils.Now.AddDays(14).ToShortDateString(), supplyOrderGeneralInformation.GetToDateValue(), "La date de fin n'a pas été modifiée");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SO_BySuppliers()
        {
            //Prepare
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            //Prepare
            string itemName = "itemSupplyOrder" + new Random().Next().ToString();
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
            HomePage homePage = LogInAsAdmin();
            try
            {
                // Création d'un item pour la supply order
                var itemPage = homePage.GoToPurchasing_ItemPage();
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier, null);

                //Act
                var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
                var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
                createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
                var supplyOrderItemPage = createSupplyOrderPage.Submit();
                supplyOrderItemPage.Filter(FilterSupplyItemType.ByName, itemName);
                supplyOrderItemPage.SelectFirstItem();
                supplyOrderItemPage.AddQuantity(itemName, "10");
                var bySuppliersPage = supplyOrderItemPage.ClickOnBySuppliers();
                Assert.IsTrue(bySuppliersPage.HasSuppliers(), "Aucun supplier n'est affiché dans la page BySuppliers.");
                Assert.IsTrue(bySuppliersPage.IsSupplierPresent(supplier), "Le supplier associé à la supply order n'est pas présent dans la liste.");
            }
            finally
            {
                var itemPage = homePage.GoToPurchasing_ItemPage();
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                var itemGeneralInformationPageDetails = itemPage.ClickOnFirstItem();
                itemGeneralInformationPageDetails.ClickOnDeleteItem();
                itemGeneralInformationPageDetails.ConfirmDelete();
                itemPage = itemGeneralInformationPageDetails.BackToList();
                var massiveDeleteModal = itemPage.MenuMassiveDelete();
                massiveDeleteModal.Filter(MassiveDeleteModal.FilterType.SearchByName, itemName);
                massiveDeleteModal.ClickSearch();
                massiveDeleteModal.ClickSelectAllButton();
                massiveDeleteModal.Delete();
            }
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SO_ExportSupplyOrderList_NewVersion()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();
            bool newVersionPrint = true;

            //Arrange
            var homePage = LogInAsAdmin();

            // Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();

            supplyOrderPage.ClearDownloads();

            supplyOrderPage.ResetFilter();
            supplyOrderPage.Filter(FilterType.DateFrom, DateUtils.Now);
            supplyOrderPage.Filter(FilterType.DateTo, DateUtils.Now.AddDays(+1));

            if (supplyOrderPage.CheckTotalNumber() < 20)
            {
                // Create
                var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
                createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
                var supplyOrderItemPage = createSupplyOrderPage.Submit();

                var itemName = supplyOrderItemPage.GetFirstItemName();
                supplyOrderItemPage.SelectFirstItem();
                supplyOrderItemPage.AddQuantity(itemName, "10");
                supplyOrderItemPage.BackToList();

                supplyOrderPage.ResetFilter();
                supplyOrderPage.Filter(FilterType.DateFrom, DateUtils.Now);
                supplyOrderPage.Filter(FilterType.DateTo, DateUtils.Now.AddDays(+1));
            }

            supplyOrderPage.Export(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = supplyOrderPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            int resultNumber = OpenXmlExcel.GetExportResultNumber(SUPPLY_ORDERS_EXCEL_SHEET_NAME, filePath);

            // On vide le répertoire de téléchargement
            DeleteAllFileDownload();

            //Assert
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
        }

        [DeploymentItem("Resources\\logo.jpg")]
        [DeploymentItem("Resources\\logo_petit.jpg")]
        [DeploymentItem("chromedriver.exe")]
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_PrintSupplyOrderList_NewVersion()
        {

            //Prepare
            string DocFileNamePdfBegin = "Print Reports_-_";
            string DocFileNameZipBegin = "All_files_";
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();

            bool newVersionPrint = true;

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.ResetFilter();

            supplyOrderPage.ClearDownloads();

            supplyOrderPage.Filter(FilterType.DateFrom, DateUtils.Now.AddDays(-1));
            supplyOrderPage.Filter(FilterType.DateTo, DateUtils.Now.AddDays(+1));
            supplyOrderPage.Filter(FilterType.ByValidationDate, true);

            if (supplyOrderPage.CheckTotalNumber() < 20)
            {
                var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
                createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
                var supplyOrderItemPage = createSupplyOrderPage.Submit();

                string itemName = supplyOrderItemPage.GetFirstItemName();
                supplyOrderItemPage.Filter(FilterSupplyItemType.ByName, itemName);
                supplyOrderItemPage.SelectFirstItem();
                supplyOrderItemPage.AddQuantity(itemName, "10");

                supplyOrderItemPage.Validate();

                supplyOrderPage = supplyOrderItemPage.BackToList();

                supplyOrderPage.ResetFilter();
                supplyOrderPage.Filter(FilterType.DateFrom, DateUtils.Now.AddDays(-1));
                supplyOrderPage.Filter(FilterType.DateTo, DateUtils.Now.AddDays(+1));
                supplyOrderPage.Filter(FilterType.ByValidationDate, true);
            }
            //get list of numbers
            var numbers = supplyOrderPage.GetSupplyOrdersNumbersList();
            //clear download folder
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();
            foreach (FileInfo file in taskFiles)
            {
                file.Delete();
            }
            var reportPage = supplyOrderPage.PrintResults(newVersionPrint);
            //supplyOrderPage.DownloadReport();
            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            // get downloaded pdf file
            string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            Assert.IsNotNull(trouve, MessageErreur.FICHIER_NON_TROUVE);
            // verifier données sur IHM et sur pdf
            foreach (var number in numbers)
            {
                var iscorrect = supplyOrderPage.VerifyPdf(trouve, number);
                Assert.IsTrue(iscorrect, "les données ne sont pas présentes dans le fichier pdf");
            }

            // 2.Verifier le logo
            PdfDocument document = PdfDocument.Open(trouve);
            var page1 = document.GetPage(1);
            List<IPdfImage> images = page1.GetImages().ToList<IPdfImage>();
            Assert.AreEqual(1, images.Count, "Pas d'image dans le PDF");
            IPdfImage image = images[0];

            FileInfo fiPdf = new FileInfo(downloadsPath + "\\logo_test.jpg");
            if (fiPdf.Exists)
            {
                fiPdf.Delete();
            }
            File.WriteAllBytes(downloadsPath + "\\logo_test.jpg", image.RawBytes.ToArray<byte>());
            fiPdf.Refresh();
            Assert.IsTrue(fiPdf.Exists, "fichier non trouvé");

            FileInfo fiTest = new FileInfo(TestContext.TestDeploymentDir + "\\logo_petit.jpg");
            Assert.IsTrue(fiTest.Exists, "fichier non trouvé (cas 2)");

            // comparaison fichiers
            var md5 = MD5.Create();
            var fsTest = File.Open(fiTest.FullName, FileMode.Open);
            var fsPdf = File.Open(fiPdf.FullName, FileMode.Open);
            var fsTest_Hash = System.Convert.ToBase64String(md5.ComputeHash(fsTest));
            var fsPdf_Hash = System.Convert.ToBase64String(md5.ComputeHash(fsPdf));
            Assert.AreEqual(System.Convert.ToBase64String(md5.ComputeHash(fsTest)), System.Convert.ToBase64String(md5.ComputeHash(fsPdf)), "fichiers différents");

        }

        //Voir la liste des items du supply Order
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_Items_Export_And_Import()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceToPartielle"].ToString();
            string itemName = TestContext.Properties["RecipeIngredient"].ToString();

            string order = "20";

            bool newVersionPrint = true;

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();

            supplyOrderPage.ClearDownloads();

            // Create
            var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
            createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
            var supplyOrderItemPage = createSupplyOrderPage.Submit();

            supplyOrderItemPage.Filter(FilterSupplyItemType.ByName, itemName);
            supplyOrderItemPage.SelectFirstItem();
            supplyOrderItemPage.AddQuantity(itemName, "10");

            WebDriver.Navigate().Refresh();
            var initValue = supplyOrderItemPage.GetFirstItemQty();
            supplyOrderItemPage.ExportItems(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = supplyOrderItemPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            OpenXmlExcel.WriteDataInColumn("Order", SUPPLY_ORDERS_EXCEL_SHEET_NAME, filePath, order, CellValues.Number);

            var importPopup = supplyOrderItemPage.Import();
            importPopup.ImportFile(correctDownloadedFile.FullName);

            supplyOrderItemPage.Filter(FilterSupplyItemType.ByName, itemName);
            var newValue = supplyOrderItemPage.GetFirstItemQty();

            //ASSERT
            Assert.AreNotEqual(initValue, newValue, "La valeur n'a pas été modifiée par l'import.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_DeleteSupplyOrder()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.ResetFilter();

            // Create
            var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
            createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
            var id = createSupplyOrderPage.GetSupplyOrderNumber();
            var supplyOrderItemPage = createSupplyOrderPage.Submit();

            supplyOrderPage = supplyOrderItemPage.BackToList();

            supplyOrderPage.Filter(FilterType.ByNumber, id);
            Assert.AreEqual(1, supplyOrderPage.CheckTotalNumber(), "Le supply order créé n'apparaît pas dans la liste des supply orders.");

            supplyOrderPage.DeleteFirstSupplyOrder();

            supplyOrderPage.ResetFilter();
            supplyOrderPage.Filter(FilterType.ByNumber, id);
            Assert.AreEqual(0, supplyOrderPage.CheckTotalNumber(), "Le supply order n'a pas été supprimé.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SO_Items_Filter_SearchByName()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();
            SupplyOrderItem supplyOrderItemPage;

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.ResetFilter();

            // Create
            var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
            createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
            supplyOrderItemPage = createSupplyOrderPage.Submit();

            var firstItemName = supplyOrderItemPage.GetFirstItemName();
            supplyOrderItemPage.Filter(FilterSupplyItemType.ByName, firstItemName);

            Assert.IsTrue(supplyOrderItemPage.VerifyName(firstItemName), String.Format(MessageErreur.FILTRE_ERRONE, "'Search by name/ref'"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SO_Items_Filter_SearchByKeyword()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();
            string itemKeyword = TestContext.Properties["Item_Keyword"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();

            // Create
            var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
            createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
            var supplyOrderItemPage = createSupplyOrderPage.Submit();

            // Ajout du keyword
            var itemName = supplyOrderItemPage.GetFirstItemName();

            var itemPageItem = supplyOrderItemPage.EditItem(itemName);
            var itemKeywordTab = itemPageItem.ClickOnKeywordItem();
            itemKeywordTab.AddKeyword(itemKeyword);
            itemPageItem.Close();

            supplyOrderItemPage.PageSize("8");
            supplyOrderItemPage.Filter(FilterSupplyItemType.ByKeyword, itemKeyword);

            Assert.IsTrue(supplyOrderItemPage.VerifyKeyword(itemKeyword), String.Format(MessageErreur.FILTRE_ERRONE, "'Keyword'"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_Items_Filter_ShowItemsNotSupplied()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();

            // Create
            var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
            createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
            var supplyOrderItemPage = createSupplyOrderPage.Submit();

            supplyOrderItemPage.ResetFilter();
            string itemName = supplyOrderItemPage.GetFirstItemName();
            supplyOrderItemPage.Filter(FilterSupplyItemType.ByName, itemName);
            supplyOrderItemPage.SelectFirstItem();
            supplyOrderItemPage.AddQuantity(itemName, "10");

            supplyOrderItemPage.ResetFilter();
            supplyOrderItemPage.Filter(FilterSupplyItemType.ShowItemsNotSupplied, false);

            Assert.IsTrue(supplyOrderItemPage.VerifySupplied(), String.Format(MessageErreur.FILTRE_ERRONE, "'Show items supplied'"));

            supplyOrderItemPage.Filter(FilterSupplyItemType.ShowItemsNotSupplied, true);

            Assert.IsFalse(supplyOrderItemPage.VerifySupplied(), String.Format(MessageErreur.FILTRE_ERRONE, "'Show items not supplied'"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_Items_Filter_SortBy()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();

            // Create
            var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
            createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
            var supplyOrderItemPage = createSupplyOrderPage.Submit();

            WebDriver.Navigate().Refresh();
            supplyOrderItemPage.ResetFilter();

            supplyOrderItemPage.Filter(FilterSupplyItemType.SortBy, "Supplier");
            supplyOrderItemPage.PageSize("100");
            Assert.IsTrue(supplyOrderItemPage.IsSortedBySupplier(), MessageErreur.FILTRE_ERRONE, "'Sort by supplier'");

            WebDriver.Navigate().Refresh();
            supplyOrderItemPage.Filter(FilterSupplyItemType.SortBy, "Group");
            supplyOrderItemPage.PageSize("100");
            Assert.IsTrue(supplyOrderItemPage.IsSortedByGroup(), MessageErreur.FILTRE_ERRONE, "'Sort by group'");

            supplyOrderItemPage.Filter(FilterSupplyItemType.SortBy, "Name");
            supplyOrderItemPage.PageSize("100");
            Assert.IsTrue(supplyOrderItemPage.IsSortedByName(), MessageErreur.FILTRE_ERRONE, "'Sort by item'");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_Items_Filter_Suppliers()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();

            // Create
            var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
            createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
            var supplyOrderItemPage = createSupplyOrderPage.Submit();

            string itemName = supplyOrderItemPage.GetFirstItemName();
            var itemSupplier = supplyOrderItemPage.GetFirstItemSupplier();

            supplyOrderItemPage.SelectFirstItem();
            supplyOrderItemPage.AddQuantity(itemName, "10");

            supplyOrderItemPage.ResetFilter();
            supplyOrderItemPage.Filter(FilterSupplyItemType.ShowItemsNotSupplied, true);
            supplyOrderItemPage.Filter(FilterSupplyItemType.Suppliers, itemSupplier);

            Assert.IsTrue(supplyOrderItemPage.VerifySupplier(itemSupplier), String.Format(MessageErreur.FILTRE_ERRONE, "'Suppliers'"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SO_Items_Filter_Groups()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();
            string itemGroup = "AIR CANADA";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.ResetFilter();

            // Create
            var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
            createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
            var supplyOrderItemPage = createSupplyOrderPage.Submit();

            // Récupération du group du premier item
            supplyOrderItemPage.Filter(FilterSupplyItemType.Groups, itemGroup);
            var itemName = supplyOrderItemPage.GetFirstItemName();
            supplyOrderItemPage.SelectFirstItem();
            var itemDetailPage = supplyOrderItemPage.EditItem(itemName);
            itemGroup = itemDetailPage.GetGroupName();
            itemDetailPage.Close();

            supplyOrderItemPage.ResetFilter();
            supplyOrderItemPage.Filter(FilterSupplyItemType.SortBy, "Group");
            supplyOrderItemPage.Filter(FilterSupplyItemType.Groups, itemGroup);

            Assert.IsTrue(supplyOrderItemPage.VerifyGroup(itemGroup), String.Format(MessageErreur.FILTRE_ERRONE, "'Group'"));
            Assert.IsTrue(supplyOrderItemPage.IsSortedByName(), "Les items du groupe ne sont pas triés par ordre alphabétique.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SO_Items_UpdateItemQuantity()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();
            string number = null;
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            try
            { // Create
                var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
                createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
                number = createSupplyOrderPage.GetSupplyOrderNumber();
                var supplyOrderItemPage = createSupplyOrderPage.Submit();

                // Récupération de la quantité du premier item
                var initQuantity = supplyOrderItemPage.GetFirstItemQty();
                var initStorage = supplyOrderItemPage.GetFirstStorageValue();
                Assert.AreEqual("", initStorage, "La valeur du storage de l'item initial est déjà affichée.");

                string itemName = supplyOrderItemPage.GetFirstItemName();
                supplyOrderItemPage.SelectFirstItem();
                supplyOrderItemPage.AddQuantity(itemName, "10");
                supplyOrderItemPage.Refresh();

                var newQuantity = supplyOrderItemPage.GetFirstItemQty();
                var newStorage = supplyOrderItemPage.GetFirstStorageValue();

                Assert.AreNotEqual(initQuantity, newQuantity, "L'update de la quantité a échoué.");
                Assert.AreNotEqual("", newStorage, "La valeur du storage n'est pas affichée malgré l'update de la quantité.");

            }
            finally
            {

                var supplyOrderPages = homePage.GoToPurchasing_SupplyOrderPage();
                supplyOrderPage.ResetFilter();
                if (!string.IsNullOrEmpty(number))
                {
                    supplyOrderPage.Filter(FilterType.ByNumber, number);
                    supplyOrderPage.DeleteFirstSupplyOrder();
                }
            }

        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SO_Items_LinkWithItems()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();

            // Create
            var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
            createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
            var supplyOrderItemPage = createSupplyOrderPage.Submit();

            var itemName = supplyOrderItemPage.GetFirstItemName();
            supplyOrderItemPage.SelectFirstItem();

            // Récupération du group du premier item
            var itemDetailPage = supplyOrderItemPage.EditItem(itemName);
            var itemGroup = itemDetailPage.GetItemName();
            itemDetailPage.Close();

            Assert.AreNotEqual("", itemGroup, "La page de l'item séléctionné n'est pas affichée.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SO_Items_SelectOtherPackaging()
        {
            //Prepare
            Random round = new Random();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string itemName = "itemSupplyOrderPackaging-" + round.Next() + DateUtils.Now.ToShortDateString();
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string packagingName2 = TestContext.Properties["Item_PackagingNameBis"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageUnit2 = "Bolsa";
            string supplierRef = itemName + "_SupplierRef";
            string storageQty = 10.ToString();
            string qty = 10.ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Création d'un item pour la supply order
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);

            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier, null, supplierRef);
                itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName2, storageQty, storageUnit2, qty, supplier, null, supplierRef);
            }

            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
            createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
            var supplyOrderItemPage = createSupplyOrderPage.Submit();

            supplyOrderItemPage.Filter(FilterSupplyItemType.ByName, itemName);
            var initPackaging = supplyOrderItemPage.GetFirstPacking();
            Assert.IsTrue(initPackaging.Contains(packagingName), "Le packaging initial n'est pas correct.");

            supplyOrderItemPage.SelectFirstItem();
            supplyOrderItemPage.AddQuantity(itemName, "10");

            bool isOtherPackagingFound = supplyOrderItemPage.SelectOtherPackaging(itemName, packagingName2);
            Assert.IsTrue(isOtherPackagingFound, "Aucun autre packaging n'a été trouvé pour l'item.");

            supplyOrderItemPage.ResetFilter();
            var newPackaging = supplyOrderItemPage.GetFirstPacking();
            Assert.IsTrue(newPackaging.Contains(packagingName2), "Le packaging initial n'a pas été renseigné.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_Items_AddComment()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();
            string comment = TestContext.Properties["CommentSupplier"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();

            // Create
            var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
            createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
            var supplyOrderItemPage = createSupplyOrderPage.Submit();

            string itemName = supplyOrderItemPage.GetFirstItemName();
            supplyOrderItemPage.SelectFirstItem();
            supplyOrderItemPage.AddQuantity(itemName, "10");

            supplyOrderItemPage.AddComment(itemName, comment);

            //Assert
            supplyOrderItemPage.Refresh();
            supplyOrderItemPage.SelectFirstItem();
            var newComment = supplyOrderItemPage.GetComment(itemName);
            Assert.AreEqual(newComment, comment, "Le commentaire n'a pas été ajouté.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_Items_DeleteItemFromSO()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();

            // Create
            var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
            createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
            var supplyOrderItemPage = createSupplyOrderPage.Submit();

            string itemName = supplyOrderItemPage.GetFirstItemName();
            supplyOrderItemPage.SelectFirstItem();
            supplyOrderItemPage.AddQuantity(itemName, "10");

            supplyOrderItemPage.Refresh();
            var initQuantity = supplyOrderItemPage.GetFirstItemQty();

            supplyOrderItemPage.SelectFirstItem();
            supplyOrderItemPage.DeleteFirstItem(itemName);

            var newQuantity = supplyOrderItemPage.GetFirstItemQty();

            Assert.AreNotEqual(initQuantity, newQuantity, "La quantité n'a pas été modifiée.");
            Assert.AreEqual("0", newQuantity, "La quantité n'a pas été effacée.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_Items_CompareItemsOK()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();

            int nbItemToCompare = 3;

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();

            // Create
            var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
            createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
            var supplyOrderItemPage = createSupplyOrderPage.Submit();

            // On renseigne les 3 premières lignes avant de comparer les items
            supplyOrderItemPage.Fill_Compare(nbItemToCompare);
            Assert.IsFalse(supplyOrderItemPage.IsErrorPopUpVisible(), "La sélection de " + nbItemToCompare.ToString() + " items pour les comparer n'est pas permise.");

            supplyOrderItemPage.Compare();
            Assert.AreEqual(supplyOrderItemPage.GetNbCompareLines(), nbItemToCompare, "Le nombre d'items à comparer est incorrect.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_Items_EraseCompareItems()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();
            int nbItemToCompare = 3;

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();

            // Create
            var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
            createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
            var supplyOrderItemPage = createSupplyOrderPage.Submit();

            var initNbLignes = supplyOrderItemPage.GetNbCompareLines();

            // On renseigne les 3 premières lignes avant de comparer les items
            supplyOrderItemPage.Fill_Compare(nbItemToCompare);
            Assert.IsFalse(supplyOrderItemPage.IsErrorPopUpVisible(), "La sélection de " + nbItemToCompare.ToString() + " items pour les comparer n'est pas permise.");

            supplyOrderItemPage.Compare();
            Assert.AreEqual(supplyOrderItemPage.GetNbCompareLines(), nbItemToCompare, "Le nombre d'items à comparer est incorrect.");

            supplyOrderItemPage.EraseCompare();
            Assert.AreEqual(supplyOrderItemPage.GetNbCompareLines(), initNbLignes, "Le nombre d'items à comparer n'a pas été modifié suite à l'annulation de la commande.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_Items_CompareItemsKO()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();
            int nbItemToCompare = 4;

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();

            // Create
            var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
            createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
            var supplyOrderItemPage = createSupplyOrderPage.Submit();

            // On renseigne les 4 premières lignes avant de comparer les items
            supplyOrderItemPage.Fill_Compare(nbItemToCompare);
            Assert.IsTrue(supplyOrderItemPage.IsErrorPopUpVisible(), "La sélection de " + nbItemToCompare.ToString() + " items pour les comparer est permise.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_Details_ExportResults_NewVersion()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();

            bool newVersionPrint = true;

            //arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();

            supplyOrderPage.ClearDownloads();

            // Create
            var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
            createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
            var supplyOrderItemPage = createSupplyOrderPage.Submit();
            supplyOrderItemPage.Filter(FilterSupplyItemType.Groups, new List<string> { "AIR CANADA", "AMERICAN AIRLINES" });
            var itemName = supplyOrderItemPage.GetFirstItemName();
            supplyOrderItemPage.SelectFirstItem();
            supplyOrderItemPage.AddQuantity(itemName, "10");

            supplyOrderItemPage.ExportItems(newVersionPrint);
            supplyOrderItemPage.WaitLoading();
            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = supplyOrderItemPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            int resultNumber = OpenXmlExcel.GetExportResultNumber(SUPPLY_ORDERS_EXCEL_SHEET_NAME, filePath);
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_Details_EraseQuantity()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();

            // Create
            var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
            createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
            var supplyOrderItemPage = createSupplyOrderPage.Submit();

            var itemName = supplyOrderItemPage.GetFirstItemName();
            supplyOrderItemPage.Filter(FilterSupplyItemType.ByName, itemName);
            supplyOrderItemPage.SelectFirstItem();
            supplyOrderItemPage.AddQuantity(itemName, "10");
            supplyOrderItemPage.Refresh();

            supplyOrderItemPage.Filter(FilterSupplyItemType.ByName, itemName);
            Assert.AreNotEqual(0, supplyOrderItemPage.GetFirstItemQty(), "La quantité ajoutée à l'item n'a pas été prise en compte.");

            supplyOrderItemPage.EraseQuantity();

            supplyOrderItemPage.Filter(FilterSupplyItemType.ByName, itemName);
            supplyOrderItemPage.Filter(FilterSupplyItemType.ShowItemsNotSupplied, true);
            Assert.AreEqual("0", supplyOrderItemPage.GetFirstItemQty(), "La quantité ajoutée à l'item n'a pas été effacée.");
        }


        [DeploymentItem("Resources\\logo.jpg")]
        [DeploymentItem("Resources\\logo_petit.jpg")]
        [DeploymentItem("chromedriver.exe")]
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SO_Details_Print_NewVersion()
        {
            //Prepare
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();
            string DocFileNamePdfBegin = "Print Report_-_";
            string DocFileNameZipBegin = "All_files_";
            bool newVersionPrint = true;

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();

            supplyOrderPage.ClearDownloads();
            //clear donwload folder
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();
            foreach (FileInfo file in taskFiles)
            {
                file.Delete();
            }
            var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
            createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
            var supplyOrderItemPage = createSupplyOrderPage.Submit();

            string itemName = supplyOrderItemPage.GetFirstItemName();
            supplyOrderItemPage.SelectFirstItem();
            supplyOrderItemPage.AddQuantity(itemName, "10");

            supplyOrderItemPage.Validate();

            var reportPage = supplyOrderItemPage.Print(newVersionPrint);
            var isGenerated = reportPage.IsReportGenerated();
            reportPage.Close();

            Assert.IsTrue(isGenerated, "Le document PDF n'a pas pu être généré par l'application.");

            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);

            // get downloaded pdf file
            string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            Assert.IsNotNull(trouve, MessageErreur.FICHIER_NON_TROUVE);


            // verifier données sur IHM et sur pdf
            var iscorrect = supplyOrderItemPage.VerifyPdf(trouve, itemName);
            Assert.IsTrue(iscorrect, "les données ne sont pas présentes dans le fichier pdf");

            // 2.Verifier le logo
            PdfDocument document = PdfDocument.Open(trouve);
            var page1 = document.GetPage(1);
            List<IPdfImage> images = page1.GetImages().ToList<IPdfImage>();
            Assert.AreEqual(1, images.Count, "Pas d'image dans le PDF");
            IPdfImage image = images[0];

            FileInfo fiPdf = new FileInfo(downloadsPath + "\\logo_test.jpg");
            if (fiPdf.Exists)
            {
                fiPdf.Delete();
            }
            File.WriteAllBytes(downloadsPath + "\\logo_test.jpg", image.RawBytes.ToArray<byte>());
            fiPdf.Refresh();
            Assert.IsTrue(fiPdf.Exists, "fichier non trouvé");

            FileInfo fiTest = new FileInfo(TestContext.TestDeploymentDir + "\\logo_petit.jpg");
            Assert.IsTrue(fiTest.Exists, "fichier non trouvé (cas 2)");

            // comparaison fichiers

            var md5 = MD5.Create();
            var fsTest = File.Open(fiTest.FullName, FileMode.Open);
            var fsPdf = File.Open(fiPdf.FullName, FileMode.Open);
            var test1 = System.Convert.ToBase64String(md5.ComputeHash(fsTest));
            var test2 = System.Convert.ToBase64String(md5.ComputeHash(fsPdf));
            Assert.AreEqual(test1, test2, "fichiers différents");


        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_Details_Validate()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();

            //Create
            var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
            createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
            var id = createSupplyOrderPage.GetSupplyOrderNumber();
            var supplyOrderItemPage = createSupplyOrderPage.Submit();

            string itemName = supplyOrderItemPage.GetFirstItemName();
            supplyOrderItemPage.SelectFirstItem();
            supplyOrderItemPage.AddQuantity(itemName, "10");

            supplyOrderItemPage.Validate();
            supplyOrderPage = supplyOrderItemPage.BackToList();

            supplyOrderPage.Filter(FilterType.ByNumber, id);

            Assert.IsTrue(supplyOrderPage.CheckValidation(true), "Le supply order n'a pas été validé.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SO_Details_GeneratePurchaseOrder()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();

            //Create
            var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
            createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
            var id = createSupplyOrderPage.GetSupplyOrderNumber();
            var supplyOrderItemPage = createSupplyOrderPage.Submit();

            supplyOrderItemPage.ResetFilter();
            string itemName = supplyOrderItemPage.GetFirstItemName();
            supplyOrderItemPage.SelectFirstItem();
            supplyOrderItemPage.AddQuantity(itemName, "10");

            supplyOrderItemPage.Validate();

            var purchaseOrderPage = supplyOrderItemPage.GeneratePurchaseOrder();
            var purchaseOrderNumber = purchaseOrderPage.GetFirstPurchaseOrderNumber();

            // Retour sur la page SupplyOrder
            supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.ResetFilter();
            supplyOrderPage.Filter(FilterType.ByNumber, id);
            supplyOrderItemPage = supplyOrderPage.SelectFirstItem();

            var supplyOrderGeneralInformationPage = supplyOrderItemPage.ClickOnGeneralInformation();
            Assert.AreEqual(supplyOrderGeneralInformationPage.GetPurchaseOrderValue(), purchaseOrderNumber, "Le purchase order n'a pas été généré pour le supply order.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_Recipes_Filter_SearchByName()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();
            string recipeName = TestContext.Properties["RecipeName"].ToString();
            string recipeName2 = TestContext.Properties["RecipeName1"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();

            //Create
            var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
            createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
            var supplyOrderItemPage = createSupplyOrderPage.Submit();

            string itemName = supplyOrderItemPage.GetFirstItemName();
            supplyOrderItemPage.SelectFirstItem();
            supplyOrderItemPage.AddQuantity(itemName, "10");

            var supplyOrderRecipePage = supplyOrderItemPage.ClickOnRecipes();
            supplyOrderRecipePage.AddRecipe(recipeName);
            supplyOrderRecipePage.AddRecipe(recipeName2);

            supplyOrderRecipePage.Filter(SupplyOrderRecipes.FilterSupplyRecipeType.SearchByName, recipeName);

            Assert.IsTrue(supplyOrderRecipePage.VerifyName(recipeName), String.Format(MessageErreur.FILTRE_ERRONE, "Search by name"));
        }

        [Ignore]//
        [TestMethod]
        public void PU_SO_Recipes_Filter_SortBy()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();
            string recipeName = TestContext.Properties["RecipeName"].ToString();
            string recipeName2 = TestContext.Properties["RecipeName1"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();

            //Create
            var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
            createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
            var supplyOrderItemPage = createSupplyOrderPage.Submit();

            string itemName = supplyOrderItemPage.GetFirstItemName();
            supplyOrderItemPage.SelectFirstItem();
            supplyOrderItemPage.AddQuantity(itemName, "10");

            var supplyOrderRecipePage = supplyOrderItemPage.ClickOnRecipes();
            supplyOrderRecipePage.AddRecipe(recipeName);
            supplyOrderRecipePage.AddRecipe(recipeName2);

            supplyOrderRecipePage.Filter(SupplyOrderRecipes.FilterSupplyRecipeType.SortBy, "Name");
            bool isSortedbyName = supplyOrderRecipePage.IsSortedByName();

            supplyOrderRecipePage.Filter(SupplyOrderRecipes.FilterSupplyRecipeType.SortBy, "Portions");
            bool isSortedbyPortion = supplyOrderRecipePage.IsSortedByPortion();

            Assert.IsTrue(isSortedbyName, String.Format(MessageErreur.FILTRE_ERRONE, "Sort by name"));
            Assert.IsTrue(isSortedbyPortion, String.Format(MessageErreur.FILTRE_ERRONE, "Search by portions"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_Recipes_AddRecipe()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();
            string recipeName = TestContext.Properties["RecipeName"].ToString();
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();

            //Create
            var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
            createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
            var supplyOrderItemPage = createSupplyOrderPage.Submit();

            string itemName = supplyOrderItemPage.GetFirstItemName();
            supplyOrderItemPage.SelectFirstItem();
            supplyOrderItemPage.AddQuantity(itemName, "10");

            var supplyOrderRecipePage = supplyOrderItemPage.ClickOnRecipes();
            supplyOrderRecipePage.AddRecipe(recipeName);

            Assert.IsTrue(supplyOrderRecipePage.IsFirstRecipeAdded(), "La recette n'a pas été ajoutée dans le supply order.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_Recipes_UpdateRecipeData()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();
            string recipeName = TestContext.Properties["RecipeName"].ToString();
            string newQty = "10";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();

            //Create
            var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
            createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
            var supplyOrderItemPage = createSupplyOrderPage.Submit();

            string itemName = supplyOrderItemPage.GetFirstItemName();
            supplyOrderItemPage.SelectFirstItem();
            supplyOrderItemPage.AddQuantity(itemName, "10");

            var supplyOrderRecipePage = supplyOrderItemPage.ClickOnRecipes();
            supplyOrderRecipePage.AddRecipe(recipeName);

            Assert.IsTrue(supplyOrderRecipePage.IsFirstRecipeAdded(), "La recette n'a pas été ajoutée dans le supply order.");

            var initNeededQty = supplyOrderRecipePage.GetFirstNeededQty();
            var initPrice = supplyOrderRecipePage.GetFirstPrice();

            supplyOrderRecipePage.SetFirstRecipeNeededQty(newQty);
            supplyOrderRecipePage.Refresh();

            var newNeededQty = supplyOrderRecipePage.GetFirstNeededQty();
            var newPrice = supplyOrderRecipePage.GetFirstPrice();

            Assert.AreNotEqual(initNeededQty, newNeededQty, "La quantité n'a pas été modifiée dans la recette.");
            Assert.AreNotEqual(initPrice, newPrice, "Le prix de la recette n'a pas été modifié.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_Recipes_DeleteRecipe()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();
            string recipeName = TestContext.Properties["RecipeName"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();

            //Create
            var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
            createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
            var supplyOrderItemPage = createSupplyOrderPage.Submit();

            string itemName = supplyOrderItemPage.GetFirstItemName();
            supplyOrderItemPage.SelectFirstItem();
            supplyOrderItemPage.AddQuantity(itemName, "10");

            var supplyOrderRecipePage = supplyOrderItemPage.ClickOnRecipes();
            supplyOrderRecipePage.AddRecipe(recipeName);

            Assert.IsTrue(supplyOrderRecipePage.IsFirstRecipeAdded(), "La recette n'a pas été ajoutée dans le supply order.");

            supplyOrderRecipePage.DeleteFirstRecipe();
            Assert.IsFalse(supplyOrderRecipePage.IsFirstRecipeAdded(), "La recette n'a pas été supprimée du supply order.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_Recipes_EditRecipe()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();
            string recipeName = TestContext.Properties["RecipeName"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();

            //Create
            var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
            createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
            var supplyOrderItemPage = createSupplyOrderPage.Submit();

            string itemName = supplyOrderItemPage.GetFirstItemName();
            supplyOrderItemPage.SelectFirstItem();
            supplyOrderItemPage.AddQuantity(itemName, "10");

            var supplyOrderRecipePage = supplyOrderItemPage.ClickOnRecipes();
            supplyOrderRecipePage.AddRecipe(recipeName);

            Assert.IsTrue(supplyOrderRecipePage.IsFirstRecipeAdded(), "La recette n'a pas été ajoutée dans le supply order.");

            var recipeVariantPage = supplyOrderRecipePage.EditFirstRecipe();
            string name = recipeVariantPage.GetRecipeName();
            recipeVariantPage.Close();

            Assert.AreNotEqual("", name, "La page d'édition de la recette est visible.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_GeneratePOfromSO()
        {
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Etre sur une SO validée
            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.ResetFilter();
            supplyOrderPage.Filter(FilterType.ShowNotValidated, false);
            string soNumber = supplyOrderPage.GetFirstSONumber();
            SupplyOrderItem soItem = supplyOrderPage.SelectFirstItem();

            //1) Etre sur une SO avec item/ quantité validée
            string itemName = soItem.GetFirstItemNameValidated();
            string itemQty = soItem.GetFirstItemQtyValidated();
            Assert.IsTrue(int.Parse(itemQty) > 0, "pas de quantité validée");

            //2) ... puis Generate Purchase Order
            //3) Choisir le supplier et la delivery date
            PurchaseOrdersPage poPage = soItem.GeneratePurchaseOrder();
            poPage.ResetFilters();
            poPage.Filter(PurchaseOrdersPage.FilterType.BySupplierOrderNumber, soNumber);

            //4) Generate Vérifier dans Gal info le numéro de la SO de base
            PurchaseOrderItem poItem = poPage.SelectFirstItem();
            PurchaseOrderGeneralInformation generalInfo = poItem.ClickOnGeneralInformation();
            Assert.AreEqual(soNumber, generalInfo.getSupplyOrderNumber());
            PurchaseOrderItem poItems = generalInfo.ClickOnItemsTab();

            //Vérifier les items et quantité (qty)
            Assert.AreEqual(itemQty, poItems.GetQuantity(), "mauvaise quantité");
            //Vérifier PO non validée
            try
            {
                // on a un <form> dans la ligne ? si oui, non validé
                poItems.SelectFirstItemPo();
            }
            catch
            {
                Assert.Fail("PO validé");
            }
            //Vérifier les items et quantité (nom)
            // "Item Test OF from SO (Item Test OF from SO_SupplierRef)"
            Assert.IsTrue(itemName.Trim().Contains(poItems.GetFirstItemName().Trim()), "mauvais nom d'item");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SO_GenerateOFFromSO()
        {
            // Prepare
            string itemName = "Item Test OF from SO";
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();

            // la place où on a posé notre stock
            string placeFrom = "Economato";
            string supplierRef = itemName + "_SupplierRef";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            ClearCache();

            // Vérifier que l'ingrédient existe
            var itemPage = homePage.GoToPurchasing_ItemPage();

            itemPage.Filter(ItemPage.FilterType.Search, itemName);

            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier, supplierRef);
            }

            //1. Créer un inventaire sur le site, ajouter 40 phys.qtty sur un item(Item Test OF from SO sur MAD)
            InventoriesPage inventory = itemPage.GoToWarehouse_InventoriesPage();
            InventoryCreateModalPage inventoryCreate = inventory.InventoryCreatePage();
            inventoryCreate.FillField_CreateNewInventory(DateUtils.Now, site, placeFrom);
            InventoryItem inventoryItems = inventoryCreate.Submit();
            inventoryItems.Filter(InventoryItem.FilterItemType.SearchByName, itemName);
            //le packaging qty (fois packaging unit) s'additionne au physQty
            inventoryItems.SelectFirstItem();
            inventoryItems.AddPhysicalQuantity(itemName, "40");
            inventoryItems.AddPhysicalPackagingQuantity(itemName, "0");
            InventoryValidationModalPage validateModal = inventoryItems.Validate();
            validateModal.ValidatePartialInventory();

            //1) Etre sur une SO avec item/ quantité validée
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.ResetFilter();
            var newSupplyOrder = supplyOrderPage.CreateNewSupplyOrder();
            newSupplyOrder.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, placeFrom, true);
            string SONb = newSupplyOrder.GetSupplyOrderNumber();
            var supplyOrderItemPage = newSupplyOrder.Submit();

            supplyOrderItemPage.Filter(FilterSupplyItemType.ByName, itemName);
            supplyOrderItemPage.SelectFirstItem();
            supplyOrderItemPage.AddQuantity(itemName, "1");
            supplyOrderItemPage.Validate();

            int itemQty = int.Parse(supplyOrderItemPage.GetFirstItemQtyValidated());
            //on récupère la valuer numérique du Packaging
            int itemPack = int.Parse(supplyOrderItemPage.GetFirstPackingValueValidated());
            Assert.IsTrue(itemQty > 0, "pas de quantité validée");
            Assert.IsTrue(itemPack > 0, "pas de packing validée");

            //2) ... puis Generate OutputForm
            //3) Choisir le supplier et la delivery date
            // pas bon cette spec, ici on sélectionne la placeFrom et la placeTo
            OutputFormItem outputForm = supplyOrderItemPage.GenerateOutputForm();

            //4) Create Vérifier dans Gal info le numéro de la SO de base
            OutputFormGeneralInformation generalInfo = outputForm.ClickOnGeneralInformationTab();
            Assert.AreEqual(SONb, generalInfo.GetSupplyOrderNumber());
            //Vérifier les items et quantité
            outputForm = generalInfo.ClickOnItemsTab();

            outputForm.Filter(OutputFormItem.FilterItemType.ShowItemsWithPhysQty, false);
            outputForm.Filter(OutputFormItem.FilterItemType.SearchByName, itemName);

            Assert.AreEqual((itemQty * itemPack).ToString(), outputForm.GetPhysicalQuantity(), "Mauvaise qty");
            //Vérifier OF non validée
            try
            {
                // si pas de <form> alors validé
                outputForm.SelectFirstItem();
            }
            catch
            {
                Assert.Fail("OF non validée");
            }
            Assert.AreEqual(itemName, outputForm.GetFirstItemName(), "mauvais item");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SO_GenerateAndValidatePOfromSO()
        {
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Etre sur une SO validée
            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.ResetFilter();
            supplyOrderPage.Filter(FilterType.ShowNotValidated, false);
            string soNumber = supplyOrderPage.GetFirstSONumber();
            SupplyOrderItem soItem = supplyOrderPage.SelectFirstItem();

            //1) Etre sur une SO avec item/ quantité validée
            string itemName = soItem.GetFirstItemNameValidated();
            string itemQty = soItem.GetFirstItemQtyValidated();
            Assert.IsTrue(int.Parse(itemQty) > 0, "pas de quantité validée");

            //2) ... puis Generate Purchase Order
            //3) Choisir le supplier et la delivery date
            //4) Generate & validate
            PurchaseOrdersPage poPage = soItem.GeneratePurchaseOrder();
            poPage.ResetFilters();
            poPage.Filter(PurchaseOrdersPage.FilterType.BySupplierOrderNumber, soNumber);
            PurchaseOrderItem poItem = poPage.SelectFirstItem();
            poItem.Validate();

            //Vérifier dans Gal info le numéro de la SO de base
            PurchaseOrderGeneralInformation generalInfo = poItem.ClickOnGeneralInformation();
            Assert.AreEqual(soNumber, generalInfo.getSupplyOrderNumberValidated());
            PurchaseOrderItem poItems = generalInfo.ClickOnItemsTab();

            //Vérifier PO validées
            //Vérifier les items et quantité (qty)
            Assert.AreEqual(itemQty, poItems.GetQuantity(), "mauvaise quantité");
            Assert.IsTrue(itemName.Contains(poItems.GetFirstItemNameValidated()), "mauvais nom d'item");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_CreateSO_PrefillCopyItemsFromAnotherSO()
        {
            //Prepare          
            string site = TestContext.Properties["Site"].ToString(); // MAD
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Etre sur une SO validée
            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();

            //Avoir créer un SO avec un item / quantité et valider
            CreateSupplyOrderModalPage newSupplyOrder = supplyOrderPage.CreateNewSupplyOrder();
            newSupplyOrder.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation, true);
            string supplyOrderNumber = newSupplyOrder.GetSupplyOrderNumber();
            var supplyOrderItemPage = newSupplyOrder.Submit();
            string itemName = supplyOrderItemPage.GetFirstItemName();
            supplyOrderItemPage.SelectFirstItem();
            supplyOrderItemPage.AddQuantity(itemName, "23");
            supplyOrderItemPage.BackToList();

            //1) Create new SO
            CreateSupplyOrderModalPage newPrefillSupplyOrder = supplyOrderPage.CreateNewSupplyOrder();
            newPrefillSupplyOrder.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation, true);
            newPrefillSupplyOrder.GetSupplyOrderNumber();

            //2) Cocher prefill 'Copy items from another supply order(s)'

            //3) Sélectionner les SO créer et valider précédemment
            newPrefillSupplyOrder.FillFromAnotherSupplyOrder(site, supplyOrderNumber);
            SupplyOrderItem items = newPrefillSupplyOrder.Submit();

            //4) Create Vérifier dans General Info le numéro de la SO copiée
            var generalInfo = items.ClickOnGeneralInformation();
            var fromSO = generalInfo.WaitForElementExists(By.XPath("//*/span[contains(text(),'From supply order(s) :')]/a"));
            Assert.AreEqual(supplyOrderNumber, fromSO.Text);
            items = generalInfo.ClickOnItemsTab();
            //Vérifier le / les items et quantité
            Assert.AreEqual(itemName, items.GetFirstItemName());
            items.SelectFirstItem();
            Assert.AreEqual("23", items.GetFirstItemQty());
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SO_CreateSO_PrefillCopyItemsFromAnotherSO_otherSite()
        {
            //Prepare          
            string site = TestContext.Properties["Site"].ToString(); // MAD
            string siteDestination = TestContext.Properties["SiteACE"].ToString(); // ACE
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Etre sur une SO validée
            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();

            //Avoir créer un SO avec un item / quantité et valider
            CreateSupplyOrderModalPage newSupplyOrder = supplyOrderPage.CreateNewSupplyOrder();
            newSupplyOrder.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now.AddDays(-1), deliveryLocation, true);
            string supplyOrderNumber = newSupplyOrder.GetSupplyOrderNumber();
            var supplyOrderItemPage = newSupplyOrder.Submit();
            string itemName = supplyOrderItemPage.GetFirstItemName();
            supplyOrderItemPage.SelectFirstItem();
            supplyOrderItemPage.AddQuantity(itemName, "23");
            supplyOrderItemPage.BackToList();

            //1) Create new SO
            CreateSupplyOrderModalPage newPrefillSupplyOrder = supplyOrderPage.CreateNewSupplyOrder();
            newPrefillSupplyOrder.FillPrincipalField_CreationNewSupplyOrder(siteDestination, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now.AddDays(-1), deliveryLocation, true);
            newPrefillSupplyOrder.GetSupplyOrderNumber();

            //2) Cocher prefill 'Copy items from another supply order(s)'

            //3) Sélectionner les SO créer et valider précédemment
            newPrefillSupplyOrder.FillFromAnotherSupplyOrder(site, supplyOrderNumber);
            SupplyOrderItem items = newPrefillSupplyOrder.Submit();

            //4) Create Vérifier dans General Info le numéro de la SO copiée
            var generalInfo = items.ClickOnGeneralInformation();
            var fromSO = generalInfo.WaitForElementExists(By.XPath("//*/span[contains(text(),'From supply order(s) :')]/a"));
            Assert.AreEqual(supplyOrderNumber, fromSO.Text);
            items = generalInfo.ClickOnItemsTab();
            //Vérifier le / les items et quantité
            Assert.AreEqual(itemName, items.GetFirstItemName());
            items.SelectFirstItem();
            Assert.AreEqual("23", items.GetFirstItemQty());
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SO_Filter_DetailItem_BySubGroup()
        {
            //Prepare data
            string subGrpName = "subgrpname";
            string subGrpCode = "subgrpcode";
            string group = "A REFERENCIA";

            //connect as admin
            LogInAsAdmin();

            //go to homePage
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //changing setting sub group to active
            homePage.SetSubGroupFunctionValue(true);
            //create new subgroup
            ParametersProduction productionPage = homePage.GoToParameters_ProductionPage();
            productionPage.GoToTab_SubGroup();
            if (!productionPage.IsGroupPresent(subGrpName))
            {
                productionPage.AddNewSubGroup(subGrpName, subGrpCode);
            }

            //Click first supply order
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.ResetFilter();
            supplyOrderPage.Filter(FilterType.ShowNotActive, true);
            var supplyOrderNumber = supplyOrderPage.GetFirstSONumber();
            var supplyOrderItems = supplyOrderPage.ClickFirstSupplyOrder();
            Thread.Sleep(1000);
            //get first purchase order item name
            supplyOrderItems.SelectFirstItem();
            var itemName = supplyOrderItems.GetFirstItemName();

            //set item subgroup
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            var itemGeneralInformationPage = itemPage.ClickOnFirstItem();
            itemGeneralInformationPage.SetGroupName(group);
            itemGeneralInformationPage.SetSubgroupName(subGrpName);
            Thread.Sleep(1000);
            //filter by subgroupe
            supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.Filter(SupplyOrderPage.FilterType.ByNumber, supplyOrderNumber);
            supplyOrderItems = supplyOrderPage.ClickFirstSupplyOrder();
            supplyOrderItems.Filter(SupplyOrderItem.FilterSupplyItemType.BySubGroup, subGrpName);
            supplyOrderItems.Filter(SupplyOrderItem.FilterSupplyItemType.ByName, itemName);
            //assert
            Assert.IsTrue(supplyOrderItems.VerifySubGroupFiltre(itemName), "erreur lors du filtrage par subgroupe");

            //changing setting sub group to inactive
            homePage.SetSubGroupFunctionValue(false);
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_ResetFilter()
        {

            //Arrange
            var homePage = LogInAsAdmin();

            //act
            //go to supply orders
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.ResetFilter();
            var number = supplyOrderPage.GetFirstSONumber();
            supplyOrderPage.Filter(FilterType.ByNumber, number);
            supplyOrderPage.ResetFilter();
            //get filter value and check if empty 

            //Assert
            var searchValue = supplyOrderPage.GetSearchValue();
            bool IsNullOrEmptySearch = String.IsNullOrEmpty(searchValue);
            Assert.IsTrue(IsNullOrEmptySearch, "Les filtres ne sont pas réinitialisés");
        }
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_DetailChangeLine()
        {
            //Prepare          
            string site = TestContext.Properties["Site"].ToString(); // MAD
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();

            //connect as admin
            LogInAsAdmin();
            //go to homePage
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            //go to supply orders
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();

            //Avoir créer un SO avec un item / quantité et valider
            CreateSupplyOrderModalPage newSupplyOrder = supplyOrderPage.CreateNewSupplyOrder();
            newSupplyOrder.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation, true);
            var supplyOrderItemPage = newSupplyOrder.Submit();
            string itemName = supplyOrderItemPage.GetFirstItemName();
            supplyOrderItemPage.SelectFirstItem();
            //Assert
            Assert.IsTrue(supplyOrderItemPage.VerifyDetailChangeLine(itemName), String.Format(MessageErreur.FILTRE_ERRONE, "ChangeLine ne fonctionne pas."));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SO_INDEX_CREATENEWSO()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();
            string articleQte = "10";
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.ResetFilter();
            // Create New Supply Order
            var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
            createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
            var supplyOrderItemPage = createSupplyOrderPage.Submit();

            var firstItemName = supplyOrderItemPage.GetFirstItemName();
            supplyOrderItemPage.SelectFirstItem();
            supplyOrderItemPage.AddQuantity(firstItemName, articleQte);
            supplyOrderItemPage.Validate();
            var itemNameValidated = supplyOrderItemPage.GetFirstItemNameValidated();
            var itemQtyValidated = supplyOrderItemPage.GetFirstItemQtyValidated();
            //Assert
            Assert.IsTrue((itemNameValidated == firstItemName) && (articleQte == itemQtyValidated), "les quantitées des articles du supply order créé n'sont pas générées.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_INDEX_AFFICHAGESO16()
        {
            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.ResetFilter();
            supplyOrderPage.PageSize("16");
            if (supplyOrderPage.ConvertStringToInt(supplyOrderPage.NombreSupplyOrdersS()) >= 16)
            {
                Assert.AreEqual(16, supplyOrderPage.GetTotalRowsForPagination(), "Le nombre de suppliers affichés dans l'index ne correspondre pas au nombre choisi dans la liste déroulante 16 ");
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_INDEX_AFFICHAGESO30()
        {
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.ResetFilter();
            supplyOrderPage.PageSize("30");
            if (supplyOrderPage.ConvertStringToInt(supplyOrderPage.NombreSupplyOrdersS()) >= 30)
            {
                Assert.AreEqual(30, supplyOrderPage.GetTotalRowsForPagination(), "Le nombre de suppliers affichés dans l'index ne correspondre pas au nombre choisi dans la liste déroulante 30 ");
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_INDEX_AFFICHAGESO50()
        {
            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.ResetFilter();
            supplyOrderPage.PageSize("50");
            if (supplyOrderPage.ConvertStringToInt(supplyOrderPage.NombreSupplyOrdersS()) >= 50)
            {
                Assert.AreEqual(50, supplyOrderPage.GetTotalRowsForPagination(), "Le nombre de suppliers affichés dans l'index ne correspondre pas au nombre choisi dans la liste déroulante 50 ");
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_INDEX_AFFICHAGESO100()
        {
            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.ResetFilter();
            supplyOrderPage.PageSize("100");
            if (supplyOrderPage.ConvertStringToInt(supplyOrderPage.NombreSupplyOrdersS()) >= 100)
            {
                Assert.AreEqual(100, supplyOrderPage.GetTotalRowsForPagination(), "Le nombre de suppliers affichés dans l'index ne correspondre pas au nombre choisi dans la liste déroulante 100 ");
            }
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SO_INDEX_CREATESOFROMMENUPLANIFICATION()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();
            string articleQte = "15";
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.ResetFilter();
            // Create New Supply Order
            var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
            createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
            createSupplyOrderPage.CheckedPrefillQtesFromMenuPlanification();
            var supplyOrderItemPage = createSupplyOrderPage.Submit();
            supplyOrderItemPage.Go_To_New_Navigate();
            var firstItemName = supplyOrderItemPage.GetFirstItemName();
            supplyOrderItemPage.SelectFirstItem();
            supplyOrderItemPage.AddQuantity(firstItemName, articleQte);
            supplyOrderItemPage.Validate();
            var itemNameValidated = supplyOrderItemPage.GetFirstItemNameValidated();
            var itemQtyValidated = supplyOrderItemPage.GetFirstItemQtyValidated();
            //Assert
            Assert.IsTrue((itemNameValidated == firstItemName) && (articleQte == itemQtyValidated), "La supply order n'est pas générée à partir de 'menu planification'.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_INDEX_ERROREXPORT()
        {
            //Prepare
            string site = TestContext.Properties["SiteMAD"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.ResetFilter();
            supplyOrderPage.Filter(FilterType.DateFrom, DateUtils.Now.AddDays(-10));
            supplyOrderPage.Filter(FilterType.DateTo, DateUtils.Now.AddDays(+20));

            if (supplyOrderPage.CheckTotalNumber() > 200)
            {
                //Assert
                bool msgErreurExist = supplyOrderPage.GetMessageExport().Contains("The maximum allowed in one time is 200 results.");
                Assert.IsTrue(msgErreurExist, "Le message d'erreur on ne peut pas exporter plus de 200 résultats n'apparait pas ");
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO()
        {

            string deliveryLocation = "ACE-Economato";

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();

            supplyOrderPage.ResetFilter();
            var totalNumber = supplyOrderPage.CheckTotalNumber();
            supplyOrderPage.ChooseDeliveryLocationOption(deliveryLocation);
            var totalNumberAfterFilter = supplyOrderPage.CheckTotalNumber();
            Assert.AreNotEqual(totalNumberAfterFilter, totalNumber);

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_DETAIL_REFRESH()
        {
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            var winrestUrl = TestContext.Properties["Winrest_URL"].ToString();
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act

            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            try
            {
                supplyOrderPage.ResetFilter();
                var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
                createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrderActiveFalse(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);

                var id = createSupplyOrderPage.GetSupplyOrderNumber();
                var supplyOrderItemPage = createSupplyOrderPage.Submit();
                string itemName = supplyOrderItemPage.GetFirstItemName();
                supplyOrderItemPage.SetFirstQuantityRefresh(itemName, "10");

                supplyOrderItemPage.OpenNewTab();
                supplyOrderItemPage.Refresh();
                var quantityAfterRefresh = supplyOrderItemPage.GetFirstItemQtyRefresh();

                Assert.AreEqual(quantityAfterRefresh, "10");
                supplyOrderItemPage.BackToList();
                supplyOrderPage.ResetFilter();
                supplyOrderPage.Filter(SupplyOrderPage.FilterType.ByNumber, id);
                supplyOrderPage.Filter(SupplyOrderPage.FilterType.ShowAll, true);

            }
            finally
            {
                supplyOrderPage.DeleteFirstSupplyOrder();
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_GENERATEPO()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            //Create
            var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
            createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
            var id = createSupplyOrderPage.GetSupplyOrderNumber();
            var supplyOrderItemPage = createSupplyOrderPage.Submit();

            supplyOrderItemPage.ResetFilter();
            string itemName = supplyOrderItemPage.GetFirstItemName();
            supplyOrderItemPage.SelectFirstItem();
            supplyOrderItemPage.AddQuantity(itemName, "10");

            supplyOrderItemPage.Validate();
            var purchaseOrderPage = supplyOrderItemPage.GeneratePurchaseOrder();
            var purchaseOrderNumbers = purchaseOrderPage.CheckTotalNumber();
            Assert.AreEqual(purchaseOrderNumbers, 1);
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SO_INDEX_CREATESOFROMNEEDS()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.ResetFilter();
            // Create New Supply Order
            var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
            createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now.AddDays(-1), deliveryLocation);
            createSupplyOrderPage.CheckFromNeeds();
            var supplyOrderItemPage = createSupplyOrderPage.Submit();
            supplyOrderItemPage.Go_To_New_Navigate();
            supplyOrderItemPage.SelectFirstItemNewTab();
            var itemName = supplyOrderItemPage.GetFirstItemName();
            var itemGeneralInformationPage = supplyOrderItemPage.EditItemNew(itemName);
            itemGeneralInformationPage.Go_To_New_Navigate(2);
            var packagingsupplie = itemGeneralInformationPage.GetFirstPackagingSupplier();
            var packagingsite = itemGeneralInformationPage.GetFirstPackagingSite();
            homePage.Navigate();
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, packagingsupplie);
            var supplieritem = suppliersPage.SelectFirstItem();
            var supplierDeliveriesTab = supplieritem.ClickOnDeliveriesTab();
            supplierDeliveriesTab.SetAmountDeliveredSites(packagingsite, "500000000");
            suppliersPage.Back_To_Navigate(1);
            supplyOrderItemPage.VerifypopupValidate();
            //Assert
            //7. Une pop-up récapitulative s'ouvre et peut afficher un message de vigileance
            if (supplyOrderItemPage.GetMessageButtonValidate() != "Validate")
            {
                //8.Cliquez sur "ignore and validate
                Assert.IsTrue(supplyOrderItemPage.GetMessageButtonValidate().Trim() == "Validate" || ((supplyOrderItemPage.GetMessageTextValidate().Trim() == "Some warning(s) have occured during validation :")
                    && (supplyOrderItemPage.GetMessageButtonValidate().Trim() == "Ignore these items and validate")), "La SO n est pas générée à partir de needs");
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SO_INDEX_CREATESOFROMPRODCO()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.ResetFilter();
            // Create New Supply Order
            var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
            createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
            createSupplyOrderPage.CheckQuantitiesFromProdCOs();
            var supplyOrderItemPage = createSupplyOrderPage.Submit();
            var itemName = supplyOrderItemPage.GetFirstItemName();
            supplyOrderItemPage.SelectFirstItem();
            supplyOrderItemPage.AddQuantity(itemName, "10");
            var itemGeneralInformationPage = supplyOrderItemPage.EditItem(itemName);
            var packagingsupplie = itemGeneralInformationPage.GetFirstPackagingSupplier();
            var packagingsite = itemGeneralInformationPage.GetFirstPackagingSite();
            homePage.Navigate();
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, packagingsupplie);
            var supplieritem = suppliersPage.SelectFirstItem();
            var supplierDeliveriesTab = supplieritem.ClickOnDeliveriesTab();
            supplierDeliveriesTab.SetAmountDeliveredSites(site, "500000000");
            suppliersPage.Back_To_Navigate(0);
            supplyOrderItemPage.VerifypopupValidate();


            //Assert
            if (supplyOrderItemPage.GetMessageButtonValidate() != "Validate")
            {
                Assert.IsTrue((supplyOrderItemPage.GetMessageTextValidate().Trim() == "Some warning(s) have occured during validation :")
                    && (supplyOrderItemPage.GetMessageButtonValidate().Trim() == "Ignore these items and validate"), "La SO n est pas générée à partir de needs");
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_INDEX_EXPORT()
        {
            //Prepare
            string site = TestContext.Properties["SiteMAD"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.ResetFilter();
            supplyOrderPage.Filter(FilterType.DateFrom, DateUtils.Now.AddDays(-10));
            supplyOrderPage.Filter(FilterType.DateTo, DateUtils.Now.AddDays(+20));
            supplyOrderPage.Filter(FilterType.Site, site);
            supplyOrderPage.Filter(FilterType.ShowNotActive, true);

            if (supplyOrderPage.CheckTotalNumber() < 200)
            {
                // On récupère les fichiers du répertoire de téléchargement
                supplyOrderPage.ClearDownloads();
                supplyOrderPage.Export(true);
                DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
                FileInfo[] taskFiles = taskDirectory.GetFiles();

                // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
                var correctDownloadedFile = supplyOrderPage.GetExportExcelFile(taskFiles);

                //Assert
                Assert.IsNotNull(correctDownloadedFile, "Le fichier excel n'est pas générer");
            }



        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SO_INDEX_CREATESOFROMCUSTOMERORDER()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();
            string articleQte = "15";
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.ResetFilter();
            // Create New Supply Order
            var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
            createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now.AddDays(-1), deliveryLocation);
            createSupplyOrderPage.CheckedPrefillQtesFromMenuCustomerorders();
            createSupplyOrderPage.SetDateFromCustomer(DateTime.ParseExact("19/07/2018", "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture));
            createSupplyOrderPage.CheckCustomer();
            var supplyOrderItemPage = createSupplyOrderPage.CreateButton();
            createSupplyOrderPage.IsErrorReport();
            supplyOrderItemPage.Go_To_New_Navigate();
            supplyOrderItemPage.SelectFirstItem();
            var firstItemName = supplyOrderItemPage.GetFirstItemName();

            supplyOrderItemPage.AddQuantity(firstItemName, articleQte);
            supplyOrderItemPage.Validate();
            var itemNameValidated = supplyOrderItemPage.GetFirstItemNameValidated();
            var itemQtyValidated = supplyOrderItemPage.GetFirstItemQtyValidated();
            //Assert
            Assert.IsTrue((itemNameValidated == firstItemName) && (articleQte == itemQtyValidated), "La supply order n'est pas générée à partir de 'Customerorders'.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SO_INDEX_CREATESOFROMPRODMANAGEMENT()
        {
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            filterAndFavorite.CreateNewSupplyOrder();
            filterAndFavorite.FillPrincipalField_GenerateSupplyOrder("ACE", DateUtils.Now, DateUtils.Now, DateUtils.Now, deliveryLocation);
            filterAndFavorite.Check_Perfill_Quantities_From_Production_Management();
            var NewSupplierNumber = filterAndFavorite.GetNewSupplierOrderNumber();
            var supplyOrderItemPage = filterAndFavorite.Submit();
            supplyOrderItemPage.Go_To_New_Navigate();
            string itemName = supplyOrderItemPage.GetFirstItemName();
            supplyOrderItemPage.SelectFirstItem();
            supplyOrderItemPage.EditQuantityRefresh(itemName, "5");
            supplyOrderItemPage.ResetFilter();
            supplyOrderItemPage.VerifypopupValidate();
            //supplyOrderItemPage.ResetFilter();
            // Retrieve and assert the warning message
            var warningMsg = supplyOrderItemPage.GetWarningValidationMessage();
            Assert.AreEqual("Some warning(s) have occured during validation :", warningMsg.Trim(), "Le message de warning attendu ne s'affiche pas.");
            supplyOrderItemPage.ClickOnIgnoreAndValidate();
            var supplyOrderPage = supplyOrderItemPage.BackToList();
            supplyOrderPage.ResetFilter();
            supplyOrderPage.Filter(SupplyOrderPage.FilterType.ByNumber, NewSupplierNumber);
            // Assert the creation and validation, since we are forcing data creation with "ignore and validate".
            Assert.IsTrue(supplyOrderPage.IsValidated(), "La SO n'est pas générée à partir de production management");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SO_CreateNewSupplyOrderFromOthers()
        {
            string keyword = TestContext.Properties["Item_Keyword"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
            createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
            var supplyOrderNumber = createSupplyOrderPage.GetSupplyOrderNumber();
            createSupplyOrderPage.FillFromAnotherSupplyOrderFucntion(site);
            var currentNumber = createSupplyOrderPage.GetSupplyOrderNumber();
            var itemPage = createSupplyOrderPage.Submit();
            supplyOrderPage = itemPage.BackToList();
            supplyOrderPage.Filter(SupplyOrderPage.FilterType.ByNumber, currentNumber);
            var getNumberItems = supplyOrderPage.CheckTotalNumber();
            Assert.IsTrue(getNumberItems > 0, "item creating error");




        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SO_Suppression_Totalw_o_Storage_Colomn()
        {
            var qty = new Random().Next(2, 20).ToString();

            HomePage homePage = LogInAsAdmin();

            // Navigate to Production Management and filter by date range
            FilterAndFavoritesPage filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.DateFrom, DateTime.Now);
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.DateTo, DateTime.Now.AddMonths(1));

            // Click on "Done" to get results
            ResultPage resultPage = filterAndFavorite.DoneToResults();

            // Create a new supply order and select the first item
            var deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();
            var supplyOrderNumber = resultPage.CreateSupplyOrder(deliveryLocation);
            SupplyOrderItem supplyOrderItem = resultPage.GenerateSupplyOrder();
            supplyOrderItem.SelectFirstItem();

            // Get the name of the first item and add quantity
            var itemName = supplyOrderItem.GetFirstItemName();
            supplyOrderItem.AddQuantity(itemName, qty);

            // Erase the quantity in the Prod. Qty column
            supplyOrderItem.EraseQuantity();

            // Get all Total w/o VAT and Storage values
            var totalItem = supplyOrderItem.GetAllTotalVat();
            var totalStorage = supplyOrderItem.GetAllStorage();

            // Assert Total w/o VAT columns are "0.0000"
            foreach (var item in totalItem)
            {
                if (double.TryParse(item, NumberStyles.Any, CultureInfo.InvariantCulture, out double itemValue))
                {
                    string formattedItem = itemValue.ToString("0.0000", CultureInfo.InvariantCulture);
                    Assert.AreEqual("0.0000", formattedItem, "La quantité n’est pas supprimée dans la colonne 'Total w/o VAT'.");
                }
                else
                {
                    Assert.Fail("Le format de la valeur 'Total w/o VAT' n’est pas valide.");
                }
            }

            // Assert Storage columns are empty or zero
            foreach (var storage in totalStorage)
            {
                Assert.IsTrue(string.IsNullOrEmpty(storage) || storage == "0", "La quantité n’est pas supprimée dans la colonne 'Storage'.");
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_CopyItems_from_another_SupplyOrders()
        {
            //Prepare
            string keyword = TestContext.Properties["Item_Keyword"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            SupplyOrderPage supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            CreateSupplyOrderModalPage createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
            createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
            createSupplyOrderPage.CopyItemsFromAnotherSupplyOrder();

            //Assert
            bool isAvailable = createSupplyOrderPage.IsDataAvailable();
            Assert.IsTrue(isAvailable, "La sélection de la Delivery Location empêche l'affichage des supply Order");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SO_Items_VerifyActualStock()
        {
            // Initialisation des variables de test
            string site = TestContext.Properties["SiteACE"].ToString();
            string deliveryLocation = TestContext.Properties["Place"].ToString();
            string quantity = "1";
            string itemName = null;
            string Number = null;


            // Connexion en tant qu'administrateur
            LogInAsAdmin();

            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Accéder à la page des Supply Orders
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            try
            {
                supplyOrderPage.ResetFilter();
                supplyOrderPage.Filter(FilterType.ShowNotValidated, true);

                // Créer une nouvelle Supply Order

                var createSupplyOrderModalPage = supplyOrderPage.CreateNewSupplyOrder();
                createSupplyOrderModalPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation, true, true);
                var supplyOrderItem = createSupplyOrderModalPage.Submit();

                // Sélectionner le premier élément de la Supply Order
                supplyOrderPage.SelectFirstItemSO();
                itemName = supplyOrderPage.GetItemName();

                // Modifier la quantité de l'article et rafraîchir la page
                supplyOrderItem.SetfirstQuantity(quantity);
                supplyOrderItem.RefreshPage();
                supplyOrderPage.SelectFirstItemSO();

                // Récupérer la quantité et l'unité depuis la Supply Order
                string combinedQtyUnit = supplyOrderPage.GetItemQuantityAndUnit();
                (string itemQty, string itemUnit) = supplyOrderPage.SplitQuantityAndUnit(combinedQtyUnit);
                supplyOrderItem.ClickOnGeneralInformation();
                Number = supplyOrderItem.GetNumber();
                // Aller à la page "Purchasing Item"
                var purchasingItemPage = homePage.GoToPurchasing_ItemPage();
                purchasingItemPage.ResetFilter();
                purchasingItemPage.Filter(ItemPage.FilterType.Search, itemName);
                purchasingItemPage.ClickOnFirstItem();

                // Récupérer la quantité et l'unité du stock réel dans la page "Purchasing Item"
                string actualQty = purchasingItemPage.GetItemQuantity();
                string actualUnit = purchasingItemPage.GetItemUnit();

                // Vérifier que la quantité et l'unité sont les mêmes que dans la Supply Order
                Assert.AreEqual(itemQty, actualQty, "La quantité de stock ne correspond pas.");
                Assert.AreEqual(itemUnit, actualUnit, "L'unité de stock ne correspond pas.");
            }
            finally
            {

                var supplyOrderPages = homePage.GoToPurchasing_SupplyOrderPage();
                supplyOrderPage.ResetFilter();
                supplyOrderPage.Filter(FilterType.ByNumber, Number);
                supplyOrderPage.DeleteFirstSupplyOrder();
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_ProdQty_Rounded()
        {
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();
            string qty = "1,5";

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.ResetFilter();
            var createSupplyOrderModalPage = supplyOrderPage.CreateNewSupplyOrder();
            createSupplyOrderModalPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation, true, true);
            var supplyOrderItem = createSupplyOrderModalPage.Submit();
            var itemName = supplyOrderItem.GetFirstItemName();
            supplyOrderItem.FIRSTITEMCLICK();
            supplyOrderItem.EditQuantityRefresh(itemName, qty);
            supplyOrderItem.WaitPageLoading();
            supplyOrderItem.Refresh();
            var roundedQuantity = supplyOrderItem.GetProdQuantity();

            Assert.AreEqual(roundedQuantity, "2", "La Prod Quantity n'a pas été arrondie correctement.");
        }

        /*[TestMethod]
        public void PU_SO_FD_NotRounded()
        {
            string location = TestContext.Properties["PlaceFrom"].ToString();

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var purchasingNeedsPage = homePage.GoToPurchasing_NeedsPage();
            var resultPageNeeds = purchasingNeedsPage.DoneToResults();
            resultPageNeeds.GenerateSupplyOrder(location, true);
            var supplyOrderItem = resultPageNeeds.Submit();

            bool isFDQuantityNotRounded = supplyOrderItem.IsFDQuantityNotRounded();
            
            Assert.IsTrue(isFDQuantityNotRounded, "Les quantités de F & D ont été arrondies comme Prod Quantity.");
        }*/

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_Export_SO()
        {
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            bool newVersionPrint = true;
            string site = "ACE";
            string coloneExcel = "Actual stock";
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.ResetFilter();

            var dayOffSet = 1;
            while (supplyOrderPage.CheckTotalNumber() > 200)
            {
                supplyOrderPage.Filter(FilterType.Site, site);
                supplyOrderPage.Filter(FilterType.DateFrom, DateUtils.Now.AddDays(-(dayOffSet + 1)));
                supplyOrderPage.Filter(FilterType.DateTo, DateUtils.Now.AddDays(-dayOffSet));
                dayOffSet++;
            }
            supplyOrderPage.ClearDownloads();
            supplyOrderPage.Export(newVersionPrint);
            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = supplyOrderPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            int resultNumber = OpenXmlExcel.GetExportResultNumber(SUPPLY_ORDERS_EXCEL_SHEET_NAME, filePath);
            var actualStockList = OpenXmlExcel.GetValuesInList(coloneExcel, SUPPLY_ORDERS_EXCEL_SHEET_NAME, filePath);
            Assert.IsTrue(actualStockList.All(res => res != string.Empty), "La colonne actual stock ne doit pas étre Vide");
            // On vide le répertoire de téléchargement
            DeleteAllFileDownload();
        }
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_ShowItems_NotSupplied_ProdQty()
        {
            var site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["Location"].ToString();

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.ResetFilter();
            var orderNumber = supplyOrderPage.GetFirstSONumber();
            supplyOrderPage.Filter(SupplyOrderPage.FilterType.ByNumber, orderNumber);

            var supplyOrderItemPage = supplyOrderPage.ClickFirstSupplyOrder();

            // Vérifier les quantités avant d'afficher les articles non fournis
            bool resultBefore = supplyOrderPage.GetQtyProd(); // Cette méthode doit retourner true si toutes les quantités >= 0
            supplyOrderPage.ClickShowItemsNotSupplied();
            supplyOrderPage.ClickFirstSupplyOrderAfter();

            // Vérifier les quantités après avoir affiché les articles non fournis
            bool resultAfter = supplyOrderPage.GetQtyProdAfterShowItem(); // Doit également retourner true si toutes les quantités >= 0

            // Assert
            Assert.IsTrue(resultBefore, "La quantité des produits avant l'affichage des articles non fournis n'est pas correcte.");
            Assert.IsTrue(resultAfter, "La quantité des produits après l'affichage des articles non fournis n'est pas correcte.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SO_Export_TotalCalculated()
        {
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            bool newVersionPrint = true;
            string site = "ACE";
            string coloneExcelW = "Total Calculated";
            string coloneExcelS = "Dispatch Qty";
            string coloneExcelU = "Packing Price (no VAT)";

            HomePage homePage = LogInAsAdmin();

            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.ResetFilter();

            var dayOffSet = 1;
            while (supplyOrderPage.CheckTotalNumber() > 200)
            {
                supplyOrderPage.Filter(FilterType.Site, site);
                supplyOrderPage.Filter(FilterType.DateFrom, DateUtils.Now.AddDays(-(dayOffSet + 1)));
                supplyOrderPage.Filter(FilterType.DateTo, DateUtils.Now.AddDays(-dayOffSet));
                dayOffSet++;
            }
            supplyOrderPage.ClearDownloads();
            supplyOrderPage.Export(newVersionPrint);
            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = supplyOrderPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            int resultNumber = OpenXmlExcel.GetExportResultNumber(SUPPLY_ORDERS_EXCEL_SHEET_NAME, filePath);
            var totalCalculated = OpenXmlExcel.GetValuesInList(coloneExcelW, SUPPLY_ORDERS_EXCEL_SHEET_NAME, filePath);
            var dispatchQty = OpenXmlExcel.GetValuesInList(coloneExcelS, SUPPLY_ORDERS_EXCEL_SHEET_NAME, filePath);
            var packingPrice = OpenXmlExcel.GetValuesInList(coloneExcelU, SUPPLY_ORDERS_EXCEL_SHEET_NAME, filePath);


            for (int i = 0; i < totalCalculated.Count; i++)
            {
                var result = supplyOrderPage.CheckTotalCalculated(totalCalculated[i], dispatchQty[i], packingPrice[i]);
            }

            bool totalCal = totalCalculated.All(res => res != string.Empty);
            Assert.IsTrue(totalCal, "La colonne actual stock ne doit pas étre Vide");
            // On vide le répertoire de téléchargement
            DeleteAllFileDownload();
        }

        [TestMethod]
        public void PU_SO_Suppression_Recipe()
        {
            string recipeName = TestContext.Properties["RecipeName"].ToString();
            string site = TestContext.Properties["SiteMAD"].ToString();


            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.Filter(FilterType.ShowNotActive, true);

            var supplyOrderItemPage = supplyOrderPage.ClickFirstSupplyOrder();

            var supplyOrderRecipePage = supplyOrderItemPage.ClickOnRecipes();

            supplyOrderRecipePage.AddRecipeSO(recipeName);
            Assert.IsTrue(supplyOrderRecipePage.IsFirstRecipeAdded(), "La recette n'a pas été ajoutée dans le supply order.");

            supplyOrderRecipePage.DeleteFirstRecipe();
            Assert.IsTrue(supplyOrderRecipePage.IsFirstRecipeDeleted(), "La recette n'a pas été supprimée correctement.");

        }
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_CreateSO_FromAnotherSupplyOrder()
        {
            //Prepare          
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();

            //1) Create new SO
            CreateSupplyOrderModalPage newSupplyOrder = supplyOrderPage.CreateNewSupplyOrder();
            newSupplyOrder.GetSupplyOrderNumber();
            newSupplyOrder.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation, true);
            var id = newSupplyOrder.GetSupplyOrderNumber();

            //2) Copy items from another Supply Order
            newSupplyOrder.CopyItemsFromAnotherSupplyOrder();
            newSupplyOrder.SelectFirstSO();

            //3) Submit the new Supply Order
            newSupplyOrder.Submit();
            newSupplyOrder.BackToList();

            //Filter on number and select the first item on the list
            supplyOrderPage.ResetFilter();
            supplyOrderPage.Filter(FilterType.ByNumber, id);

            //Assert
            Assert.AreEqual(1, supplyOrderPage.CheckTotalNumber(), "Le supply order créé n'apparaît pas dans la liste des supply orders.");

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_Suppression_Index()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            SupplyOrderPage supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();

            supplyOrderPage.ResetFilter();

            var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
            createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
            var id = createSupplyOrderPage.GetSupplyOrderNumber();
            var supplyOrderItemPage = createSupplyOrderPage.Submit();
            try
            {
                supplyOrderPage = supplyOrderItemPage.BackToList();
                //Filter on number and select the first item on the list
                supplyOrderPage.ResetFilter();
                supplyOrderPage.Filter(FilterType.ByNumber, id);
                supplyOrderPage.DeleteFirstSupplyOrder();
                supplyOrderPage.ResetFilter();
                supplyOrderPage.Filter(FilterType.ByNumber, id);
                Assert.AreEqual(0, supplyOrderPage.CheckTotalNumber(), "Le supply order n'a pas été supprimé.");
            }
            finally
            {
                supplyOrderPage.ResetFilter();
                supplyOrderPage.Filter(FilterType.ByNumber, id);
                if (supplyOrderPage.CheckTotalNumber() > 0)
                {
                    supplyOrderPage.DeleteFirstSupplyOrder();
                }
            }
        }
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_From_SO_Other_Site_ProdQty()
        {
            //Prepare          
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();

            //1) Create new SO first
            CreateSupplyOrderModalPage newSupplyOrder = supplyOrderPage.CreateNewSupplyOrder();
            newSupplyOrder.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation, true);

            //Number de so crée 
            var NumberSO = newSupplyOrder.GetSupplyOrderNumber();

            //2) Copy items from another Supply Order
            newSupplyOrder.CopyItemsFromAnotherSupplyOrder();

            var itemPage = new ItemPage(WebDriver, TestContext);
            newSupplyOrder.SelectFirstSO();

            //3) Submit the new Supply Order
            newSupplyOrder.Submit();

            //Quantity de so Create First
            var QtyNumberSOCreatedFirst = supplyOrderPage.GetItemQuantityProd();
            newSupplyOrder.BackToList();
            supplyOrderPage.ResetFilter();

            //4) Create new SO second
            supplyOrderPage.CreateNewSupplyOrder();
            newSupplyOrder.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation, true);

            //5) Copy items from another Supply Order
            newSupplyOrder.CopyItemsFromAnotherSupplyOrder();

            //Filter with Number so original 
            newSupplyOrder.filterNumberInPageCreateSO(NumberSO);
            newSupplyOrder.SelectFirstSO();
            //6) Submit the new Supply Order
            newSupplyOrder.Submit();

            //Quantity so Created Second
            var QtyNumberSOCreatedSecond = supplyOrderPage.GetItemQuantityProd();
            newSupplyOrder.BackToList();
            supplyOrderPage.ResetFilter();

            // ASSERTIONS: Verify that the quantities match between the original SO and the new SOs
            Assert.AreEqual(QtyNumberSOCreatedFirst, QtyNumberSOCreatedSecond, "The quantity of the original SO does not match the new SO quantity.");
        }
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_Suppression_ProdQty_Column()
        {
            // Récupérer les paramètres du contexte de test
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();
            DateTime dtfrom = DateTime.Today.AddDays(-11);
            DateTime dtto = DateTime.Today;

            // Arrange
            var homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.DateFrom, dtfrom);
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.DateTo, dtto);
            QuantityAdjustmentsPage qtyAdjustmentPage = filterAndFavorite.DoneToQtyAjustement();
            CreateSupplyOrderModalPage createSupplyOrderModalPage = filterAndFavorite.CreateNewSupplyOrder();
            createSupplyOrderModalPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
            string originalWindow = WebDriver.CurrentWindowHandle;
            SupplyOrderItem supplyOrderItem = createSupplyOrderModalPage.Submit();
            supplyOrderItem.SwitchToNewWindow(originalWindow);
            supplyOrderItem.WaitPageLoading();
            supplyOrderItem.EditFirstItem("10");
            supplyOrderItem.Refresh();
            var ProdQtyBeforeErase = supplyOrderItem.GetProdQty();
            SupplyOrderRecipes supplyOrderRecipes = supplyOrderItem.ClickOnRecipes();
            supplyOrderRecipes.ClickOnItems();
            supplyOrderItem.EraseQuantity();
            var ProdQtyAfterErase = supplyOrderItem.GetProdQty();
            Assert.AreNotEqual(ProdQtyAfterErase, ProdQtyBeforeErase, "La quantité n'est pas modifiée");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_RefreshCountSO()
        {
            string dateToString = DateTime.Now.Date.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
            DateTime dateTo = DateTime.ParseExact(dateToString, "dd/MM/yyyy", CultureInfo.InvariantCulture);
            string dateFromString = DateTime.Now.Date.AddDays(-7).ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
            DateTime dateFrom = DateTime.ParseExact(dateFromString, "dd/MM/yyyy", CultureInfo.InvariantCulture);
            HomePage homePage = LogInAsAdmin();
            SupplyOrderPage supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.ResetFilter();
            int totalCountBeforeFilter = supplyOrderPage.CheckTotalNumber();
            supplyOrderPage.Filter(FilterType.ShowNotActive, true);
            supplyOrderPage.Filter(FilterType.DateFrom, dateFrom);
            supplyOrderPage.Filter(FilterType.DateTo, dateTo);
            int totalCountAfterFilter = supplyOrderPage.CheckTotalNumber();
            Assert.AreNotEqual(totalCountAfterFilter, totalCountBeforeFilter, "Le filtre ne fonctionne pas correctely !");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_Modification_Qty_Colomn_F_D()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["Location"].ToString();
            string qty = new Random().Next(2, 20).ToString();
            DateTime dateFrom = DateTime.Now.AddDays(-1);
            DateTime dateTo = DateTime.Now.AddDays(1);
            //Arrange
            var homePage = LogInAsAdmin();

            var filterAndFavorite = homePage.GoToProduction_ProductionManagemenentPage();
            filterAndFavorite.ResetFilter();
            filterAndFavorite.Filter(FilterAndFavoritesPage.FilterType.Site, site);
            var qtyAjustementPage = filterAndFavorite.DoneToQtyAjustement();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.DateFrom, dateFrom);

            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.DateTo, dateTo);

            var resultPage = qtyAjustementPage.GoToResultPage();
            // Génération supply order
            resultPage.CreateSupplyOrder(deliveryLocation);

            SupplyOrderItem supplyOrderItem = resultPage.GenerateSupplyOrder();
            var FirstItemFD_Before = supplyOrderItem.GetFirstItemFD();
            supplyOrderItem.SelectFirstItem();

            // Get the name of the first item and add quantity           
            var itemName = supplyOrderItem.GetFirstItemName();
            supplyOrderItem.AddQuantity(itemName, qty);
            supplyOrderItem.Refresh();
            var FirstItemFD_After = supplyOrderItem.GetFirstItemFD();
            Assert.AreEqual(FirstItemFD_Before, FirstItemFD_After, "la quantité du F&D a été changé");

        }
     
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_SearchStartDate()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.ResetFilter();

            if (supplyOrderPage.CheckTotalNumber() < 20)
            {
                //Create SupplyOrder 1
                var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
                createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(20), DateUtils.Now, deliveryLocation);
                var supplyOrderItemPage = createSupplyOrderPage.Submit();

                var itemName = supplyOrderItemPage.GetFirstItemName();
                supplyOrderItemPage.SelectFirstItem();
                supplyOrderItemPage.AddQuantity(itemName, "10");

                supplyOrderPage = supplyOrderItemPage.BackToList();
               
                //Create SupplyOrder 2
                createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
                createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now.AddDays(1), DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
                supplyOrderItemPage = createSupplyOrderPage.Submit();

                itemName = supplyOrderItemPage.GetFirstItemName();
                supplyOrderItemPage.SelectFirstItem();
                supplyOrderItemPage.AddQuantity(itemName, "10");

                supplyOrderPage = supplyOrderItemPage.BackToList();
                supplyOrderPage.ResetFilter();
            }

            if (!supplyOrderPage.isPageSizeEqualsTo100())
            {
                supplyOrderPage.PageSize("8");
                supplyOrderPage.PageSize("100");
            }
            //SortBy StartDate
            supplyOrderPage.Filter(FilterType.SortBy, "DATE");
            var isSortedByStartDateOK = supplyOrderPage.IsSortedByStartDate();

            //Assert
            Assert.IsTrue(isSortedByStartDateOK, "Les supply orders ne sont pas rangées par start date.");


        }
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_FilterSearchNumber()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.ResetFilter();

            if (supplyOrderPage.CheckTotalNumber() < 20)
            {
                //Create SupplyOrder 1
                var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
                createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(20), DateUtils.Now, deliveryLocation);
                var supplyOrderItemPage = createSupplyOrderPage.Submit();

                var itemName = supplyOrderItemPage.GetFirstItemName();
                supplyOrderItemPage.SelectFirstItem();
                supplyOrderItemPage.AddQuantity(itemName, "10");

                supplyOrderPage = supplyOrderItemPage.BackToList();

                //Create SupplyOrder 2
                createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
                createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now.AddDays(1), DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
                supplyOrderItemPage = createSupplyOrderPage.Submit();

                itemName = supplyOrderItemPage.GetFirstItemName();
                supplyOrderItemPage.SelectFirstItem();
                supplyOrderItemPage.AddQuantity(itemName, "10");

                supplyOrderPage = supplyOrderItemPage.BackToList();
                supplyOrderPage.ResetFilter();
            }

            if (!supplyOrderPage.isPageSizeEqualsTo100())
            {
                supplyOrderPage.PageSize("8");
                supplyOrderPage.PageSize("100");
            }
            //SortBy StartDate
            supplyOrderPage.Filter(FilterType.SortBy, "NUMBER");
            var isSortedByNumberOK = supplyOrderPage.IsSortedByNumber();

            //Assert
            Assert.IsTrue(isSortedByNumberOK, "Les supply orders ne sont pas rangées par Number.");


        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_SO_FilterSearchEndDate()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string deliveryLocation = TestContext.Properties["PlaceFrom"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.ResetFilter();

            if (supplyOrderPage.CheckTotalNumber() < 20)
            {
                //Create SupplyOrder 1
                var createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
                createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now, DateUtils.Now.AddDays(20), DateUtils.Now, deliveryLocation);
                var supplyOrderItemPage = createSupplyOrderPage.Submit();

                var itemName = supplyOrderItemPage.GetFirstItemName();
                supplyOrderItemPage.SelectFirstItem();
                supplyOrderItemPage.AddQuantity(itemName, "10");

                supplyOrderPage = supplyOrderItemPage.BackToList();

                //Create SupplyOrder 2
                createSupplyOrderPage = supplyOrderPage.CreateNewSupplyOrder();
                createSupplyOrderPage.FillPrincipalField_CreationNewSupplyOrder(site, DateUtils.Now.AddDays(1), DateUtils.Now.AddDays(30), DateUtils.Now, deliveryLocation);
                supplyOrderItemPage = createSupplyOrderPage.Submit();

                itemName = supplyOrderItemPage.GetFirstItemName();
                supplyOrderItemPage.SelectFirstItem();
                supplyOrderItemPage.AddQuantity(itemName, "10");

                supplyOrderPage = supplyOrderItemPage.BackToList();
                supplyOrderPage.ResetFilter();
            }

            if (!supplyOrderPage.isPageSizeEqualsTo100())
            {
                supplyOrderPage.PageSize("8");
                supplyOrderPage.PageSize("100");
            }
            //SortBy EndDate
            supplyOrderPage.Filter(FilterType.SortBy, "ENDDATE");
            var isSortedByEndDateOK = supplyOrderPage.IsSortedByEndDate();

            //Assert
            Assert.IsTrue(isSortedByEndDateOK, "Les supply orders ne sont pas rangées par EndDate.");

        }
       

    }
   
}