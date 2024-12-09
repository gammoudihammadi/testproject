using DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.PriceList;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Recipes;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using System;
using System.IO;
using System.IO.Packaging;
using System.Runtime.InteropServices.ComTypes;
using System.Security.Policy;
using System.Threading;
using System.Web.UI;
using System.Windows.Forms;

namespace Newrest.Winrest.FunctionalTests.Customers
{
    [TestClass]
    public class PriceListTest : TestBase
    {

        private const int _timeout = 600000;
        private const string PRICE_LIST_EXCEL_SHEET_NAME = "Sheet 1";
        private readonly string PRICE_PERIOD_ERROR_MESSAGE = "There is already an existing default pricing for this site within this period";
        private readonly string priceNameToday = "Price-" + DateUtils.Now.ToString("dd/MM/yyyy") + "-";
        private readonly string itemNameForForeignTest = "Item Foreign Test";
        private readonly string recipeNameForForeignTest = "Recipe Foreign Test";
        private readonly string itemCommercialNameForForeignTest = "Item Commercial Name Foreign Test";
        private readonly string recipeCommercialNameForForeignTest = "Recipe Commercial Name Foreign Test";
        private readonly string DELIVERY_NAME = "Test CU_PRLIST";
        private readonly string priceListName = "Price-" + DateUtils.Now.ToString("dd/MM/yyyy") + "-foreignTest";
        private readonly string COMMERCIAL_NAME_FOR_CO_TEST = "Item CN-" + DateUtils.Now.ToString("dd/MM/yyyy") + "-coTest";


        [TestMethod]
        [Priority(0)]
        [Timeout(_timeout)]
        public void CU_PRLIST_SetConfigWinrest4_0()
        {
            //Prepare
            var keyword = TestContext.Properties["Item_Keyword"].ToString();
            var site = TestContext.Properties["Site"].ToString();

            //Arrange
            HomePage homePage =  LogInAsAdmin();

            // New version search
            homePage.SetNewVersionSearchValue(true);

            // New keyword search
            homePage.SetNewVersionKeywordValue(true);

            // New VersionItemDetails
            //homePage.SetVersionItemDetailValue(true);

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

        }

        [Priority(1)]
        [TestMethod]
        [Timeout(_timeout)]
        public void CU_PRLIST_CreateItemsForPriceList()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string itemName = TestContext.Properties["ItemNamePriceList"].ToString();
            string itemName2 = TestContext.Properties["ItemNamePriceListBis"].ToString();

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

            // ----------------------------------------- Item 1 ---------------------------------------------------------------
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);

            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);

                // 1 packaging pour le site ACE
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);

                // 1 packaging pour le site MAD
                itemGeneralInformationPage.NewPackaging();
                itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                itemPage = itemGeneralInformationPage.BackToList();
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
            }

            string item1 = itemPage.GetFirstItemName();

            // ----------------------------------------- Item 2 ---------------------------------------------------------------
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName2);

            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName2, group, workshop, taxType, prodUnit);

                // 1 packaging pour le site ACE
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);

                // 1 packaging pour le site MAD
                itemGeneralInformationPage.NewPackaging();
                itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                itemPage = itemGeneralInformationPage.BackToList();
                itemPage.Filter(ItemPage.FilterType.Search, itemName2);
            }

            string item2 = itemPage.GetFirstItemName();


            // Assert
            Assert.AreEqual(itemName, item1, "L'item " + itemName + " n'existe pas.");
            Assert.AreEqual(itemName2, item2, "L'item " + itemName2 + " n'existe pas.");
        }

        [Priority(2)]
        [TestMethod]
        [Timeout(_timeout)]
        public void CU_PRLIST_CreateRecipeForPriceList()
        {
            //Prepare recipe
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariant"].ToString();
            string ingredient = TestContext.Properties["ItemNamePriceList"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string recipeName = TestContext.Properties["RecipePriceList"].ToString();
            int nbPortions = new Random().Next(1, 30);

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);

            if (recipesPage.CheckTotalNumber() > 0)
            {
                recipesPage.MassiveDeleteRecipes(recipeName, site, recipeType);

            }
            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());
            recipeGeneralInfosPage.AddVariantWithSite(site, variant);
            var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
            recipeVariantPage.AddIngredient(ingredient);
            recipesPage = recipeVariantPage.BackToList();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);

            //Assert
            var firstRecipeName = recipesPage.GetFirstRecipeName();
            Assert.AreEqual(recipeName, firstRecipeName, "La recette n'a pas été ajoutée à la liste.");
        }

        // Créer une nouvelle liste de prix
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_PRLIST_Create_New_Pricing()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();

            Random rnd = new Random();
            string priceName = priceNameToday + rnd.Next(1, 10000).ToString();
            DateTime startDate = DateUtils.Now;
            DateTime endDate = DateUtils.Now.AddDays(1);


            // Arrange
            HomePage homePage=LogInAsAdmin();
          

            // Act
            var priceListPage = homePage.GoToCustomers_PriceListPage();
            priceListPage.ResetFilter();
            priceListPage.Filter(PriceListPage.FilterType.Site, site);
            try
            {
                var pricingCreateModalpage = priceListPage.CreateNewPricing();
                var priceListDetails = pricingCreateModalpage.FillField_CreateNewPricing(priceName, site, startDate, endDate);
                priceListPage = priceListDetails.BackToList();

                priceListPage.Filter(PriceListPage.FilterType.Site, site);
                priceListPage.Filter(PriceListPage.FilterType.Search, priceName);

                // Assert
                Assert.AreEqual(priceName, priceListPage.GetFirstPricingName(), "Le price list créé n'apparaît pas dans la liste.");
            }
            finally
            {
                priceListPage.Filter(PriceListPage.FilterType.Site, site);
                priceListPage.Filter(PriceListPage.FilterType.Search, priceName);
                priceListPage.DeletePricing(priceName);
            }
        }

        //Créer une  liste de prix déjà existante
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_PRLIST_Create_New_Pricing_Already_Existing()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();

            Random rnd = new Random();
            string priceName = priceNameToday + rnd.Next(1, 10000).ToString();
            DateTime startDate = DateUtils.Now;
            DateTime endDate = DateUtils.Now.AddDays(1);


            // Arrange
            HomePage homePage=LogInAsAdmin();
           

            // Act
            var priceListPage = homePage.GoToCustomers_PriceListPage();
            priceListPage.ResetFilter();
            priceListPage.Filter(PriceListPage.FilterType.Site, site);

            if (priceListPage.CheckTotalNumber() > 0)
            {
                priceListPage.DeleteAllPriceList();
            }

            try
            {
                var pricingCreateModalpage = priceListPage.CreateNewPricing();
                var priceListDetails = pricingCreateModalpage.FillField_CreateNewPricing(priceName, site, startDate, endDate, true);
                priceListPage = priceListDetails.BackToList();

                pricingCreateModalpage = priceListPage.CreateNewPricing();
                pricingCreateModalpage.FillField_CreateNewPricing(priceName, site, startDate, endDate, true);

                // Assert
                Assert.IsTrue(pricingCreateModalpage.ErrorMessageNameAlreadyExists(), "Aucun message d'erreur renvoyé alors qu'un price list identique existe déjà.");
            }
            finally
            {
                DeletePriceList(homePage, site);
            }
        }

        // Créer une nouvelle liste de prix sans remplir les champs obligatoires
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_PRLIST_Create_New_Pricing_Without_Fill_Fields()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            HomePage homePage =   LogInAsAdmin();

            // Act
            var priceListPage = homePage.GoToCustomers_PriceListPage();
            priceListPage.ResetFilter();
            priceListPage.Filter(PriceListPage.FilterType.Site, site);

            if (priceListPage.CheckTotalNumber() > 0)
            {
                priceListPage.DeleteAllPriceList();
            }

            try
            {
                var pricingCreateModalpage = priceListPage.CreateNewPricing();
                pricingCreateModalpage.CreatePricing();

                // Assert
                Assert.IsTrue(pricingCreateModalpage.ErrorMessageNameRequired(), "Aucun message d'erreur alors que les champs obligatoires n'ont pas été renseignés.");
            }
            finally
            {
                DeletePriceList(homePage, site);
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_PRLIST_Filter_Search()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();

            Random rnd = new Random();
            string priceName = priceNameToday + rnd.Next(1, 10000).ToString();
            DateTime startDate = DateUtils.Now;
            DateTime endDate = DateUtils.Now.AddDays(1);


            // Arrange
            HomePage homePage =  LogInAsAdmin();

            // Act
            var priceListPage = homePage.GoToCustomers_PriceListPage();
            priceListPage.ResetFilter();
            priceListPage.Filter(PriceListPage.FilterType.Site, site);

            if (priceListPage.CheckTotalNumber() > 0)
            {
                priceListPage.DeleteAllPriceList();
            }

            try
            {
                var pricingCreateModalpage = priceListPage.CreateNewPricing();
                var priceListDetails = pricingCreateModalpage.FillField_CreateNewPricing(priceName, site, startDate, endDate);
                priceListPage = priceListDetails.BackToList();

                priceListPage.Filter(PriceListPage.FilterType.Site, site);
                priceListPage.Filter(PriceListPage.FilterType.Search, priceName);

                // Assert
                Assert.AreEqual(priceName, priceListPage.GetFirstPricingName(), String.Format(MessageErreur.FILTRE_ERRONE, "Search"));
            }
            finally
            {
                DeletePriceList(homePage, site);
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_PRLIST_Filter_Site()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string siteBis = TestContext.Properties["SiteLP"].ToString();

            Random rnd = new Random();
            string priceName = priceNameToday + rnd.Next(1, 10000).ToString();
            string priceNameBis = priceNameToday + rnd.Next(1, 10000).ToString();

            DateTime startDate = DateUtils.Now;
            DateTime endDate = DateUtils.Now.AddDays(1);


            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var priceListPage = homePage.GoToCustomers_PriceListPage();
            priceListPage.ResetFilter();
            priceListPage.Filter(PriceListPage.FilterType.Site, site);

            if (priceListPage.CheckTotalNumber() > 0)
            {
                priceListPage.DeleteAllPriceList();
            }

            priceListPage.Filter(PriceListPage.FilterType.Site, siteBis);

            if (priceListPage.CheckTotalNumber() > 0)
            {
                priceListPage.DeleteAllPriceList();
            }

            try
            {
                var pricingCreateModalpage = priceListPage.CreateNewPricing();
                var priceListDetails = pricingCreateModalpage.FillField_CreateNewPricing(priceName, site, startDate, endDate);
                priceListPage = priceListDetails.BackToList();

                pricingCreateModalpage = priceListPage.CreateNewPricing();
                priceListDetails = pricingCreateModalpage.FillField_CreateNewPricing(priceNameBis, siteBis, startDate, endDate);
                priceListPage = priceListDetails.BackToList();


                priceListPage.Filter(PriceListPage.FilterType.Site, site);
                priceListPage.Filter(PriceListPage.FilterType.Search, priceName);

                Assert.AreEqual(1, priceListPage.CheckTotalNumber(), "Le nombre de price list trouvé est incorrect pour les filtres appliqués.");
                Assert.AreEqual(priceName, priceListPage.GetFirstPricingName(), String.Format(MessageErreur.FILTRE_ERRONE, "Site"));

                priceListPage.ResetFilter();
                priceListPage.Filter(PriceListPage.FilterType.Site, siteBis);
                priceListPage.Filter(PriceListPage.FilterType.Search, priceNameBis);

                Assert.AreEqual(1, priceListPage.CheckTotalNumber(), "Le nombre de price list trouvé est incorrect pour les filtres appliqués.");
                Assert.AreEqual(priceNameBis, priceListPage.GetFirstPricingName(), String.Format(MessageErreur.FILTRE_ERRONE, "Site"));
            }
            finally
            {
                DeletePriceList(homePage, site);

                DeletePriceList(homePage, siteBis);
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_PRLIST_Filter_SortByName()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();

            Random rnd = new Random();
            string priceName = priceNameToday + rnd.Next(1, 10000).ToString();
            string priceNameBis = priceNameToday + rnd.Next(1, 10000).ToString();
            string filterName = "Name";
            DateTime startDate = DateUtils.Now;
            DateTime endDate = DateUtils.Now.AddDays(1);

            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var priceListPage = homePage.GoToCustomers_PriceListPage();
            priceListPage.ResetFilter();
            priceListPage.Filter(PriceListPage.FilterType.Site, site);
            try
            {
                // creer 1er price
                var pricingCreateModalpage = priceListPage.CreateNewPricing();
                var priceListDetails = pricingCreateModalpage.FillField_CreateNewPricing(priceName, site, startDate, endDate);
                priceListPage = priceListDetails.BackToList();
                // creér 2eme price
                pricingCreateModalpage = priceListPage.CreateNewPricing();
                priceListDetails = pricingCreateModalpage.FillField_CreateNewPricing(priceNameBis, site, startDate, endDate);
                priceListPage = priceListDetails.BackToList();
                priceListPage.Filter(PriceListPage.FilterType.SortBy, filterName);
                bool isSortedByName = priceListPage.IsSortedByName();
                Assert.IsTrue(isSortedByName, String.Format(MessageErreur.FILTRE_ERRONE, "Sort by name"));
            }
            finally
            {
                priceListPage.Filter(PriceListPage.FilterType.Search, priceName);
                priceListPage.DeletePricing(priceName);
                priceListPage.Filter(PriceListPage.FilterType.Search, priceNameBis);
                priceListPage.DeletePricing(priceNameBis);
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_PRLIST_Filter_SortBySite()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string site2 = TestContext.Properties["SiteACE"].ToString();

            Random rnd = new Random();
            string priceName = priceNameToday + rnd.Next(1, 1000).ToString();
            string priceNameBis = priceNameToday + rnd.Next(1001, 2000).ToString();
            string priceName2 = "Price" + rnd.Next(2001, 3000).ToString();
            string filterSite = "Site";
            DateTime startDate = DateUtils.Now;
            DateTime endDate = DateUtils.Now.AddDays(1);

            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var priceListPage = homePage.GoToCustomers_PriceListPage();
            priceListPage.ResetFilter();
            priceListPage.Filter(PriceListPage.FilterType.Site, site);
            try
            {
                // creer 1er price
                var pricingCreateModalpage = priceListPage.CreateNewPricing();
                var priceListDetails = pricingCreateModalpage.FillField_CreateNewPricing(priceName, site, startDate, endDate);
                priceListPage = priceListDetails.BackToList();
                // creér 2eme price
                priceListPage.Filter(PriceListPage.FilterType.Site, site2);
                pricingCreateModalpage = priceListPage.CreateNewPricing();
                priceListDetails = pricingCreateModalpage.FillField_CreateNewPricing(priceNameBis, site2, startDate, endDate);
                priceListPage = priceListDetails.BackToList();
                //creér 3eme price
                pricingCreateModalpage = priceListPage.CreateNewPricing();
                priceListDetails = pricingCreateModalpage.FillField_CreateNewPricing(priceName2, site2, startDate, endDate);
                priceListPage = priceListDetails.BackToList();
                priceListPage.Filter(PriceListPage.FilterType.SortBy, filterSite);
                bool isSortedBySite = priceListPage.IsSortedBySite();
                Assert.IsTrue(isSortedBySite, String.Format(MessageErreur.FILTRE_ERRONE, "Sort by site"));
            }
            finally
            {
                priceListPage.Filter(PriceListPage.FilterType.Site, site);
                priceListPage.Filter(PriceListPage.FilterType.Search, priceName);
                priceListPage.DeletePricing(priceName);
                priceListPage.Filter(PriceListPage.FilterType.Site, site2);
                priceListPage.Filter(PriceListPage.FilterType.Search, priceNameBis);
                priceListPage.DeletePricing(priceNameBis);
                priceListPage.Filter(PriceListPage.FilterType.Search, priceName2);
                priceListPage.DeletePricing(priceName2);

            }
        }


        [Timeout(_timeout)]
        [TestMethod]
        public void CU_PRLIST_Change_General_Informations()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();

            Random rnd = new Random();
            string priceName = priceNameToday + rnd.Next(1, 10000).ToString();
            DateTime startDate = DateUtils.Now;
            DateTime endDate = DateUtils.Now.AddDays(1);

            string newPriceName = priceNameToday + "modified";
            DateTime newStartDate = DateUtils.Now.AddMonths(1);
            DateTime newEndDate = DateUtils.Now.AddMonths(2);


            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var dateFormat = homePage.GetDateFormatPickerValue();

            // Act
            var priceListPage = homePage.GoToCustomers_PriceListPage();
            priceListPage.ResetFilter();
            priceListPage.Filter(PriceListPage.FilterType.Site, site);

            if (priceListPage.CheckTotalNumber() > 0)
            {
                priceListPage.DeleteAllPriceList();
            }

            try
            {
                var pricingCreateModalpage = priceListPage.CreateNewPricing();
                var priceListDetails = pricingCreateModalpage.FillField_CreateNewPricing(priceName, site, startDate, endDate);

                var generalInformationPage = priceListDetails.ClickOnGeneralInformation();

                generalInformationPage.ModifyGeneralInformations(newPriceName, newStartDate, newEndDate);
                priceListPage = generalInformationPage.BackToList();

                priceListPage.ResetFilter();
                priceListPage.Filter(PriceListPage.FilterType.Site, site);
                priceListPage.Filter(PriceListPage.FilterType.Search, newPriceName);

                priceListDetails = priceListPage.ClickOnFirstPricing();
                generalInformationPage = priceListDetails.ClickOnGeneralInformation();

                Assert.AreEqual(newPriceName, generalInformationPage.GetPriceName(), "La modification du nom du price list n'est pas effective.");
                Assert.AreEqual(newStartDate.Date, generalInformationPage.GetPriceStartDate(dateFormat).Date, "La modification de la start date du price list n'est pas effective.");
                Assert.AreEqual(newEndDate.Date, generalInformationPage.GetPriceEndDate(dateFormat).Date, "La modification de la end date du price list n'est pas effective.");
            }
            finally
            {
                DeletePriceList(homePage, site);
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_PRLIST_Add_Item()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string itemName = TestContext.Properties["ItemNamePriceList"].ToString();

            Random rnd = new Random();
            string priceName = priceNameToday + rnd.Next(1, 10000).ToString();
            DateTime startDate = DateUtils.Now;
            DateTime endDate = DateUtils.Now.AddDays(1);


            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var priceListPage = homePage.GoToCustomers_PriceListPage();
            priceListPage.ResetFilter();
            priceListPage.Filter(PriceListPage.FilterType.Site, site);

            if (priceListPage.CheckTotalNumber() > 0)
            {
                priceListPage.DeleteAllPriceList();
            }

            try
            {
                var pricingCreateModalpage = priceListPage.CreateNewPricing();
                var priceListDetails = pricingCreateModalpage.FillField_CreateNewPricing(priceName, site, startDate, endDate);

                priceListDetails.AddItem(itemName, true);
                WebDriver.Navigate().Refresh();

                Assert.IsTrue(priceListDetails.IsItemAdded(itemName), "L'item " + itemName + "n'a pas été ajouté au price list.");
            }
            finally
            {
                DeletePriceList(homePage, site);
            }
        }

        // Ajouter une recette
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_PRLIST_Add_Recipe()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string recipeName = TestContext.Properties["RecipePriceList"].ToString();

            Random rnd = new Random();
            string priceName = priceNameToday + rnd.Next(1, 10000).ToString();
            DateTime startDate = DateUtils.Now;
            DateTime endDate = DateUtils.Now.AddDays(1);


            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var priceListPage = homePage.GoToCustomers_PriceListPage();
            priceListPage.ResetFilter();
            priceListPage.Filter(PriceListPage.FilterType.Site, site);

            if (priceListPage.CheckTotalNumber() > 0)
            {
                priceListPage.DeleteAllPriceList();
            }

            try
            {
                var pricingCreateModalpage = priceListPage.CreateNewPricing();
                var priceListDetails = pricingCreateModalpage.FillField_CreateNewPricing(priceName, site, startDate, endDate);

                priceListDetails.AddItem(recipeName, false);
                WebDriver.Navigate().Refresh();

                Assert.IsTrue(priceListDetails.IsItemAdded(recipeName), "La recette " + recipeName + "n'a pas été ajoutée au price list.");
            }
            finally
            {
                DeletePriceList(homePage, site);
            }
        }

        // Modifier les informations d'un item
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_PRLIST_Update_Item()
        {

            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string itemName = TestContext.Properties["ItemNamePriceList"].ToString();

            Random rnd = new Random();
            string priceName = priceNameToday + rnd.Next(1, 10000).ToString();
            DateTime startDate = DateUtils.Now;
            DateTime endDate = DateUtils.Now.AddDays(1);

            double initialPrice = 55;

            // Log in
            HomePage homePage = LogInAsAdmin();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            // Act
            var priceListPage = homePage.GoToCustomers_PriceListPage();
            priceListPage.ResetFilter();
            priceListPage.Filter(PriceListPage.FilterType.Site, site);

            if (priceListPage.CheckTotalNumber() > 0)
            {
                priceListPage.DeleteAllPriceList();
            }

            try
            {
                var pricingCreateModalpage = priceListPage.CreateNewPricing();
                var priceListDetails = pricingCreateModalpage.FillField_CreateNewPricing(priceName, site, startDate, endDate);

                priceListDetails.AddItem(itemName, true);
                WebDriver.Navigate().Refresh();

                Assert.IsTrue(priceListDetails.IsItemAdded(itemName), "L'item " + itemName + "n'a pas été ajouté au price list.");

                priceListDetails.UpdateItem(initialPrice.ToString(), 0);

                Assert.AreEqual(initialPrice, priceListDetails.GetInitialPrice(decimalSeparatorValue), "L'update de l'initial price de l'item a échoué.");

            }
            finally
            {
                DeletePriceList(homePage, site);
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_PRLIST_Add_Production_Comment()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string itemName = TestContext.Properties["ItemNamePriceList"].ToString();
            string comment = "This is a comment";

            Random rnd = new Random();
            string priceName = priceNameToday + rnd.Next(1, 10000).ToString();
            DateTime startDate = DateUtils.Now;
            DateTime endDate = DateUtils.Now.AddDays(1);

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var priceListPage = homePage.GoToCustomers_PriceListPage();
            priceListPage.ResetFilter();
            priceListPage.Filter(PriceListPage.FilterType.Site, site);

            if (priceListPage.CheckTotalNumber() > 0)
            {
                priceListPage.DeleteAllPriceList();
            }

            try
            {
                var pricingCreateModalpage = priceListPage.CreateNewPricing();
                var priceListDetails = pricingCreateModalpage.FillField_CreateNewPricing(priceName, site, startDate, endDate);

                priceListDetails.AddItem(itemName, true);
                WebDriver.Navigate().Refresh();

                Assert.IsTrue(priceListDetails.IsItemAdded(itemName), "L'item " + itemName + "n'a pas été ajouté au price list.");

                priceListDetails.AddProductionComment(comment);
                WebDriver.Navigate().Refresh();

                Assert.AreEqual(comment, priceListDetails.GetProductionComment(), "Le production comment n'a pas été ajouté.");
            }
            finally
            {
                DeletePriceList(homePage, site);
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_PRLIST_Add_Billing_Comment()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string itemName = TestContext.Properties["ItemNamePriceList"].ToString();
            string comment = "This is a comment";

            Random rnd = new Random();
            string priceName = priceNameToday + rnd.Next(1, 10000).ToString();
            DateTime startDate = DateUtils.Now;
            DateTime endDate = DateUtils.Now.AddDays(1);


            // Arrange
            HomePage homePage =  LogInAsAdmin();

            // Act
            var priceListPage = homePage.GoToCustomers_PriceListPage();
            priceListPage.ResetFilter();
            priceListPage.Filter(PriceListPage.FilterType.Site, site);

            if (priceListPage.CheckTotalNumber() > 0)
            {
                priceListPage.DeleteAllPriceList();
            }

            try
            {
                var pricingCreateModalpage = priceListPage.CreateNewPricing();
                var priceListDetails = pricingCreateModalpage.FillField_CreateNewPricing(priceName, site, startDate, endDate);

                priceListDetails.AddItem(itemName, true);
                WebDriver.Navigate().Refresh();

                Assert.IsTrue(priceListDetails.IsItemAdded(itemName), "L'item " + itemName + "n'a pas été ajouté au price list.");

                priceListDetails.AddBillingComment(comment);
                WebDriver.Navigate().Refresh();

                Assert.AreEqual(comment, priceListDetails.GetBillingComment(), "Le billing comment n'a pas été ajouté.");
            }
            finally
            {
                DeletePriceList(homePage, site);
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_PRLIST_Delete_Item()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string itemName = TestContext.Properties["ItemNamePriceList"].ToString();

            Random rnd = new Random();
            string priceName = priceNameToday + rnd.Next(1, 10000).ToString();
            DateTime startDate = DateUtils.Now;
            DateTime endDate = DateUtils.Now.AddDays(1);

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var priceListPage = homePage.GoToCustomers_PriceListPage();
            priceListPage.ResetFilter();
            priceListPage.Filter(PriceListPage.FilterType.Site, site);

            if (priceListPage.CheckTotalNumber() > 0)
            {
                priceListPage.DeleteAllPriceList();
            }

            try
            {
                var pricingCreateModalpage = priceListPage.CreateNewPricing();
                var priceListDetails = pricingCreateModalpage.FillField_CreateNewPricing(priceName, site, startDate, endDate);

                priceListDetails.AddItem(itemName, true);
                WebDriver.Navigate().Refresh();

                Assert.IsTrue(priceListDetails.IsItemAdded(itemName), "L'item " + itemName + "n'a pas été ajouté au price list.");

                priceListDetails.DeleteItem();
                WebDriver.Navigate().Refresh();

                Assert.IsFalse(priceListDetails.IsItemAdded(itemName), "L'item " + itemName + "n'a pas été supprimé du price list.");
            }
            finally
            {
                DeletePriceList(homePage, site);
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_PRLIST_Export_Prices_NewVersion()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string itemName = TestContext.Properties["ItemNamePriceList"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            Random rnd = new Random();
            string priceName = priceNameToday + rnd.Next(1, 10000).ToString();
            DateTime startDate = DateUtils.Now;
            DateTime endDate = DateUtils.Now.AddDays(1);

            bool newVersionPrint = true;

            // Arrange
            HomePage homePage= LogInAsAdmin();
          

            ClearCache();

            // Act
            var priceListPage = homePage.GoToCustomers_PriceListPage();

            priceListPage.ClearDownloads();
            

            priceListPage.ResetFilter();
            try
            {
                var pricingCreateModalpage = priceListPage.CreateNewPricing();
                var priceListDetails = pricingCreateModalpage.FillField_CreateNewPricing(priceName, site, startDate, endDate);

                priceListDetails.AddItem(itemName, true);
                priceListDetails.ExportPriceList(newVersionPrint);

                DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
                FileInfo[] taskFiles = taskDirectory.GetFiles();

                var correctDownloadedFile = priceListDetails.GetExportExcelFile(taskFiles);
                Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

                var fileName = correctDownloadedFile.Name;

                // On récupère le nombre de résultats issus du filtre et l'URL du fichier Excel à ouvrir
                var filePath = Path.Combine(downloadsPath, fileName);

                // Exploitation du fichier Excel
                int resultNumber = OpenXmlExcel.GetExportResultNumber(PRICE_LIST_EXCEL_SHEET_NAME, filePath);
                bool result = OpenXmlExcel.ReadAllDataInColumn("Name", PRICE_LIST_EXCEL_SHEET_NAME, filePath, itemName);

                //Assert
                Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);
                Assert.IsTrue(result, MessageErreur.EXCEL_DONNEES_KO);

            }
            finally
            {
                DeletePriceList(homePage, site);
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_PRLIST_Import_Prices()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();

            Random rnd = new Random();
            string priceName = priceNameToday + rnd.Next(1, 10000).ToString();
            DateTime startDate = DateUtils.Now;
            DateTime endDate = DateUtils.Now.AddDays(1);

            var path = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase);
            path = path.Substring(6);
            var ExcelPath = path + "\\PageObjects\\Customers\\PriceList\\PR_PRLI_import_test.xlsx";

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var priceListPage = homePage.GoToCustomers_PriceListPage();
            priceListPage.ResetFilter();
            priceListPage.Filter(PriceListPage.FilterType.Site, site);

            if (priceListPage.CheckTotalNumber() > 0)
            {
                priceListPage.DeleteAllPriceList();
            }

            try
            {
                var pricingCreateModalpage = priceListPage.CreateNewPricing();
                var priceListDetails = pricingCreateModalpage.FillField_CreateNewPricing(priceName, site, startDate, endDate);

                Assert.AreEqual(0, priceListDetails.CountItems(), "Il y a déjà un item dans le price list.");

                var pricingImportModalPage = priceListDetails.PriceListImportPage();
                priceListDetails = pricingImportModalPage.ImportFile(ExcelPath);

                Assert.AreEqual(1, priceListDetails.CountItems(), "L'item n'a pas été importé.");
                Assert.IsTrue(priceListDetails.GetItemName().Contains("TestImport"), "L'item importé ne correspond pas à celui du fichier.");

            }
            finally
            {
                DeletePriceList(homePage, site);
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_PRLIST_Add_New_Period()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string itemName = TestContext.Properties["ItemNamePriceList"].ToString();

            Random rnd = new Random();
            string priceName = priceNameToday + rnd.Next(1, 10000).ToString();
            DateTime startDate = DateUtils.Now;
            DateTime endDate = DateUtils.Now.AddDays(1);

            DateTime startDatePeriod = DateUtils.Now.AddDays(+760);
            DateTime endDatePeriod = DateUtils.Now.AddDays(+800);

            // Arrange
            HomePage homePage =  LogInAsAdmin();

            // Act
            var priceListPage = homePage.GoToCustomers_PriceListPage();
            priceListPage.ResetFilter();
            priceListPage.Filter(PriceListPage.FilterType.Site, site);

            if (priceListPage.CheckTotalNumber() > 0)
            {
                priceListPage.DeleteAllPriceList();
            }

            try
            {
                var pricingCreateModalpage = priceListPage.CreateNewPricing();
                var priceListDetails = pricingCreateModalpage.FillField_CreateNewPricing(priceName, site, startDate, endDate);

                priceListDetails.AddItem(itemName, true);

                string initPeriod = priceListDetails.GetPeriod();

                Assert.AreEqual(1, priceListDetails.CountPeriod(), "Le price list contient déjà plusieurs périodes.");

                priceListDetails.AddNewPeriod(startDatePeriod, endDatePeriod);
                string finalPeriod = priceListDetails.GetPeriod();

                Assert.AreEqual(2, priceListDetails.CountPeriod(), "La nouvelle période n'a pas été ajoutée au price list.");

                // Assert 
                Assert.IsTrue(initPeriod.Contains(startDate.ToString("dd/MM/yyyy")) && initPeriod.Contains(endDate.ToString("dd/MM/yyyy")), "La période initiale ne correspond pas aux dates définies pour le price list.");
                Assert.IsTrue(finalPeriod.Contains(startDatePeriod.ToString("dd/MM/yyyy")) && finalPeriod.Contains(endDatePeriod.ToString("dd/MM/yyyy")), "La période finale ne correspond pas aux dates définies pour l'ajout de la nouvelle période.");
            }
            finally
            {
                DeletePriceList(homePage, site);
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_PRLIST_Add_New_Period_Already_Existing()
        {

            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string itemName = TestContext.Properties["ItemNamePriceList"].ToString();

            Random rnd = new Random();
            string priceName = priceNameToday + rnd.Next(1, 10000).ToString();
            DateTime startDate = DateUtils.Now;
            DateTime endDate = DateUtils.Now.AddDays(1);

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var priceListPage = homePage.GoToCustomers_PriceListPage();
            priceListPage.ResetFilter();
            priceListPage.Filter(PriceListPage.FilterType.Site, site);

            if (priceListPage.CheckTotalNumber() > 0)
            {
                priceListPage.DeleteAllPriceList();
            }

            try
            {
                var pricingCreateModalpage = priceListPage.CreateNewPricing();
                var priceListDetails = pricingCreateModalpage.FillField_CreateNewPricing(priceName, site, startDate, endDate);

                priceListDetails.AddItem(itemName, true);

                priceListDetails.AddNewPeriod(startDate, endDate);

                Assert.IsTrue(priceListDetails.ErrorMessageNewPeriod(PRICE_PERIOD_ERROR_MESSAGE), "l'ajout d'une période existante n'entraîne pas l'apparition de message d'erreur.");
            }
            finally
            {
                DeletePriceList(homePage, site);
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_PRLIST_Change_Period()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string itemName = TestContext.Properties["ItemNamePriceList"].ToString();

            Random rnd = new Random();
            string priceName = priceNameToday + rnd.Next(1, 10000).ToString();
            DateTime startDate = DateUtils.Now;
            DateTime endDate = DateUtils.Now.AddDays(1);

            DateTime startDatePeriod = DateUtils.Now.AddDays(+760);
            DateTime endDatePeriod = DateUtils.Now.AddDays(+800);

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var priceListPage = homePage.GoToCustomers_PriceListPage();
            priceListPage.ResetFilter();
            priceListPage.Filter(PriceListPage.FilterType.Site, site);

            if (priceListPage.CheckTotalNumber() > 0)
            {
                priceListPage.DeleteAllPriceList();
            }

            try
            {
                var pricingCreateModalpage = priceListPage.CreateNewPricing();
                var priceListDetails = pricingCreateModalpage.FillField_CreateNewPricing(priceName, site, startDate, endDate);

                priceListDetails.AddItem(itemName, true);

                priceListDetails.AddNewPeriod(startDatePeriod, endDatePeriod);
                string initPeriod = priceListDetails.GetPeriod();

                Assert.AreEqual(2, priceListDetails.CountPeriod(), "La nouvelle période n'a pas été ajoutée au price list.");
                Assert.IsTrue(initPeriod.Contains(startDatePeriod.ToString("dd/MM/yyyy")) && initPeriod.Contains(endDatePeriod.ToString("dd/MM/yyyy")), "La période ajoutée n'est pas visible pour le price list.");

                priceListDetails.ChangePeriod(startDate.ToString("dd/MM/yyyy"));
                string finalPeriod = priceListDetails.GetPeriod();

                // Assert 
                Assert.IsTrue(finalPeriod.Contains(startDate.ToString("dd/MM/yyyy")) && finalPeriod.Contains(endDate.ToString("dd/MM/yyyy")), "La période visualisée n'a pas été modifiée.");

            }
            finally
            {
                DeletePriceList(homePage, site);
            }
        }

        // Déployer les prix

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_PRLIST_Unfold_Prices()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();

            Random rnd = new Random();
            string priceName = priceNameToday + rnd.Next(1, 10000).ToString();
            DateTime startDate = DateUtils.Now;
            DateTime endDate = DateUtils.Now.AddDays(1);


            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var priceListPage = homePage.GoToCustomers_PriceListPage();
            priceListPage.ResetFilter();
            priceListPage.Filter(PriceListPage.FilterType.Site, site);

            if (priceListPage.CheckTotalNumber() > 0)
            {
                priceListPage.DeleteAllPriceList();
            }

            try
            {
                var pricingCreateModalpage = priceListPage.CreateNewPricing();
                var priceListDetails = pricingCreateModalpage.FillField_CreateNewPricing(priceName, site, startDate, endDate);
                priceListPage = priceListDetails.BackToList();

                priceListPage.Filter(PriceListPage.FilterType.Site, site);
                priceListPage.Filter(PriceListPage.FilterType.Search, priceName);

                priceListPage.UnfoldFold();

                // Assert
                Assert.IsTrue(priceListPage.IsUnfold(), "Les prix ne sont pas déployés.");
            }
            finally
            {
                DeletePriceList(homePage, site);
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_PRLIST_Fold_Prices()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();

            Random rnd = new Random();
            string priceName = priceNameToday + rnd.Next(1, 10000).ToString();
            DateTime startDate = DateUtils.Now;
            DateTime endDate = DateUtils.Now.AddDays(1);


            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var priceListPage = homePage.GoToCustomers_PriceListPage();
            priceListPage.ResetFilter();
            priceListPage.Filter(PriceListPage.FilterType.Site, site);

            if (priceListPage.CheckTotalNumber() > 0)
            {
                priceListPage.DeleteAllPriceList();
            }

            try
            {
                var pricingCreateModalpage = priceListPage.CreateNewPricing();
                var priceListDetails = pricingCreateModalpage.FillField_CreateNewPricing(priceName, site, startDate, endDate);
                priceListPage = priceListDetails.BackToList();

                priceListPage.Filter(PriceListPage.FilterType.Site, site);
                priceListPage.Filter(PriceListPage.FilterType.Search, priceName);

                priceListPage.UnfoldFold();
                Assert.IsTrue(priceListPage.IsUnfold(), "Les prix ne sont pas déployés.");

                priceListPage.UnfoldFold();
                Assert.IsTrue(priceListPage.IsFold(), "Les prix ne sont pas masqués.");
            }
            finally
            {
                DeletePriceList(homePage, site);
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_PRLIST_Delete_Price()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();

            Random rnd = new Random();
            string priceName = priceNameToday + rnd.Next(1, 10000).ToString();
            DateTime startDate = DateUtils.Now;
            DateTime endDate = DateUtils.Now.AddDays(1);


            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var priceListPage = homePage.GoToCustomers_PriceListPage();
            priceListPage.ResetFilter();
            priceListPage.Filter(PriceListPage.FilterType.Site, site);

            if (priceListPage.CheckTotalNumber() > 0)
            {
                priceListPage.DeleteAllPriceList();
            }

            try
            {
                var pricingCreateModalpage = priceListPage.CreateNewPricing();
                var priceListDetails = pricingCreateModalpage.FillField_CreateNewPricing(priceName, site, startDate, endDate);
                priceListPage = priceListDetails.BackToList();

                priceListPage.Filter(PriceListPage.FilterType.Site, site);
                priceListPage.Filter(PriceListPage.FilterType.Search, priceName);

                Assert.AreEqual(1, priceListPage.CheckTotalNumber(), "Aucun item ne correspond aux critères.");

                priceListPage.DeletePricing(priceName);

                Assert.AreEqual(0, priceListPage.CheckTotalNumber(), "L'item n'a pas été supprimé.");
            }
            finally
            {
                DeletePriceList(homePage, site);
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_PRLIST_Delete_Period()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string itemName = TestContext.Properties["ItemNamePriceList"].ToString();

            Random rnd = new Random();
            string priceName = priceNameToday + rnd.Next(1, 10000).ToString();
            DateTime startDate = DateUtils.Now;
            DateTime endDate = DateUtils.Now.AddDays(1);

            DateTime startDatePeriod = DateUtils.Now.AddDays(+760);
            DateTime endDatePeriod = DateUtils.Now.AddDays(+800);

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var priceListPage = homePage.GoToCustomers_PriceListPage();
            priceListPage.ResetFilter();
            priceListPage.Filter(PriceListPage.FilterType.Site, site);

            if (priceListPage.CheckTotalNumber() > 0)
            {
                priceListPage.DeleteAllPriceList();
            }

            try
            {
                var pricingCreateModalpage = priceListPage.CreateNewPricing();
                var priceListDetails = pricingCreateModalpage.FillField_CreateNewPricing(priceName, site, startDate, endDate);

                priceListDetails.AddItem(itemName, true);
                priceListDetails.AddNewPeriod(startDatePeriod, endDatePeriod);
                Assert.AreEqual(2, priceListDetails.CountPeriod(), "La nouvelle période n'a pas été ajoutée au price list.");

                priceListPage = priceListDetails.BackToList();
                priceListPage.Filter(PriceListPage.FilterType.Site, site);
                priceListPage.Filter(PriceListPage.FilterType.Search, priceName);

                priceListPage.UnfoldFold();

                int periodNumber = priceListPage.CountPeriod();
                Assert.AreNotEqual(0, periodNumber, "Le price list possède au moins une période.");

                priceListPage.DeletePeriod();

                int finalPeriodNumber = priceListPage.CountPeriod();
                Assert.AreNotEqual(finalPeriodNumber, periodNumber, "La période n'a pas été supprimée.");
            }
            finally
            {
                DeletePriceList(homePage, site);
            }
        }

        // Filtre Search Price
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_PRLIST_Filter_Price_Search()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string itemName = TestContext.Properties["ItemNamePriceList"].ToString();
            string secondItem = TestContext.Properties["ItemNamePriceListBis"].ToString();

            Random rnd = new Random();
            string priceName = priceNameToday + rnd.Next(1, 10000).ToString();
            DateTime startDate = DateUtils.Now;
            DateTime endDate = DateUtils.Now.AddDays(1);

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var priceListPage = homePage.GoToCustomers_PriceListPage();
            priceListPage.ResetFilter();
            priceListPage.Filter(PriceListPage.FilterType.Site, site);

            if (priceListPage.CheckTotalNumber() > 0)
            {
                priceListPage.DeleteAllPriceList();
            }

            try
            {
                var pricingCreateModalpage = priceListPage.CreateNewPricing();
                var priceListDetails = pricingCreateModalpage.FillField_CreateNewPricing(priceName, site, startDate, endDate);

                priceListDetails.AddItem(itemName, true);
                priceListDetails.AddItem(secondItem, true);
                WebDriver.Navigate().Refresh();

                Assert.IsTrue(priceListDetails.IsItemAdded(itemName), "L'item " + itemName + "n'a pas été ajouté au price list.");
                Assert.IsTrue(priceListDetails.IsItemAdded(secondItem), "L'item " + secondItem + "n'a pas été ajouté au price list.");

                // On filtre pour obtenir que le premier item
                priceListDetails.ResetFilter();
                priceListDetails.Filter(PriceListDetailsPage.FilterType.Search, itemName);
                Thread.Sleep(2000);

                Assert.IsTrue((priceListDetails.CountItems() == 1) && (priceListDetails.GetItemName().Contains(itemName)), string.Format(MessageErreur.FILTRE_ERRONE, "Search items"));

                // On filtre pour obtenir que le second item
                priceListDetails.ResetFilter();
                priceListDetails.Filter(PriceListDetailsPage.FilterType.Search, secondItem);
                Thread.Sleep(2000);

                Assert.IsTrue((priceListDetails.CountItems() == 1) && (priceListDetails.GetItemName().Contains(secondItem)), string.Format(MessageErreur.FILTRE_ERRONE, "Search items"));
            }
            finally
            {
                DeletePriceList(homePage, site);
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_PRLIST_Filter_Price_Show_Without_Price()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string itemName = TestContext.Properties["ItemNamePriceList"].ToString();
            string secondItem = TestContext.Properties["ItemNamePriceListBis"].ToString();

            Random rnd = new Random();
            string priceName = priceNameToday + rnd.Next(1, 10000).ToString();
            DateTime startDate = DateUtils.Now;
            DateTime endDate = DateUtils.Now.AddDays(1);

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var priceListPage = homePage.GoToCustomers_PriceListPage();
            priceListPage.ResetFilter();
            priceListPage.Filter(PriceListPage.FilterType.Site, site);

            if (priceListPage.CheckTotalNumber() > 0)
            {
                priceListPage.DeleteAllPriceList();
            }

            try
            {
                var pricingCreateModalpage = priceListPage.CreateNewPricing();
                var priceListDetails = pricingCreateModalpage.FillField_CreateNewPricing(priceName, site, startDate, endDate);

                priceListDetails.ResetFilter();
                priceListDetails.AddItem(itemName, true);
                priceListDetails.AddCoef("1", itemName);
                priceListDetails.UpdateItem("2", 0);

                priceListDetails.AddItem(secondItem, true);
                priceListDetails.AddCoef("1", secondItem);
                priceListDetails.UpdateItem("0", 1);


                WebDriver.Navigate().Refresh();

                Assert.IsTrue(priceListDetails.IsItemAdded(itemName), "L'item " + itemName + "n'a pas été ajouté au price list.");
                Assert.IsTrue(priceListDetails.IsItemAdded(secondItem), "L'item " + secondItem + "n'a pas été ajouté au price list.");

                // On filtre pour obtenir que le premier item
                priceListDetails.Filter(PriceListDetailsPage.FilterType.ShowWithoutPriceOnly, true);
                Thread.Sleep(2000);
                Assert.IsTrue((priceListDetails.CountItems() == 1) && (priceListDetails.GetItemName().Contains(secondItem)), string.Format(MessageErreur.FILTRE_ERRONE, "ShowWithoutPriceOnly"));
            }
            finally
            {
                DeletePriceList(homePage, site);
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_PRLIST_Filter_Price_Show_Negative_Markup()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string itemName = TestContext.Properties["ItemNamePriceList"].ToString();

            Random rnd = new Random();
            string priceName = priceNameToday + rnd.Next(1, 10000).ToString();
            DateTime startDate = DateUtils.Now;
            DateTime endDate = DateUtils.Now.AddDays(1);
            string coef = "0.8";

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            // Act
            var priceListPage = homePage.GoToCustomers_PriceListPage();
            priceListPage.ResetFilter();
            priceListPage.Filter(PriceListPage.FilterType.Site, site);

            if (priceListPage.CheckTotalNumber() > 0)
            {
                priceListPage.DeleteAllPriceList();
            }

            try
            {
                var pricingCreateModalpage = priceListPage.CreateNewPricing();
                var priceListDetails = pricingCreateModalpage.FillField_CreateNewPricing(priceName, site, startDate, endDate);

                priceListDetails.AddItem(itemName, true);
                WebDriver.Navigate().Refresh();

                Assert.IsTrue(priceListDetails.IsItemAdded(itemName), "L'item " + itemName + "n'a pas été ajouté au price list.");
                Assert.IsTrue(priceListDetails.GetMarkup(decimalSeparatorValue) == 0, "L'item ne possède pas un markup égal à 0 par défaut.");

                priceListDetails.Filter(PriceListDetailsPage.FilterType.ShowNegativeMarkup, true);
                priceListDetails.WaitPageLoading();
                Assert.AreEqual(0, priceListDetails.CountItems(), "L'item de markup égal à 0 apparaît alors que le filtre Show negative markup est activé.");

                priceListDetails.Filter(PriceListDetailsPage.FilterType.ShowNegativeMarkup, false);

                // Ajout d'un coef < 1 pour l'item pour afficher un markup négatif
                priceListDetails.AddCoef(coef, itemName);
                WebDriver.Navigate().Refresh();

                Assert.IsTrue(priceListDetails.GetMarkup(decimalSeparatorValue) < 0, "L'item ne possède pas un markup égal négatif.");

                // On applique le filtre
                priceListDetails.Filter(PriceListDetailsPage.FilterType.ShowNegativeMarkup, true);

                Assert.IsTrue((priceListDetails.CountItems() == 1) && (priceListDetails.GetItemName().Contains(itemName)), string.Format(MessageErreur.FILTRE_ERRONE, "Show negative markup"));
            }
            finally
            {
                DeletePriceList(homePage, site);
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_PRLIST_Filter_Price_Element_Types()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string itemName = TestContext.Properties["ItemNamePriceList"].ToString();
            string recette = TestContext.Properties["RecipePriceList"].ToString();

            Random rnd = new Random();
            string priceName = priceNameToday + rnd.Next(1, 10000).ToString();
            DateTime startDate = DateUtils.Now;
            DateTime endDate = DateUtils.Now.AddDays(1);

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var priceListPage = homePage.GoToCustomers_PriceListPage();
            priceListPage.ResetFilter();
            priceListPage.Filter(PriceListPage.FilterType.Site, site);

            if (priceListPage.CheckTotalNumber() > 0)
            {
                priceListPage.DeleteAllPriceList();
            }

            try
            {
                var pricingCreateModalpage = priceListPage.CreateNewPricing();
                var priceListDetails = pricingCreateModalpage.FillField_CreateNewPricing(priceName, site, startDate, endDate);

                priceListDetails.AddItem(itemName, true);
                priceListDetails.AddItem(recette, false);
                WebDriver.Navigate().Refresh();

                Assert.IsTrue(priceListDetails.IsItemAdded(itemName), "L'item " + itemName + "n'a pas été ajouté au price list.");
                Assert.IsTrue(priceListDetails.IsItemAdded(recette), "La recette " + recette + "n'a pas été ajoutée au price list.");

                // On filtre pour obtenir que l'item
                priceListDetails.ShowElementTypesFilter();
                priceListDetails.Filter(PriceListDetailsPage.FilterType.ElementTypeItems, true);
                priceListDetails.WaitPageLoading();
                Assert.IsTrue((priceListDetails.CountItems() == 1) && (priceListDetails.GetItemName().Contains(itemName)), string.Format(MessageErreur.FILTRE_ERRONE, "Element types Items"));

                priceListDetails.Filter(PriceListDetailsPage.FilterType.ElementTypeRecipes, true);
                priceListDetails.WaitPageLoading();
                Assert.IsTrue((priceListDetails.CountItems() == 1) && (priceListDetails.GetItemName().Contains(recette)), string.Format(MessageErreur.FILTRE_ERRONE, "Element types Recipes"));

                priceListDetails.Filter(PriceListDetailsPage.FilterType.ElementTypeFreePrices, true);
                priceListDetails.WaitPageLoading();
                Assert.IsTrue((priceListDetails.CountItems() == 0), string.Format(MessageErreur.FILTRE_ERRONE, "Element types Free prices"));

                priceListDetails.Filter(PriceListDetailsPage.FilterType.ElementTypeAll, true);
                priceListDetails.WaitPageLoading();
                Assert.IsTrue((priceListDetails.CountItems() == 2), string.Format(MessageErreur.FILTRE_ERRONE, "Element types All"));
            }
            finally
            {
                DeletePriceList(homePage, site);
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_PRLIST_item_update()
        {
            //Prepare
            string itemSubGroup = TestContext.Properties["SupplierInvoiceItemGroupBis"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            Random rnd = new Random();
            int randomNumber = rnd.Next(1000, 10000);
            var newPriceListName = "newPriceList" + "-" + randomNumber;
            var newItemName = "newItem" + "-" + randomNumber;
            var updatedItemName = "updatedItem" + "-" + randomNumber;
            string supplierName = "A REFERENCIAR";

            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string storageUnit = "KG";
            string supplierRef = newItemName + "_SupplierRef";
            string storageQty = 2.32.ToString();
            var site = "ACE";

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            var itemCreatePage = itemPage.ItemCreatePage();
            var itemGeneralInformation = itemCreatePage.FillField_CreateNewItem(newItemName, itemSubGroup, workshop, taxType, prodUnit);
            var itemCreateNewPackaging = itemGeneralInformation.NewPackaging();
            // FIXME storageQty avec une virgule... 2,32
            itemCreateNewPackaging.FillField_CreateNewPackaging(
                    site, packagingName, 2.32.ToString(), storageUnit, 4.2.ToString(), supplierName, (string)null, supplierRef);

            var priceListPage = homePage.GoToCustomers_PriceListPage();
            var priceCreatePage = priceListPage.CreateNewPricing();
            var priceListDetailPage = priceCreatePage.FillField_CreateNewPricing(newPriceListName, site, DateTime.Now, DateTime.Now.AddDays(1));
            priceListDetailPage.AddItem(newItemName, true);

            itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.Filter(ItemPage.FilterType.Search, newItemName);
            itemGeneralInformation = itemPage.ClickOnFirstItem();
            itemGeneralInformation.UpdateItem(updatedItemName, itemSubGroup, workshop, taxType, prodUnit);

            priceListPage = homePage.GoToCustomers_PriceListPage();
            priceListPage.Filter(PriceListPage.FilterType.Search, newPriceListName);
            priceListDetailPage = priceListPage.ClickOnFirstPricing();
            var priceListItemName = priceListDetailPage.GetItemName();
            var result = priceListItemName.Contains(updatedItemName);

            //Assert
            Assert.IsTrue(result, "item name is not same after update");
        }


        [TestMethod]
        [Priority(3)]
        [Timeout(_timeout)]
        public void CU_PRLIST_CreateItemAndRecipeWithDifferentCommercialName()
        {
            //Prepare
            //Item
            string site = TestContext.Properties["Site"].ToString(); //MAD
            string itemName = itemNameForForeignTest;
            string itemCommercialName = itemCommercialNameForForeignTest;

            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();

            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();
            //Recipe
            string recipe = recipeNameForForeignTest;
            string recipeCommercialName = recipeCommercialNameForForeignTest;
            string variant = TestContext.Properties["RecipeVariant"].ToString();
            string ingredient = TestContext.Properties["ItemNamePriceList"].ToString();
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            int nbPortions = new Random().Next(1, 30);

            //Arrange
            HomePage homePage =  LogInAsAdmin();

            //Act
            //Go to ItemPage and create a new Item if it does not exist
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.ShowAll, true);
            itemPage.Filter(ItemPage.FilterType.Search, itemName);

            if (itemPage.CheckTotalNumber() == 0)
            {
                var createPage = itemPage.ItemCreatePage();
                var filledForm = createPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit, commercialName: itemCommercialName);

                // 1 packaging pour le site MAD
                var itemCreatePackagingPage = filledForm.NewPackaging();
                var itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                itemPage = itemGeneralInformationPage.BackToList();
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
            }

            //Verify Item has been created and has packaging
            Assert.AreEqual(itemName, itemPage.GetFirstItemName());

            //Go to RecipePage and create a new recipe if it does not exist
            var recipePage = homePage.GoToMenus_Recipes();
            recipePage.Filter(RecipesPage.FilterType.ShowAll, true);
            recipePage.Filter(RecipesPage.FilterType.SearchRecipe, recipe);

            if (recipePage.CheckTotalNumber() == 0)
            {
                var createRecipePage = recipePage.CreateNewRecipe();
                var recipeGeneralInfosPage = createRecipePage.FillField_CreateNewRecipe(recipe, recipeType, nbPortions.ToString(), commercialName: recipeCommercialName);

                recipeGeneralInfosPage.AddVariantWithSite(site, variant);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                recipeVariantPage.AddIngredient(ingredient);
                recipePage = recipeVariantPage.BackToList();
                recipePage.ResetFilter();
                recipePage.Filter(RecipesPage.FilterType.SearchRecipe, recipe);
            }

            //Verify Recipe has been created
            Assert.AreEqual(recipe, recipePage.GetFirstRecipeName());
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_PRLIST_Foreign()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string itemCommercialName = itemCommercialNameForForeignTest;
            string recipeCommercialName = recipeCommercialNameForForeignTest;

            DateTime startDate = DateUtils.Now;
            DateTime endDate = DateUtils.Now.AddDays(1);

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var priceListPage = homePage.GoToCustomers_PriceListPage();
            priceListPage.ResetFilter();
            priceListPage.Filter(PriceListPage.FilterType.Site, site);
            priceListPage.Filter(PriceListPage.FilterType.Search, priceListName);

            //Create pricing if it does not exist
            try
            {
                if (priceListPage.CheckTotalNumber() == 0)
                {
                    var pricingModal = priceListPage.CreateNewPricing();
                    var priceListDetails = pricingModal.FillField_CreateNewPricing(priceListName, site, startDate, endDate, false);

                    priceListDetails.AddItem(itemNameForForeignTest, true);
                    priceListDetails.AddItem(recipeNameForForeignTest, false);

                    priceListDetails.AddNewPeriod(startDate, endDate);
                    if (priceListDetails.ErrorMessageNewPeriod(PRICE_PERIOD_ERROR_MESSAGE))
                    {
                        priceListDetails.ClosePeriodPopup();
                    }
                    priceListDetails.BackToList();
                    priceListPage.ResetFilter();
                    priceListPage.Filter(PriceListPage.FilterType.Site, site);
                    priceListPage.Filter(PriceListPage.FilterType.Search, priceListName);
                }
            }
            catch (Exception)
            {
                DeletePriceList(homePage, site);
                throw;
            }

            PriceListDetailsPage detailsPage = priceListPage.ClickOnFirstPricing();
            detailsPage.ClickOnForeign();

            //Verify if the Name 'N' displayed is the commercialName of either the recipe or the item (item will be first in list)
            var itemCommercialNameDisplayed = detailsPage.GetNameByRow(1);
            var recipeCommercialNameDisplayed = detailsPage.GetNameByRow(2);

            Assert.AreEqual(itemCommercialNameDisplayed, itemCommercialName);
            Assert.AreEqual(recipeCommercialNameDisplayed, recipeCommercialNameForForeignTest);
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_PRLIST_item_CN_in_CO()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString(); //MAD
            string customer = "A.F.B. LANZAROTE"; // TestContext.Properties["CustomerDelivery"].ToString();
            DateTime startDate = DateUtils.Now;
            DateTime endDate = DateUtils.Now.AddDays(1);

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //handle delivery
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, DELIVERY_NAME);

            if (deliveryPage.CheckTotalNumber() == 0)
            {
                var createDeliveryPage = deliveryPage.DeliveryCreatePage();
                createDeliveryPage.FillFields_CreateDeliveryModalPage(DELIVERY_NAME, customer, site, true);
                createDeliveryPage.Create();
            }

            //create price list if it does not exist today
            // Act
            var priceListPage = homePage.GoToCustomers_PriceListPage();
            priceListPage.ResetFilter();
            priceListPage.Filter(PriceListPage.FilterType.Site, site);
            priceListPage.Filter(PriceListPage.FilterType.Search, priceListName);

            //Create pricing if it does not exist
            try
            {
                if (priceListPage.CheckTotalNumber() == 0)
                {
                    var pricingModal = priceListPage.CreateNewPricing();
                    var priceListDetails = pricingModal.FillField_CreateNewPricing(priceListName, site, startDate, endDate, false);

                    priceListDetails.AddItem(itemNameForForeignTest, true);
                    priceListDetails.AddItem(recipeNameForForeignTest, false);
                    priceListDetails.AddCommercialNameByRow(COMMERCIAL_NAME_FOR_CO_TEST, 1);

                    priceListDetails.AddNewPeriod(startDate, endDate);
                    if (priceListDetails.ErrorMessageNewPeriod(PRICE_PERIOD_ERROR_MESSAGE))
                    {
                        priceListDetails.ClosePeriodPopup();
                    }
                    priceListDetails.BackToList();
                }
                else
                {
                    //handle PriceList Item - CommercialName
                    var priceDetailPage = priceListPage.ClickOnFirstPricing();
                    priceDetailPage.AddCommercialNameByRow(COMMERCIAL_NAME_FOR_CO_TEST, 1);
                }
            }
            catch (Exception)
            {
                DeletePriceList(homePage, site);
                throw;
            }

            //handle customerOrder link with PriceList
            var customerPage = homePage.GoToCustomers_CustomerPage();
            customerPage.Filter(PageObjects.Customers.Customer.CustomerPage.FilterType.Search, customer);
            var customerDetailPage = customerPage.SelectFirstCustomer();
            var custpricePage = customerDetailPage.ClickPriceListTab();
            if (!custpricePage.CheckSiteExists(site))
            {
                custpricePage.AddPrice(site, priceListName);
            }

            //create customerOrder
            var prodCOPage = homePage.GoToProduction_CustomerOrderPage();
            var createCOPage = prodCOPage.CustomerOrderCreatePage();
            createCOPage.FillField_CreatNewCustomerOrderWOUTflight(site, customer, DELIVERY_NAME);
            var customerOrderDetail = createCOPage.Submit();
            //Setting customer to foreign
            var geninfoTab = customerOrderDetail.ClickOnGeneralInformationTab();
            if (geninfoTab.GetLocalForeignValue() == "true")
            {
                geninfoTab.UpdateLocalForeign();
            }

            try
            {
                Assert.IsTrue(customerOrderDetail.AddNewItem(COMMERCIAL_NAME_FOR_CO_TEST + "-" + new Random().Next(), "5"));
            }
            finally
            {
                prodCOPage = customerOrderDetail.BackToList();
                prodCOPage.DeleteCustomerOrder();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_PRLIST_Local()
        {
            string site = TestContext.Properties["Site"].ToString();

            Random rnd = new Random();
            string priceName = priceNameToday + rnd.Next(1, 10000).ToString();
            DateTime startDate = DateUtils.Now;
            DateTime endDate = DateUtils.Now.AddDays(10);
            string itemName = TestContext.Properties["ItemNamePriceList"].ToString();
            bool priceListIsDefault = false;
            bool avoidAlreadyExistPriceList = true;

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Get Name from General Informations of the item
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            var Item = itemPage.ClickOnFirstItem();
            var ItemName = Item.GetItemName();

            // Get Name from General Informations of the recipe
            var recipePage = homePage.GoToMenus_Recipes();
            recipePage.Filter(RecipesPage.FilterType.SearchRecipe, recipeNameForForeignTest);
            var recipe = recipePage.SelectFirstRecipe();
            var recipeName = recipe.GetRecipeName();


            var priceListPage = homePage.GoToCustomers_PriceListPage();
            try
            {
                priceListPage.Filter(PriceListPage.FilterType.Site, site);
                var countPriceListPage = priceListPage.CheckTotalNumber();
                if (countPriceListPage == 0)
                {
                    priceListIsDefault = true;
                }
                var pricingCreateModalpage = priceListPage.CreateNewPricing();
                var priceListDetails = pricingCreateModalpage.FillField_CreateNewPricing(priceName, site, startDate, endDate, priceListIsDefault, avoidAlreadyExistPriceList);
                priceListDetails.AddItem(itemName, false);
                WebDriver.Navigate().Refresh();
                var priceItemName = priceListDetails.GetItemName();
                priceListDetails.DeleteItem();

                priceListDetails.AddItem(recipeNameForForeignTest, false);
                WebDriver.Navigate().Refresh();
                var priceRecipeName = priceListDetails.GetItemName();

                priceListDetails.AddItem(itemName, false);
                WebDriver.Navigate().Refresh();
                Assert.IsTrue(priceItemName.Contains(itemName), "le itemName ne correspond pas a son Name");
                Assert.IsTrue(priceRecipeName.Contains(recipeNameForForeignTest), "le recipeName ne correspond pas a son Name");
                var isLocal = priceListPage.isPriceListLocal();
                Assert.IsTrue(priceListPage.isPriceListLocal(), "la liste des prices doit etre en Local");
            }
            finally
            {
                priceListPage = homePage.GoToCustomers_PriceListPage();
                priceListPage.ResetFilter();
                priceListPage.Filter(PriceListPage.FilterType.Search, priceName);
                priceListPage.Filter(PriceListPage.FilterType.Site, site);
                if (priceListPage.CheckTotalNumber() > 0)
                {
                    priceListPage.DeletePricing(priceName);
                }
            }

        }
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_PRLIST_item_commercial_name()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string itemName = TestContext.Properties["ItemNamePriceList"].ToString();

            Random rnd = new Random();
            string priceName = priceNameToday + rnd.Next(1, 10000).ToString();
            DateTime startDate = DateUtils.Now;
            DateTime endDate = DateUtils.Now.AddDays(1);
            //Prepare recipe
            string recipeName = TestContext.Properties["RecipePriceList"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var priceListPage = homePage.GoToCustomers_PriceListPage();
            priceListPage.ResetFilter();
            priceListPage.Filter(PriceListPage.FilterType.Site, site);

            if (priceListPage.CheckTotalNumber() > 0)
            {
                priceListPage.DeleteAllPriceList();
            }

            try
            {
                var pricingCreateModalpage = priceListPage.CreateNewPricing();
                var priceListDetails = pricingCreateModalpage.FillField_CreateNewPricing(priceName, site, startDate, endDate);
                priceListDetails.AddItem(itemName, true);
                priceListDetails.AddItem(recipeName, false);
                priceListDetails.AddCommercialNameByRow(COMMERCIAL_NAME_FOR_CO_TEST, 1);
                Thread.Sleep(2000);
                WebDriver.Navigate().Refresh();
                Assert.IsTrue(priceListDetails.GetCommercialNameByRowNumber(1) == COMMERCIAL_NAME_FOR_CO_TEST, "Le Commercial Name " + COMMERCIAL_NAME_FOR_CO_TEST + "n'a pas été ajouté .");
            }
            finally
            {
                DeletePriceList(homePage, site);
                WebDriver.Navigate().Refresh();
            }

        }

        // ________________________________________ Utilitaire _____________________________________________

        public void DeletePriceList(HomePage homePage, string site)
        {
            var priceListPage = homePage.GoToCustomers_PriceListPage();
            priceListPage.ResetFilter();
            priceListPage.Filter(PriceListPage.FilterType.Site, site);

            if (priceListPage.CheckTotalNumber() > 0)
            {
                priceListPage.DeleteAllPriceList();
            }
        }



    }
}
