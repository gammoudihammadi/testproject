using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Datasheet;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Recipes;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Item;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.Inventory;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.OutputForm;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.ReceiptNotes;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using UglyToad.PdfPig;
using static Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.ItemGeneralInformationPage;
using Page = UglyToad.PdfPig.Content.Page;
using System.Threading;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Logs;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Sites;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using System.Security.Policy;
using System.Web.UI;
using System.Reflection;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
namespace Newrest.Winrest.FunctionalTests.Purchasing
{
    [TestClass]
    public class ItemTests : TestBase
    {
        private const int _timeout = 600000;
        public const string SHEET1 = "Sheet 1";
        private readonly string itemNameToday = "Item-" + DateUtils.Now.ToString("dd/MM/yyyy");
        private readonly string allergenNameToday = "Allergen-" + DateUtils.Now.ToString("dd/MM/yyyy");
        private readonly string NewallergenName = "NewAllergen-" + DateUtils.Now.ToString("dd/MM/yyyy");


        /// <summary>
        /// 
        /// Mise en place du paramétrage pour la configuration Winrest 4.0 
        /// 
        /// </summary>
        /// 
        [TestMethod]
        [Priority(0)]
        [Timeout(_timeout)]
        public void PU_ITEM_SetConfigWinrest4_0()
        {
            //Prepare
            var keyword = TestContext.Properties["Item_Keyword"].ToString();
            var site = TestContext.Properties["Site"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            ClearCache();

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
            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();
            var receiptNotesItem = receiptNotesPage.SelectFirstReceiptNoteItem();
            Assert.IsTrue(receiptNotesItem.IsGroupDisplayActive(), "Le paramètre 'NewGroupDisplay' n'est pas activé.");
        }

        //Créer un nouvel item
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Create_New_Item()
        {
            //Prepare
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
            string itemName = itemNameToday + "-" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
            itemPage = itemGeneralInformationPage.BackToList();

            // Assert 
            itemPage.Filter(ItemPage.FilterType.Search, itemName.ToString());
            string firstItemName = itemPage.GetFirstItemName();
            Assert.AreEqual(itemName, firstItemName, "Le nouvel item n'a pas été créé et ajouté à la liste.");
        }

        //Créer un nouvel item existant
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Create_New_Item_Already_Existing()
        {
            //Prepare
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string itemName = itemNameToday + "-" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();

            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
            itemPage = itemGeneralInformationPage.BackToList();

            itemCreateModalPage = itemPage.ItemCreatePage();
            itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);

            bool nameAlreadyExistMsgError = itemCreateModalPage.ErrorMessageNameAlreadyExists();
            // Assert 
            Assert.IsTrue(nameAlreadyExistMsgError, "Aucune erreur n'est survenue alors que l'item créé existe déjà.");
        }

        //Créer un nouvel item déjà sans remplir les champs obligatoires
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Create_New_Item_No_Fill_Fields()
        {
            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            var itemCreateModalPage = itemPage.ItemCreatePage();
            itemCreateModalPage.Save();

            bool nameRequiredMsgError = itemCreateModalPage.ErrorMessageNameRequired();
            // Assert 
            Assert.IsTrue(nameRequiredMsgError, "Aucune erreur n''st survenue alors que les champs obligatoires n'ont pas été renseignés.");
        }

        //Créer un packaging pour un nouvel item
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Create_Item_New_Packaging()
        {
            //Prepare
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
            string packPrice = 100.ToString();
            string itemName = itemNameToday + "-" + new Random().Next().ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();
            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();
            //Act
            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();

            ItemCreateModalPage itemCreateModalPage = itemPage.ItemCreatePage();
            ItemGeneralInformationPage itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            ItemCreateNewPackagingModalPage itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier, qty);

            // Assert 
            Assert.IsTrue(itemGeneralInformationPage.IsPackagingVisible(), "Le packaging n'a pas été créé.");
            Assert.IsTrue(itemGeneralInformationPage.VerifyPackagingSite(site), "Le site du packaging crée n'est pas le bon.");
            Assert.IsTrue(itemGeneralInformationPage.IsMainChecked(supplier), "Le site du packaging crée n'est pas le bon.");
            Assert.IsTrue(itemGeneralInformationPage.FindPackagingName(packagingName), "Le Packaging du packaging crée n'est pas le bon.");
            Assert.AreEqual(Convert.ToDouble(storageQty).ToString(), itemGeneralInformationPage.GetFirstPackagingStorageQty(decimalSeparatorValue).ToString(), "Le Packaging du packaging crée n'est pas le bon.");
            Assert.AreEqual(storageUnit, itemGeneralInformationPage.GetFirstPackagingStorUnit(), "La Stor. Unit du packaging crée n'est pas la bonne.");
            Assert.AreEqual(qty, itemGeneralInformationPage.GetFirstPackagingProdQty(), "La ProdQty du packaging crée n'est pas la bonne.");
            Assert.AreEqual(qty, itemGeneralInformationPage.GetFirstPackagingUnitPrice(), "La Unit. Price du packaging crée n'est pas la bonne.");
            Assert.AreEqual(packPrice, itemGeneralInformationPage.GetFirstPackagingPackPrice(), "La Pack. Price du packaging crée n'est pas la bonne.");
            Assert.AreEqual(qty + " " + storageUnit, itemGeneralInformationPage.GetFirstPackagingLimitQty(), "La LimitQty du packaging crée n'est pas la bonne.");
            Assert.AreEqual("97 %", itemGeneralInformationPage.GetFirstPackagingYield(), "La Yield du packaging crée n'est pas la bonne.");
        }

        //Créer un packaging multisite pour un nouvel
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Create_New_Packaging_MultiSites()
        {
            //Prepare
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site1 = TestContext.Properties["SiteACE"].ToString();
            string site2 = TestContext.Properties["SiteBCN"].ToString();
            string site3 = TestContext.Properties["SiteMAD"].ToString();
            string[] sites = { site1, site2, site3 };
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();

            string itemName = itemNameToday + "-" + new Random().Next().ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();
            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();

            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackagingMultiSites(sites, packagingName, storageQty, storageUnit, qty, supplier);

            // Assert

            foreach (var site in sites)
            {
                Assert.IsTrue(itemGeneralInformationPage.VerifyPackagingSite(site), "Le packaging multisite n'a pas été créé.");
            }


        }

        //Créer un packaging pour un nouvel item sans remplir les champs obligatoires
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Create_Item_New_Packaging_No_Fill_Fields()
        {
            //Prepare
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string itemName = itemNameToday + "-" + new Random().Next().ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();

            ItemCreateModalPage itemCreateModalPage = itemPage.ItemCreatePage();
            ItemGeneralInformationPage itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
            ItemCreateNewPackagingModalPage itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemCreatePackagingPage.NoFillFieldPackaging();

            //Assert
            Assert.IsTrue(itemCreatePackagingPage.ErrorMessageSiteRequired(), "Pas de message d'erreur alors que le site n'est pas renseigné.");
            Assert.IsTrue(itemCreatePackagingPage.ErrorMessagePackagingRequired(), "Pas de message d'erreur alors que le packaging n'est pas renseigné.");
            Assert.IsTrue(itemCreatePackagingPage.ErrorMessageQuantityRequired(), "Pas de message d'erreur alors que la quantité n'est pas renseignée.");
            Assert.IsTrue(itemCreatePackagingPage.ErrorMessageStorageUnitRequired(), "Pas de message d'erreur alors que la storage unit n'est pas renseignée.");
            Assert.IsTrue(itemCreatePackagingPage.ErrorMessageSupplierRequired(), "Pas de message d'erreur alors que le supplier n'est pas renseigné.");
        }

        //Modifier le nom d'un item
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_GeneralInfo_Update_Name()
        {
            //Prepare
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = "10";
            string qty = "10";

            string itemName = itemNameToday + "-" + new Random().Next().ToString();
            string newName = itemName + "_changed";

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            try
            {
                ItemCreateModalPage itemCreateModalPage = itemPage.ItemCreatePage();
                ItemGeneralInformationPage itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                itemGeneralInformationPage.SetItemName(newName);
                itemPage = itemGeneralInformationPage.BackToList();

                //Assert
                itemPage.Filter(ItemPage.FilterType.Search, newName);
                itemGeneralInformationPage = itemPage.ClickFirstItem();
                string editedName = itemGeneralInformationPage.GetItemName();
                Assert.AreEqual(newName, editedName, "Le nom de l'item n'a pas été modifié.");
            }
            finally
            {

                homePage.GoToPurchasing_ItemPage();
                itemPage.ResetFilter();
                //Desactivate packaging
                itemPage.Filter(ItemPage.FilterType.Search, newName);
                ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
                generalInfo.DeactivatePrice(site, supplier, packagingName);
                //Delete Item
                generalInfo.BackToList();
                MassiveDeleteModal massiveDelete = itemPage.MenuMassiveDelete();
                massiveDelete.Filter(MassiveDeleteModal.FilterType.SearchByName, newName);
                massiveDelete.Filter(MassiveDeleteModal.FilterType.ShowInactiveSuppliers, true);
                massiveDelete.Filter(MassiveDeleteModal.FilterType.SupplierMultiple, supplier);
                massiveDelete.ClickSearch();
                massiveDelete.ClickSelectAllButton();
                massiveDelete.Delete();
            }

        }

        //Modifier le groupe d'un item
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_GeneralInfo_Update_Group()
        {
            //Prepare
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = "10";
            string qty = "10";
            string newGroupName = "NORWEGIAN";
            string itemName = itemNameToday + "-" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            try
            {
                ItemCreateModalPage itemCreateModalPage = itemPage.ItemCreatePage();
                ItemGeneralInformationPage itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                itemGeneralInformationPage.SetItemGroup(newGroupName);
                itemPage = itemGeneralInformationPage.BackToList();
                //Assert
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                itemGeneralInformationPage = itemPage.ClickFirstItem();
                string editedGroupName = itemGeneralInformationPage.GetGroupName();
                Assert.AreEqual(newGroupName, editedGroupName, "Le groupe de l'item n'a pas été modifié.");
            }
            finally
            {
                homePage.GoToPurchasing_ItemPage();
                itemPage.ResetFilter();
                //Desactivate packaging
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
                generalInfo.DeactivatePrice(site, supplier, packagingName);
                //Delete Item
                generalInfo.BackToList();
                MassiveDeleteModal massiveDelete = itemPage.MenuMassiveDelete();
                massiveDelete.Filter(MassiveDeleteModal.FilterType.SearchByName, itemName);
                massiveDelete.Filter(MassiveDeleteModal.FilterType.ShowInactiveSuppliers, true);
                massiveDelete.Filter(MassiveDeleteModal.FilterType.SupplierMultiple, supplier);
                massiveDelete.ClickSearch();
                massiveDelete.ClickSelectAllButton();
                massiveDelete.Delete();
            }

        }

        //Modifier le nom commercial 1 d'un item
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_GeneralInfo_Update_CommercialName1()
        {
            //Prepare
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = "10";
            string qty = "10";
            string itemName = itemNameToday + "-" + new Random().Next().ToString();
            string newCommercialName1 = "Commercial_Name_1_changed";

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            try
            {
                ItemCreateModalPage itemCreateModalPage = itemPage.ItemCreatePage();
                ItemGeneralInformationPage itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                itemGeneralInformationPage.SetItemCommercialName1(newCommercialName1);
                itemPage = itemGeneralInformationPage.BackToList();
                //Assert
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                itemGeneralInformationPage = itemPage.ClickFirstItem();
                string editedCommercialName1 = itemGeneralInformationPage.GetItemCommercialName1();
                Assert.AreEqual(newCommercialName1, editedCommercialName1, "Le nom commercial 1 de l'item  n'a pas été modifié.");
            }
            finally
            {
                homePage.GoToPurchasing_ItemPage();
                itemPage.ResetFilter();
                //Desactivate packaging
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
                generalInfo.DeactivatePrice(site, supplier, packagingName);
                //Delete Item
                generalInfo.BackToList();
                MassiveDeleteModal massiveDelete = itemPage.MenuMassiveDelete();
                massiveDelete.Filter(MassiveDeleteModal.FilterType.SearchByName, itemName);
                massiveDelete.Filter(MassiveDeleteModal.FilterType.ShowInactiveSuppliers, true);
                massiveDelete.Filter(MassiveDeleteModal.FilterType.SupplierMultiple, supplier);
                massiveDelete.ClickSearch();
                massiveDelete.ClickSelectAllButton();
                massiveDelete.Delete();
            }

        }


        //Modifier le nom commercial 2 d'un item
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_GeneralInfo_Update_CommercialName2()
        {
            //Prepare
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = "10";
            string qty = "10";
            string itemName = itemNameToday + "-" + new Random().Next().ToString();
            string CommercialName2 = "Commercial_Name_2";
            string newCommercialName2 = CommercialName2 + "_changed";

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            try
            {
                ItemCreateModalPage itemCreateModalPage = itemPage.ItemCreatePage();
                ItemGeneralInformationPage itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit, null, null, null, CommercialName2);

                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                itemGeneralInformationPage.SetItemCommercialName2(newCommercialName2);
                itemPage = itemGeneralInformationPage.BackToList();

                //Assert
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                itemGeneralInformationPage = itemPage.ClickFirstItem();
                string editedCommercialName2 = itemGeneralInformationPage.GetItemCommercialName2();
                Assert.AreEqual(newCommercialName2, editedCommercialName2, "Le nom commercial 2 de l'item n'a pas été modifié.");
            }
            finally
            {
                homePage.GoToPurchasing_ItemPage();
                itemPage.ResetFilter();
                //Desactivate packaging
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
                generalInfo.DeactivatePrice(site, supplier, packagingName);
                //Delete Item
                generalInfo.BackToList();
                MassiveDeleteModal massiveDelete = itemPage.MenuMassiveDelete();
                massiveDelete.Filter(MassiveDeleteModal.FilterType.SearchByName, itemName);
                massiveDelete.Filter(MassiveDeleteModal.FilterType.ShowInactiveSuppliers, true);
                massiveDelete.Filter(MassiveDeleteModal.FilterType.SupplierMultiple, supplier);
                massiveDelete.ClickSearch();
                massiveDelete.ClickSelectAllButton();
                massiveDelete.Delete();
            }

        }

        //Modifier le workshop d'un item
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_ITEM_GeneralInfo_Update_Workshop()
        {
            //Prepare
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = "10";
            string qty = "10";
            string itemName = itemNameToday + "-" + new Random().Next().ToString();
            string newWorkshopName = "Corte";

            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            try
            {
                ItemCreateModalPage itemCreateModalPage = itemPage.ItemCreatePage();
                ItemGeneralInformationPage itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                itemGeneralInformationPage.SetItemWorkshop(newWorkshopName);
                itemPage = itemGeneralInformationPage.BackToList();
                //Assert
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                itemGeneralInformationPage = itemPage.ClickFirstItem();
                string editedworkshopName = itemGeneralInformationPage.GetWorkshop();
                Assert.AreEqual(newWorkshopName, editedworkshopName, "Le Workshop de l'item n'a pas été modifié.");
            }
            finally
            {
                homePage.GoToPurchasing_ItemPage();
                itemPage.ResetFilter();
                //Desactivate packaging
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
                generalInfo.DeactivatePrice(site, supplier, packagingName);
                //Delete Item
                generalInfo.BackToList();
                MassiveDeleteModal massiveDelete = itemPage.MenuMassiveDelete();
                massiveDelete.Filter(MassiveDeleteModal.FilterType.SearchByName, itemName);
                massiveDelete.Filter(MassiveDeleteModal.FilterType.ShowInactiveSuppliers, true);
                massiveDelete.Filter(MassiveDeleteModal.FilterType.SupplierMultiple, supplier);
                massiveDelete.ClickSearch();
                massiveDelete.ClickSelectAllButton();
                massiveDelete.Delete();
            }

        }

        //Modifier la référence d'un item
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_ITEM_GeneralInfo_Update_Reference()
        {
            //Prepare
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = "10";
            string qty = "10";
            string itemName = itemNameToday + "-" + new Random().Next().ToString();
            string reference = "REF";
            string newReference = reference + "_changed";

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            try
            {
                ItemCreateModalPage itemCreateModalPage = itemPage.ItemCreatePage();
                ItemGeneralInformationPage itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit, null, null, null, null, reference);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                itemGeneralInformationPage.SetItemReference(newReference);
                itemPage = itemGeneralInformationPage.BackToList();

                //Assert
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                itemGeneralInformationPage = itemPage.ClickFirstItem();
                string editedReference = itemGeneralInformationPage.GetItemReference();
                Assert.AreEqual(newReference, editedReference, "La référence de l'item n'a pas été modifié.");
            }
            finally
            {
                homePage.GoToPurchasing_ItemPage();
                itemPage.ResetFilter();
                //Desactivate packaging
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
                generalInfo.DeactivatePrice(site, supplier, packagingName);
                //Delete Item
                generalInfo.BackToList();
                MassiveDeleteModal massiveDelete = itemPage.MenuMassiveDelete();
                massiveDelete.Filter(MassiveDeleteModal.FilterType.SearchByName, itemName);
                massiveDelete.Filter(MassiveDeleteModal.FilterType.ShowInactiveSuppliers, true);
                massiveDelete.Filter(MassiveDeleteModal.FilterType.SupplierMultiple, supplier);
                massiveDelete.ClickSearch();
                massiveDelete.ClickSelectAllButton();
                massiveDelete.Delete();
            }

        }

        //Modifier le VAT d'un item
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_ITEM_GeneralInfo_Update_VAT()
        {
            //Prepare
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = "10";
            string qty = "10";
            string itemName = itemNameToday + "-" + new Random().Next().ToString();
            string newVAT = "2-General";

            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            try
            {
                ItemCreateModalPage itemCreateModalPage = itemPage.ItemCreatePage();
                ItemGeneralInformationPage itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                itemGeneralInformationPage.SetItemVAT(newVAT);
                itemPage = itemGeneralInformationPage.BackToList();
                //Assert
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                itemGeneralInformationPage = itemPage.ClickFirstItem();
                string editedVAT = itemGeneralInformationPage.GetVatName();
                Assert.AreEqual(newVAT, editedVAT, "Le VAT de l'item n'a pas été modifié.");
            }
            finally
            {
                homePage.GoToPurchasing_ItemPage();
                itemPage.ResetFilter();
                //Desactivate packaging
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
                generalInfo.DeactivatePrice(site, supplier, packagingName);
                //Delete Item
                generalInfo.BackToList();
                MassiveDeleteModal massiveDelete = itemPage.MenuMassiveDelete();
                massiveDelete.Filter(MassiveDeleteModal.FilterType.SearchByName, itemName);
                massiveDelete.Filter(MassiveDeleteModal.FilterType.ShowInactiveSuppliers, true);
                massiveDelete.Filter(MassiveDeleteModal.FilterType.SupplierMultiple, supplier);
                massiveDelete.ClickSearch();
                massiveDelete.ClickSelectAllButton();
                massiveDelete.Delete();
            }

        }

        //Modifier L'unité de production d'un item
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_ITEM_GeneralInfo_Update_ProdUnit()
        {
            //Prepare
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = "10";
            string qty = "10";
            string itemName = itemNameToday + "-" + new Random().Next().ToString();
            string newProdUnit = "ML";

            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            try
            {
                ItemCreateModalPage itemCreateModalPage = itemPage.ItemCreatePage();
                ItemGeneralInformationPage itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                itemGeneralInformationPage.SetItemProdUnit(newProdUnit);
                itemPage = itemGeneralInformationPage.BackToList();
                //Assert
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                itemGeneralInformationPage = itemPage.ClickFirstItem();
                string editedProdUnit = itemGeneralInformationPage.GetProdUnit();
                Assert.AreEqual(newProdUnit, editedProdUnit, "L'unité de production de l'item n'a pas été modifié.");
            }
            finally
            {
                homePage.GoToPurchasing_ItemPage();
                itemPage.ResetFilter();
                //Desactivate packaging
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
                generalInfo.DeactivatePrice(site, supplier, packagingName);
                //Delete Item
                generalInfo.BackToList();
                MassiveDeleteModal massiveDelete = itemPage.MenuMassiveDelete();
                massiveDelete.Filter(MassiveDeleteModal.FilterType.SearchByName, itemName);
                massiveDelete.Filter(MassiveDeleteModal.FilterType.ShowInactiveSuppliers, true);
                massiveDelete.Filter(MassiveDeleteModal.FilterType.SupplierMultiple, supplier);
                massiveDelete.ClickSearch();
                massiveDelete.ClickSelectAllButton();
                massiveDelete.Delete();
            }

        }

        //Modifier le poids en grammes d'un item
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_ITEM_GeneralInfo_Update_WeightInGram()
        {
            //Prepare
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = "10";
            string qty = "10";
            string itemName = itemNameToday + "-" + new Random().Next().ToString();
            string newWeightInGram = "3";

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            try
            {
                ItemCreateModalPage itemCreateModalPage = itemPage.ItemCreatePage();
                ItemGeneralInformationPage itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                itemGeneralInformationPage.SetItemWeightInGram(newWeightInGram);
                itemPage = itemGeneralInformationPage.BackToList();

                //Assert
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                itemGeneralInformationPage = itemPage.ClickFirstItem();
                string editedWeightInGramt = itemGeneralInformationPage.GetItemWeightInGram();
                Assert.AreEqual(newWeightInGram, editedWeightInGramt, "Le poids en grammes de l'item n'a pas été modifié.");
            }
            finally
            {
                homePage.GoToPurchasing_ItemPage();
                itemPage.ResetFilter();
                //Desactivate packaging
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
                generalInfo.DeactivatePrice(site, supplier, packagingName);
                //Delete Item
                generalInfo.BackToList();
                MassiveDeleteModal massiveDelete = itemPage.MenuMassiveDelete();
                massiveDelete.Filter(MassiveDeleteModal.FilterType.SearchByName, itemName);
                massiveDelete.Filter(MassiveDeleteModal.FilterType.ShowInactiveSuppliers, true);
                massiveDelete.Filter(MassiveDeleteModal.FilterType.SupplierMultiple, supplier);
                massiveDelete.ClickSearch();
                massiveDelete.ClickSelectAllButton();
                massiveDelete.Delete();
            }

        }

        //Modifier le Checkbox IsSanitization d'un item
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_GeneralInfo_Update_IsSanitization()
        {
            //Prepare
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = "10";
            string qty = "10";
            string itemName = itemNameToday + "-" + new Random().Next().ToString();
            bool isSanitization = false;
            bool newIsSanitization = true;

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            try
            {
                ItemCreateModalPage itemCreateModalPage = itemPage.ItemCreatePage();
                ItemGeneralInformationPage itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit, null, null, null, null, null, isSanitization);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                itemGeneralInformationPage.SetCheckoxIsSanitization(newIsSanitization);
                itemPage = itemGeneralInformationPage.BackToList();

                //Assert
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                itemGeneralInformationPage = itemPage.ClickFirstItem();
                bool editedCheckboxSanitization = itemGeneralInformationPage.GetCheckboxIsSantization();
                Assert.AreEqual(newIsSanitization, editedCheckboxSanitization, "Le Checkbox IsSanitization de l'item n'a pas été modifié.");
            }
            finally
            {

                homePage.GoToPurchasing_ItemPage();
                itemPage.ResetFilter();
                //Desactivate packaging
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
                generalInfo.DeactivatePrice(site, supplier, packagingName);
                //Delete Item
                generalInfo.BackToList();
                MassiveDeleteModal massiveDelete = itemPage.MenuMassiveDelete();
                massiveDelete.Filter(MassiveDeleteModal.FilterType.SearchByName, itemName);
                massiveDelete.Filter(MassiveDeleteModal.FilterType.ShowInactiveSuppliers, true);
                massiveDelete.Filter(MassiveDeleteModal.FilterType.SupplierMultiple, supplier);
                massiveDelete.ClickSearch();
                massiveDelete.ClickSelectAllButton();
                massiveDelete.Delete();
            }

        }

        //Modifier le Checkbox IsThawing d'un item
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_GeneralInfo_Update_IsThawing()
        {
            //Prepare
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = "10";
            string qty = "10";
            string itemName = itemNameToday + "-" + new Random().Next().ToString();
            bool IsThawing = false;
            bool newIsThawing = true;

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            try
            {
                ItemCreateModalPage itemCreateModalPage = itemPage.ItemCreatePage();
                ItemGeneralInformationPage itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit, null, null, null, null, null, IsThawing);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                itemGeneralInformationPage.SetCheckoxIsThawing(newIsThawing);
                itemPage = itemGeneralInformationPage.BackToList();

                //Assert
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                itemGeneralInformationPage = itemPage.ClickFirstItem();
                bool editedIsThawing = itemGeneralInformationPage.GetCheckboxIsThawing();
                Assert.AreEqual(newIsThawing, editedIsThawing, "Le Checkbox IsThawing de l'item n'a pas été modifié.");
            }
            finally
            {

                homePage.GoToPurchasing_ItemPage();
                itemPage.ResetFilter();
                //Desactivate packaging
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
                generalInfo.DeactivatePrice(site, supplier, packagingName);
                //Delete Item
                generalInfo.BackToList();
                MassiveDeleteModal massiveDelete = itemPage.MenuMassiveDelete();
                massiveDelete.Filter(MassiveDeleteModal.FilterType.SearchByName, itemName);
                massiveDelete.Filter(MassiveDeleteModal.FilterType.ShowInactiveSuppliers, true);
                massiveDelete.Filter(MassiveDeleteModal.FilterType.SupplierMultiple, supplier);
                massiveDelete.ClickSearch();
                massiveDelete.ClickSelectAllButton();
                massiveDelete.Delete();
            }

        }


        //Modifier le Checkbox NoHACCPRecord d'un item
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_GeneralInfo_Update_NoHACCPRecord()
        {
            //Prepare
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = "10";
            string qty = "10";
            string itemName = itemNameToday + "-" + new Random().Next().ToString();
            bool NoHACCPRecord = false;
            bool newNoHACCPRecord = true;

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            try
            {
                ItemCreateModalPage itemCreateModalPage = itemPage.ItemCreatePage();
                ItemGeneralInformationPage itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit, null, null, null, null, null, NoHACCPRecord);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                itemGeneralInformationPage.SetCheckoxNoHACCPRecord(newNoHACCPRecord);
                itemPage = itemGeneralInformationPage.BackToList();

                //Assert
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                itemGeneralInformationPage = itemPage.ClickFirstItem();
                bool editedNoHACCPRecord = itemGeneralInformationPage.GetCheckboxNoHACCPRecord();
                Assert.AreEqual(newNoHACCPRecord, editedNoHACCPRecord, "Le Checkbox NoHACCPRecord de l'item n'a pas été modifié.");
            }
            finally
            {

                homePage.GoToPurchasing_ItemPage();
                itemPage.ResetFilter();
                //Desactivate packaging
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
                generalInfo.DeactivatePrice(site, supplier, packagingName);
                //Delete Item
                generalInfo.BackToList();
                MassiveDeleteModal massiveDelete = itemPage.MenuMassiveDelete();
                massiveDelete.Filter(MassiveDeleteModal.FilterType.SearchByName, itemName);
                massiveDelete.Filter(MassiveDeleteModal.FilterType.ShowInactiveSuppliers, true);
                massiveDelete.Filter(MassiveDeleteModal.FilterType.SupplierMultiple, supplier);
                massiveDelete.ClickSearch();
                massiveDelete.ClickSelectAllButton();
                massiveDelete.Delete();
            }

        }

        //Modifier le Checkbox IsVegetable d'un item
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_GeneralInfo_Update_IsVegetable()
        {
            //Prepare
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = "10";
            string qty = "10";
            string itemName = itemNameToday + "-" + new Random().Next().ToString();
            bool IsVegetable = false;
            bool newIsVegetable = true;

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            try
            {
                ItemCreateModalPage itemCreateModalPage = itemPage.ItemCreatePage();
                ItemGeneralInformationPage itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit, null, null, null, null, null, IsVegetable);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                itemGeneralInformationPage.SetCheckoxIsVegetable(newIsVegetable);
                itemPage = itemGeneralInformationPage.BackToList();

                //Assert
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                itemGeneralInformationPage = itemPage.ClickFirstItem();
                bool editedIsVegetable = itemGeneralInformationPage.GetCheckboxIsVegetable();
                Assert.AreEqual(newIsVegetable, editedIsVegetable, "Le Checkbox IsVegetable de l'item n'a pas été modifié.");
            }
            finally
            {

                homePage.GoToPurchasing_ItemPage();
                itemPage.ResetFilter();
                //Desactivate packaging
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
                generalInfo.DeactivatePrice(site, supplier, packagingName);
                //Delete Item
                generalInfo.BackToList();
                MassiveDeleteModal massiveDelete = itemPage.MenuMassiveDelete();
                massiveDelete.Filter(MassiveDeleteModal.FilterType.SearchByName, itemName);
                massiveDelete.Filter(MassiveDeleteModal.FilterType.ShowInactiveSuppliers, true);
                massiveDelete.Filter(MassiveDeleteModal.FilterType.SupplierMultiple, supplier);
                massiveDelete.ClickSearch();
                massiveDelete.ClickSelectAllButton();
                massiveDelete.Delete();
            }

        }

        //Desactiver le prix d'un item
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Item_Deactivate_Price()
        {
            //Prepare
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
            string itemName = itemNameToday + "-" + new Random().Next().ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
            itemGeneralInformationPage.DeactivatePrice(site, supplier, packagingName);

            // Assert 
            Assert.IsFalse(itemGeneralInformationPage.IsPackagingVisible(), "Le packaging est toujours visible alors que le prix a été désactivé.");
        }

        //Vérifier les informations d'un packaging
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Item_Packaging_Check_Information()
        {
            //Prepare
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
            string itemName = "Item_Packaging_Check_Information";

            //Arrange
            var homePage = LogInAsAdmin();
            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformation = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformation.NewPackaging();
                itemGeneralInformation = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                itemGeneralInformation.BackToList();
                itemPage.Filter(ItemPage.FilterType.Search, itemName);

            }

            var itemGeneralInformationPage = itemPage.ClickFirstItem();
            // Assert 
            Assert.AreEqual(supplier, itemGeneralInformationPage.GetFirstPackagingSupplier(), "Le supplier n'est pas celui attendu.");
            Assert.AreEqual(packagingName, itemGeneralInformationPage.GetFirstPackagingInformation(), "Le ackaging n'est pas celui attendu.");
            Assert.AreEqual(Convert.ToDouble(storageQty).ToString(), itemGeneralInformationPage.GetFirstPackagingStorageQty(decimalSeparatorValue).ToString(), "La storage qty n'est pas celle attendue.");
        }

        //L'item est achetable
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Item_Packaging_Purchasable()
        {
            //Prepare
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
            string itemName = new Random().Next().ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
            itemGeneralInformationPage.SetPurchasable(false);
            itemPage = itemGeneralInformationPage.BackToList();

            // On recherche l'item créé dans la liste des items not purchasable
            itemPage.Filter(ItemPage.FilterType.ActiveNotPurchasable, true);
            itemPage.Filter(ItemPage.FilterType.Search, itemName.ToString());
            Assert.AreEqual(itemName.ToString(), itemPage.GetFirstItemName(), "L'item est toujours Purchasable malgré la désactivation de la fonctionnalité.");

            // On remodifie l'item
            itemPage.ClickOnFirstItem();
            itemGeneralInformationPage.SetPurchasable(true);
            itemPage = itemGeneralInformationPage.BackToList();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName.ToString());
            itemPage.Filter(ItemPage.FilterType.ActiveNotPurchasable, true);

            Assert.AreEqual(0, itemPage.CheckTotalNumber(), "L'item est toujours notPurchasable malgré la réactivation de la fonctionnalité.");
        }

        //Dupliquer les détails de l'item
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Item_Packaging_Duplicate()
        {
            //Prepare
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string siteDuplicate = TestContext.Properties["SiteLP"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string itemName = itemNameToday + "-" + new Random().Next().ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, 2.32.ToString(), storageUnit, 4.2.ToString(), supplier);
            itemGeneralInformationPage.DuplicateItem(siteDuplicate);

            // Assert 
            Assert.AreEqual(2, itemGeneralInformationPage.CountPackaging(), "Le packaging n'a pas été dupliqué.");
        }

        //Ajouter un code barre
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Item_Packaging_Bar_Code()
        {
            //Prepare
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
            string itemName = itemNameToday + "-" + new Random().Next().ToString();
            string codeBar = itemName + 1.ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
            itemGeneralInformationPage.AddBarCode(codeBar);

            // Assert
            var isCodeBarAdded = itemGeneralInformationPage.IsCodeBarAdded(codeBar);
            Assert.IsTrue(isCodeBarAdded, "Le code bar n'a pas été ajouté.");
        }

        //Ajouter un commentaire
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Item_Packaging_Add_Comment()
        {
            //Prepare
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
            string itemName = itemNameToday + "-" + new Random().Next().ToString();
            string comment = "test comment";

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
            var commentPage = itemGeneralInformationPage.ClickOnCommentItem();
            commentPage.FillComment(comment);
            itemGeneralInformationPage = commentPage.Save();

            // Assert 
            commentPage = itemGeneralInformationPage.ClickOnCommentItem();
            Assert.IsTrue(commentPage.IsCommented(comment), "Le commentaire n'a pas été ajouté.");
        }

        //Ajouter delete item
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Item_Packaging_Delete_Item()
        {
            //Prepare
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string delivery = TestContext.Properties["PlaceFrom"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();
            string itemName = itemNameToday + "-" + new Random().Next().ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
            var ReceiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            ReceiptNotesPage.ResetFilter();

            // Create
            var receiptNotesCreateModalpage = ReceiptNotesPage.ReceiptNotesCreatePage();
            receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, delivery));
            var receiptNotesItem = receiptNotesCreateModalpage.Submit();
            receiptNotesItem.Filter(ReceiptNotesItem.FilterItemType.SearchByName, itemName);
            receiptNotesItem.SelectFirstItem();
            receiptNotesItem.AddReceived(itemName, "2");
            var qualityChecks = receiptNotesItem.ClickOnChecksTab();
            qualityChecks.DeliveryAccepted();

            if (qualityChecks.CanClickOnSecurityChecks())
            {
                qualityChecks.CanClickOnSecurityChecks();
                qualityChecks.SetSecurityChecks("No");
                qualityChecks.SetQualityChecks();
                receiptNotesItem = qualityChecks.ClickOnReceiptNoteItemTab();
            }
            else
            {
                receiptNotesItem = qualityChecks.ClickOnReceiptNoteItemTab();
            }

            receiptNotesItem.Validate();

            itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.Filter(ItemPage.FilterType.Search, itemName.ToString());
            itemPage.ClickOnFirstItem();

            itemGeneralInformationPage.ClickOnDeleteItem();
            // Assert 
            Assert.IsTrue(itemGeneralInformationPage.ErrorMessageCannotBeDeleted(supplier, site), "Le packaging peut être supprimé alors qu'il est associé à un RN.");
        }

        //Rendre les packaging unpurchasable
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Item_SetUnpurchasableItem()
        {
            //Prepare
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
            string itemName = itemNameToday + "-" + new Random().Next().ToString();

            //Arrange

            HomePage homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
            Assert.IsTrue(itemGeneralInformationPage.IsPurchasable(site, supplier), "Le packaging n'est pas purchasable.");
            itemGeneralInformationPage.SetUnpurchasableItem(site, supplier, packagingName);
            WebDriver.Navigate().Refresh();

            // Assert 
            Assert.IsFalse(itemGeneralInformationPage.IsPurchasable(site, supplier), "La fonctionnalité 'SetUnpurchasable items' ne fonctionne pas.");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Item_FilterPackagingBySite()
        {
            //Prepare
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string site2 = TestContext.Properties["SiteBis"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();
            string itemName = itemNameToday + "-" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
            itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site2, packagingName, storageQty, storageUnit, qty, supplier);
            Assert.AreEqual(2, itemGeneralInformationPage.CountPackaging(), "Les packaging n'ont pas été créés.");
            itemGeneralInformationPage.SearchPackaging(site);

            // Assert
            var countPackaging = itemGeneralInformationPage.CountPackaging();
            var verifyPackagingSite = itemGeneralInformationPage.VerifyPackagingSite(site);
            Assert.AreEqual(1, countPackaging, "Les packaging n'ont pas été filtrés.");
            Assert.IsTrue(verifyPackagingSite, String.Format(MessageErreur.FILTRE_ERRONE, "Search packaging"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Item_FilterPackagingBySupplier()
        {
            //Prepare
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["Prodman_Needs_Item_Supplier2"].ToString();
            string supplier2 = TestContext.Properties["Prodman_Needs_Item_Supplier1"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();
            string itemName = itemNameToday + "-" + new Random().Next().ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();

            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
            itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier2);

            Assert.AreEqual(2, itemGeneralInformationPage.CountPackaging(), "Les packaging n'ont pas été créés.");

            itemGeneralInformationPage.SearchPackaging(supplier);

            // Assert 
            Assert.AreEqual(1, itemGeneralInformationPage.CountPackaging(), "Les packaging n'ont pas été filtrés.");
            Assert.IsTrue(itemGeneralInformationPage.VerifyPackagingSupplier(supplier), String.Format(MessageErreur.FILTRE_ERRONE, "Search packaging"));
        }

        //Voir tous les packagings d'un item
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Item_Packaging_Show_All()
        {
            //Prepare
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string site2 = TestContext.Properties["SiteToFlight"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();
            string itemName = itemNameToday + "-" + new Random().Next().ToString();

            //Arrange

            HomePage homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
            itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site2, packagingName, storageQty, storageUnit, qty, supplier);
            itemGeneralInformationPage.DeactivatePrice(site2, supplier, packagingName);
            int valueBeforeShowAll = itemGeneralInformationPage.CountPackaging();
            itemGeneralInformationPage.ShowAllPackaging();
            int valueShowAll = itemGeneralInformationPage.CountPackaging();

            // Assert 
            Assert.AreNotEqual(valueBeforeShowAll, valueShowAll, "le nombre de packaging visibles est identique après avoir activé l'option show all");
        }

        //Print item - new version
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_ITEM_Item_Packaging_Print_New_Version()
        {
            //Prepare
            bool newversionPrint = true;

            //Prepare
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

            string itemName = itemNameToday + "-" + new Random().Next().ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();

            itemPage.ClearDownloads();

            ItemGeneralInformationPage itemGeneralInformationPage;

            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
            }
            else
            {
                itemGeneralInformationPage = itemPage.ClickOnFirstItem();
            }

            var printPage = itemGeneralInformationPage.Print(newversionPrint, site);
            bool isGenerated = printPage.IsReportGenerated();
            printPage.Close();

            //Assert
            Assert.IsTrue(isGenerated, "Le fichier PDF n'a pas été généré.");
        }

        //Créer un nouvel item
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Items_Purchase_Fold_Unfold()
        {
            //Prepare
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

            string itemName = itemNameToday + "-" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();
            homePage.Navigate();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();

            if (itemPage.CheckTotalNumber() < 5)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                itemPage = itemGeneralInformationPage.BackToList();
            }

            //Assert - unfold
            itemPage.UnfoldAll();
            bool isUnfoldAll = itemPage.IsUnfoldAll();
            Assert.IsTrue(isUnfoldAll, "Le détail des items n'est pas affiché.");

            //Assert - fold
            itemPage.FoldAll();
            bool isFoldAll = itemPage.IsFoldAll();
            Assert.IsTrue(isFoldAll, "Le détail des items n'est pas caché.");
        }

        //Créer un nouvel item
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Items_Storage_FoldAll()
        {
            //Prepare            
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

            string itemName = itemNameToday + "-" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.ClickOnStorage();
            int totalNumber = itemPage.CheckTotalNumber();

            try
            {
                if (totalNumber < 5)
                {
                    var itemCreateModalPage = itemPage.ItemCreatePage();
                    var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
                    var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                    itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                    itemPage = itemGeneralInformationPage.BackToList();
                    itemPage.ClickOnStorage();
                }

                //unfold
                itemPage.UnfoldAll();

                //Assert - fold
                itemPage.FoldAll();
                bool isFoldAll = itemPage.IsFoldAll();
                Assert.IsTrue(isFoldAll, "Le détail des items n'est pas masqué.");
            }

            finally
            {
                if (totalNumber < 5)
                {
                    homePage.GoToPurchasing_ItemPage();
                    itemPage.ResetFilter();
                    //Desactivate packaging
                    itemPage.Filter(ItemPage.FilterType.Search, itemName);
                    ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
                    generalInfo.DeactivatePrice(site, supplier, packagingName);
                    //Delete Item
                    generalInfo.BackToList();
                    MassiveDeleteModal massiveDelete = itemPage.MenuMassiveDelete();
                    massiveDelete.Filter(MassiveDeleteModal.FilterType.SearchByName, itemName);
                    massiveDelete.Filter(MassiveDeleteModal.FilterType.ShowInactiveSuppliers, true);
                    massiveDelete.Filter(MassiveDeleteModal.FilterType.SupplierMultiple, supplier);
                    massiveDelete.ClickSearch();
                    massiveDelete.ClickSelectAllButton();
                    massiveDelete.Delete();
                }
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Items_Storage_UnfoldAll()
        {
            //Prepare            
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

            string itemName = itemNameToday + "-" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.ClickOnStorage();
            int totalNumber = itemPage.CheckTotalNumber();

            try
            {
                if (totalNumber < 5)
                {
                    var itemCreateModalPage = itemPage.ItemCreatePage();
                    var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
                    var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                    itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                    itemPage = itemGeneralInformationPage.BackToList();
                    itemPage.ClickOnStorage();
                }

                //unfold
                itemPage.UnfoldAll();
                bool isUnfoldAll = itemPage.IsUnfoldAll();

                //Assert
                Assert.IsTrue(isUnfoldAll, "Le détail des items n'est pas affiché.");
            }
            finally
            {
                if (totalNumber < 5)
                {
                    homePage.GoToPurchasing_ItemPage();
                    itemPage.ResetFilter();
                    //Desactivate packaging
                    itemPage.Filter(ItemPage.FilterType.Search, itemName);
                    ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
                    generalInfo.DeactivatePrice(site, supplier, packagingName);
                    //Delete Item
                    generalInfo.BackToList();
                    MassiveDeleteModal massiveDelete = itemPage.MenuMassiveDelete();
                    massiveDelete.Filter(MassiveDeleteModal.FilterType.SearchByName, itemName);
                    massiveDelete.Filter(MassiveDeleteModal.FilterType.ShowInactiveSuppliers, true);
                    massiveDelete.Filter(MassiveDeleteModal.FilterType.SupplierMultiple, supplier);
                    massiveDelete.ClickSearch();
                    massiveDelete.ClickSelectAllButton();
                    massiveDelete.Delete();
                }
            }

        }

        //Filtre nom d'item
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_FilterByName()
        {
            //Prepare
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string supplierRef = "coucou";
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();
            string itemName = itemNameToday + "-" + new Random().Next().ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();
            //Act
            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
            ItemCreateModalPage itemCreateModalPage = itemPage.ItemCreatePage();
            ItemGeneralInformationPage itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            ItemCreateNewPackagingModalPage itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier, null, supplierRef);
            itemPage = itemGeneralInformationPage.BackToList();
            itemPage.ResetFilter();
            // Assert 
            itemPage.Filter(ItemPage.FilterType.Search, itemName.ToString());
            Assert.AreEqual(itemName.ToString(), itemPage.GetFirstItemName(), String.Format(MessageErreur.FILTRE_ERRONE, "Search"));

            itemPage.Filter(ItemPage.FilterType.Supplier, supplier);
            Assert.AreEqual(itemName.ToString(), itemPage.GetFirstItemName(), String.Format(MessageErreur.FILTRE_ERRONE, "Search Supplier"));

            itemPage.Filter(ItemPage.FilterType.SupplierRef, supplierRef);
            Assert.AreEqual(itemName.ToString(), itemPage.GetFirstItemName(), String.Format(MessageErreur.FILTRE_ERRONE, "Search SupplierRef"));
        }

        //Filtre SortBy
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Filter_SortBy()
        {
            //Prepare
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
            string itemName = itemNameToday + "-" + new Random().Next().ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();

            if (itemPage.CheckTotalNumber() < 20)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                itemPage = itemGeneralInformationPage.BackToList();
            }

            if (!itemPage.isPageSizeEqualsTo100())
            {
                itemPage.PageSize("8");
                itemPage.PageSize("100");
            }

            itemPage.Filter(ItemPage.FilterType.SortBy, "Group");
            var isSortedByItemGroup = itemPage.IsSortedByItemGroup();

            itemPage.Filter(ItemPage.FilterType.SortBy, "Name");
            var isSortedByName = itemPage.IsSortedByName();

            //Assert
            Assert.IsTrue(isSortedByItemGroup, MessageErreur.FILTRE_ERRONE, "Sort by Item group");
            Assert.IsTrue(isSortedByName, MessageErreur.FILTRE_ERRONE, "Sort by Name");
        }

        //Filtre Inactive
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Filter_ShowInactive()
        {
            //Prepare
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
            string itemName = itemNameToday + "-" + new Random().Next().ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.ShowOnlyInactive, true);

            if (itemPage.CheckTotalNumber() < 20)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                itemPage = itemGeneralInformationPage.BackToList();
            }

            //Act
            itemPage.Filter(ItemPage.FilterType.ShowOnlyInactive, true);

            if (!itemPage.isPageSizeEqualsTo100())
            {
                itemPage.PageSize("8");
                itemPage.PageSize("100");
            }

            //Assert
            Assert.IsFalse(itemPage.CheckStatus(false), string.Format(MessageErreur.FILTRE_ERRONE, "Show inactive only"));
        }

        //Filtre Active
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Filter_ShowActive()
        {
            //Prepare
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
            string itemName = itemNameToday + "-" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();
            homePage.Navigate();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.ShowOnlyActive, true);

            if (itemPage.CheckTotalNumber() < 20)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                itemPage = itemGeneralInformationPage.BackToList();
                itemPage.ResetFilter();
                itemPage.Filter(ItemPage.FilterType.ShowOnlyActive, true);
            }

            if (!itemPage.isPageSizeEqualsTo100())
            {
                itemPage.PageSize("8");
                itemPage.PageSize("100");
            }

            //Assert
            Assert.IsTrue(itemPage.CheckStatus(true), string.Format(MessageErreur.FILTRE_ERRONE, "Show active only"));
        }

        //Filtre Show All
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_ITEM_Filter_ShowAll()
        {
            //Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
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
            bool newVersionPrint = true;
            string itemName = itemNameToday + "-" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ClearDownloads();
            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
            itemPage = itemGeneralInformationPage.BackToList();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.ShowAll, true);
            var totalNbreAll = itemPage.CheckTotalNumber();
            itemPage.Filter(ItemPage.FilterType.ShowOnlyInactive, true);
            var totalNbreInactiveOnly = itemPage.CheckTotalNumber();
            itemPage.Filter(ItemPage.FilterType.ShowOnlyActive, true);
            var totalNbreActiveOnly = itemPage.CheckTotalNumber();
            itemPage.Filter(ItemPage.FilterType.ActiveNotPurchasable, true);
            var totalNbreActiveNotPurchasable = itemPage.CheckTotalNumber();
            // Assert
            Assert.AreEqual(totalNbreAll, (totalNbreInactiveOnly + totalNbreActiveOnly + totalNbreActiveNotPurchasable), String.Format(MessageErreur.FILTRE_ERRONE, "'Show all'"));

            itemPage.Filter(ItemPage.FilterType.Site, site);
            itemPage.Filter(ItemPage.FilterType.Supplier, supplier);
            itemPage.Filter(ItemPage.FilterType.Group, group);

            // On exporte les résultats sous la forme d'un fichier Excel (dont on récupère le nom)
            // Export du fichier au format Excel
            itemPage.Export(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = itemPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;

            // On récupère le nombre de résultats issus du filtre et l'URL du fichier Excel à ouvrir
            var filePath = Path.Combine(downloadsPath, fileName);

            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber(SHEET1, filePath);

            //Assert
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
        }


        //Filtre Active Not purchasable     
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_ITEM_Filter_ActiveNotPurchasable()
        {
            //Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            //Prepare
            string itemName = $"NotPurchasableItem-{DateTime.Now.ToString("dd/MM/yyyy/HH")}";
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
            bool newVersionPrint = true;
            //Arrange
            var homePage = LogInAsAdmin();
            homePage.Navigate();
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ClearDownloads();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.ActiveNotPurchasable, true);
            itemPage.Filter(ItemPage.FilterType.Group, group);
            itemPage.Filter(ItemPage.FilterType.Site, site);
            itemPage.Filter(ItemPage.FilterType.Supplier, supplier);
            itemPage.PageUp();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            if (!itemPage.IsItemExist(itemName))
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                itemGeneralInformationPage.SetUnpurchasableItem(site, supplier, packagingName);
                itemGeneralInformationPage.ConfirmSetUnpurchasableReportForSupplier();
                itemGeneralInformationPage.BackToList();
            }
            else
            {
                var itemGeneralInformationPage = itemPage.ClickOnFirstItem();
                if (!itemGeneralInformationPage.IsPackageExist(packagingName))
                {
                    var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                    itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                    itemGeneralInformationPage.SetUnpurchasableItem(site, supplier, packagingName);
                    itemGeneralInformationPage.ConfirmSetUnpurchasableReportForSupplier();
                }
                itemGeneralInformationPage.BackToList();
            }
            // On exporte les résultats sous la forme d'un fichier Excel (dont on récupère le nom)
            // Export du fichier au format Excel
            itemPage.Export(newVersionPrint);
            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();
            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = itemPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);
            var fileName = correctDownloadedFile.Name;
            // On récupère le nombre de résultats issus du filtre et l'URL du fichier Excel à ouvrir
            var filePath = Path.Combine(downloadsPath, fileName);
            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber("Sheet 1", filePath);
            bool result = OpenXmlExcel.ReadAllDataInColumn("Is purchasable", "Sheet 1", filePath, "false");
            //Assert
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.IsTrue(result, MessageErreur.EXCEL_DONNEES_KO);
        }

        //Filtre Group
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Filter_Group()
        {
            //Prepare
            string itemName = itemNameToday + "-" + new Random().Next().ToString();
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
            //Arrange
            var homePage = LogInAsAdmin();
            homePage.Navigate();
            homePage.SetNewVersionKeywordValue(true);

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Group, group);

            if (itemPage.CheckTotalNumber() < 20)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                itemPage = itemGeneralInformationPage.BackToList();
                itemPage.ResetFilter();
                itemPage.Filter(ItemPage.FilterType.Group, group);
            }

            if (!itemPage.isPageSizeEqualsTo100())
            {
                itemPage.PageSize("8");
                itemPage.PageSize("100");
            }

            //Assert
            Assert.IsTrue(itemPage.VerifyGroup(group), string.Format(MessageErreur.FILTRE_ERRONE, "Group"));
        }

        //Filtre Site
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_ITEM_Filter_Site()
        {
            //Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string itemName = itemNameToday + "-" + new Random().Next().ToString();
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
            bool newVersion = true;
            //Arrange
            var homePage = LogInAsAdmin();
            homePage.Navigate();
            homePage.SetNewVersionKeywordValue(true);
            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.ClearDownloads();
            itemPage.Filter(ItemPage.FilterType.Site, site);
            itemPage.Filter(ItemPage.FilterType.Supplier, supplier);
            if (itemPage.CheckTotalNumber() < 20)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                itemPage = itemGeneralInformationPage.BackToList();
                itemPage.ResetFilter();
                itemPage.Filter(ItemPage.FilterType.Site, site);
                itemPage.Filter(ItemPage.FilterType.Supplier, supplier);
            }

            // On exporte les résultats sous la forme d'un fichier Excel (dont on récupère le nom)
            // Export du fichier au format Excel
            itemPage.Export(newVersion);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = itemPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;

            // On récupère le nombre de résultats issus du filtre et l'URL du fichier Excel à ouvrir
            var filePath = Path.Combine(downloadsPath, fileName);

            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber(SHEET1, filePath);
            bool result = OpenXmlExcel.ReadAllDataInColumn("Site", SHEET1, filePath, site);

            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.IsTrue(result, MessageErreur.EXCEL_DONNEES_KO);
        }

        //Filtre Site
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Filter_Supplier()
        {
            //Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string itemName = itemNameToday + "-" + new Random().Next().ToString();
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
            bool newVersion = true;
            //Arrange
            var homePage = LogInAsAdmin();
            homePage.Navigate();
            homePage.SetNewVersionKeywordValue(true);

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.ClearDownloads();
            itemPage.Filter(ItemPage.FilterType.Supplier, supplier);
            itemPage.Filter(ItemPage.FilterType.Site, site);

            if (itemPage.CheckTotalNumber() < 20)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                itemPage = itemGeneralInformationPage.BackToList();

                itemPage.ResetFilter();
                itemPage.Filter(ItemPage.FilterType.Supplier, supplier);
            }

            // On exporte les résultats sous la forme d'un fichier Excel (dont on récupère le nom)
            // Export du fichier au format Excel
            itemPage.Export(newVersion);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = itemPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;

            // On récupère le nombre de résultats issus du filtre et l'URL du fichier Excel à ouvrir
            var filePath = Path.Combine(downloadsPath, fileName);

            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber(SHEET1, filePath);
            bool result = OpenXmlExcel.ReadAllDataInColumn("Supplier", SHEET1, filePath, supplier);

            //Assert
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.IsTrue(result, MessageErreur.EXCEL_DONNEES_KO);
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_ITEM_Filter_Allergen()
        {
            //Prepare
            string itemName = itemNameToday + "-" + new Random().Next().ToString();
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
            string allergens = "Gluten/Gluten";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            homePage.SetNewVersionKeywordValue(true);

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            //Create
            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
            var itemIntolerancePage = itemGeneralInformationPage.ClickOnIntolerancePage();
            itemIntolerancePage.AddAllergen(allergens);
            itemIntolerancePage.SaveAllergen();
            itemIntolerancePage.WaitLoading();
            //Filter
            itemPage = itemGeneralInformationPage.BackToList();
            itemPage.Filter(ItemPage.FilterType.Allergens, allergens);
            itemPage.PageUp();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            //Assert
            string firstItemName = itemPage.GetFirstItemName();
            Assert.AreEqual(itemName, firstItemName, string.Format(MessageErreur.FILTRE_ERRONE, "Allergen"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Filter_Keyword()
        {
            //Prepare
            var keyword = TestContext.Properties["Item_Keyword"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            bool newVersionPrint = true;

            //Arrange
            var homePage = LogInAsAdmin();

            homePage.SetNewVersionKeywordValue(true);

            //GO TO ITEM PAGE 
            var itemPage = homePage.GoToPurchasing_ItemPage();

            //FILTER KEYWORD 
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

            //EXPORT EXCEL 

            itemPage.ClearDownloads();

            // On exporte les résultats sous la forme d'un fichier Excel (dont on récupère le nom)
            // Export du fichier au format Excel
            itemPage.Export(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = itemPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;

            // On récupère le nombre de résultats issus du filtre et l'URL du fichier Excel à ouvrir
            var filePath = Path.Combine(downloadsPath, fileName);

            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber("Sheet 1", filePath);
            bool result = OpenXmlExcel.ReadAllDataInColumn("Keyword", "Sheet 1", filePath, "TEST_KEY");

            //Assert
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.IsTrue(result, MessageErreur.EXCEL_DONNEES_KO);
        }

        //Exporter les données d'1 groupe d'article
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Export_Items_From_Site_And_Group_NewVersion()
        {
            //Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string group = TestContext.Properties["Item_Group"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            bool newVersionPrint = true;

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ClearDownloads();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Site, site);
            itemPage.Filter(ItemPage.FilterType.Group, group);
            itemPage.Export(newVersionPrint);

            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = itemPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber(SHEET1, filePath);
            bool resultSite = OpenXmlExcel.ReadAllDataInColumn("Site", SHEET1, filePath, site);
            bool resultGroup = OpenXmlExcel.ReadAllDataInColumn("Group", SHEET1, filePath, group);

            //Assert
            Assert.AreNotEqual(0, resultNumber);
            Assert.IsTrue(resultSite);
            Assert.IsTrue(resultGroup);
        }

        //Exporter les données d'1 seul site
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Export_Items_From_Site_NewVersion()
        {
            //Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            bool newVersionPrint = true;

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.ClearDownloads();

            itemPage.Filter(ItemPage.FilterType.Site, site);
            int numberOfRows = itemPage.CheckTotalNumber();
            if (numberOfRows > 0)
            {
                var firstRowGroup = itemPage.GetFirstItemGroup();
                itemPage.Filter(ItemPage.FilterType.Group, firstRowGroup);
                numberOfRows = itemPage.CheckTotalNumber();
            }
            itemPage.Export(newVersionPrint);

            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = itemPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber(SHEET1, filePath);
            bool resultSite = OpenXmlExcel.ReadAllDataInColumn("Site", SHEET1, filePath, site);

            //Assert
            if (numberOfRows > 0)
            {
                Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
            }
            Assert.IsTrue(resultSite, MessageErreur.EXCEL_DONNEES_KO);


        }

        //Exporter les données d'1 fournisseur sur plusieurs sites       
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_ITEM_Export_Items_From_Sites_And_Supplier_NewVersion()
        {
            //Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString(); // C:\ChromeDriverDownloads
            string siteMAD = TestContext.Properties["Site"].ToString();
            string siteBCN = TestContext.Properties["InvoiceSite"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string group = "PLATOS CALIENTES CONG";

            //Arrange
            var homePage = LogInAsAdmin();
            bool newVersionPrint = true;

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            DeleteAllFileDownload();
            itemPage.ClearDownloads();
            itemPage.DoubleFilterSite(siteMAD, siteBCN);
            itemPage.Filter(ItemPage.FilterType.Supplier, supplier);
            itemPage.Filter(ItemPage.FilterType.Group, group);
            itemPage.Export(newVersionPrint);

            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = itemPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber(SHEET1, filePath);
            bool resultSupplier = OpenXmlExcel.ReadAllDataInColumn("Supplier", SHEET1, filePath, supplier);

            //Assert
            Assert.AreNotEqual(0, resultNumber);
            Assert.IsTrue(resultSupplier);
        }
        
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Items_Filter_UseCase_Filter_Search()
        {
            //Prepare
            string recipeName = "Recipe for Usecase 1";
            string itemName = "Item for useCase";
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string siteAce = TestContext.Properties["SiteLP"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();
            string itemForUseCase = "Item for useCase";
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string recipeTypeBis = TestContext.Properties["RecipeTypeBis"].ToString();
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string variantBis = TestContext.Properties["RecipeVariant3"].ToString();
            string recipeForUseCase1 = "Recipe for Usecase 1";
            int nbPortions = new Random().Next(1, 30);
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string serviceCategorie = TestContext.Properties["Production_Service1"].ToString();
            string serviceType = TestContext.Properties["Production_ServiceType"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                itemGeneralInformationPage.NewPackaging();
                itemCreatePackagingPage.FillField_CreateNewPackaging(siteAce, packagingName, storageQty, storageUnit, qty, supplier);
                var itemUseCasePage = itemGeneralInformationPage.ClickOnUseCasePage();
                itemUseCasePage.ResetFilter();
                var recipesPage = homePage.GoToMenus_Recipes();
                recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeForUseCase1);
                if (recipesPage.CheckTotalNumber() == 0)
                {
                    var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                    var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeForUseCase1, recipeType, nbPortions.ToString());
                    recipeGeneralInfosPage.AddVariantWithSite(site, variant);
                    var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
                    recipeVariantPage.AddIngredient(itemForUseCase);
                    recipeVariantPage.BackToList();
                }
                else
                {
                    var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
                    var recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();
                    if (recipeVariantPage.IsIngredientDisplayed())
                    {
                        recipeVariantPage.DeleteFirstIngredient(false);
                    }
                    recipeVariantPage.AddIngredient(itemForUseCase);
                    recipeVariantPage.BackToList();
                }
            }
            var itemGeneralInformationsPage = itemPage.ClickOnFirstItem();
            var ItemUseCasePage = itemGeneralInformationsPage.ClickOnUseCasePage();
            ItemUseCasePage.ResetFilter();

            // on utilise le filtre search sur le recipe name du premier élément de la liste
            ItemUseCasePage.Filter(ItemUseCasePage.FilterType.Search, recipeName);

            // on compare tous les éléments de la liste obtenue avec la valeur sur laquelle nous avons appliqué le filtre
            var ok = ItemUseCasePage.IsSearchFilterOK(recipeName);

            //Assert
            Assert.IsTrue(ok, String.Format(MessageErreur.FILTRE_ERRONE, "Search"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Items_UseCase_ResetFilter()
        {
            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            string itemName = "Item for useCase";
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            var itemGeneralInformationsPage = itemPage.ClickOnFirstItem();
            var itemUseCasePage = itemGeneralInformationsPage.ClickOnUseCasePage();
            itemUseCasePage.ResetFilter();

            //value of filters coming on the page
            var searchText = itemUseCasePage.GetSearchFilterText();
            var inactiveRecipesValues = itemUseCasePage.GetInactiveRecipesBool();
            var siteSelectedNumber = itemUseCasePage.GetSiteSelectedNumber();
            var recipeTypeSelectedNumber = itemUseCasePage.GetRecipeTypesSelectedNumber();
            var variantSelectedNumber = itemUseCasePage.GetVariantSelectedNumber();

            // on filtre sur le premier élément de la liste

            var firstUC_recipeName = itemUseCasePage.GetFirstUseCaseRecipeName();
            itemUseCasePage.Filter(ItemUseCasePage.FilterType.Search, firstUC_recipeName);
            itemUseCasePage.Filter(ItemUseCasePage.FilterType.ShowInactiveRecipe, true);
            var firstUC_site = itemUseCasePage.GetFirstUseCaseSite();
            itemUseCasePage.Filter(ItemUseCasePage.FilterType.Site, firstUC_site);

            var firstUC_recipeType = itemUseCasePage.GetFirstUseCaseRecipeType();
            itemUseCasePage.Filter(ItemUseCasePage.FilterType.RecipeTypes, firstUC_recipeType);

            var firstUC_variant = itemUseCasePage.GetFirstUseCaseVariant();
            itemUseCasePage.Filter(ItemUseCasePage.FilterType.Variant, firstUC_variant);

            //value of filters after filter
            var searchTextAfterFilter = itemUseCasePage.GetSearchFilterText();
            var inactiveRecipesValuesAfterFilter = itemUseCasePage.GetInactiveRecipesBool();
            var siteSelectedNumberAfterFilter = itemUseCasePage.GetSiteSelectedNumber();
            var recipeTypeSelectedNumberAfterFilter = itemUseCasePage.GetRecipeTypesSelectedNumber();
            var variantSelectedNumberAfterFilter = itemUseCasePage.GetVariantSelectedNumber();

            //Reset
            itemUseCasePage.ResetFilter();

            //value of filters after restefilters
            var searchTextResetFilter = itemUseCasePage.GetSearchFilterText();
            var inactiveRecipesValuesResetFilter = itemUseCasePage.GetInactiveRecipesBool();
            var siteSelectedNumberResetFilter = itemUseCasePage.GetSiteSelectedNumber();
            var recipeTypeSelectedNumberResetFilter = itemUseCasePage.GetRecipeTypesSelectedNumber();
            var variantSelectedNumberResetFilter = itemUseCasePage.GetVariantSelectedNumber();

            //asserts
            Assert.AreNotEqual(searchTextAfterFilter, searchTextResetFilter, "Le filtre Search n'a pas été réinitialisé.");
            Assert.AreEqual(searchText, searchTextResetFilter, "Le filtre Search n'a pas été réinitialisé.");

            Assert.AreNotEqual(inactiveRecipesValuesAfterFilter, inactiveRecipesValuesResetFilter, "Le filtre Show inactive recipes n'a pas été réinitialisé.");
            Assert.AreEqual(inactiveRecipesValues, inactiveRecipesValuesResetFilter, "Le filtre Show inactive recipes n'a pas été réinitialisé.");

            Assert.AreNotEqual(siteSelectedNumberAfterFilter, siteSelectedNumberResetFilter, "Le filtre Site n'a pas été réinitialisé.");
            Assert.AreEqual(siteSelectedNumber, siteSelectedNumberResetFilter, "Le filtre Site n'a pas été réinitialisé.");

            Assert.AreNotEqual(recipeTypeSelectedNumberAfterFilter, recipeTypeSelectedNumberResetFilter, "Le filtre Recipe Tyes n'a pas été réinitialisé.");
            Assert.AreEqual(recipeTypeSelectedNumber, recipeTypeSelectedNumberResetFilter, "Le filtre Recipe Tyes n'a pas été réinitialisé.");

            Assert.AreNotEqual(variantSelectedNumberAfterFilter, variantSelectedNumberResetFilter, "Le filtre Variant n'a pas été réinitialisé.");
            Assert.AreEqual(variantSelectedNumber, variantSelectedNumberResetFilter, "Le filtre Variant n'a pas été réinitialisé.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Items_Filter_UseCase_ShowInactive()
        {
            //Arrange

            HomePage homePage = LogInAsAdmin();
            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            string itemName = "Item for useCase";
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            var itemGeneralInformationsPage = itemPage.ClickOnFirstItem();
            var itemUseCasePage = itemGeneralInformationsPage.ClickOnUseCasePage();
            itemUseCasePage.ResetFilter();
            var totalUseCase = itemUseCasePage.CheckTotalNumber();
            itemUseCasePage.Filter(ItemUseCasePage.FilterType.ShowInactiveRecipe, true);
            //Assert
            Assert.AreNotEqual(itemUseCasePage.CheckTotalNumber(), totalUseCase, MessageErreur.FILTRE_ERRONE, "Show inactive recipes");
            itemUseCasePage.Filter(ItemUseCasePage.FilterType.ShowInactiveRecipe, false);
            bool value = itemUseCasePage.verifyShowInactive(false);
            Assert.IsTrue(value, MessageErreur.FILTRE_ERRONE, "Show inactive recipes");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Items_Filter_UseCase_Site()
        {
            //Prepare

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            string itemName = "Item for useCase";
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            var itemGeneralInformationsPage = itemPage.ClickOnFirstItem();
            var itemUseCasePage = itemGeneralInformationsPage.ClickOnUseCasePage();
            itemUseCasePage.ResetFilter();
            var firstUC_site = itemUseCasePage.GetFirstUseCaseSite();
            itemUseCasePage.Filter(ItemUseCasePage.FilterType.Site, firstUC_site);
            bool isSiteExist = itemUseCasePage.IsSiteFilterOK(firstUC_site);

            //Assert
            Assert.IsTrue(isSiteExist, String.Format(MessageErreur.FILTRE_ERRONE, "Site"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Items_Filter_UseCase_RecipeType()
        {
            //Arrange

            HomePage homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            string itemName = "Item for useCase";
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            var itemGeneralInformationsPage = itemPage.ClickOnFirstItem();
            var itemUseCasePage = itemGeneralInformationsPage.ClickOnUseCasePage();
            itemUseCasePage.ResetFilter();
            var firstUC_recipeType = itemUseCasePage.GetFirstUseCaseRecipeType();
            itemUseCasePage.Filter(ItemUseCasePage.FilterType.RecipeTypes, firstUC_recipeType);
            bool isrecipetypeExist = itemUseCasePage.isRecipeTypesFilterOK(firstUC_recipeType);
            //Assert
            Assert.IsTrue(isrecipetypeExist, String.Format(MessageErreur.FILTRE_ERRONE, "Recipe Types"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Items_Filter_UseCase_Variant()
        {
            //Arrange

            HomePage homePage = LogInAsAdmin();
            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            string itemName = "Item for useCase";
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            var itemGeneralInformationsPage = itemPage.ClickOnFirstItem();
            var itemUseCasePage = itemGeneralInformationsPage.ClickOnUseCasePage();
            itemUseCasePage.ResetFilter();
            var firstUC_variant = itemUseCasePage.GetFirstUseCaseVariant();
            itemUseCasePage.Filter(ItemUseCasePage.FilterType.Variant, firstUC_variant);
            bool isVariantExist = itemUseCasePage.isVariantFilterOK(firstUC_variant);

            //Assert
            Assert.IsTrue(isVariantExist, String.Format(MessageErreur.FILTRE_ERRONE, "Site"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Dietetic_Change_Informations()
        {
            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            var itemGeneralInformationsPage = itemPage.ClickOnFirstItem();
            var itemDieteticPage = itemGeneralInformationsPage.ClickOnDieteticPage();
            string numericalTestValue = "1500";
            itemDieteticPage.SetEnergyKJ_numericalValue(numericalTestValue);
            itemGeneralInformationsPage.ClickOnUseCasePage();
            itemGeneralInformationsPage.ClickOnDieteticPage();
            //Assert
            Assert.AreEqual(numericalTestValue, itemDieteticPage.GetEnergyKJ_numericalValue(), (MessageErreur.MESSAGE_ERREUR_MODIFICATION_NON_ENREGISTREE + " EnergyKJ"));
            itemDieteticPage.SetEnergyKJ_numericalValue("0");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Apply_CLIQUAL()
        {
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
            string itemName = itemNameToday + "-" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
            itemPage = itemGeneralInformationPage.BackToList();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            //
            var itemGeneralInformationsPage = itemPage.ClickOnFirstItem();
            var itemDieteticPage = itemGeneralInformationsPage.ClickOnDieteticPage();
            var dieteticPage = itemGeneralInformationsPage.ClickOnDieteticPage();

            //on recupere la valeur initiale
            var initialValue = dieteticPage.GetEnergyKJ_numericalValue();
            var searchCiqualModal = itemDieteticPage.ClickOnCiqualButton();
            searchCiqualModal.SearchCiqualItem("a");
            searchCiqualModal.ClickOnFirstSuggestedCiqualDisplayed();
            searchCiqualModal.SelectFirstCiqualDisplayed();
            searchCiqualModal.ApplyCiqualOnItem();

            //recupere la nouvelle valeur
            var newValue = dieteticPage.GetEnergyKJ_numericalValue();

            Assert.AreNotEqual(initialValue, newValue, MessageErreur.MESSAGE_ERREUR_MODIFICATION_NON_ENREGISTREE, "APPLY CIQUAL");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Intolerance_Apply()
        {
            //Arrange

            HomePage homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            var itemGeneralInformationsPage = itemPage.ClickOnFirstItem();
            var itemIntolerancePage = itemGeneralInformationsPage.ClickOnIntolerancePage();
            string allergen = TestContext.Properties["RecipeAllergen"].ToString();
            itemIntolerancePage.AddAllergen(allergen);
            itemGeneralInformationsPage.ClickOnUseCasePage();
            itemGeneralInformationsPage.ClickOnIntolerancePage();
            Assert.IsTrue(itemIntolerancePage.IsAllergenChecked(allergen), MessageErreur.MESSAGE_ERREUR_MODIFICATION_NON_ENREGISTREE, "Intolerance : ", allergen);
            itemIntolerancePage.UncheckAllergen(allergen);
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Label_Change()
        {
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
            string itemName = itemNameToday + "-" + new Random().Next().ToString();

            //Arrange

            HomePage homePage = LogInAsAdmin();


            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
            itemPage = itemGeneralInformationPage.BackToList();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            var itemGeneralInformationsPage = itemPage.ClickOnFirstItem();
            var itemLabelPage = itemGeneralInformationsPage.ClickOnLabelPage();
            itemLabelPage.EditLabel("someText");
            itemLabelPage.Validate();
            itemGeneralInformationsPage.ClickOnDieteticPage();
            itemGeneralInformationsPage.ClickOnLabelPage();

            Assert.IsTrue(itemLabelPage.IsEditLabel("someText"), "someText", (MessageErreur.MESSAGE_ERREUR_MODIFICATION_NON_ENREGISTREE, "Intolerance : ", "someText"));
        }

        [TestMethod]
        [Priority(3)]
        [Timeout(_timeout)]
        public void PU_ITEM_Show_Last_RN()
        {
            //prepare
            //string itemName = "LastItemReceipt" + DateUtils.Now.ToString("dd/MM/yyyy");
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            //Search Item 
            itemPage.Filter(ItemPage.FilterType.ShowAll, true);
            //itemPage.Filter(ItemPage.FilterType.Search, itemName);
            string itemName = itemPage.GetFirstItemName();
            var itemGeneralInformationsPage = itemPage.ClickOnFirstItem();
            var lastReceiptNotesPage = itemGeneralInformationsPage.ClickOnLastReceiptNotesPage();
            // Create receipt note if not created during PU_ITEM_Items_CreateItemForReceiptNote()
            if (!lastReceiptNotesPage.IsLinkToRNOK())
            {
                var purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();
                var createPurchaseOrderPage = purchaseOrdersPage.CreateNewPurchaseOrder();
                createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+10), true);
                var purchaseOrderItemPage = createPurchaseOrderPage.Submit();
                itemName = purchaseOrderItemPage.GetFirstItemName();
                purchaseOrderItemPage.Filter(PurchaseOrderItem.FilterItemType.ByName, itemName);
                purchaseOrderItemPage.SelectFirstItemPo();
                purchaseOrderItemPage.AddQuantity("2");
                purchaseOrderItemPage.Validate();
                purchaseOrderItemPage.GenerateReceiptNote(true);
                var receiptNotesItem = purchaseOrderItemPage.ValidateReceiptNoteCreation();
                var qualityChecks = receiptNotesItem.ClickOnChecksTab();
                qualityChecks.DeliveryAccepted();
                if (qualityChecks.CanClickOnSecurityChecks())
                {
                    qualityChecks.CanClickOnSecurityChecks();
                    qualityChecks.SetSecurityChecks("No");
                    qualityChecks.SetQualityChecks();
                    receiptNotesItem = qualityChecks.ClickOnReceiptNoteItemTab();
                }
                else
                {
                    receiptNotesItem = qualityChecks.ClickOnReceiptNoteItemTab();
                }

                receiptNotesItem.Validate();
                itemPage = homePage.GoToPurchasing_ItemPage();
                itemPage.ResetFilter();
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                itemGeneralInformationsPage = itemPage.ClickOnFirstItem();
                lastReceiptNotesPage = itemGeneralInformationsPage.ClickOnLastReceiptNotesPage();
            }
            var isLinkToRNOK = lastReceiptNotesPage.IsLinkToRNOK();
            Assert.IsTrue(isLinkToRNOK, (MessageErreur.MESSAGE_ERREUR_LAST_RECEIPT_NOTE));

        }

        [TestMethod]
        [Priority(4)]
        [Timeout(_timeout)]
        public void PU_ITEM_Add_Picture()
        {
            //prepare
            string itemName = "ZUMOS FRUTAS NATURALES";
            var path = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase);
            path = path.Substring(6);
            var imagePath = path + "\\PageObjects\\Purchasing\\Item\\test.jpg";
            Assert.IsTrue(new FileInfo(imagePath).Exists, "Fichier " + imagePath + " non trouvé");

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            var itemGeneralInformationsPage = itemPage.ClickOnFirstItem();
            var picturePage = itemGeneralInformationsPage.ClickOnPicturePage();
            if (picturePage.IsPictureAdded())
            {
                picturePage.DeletePicture();
            }

            picturePage.AddPicture(imagePath);
            Assert.IsTrue(picturePage.IsPictureAdded(), MessageErreur.ERREUR_UPLOAD_PICTURE, "IMAGE NON AJOUTEE");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Delete_Picture()
        {
            //Prepare
            string itemName = "Item for useCase";


            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            var itemGeneralInformationsPage = itemPage.ClickOnFirstItem();

            var picturePage = itemGeneralInformationsPage.ClickOnPicturePage();

            if (picturePage.IsPictureAdded())
            {
                picturePage.DeletePicture();
            }

            Assert.IsTrue(picturePage.IsPictureDeleted(), MessageErreur.ERREUR_UPLOAD_PICTURE, "IMAGE NON SUPPRIMEE");
        }

        [TestMethod]
        [Priority(1)]
        [Timeout(_timeout)]
        public void PU_ITEM_Items_CreateItemWithRecipes()
        {
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string siteAce = TestContext.Properties["SiteLP"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();
            string itemForUseCase = "Item for useCase";
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string recipeTypeBis = TestContext.Properties["RecipeTypeBis"].ToString();
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string variantBis = TestContext.Properties["RecipeVariant3"].ToString();
            string recipeForUseCase1 = "Recipe for Usecase 1";
            string recipeForUseCase2 = "Recipe for Usecase 2";
            int nbPortions = new Random().Next(1, 30);
            string datasheetForUseCase = "Datasheet for Usecase";
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string serviceName = "Service for Usecase";
            string serviceCategorie = TestContext.Properties["Production_Service1"].ToString();
            string serviceType = TestContext.Properties["Production_ServiceType"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            //1.Creation Item
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemForUseCase);
            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemForUseCase, group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                itemGeneralInformationPage.NewPackaging();
                itemCreatePackagingPage.FillField_CreateNewPackaging(siteAce, packagingName, storageQty, storageUnit, qty, supplier);
                var itemUseCasePage = itemGeneralInformationPage.ClickOnUseCasePage();
                itemUseCasePage.ResetFilter();
            }
            //2.Add item in recipes
            //first recipe
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeForUseCase1);
            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeForUseCase1, recipeType, nbPortions.ToString());
                recipeGeneralInfosPage.AddVariantWithSite(site, variant);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
                recipeVariantPage.AddIngredient(itemForUseCase);
                recipeVariantPage.BackToList();
            }
            else
            {
                var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
                var recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();
                if (recipeVariantPage.IsIngredientDisplayed())
                {
                    recipeVariantPage.DeleteFirstIngredient(false);
                }
                recipeVariantPage.AddIngredient(itemForUseCase);
                recipeVariantPage.BackToList();
            }
            //second recipe inactive
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeForUseCase2);
            recipesPage.Filter(RecipesPage.FilterType.ShowAll, true);
            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeForUseCase2, recipeTypeBis, nbPortions.ToString(), false);
                recipeGeneralInfosPage.AddVariantWithSite(siteAce, variantBis);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
                recipeVariantPage.AddIngredient(itemForUseCase);
            }
            else
            {
                var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
                var recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();
                if (recipeVariantPage.IsIngredientDisplayed())
                {
                    recipeVariantPage.DeleteFirstIngredient(false);
                }
                recipeVariantPage.AddIngredient(itemForUseCase);
            }
            //3.Add recipe in datasheet
            var datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetForUseCase);
            if (datasheetPage.CheckTotalNumber() == 0)
            {
                var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
                var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetForUseCase, guestType, site);
                datasheetDetailPage.AddRecipe(recipeForUseCase1);
                //Assert
                Assert.IsTrue(datasheetDetailPage.IsRecipeAdded(), "La recette n'a pas été rajoutée à la datasheet.");
            }
            //4.Create service link to datsheet and check dates
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, null, null, serviceCategorie, null, serviceType);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                serviceGeneralInformationsPage.SetProduced(true);
                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                priceModalPage.FillFields_CustomerPrice(site, supplier, DateUtils.Now.AddDays(-20), DateUtils.Now.AddMonths(2), "10", datasheetForUseCase);
            }
            else
            {
                var pricePage = servicePage.ClickOnFirstService();
                pricePage.ToggleFirstPrice();
                var serviceCreatePriceModalPage = pricePage.EditFirstPrice(site, supplier);
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
            }
            // Assert 
            homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemForUseCase.ToString());
            Assert.AreEqual(itemForUseCase, itemPage.GetFirstItemName(), "Le nouvel item n'a pas été créé et ajouté à la liste.");

        }

        [TestMethod]
        [Priority(2)]
        [Timeout(_timeout)]
        public void PU_ITEM_Items_CreateItemForReceiptNote()
        {
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string siteAce = TestContext.Properties["SiteLP"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();
            string itemName = "LastItemReceipt" + DateUtils.Now.ToString("dd/MM/yyyy");
            var location = TestContext.Properties["Location"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            if (itemPage.CheckTotalNumber() == 0)
            {
                // not Purchasable (no packaging)
                itemPage.Filter(ItemPage.FilterType.ShowAll, true);

                ItemGeneralInformationPage itemGeneralInformationPage;
                if (itemPage.CheckTotalNumber() == 0)
                {
                    //Creation Item
                    var itemCreateModalPage = itemPage.ItemCreatePage();
                    itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
                }
                else
                {
                    itemGeneralInformationPage = itemPage.ClickOnFirstItem();
                }

                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                itemGeneralInformationPage.NewPackaging();
                itemCreatePackagingPage.FillField_CreateNewPackaging(siteAce, packagingName, storageQty, storageUnit, qty, supplier);
                //Act
                var purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();
                var createPurchaseOrderPage = purchaseOrdersPage.CreateNewPurchaseOrder();
                createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+10), true);
                var purchaseOrderItemPage = createPurchaseOrderPage.Submit();
                purchaseOrderItemPage.Filter(PurchaseOrderItem.FilterItemType.ByName, itemName);
                purchaseOrderItemPage.SelectFirstItemPo();
                purchaseOrderItemPage.AddQuantity("2");
                purchaseOrderItemPage.Validate();
                purchaseOrderItemPage.GenerateReceiptNote(true);
                var receiptNotesItem = purchaseOrderItemPage.ValidateReceiptNoteCreation();
                var qualityChecks = receiptNotesItem.ClickOnChecksTab();
                qualityChecks.DeliveryAccepted();
                if (qualityChecks.CanClickOnSecurityChecks())
                {
                    qualityChecks.CanClickOnSecurityChecks();
                    qualityChecks.SetSecurityChecks("No");
                    qualityChecks.SetQualityChecks();
                    receiptNotesItem = qualityChecks.ClickOnReceiptNoteItemTab();
                }
                else
                {
                    receiptNotesItem = qualityChecks.ClickOnReceiptNoteItemTab();
                }
                receiptNotesItem.Validate();
                homePage.GoToPurchasing_ItemPage();
                itemPage.ResetFilter();

            }

            // Assert 
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            Assert.AreEqual(itemName, itemPage.GetFirstItemName(), "Le nouvel item n'a pas été créé et ajouté à la liste.");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Item_PrintUseCasesPDF()
        {
            // pas d'impression, ça bloque les tests
            //Assert.Fail("à corriger après le ticket 18727");

            //Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();
            homePage.ClearDownloads();
            homePage.PurgeDownloads();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();

            // purge
            DirectoryInfo directory = new DirectoryInfo(downloadsPath);
            foreach (FileInfo fi in directory.GetFiles())
            {
                if (fi.Name.StartsWith("Item_") && fi.Name.EndsWith(".pdf"))
                {
                    fi.Delete();
                }
            }

            //Etre sur le détail d'un item / Onglet Use Case
            string firstItemName = "Item for useCase";
            itemPage.Filter(ItemPage.FilterType.Search, firstItemName);
            ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
            ItemUseCasePage useCase = generalInfo.ClickOnUseCasePage();
            useCase.Filter(ItemUseCasePage.FilterType.Site, "MAD");
            if (useCase.CheckTotalNumber() == 0)
            {
                //create a recipe that contains the item 
                var recipesPage = homePage.GoToMenus_Recipes();
                var recipe = recipesPage.SelectFirstRecipe();
                recipe.ClickFirstFlight();
                var firstIngredient = recipe.GetFirstIngredient();
                if (firstIngredient != null)
                {
                    firstItemName = firstIngredient;
                }
                else
                {
                    recipe.AddIngredient(firstItemName);
                }
                itemPage = homePage.GoToPurchasing_ItemPage();
                itemPage.Filter(ItemPage.FilterType.Search, firstItemName);
                generalInfo = itemPage.ClickOnFirstItem();
                useCase = generalInfo.ClickOnUseCasePage();
            }
            //Cliquer sur "Case Print use cases of this item in datasheets"
            useCase.PrintUseCaseReport();
            //Un fichier PDF et générer
            FileInfo filePdf = null;
            directory = new DirectoryInfo(downloadsPath);
            foreach (FileInfo fi in directory.GetFiles())
            {
                if (fi.Name.StartsWith("Item_") && fi.Name.EndsWith(".pdf"))
                {
                    filePdf = fi;
                    break;
                }
            }
            Assert.IsNotNull(filePdf, "Fichier non généré");
            Assert.IsTrue(filePdf.Exists, "Fichier non généré - cas 2");
            //Vérifier les données du fichier PDF
            useCase.CheckPrintUseCaseReport(filePdf);
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_PrintTheoricalReport_PDF()
        {
            //Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            var siteACE = TestContext.Properties["SiteACE"].ToString();
            var siteMAD = TestContext.Properties["SiteMAD"].ToString();
            string groupRYANAIR = "RYANAIR";
            //Purchasing Item Theorical Print_-_440279_-_20220701132827.pdf
            string DocFileNamePdfBegin = "Purchasing Item Theorical Print_-_";
            //All_files_20220225_102148.zip
            string DocFileNameZipBegin = "All_files_";

            //Arrange
            var homePage = LogInAsAdmin();

            homePage.ClearDownloads();
            DeleteAllFileDownload();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();

            //1. etre sur ongletg storage

            itemPage.Filter(ItemPage.FilterType.Site, siteACE);
            itemPage.Filter(ItemPage.FilterType.Group, groupRYANAIR);
            itemPage.ClickOnStorage();

            Assert.IsTrue(itemPage.CheckTotalNumber() > 0, "ACE n'a pas de RYANAIR");
            //1.Survoler les "…"
            //2.Cliquer sur "Print Théorical Report"
            //3.Choisir un site avec plusieurs storage place
            //4.sélectionner les storages place voulus
            // ???
            //5.Cliquer sur Print

            PrintReportPage printPage = itemPage.PrintTheoricalReport(siteACE);
            bool isGenerated = printPage.IsReportGenerated();
            printPage.Close();
            Assert.IsTrue(isGenerated, "document ACE non généré");
            printPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            string trouve = printPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            FileInfo pdfACE = new FileInfo(trouve);
            Assert.IsTrue(pdfACE.Exists, "document ACE non généré - cas 2");
            printPage.ClosePrintButton();

            //Le fichier PDF est généré et vérifier les données
            PdfDocument document = PdfDocument.Open(pdfACE.FullName);
            List<string> mots = new List<string>();
            foreach (Page p in document.GetPages())
            {
                foreach (var mot in p.GetWords())
                {
                    mots.Add(mot.Text);
                }
            }
            Assert.AreNotEqual(0, mots.Select(m => m == groupRYANAIR).Count(), "Le print theorical quantity sur le site ACE n'a pas d'items de group RYANAIR alors que sur l'onglet Purchase, on retrouve bien des items de group RYANAIR.");
            Assert.AreNotEqual(0, mots[mots.LastIndexOf("RYANAIR") + 3], "Sur le print theorical quantity, la Theo. value du groupe RYANAIR n'est pas positif alors qu'il n'y a des items sur l'onglet Storage.");

            itemPage.ClearDownloads();
            DeleteAllFileDownload();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Site, siteMAD);
            itemPage.Filter(ItemPage.FilterType.Group, groupRYANAIR);
            itemPage.Filter(ItemPage.FilterType.ShowAll, true);
            itemPage.ClickOnStorage();

            Assert.AreEqual(0, itemPage.CheckTotalNumber(), "MAD a du RYANAIR");

            printPage = itemPage.PrintTheoricalReport(siteMAD);
            isGenerated = printPage.IsReportGenerated();
            printPage.Close();
            Assert.IsTrue(isGenerated, "document MAD non généré");
            printPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            trouve = printPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            FileInfo pdfMAD = new FileInfo(trouve);
            Assert.IsTrue(pdfMAD.Exists, "document MAD non généré - cas 2");

            //Le fichier PDF est généré et vérifier les données
            document = PdfDocument.Open(pdfMAD.FullName);
            mots = new List<string>();
            foreach (Page p in document.GetPages())
            {
                foreach (var mot in p.GetWords())
                {
                    mots.Add(mot.Text);
                }
            }

            //il y a des items sur ce site donc on retrouve bien les items dans le pdf, cf compare à purchase, mais 0 quantité en storage d'où 0 items sur onglet storage
            Assert.AreNotEqual(0, mots.Select(m => m == groupRYANAIR).Count(), "Le print theorical quantity sur le site MAD n'a pas d'items de group RYANAIR alors que sur l'onglet Purchase, on retrouve bien des items de group RYANAIR.");
            Assert.AreEqual("0,0000", mots[mots.LastIndexOf("RYANAIR") + 3], "Sur le print theorical quantity, la Theo. value du groupe RYANAIR n'est pas à 0 alors qu'il n'y a pas d'items sur l'onglet Storage.");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Items_Filter_UseCase_Export()
        {
            //Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            //Arrange

            HomePage homePage = LogInAsAdmin();
            homePage.ClearDownloads();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            //Etre dans un item / Onglet Use case
            string itemName = "Item for useCase";
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
            ItemUseCasePage useCase = generalInfo.ClickOnUseCasePage();
            useCase.ResetFilter();

            // purge
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            foreach (FileInfo fi in taskDirectory.GetFiles())
            {
                if (fi.Name.StartsWith("Item_") && fi.Name.EndsWith(".xlsx"))
                {
                    fi.Delete();
                }
            }

            //Cliquer sur Export (téléchargement direct)
            PrintReportPage reportPage = useCase.Export(true);
            reportPage.Close();
            reportPage.ClosePrintButton();
            // n'a pas aimé le refresh du ClosePrintButton...
            useCase = generalInfo.ClickOnUseCasePage();

            //Fichier Excel est exporté et vérifier les données
            taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo trouve = null;
            foreach (FileInfo fi in taskDirectory.GetFiles())
            {
                if (fi.Name.StartsWith("Item_") && fi.Name.EndsWith(".xlsx"))
                {
                    trouve = fi;
                }
            }
            Assert.IsNotNull(trouve, "Fichier non généré");
            useCase.CheckExportUseCase(trouve, itemName);
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Item_UseCase_SortByRecipe()
        {
            //Prepare

            //Arrange
            HomePage homePage = LogInAsAdmin();
            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            //Etre sur un item, onglet Use Case et avoir des Recipe Types disponibles
            string itemName = "Item for useCase";
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
            ItemUseCasePage useCase = generalInfo.ClickOnUseCasePage();
            //Appliquer le filtre Sort By Recipe
            useCase.ResetFilter();
            useCase.Filter(ItemUseCasePage.FilterType.SortBy, "Recipe");
            useCase.Filter(ItemUseCasePage.FilterType.ShowInactiveRecipe, true);
            //Les données correspondent aux filtre appliquer
            var recipes = WebDriver.FindElements(By.XPath("//*/form/div[3]/div[2]/div[1]/span[1]"));
            List<string> avantTri = new List<string>();
            List<string> apresTri = new List<string>();
            foreach (var recipe in recipes)
            {
                avantTri.Add(recipe.Text);
                apresTri.Add(recipe.Text);
            }
            apresTri.Sort();
            for (int i = 0; i < avantTri.Count; i++)
            {
                Assert.AreEqual(avantTri[i], apresTri[i]);
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Item_UseCase_SelectAll()
        {
            //Prepare
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string itemName = itemNameToday + "-" + new Random().Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string recipeName1 = "Recipe1" + "-" + DateTime.Now.ToString();
            string recipeName2 = "Recipe2" + "-" + DateTime.Now.ToString();
            int nbPortions = new Random().Next(1, 30);

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            try
            {
                //Create an item
                ItemCreateModalPage itemCreateModalPage = itemPage.ItemCreatePage();
                ItemGeneralInformationPage itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
                ItemCreateNewPackagingModalPage itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);

                //Create Receipe 1
                RecipesPage recipesPage = homePage.GoToMenus_Recipes();
                recipesPage.ResetFilter();
                RecipesCreateModalPage recipesCreateModalPage = recipesPage.CreateNewRecipe();
                RecipeGeneralInformationPage recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName1, recipeType, nbPortions.ToString());
                recipeGeneralInfosPage.AddVariantWithSite(site, variant);
                if (!recipeGeneralInfosPage.IsFirstVariantIsVisible())
                {
                    recipeGeneralInfosPage.Validate();
                }
                RecipesVariantPage recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
                recipeVariantPage.AddIngredient(itemName);
                recipeVariantPage.BackToList();

                //Create New Receipe 2
                recipesPage = homePage.GoToMenus_Recipes();
                recipesPage.ResetFilter();
                recipesCreateModalPage = recipesPage.CreateNewRecipe();
                recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName2, recipeType, nbPortions.ToString());
                recipeGeneralInfosPage.AddVariantWithSite(site, variant);
                if (!recipeGeneralInfosPage.IsFirstVariantIsVisible())
                {
                    recipeGeneralInfosPage.Validate();
                }
                recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
                recipeVariantPage.AddIngredient(itemName);
                recipeVariantPage.BackToList();


                itemPage = homePage.GoToPurchasing_ItemPage();
                itemPage.ResetFilter();
                //Etre sur un item, onglet Use Case et avoir des Recipe Types disponibles

                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
                ItemUseCasePage useCase = generalInfo.ClickOnUseCasePage();
                //1. Cliquer sur Select All (tout est sélectionné)
                useCase.SelectAllRecipeTypes();
                //Tout se sélectionne 
                var totalNumber = useCase.CheckTotalNumber();
                var selectedNumber = useCase.CheckSelectedNumber();
                Assert.AreEqual(selectedNumber, totalNumber, "SelectAll ne fonctionne pas correctement");
            }

            finally
            {
                //Delete Receipe 1
                RecipesPage recipesPage = homePage.GoToMenus_Recipes();
                if (!string.IsNullOrEmpty(recipeName1))
                {
                    recipesPage.MassiveDeleteRecipes(recipeName1, site, recipeType);
                }
                homePage.GoToPurchasing_ItemPage();
                itemPage.ResetFilter();

                //Delete Receipe2
                recipesPage = homePage.GoToMenus_Recipes();
                if (!string.IsNullOrEmpty(recipeName2))
                {
                    recipesPage.MassiveDeleteRecipes(recipeName2, site, recipeType);
                }
                homePage.GoToPurchasing_ItemPage();
                itemPage.ResetFilter();

                //Desactivate packaging
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
                generalInfo.DeactivatePrice(site, supplier, packagingName);

                //Delete Item
                generalInfo.BackToList();
                MassiveDeleteModal massiveDelete = itemPage.MenuMassiveDelete();
                massiveDelete.Filter(MassiveDeleteModal.FilterType.SearchByName, itemName);
                massiveDelete.Filter(MassiveDeleteModal.FilterType.ShowInactiveSuppliers, true);
                massiveDelete.Filter(MassiveDeleteModal.FilterType.SupplierMultiple, supplier);
                massiveDelete.ClickSearch();
                massiveDelete.ClickSelectAllButton();
                massiveDelete.Delete();
            }
        }

        /**
         * je cherche à remplir UseCase d'un item
         * Réponse :
         * il faut créer une recette,
         * y ajouter un variant et sur ce variant ajouter l'item
         * (il faut bien que ton item ai un packaging sur le même site que ta recette)
         */
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Item_UseCase_ReplaceByAnotherItem()
        {
            //Prepare
            //Attention données créées ici PU_ITEM_Items_CreateItemWithRecipes et vérifier si use case ne sont pas rester sur TOnic
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string recipeForUseCase1 = "Recipe for Usecase 1";
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            int nbPortions = new Random().Next(1, 30);
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            //Etre sur un item, onglet Use Case et avoir des Recipe Types disponibles
            string itemName = "Item for useCase";
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
            string itemNameFrom = generalInfo.GetItemName();
            Console.WriteLine(DateUtils.Now.ToString("dd/MM/yyyy") + " itemName from :" + itemNameFrom);
            ItemUseCasePage useCase = generalInfo.ClickOnUseCasePage();
            var numberItems = useCase.GetNumberOfItems();
            if (numberItems == 0)
            {
                //.Add item in recipes
                //first recipe
                var recipesPage = homePage.GoToMenus_Recipes();
                recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeForUseCase1);
                if (recipesPage.CheckTotalNumber() == 0)
                {
                    var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                    var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeForUseCase1, recipeType, nbPortions.ToString());
                    recipeGeneralInfosPage.AddVariantWithSite(site, variant);
                    var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
                    recipeVariantPage.AddIngredient(itemNameFrom);
                    recipeVariantPage.BackToList();
                }
                else
                {
                    var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
                    var recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();
                    recipeVariantPage.DeleteFirstIngredient(false);
                    recipeVariantPage.AddIngredient(itemNameFrom);
                    recipeVariantPage.BackToList();
                }
                homePage.GoToPurchasing_ItemPage();
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                itemPage.ClickOnFirstItem();
                generalInfo.ClickOnUseCasePage();
            }

            //1. Cocher une Recipe Type
            var recipeName = useCase.GetFirstUseCaseRecipeName();
            Console.WriteLine(DateUtils.Now.ToString("dd/MM/yyyy") + " RecipeName voyage :" + recipeName);
            useCase.SelectBoxFirstUseCase();

            //2.Cliquer sur "Replace by another item"
            //3.Rentre les primière lettre d'un item et selectionner le
            var itemNameTo = useCase.ReplaceByAnotherItem("AGUA INDIAN TONIC 15 CL LATA", true);
            Console.WriteLine(DateUtils.Now.ToString("dd/MM/yyyy") + " itemName to :" + itemNameTo);

            //4.Cliquer sur "Replace with this item" et Confirmer

            //Vérifier que l'item a bien été remplacer dans la liste.
            itemPage = useCase.BackToList();
            itemPage.Filter(ItemPage.FilterType.ShowAll, true);
            itemPage.Filter(ItemPage.FilterType.Search, itemNameTo);
            Assert.AreEqual(itemNameTo, itemPage.GetFirstItemName());
            generalInfo = itemPage.ClickOnFirstItem();
            useCase = generalInfo.ClickOnUseCasePage();
            // recherche recipeName dans la liste UseCase
            useCase.SelectUseCase(recipeName);
            useCase.ReplaceByAnotherItem(itemNameFrom, true);
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Item_UseCase_MultipleUpdate()
        {
            //Prepare

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            //Etre sur un item, onglet Use Case et avoir des Recipe Types disponibles
            string itemName = "Item for useCase";
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
            ItemUseCasePage useCase = generalInfo.ClickOnUseCasePage();
            //1.Cocher une Recipe Type
            useCase.SelectBoxFirstUseCase();
            useCase.MultipleUpdate("Yield", "14");
            //5.Les données sont modifiées Vérifier que les modifications ont été prises en compte
            var yield = useCase.GetColumnFirstValue(ItemUseCasePage.ColumnName.Yield);
            Assert.AreEqual("14", yield, "multi update Yield KO");
            useCase.SelectBoxFirstUseCase();
            useCase.MultipleUpdate("Net weight", "14");
            var netWeight = useCase.GetColumnFirstValue(ItemUseCasePage.ColumnName.NetWeight);
            Assert.AreEqual("14", netWeight, "multi update Net Weight KO");
            useCase.SelectBoxFirstUseCase();
            useCase.MultipleUpdate("Net quantity", "14");
            var netQty = useCase.GetColumnFirstValue(ItemUseCasePage.ColumnName.NetQty);
            Assert.AreEqual("14", netQty, "multi update Net Quantity KO");
            useCase.SelectBoxFirstUseCase();
            useCase.MultipleUpdate("Gross quantity", "14");
            var grossQty = useCase.GetColumnFirstValue(ItemUseCasePage.ColumnName.GrossQty);
            Assert.AreEqual("14", grossQty, "multi update Gross Quantity KO");
            useCase.SelectBoxFirstUseCase();
            useCase.MultipleUpdate("Workshop", "Corte");
            var workshop = useCase.GetColumnFirstValue(ItemUseCasePage.ColumnName.Workshop);
            Assert.AreEqual("Corte", workshop, "multi update Workshop KO");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Item_UseCase_ModifyValues()
        {
            //Prepare

            //Arrange

            HomePage homePage = LogInAsAdmin();


            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            //Etre sur un item, onglet Use Case et avoir des Recipe Types disponibles
            string itemName = "Item for useCase";
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
            ItemUseCasePage useCase = generalInfo.ClickOnUseCasePage();
            //Modifier les valeurs sur une ligne et vérifier que tout ce recalcule correctement
            useCase.SelectFirstUseCase();

            useCase.SetColumnFirstValueForm(ItemUseCasePage.ColumnName.Yield, "100");
            useCase.SetColumnFirstValueForm(ItemUseCasePage.ColumnName.NetWeight, "10");
            useCase.SetColumnFirstValueForm(ItemUseCasePage.ColumnName.NetQty, "10");
            useCase.SetColumnFirstValueForm(ItemUseCasePage.ColumnName.GrossQty, "10");

            Assert.AreEqual("100", useCase.GetColumnFirstValueForm(ItemUseCasePage.ColumnName.Yield));
            Assert.AreEqual("10", useCase.GetColumnFirstValueForm(ItemUseCasePage.ColumnName.NetWeight));
            Assert.AreEqual("10", useCase.GetColumnFirstValueForm(ItemUseCasePage.ColumnName.NetQty));
            Assert.AreEqual("10", useCase.GetColumnFirstValueForm(ItemUseCasePage.ColumnName.GrossQty));

            //Vérifier les recalculs
            useCase.SetColumnFirstValueForm(ItemUseCasePage.ColumnName.Yield, "50");

            Assert.AreEqual("50", useCase.GetColumnFirstValueForm(ItemUseCasePage.ColumnName.Yield));
            Assert.AreEqual("10", useCase.GetColumnFirstValueForm(ItemUseCasePage.ColumnName.NetWeight));
            Assert.AreEqual("10", useCase.GetColumnFirstValueForm(ItemUseCasePage.ColumnName.NetQty));
            Assert.AreEqual("20", useCase.GetColumnFirstValueForm(ItemUseCasePage.ColumnName.GrossQty));

            useCase.SetColumnFirstValueForm(ItemUseCasePage.ColumnName.NetQty, "20");
            Assert.AreEqual("20", useCase.GetColumnFirstValueForm(ItemUseCasePage.ColumnName.NetWeight));
            Assert.AreEqual("40", useCase.GetColumnFirstValueForm(ItemUseCasePage.ColumnName.GrossQty));

            useCase.SetColumnFirstValueForm(ItemUseCasePage.ColumnName.GrossQty, "20");
            Assert.AreEqual("10", useCase.GetColumnFirstValueForm(ItemUseCasePage.ColumnName.NetWeight));
            Assert.AreEqual("10", useCase.GetColumnFirstValueForm(ItemUseCasePage.ColumnName.NetQty));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Item_UseCase_LinkRecipe()
        {
            // Prepare

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            string itemName = "Item for useCase";
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            //Etre sur un item, onglet Use Case et avoir des Recipe Types disponibles
            ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
            generalInfo.GetItemName();
            ItemUseCasePage useCase = generalInfo.ClickOnUseCasePage();
            //1. Cliquer sur l'icone Stylo d'une Recipe
            //2.Une fenetre s'ouvre vers la recpie
            useCase.SelectFirstUseCase();
            string recipeName = useCase.GetFirstRecipeName();
            bool isRecipe;
            RecipesVariantPage variantPage = useCase.ClickStyloFirstRecipe(out isRecipe);
            if (isRecipe)
            {
                //Vérifier que l'ouverture de la recipe correspond au même nom
                Assert.IsTrue(variantPage.GetRecipeName().Contains(recipeName.ToUpper()), "Mauvais nom Recipe");
            }
            else
            {
                //Datasheet
                var recipes = WebDriver.FindElements(By.XPath("//*/p[@class='datasheet-recipe-name']"));
                bool trouve = false;
                foreach (var recipe in recipes)
                {
                    if (recipeName.Contains(recipe.Text))
                    {
                        trouve = true;
                        break;
                    }
                }
                Assert.IsTrue(trouve, "Le Datasheet lié n'a pas ce recipe");
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Items_Storage_SelectAndUnselectAll()
        {    //Prepare

            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string packagingName1 = TestContext.Properties["Item_PackagingNameBis"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = "10";
            string qty = "10";
            string place = TestContext.Properties["InventoryPlace"].ToString();
            var adminName = TestContext.Properties["Admin_UserName"].ToString();
            var site = "SiteTestItems_Storage_SelectAndUnselectAll";
            string site1 = "SiteItems_Storage_SelectAndUnselectAll";
            string organizationNameFrom = "TestSitePlaceFrom";
            string organizationNameTo = "TestSitePlaceTo";
            string itemName = "Items_Storage_SelectAndUnselectAll";
            string userName = adminName.Substring(0, adminName.IndexOf("@"));
            string number = new Random().Next(1, 200).ToString();
            string siteZipCode = new Random().Next(10000, 99999).ToString();
            string siteAddress = number + " calle de " + GenerateName(10);
            var cities = new List<string> { "Barcelona", "Madrid", "San Sebastian", "Bilbao", "Cadiz", "Majorqua", "Albacete", "Reus", "Sevilla", "Granada", "Ourense", "Zaragoza" };
            int index = new Random().Next(cities.Count);
            string siteCity = cities[index];

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);

            if (itemPage.CheckTotalNumber() == 0)
            {
                var parametersSites = homePage.GoToParameters_Sites();
                var sitesModalPage = parametersSites.ClickOnNewSite();
                parametersSites = sitesModalPage.FillPrincipalField_CreationNewSite(site1, siteAddress, siteZipCode, siteCity);
                parametersSites.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, site1);
                var siteId1 = parametersSites.CollectNewSiteID();
                var parametersUser = homePage.GoToParameters_User();
                parametersUser.SearchAndSelectUser(userName);
                parametersUser.ClickOnAffectedSite();
                parametersUser.GiveSiteRightsToUser(siteId1, true, site1);

                homePage.GoToParameters_Sites();
                parametersSites.ClickOnNewSite();
                parametersSites = sitesModalPage.FillPrincipalField_CreationNewSite(site, siteAddress, siteZipCode, siteCity);
                parametersSites.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, site);
                var siteId = parametersSites.CollectNewSiteID();
                homePage.GoToParameters_User();
                parametersUser.SearchAndSelectUser(userName);
                parametersUser.ClickOnAffectedSite();
                parametersUser.GiveSiteRightsToUser(siteId, true, site);

                // Créer les organisations pour le site
                homePage.GoToParameters_Sites();
                parametersSites.Filter(ParametersSites.FilterType.SearchSite, site);
                parametersSites.WaitPageLoading();
                parametersSites.ClickOnFirstSite();
                parametersSites.ClickToOrganization();
                parametersSites.CreateNewOrganization(organizationNameFrom, "storage place");
                parametersSites.CreateNewOrganization(organizationNameTo, "storage place");

                // Créer les organisations pour le site
                homePage.GoToParameters_Sites();
                parametersSites.Filter(ParametersSites.FilterType.SearchSite, site1);
                parametersSites.WaitPageLoading();
                parametersSites.ClickOnFirstSite();
                parametersSites.ClickToOrganization();
                parametersSites.CreateNewOrganization(organizationNameFrom, "storage place");
                parametersSites.CreateNewOrganization(organizationNameTo, "storage place");
                // Créer un item
                homePage.GoToPurchasing_ItemPage();
                itemPage.ResetFilter();
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
                itemGeneralInformationPage.WaitPageLoading();
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier, qty);
                itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site1, packagingName1, storageQty, storageUnit, qty, supplier, qty);

                homePage.GoToPurchasing_ItemPage();
                // Créer deux  inventaires
                var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();
                var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(DateTime.Now.AddDays(2), site, organizationNameFrom);
                var inventoryItem = inventoryCreateModalpage.Submit();
                inventoryItem.Filter(InventoryItem.FilterItemType.SearchByName, itemName);
                inventoryItem.SelectFirstItem();
                inventoryItem.WaitPageLoading();
                inventoryItem.AddPhysicalQuantity(itemName, "10");
                inventoryItem.WaitPageLoading();
                InventoryGeneralInformation InventoryGeneralInformation = inventoryItem.ClickOnGeneralInformationTab();
                inventoryItem.WaitPageLoading();
                InventoryGeneralInformation.ClickOnItems();
                inventoryItem.Validate();
                inventoryItem.ConfirmValidation();
                homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();
                inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(DateTime.Now.AddDays(2), site1, organizationNameFrom);
                inventoryCreateModalpage.Submit();
                inventoryItem.Filter(InventoryItem.FilterItemType.SearchByName, itemName);
                inventoryItem.SelectFirstItem();
                inventoryItem.WaitPageLoading();
                inventoryItem.AddPhysicalQuantity(itemName, "10"); inventoryItem.WaitPageLoading();
                inventoryItem.WaitPageLoading();
                inventoryItem.ClickOnGeneralInformationTab();
                InventoryGeneralInformation.ClickOnItems();
                inventoryItem.Validate();
                inventoryItem.ConfirmValidation();
            }
            //Act
            //Etre sur le détail d'un item, onglet Storage et avoir au moins 2 items disponibles
            homePage.GoToPurchasing_ItemPage();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
            ItemStoragePage storage = generalInfo.ClickOnStorage();
            storage.ResetFilters();
            //Cliquer sur Select All et Unselect All
            //Les items de sélectionnent et se désélectionnent
            Assert.IsTrue(storage.SelectAll(), "le boutton selectAll n'est pas fonctionnel");
            Assert.IsTrue(storage.SelectNone(), "le boutton selectNone n'est pas fonctionnel");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Items_Storage_ResetFilter()
        {
            //Prepare
            string itemName = "Items_Storage_ResetFilter";
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
            var location = TestContext.Properties["Location"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                itemPage = itemGeneralInformationPage.BackToList();
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
            }
            var itemGeneralInformationsPage = itemPage.ClickOnFirstItem();
            var lastStorage = itemGeneralInformationsPage.ClickOnStorage();

            if (!lastStorage.IsLinkToStorageOK())
            {
                var purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();
                var createPurchaseOrderPage = purchaseOrdersPage.CreateNewPurchaseOrder();
                createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+10), true);
                var purchaseOrderItemPage = createPurchaseOrderPage.Submit();
                purchaseOrderItemPage.Filter(PurchaseOrderItem.FilterItemType.ByName, itemName);
                purchaseOrderItemPage.SelectFirstItemPo();
                purchaseOrderItemPage.AddQuantity("2");
                purchaseOrderItemPage.Validate();

            }

            homePage.GoToPurchasing_ItemPage();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
            ItemStoragePage storage = generalInfo.ClickOnStorage();
            storage.Filter(ItemStoragePage.FilterType.Site, "MAD - MAD");
            storage.Filter(ItemStoragePage.FilterType.StorageUnits, "KG");

            //Cliquer sur Reset Filter
            storage.ResetFilters();

            //Les filtres sont réinitialisé
            var checkSite = storage.CheckAllSites();
            var checkAllStorageUnits = storage.CheckAllStorageUnits();
            Assert.IsTrue(checkSite, "filtre Site non réinitialisé");
            Assert.IsTrue(checkAllStorageUnits, "filtre StorageUnits non réinitialisé");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Items_Storage_FilterSite()
        {
            //Arrange
            var homePage = LogInAsAdmin();
            homePage.Navigate();
            //Act
            //Etre sur le détail d'un item onglet Storage
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            ItemStoragePage storage = itemPage.ClickOnStorage();
            ItemGeneralInformationPage generalInformationPage = itemPage.ClickOnFirstItems();
            itemPage.ClickOnStorageItem();
            //Appliquer le filtre Site
            storage.Filter(ItemStoragePage.FilterType.Site, "MAD - MAD");
            var sitesList = storage.GetSitesListFiltred();
            foreach (var site in sitesList)
            {
                Assert.AreEqual("MAD", site, $"Le filtre des sites ne fonctionne pas correctement: {site}");
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Items_Storage_FilterStorageunits()
        {

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            //Etre sur le détail d'un item onglet Storage
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
            ItemStoragePage storage = generalInfo.ClickOnStorage();
            storage.ResetFilters();
            var storageNumbers = storage.CheckTotalNumber();
            if (storageNumbers > 0)
            {
                var firstStorageUnit = storage.GetFirstUnitFromLine();
                //Appliquer le filtre StorageUnits
                storage.Filter(ItemStoragePage.FilterType.StorageUnits, firstStorageUnit);
                //Check Number
                var storageUnitNumbers = storage.CheckTotalNumber();
                Assert.AreNotEqual(0, storageUnitNumbers, "Le filtre Storageunits ne fonctionne plus");
                Assert.IsTrue(storageNumbers >= storageUnitNumbers, "Les résultats ne sont pas filtrés");
                storage.ResetFilters();
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_ITEM_Items_Storage()
        {
            //Prepare
            string itemName = "LastItemReceipt" + DateUtils.Now.ToString("dd/MM/yyyy");

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            //Etre sur le détail d'un item
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            //1. Aller sur un item
            ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
            //2. Cliquer sur l'onglet Storage
            generalInfo.ClickOnStorage();
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Items_LastReceiptNotes()
        {
            //Prepare
            string itemName = "LastItemReceipt" + DateUtils.Now.ToString("dd/MM/yyyy");

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            //Etre sur le détail d'un item
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            //1. Aller sur un item
            ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
            //2. Cliquer sur l'onglet Last receipt notes	
            generalInfo.ClickOnLastReceiptNotesPage();
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Items_LastOutputForms()
        {
            //Prepare
            string itemName = "LastItemReceipt" + DateUtils.Now.ToString("dd/MM/yyyy");

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            //Etre sur le détail d'un item
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            //1. Aller sur un item
            ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
            //2. Cliquer sur l'onglet Last Output Forms
            generalInfo.ClickOnLastOutputForms();
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Items_LastOrdersNotReceived()
        {
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var itemPage = homePage.GoToPurchasing_ItemPage();
            var firstItem = itemPage.GetFirstItemName();
            itemPage.Filter(ItemPage.FilterType.Search, firstItem);
            ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
            generalInfo.ClickOnLastOrdersNotReceived();
            bool isTabOpned = generalInfo.IsLastOrdersNotReceivedTabOpened();
            Assert.IsTrue(isTabOpned, "L'onglet 'Last orders not received'est cliquable et s'ouvre correctement.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Item_FilterPackaging()
        {
            // Prepare
            var rdm = new Random().Next();
            string itemName = "LastItemReceiptPackaging" + DateUtils.Now.ToString("dd/MM/yyyy") + "_";
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string siteMAD = TestContext.Properties["Site"].ToString();
            string siteACE = TestContext.Properties["SiteACE"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = "10";
            string qty = "10";

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            // Aller sur la page des items et réinitialiser les filtres
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();

            // Vérifier si l'item existe déjà avant de le créer
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            if (itemPage.CheckTotalNumber() == 0)
            {
                // Créer un nouvel item si non présent
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);

                // Créer deux packagings pour cet item
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(siteMAD, packagingName, storageQty, storageUnit, qty, supplier);
                itemGeneralInformationPage.NewPackaging();
                itemCreatePackagingPage.FillField_CreateNewPackaging(siteACE, packagingName, storageQty, storageUnit, qty, supplier);
                // Retourner à la page des items
                itemPage = itemGeneralInformationPage.BackToList();
                itemPage.ResetFilter();
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
            }

            // Ouvrir l'item créé ou trouvé et vérifier les packagings
            ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();

            // Vérification du nombre total de packagings (devrait être 2)
            Assert.AreEqual(2, generalInfo.CountPackaging(), "Le nombre de packagings est incorrect.");

            generalInfo.SearchPackaging("MAD");

            // Vérifier que le nombre de packagings de dite MAD après filtrage est 1
            Assert.AreEqual(1, generalInfo.CountPackaging(), "Le nombre de packagings après filtrage est incorrect.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_BestPrice()
        {
            string price = 10.ToString();
            string itemName = "itemBestPrice _" + new Random().Next().ToString();
            string group = "A REFERENCIA";
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string packagingName = "KG";
            string storageUnit = "KG";
            string storageQty = 2.ToString();
            string qty = 100.ToString();
            string qty2 = 177.ToString();
            string supplierbestPrice = TestContext.Properties["SupplierInvoice"].ToString();
            string unitPrice = 177.ToString();
            string unitPrice1 = 100.ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["Supplier"].ToString();
            //Arrange
            HomePage homePage = LogInAsAdmin();

            try
            {
                ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
                ItemCreateModalPage itemCreateModalPage = itemPage.ItemCreatePage();
                ItemGeneralInformationPage itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
                // 1 packaging pour le site ACE
                ItemCreateNewPackagingModalPage itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier, "177", null, unitPrice);
                itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty2, supplierbestPrice, "100", null, unitPrice1);
                itemPage = homePage.GoToPurchasing_ItemPage();
                itemPage.ResetFilter();
                itemPage.Filter(ItemPage.FilterType.Site, site);
                BestPriceModal bestPrice = itemPage.MenuBestPrice();
                //2.Une pop'up s'ouvre, Renseigner Sites et GROUP
                bestPrice.Fill(site, null, group);
                //3.Cliquer sur Search
                bestPrice.ClickSearch();
                bool isSelected=bestPrice.SelectCheckBoxByItemName(itemName);
                Assert.IsTrue(isSelected, $"Checkbox for '{itemName}' was not selected as the main item.");
                bestPrice.ClickUpdate();
                bestPrice.ClickOnOKUpdate();
                WebDriver.Navigate().Refresh();
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                itemGeneralInformationPage = itemPage.ClickOnFirstItem();
                itemGeneralInformationPage.SearchPackaging(supplierbestPrice);
                bool isMainChecked = itemGeneralInformationPage.IsMainChecked(supplierbestPrice);
                Assert.IsTrue(isMainChecked, "Packaging avec supplier best price n'est pas le packaging principal");
            }
            finally
            {
                ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                ItemGeneralInformationPage itemGeneralInformationPageDetails = itemPage.ClickOnFirstItem();
                itemGeneralInformationPageDetails.ClickOnDeleteItem();
                itemGeneralInformationPageDetails.ConfirmDelete();
                itemGeneralInformationPageDetails.ClickOnDeleteItem();
                itemGeneralInformationPageDetails.ConfirmDelete();
                itemPage = itemGeneralInformationPageDetails.BackToList();
                //Delete first item 
                MassiveDeleteModal massiveDeleteModal = itemPage.MenuMassiveDelete();
                massiveDeleteModal.Filter(MassiveDeleteModal.FilterType.SearchByName, itemName);
                massiveDeleteModal.ClickSearch();
                var oldTotal = massiveDeleteModal.GetTotalRowsForPagination();
                massiveDeleteModal.ClickSelectAllButton();
                massiveDeleteModal.Delete();
                massiveDeleteModal = itemPage.MenuMassiveDelete();
                massiveDeleteModal.Filter(MassiveDeleteModal.FilterType.SearchByName, itemName);
                massiveDeleteModal.ClickSearch();
                var newTotal = massiveDeleteModal.GetTotalRowsForPagination();
                //verify the item is deleted
                Assert.IsTrue(newTotal == oldTotal - 1, "La suppression d'un item ne fonctionne pas.");
            }

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Detail_Keyword()
        {
            //Prepare

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            //Etre sur l'index des Purchasing items
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            // Etre sur le détail d'un item
            ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
            // Cliquer sur l'onglet Keyword
            ItemKeywordPage keyword = generalInfo.ClickOnKeywordItem();
            Assert.IsNotNull(keyword, "Pas d'onglet Keyword");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Detail_Dietetic()
        {
            //Prepare

            //Arrange  
            var homePage = LogInAsAdmin();

            //Act
            //Etre sur le détail d'un item / Onglet Dietetic
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
            //1.Cliquer sur l'onglet Dietetic
            ItemDieteticPage dietetic = generalInfo.ClickOnDieteticPage();
            //2.Modifier les Macro Nutriments
            string avant = dietetic.GetMacroNutriments("Alcohol");
            dietetic.SetMacroNutriments("Alcohol", "20");
            generalInfo = dietetic.ClickOnGeneralInfo();
            dietetic = generalInfo.ClickOnDieteticPage();
            string apres = dietetic.GetMacroNutriments("Alcohol");
            //Assert
            Assert.AreEqual("20", apres);
            dietetic.SetMacroNutriments("Alcohol", avant);
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_AddKeyword_ExportReceiptNote()
        {
            //Prepare
            //Avoir un site et un supplier ou les créer
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["Prodman_Needs_Item_Supplier1"].ToString();
            var rdm = new Random().Next();
            string itemName = "Item" + DateUtils.Now.ToString("yyyy-MM-dd") + "_" + rdm;
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();
            string keyword1 = "TEST_KEY";
            string deliveryPlace = TestContext.Properties["InventoryPlace"].ToString();
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            homePage.ClearDownloads();
            //1. Sur purchasing Item
            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
            //Créer un item
            ItemCreateModalPage itemCreate = itemPage.ItemCreatePage();
            ItemGeneralInformationPage generalInfo = itemCreate.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
            //Ajouter un packaging sur un site et un supplier
            ItemCreateNewPackagingModalPage pack = generalInfo.NewPackaging();
            generalInfo = pack.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);

            //Onglet Keyword : ajouter un keyword sur le haut ou sur la ligne (exemple plastic free ou local)
            ItemKeywordPage keywordPage = generalInfo.ClickOnKeywordItem();

            keywordPage.AddKeyword(keyword1);

            generalInfo = keywordPage.ClickOnGeneralInfo();

            ClearCache();

            //2. Aller sur Warehouse > Receipt note
            ReceiptNotesPage rn = generalInfo.GoToWarehouse_ReceiptNotesPage();
            //Créer une receipt note, sur le même site et supplier
            ReceiptNotesCreateModalPage createRN = rn.ReceiptNotesCreatePage();
            createRN.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, deliveryPlace));
            ReceiptNotesItem itemsRN = createRN.Submit();
            string RNName = itemsRN.GetReceiptNoteNumber();
            //Ajouter un quantity sur le Received de l'item
            itemsRN.Filter(ReceiptNotesItem.FilterItemType.SearchByName, itemName);
            itemsRN.Filter(ReceiptNotesItem.FilterItemType.ByGroup, group);
            itemsRN.Filter(ReceiptNotesItem.FilterItemType.SearchByKeyword, keyword1);
            itemsRN.SelectFirstItem();
            itemsRN.AddReceived(itemName, "2");
            //Valider les Checks
            var qualityChecks = itemsRN.ClickOnChecksTab();
            qualityChecks.DeliveryAccepted();

            if (qualityChecks.CanClickOnSecurityChecks())
            {
                qualityChecks.CanClickOnSecurityChecks();
                qualityChecks.SetSecurityChecks("No");
                qualityChecks.SetQualityChecks();
                itemsRN = qualityChecks.ClickOnReceiptNoteItemTab();
            }
            else
            {
                itemsRN = qualityChecks.ClickOnReceiptNoteItemTab();
            }

            //A venir : Vérifier dans l'onglet Keyword sur la RN avant validation
            itemsRN.Filter(ReceiptNotesItem.FilterItemType.SearchByKeyword, keyword1);
            itemsRN.Filter(ReceiptNotesItem.FilterItemType.ByGroup, group);

            //Valider la RN
            itemsRN.Validate();

            DirectoryInfo directory = new DirectoryInfo(downloadsPath);
            foreach (FileInfo fi in directory.GetFiles())
            {
                //purge dossier downloadsPath
                fi.Delete();
            }

            //Faire un export au niveau de l'index sur cette RN : vérifier le keyword dans la colonne Keyword
            rn = itemsRN.BackToList();
            rn.Filter(ReceiptNotesPage.FilterType.ByNumber, RNName);
            rn.Filter(ReceiptNotesPage.FilterType.Supplier, supplier);
            rn.Filter(ReceiptNotesPage.FilterType.Site, site);
            // un mois maximum
            rn.Filter(ReceiptNotesPage.FilterType.DateFrom, DateUtils.Now.AddDays(-20));
            rn.Filter(ReceiptNotesPage.FilterType.DateTo, DateUtils.Now.AddDays(1));
            rn.Filter(ReceiptNotesPage.FilterType.ShowAll, true);
            rn.ExportExcelFile(true);
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();
            FileInfo trouveExcel = rn.GetExportExcelFile(taskFiles);
            // Exploitation du fichier Excel
            List<String> colonneKeyword = OpenXmlExcel.GetValuesInList("Keywords", "Receipt notes", trouveExcel.FullName);
            //Assert
            Assert.AreNotEqual(0, colonneKeyword.Count, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.IsTrue(colonneKeyword.Contains(keyword1), keyword1 + " non trouvé dans la colonne Keyword");

            List<String> colonneNumber = OpenXmlExcel.GetValuesInList("Number", "Receipt notes", trouveExcel.FullName);
            Assert.IsTrue(colonneNumber.Contains(RNName), RNName + " non trouvé dans la colonne Number");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_AddIntolerance()
        {
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string recipeName1 = "Recette1" + new Random().Next().ToString();
            string recipeVariant = TestContext.Properties["RecipeVariant"].ToString();
            int nbPortions = new Random().Next(1, 30);
            string site = TestContext.Properties["Site"].ToString();
            HomePage homePage = LogInAsAdmin();
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            var itemGeneralInformations = itemPage.ClickOnFirstItem();
            var itemName = itemGeneralInformations.GetItemName();
            itemGeneralInformations.ClickOnIntolerancePage();
            var allName = itemGeneralInformations.CheckIntoleranceIfNotChecked();
            var recipesPage = homePage.GoToMenus_Recipes();
            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName1, recipeType, nbPortions.ToString());
            recipeGeneralInfosPage.AddVariantWithSite(site, recipeVariant);
            var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
            var intolPage = recipeVariantPage.ClickOnIntoleranceTab();
            var value = intolPage.IsRecipeAllergenChecked(allName);
            Assert.IsFalse(value, "La recette a l'intolérance");
            recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName1);
            recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
            recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();
            recipeVariantPage.AddIngredient(itemName);
            intolPage = recipeVariantPage.ClickOnIntoleranceTab();
            value = intolPage.IsRecipeAllergenChecked(allName);
            Assert.IsTrue(value, "La recette n'a pas l'intolérance");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Filter_BySubGroup()
        {

            //Prepare data
            string subGrpName = "subgrpname";
            //connect as admin
            var homePage = LogInAsAdmin();

            homePage.Navigate();

            //changing setting sub group to active
            homePage.SetSubGroupFunctionValue(true);
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.Filter(ItemPage.FilterType.SubGroup, subGrpName);
            var itemGeneralInformation = itemPage.ClickOnFirstItem();
            var name = itemGeneralInformation.getSubGroupeName();
            Assert.AreEqual(subGrpName, name, "erreur de filtrage par sub groupe");
            //desactivate sub group functionality 
            homePage.SetSubGroupFunctionValue(false);
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Filter_ShowStorageInAlertThreshold()
        {
            //prepare
            string itemName = "ItemTestAlertTreshold";
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            // Prepare items            
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();
            string limitQty = 10.ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();
            string inventoryItemPhyQty = 5.ToString();

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            //1)
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName.ToString());

            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);

                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                itemPage.ClickOnFirstPackagingAndSetLimitQty(limitQty);
                itemGeneralInformationPage.BackToList();

            }
            else //modifier l'inventaire à la quantité limit
            {
                var inventoryPage = homePage.GoToWarehouse_InventoriesPage();
                var inventoryCreateModalPage = inventoryPage.InventoryCreatePage();
                inventoryCreateModalPage.FillField_CreateNewInventory(DateUtils.Now, site, place);
                var itemInventory = inventoryCreateModalPage.Submit();
                itemInventory.Filter(InventoryItem.FilterItemType.SearchByName, itemName);
                itemInventory.SelectFirstItem();
                itemInventory.AddPhysicalQuantity(itemName, limitQty);
                InventoryValidationModalPage modal2 = itemInventory.Validate();
                modal2.ValidatePartialInventory();
            }
            //2)
            itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ClickOnStorage();
            itemPage.Filter(ItemPage.FilterType.ShowStorageInAlertThreshold, true);
            // Assert 
            itemPage.Filter(ItemPage.FilterType.Search, itemName.ToString());
            Assert.AreNotEqual(itemName, itemPage.GetFirstItemName(), "L'item {0} est présent dans la liste alors que son stockage est => à la limit quantity.", itemName);
            //3)
            var inventoryPageBis = homePage.GoToWarehouse_InventoriesPage();
            var inventoryCreateModalPageBis = inventoryPageBis.InventoryCreatePage();
            inventoryCreateModalPageBis.FillField_CreateNewInventory(DateUtils.Now, site, place);
            var itemInventoryBis = inventoryCreateModalPageBis.Submit();
            itemInventoryBis.Filter(InventoryItem.FilterItemType.SearchByName, itemName);
            itemInventoryBis.SelectFirstItem();
            itemInventoryBis.AddPhysicalQuantity(itemName, inventoryItemPhyQty);
            InventoryValidationModalPage modal2Bis = itemInventoryBis.Validate();
            modal2Bis.ValidatePartialInventory();

            // 4) 
            itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ClickOnStorage();
            itemPage.Filter(ItemPage.FilterType.ShowStorageInAlertThreshold, true);
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            var searchedItemName = itemPage.GetFirstItem();
            Assert.AreEqual(itemName, searchedItemName, "L'item {0} n'est pas présent dans la liste alors que son stockage est < à la limit quantity.", itemName);
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Item_UseCases_MultipleUpdate()
        {
            //connect as admin

            HomePage homePage = LogInAsAdmin();
            //go to items page 
            var itemsPage = homePage.GoToPurchasing_ItemPage();
            string itemName = "Item for useCase";
            itemsPage.ResetFilter();
            itemsPage.Filter(ItemPage.FilterType.Search, itemName);
            var itemGeneralInformationPage = itemsPage.ClickOnFirstItem();
            itemGeneralInformationPage.ClickOnUseCasePage();
            var useCasesNbr = itemGeneralInformationPage.GetNumberOfUseCases();
            Assert.IsTrue(useCasesNbr > 0, "l'item n'a pas de use cases");
            itemGeneralInformationPage.SelectAll();
            itemGeneralInformationPage.MultipleUpdate("Yield", "50");
            Assert.IsTrue(itemGeneralInformationPage.VerifyMultipleUpdate(Fields.Yield, "50"));
            itemGeneralInformationPage.MultipleUpdate("Net weight", "10");
            Assert.IsTrue(itemGeneralInformationPage.VerifyMultipleUpdate(Fields.NetWeight, "10"));
            itemGeneralInformationPage.MultipleUpdate("Net quantity", "20");
            Assert.IsTrue(itemGeneralInformationPage.VerifyMultipleUpdate(Fields.NetQuantity, "20"));
            itemGeneralInformationPage.MultipleUpdate("Gross quantity", "30");
            Assert.IsTrue(itemGeneralInformationPage.VerifyMultipleUpdate(Fields.GrossQuantity, "30"));
            itemGeneralInformationPage.MultipleUpdate("Workshop", "Corte");
            Assert.IsTrue(itemGeneralInformationPage.VerifyMultipleUpdate(Fields.Workshop, "Corte"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_FilterKeywords()
        {
            //Prepare
            string keyword = "TEST_KEY";

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            //Etre sur l'index des items et avoir des données
            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            int avant = itemPage.CheckTotalNumber();
            Assert.IsTrue(avant > 0, "Pas de donnée");
            //Filtrer par Keywords
            itemPage.Filter(ItemPage.FilterType.Keyword, keyword);
            //Vérifier que les résultats s'accordent bien au filtre appliqué
            int apres = itemPage.CheckTotalNumber();
            Assert.IsTrue(apres > 0, "Pas de donnée");
            Assert.IsTrue(avant >= apres, "Inconhérence");
            ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
            ItemKeywordPage keywordPage = generalInfo.ClickOnKeywordItem();

            // Assert
            bool IsKeywordPresent = keywordPage.IsKeywordPresent(keyword);
            Assert.IsTrue(IsKeywordPresent, "keyword non présent dans l'onglet keyboard");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_FilterSearch()
        {
            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            //Etre sur l'index des items et avoir des données
            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            int avant = itemPage.CheckTotalNumber();
            Assert.IsTrue(avant > 0, "Pas de donnée");
            //Filtrer SEARCH
            string firstItemName = itemPage.GetFirstItemName();
            itemPage.Filter(ItemPage.FilterType.Search, firstItemName);
            //Vérifier que les résultats s'accordent bien au filtre appliqué
            Assert.AreEqual(1, itemPage.CheckTotalNumber());
            Assert.AreEqual(firstItemName, itemPage.GetFirstItemName());
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_FilterSupplierRef()
        {
            string supplier = TestContext.Properties["Prodman_Needs_Item_Supplier1"].ToString();
            string supplierRef = TestContext.Properties["SupplierReferenceTest"].ToString();
            var homePage = LogInAsAdmin();
            homePage.Navigate();
            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            int avant = itemPage.CheckTotalNumber();
            Assert.IsTrue(avant > 0, "Pas de donnée");
            itemPage.Filter(ItemPage.FilterType.SupplierRef, supplierRef);
            var itemGeneralInformationPage = itemPage.ClickFirstItem();
            var result = itemGeneralInformationPage.VerifySupplierRef(supplierRef);
            Assert.IsTrue(result, "erreur de filtrage de supplier ref");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Detail_BestPrice()
        {
            //Prepare
            string site = TestContext.Properties["SiteLpCart"].ToString();
            string price = 10.ToString();
            string itemName = "itemBestPrice _" + new Random().Next().ToString();
            string group = "A REFERENCIA";
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string supplier = TestContext.Properties["Supplier"].ToString();
            string packagingName = "KG";
            string storageUnit = "KG";
            string storageQty = 2.ToString();
            string qty = 100.ToString();
            string qty2 = 177.ToString();
            string supplierbestPrice = TestContext.Properties["SupplierInvoice"].ToString();
            string unitPrice = 177.ToString();
            string unitPrice1 = 100.ToString();
            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            //Etre sur l'index des Purchasing items
            try
            {
                var itemPage = homePage.GoToPurchasing_ItemPage();
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
                // 1 packaging pour le site ACE
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier, "177", null, unitPrice);
                itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty2, supplierbestPrice, "100", null, unitPrice1);

                BestPriceModal bestPriceModal = itemGeneralInformationPage.ClickOnBestPrice();
                bestPriceModal.Fill(site);
                bestPriceModal.ClickSearch();
                string BestPriceSupplier = bestPriceModal.GetSupplierNameFromBestPriceTable(1);
                string BestPriceSite = bestPriceModal.GetSiteFromBestPriceTable(1);
                bestPriceModal.SelectFirstLine();
                bestPriceModal.ClickUpdate();
                bestPriceModal.ClickOnOKUpdate();
                Assert.AreEqual(supplierbestPrice, BestPriceSupplier, "Best Price Supplier ne correspond pas à la formule du best price.");
                Assert.AreEqual(site, BestPriceSite, "Best Price Site ne correspond pas à la formule du best price.");
                WebDriver.Navigate().Refresh();
                itemGeneralInformationPage.SearchPackaging(BestPriceSupplier);
                bool isMainChecked = itemGeneralInformationPage.IsMainChecked(BestPriceSupplier);
                Assert.IsTrue(isMainChecked, "Packaging avec supplier best price n'est pas le packaging principal");

            }
            finally
            {
                //delete first packaging of item 1 
                var itemPage = homePage.GoToPurchasing_ItemPage();
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                var itemGeneralInformationPageDetails = itemPage.ClickOnFirstItem();
                itemGeneralInformationPageDetails.ClickOnDeleteItem();
                itemGeneralInformationPageDetails.ConfirmDelete();
                itemGeneralInformationPageDetails.ClickOnDeleteItem();
                itemGeneralInformationPageDetails.ConfirmDelete();
                itemPage = itemGeneralInformationPageDetails.BackToList();
                //Delete first item 
                var massiveDeleteModal = itemPage.MenuMassiveDelete();
                massiveDeleteModal.Filter(MassiveDeleteModal.FilterType.SearchByName, itemName);
                massiveDeleteModal.ClickSearch();
                var oldTotal = massiveDeleteModal.GetTotalRowsForPagination();
                massiveDeleteModal.ClickSelectAllButton();
                massiveDeleteModal.Delete();
                massiveDeleteModal = itemPage.MenuMassiveDelete();
                massiveDeleteModal.Filter(MassiveDeleteModal.FilterType.SearchByName, itemName);
                massiveDeleteModal.ClickSearch();
                var newTotal = massiveDeleteModal.GetTotalRowsForPagination();
                //verify the item is deleted
                Assert.IsTrue(newTotal == oldTotal - 1, "La suppression d'un item ne fonctionne pas.");

            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Detail_DeactivatePackagingDetail()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string itemName = "packaging déactivé _" + new Random().Next().ToString();
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string supplier = TestContext.Properties["Prodman_Needs_Item_Supplier1"].ToString();
            string packagingName = "Caja";
            string storageUnit = "KG";
            string storageQty = 100.ToString();
            string qty = 1.ToString();
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            //Etre sur le détail d'un item / Onglet Dietetic
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();

            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);

            // 1 packaging pour le site ACE
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
            itemGeneralInformationPage.DeactivatePrice(site, supplier, packagingName);
            Assert.IsFalse(itemGeneralInformationPage.FindPackagingName(packagingName), "les packaging ne sont pas déactivé");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Item_UseCase_ModifyYieldDatasheet()
        {
            //Prepare
            Random rnd = new Random();
            string currency = TestContext.Properties["Currency"].ToString();
            string itemName = "Item for useCase";
            // double editedYeled = rnd.Next(1, 100);

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();
            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            //Etre sur un item, onglet Use Case et avoir des Recipe Types disponibles
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
            ItemUseCasePage useCase = generalInfo.ClickOnUseCasePage();
            var newYeiled = rnd.Next(1, 100);
            useCase.SelectFirstUseCase();
            string recipeName = useCase.GetFirstRecipeName();
            bool isdatasheet;
            var datasheetDetailPage = useCase.ClickOnUseCaseDatasheetDetailPage(out isdatasheet);
            var recipeVariantPage = datasheetDetailPage.EditRecipe();
            recipeVariantPage.ClickOnFirstIngredientDatasheet();
            var oldPrice = recipeVariantPage.GetPrice(decimalSeparatorValue, currency);
            datasheetDetailPage = recipeVariantPage.CloseWindow();
            datasheetDetailPage.WaitPageLoading();
            itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            generalInfo = itemPage.ClickOnFirstItem();
            useCase = generalInfo.ClickOnUseCasePage();
            useCase.SelectFirstUseCase();
            useCase.SetColumnFirstValueForm(ItemUseCasePage.ColumnName.Yield, newYeiled.ToString());
            datasheetDetailPage = useCase.ClickOnUseCaseDatasheetDetailPage(out isdatasheet);
            recipeVariantPage = datasheetDetailPage.EditRecipe();
            recipeVariantPage.ClickOnFirstIngredientDatasheet();
            var newPrice = recipeVariantPage.GetPrice(decimalSeparatorValue, currency);
            datasheetDetailPage = recipeVariantPage.CloseWindow();
            WebDriver.Navigate().Refresh();
            //Assert
            Assert.AreNotEqual(oldPrice, newPrice, "Le prix de la sous recette n'est pas actualisé .");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_SlowMovingItem()
        {
            //Prepare
            var adminName = TestContext.Properties["Admin_UserName"].ToString();
            var site = "SlowMoving" + new Random().Next(1000, 9999).ToString() + "TestSite";
            double inventoryValue = 0;
            double outputFormValue = 0;
            string userName = adminName.Substring(0, adminName.IndexOf("@"));
            string number = new Random().Next(1, 200).ToString();
            string siteZipCode = new Random().Next(10000, 99999).ToString();
            string siteAddress = number + " calle de " + GenerateName(10);
            var cities = new List<string> { "Barcelona", "Madrid", "San Sebastian", "Bilbao", "Cadiz", "Majorqua", "Albacete", "Reus", "Sevilla", "Granada", "Ourense", "Zaragoza" };
            int index = new Random().Next(cities.Count);
            string siteCity = cities[index];
            string organizationNameFrom = "SlowMovingTestSitePlaceFrom";
            string organizationNameTo = "SlowMovingTestSitePlaceTo";
            string itemName = "ITEM-" + GenerateName(4) + "-" + new Random().Next(1000, 9999).ToString();
            string itemName1 = "ITEM-" + GenerateName(4) + "-_" + new Random().Next(1000, 9999).ToString();
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string packagingName1 = TestContext.Properties["Item_PackagingNameBis"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = "10";
            string qty = "10";

            // Arrange

            HomePage homePage = LogInAsAdmin();

            var parametersSites = homePage.GoToParameters_Sites();

            // Vérification de l'existence du site avant de le créer
            parametersSites.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, site);
            string siteId;

            if (!parametersSites.IsSiteExisting(site))
            {
                // Si le site n'existe pas, le créer
                var sitesModalPage = parametersSites.ClickOnNewSite();
                parametersSites = sitesModalPage.FillPrincipalField_CreationNewSite(site, siteAddress, siteZipCode, siteCity);
                parametersSites.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, site);
                siteId = parametersSites.CollectNewSiteID();

                // Assigner le site à l'utilisateur
                var parametersUser = homePage.GoToParameters_User();
                parametersUser.SearchAndSelectUser(userName);
                parametersUser.ClickOnAffectedSite();
                parametersUser.GiveSiteRightsToUser(siteId, true, site);

                // Créer les organisations pour le site
                homePage.GoToParameters_Sites();
                parametersSites.Filter(ParametersSites.FilterType.SearchSite, site);
                parametersSites.ClickOnFirstSite();
                parametersSites.ClickToOrganization();
                parametersSites.CreateNewOrganization(organizationNameFrom, "storage place");
                parametersSites.CreateNewOrganization(organizationNameTo, "storage place");

                // Créer deux items
                var itemPage = homePage.GoToPurchasing_ItemPage();
                itemPage.ResetFilter();
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
                itemGeneralInformationPage.WaitPageLoading();
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier, qty);
                homePage.GoToPurchasing_ItemPage();
                itemPage.ItemCreatePage();
                itemCreateModalPage.FillField_CreateNewItem(itemName1, group, workshop, taxType, prodUnit);
                itemGeneralInformationPage.WaitPageLoading();
                itemGeneralInformationPage.NewPackaging();
                itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName1, storageQty, storageUnit, qty, supplier, qty);

                // Créer un inventaire
                var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();
                var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(DateTime.Now, site, organizationNameFrom);
                var inventoryItem = inventoryCreateModalpage.Submit();
                inventoryItem.Filter(InventoryItem.FilterItemType.SearchByName, itemName);
                inventoryItem.SelectFirstItem();
                inventoryItem.AddPhysicalQuantity(itemName, "10");
                InventoryGeneralInformation InventoryGeneralInformation = inventoryItem.ClickOnGeneralInformationTab();
                InventoryGeneralInformation.ClickOnItems();
                inventoryItem.Filter(InventoryItem.FilterItemType.SearchByName, itemName1);
                inventoryItem.SelectFirstItem();
                inventoryItem.AddPhysicalQuantity(itemName1, "10");
                inventoryValue = inventoryItem.GetInvenoryValue();
                inventoryItem.Validate();
                inventoryItem.ConfirmValidation();

                // Créer le formulaire de sortie
                var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
                var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
                outputFormCreateModalpage.FillField_CreatNewOutputForm(DateTime.Now, site, organizationNameFrom, organizationNameTo);
                var outputFormItem = outputFormCreateModalpage.Submit();
                outputFormItem.Filter(OutputFormItem.FilterItemType.SearchByName, itemName);
                outputFormItem.SelectFirstItem();
                outputFormItem.AddPhysicalQuantity(itemName, "10");
                outputFormValue = outputFormItem.GetPhysicalValue();
                outputFormItem.Validate();
            }
            //else
            //{
            //    var inventoryPage = homePage.GoToWarehouse_InventoriesPage();
            //    inventoryPage.ResetFilters();
            //    inventoryPage.Filter(InventoriesPage.FilterType.Site, site);
            //    var inventoryItemsPage = inventoryPage.ClickFirstInventory();
            //    inventoryValue = inventoryItemsPage.GetInvenoryValue();

            //    var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            //    outputFormPage.Filter(OutputFormPage.FilterType.CombosOptions, site, OutputFormPage.TypeOfShowGroup.NoShowGroupOption, OutputFormPage.TypeCombosOptions.Sites);
            //    var outputFormItemPage = outputFormPage.ClickOnFirstOF();
            //    outputFormValue = outputFormItemPage.GetPhysicalValue();

            //}
            // Calcul du pourcentage de slow-moving items
            var slowMovingPercentage = ((inventoryValue - outputFormValue) / inventoryValue) * 100;
            // Vérification du pourcentage dans le tableau de bord
            homePage.Navigate();
            var dashboardUsagePage = homePage.ClickDashboardUsage();
            dashboardUsagePage.ResetFilters();
            dashboardUsagePage.Filter(PageObjects.Admin.DashboardUsagePage.FilterType.Site, site);
            dashboardUsagePage.Filter(PageObjects.Admin.DashboardUsagePage.FilterType.From, DateUtils.Now.AddDays(-1));
            dashboardUsagePage.Filter(PageObjects.Admin.DashboardUsagePage.FilterType.To, DateUtils.Now.AddDays(1));
            var result = dashboardUsagePage.VerifySlowMovingPercentage(slowMovingPercentage);

            Assert.IsTrue(result, "Wrong slow-moving items percentage.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Detail_SetImpurchasableItems()
        {
            //Prepare
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
            string itemName = itemNameToday + "-" + new Random().Next().ToString();


            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();

            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);

            itemGeneralInformationPage.SetUnpurchasableItem(site, supplier, packagingName);
            WebDriver.Navigate().Refresh();

            // Assert 
            Assert.IsFalse(itemGeneralInformationPage.IsPurchasable(site, supplier), "La fonctionnalité 'SetUnpurchasable items' ne fonctionne pas.");


        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Massivedelete_Suppliersearch()
        {
            //Prepare

            string supplier = TestContext.Properties["Prodman_Needs_Item_Supplier1"].ToString();
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            MassiveDeleteModal massiveDelete = itemPage.MenuMassiveDelete();
            massiveDelete.Fill(supplier);
            //3.Cliquer sur Search
            massiveDelete.ClickSearch();

            // mémoriser le nom de l'item
            string supplierName = massiveDelete.GetSupplierName();
            //Assert 
            Assert.AreEqual(supplier, supplierName, "Le filter SUPPLIER ne remet pas");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Massivedelete_namesearch()
        {
            //Prepare
            string searchName = "095T BANDEJA 1/1";
            string supplier = TestContext.Properties["Supplier"].ToString();
            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            //Etre sur l'index des Purchasing items
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            MassiveDeleteModal massiveDelete = itemPage.MenuMassiveDelete();
            massiveDelete.Filter(MassiveDeleteModal.FilterType.SearchByName, searchName);
            //3.Cliquer sur Search
            massiveDelete.ClickSearch();

            //Assert
            // mémoriser le nom de l'item
            string itemName = massiveDelete.GetItemName();
            Assert.AreEqual(searchName, itemName, "Seulement les Item correspondants au text search doivent apparaitre.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Massivedelete_InactiveSupplierCheck()
        {
            //Prepare
            string filter = "Inactive";

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            MassiveDeleteModal massiveDelete = itemPage.MenuMassiveDelete();
            massiveDelete.Filter(MassiveDeleteModal.FilterType.ShowInactiveSuppliers, true);
            massiveDelete.Filter(MassiveDeleteModal.FilterType.Supplier, filter);
            var ordersList = massiveDelete.GetSuppliersListFiltred();

            //Assert
            foreach (var order in ordersList)
            {
                Assert.IsTrue(order.StartsWith("Inactive -"), $"Order does not start with 'Inactive -': {order}");
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Massivedelete_pagination()
        {
            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();

            MassiveDeleteModal massiveDeleteModal = itemPage.MenuMassiveDelete();
            massiveDeleteModal.ClickSearch();
            massiveDeleteModal.ClickSelectAllButton();

            //pagination processing
            int tot = massiveDeleteModal.CheckTotalSelectCount();

            massiveDeleteModal.PageSize("16");

            //Assert
            Assert.IsTrue(massiveDeleteModal.IsPageSizeEqualsTo("16"), "Pagination ne fonctionne pas.");
            Assert.AreEqual(massiveDeleteModal.GetTotalRowsForPagination(), tot >= 16 ? 16 : tot, "Pagination ne fonctionne pas.");

            massiveDeleteModal.PageSize("30");

            //Assert
            Assert.IsTrue(massiveDeleteModal.IsPageSizeEqualsTo("30"), "Pagination ne fonctionne pas.");
            Assert.AreEqual(massiveDeleteModal.GetTotalRowsForPagination(), tot >= 30 ? 30 : tot, "Pagination ne fonctionne pas.");

            massiveDeleteModal.PageSize("50");

            //Assert
            Assert.IsTrue(massiveDeleteModal.IsPageSizeEqualsTo("50"), "Pagination ne fonctionne pas.");
            Assert.AreEqual(massiveDeleteModal.GetTotalRowsForPagination(), tot >= 50 ? 50 : tot, "Pagination ne fonctionne pas.");

            massiveDeleteModal.PageSize("100");

            //Assert
            Assert.IsTrue(massiveDeleteModal.IsPageSizeEqualsTo("100"), "Pagination ne fonctionne pas.");
            Assert.AreEqual(massiveDeleteModal.GetTotalRowsForPagination(), tot >= 100 ? 100 : tot, "Pagination ne fonctionne pas.");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Massivedelete_sitesearch()
        {
            //Prepare
            string site1 = "ACE";

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            var massiveDeletePage = itemPage.MenuMassiveDelete();
            massiveDeletePage.Filter(MassiveDeleteModal.FilterType.Site, site1);
            massiveDeletePage.ClickSearch();
            massiveDeletePage.SetPageSize("50");

            //Assert
            var filteredSite = massiveDeletePage.GetAllSites();
            Assert.IsTrue(filteredSite.All(sites => site1 == sites), String.Format(MessageErreur.FILTRE_ERRONE, "Seuls les Item du site sélectionné doivent ressortir."));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Massivedelete_organizebysite()
        {
            //Prepare
            string site = TestContext.Properties["SiteACE"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            MassiveDeleteModal massiveDelete = itemPage.MenuMassiveDelete();
            massiveDelete.Filter(MassiveDeleteModal.FilterType.Site, site);

            massiveDelete.ClickSearch();
            massiveDelete.ClickSite();
            List<string> itemsResults = massiveDelete.GetPurchasingItemSite();

            //Assert
            Assert.IsTrue(itemsResults.SequenceEqual(itemsResults.OrderBy(nameSite => nameSite)) || itemsResults.SequenceEqual(itemsResults.OrderByDescending(nameSite => nameSite)), "L'organisation par site n'est pas en ordre alphabétique");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Massivedelete_inactivesitescheck()
        {
            //Arrange
            var homePage = LogInAsAdmin();
            homePage.Navigate();
            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            MassiveDeleteModal massiveDeleteModal = itemPage.MenuMassiveDelete();
            massiveDeleteModal.Filter(MassiveDeleteModal.FilterType.ShowInactiveSites, true);
            massiveDeleteModal.Filter(MassiveDeleteModal.FilterType.Site, "Inactive");
            var ordersList = massiveDeleteModal.GetSitesListFiltred();
            foreach (var order in ordersList)
            {
                Assert.IsTrue(order.StartsWith("Inactive -"), $"Les filtres ne fonctionnent pas et les sites ne commencent pas par 'Inactive': {order}");
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Massivedelete_OrganizeByStatus()
        {
            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            MassiveDeleteModal massiveDelete = itemPage.MenuMassiveDelete();
            massiveDelete.ClickSearch();
            massiveDelete.ClickStatus();
            List<string> itemsResults = massiveDelete.GetPurchasingItemStatus();

            //Assert
            Assert.IsTrue(itemsResults.SequenceEqual(itemsResults.OrderBy(status => status)) || itemsResults.SequenceEqual(itemsResults.OrderByDescending(status => status)), "L'organisation par status n'est pas en ordre alphabétique");


        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Massivedelete_organizebyItemName()
        {
            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            var massiveDeletePage = itemPage.MenuMassiveDelete();
            var isAscendingSorted = massiveDeletePage.IsSortedAlphabetically(ascending: true);

            massiveDeletePage.SortByItemBtn();

            //Assert
            var isDescendingSorted = massiveDeletePage.IsSortedAlphabetically(ascending: false);
            Assert.IsTrue(isAscendingSorted, "Items are not sorted in ascending alphabetical order.");
            Assert.IsTrue(isDescendingSorted, "Items are not sorted in descending alphabetical order.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Massivedelete_select()
        {
            //Prepare
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = "ACE";
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();
            string itemName = "AUTO_TEST_ITEM_" + DateTime.Now.ToString("dd_MM_yyyy-") + new Random().Next().ToString();
            string itemName1 = itemName + "_1";
            string itemName2 = itemName + "_2";

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();

            // Create 2 items to test Massive Delete select
            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName1.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
            itemPage = itemGeneralInformationPage.BackToList();

            itemCreateModalPage = itemPage.ItemCreatePage();
            itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName2.ToString(), group, workshop, taxType, prodUnit);
            itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
            itemPage = itemGeneralInformationPage.BackToList();

            //Find 2 ITEMS created
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            var itemsNumber = itemPage.CheckTotalNumber();
            if (itemsNumber == 2)
            {
                //delete first packaging of item 1 
                itemPage.Filter(ItemPage.FilterType.Search, itemName1);
                var itemGeneralInformationPageDetails = itemPage.ClickOnFirstItem();
                itemGeneralInformationPageDetails.ClickOnDeleteItem();
                itemGeneralInformationPageDetails.ConfirmDelete();
                itemPage = itemGeneralInformationPageDetails.BackToList();

                //delete first packaging of item 2 
                itemPage.Filter(ItemPage.FilterType.Search, itemName2);
                itemGeneralInformationPageDetails = itemPage.ClickOnFirstItem();
                itemGeneralInformationPageDetails.ClickOnDeleteItem();
                itemGeneralInformationPageDetails.ConfirmDelete();
                itemPage = itemGeneralInformationPageDetails.BackToList();

                itemPage.ResetFilter();

                //Delete first item 
                var massiveDeleteModal = itemPage.MenuMassiveDelete();
                massiveDeleteModal.Filter(MassiveDeleteModal.FilterType.SearchByName, itemName);
                massiveDeleteModal.ClickSearch();
                var oldTotal = massiveDeleteModal.GetTotalRowsForPagination();
                massiveDeleteModal.DeleteFirstService();
                massiveDeleteModal = itemPage.MenuMassiveDelete();
                massiveDeleteModal.Filter(MassiveDeleteModal.FilterType.SearchByName, itemName);
                massiveDeleteModal.ClickSearch();
                var newTotal = massiveDeleteModal.GetTotalRowsForPagination();
                //verify the item is deleted
                Assert.IsTrue(newTotal == oldTotal - 1, "La suppression d'un item ne fonctionne pas.");
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Massivedelete_inactivesitessearch()
        {
            //Prepare
            string site = TestContext.Properties["SiteACE"].ToString();
            string siteInactive = "Inactive - ANES";

            //arrange
            var homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.ShowOnlyInactive, true);
            itemPage.WaitPageLoading();
            MassiveDeleteModal massiveDeleteModal = itemPage.MenuMassiveDelete();
            massiveDeleteModal.ClickOnShowInactiveSiteCheckbox();
            massiveDeleteModal.SelectMultipleInactiveSites(site, siteInactive);
            massiveDeleteModal.ClickSearch();
            massiveDeleteModal.ClickSelectAllButton();
            var totalNumberBeforeDelete = massiveDeleteModal.CheckTotalSelectCount();
            massiveDeleteModal.ClickUnselectAllButton();
            massiveDeleteModal.UncheckAllSites();
            massiveDeleteModal.WaitPageLoading();
            massiveDeleteModal.SelectInactiveSites(siteInactive);
            massiveDeleteModal.ClickSearch();
            massiveDeleteModal.ClickSelectAllButton();
            var totalNumberAfterDelete = massiveDeleteModal.CheckTotalSelectCount();
            //Assert
            Assert.IsTrue(totalNumberAfterDelete <= totalNumberBeforeDelete, "Seuls les Item des sites inactifs sélectionnés doivent ressortir.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Massivedelete()
        {
            //prepare
            string site = TestContext.Properties["Site"].ToString();
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = "10";
            string qty = "10";
            string itemName = "ItemToDelete_" + DateTime.Now.ToString("dd_MM_yyyy-") + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);

            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                itemPage = itemGeneralInformationPage.BackToList();
                itemPage.ResetFilter();
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
            }
            ItemGeneralInformationPage itemGeneralInformation = itemPage.ClickOnFirstItem();

            if (itemGeneralInformation.CountPackaging() == 0)
            {
                var itemCreatePackagingPage = itemGeneralInformation.NewPackaging();
                itemGeneralInformation = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
            }
            int itemPackagingNumberBeforeDelete = itemGeneralInformation.CountPackaging();
            itemGeneralInformation.ClickOnDeleteItem();
            itemGeneralInformation.ConfirmDelete();
            itemPage = itemGeneralInformation.BackToList();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.ShowAll, true);
            MassiveDeleteModal massiveDeleteModal = itemPage.MenuMassiveDelete();
            massiveDeleteModal.Filter(MassiveDeleteModal.FilterType.SearchByName, itemName);
            massiveDeleteModal.SelectFirstItem(site);
            massiveDeleteModal.ClickSearch();
            massiveDeleteModal.ClickSelectAllButton();
            massiveDeleteModal.Delete();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.ShowAll, true);
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            itemGeneralInformation = itemPage.ClickOnFirstItem();
            int itemPackagingNumberAfterDelete = itemGeneralInformation.CountPackaging();
            //This test focused only on deleting item packaging until 'Show items without price' was reactivated.
            Assert.IsTrue(itemPackagingNumberAfterDelete < itemPackagingNumberBeforeDelete, "La suppression massive ne fonctionnent pas!");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Massivedelete_selectall()
        {

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            MassiveDeleteModal massiveDeleteModal = itemPage.MenuMassiveDelete();
            massiveDeleteModal.ClickSearch();

            // Select All Processing
            var allSelectResult = massiveDeleteModal.VerifyRowsAreSelected();
            var totalNumber = massiveDeleteModal.CheckTotalNumber();

            //Assert
            Assert.IsTrue(totalNumber > 0, "L'affichage du nombre de Selected datasheet ne s'est pas mis à jour");
            Assert.IsTrue(allSelectResult, "Erreur de Select All,Le nombre de lignes cochées et l'affichage du nombre de sélections ne sont pas égaux.");


        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Massivedelete_link()
        {
            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            var massiveDeletePage = itemPage.MenuMassiveDelete();
            massiveDeletePage.SelectAllSites();
            massiveDeletePage.SelectAllSupplier();
            massiveDeletePage.ClickSearch();
            var Item = massiveDeletePage.GetFirstItem();
            var ItemGeneralInformationPage = massiveDeletePage.ClickOnItemLinkFromRow(1);
            string ItemName = ItemGeneralInformationPage.GetItemName().Trim();
            bool isNewTabAdded = massiveDeletePage.GetTotalTabs();

            // Assert
            Assert.IsTrue(isNewTabAdded, "Aucun Onglet n'a été ouvert");
            Assert.AreEqual(Item, ItemName, "La page detail item n'est pas visible.");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Massivedelete_OrganizeBySupplier()
        {
            string supplier = TestContext.Properties["Prodman_Needs_Item_Supplier1"].ToString();
            var homePage = LogInAsAdmin();
            homePage.Navigate();
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            MassiveDeleteModal massiveDelete = itemPage.MenuMassiveDelete();
            massiveDelete.WaitPageLoading();
            massiveDelete.Filter(MassiveDeleteModal.FilterType.Supplier, supplier);
            massiveDelete.ClickSearch();
            massiveDelete.ClickSupplier();
            List<string> itemsResults = massiveDelete.GetPurchasingItemSupplier();
            var result = itemsResults.SequenceEqual(itemsResults.OrderBy(nameSupplier => nameSupplier)) || itemsResults.SequenceEqual(itemsResults.OrderByDescending(nameSupplier => nameSupplier));
            Assert.IsTrue(result, "L'organisation par supplier n'est pas en ordre alphabétique");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Massivedelete_InactiveSupplierSearch()
        {
            //prepare
            Random rnd = new Random();
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplierTest = "test-supplier-" + rnd.Next(1000, 9999).ToString();
            string storageUnit = "KG";
            string storageQty = "10";
            string qty = "10";
            string itemName = itemNameToday + "-" + rnd.Next(1000, 9999).ToString();
            //arrange
            var homePage = LogInAsAdmin();

            //Act
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, supplierTest);

            if (suppliersPage.CheckTotalNumber() == 0)
            {
                var supplierCreateModalpage = suppliersPage.SupplierCreatePage();
                supplierCreateModalpage.FillField_CreatNewSupplier(supplierTest, true, false);
                var supplierItem = supplierCreateModalpage.Submit();
                supplierItem.BackToList();
            }

            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();

            // If no item exists, create a new item with packaging
            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGenerlInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGenerlInformationPage.NewPackaging();
                itemGenerlInformationPage = itemCreatePackagingPage.FillField_CreateNewPackagingWithoutPrice(site, packagingName, storageQty, storageUnit, qty, supplierTest);
                itemPage = itemGenerlInformationPage.BackToList();
            }

            //Desactiver supplier
            suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, supplierTest);
            var supplierGeneralInfo = suppliersPage.SelectFirstSupplier();
            supplierGeneralInfo.SetSupplierInactive();

            homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();

            MassiveDeleteModal massiveDelete = itemPage.MenuMassiveDelete();
            massiveDelete.Filter(MassiveDeleteModal.FilterType.SearchByName, itemName);
            massiveDelete.Filter(MassiveDeleteModal.FilterType.ShowInactiveSuppliers, true);
            massiveDelete.Filter(MassiveDeleteModal.FilterType.SupplierMultiple, supplierTest);

            massiveDelete.ClickSearch();
            List<string> itemsResults = massiveDelete.GetPurchasingItemIDStatus();
            foreach (var item in itemsResults)
            {
                Assert.IsTrue(item == "Inactive", $" les Supplier ne sont pas inactive");
            }
            massiveDelete.ClickSelectAllButton();
            massiveDelete.Delete();
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_ITEMS_Doublon()
        {
            //Prepare
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string siteACE = TestContext.Properties["SiteACE"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = 2.32.ToString();
            string qty = 2.32.ToString();
            string itemName = "ITEM_Doublon-" + new Random().Next().ToString();

            // Login and go to page Home
            var homePage = LogInAsAdmin();

            // Go to Purchasing Item Page
            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();

            try
            {
                // Create New Item and Packaging 
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(siteACE, packagingName, storageQty, storageUnit, qty, supplier);

                // Back to list and filter with item Created
                itemPage = itemGeneralInformationPage.BackToList();
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                itemPage.ResetFilter();

                /* Click on new item */
                ItemCreateModalPage itemCreatePage = itemPage.ItemCreatePage();

                /* Fill the name case with the fisr name on on the list */
                itemCreatePage.FillField_Name(itemName);

                /* Verify the error message */
                bool is_error_message_double_name_exists = itemCreatePage.ErrorMessageNameDouble();

                /* Assert */
                Assert.IsTrue(is_error_message_double_name_exists, "The message error don't show up");
            }
            finally
            {
                // Go to Purchasing Item Page
                homePage.GoToPurchasing_ItemPage();
                itemPage.ResetFilter();

                //Desactivate packaging
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
                generalInfo.DeactivatePrice(siteACE, supplier, packagingName);

                //Delete Item
                generalInfo.BackToList();
                MassiveDeleteModal massiveDelete = itemPage.MenuMassiveDelete();
                massiveDelete.Filter(MassiveDeleteModal.FilterType.SearchByName, itemName);
                massiveDelete.Filter(MassiveDeleteModal.FilterType.ShowInactiveSuppliers, true);
                massiveDelete.Filter(MassiveDeleteModal.FilterType.SupplierMultiple, supplier);
                massiveDelete.ClickSearch();
                massiveDelete.ClickSelectAllButton();
                massiveDelete.Delete();
            }

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Edit_Item()
        {
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string prodUnitToUpdate = "L";
            string itemName = itemNameToday + "-" + new Random().Next().ToString();
            LogInAsAdmin();
            HomePage homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.ShowAll, true);
            ItemCreateModalPage itemCreateModalPage = itemPage.ItemCreatePage();
            ItemGeneralInformationPage itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
            itemPage = itemGeneralInformationPage.BackToList();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            itemGeneralInformationPage = itemPage.ClickOnFirstItem();
            itemGeneralInformationPage.UpdateItem(itemName, group, workshop, taxType, prodUnitToUpdate);
            itemPage = itemGeneralInformationPage.BackToList();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.ShowAll, true);
            itemPage.Filter(ItemPage.FilterType.Search, itemName.ToString());
            string newProdUnit = itemPage.GetFirstItemUnit();
            Assert.AreEqual(newProdUnit, prodUnitToUpdate, "La modification n'a pas été effectuer !");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Label_Validation()
        {
            //Prepare
            //Prepare
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
            string itemName = itemNameToday + "-" + new Random().Next().ToString();
            //Arrange

            LogInAsAdmin();

            HomePage homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();

            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);

            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
            itemPage = itemGeneralInformationPage.BackToList();
            itemPage.ResetFilter();
            itemPage.ClickOnPurchaseTab();
            itemPage.Filter(ItemPage.FilterType.Search, itemName.ToString());
            itemPage.ClickOnFirstItem();
            itemPage.ClickOnLabelTab();
            string expectedComment = "TESTTEST";
            itemPage.EnterComment(expectedComment);
            itemPage.ClickOnValidateButton();
            itemPage.ClickOnUseCase();
            itemPage.ClickOnLabelTab();

            string actualComment = itemPage.GetCommentText();
            Assert.AreEqual(expectedComment, actualComment, "Le commentaire n'a pas été enregistré correctement.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_purchasebook_Column_Action_Duplicate()
        {
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
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            HomePage homePage = LogInAsAdmin();
            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemNameToday);
            if (itemPage.CheckTotalNumber() == 0)
            {
                ItemCreateModalPage itemCreateModalPage = itemPage.ItemCreatePage();
                ItemGeneralInformationPage itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemNameToday, group, workshop, taxType, prodUnit, "200");
                ItemCreateNewPackagingModalPage itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                itemPage = itemGeneralInformationPage.BackToList();
                itemPage.ResetFilter();
                itemPage.Filter(ItemPage.FilterType.Search, itemNameToday);
            }
            itemPage.ClearDownloads();
            foreach (FileInfo file in new DirectoryInfo(downloadsPath).GetFiles()) file.Delete();
            itemPage.Export(true);
            //itemPage.DownloadFirstFile();
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();
            FileInfo correctDownloadedFile = itemPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "fichier non généré");
            string fileName = correctDownloadedFile.Name;
            string filePath = System.IO.Path.Combine(downloadsPath, fileName);
            List<string> header = OpenXmlExcel.ReadFirstRow("Sheet 1", filePath);
            int actionColumnCount = header.Count(column => column.Equals("Action", StringComparison.OrdinalIgnoreCase));
            Assert.IsTrue(actionColumnCount == 1, $"Échec : Le fichier Excel doit contenir exactement une colonne nommée 'Action'. Nombre trouvé : {actionColumnCount}");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Search_Recpe_Datasheet()
        {

            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string itemForUseCase = "Item for useCase";
            string recipeForUseCase2 = "Recipe for Usecase 2";
            int nbPortions = new Random().Next(1, 30);
            DateTime now = DateTime.Now;
            string formattedDateTime = now.ToString("yyyy-MM-dd_HH-mm-ss");
            string datasheetForUseCase = $"Datasheet_for_Testing_{formattedDateTime}";
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string recipeForUseCase1 = $"Recipe_{formattedDateTime}";
            string serviceName = "Service for Usecase";
            string serviceCategorie = TestContext.Properties["Production_Service1"].ToString();
            string serviceType = TestContext.Properties["Production_ServiceType"].ToString();
            string itemName = itemNameToday + "-" + new Random().Next().ToString();
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var datasheetPage = homePage.GoToMenus_Datasheet();

            var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
            var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetForUseCase, guestType, site);
            var recipesCreateModalPage = datasheetDetailPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillFields_AddNewRecipeToDatasheetCommercial2(recipeForUseCase1, recipeForUseCase1, "b");
            datasheetDetailPage.Firstreciepclick();
            datasheetDetailPage.EditButton();
            string ingredientname = datasheetDetailPage.getingredientname();
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, ingredientname);
            itemPage.Filter(ItemPage.FilterType.ShowAll, true);

            itemPage.ClickOnFirstItem();
            itemPage.ClickOnUseCase();
            itemPage.Filter(ItemPage.FilterType.Search, recipeForUseCase1);

            string reciepe = itemPage.GetreciepnametName();

            Assert.AreEqual(recipeForUseCase1, reciepe);

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_WeightInGram()
        {
            //Prepare
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
            int weightInitial = 200;
            string itemName = "itemWeightInGram";

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit, "200");
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                itemPage = itemGeneralInformationPage.BackToList();
                itemPage.ResetFilter();
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
            }

            ItemGeneralInformationPage GeneralInformation = itemPage.ClickOnFirstItem();
            string weight = GeneralInformation.GetWeightInGram();

            Assert.AreEqual(weightInitial, int.Parse(weight), "La case Weigh in gram n'affiche pas correctement");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Best_price()
        {
            //Prepare
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string itemName = itemNameToday + "-" + new Random().Next().ToString();
            string packagingNotPurchasableName = TestContext.Properties["Item_PackagingName"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();
            string supplierRef = "coucou";
            string packagingPurchasableName = "Bandeja";
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["Prodman_Needs_Item_Supplier1"].ToString();
            string limitQty = 10.ToString();
            bool roundingToSu = false;
            int firstLine = 1;

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit, "200");
            ItemCreateNewPackagingModalPage packNotPurchasable = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = packNotPurchasable.FillField_CreateNewPackaging(site, packagingNotPurchasableName, storageQty, storageUnit, qty, supplier, limitQty, supplierRef, "10", roundingToSu);
            itemGeneralInformationPage.SelectFirstPackaging();
            itemGeneralInformationPage.SetPurchasable(false);
            ItemCreateNewPackagingModalPage packPurchasable = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = packPurchasable.FillField_CreateNewPackaging(site, packagingPurchasableName, storageQty, storageUnit, qty, supplier, limitQty, supplierRef, "10", roundingToSu);
            itemGeneralInformationPage.SelectFirstPackaging();
            BestPriceModal bestPriceModal = itemGeneralInformationPage.ClickOnBestPrice();
            bestPriceModal.Fill(site, supplier, group);
            bestPriceModal.ClickSearch();
            Assert.IsTrue(bestPriceModal.GetSiteFromBestPriceTable(1) == site && bestPriceModal.GetPackFromBestPriceTable(1).Contains(packagingPurchasableName) && bestPriceModal.GetItemNameFromBestPriceTable(1) == itemName && bestPriceModal.GetSupplierNameFromBestPriceTable(1) == supplier, "la recherche ne s'affiche pas correctement");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Filter_By_keywords()
        {
            //Prepare
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string supplierRef = "coucou";
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();

            string itemName = itemNameToday + "-" + new Random().Next(100, 999).ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //act
            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            string Keyword = TestContext.Properties["Item_Keyword"].ToString();
            // count 
            var itemCreateModalPage = itemPage.ItemCreatePage();
            ItemGeneralInformationPage itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
            ItemKeywordPage itemKeywordPage = itemGeneralInformationPage.ClickOnKeywordItem();
            itemKeywordPage.AddKeyword(Keyword);
            itemPage = itemGeneralInformationPage.BackToList();
            itemPage.Filter(ItemPage.FilterType.Keyword, Keyword.ToString());
            itemPage.Filter(ItemPage.FilterType.Search, itemName.ToString());

            // Assert 
            int numberSearch = itemPage.CheckTotalNumber();
            Assert.AreEqual(numberSearch, 1, MessageErreur.FILTRE_ERRONE, "Le Filter By keywords ne fonctionne pas.");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_UseCase_replace()
        {
            //prepare
            string recipeNameToday = "Rec-" + DateUtils.Now.ToString("dd/MM/yyyy");
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
            string itemName = itemNameToday + "-" + new Random().Next(100, 999).ToString() + DateTime.Now;
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            try
            {
                //create New Item 
                var itemCreateModalPage = itemPage.ItemCreatePage();
                ItemGeneralInformationPage itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);

                //Create New Receipe
                var recipesPage = homePage.GoToMenus_Recipes();
                recipesPage.ResetFilter();
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());
                recipeGeneralInfosPage.AddVariantWithSite(site, variant);
                if (!recipeGeneralInfosPage.IsFirstVariantIsVisible())
                {
                    recipeGeneralInfosPage.Validate();
                }
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
                recipeVariantPage.AddIngredient(itemName);
                recipeVariantPage.BackToList();

                itemPage = homePage.GoToPurchasing_ItemPage();
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                itemPage.ClickOnFirstItem();
                itemPage.ClickOnUseCase();
                itemPage.SelectALL();
                itemPage.ReplaceWithAnOtherItem();
                itemPage.InputReplaceWithAnOtherItem("a");
                itemPage.clickOnItemCheck();
                //itemPage.FirstLineClick();
                itemPage.Replaceitem();
                bool result = itemPage.ConfirmeButton();
                //itemPage.WaitPageLoading();
                //bool result = itemPage.ReplacedSuccess();

                //Assert
                Assert.IsTrue(result, "NOT REPLACED SUCCESSFULY");

            }
            finally
            {
                //Delete Receipe
                var recipesPage = homePage.GoToMenus_Recipes();
                if (!string.IsNullOrEmpty(recipeName))
                {
                    recipesPage.MassiveDeleteRecipes(recipeName, site, recipeType);
                }
            }

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Right_to_read()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            var itemCreateModalPage = itemPage.ClickOnFirstItem();
            var ItemName = itemCreateModalPage.GetItemName();
            Assert.IsTrue(itemCreateModalPage.GetItemName() != null, "Reading worked fine");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_SavePictureItm()
        {
            string group = "A REFERENCIA";
            var path = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase);
            path = path.Substring(6);
            var imagePath = path + "\\PageObjects\\Purchasing\\Item\\test.jpg";
            Assert.IsTrue(new FileInfo(imagePath).Exists, "Fichier " + imagePath + " non trouvé");

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var itemPage = homePage.GoToPurchasing_ItemPage();
            var itemCreateModalPage = itemPage.ClickOnFirstItem();
            var picturePage = itemCreateModalPage.ClickOnPicturePage();

            if (picturePage.IsPictureAdded())
            {
                picturePage.DeletePicture();
            }

            picturePage.AddPicture(imagePath);

            // Fonction à créer pour vérifier la présence de l'image sur l'onglet "Picture" aprés l'ajout
            Assert.IsTrue(picturePage.IsPictureAdded(), "IMAGE NON AJOUTEE");

            //Aller à l'onglet "General Information" et modifier une donnée (comme "Workshop")
            var generalInfoPage = itemCreateModalPage.ClickOnGeneralInformationPage();
            generalInfoPage.SetGroupName(group);

            //Revenir à l'onglet "Picture" et vérifier si l'image est toujours présente
            picturePage = itemCreateModalPage.ClickOnPicturePage();
            Assert.IsTrue(picturePage.IsPictureAdded(), "L'image a été supprimée après la modification des données de l'item");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_UseCaseDesactivateReplaceByAnotherItem()
        {
            //Prepare
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string itemName = itemNameToday + "-" + new Random().Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string recipeName = "Recipe" + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            try
            {
                //Create an item
                itemPage = homePage.GoToPurchasing_ItemPage();
                ItemCreateModalPage itemCreateModalPage = itemPage.ItemCreatePage();
                ItemGeneralInformationPage itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);

                //Create New Receipe
                var recipesPage = homePage.GoToMenus_Recipes();
                recipesPage.ResetFilter();
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());
                recipeGeneralInfosPage.AddVariantWithSite(site, variant);
                if (!recipeGeneralInfosPage.IsFirstVariantIsVisible())
                {
                    recipeGeneralInfosPage.Validate();
                }
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
                recipeVariantPage.AddIngredient(itemName);
                recipeVariantPage.BackToList();

                //Aller sur item - Use Case
                itemPage = homePage.GoToPurchasing_ItemPage();
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
                ItemUseCasePage useCase = generalInfo.ClickOnUseCasePage();
                var totalUseCase = useCase.CheckTotalNumber();
                if (totalUseCase > 0)
                {
                    useCase.PageSize("100");
                    var selectedusecase = useCase.verifyUseCaseIsNotCheked();
                    var replacebyanotheritembuttonIsDisabled = useCase.ReplaceByAnotherItemButtonIsDisabled();
                    if (selectedusecase)
                    {
                        Assert.IsTrue(replacebyanotheritembuttonIsDisabled, "Le bouton Replace by another item apparaître activé pourtant aucune ligne la liste des items sous l'onglet use case n'est cochée.");
                        useCase.SelectBoxFirstUseCase();
                        useCase.UnSelectBoxFirstUseCase();
                        replacebyanotheritembuttonIsDisabled = useCase.ReplaceByAnotherItemButtonIsDisabled();
                        Assert.IsTrue(replacebyanotheritembuttonIsDisabled, "Le bouton Replace by another item apparaître activés pourtant aucune ligne la liste des items sous l'onglet use case n'est cochée.");
                    }
                    else
                    {
                        Assert.IsFalse(replacebyanotheritembuttonIsDisabled, "Le bouton Replace by another item apparaître désactivé pourtant une ligne de la liste des items sous l'onglet use case est cochée.");
                    }
                }
            }
            finally
            {
                //Delete Receipe
                var recipesPage = homePage.GoToMenus_Recipes();
                recipesPage.MassiveDeleteRecipes(recipeName, site, recipeType);
                homePage.GoToPurchasing_ItemPage();
                itemPage.ResetFilter();
                //Desactivate packaging
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
                generalInfo.DeactivatePrice(site, supplier, packagingName);

                //Delete Item
                generalInfo.BackToList();
                MassiveDeleteModal massiveDelete = itemPage.MenuMassiveDelete();
                massiveDelete.Filter(MassiveDeleteModal.FilterType.SearchByName, itemName);
                massiveDelete.Filter(MassiveDeleteModal.FilterType.ShowInactiveSuppliers, true);
                massiveDelete.Filter(MassiveDeleteModal.FilterType.SupplierMultiple, supplier);
                massiveDelete.ClickSearch();
                massiveDelete.ClickSelectAllButton();
                massiveDelete.Delete();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_caracteristics_ExportItemsPurchase()
        {
            // Préparation des paramètres
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            bool newVersionPrint = true;
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();
            string itemName = itemNameToday + "-" + new Random().Next().ToString();
            string caracteristics = "TESTTEST" + string.Concat(Enumerable.Range(0, 4).Select(_ => new Random().Next(0, 10).ToString()));
            // Connexion et navigation
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Création de l'élément
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
            itemPage = itemGeneralInformationPage.BackToList();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            itemGeneralInformationPage = itemPage.ClickOnFirstItem();
            var itemLabelPage = itemGeneralInformationPage.ClickOnLabelPage();
            // Ajout des caractéristiques
            itemLabelPage.EditLabel(caracteristics);
            itemLabelPage.Validate();
            itemPage = itemGeneralInformationPage.BackToList();

            // Filtrer et exporter les résultats
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.ShowAll, true);
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            itemPage.Filter(ItemPage.FilterType.Site, site);
            itemPage.Filter(ItemPage.FilterType.Supplier, supplier);
            itemPage.Filter(ItemPage.FilterType.Group, group);
            DeleteAllFileDownload();
            itemPage.ClearDownloads();
            itemPage.Export(newVersionPrint);
            // Récupérer le fichier exporté
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();
            var correctDownloadedFile = itemPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);
            int resultNumber = OpenXmlExcel.GetExportResultNumber("Sheet 1", filePath);
            var listResult = OpenXmlExcel.GetValuesInList("Caracteristics", "Sheet 1", filePath);
            //Assert
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.IsTrue(listResult.Contains(caracteristics), MessageErreur.EXCEL_DONNEES_KO);

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_ALLERGENE_Name_Change()
        {
            string itemName = "Item-" + DateUtils.Now.ToString("dd/MM/yyyy") + "-" + new Random().Next().ToString();
            string allergenName = "Allergen-" + DateUtils.Now.ToString("dd/MM/yyyy") + "-" + new Random().Next().ToString();
            string nameEnglish = "Allergen-" + DateUtils.Now.ToString("dd/MM/yyyy") + "-" + new Random().Next().ToString();
            string newAllergenName = "NewAllergen-" + DateUtils.Now.ToString("dd/MM/yyyy") + "-" + new Random().Next().ToString();
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = "10";
            string qty = "10";
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            bool newVersionPrint = true;

            //arrange
            var homePage = LogInAsAdmin();
            //act
            var parametersProduction = homePage.GoToParameters_ProductionPage();
            // Allergen Tab
            parametersProduction.GoToTab_Allergen();
            // Create a new allergen
            parametersProduction.CreateNewAllergen(allergenName, nameEnglish);

            // Edit allergen and change name
            parametersProduction.ClickEditAllergen(allergenName);
            parametersProduction.ChangeAllergenName(newAllergenName);

            // Create a new item and link it to the allergen
            var itemsPage = parametersProduction.GoToPurchasing_ItemPage();
            itemsPage.ClearDownloads();
            itemsPage.ResetFilter();
            try
            {
                // Create New Item With Packaging
                var itemCreateModalPage = itemsPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);

                // Add the new allergen to the item
                itemsPage = itemGeneralInformationPage.BackToList();
                itemsPage.Filter(ItemPage.FilterType.Search, itemName);
                itemGeneralInformationPage = itemsPage.ClickOnFirstItem();
                var itemIntolerancePage = itemsPage.ClickOnIntolerance();
                itemIntolerancePage.AddAllergen(newAllergenName);
                itemGeneralInformationPage.BackToList();

                // Export item details
                itemsPage.ResetFilter();
                itemsPage.Filter(ItemPage.FilterType.Search, itemName);
                itemsPage.Export(newVersionPrint);

                // Verify the export in Excel
                DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
                FileInfo[] taskFiles = taskDirectory.GetFiles();
                var correctDownloadedFile = itemsPage.GetExportExcelFile(taskFiles);
                Assert.IsNotNull(correctDownloadedFile);

                var fileName = correctDownloadedFile.Name;
                var filePath = Path.Combine(downloadsPath, fileName);

                string expectedColumnName = "Allergen " + newAllergenName;

                // Check Excel file for allergen data
                int resultNumber = OpenXmlExcel.GetExportResultNumber("Sheet 1", filePath);
                Assert.IsNotNull(resultNumber);

                bool result = OpenXmlExcel.ReadAllDataInColumn(expectedColumnName, "Sheet 1", filePath, "X");
                Assert.IsTrue(result, MessageErreur.EXCEL_DONNEES_KO);
            }
            finally
            {
                // Supprimer l'item 
                itemsPage.ResetFilter();
                itemsPage.Filter(ItemPage.FilterType.Search, itemName);
                var itemGeneralInformationPage = itemsPage.ClickOnFirstItem();
                itemGeneralInformationPage.ClickOnDeleteItem();
                itemGeneralInformationPage.ConfirmDelete();

                // Supprimer l'allergen
                parametersProduction.GoToParameters_ProductionPage();
                parametersProduction.GoToTab_Allergen();
                parametersProduction.ClickDeleteAllergen(newAllergenName);
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_ExpireDatedecoupe()
        {

            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();
            // Arrange
            LogInAsAdmin();

            // Aller sur la page Purchasing > Item
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();

            //Search Item 
            itemPage.Filter(ItemPage.FilterType.ShowAll, true);

            string itemName = itemPage.GetFirstItemName();
            var itemGeneralInformationsPage = itemPage.ClickOnFirstItem();
            var lastReceiptNotesPage = itemGeneralInformationsPage.ClickOnLastReceiptNotesPage();
            // Create receipt note if not created during PU_ITEM_Items_CreateItemForReceiptNote()
            if (!lastReceiptNotesPage.IsLinkToRNOK())
            {
                var purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();

                var createPurchaseOrderPage = purchaseOrdersPage.CreateNewPurchaseOrder();
                createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+10), true);
                var purchaseOrderItemPage = createPurchaseOrderPage.Submit();
                itemName = purchaseOrderItemPage.GetFirstItemName();
                purchaseOrderItemPage.Filter(PurchaseOrderItem.FilterItemType.ByName, itemName);
                purchaseOrderItemPage.SelectFirstItemPo();
                purchaseOrderItemPage.AddQuantity("2");
                purchaseOrderItemPage.Validate();
                purchaseOrderItemPage.GenerateReceiptNote(true);
                var receiptNotesItem = purchaseOrderItemPage.ValidateReceiptNoteCreation();
                var qualityChecks = receiptNotesItem.ClickOnChecksTab();
                qualityChecks.DeliveryAccepted();
                if (qualityChecks.CanClickOnSecurityChecks())
                {
                    qualityChecks.CanClickOnSecurityChecks();
                    qualityChecks.SetSecurityChecks("No");
                    qualityChecks.SetQualityChecks();
                    receiptNotesItem = qualityChecks.ClickOnReceiptNoteItemTab();
                }
                else
                {
                    receiptNotesItem = qualityChecks.ClickOnReceiptNoteItemTab();
                }

                receiptNotesItem.Validate();
                itemPage = homePage.GoToPurchasing_ItemPage();
                itemPage.ResetFilter();
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                itemGeneralInformationsPage = itemPage.ClickOnFirstItem();
                lastReceiptNotesPage = itemGeneralInformationsPage.ClickOnLastReceiptNotesPage();
            }
            // Act
            // Vérifier la visibilité de la colonne "Expire date" sur un écran de grande taille
            WebDriver.Manage().Window.Size = new System.Drawing.Size(1920, 1080);
            var isExpireDateVisibleOnLargeScreen = lastReceiptNotesPage.IsExpireDateColumnVisible();
            Assert.IsTrue(isExpireDateVisibleOnLargeScreen, "La colonne 'Expire date' ne s'affiche pas correctement sur un grand écran.");

            // Changer la résolution à celle d'un écran de PC portable
            WebDriver.Manage().Window.Size = new System.Drawing.Size(1366, 768); // Simulation d'un écran de PC portable
            var isExpireDateVisibleOnSmallScreen = lastReceiptNotesPage.IsExpireDateColumnVisible();
            Assert.IsTrue(isExpireDateVisibleOnSmallScreen, "La colonne 'Expire date' ne s'affiche pas correctement sur un écran de PC portable.");

            // Assert : Vérifier que la date n'est pas découpée, quelle que soit la résolution d'écran
            var expirationDate = lastReceiptNotesPage.GetExpirationDate();
            Assert.IsFalse(lastReceiptNotesPage.IsExpirationDateTruncated(), "La date de la colonne 'Expire date' est coupée ou tronquée sur un petit écran.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Import_Purchase_Book_Zero()
        {

            //Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string itemName = itemNameToday + "-" + new Random().Next().ToString();
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string supplierRef = itemName + "_SupplierRef";
            string storageQty = 10.ToString();
            string qty = 10.ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();
            bool newVersionPrint = true;

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.ClearDownloads();

            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier, null, supplierRef);
            itemPage = itemGeneralInformationPage.BackToList();
            itemPage.Filter(ItemPage.FilterType.Search, itemName.ToString());
            itemPage.Export(newVersionPrint);


            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = itemPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Le fichier généré est vide");

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string idColumnName, string id, string sheetName, string fileName, string value
            OpenXmlExcel.WriteDataInCell("Action", "Supplier", supplier, SHEET1, filePath, "M", CellValues.SharedString);
            OpenXmlExcel.WriteDataInCell("Prod. weight", "Supplier", supplier, SHEET1, filePath, "0", CellValues.String);
            OpenXmlExcel.WriteDataInCell("Prod. qty", "Supplier", supplier, SHEET1, filePath, "0", CellValues.String);
            OpenXmlExcel.WriteDataInCell("Storage qty", "Supplier", supplier, SHEET1, filePath, "0", CellValues.String);
            OpenXmlExcel.WriteDataInCell("Is main supplier", "Supplier", supplier, SHEET1, filePath, "VRAI", CellValues.Boolean);
            OpenXmlExcel.WriteDataInCell("Supplier", "Supplier", supplier, SHEET1, filePath, "A REFERENCIAR", CellValues.SharedString);

            WebDriver.Navigate().Refresh();
            var importPopup = itemPage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);
            bool verifyProdQty = itemPage.VerifyProdQty();
            bool verifyProdWeight = itemPage.VerifyProdWeight();
            bool verifyStorageQty = itemPage.VerifyStorageQty();

            Assert.IsTrue(verifyProdQty, "Le message d'erreur ne s'affiche pas");
            Assert.IsTrue(verifyProdWeight, "Le message d'erreur ne s'affiche pas");
            Assert.IsTrue(verifyStorageQty, "Le message d'erreur ne s'affiche pas");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Packaging_CheckBoxAlignment()
        {
            // Prepare
            string itemName = itemNameToday + "-" + new Random().Next().ToString();
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = "10";
            string qty = "10";

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            homePage.SetNewVersionKeywordValue(true);

            // Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
            itemPage = itemGeneralInformationPage.BackToList();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);

            // Click on first item and get the item details container
            itemPage.ClickOnFirstItem();
            var itemDetails = itemPage.GetItemDetailsContainer(); // Utilise la méthode pour récupérer le conteneur

            // Vérifications dans la méthode principale
            var roundToUsPosition = itemPage.GetCheckBoxPosition(itemDetails, "detail.IsRounded");
            var roundToUsLabelPosition = itemPage.GetLabelPosition(itemDetails, "Round. To US");

            var purchasablePosition = itemPage.GetCheckBoxPosition(itemDetails, "detail.IsPurchasable");
            var purchasableLabelPosition = itemPage.GetLabelPosition(itemDetails, "Purchasable");

            var unavailablePosition = itemPage.GetCheckBoxPosition(itemDetails, "detail.IsUnavailable");
            var unavailableLabelPosition = itemPage.GetLabelPosition(itemDetails, "Unavailable");

            // Vérification pour Round to US
            Assert.IsTrue(Math.Abs(roundToUsLabelPosition.X - roundToUsPosition.X) <= 15,
                "Round to US checkbox is not aligned under its title.");

            // Vérification pour Purchasable
            Assert.IsTrue(Math.Abs(purchasableLabelPosition.X - purchasablePosition.X) <= 15,
                "Purchasable checkbox is not aligned under its title.");

            // Vérification pour Unavailable
            Assert.IsTrue(Math.Abs(unavailableLabelPosition.X - unavailablePosition.X) <= 15,
                "Unavailable checkbox is not aligned under its title.");

            // Vérification des positions alignées (Y position)
            Assert.IsTrue(itemPage.AreCheckBoxesAligned(roundToUsPosition.Y, purchasablePosition.Y, unavailablePosition.Y),
                "Les cases à cocher ne sont pas alignées correctement verticalement.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Duplicates_Currency_Packaging()
        {
            //Prepare
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
            string currency = TestContext.Properties["Currency"].ToString();
            string itemName = itemNameToday + "-" + new Random().Next().ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
            itemPage = itemGeneralInformationPage.BackToList();

            // Assert 
            itemPage.Filter(ItemPage.FilterType.Search, itemName.ToString());
            Assert.AreEqual(itemName, itemPage.GetFirstItemName(), "Le nouvel item n'a pas été créé et ajouté à la liste.");
            var itemDetail = itemPage.ClickOnItemDetails();
            bool itemVerifiedcurrency = itemPage.GetPriceCurrency(decimalSeparatorValue, currency);
            //Assert
            Assert.IsTrue(itemVerifiedcurrency, "Item currency is not verified.");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_AccessItemIndex()

        {   //Prepare
            string Area = "Purchasing";
            string Controller = "Item";
            DateTime dateFrom = DateTime.Now;
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            ParametersLogs logsParameter = homePage.GoToParameters_Logs();
            logsParameter.WaitPageLoading();
            logsParameter.Filter(ParametersLogs.FilterType.minDate, dateFrom);
            bool isVerifiedLogs = logsParameter.VerifiedLogs(Area, Controller);

            //Assert
            Assert.IsTrue(isVerifiedLogs, "Il ne faut pas avoir un message d'erreur de timeout.");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_DesactivateNotPurchasable()
        {
            var homePage = LogInAsAdmin();

            // Création de l'élément
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.ActiveNotPurchasable, true);
            itemPage.DeactivateNotPurchasable();

            bool isPopupVisible = itemPage.IsItemsDeactivationReportPopupVisible();
            bool areButtonsVisible = itemPage.AreCancelAndSaveButtonsVisible();

            Assert.IsTrue(isPopupVisible, "La pop-up 'Items deactivation report' n'est pas apparue.");
            Assert.IsTrue(areButtonsVisible, "Les boutons 'Cancel' ou 'Save' ne sont pas visibles.");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_ShowAveragePrice()
        {
            //Prepare
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
            string place = TestContext.Properties["InventoryPlace"].ToString();
            string itemName = itemNameToday + "-" + new Random().Next().ToString();
            string value = 10.ToString();

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Création de l'élément
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();


            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
            itemGeneralInformationPage.BackToList();
            var inventoryPage = homePage.GoToWarehouse_InventoriesPage();
            var inventoryCreateModalPage = inventoryPage.InventoryCreatePage();
            inventoryCreateModalPage.FillField_CreateNewInventory(DateUtils.Now, site, place);
            var itemInventory = inventoryCreateModalPage.Submit();
            itemInventory.Filter(InventoryItem.FilterItemType.SearchByName, itemName);
            itemInventory.SelectFirstItem();
            itemInventory.AddPhysicalPackagingQuantity(itemName, value);
            homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            itemPage.ClickOnFirstItem();
            itemGeneralInformationPage.ScrollDown();
            var AvPrice = itemGeneralInformationPage.GetFirstPackagingAveragePrice();
            itemGeneralInformationPage.SelectFirstPackaging();
            var AvPriceAfterClick = itemGeneralInformationPage.GetFirstPackagingAveragePrice();

            Assert.AreEqual(AvPrice, AvPriceAfterClick, "Average Price ne doit pas changé !");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_GroupingSameSite_Supplier()
        {
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            string supplier = TestContext.Properties["Prodman_Needs_Item_Supplier1"].ToString();

            var packagingNames = new List<string>
       {
         "Bandeja", "Bolsa"
       };
            string itemName = itemNameToday + "-" + new Random().Next().ToString();
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            foreach (var packagingName in packagingNames)
            {
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
            }
            itemPage = itemGeneralInformationPage.BackToList();
            itemPage.Filter(ItemPage.FilterType.Search, itemName.ToString());
            itemPage.ClickFlecheAGauche();

            var sitegroup = itemPage.GetSite();
            var supplier1 = itemPage.GetSupplier1();
            var supplier2 = itemPage.GetSupplier();


            Assert.AreEqual(sitegroup, site, "le site est n'est pas le meme site cree ");
            Assert.AreEqual(supplier1, supplier2, "Les packagings ne sont pas regroupés correctement pour le site '{site}' et le fournisseur '{supplier}'. ");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Filte_KeywordsDetails()
        {
            //Prepare
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string supplierRef = "coucou";
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();

            string itemName = itemNameToday + "-" + new Random().Next(100, 999).ToString() + DateTime.Now;

            //arrange
            HomePage homePage = LogInAsAdmin();

            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            string Keyword = TestContext.Properties["Item_Keyword"].ToString();
            // count 
            var itemCreateModalPage = itemPage.ItemCreatePage();
            ItemGeneralInformationPage itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
            ItemKeywordPage itemKeywordPage = itemGeneralInformationPage.ClickOnKeywordItem();

            // add keyword details //
            itemKeywordPage.AddKeywordDetails(Keyword);

            itemPage = itemGeneralInformationPage.BackToList();
            itemPage.Filter(ItemPage.FilterType.Keyword, Keyword.ToString());
            itemPage.Filter(ItemPage.FilterType.Search, itemName.ToString());

            // Assert 
            Assert.AreEqual(itemPage.CheckTotalNumber(), 1, MessageErreur.FILTRE_ERRONE, "Le filtre KEYWORDS doit filtrer avec les Keyword Details.");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Duplication_item_Duplication_keywordsdetails()
        {

            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            string siteDuplicated = TestContext.Properties["SiteToFlightBob"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();
            string itemName = itemNameToday + "-" + new Random().Next().ToString();

            LogInAsAdmin();
            HomePage homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
            itemGeneralInformationPage.DuplicateItem(siteDuplicated);
            ItemKeywordPage itemKeywordPage = itemGeneralInformationPage.ClickOnKeywordItem();
            itemKeywordPage.Filter(ItemKeywordPage.FilterType.Search, siteDuplicated.ToString());
            bool isDuplicated = itemKeywordPage.IsPackagingDuplicated();
            Assert.IsTrue(isDuplicated, "La Duplication de item detail n'est pas pris en compte les keywords details");
        }
 
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_ReplaceByAnotherItemMultipleItems()
        {
            // Préparer les variables
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
            string itemName = "itemNAmeforTEST";
            string itemName1 = "itemNAmeforTEST1";
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();

            string recipeName = "recipeNameTest";
            int nbPortions = new Random().Next(1, 30);

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            if (itemPage.CheckTotalNumber() == 0)
            {
                // Création de deux items
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGenerlInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);

                var itemCreatePackagingPage = itemGenerlInformationPage.NewPackaging();
                itemGenerlInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                itemPage = itemGenerlInformationPage.BackToList();
            }
            itemPage.Filter(ItemPage.FilterType.Search, itemName1);
            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneraInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName1, group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneraInformationPage.NewPackaging();
                itemGeneraInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                itemPage = itemGeneraInformationPage.BackToList();
            }
            // Création d'une recette
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);
            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

                recipeGeneralInfosPage.AddVariantWithSite(site, variant);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();

                // ajouter 20 fois le même ingrédient
                for (int i = 0; i < 20; i++)
                {
                    recipeVariantPage.AddIngredient(itemName);
                }

                recipesPage = recipeVariantPage.BackToList();
            }
            // deplacer les igrendients 
            homePage.GoToPurchasing_ItemPage();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            var itemGeneralInformationPage = itemPage.ClickFirstItem();
            ItemUseCasePage itemUseCasePage = itemGeneralInformationPage.ClickOnUseCasePage();
            itemUseCasePage.WaitPageLoading();
            itemUseCasePage.SelectAll();
            var nombreIngredients = itemUseCasePage.CheckTotalNumber() == 20;
            Assert.IsTrue(nombreIngredients, "le nombre des ingrédient n'est pas 20");
            itemUseCasePage.ReplaceByAnotherItem(itemName1);
            var resultat = itemUseCasePage.CheckTotalNumber() == 0;
            //assert
            Assert.IsTrue(resultat, "les igrendients ne sont pas remplacé à l'autre item ");
            homePage.GoToPurchasing_ItemPage();
            itemPage.Filter(ItemPage.FilterType.Search, itemName1);
            itemPage.ClickFirstItem();
            itemGeneralInformationPage.ClickOnUseCasePage();
            itemUseCasePage.SelectAll();
            itemUseCasePage.ReplaceByAnotherItem(itemName);
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_ITEM_Import_ErrorMessageScroll()
        {
            // Prepare
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
            string itemName = "itemErrorMessageScroll";

            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();
            homePage.Navigate();

            bool newVersionPrint = true;

            // Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);

            if (itemPage.CheckTotalNumber() == 0)
            {
                // Création d'un item
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGenerlInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGenerlInformationPage.NewPackaging();
                itemGenerlInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                itemPage = itemGenerlInformationPage.BackToList();
            }

            homePage.GoToPurchasing_ItemPage();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            itemPage.ClearDownloads();

            foreach (var file in new DirectoryInfo(downloadsPath).GetFiles())
            {
                file.Delete();
            }

            itemPage.Export(newVersionPrint);

            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            var correctDownloadedFile = itemPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Écriture des données dans le fichier Excel exporté
            OpenXmlExcel.WriteDataInCell("Action", "Supplier", supplier, SHEET1, filePath, "A", CellValues.SharedString);
            OpenXmlExcel.WriteDataInCell("Is main supplier", "Supplier", supplier, SHEET1, filePath, "VRAI", CellValues.Boolean);
            OpenXmlExcel.WriteDataInCell("Price id", "Supplier", supplier, SHEET1, filePath, "M", CellValues.String);
            OpenXmlExcel.WriteDataInCell("Supplier", "Supplier", supplier, SHEET1, filePath, "M", CellValues.SharedString);

            WebDriver.Navigate().Refresh();

            var importPopup = itemPage.Import();
            importPopup.ImportExcel(correctDownloadedFile.FullName);

            var resultat = itemPage.IsElementVisible();

            // Assert
            Assert.IsTrue(resultat, "L'importation a été effectuée avec succès.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_DeleteItemUnpurchasableForDatasheet()
        {
            //Prepare
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
            string datasheetForUseCase = "Datasheet for Usecase";
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string itemName = itemNameToday + "-" + new Random().Next().ToString();
            string recipeName = "Recipe for Usecase 1";
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            int nbPortions = new Random().Next(1, 30);

            RecipesVariantPage recipesVariantPage;
            DatasheetDetailsPage datasheetDetailPage;
            ItemGeneralInformationPage itemGeneralInformationPage;
            //arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();

            //Add Item 
            var itemCreateModalPage = itemPage.ItemCreatePage();
            itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
            itemGeneralInformationPage.SetUnpurchasableItem(site, supplier, packagingName);
            WebDriver.Navigate().Refresh();

            // Assert 
            Assert.IsFalse(itemGeneralInformationPage.IsPurchasable(site, supplier), "La fonctionnalité 'SetUnpurchasable items' ne fonctionne pas.");

            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.ResetFilter();
            recipesPage.Filter(RecipesPage.FilterType.SearchRecipe, recipeName);
            if (recipesPage.CheckTotalNumber() == 0)
            {
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());
                recipesPage = recipeGeneralInfosPage.BackToList();
                recipesPage.ResetFilter();
                recipesPage.Filter(RecipesPage.FilterType.ShowWithoutVariants, true);
            }


            var datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetForUseCase);
            datasheetPage.Filter(DatasheetPage.FilterType.Sites, site);
            datasheetPage.Filter(DatasheetPage.FilterType.SortBy, "Name");

            if (datasheetPage.CheckTotalNumber() == 0)
            {
                var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
                datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetForUseCase, guestType, site);
                datasheetDetailPage.BackToList();

                datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetForUseCase);
                datasheetPage.Filter(DatasheetPage.FilterType.Sites, site);
                datasheetPage.Filter(DatasheetPage.FilterType.SortBy, "Name");
            }

            datasheetDetailPage = datasheetPage.SelectFirstDatasheet();
            var recipeVariantPage = datasheetDetailPage.EditRecipe();
            recipeVariantPage.AjouterIngredient(itemName);
            datasheetDetailPage.CloseDetail();
            datasheetDetailPage.ShowDetails();
            var isRecipeDetailsVisible = datasheetDetailPage.IsNewSubRecipeAdded(itemName);
            Assert.IsTrue(isRecipeDetailsVisible, "La sous-recette n'a pas été ajoutée à la recette.");
            recipesVariantPage = datasheetDetailPage.EditRecipe();
            bool verifyIsExistFirstIngredient = recipesVariantPage.verifyIsExistIngredients(itemName);
            recipesVariantPage.DeleteIngredient(itemName);
            //Assert
            Assert.IsTrue(verifyIsExistFirstIngredient, "La suppression de l'item Unpurchasable doit se faire depuis la datasheet, pas de message d'erreur affiché.");

        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_UseCaseDesactivateMultipleUpdate()
        {
            //Prepare
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string itemName = itemNameToday + "-" + new Random().Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string recipeName = "Recipe" + "-" + new Random().Next().ToString();
            int nbPortions = new Random().Next(1, 30);

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            try
            {
                //Create an item
                itemPage = homePage.GoToPurchasing_ItemPage();
                ItemCreateModalPage itemCreateModalPage = itemPage.ItemCreatePage();
                ItemGeneralInformationPage itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);

                //Create New Receipe
                var recipesPage = homePage.GoToMenus_Recipes();
                recipesPage.ResetFilter();
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());
                recipeGeneralInfosPage.AddVariantWithSite(site, variant);
                if (!recipeGeneralInfosPage.IsFirstVariantIsVisible())
                {
                    recipeGeneralInfosPage.Validate();
                }
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
                recipeVariantPage.AddIngredient(itemName);
                recipeVariantPage.BackToList();

                //Aller sur item - Use Case
                itemPage = homePage.GoToPurchasing_ItemPage();
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
                ItemUseCasePage useCase = generalInfo.ClickOnUseCasePage();
                var totalUseCase = useCase.CheckTotalNumber();
                if (totalUseCase > 0)
                {
                    useCase.PageSize("100");
                    var selectedusecase = useCase.verifyUseCaseIsNotCheked();
                    var multipleupdateIsdisabled = useCase.MultipleUpdateIsDisabled();
                    if (selectedusecase)
                    {
                        Assert.IsTrue(multipleupdateIsdisabled, "Le bouton multiple update apparaître activé pourtant aucune ligne la liste des items sous l'onglet use case n'est cochée.");

                        useCase.SelectBoxFirstUseCase();
                        useCase.UnSelectBoxFirstUseCase();

                        multipleupdateIsdisabled = useCase.ReplaceByAnotherItemButtonIsDisabled();
                        Assert.IsTrue(multipleupdateIsdisabled, "Le bouton multiple update apparaître activé pourtant aucune ligne la liste des items sous l'onglet use case n'est cochée.");
                    }
                    else
                    {
                        Assert.IsFalse(multipleupdateIsdisabled, "Le bouton multiple update apparaître désactivé pourtant une ligne de la liste des items sous l'onglet use case est cochée.");
                    }
                }
            }
            finally
            {
                //Delete Receipe
                var recipesPage = homePage.GoToMenus_Recipes();
                recipesPage.MassiveDeleteRecipes(recipeName, site, recipeType);
                homePage.GoToPurchasing_ItemPage();
                itemPage.ResetFilter();

                //Desactivate packaging
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
                generalInfo.DeactivatePrice(site, supplier, packagingName);

                //Delete Item
                generalInfo.BackToList();
                MassiveDeleteModal massiveDelete = itemPage.MenuMassiveDelete();
                massiveDelete.Filter(MassiveDeleteModal.FilterType.SearchByName, itemName);
                massiveDelete.Filter(MassiveDeleteModal.FilterType.ShowInactiveSuppliers, true);
                massiveDelete.Filter(MassiveDeleteModal.FilterType.SupplierMultiple, supplier);
                massiveDelete.ClickSearch();
                massiveDelete.ClickSelectAllButton();
                massiveDelete.Delete();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Item_UseCase_UnselectAll()
        {
            //Prepare
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string itemName = itemNameToday + "-" + new Random().Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string recipeName1 = "Recipe1" + "-" + DateTime.Now.ToString();
            string recipeName2 = "Recipe2" + "-" + DateTime.Now.ToString();
            int nbPortions = new Random().Next(1, 30);

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            try
            {
                //Create an item
                ItemCreateModalPage itemCreateModalPage = itemPage.ItemCreatePage();
                ItemGeneralInformationPage itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);

                //Create Receipe 1
                var recipesPage = homePage.GoToMenus_Recipes();
                recipesPage.ResetFilter();
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName1, recipeType, nbPortions.ToString());
                recipeGeneralInfosPage.AddVariantWithSite(site, variant);
                if (!recipeGeneralInfosPage.IsFirstVariantIsVisible())
                {
                    recipeGeneralInfosPage.Validate();
                }
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
                recipeVariantPage.AddIngredient(itemName);
                recipeVariantPage.BackToList();

                //Create New Receipe 2
                recipesPage = homePage.GoToMenus_Recipes();
                recipesPage.ResetFilter();
                recipesCreateModalPage = recipesPage.CreateNewRecipe();
                recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName2, recipeType, nbPortions.ToString());
                recipeGeneralInfosPage.AddVariantWithSite(site, variant);
                if (!recipeGeneralInfosPage.IsFirstVariantIsVisible())
                {
                    recipeGeneralInfosPage.Validate();
                }
                recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
                recipeVariantPage.AddIngredient(itemName);
                recipeVariantPage.BackToList();
                itemPage = homePage.GoToPurchasing_ItemPage();
                itemPage.ResetFilter();
                //Etre sur un item, onglet Use Case et avoir des Recipe Types disponibles

                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
                ItemUseCasePage useCase = generalInfo.ClickOnUseCasePage();

                //Cliquer sur UnSelect All(tout est désélectionné)
                useCase.SelectAllRecipeTypes();
                useCase.UnSelectAllRecipeTypes();
                var selectedNumber = useCase.CheckSelectedNumber();
                //Assert
                Assert.AreEqual(0, selectedNumber, "UnselectAll ne fonctionne pas correctement");
            }

            finally
            {
                //Delete Receipe 1
                var recipesPage = homePage.GoToMenus_Recipes();
                if (!string.IsNullOrEmpty(recipeName1))
                {
                    recipesPage.MassiveDeleteRecipes(recipeName1, site, recipeType);
                }
                homePage.GoToPurchasing_ItemPage();
                itemPage.ResetFilter();

                //Delete Receipe2
                recipesPage = homePage.GoToMenus_Recipes();
                if (!string.IsNullOrEmpty(recipeName2))
                {
                    recipesPage.MassiveDeleteRecipes(recipeName2, site, recipeType);
                }
                homePage.GoToPurchasing_ItemPage();
                itemPage.ResetFilter();

                //Desactivate packaging
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
                generalInfo.DeactivatePrice(site, supplier, packagingName);

                //Delete Item
                generalInfo.BackToList();
                MassiveDeleteModal massiveDelete = itemPage.MenuMassiveDelete();
                massiveDelete.Filter(MassiveDeleteModal.FilterType.SearchByName, itemName);
                massiveDelete.Filter(MassiveDeleteModal.FilterType.ShowInactiveSuppliers, true);
                massiveDelete.Filter(MassiveDeleteModal.FilterType.SupplierMultiple, supplier);
                massiveDelete.ClickSearch();
                massiveDelete.ClickSelectAllButton();
                massiveDelete.Delete();
            }
        }


        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_SupplierFilter()
        {
            //Prepare
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
            string itemName = itemNameToday + "-" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();

            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
            itemPage = itemGeneralInformationPage.BackToList();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName.ToString());
            string firstItemName = itemPage.GetFirstItemName();
            Assert.AreEqual(itemName, firstItemName, "Le nouvel item n'a pas été créé et ajouté à la liste.");
            itemPage.Filter(ItemPage.FilterType.Supplier, supplier);
            itemPage.Filter(ItemPage.FilterType.Search, itemName.ToString());
            Assert.AreEqual(itemName, firstItemName, "le filtre supplier ne fonctionne pas ");
            itemPage.ClickFlecheAGauche();
            var supplierIndex = itemPage.GetSupplierFromIndex();
            Assert.AreEqual(supplierIndex, supplier, "le filtre supplier ne fonctionne pas");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_Dropdown()
        {
            //Prepare
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site1 = TestContext.Properties["SiteACE"].ToString();
            string site2 = TestContext.Properties["SiteBCN"].ToString();
            string site3 = TestContext.Properties["SiteMAD"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = "10";
            string qty = "10";
            string itemName = "Item-" + DateUtils.Now.ToString("dd/MM/yyyy") + "-" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            try
            {
                itemPage.ResetFilter();

                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site1, packagingName, storageQty, storageUnit, qty, supplier);
                itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site2, packagingName, storageQty, storageUnit, qty, supplier);
                itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site3, packagingName, storageQty, storageUnit, qty, supplier);
                List<string> sitesPackagingTotal = itemGeneralInformationPage.GetAllSites();
                itemPage = itemGeneralInformationPage.BackToList();
                itemPage.ResetFilter();
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                string firstItemName = itemPage.GetFirstItemName();
                Assert.AreEqual(itemName, firstItemName, "Le nouvel item n'a pas été créé et ajouté à la liste.");

                itemPage.ClickFlecheAGauche();
                bool ItemDetailsIsVisible = itemPage.ItemDetailsIsVisible();
                Assert.IsTrue(ItemDetailsIsVisible, "Unfold Item Not Clickable");

                List<string> sitesPackagingInItem = itemPage.GetAllSites();
                foreach (string site in sitesPackagingInItem)
                {
                    Assert.IsTrue(sitesPackagingTotal.Contains(site), "Unfold Item for check Site " + site + " not in details Item");
                }
                itemPage.ClickFlecheAGauche();
                itemPage.WaitPageLoading();
                bool ItemDetailsIsVisibleAfterFold = itemPage.ItemDetailsIsVisible();

                Assert.IsFalse(ItemDetailsIsVisibleAfterFold, "Fold Item Not Clickable");
            }
            finally
            {
                itemPage.ResetFilter();
                //Desactivate packaging
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
                generalInfo.DeactivatePrice(site1, supplier, packagingName);
                generalInfo.DeactivatePrice(site2, supplier, packagingName);
                generalInfo.DeactivatePrice(site3, supplier, packagingName);

                //Delete Item
                generalInfo.BackToList();
                MassiveDeleteModal massiveDelete = itemPage.MenuMassiveDelete();
                massiveDelete.Filter(MassiveDeleteModal.FilterType.SearchByName, itemName);
                massiveDelete.Filter(MassiveDeleteModal.FilterType.ShowInactiveSuppliers, true);
                massiveDelete.Filter(MassiveDeleteModal.FilterType.SupplierMultiple, supplier);
                massiveDelete.ClickSearch();
                massiveDelete.ClickSelectAllButton();
                massiveDelete.Delete();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_UseCase_GrossQuantity()
        {
            //Prepare
            string recipeName = "Recipe for Usecase 1" + "-" +new Random().Next(100, 300);
            string itemName = "Item for useCase" + "-"+ new Random().Next(1, 30);
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();

            string site = TestContext.Properties["Site"].ToString();
            string siteAce = TestContext.Properties["SiteLP"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();

            string itemForUseCase = "Item for useCase" + new Random().Next(100, 3000); 

            string recipeType = TestContext.Properties["RecipeType"].ToString();
            
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string variantBis = TestContext.Properties["RecipeVariant3"].ToString();
            
            int nbPortions = new Random().Next(1, 30);
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string serviceCategorie = TestContext.Properties["Production_Service1"].ToString();
            string serviceType = TestContext.Properties["Production_ServiceType"].ToString();
            var value = new Random().Next(1000, 9999);
            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
            itemGeneralInformationPage.NewPackaging();
            itemCreatePackagingPage.FillField_CreateNewPackaging(siteAce, packagingName, storageQty, storageUnit, qty, supplier);
            var recipesPage = homePage.GoToMenus_Recipes();
            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());
            recipeGeneralInfosPage.AddVariantWithSite(site, variant);
            var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
            recipeVariantPage.AddIngredient(itemName);
            recipeVariantPage.BackToList();
            homePage.GoToPurchasing_ItemPage();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            itemPage.ClickFirstItem();
            var itemUseCasePage = itemGeneralInformationPage.ClickOnUseCasePage();
            itemUseCasePage.ResetFilter();
            itemUseCasePage.SelectBoxFirstUseCase();
            var originalGrossQty = itemUseCasePage.GetGrossQty();
            itemUseCasePage.SetGrossQty(value);
            itemGeneralInformationPage.ClickOnUseCasePage();
            var grossQtyModified = itemUseCasePage.GetGrossQty();
            Assert.AreEqual(value, grossQtyModified, "La quantité brute n'a pas été mise à jour correctement.");
            Assert.AreNotEqual(originalGrossQty, grossQtyModified, "La quantité brute n'a pas changé par rapport à la valeur originale.");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_UseCase_NetWeight()
        {
            //Prepare
            string recipeName = "Recipe for Usecase 1" + "-" + new Random().Next(100, 300);
            string itemName = "Item for useCase" + "***" + new Random().Next(1000, 8000);
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();

            //string site = TestContext.Properties["Site"].ToString();
            string siteAce = TestContext.Properties["SiteLP"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();

            string itemForUseCase = "Item for useCase" + new Random().Next(100, 3000);

            string recipeType = TestContext.Properties["RecipeType"].ToString();

            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string variantBis = TestContext.Properties["RecipeVariant3"].ToString();

            int nbPortions = new Random().Next(1, 30);
            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();
            string serviceCategorie = TestContext.Properties["Production_Service1"].ToString();
            string serviceType = TestContext.Properties["Production_ServiceType"].ToString();
            var value = new Random().Next(1000, 9999);
            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            try
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(siteAce, packagingName, storageQty, storageUnit, qty, supplier);
                var recipesPage = homePage.GoToMenus_Recipes();
                var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());
                recipeGeneralInfosPage.AddVariantWithSite(siteAce, variant);
                var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
                recipeVariantPage.AddIngredient(itemName);
                recipeVariantPage.BackToList();
                homePage.GoToPurchasing_ItemPage();
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                itemPage.ClickFirstItem();
                var itemUseCasePage = itemGeneralInformationPage.ClickOnUseCasePage();
                itemUseCasePage.ResetFilter();
                itemUseCasePage.SelectBoxFirstUseCase();
                var originalNetWeight = itemUseCasePage.GetNetWeight();
                itemUseCasePage.SetNetWeight(value);
                itemGeneralInformationPage.ClickOnUseCasePage();
                var netWeightModified = itemUseCasePage.GetNetWeight();
                Assert.AreEqual(value, netWeightModified, "La quantité brute n'a pas été mise à jour correctement.");
                Assert.AreNotEqual(originalNetWeight, netWeightModified, "La quantité brute n'a pas changé par rapport à la valeur originale.");
            }
            
            finally
            {
                RecipesPage recipesPage = homePage.GoToMenus_Recipes();
                    recipesPage.MassiveDeleteRecipes(recipeName, siteAce, recipeType);
                homePage.GoToPurchasing_ItemPage();
                itemPage.ResetFilter();
                //Desactivate packaging
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
                generalInfo.DeactivatePrice(siteAce, supplier, packagingName);
                //Delete Item
                generalInfo.BackToList();
                MassiveDeleteModal massiveDelete = itemPage.MenuMassiveDelete();
                massiveDelete.Filter(MassiveDeleteModal.FilterType.SearchByName, itemName);
                massiveDelete.Filter(MassiveDeleteModal.FilterType.ShowInactiveSuppliers, true);
                massiveDelete.Filter(MassiveDeleteModal.FilterType.SupplierMultiple, supplier);
                massiveDelete.ClickSearch();
                massiveDelete.ClickSelectAllButton();
                massiveDelete.Delete();
            }
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_PrintPriceHistory()
        {
            // Prepare
            bool Print = true;
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNamePdfBegin = "Price History Movement_-_";
            string DocFileNameZipBegin = "All_files_";
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = "10";
            string qty = "10";
            string itemName = $"{itemNameToday}-{new Random().Next()}";
            DateTime fromDatePrint = DateTime.Now;
            DateTime toDatePrint = DateTime.Now.AddDays(20);

            // Date format changed to dd/MM/yyyy
            string fromDatePrintStr = fromDatePrint.ToString("dd/MM/yyyy");
            string toDatePrintStr = toDatePrint.ToString("dd/MM/yyyy");

            // Segments à valider dans le PDF
            string[] addressSegments = new string[] { itemName, fromDatePrintStr, toDatePrintStr };

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.ClearDownloads();

            // Create Item
            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);

            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);

            itemPage = itemGeneralInformationPage.BackToList();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            itemPage.ClickOnFirstItem();

            // Generate and Validate PDF
            var printPage = itemGeneralInformationPage.printHistory(Print);
            bool isGenerated = printPage.IsReportGenerated();
            printPage.Close();

            // Assert - Validate PDF generation
            Assert.IsTrue(isGenerated, "Le fichier PDF n'a pas été généré.");

            // Purge old files and retrieve new file
            printPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            string generatedFilePath = printPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);

            FileInfo fileInfo = new FileInfo(generatedFilePath);
            fileInfo.Refresh();
            Assert.IsTrue(fileInfo.Exists, $"{generatedFilePath} non généré.");

            PdfDocument document = PdfDocument.Open(fileInfo.FullName);
            List<string> words = document.GetPages()
                                         .SelectMany(page => page.GetWords())
                                         .Select(word => word.Text)
                                         .ToList();

            foreach (string segment in addressSegments)
            {
                bool segmentPresent = words.Any(word => word.Contains(segment));
                Assert.IsTrue(segmentPresent, $"Le segment '{segment}' n'a pas été trouvé dans le PDF.");
            }

            // Validate item name in PDF
            bool itemNamePresent = words.Any(word => word.Contains(itemName));
            Assert.IsTrue(itemNamePresent, $"L'article '{itemName}' n'a pas été trouvé dans le PDF.");

            // Validate from date in PDF
            bool dateFromPresent = words.Any(word => word.Contains(fromDatePrintStr));
            Assert.IsTrue(dateFromPresent, $"La date de début '{fromDatePrintStr}' n'a pas été trouvée dans le PDF.");

            // Validate to date in PDF
            bool dateToPresent = words.Any(word => word.Contains(toDatePrintStr));
            Assert.IsTrue(dateToPresent, $"La date de fin '{toDatePrintStr}' n'a pas été trouvée dans le PDF.");
        }


        [Timeout(_timeout)]
        [TestMethod]
        public void PU_ITEM_GeneralInformation_SuppRef()
        {
                //Prepare
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
                string itemName = itemNameToday + "-" + new Random().Next().ToString();

                //Arrange
                var homePage = LogInAsAdmin();
                string newRefSupplier = "10";
                ItemPage itemPage = homePage.GoToPurchasing_ItemPage();

                //Act
                itemPage.ResetFilter();

                //Create item
                ItemCreateModalPage itemCreateModalPage = itemPage.ItemCreatePage();
                ItemGeneralInformationPage itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                itemPage = itemGeneralInformationPage.BackToList();

                //Search item
                itemPage.ResetFilter();
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                ItemGeneralInformationPage ItemGeneralInformationPage = itemPage.ClickFirstItem();
                ItemGeneralInformationPage.SelectFirstPackaging();

                //get supplier
                string refSupplier = ItemGeneralInformationPage.GetRefSupplier();

                //set supplier
                ItemGeneralInformationPage.SetSupplierRef(newRefSupplier);

                //Assert
                Assert.AreNotEqual(newRefSupplier, refSupplier, "n'a pas changé par rapport à la valeur originale..");

                itemPage.ClickOnUseCase();
                itemPage.ClickOngeneralInformation();
                ItemGeneralInformationPage.SelectFirstPackaging();

                string supplierUpdated = ItemGeneralInformationPage.GetRefSupplier();

                //Assert
                Assert.AreEqual(newRefSupplier, supplierUpdated, "n'a pas changé par rapport à la valeur originale..");

        }
    }
}
