using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Admin;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Sites;
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
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using static Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.Inventory.InventoryItem;

namespace Newrest.Winrest.FunctionalTests.Warehouse
{
    [TestClass]
    public class InventoryTests : TestBase
    {
        private const int _timeout = 600000;
        //______________________________________________UTILITAIRE_________________________________________________________________ 

        /// <summary>
        /// 
        /// Mise en place du paramétrage pour la configuration Winrest 4.0 
        /// 
        /// </summary>
        [TestMethod]
        [Priority(0)]
        [Timeout(_timeout)]
        public void WA_INVE_SetConfigWinrest4_0()
        {
            //Arrange
            var homePage = LogInAsAdmin();
            ClearCache();

            // New version search
            homePage.SetNewVersionSearchValue(true);

            // New group display
            homePage.SetNewGroupDisplayValue(true);

            // Vérifier que c'est activé
            var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            inventoriesPage.Filter(InventoriesPage.FilterType.ShowNotValidated, true);
            var inventoryItem = inventoriesPage.SelectFirstInventory();

            try
            {
                var itemName = inventoryItem.GetFirstItemName();

                // Filter  item name 
                inventoryItem.Filter(FilterItemType.SearchByName, itemName);
            }
            catch
            {
                throw new Exception("La recherche n'a pas pu être effectuée, le NewSearchMode est inactif.");
            }

            // vérifier new group display
            Assert.IsTrue(inventoryItem.IsGroupDisplayActive(), "Le paramètre 'NewGroupDisplay' n'est pas activé.");
        }

        [Priority(1)]
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_DeactivateSubGroupForTests()
        {
            //Arrange
            var homePage = LogInAsAdmin();

            homePage.SetSubGroupFunctionValue(false);

            Assert.IsFalse(homePage.GetSubGroupFunctionValue(), "La fonctionnalité de subGroup est toujours activée.");
        }

        [Priority(2)]
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_CreateItemForInventory()
        {
            string itemName = TestContext.Properties["Item_Inventory"].ToString();
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
            HomePage homePage = LogInAsAdmin();
            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            if (itemPage.CheckTotalNumber() == 0)
            {
                ItemCreateModalPage itemCreateModalPage = itemPage.ItemCreatePage();
                ItemGeneralInformationPage itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
                ItemCreateNewPackagingModalPage itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                itemPage = itemGeneralInformationPage.BackToList();
                itemPage.Filter(ItemPage.FilterType.Search, itemName.ToString());
            }
            string firstItemName = itemPage.GetFirstItemName();
            Assert.AreEqual(itemName, firstItemName, $"L'item {itemName} n'est pas présent dans la liste des items disponibles.");
        }

        [Priority(3)]
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_PrepareExportSageConfig()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string itemName = TestContext.Properties["Item_Inventory"].ToString();
            string journalInventory = TestContext.Properties["Journal_Inventory"].ToString();

            string taxName = TestContext.Properties["Item_TaxType"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            // Récupération du groupe de l'item
            string itemGroup = GetItemGroup(homePage, itemName);

            // Vérification du paramétrage
            // --> Admin Settings
            bool isAppSettingsOK = SetApplicationSettingsForSageAuto(homePage);

            // Sites -- > Analytical plan et section
            bool isAnalyticalPlanOK = VerifySiteAnalyticalPlanSection(homePage, site);

            // Parameter - Accounting --> Service categories & VAT
            bool isGroupAndVatOK = VerifyGroupAndVAT(homePage, itemGroup, taxName);

            // Parameter - Accounting --> Journal
            bool isJournalOk = VerifyAccountingJournal(homePage, site, journalInventory);

            // IntegrationDate
            var date = VerifyIntegrationDate(homePage);

            // Assert
            Assert.AreNotEqual("", itemGroup, "Le groupe de l'item n'a pas été récupéré.");
            Assert.IsTrue(isAppSettingsOK, "Les application settings pour TL ne sont pas configurés correctement.");
            Assert.IsTrue(isAnalyticalPlanOK, "La configuration des analytical plan du site n'est pas effectuée.");
            Assert.IsTrue(isGroupAndVatOK, "La configuration du group and VAT de l'item n'est pas effectuée.");
            Assert.IsTrue(isJournalOk, "La catégorie du accounting journal n'a pas été effectuée.");
            Assert.IsNotNull(date, "La date d'intégration est nulle.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_CreateInventory()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            // On va cocher la date du jour dans les settings inventory pour pouvoir lancer le test
            bool initialInventoryValue = SetInventoryDayValue(homePage);

            try
            {
                //Act
                var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();

                // Create
                var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place);
                String ID = inventoryCreateModalpage.GetInventoryNumber();
                var inventoryItem = inventoryCreateModalpage.Submit();

                var itemName = inventoryItem.GetFirstItemName();
                inventoryItem.SelectFirstItem();
                inventoryItem.AddPhysicalQuantity(itemName, "10");

                inventoriesPage = inventoryItem.BackToList();
                inventoriesPage.ResetFilters();
                inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, ID);

                //Assert
                Assert.AreEqual(ID, inventoriesPage.GetFirstID(), String.Format(MessageErreur.OBJET_NON_CREE, "L'inventory"));
            }
            finally
            {
                // Si au début du test, la case due l'inventory du jour était décochée, on vient la redécocher de nouveau (retour état initial)
                if (!initialInventoryValue)
                {
                    RemoveInventoryValue(homePage);
                }
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_CreateMobileInventory()
        {
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // On va cocher la date du jour dans les settings inventory pour pouvoir lancer le test
            bool initialInventoryValue = SetInventoryDayValue(homePage);

            try
            {
                //Act
                var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();

                // Create
                var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place);
                String ID = inventoryCreateModalpage.GetInventoryNumber();
                var inventoryGeneralInformations = inventoryCreateModalpage.SubmitMobile();

                var inventoryItem = inventoryGeneralInformations.ClickOnItems();

                var itemName = inventoryItem.GetFirstItemName();
                inventoryItem.SelectFirstItem();
                inventoryItem.AddPhysicalQuantity(itemName, "10");

                inventoriesPage = inventoryItem.BackToList();
                inventoriesPage.ResetFilters();
                inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, ID);

                //Assert
                Assert.AreEqual(ID, inventoriesPage.GetFirstID(), String.Format(MessageErreur.OBJET_NON_CREE, "L'inventory"));
            }
            finally
            {
                // Si au début du test, la case due l'inventory du jour était décochée, on vient la redécocher de nouveau (retour état initial)
                if (!initialInventoryValue)
                {
                    RemoveInventoryValue(homePage);
                }
            }
        }

        //______________________________________________FIN CREATE INVENTORY____________________________________________________

        //______________________________________________FILTRES INVENTORY_______________________________________________________

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_Filter_SearchByNumber()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            // On va cocher la date du jour dans les settings inventory pour pouvoir lancer le test
            bool initialInventoryValue = SetInventoryDayValue(homePage);

            try
            {
                //Act
                var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();

                // Create
                var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place);
                String ID = inventoryCreateModalpage.GetInventoryNumber();
                var inventoryItem = inventoryCreateModalpage.Submit();

                var itemName = inventoryItem.GetFirstItemName();
                inventoryItem.SelectFirstItem();
                inventoryItem.AddPhysicalQuantity(itemName, "10");

                inventoriesPage = inventoryItem.BackToList();
                inventoriesPage.ResetFilters();
                inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, ID);

                string firstID = inventoriesPage.GetFirstID();
                //Assert
                Assert.AreEqual(ID, firstID, String.Format(MessageErreur.FILTRE_ERRONE, "Search by number"));
            }
            finally
            {
                // Si au début du test, la case due l'inventory du jour était décochée, on vient la redécocher de nouveau (retour état initial)
                if (!initialInventoryValue)
                {
                    RemoveInventoryValue(homePage);
                }
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_Filter_SortByNumber()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            var dateFormat = homePage.GetDateFormatPickerValue();

            // On va cocher la date du jour dans les settings inventory pour pouvoir lancer le test
            bool initialInventoryValue = SetInventoryDayValue(homePage);

            var totalNumber = 0;
            string firstNumber = "";
            try
            {
                //Act
                var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();
                totalNumber = inventoriesPage.CheckTotalNumber();
                if (totalNumber < 20)
                {
                    // Create
                    var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                    inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place);
                    var inventoryItem = inventoryCreateModalpage.Submit();

                    var itemName = inventoryItem.GetFirstItemName();
                    inventoryItem.SelectFirstItem();
                    inventoryItem.AddPhysicalQuantity(itemName, "10");

                    inventoriesPage = inventoryItem.BackToList();
                    firstNumber = inventoriesPage.GetFirstInventoryNumber();
                    inventoriesPage.ResetFilters();
                }

                if (!inventoriesPage.isPageSizeEqualsTo100())
                {
                    inventoriesPage.PageSize("8");
                    inventoriesPage.PageSize("100");
                }

                inventoriesPage.Filter(InventoriesPage.FilterType.SortBy, "NUMBER");
                var isSortedByNumber = inventoriesPage.IsSortedByNumber();

                //Assert
                Assert.IsTrue(isSortedByNumber, String.Format(MessageErreur.FILTRE_ERRONE, "Sort By number"));
            }
            finally
            {
                // Si au début du test, la case due l'inventory du jour était décochée, on vient la redécocher de nouveau (retour état initial)
                if (!initialInventoryValue)
                {
                    RemoveInventoryValue(homePage);
                }
                if (totalNumber < 20)
                {
                    var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                    inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, firstNumber);
                    inventoriesPage.deleteInventory();

                }
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_Filter_SortByDate()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            var dateFormat = homePage.GetDateFormatPickerValue();

            // On va cocher la date du jour dans les settings inventory pour pouvoir lancer le test
            bool initialInventoryValue = SetInventoryDayValue(homePage);

            var totalNumber = 0;
            string firstNumber = "";
            try
            {
                //Act
                var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();
                totalNumber = inventoriesPage.CheckTotalNumber();
                if (totalNumber < 20)
                {
                    // Create
                    var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                    inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place);
                    var inventoryItem = inventoryCreateModalpage.Submit();

                    var itemName = inventoryItem.GetFirstItemName();
                    inventoryItem.SelectFirstItem();
                    inventoryItem.AddPhysicalQuantity(itemName, "10");

                    inventoriesPage = inventoryItem.BackToList();
                    firstNumber = inventoriesPage.GetFirstInventoryNumber();
                    inventoriesPage.ResetFilters();
                }

                if (!inventoriesPage.isPageSizeEqualsTo100())
                {
                    inventoriesPage.PageSize("8");
                    inventoriesPage.PageSize("100");
                }

                inventoriesPage.Filter(InventoriesPage.FilterType.SortBy, "DATE");
                var isSortedByDate = inventoriesPage.IsSortedByDate(dateFormat);

                //Assert
                Assert.IsTrue(isSortedByDate, String.Format(MessageErreur.FILTRE_ERRONE, "By date"));
            }
            finally
            {
                // Si au début du test, la case due l'inventory du jour était décochée, on vient la redécocher de nouveau (retour état initial)
                if (!initialInventoryValue)
                {
                    RemoveInventoryValue(homePage);
                }
                if (totalNumber < 20)
                {
                    var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                    inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, firstNumber);
                    inventoriesPage.deleteInventory();

                }
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_Filter_Show_Items_Validated()
        {
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();
            HomePage homePage = LogInAsAdmin();
            bool initialInventoryValue = SetInventoryDayValue(homePage);
            try
            {
                InventoriesPage inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();
                inventoriesPage.Filter(InventoriesPage.FilterType.ShowItemsNotValidated, false);
                if (inventoriesPage.CheckTotalNumber() < 20)
                {
                    InventoryCreateModalPage inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                    inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place);
                    InventoryItem inventoryItem = inventoryCreateModalpage.Submit();
                    var itemName = inventoryItem.GetFirstItemName();
                    inventoryItem.SelectFirstItem();
                    inventoryItem.AddPhysicalQuantity(itemName, "10");
                    InventoryValidationModalPage validateInventory = inventoryItem.Validate();
                    inventoryItem = validateInventory.ValidatePartialInventory();
                    inventoriesPage = inventoryItem.BackToList();
                    inventoriesPage.ResetFilters();
                    inventoriesPage.Filter(InventoriesPage.FilterType.ShowItemsNotValidated, false);
                }
                if (!inventoriesPage.isPageSizeEqualsTo100())
                {
                    inventoriesPage.PageSize("8");
                    inventoriesPage.PageSize("100");
                }
                bool isChecked = inventoriesPage.CheckShowItemsNotValidated(false);
                Assert.IsFalse(isChecked, String.Format(MessageErreur.FILTRE_ERRONE, "'Show items validated'"));
            }
            finally
            {
                // Si au début du test, la case due l'inventory du jour était décochée, on vient la redécocher de nouveau (retour état initial)
                if (!initialInventoryValue)
                {
                    RemoveInventoryValue(homePage);
                }
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_Filter_Show_Items_NotValidated()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // On va cocher la date du jour dans les settings inventory pour pouvoir lancer le test
            bool initialInventoryValue = SetInventoryDayValue(homePage);

            try
            {
                //Act
                var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();
                inventoriesPage.Filter(InventoriesPage.FilterType.ShowItemsNotValidated, true);

                if (inventoriesPage.CheckTotalNumber() < 20)
                {
                    // Create
                    var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                    inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place);
                    var inventoryItem = inventoryCreateModalpage.Submit();

                    var itemName = inventoryItem.GetFirstItemName();
                    inventoryItem.SelectFirstItem();
                    inventoryItem.AddPhysicalQuantity(itemName, "10");

                    // Retour à la liste des inventories
                    inventoriesPage = inventoryItem.BackToList();

                    inventoriesPage.ResetFilters();
                    inventoriesPage.Filter(InventoriesPage.FilterType.ShowItemsNotValidated, true);
                }

                if (!inventoriesPage.isPageSizeEqualsTo100())
                {
                    inventoriesPage.PageSize("8");
                    inventoriesPage.PageSize("100");
                }

                //Assert
                Assert.IsTrue(inventoriesPage.CheckShowItemsNotValidated(true), String.Format(MessageErreur.FILTRE_ERRONE, "'Show items not validated'"));
            }
            finally
            {
                // Si au début du test, la case due l'inventory du jour était décochée, on vient la redécocher de nouveau (retour état initial)
                if (!initialInventoryValue)
                {
                    RemoveInventoryValue(homePage);
                }
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_Filter_ShowAll()
        {
            HomePage homePage = LogInAsAdmin();
            bool initialInventoryValue = SetInventoryDayValue(homePage);
            InventoriesPage inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            inventoriesPage.ResetFilters();
            // Filter with Show active Only and check Total number
            inventoriesPage.Filter(InventoriesPage.FilterType.ShowActive, true);
            int active = inventoriesPage.CheckTotalNumber();
            // Filter with Show inactive Only and check Total number
            inventoriesPage.Filter(InventoriesPage.FilterType.ShowInactive, true);
            int inactive = inventoriesPage.CheckTotalNumber();
            // Filter with Show All and check Total number
            inventoriesPage.Filter(InventoriesPage.FilterType.ShowAll, true);
            int all = inventoriesPage.CheckTotalNumber();
            // Assert
            Assert.AreEqual(all, (inactive + active), string.Format(MessageErreur.FILTRE_ERRONE, "'Show all'"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_Filter_ShowInactive()
        {
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();
            HomePage homePage = LogInAsAdmin();
            bool initialInventoryValue = SetInventoryDayValue(homePage);
            try
            {
                InventoriesPage inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();
                inventoriesPage.Filter(InventoriesPage.FilterType.ShowInactive, true);
                if (inventoriesPage.CheckTotalNumber() < 20)
                {
                    InventoryCreateModalPage inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                    inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place, false);
                    InventoryItem inventoryItem = inventoryCreateModalpage.Submit();
                    string itemName = inventoryItem.GetFirstItemName();
                    inventoryItem.SelectFirstItem();
                    inventoryItem.AddPhysicalQuantity(itemName, "10");
                    inventoriesPage = inventoryItem.BackToList();
                    inventoriesPage.ResetFilters();
                    inventoriesPage.Filter(InventoriesPage.FilterType.ShowInactive, true);
                }
                if (!inventoriesPage.isPageSizeEqualsTo100())
                {
                    inventoriesPage.PageSize("8");
                    inventoriesPage.PageSize("100");
                }
                bool isChecked = inventoriesPage.CheckStatus(false);
                Assert.IsFalse(isChecked, string.Format(MessageErreur.FILTRE_ERRONE, "'Show only inactive'"));
            }
            finally
            {
                if (!initialInventoryValue)
                    RemoveInventoryValue(homePage);
            }

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_Filter_ShowActive()
        {
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();
            HomePage homePage = LogInAsAdmin();
            bool initialInventoryValue = SetInventoryDayValue(homePage);
            try
            {
                InventoriesPage inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();
                inventoriesPage.Filter(InventoriesPage.FilterType.ShowActive, true);
                if (inventoriesPage.CheckTotalNumber() < 20)
                {
                    InventoryCreateModalPage inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                    inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place, true);
                    InventoryItem inventoryItem = inventoryCreateModalpage.Submit();
                    string itemName = inventoryItem.GetFirstItemName();
                    inventoryItem.SelectFirstItem();
                    inventoryItem.AddPhysicalQuantity(itemName, "10");
                    inventoriesPage = inventoryItem.BackToList();
                    inventoriesPage.ResetFilters();
                    inventoriesPage.Filter(InventoriesPage.FilterType.ShowActive, true);
                }
                if (!inventoriesPage.isPageSizeEqualsTo100())
                {
                    inventoriesPage.PageSize("8");
                    inventoriesPage.PageSize("100");
                }
                bool isChecked = inventoriesPage.CheckStatus(true);
                Assert.IsTrue(isChecked, string.Format(MessageErreur.FILTRE_ERRONE, "'Show only active'"));
            }
            finally
            {
                if (!initialInventoryValue)
                    RemoveInventoryValue(homePage);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_Filter_ShowNotValidatedOnly()
        {
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();
            HomePage homePage = LogInAsAdmin();
            bool initialInventoryValue = SetInventoryDayValue(homePage);
            try
            {
                InventoriesPage inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();
                inventoriesPage.Filter(InventoriesPage.FilterType.ShowNotValidated, true);
                if (inventoriesPage.CheckTotalNumber() < 20)
                {
                    InventoryCreateModalPage inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                    inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place);
                    InventoryItem inventoryItem = inventoryCreateModalpage.Submit();
                    string itemName = inventoryItem.GetFirstItemName();
                    inventoryItem.SelectFirstItem();
                    inventoryItem.AddPhysicalQuantity(itemName, "10");
                    inventoriesPage = inventoryItem.BackToList();
                    inventoriesPage.ResetFilters();
                    inventoriesPage.Filter(InventoriesPage.FilterType.ShowNotValidated, true);
                }
                if (!inventoriesPage.isPageSizeEqualsTo100())
                {
                    inventoriesPage.PageSize("8");
                    inventoriesPage.PageSize("100");
                }
                bool isChecked = inventoriesPage.CheckValidation(false);
                Assert.IsFalse(isChecked, string.Format(MessageErreur.FILTRE_ERRONE, "'Show not validated only'"));
            }
            finally
            {
                if (!initialInventoryValue)
                    RemoveInventoryValue(homePage);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_Filter_ShowSentToSAGEOnly()
        {
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();

            HomePage homePage = LogInAsAdmin();
            var dateFormat = homePage.GetDateFormatPickerValue();
            InventoriesPage inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            inventoriesPage.ResetFilters();

            InventoryCreateModalPage inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
            inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place);
            string ID = inventoryCreateModalpage.GetInventoryNumber();
            InventoryItem inventoryItem = inventoryCreateModalpage.Submit();
            inventoriesPage = inventoryItem.BackToList();
            inventoriesPage.ResetFilters();
            inventoriesPage.Filter(InventoriesPage.FilterType.ShowAll, true);
            var inventaireNumberShowAll = inventoriesPage.CheckTotalNumber();
            bool isNotEqual = inventaireNumberShowAll != 0;
            inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, ID);
            var totalNumberByiD = inventoriesPage.CheckTotalNumber();
            //Assert
            Assert.AreEqual(1, totalNumberByiD, "Search By Number ne fonctionne Pas !");
            inventoriesPage.ResetFilters();

            inventoriesPage.Filter(InventoriesPage.FilterType.SortBy, "NUMBER");
            var isSortedByNumber = inventoriesPage.IsSortedByNumber();
            //Assert
            Assert.IsTrue(isSortedByNumber, String.Format(MessageErreur.FILTRE_ERRONE, "Sort By number"));
            inventoriesPage.Filter(InventoriesPage.FilterType.SortBy, "DATE");
            var isSortedByDate = inventoriesPage.IsSortedByDate(dateFormat);
            //Assert
            Assert.IsTrue(isSortedByDate, String.Format(MessageErreur.FILTRE_ERRONE, "By date"));

            inventoriesPage.Filter(InventoriesPage.FilterType.ShowActive, true);
            int nbActive = inventoriesPage.CheckTotalNumber();

            inventoriesPage.Filter(InventoriesPage.FilterType.ShowInactive, true);
            int nbInactive = inventoriesPage.CheckTotalNumber();
            //Assert
            Assert.AreEqual(nbInactive, inventaireNumberShowAll - nbActive, "Search By Number ne fonctionne Pas !");

            inventoriesPage.Filter(InventoriesPage.FilterType.ShowNotValidated, true);
            bool isChecked = inventoriesPage.CheckValidation(false);
            //Assert
            Assert.IsFalse(isChecked, string.Format(MessageErreur.FILTRE_ERRONE, "'Show not validated only'"));

            inventoriesPage.Filter(InventoriesPage.FilterType.DateFrom, DateTime.Now.AddDays(5));
            int inventaireNumberDateFrom = inventoriesPage.CheckTotalNumber();
            isNotEqual = inventaireNumberShowAll != inventaireNumberDateFrom;
            //Assert
            Assert.IsTrue(isNotEqual, MessageErreur.FILTRE_ERRONE);

            inventoriesPage.Filter(InventoriesPage.FilterType.DateTo, DateTime.Now.AddDays(-5));
            int inventaireNumberDateTo = inventoriesPage.CheckTotalNumber();
            isNotEqual = inventaireNumberShowAll != inventaireNumberDateTo;
            //Assert
            Assert.IsTrue(isNotEqual, MessageErreur.FILTRE_ERRONE);

            inventoriesPage.Filter(InventoriesPage.FilterType.Site, TestContext.Properties["Site"].ToString());
            int inventaireNumberSite = inventoriesPage.CheckTotalNumber();
            isNotEqual = inventaireNumberShowAll != inventaireNumberSite;
            //Assert
            Assert.IsTrue(isNotEqual, MessageErreur.FILTRE_ERRONE);

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_Filter_ShowValidatedNotSentToSAGE()
        {
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();
            string group = TestContext.Properties["Item_Group"].ToString();
            HomePage homePage = LogInAsAdmin();
            bool initialInventoryValue = SetInventoryDayValue(homePage);
            try
            {
                InventoriesPage inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();
                inventoriesPage.Filter(InventoriesPage.FilterType.ShowValidatedNotSentToSAGE, true);
                if (inventoriesPage.CheckTotalNumber() < 20)
                {
                    homePage.GoToWarehouse_InventoriesPage();
                    InventoryCreateModalPage inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                    inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place);
                    InventoryItem inventoryItem = inventoryCreateModalpage.Submit();
                    inventoryItem.Filter(FilterItemType.ByGroup, group);
                    string itemName = inventoryItem.GetFirstItemName();
                    inventoryItem.SelectFirstItem();
                    inventoryItem.AddPhysicalQuantity(itemName, "10");
                    InventoryValidationModalPage validateInventory = inventoryItem.Validate();
                    validateInventory.ValidatePartialInventory();
                    inventoriesPage = inventoryItem.BackToList();
                    inventoriesPage.ResetFilters();
                    inventoriesPage.Filter(InventoriesPage.FilterType.ShowValidatedNotSentToSAGE, true);
                }
                if (!inventoriesPage.isPageSizeEqualsTo100())
                {
                    inventoriesPage.PageSize("8");
                    inventoriesPage.PageSize("100");
                }
                bool isChecked = inventoriesPage.CheckValidation(true);
                Assert.IsTrue(isChecked, string.Format(MessageErreur.FILTRE_ERRONE, "'Show validated & not sent to SAGE'"));
                bool isSent = inventoriesPage.IsSentToSAGE();
                Assert.IsFalse(isSent, string.Format(MessageErreur.FILTRE_ERRONE, "'Show validated & not sent to SAGE'"));
            }
            finally
            {
                if (!initialInventoryValue)
                    RemoveInventoryValue(homePage);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_Filter_ExportedForSageManually()
        {
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();
            string itemName = TestContext.Properties["Item_Inventory"].ToString();
            HomePage homePage = LogInAsAdmin();
            bool initialInventoryValue = SetInventoryDayValue(homePage);
            try
            {
                InventoriesPage inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();
                inventoriesPage.Filter(InventoriesPage.FilterType.ShowExportedForSageManually, true);
                /*        if (inventoriesPage.CheckTotalNumber() < 20)
                        {
                            InventoryCreateModalPage inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                            inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place);
                            string ID = inventoryCreateModalpage.GetInventoryNumber();
                            InventoryItem inventoryItem = inventoryCreateModalpage.Submit();
                            inventoryItem.Filter(FilterItemType.SearchByName, itemName);
                            inventoryItem.SelectFirstItem();
                            inventoryItem.AddPhysicalQuantity(itemName, "2");
                            InventoryValidationModalPage validateInventory = inventoryItem.Validate();
                            validateInventory.ValidatePartialInventory();
                            inventoriesPage = inventoryItem.BackToList();
                            inventoriesPage.ResetFilters();
                            inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, ID);
                            inventoriesPage.ManualExportResultsForSage(true);
                            inventoriesPage.EnableExportForSage();
                            inventoriesPage.ResetFilters();
                            inventoriesPage.Filter(InventoriesPage.FilterType.ShowExportedForSageManually, true);
                        }*/
                if (!inventoriesPage.isPageSizeEqualsTo100())
                {
                    inventoriesPage.PageSize("8");
                    inventoriesPage.PageSize("100");
                }
                bool isChecked = inventoriesPage.CheckValidation(true);
                Assert.IsTrue(isChecked, string.Format(MessageErreur.FILTRE_ERRONE, "'Exported for sage manually'"));
                //bool isSent = inventoriesPage.IsSentToSageManually();
                bool isAccounted = inventoriesPage.isAccounted();
                Assert.IsTrue(isAccounted, string.Format(MessageErreur.FILTRE_ERRONE, "'Exported for sage manually'"));
            }
            finally
            {
                if (!initialInventoryValue)
                    RemoveInventoryValue(homePage);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_Filter_DateFrom()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();
            DateTime date1 = DateUtils.Now;
            DateTime date2 = DateUtils.Now.AddDays(+1);
            DateTime date3 = DateUtils.Now.AddDays(-1);
            bool isActive = true;
            DateTime date = DateUtils.Now.AddMonths(+12);
            var  itemNumber1 = "";
            var  itemNumber2 = "";
            var  itemNumber3 = "";
            //Arrange
            var homePage = LogInAsAdmin();
            // On va cocher la date du jour dans les settings inventory pour pouvoir lancer le test
            bool initialInventoryValue = SetInventoryDayValue(homePage);
            var dateFormat = homePage.GetDateFormatPickerValue();
            var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            try
            {
                // Create inventory 1
                var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(date1, site, place, isActive);
                itemNumber1= inventoryCreateModalpage.GetInventoryNumber();
                var inventoryItem = inventoryCreateModalpage.Submit();
                inventoriesPage = inventoryItem.BackToList();
                 // Create inventory 2
                inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(date2, site, place, isActive);
                itemNumber2 = inventoryCreateModalpage.GetInventoryNumber();
                inventoryItem = inventoryCreateModalpage.Submit();
                inventoriesPage = inventoryItem.BackToList();
                // Create inventory 3
                inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(date3, site, place, isActive);
                itemNumber3 = inventoryCreateModalpage.GetInventoryNumber();
                inventoryItem = inventoryCreateModalpage.Submit();
                inventoriesPage = inventoryItem.BackToList();
                inventoriesPage.ResetFilters();
                inventoriesPage.Filter(InventoriesPage.FilterType.DateFrom, date1);
                homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, itemNumber1);
                inventoriesPage.WaitPageLoading();
                int check = inventoriesPage.CheckTotalNumber();
                Assert.AreEqual(check, 1, String.Format(MessageErreur.FILTRE_ERRONE, "'From'"));
                inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, itemNumber2);
                inventoriesPage.WaitPageLoading();
                check = inventoriesPage.CheckTotalNumber();
                Assert.AreEqual(check, 1, String.Format(MessageErreur.FILTRE_ERRONE, "'From'"));
                inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, itemNumber3);
                inventoriesPage.WaitPageLoading();
                check = inventoriesPage.CheckTotalNumber();
                Assert.AreEqual(check,0, String.Format(MessageErreur.FILTRE_ERRONE, "'From'"));
            }
            finally
            {
                homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();
                inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, itemNumber1);
                inventoriesPage.deleteInventory();
                inventoriesPage.ResetFilters();
                inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, itemNumber2);
                inventoriesPage.deleteInventory();
                inventoriesPage.ResetFilters();
                inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, itemNumber3);
                inventoriesPage.deleteInventory();
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_Filter_DateTo()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();
            DateTime date1 = DateUtils.Now;
            DateTime date2 = DateUtils.Now.AddDays(+1);
            DateTime date3 = DateUtils.Now.AddDays(-1);
            bool isActive = true;
            DateTime date = DateUtils.Now.AddMonths(-12);
            var itemNumber1 = "";
            var itemNumber2 = "";
            var itemNumber3 = "";
            //Arrange
            var homePage = LogInAsAdmin();
            // On va cocher la date du jour dans les settings inventory pour pouvoir lancer le test
            bool initialInventoryValue = SetInventoryDayValue(homePage);
            var dateFormat = homePage.GetDateFormatPickerValue();
            var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            try
            {
                // Create inventory 1
                var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(date1, site, place, isActive);
                itemNumber1 = inventoryCreateModalpage.GetInventoryNumber();
                var inventoryItem = inventoryCreateModalpage.Submit();
                inventoriesPage = inventoryItem.BackToList();
                // Create inventory 2
                inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(date2, site, place, isActive);
                itemNumber2 = inventoryCreateModalpage.GetInventoryNumber();
                inventoryItem = inventoryCreateModalpage.Submit();
                inventoriesPage = inventoryItem.BackToList();
                // Create inventory 3
                inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(date3, site, place, isActive);
                itemNumber3 = inventoryCreateModalpage.GetInventoryNumber();
                inventoryItem = inventoryCreateModalpage.Submit();
                inventoriesPage = inventoryItem.BackToList();
                inventoriesPage.ResetFilters();
                inventoriesPage.Filter(InventoriesPage.FilterType.DateTo, date1);
                homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, itemNumber1);
                inventoriesPage.WaitPageLoading();
                int check = inventoriesPage.CheckTotalNumber();
                Assert.AreEqual(check, 1, String.Format(MessageErreur.FILTRE_ERRONE, "'To'"));
                inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, itemNumber2);
                inventoriesPage.WaitPageLoading();
                check = inventoriesPage.CheckTotalNumber();
                Assert.AreEqual(check, 0, String.Format(MessageErreur.FILTRE_ERRONE, "'To'"));
                inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, itemNumber3);
                inventoriesPage.WaitPageLoading();
                check = inventoriesPage.CheckTotalNumber();
                Assert.AreEqual(check,1, String.Format(MessageErreur.FILTRE_ERRONE, "'To'"));                
            }
            finally
            {
                homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();
                inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, itemNumber1);
                inventoriesPage.deleteInventory();
                inventoriesPage.ResetFilters();
                inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, itemNumber2);
                inventoriesPage.deleteInventory();
                inventoriesPage.ResetFilters();
                inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, itemNumber3);
                inventoriesPage.deleteInventory();
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_Filter_Site()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // On va cocher la date du jour dans les settings inventory pour pouvoir lancer le test
            bool initialInventoryValue = SetInventoryDayValue(homePage);

            try
            {
                //Act
                var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();
                inventoriesPage.Filter(InventoriesPage.FilterType.Site, site);

                if (inventoriesPage.CheckTotalNumber() < 20)
                {
                    // Create
                    var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                    inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place, true);
                    var inventoryItem = inventoryCreateModalpage.Submit();

                    var itemName = inventoryItem.GetFirstItemName();
                    inventoryItem.SelectFirstItem();
                    inventoryItem.AddPhysicalQuantity(itemName, "10");

                    // Retour à la liste des inventories
                    inventoriesPage = inventoryItem.BackToList();

                    inventoriesPage.ResetFilters();
                    inventoriesPage.Filter(InventoriesPage.FilterType.Site, site);
                }

                if (!inventoriesPage.isPageSizeEqualsTo100())
                {
                    inventoriesPage.PageSize("8");
                    inventoriesPage.PageSize("100");
                }

                //Assert
                Assert.IsTrue(inventoriesPage.VerifySite(site), String.Format(MessageErreur.FILTRE_ERRONE, "'Sites'"));
            }
            finally
            {
                // Si au début du test, la case due l'inventory du jour était décochée, on vient la redécocher de nouveau (retour état initial)
                if (!initialInventoryValue)
                {
                    RemoveInventoryValue(homePage);
                }
            }
        }

        //______________________________________________FIN FILTRES INVENTORY____________________________________________________

        //______________________________________________FILTRES ITEMS INVENTORY__________________________________________________

        [TestMethod]

        [Timeout(_timeout)]
        public void WA_INVE_Items_Filter_SearchByName()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            // On va cocher la date du jour dans les settings inventory pour pouvoir lancer le test
            bool initialInventoryValue = SetInventoryDayValue(homePage);

            try
            {
                //Act
                var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();

                // Create
                var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place, true);
                var inventoryItem = inventoryCreateModalpage.Submit();

                var itemName = inventoryItem.GetFirstItemName();

                // Filter  item name 
                inventoryItem.PageSize("8");
                inventoryItem.Filter(FilterItemType.SearchByName, itemName);
                Assert.IsTrue(inventoryItem.VerifyName(itemName), String.Format(MessageErreur.FILTRE_ERRONE, "'Search by name'"));
            }
            finally
            {
                // Si au début du test, la case due l'inventory du jour était décochée, on vient la redécocher de nouveau (retour état initial)
                if (!initialInventoryValue)
                {
                    RemoveInventoryValue(homePage);
                }
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_Items_Filter_ShowItemsWithQty()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // On va cocher la date du jour dans les settings inventory pour pouvoir lancer le test
            bool initialInventoryValue = SetInventoryDayValue(homePage);

            try
            {
                var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

                // Récupération du type de séparateur (, ou . selon les pays)
                CultureInfo ci = decimalSeparatorValue.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

                //Act
                var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();

                // Create
                var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place, true);
                var inventoryItem = inventoryCreateModalpage.Submit();

                var itemName = inventoryItem.GetFirstItemName();
                inventoryItem.SelectFirstItem();
                inventoryItem.AddPhysicalQuantity(itemName, "10");


                // Filter show items with theo/phys/qty only    
                inventoryItem.Filter(FilterItemType.ShowItemsWithQtyOnly, true);
                Assert.IsTrue(inventoryItem.IsWithTheoOrPhysQtyOnly(), String.Format(MessageErreur.FILTRE_ERRONE, "'Show items with theo/phys qty only'"));

                itemName = inventoryItem.GetFirstItemName();
                string theoQty = inventoryItem.GetTheoricalQuantity();
                double theoricalQty = Convert.ToDouble(theoQty, ci);
                double physicalQty = theoricalQty + 1;

                inventoryItem.SelectFirstItem();
                inventoryItem.AddPhysicalQuantity(itemName, physicalQty.ToString());
                // data binding
                Thread.Sleep(2000);
                inventoryItem.ResetFilter();

                // Filter show items with qty difference only
                inventoryItem.Filter(FilterItemType.ShowItemsWithDifferenceOnly, true);
                inventoryItem.PageSize("100");
                Assert.IsTrue(inventoryItem.IsWithQtyDifferenceOnly(), String.Format(MessageErreur.FILTRE_ERRONE, "'Show items with qty difference only'"));

            }
            finally
            {
                // Si au début du test, la case due l'inventory du jour était décochée, on vient la redécocher de nouveau (retour état initial)
                if (!initialInventoryValue)
                {
                    RemoveInventoryValue(homePage);
                }
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_Items_Filter_Group()
        {
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();
            var homePage = LogInAsAdmin();
            bool initialInventoryValue = SetInventoryDayValue(homePage);
            try
            {
                InventoriesPage inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();
                InventoryCreateModalPage inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place, true);
                InventoryItem inventoryItem = inventoryCreateModalpage.Submit();
                string itemName = inventoryItem.GetFirstItemName();
                inventoryItem.SelectFirstItem();
                ItemGeneralInformationPage itemPageItem = inventoryItem.EditItem(itemName);
                string itemGroup = itemPageItem.GetGroupName();
                itemPageItem.Close();
                inventoryItem.Filter(FilterItemType.ByGroup, itemGroup);
                inventoryItem.PageSize("8");
                bool isVerified = inventoryItem.VerifyGroup(itemGroup);
                Assert.IsTrue(isVerified, string.Format(MessageErreur.FILTRE_ERRONE, "'Group'"));
            }
            finally
            {
                if (!initialInventoryValue)
                    RemoveInventoryValue(homePage);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_Items_Filters_SortByName()
        {
            // Prepare
            string site = TestContext.Properties["Production_Site1"].ToString();
            string place = TestContext.Properties["Location"].ToString();
            string ID = "";
            //Arrange
            var homePage = LogInAsAdmin();

            // On va cocher la date du jour dans les settings inventory pour pouvoir lancer le test
            bool initialInventoryValue = SetInventoryDayValue(homePage);
            //Act
            var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            inventoriesPage.ResetFilters();

            try
            {
                // Create
                var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place, true);
                ID = inventoryCreateModalpage.GetInventoryNumber();
                var inventoryItem = inventoryCreateModalpage.Submit();

                // Filter show items with theo/phys/qty only    
                if (!inventoriesPage.isPageSizeEqualsTo100WidhoutTotalNumber())
                {
                    inventoriesPage.PageSize("8");
                    inventoriesPage.PageSize("100");
                }

                inventoryItem.Filter(FilterItemType.SortBy, "Name");
                bool isSorted = inventoryItem.IsSortedByName(); 

                //Assert
                Assert.IsTrue(isSorted, string.Format(MessageErreur.FILTRE_ERRONE, "'Sort by Name'"));
            }
            finally
            {
                // Si au début du test, la case due l'inventory du jour était décochée, on vient la redécocher de nouveau (retour état initial)
                if (!initialInventoryValue)
                {
                    RemoveInventoryValue(homePage);
                }
                inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, ID);
                inventoriesPage.deleteInventory();
                inventoriesPage.ResetFilters();
            }
        }


        [TestMethod]

        [Timeout(_timeout)]
        public void WA_INVE_Items_Filters_SortByItemGroup()
        {
            // Prepare
            string site = TestContext.Properties["Production_Site1"].ToString();
            string place = TestContext.Properties["Location"].ToString();
            string ID = string.Empty;
            //Arrange
            var homePage = LogInAsAdmin();

            // On va cocher la date du jour dans les settings inventory pour pouvoir lancer le test
            bool initialInventoryValue = SetInventoryDayValue(homePage);
            //Act
            var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            inventoriesPage.ResetFilters();

            try
            {
                // Create
                var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place, true);
                ID = inventoryCreateModalpage.GetInventoryNumber();
                var inventoryItem = inventoryCreateModalpage.Submit();

                // Filter show items with theo/phys/qty only    
                if (!inventoriesPage.isPageSizeEqualsTo100WidhoutTotalNumber())
                {
                    inventoriesPage.PageSize("8");
                    inventoriesPage.PageSize("100");
                }

                inventoryItem.Filter(FilterItemType.SortBy, "Group");
                bool isSorted = inventoryItem.IsSortedByGroup();

                //Assert
                Assert.IsTrue(isSorted, string.Format(MessageErreur.FILTRE_ERRONE, "'Sort by Item Group'"));
            }
            finally
            {
                // Si au début du test, la case due l'inventory du jour était décochée, on vient la redécocher de nouveau (retour état initial)
                if (!initialInventoryValue)
                {
                    RemoveInventoryValue(homePage);
                }
                inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, ID);
                inventoriesPage.deleteInventory();
                inventoriesPage.ResetFilters();
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_UpdatePhysicalQuantity()
        {
            //prepare
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();
            var currency = TestContext.Properties["Currency"].ToString();
            string id = "";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // On va cocher la date du jour dans les settings inventory pour pouvoir lancer le test
            bool initialInventoryValue = SetInventoryDayValue(homePage);
            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            //Act
            var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            inventoriesPage.ResetFilters();
            try
            {
                //Create a new Inventory NOT validated
                var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place, true);
                var inventoryDetailPage = inventoryCreateModalpage.Submit();
                id = inventoryDetailPage.GetInventoryNumber();

                // get initial values to compare with
                var initialPhysicalValue = inventoryDetailPage.GetPhysicalValue();
                var initialDifferenceValue = inventoryDetailPage.GetDiffPriceValue(currency);

                // Add physical quantity to item and get new values 
                var itemName = inventoryDetailPage.GetFirstItemName();
                inventoryDetailPage.SelectFirstItem();
                var newValueQty = string.IsNullOrEmpty(initialPhysicalValue) ? 10 : Convert.ToDouble(initialPhysicalValue) + 10;
                var newValueDiff = string.IsNullOrEmpty(initialDifferenceValue) ? 10 : Convert.ToDouble(initialDifferenceValue) + 10;
                inventoryDetailPage.AddPhysicalQuantity(itemName, newValueQty.ToString());
                inventoryDetailPage.SetPrice(itemName, newValueDiff.ToString());

                // Get new values
                inventoryDetailPage.Filter(InventoryItem.FilterItemType.SearchByName, itemName);
                string newPhysValue = inventoryDetailPage.GetPhysicalValue();
                string newDiffValue = inventoryDetailPage.GetDiffPriceValue(currency);

                //Assert 
                Assert.AreEqual(newValueQty.ToString(), newPhysValue, "La quantité n'a pas mis à jour");
                Assert.AreEqual(newValueDiff.ToString(), newDiffValue, "La difference n'a pas mis à jour");

                //Symbol IsEdited
                inventoryDetailPage.SelectFirstItem();
                Assert.IsTrue(inventoryDetailPage.IsEdited(itemName));
            }
            finally
            {
                homePage.GoToWarehouse_InventoriesPage();
                if (!string.IsNullOrEmpty(id))
                {
                    inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, id);
                    inventoriesPage.deleteInventory();
                }
                // Si au début du test, la case due l'inventory du jour était décochée, on vient la redécocher de nouveau (retour état initial)
                if (!initialInventoryValue)
                {
                    RemoveInventoryValue(homePage);
                }
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_UpdatePrice()
        {
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();
            var currency = TestContext.Properties["Currency"].ToString();
            string quantity = "5";

            //Arrange
            LogInAsAdmin();

            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // On va cocher la date du jour dans les settings inventory pour pouvoir lancer le test
            bool initialInventoryValue = SetInventoryDayValue(homePage);

            try
            {
                var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

                //Act
                var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();

                //Create a new Inventory
                var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place, true);
                var inventoryDetailPage = inventoryCreateModalpage.Submit();

                if (inventoryDetailPage.GetTheoricalQuantity() == "0")
                {
                    EnsureTheoricalQuantity();
                }

                inventoryDetailPage.Filter(FilterItemType.ShowItemsWithQtyOnly, true);

                // get initial values to compare with
                double initialTheoricalValue = inventoryDetailPage.GetTheoricalValue(currency, decimalSeparatorValue);
                double initialDifferenceValue = inventoryDetailPage.GetDifferenceValue(currency, decimalSeparatorValue);

                double initialPrice = inventoryDetailPage.GetPrice(currency, decimalSeparatorValue);
                var itemName = inventoryDetailPage.GetFirstItemName();

                inventoryDetailPage.SelectFirstItem();
                inventoryDetailPage.AddPhysicalQuantity(itemName, quantity);
                inventoryDetailPage.SetPrice(itemName, (initialPrice + 1).ToString());
                // data-binding
                inventoryDetailPage.WaitPageLoading();
                inventoryDetailPage.Refresh();

                double newTheoValue = inventoryDetailPage.GetTheoricalValue(currency, decimalSeparatorValue);
                double newDiffValue = inventoryDetailPage.GetDifferenceValue(currency, decimalSeparatorValue);
                double newPrice = inventoryDetailPage.GetPrice(currency, decimalSeparatorValue);

                //Assert 
                Assert.AreEqual(initialPrice + 1, newPrice, "La valeur du Price n'est pas mise à jour correctement");

                Assert.AreEqual(initialTheoricalValue, newTheoValue, "La valeur Theorical value est modifiée");
                Assert.AreNotEqual(initialDifferenceValue, newDiffValue, "La valeur du Difference value n'est pas mise à jour");
                Assert.AreNotEqual(initialPrice, newPrice, "La valeur du Price n'est pas mise à jour");

                inventoryDetailPage.SelectFirstItem();
                bool isEdited = inventoryDetailPage.IsEdited(itemName);
                Assert.IsTrue(isEdited, "Les modifications ne sont pas correctes");
            }
            finally
            {
                // Si au début du test, la case due l'inventory du jour était décochée, on vient la redécocher de nouveau (retour état initial)
                if (!initialInventoryValue)
                {
                    RemoveInventoryValue(homePage);
                }
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_UpdatePhysPackaging()
        {
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();

            // Item  to test 
            string value = "10";

            //Arrange
            var homePage = LogInAsAdmin();

            // On va cocher la date du jour dans les settings inventory pour pouvoir lancer le test
            bool initialInventoryValue = SetInventoryDayValue(homePage);

            try
            {
                //Act
                var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();

                //Create a new Inventory
                var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place, true);
                var inventoryDetailPage = inventoryCreateModalpage.Submit();

                // get initial values to compare with
                string initialPhysPackQty = inventoryDetailPage.GetPhysicalPackagingQuantity();
                string initialTotalQty = inventoryDetailPage.GetTotalPhysQty();

                var itemName = inventoryDetailPage.GetFirstItemName();
                inventoryDetailPage.SelectFirstItem();

                // Add physical packaging 
                inventoryDetailPage.AddPhysicalPackagingQuantity(itemName, value);
                inventoryDetailPage.Refresh();

                string newTotalQty = inventoryDetailPage.GetTotalPhysQty();
                string newPhysPackQty = inventoryDetailPage.GetPhysicalPackagingQuantity();

                //Assert 
                Assert.AreEqual(value, newPhysPackQty);

                Assert.AreNotEqual(initialPhysPackQty, newPhysPackQty);
                Assert.AreNotEqual(initialTotalQty, newTotalQty);

                inventoryDetailPage.SelectFirstItem();
                Assert.IsTrue(inventoryDetailPage.IsEdited(itemName));
            }
            finally
            {
                // Si au début du test, la case due l'inventory du jour était décochée, on vient la redécocher de nouveau (retour état initial)
                if (!initialInventoryValue)
                {
                    RemoveInventoryValue(homePage);
                }
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_AddComment()
        {
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();
            string value = "10";
            string comment = "Comment test";


            //Arrange
            var homePage = LogInAsAdmin();

            // On va cocher la date du jour dans les settings inventory pour pouvoir lancer le test
            bool initialInventoryValue = SetInventoryDayValue(homePage);

            try
            {
                //Act
                var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();

                //Create a new Inventory NOT validated
                var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place, true);
                var inventoryDetailPage = inventoryCreateModalpage.Submit();

                var itemName = inventoryDetailPage.GetFirstItemName();
                inventoryDetailPage.SelectFirstItem();
                inventoryDetailPage.AddPhysicalQuantity(itemName, value);

                //Add comment
                inventoryDetailPage.AddComment(itemName, comment);
                WebDriver.Navigate().Refresh();

                //Assert
                inventoryDetailPage.SelectFirstItem();
                string commentFinale = inventoryDetailPage.GetComment(itemName);
                Assert.AreEqual(comment, commentFinale, "Le commentaire final ne correspond pas au commentaire attendu !");
            }
            finally
            {
                // Si au début du test, la case due l'inventory du jour était décochée, on vient la redécocher de nouveau (retour état initial)
                if (!initialInventoryValue)
                {
                    RemoveInventoryValue(homePage);
                }
            }
        }

        //______________________________________________FIN FILTRES ITEMS INVENTORY______________________________________________

        //______________________________________________VALIDATE INVENTORY_______________________________________________________

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_ValidateInventoryKO()
        {
            // Prepare
            string site = TestContext.Properties["SiteACE"].ToString();
            string place = TestContext.Properties["Place"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            // On va cocher la date du jour dans les settings inventory pour pouvoir lancer le test
            bool initialInventoryValue = SetInventoryDayValue(homePage);

            try
            {
                //Act
                var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();

                // Create
                var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place, true);
                var inventoryItem = inventoryCreateModalpage.Submit();

                // On valide l'inventory créé avant de le filtrer par ItemGroup
                var validateInventory = inventoryItem.Validate();

                bool canUpdateItems = validateInventory.CanUpdateItems();
                //Assert
                Assert.IsFalse(canUpdateItems);
            }
            finally
            {
                // Si au début du test, la case due l'inventory du jour était décochée, on vient la redécocher de nouveau (retour état initial)
                if (!initialInventoryValue)
                {
                    RemoveInventoryValue(homePage);
                }
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_ValidateTotalInventory()
        {
            // Prepare
            string site = "ACE";
            string place = TestContext.Properties["Place"].ToString();

            //Arrange
            LogInAsAdmin();

            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // On va cocher la date du jour dans les settings inventory pour pouvoir lancer le test
            bool initialInventoryValue = SetInventoryDayValue(homePage);

            try
            {
                //Act
                var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();

                //Create a new Inventory
                var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place, true);
                String ID = inventoryCreateModalpage.GetInventoryNumber();
                var inventoryDetailPage = inventoryCreateModalpage.Submit();

                var itemName = inventoryDetailPage.GetFirstItemName();
                inventoryDetailPage.SelectFirstItem();
                inventoryDetailPage.AddPhysicalQuantity(itemName, "10");

                //validate 
                var validateInventory = inventoryDetailPage.Validate();
                validateInventory.ValidateTotalInventory();
                inventoryDetailPage.BackToList();

                inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, ID);

                //Assert
                Assert.AreEqual(ID, inventoriesPage.GetFirstID());
                Assert.IsTrue(inventoriesPage.CheckValidation(true), "Les inventories n'ont pas été validées.");
            }
            finally
            {

                // On créé un autre inventory pour repeupler les premiers items
                CreateOtherInventory(homePage, site, place);

                // Si au début du test, la case due l'inventory du jour était décochée, on vient la redécocher de nouveau (retour état initial)
                if (!initialInventoryValue)
                {
                    RemoveInventoryValue(homePage);
                }
            }
        }

        public void CreateOtherInventory(HomePage homePage, string site, string place)
        {
            var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            // Création d'un nouvel Inventory
            var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
            inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place, true);
            var inventoryDetailPage = inventoryCreateModalpage.Submit();

            // On ajoute une quantité de 10 pour les 5 premiers items de la liste
            inventoryDetailPage.AddPhysicalQtyToInventory("10");
            var validateInventory = inventoryDetailPage.Validate();
            validateInventory.ValidateTotalInventory();
            inventoryDetailPage.BackToList();
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_ValidatePartialInventory()
        {
            // Prepare
            string site = TestContext.Properties["SiteACE"].ToString();
            string place = TestContext.Properties["Place"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            // On va cocher la date du jour dans les settings inventory pour pouvoir lancer le test
            bool initialInventoryValue = SetInventoryDayValue(homePage);

            try
            {
                //Act
                var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();

                // Create a new Inventory
                var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place, true);
                String ID = inventoryCreateModalpage.GetInventoryNumber();
                var inventoryDetailPage = inventoryCreateModalpage.Submit();

                var itemName = inventoryDetailPage.GetFirstItemName();
                inventoryDetailPage.SelectFirstItem();
                inventoryDetailPage.AddPhysicalQuantity(itemName, "10");

                //validate 
                var validateInventory = inventoryDetailPage.Validate();
                validateInventory.ValidatePartialInventory();
                inventoryDetailPage.BackToList();

                inventoriesPage.Filter(InventoriesPage.FilterType.ShowItemsNotValidated, false);
                inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, ID);

                //Assert
                Assert.AreEqual(ID, inventoriesPage.GetFirstID());
                Assert.IsTrue(inventoriesPage.CheckValidation(true));
            }
            finally
            {
                // Si au début du test, la case due l'inventory du jour était décochée, on vient la redécocher de nouveau (retour état initial)
                if (!initialInventoryValue)
                {
                    RemoveInventoryValue(homePage);
                }
            }
        }

        //______________________________________________FIN VALIDATE INVENTORY___________________________________________________

        //______________________________________________REFRESH INVENTORY________________________________________________________

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_RefreshInventory()
        {
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();

            // Item  to test 
            string value = "10";

            //Arrange
            LogInAsAdmin();

            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // On va cocher la date du jour dans les settings inventory pour pouvoir lancer le test
            bool initialInventoryValue = SetInventoryDayValue(homePage);

            try
            {
                //Act
                var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();

                //Create a new Inventory NOT validated
                var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place, true);
                var inventoryDetailPage = inventoryCreateModalpage.Submit();

                if (inventoryDetailPage.GetTheoricalQuantity() == "0")
                {
                    EnsureTheoricalQuantity();
                }

                inventoryDetailPage.Filter(FilterItemType.ShowItemsWithQtyOnly, true);

                var initPhyQye = inventoryDetailPage.GetPhysicalQuantity();

                // Add physical quantity  
                var itemName = inventoryDetailPage.GetFirstItemName();
                inventoryDetailPage.SelectFirstItem();
                inventoryDetailPage.AddPhysicalQuantity(itemName, value);
                // data binding

                inventoryDetailPage.Refresh();

                var newPhyQye = inventoryDetailPage.GetPhysicalQuantity();

                //Assert 
                Assert.AreNotEqual(initPhyQye, newPhyQye, "Les valeurs ne sont pas rafraichies.");
                Assert.AreEqual(value, newPhyQye, "Les valeurs ne sont pas rafraichies correctement.");
            }
            finally
            {
                // Si au début du test, la case due l'inventory du jour était décochée, on vient la redécocher de nouveau (retour état initial)
                if (!initialInventoryValue)
                {
                    RemoveInventoryValue(homePage);
                }
            }
        }

        //______________________________________________FIN REFRESH INVENTORY____________________________________________________

        //______________________________________________EXPORT INVENTORY_________________________________________________________

        /*
         * Export de la liste des items d'un inventory avec newVersionPrint = true
        */
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_ExportInventoryItemsNewVersion()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();

            bool newVersionPrint = true;

            // Log in
            var homePage = LogInAsAdmin();
            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            // On va cocher la date du jour dans les settings inventory pour pouvoir lancer le test
            bool initialInventoryValue = SetInventoryDayValue(homePage);

            try
            {
                //Act
                var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();

                if (newVersionPrint)
                {
                    inventoriesPage.ClearDownloads();
                }

                //Create a new Inventory
                var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place, true);
                string inventoryNumber = inventoryCreateModalpage.GetInventoryNumber();
                var inventoryDetailPage = inventoryCreateModalpage.Submit();

                // Add physical quantity to item to enable the comment 
                var itemName = inventoryDetailPage.GetFirstItemName();
                inventoryDetailPage.SelectFirstItem();
                inventoryDetailPage.AddPhysicalQuantity(itemName, "10");

                var validateInventory = inventoryDetailPage.Validate();
                validateInventory.ValidatePartialInventory();
                DeleteAllFileDownload();
                // Export de la liste des Items
                FileInfo trouve = ExportGeneriqueItems(inventoryDetailPage, newVersionPrint);
                inventoryDetailPage.PageSize("100");
                inventoryDetailPage.CheckExport(trouve, site, inventoryNumber, decimalSeparatorValue);
            }
            finally
            {
                // Si au début du test, la case due l'inventory du jour était décochée, on vient la redécocher de nouveau (retour état initial)
                if (!initialInventoryValue)
                {
                    RemoveInventoryValue(homePage);
                }
            }
        }

        private FileInfo ExportGeneriqueItems(InventoryItem inventoryItem, bool printValue)
        {
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            DeleteAllFileDownload();

            // On exporte les résultats sous la forme d'un fichier Excel (dont on récupère le nom)
            inventoryItem.ExportExcelFile(printValue, downloadsPath);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = inventoryItem.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

            // Récupération du nom du fichier et construction de l'URL du fichier Excel à ouvrir   
            var fileName = correctDownloadedFile.Name;
            var filePath = System.IO.Path.Combine(downloadsPath, fileName);

            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber("Inventory", filePath);

            //Assert
            Assert.AreNotEqual(0, resultNumber);

            return new FileInfo(filePath);
        }

        /*
         * Export de la liste des items d'un inventory avec newVersionPrint = true
        */
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_ExportInventoryResultsNewVersion()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            bool newVersionPrint = true;

            // Log in
            var homePage = LogInAsAdmin();

            // On va cocher la date du jour dans les settings inventory pour pouvoir lancer le test
            bool initialInventoryValue = SetInventoryDayValue(homePage);

            try
            {
                //Act
                var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();

                if (newVersionPrint)
                {
                    inventoriesPage.ClearDownloads();
                }

                inventoriesPage.ResetFilters();
                inventoriesPage.Filter(InventoriesPage.FilterType.DateFrom, DateUtils.Now.AddDays(-1));
                inventoriesPage.Filter(InventoriesPage.FilterType.DateTo, DateUtils.Now);

                if (inventoriesPage.CheckTotalNumber() < 20)
                {
                    //Create a new Inventory
                    var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                    inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place, true);
                    var inventoryDetailPage = inventoryCreateModalpage.Submit();

                    // Add physical quantity to item to enable the comment 
                    var itemName = inventoryDetailPage.GetFirstItemName();
                    inventoryDetailPage.SelectFirstItem();
                    inventoryDetailPage.AddPhysicalQuantity(itemName, "10");

                    var validateInventory = inventoryDetailPage.Validate();
                    validateInventory.ValidatePartialInventory();
                    inventoriesPage = inventoryDetailPage.BackToList();

                    inventoriesPage.ResetFilters();
                    inventoriesPage.Filter(InventoriesPage.FilterType.DateFrom, DateUtils.Now.AddDays(-1));
                    inventoriesPage.Filter(InventoriesPage.FilterType.DateTo, DateUtils.Now);
                }
                DeleteAllFileDownload();
                // On exporte les résultats sous la forme d'un fichier Excel (dont on récupère le nom)
                inventoriesPage.ExportExcelFile(newVersionPrint);

                // On récupère les fichiers du répertoire de téléchargement
                DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
                FileInfo[] taskFiles = taskDirectory.GetFiles();

                // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
                var correctDownloadedFile = inventoriesPage.GetExportExcelFile(taskFiles);
                Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

                // Récupération du nom du fichier et construction de l'URL du fichier Excel à ouvrir   
                var fileName = correctDownloadedFile.Name;
                var filePath = System.IO.Path.Combine(downloadsPath, fileName);

                // Exploitation du fichier Excel
                int resultNumber = OpenXmlExcel.GetExportResultNumber("Inventory", filePath);

                //Assert
                Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
                inventoriesPage.PageSize("1000");
                inventoriesPage.CheckExport(new FileInfo(filePath));
            }
            finally
            {
                // Si au début du test, la case due l'inventory du jour était décochée, on vient la redécocher de nouveau (retour état initial)
                if (!initialInventoryValue)
                {
                    RemoveInventoryValue(homePage);
                }
            }
        }

        //______________________________________________FIN EXPORT INVENTORY_____________________________________________________

        //______________________________________________PRINT INVENTORY__________________________________________________________

        /*
         * Test d'impression des items d'un inventory avec newVersionPrint = true
        */
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_PrintInventoryNewVersion()
        {
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();
            string delivery = TestContext.Properties["PlaceFrom"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNamePdfBegin = "Print Report_-_";
            string DocFileNameZipBegin = "All_files_";
            bool newVersionPrint = true;
            HomePage homePage = LogInAsAdmin();
            bool initialInventoryValue = SetInventoryDayValue(homePage);
            try
            {
                InventoriesPage inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();
                if (newVersionPrint) inventoriesPage.ClearDownloads();
                InventoryCreateModalPage inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place, true);
                InventoryItem inventoryDetailPage = inventoryCreateModalpage.Submit();
                if (inventoryDetailPage.GetTheoricalQuantity() == "0")
                    EnsureTheoricalQuantity();
                inventoryDetailPage.Filter(FilterItemType.ShowItemsWithQtyOnly, true);
                string itemName = inventoryDetailPage.GetFirstItemName();
                inventoryDetailPage.AddPhysicalQuantityByThQty("0");
                inventoryDetailPage.Refresh();
                PrintReportPage reportPage = inventoryDetailPage.PrintInventoryItems(newVersionPrint);
                bool isReportNotGenerated = reportPage.IsReportGenerated();
                reportPage.Close();
                Assert.IsFalse(isReportNotGenerated, "Le fichier ne doit pas générer de fichier lorsque la valeur physical quantity est 0");
                inventoryDetailPage.AddPhysicalQuantity(itemName, "10");
                inventoryDetailPage.Refresh();
                reportPage = inventoryDetailPage.PrintInventoryItems(newVersionPrint);
                bool isReportGenerated = reportPage.IsReportGenerated();
                reportPage.Close();
                Assert.IsTrue(isReportGenerated, "L'application n'a pas pu générer le fichier attendu.");
                InventoryGeneralInformation generalInfo = inventoryDetailPage.ClickOnGeneralInformationTab();
                string inventoryNumber = generalInfo.GetInventoryNumber();
                reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
                string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
                FileInfo fi = new FileInfo(trouve);
                fi.Refresh();
                Assert.IsTrue(fi.Exists, trouve + " non généré");
                PdfDocument document = PdfDocument.Open(fi.FullName);
                Page page1 = document.GetPage(1);
                IEnumerable<Word> words = page1.GetWords();
                int verifySite = words.Count(w => w.Text == site + "(" + site + ")");
                Assert.AreEqual(1, verifySite, site + "(" + site + ") non présent dans le Pdf");
                int verifyDate = words.Count(w => w.Text == DateUtils.Now.Date.ToString("dd/MM/yyyy"));
                Assert.AreEqual(1, verifyDate, DateUtils.Now.Date.ToString("dd/MM/yyyy") + " non présent dans le Pdf");
                int verifyInventoryNumber = words.Count(w => w.Text == inventoryNumber);
                Assert.AreEqual(1, verifyInventoryNumber, inventoryNumber + " non présent dans le Pdf");
            }
            finally
            {
                if (!initialInventoryValue)
                    RemoveInventoryValue(homePage);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_PrintPreparationSheetInventoryNewVersion()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            //Print Preparation Sheet_-_437549_-_20220316103924.pdf
            string DocFileNamePdfBegin = "Print Preparation Sheet_-_";
            //All_files_20220225_102148.zip
            string DocFileNameZipBegin = "All_files_";

            bool newVersionPrint = true;

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // On va cocher la date du jour dans les settings inventory pour pouvoir lancer le test
            bool initialInventoryValue = SetInventoryDayValue(homePage);

            try
            {
                //Act
                var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();

                if (newVersionPrint)
                {
                    inventoriesPage.ClearDownloads();
                }

                //Create a new Inventory
                var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place, true);
                var inventoryDetailPage = inventoryCreateModalpage.Submit();

                // Add physical quantity to item to enable the comment 
                var itemName = inventoryDetailPage.GetFirstItemName();
                inventoryDetailPage.SelectFirstItem();
                inventoryDetailPage.AddPhysicalQuantity(itemName, "10");
                inventoryDetailPage.Refresh();

                var reportPage = inventoryDetailPage.PrintPreparationSheet(newVersionPrint);
                var isReportGenerated = reportPage.IsReportGenerated();
                reportPage.Close();

                //Assert  
                Assert.IsTrue(isReportGenerated);

                reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);

                InventoryGeneralInformation generalInfo = inventoryDetailPage.ClickOnGeneralInformationTab();
                string inventoryNumber = generalInfo.GetInventoryNumber();

                // cliquer sur All
                string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
                FileInfo fi = new FileInfo(trouve);
                fi.Refresh();
                Assert.IsTrue(fi.Exists, trouve + " non généré");

                PdfDocument document = PdfDocument.Open(fi.FullName);
                Page page1 = document.GetPage(1);
                IEnumerable<Word> words = page1.GetWords();
                Assert.AreEqual(1, words.Count(w => w.Text == site + "(" + site + ")"), site + "(" + site + ") non présent dans le Pdf");
                Assert.AreEqual(1, words.Count(w => w.Text == DateUtils.Now.Date.ToString("dd/MM/yyyy")), DateUtils.Now.Date.ToString("dd/MM/yyyy") + " non présent dans le Pdf");
                Assert.AreEqual(1, words.Count(w => w.Text == inventoryNumber), inventoryNumber + " non présent dans le Pdf");
            }
            finally
            {
                // Si au début du test, la case due l'inventory du jour était décochée, on vient la redécocher de nouveau (retour état initial)
                if (!initialInventoryValue)
                {
                    RemoveInventoryValue(homePage);
                }
            }
        }

        //______________________________________________FIN PRINT INVENTORY______________________________________________________

        //______________________________________________GENERAL INFOS INVENTORY__________________________________________________

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_InventoryGeneralInformation()
        {
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            // On va cocher la date du jour dans les settings inventory pour pouvoir lancer le test
            bool initialInventoryValue = SetInventoryDayValue(homePage);

            try
            {

                //Act
                var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();

                //Create a new Inventory NOT validated
                var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place, true);
                String ID = inventoryCreateModalpage.GetInventoryNumber();
                var inventoryDetailPage = inventoryCreateModalpage.Submit();

                var itemName = inventoryDetailPage.GetFirstItemName();
                inventoryDetailPage.SelectFirstItem();
                inventoryDetailPage.AddPhysicalQuantity(itemName, "10");

                InventoryGeneralInformation generalInfo = inventoryDetailPage.ClickOnGeneralInformationTab();

                //2. Ajouter un commentaire et vérifier sa prise à compte
                string comment = "Mon commentaire";
                generalInfo.SetComment(comment);
                inventoryDetailPage = generalInfo.ClickOnItems();
                generalInfo = inventoryDetailPage.ClickOnGeneralInformationTab();
                Assert.AreEqual(comment, generalInfo.GetComment(), "commentaire non sauvegardé");

                inventoriesPage = generalInfo.BackToList();

                inventoriesPage.ResetFilters();
                inventoriesPage.Filter(InventoriesPage.FilterType.ShowActive, true);
                inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, ID);
                Assert.AreEqual(ID, inventoriesPage.GetFirstID());

                //Vérifier sur l'index que l'icone commentaire est vert pour cette inventory...
                Assert.IsTrue(inventoriesPage.IsCommentBubbleGreen(), "Comment bubble non vert");
                //...et le commentaire en survolant cet icone
                Assert.AreEqual(comment, inventoriesPage.GetCommentBubble(), "Commentaire non pris en compte");

                inventoryDetailPage = inventoriesPage.SelectFirstInventory();
                var inventoryGeneralInformation = inventoryDetailPage.ClickOnGeneralInformationTab();

                // Assert 
                Assert.AreEqual(inventoryGeneralInformation.GetInventoryNumber(), ID);
                Assert.AreEqual(inventoryGeneralInformation.GetInventorySite(), site);
                Assert.AreEqual(inventoryGeneralInformation.GetInventoryPlace(), place);

                // Deactivate 
                inventoryGeneralInformation.ChangeStatusInventory(false);
                Assert.IsFalse(inventoryGeneralInformation.CanValidate());

                // Check if the inventory is deactivated 
                inventoriesPage = inventoryGeneralInformation.BackToList();
                inventoriesPage.ResetFilters();
                inventoriesPage.Filter(InventoriesPage.FilterType.ShowInactive, true);
                inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, ID);
                Assert.AreEqual(ID, inventoriesPage.GetFirstID());
            }
            finally
            {
                // Si au début du test, la case due l'inventory du jour était décochée, on vient la redécocher de nouveau (retour état initial)
                if (!initialInventoryValue)
                {
                    RemoveInventoryValue(homePage);
                }
            }
        }

        // __________________________________________________ EXPORT_SAGE _____________________________________________________

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_SageManual_Details_ExportSAGEKONewVersion()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();
            string itemName = TestContext.Properties["Item_Inventory"].ToString();
            string journalInventory = TestContext.Properties["Journal_Inventory"].ToString();

            bool newVersionPrint = true;

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // On va cocher la date du jour dans les settings inventory pour pouvoir lancer le test
            bool initialInventoryValue = SetInventoryDayValue(homePage);

            // Config export sage manuel
            homePage.SetSageAutoEnabled(site, false);

            try
            {
                // Config journal KO pour le test
                Assert.IsTrue(VerifyAccountingJournal(homePage, site, ""));

                //Act
                var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();

                if (newVersionPrint)
                {
                    inventoriesPage.ClearDownloads();
                }

                // Create
                var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place);
                string ID = inventoryCreateModalpage.GetInventoryNumber();
                var inventoryItem = inventoryCreateModalpage.Submit();

                inventoryItem.Filter(FilterItemType.SearchByName, itemName);
                inventoryItem.SelectFirstItem();
                inventoryItem.AddPhysicalQuantity(itemName, "2");

                // On valide l'inventory créé
                var validateInventory = inventoryItem.Validate();
                validateInventory.ValidatePartialInventory();

                inventoryItem.ManualExportSageError(newVersionPrint);
                inventoriesPage = inventoryItem.BackToList();
                inventoriesPage.ResetFilters();
                inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, ID);

                Assert.IsFalse(inventoriesPage.IsSentToSAGE(), "L'inventory est notifiée comme envoyée vers le SAGE malgré l'erreur.");

            }
            finally
            {
                VerifyAccountingJournal(homePage, site, journalInventory);

                // Si au début du test, la case due l'inventory du jour était décochée, on vient la redécocher de nouveau (retour état initial)
                if (!initialInventoryValue)
                {
                    RemoveInventoryValue(homePage);
                }
            }

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_SageManual_Details_ExportSAGENewVersion()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();
            string itemName = TestContext.Properties["Item_Inventory"].ToString();
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string currency = TestContext.Properties["Currency"].ToString();
            string journalInventory = TestContext.Properties["Journal_Inventory"].ToString();

            bool newVersionPrint = true;

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            // On va cocher la date du jour dans les settings inventory pour pouvoir lancer le test
            bool initialInventoryValue = SetInventoryDayValue(homePage);

            // Config export sage manuel
            homePage.SetSageAutoEnabled(site, false);

            try
            {
                // Config journal KO pour le test
                VerifyAccountingJournal(homePage, site, journalInventory);

                //Act
                var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();

                if (newVersionPrint)
                {
                    inventoriesPage.ClearDownloads();
                }

                // Create
                var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place);
                String ID = inventoryCreateModalpage.GetInventoryNumber();
                var inventoryItem = inventoryCreateModalpage.Submit();

                inventoryItem.Filter(FilterItemType.SearchByName, itemName);
                inventoryItem.SelectFirstItem();
                inventoryItem.AddPhysicalQuantity(itemName, "2");

                // On valide l'inventory créé
                var validateInventory = inventoryItem.Validate();
                validateInventory.ValidateTotalInventory();

                var inventoryFooterPage = inventoryItem.ClickOnFooterTab();
                double montantInventory = inventoryFooterPage.GetInventoryTotalHT(currency, decimalSeparatorValue);

                inventoryItem.ExportForSage(newVersionPrint);

                // On récupère les fichiers du répertoire de téléchargement
                DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
                FileInfo[] taskFiles = taskDirectory.GetFiles();

                var correctDownloadedFile = inventoryItem.GetExportSAGEFile(taskFiles);
                Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

                // Récupération du nom du fichier et construction de l'URL du fichier txt à exploiter   
                var fileName = correctDownloadedFile.Name;
                var filePath = System.IO.Path.Combine(downloadsPath, fileName);

                // On n'exploite que les lignes avec contenu "général" --> "G"
                double[] contenuFichier = ExploitTextFiles.VerifySAGEFileContent(filePath, "G", decimalSeparatorValue);

                inventoriesPage = inventoryItem.BackToList();
                inventoriesPage.ResetFilters();
                inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, ID);

                Assert.AreEqual(contenuFichier[0].ToString(), contenuFichier[1].ToString(), "Le contenu du fichier SAGE est incorrect.");

                // Remarque : pour les Inventory, le montant issu du fichier SAGE est égal à 2 fois le montant HT de l'inventory
                Assert.AreEqual((2 * montantInventory).ToString(), contenuFichier[0].ToString(), "Le montant issu du fichier SAGE n'est pas égal au montant de l'Inventory défini dans l'application.");
                Assert.IsTrue(inventoriesPage.IsSentToSAGE(), "L'inventory n'est pas notifiée comme envoyée vers le SAGE.");
            }
            finally
            {
                VerifyAccountingJournal(homePage, site, journalInventory);

                // Si au début du test, la case due l'inventory du jour était décochée, on vient la redécocher de nouveau (retour état initial)
                if (!initialInventoryValue)
                {
                    RemoveInventoryValue(homePage);
                }
            }

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_Details_EnableSAGEExport()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();
            string group = TestContext.Properties["Item_Group"].ToString();

            bool newVersionPrint = true;

            //Arrange
            var homePage = LogInAsAdmin();

            // Config export sage manuel
            homePage.SetSageAutoEnabled(site, false);

            // On va cocher la date du jour dans les settings inventory pour pouvoir lancer le test
            bool initialInventoryValue = SetInventoryDayValue(homePage);

            try
            {
                //Act
                var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();

                if (newVersionPrint)
                {
                    inventoriesPage.ClearDownloads();
                }

                // Create
                var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place);
                var inventoryItem = inventoryCreateModalpage.Submit();
                inventoryItem.ResetFilter();
                inventoryItem.Filter(FilterItemType.ByGroup, group);

                var itemName = inventoryItem.GetFirstItemName();
                inventoryItem.SelectFirstItem();
                inventoryItem.AddPhysicalQuantity(itemName, "10");

                // On valide l'inventory créé
                var validateInventory = inventoryItem.Validate();
                validateInventory.ValidatePartialInventory();

                bool isEnableOK = inventoryItem.CanClickOnEnableSAGE();
                Assert.IsFalse(isEnableOK, "Il est possible de cliquer sur la fonctionnalité 'Enable export for SAGE' "
                    + "pour un inventory non envoyé au SAGE.");

                inventoryItem.ExportForSage(newVersionPrint);
                inventoryItem.ClosePrintButton();

                bool isSAGEOK = inventoryItem.CanClickOnSAGE();
                isEnableOK = inventoryItem.CanClickOnEnableSAGE();

                Assert.IsFalse(isSAGEOK, "Il est possible de cliquer sur la fonctionnalité 'Export for SAGE' "
                    + "après avoir réalisé un export SAGE.");

                Assert.IsTrue(isEnableOK, "Il est impossible de cliquer sur la fonctionnalité 'Enable export for SAGE' "
                    + "pour un invoice envoyé au SAGE.");

                inventoryItem.EnableExportForSage();

                isSAGEOK = inventoryItem.CanClickOnSAGE();
                isEnableOK = inventoryItem.CanClickOnEnableSAGE();

                Assert.IsTrue(isSAGEOK, "Il est impossible de cliquer sur la fonctionnalité 'Export for SAGE' "
                    + "après avoir cliqué un export SAGE.");

                Assert.IsFalse(isEnableOK, "Il est possible de cliquer sur la fonctionnalité 'Enable export for SAGE' "
                    + "pour un inventory envoyé au SAGE.");

            }
            finally
            {
                // Si au début du test, la case due l'inventory du jour était décochée, on vient la redécocher de nouveau (retour état initial)
                if (!initialInventoryValue)
                {
                    RemoveInventoryValue(homePage);
                }
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_SageManual_Index_ExportSAGEKONewVersion()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();
            string itemName = TestContext.Properties["Item_Inventory"].ToString();
            string journalInventory = TestContext.Properties["Journal_Inventory"].ToString();

            bool newVersionPrint = true;

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // On va cocher la date du jour dans les settings inventory pour pouvoir lancer le test
            bool initialInventoryValue = SetInventoryDayValue(homePage);

            // Config export sage manuel
            homePage.SetSageAutoEnabled(site, false);

            try
            {
                // Config journal KO pour le test
                VerifyAccountingJournal(homePage, site, "");

                //Act
                var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();

                if (newVersionPrint)
                {
                    inventoriesPage.ClearDownloads();
                }

                // Create
                var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place);
                String ID = inventoryCreateModalpage.GetInventoryNumber();
                var inventoryItem = inventoryCreateModalpage.Submit();

                inventoryItem.Filter(FilterItemType.SearchByName, itemName);
                inventoryItem.SelectFirstItem();
                inventoryItem.AddPhysicalQuantity(itemName, "2");

                // On valide l'inventory créé
                var validateInventory = inventoryItem.Validate();
                validateInventory.ValidateTotalInventory();

                inventoriesPage = inventoryItem.BackToList();
                inventoriesPage.ResetFilters();
                inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, ID);

                inventoriesPage.ManualExportSageError(newVersionPrint);

                WebDriver.Navigate().Refresh();

                Assert.IsFalse(inventoriesPage.IsSentToSAGE(), "L'inventory est notifiée comme envoyée vers le SAGE malgré l'erreur.");

            }
            finally
            {
                VerifyAccountingJournal(homePage, site, journalInventory);

                // Si au début du test, la case due l'inventory du jour était décochée, on vient la redécocher de nouveau (retour état initial)
                if (!initialInventoryValue)
                {
                    RemoveInventoryValue(homePage);
                }
            }

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_SageManual_Index_ExportSAGENewVersion()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();
            string itemName = TestContext.Properties["Item_Inventory"].ToString();
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string currency = TestContext.Properties["Currency"].ToString();
            string journalInventory = TestContext.Properties["Journal_Inventory"].ToString();

            bool newVersionPrint = true;

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            // On va cocher la date du jour dans les settings inventory pour pouvoir lancer le test
            bool initialInventoryValue = SetInventoryDayValue(homePage);

            // Config export sage manuel
            homePage.SetSageAutoEnabled(site, false);

            VerifyAccountingJournal(homePage, site, journalInventory);

            try
            {
                //Act
                var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();

                if (newVersionPrint)
                {
                    inventoriesPage.ClearDownloads();
                }

                // Create
                var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place);
                String ID = inventoryCreateModalpage.GetInventoryNumber();
                var inventoryItem = inventoryCreateModalpage.Submit();

                inventoryItem.Filter(FilterItemType.SearchByName, itemName);
                inventoryItem.SelectFirstItem();
                inventoryItem.AddPhysicalQuantity(itemName, "2");

                // On valide l'inventory créé
                var validateInventory = inventoryItem.Validate();
                validateInventory.ValidateTotalInventory();

                var inventoryFooterPage = inventoryItem.ClickOnFooterTab();
                double montantInventory = inventoryFooterPage.GetInventoryTotalHT(currency, decimalSeparatorValue);

                inventoriesPage = inventoryItem.BackToList();
                inventoriesPage.ResetFilters();
                inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, ID);

                inventoriesPage.ManualExportResultsForSage(newVersionPrint);

                // On récupère les fichiers du répertoire de téléchargement
                DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
                FileInfo[] taskFiles = taskDirectory.GetFiles();

                var correctDownloadedFile = inventoriesPage.GetExportForSageFile(taskFiles);
                Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

                // Récupération du nom du fichier et construction de l'URL du fichier txt à exploiter   
                var fileName = correctDownloadedFile.Name;
                var filePath = System.IO.Path.Combine(downloadsPath, fileName);

                // On n'exploite que les lignes avec contenu "général" --> "G"
                double[] contenuFichier = ExploitTextFiles.VerifySAGEFileContent(filePath, "G", decimalSeparatorValue);

                Assert.AreEqual(contenuFichier[0].ToString(), contenuFichier[1].ToString(), "Le contenu du fichier SAGE est incorrect.");

                // Remarque : pour les Inventory, le montant issu du fichier SAGE est égal à 2 fois le montant HT de l'inventory
                Assert.AreEqual((2 * montantInventory).ToString(), contenuFichier[0].ToString(), "Le montant issu du fichier SAGE n'est pas égal au montant de l'Inventory défini dans l'application.");

                WebDriver.Navigate().Refresh();
                inventoriesPage.ResetFilters();
                inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, ID);
                Assert.IsTrue(inventoriesPage.IsSentToSAGE(), "L'inventory n'est pas notifiée comme envoyée vers le SAGE.");
            }
            finally
            {
                // Si au début du test, la case due l'inventory du jour était décochée, on vient la redécocher de nouveau (retour état initial)
                if (!initialInventoryValue)
                {
                    RemoveInventoryValue(homePage);
                }
            }

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_Index_EnableSAGEExport()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();
            string itemName = TestContext.Properties["Item_Inventory"].ToString();

            bool newVersionPrint = true;

            //Arrange
            var homePage = LogInAsAdmin();

            // Config export sage manuel
            homePage.SetSageAutoEnabled(site, false);

            // On va cocher la date du jour dans les settings inventory pour pouvoir lancer le test
            bool initialInventoryValue = SetInventoryDayValue(homePage);

            try
            {
                //Act
                var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();

                if (newVersionPrint)
                {
                    inventoriesPage.ClearDownloads();
                    DeleteAllFileDownload();
                }

                // Create
                var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place);
                string ID = inventoryCreateModalpage.GetInventoryNumber();
                var inventoryItem = inventoryCreateModalpage.Submit();

                inventoryItem.Filter(FilterItemType.SearchByName, itemName);
                inventoryItem.SelectFirstItem();
                inventoryItem.AddPhysicalQuantity(itemName, "2");

                // On valide l'inventory créé
                var validateInventory = inventoryItem.Validate();
                validateInventory.ValidatePartialInventory();

                inventoriesPage = inventoryItem.BackToList();
                inventoriesPage.ResetFilters();
                inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, ID);

                inventoriesPage.ManualExportResultsForSage(newVersionPrint);
                inventoryItem = inventoriesPage.SelectFirstInventory();

                Assert.IsTrue(inventoryItem.CanClickOnEnableSAGE(), "Il n'est pas possible de cliquer sur la fonctionnalité 'Enable export for SAGE' "
                + "après avoir exporté l'inventory vers SAGE depuis la page Index.");

                Assert.IsFalse(inventoryItem.CanClickOnSAGE(), "Il est possible de cliquer sur la fonctionnalité 'Export for SAGE' "
                    + "après avoir exporté l'inventory vers SAGE depuis la page Index.");

                inventoriesPage = inventoryItem.BackToList();
                inventoriesPage.EnableExportForSage();
                inventoryItem = inventoriesPage.SelectFirstInventory();

                Assert.IsFalse(inventoryItem.CanClickOnEnableSAGE(), "Il est possible de cliquer sur la fonctionnalité 'Enable export for SAGE' "
                + "après avoir activé la fonctionnalité depuis la page Index.");

                Assert.IsTrue(inventoryItem.CanClickOnSAGE(), "Il est impossible de cliquer sur la fonctionnalité 'Export for SAGE' "
                    + "après avoir cliqué sur 'Enable export for SAGE' depuis la page Index.");
            }
            finally
            {
                // Si au début du test, la case due l'inventory du jour était décochée, on vient la redécocher de nouveau (retour état initial)
                if (!initialInventoryValue)
                {
                    RemoveInventoryValue(homePage);
                }
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_Details_GenerateSageTxt()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();
            string itemName = TestContext.Properties["Item_Inventory"].ToString();
            var currency = TestContext.Properties["Currency"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            bool newVersionPrint = true;

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            try
            {
                var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();

                if (newVersionPrint)
                {
                    inventoriesPage.ClearDownloads();
                }

                // Create
                var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place);
                string ID = inventoryCreateModalpage.GetInventoryNumber();
                var inventoryItem = inventoryCreateModalpage.Submit();

                inventoryItem.Filter(FilterItemType.SearchByName, itemName);
                inventoryItem.SelectFirstItem();
                inventoryItem.AddPhysicalQuantity(itemName, "2");

                // On valide l'inventory créé
                var validateInventory = inventoryItem.Validate();
                validateInventory.ValidateTotalInventory();

                var inventoryFooter = inventoryItem.ClickOnFooterTab();
                double montantInventory = inventoryFooter.GetInventoryTotalHT(currency, decimalSeparatorValue);

                // Calcul du montant de la facture transmise à TL
                var inventoryAccounting = inventoryFooter.ClickOnAccounting();
                double montantGlobal = inventoryAccounting.GetInventoryGrossAmount("G", decimalSeparatorValue);
                double montantDetaille = inventoryAccounting.GetInventoryDetailAmount("G", decimalSeparatorValue);

                inventoryItem = inventoryAccounting.ClickOnItems();
                inventoryItem.ExportForSage(newVersionPrint);

                // On récupère les fichiers du répertoire de téléchargement
                DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
                FileInfo[] taskFiles = taskDirectory.GetFiles();

                var correctDownloadedFile = inventoryItem.GetExportSAGEFile(taskFiles);
                Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

                // Récupération du nom du fichier et construction de l'URL du fichier txt à exploiter   
                var fileName = correctDownloadedFile.Name;
                var filePath = System.IO.Path.Combine(downloadsPath, fileName);

                // On n'exploite que les lignes avec contenu "général" --> "G"
                double[] contenuFichier = ExploitTextFiles.VerifySAGEFileContent(filePath, "G", decimalSeparatorValue);

                inventoriesPage = inventoryItem.BackToList();
                inventoriesPage.ResetFilters();
                inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, ID);

                Assert.AreEqual(contenuFichier[0].ToString(), contenuFichier[1].ToString(), "Le contenu du fichier SAGE généré est incorrect.");
                Assert.AreEqual((2 * montantInventory).ToString(), contenuFichier[0].ToString(), "Le montant issu du fichier SAGE n'est pas égal au montant de la RN défini dans l'application.");
                Assert.AreEqual(montantGlobal.ToString(), montantDetaille.ToString(), "Les montants AmountDebit et AmountCredit de la RN envoyée vers SAGE ne sont pas les mêmes dans l'onglet Accounting.");
                Assert.AreEqual((2 * montantInventory).ToString(), montantGlobal.ToString(), "Le montant issu du fichier SAGE n'est pas égal au montant de la RN défini dans l'application.");
                Assert.IsTrue(inventoriesPage.isAccounted(), "L'inventory n'a pas été envoyée manuellement vers le SAGE.");
            }
            finally
            {
                homePage.SetSageAutoEnabled(site, false);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_Index_GenerateSageTxt()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();
            string itemName = TestContext.Properties["Item_Inventory"].ToString();
            var currency = TestContext.Properties["Currency"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            bool newVersionPrint = true;

            // Log in
            var homePage = LogInAsAdmin();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            // Config pour export SAGE auto mais pas pour le site
            homePage.SetSageAutoEnabled(site, true, "Inventory", false);

            try
            {
                var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();

                if (newVersionPrint)
                {
                    inventoriesPage.ClearDownloads();
                }

                // Create
                var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place);
                string ID = inventoryCreateModalpage.GetInventoryNumber();
                var inventoryItem = inventoryCreateModalpage.Submit();

                inventoryItem.Filter(FilterItemType.SearchByName, itemName);
                inventoryItem.SelectFirstItem();
                inventoryItem.AddPhysicalQuantity(itemName, "2");

                // On valide l'inventory créé
                var validateInventory = inventoryItem.Validate();
                validateInventory.ValidateTotalInventory();

                var inventoryFooter = inventoryItem.ClickOnFooterTab();
                double montantInventory = inventoryFooter.GetInventoryTotalHT(currency, decimalSeparatorValue);

                // Calcul du montant de la facture transmise à TL
                var inventoryAccounting = inventoryFooter.ClickOnAccounting();
                double montantGlobal = inventoryAccounting.GetInventoryGrossAmount("G", decimalSeparatorValue);
                double montantDetaille = inventoryAccounting.GetInventoryDetailAmount("G", decimalSeparatorValue);

                inventoriesPage = inventoryAccounting.BackToList();
                inventoriesPage.ResetFilters();
                inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, ID);

                inventoriesPage.GenerateSageTxt();

                // On récupère les fichiers du répertoire de téléchargement
                DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
                FileInfo[] taskFiles = taskDirectory.GetFiles();

                var correctDownloadedFile = inventoryItem.GetExportSAGEFile(taskFiles);
                Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

                // Récupération du nom du fichier et construction de l'URL du fichier txt à exploiter   
                var fileName = correctDownloadedFile.Name;
                var filePath = System.IO.Path.Combine(downloadsPath, fileName);

                // On n'exploite que les lignes avec contenu "général" --> "G"
                double[] contenuFichier = ExploitTextFiles.VerifySAGEFileContent(filePath, "G", decimalSeparatorValue);

                Assert.AreEqual(contenuFichier[0].ToString(), contenuFichier[1].ToString(), "Le contenu du fichier SAGE généré est incorrect.");
                Assert.AreEqual((2 * montantInventory).ToString(), contenuFichier[0].ToString(), "Le montant issu du fichier SAGE n'est pas égal au montant de la RN défini dans l'application.");
                Assert.AreEqual(montantGlobal.ToString(), montantDetaille.ToString(), "Les montants AmountDebit et AmountCredit de la RN envoyée vers SAGE ne sont pas les mêmes dans l'onglet Accounting.");
                Assert.AreEqual((2 * montantInventory).ToString(), montantGlobal.ToString(), "Le montant issu du fichier SAGE n'est pas égal au montant de la RN défini dans l'application.");

                WebDriver.Navigate().Refresh();
                inventoriesPage.ResetFilters();
                inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, ID);
                bool isAccounted = inventoriesPage.isAccounted();
                Assert.IsTrue(isAccounted, string.Format(MessageErreur.FILTRE_ERRONE, "'Exported for sage manually'"));
            }
            finally
            {
                homePage.SetSageAutoEnabled(site, false);
            }
        }

        // aucun pays n'utilise SAGE AUTO
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_SageAuto_ExportSageItemOK()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();
            string itemName = TestContext.Properties["Item_Inventory"].ToString();
            var currency = TestContext.Properties["Currency"].ToString();

            bool newVersionPrint = true;

            // Log in
            var homePage = LogInAsAdmin();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            // Config pour export SAGE auto mais pas pour le site
            homePage.SetSageAutoEnabled(site, true, "Inventory");

            try
            {
                var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();

                if (newVersionPrint)
                {
                    inventoriesPage.ClearDownloads();
                }

                // Create
                var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place);
                var inventoryItem = inventoryCreateModalpage.Submit();

                inventoryItem.Filter(FilterItemType.SearchByName, itemName);
                inventoryItem.SelectFirstItem();
                inventoryItem.AddPhysicalQuantity(itemName, "2");

                // On valide l'inventory créé
                var validateInventory = inventoryItem.Validate();
                validateInventory.ValidatePartialInventory();

                var inventoryFooter = inventoryItem.ClickOnFooterTab();
                double montantInventory = inventoryFooter.GetInventoryTotalHT(currency, decimalSeparatorValue);

                // Calcul du montant de la facture transmise à TL
                var inventoryAccounting = inventoryFooter.ClickOnAccounting();
                double montantGlobal = inventoryAccounting.GetInventoryGrossAmount("G", decimalSeparatorValue);
                double montantDetaille = inventoryAccounting.GetInventoryDetailAmount("G", decimalSeparatorValue);

                Assert.AreEqual(montantGlobal.ToString(), montantDetaille.ToString(), "Les montants AmountDebit et AmountCredit de l'inventory envoyée vers SAGE ne sont pas les mêmes dans l'onglet Accounting.");
                Assert.AreEqual((2 * montantInventory).ToString(), montantGlobal.ToString(), "Le montant issu du fichier SAGE n'est pas égal au montant de l'inventory défini dans l'application.");

            }
            finally
            {
                homePage.SetSageAutoEnabled(site, false);
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_SageManuel_ExportSAGEItemKO_CodeJournal()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();
            string itemName = TestContext.Properties["Item_Inventory"].ToString();
            string journalInventory = TestContext.Properties["Journal_Inventory"].ToString();

            bool newVersionPrint = true;

            // Log in
            var homePage = LogInAsAdmin();

            // Config pour export SAGE manuel
            homePage.SetSageAutoEnabled(site, false);

            try
            {

                // Parameter - Accounting --> Journal KO pour test
                VerifyAccountingJournal(homePage, site, "");

                var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();

                if (newVersionPrint)
                {
                    inventoriesPage.ClearDownloads();
                }

                // Create
                var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place);
                var inventoryItem = inventoryCreateModalpage.Submit();

                inventoryItem.Filter(FilterItemType.SearchByName, itemName);
                inventoryItem.SelectFirstItem();
                inventoryItem.AddPhysicalQuantity(itemName, "2");

                // On valide l'inventory créé
                var validateInventory = inventoryItem.Validate();
                validateInventory.ValidatePartialInventory();

                // Vérifie qu'un message d'erreur est présent dans l'onglet Accounting
                var inventoryAccounting = inventoryItem.ClickOnAccountingTab();
                string erreur = inventoryAccounting.GetErrorMessage();

                Assert.AreNotEqual("", erreur, "Aucun message d'erreur n'apparaît alors que le paramétrage relatif au code journal est KO.");
                Assert.IsTrue(erreur.Contains("no Inventory journal value set"), "Le message d'erreur ne concerne pas le paramétrage Code journal.");
            }
            finally
            {
                // Parameter - Accounting --> Journal KO pour test
                VerifyAccountingJournal(homePage, site, journalInventory);
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_SageManuel_ExportSAGEItemKO_SiteAnalyticalPlan()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();
            string itemName = TestContext.Properties["Item_Inventory"].ToString();

            bool newVersionPrint = true;

            // Log in
            var homePage = LogInAsAdmin();

            // Config pour export SAGE manuel mais pas pour le site
            homePage.SetSageAutoEnabled(site, false);

            try
            {

                // Sites -- > Analytical plan et section KO pour test
                VerifySiteAnalyticalPlanSection(homePage, site, false);

                var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();

                if (newVersionPrint)
                {
                    inventoriesPage.ClearDownloads();
                }

                // Create
                var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place);
                var inventoryItem = inventoryCreateModalpage.Submit();

                inventoryItem.Filter(FilterItemType.SearchByName, itemName);
                inventoryItem.SelectFirstItem();
                inventoryItem.AddPhysicalQuantity(itemName, "2");

                // On valide l'inventory créé
                var validateInventory = inventoryItem.Validate();
                validateInventory.ValidatePartialInventory();

                // Vérifie qu'un message d'erreur est présent dans l'onglet Accounting
                var inventoryAccounting = inventoryItem.ClickOnAccountingTab();
                string erreur = inventoryAccounting.GetErrorMessage();

                Assert.AreNotEqual("", erreur, "Aucun message d'erreur n'apparaît alors que le paramétrage relatif au 'Analytic plan' du site est KO.");
                Assert.IsTrue(erreur.Contains($"Accounting analytic plan of the site {site} cannot be empty"), "Le message d'erreur ne concerne pas le paramétrage relatif au 'Analytic plan' du site.");
            }
            finally
            {
                VerifySiteAnalyticalPlanSection(homePage, site);
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_SageManuel_ExportSAGEItemKO_NoGroupVat()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();
            string itemName = TestContext.Properties["Item_Inventory"].ToString();
            string taxName = TestContext.Properties["TaxTypeInvoicesExportSage"].ToString();
            bool newVersionPrint = true;

            // Log in
            var homePage = LogInAsAdmin();

            // Config pour export SAGE Manuel mais pas pour le site
            homePage.SetSageAutoEnabled(site, false);

            // Récupération du groupe de l'item
            string itemGroup = GetItemGroup(homePage, itemName);

            try
            {
                // Parameter - Accounting --> Group & VAT KO pour test
                VerifyGroupAndVAT(homePage, itemGroup, taxName, false);

                var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
                inventoriesPage.ResetFilters();

                if (newVersionPrint)
                {
                    inventoriesPage.ClearDownloads();
                }

                // Create
                var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place);
                var inventoryItem = inventoryCreateModalpage.Submit();

                inventoryItem.Filter(FilterItemType.SearchByName, itemName);
                inventoryItem.SelectFirstItem();
                inventoryItem.AddPhysicalQuantity(itemName, "2");

                // On valide l'inventory créé
                var validateInventory = inventoryItem.Validate();
                validateInventory.ValidatePartialInventory();

                // Vérifie qu'un message d'erreur est présent dans l'onglet Accounting
                var inventoryAccounting = inventoryItem.ClickOnAccountingTab();
                string erreur = inventoryAccounting.GetErrorMessage();

                Assert.AreNotEqual("", erreur, "Aucun message d'erreur n'apparaît alors que le paramétrage relatif au 'Group & VAT' de l'item est KO.");
                Assert.IsTrue(erreur.Contains($"Inventory variation account for group {itemGroup} is not set"), "Le message d'erreur ne concerne pas le paramétrage relatif au 'Group & VAT' de l'item.");
            }
            finally
            {
                VerifyGroupAndVAT(homePage, itemGroup, taxName);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_Items_ItemExpiryDate()
        {
            // Log in
            HomePage homePage= LogInAsAdmin();

            //Act         
            var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            inventoriesPage.ResetFilters();
            inventoriesPage.Filter(InventoriesPage.FilterType.ShowNotValidated, true);
            InventoryItem inventoryItems = inventoriesPage.SelectFirstInventory();
            string itemName = inventoryItems.GetFirstItemName();
            Assert.IsNotNull(itemName, "echec itemName");
            inventoryItems.SelectFirstItem();
            inventoryItems.AddPhysicalPackagingQuantity(itemName, "0");
            inventoryItems.AddPhysicalQuantity(itemName, "2");
            inventoryItems.SetPrice(itemName, "10");
            Thread.Sleep(2000);

            // unselect
            inventoryItems.ResetFilter();
            string totalQty = inventoryItems.GetTotalPhysQty();
            string physQty = inventoryItems.GetPhysicalQuantity();
            // Attendu : <2 KG(1000 G)>, Réel : <2>
            // - vérifier que le total quantity est égal à la physical quantity de l'item
            Assert.IsTrue(totalQty.StartsWith(physQty), "total quantity pas égal à phys quantity");

            inventoryItems.SelectFirstItem();

            // vide expiry
            InventoryExpiry expiryDate = inventoryItems.ShowFirstExpiryDate();
            expiryDate.DeleteAllExpiryDate();
            expiryDate.Save();

            //1.Cliquer sur le expiry date du premier item
            expiryDate = inventoryItems.ShowFirstExpiryDate();
            // vérifier le nom de l'item
            string expiryItemName = expiryDate.WaitForElementIsVisible(By.XPath("//*[@id=\"formSaveDates\"]/div[2]/div[1]/div[2]/p/b")).Text;
            Assert.AreEqual(itemName, expiryItemName);
            //vérifier que le total quantity est égal à la physical quantity de l'item
            Assert.AreEqual(physQty, inventoryItems.GetExpiryDateQuantity(), "Le total quantity dans l'expiry date n'est pas égal à la physical quantity.");

            //2. ajouter une nouvelle expiry date
            expiryDate.DeleteAllExpiryDate();
            // - quantity supérieur au total quatity
            // --> message d'erreur 'The sum of individual quantities is greater than the total output quantity.'
            expiryDate.AddExpiryDate("7", DateUtils.Now.AddDays(2));
            // validator
            expiryDate.Save();
            string expiryMessageErreur = expiryDate.WaitForElementIsVisible(By.XPath("//*[@id=\"divWarningQuantityOverflow\"]/label")).Text;
            Assert.AreEqual("The sum of individual quantities is greater than the total inventory quantity.", expiryMessageErreur, "mauvais message d'erreur");
            //-quantity inférieur puis save
            //-- > logo devenu vert et données sauvées bien dans la popup
            expiryDate.DeleteAllExpiryDate();
            expiryDate.Save();
            expiryDate = inventoryItems.ShowFirstExpiryDate();

            expiryDate.AddExpiryDate("2", DateUtils.Now.AddDays(2));
            expiryDate.Save();
            Assert.IsTrue(expiryDate.CheckGreenIcon(), "icone non verte cas 1");
            InventoryGeneralInformation generalInfo = inventoryItems.ClickOnGeneralInformationTab();
            inventoryItems = generalInfo.ClickOnItems();
            inventoryItems.SelectFirstItem();
            Assert.IsTrue(expiryDate.CheckGreenIcon(), "icone non verte cas 2");

            //3. recliquer sur expiry date et supprimer
            //--> logo redevenu comme avant et plus de date dans la popup
            expiryDate = inventoryItems.ShowFirstExpiryDate();
            expiryDate.DeleteAllExpiryDate();
            expiryDate.Save();
            Assert.IsFalse(expiryDate.CheckGreenIcon(), "icone verte cas 1");
            generalInfo = inventoryItems.ClickOnGeneralInformationTab();
            inventoryItems = generalInfo.ClickOnItems();
            inventoryItems.SelectFirstItem();
            Assert.IsFalse(expiryDate.CheckGreenIcon(), "icone verte cas 2");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_Items_EditItem()
        {
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            inventoriesPage.ResetFilters();

            // Create
            var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
            inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place);
            var inventoryItem = inventoryCreateModalpage.Submit();
            string itemName = inventoryItem.GetFirstItemName();
            inventoryItem.SelectFirstItem();
            ItemGeneralInformationPage itemGeneralInfo = inventoryItem.EditItem();
            Assert.AreEqual(itemName, itemGeneralInfo.GetItemName(), "mauvais item sélectionné");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_Items_ResetQuantities()
        {
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();
            var currency = TestContext.Properties["Currency"].ToString();

            // Log in
            var homePage = LogInAsAdmin();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            //Act
            var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            inventoriesPage.ResetFilters();

            //Etre sur un inventaire non validé

            //1) Créer un inventaire
            var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
            inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place);
            var inventoryItem = inventoryCreateModalpage.Submit();
            string itemName = inventoryItem.GetFirstItemName();
            double priceOriginal = inventoryItem.GetPrice(currency, decimalSeparatorValue);
            inventoryItem.SelectFirstItem();

            //2) Ajouter une physical packaging quantity, physical quantity, price, expiry date, comment sur un item
            inventoryItem.AddPhysicalPackagingQuantity(itemName, "5");
            inventoryItem.AddPhysicalQuantity(itemName, "4");
            inventoryItem.SetPrice(itemName, "10");
            InventoryExpiry expiry = inventoryItem.ShowFirstExpiryDate();
            expiry.AddExpiryDate("4", DateUtils.Now.AddDays(7));
            expiry.Save();

            inventoryItem.Refresh();
            //Vérifier que les informations sur l'item modifié (calculs)
            string th = inventoryItem.GetTheoricalQuantity();
            //- Physical packaging à 5
            Assert.AreEqual("5", inventoryItem.GetPhysicalPackagingQuantity(), "Phys pack qty non saisie");
            //- phys.qty à 4
            Assert.AreEqual("4", inventoryItem.GetPhysicalQuantity(), "Phys qty non saisie");
            //-price à 10
            Assert.AreEqual(10.0, inventoryItem.GetPrice(currency, decimalSeparatorValue), 0.00001, "Prix non remis à la valeur initiale");
            //-total phys qty, qty dif calculé
            string storageUnit = inventoryItem.GetStorageUnit();
            Assert.AreEqual("9 " + storageUnit, inventoryItem.GetTotalPhysQty(), "total qty");
            //influencé par Theo. qty, Phys. pckging qty, et Phys qty, mais calcul incompréhensible.
            double qtyDiff = double.Parse(inventoryItem.GetPhysicalPackagingQuantity()) + double.Parse(inventoryItem.GetPhysicalQuantity()) - double.Parse(th);
            Assert.AreEqual(qtyDiff, double.Parse(inventoryItem.GetQtyDiff()), 0.001, "diff value non calculé");
            //-comment(de l'icone et du comment)
            inventoryItem.SelectFirstItem();

            //-la date d'expiration et l'icone en vert
            expiry = inventoryItem.ShowFirstExpiryDate();
            Assert.AreEqual(1, expiry.CountExpiryDate(), "ExpiryDate non saisie");
            expiry.Save();
            Assert.IsTrue(expiry.CheckGreenIcon(), "icone non verte");

            //3) ... puis Reset quantities
            inventoryItem.ShowExtendedMenu();
            //allonger temps d'attente après reset quantitiés
            //ticket 12573
            inventoryItem.ResetQty();

            inventoryItem.Refresh();

            th = inventoryItem.GetTheoricalQuantity();
            //Vérifier que les informations sur l'item modifier sont revenues à 0
            //- Physical packaging qty vide
            Assert.AreEqual("0", inventoryItem.GetPhysicalPackagingQuantity(), "Phys pack qty non remis à zéro");
            //- phys.qty à 0
            Assert.AreEqual("0", inventoryItem.GetPhysicalQuantity(), "Phys qty non remis à zéro");
            //-price revenu à son prix
            Assert.AreEqual(priceOriginal, inventoryItem.GetPrice(currency, decimalSeparatorValue), 0.00001, "Prix non remis à la valeur initiale");
            //-total phys qty, qty dif, value à 0
            Assert.AreEqual("0 " + storageUnit, inventoryItem.GetTotalPhysQty(), "total qty non remis à zéro");
            Assert.AreEqual(-double.Parse(th), double.Parse(inventoryItem.GetQtyDiff()), 0.001, "diff value non remis à zéro");
            //-suppression du comment(de l'icone et du comment)
            inventoryItem.SelectFirstItem();

            //-suppression de la date d'expiration et retour de l'icone en noir
            expiry = inventoryItem.ShowFirstExpiryDate();
            Assert.AreEqual(0, expiry.CountExpiryDate(), "ExpiryDate non remis à zéro");
            expiry.Cancel();
            Assert.IsFalse(expiry.CheckGreenIcon(), "icone non noir");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_ExportCompareOrderInventoryItems()
        {
            //prepare
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            // arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            DeleteAllFileDownload();
            homePage.ClearDownloads();

            //Act
            var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            inventoriesPage.ResetFilters();

            // Evite message d'erreur "que sur un mois" (plage ?)
            inventoriesPage.Filter(InventoriesPage.FilterType.ShowItemsNotValidated, false);
            var firstDateValue = inventoriesPage.GetFirstDate();
            var firstDate = !String.IsNullOrEmpty(firstDateValue) ? DateTime.ParseExact(firstDateValue, "dd/MM/yyyy", null) : DateUtils.Now;
            var dateTo = firstDate.AddDays(2);
            inventoriesPage.Filter(InventoriesPage.FilterType.DateFrom, firstDate);
            inventoriesPage.Filter(InventoriesPage.FilterType.DateTo, dateTo);
            //Export Proccess
            inventoriesPage.ExportExcelFile(true);
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            FileInfo correctDownloadedFile = inventoriesPage.GetExportExcelFile(taskFiles);

            //test si le fichier est exporté
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

            //Assert
            var itemsisSame = inventoriesPage.InventoryItemsiSame(correctDownloadedFile);
            Assert.IsTrue(itemsisSame, "Les items ne sont pas dans le même ordre");

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_Display_Theorical_Value()
        {
            // prepare
            string site = TestContext.Properties["Site"].ToString();
            string delivery = TestContext.Properties["PlaceFrom"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();
            string itemName1 = "belzebuth1f";
            string itemName2 = "belzebuth2f";
            string itemName3 = "belzebuth3f";
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            var currency = TestContext.Properties["Currency"].ToString();

            // arrange

            var homePage = LogInAsAdmin();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            //-Avoir des items avec un stock positif sur un storage place
            //Je prend les 3 premiers items
            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.ShowAll, true);
            ItemGeneralInformationPage itemGeneralInfo;
            List<string> itemNames = new List<string>() { itemName1, itemName2, itemName3 };
            itemPage.Filter(ItemPage.FilterType.Search, itemName1);
            if (itemPage.CheckTotalNumber() == 0)
            {
                ItemCreateModalPage newItem = itemPage.ItemCreatePage();
                itemGeneralInfo = newItem.FillField_CreateNewItem(itemName1, group, workshop, taxType, prodUnit);
                itemPage = itemGeneralInfo.BackToList();
            }
            itemPage.Filter(ItemPage.FilterType.Search, itemName2);
            if (itemPage.CheckTotalNumber() == 0)
            {
                ItemCreateModalPage newItem = itemPage.ItemCreatePage();
                itemGeneralInfo = newItem.FillField_CreateNewItem(itemName2, group, workshop, taxType, prodUnit);
                itemPage = itemGeneralInfo.BackToList();
            }
            itemPage.Filter(ItemPage.FilterType.Search, itemName3);
            if (itemPage.CheckTotalNumber() == 0)
            {
                ItemCreateModalPage newItem = itemPage.ItemCreatePage();
                itemGeneralInfo = newItem.FillField_CreateNewItem(itemName3, group, workshop, taxType, prodUnit);
                itemPage = itemGeneralInfo.BackToList();
            }
            itemPage.Filter(ItemPage.FilterType.Search, itemName1);

            itemGeneralInfo = itemPage.ClickOnItem(0);
            if (!itemGeneralInfo.ScanPackaging(site, supplier))
            {
                ItemCreateNewPackagingModalPage pack0 = itemGeneralInfo.NewPackaging();
                itemGeneralInfo = pack0.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
            }
            itemPage = itemGeneralInfo.BackToList();

            itemPage.Filter(ItemPage.FilterType.Search, itemName2);
            itemGeneralInfo = itemPage.ClickOnItem(0);
            if (!itemGeneralInfo.ScanPackaging(site, supplier))
            {
                ItemCreateNewPackagingModalPage pack1 = itemGeneralInfo.NewPackaging();
                itemGeneralInfo = pack1.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
            }
            itemPage = itemGeneralInfo.BackToList();

            itemPage.Filter(ItemPage.FilterType.Search, itemName3);
            itemGeneralInfo = itemPage.ClickOnItem(0);
            if (!itemGeneralInfo.ScanPackaging(site, supplier))
            {
                ItemCreateNewPackagingModalPage pack2 = itemGeneralInfo.NewPackaging();
                itemGeneralInfo = pack2.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
            }
            //1 inventory avec les 3 items en question, assert storage pas vide
            InventoriesPage inventoryPage = homePage.GoToWarehouse_InventoriesPage();
            InventoryCreateModalPage createInventory = inventoryPage.InventoryCreatePage();
            createInventory.FillField_CreateNewInventory(DateUtils.Now, site, place);
            InventoryItem inventoryItem = createInventory.Submit();

            inventoryItem.ResetFilter();

            inventoryItem.Filter(FilterItemType.SearchByName, itemNames[0]);
            inventoryItem.SelectFirstItem();
            inventoryItem.AddPhysicalQuantity(itemNames[0], "10");
            //Thread.Sleep(4000);

            inventoryItem.Filter(FilterItemType.SearchByName, itemNames[1]);
            inventoryItem.SelectFirstItem();
            inventoryItem.AddPhysicalQuantity(itemNames[1], "10");
            //Thread.Sleep(4000);

            inventoryItem.Filter(FilterItemType.SearchByName, itemNames[2]);
            inventoryItem.SelectFirstItem();
            inventoryItem.AddPhysicalQuantity(itemNames[2], "10");
            //Thread.Sleep(4000);

            inventoryItem.Filter(FilterItemType.SearchByName, "");
            inventoryItem.Filter(FilterItemType.ShowItemsWithQtyOnly, true);

            InventoryValidationModalPage validation = inventoryItem.Validate();
            inventoryItem = validation.ValidateTotalInventory();

            //-Les items doivent avoir un prix différent de 0 sur le site qui contient le storage place
            itemPage = inventoryItem.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName1);
            ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
            ItemStoragePage storage = generalInfo.ClickOnStorage();
            Assert.IsTrue(storage.CheckTotalNumber() >= 1, "no storage " + itemName1);
            var firstPrice = storage.WaitForElementIsVisible(By.XPath("//*[@id='table-itemDetailsStorage']/tbody/tr[2]/td[6]"));
            var firstPriceValue = firstPrice.Text.Trim();
            Assert.AreEqual("€ 10", firstPriceValue, "il faut un prix différent de 0");

            //-Avoir toute les RN, les OF et les inventaires validés sur le site
            ReceiptNotesPage rn = storage.GoToWarehouse_ReceiptNotesPage();
            ReceiptNotesCreateModalPage rnCreate = rn.ReceiptNotesCreatePage();
            rnCreate.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, delivery));
            ReceiptNotesItem rnItems = rnCreate.Submit();
            rnItems.ResetFilters();
            rnItems.Filter(ReceiptNotesItem.FilterItemType.SearchByName, itemName1);
            rnItems.SetFirstReceivedQuantity("2");
            ReceiptNotesQualityChecks rnChecks = rnItems.ClickOnChecksTab();
            rnChecks.DeliveryAccepted();
            if (rnChecks.CanClickOnSecurityChecks())
            {
                rnChecks.CanClickOnSecurityChecks();
                rnChecks.SetSecurityChecks("No");
                rnChecks.SetQualityChecks();
                rnItems = rnChecks.ClickOnReceiptNoteItemTab();
            }
            else
            {
                rnItems = rnChecks.ClickOnReceiptNoteItemTab();
            }
            rnItems.Validate();
            OutputFormPage of = rnItems.GoToWarehouse_OutputFormPage();
            OutputFormCreateModalPage ofCreate = of.OutputFormCreatePage();
            ofCreate.FillField_CreatNewOutputForm(DateUtils.Now, site, "MAD4", "Economato");
            OutputFormItem ofItems = ofCreate.Submit();
            ofItems.Filter(OutputFormItem.FilterItemType.SearchByName, itemName2);
            ofItems.AddPhyQuantity(itemName2, "3");
            OutputFormQualityChecks ofChecks = ofItems.ClickOnChecksTab();
            ofChecks.DeliveryAccepted();
            ofItems = ofChecks.ClickOnItemsTab();
            ofItems.Validate();

            //- Ouvrir un nouveau inventaire
            inventoryPage = ofItems.GoToWarehouse_InventoriesPage();
            createInventory = inventoryPage.InventoryCreatePage();
            createInventory.FillField_CreateNewInventory(DateUtils.Now, site, place);
            inventoryItem = createInventory.Submit();
            inventoryItem.ResetFilter();
            inventoryItem.Filter(FilterItemType.SearchByName, itemName1);
            inventoryItem.SelectFirstItem();
            inventoryItem.AddPhysicalQuantity(itemName1, "5");
            Thread.Sleep(4000);
            var theoricalQuantity = inventoryItem.GetTheoricalQuantity();
            Assert.AreEqual("10", theoricalQuantity, "RN theory 10 items raté");
            inventoryItem.Filter(FilterItemType.SearchByName, itemName2);
            inventoryItem.SelectFirstItem();
            inventoryItem.AddPhysicalQuantity(itemName2, "5");
            Thread.Sleep(4000);
            theoricalQuantity = inventoryItem.GetTheoricalQuantity();
            Assert.AreEqual("7", theoricalQuantity, "OF theory 10-3 items raté");
            inventoryItem.Filter(FilterItemType.SearchByName, itemName3);
            inventoryItem.SelectFirstItem();
            inventoryItem.AddPhysicalQuantity(itemName3, "5");
            Thread.Sleep(4000);
            theoricalQuantity = inventoryItem.GetTheoricalQuantity();
            Thread.Sleep(4000);
            Assert.AreEqual("10", theoricalQuantity, "Inventory 10 items raté");
            validation = inventoryItem.Validate();
            inventoryItem = validation.ValidateTotalInventory();
            // Vérifier si lors de l'ouverture de l'inventaire si on a un stock positif qui a un prix alors il y a l'affichage en haut à gauche d'une valeur pour le theorical value.            

            inventoryItem.ClearDownloads();
            string inventoryNumber = inventoryItem.GetInventoryNumber();
            var equationToEqual = 10.0 * 20.0 + 10.0 * 7.0;
            var theoricalValue = inventoryItem.GetTheoricalValue(currency, decimalSeparatorValue);
            Assert.AreEqual(equationToEqual, theoricalValue, 0.001);
            inventoryItem.ExportExcelFile(true, downloadsPath);
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            FileInfo correctDownloadedFile = inventoryItem.GetExportExcelFile(taskFiles);
            Assert.IsTrue(correctDownloadedFile.Exists);
            inventoryItem.CheckExport(correctDownloadedFile, site, inventoryNumber, decimalSeparatorValue);
            //De même dans l'export de l'inventaire au format excel doit apparaître des valeurs dans la colonne "theorical value"
            int resultNumber = OpenXmlExcel.GetExportResultNumber("Inventory", correctDownloadedFile.FullName);
            Assert.AreNotEqual(0, resultNumber, "Fichier Excel vide");
            List<string> colThQty = OpenXmlExcel.GetValuesInList("Theo. qty", "Inventory", correctDownloadedFile.FullName);
            Assert.AreEqual(3, colThQty.Count, "Mauvais nombre de lignes Excel");
            Thread.Sleep(4000);
            int[] sumThQty = new int[3];
            int i = 0;
            foreach (string col in colThQty)
            {
                if (i == 1)
                {
                    Assert.AreEqual("7", col);
                    sumThQty[i] = 7;
                }
                else
                {
                    Assert.AreEqual("10", col);
                    sumThQty[i] = 10;
                }
                i++;
            }
            //La theorical value doit être égale à la somme pour chaque items de sa theorical qty multiplié par son prix.
            List<string> colPrice = OpenXmlExcel.GetValuesInList("Average price", "Inventory", correctDownloadedFile.FullName);
            Assert.AreEqual(3, colPrice.Count, "Mauvais nombre de lignes Excel");
            int[] sumPrice = new int[3];
            i = 0;
            foreach (string col in colPrice)
            {
                Assert.AreEqual("10", col);
                sumPrice[i] = 10;
                i++;
            }
            var sumPricee = (int)(10.0 * 20.0 + 10.0 * 7.0);
            var sumThQtyPrices = sumThQty[0] * sumPrice[0] + sumThQty[1] * sumPrice[1] + sumThQty[2] * sumPrice[2];
            Assert.AreEqual(sumPricee, sumThQtyPrices, "les valeurs ne sont pas egaux");
            //La theorical value doit être égale entre la valeur en haut à gauche et la somme de la colonne "theorical value" de l'export excel
            List<string> colThValue = OpenXmlExcel.GetValuesInList("Theo. value", "Inventory", correctDownloadedFile.FullName);
            Assert.AreEqual(3, colThValue.Count, "Mauvais nombre de lignes Excel");
            int[] sumThValue = new int[3];
            i = 0;
            foreach (string col in colThValue)
            {
                if (i == 1)
                {
                    Assert.AreEqual("70", col);
                    sumThValue[i] = 70;
                }
                else
                {
                    Assert.AreEqual("100", col);
                    sumThValue[i] = 100;
                }
                i++;
            }
            var somme = (int)(10.0 * 20.0 + 10.0 * 7.0);
            var sommeThValues = sumThValue[0] + sumThValue[1] + sumThValue[2];
            Assert.AreEqual(somme, sommeThValues, "La somme des variables et la somme des Theo. value ne sont pas égaux  ");

            // on épuise l'inventary
            of = rnItems.GoToWarehouse_OutputFormPage();
            ofCreate = of.OutputFormCreatePage();
            ofCreate.FillField_CreatNewOutputForm(DateUtils.Now, site, "MAD4", "Economato");
            ofItems = ofCreate.Submit();
            ofItems.Filter(OutputFormItem.FilterItemType.SearchByName, itemName1);
            ofItems.AddPhyQuantity(itemName1, "5");
            ofItems.Filter(OutputFormItem.FilterItemType.SearchByName, itemName2);
            ofItems.AddPhyQuantity(itemName2, "5");
            ofItems.Filter(OutputFormItem.FilterItemType.SearchByName, itemName3);
            ofItems.AddPhyQuantity(itemName3, "5");
            ofChecks = ofItems.ClickOnChecksTab();
            ofChecks.DeliveryAccepted();
            ofItems = ofChecks.ClickOnItemsTab();
            ofItems.Validate();

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_Inventory_OF()
        {
            //Site actif
            string site = TestContext.Properties["Site"].ToString();
            //Storage place actif
            //Avoir des items avec des quantités et des prix supérieur à 0
            string itemOutputForm = TestContext.Properties["Item_OutputForm"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();

            var currency = TestContext.Properties["Currency"].ToString();

            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();

            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxName = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();

            // Log in
            HomePage homePage= LogInAsAdmin();
           
            homePage.ClearDownloads();
            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            //Act

            // avoir un item itemOutputForm
            // Création items
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemOutputForm);
            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemOutputForm, group, workshop, taxName, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);

                itemPage = itemGeneralInformationPage.BackToList();
                itemPage.ResetFilter();
                itemPage.Filter(ItemPage.FilterType.Search, itemOutputForm);
            }
            Assert.AreEqual(itemOutputForm, itemPage.GetFirstItemName(), $"L'item {itemOutputForm} n'est pas présent dans la liste des items disponibles.");

            // création d'une qty th : créer un inventaire avec phys qty à 10, le valider,
            var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            inventoriesPage.ResetFilters();
            //Ouvrir un inventaire sans le valider
            InventoryCreateModalPage invCreate = inventoriesPage.InventoryCreatePage();
            invCreate.FillField_CreateNewInventory(DateUtils.Now, site, place);
            string invNumber = invCreate.GetInventoryNumber();
            InventoryItem invItem = invCreate.Submit();
            invItem.Filter(FilterItemType.SearchByName, itemOutputForm);
            invItem.SelectFirstItem();
            invItem.AddPhysicalQuantity(itemOutputForm, "200");
            InventoryValidationModalPage validateModal = invItem.Validate();
            validateModal.ValidateTotalInventory();

            // créer ensuite un inventaire, là th qty a une valeur
            inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            inventoriesPage.ResetFilters();
            //Ouvrir un inventaire sans le valider
            invCreate = inventoriesPage.InventoryCreatePage();
            invCreate.FillField_CreateNewInventory(DateUtils.Now, site, place);
            invNumber = invCreate.GetInventoryNumber();
            invItem = invCreate.Submit();
            //Après avoir ouvert l'inventaire sans être validé, sauvegarder les valeurs Theorical value (A)
            invItem.Filter(FilterItemType.ShowItemsWithQtyOnly, true);
            invItem.PageSize("100");
            Dictionary<string, string> allThA = invItem.GetAllTheoricalValues();
            Assert.AreNotEqual(0, allThA.Count, "Pas de valeurs th");
            //Filtrer l'inventaire pour voir un item (i) avec un stock théorique et un prix positif, sauvegarder la valeur du stock théorique (B)
            invItem.Filter(FilterItemType.SearchByName, itemOutputForm);
            invItem.SelectFirstItem();
            double saveThB = invItem.GetTheoricalValue(currency, decimalSeparatorValue);
            Assert.AreNotEqual(0, saveThB, 0.01, "ThValue vide");

            //Réaliser un OF sur le site actif depuis le storage place actif vers un autre place du site
            OutputFormPage of = invItem.GoToWarehouse_OutputFormPage();
            OutputFormCreateModalPage ofModal = of.OutputFormCreatePage();
            ofModal.FillField_CreatNewOutputForm(DateUtils.Now, site, "MAD4", "Economato");
            OutputFormItem ofItem = ofModal.Submit();
            // contenant l'item (i) avec une quantité,
            ofItem.Filter(OutputFormItem.FilterItemType.SearchByName, itemOutputForm);
            ofItem.SelectFirstItem();
            ofItem.AddPhysicalQuantity(itemOutputForm, "10");
            // valider l'OF.
            ofItem.Validate();

            //Ouvrir l'inventaire et faire refresh, voir l'item (i) et son stock théorique (C), voir la theorical value (D)
            inventoriesPage = ofItem.GoToWarehouse_InventoriesPage();
            inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, invNumber);
            invItem = inventoriesPage.SelectFirstInventory();
            invItem.Refresh();
            invItem.Filter(FilterItemType.SearchByName, itemOutputForm);
            string thQtyApresC = invItem.GetTheoricalQuantity();
            double thValApresD = invItem.GetTheoricalValue(currency, decimalSeparatorValue);

            // faut valider l'inventaire sinon le th qty s'efface.
            validateModal = invItem.Validate();
            validateModal.ValidateTotalInventory();

            //Vérifier que(A) est différent de(D)
            Assert.AreNotEqual(allThA[itemOutputForm], thValApresD, "A pas différent de D");
            //Vérifier que(B) est plus grand que(C)
            Assert.AreNotEqual(saveThB > Double.Parse(thQtyApresC), 0.01, "B pas plus grand que C");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_Top10_Difference()
        {
            string site = TestContext.Properties["Site"].ToString();
            //Storage place actif
            //Avoir des items avec des quantités et des prix supérieur à 0
            string place = TestContext.Properties["InventoryPlace"].ToString();

            // Log in
            var homePage = LogInAsAdmin();
            homePage.ClearDownloads();
            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();
            var currency = TestContext.Properties["Currency"].ToString();

            //Act
            CreateOtherInventory(homePage, site, place);

            var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            inventoriesPage.ResetFilters();
            //Ouvrir un inventaire sans le valider
            InventoryCreateModalPage invCreate = inventoriesPage.InventoryCreatePage();
            invCreate.FillField_CreateNewInventory(DateUtils.Now, site, place);
            InventoryItem invItem = invCreate.Submit();
            //Avoir du stock sur le storage place
            //Avoir des prix pour les items strictement supérieur à 0
            //Remplir un inventaire avec des valeurs physiques aléatoires pour chaque ligne
            invItem.Filter(FilterItemType.ShowItemsWithQtyOnly, true);
            invItem.Refresh();
            if (!invItem.isPageSizeEqualsTo100WidhoutTotalNumber())
            {
                invItem.PageSize("8");
                invItem.PageSize("100");
            }

            Dictionary<string, string> allTh = invItem.GetAllTheoricalValues();
            Dictionary<string, double> allPrice = new Dictionary<string, double>();
            int counter = 11;
            foreach (var key in allTh.Keys)
            {
                invItem.SelectItem(key);
                string thQty = invItem.GetTheoricalQuantityEdit();
                CultureInfo ci = decimalSeparatorValue.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");
                double thQtyDouble = double.Parse(thQty.Replace(currency, ""), ci);
                int thQtyInt = (int)Math.Floor(thQtyDouble);
                int randomQty = new Random().Next(1, thQtyInt);
                invItem.AddPhysicalQuantity(key, randomQty.ToString());
                double price = invItem.GetPriceEdit(currency, decimalSeparatorValue);
                double absDiffPrice = (thQtyDouble * price) - (randomQty * price);
                allPrice.Add(key, absDiffPrice);
                Console.WriteLine(key + " (thQtyDouble * price) - (randomQty * price)");
                Console.WriteLine(key + " (" + thQtyDouble + " * " + price + ") - (" + randomQty + " * " + price + ") = " + absDiffPrice);
                counter--;
                if (counter == 0) break;
            }
            invItem.Filter(FilterItemType.ShowItemsWithQtyOnly, true);
            //faut annuler la validation de ce Inventory sinon ça grignotera nos qty th.
            InventoryValidationModalPage validationModal = invItem.Validate();

            //Au moment de valider l'inventaire, visualiser un top 10 d'items avec les différences d'inventaire
            Dictionary<string, double> itemPrice = validationModal.GetTop10(currency, decimalSeparatorValue);
            foreach (var key in itemPrice.Keys)
            {
                Console.WriteLine("Modal " + key + " : " + itemPrice[key]);
            }
            foreach (var key in itemPrice.Keys)
            {
                Assert.AreEqual(allPrice[key], itemPrice[key], 0.01, key + " avant : " + allPrice[key] + " , apres : " + itemPrice[key]);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_Verification_Theovalue2()
        {
            // Prepare
            //Avoir un site actif
            string site = TestContext.Properties["Site"].ToString();
            //Avoir un storage place actif
            string place = TestContext.Properties["Place"].ToString(); // Economato
            string placeFrom = TestContext.Properties["PlaceFrom"].ToString(); // Economato
            string placeTo = TestContext.Properties["PlaceTo"].ToString(); // Producción
            //Avoir des items avec des qty et des prix
            string itemName = TestContext.Properties["Item_Inventory"].ToString();

            string supplier = TestContext.Properties["SupplierPo"].ToString(); // AIR CPU, S.L.

            string invQtyA = "77";
            string rnQtyC = "10";
            string ofQtyThD = "40";

            // Log in
            var homePage = LogInAsAdmin();
            homePage.ClearDownloads();

            // Act
            //Avoir un inventaire validé ayant une physical value(A) strictement supérieure à 0
            var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            inventoriesPage.ResetFilters();
            InventoryCreateModalPage invCreate = inventoriesPage.InventoryCreatePage();
            invCreate.FillField_CreateNewInventory(DateUtils.Now, site, place);
            string invNumber = invCreate.GetInventoryNumber();
            InventoryItem invItem = invCreate.Submit();
            invItem.Filter(FilterItemType.SearchByName, itemName);
            invItem.SelectFirstItem();
            invItem.AddPhysicalQuantity(itemName, invQtyA);
            InventoryValidationModalPage modalValidate = invItem.Validate();
            invItem = modalValidate.ValidatePartialInventory();
            //Ne pas réaliser d'autre inventaire sur le storage place.

            //Réaliser un RN validé sur le storage place, d'une valeur (C)
            ReceiptNotesPage rnPage = invItem.GoToWarehouse_ReceiptNotesPage();
            ReceiptNotesCreateModalPage rnCreate = rnPage.ReceiptNotesCreatePage();
            rnCreate.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, place));
            ReceiptNotesItem rnItems = rnCreate.Submit();
            ReceiptNotesQualityChecks rnChecks = rnItems.ClickOnChecksTab();
            rnChecks.DeliveryAccepted();
            if (rnChecks.CanClickOnSecurityChecks())
            {
                rnChecks.CanClickOnSecurityChecks();
                rnChecks.SetSecurityChecks("No");
                rnChecks.SetQualityChecks();
                rnItems = rnChecks.ClickOnReceiptNoteItemTab();
            }
            else
            {
                rnItems = rnChecks.ClickOnReceiptNoteItemTab();
            }
            rnItems.Filter(ReceiptNotesItem.FilterItemType.SearchByName, itemName);
            rnItems.SetFirstReceivedQuantity(rnQtyC);
            string rnPackSize = rnItems.GetFormPackaging();
            Assert.AreEqual("bolsa 1Kg 10 KG (10 UD)", rnPackSize);
            rnItems.Validate();
            //Réaliser un OF validé sur le storage place, d'une valeur (D)
            OutputFormPage ofPage = rnItems.GoToWarehouse_OutputFormPage();
            OutputFormCreateModalPage ofCreate = ofPage.OutputFormCreatePage();
            ofCreate.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo);
            OutputFormItem ofItems = ofCreate.Submit();
            ofItems.Filter(OutputFormItem.FilterItemType.SearchByName, itemName);
            ofItems.SelectFirstItem();
            ofItems.AddPhysicalQuantity(itemName, ofQtyThD);
            ofItems.Validate();
            //Ouvrir un inventaire sur le storage place, voir la valeur théorique de l'inventaire (B)
            inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            invCreate = inventoriesPage.InventoryCreatePage();
            invCreate.FillField_CreateNewInventory(DateUtils.Now, site, place);
            invNumber = invCreate.GetInventoryNumber();
            invItem = invCreate.Submit();
            invItem.Filter(FilterItemType.SearchByName, itemName);
            string invQtyB = invItem.GetTheoricalQuantity();

            //B= A+C-D : Si ce n'est pas le cas alors KO	
            Assert.AreEqual(int.Parse(invQtyB), int.Parse(invQtyA) + int.Parse(rnQtyC) * 10 - int.Parse(ofQtyThD), "B = A + C - D : " + invQtyB + " = " + invQtyA + " + " + rnQtyC + "*10 - " + ofQtyThD);
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_Verification_Theo_Value()
        {
            // Prepare
            //Avoir un site actif ex : "BCN"
            string site = TestContext.Properties["SiteBCN"].ToString();
            //Avoir un storage actif ex: "Economat"
            string place = TestContext.Properties["Place"].ToString(); // Economato
            string placeFrom = TestContext.Properties["PlaceFrom"].ToString(); // Economato
            string placeTo = TestContext.Properties["PlaceTo"].ToString(); // Producción
            string itemName = TestContext.Properties["Item_Inventory"].ToString();
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string currency = TestContext.Properties["Currency"].ToString();
            string invNumber = string.Empty;
            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            homePage.ClearDownloads();
            var decimalSeparator = homePage.GetDecimalSeparatorValue();

            // Act
            //Dans le storage actif, avoir des items en stock avec un prix supérieur à 0.
            var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            inventoriesPage.ResetFilters();
            try
            {
                //Ouvrir un inventaire sans le valider et visualiser la theorical value, sauvegarder cette valeur (A)
                var invCreate = inventoriesPage.InventoryCreatePage();
                invCreate.FillField_CreateNewInventory(DateUtils.Now, site, place);
                invNumber = invCreate.GetInventoryNumber();
                var invItem = invCreate.Submit();

                double qtyThValA = invItem.GetValueEdit(currency, decimalSeparator);

                var itemPage = homePage.GoToPurchasing_ItemPage();
                itemPage.ClickOnStorage();

                itemPage.Filter(ItemPage.FilterType.Site, site);
                itemPage.ClickOnStorage();
                itemPage.Export(true);
                DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
                FileInfo[] taskFiles = taskDirectory.GetFiles();
                // StorageBook 2023-03-15 15-20-38.xlsx
                FileInfo trouve = itemPage.GetExportExcelFile(taskFiles, true);
                Assert.IsNotNull(trouve);
                Assert.IsTrue(trouve.Exists);
                //Faire la somme de la colonne Theo Value, sauvegarder cette valeur(B)
                List<string> colonnePlace = OpenXmlExcel.GetValuesInList("Place", "Actual stock", trouve.FullName);
                List<string> colonneThVal = OpenXmlExcel.GetValuesInList("Theo. value", "Actual stock", trouve.FullName);
                double totale_excel = itemPage.GetTotalColFromExcel(colonneThVal, colonnePlace, place, decimalSeparator);
                //Le résultat est que A=B
                //Si il y a une différence alors le test est KO
                //Assert
                Assert.AreEqual(qtyThValA, totale_excel, "les deux valeurs qtyThValA et thValB ne sont pas égaux");

            }
            finally
            {
                homePage.GoToWarehouse_InventoriesPage();
                if (!string.IsNullOrEmpty(invNumber))
                {
                    inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, invNumber);
                    inventoriesPage.deleteInventory();

                }

            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_DeleteInventory()
        {
            //Avoir un site actif ex : "BCN"
            string site = TestContext.Properties["Site"].ToString();
            //Avoir un storage actif ex: "Economat"
            string place = TestContext.Properties["Place"].ToString();

            // Log in
            var homePage = LogInAsAdmin();

            var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            inventoriesPage.ResetFilters();
            InventoryCreateModalPage invCreate = inventoriesPage.InventoryCreatePage();
            invCreate.FillField_CreateNewInventory(DateUtils.Now, site, place);
            string invNumber = invCreate.GetInventoryNumber();
            InventoryItem invItem = invCreate.Submit();
            invItem.BackToList();
            var totalNumberAfterCreate = inventoriesPage.CheckTotalNumber();
            inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, invNumber);
            inventoriesPage.deleteInventory();
            inventoriesPage.ResetFilters();
            var totalNumberAfterDelete = inventoriesPage.CheckTotalNumber();
            Assert.AreEqual(totalNumberAfterDelete, totalNumberAfterCreate - 1, "La suppression ne fonctionne pas.");
        }
        [TestMethod]

        [Timeout(_timeout)]
        public void WA_INV_DetailChangeLine()
        {
            //Avoir un site actif ex : "BCN"
            string site = TestContext.Properties["Site"].ToString();
            //Avoir un storage actif ex: "Economat"
            string place = TestContext.Properties["Place"].ToString();

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            inventoriesPage.ResetFilters();
            InventoryCreateModalPage invCreate = inventoriesPage.InventoryCreatePage();
            invCreate.FillField_CreateNewInventory(DateUtils.Now, site, place);
            string invNumber = invCreate.GetInventoryNumber();
            InventoryItem invItem = invCreate.Submit();
            Assert.IsTrue(invItem.VerifyDetailChangeLine("12"), "Le curseur ne passe pas à la ligne suivante.");
        }
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_CreateMobileInventoryNotValidateAndVerify()
        {
            //Avoir un site actif ex : "BCN"
            string site = TestContext.Properties["Site"].ToString();
            //Avoir un storage actif ex: "Economat"
            string place = TestContext.Properties["Place"].ToString();
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            inventoriesPage.ResetFilters();
            inventoriesPage.ClearDownloads();
            InventoryCreateModalPage invCreate = inventoriesPage.InventoryCreatePage();
            invCreate.FillField_CreateNewInventory(DateUtils.Now, site, place);
            string invNumber = invCreate.GetInventoryNumber();
            InventoryGeneralInformation inventoryGeneralInformation = invCreate.SubmitMobile();
            InventoryItem inventoryItem = inventoryGeneralInformation.ClickOnItems();
            inventoryItem.SelectFirstItem();
            var itemName = inventoryItem.GetFirstItemName();
            inventoryItem.WaitLoading();
            inventoryItem.AddPhysicalQuantity(itemName, "2");
            inventoryItem.WaitLoading();
            inventoryGeneralInformation.BackToList();
            inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, invNumber);
            inventoriesPage.Filter(InventoriesPage.FilterType.DateFrom, DateUtils.Now);
            inventoriesPage.Filter(InventoriesPage.FilterType.DateTo, DateUtils.Now.AddMonths(1));
            inventoriesPage.ExportExcelFile(true);
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            FileInfo trouve = inventoriesPage.GetExportExcelFile(taskFiles, true);
            Assert.IsNotNull(trouve);
            Assert.IsTrue(trouve.Exists);
            var correctDownloadedFile = inventoriesPage.GetExportExcelFile(taskFiles);
            Assert.IsTrue(correctDownloadedFile != null && correctDownloadedFile.Exists, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

            // Récupération du nom du fichier et construction de l'URL du fichier Excel à ouvrir   
            var fileName = correctDownloadedFile.Name;
            var filePath = System.IO.Path.Combine(downloadsPath, fileName);

            int resultNumber = OpenXmlExcel.GetExportResultNumber("Inventory", filePath);
            Assert.IsTrue(inventoriesPage.VerifyNotValidated(), "Le nouvel inventaire est validée ");
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);

        }
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_CreateMobileInventoryValidateAndVerify()
        {
            //Avoir un site actif ex : "BCN"
            string site = TestContext.Properties["Site"].ToString();
            //Avoir un storage actif ex: "Economat"
            string place = TestContext.Properties["Place"].ToString();
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            inventoriesPage.ResetFilters();
            inventoriesPage.ClearDownloads();
            InventoryCreateModalPage invCreate = inventoriesPage.InventoryCreatePage();
            invCreate.FillField_CreateNewInventory(DateUtils.Now, site, place);
            string invNumber = invCreate.GetInventoryNumber();
            InventoryGeneralInformation inventoryGeneralInformation = invCreate.SubmitMobile();
            inventoryGeneralInformation.BackToList();
            inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, invNumber);
            InventoryItem inventoryItem = inventoriesPage.SelectFirstInventory();
            inventoryItem.ResetFilter();
            inventoryItem.Filter(FilterItemType.ShowItemsWithQtyOnly, true);
            var inventoryItems = inventoryItem.GetAllItemsName();
            InventoryValidationModalPage validateInventory = inventoryItem.Validate();
            inventoryItem = validateInventory.ValidatePartialInventory();
            inventoriesPage = inventoryItem.BackToList();
            inventoriesPage.ResetFilters();
            inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, invNumber);
            var isInventoryValidated = inventoriesPage.CheckValidation(true);
            Assert.IsTrue(isInventoryValidated, "L'inventaire' n'est pas validé");
            inventoriesPage.Filter(InventoriesPage.FilterType.DateFrom, DateUtils.Now);
            inventoriesPage.Filter(InventoriesPage.FilterType.DateTo, DateUtils.Now.AddMonths(1));
            inventoriesPage.ExportExcelFile(true);
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();
            FileInfo trouve = inventoriesPage.GetExportExcelFile(taskFiles, true);
            Assert.IsNotNull(trouve);
            Assert.IsTrue(trouve.Exists);
            var correctDownloadedFile = inventoriesPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");
            // Récupération du nom du fichier et construction de l'URL du fichier Excel à ouvrir   
            var fileName = correctDownloadedFile.Name;
            var filePath = System.IO.Path.Combine(downloadsPath, fileName);
            int resultNumber = OpenXmlExcel.GetExportResultNumber("Inventory", filePath);
            var idsExcel = OpenXmlExcel.GetValuesInList("Id", "Inventory", filePath);
            var itemsExcel = OpenXmlExcel.GetValuesInList("Item", "Inventory", filePath);
            Assert.IsTrue(isInventoryValidated, "Le nouvel inventaire est validée ");
            Assert.AreEqual(inventoryItems.Count, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.IsTrue(idsExcel.All(id => int.Parse(id) == int.Parse(invNumber)), "Les IDs d'article exportés ne sont pas conformes aux ID de l’inventaire.");
            Assert.IsTrue(itemsExcel.SequenceEqual(inventoryItems), "Les articles exportés ne sont pas conformes dans l’inventaire.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_PrintPreparationSheetNotValidateInventoryNewVersion()
        {

            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            //Print Preparation Sheet_-_437549_-_20220316103924.pdf
            string DocFileNamePdfBegin = "Print Preparation Sheet_-_";
            //All_files_20220225_102148.zip
            string DocFileNameZipBegin = "All_files_";

            // Log in
            var homePage = LogInAsAdmin();

            //Act
            var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            inventoriesPage.ClearDownloads();
            inventoriesPage.ResetFilters();
            InventoryCreateModalPage invCreate = inventoriesPage.InventoryCreatePage();
            invCreate.FillField_CreateNewInventory(DateUtils.Now, site, place);
            string inventoryNumber = invCreate.GetInventoryNumber();
            InventoryItem inventoryInventoryItem = inventoriesPage.Submit();
            inventoryInventoryItem.BackToList();
            inventoriesPage.Filter(InventoriesPage.FilterType.ShowNotValidated, false);
            inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, inventoryNumber);
            Assert.AreEqual(1, inventoriesPage.CheckTotalNumber(), "Inventory n'est pas crée");
            inventoriesPage.SelectFirstInventory();
            var reportPage = inventoryInventoryItem.PrintPreparationSheet(true);

            var isReportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();

            //Assert  
            Assert.IsTrue(isReportGenerated, "le report n'est pas généré");

            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            // cliquer sur All
            string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            FileInfo fi = new FileInfo(trouve);
            fi.Refresh();
            Assert.IsTrue(fi.Exists, trouve + " non généré");

            PdfDocument document = PdfDocument.Open(fi.FullName);
            Page page1 = document.GetPage(1);
            IEnumerable<Word> words = page1.GetWords();
            Assert.AreEqual(1, words.Count(w => w.Text == site + "(" + site + ")"), site + "(" + site + ") non présent dans le Pdf");
            Assert.AreEqual(1, words.Count(w => w.Text == place), place + " non présent dans le Pdf");
            Assert.AreEqual(1, words.Count(w => w.Text == DateUtils.Now.Date.ToString("dd/MM/yyyy")), DateUtils.Now.Date.ToString("dd/MM/yyyy") + " non présent dans le Pdf");
            Assert.AreEqual(1, words.Count(w => w.Text == inventoryNumber), inventoryNumber + " non présent dans le Pdf");



        }
        
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_DateChanged()
        {
            var homePage = LogInAsAdmin();

            InventoriesPage inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            inventoriesPage.ResetFilters();
            inventoriesPage.Filter(InventoriesPage.FilterType.ShowNotValidated, true);

            InventoryItem inventory = inventoriesPage.SelectFirstInventory();
            inventory.ClickOnGeneralInformationTab();
            string invNumber = inventory.GetInventoryNumber();
            string invDate = inventory.GetInventoryDate();
            DateTime date = DateTime.ParseExact(invDate, "dd/MM/yyyy", null);

            inventory.EditInventoryDate(date.AddDays(1));
            InventoriesPage invPage = inventory.BackToList();

            invPage.Filter(InventoriesPage.FilterType.ByNumber, invNumber);
            InventoryItem inventoryAfterModif = inventoriesPage.SelectFirstInventory();
            inventoryAfterModif.ClickOnGeneralInformationTab();
            string invDateModified = inventoryAfterModif.GetInventoryDate();
            bool isDateModified = !invDate.Equals(invDateModified);
            Assert.IsTrue(isDateModified, "La modification de la date de l'Inventaire a echouée");
        }
        
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_DifferenceValue()
        {
            var homePage = LogInAsAdmin();

            var currency = TestContext.Properties["Currency"].ToString();
            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            InventoriesPage inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            inventoriesPage.ResetFilters();

            InventoryItem inventory = inventoriesPage.SelectFirstInventory();
            double theoValue = inventory.GetTheoricalValue(currency, decimalSeparatorValue);
            double phyValue = inventory.GetPhysicalValue(currency, decimalSeparatorValue);
            double diffValue = inventory.GetDifferenceValue(currency, decimalSeparatorValue);

            Assert.IsTrue(diffValue == phyValue - theoValue, "Le calcul affiché dans Difference Value n'est pas correct");
        }
        
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_DifferenceInPercentage()
        {
            string physQty = "5";

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var currency = TestContext.Properties["Currency"].ToString();
            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            InventoriesPage inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            inventoriesPage.ResetFilters();
            inventoriesPage.Filter(InventoriesPage.FilterType.ShowNotValidated, true);

            InventoryItem inventoryItem = inventoriesPage.SelectFirstInventory();
            var physicalQuantitie = inventoryItem.GetAllPhysicalQuantity();

            if (physicalQuantitie.All(qty => string.IsNullOrEmpty(qty)))
            {
                string itemName = inventoryItem.GetFirstItemName();
                inventoryItem.SelectFirstItem();
                inventoryItem.AddPhysicalQuantity(itemName, physQty);
            }

            double theoValue = inventoryItem.GetTheoricalValue(currency, decimalSeparatorValue);
            double diffValue = inventoryItem.GetDifferenceValue(currency, decimalSeparatorValue);
            double diffPercentage = inventoryItem.GetDifferencePercentage(decimalSeparatorValue);

            double calcDiffPercentage = inventoryItem.ParseDiffPercentage((diffValue / theoValue) * 100);
            Assert.AreEqual(diffPercentage, calcDiffPercentage, "Le calcul affiché dans Difference in Percentage n'est pas correct");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_Filter_Keyword()
        {
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();
            string itemKeyword = TestContext.Properties["Item_Keyword"].ToString();
            string itemKeywordNone = "_None";
            string qty = "10";
            HomePage homePage = LogInAsAdmin();
            InventoriesPage inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            InventoryCreateModalPage inventoryCreateModalPage = inventoriesPage.InventoryCreatePage();
            inventoryCreateModalPage.FillField_CreateNewInventory(DateUtils.Now, site, place);
            string inventoryNumber = inventoryCreateModalPage.GetInventoryNumber();
            InventoryItem inventoryItem = inventoryCreateModalPage.Submit();
            inventoryItem.Filter(FilterItemType.Keywords, itemKeyword);
            inventoryItem.SelectFirstItem();
            string itemName = inventoryItem.GetFirstItemName();
            inventoryItem.AddPhysicalQuantity(itemName, qty);
            inventoryItem.Refresh();
            inventoryItem.ResetFilter();
            inventoryItem.Filter(FilterItemType.Keywords, itemKeywordNone);
            inventoryItem.SelectFirstItem();
            itemName = inventoryItem.GetFirstItemName();
            inventoryItem.AddPhysicalQuantity(itemName, qty);
            inventoryItem.Refresh();
            inventoryItem.Validate();
            inventoryItem.WaitPageLoading();
            inventoryItem.CopyTheoricalQty(true);
            inventoriesPage = inventoryItem.BackToList();
            inventoriesPage.ResetFilters();
            inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, inventoryNumber);
            inventoryItem = inventoriesPage.SelectFirstInventory();
            int nbreItemsbeforeFilter = inventoryItem.GetNumberItems();
            inventoryItem.Filter(FilterItemType.Keywords, itemKeyword);
            int nbreItemsafterFilter = inventoryItem.GetNumberItems();
            Assert.AreNotEqual(nbreItemsbeforeFilter, nbreItemsafterFilter, "Le Filter_Keyword ne fonctionne pas.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_Filter_Subgroup()
        {
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();
            string item_Group = TestContext.Properties["Item_GroupBis"].ToString();
            string inventoryNumber = "";
            HomePage homePage = LogInAsAdmin();
            InventoriesPage inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            try
            {
                inventoriesPage.ResetFilters();
                InventoryCreateModalPage inventoryCreateModalPage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalPage.FillField_CreateNewInventory(DateUtils.Now, site, place);
                inventoryNumber = inventoryCreateModalPage.GetInventoryNumber();
                InventoryItem inventoryItem = inventoryCreateModalPage.Submit();
                inventoryItem.Filter(FilterItemType.ByGroup, item_Group);
                bool isVerified = inventoryItem.VerifyGroupFilter(item_Group);
                Assert.IsTrue(isVerified, "Le Filter_SubGroup ne fonctionne pas.");
                inventoryItem.BackToList();
                inventoriesPage.ResetFilters();
                inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, inventoryNumber);
            }
            finally
            {
                inventoriesPage.deleteInventory();
                inventoriesPage.ResetFilters();
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_Filter_ShowItemWithQtyDifference()
        {
            // prepare
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();

            //arrange
            var homePage = LogInAsAdmin();
            InventoriesPage inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            try
            {
                // Create New Inventory
                inventoriesPage.ResetFilters();
                InventoryCreateModalPage createInventory = inventoriesPage.InventoryCreatePage();
                createInventory.FillField_CreateNewInventory(DateUtils.Now, site, place);
                string ID = createInventory.GetInventoryNumber();
                InventoryItem inventoryItem = createInventory.Submit();

                // Select the First Item in the Created Inventory
                inventoryItem.SelectFirstItem();
                var itemName = inventoryItem.GetFirstItemName();

                // Add Physical Quantity & Physical Packaging Quantity
                inventoryItem.AddPhysicalQuantity(itemName, "10");
                inventoryItem.AddPhysicalPackagingQuantity(itemName, "10");

                // Filter show items with qty difference only
                inventoryItem.Filter(FilterItemType.ShowItemsWithDifferenceOnly, true);

                // Verify if the filter is set to show differences only
                bool isWithQtyDifferenceOnly = inventoryItem.IsWithQtyDifferenceOnly();
                // Assert
                Assert.IsTrue(isWithQtyDifferenceOnly, String.Format(MessageErreur.FILTRE_ERRONE, "'Show items with qty difference only'"));

                inventoryItem.BackToList();
                inventoriesPage.ResetFilters();
                inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, ID);
            }
            finally
            {
                // Delete the Inventory already created
                inventoriesPage.deleteInventory();
                inventoriesPage.ResetFilters();
            }

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_Filter_ShowItemWithThoeAndPhyQty()
        {
            HomePage homePage = LogInAsAdmin();
            InventoriesPage inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            inventoriesPage.ResetFilters();
            inventoriesPage.Filter(InventoriesPage.FilterType.ShowItemsNotValidated, true);
            inventoriesPage.Filter(InventoriesPage.FilterType.ShowNotValidated, true);
            InventoryItem inventoryItem = inventoriesPage.SelectFirstInventory();
            string itemName = inventoryItem.GetFirstItemName();
            inventoryItem.SelectFirstItem();
            inventoryItem.AddPhysicalQuantity(itemName, "10");
            inventoryItem.Filter(FilterItemType.ShowItemsWithQtyOnly, true);
            inventoryItem.PageSize("100");
            bool isWithTheoOrPhysQtyOnly = inventoryItem.IsWithTheoOrPhysQtyOnly();
            Assert.IsTrue(isWithTheoOrPhysQtyOnly, string.Format(MessageErreur.FILTRE_ERRONE, "'Show items with qty difference only'"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_Filter_ExoprtedForCegidAutomatically()
        {
            HomePage homePage = LogInAsAdmin();
            InventoriesPage inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            inventoriesPage.ResetFilters();
            inventoriesPage.Filter(InventoriesPage.FilterType.ShowAll, true);
            int nombreTotalInventory = inventoriesPage.CheckTotalNumber();
            inventoriesPage.Filter(InventoriesPage.FilterType.ExportedForSAGECEGIDAutomatically, true);
            int nombreExportedForSageCegidAuto = inventoriesPage.CheckTotalNumber();
            Assert.AreNotEqual(nombreTotalInventory, nombreExportedForSageCegidAuto, "Filtre ne fonctionne pas");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_Filter_Show_Sent_To_Cegid_Only()
        {
            HomePage homePage = LogInAsAdmin();
            InventoriesPage inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            inventoriesPage.ResetFilters();
            inventoriesPage.Filter(InventoriesPage.FilterType.ShowAll, true);
            int nombreTotalInventory = inventoriesPage.CheckTotalNumber();
            inventoriesPage.Filter(InventoriesPage.FilterType.ShowSentToSAGECEDIDOnly, true);
            int nombreSentToSAGECEDIDOnly = inventoriesPage.CheckTotalNumber();
            Assert.AreNotEqual(nombreTotalInventory, nombreSentToSAGECEDIDOnly, "Filtre ne fonctionne pas");
        }
        
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_Filter_ValidatedNotSentToCegid()
        {
            HomePage homePage = LogInAsAdmin();
            InventoriesPage inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            inventoriesPage.ResetFilters();
            inventoriesPage.Filter(InventoriesPage.FilterType.ShowValidatedNotSentToSAGE, true);
            inventoriesPage.PageSize("100");
            bool areAllInveValidatedAndNoCegid = inventoriesPage.AreAllInveValidatedAndNoCegid();
            Assert.IsTrue(areAllInveValidatedAndNoCegid, "Some Inventories rows in the filter result are either not validated or with a CEGID icon. It should be validated and no CEGID icon.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_TotalFlightIndex()
        {
            var homePage = LogInAsAdmin();
            InventoriesPage inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            inventoriesPage.ResetFilters();
            inventoriesPage.Filter(InventoriesPage.FilterType.ShowAll, true);
            var inventaireNumberShowAll = inventoriesPage.CheckTotalNumber();
            inventoriesPage.Filter(InventoriesPage.FilterType.ShowActive, true);
            var inventaireNumberShowActive = inventoriesPage.CheckTotalNumber();
            inventoriesPage.Filter(InventoriesPage.FilterType.ShowInactive, true);
            var inventaireNumberShowInactive = inventoriesPage.CheckTotalNumber();
            Assert.IsTrue(inventaireNumberShowAll == inventaireNumberShowActive + inventaireNumberShowInactive, MessageErreur.FILTRE_ERRONE);
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_ResetFilter()
        {
            HomePage homePage = LogInAsAdmin();
            InventoriesPage inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            inventoriesPage.ResetFilters();
            inventoriesPage.Filter(InventoriesPage.FilterType.ShowAll, true);
            int inventaireNumberShowAll = inventoriesPage.CheckTotalNumber();
            bool isNotEqual = inventaireNumberShowAll != 0;
            Assert.IsTrue(isNotEqual, MessageErreur.FILTRE_ERRONE);
            inventoriesPage.Filter(InventoriesPage.FilterType.ShowActive, true);
            int inventaireNumberShowActive = inventoriesPage.CheckTotalNumber();
            isNotEqual = inventaireNumberShowAll != inventaireNumberShowActive;
            Assert.IsTrue(isNotEqual, MessageErreur.FILTRE_ERRONE);
            inventoriesPage.Filter(InventoriesPage.FilterType.ShowInactive, true);
            int inventaireNumberShowInactive = inventoriesPage.CheckTotalNumber();
            isNotEqual = inventaireNumberShowAll != inventaireNumberShowInactive;
            Assert.IsTrue(isNotEqual, MessageErreur.FILTRE_ERRONE);
            inventoriesPage.Filter(InventoriesPage.FilterType.ShowItemsNotValidated, true);
            int inventaireNumberShowItemsNotValidated = inventoriesPage.CheckTotalNumber();
            isNotEqual = inventaireNumberShowAll != inventaireNumberShowItemsNotValidated;
            Assert.IsTrue(isNotEqual, MessageErreur.FILTRE_ERRONE);
            inventoriesPage.Filter(InventoriesPage.FilterType.ShowExportedForSageManually, true);
            int inventaireNumberShowExportedForSageManually = inventoriesPage.CheckTotalNumber();
            isNotEqual = inventaireNumberShowAll != inventaireNumberShowExportedForSageManually;
            Assert.IsTrue(isNotEqual, MessageErreur.FILTRE_ERRONE);
            inventoriesPage.Filter(InventoriesPage.FilterType.ShowValidatedNotSentToSAGE, true);
            int inventaireNumberShowValidatedNotSentToSAGE = inventoriesPage.CheckTotalNumber();
            isNotEqual = inventaireNumberShowAll != inventaireNumberShowValidatedNotSentToSAGE;
            Assert.IsTrue(isNotEqual, MessageErreur.FILTRE_ERRONE);
            inventoriesPage.Filter(InventoriesPage.FilterType.ExportedForSAGECEGIDAutomatically, true);
            int inventaireNumberExportedForSAGECEGIDAutomatically = inventoriesPage.CheckTotalNumber();
            isNotEqual = inventaireNumberShowAll != inventaireNumberExportedForSAGECEGIDAutomatically;
            Assert.IsTrue(isNotEqual, MessageErreur.FILTRE_ERRONE);
            inventoriesPage.Filter(InventoriesPage.FilterType.DateFrom, DateTime.Now.AddDays(5));
            int inventaireNumberDateFrom = inventoriesPage.CheckTotalNumber();
            isNotEqual = inventaireNumberShowAll != inventaireNumberDateFrom;
            Assert.IsTrue(isNotEqual, MessageErreur.FILTRE_ERRONE);
            inventoriesPage.Filter(InventoriesPage.FilterType.DateTo, DateTime.Now.AddDays(-5));
            int inventaireNumberDateTo = inventoriesPage.CheckTotalNumber();
            isNotEqual = inventaireNumberShowAll != inventaireNumberDateTo;
            Assert.IsTrue(isNotEqual, MessageErreur.FILTRE_ERRONE);
            inventoriesPage.Filter(InventoriesPage.FilterType.Site, TestContext.Properties["Site"].ToString());
            int inventaireNumberSite = inventoriesPage.CheckTotalNumber();
            isNotEqual = inventaireNumberShowAll != inventaireNumberSite;
            Assert.IsTrue(isNotEqual, MessageErreur.FILTRE_ERRONE);
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_Filter_Validation1()
        {
            //prepare
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();

            //arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            InventoriesPage inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            inventoriesPage.ResetFilters();
            var inventoryCreateModalPage = inventoriesPage.InventoryCreatePage();
            inventoryCreateModalPage.FillField_CreateNewInventory(DateUtils.Now, site, place);
            var inventoryItem = inventoryCreateModalPage.Submit();
            inventoryItem.Validate();
            inventoryItem.CopyTheoricalQty(true);
            var theoraticalQuantity = inventoryItem.GetTheoricalQuantity().Replace(" ", "");
            var physicalQuantity = inventoryItem.GetPhysicalQuantity();
            bool isPhysQtyUpdated = theoraticalQuantity.Equals(physicalQuantity);
            Assert.IsTrue(isPhysQtyUpdated, "Les quantités théoriques et les quantités physiques ne sont pas les mêmes.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_Filter_Validation11()
        {
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();

            var homePage = LogInAsAdmin();
            InventoriesPage inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            inventoriesPage.ResetFilters();
            var inventoryCreateModalPage = inventoriesPage.InventoryCreatePage();
            inventoryCreateModalPage.FillField_CreateNewInventory(DateUtils.Now, site, place);
            var inventoryItem = inventoryCreateModalPage.Submit();
            //inventoryItem.SetQtysToZero(true);
            //1.Cliquer sur l’icon verte
            //2.Cliquer sur “validate this inventory”
            var inventoryValidationModalPage = inventoryItem.Validate();

            //3.une pop up s’ouvre
            //4.Choisir l’option « Set physical quantity to zero for each of these items”
            //5.Cliquer sur “update items”
            inventoryItem = inventoryValidationModalPage.ValidateTotalInventory();

            //Après validation, les filtres sont inactif, et il y a aucune ligne d'affiché
            var physicalQuantity = inventoryItem.GetAllPhysicalQuantity();
            //Assert.AreEqual(0, physicalQuantity.Count);
            for (int i = 0; i < physicalQuantity.Count; i++)
            {
                Assert.AreEqual("0", physicalQuantity[i], "Physical Quantity ligne " + i + " n'est pas zero");
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_WMSExport()
        {
            //prepare
            string wmsInvetory = "WmsInventoryExportPath";

            //arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //act

            ApplicationSettingsPage applicationSettingsPage = homePage.GoToApplicationSettings();
            var isWMSInventoryExportExist = applicationSettingsPage.IsWMSInventoryExportExist(wmsInvetory);
            Assert.IsTrue(isWMSInventoryExportExist, "Le WMS pour l'inventaire n'existe pas");

            InventoriesPage inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            inventoriesPage.ResetFilters();
            var message = inventoriesPage.GetExportWMSMessage();

            //Assert
            var isContainsInventoryExport = message.Contains("WMSInventoryExportMultiSuccess");
            Assert.IsTrue(isContainsInventoryExport, "la message " + message + " est incorrect");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INVE_Items_Filter_ShowAllSeizableItem()
        {
            HomePage homePage = LogInAsAdmin();
            InventoriesPage inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            inventoriesPage.ResetFilters();
            inventoriesPage.Filter(InventoriesPage.FilterType.ShowNotValidated, true);
            InventoryItem inventoryItemPage = inventoriesPage.ClickFirstInventory();
            inventoryItemPage.Filter(FilterItemType.SearchByName, "aze");
            inventoryItemPage.Filter(FilterItemType.ShowAllSeizableItems, true);
            int nombreTotalShowAllSeizableItems = inventoryItemPage.CountInvenories();
            inventoryItemPage.Filter(FilterItemType.ShowItemsWithQtyOnly, true);
            int nombreShowItemsWithQtyOnly = inventoryItemPage.CountInvenories();
            inventoryItemPage.Filter(FilterItemType.ShowAllSeizableItems, true);
            int nombreTotalShowAllSeizableItemsAfterChangeFilter = inventoryItemPage.CountInvenories();
            Assert.AreEqual(nombreTotalShowAllSeizableItems, nombreTotalShowAllSeizableItemsAfterChangeFilter, "le filtre n'est pas appliqué correctement");
        }

        //________________________________________ Utilitaire ______________________________________________

        public bool SetInventoryDayValue(HomePage homePage)
        {

            var productionPage = homePage.GoToParameters_ProductionPage();
            productionPage.GoToTab_Inventory();

            // Récupération du numéro du jour en cours
            var jour = DateUtils.Now.Day;

            return productionPage.SetInventoryValue(jour);
        }

        public void RemoveInventoryValue(HomePage homePage)
        {
            var productionPage = homePage.GoToParameters_ProductionPage();
            productionPage.GoToTab_Inventory();

            // Récupération du numéro du jour en cours
            var jour = DateUtils.Now.Day;

            productionPage.RemoveInventoryValue(jour);
        }

        private string GetItemGroup(HomePage homePage, string itemName)
        {
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);

            var itemGeneralInfo = itemPage.ClickOnFirstItem();

            return itemGeneralInfo.GetGroupName();
        }

        private bool SetApplicationSettingsForSageAuto(HomePage homePage)
        {
            string environnement = TestContext.Properties["Winrest_Environnement"].ToString().ToUpper();

            try
            {
                var applicationSettings = homePage.GoToApplicationSettings();
                var versionBDD = applicationSettings.GetApplicationDbVersion();

                // Country code
                var appSettingsModalPage = applicationSettings.GetWinrestTLCountryCodePage();
                appSettingsModalPage.SetWinrestTLCountryCode(environnement);
                applicationSettings = appSettingsModalPage.Save();

                // BDD
                //appSettingsModalPage = applicationSettings.GetWinrestExportTLSageDbOverloadPage();
                //appSettingsModalPage.SetWinrestExportTLSageDbOverload(versionBDD);
                //applicationSettings = appSettingsModalPage.Save();

                // Override countryCode
                //appSettingsModalPage = applicationSettings.GetWinrestExportTLSageCountryCodeOverloadPage();
                //appSettingsModalPage.SetWinrestExportTLSageCountryCodeOverload(environnement);
                //appSettingsModalPage.Save();
            }
            catch
            {
                return false;
            }

            return true;
        }

        private bool VerifySiteAnalyticalPlanSection(HomePage homePage, string site, bool isOK = true)
        {
            string analyticalPlan = isOK ? "1" : "";
            string analyticalSection = isOK ? "314" : "";

            try
            {
                var settingsSitesPage = homePage.GoToParameters_Sites();
                settingsSitesPage.Filter(ParametersSites.FilterType.SearchSite, site);
                settingsSitesPage.ClickOnFirstSite();

                settingsSitesPage.ClickToInformations();
                settingsSitesPage.SetAnalyticPlan(analyticalPlan);
                settingsSitesPage.SetAnalyticSection(analyticalSection);
            }
            catch
            {
                return false;
            }

            return true;
        }

        public bool VerifyGroupAndVAT(HomePage homePage, string group, string vat, bool isOK = true)
        {
            // Prepare
            string account = "60105100";
            string exoAccount = "60105100";
            string inventoryAccount = isOK ? "60105100" : "";
            string inventoryVariationAccount = isOK ? "60105100" : "";

            try
            {
                // Act
                var accountingParametersPage = homePage.GoToParameters_AccountingPage();
                accountingParametersPage.GoToTab_GroupVats();

                if (!accountingParametersPage.IsGroupPresent(group))
                {
                    accountingParametersPage.CreateNewGroup(group, vat);
                }

                accountingParametersPage.SearchGroup(group, "");
                accountingParametersPage.EditInventoryAccounts(account, exoAccount, inventoryAccount, inventoryVariationAccount);
            }
            catch
            {
                return false;
            }

            return true;
        }

        private bool VerifyAccountingJournal(HomePage homePage, string site, string journalInventory)
        {
            try
            {
                var accountingJournalPage = homePage.GoToParameters_AccountingPage();
                accountingJournalPage.GoToTab_Journal();
                accountingJournalPage.EditJournal(site, null, null, null, journalInventory);
            }
            catch
            {
                return false;
            }

            return true;
        }

        public DateTime VerifyIntegrationDate(HomePage homePage)
        {
            // Act
            var accountingParametersPage = homePage.GoToParameters_AccountingPage();

            accountingParametersPage.GoToTab_MonthlyClosingDays();

            return accountingParametersPage.GetSageClosureMonthIndex();
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INV_AddNewExpiryDate()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();
            string value = "12";
            string inventoryNumber = "";

            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            inventoriesPage.ResetFilters();
            try
            {
                // Create
                var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, place);
                inventoryNumber = inventoryCreateModalpage.GetInventoryNumber();
                var inventoryItem = inventoryCreateModalpage.Submit();
                // Filter Proccess
                inventoryItem.Filter(FilterItemType.ShowItemsWithQtyOnly, true);
                inventoryItem.SelectFirstItem();
                var itemName = inventoryItem.GetFirstItemName();
                //Add physicalQuantity
                inventoryItem.AddPhysicalQuantity(itemName, value);
                InventoryExpiry expiry = inventoryItem.ShowFirstExpiryDate();
                expiry.AddExpiryDate("6,5", DateUtils.Now.AddDays(2));
                expiry.AddExpiryDate("4,3", DateUtils.Now.AddDays(3));
                expiry.AddExpiryDate("1,2", DateUtils.Now.AddDays(4));
                expiry.Save();
                //Assert
                Assert.IsTrue(inventoryItem.GetExpiryDateCssAfterSaves().Contains("green-text"), " impossible de faire La somme total de différentes quantités avec decimales.");
                //back to list
                inventoryItem.BackToList();

            }
            finally
            {
                inventoriesPage.ResetFilters();
                inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, inventoryNumber);
                inventoriesPage.deleteInventory();
                inventoriesPage.ResetFilters();
            }
        }

        /// <summary>
        /// Update inventory's first item allergens and check the state 
        /// of the allergens button in the inventory details.
        /// </summary>
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_INV_CheckAllergensUpdate()
        {
            string allergen1 = "Cacahuetes/Peanuts";
            string allergen2 = "Frutos de cáscara- Macadamias/Nuts-Macadamia";
            HomePage homePage = LogInAsAdmin();
            InventoriesPage inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            inventoriesPage.ResetFilters();
            inventoriesPage.Filter(InventoriesPage.FilterType.ShowNotValidated, true);
            string inventoryNumber = inventoriesPage.GetFirstInventoryNumber(); // change here
            inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, inventoryNumber);
            InventoryItem inventory = inventoriesPage.SelectFirstInventory();
            string itemName = inventory.GetFirstItemName();
            inventory.ResetFilter();
            inventory.Filter(InventoryItem.FilterItemType.SearchByName, itemName);
            inventory.SelectFirstItem();
            inventory.AddPhysicalQuantity(itemName, "2");
            inventory.Refresh();
            inventory.ResetFilter();
            inventory.Filter(InventoryItem.FilterItemType.SearchByName, itemName);
            inventory.SelectFirstItem();
            ItemGeneralInformationPage itemPage = inventory.EditItem(itemName);
            ItemIntolerancePage itemIntolerancePage = itemPage.ClickOnIntolerancePage();
            itemIntolerancePage.AddAllergen(allergen1);
            itemIntolerancePage.AddAllergen(allergen2);
            inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            inventoriesPage.ResetFilters();
            inventoriesPage.Filter(InventoriesPage.FilterType.ShowNotValidated, true);
            inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, inventoryNumber);
            inventory = inventoriesPage.SelectFirstInventory();
            inventory.ResetFilter();
            inventory.Filter(InventoryItem.FilterItemType.SearchByName, itemName);
            bool isIconGreen = inventory.IsAllergenIconGreen(itemName);
            List<string> allergensInItem = inventory.GetAllergens(itemName);
            bool containsAllergen1 = allergensInItem.Contains(allergen1);
            bool containsAllergen2 = allergensInItem.Contains(allergen2);
            Assert.IsTrue(isIconGreen, "L'icon n'est pas vert!");
            Assert.IsTrue(containsAllergen1 && containsAllergen2, "Allergens n'ont pas été ajoutés");
        }

        public void EnsureTheoricalQuantity()
        {
            var homePage = new HomePage(WebDriver, TestContext);
            var inventoriesPage = new InventoriesPage(WebDriver, TestContext);
            var inventoryDetailPage = new InventoryItem(WebDriver, TestContext);
            string site = TestContext.Properties["Site"].ToString();
            string delivery = TestContext.Properties["PlaceFrom"].ToString();
            string quantity = "5";
            var inventNumber = inventoryDetailPage.GetInventoryNumber();

            var firstItem = inventoryDetailPage.GetFirstItemName();
            inventoryDetailPage.Filter(FilterItemType.SearchByName, firstItem);
            inventoryDetailPage.SelectFirstItem();
            var editItem = inventoryDetailPage.EditItem(firstItem);
            var supplier = editItem.GetPackagingSupplierBySite(site);
            homePage.Navigate();
            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            var receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
            receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, delivery));
            var receiptNotesItem = receiptNotesCreateModalpage.Submit();
            receiptNotesItem.SetFirstReceivedQuantity(quantity);
            var checksTab = receiptNotesItem.ClickOnChecksTab();
            checksTab.SetNotApplicable();
            checksTab.DeliveryAccepted();
            checksTab.ValidateQualityChecks();
            checksTab.WaitPageLoading();
            homePage.Navigate();
            inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            inventoriesPage.ResetFilters();
            inventoriesPage.Filter(InventoriesPage.FilterType.ByNumber, inventNumber);
            inventoryDetailPage = inventoriesPage.ClickFirstInventory();
        }
    }
}