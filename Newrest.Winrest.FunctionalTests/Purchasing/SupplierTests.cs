using DocumentFormat.OpenXml.Office2016.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using iText.Layout.Properties;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Delivery;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.Flight;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Recipes;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Production;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Sites;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Item;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Suppliers;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.PageObjects.ToDoList.Scheduler;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.Inventory;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.ReceiptNotes;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Policy;
using System.Threading;
using System.Web;
using System.Web.UI;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using static Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.ItemUseCasePage;
using Page = UglyToad.PdfPig.Content.Page;

namespace Newrest.Winrest.FunctionalTests.Purchasing
{
    [TestClass]
    public class SupplierTests : TestBase
    {
        private static Random random = new Random();
        private string supplierName = "Supplier_" + DateUtils.Now.ToString("dd/MM/yyyy") + "_" + random.Next(1000, 9999);
        private string itemName1 = "1-suppItem" + random.Next().ToString();
        private string itemName2 = "2-suppItem" + random.Next().ToString();
        private string supplierType = "MonSupplierType";
        private const int _timeout = 600000;
        string nameItem = "";
        /// <summary>
        /// 
        /// Mise en place du paramétrage pour la configuration Winrest 4.0 
        /// 
        /// </summary>
        /// 
        [TestInitialize]
        public override void TestInitialize()
        {
            base.TestInitialize();

            var testMethod = TestContext.TestName;
            switch (testMethod)
            {
                case nameof(PU_SUPP_FilterSupplierType):
                    TestInitialize_Create_SupplierType();
                    TestInitialize_Create_Supplier();
                    break;

                case nameof(PU_SUPP_Deliv_TimeLimit):
                    TestInitialize_Create_SupplierType();
                    TestInitialize_Create_Supplier();
                    break;
                case nameof(PU_SUPP_DELI_MINAmount):
                    TestInitialize_Create_SupplierType();
                    TestInitialize_Create_Supplier();
                    break;
                case nameof(PU_SUPP_Storage_FilterItemName):
                    TestInitialize_Create_SupplierType();
                    TestInitialize_Create_Supplier();
                    TestInitialize_Create_Two_Items_With_Inventory(); 
                    break;
                case nameof(PU_SUPP_DELI_ShippingCost):
                    TestInitialize_Create_SupplierType();
                    TestInitialize_Create_Supplier();
                    break;
              
                 
                default:
                    break;
            }
        }

        public void TestInitialize_Create_SupplierType()
        {
            HomePage homePage = LogInAsAdmin();
            ParametersPurchasing purchasing = homePage.GoToParameters_PurchasingPage();
            purchasing.GoToTab_Supplier();
            bool IsExistingType = purchasing.CheckForExistingType(supplierType);
            if (!IsExistingType)
            {
                purchasing.CreateNewType(supplierType);
            }
            bool isSupplierTypeCreated = purchasing.CheckForExistingType(supplierType);
            Assert.IsTrue(isSupplierTypeCreated, "Le type de fournisseur n'est pas créé.");
        }
        public void TestInitialize_Create_Supplier()
        {
            //Prepare
            bool isActivated = true;
            bool isAudited = false;
            //Arrange
            HomePage homePage = LogInAsAdmin();
            //Act
            SuppliersPage suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            SupplierCreateModalPage supplierCreateModalpage = suppliersPage.SupplierCreatePage();
            supplierCreateModalpage.FillField_CreatNewSupplier(supplierName, isActivated, isAudited);
            SupplierItem supplierItem = supplierCreateModalpage.Submit();
            SupplierGeneralInfoTab suppliergeneralInfo = supplierItem.ClickOnGeneralInformationTab();
            suppliergeneralInfo.SetSupplierType(supplierType);
            CreateNewItem(supplierItem, supplierName);
            suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, supplierName);
            var displayedSupplier = suppliersPage.GetFirstSupplierName();
            Assert.AreEqual(supplierName, displayedSupplier, "Le supplier créé n'est pas présent dans la liste.");
        }
        public void TestInitialize_Create_Two_Items_With_Inventory()
        {
            //Prepare 
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string storageUnit = "KG";
            string supplierRef = itemName1 + "_SupplierRef";
            string storageQty = 2.32.ToString();
            string storageQty1 = 4.2.ToString();
            string limitQty = 10.ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();
            DateTime date = DateUtils.Now;

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
            //Create Item 1
            ItemCreateModalPage itemCreatePage = itemPage.ItemCreatePage();
            ItemGeneralInformationPage itemGeneralInformationPage = itemCreatePage.FillField_CreateNewItem(itemName1, group, workshop, taxType, prodUnit);
            ItemCreateNewPackagingModalPage itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, storageQty1, supplierName, null, supplierRef, storageQty);

            //Create Item 2
            itemPage = homePage.GoToPurchasing_ItemPage();
            itemCreatePage = itemPage.ItemCreatePage();
            itemGeneralInformationPage = itemCreatePage.FillField_CreateNewItem(itemName2, group, workshop, taxType, prodUnit);
            itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, storageQty1, supplierName, null, supplierRef, storageQty);

            //Create Inventory for Item 1
            InventoriesPage inventoryPage = homePage.GoToWarehouse_InventoriesPage();
            InventoryCreateModalPage inventoryCreateModalPage = inventoryPage.InventoryCreatePage();
            inventoryCreateModalPage.FillField_CreateNewInventory(date, site, place);
            InventoryItem itemInventory = inventoryCreateModalPage.Submit();
            itemInventory.Filter(InventoryItem.FilterItemType.SearchByName, itemName1);
            itemInventory.SelectFirstItem();
            itemInventory.AddPhysicalQuantity(itemName1, limitQty);
            InventoryValidationModalPage modal1 = itemInventory.Validate();
            modal1.ValidatePartialInventory();

            //Create Inventory for Item 2
            inventoryPage = homePage.GoToWarehouse_InventoriesPage();
            inventoryCreateModalPage = inventoryPage.InventoryCreatePage();
            inventoryCreateModalPage.FillField_CreateNewInventory(date, site, place);
            itemInventory = inventoryCreateModalPage.Submit();
            itemInventory.Filter(InventoryItem.FilterItemType.SearchByName, itemName2);
            itemInventory.SelectFirstItem();
            itemInventory.AddPhysicalQuantity(itemName2, limitQty);
            InventoryValidationModalPage modal2 = itemInventory.Validate();
            modal2.ValidatePartialInventory();
        }
        [TestMethod]
        [Priority(0)]
        [Timeout(_timeout)]
        public void PU_SUPP_SetConfigWinrest4_0()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            ClearCache();

            // New version search
            homePage.SetNewVersionSearchValue(true);

            // New VersionItemDetails
            //homePage.SetVersionItemDetailValue(true);

            // Vérifier que c'est activé
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            var firstSupplier = suppliersPage.GetFirstSupplierName();

            suppliersPage.Filter(SuppliersPage.FilterType.Search, firstSupplier);
            Assert.AreEqual(1, suppliersPage.CheckTotalNumber(), "La recherche a pu être effectuée, le NewSearchMode est actif.");

            // vérifier new version item detail
            try
            {
                var itemPage = homePage.GoToPurchasing_ItemPage();
                itemPage.ResetFilter();
                var itemDetails = itemPage.ClickOnFirstItem();
                itemDetails.SearchPackaging(site);
            }
            catch
            {
                throw new Exception("La recherche d'un packaging n'a pas pu être effectuée, le NewVersionItemDetails est inactif.");
            }
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_Create_Supplier()
        {
            //Prepare
            string random_supplier_name = "test-supplier-" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            var supplierCreateModalpage = suppliersPage.SupplierCreatePage();
            supplierCreateModalpage.FillField_CreatNewSupplier(random_supplier_name, true, false);
            var supplierItem = supplierCreateModalpage.Submit();

            CreateNewItem(supplierItem, random_supplier_name);

            suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, random_supplier_name);

            // Assert
            Assert.AreEqual(random_supplier_name, suppliersPage.GetFirstSupplierName(), "Le supplier créé n'est pas présent dans la liste.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_Create_Supplier_General_Informations()
        {
            string random_supplier_name = "test-supplier-" + new Random().Next().ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();

            // Create
            var supplierCreateModalpage = suppliersPage.SupplierCreatePage();
            supplierCreateModalpage.FillField_CreatNewSupplier(random_supplier_name, true, true);
            var supplierItem = supplierCreateModalpage.Submit();

            // Get the value of SanitaryNumber and SiretNumber of the supplier
            string initSanitaryNumber = supplierItem.SanitaryNumber();
            string initSiretNumber = supplierItem.SiretNumber();
            string initLastAuditDate = supplierItem.LastAuditDate();
            string initNextAuditDate = supplierItem.NextAuditDate();

            // Change supplier infos
            supplierItem.CompletGeneralInformations();

            CreateNewItem(supplierItem, random_supplier_name);

            suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();

            // Search for the created supplier to get the new values of SanitaryNumber and SiretNumber
            suppliersPage.Filter(SuppliersPage.FilterType.Search, random_supplier_name);
            suppliersPage.SelectFirstItem();

            string newSanitaryNumber = supplierItem.SanitaryNumber();
            string newSiretNumber = supplierItem.SiretNumber();
            string newLastAuditDate = supplierItem.LastAuditDate();
            string newNextAuditDate = supplierItem.NextAuditDate();

            //Assert
            Assert.AreNotEqual(initSanitaryNumber, newSanitaryNumber, "La valeur du SanitaryNumber n'a pas été modifiée.");
            Assert.AreNotEqual(initSiretNumber, newSiretNumber, "La valeur du SiretNumber n'a pas été modifiée.");
            Assert.AreNotEqual(initLastAuditDate, newLastAuditDate, "La valeur du LastAuditDate n'a pas été modifiée.");
            Assert.AreNotEqual(initNextAuditDate, newNextAuditDate, "La valeur du NextAuditDate n'a pas été modifiée.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_UpdateInfos_WithImport()
        {
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string random_supplier_name = "test-supplier-" + new Random().Next().ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();

            suppliersPage.ClearDownloads();

            // Create
            var SupplierCreateModalpage = suppliersPage.SupplierCreatePage();
            SupplierCreateModalpage.FillField_CreatNewSupplier(random_supplier_name, true, false);
            var supplierItem = SupplierCreateModalpage.Submit();

            // 1. New Item
            CreateNewItem(supplierItem, random_supplier_name);

            suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();

            suppliersPage.Filter(SuppliersPage.FilterType.Search, random_supplier_name);
            suppliersPage.SelectFirstItem();

            supplierItem.GetItemsTab();
            supplierItem.ResetFilter();

            Assert.AreNotEqual("", supplierItem.getItemName(), "L'item n'a pas été ajouté au supplier.");


            // 2. Export Item
            var initValue = supplierItem.getItemName();
            string new_item_name = initValue + "_changed_From_My_Import";

            // Export of the supplier items
            supplierItem.ExportItems();

            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = supplierItem.GetExportExcelFileItem(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, MessageErreur.FICHIER_NON_TROUVE);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            var resultNumber = OpenXmlExcel.GetExportResultNumber("Sheet 1", filePath);
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);


            //3. Import Item
            OpenXmlExcel.WriteDataInCell("Action", "Supplier", random_supplier_name, "Sheet 1", filePath, "M", CellValues.SharedString);
            OpenXmlExcel.WriteDataInCell("Item name", "Supplier", random_supplier_name, "Sheet 1", filePath, new_item_name, CellValues.SharedString);

            var importPopup = supplierItem.Import();
            bool isImported = importPopup.ImportFile(correctDownloadedFile.FullName);
            Assert.IsTrue(isImported, "Erreur lors de l'import du fichier.");

            supplierItem.Filter(SupplierItem.FilterType.Search, new_item_name);
            var newValue = supplierItem.getItemName();
            Assert.AreNotEqual(initValue, newValue, "Le nom de l'item n'a pas été modifié.");


            //4. Set Unpurchasable Items
            supplierItem.ResetFilter();
            supplierItem.SetUnpurchasableItems();
            supplierItem.ResetFilter();
            supplierItem.Filter(SupplierItem.FilterType.Search, new_item_name);
            Assert.IsFalse(supplierItem.IsItemPresent(), "L'item n'est pas unpurchasable.");

            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.ShowAll, true);
            itemPage.Filter(ItemPage.FilterType.Search, new_item_name);


            var itemGeneralInformationPage = itemPage.ClickOnFirstItem();
            itemGeneralInformationPage.SetPurchasable(true);

            suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();

            suppliersPage.Filter(SuppliersPage.FilterType.Search, random_supplier_name);
            suppliersPage.SelectFirstItem();

            supplierItem.GetItemsTab();
            supplierItem.ResetFilter();
            supplierItem.Filter(SupplierItem.FilterType.Search, new_item_name);

            Assert.AreNotEqual("", supplierItem.getItemName(), "L'item n'est pas purchasable de nouveau.");

            // 5. DeactivateItem
            supplierItem.DeactivateItems();
            supplierItem.ResetFilter();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, new_item_name);
            Assert.IsFalse(supplierItem.IsItemPresent(), "L'item est toujours présent malgré le fait qu'il soit désactivé.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_Supplier_Deliveries()
        {
            //Prepare
            string comment = TestContext.Properties["CommentSupplier"].ToString();
            string random_supplier_name = "test-supplier-" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            var supplierCreateModalpage = suppliersPage.SupplierCreatePage();
            supplierCreateModalpage.FillField_CreatNewSupplier(random_supplier_name, false, false);
            var supplierItem = supplierCreateModalpage.Submit();

            // Add of deliveryDays and Comment
            supplierItem.CompletDeliveriesDaysAndComment();
            supplierItem.ClickToDeliveryBtn();

            bool isAreDaysAdded = supplierItem.AreDaysAdded();
            bool isDeliveryCommentFilled = supplierItem.IsDeliveryCommentFilled(comment);
            //Assert
            Assert.IsTrue(isAreDaysAdded, "Tous les jours n'ont pas été ajoutés.");
            Assert.IsTrue(isDeliveryCommentFilled, "Le commentaire n'a pas correctement été ajouté.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_Supplier_Contacts()
        {
            //Prepare
            string random_supplier_name = "test-supplier-" + new Random().Next().ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            var supplierCreateModalpage = suppliersPage.SupplierCreatePage();
            supplierCreateModalpage.FillField_CreatNewSupplier(random_supplier_name, false, false);
            var supplierItem = supplierCreateModalpage.Submit();


            supplierItem.GoToContactTab();
            if (!supplierItem.RowContactsIsVisible())
            {
                // Create a new contact if contact details are not visible
                supplierItem.AddContact("test@mail.com");
            }
            supplierItem.ClickUnfoldAllButton();
            Assert.IsTrue(supplierItem.DetailsContactsIsVisible(), "Les détails des contacts devraient être affichés après 'Unfold All'");
            supplierItem.ClickFoldAllButton();
            Assert.IsFalse(supplierItem.DetailsContactsIsVisible(), "Les détails des contacts devraient être cachés après 'Fold All'");

            supplierItem.DeleteAllContacts();
            // Add a contact to the supplier
            var nameGenerate = supplierItem.AddContact("test@mail.com");

            //Assert
            Assert.AreEqual(nameGenerate + " ", supplierItem.GetFirstContact(), "Le contact n'a pas été ajouté au supplier.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_Supplier_Accounting()
        {
            //Prepare
            string random_supplier_name = "test-supplier-" + new Random().Next().ToString();
            string site = TestContext.Properties["Site"].ToString();

            string thirdId = "10";
            string accountingId = "10";
            string thirdIdDTI = "10";
            string accountingIdDTI = "10";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            var supplierCreateModalpage = suppliersPage.SupplierCreatePage();
            supplierCreateModalpage.FillField_CreatNewSupplier(random_supplier_name, false, false);
            var supplierItem = supplierCreateModalpage.Submit();

            // Go to accounting tab to add an accounting
            supplierItem.ClickOnAccountingTab();

            supplierItem.CreateNewAccounting(thirdId, accountingId, thirdIdDTI, accountingIdDTI);
            supplierItem.SearchAccounting(site);
            supplierItem.EditAccounting(thirdId, accountingId, thirdIdDTI, accountingIdDTI);

            //Assert
            Assert.IsTrue(supplierItem.IsAccountingPresent(site), "La comptabilité fournisseur n'a pas été ajoutée.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_Supplier_Certification()
        {
            //Prepare
            string certification = TestContext.Properties["Certification"].ToString();
            string random_supplier_name = "test-supplier-" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            var supplierCreateModalpage = suppliersPage.SupplierCreatePage();
            supplierCreateModalpage.FillField_CreatNewSupplier(random_supplier_name, false, false);
            var supplierItem = supplierCreateModalpage.Submit();

            // Go to certification tab
            supplierItem.GetCertificationTab();

            bool addNewCertif = supplierItem.AddNewCertification(certification);
            bool delFirstCertif = supplierItem.DelFirstCertification();
            //Assert
            Assert.IsTrue(addNewCertif, "La certification n'a pas été ajoutée.");
            Assert.IsTrue(delFirstCertif, "La certification n'a pas été supprimée.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_Supplier_Fold_Unfold_Items()
        {
            //Prepare
            string random_supplier_name = "test-supplier-" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            var supplierCreateModalpage = suppliersPage.SupplierCreatePage();
            supplierCreateModalpage.FillField_CreatNewSupplier(random_supplier_name, true, false);
            var supplierItem = supplierCreateModalpage.Submit();

            // Creation of an item
            CreateNewItem(supplierItem, random_supplier_name);

            suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();

            suppliersPage.Filter(SuppliersPage.FilterType.Search, random_supplier_name);
            suppliersPage.SelectFirstItem();

            supplierItem.GetItemsTab();

            // Assert - unfold
            supplierItem.ResetFilter();
            // hack ticket 12582 removed
            Assert.IsTrue(supplierItem.IsFoldAll(), "Le détail des items n'est pas masqué.");
            supplierItem.UnFoldorFoldItems();
            // FoldAll => UnFoldAll => FoldAll

            // Assert - fold
            supplierItem.UnFoldorFoldItems();
            Assert.IsTrue(supplierItem.IsFoldAll(), "Le détail des items n'est pas masqué...");
        }

        // _____________________________________________ Filtres __________________________________________________
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_Filter_Search()
        {
            //Prepare
            string random_supplier_name = "test-supplier-" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            // Create
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            var supplierCreateModalpage = suppliersPage.SupplierCreatePage();
            supplierCreateModalpage.FillField_CreatNewSupplier(random_supplier_name, true, false);
            var supplierItem = supplierCreateModalpage.Submit();

            CreateNewItem(supplierItem, random_supplier_name);

            suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();

            if (!suppliersPage.isPageSizeEqualsTo100())
            {
                suppliersPage.PageSize("8");
                suppliersPage.PageSize("100");
            }

            suppliersPage.Filter(SuppliersPage.FilterType.Search, random_supplier_name);

            string firstSupplierName = suppliersPage.GetFirstSupplierName();
            //Assert
            Assert.AreEqual(firstSupplierName, random_supplier_name, String.Format(MessageErreur.FILTRE_ERRONE, "'Search by name'"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_Filter_SortByName()
        {
            //Prepare
            string random_supplier_name = "test-supplier-" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            // Create
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();

            if (suppliersPage.CheckTotalNumber() < 20)
            {
                var supplierCreateModalpage = suppliersPage.SupplierCreatePage();
                supplierCreateModalpage.FillField_CreatNewSupplier(random_supplier_name, true, false);
                var supplierItem = supplierCreateModalpage.Submit();

                CreateNewItem(supplierItem, random_supplier_name);

                suppliersPage = homePage.GoToPurchasing_SuppliersPage();
                suppliersPage.ResetFilter();
            }

            if (!suppliersPage.isPageSizeEqualsTo100())
            {
                suppliersPage.PageSize("8");
                suppliersPage.PageSize("100");
            }

            suppliersPage.Filter(SuppliersPage.FilterType.SortBy, "Sort by name");

            bool isSortedByName = suppliersPage.IsSortedByName();
            //Assert
            Assert.IsTrue(isSortedByName, String.Format(MessageErreur.FILTRE_ERRONE, "'Sort by name'"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_Filter_SortByEndOfContract()
        {
            // Prepare
            string random_supplier_name = "test-supplier-" + new Random().Next().ToString();
            string downloadPath = TestContext.Properties["DownloadsPath"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            bool newVersionPrint = true;

            // Log in
            var homePage = LogInAsAdmin();

            // Create
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            suppliersPage.Filter(SuppliersPage.FilterType.Site, site);//filter to get less supllier for export

            suppliersPage.ClearDownloads();

            if (suppliersPage.CheckTotalNumber() < 20)
            {
                var supplierCreateModalpage = suppliersPage.SupplierCreatePage();
                supplierCreateModalpage.FillField_CreatNewSupplier(random_supplier_name, true, false);
                var supplierItem = supplierCreateModalpage.Submit();

                CreateNewItem(supplierItem, random_supplier_name);

                suppliersPage = homePage.GoToPurchasing_SuppliersPage();
                suppliersPage.ResetFilter();
                suppliersPage.Filter(SuppliersPage.FilterType.Site, site);
            }
            suppliersPage.Filter(SuppliersPage.FilterType.SortBy, "Sort by end of contract");

            DirectoryInfo taskDirectory = new DirectoryInfo(downloadPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();
            foreach (FileInfo fi in taskFiles)
            {
                //purge
                fi.Delete();
            }
            // On exporte les résultats sous la forme d'un fichier Excel (dont on récupère le nom)
            // Export du fichier au format Excel
            suppliersPage.Export(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            taskDirectory = new DirectoryInfo(downloadPath);
            taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = suppliersPage.GetExportExcelFile(taskFiles, true);
            Assert.IsNotNull(correctDownloadedFile);

            // Récupération du nom du fichier et construction de l'URL du fichier Excel à ouvrir   
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadPath, fileName);

            int resultNumber = OpenXmlExcel.GetExportResultNumber("Suppliers", filePath);
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);

            List<string> totalList = OpenXmlExcel.GetValuesInList("End Of Contract", "Suppliers", filePath);
            Assert.IsTrue(suppliersPage.IsListSorted(totalList), String.Format(MessageErreur.FILTRE_ERRONE, "Sort by end of contract"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_Filter_SortByInternalID()
        {
            // Prepare
            string random_supplier_name = "test-supplier-" + new Random().Next().ToString();
            string downloadPath = TestContext.Properties["DownloadsPath"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            bool newVersionPrint = true;

            // Log in
            var homePage = LogInAsAdmin();

            // Create
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            suppliersPage.Filter(SuppliersPage.FilterType.Site, site);//filter to get less supllier for export

            suppliersPage.ClearDownloads();

            if (suppliersPage.CheckTotalNumber() < 20)
            {
                var supplierCreateModalpage = suppliersPage.SupplierCreatePage();
                supplierCreateModalpage.FillField_CreatNewSupplier(random_supplier_name, true, false);
                var supplierItem = supplierCreateModalpage.Submit();

                CreateNewItem(supplierItem, random_supplier_name);

                suppliersPage = homePage.GoToPurchasing_SuppliersPage();
                suppliersPage.ResetFilter();
                suppliersPage.Filter(SuppliersPage.FilterType.Site, site);//filter to get less supllier for export
            }

            if (!suppliersPage.isPageSizeEqualsTo100())
            {
                suppliersPage.PageSize("8");
                suppliersPage.PageSize("100");
            }

            suppliersPage.Filter(SuppliersPage.FilterType.SortBy, "Sort by internal ID");


            DirectoryInfo taskDirectory = new DirectoryInfo(downloadPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();
            foreach (FileInfo fi in taskFiles)
            {
                //purge
                fi.Delete();
            }
            // On exporte les résultats sous la forme d'un fichier Excel (dont on récupère le nom)
            // Export du fichier au format Excel
            suppliersPage.Export(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            taskDirectory = new DirectoryInfo(downloadPath);
            taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = suppliersPage.GetExportExcelFile(taskFiles, true);
            Assert.IsNotNull(correctDownloadedFile);

            // Récupération du nom du fichier et construction de l'URL du fichier Excel à ouvrir   
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadPath, fileName);

            int resultNumber = OpenXmlExcel.GetExportResultNumber("Suppliers", filePath);
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);

            List<string> totalList = OpenXmlExcel.GetValuesInList("Supplier No", "Suppliers", filePath);
            Assert.IsTrue(suppliersPage.IsListSorted(totalList), String.Format(MessageErreur.FILTRE_ERRONE, "Sort by internal ID"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_Filter_ShowAll()
        {
            //Prepare           
            string random_supplier_name = "new-test-supplier-" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();

            if (suppliersPage.CheckTotalNumber() < 20)
            {
                var supplierCreateModalpage = suppliersPage.SupplierCreatePage();
                supplierCreateModalpage.FillField_CreatNewSupplier(random_supplier_name, true, false);
                var supplierItem = supplierCreateModalpage.Submit();

                CreateNewItem(supplierItem, random_supplier_name);

                suppliersPage = homePage.GoToPurchasing_SuppliersPage();
                suppliersPage.ResetFilter();
            }

            if (!suppliersPage.isPageSizeEqualsTo100())
            {
                suppliersPage.PageSize("8");
                suppliersPage.PageSize("100");
            }

            suppliersPage.Filter(SuppliersPage.FilterType.ShowOnlyActive, true);
            var nbr1 = suppliersPage.CheckTotalNumber();

            suppliersPage.Filter(SuppliersPage.FilterType.ShowOnlyInactive, true);
            var nbr2 = suppliersPage.CheckTotalNumber();

            suppliersPage.Filter(SuppliersPage.FilterType.ShowAll, true);
            var realNbr = suppliersPage.CheckTotalNumber();

            //Assert
            Assert.AreEqual(nbr1 + nbr2, realNbr, String.Format(MessageErreur.FILTRE_ERRONE, "Show all"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_Filter_ShowOnlyActive()
        {
            string random_supplier_name = "test-supplier-" + new Random().Next().ToString();
            int nbrAllSupplier, nbrActiveSuppliers, nbrInactiveSupplier;
            HomePage homePage = LogInAsAdmin();
            SuppliersPage suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            suppliersPage.Filter(SuppliersPage.FilterType.ShowOnlyActive, true);
            nbrActiveSuppliers = suppliersPage.CheckTotalNumber();
            if (nbrActiveSuppliers == 0)
            {
                SupplierCreateModalPage supplierCreateModalpage = suppliersPage.SupplierCreatePage();
                supplierCreateModalpage.FillField_CreatNewSupplier(random_supplier_name, true, false);
                SupplierItem supplierItem = supplierCreateModalpage.Submit();
                CreateNewItem(supplierItem, random_supplier_name);
                suppliersPage = homePage.GoToPurchasing_SuppliersPage();
                suppliersPage.ResetFilter();
                suppliersPage.Filter(SuppliersPage.FilterType.ShowOnlyActive, true);
                nbrActiveSuppliers = suppliersPage.CheckTotalNumber();
            }

            suppliersPage.Filter(SuppliersPage.FilterType.ShowOnlyInactive, true);
            nbrInactiveSupplier = suppliersPage.CheckTotalNumber();
            if (nbrInactiveSupplier == 0)
            {
                var supplierCreateModalpage = suppliersPage.SupplierCreatePage();
                supplierCreateModalpage.FillField_CreatNewSupplier(random_supplier_name, true, false);
                var supplierItem = supplierCreateModalpage.Submit();
                CreateNewItem(supplierItem, random_supplier_name);
                suppliersPage = homePage.GoToPurchasing_SuppliersPage();
                suppliersPage.ResetFilter();
                suppliersPage.Filter(SuppliersPage.FilterType.Search, random_supplier_name);
                suppliersPage.SelectFirstItem();
                supplierItem.GetItemsTab();
                supplierItem.DeactivateItems();
                supplierItem.ClickOnGeneralInformationTab();
                supplierItem.DeactivateSupplier();
                suppliersPage = supplierItem.BackToList();
                suppliersPage.ResetFilter();
                suppliersPage.Filter(SuppliersPage.FilterType.ShowOnlyInactive, true);
            }
            suppliersPage.Filter(SuppliersPage.FilterType.ShowAll, true);
            nbrAllSupplier = suppliersPage.CheckTotalNumber();
            Assert.AreEqual(nbrAllSupplier - nbrInactiveSupplier, nbrActiveSuppliers, String.Format(MessageErreur.FILTRE_ERRONE, "'Show only active'"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_Filter_ShowOnlyInactive()
        {
            string random_supplier_name = "test-supplier-" + new Random().Next().ToString();
            int nbrAllSupplier, nbrActiveSuppliers, nbrInactiveSupplier;
            HomePage homePage = LogInAsAdmin();
            SuppliersPage suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            suppliersPage.Filter(SuppliersPage.FilterType.ShowOnlyInactive, true);
            nbrInactiveSupplier = suppliersPage.CheckTotalNumber();
            if (nbrInactiveSupplier == 0)
            {
                var supplierCreateModalpage = suppliersPage.SupplierCreatePage();
                supplierCreateModalpage.FillField_CreatNewSupplier(random_supplier_name, true, false);
                var supplierItem = supplierCreateModalpage.Submit();
                CreateNewItem(supplierItem, random_supplier_name);
                suppliersPage = homePage.GoToPurchasing_SuppliersPage();
                suppliersPage.ResetFilter();
                suppliersPage.Filter(SuppliersPage.FilterType.Search, random_supplier_name);
                suppliersPage.SelectFirstItem();
                supplierItem.GetItemsTab();
                supplierItem.DeactivateItems();
                supplierItem.ClickOnGeneralInformationTab();
                supplierItem.DeactivateSupplier();
                suppliersPage = supplierItem.BackToList();
                suppliersPage.ResetFilter();
                suppliersPage.Filter(SuppliersPage.FilterType.ShowOnlyInactive, true);
                nbrInactiveSupplier = suppliersPage.CheckTotalNumber();
            }
            suppliersPage.Filter(SuppliersPage.FilterType.ShowOnlyActive, true);
            nbrActiveSuppliers = suppliersPage.CheckTotalNumber();
            if (nbrActiveSuppliers == 0)
            {
                SupplierCreateModalPage supplierCreateModalpage = suppliersPage.SupplierCreatePage();
                supplierCreateModalpage.FillField_CreatNewSupplier(random_supplier_name, true, false);
                SupplierItem supplierItem = supplierCreateModalpage.Submit();
                CreateNewItem(supplierItem, random_supplier_name);
                suppliersPage = homePage.GoToPurchasing_SuppliersPage();
                suppliersPage.ResetFilter();
                suppliersPage.Filter(SuppliersPage.FilterType.ShowOnlyActive, true);
            }
            suppliersPage.Filter(SuppliersPage.FilterType.ShowAll, true);
            nbrAllSupplier = suppliersPage.CheckTotalNumber();
            Assert.AreEqual(nbrAllSupplier - nbrActiveSuppliers, nbrInactiveSupplier, String.Format(MessageErreur.FILTRE_ERRONE, "'Show only inactive'"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_Filter_DeliveryDays()
        {
            //Prepare           
            string random_supplier_name = "test-supplier-" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();

            // Create
            var SupplierCreateModalpage = suppliersPage.SupplierCreatePage();
            SupplierCreateModalpage.FillField_CreatNewSupplier(random_supplier_name, true, false);
            var supplierItem = SupplierCreateModalpage.Submit();

            // Création d'un item
            CreateNewItem(supplierItem, random_supplier_name);

            suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();

            suppliersPage.Filter(SuppliersPage.FilterType.Search, random_supplier_name);
            supplierItem = suppliersPage.SelectFirstItem();

            // On ajoute le mardi en deliveryDays
            supplierItem.AddDeliveryDays(false, true);
            suppliersPage = supplierItem.BackToList();

            suppliersPage.ResetFilter();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, random_supplier_name);
            suppliersPage.Filter(SuppliersPage.FilterType.DeliveryDays, "Tuesday");

            Assert.AreEqual(1, suppliersPage.CheckTotalNumber(), String.Format(MessageErreur.FILTRE_ERRONE, "DeliveryDays"));

            suppliersPage.Filter(SuppliersPage.FilterType.DeliveryDays, "Sunday");
            Assert.AreEqual(0, suppliersPage.CheckTotalNumber(), String.Format(MessageErreur.FILTRE_ERRONE, "DeliveryDays"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_Filter_Site()
        {
            //Prepare           
            string site = TestContext.Properties["Site"].ToString();
            string random_supplier_name = "test-supplier-" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            var supplierCreateModalpage = suppliersPage.SupplierCreatePage();
            supplierCreateModalpage.FillField_CreatNewSupplier(random_supplier_name, true, false);
            var supplierItem = supplierCreateModalpage.Submit();

            CreateNewItem(supplierItem, random_supplier_name);
            suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, random_supplier_name);
            suppliersPage.Filter(SuppliersPage.FilterType.Site, site);

            var firstSupplier = suppliersPage.GetFirstSupplierName();
            Assert.AreEqual(firstSupplier, random_supplier_name, "Le nom du fournisseur ne correspond pas. ");

            var nbrSupplier = suppliersPage.CheckTotalNumber();
            Assert.AreEqual(1, nbrSupplier, "Le nombre de fournisseurs trouvés est incorrect.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_ExportIndex()
        {
            string downloadPath = TestContext.Properties["DownloadsPath"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            homePage.ClearDownloads();
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadPath);
            foreach (var file in taskDirectory.GetFiles())
            {
                // purge
                if (file.Name.EndsWith(".xlsx"))
                {
                    file.Delete();
                }
            }

            //Etre sur l'index des Supplier
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();

            //1. Filtrer un peu la liste de Supplier
            suppliersPage.ResetFilter();
            suppliersPage.Filter(SuppliersPage.FilterType.ShowOnlyActive, true);
            suppliersPage.Filter(SuppliersPage.FilterType.Site, "ACE");
            int totalLignes = suppliersPage.CheckTotalNumber();

            //2.Survoler... et cliquer sur Export
            // On exporte les résultats sous la forme d'un fichier Excel (dont on récupère le nom)
            // Export du fichier au format Excel
            suppliersPage.Export(true);
            // On récupère les fichiers du répertoire de téléchargement
            taskDirectory = new DirectoryInfo(downloadPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();
            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = suppliersPage.GetExportExcelFile(taskFiles, true);
            //Le fichier Excel est généré
            Assert.IsNotNull(correctDownloadedFile, "fichier non généré cas 1");
            Assert.IsTrue(correctDownloadedFile.Exists, "fichier non généré cas 2");

            //Vérifier les données
            var resultNumber = OpenXmlExcel.GetExportResultNumber("Suppliers", correctDownloadedFile.FullName);
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
            List<string> valeurs = OpenXmlExcel.GetValuesInList("Supplier Name", "Suppliers", correctDownloadedFile.FullName);
            Assert.AreEqual(totalLignes, valeurs.Count);

            //suppliersPage.ClearDownloads();

            string supplier = valeurs.Where(r => r.StartsWith("B")).FirstOrDefault();
            int offset = valeurs.IndexOf(supplier);

            suppliersPage.Filter(SuppliersPage.FilterType.Search, supplier);
            SupplierItem items = suppliersPage.SelectFirstItem();

            string SanitaryNumber = OpenXmlExcel.GetValuesInList("Sanitary Number", "Suppliers", correctDownloadedFile.FullName)[offset];
            string SIRET = OpenXmlExcel.GetValuesInList("SIRET", "Suppliers", correctDownloadedFile.FullName)[offset];
            string EdiID = OpenXmlExcel.GetValuesInList("Edi ID", "Suppliers", correctDownloadedFile.FullName)[offset];
            string EdifactField = OpenXmlExcel.GetValuesInList("Edifact Field", "Suppliers", correctDownloadedFile.FullName)[offset];
            string EdifactFormat = OpenXmlExcel.GetValuesInList("Edifact Format", "Suppliers", correctDownloadedFile.FullName)[offset];
            string PaymentTerm = OpenXmlExcel.GetValuesInList("Payment Term", "Suppliers", correctDownloadedFile.FullName)[offset];
            string EndOfContract = OpenXmlExcel.GetValuesInList("End Of Contract", "Suppliers", correctDownloadedFile.FullName)[offset];
            string Comment = OpenXmlExcel.GetValuesInList("Comment", "Suppliers", correctDownloadedFile.FullName)[offset];
            // Booléans
            string VatFree = OpenXmlExcel.GetValuesInList("Vat Free", "Suppliers", correctDownloadedFile.FullName)[offset];
            string Address1 = OpenXmlExcel.GetValuesInList("Address 1", "Suppliers", correctDownloadedFile.FullName)[offset];
            string Address2 = OpenXmlExcel.GetValuesInList("Address 2", "Suppliers", correctDownloadedFile.FullName)[offset];
            string ZipCode = OpenXmlExcel.GetValuesInList("Zip Code", "Suppliers", correctDownloadedFile.FullName)[offset];
            string City = OpenXmlExcel.GetValuesInList("City", "Suppliers", correctDownloadedFile.FullName)[offset];
            string Country = OpenXmlExcel.GetValuesInList("Country", "Suppliers", correctDownloadedFile.FullName)[offset];
            string Currency = OpenXmlExcel.GetValuesInList("Currency", "Suppliers", correctDownloadedFile.FullName)[offset];
            string LevelOfRisk = OpenXmlExcel.GetValuesInList("Level Of Risk", "Suppliers", correctDownloadedFile.FullName)[offset];
            string IntentedUse = OpenXmlExcel.GetValuesInList("Intented Use", "Suppliers", correctDownloadedFile.FullName)[offset];
            string Riskassessmentscore = OpenXmlExcel.GetValuesInList("Risk assessment score", "Suppliers", correctDownloadedFile.FullName)[offset];
            string QuantityVolume = OpenXmlExcel.GetValuesInList("Quantity Volume", "Suppliers", correctDownloadedFile.FullName)[offset];
            // Booléans
            string Tobeaudited = OpenXmlExcel.GetValuesInList("To be audited", "Suppliers", correctDownloadedFile.FullName)[offset];
            string LastAuditDate = OpenXmlExcel.GetValuesInList("Last Audit Date", "Suppliers", correctDownloadedFile.FullName)[offset];
            string NextAuditDate = OpenXmlExcel.GetValuesInList("Next Audit Date", "Suppliers", correctDownloadedFile.FullName)[offset];
            string Knownsupplier = OpenXmlExcel.GetValuesInList("Known supplier", "Suppliers", correctDownloadedFile.FullName)[offset];
            string SupplierType = OpenXmlExcel.GetValuesInList("Supplier Type", "Suppliers", correctDownloadedFile.FullName)[offset];

            Assert.AreEqual(SanitaryNumber, items.GetSanitaryNumber());
            Assert.AreEqual(SIRET, items.GetSIRET());
            Assert.AreEqual("", EdiID);
            Assert.AreEqual("NumVat", EdifactField);
            Assert.AreEqual("None", EdifactFormat);

            Assert.AreEqual(PaymentTerm, items.GetPaymentTerm());
            Assert.AreEqual(DateTime.FromOADate(double.Parse(EndOfContract)).ToString("dd/MM/yyyy"), items.GetEndOfContract());
            Assert.AreEqual(Comment, items.GetComment());
            Assert.AreEqual("FALSE", VatFree); // checkbox Supplier_IsExo
            Assert.AreEqual(Address1, items.GetAddress1());
            Assert.AreEqual(Address2, items.GetAddress2());
            Assert.AreEqual(ZipCode, items.GetZipCode());
            Assert.AreEqual(City, items.GetCity());
            Assert.AreEqual(Country, items.GetCountry());
            Assert.AreEqual(Currency, items.GetCurrency());
            Assert.AreEqual(LevelOfRisk, items.GetLevelOfRisk());
            Assert.AreEqual(IntentedUse, items.GetIntentedUse());
            // Low vs Low risk supplier
            Assert.AreEqual(Riskassessmentscore, items.GetRiskassessmentscore());
            Assert.AreEqual(QuantityVolume, items.GetQuantityVolume());
            Assert.AreEqual("FALSE", Tobeaudited); // checkbox check-isaudited
            Assert.AreEqual(LastAuditDate, items.GetLastAuditDate());
            Assert.AreEqual(NextAuditDate, items.GetNextAuditDate());
            Assert.AreEqual("", Knownsupplier);
            Assert.AreEqual("", SupplierType);

            //Deliveries
            List<string> valeursDeliveries = OpenXmlExcel.GetValuesInList("Supplier Name", "Deliveries", correctDownloadedFile.FullName);
            int offsetDeliveries = valeursDeliveries.IndexOf(supplier);
            string SiteName = OpenXmlExcel.GetValuesInList("Site Name", "Deliveries", correctDownloadedFile.FullName)[offsetDeliveries];
            string SiteCode = OpenXmlExcel.GetValuesInList("Site Code", "Deliveries", correctDownloadedFile.FullName)[offsetDeliveries];
            string SiteBranchCode = OpenXmlExcel.GetValuesInList("Site Branch Code", "Deliveries", correctDownloadedFile.FullName)[offsetDeliveries];
            string MinAmount = OpenXmlExcel.GetValuesInList("Min Amount", "Deliveries", correctDownloadedFile.FullName)[offsetDeliveries];
            string ShippingCost = OpenXmlExcel.GetValuesInList("Shipping Cost", "Deliveries", correctDownloadedFile.FullName)[offsetDeliveries];
            string Comment2 = OpenXmlExcel.GetValuesInList("Comment", "Deliveries", correctDownloadedFile.FullName)[offsetDeliveries];
            string Monday = OpenXmlExcel.GetValuesInList("Monday", "Deliveries", correctDownloadedFile.FullName)[offsetDeliveries];
            string Tuesday = OpenXmlExcel.GetValuesInList("Tuesday", "Deliveries", correctDownloadedFile.FullName)[offsetDeliveries];
            string Wednesday = OpenXmlExcel.GetValuesInList("Wednesday", "Deliveries", correctDownloadedFile.FullName)[offsetDeliveries];
            string Thursday = OpenXmlExcel.GetValuesInList("Thursday", "Deliveries", correctDownloadedFile.FullName)[offsetDeliveries];
            string Friday = OpenXmlExcel.GetValuesInList("Friday", "Deliveries", correctDownloadedFile.FullName)[offsetDeliveries];
            string Saturday = OpenXmlExcel.GetValuesInList("Saturday", "Deliveries", correctDownloadedFile.FullName)[offsetDeliveries];
            string Sunday = OpenXmlExcel.GetValuesInList("Sunday", "Deliveries", correctDownloadedFile.FullName)[offsetDeliveries];

            SupplierDeliveriesTab deliveriesTab = items.ClickOnDeliveriesTab();

            Assert.AreEqual("ACE", SiteName);
            Assert.AreEqual("", SiteCode);
            Assert.AreEqual("", SiteBranchCode);
            Assert.AreEqual(MinAmount == "" ? "0" : MinAmount, deliveriesTab.GetMinAmount());
            Assert.AreEqual(ShippingCost == "" ? "0" : ShippingCost, deliveriesTab.GetShippingCost());
            Assert.AreEqual(Comment2, deliveriesTab.GetComment(SiteName));
            Assert.AreEqual(Monday, deliveriesTab.GetMonday());
            Assert.AreEqual(Tuesday, deliveriesTab.GetTuesday());
            Assert.AreEqual(Wednesday, deliveriesTab.GetWednesday());
            Assert.AreEqual(Thursday, deliveriesTab.GetThursday());
            Assert.AreEqual(Friday, deliveriesTab.GetFriday());
            Assert.AreEqual(Saturday, deliveriesTab.GetSaturday());
            Assert.AreEqual(Sunday, deliveriesTab.GetSunday());

            List<string> valeursContacts = OpenXmlExcel.GetValuesInList("Supplier Name", "Contacts", correctDownloadedFile.FullName);
            int offsetContacts = valeursContacts.IndexOf(supplier);
            string Sites = OpenXmlExcel.GetValuesInList("Sites", "Contacts", correctDownloadedFile.FullName)[offsetContacts];
            string Name = OpenXmlExcel.GetValuesInList("Name", "Contacts", correctDownloadedFile.FullName)[offsetContacts];
            string Phone = OpenXmlExcel.GetValuesInList("Phone", "Contacts", correctDownloadedFile.FullName)[offsetContacts];
            string Fax = OpenXmlExcel.GetValuesInList("Fax", "Contacts", correctDownloadedFile.FullName)[offsetContacts];
            string Mail = OpenXmlExcel.GetValuesInList("Mail", "Contacts", correctDownloadedFile.FullName)[offsetContacts];
            string Adress1 = OpenXmlExcel.GetValuesInList("Adress1", "Contacts", correctDownloadedFile.FullName)[offsetContacts];
            string Adress2 = OpenXmlExcel.GetValuesInList("Adress2", "Contacts", correctDownloadedFile.FullName)[offsetContacts];
            string ZipCode2 = OpenXmlExcel.GetValuesInList("Zip Code", "Contacts", correctDownloadedFile.FullName)[offsetContacts];
            string City2 = OpenXmlExcel.GetValuesInList("City", "Contacts", correctDownloadedFile.FullName)[offsetContacts];
            string ForPO = OpenXmlExcel.GetValuesInList("ForPO", "Contacts", correctDownloadedFile.FullName)[offsetContacts];
            string ForClaim = OpenXmlExcel.GetValuesInList("ForClaim", "Contacts", correctDownloadedFile.FullName)[offsetContacts];

            // Contacts
            SupplierContactTab contactsTab = deliveriesTab.ClickOnContactsTab();
            contactsTab.Deplier("ACE");

            Assert.AreEqual("ACE", Sites);
            Assert.AreEqual(Name, supplier);
            Assert.AreEqual(Phone, contactsTab.GetPhone());
            Assert.AreEqual(Fax, contactsTab.GetFax());
            Assert.AreEqual(Mail, contactsTab.GetMail());
            string[] tableau = contactsTab.GetAdress();
            if (tableau.Length == 6)
            {
                Assert.AreEqual(Adress1, tableau[0]);
                Assert.AreEqual("", Adress2);
                Assert.AreEqual(ZipCode2, tableau[1]);
                Assert.AreEqual(City2, tableau[2]);
            }
            else
            {
                Assert.AreEqual(Adress1, tableau[0]);
                Assert.AreEqual(Adress2, tableau[1]);
                Assert.AreEqual(ZipCode2, tableau[2]);
                Assert.AreEqual(City2, tableau[3]);
            }
            Assert.AreEqual(ForPO == "TRUE" ? "For Purchase Orders" : "", contactsTab.GetForPO());
            Assert.AreEqual(ForClaim == "TRUE" ? "For Claims" : "", contactsTab.GetForClaim());

            //Accountings
            List<string> valeursAccountings = OpenXmlExcel.GetValuesInList("Supplier Name", "Accountings", correctDownloadedFile.FullName);
            int offsetAccountings = valeursAccountings.IndexOf(supplier);
            if (offsetAccountings == -1)
            {
                Assert.Fail("Pas de Accountings pour le Purchasing Supplier " + supplier);
            }
            string Sites2 = OpenXmlExcel.GetValuesInList("Sites", "Accountings", correctDownloadedFile.FullName)[offsetAccountings];
            string ThirdPart = OpenXmlExcel.GetValuesInList("Third Part", "Accountings", correctDownloadedFile.FullName)[offsetAccountings];
            string AccountingId = OpenXmlExcel.GetValuesInList("Accounting Id", "Accountings", correctDownloadedFile.FullName)[offsetAccountings];
            string ThirdPartDti = OpenXmlExcel.GetValuesInList("Third Part Dti", "Accountings", correctDownloadedFile.FullName)[offsetAccountings];
            string AccountingIdDti = OpenXmlExcel.GetValuesInList("Accounting Id Dti", "Accountings", correctDownloadedFile.FullName)[offsetAccountings];
            string SupplierAccountingId = OpenXmlExcel.GetValuesInList("SupplierAccountingId", "Accountings", correctDownloadedFile.FullName)[offsetAccountings];
            string TiersdetailsId = OpenXmlExcel.GetValuesInList("Tiers details Id", "Accountings", correctDownloadedFile.FullName)[offsetAccountings];
            string TiersName = OpenXmlExcel.GetValuesInList("Tiers Name", "Accountings", correctDownloadedFile.FullName)[offsetAccountings];
            string AddressIdentifier = OpenXmlExcel.GetValuesInList("Address Identifier", "Accountings", correctDownloadedFile.FullName)[offsetAccountings];
            string Domiciliationidentifier = OpenXmlExcel.GetValuesInList("Domiciliation identifier", "Accountings", correctDownloadedFile.FullName)[offsetAccountings];

            SupplierAccountingsTab accountingsTab = contactsTab.ClickOnAccountingsTab();

            Assert.AreEqual(Sites2, accountingsTab.GetSite());
            Assert.AreEqual(ThirdPart, accountingsTab.GetThirdPart());
            Assert.AreEqual(AccountingId, accountingsTab.GetAccountingId());
            Assert.AreEqual(ThirdPartDti, accountingsTab.GetThirdPartDti());
            Assert.AreEqual(AccountingIdDti, accountingsTab.GetAccountingIdDti());
            Assert.AreEqual(SupplierAccountingId, accountingsTab.GetSupplierAccountingId());
            Assert.AreEqual("", TiersdetailsId);
            Assert.AreEqual("", TiersName);
            Assert.AreEqual(AddressIdentifier, accountingsTab.GetAddressIdentifier());
            Assert.AreEqual(Domiciliationidentifier, accountingsTab.GetDomiciliationidentifier());

            // Certifications

            List<string> valeursCertifications = OpenXmlExcel.GetValuesInList("Supplier Name", "Certifications", correctDownloadedFile.FullName);

            int lengthCertifications = valeursCertifications.FindAll(x => x.Equals(supplier)).Count;
            int offsetCertifications = valeursCertifications.IndexOf(supplier);

            SupplierCertificationTab certificationsTab = accountingsTab.ClickOnCertificationsTab();

            for (int i = 0; i < lengthCertifications; i++)
            {
                offset = offsetCertifications + i;
                string Name2 = OpenXmlExcel.GetValuesInList("Name", "Certifications", correctDownloadedFile.FullName)[offset];
                string IsCopyMandatory = OpenXmlExcel.GetValuesInList("Is Copy Mandatory", "Certifications", correctDownloadedFile.FullName)[offset];
                string CopyGiven = OpenXmlExcel.GetValuesInList("Copy Given", "Certifications", correctDownloadedFile.FullName)[offset];
                string ExpireDateCheck = OpenXmlExcel.GetValuesInList("Expire Date Check", "Certifications", correctDownloadedFile.FullName)[offset];
                string ExpirationDate = OpenXmlExcel.GetValuesInList("Expiration Date", "Certifications", correctDownloadedFile.FullName)[offset];

                // ISO 14001 / ISO 9001 (tri plus ou moins intélligent)
                string Name2Tableau = certificationsTab.GetName(i);
                if (Name2.Contains(" "))
                {
                    Name2 = Name2.Substring(0, Name2.IndexOf(" ") + 1);
                    Name2Tableau = Name2Tableau.Substring(0, Name2Tableau.IndexOf(" ") + 1);
                }
                Assert.AreEqual(Name2, Name2Tableau, "Certification name");
                Assert.AreEqual(IsCopyMandatory == "TRUE" ? "Yes" : "No", certificationsTab.GetIsCopyMandatory(i), "Certification IsCopyMandatory");
                Assert.AreEqual(CopyGiven == "TRUE" ? "true" : "false", certificationsTab.GetCopyGiven(i), "Certification CopyGiven");
                Assert.AreEqual(ExpireDateCheck == "TRUE" ? "Yes" : "No", certificationsTab.GetExpireDateCheck(i), "Certification ExpireDateCheck");
                Assert.AreEqual(ExpirationDate == "" ? "" : DateTime.FromOADate(double.Parse(ExpirationDate)).ToString("dd/MM/yyyy"), certificationsTab.GetExpirationDate(i), "Certification ExpirationDate");
            }

            // AccountingsTiersDetails
            List<string> valeursAccountingsTiers = OpenXmlExcel.GetValuesInList("Supplier Name", "AccountingsTiersDetails", correctDownloadedFile.FullName);

            int lengthAccountingsTiers = valeursAccountingsTiers.FindAll(x => x.Equals(supplier)).Count;
            int offsetAccountingsTiers = valeursAccountingsTiers.IndexOf(supplier);

            SupplierAccountingsTiersDetailsTab accountingsTiersDetailsTab = certificationsTab.ClickOnAccountingsTiersDetailsTab();

            for (int i = 0; i < lengthAccountingsTiers; i++)
            {
                offset = offsetAccountingsTiers + i;
                string TiersDetailId = OpenXmlExcel.GetValuesInList("Tiers Detail Id", "AccountingsTiersDetails", correctDownloadedFile.FullName)[offset];
                string TiersCode = OpenXmlExcel.GetValuesInList("Tiers Code", "AccountingsTiersDetails", correctDownloadedFile.FullName)[offset];
                string TiersName2 = OpenXmlExcel.GetValuesInList("Tiers Name", "AccountingsTiersDetails", correctDownloadedFile.FullName)[offset];
                string TiersAddressIdentifier = OpenXmlExcel.GetValuesInList("Tiers Address Identifier", "AccountingsTiersDetails", correctDownloadedFile.FullName)[offset];
                string TiersAddress = OpenXmlExcel.GetValuesInList("Tiers Address", "AccountingsTiersDetails", correctDownloadedFile.FullName)[offset];
                string TiersDomiciliationIdentifier = OpenXmlExcel.GetValuesInList("Tiers Domiciliation Identifier", "AccountingsTiersDetails", correctDownloadedFile.FullName)[offset];


                Assert.AreEqual(TiersDetailId, accountingsTiersDetailsTab.GetTiersDetailId(i));
                Assert.AreEqual(TiersCode, accountingsTiersDetailsTab.GetTiersCode(i));
                Assert.AreEqual(TiersName2, accountingsTiersDetailsTab.GetTiersName(i));
                Assert.AreEqual(TiersAddressIdentifier, accountingsTiersDetailsTab.GetTiersAddressIdentifier(i));
                Assert.AreEqual(TiersAddress, accountingsTiersDetailsTab.GetTiersAddress(i));
                Assert.AreEqual(TiersDomiciliationIdentifier, accountingsTiersDetailsTab.GetTiersDomiciliationIdentifier(i));


            }
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_ImportIndex()
        {
            //Etre sur l'index des Supplier et avoir un fichier exporter
            string downloadPath = TestContext.Properties["DownloadsPath"].ToString();
            string random_supplier_name = "test-supplier-" + new Random().Next().ToString();
            string newSupplierName = "test-import-supplier-modifyName" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            homePage.ClearDownloads();
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadPath);
            foreach (var file in taskDirectory.GetFiles())
            {
                // purge
                if (file.Name.EndsWith(".xlsx"))
                {
                    file.Delete();
                }
            }

            //Act
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            var supplierCreateModalpage = suppliersPage.SupplierCreatePage();
            supplierCreateModalpage.FillField_CreatNewSupplier(random_supplier_name, true, false);
            var supplierItem = supplierCreateModalpage.Submit();

            CreateNewItem(supplierItem, random_supplier_name);

            //1. Filtrer un peu la liste de Supplier
            suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, random_supplier_name);
            Assert.AreEqual(1, suppliersPage.CheckTotalNumber());

            //2.Survoler... et cliquer sur Export
            // On exporte les résultats sous la forme d'un fichier Excel (dont on récupère le nom)
            // Export du fichier au format Excel
            suppliersPage.Export(true);
            // On récupère les fichiers du répertoire de téléchargement
            taskDirectory = new DirectoryInfo(downloadPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();
            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            FileInfo correctDownloadedFile = suppliersPage.GetExportExcelFile(taskFiles, true);
            //Le fichier Excel est généré
            Assert.IsNotNull(correctDownloadedFile, "fichier non généré cas 1");
            Assert.IsTrue(correctDownloadedFile.Exists, "fichier non généré cas 2");

            try
            {
                suppliersPage.ResetFilter(); // on ferme la fenêtre du service d'impression

                OpenXmlExcel.WriteDataInColumn("Action", "Suppliers", correctDownloadedFile.FullName, "M", CellValues.SharedString, true);
                OpenXmlExcel.WriteDataInColumn("Supplier Name", "Suppliers", correctDownloadedFile.FullName, newSupplierName, CellValues.SharedString, true);
                OpenXmlExcel.WriteDataInColumn("SIRET", "Suppliers", correctDownloadedFile.FullName, "ZZ999", CellValues.SharedString, true);

                //1. Survoler ... et cliquer sur importer
                Assert.IsTrue(suppliersPage.Import(correctDownloadedFile), "une erreur popup lors de l'import");
                //Le fichier est importer
                suppliersPage.ResetFilter();
                suppliersPage.Filter(SuppliersPage.FilterType.Search, newSupplierName);
                //Vérifier l'importation
                Assert.AreEqual(1, suppliersPage.CheckTotalNumber(), "La modification du nom d'un supplier existant lors de l'import n'as PageSetup fonctionné");
            }
            finally
            {
                suppliersPage.ResetFilter();
                suppliersPage.Filter(SuppliersPage.FilterType.Search, newSupplierName);
                if (suppliersPage.CheckTotalNumber() > 0)
                {
                    suppliersPage.SelectFirstItem();
                    SupplierGeneralInfoTab generalInfo = supplierItem.ClickOnGeneralInformationTab();
                    generalInfo.SetName(random_supplier_name);
                }
            }

        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_Detail_Filter_SortByItemGroup()
        {
            //Prepare
            string random_supplier_name = "test-supplier-" + new Random().Next().ToString();

            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string itemName = "supplier-" + new Random().Next().ToString();
            string storageUnit = "KG";
            string supplierRef = itemName + "_SupplierRef";
            string storageQty = 2.32.ToString();

            //Arrange
            var homePage = LogInAsAdmin();
            //Etre sur un supplier dans l'onglet items
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            // Create
            var SupplierCreateModalpage = suppliersPage.SupplierCreatePage();
            SupplierCreateModalpage.FillField_CreatNewSupplier(random_supplier_name, true, false);
            var supplierItem = SupplierCreateModalpage.Submit();

            // 1. New Item || Site By Default: MAD
            supplierItem.GetItemsTab();
            // Item and packaging creation
            var itemCreatePage = supplierItem.ItemCreatePage();
            var itemGeneralInformationPage = itemCreatePage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemCreatePackagingPage.FillField_CreateNewPackaging(
                site, packagingName, 2.32.ToString(), storageUnit, 4.2.ToString(), random_supplier_name, null, supplierRef, storageQty);

            //CreateNewItem(supplierItem, random_supplier_name);
            suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, random_supplier_name);
            suppliersPage.SelectFirstItem();
            supplierItem.GetItemsTab();

            //filter by Name
            supplierItem.Filter(SupplierItem.FilterType.SortBy, "Item group");
            var itemGroupAfterSort = supplierItem.GetItemGroup();
            Assert.AreEqual(group, itemGroupAfterSort, "erreur de filtrage By Item group");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_Detail_Filter_SortByName()
        {
            //Prepare
            string random_supplier_name = "test-supplier-" + new Random().Next().ToString();

            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string itemName = "supplier-" + new Random().Next().ToString();
            string storageUnit = "KG";
            string supplierRef = itemName + "_SupplierRef";
            string storageQty = 2.32.ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Etre sur un supplier dans l'onglet items
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();

            // Create
            var SupplierCreateModalpage = suppliersPage.SupplierCreatePage();
            SupplierCreateModalpage.FillField_CreatNewSupplier(random_supplier_name, true, false);
            var supplierItem = SupplierCreateModalpage.Submit();
            // 1. New Item || Site By Default: MAD

            supplierItem.GetItemsTab();
            // Item and packaging creation
            var itemCreatePage = supplierItem.ItemCreatePage();
            var itemGeneralInformationPage = itemCreatePage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemCreatePackagingPage.FillField_CreateNewPackaging(
                site, packagingName, 2.32.ToString(), storageUnit, 4.2.ToString(), random_supplier_name, null, supplierRef, storageQty);

            //CreateNewItem(supplierItem, random_supplier_name);
            suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, random_supplier_name);
            suppliersPage.SelectFirstItem();
            supplierItem.GetItemsTab();

            //filter by Name
            supplierItem.Filter(SupplierItem.FilterType.SortBy, "Name");
            var itemNameAfterSort = supplierItem.GetName();
            Assert.AreEqual(itemName.ToString(), itemNameAfterSort, "erreur de filtrage By Name");


        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_Detail_FilterShowOnlyActive()
        {
            string site = TestContext.Properties["SiteLpCart"].ToString();
            string itemName = "itemToFilter" + new Random().Next().ToString();
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();
            string limitQty = 177.ToString();
            string unitPrice = 100.ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            //Arrange

            HomePage homePage = LogInAsAdmin();

            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            try
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGenerlInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGenerlInformationPage.NewPackaging();
                itemGenerlInformationPage = itemCreatePackagingPage.FillField_CreateNewPackagingWithoutPrice(site, packagingName, storageQty, storageUnit, qty, supplier);
                itemPage = itemGenerlInformationPage.BackToList();
                var supplierspage = homePage.GoToPurchasing_SuppliersPage();
                supplierspage.Filter(SuppliersPage.FilterType.Search, supplier);
                SupplierItem supplieritem = supplierspage.SelectFirstItem();
                supplieritem.ClickOnItemsTab();
                supplieritem.ResetFilter();
                supplieritem.Filter(SupplierItem.FilterType.Search, itemName);
                supplieritem.Filter(SupplierItem.FilterType.ShowOnlyActive, true);
                var firstItemname = supplieritem.getItemName();
                Assert.IsTrue(supplieritem.VerifyFilterSearch(itemName), "Erreur de filtrage lors de la recherche par nom");
                Assert.AreEqual(firstItemname, itemName, "Erreur de filtrage lors de la recherche par show active only");
            }
            finally
            {

                itemPage = homePage.GoToPurchasing_ItemPage();
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
        public void PU_SUPP_Detail_FilterShowOnlyInactive()
        {
            //Arrange
            var homePage = LogInAsAdmin();

            //Etre sur un supplier dans l'onglet items
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            SupplierItem items = suppliersPage.SelectFirstItem();
            //Appliquer le filtre Show Only Active
            items.ClickOnItemsTab();
            items.ResetFilter();
            items.PageSize("100");
            items.PageUp();
            items.Filter(SupplierItem.FilterType.ShowOnlyInactive, true);

            //Les données sont filtrées
            int tentative = 0;
            while (items.IsFoldAll() && tentative < 10)
            {
                items.UnFoldorFoldItems();
                tentative++;
            }
            var inactives = WebDriver.FindElements(By.XPath("//*/table/tbody/tr[2]/td[6]/span[text()='Inactive']"));
            var actives = WebDriver.FindElements(By.XPath("//*/table/tbody/tr[2]/td[6]/a"));
            Assert.AreEqual(0, actives.Count, "Il y a des actives");
            Assert.IsTrue(inactives.Count > 0, "Il n'y a pas d'inactives");

        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_Detail_AddCertifications()
        {
            //prepare
            string certification = TestContext.Properties["Certification"].ToString();
            var homePage = LogInAsAdmin();
            //Etre sur un supplier dans l'onglet items
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            SupplierItem item = suppliersPage.SelectFirstItem();
            item.GetCertificationTab();
            var isadded = item.AddNewCertification(certification);
            Assert.IsTrue(isadded, "La certification n'a pas été ajoutée.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_Detail_DeleteCertifications()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            //Etre sur un supplier dans l'onglet items
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            SupplierItem item = suppliersPage.SelectFirstItem();
            item.GetCertificationTab();
            var isDeleted = item.DelFirstCertification();
            Assert.IsTrue(isDeleted, "La certification n'a pas été supprimée.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_Detail_EDI()
        {
            var idEdi = "idEdi";
            var repository = "repository";
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            //Etre sur un supplier dans l'onglet items
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            var supplierItem = suppliersPage.SelectFirstItem();
            supplierItem.ClickOnEDITab();
            supplierItem.FillEdiData(idEdi, repository);
            supplierItem.ClickOnDeliveriesTab();
            supplierItem.ClickOnEDITab();
            var id = supplierItem.GetIdEdi();
            var repo = supplierItem.GetRepository();
            Assert.AreEqual(idEdi, id, "id n'a pas été mis a jour");
            Assert.AreEqual(repository, repo, "repository n'a pas été mis a jour");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_DeliveredSites()
        {
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["SiteBis"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string itemName = "supplier-" + DateUtils.Now.ToString("dd / MM / yyyy") + " - " + new Random().Next().ToString();
            string storageUnit = "KG";
            string supplierRef = itemName + "_SupplierRef";
            string storageQty = "2.32";
            string random_supplier_name = "test-supplier-" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            // Create New Supplier
            var supplierCreateModalpage = suppliersPage.SupplierCreatePage();
            supplierCreateModalpage.FillField_CreatNewSupplier(random_supplier_name, true, false);
            var supplierItem = supplierCreateModalpage.Submit();
            supplierItem.PageUp();
            supplierItem.GetItemsTab();
            // Create New Item and Packaging
            var itemCreatePage = supplierItem.ItemCreatePage();
            var itemGeneralInformationPage = itemCreatePage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemCreatePackagingPage.FillField_CreateNewPackaging(
                site, packagingName, "2.32", storageUnit, "4.2", random_supplier_name, null, supplierRef, storageQty);

            homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();

            // Filter supplier name 
            suppliersPage.Filter(SuppliersPage.FilterType.Search, random_supplier_name);
            suppliersPage.SelectFirstItem();
            var firstSupplierName = supplierItem.getSupplierName();
            var supplierDeliveriesTab = supplierItem.ClickOnDeliveriesTab();

            supplierDeliveriesTab.SelectedDeliveredSites(site, false);

            suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            //Cliquer sur Delivered sites
            suppliersPage.ClickDeliverySites();
            suppliersPage.SetDeliveredSiteOnSupplier(site, firstSupplierName);

            // verify Site in Delivery Tab
            suppliersPage.Filter(SuppliersPage.FilterType.Search, random_supplier_name);
            suppliersPage.SelectFirstItem();
            supplierItem.ClickOnDeliveriesTab();

            // Assert
            bool isSelectedDeliveredSite = supplierDeliveriesTab.CheckSelectedDeliveredSites(site);
            Assert.IsTrue(isSelectedDeliveredSite, "La case 'Delivered' pour le site sélectionné n'est pas cochée");

        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_Detail_FilterActiveNotPurchasable()
        {
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string itemName = "supplier-" + new Random().Next().ToString();
            string storageUnit = "KG";
            string supplierRef = itemName + "_SupplierRef";
            string storageQty = 2.32.ToString();
            string random_supplier_name = "test-supplier-" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            var supplierCreateModalpage = suppliersPage.SupplierCreatePage();
            supplierCreateModalpage.FillField_CreatNewSupplier(random_supplier_name, true, false);
            var supplierItem = supplierCreateModalpage.Submit();
            supplierItem.PageUp();
            supplierItem.GetItemsTab();
            var itemCreatePage = supplierItem.ItemCreatePage();
            var itemGeneralInformationPage = itemCreatePage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemCreatePackagingPage.FillField_CreateNewPackaging(
                site, packagingName, 2.32.ToString(), storageUnit, 4.2.ToString(), random_supplier_name, null, supplierRef, storageQty);

            itemGeneralInformationPage.SetUnpurchasableItem(site, random_supplier_name, packagingName);
            WebDriver.Navigate().Refresh();

            homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();

            suppliersPage.Filter(SuppliersPage.FilterType.Search, random_supplier_name);
            suppliersPage.SelectFirstItem();
            supplierItem.ClickOnItemsTab();

            supplierItem.Filter(SupplierItem.FilterType.ActiveNotPurchasable, true);

            supplierItem.SelectFirstItem();
            var verifySupplieritems = supplierItem.AtleastOnePackagingIsNotPurchasable();
            Assert.IsTrue(verifySupplieritems, "erreur de filtrage active not purchasable");

        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_Detail_FilterByGroup()
        {
            //Prepare
            string random_supplier_name = "test-supplier-" + new Random().Next().ToString();
            //login as admin
            var homePage = LogInAsAdmin();
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();

            // Create
            var SupplierCreateModalpage = suppliersPage.SupplierCreatePage();
            SupplierCreateModalpage.FillField_CreatNewSupplier(random_supplier_name, true, false);
            var supplierItem = SupplierCreateModalpage.Submit();
            // 1. New Item || Site By Default: MAD
            CreateNewItem(supplierItem, random_supplier_name);
            var groupSelectedItem = supplierItem.GetItemGroupNameSelected();
            suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, random_supplier_name);
            suppliersPage.SelectFirstItem();
            supplierItem.GetItemsTab();

            //filter by groupe
            supplierItem.Filter(SupplierItem.FilterType.ByGroup, groupSelectedItem);
            var firstitemGrpName = supplierItem.GetFirstItemGroupName();
            Assert.AreEqual(groupSelectedItem, firstitemGrpName, "erreur de filtrage By Groupe");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_Detail_FilterBySite()
        {
            // Prepare
            string siteACE = TestContext.Properties["SiteACE"].ToString();
            string siteMAD = TestContext.Properties["SiteMAD"].ToString();
            string random_supplier_name = "test-supplier-" + new Random().Next().ToString();
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string itemName = "supplier-" + new Random().Next().ToString();
            string storageUnit = "KG";
            string supplierRef = itemName + "_SupplierRef";
            string storageQty = "2.32";

            // Login as admin
            var homePage = LogInAsAdmin();
            // Go to suppliers page
            var supplierspage = homePage.GoToPurchasing_SuppliersPage();

            //Create or open a supplier
            SupplierItem supplierItem;

            try
            {
                // Create New Supplier
                var supplierCreateModalpage = supplierspage.SupplierCreatePage();
                supplierCreateModalpage.FillField_CreatNewSupplier(random_supplier_name, true, false);
                supplierItem = supplierCreateModalpage.Submit();

                // Go to Item Page
                supplierItem.PageUp();
                supplierItem.GetItemsTab();

                // Create Item and two packaging with different Site (ACE / MAD) 
                var itemCreatePage = supplierItem.ItemCreatePage();
                var itemGeneralInformationPage = itemCreatePage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemCreatePackagingPage.FillField_CreateNewPackaging(
                    siteACE, packagingName, "2.32", storageUnit, "4.2", random_supplier_name, null, supplierRef, storageQty);
                itemGeneralInformationPage.NewPackaging();
                itemCreatePackagingPage.FillField_CreateNewPackaging(
                    siteMAD, packagingName, "2.32", storageUnit, "4.2", random_supplier_name, null, supplierRef, storageQty);

                // Go to Purchasing Suppliers Page
                supplierspage = homePage.GoToPurchasing_SuppliersPage();
                supplierspage.Filter(SuppliersPage.FilterType.Search, random_supplier_name);

                // Select first supplier
                var supplier = supplierspage.SelectFirstItem();
                // Go to items tab
                supplier.ClickOnItemsTab();
                supplier.ResetFilter();

                // Filter by site
                supplier.Filter(SupplierItem.FilterType.BySite, siteACE);

                // UnfoldAll and Get List Of SitesAfterFilter
                supplier.UnfoldAll();
                List<string> sitesAfterFilter = supplier.GetSiteItem();

                bool verifySiteAfterFilterExist = supplier.VerifySiteAfterFilterExist(siteACE, sitesAfterFilter);
                // Assert
                Assert.IsTrue(verifySiteAfterFilterExist, "erreur de filtrage By Site");
            }
            finally
            {
                // Go to Purchasing Item Page
                var itemPage = homePage.GoToPurchasing_ItemPage();
                itemPage.ResetFilter();

                //Desactivate packaging
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
                generalInfo.DeactivatePrice(siteACE, random_supplier_name, packagingName);
                generalInfo.DeactivatePrice(siteMAD, random_supplier_name, packagingName);

                //Delete Item
                generalInfo.BackToList();
                MassiveDeleteModal massiveDelete = itemPage.MenuMassiveDelete();
                massiveDelete.Filter(MassiveDeleteModal.FilterType.SearchByName, itemName);
                massiveDelete.Filter(MassiveDeleteModal.FilterType.ShowInactiveSuppliers, true);
                massiveDelete.Filter(MassiveDeleteModal.FilterType.SupplierMultiple, random_supplier_name);
                massiveDelete.ClickSearch();
                massiveDelete.ClickSelectAllButton();
                massiveDelete.Delete();
            }


        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_Detail_FilterBySubGroup()
        {
            //Prepare data
            string subGrpName = "subgrpname";
            string subGrpCode = "subgrpcode";
            string group = "A REFERENCIA";
            //login as admin
            var homePage = LogInAsAdmin();

            //activate subgrp option
            homePage.SetSubGroupFunctionValue(true);

            //create new subgroup
            ParametersProduction productionPage = homePage.GoToParameters_ProductionPage();
            productionPage.GoToTab_SubGroup();
            if (!productionPage.IsGroupPresent(subGrpName))
            {
                productionPage.AddNewSubGroup(subGrpName, subGrpCode);
            }
            //set item subgroup
            var itemPage = homePage.GoToPurchasing_ItemPage();
            var itemName = itemPage.GetFirstItemName();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            var itemGeneralInformationPage = itemPage.ClickOnFirstItem();
            itemGeneralInformationPage.SetGroupName(group);
            itemGeneralInformationPage.SetSubgroupName(subGrpName);
            var supplierName = itemGeneralInformationPage.GetFirstPackagingSupplier();

            var supplierspage = homePage.GoToPurchasing_SuppliersPage();
            //select supplier
            supplierspage.ResetFilter();
            supplierspage.Filter(SuppliersPage.FilterType.Search, supplierName);
            var supplier = supplierspage.SelectFirstItem();
            //go to items tab
            supplier.ClickOnItemsTab();
            //filter by subgroupe
            supplier.Filter(SupplierItem.FilterType.BySubGroupe, subGrpName);
            var firstitemSubGrp = supplier.GetFirstItemSubGrp();
            Assert.AreEqual(subGrpName, firstitemSubGrp, "erreur de filtrage By subgrp");

            //changing setting sub group to inactive
            homePage.SetSubGroupFunctionValue(false);
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_Detail_FilterSearch()
        {
            string site = TestContext.Properties["SiteLpCart"].ToString();
            string itemName = "itemToFilter" + new Random().Next().ToString();
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();

            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();
            string limitQty = 177.ToString();
            string unitPrice = 100.ToString();
            //login as admin
            HomePage homePage = LogInAsAdmin();
            var supplierspage = homePage.GoToPurchasing_SuppliersPage();
            supplierspage.Filter(SuppliersPage.FilterType.Search, supplier);
            SupplierItem supplieritem = supplierspage.SelectFirstItem();
            supplieritem.ClickOnItemsTab();
            supplieritem.ResetFilter();
            if (supplieritem.ItemsCount() == 0)
            {
                ItemCreateModalPage createItem = supplieritem.ItemCreatePage();
                ItemGeneralInformationPage itemGeneralInfo = createItem.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInfo.NewPackaging();
                itemGeneralInfo = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier, limitQty, null, unitPrice);
                supplieritem.BackToList();
                supplierspage = homePage.GoToPurchasing_SuppliersPage();
                supplierspage.Filter(SuppliersPage.FilterType.Search, supplier);
                supplieritem = supplierspage.SelectFirstItem();
                supplieritem.ClickOnItemsTab();
                supplieritem.ResetFilter();

            }
            var firstItemname = supplieritem.getItemName();
            supplieritem.Filter(SupplierItem.FilterType.Search, firstItemname);
            Assert.IsTrue(supplieritem.VerifyFilterSearch(firstItemname), "Erreur de filtrage lors de la recherche par nom");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_Detail_ModifyCertifications()
        {
            var expiryDate = "25/05/2022";
            //login as admin
            LogInAsAdmin();
            //go to suppliers page
            var homePage = new HomePage(WebDriver, TestContext);
            var supplierspage = homePage.GoToPurchasing_SuppliersPage();
            //select first supplier
            var supplier = supplierspage.SelectFirstItem();
            //go to certifications tab
            supplier.GetCertificationTab();
            supplier.ModifyExpirationDateCertification(expiryDate);
            supplier.ClickOnItemsTab();
            supplier.GetCertificationTab();
            var dateAfterModif = supplier.GetFirstExpirationDate();
            Assert.AreEqual(expiryDate, dateAfterModif, "erreur de mise a jour de certification");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_Supplier_AccountingTiersDetails()
        {
            //Prepare
            string accountingTiersDetailsSupplier = "Accounting Tiers Details Supplier";

            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            //Créer ou ouvrir un supplier
            SupplierItem supplierItem;
            suppliersPage.Filter(SuppliersPage.FilterType.Search, accountingTiersDetailsSupplier);
            suppliersPage.Filter(SuppliersPage.FilterType.ShowAll, true);
            if (suppliersPage.CheckTotalNumber() == 0)
            {
                var supplierCreateModalpage = suppliersPage.SupplierCreatePage();
                supplierCreateModalpage.FillField_CreatNewSupplier(accountingTiersDetailsSupplier, true, false);
                supplierItem = supplierCreateModalpage.Submit();
                CreateNewItem(supplierItem, accountingTiersDetailsSupplier);
                suppliersPage = homePage.GoToPurchasing_SuppliersPage();
                suppliersPage.Filter(SuppliersPage.FilterType.Search, accountingTiersDetailsSupplier);
            }

            supplierItem = suppliersPage.SelectFirstItem();
            //1. Je cliquer sur "New accounting entry" et crée un nouveau supplier accounting
            supplierItem.ClickOnAccountingTiersDetailsTab();
            supplierItem.DeleteAccountingTiersDetails("test");

            int avantAvant = supplierItem.AccountingCount();
            supplierItem.CreateNewAccountingTiersDetails("test", "test", "test", "test", "test");
            //2.Je vérifier les infos crée

            int avant = supplierItem.AccountingCount();
            Assert.AreEqual(avantAvant + 1, avant);
            //3.Edit le new et vérifier les infos éditées
            /*
             * testbis est un mot trop long pour le champs en BDD
             * {Alert text : An unexpected error occurred in Winrest.
             * Message : String or binary data would be truncated in table &#x27;winrest-testauto6-patch-db.dbo.AccountingTiersDetails&#x27;, column &#x27;TiersDomiciliationIdentifier&#x27;. Truncated value: &#x27;testbi&#x27;.&#xD;&#xA;The statement has been terminated.
             */
            supplierItem.EditAccountingTiersDetails("test", "test2", "test2", "test2", "test2", "test2");

            int apres = supplierItem.AccountingCount();
            Assert.AreEqual(avant, apres, "Ce n'est pas une edition");

            //4.Supprimer et vérifier la suppresion
            supplierItem.DeleteAccountingTiersDetails("test2");
            apres = supplierItem.AccountingCount();
            Assert.AreEqual(avantAvant, apres, "Ce n'est pas une suppression");
            //Il faut vérifier que la création/l'édition/la suppresion des données comptables fonctionne
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_ResetFilter()
        {
            //Prepare

            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            //Etre sur l'index des suppliers
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            int avant = suppliersPage.CheckTotalNumber();
            //1. Appliquer quelques filtres
            suppliersPage.Filter(SuppliersPage.FilterType.Site, "MAD");
            int suivant1 = suppliersPage.CheckTotalNumber();
            Assert.AreNotEqual(avant, suivant1);
            suppliersPage.Filter(SuppliersPage.FilterType.ShowOnlyActive, true);
            int suivant2 = suppliersPage.CheckTotalNumber();
            Assert.IsTrue(suivant1 <= suivant2, "Filtre ShowOnlyActive KO");
            suppliersPage.Filter(SuppliersPage.FilterType.DeliveryDays, "Thursday");
            int suivant3 = suppliersPage.CheckTotalNumber();
            Assert.AreNotEqual(suivant2, suivant3);
            //2.Cliquer sur Reset filter
            suppliersPage.ResetFilter();
            int apres = suppliersPage.CheckTotalNumber();
            //Les filtres sont réinitialisé
            Assert.AreEqual(avant, apres);
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_PrintResults()
        {
            //Prepare
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string docFileNamePdfBegin = "Supplier List_-_";
            string docFileNameZipBegin = "All_files_";
            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            homePage.ClearDownloads();
            //Etre sur l'index des suppliers
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            //1.Survoler...
            //2.Ciquer sur Print Results
            //3.Ouvrir le PDF à partir de la print list
            //Le fichier PDF est généré dans un autre onglet
            PrintReportPage reportPage = suppliersPage.Print(true);
            reportPage.Purge(downloadsPath, docFileNamePdfBegin, docFileNameZipBegin);
            string trouvePrintAllZip = reportPage.PrintAllZip(downloadsPath, docFileNamePdfBegin, docFileNameZipBegin);
            FileInfo trouve = new FileInfo(trouvePrintAllZip);

            Assert.IsTrue(trouve.Exists, "Fichier PDF non trouvé");

            PdfDocument document = PdfDocument.Open(trouvePrintAllZip);
            Page page1 = document.GetPage(1);
            IEnumerable<Word> words = page1.GetWords();
            Assert.AreEqual(2, words.Where(w => w.Text.Contains("Suppliers")).Count(), "'Suppliers' non présent dans le Pdf");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_PrintNberOfItems()
        {
            //Prepare
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string fileName = "export-suppliers -active.xlsx";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            homePage.ClearDownloads();
            new FileInfo(Path.Combine(downloadsPath, fileName)).Delete();

            //Etre sur l'index des suppliers
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();

            //1. Survoler ...
            //2. Cliquer sur Nber Of Items
            suppliersPage.NberOfitems();

            //3. Ouvrir le fichier Excel dans la print list
            //Le fichier Excel est généré
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();
            var correctDownloadedFile = suppliersPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);
            int resultNumber = OpenXmlExcel.GetExportResultNumber("Suppliers Active Items", correctDownloadedFile.FullName);
            Assert.AreNotEqual(0, resultNumber, "Le fichier excel est vide");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_DetailItemResetFilter()
        {
            //login as admin
            var homePage = LogInAsAdmin();
            var supplierspage = homePage.GoToPurchasing_SuppliersPage();
            //select first supplier
            var supplier = supplierspage.SelectFirstItem();
            //go to items tab
            supplier.ClickOnItemsTab();
            var itemName = supplier.getItemName();
            supplier.Filter(SupplierItem.FilterType.Search, itemName);
            supplier.ResetFilter();
            bool verifResetFilter = supplier.VerifyResetFilter();
            Assert.IsTrue(verifResetFilter, "erreur de réinitialisation du filtre");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_FastCopyPrice()
        {
            //login as admin
            LogInAsAdmin();
            //go to suppliers page
            var homePage = new HomePage(WebDriver, TestContext);
            var supplierspage = homePage.GoToPurchasing_SuppliersPage();
            supplierspage.CopyPriceUpdate();
            Assert.IsTrue(supplierspage.VerifyFastCopyPrice(), "echec de fast copy price");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_DetailItems()
        {
            //login as admin
            LogInAsAdmin();
            //go to suppliers page
            var homePage = new HomePage(WebDriver, TestContext);
            var supplierspage = homePage.GoToPurchasing_SuppliersPage();
            supplierspage.ResetFilter();
            //Etre dans le détail d'un suppliers
            string supplierName = supplierspage.GetFirstSupplierName();
            SupplierItem items = supplierspage.SelectFirstItem();
            //Cliquer sur l'onglet items
            ItemGeneralInformationPage generalInfo = items.ClickOnGeneralInformationTab();
            items = generalInfo.ClickOnItemsTab();
            //l'onglet s'affiche
            Assert.AreEqual(items.getSupplierName(), supplierName);
            Assert.IsTrue(items.IsItemsPage(), "Pas sur l'onglet Items");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_Detail_DeleteItem()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();

            string itemName = "CARAMBAR";
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();

            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();

            //login as admin
            LogInAsAdmin();
            //go to suppliers page
            var homePage = new HomePage(WebDriver, TestContext);
            var supplierspage = homePage.GoToPurchasing_SuppliersPage();
            supplierspage.ResetFilter();
            supplierspage.Filter(SuppliersPage.FilterType.Site, site);
            supplierspage.Filter(SuppliersPage.FilterType.Search, supplier);
            //Etre sur un supplier dans l'onglet items
            var generalInfo = supplierspage.SelectFirstItem();
            SupplierItem items = generalInfo.ClickOnItemsTab();
            // Aller sur un item et le supprimer
            // créer un item/pack vers site afin de l'effacer ensuite
            items.Filter(SupplierItem.FilterType.Search, itemName);
            items.Filter(SupplierItem.FilterType.ShowOnlyActive, true);

            if (items.isElementVisible(By.XPath("//*/p[text()='No item']")))
            {
                ItemPage itemsPage = items.GoToPurchasing_ItemPage();
                itemsPage.ResetFilter();
                itemsPage.Filter(ItemPage.FilterType.Search, itemName);
                itemsPage.Filter(ItemPage.FilterType.ShowAll, true);
                ItemGeneralInformationPage itemGeneralInfo = null;
                if (itemsPage.CheckTotalNumber() == 0)
                {
                    ItemCreateModalPage createItem = itemsPage.ItemCreatePage();
                    itemGeneralInfo = createItem.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
                }
                else
                {
                    itemGeneralInfo = itemsPage.ClickOnFirstItem();
                }
                itemGeneralInfo.ShowAllPackaging();
                itemGeneralInfo.SearchPackaging("AIR CPU, S.L.");
                if (itemGeneralInfo.CountPackaging() > 0)
                {
                    // var menu = itemGeneralInfo.WaitForElementIsVisible(By.XPath("//*/a[contains(@class,'open-item-packaging')]"));
                    //menu.Click();
                    itemGeneralInfo.showfirstpackaging();
                }
                else
                {
                    var itemCreatePackagingPage = itemGeneralInfo.NewPackaging();
                    itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                }
            }
            supplierspage = items.GoToPurchasing_SuppliersPage();
            supplierspage.Filter(SuppliersPage.FilterType.Site, site);
            supplierspage.Filter(SuppliersPage.FilterType.Search, supplier);

            generalInfo = supplierspage.SelectFirstItem();
            items = generalInfo.ClickOnItemsTab();
            items.Filter(SupplierItem.FilterType.Search, itemName);

            //var unFold =items.WaitForElementIsVisible(By.Id("unfoldBtn"));
            //if (!unFold.GetAttribute("class").Contains("unfolded"))
            //{
            //    unFold.Click();
            //    Thread.Sleep(2000);
            //}
            items.UnfoldAllForDelete();
            items.UnfoldAllForDelete();

            items.DeleteButton();
            var confirmDelete = items.WaitForElementIsVisible(By.Id("dataConfirmOK"));
            confirmDelete.Click();
            ItemGeneralInformationPage generalInfoTab = items.ClickOnGeneralInformationTab();
            items = generalInfoTab.ClickOnItemsTab();
            items.Filter(SupplierItem.FilterType.Search, itemName);
            Assert.IsTrue(items.isElementVisible(By.XPath("//*/p[text()='No item']")));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_ExportIndexGeneralInformation()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string downloadPath = TestContext.Properties["DownloadsPath"].ToString();

            //login as admin
            LogInAsAdmin();
            //go to suppliers page
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.ClearDownloads();
            var supplierspage = homePage.GoToPurchasing_SuppliersPage();
            supplierspage.ResetFilter();
            supplierspage.Filter(SuppliersPage.FilterType.Site, site);
            supplierspage.Filter(SuppliersPage.FilterType.Search, supplier);
            //Etre sur un supplier dans l'onglet items
            SupplierGeneralInfoTab supplierGeneralInfo = supplierspage.SelectFirstSupplier();
            IEnumerable<string> knownSupplierSites = supplierGeneralInfo.GetSites();
            supplierGeneralInfo.BackToList();
            supplierspage.Filter(SuppliersPage.FilterType.Site, site);
            supplierspage.Filter(SuppliersPage.FilterType.Search, supplier);
            supplierspage.Export(true);
            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadPath);
            taskDirectory = new DirectoryInfo(downloadPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();
            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            FileInfo correctDownloadedFile = supplierspage.GetExportExcelFile(taskFiles, true);
            //Le fichier Excel est généré
            Assert.IsNotNull(correctDownloadedFile, "fichier non généré cas 1");
            Assert.IsTrue(correctDownloadedFile.Exists, "fichier non généré cas 2");
            //Vérifier les données
            var resultNumber = OpenXmlExcel.GetExportResultNumber("Suppliers", correctDownloadedFile.FullName);
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
            List<string> valeurs = OpenXmlExcel.GetValuesInList("Known supplier", "Suppliers", correctDownloadedFile.FullName);

            Assert.IsTrue(supplierspage.verifyDonneeExcel(valeurs, knownSupplierSites), MessageErreur.EXCEL_DONNEES_KO);
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_RESETFILTERINDEX()
        {
            //Prepare
            string siteMAD = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            string defaultName = suppliersPage.GetNameSearched();
            string defaultSortFilter = suppliersPage.GetSortTypeSelected();
            string defaultShowChecked = suppliersPage.GetActiveOrInactiveChecked();
            string defaultDeliveryDaysSelected = suppliersPage.GetDeliveryDaySelected();
            string defaultSitesSelected = suppliersPage.GetSiteSelected();
            string defaultSupplierTypesSelected = suppliersPage.GetSupplierTypesSelected();
            // Appliquer quelques filtres
            suppliersPage.Filter(SuppliersPage.FilterType.Search, supplier);
            suppliersPage.Filter(SuppliersPage.FilterType.SortBy, "Sort by internal ID");
            suppliersPage.Filter(SuppliersPage.FilterType.ShowAll, true);
            suppliersPage.Filter(SuppliersPage.FilterType.Site, siteMAD);
            suppliersPage.Filter(SuppliersPage.FilterType.DeliveryDays, "Thursday");

            //2.Cliquer sur Reset filter
            suppliersPage.ResetFilter();
            //Les filtres sont supprimés
            string sortAfterResetFilter = suppliersPage.GetSortTypeSelected();
            Assert.AreEqual(defaultName, suppliersPage.GetNameSearched(), "le filter Search By Name n'est pas supprimé");
            Assert.AreEqual(defaultSortFilter, sortAfterResetFilter, "le filter Sort By  n'est pas supprimé");
            Assert.AreEqual(defaultShowChecked, suppliersPage.GetActiveOrInactiveChecked(), "le filter Show n'est pas supprimé");
            Assert.AreEqual(defaultDeliveryDaysSelected, suppliersPage.GetDeliveryDaySelected(), "le filter par DELIVERY DAYS n'est pas supprimé");
            Assert.AreEqual(defaultSitesSelected, suppliersPage.GetSiteSelected(), "le filter par SITES n'est pas supprimé");
            Assert.AreEqual(defaultSupplierTypesSelected, suppliersPage.GetSupplierTypesSelected(), "le filter par SUPPLIER TYPES n'est pas supprimé");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_DETAIL_CREATECONTACT()
        {
            //Prepare
            string random_supplier_name = "test-supplier-" + new Random().Next().ToString();
            string contactName = "newRest@test.fr";
            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            var supplierCreateModalpage = suppliersPage.SupplierCreatePage();
            supplierCreateModalpage.FillField_CreatNewSupplier(random_supplier_name, true, false);
            var supplierItem = supplierCreateModalpage.Submit();
            string newContact = supplierItem.AddContact(contactName);
            var firstContact = supplierItem.GetFirstContact();
            Assert.AreEqual(newContact.Trim(), firstContact.Trim(), "Le nouveau contact n'est pas ajouté à la liste de contact du fournisseur");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_DETAIL_FASTCOPYPRICE()
        {
            //Prepare
            string random_supplier_name = "test-supplier-" + new Random().Next().ToString();
            string siteACE = TestContext.Properties["SiteACE"].ToString();
            string siteMAD = TestContext.Properties["SiteMAD"].ToString();
            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();

            // Create
            var SupplierCreateModalpage = suppliersPage.SupplierCreatePage();
            SupplierCreateModalpage.FillField_CreatNewSupplier(random_supplier_name, true, false);
            var supplierItem = SupplierCreateModalpage.Submit();
            // 1. New Item || Site By Default: MAD
            CreateNewItem(supplierItem, random_supplier_name);
            suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, random_supplier_name);
            suppliersPage.SelectFirstItem();
            supplierItem.GetItemsTab();
            supplierItem.ResetFilter();
            supplierItem.Filter(SupplierItem.FilterType.BySite, siteMAD);
            var itemNameToCopy = supplierItem.getItemName();
            supplierItem.ResetFilter();
            supplierItem.Filter(SupplierItem.FilterType.BySite, siteACE);
            supplierItem.FastCopyPrice("MAD", siteACE);
            supplierItem.ResetFilter();
            supplierItem.Filter(SupplierItem.FilterType.BySite, siteACE);
            var itemNameCopied = supplierItem.getItemName();
            Assert.AreEqual(itemNameToCopy, itemNameCopied, "Les prix du supplier selectionné n'ont pas été copié");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_DETAILSRORAGE_FILTERITEMGROUP()
        {
            string subGrpName = "SubGroupOF";
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string itemName = "supplier-" + new Random().Next().ToString();
            string storageUnit = "KG";
            string supplierRef = itemName + "_SupplierRef";
            string storageQty = 2.32.ToString();
            string random_supplier_name = "test-supplier-" + new Random().Next().ToString();
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();
            string weigh = null;

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //activate subgrp option
            homePage.SetSubGroupFunctionValue(true);

            try
            {
                //Etre sur un supplier dans l'onglet items
                //SupplierItem items = suppliersPage.SelectFirstItem();

                //items.ClickOnItemsTab();
                var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
                var supplierCreateModalpage = suppliersPage.SupplierCreatePage();
                supplierCreateModalpage.FillField_CreatNewSupplier(random_supplier_name, true, false);
                var supplierItem = supplierCreateModalpage.Submit();
                supplierItem.PageUp();
                supplierItem.GetItemsTab();
                var itemCreatePage = supplierItem.ItemCreatePage();
                var itemGeneralInformationPage = itemCreatePage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit, weigh, subGrpName);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemCreatePackagingPage.FillField_CreateNewPackaging(
                    site, packagingName, 2.32.ToString(), storageUnit, 4.2.ToString(), random_supplier_name, null, supplierRef, storageQty);
                homePage.GoToPurchasing_SuppliersPage();
                suppliersPage.Filter(SuppliersPage.FilterType.Search, random_supplier_name);
                var SupplierItem = suppliersPage.SelectFirstItem();
                SupplierItem.ClickOnItemsTab();
                supplierItem.ResetFilter();
                supplierItem.Filter(SupplierItem.FilterType.ByGroup, group);
                supplierItem.PageSize("100");
                supplierItem.PageUp();
                supplierItem.UnFoldorFoldItems();



                var groupNameList = supplierItem.GetListGroupName();

                foreach (var item in groupNameList)
                {
                    Assert.AreEqual(group, item, "Les groups ne sont pas le meme.");
                }

                //Appliquer le filtre by sub group

                supplierItem.ResetFilter();
                supplierItem.Filter(SupplierItem.FilterType.BySubGroupe, subGrpName);

                var subGroupName = supplierItem.GetFirstItemSubGrp();

                Assert.AreEqual(subGrpName, subGroupName, "Les subgroups ne sont pas le meme.");
            }
            finally
            {
                //changing setting sub group to inactive
                homePage.SetSubGroupFunctionValue(false);
            }
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_DETAIL_FOLDCONTACT()
        {
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            SupplierItem supplierItem = suppliersPage.SelectFirstItem();
            supplierItem.GoToContactTab();
            if (!supplierItem.RowContactsIsVisible())
            {
                // Create a new contact if contact details are not visible
                supplierItem.AddContact("test@mail.com");
            }
            supplierItem.ClickUnfoldAllButton();
            Assert.IsTrue(supplierItem.DetailsContactsIsVisible(), "Les détails des contacts devraient être affichés après 'Unfold All'");
            supplierItem.ClickFoldAllButton();
            Assert.IsFalse(supplierItem.DetailsContactsIsVisible(), "Les détails des contacts devraient être cachés après 'Fold All'");
        }
        // ________________________________________ Utilitaire _________________________________________________
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_INDEX_EXPORTNBEROFITEM()
        {
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            //login as admin
            var homePage = LogInAsAdmin();
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            suppliersPage.NberOfitems();
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();
            var correctDownloadedFile = suppliersPage.GetExportExcelFile(taskFiles, true);
            Assert.IsNotNull(correctDownloadedFile);
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_INDEX_AFFICHAGESUPPLIER50()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);

            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();

            suppliersPage.ResetFilter();
            suppliersPage.PageSize("50");
            bool pageSize100 = suppliersPage.IsPageItemsCountEqualsTo50();

            Assert.IsTrue(pageSize100);
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_INDEX_AFFICHAGESUPPLIER100()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);

            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();

            suppliersPage.ResetFilter();
            suppliersPage.PageSize("100");
            bool pageSize100 = suppliersPage.IsPageItemsCountEqualsTo100();

            Assert.IsTrue(pageSize100);
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_DETAIL_EXPORTITEM()
        {
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string itemName = $"supplier-{new Random().Next()}";
            string storageUnit = "KG";
            string supplierRef = $"{itemName}_SupplierRef";
            string storageQty = 2.32.ToString();
            string random_supplier_name = $"test-supplier-{new Random().Next()}";
            //Arrange
            HomePage homePage = LogInAsAdmin();
            homePage.ClearDownloads();
            homePage.PurgeDownloads();
            //Act
            SuppliersPage suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            SupplierCreateModalPage supplierCreateModalpage = suppliersPage.SupplierCreatePage();
            supplierCreateModalpage.FillField_CreatNewSupplier(random_supplier_name, true, false);
            SupplierItem supplierItem = supplierCreateModalpage.Submit();
            supplierItem.PageUp();
            supplierItem.GetItemsTab();
            ItemCreateModalPage itemCreatePage = supplierItem.ItemCreatePage();
            ItemGeneralInformationPage itemGeneralInformationPage = itemCreatePage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            ItemCreateNewPackagingModalPage itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemCreatePackagingPage.FillField_CreateNewPackaging(
                site, packagingName, 2.32.ToString(), storageUnit, 4.2.ToString(), random_supplier_name, null, supplierRef, storageQty);
            homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, random_supplier_name);
            suppliersPage.SelectFirstItem();
            supplierItem.ClickOnItemsTab();
            supplierItem.ExportItems();
            bool isExported = supplierItem.IsExported();
            Assert.IsTrue(isExported, "L'export n'est pas sauvegardée");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_DETAIL_NEWCERTIFICATION()
        {
            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            SupplierItem supplierItem = suppliersPage.SelectFirstItem();
            supplierItem.GetCertificationTab();
            supplierItem.GetListGroupName();
            supplierItem.DelCertification("IFS");
            var countCertifsBeforeAdd = supplierItem.certificationsCount();
            var isAddedCertificate = supplierItem.AddNewCertification("IFS");
            var countCertifsAfterAdd = supplierItem.certificationsCount();

            Assert.AreNotEqual(countCertifsBeforeAdd, countCertifsAfterAdd, "Le certification n'est pas sauvegardée");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_INDEX_SETUNPURCHASABLEITEM()
        {
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string suppliertTestDeleteItem = "SupplierTestDeleteItem";

            string itemName = TestContext.Properties["ItemNamePriceList"].ToString();
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            ClearCache();

            // Act
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();

            suppliersPage.Filter(SuppliersPage.FilterType.Search, "SUPPLIERTESTDELETEITEM");

            var supplierItem = suppliersPage.ClickAndGoFirstSupplier();
            supplierItem.ClickOnItemsTab();
            int numberItemsCountBefore = supplierItem.ItemsCount();

            var isFirstItemExists = supplierItem.ClickOnFirstItems();
            if (!isFirstItemExists)
            {
                CreateNewItem(supplierItem, suppliertTestDeleteItem);
                suppliersPage = homePage.GoToPurchasing_SuppliersPage();
                suppliersPage.Filter(SuppliersPage.FilterType.Search, suppliertTestDeleteItem);

                supplierItem = suppliersPage.ClickAndGoFirstSupplier();
                supplierItem.ClickOnItemsTab();

                ItemCreateModalPage modal = supplierItem.ItemCreatePage();
                string site = "MAD";
                string packagingName = "KG";
                string storageQty = "5";
                string storageUnit = "KG";
                string qty = "58";
                string supplier = "SupplierTestDeleteItem";
                ItemGeneralInformationPage generalInfo = modal.FillField_CreateNewItem(itemName + new Random().NextDouble(), group, workshop, taxType, prodUnit);

                var itemCreatePackagingPage = generalInfo.NewPackaging();
                generalInfo = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);

                suppliersPage = homePage.GoToPurchasing_SuppliersPage();
                suppliersPage.Filter(SuppliersPage.FilterType.Search, "SupplierTestDeleteItem");
                supplierItem = suppliersPage.ClickAndGoFirstSupplier();
                supplierItem.ClickOnItemsTab();
                numberItemsCountBefore = supplierItem.ItemsCount();
            }

            suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, "SUPPLIERTESTDELETEITEM");

            suppliersPage.ClickUNPURCHASABLEITEM();
            suppliersPage.SetUnpurchasableItem();
            supplierItem = suppliersPage.ClickAndGoFirstSupplier();
            supplierItem.ClickOnItemsTab();
            var numberItemsAfter = supplierItem.ItemsCount();
            Assert.AreEqual(0, numberItemsAfter, "L'item n'est pas unpurchasable");
            Assert.AreNotEqual(numberItemsCountBefore, numberItemsAfter, "L'item n'est pas unpurchasable.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_DETAIL_DESACTIVATEITEM()
        {
            var firstItemsCount = (int)decimal.Zero;
            string itemName = "supplier-" + new Random().Next().ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            var firstSupplierName = suppliersPage.GetFirstSupplierName();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, firstSupplierName);

            var supplierItem = suppliersPage.ClickAndGoFirstSupplier();
            supplierItem.ClickOnItemsTab();

            CreateNewItem(supplierItem, firstSupplierName);
            ClearCache();
            Thread.Sleep(2000);

            suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, firstSupplierName);

            supplierItem = suppliersPage.ClickAndGoFirstSupplier();
            supplierItem.ClickOnItemsTab();
            // supplier-1455071880

            suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, firstSupplierName);

            supplierItem = suppliersPage.ClickAndGoFirstSupplier();
            supplierItem.ClickOnItemsTab();
            supplierItem.DeactivateItems();
            //var secondItemsCount = supplierItem.ItemsCount();
            bool verifdesactivated = supplierItem.VerifDeactivateItems();

            //Assert.AreNotEqual(firstItemsCount, secondItemsCount, "Deactivate items for supplier SupplierTestDeleteItem does not work.");
            Assert.IsTrue(verifdesactivated, "n'est pas encore deactivé");

        }

        // ________________________________________ Utilitaire _________________________________________________

        private void CreateNewItem(SupplierItem supplierItem, string supplierName)
        {
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string itemName = "supplier-" + new Random().Next().ToString();
            string storageUnit = "KG";
            string supplierRef = itemName + "_SupplierRef";
            string storageQty = 2.32.ToString();

            // Go to Item Tab in Supplier Item Page
            supplierItem.PageUp();
            supplierItem.GetItemsTab();

            // Item and packaging creation
            var itemCreatePage = supplierItem.ItemCreatePage();
            var itemGeneralInformationPage = itemCreatePage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemCreatePackagingPage.FillField_CreateNewPackaging(
                site, packagingName, 2.32.ToString(), storageUnit, 4.2.ToString(), supplierName, null, supplierRef, storageQty);
        }       
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_DETAIL_DELIVERYDAY()
        {
            string random_supplier_name = "test-supplier-" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            var supplierCreateModalpage = suppliersPage.SupplierCreatePage();
            supplierCreateModalpage.FillField_CreatNewSupplier(random_supplier_name, false, false);
            var supplierItem = supplierCreateModalpage.Submit();
            supplierItem.SelectAllDeliveryweekdays();
            bool isAlldeliveryWeekDays = supplierItem.AreAllDeliveryweekdays();
            Assert.IsTrue(isAlldeliveryWeekDays, "Le choix des jours de livraison n'est pas sauvegardée");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_DETAIL_NEWITEM()
        {
            //Prepare
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();

            string itemName = "supplier-" + new Random().Next().ToString();
            string storageUnit = "KG";
            string supplierRef = itemName + "_SupplierRef";
            string storageQty = 2.32.ToString();

            string random_supplier_name = "test-supplier-" + new Random().Next().ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            // Create
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            var supplierCreateModalpage = suppliersPage.SupplierCreatePage();
            supplierCreateModalpage.FillField_CreatNewSupplier(random_supplier_name, true, false);
            var supplierItem = supplierCreateModalpage.Submit();
            supplierItem.PageUp();
            supplierItem.GetItemsTab();
            var itemCreatePage = supplierItem.ItemCreatePage();
            var itemGeneralInformationPage = itemCreatePage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemCreatePackagingPage.FillField_CreateNewPackaging(
                site, packagingName, 2.32.ToString(), storageUnit, 4.2.ToString(), random_supplier_name, null, supplierRef, storageQty);
            suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, random_supplier_name);
            suppliersPage.SelectFirstItem();
            supplierItem.GetItemsTab();
            supplierItem.ResetFilter();
            supplierItem.Filter(SupplierItem.FilterType.Search, itemName);
            string filtredItem = supplierItem.getItemName();
            Assert.AreEqual(itemName, filtredItem, "L'item n'a pas été crée.");
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            Assert.AreEqual(itemName, itemPage.GetFirstItemName(), "L'item n'a pas été crée.");

        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_INDEX_DELIVERYDAY()
        {
            string AGPsite = TestContext.Properties["SiteToFlightBob"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            bool change = true;
            //Arrange
            var homePage = LogInAsAdmin();
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.DeliveryDay();
            suppliersPage.SetSiteOnDeliveryDay(AGPsite);
            suppliersPage.SetSupplierOnDeliveryDay(supplier);
            string[] deliveryday = suppliersPage.SelectedDeliveriesDays();
            suppliersPage.ResetFilter();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, supplier);
            var supplierItem = suppliersPage.SelectFirstItem();
            var supplierDeliveriesTab = supplierItem.ClickOnDeliveriesTab();
            string[] deliverydayAfterSave = supplierDeliveriesTab.GetAllColorsDeliveryDay(AGPsite);

            for (int i = 0; i < 7; i++)
            {
                if (!deliveryday[i].Equals(deliverydayAfterSave[i]))
                {
                    change = false;
                    break;
                }
            }

            Assert.IsTrue(change, "Le choix des jours de livraison n'est pas sauvegardée");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_INDEX_DUPLICATESUPPLIER()
        {
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string random_supplier_name = "test-supplier-" + new Random().Next().ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            ClearCache();
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            suppliersPage.Filter(SuppliersPage.FilterType.ShowAll, true);
            var realNbrSupplier = suppliersPage.CheckTotalNumber();

            suppliersPage.DuplicateSupplier();
            suppliersPage.SetFromSupplierDuplicateSupplier(supplier);
            suppliersPage.SetNewSupplierDuplicateSupplier(random_supplier_name);
            //FIXME
            suppliersPage.UncheckEDI();

            //OK FIXME mauvaise page de retour : on est dans General Info
            SupplierItem items = suppliersPage.ClickButtonUpdate();
            items.ClickOnGeneralInformationTab().ClickOnItemsTab();

            ItemCreateModalPage newItemModal = items.ItemCreatePage();

            ItemGeneralInformationPage generalInfo = newItemModal.FillField_CreateNewItem("item_" + random_supplier_name, "AIR CANADA", "Cocina Caliente", "1-Reducido", "KG");
            var packModal = generalInfo.NewPackaging();
            generalInfo = packModal.FillField_CreateNewPackaging("ACE", "KG", "5", "KG", "58", random_supplier_name);

            suppliersPage = items.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            //OK FIXME : supplier sans item donc non affiché dans l'index
            suppliersPage.Filter(SuppliersPage.FilterType.ShowAll, true);

            var realNbr = suppliersPage.CheckTotalNumber();
            //Assert
            Assert.AreEqual(realNbrSupplier + 1, realNbr, " Le fournisseur n'a pas été dupliqué de nouveau cas 1");
            suppliersPage.Filter(SuppliersPage.FilterType.Search, random_supplier_name);
            Assert.AreEqual(suppliersPage.GetFirstSupplierName(), random_supplier_name, " Le fournisseur n'a pas été dupliqué de nouveau cas 2");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_DETAILSTORAGE_FILTERSITE()
        {
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string itemName = "supplier-" + new Random().Next().ToString();
            string storageUnit = "KG";
            string supplierRef = itemName + "_SupplierRef";
            string storageQty = 2.32.ToString();
            string random_supplier_name = "test-supplier-" + new Random().Next().ToString();
            string limitQty = 10.ToString();
            string place = TestContext.Properties["InventoryPlace"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            var supplierCreateModalpage = suppliersPage.SupplierCreatePage();
            supplierCreateModalpage.FillField_CreatNewSupplier(random_supplier_name, true, false);
            var supplierItem = supplierCreateModalpage.Submit();
            supplierItem.PageUp();
            supplierItem.GetItemsTab();
            var itemCreatePage = supplierItem.ItemCreatePage();
            var itemGeneralInformationPage = itemCreatePage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemCreatePackagingPage.FillField_CreateNewPackaging(
                site, packagingName, 2.32.ToString(), storageUnit, 4.2.ToString(), random_supplier_name, null, supplierRef, storageQty);

            var inventoryPage = homePage.GoToWarehouse_InventoriesPage();
            var inventoryCreateModalPage = inventoryPage.InventoryCreatePage();
            inventoryCreateModalPage.FillField_CreateNewInventory(DateUtils.Now, site, place);
            var itemInventory = inventoryCreateModalPage.Submit();
            itemInventory.Filter(InventoryItem.FilterItemType.SearchByName, itemName);
            itemInventory.SelectFirstItem();
            itemInventory.AddPhysicalQuantity(itemName, limitQty);
            InventoryValidationModalPage modal2 = itemInventory.Validate();
            modal2.ValidatePartialInventory();

            homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, random_supplier_name);
            suppliersPage.SelectFirstItem();
            supplierItem.GoToStorageTab();

            supplierItem.FilterBySitesOnStorageTab(site);
            Assert.IsTrue(supplierItem.AreFiltredSitesStorageTabCorrect(site), "Les résultats  affichés  ne correspondent pas au filtre choisi.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_DETAILSRORAGE_SELECTALL()
        {
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string itemName = "supplier-" + new Random().Next().ToString();
            string storageUnit = "KG";
            string supplierRef = itemName + "_SupplierRef";
            string storageQty = 2.32.ToString();
            string random_supplier_name = "test-supplier-" + new Random().Next().ToString();
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            //Act
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            var supplierCreateModalpage = suppliersPage.SupplierCreatePage();
            supplierCreateModalpage.FillField_CreatNewSupplier(random_supplier_name, true, false);
            var supplierItem = supplierCreateModalpage.Submit();
            supplierItem.PageUp();
            supplierItem.GetItemsTab();
            var itemCreatePage = supplierItem.ItemCreatePage();
            var itemGeneralInformationPage = itemCreatePage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemCreatePackagingPage.FillField_CreateNewPackaging(
                site, packagingName, 2.32.ToString(), storageUnit, 4.2.ToString(), random_supplier_name, null, supplierRef, storageQty);
            //SupplierItem supplierItem = suppliersPage.SelectFirstItem();
            ReceiptNotesPage receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();
            ReceiptNotesCreateModalPage receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
            receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, random_supplier_name, deliveryPlace));
            ReceiptNotesItem receiptNotesItem = receiptNotesCreateModalpage.Submit();
            receiptNotesItem.SetFirstReceivedQuantity("50");
            var ReceiptNotesQualityChecks = receiptNotesItem.ClickOnChecksTab();
            ReceiptNotesQualityChecks.SetQualityChecks();
            ReceiptNotesQualityChecks.DeliveryAccepted();
            receiptNotesItem.Validate();
            homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, random_supplier_name);
            suppliersPage.SelectFirstItem();
            supplierItem.GoToStorageTab();
            supplierItem.ClickUnSelectAll();
            Assert.IsTrue(supplierItem.AreAllItemsUnSelected(), "Tous les items doivent être désélectionnés");
            supplierItem.ClickSelectAll();
            Assert.IsTrue(supplierItem.AreAllItemsSelected(), "Tous les items doivent être sélectionnés");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_INDEX_AFFICHAGESUPPLIER30()
        {
            var pageSize30 = "30";
            //Arrange
            HomePage homePage = LogInAsAdmin();
            //Act
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            suppliersPage.PageSize(pageSize30);
            Assert.IsTrue(suppliersPage.IsSuppliersEqualsToSizePage(int.Parse(pageSize30)), "Le nombre de suppliers affichés doit correspondre à la taille de la page ");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_INDEX_AFFICHAGESUPPLIER16()
        {
            var pageSize16 = "16";
            //Arrange
            HomePage homePage = LogInAsAdmin();
            //Act
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            suppliersPage.PageSize(pageSize16);
            Assert.IsTrue(suppliersPage.IsSuppliersEqualsToSizePage(int.Parse(pageSize16)), $"Le nombre de suppliers affichés doit correspondre à la taille de la page {pageSize16}");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_INDEX_DESACTIVATEITEM()
        {
            string AGPsite = TestContext.Properties["SiteToFlightBob"].ToString();
            string random_supplier_name = TestContext.Properties["SupplierDeleteItem"].ToString();
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string itemName = "supplier-" + new Random().Next().ToString();
            string storageUnit = "KG";
            string supplierRef = itemName + "_SupplierRef";
            string storageQty = 2.32.ToString();

            //Arrange
            var homePage = LogInAsAdmin();
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, random_supplier_name);
            var supplierItem = suppliersPage.SelectFirstItem();
            supplierItem.PageUp();
            supplierItem.GetItemsTab();
            var itemCreatePage = supplierItem.ItemCreatePage();
            var itemGeneralInformationPage = itemCreatePage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemCreatePackagingPage.FillField_CreateNewPackaging(
                AGPsite, packagingName, 2.32.ToString(), storageUnit, 4.2.ToString(), random_supplier_name, null, supplierRef, storageQty);
            suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.DesactivateItemr();
            suppliersPage.SetSiteDesactivateItem(AGPsite);
            suppliersPage.SetSupplierOnDeliveryDay(random_supplier_name);
            suppliersPage.ClicktCategorieItem();
            suppliersPage.ClickDeleteDesactivateItem();
            suppliersPage.ClickSaveDesactivateItem();
            suppliersPage.ResetFilter();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, random_supplier_name);
            suppliersPage.SelectFirstItem();
            supplierItem.GetItemsTab();
            supplierItem.ResetFilter();
            supplierItem.Filter(SupplierItem.FilterType.BySite, AGPsite);
            Assert.AreEqual(0, supplierItem.NombreItem(), "La catégorie d'item pour le site et le supplier séléctionné n a pas été supprimé.");

        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_DETAIL_SETUNPURCHASABLEITEM()
        {
            //Prepare
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();

            string itemName = "supplier-" + new Random().Next().ToString();
            string storageUnit = "KG";
            string supplierRef = itemName + "_SupplierRef";
            string storageQty = 2.32.ToString();

            string random_supplier_name = "test-supplier-" + new Random().Next().ToString();
            //Arrange
            var homePage = LogInAsAdmin();
            // Create
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            var supplierCreateModalpage = suppliersPage.SupplierCreatePage();
            supplierCreateModalpage.FillField_CreatNewSupplier(random_supplier_name, true, false);
            var supplierItem = supplierCreateModalpage.Submit();
            supplierItem.PageUp();
            supplierItem.GetItemsTab();
            var itemCreatePage = supplierItem.ItemCreatePage();
            var itemGeneralInformationPage = itemCreatePage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemCreatePackagingPage.FillField_CreateNewPackaging(
                site, packagingName, 2.32.ToString(), storageUnit, 4.2.ToString(), random_supplier_name, null, supplierRef, storageQty);
            suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, random_supplier_name);
            suppliersPage.SelectFirstItem();
            supplierItem.GetItemsTab();
            supplierItem.ResetFilter();
            supplierItem.Filter(SupplierItem.FilterType.Search, itemName);
            supplierItem.BTN_SetUnpurchasableItems();
            supplierItem.Set_Site_SetUnpurchasableItems(site);
            supplierItem.Click_SetUnpurchasableItems();
            supplierItem.Click_OK_SetUnpourchasableItems();
            suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, random_supplier_name);
            suppliersPage.SelectFirstItem();
            supplierItem.GetItemsTab();
            supplierItem.ResetFilter();
            supplierItem.Filter(SupplierItem.FilterType.Search, itemName);
            var numberItem = supplierItem.NombreItem();
            Assert.AreEqual(0, numberItem, "Le changement d'articles inachetables n'a pas été réalisé.");


        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_FastCopy()
        {

            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            SupplierItem supplierItem = suppliersPage.SelectFirstItem();
            supplierItem.GetItemsTab();
            supplierItem.ClickOnFirstCopyPrice();
            supplierItem.ClickOnSelectAllForCopyPrice();
            supplierItem.ScrollUntilCopyIsVisible();
            supplierItem.ClickInCopyBtn();

            /* Get the message from the pop up */
            string itemEditedMessage = supplierItem.CopyPopUpMessage();

            bool isItemEditedMessage = itemEditedMessage.Contains("Item edited : ");
            // Assert 
            Assert.IsTrue(isItemEditedMessage, "The Item edited message don't show up");

        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_Deactivate_items()
        {
            string ItemFiltred = "IsActifAndNotPurchasable";
            //Arrange
            HomePage homePage = LogInAsAdmin();
            //Act
            SuppliersPage suppliersPage = homePage.GoToPurchasing_SuppliersPage();

            suppliersPage.DesactivateItemr();

            /* Fill the Deactivate items for supplier form*/
            suppliersPage.SetSiteDesactivateItem_Chceck_All();
            suppliersPage.SetSUPPLIERSDesactivateItem_Chceck_All();
            suppliersPage.SetIteams_DeactivateItems(ItemFiltred);

            /* Click on Delete for the Deactivate items for supplier*/
            suppliersPage.ClickDeleteDesactivateItem();
            /* Get if the pop up is visible based on its title */
            bool is_visible = suppliersPage.ItemsDeactivationReport_IsVisible();

            // Assert
            Assert.IsTrue(is_visible, "The iteams deactivations report don't show up ");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_FastCopy_TimeOut()
        {
            string site = TestContext.Properties["SiteLpCart"].ToString();
            string itemName = "itemToFilter" + new Random().Next().ToString();
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();

            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();
            string limitQty = 177.ToString();
            string unitPrice = 100.ToString();

            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, supplier);
            SupplierItem supplierItem = suppliersPage.SelectFirstItem();
            supplierItem.ClickOnItemsTab();
            supplierItem.ResetFilter();
            if (supplierItem.ItemsCount() == 0)
            {
                ItemCreateModalPage createItem = supplierItem.ItemCreatePage();
                ItemGeneralInformationPage itemGeneralInfo = createItem.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInfo.NewPackaging();
                itemGeneralInfo = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier, limitQty, null, unitPrice);
                supplierItem.BackToList();
                suppliersPage = homePage.GoToPurchasing_SuppliersPage();
                suppliersPage.Filter(SuppliersPage.FilterType.Search, supplier);
                supplierItem = suppliersPage.SelectFirstItem();
                supplierItem.ClickOnItemsTab();
                supplierItem.ResetFilter();

            }
            var firstItemname = supplierItem.getItemName();
            supplierItem.Filter(SupplierItem.FilterType.Search, firstItemname);
            supplierItem.ClickOnFirstCopyPrice();
            supplierItem.SetSiteForFastCopy();
            supplierItem.SelectCheckBoxes(3);
            supplierItem.ScrollUntilCopyIsVisible();
            supplierItem.ClickInCopyBtn();
            string itemEditedMessage = supplierItem.CopyPopUpMessage();
            bool isItemEditedMessage = itemEditedMessage.Contains("Item edited : ");
            Assert.IsTrue(isItemEditedMessage, "The Item edited message don't show up");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_FastCopy_keyword_details()
        {
            string site = TestContext.Properties["SiteLP"].ToString();

            string itemName = "test-item-" + new Random().Next().ToString();

            string group = "A REFERENCIA";
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();

            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            //string supplier = TestContext.Properties["Supplier"].ToString();
            string supplier = "ACEITES CONDE DE TORREPALMA, SCA";
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();


            //Act
            ItemPage itemsPage = homePage.GoToPurchasing_ItemPage();

            ItemGeneralInformationPage itemGeneralInfo = null;

            ItemCreateModalPage createItem = itemsPage.ItemCreatePage();
            itemGeneralInfo = createItem.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInfo.NewPackaging();
            itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
            var PurchchasingSupplier = itemGeneralInfo.GoToPurchasing_SuppliersPage();

            PurchchasingSupplier.Filter(SuppliersPage.FilterType.Search, supplier);

            var supplierItem = PurchchasingSupplier.ClickAndGoFirstSupplier();
            PurchchasingSupplier.GoToItemsInSupplier();
            PurchchasingSupplier.ResetFilter();
            PurchchasingSupplier.FilterItemInSupplier(SuppliersPage.FilterType.Site, site);
            PurchchasingSupplier.FilterItemInSupplier(SuppliersPage.FilterType.Search, itemName);

            int NumItemInSupplierACP = PurchchasingSupplier.GetItemsInSupplierCount();
            PurchchasingSupplier.ResetFilter();
            PurchchasingSupplier.FilterItemInSupplier(SuppliersPage.FilterType.Site, "AGP");

            PurchchasingSupplier.FilterItemInSupplier(SuppliersPage.FilterType.Search, itemName);

            int NumItemInSupplierAGP = PurchchasingSupplier.GetItemsInSupplierCount();
            Assert.AreEqual(1, NumItemInSupplierACP);
            Assert.AreEqual(0, NumItemInSupplierAGP);

            supplierItem.FastCopyPrice("ACE", "AGP");
            PurchchasingSupplier.ResetFilter();
            PurchchasingSupplier.FilterItemInSupplier(SuppliersPage.FilterType.Site, "AGP");

            PurchchasingSupplier.FilterItemInSupplier(SuppliersPage.FilterType.Search, itemName);

            int NumItemInSupplierAGPAFTER = PurchchasingSupplier.GetItemsInSupplierCount();

            Assert.AreEqual(1, NumItemInSupplierAGPAFTER);
            supplierItem.ClickOnFirstItems();
            supplierItem.ClickKeyword();
            supplierItem.SearchKeywords("AGP");
            bool VerifKeyword = supplierItem.CountKEYWORD();
            Assert.IsTrue(VerifKeyword, "Keyword n'est pas ajouté");

        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_Contact_Duplicate()
        {
            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();

            // Select a random supplier item
            SupplierItem supplierItem = suppliersPage.SelectRandomItem();

            // Navigate to the contact tab
            //var supplierContactTab = new SupplierContactTab(WebDriver, TestContext).GoToContact();
            supplierItem.GoToContactTab();
            // Check if there are existing contacts
            var existingContacts = supplierItem.GetContactList();

            if (existingContacts.Any())
            {
                supplierItem.DeleteAllContacts();
                while (supplierItem.GetContactList().Any())
                {
                    supplierItem.DeleteAllContacts();
                }
            }
            var supplierContactTab = supplierItem.ClickOnNewContactButton();
            supplierContactTab.FillNewContactForm("Nom du contact", "email@example.com", "123456789", "Adresse", "Code postal", "Ville");
            supplierContactTab.SubmitNewContactForm();

            // Wait for the new contact to appear in the list
            WebDriverWait waitNewContact = new WebDriverWait(WebDriver, TimeSpan.FromSeconds(60));
            waitNewContact.Until(d => supplierContactTab.GetContactList().Any(c => c.Contains("Nom du contact")));

            // Retrieve the list of contacts after creation
            var contactList = supplierContactTab.GetContactList();

            // Assert - Verify that the contact is not duplicated
            int contactCount = contactList.Count(c => c.Contains("Nom du contact"));
            Assert.AreEqual(1, contactCount, "Le contact créé est dupliqué, ce qui n'était pas attendu.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_Storage_TimeOut()
        {
            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();

            var supplierItem = suppliersPage.SelectFirstItem();
            var stopwatch = new Stopwatch();
            stopwatch.Start();
            supplierItem.GoToStorageTab();
            stopwatch.Stop();

            //Assert - Vérification que l'opération ne prend pas trop de temps
            double maxAllowedTime = 5000;
            Assert.IsTrue(stopwatch.ElapsedMilliseconds < maxAllowedTime,
                $"L'opération a pris trop de temps : {stopwatch.ElapsedMilliseconds}ms, alors que le maximum autorisé est {maxAllowedTime}ms.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_IBAN()
        {
            // Préparation
            string random_supplier_name = "test-supplier-" + new Random().Next().ToString();
            string iban = "FR7630006000011234567890189"; // Exemple d'IBAN

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            var supplierCreateModalpage = suppliersPage.SupplierCreatePage();
            supplierCreateModalpage.FillField_CreatNewSupplier(random_supplier_name, true, false);
            var supplierItem = supplierCreateModalpage.Submit();

            CreateNewItem(supplierItem, random_supplier_name);

            suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, random_supplier_name);
            suppliersPage.ClickAndGoFirstSupplier();
            supplierItem.ClickOnGeneralInformationTab();

            // Remplir le champ IBAN
            supplierItem.FillIBANField(iban);

            supplierItem.BackToList();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, random_supplier_name);
            suppliersPage.ClickAndGoFirstSupplier();

            // Assert pour vérifier que l'IBAN est bien enregistré
            string savedIban = supplierItem.GetSavedIBAN();

            Assert.AreEqual(iban, savedIban, "L'IBAN n'a pas été enregistré correctement.");
        }

        //[TestMethod]
        //public void PU_SUPP_ExportImportAutoClose()
        //{
        //    string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
        //    LogInAsAdmin();
        //    HomePage homePage = new HomePage(WebDriver, TestContext);
        //    homePage.Navigate();
        //    SuppliersPage suppliersPage = homePage.GoToPurchasing_SuppliersPage();
        //    suppliersPage.ResetFilter();
        //    string supplierName = suppliersPage.GetFirstSupplierName();
        //    suppliersPage.Filter(SuppliersPage.FilterType.Search, supplierName);
        //    SupplierItem supplierItem = suppliersPage.SelectFirstItem();
        //    SupplierGeneralInfoTab supplierGeneralInfoTab = supplierItem.ClickOnGeneralInformationTab();
        //    supplierGeneralInfoTab.CheckAutoCloseRecipeNote();
        //    supplierGeneralInfoTab.BackToList();
        //    suppliersPage.ResetFilter();
        //    suppliersPage.Filter(SuppliersPage.FilterType.Search, supplierName);
        //    suppliersPage.ClearDownloads();
        //    DeleteAllFileDownload();
        //    suppliersPage.Export(true);
        //    DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
        //    FileInfo[] taskFiles = taskDirectory.GetFiles();
        //    FileInfo correctDownloadedFile = suppliersPage.GetExportExcelFile(taskFiles, true);
        //    Assert.IsNotNull(correctDownloadedFile);
        //    string fileName = correctDownloadedFile.Name;
        //    string filePath = Path.Combine(downloadsPath, fileName);
        //    string x = OpenXmlExcel.GetCellValue("a", "Output forms", );
        //}
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_FastCopyPriceIsMain()
        {
            // Arrange 
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
            string itemName = "itemFastCopyPriceIsMain";
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["RecipeVariantBis"].ToString();
            string ingredient = TestContext.Properties["RecipeIngredient"].ToString();
            int nbPortions = new Random().Next(1, 30);

            // Log in as Admin
            HomePage homePage = LogInAsAdmin();

            // Act 
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);

            // If no item exists, create a new item with packaging
            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGenerlInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGenerlInformationPage.NewPackaging();
                itemGenerlInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                itemPage = itemGenerlInformationPage.BackToList();
            }

            homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);

            // Get the state of the 'IsMain' checkbox before fast copy
            ItemGeneralInformationPage itemGeneralInformationPage = itemPage.ClickOnFirstItem();
            var isMainBeforeFastCopy = itemGeneralInformationPage.IsCheckboxChecked();

            // Perform Fast Copy Price operation
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, supplier);
            SupplierItem supplierItem = suppliersPage.SelectFirstItem();
            SupplierGeneralInfoTab supplierGeneralInfoTab = supplierItem.ClickOnGeneralInformationTab();
            supplierGeneralInfoTab.ClickOnItemTab();
            supplierItem.Filter(SupplierItem.FilterType.Search, itemName);
            supplierItem.Filter(SupplierItem.FilterType.BySite, site);
            supplierItem.ClickOnFastCopyPrice();
            supplierItem.FilterPopup(SupplierItem.FilterTypePopup.Site, site);
            supplierItem.FilterPopup(SupplierItem.FilterTypePopup.Search, site);
            supplierItem.CopyPriceFromSupplier();

            // Re-check the item and compare 'IsMain' checkbox state after fast copy
            homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            itemGeneralInformationPage = itemPage.ClickOnFirstItem();
            var isMainAfterFastCopy = itemGeneralInformationPage.IsCheckboxChecked();

            // Assert 
            Assert.AreEqual(isMainAfterFastCopy, isMainBeforeFastCopy, "The 'IsMain' checkbox state changed after fast copy.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_ExportContacts()
        {
            //Prepare
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string downloadPath = TestContext.Properties["DownloadsPath"].ToString();

            //arrange
            HomePage homePage = LogInAsAdmin();

            homePage.ClearDownloads();
            var supplierspage = homePage.GoToPurchasing_SuppliersPage();
            supplierspage.ResetFilter();

            supplierspage.Filter(SuppliersPage.FilterType.Search, supplier);
            int totalLignes = supplierspage.CheckTotalNumber();

            if (supplierspage.CheckTotalNumber() == 0)
            {
                var supplierCreateModalpage = supplierspage.SupplierCreatePage();
                supplierCreateModalpage.FillField_CreatNewSupplier(supplier, true, false);
                SupplierItem supplierItem = supplierCreateModalpage.Submit();
                CreateNewItem(supplierItem, supplier);
                supplierspage = homePage.GoToPurchasing_SuppliersPage();
                supplierspage.Filter(SuppliersPage.FilterType.Search, supplier);
                totalLignes = supplierspage.CheckTotalNumber();
            }

            SupplierItem supplierItemNew = supplierspage.SelectFirstItem();
            // Add a contact to the supplier
            string contactGenerate = supplierItemNew.AddContactWithoutSite("test@mail.com");

            supplierItemNew.BackToList();
            supplierspage.Filter(SuppliersPage.FilterType.Search, supplier);

            supplierspage.Export(true);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadPath);
            taskDirectory = new DirectoryInfo(downloadPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            FileInfo correctDownloadedFile = supplierspage.GetExportExcelFile(taskFiles, true);

            // Récupération du nom du fichier et construction de l'URL du fichier Excel à ouvrir   
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadPath, fileName);

            //Le fichier Excel est généré
            Assert.IsNotNull(correctDownloadedFile, "fichier non généré cas 1");
            Assert.IsTrue(correctDownloadedFile.Exists, "fichier non généré cas 2");

            //Vérifier les données
            var resultNumber = OpenXmlExcel.GetExportResultNumber("Suppliers", correctDownloadedFile.FullName);
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
            List<string> valeurs = OpenXmlExcel.GetValuesInList("Supplier Name", "Suppliers", correctDownloadedFile.FullName);
            Assert.AreEqual(totalLignes, valeurs.Count);

            int offset = valeurs.IndexOf(supplier);
            supplierspage.ClosePrintButton();

            //supplierspage.Filter(SuppliersPage.FilterType.Search, supplier);
            SupplierItem item = supplierspage.SelectFirstItem();
            SupplierContactTab contactsTab = item.ClickOnContactTab();

            List<string> valeursContacts = OpenXmlExcel.GetValuesInList("Name", "Contacts", correctDownloadedFile.FullName);

            // Assert
            bool isNameEqual = valeursContacts.Contains(contactGenerate.Trim());
            Assert.IsTrue(isNameEqual, "Il faut définir un site sur un contact de fournisseur, pour qu'il  figure sur l'onglet Contact du fichier Excel.");


        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_Export_KnownSupplier()
        {
            string downloadPath = TestContext.Properties["DownloadsPath"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            HomePage homePage = LogInAsAdmin();
            homePage.ClearDownloads();
            DeleteAllFileDownload();
            SuppliersPage supplierspage = homePage.GoToPurchasing_SuppliersPage();
            supplierspage.ResetFilter();
            supplierspage.Filter(SuppliersPage.FilterType.Search, supplier);
            SupplierItem supplierItem = supplierspage.SelectFirstItem();
            List<string> checkedKnownSupplier = supplierItem.GetCheckedKnownSupplier();
            supplierItem.BackToList();
            supplierspage.ResetFilter();
            supplierspage.Filter(SuppliersPage.FilterType.Search, supplier);
            supplierspage.Export(true);
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadPath);
            taskDirectory = new DirectoryInfo(downloadPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();
            FileInfo correctDownloadedFile = supplierspage.GetExportExcelFile(taskFiles, true);
            Assert.IsNotNull(correctDownloadedFile, "fichier non généré cas 1");
            string fileName = correctDownloadedFile.Name;
            string filePath = Path.Combine(downloadPath, fileName);
            List<string> knownSuppliers = OpenXmlExcel.GetValuesInList("Known supplier", "Supplier", filePath);

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_GeneralInfo_AutoCloseRN()
        {
            //Prepare           
            string site = TestContext.Properties["Site"].ToString();
            string random_supplier_name = "test-supplier-" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();

            if (suppliersPage.CheckTotalNumber() == 0)
            {
                var supplierCreateModalpage = suppliersPage.SupplierCreatePage();
                supplierCreateModalpage.FillField_CreatNewSupplier(random_supplier_name, true, false);
                var supplierItem = supplierCreateModalpage.Submit();
            }

            var firstSupplier = suppliersPage.GetFirstSupplierName();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, firstSupplier);
            var generalInfo = suppliersPage.ClickAndGoFirstSupplier();

            bool status = generalInfo.GetCheckAutoCloseStatus();
            generalInfo.CheckAutoClose();
            suppliersPage = generalInfo.BackToList();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, firstSupplier);
            generalInfo = suppliersPage.ClickAndGoFirstSupplier();

            if (status == true)
            {
                status = generalInfo.GetCheckAutoCloseStatus();
                Assert.IsFalse(status, "AutoCloseRN ne fonctionne pas");
            }
            else
            {
                status = generalInfo.GetCheckAutoCloseStatus();
                Assert.IsTrue(status, "AutoCloseRN ne fonctionne pas");
            }
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_FilterSupplierType()
        {
            //Arrange
            HomePage homePage = LogInAsAdmin();
            SuppliersPage suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, supplierName);
            suppliersPage.Filter(SuppliersPage.FilterType.SupplierType, supplierType);
            string firstSupplierName = suppliersPage.GetFirstSupplierName();
            //Assert
            Assert.AreEqual(firstSupplierName, supplierName, String.Format(MessageErreur.FILTRE_ERRONE, "'Search by Supplier Type'"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_ExportImportAutoCloseRN()
        {
            //Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString(); // C:\ChromeDriverDownloads
            string random_supplier_name = "test-supplier-" + new Random().Next().ToString();
            bool newVersionPrint = true;

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            suppliersPage.ClearDownloads();

            if (suppliersPage.CheckTotalNumber() == 0)
            {
                SupplierCreateModalPage supplierCreateModalpage = suppliersPage.SupplierCreatePage();
                supplierCreateModalpage.FillField_CreatNewSupplier(random_supplier_name, true, false);
                SupplierItem supplierItem = supplierCreateModalpage.Submit();
            }

            string firstSupplier = suppliersPage.GetFirstSupplierName();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, firstSupplier);
            SupplierItem generalInfo = suppliersPage.ClickAndGoFirstSupplier();

            bool status = generalInfo.GetCheckAutoCloseStatus();
            if (!status)
            {
                //cochez la case "Auto Close Receipt Note"
                generalInfo.CheckAutoClose();
            }
            //Exporter le fichier
            suppliersPage = generalInfo.BackToList();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, firstSupplier);
            suppliersPage.Export(newVersionPrint);

            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté.
            FileInfo correctDownloadedFile = suppliersPage.GetExportExcelFile(taskFiles, true);
            Assert.IsNotNull(correctDownloadedFile);

            string fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);
            bool vraiExist = OpenXmlExcel.ReadAllDataInColumn("Auto Close", "Suppliers", filePath, "VRAI");

            //Assert
            Assert.IsTrue(vraiExist, "L'export la colonne Auto close n'affiche pas la valeur 'VRAI'");

            //Modifier le fichier à importer
            OpenXmlExcel.WriteDataInCell("Auto Close", "Supplier", random_supplier_name, "Suppliers", filePath, "Faux", CellValues.SharedString);
            WebDriver.Navigate().Refresh();
            bool importPopup = suppliersPage.Import(correctDownloadedFile);
            Assert.IsTrue(importPopup, " Erreur dans l'importation du fichier");

            suppliersPage.Filter(SuppliersPage.FilterType.Search, firstSupplier);
            generalInfo = suppliersPage.ClickAndGoFirstSupplier();
            status = generalInfo.GetCheckAutoCloseStatus();

            Assert.IsTrue(status, "La case Auto Close Receipt Note a changé après modifications sur le fichier d'import");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_ReverseVAT()
        {
            //Prepare
            bool value = true;
            string random_supplier_name = "test-supplier-" + new Random().Next().ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            SuppliersPage suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            // Create
            SupplierCreateModalPage supplierCreateModalpage = suppliersPage.SupplierCreatePage();
            supplierCreateModalpage.FillField_CreatNewSupplier(random_supplier_name, value, value);
            SupplierItem supplierItem = supplierCreateModalpage.Submit();

            //check reverse VAT 
            supplierItem.checkReverseVAT();
            supplierItem.WaitPageLoading();
            WebDriver.Navigate().Refresh();
            //Assert
            bool isreverseVatCheckedOk = supplierItem.IsReverseVatChecked();
            Assert.IsTrue(isreverseVatCheckedOk, "La case Reverse VAT soit toujours cochée.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_VegetableItem()
        {

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
            string unitPrice = "53";
            string itemName = "itemIsVegetable" + "-" + new Random().Next().ToString();
            bool itemIsVegetable = true;

            HomePage homePage = LogInAsAdmin();
            try
            {
                ParametersSites siteParameterPage = homePage.GoToParameters_Sites();
                siteParameterPage.Filter(ParametersSites.FilterType.SearchSite, site);
                siteParameterPage.CheckIfFirstSiteIsActive();
                SuppliersPage suppliersPage = homePage.GoToPurchasing_SuppliersPage();
                suppliersPage.Filter(SuppliersPage.FilterType.ShowOnlyActive, true);
                suppliersPage.Filter(SuppliersPage.FilterType.Search, supplier);
                SupplierItem supplierItem = suppliersPage.SelectFirstItem();
                supplierItem.ClickOnItemsTab();
                supplierItem.ResetFilter();
                ItemCreateModalPage createItem = supplierItem.ItemCreatePage();
                ItemGeneralInformationPage itemGeneralInfo = createItem.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInfo.NewPackaging();
                itemGeneralInfo = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier, null, null, unitPrice);
                itemGeneralInfo.SetCheckoxIsVegetable(itemIsVegetable);
                ItemDieteticPage dietetic = itemGeneralInfo.ClickOnDieteticPage();
                dietetic.SetMacroNutriments("Fibres", "20");
                itemGeneralInfo = dietetic.ClickOnGeneralInfo();
                bool verifIsVegetable = itemGeneralInfo.GetCheckboxIsVegetable();
                Assert.AreEqual(itemIsVegetable, verifIsVegetable, "Le Checkbox IsVegetable de l'item n'a pas été modifié.");
            }
            finally
            {

                ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
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
        public void PU_SUPP_Deliv_TimeLimit()
        {
            //Prepare 
            string updateTime = "1100AM";

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            SuppliersPage suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, supplierName);
            suppliersPage.Filter(SuppliersPage.FilterType.SupplierType, supplierType);

            //Click on first Supplier
            SupplierItem supplierItem = suppliersPage.SelectFirstItem();
            supplierItem.ClickToDeliveryBtn();

            //Get TimeLimit before Update 
            string originStartTime = supplierItem.GetStartTime();
            //Edit TimeLimit
            supplierItem.EditStartTime(updateTime);
            //Refresh Page 
            WebDriver.Navigate().Refresh();
            supplierItem.ClickToDeliveryBtn();
            //Get TimeLimit after Update 
            string editedStartTime = supplierItem.GetStartTime();

            //Assert
            Assert.AreNotEqual(originStartTime, editedStartTime, "Time limit was not edited correctly.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_DELI_MINAmount()
        {
            //Prepare 
            string minAmountToSet = "50";

            //Arrange
            HomePage homePage = LogInAsAdmin();
            //Act
            SuppliersPage suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, supplierName);
            suppliersPage.Filter(SuppliersPage.FilterType.SupplierType, supplierType);

            //Click on first Supplier
            SupplierItem supplierItem = suppliersPage.SelectFirstItem();
            supplierItem.ClickToDeliveryBtn();
            //Edit Min Amount
            supplierItem.EditMinAmount(minAmountToSet);
            //Refresh Page 
            supplierItem.GoToContactTab();
            supplierItem.ClickToDeliveryBtn();
            //Get TimeLimit after Update 
            string diplayedMinAmount = supplierItem.GetMinAmount();

            //Assert
            Assert.AreEqual(diplayedMinAmount, minAmountToSet, "La modification du champ Min Amount n'a pas été enregistrée.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_Storage_FilterItemName()
        {
            //Prepare
            string startName = itemName1.Substring(0, 4);

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            SuppliersPage suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, supplierName);
            suppliersPage.Filter(SuppliersPage.FilterType.SupplierType, supplierType);

            //Click on first Supplier
            SupplierItem supplierItem = suppliersPage.SelectFirstItem();
            supplierItem.GoToStorageTab();
         
            //Filter by name
            supplierItem.FilterByNameOnStorageTab(startName);
            var newValueName = supplierItem.GetItemNameSupplier();
           
            //Assert
            Assert.AreEqual(itemName1, newValueName, "le nom de l'item affiché ne correspond pas au nom mis au filtre.");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_DELI_ShippingCost()
        {
            //Prepare 
            string shippingCostToSet = "22";

            //Arrange
            HomePage homePage = LogInAsAdmin();
            //Act
            SuppliersPage suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, supplierName);
            suppliersPage.Filter(SuppliersPage.FilterType.SupplierType, supplierType);

            //Click on first Supplier
            SupplierItem supplierItem = suppliersPage.SelectFirstItem();
            supplierItem.ClickToDeliveryBtn();
            //Edit shipping Cost
            supplierItem.EditShippingCost(shippingCostToSet);
            //Refresh Page 
            supplierItem.GoToContactTab();
            supplierItem.ClickToDeliveryBtn();
            //Get  shipping Cost after Update 
            string diplayedshippingCost = supplierItem.GetShippingCost();

            //Assert
            Assert.AreEqual(diplayedshippingCost, shippingCostToSet, "La modification du champ SHIPPING COST n'a pas été enregistrée.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_SUPP_NewItem_IsCSR()
        {
            //prepare
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string itemName = "supplier-" + new Random().Next().ToString();
            string storageUnit = "KG";
            string supplierRef = itemName + "_SupplierRef";
            string storageQty = "2.32";
            bool IsCSR = true;

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            SuppliersPage suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();

            //Click on first Supplier active
            suppliersPage.Filter(SuppliersPage.FilterType.ShowOnlyActive, true);
            string displayedSupplier = suppliersPage.GetFirstSupplierName();
            SupplierItem supplierItem = suppliersPage.SelectFirstItem();

            //item
            suppliersPage.GoToItemsInSupplier();

            // Go to Item Tab in Supplier Item Page
            supplierItem.PageUp();
            supplierItem.GetItemsTab();

            // Item and packaging creation
            var itemCreatePage = supplierItem.ItemCreatePage();
            var itemGeneralInformationPage = itemCreatePage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit, null, null, null, null, null, false, false, false, false, IsCSR);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();

            itemCreatePackagingPage.FillField_CreateNewPackagingAllSites(packagingName, storageQty, storageUnit, storageQty, displayedSupplier, null, supplierRef);

            //get item name
            nameItem = itemGeneralInformationPage.GetItemName();
            suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            supplierItem = suppliersPage.SelectFirstItem();
            suppliersPage.GoToItemsInSupplier();
     
           //Filter
            supplierItem.Filter(SupplierItem.FilterType.Search, nameItem);
            var itemGeneralInfo =  supplierItem.SelectItem();

            //Assert
            bool isCsr = itemGeneralInfo.GetCheckboxisCsr();
            Assert.IsTrue(isCsr, "La case 'Is CSR' n'est pas cochée");
        }
    }
}
