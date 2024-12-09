using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.ExtendedProperties;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Production;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Sites;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.Production;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.ProductionCO;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.ProductionManagement;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.Claims;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.ReceiptNotes;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography.Xml;
using System.Security.Policy;
using System.Threading;
using System.Web.UI;
using UglyToad.PdfPig;
using static Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.SupplyOrderPage;
using static Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.PurchaseOrderItem;
using static Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.SupplyOrderItem;
using Page = UglyToad.PdfPig.Content.Page;

namespace Newrest.Winrest.FunctionalTests.Purchasing
{
    [TestClass]
    public class PurchaseOrderTests : TestBase
    {
        private const int _timeout = 600000;
        private const string PURCHASE_ORDERS_EXCEL_SHEET_NAME = "Purchase Orders";
        private readonly string itemNameToday = "Item-" + DateUtils.Now.ToString("dd/MM/yyyy");

        /// <summary>
        /// 
        /// Mise en place du paramétrage pour la configuration Winrest 4.0 
        /// 
        /// </summary>
        [TestMethod]
        [Priority(0)]
        [Timeout(_timeout)]
        public void PU_PO_SetConfigWinrest4_0()
        {
            // Prepare
            var keyword = TestContext.Properties["Item_Keyword"].ToString();
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            ClearCache();

            // New version search
            homePage.SetNewVersionSearchValue(true);

            // New keyword search
            homePage.SetNewVersionKeywordValue(true);

            // New group display
            homePage.SetNewGroupDisplayValue(true);

            // Vérifier que c'est activé
            var purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            purchaseOrdersPage.ResetFilters();
            purchaseOrdersPage.Filter(PurchaseOrdersPage.FilterType.FilterShow, "Show not validated only");
            var purchaseOrderItemPage = purchaseOrdersPage.SelectFirstItem();

            // Vérifier New version search
            try
            {
                var itemSearched = purchaseOrderItemPage.GetFirstItemName();
                if (itemSearched == "")
                {
                    purchaseOrderItemPage.BackToList();
                    var createPurchaseOrderPage = purchaseOrdersPage.CreateNewPurchaseOrder();
                    createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+10), true);
                    purchaseOrderItemPage = createPurchaseOrderPage.Submit();
                    var ID = purchaseOrderItemPage.GetPurchaseOrderNumber();
                    purchaseOrderItemPage.SelectFirstItemPo();
                    purchaseOrderItemPage.AddQuantity("2");
                }
                itemSearched = purchaseOrderItemPage.GetFirstItemName();
                purchaseOrderItemPage.Filter(FilterItemType.ByName, itemSearched);
            }
            catch
            {
                throw new Exception("La recherche a pu être effectuée, le NewSearchMode est actif.");
            }

            // vérifier new keyword search
            try
            {
                purchaseOrderItemPage.ResetFilter();
                purchaseOrderItemPage.Filter(FilterItemType.ByKeyword, keyword);
            }
            catch
            {
                throw new Exception("La recherche par keyword n'a pas pu être effectuée, le NewKeywordMode est inactif.");
            }

            // vérifier new group display
            Assert.IsTrue(purchaseOrderItemPage.IsGroupDisplayActive(), "Le paramètre 'NewGroupDisplay' n'est pas activé.");
        }

        [TestMethod]
        [Priority(1)]
        [Timeout(_timeout)]
        public void PU_PO_Create_Supplier_For_PurchaseOrder()
        {
            //Prepare
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var site = TestContext.Properties["Site"].ToString();
            var contactEmail = TestContext.Properties["Admin_UserName"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, supplier);

            if (suppliersPage.CheckTotalNumber() == 0)
            {
                var supplierCreateModalpage = suppliersPage.SupplierCreatePage();
                supplierCreateModalpage.FillField_CreatNewSupplier(supplier, false, false);
                var supplierItem = supplierCreateModalpage.Submit();
                suppliersPage = supplierItem.BackToList();
            }

            //check if contact email adress wr.testauto
            var supplierItemPage = suppliersPage.SelectFirstItem();
            supplierItemPage.GoToContactTab();
            supplierItemPage.SearchAndEditContact(supplier, site, contactEmail);
            supplierItemPage.BackToList();

            suppliersPage.ResetFilter();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, supplier);

            // Assert
            Assert.AreEqual(supplier, suppliersPage.GetFirstSupplierName(), "Le supplier créé n'est pas présent dans la liste.");
        }
        [Timeout(_timeout)]
        //Créer un nouveau Purchase Order       
        [TestMethod]
        public void PU_PO_Create_New_Purchase_Order()
        {
            //Prepare
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();

            //Arrange
            LogInAsAdmin();

            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();

            var createPurchaseOrderPage = purchaseOrdersPage.CreateNewPurchaseOrder();
            createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+10), true);
            var purchaseOrderItemPage = createPurchaseOrderPage.Submit();
            var ID = purchaseOrderItemPage.GetPurchaseOrderNumber();
            purchaseOrderItemPage.SelectFirstItemPo();
            purchaseOrderItemPage.AddQuantity("2");
            purchaseOrdersPage = purchaseOrderItemPage.BackToList();

            purchaseOrdersPage.Filter(PurchaseOrdersPage.FilterType.ByNumber, ID);

            //Assert
            Assert.AreEqual(ID, purchaseOrdersPage.GetFirstPurchaseOrderNumber(), "Le purchase order créé n'a pas été ajouté à la liste.");
        }
        [Timeout(_timeout)]
        //Créer un nouveau Purchase Order à partir du Purchase Order déjà existant
        [TestMethod]
        public void PU_PO_CreationNewPurchaseOrderFromPO()
        {
            //Prepare
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();
            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            var purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            // On créé un purchase order
            var purchaseOrderCreatePage = purchaseOrdersPage.CreateNewPurchaseOrder();
            purchaseOrderCreatePage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+1), true);
            var poItemPage = purchaseOrderCreatePage.Submit();
            string fromPONumber = poItemPage.GetPurchaseOrderNumber();
            string itemName = poItemPage.GetFirstItemName();
            poItemPage.SelectFirstItemPo();
            poItemPage.AddQuantity("2");
            // unselect line
            poItemPage.ResetFilter();
            var pricePOFrom = poItemPage.GetTotalSum();
            purchaseOrdersPage = poItemPage.BackToList();
            var createPurchaseOrderPage = purchaseOrdersPage.CreateNewPurchaseOrder();
            createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+1), true, true, fromPONumber);
            var purchaseOrderItemPage = createPurchaseOrderPage.Submit();
            var pricePOTo = purchaseOrderItemPage.GetTotalSum();
            Assert.AreEqual(pricePOFrom, pricePOTo, "Mauvais tarif");
            string ID = purchaseOrderItemPage.GetPurchaseOrderNumber();
            var firstItemName = purchaseOrderItemPage.GetFirstItemName();
            Assert.AreNotEqual("", firstItemName, "Le purchase order contient bien des items");
            purchaseOrderItemPage.Filter(FilterItemType.ByName, itemName);
            Assert.IsTrue(purchaseOrderItemPage.VerifyName(itemName), "Le purchase order créé ne contient pas les mêmes items que celui à partir duquel il a été créé.");
            purchaseOrdersPage = purchaseOrderItemPage.BackToList();
            purchaseOrdersPage.Filter(PurchaseOrdersPage.FilterType.ByNumber, ID);
            var firstPurchaseOrderNumber = purchaseOrdersPage.GetFirstPurchaseOrderNumber();
            Assert.AreEqual(ID, firstPurchaseOrderNumber, "Le purchase order créé n'a pas été ajouté à la liste.");
        }
        [Timeout(_timeout)]
        //Modifier les informations présentes dans l'onglet general information
        [TestMethod]
        public void PU_PO_GeneralDetailModification()
        {
            //Prepare
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();

            //Arrange
            LogInAsAdmin();

            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            var createPurchaseOrderPage = purchaseOrdersPage.CreateNewPurchaseOrder();
            createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+1), true);
            var purchaseOrderItemPage = createPurchaseOrderPage.Submit();
            var ID = purchaseOrderItemPage.GetPurchaseOrderNumber();


            var puchaseOrderGeneralInformation = purchaseOrderItemPage.ClickOnGeneralInformation();

            // Récupération des valeurs à modifier
            var oldDeliveryDate = puchaseOrderGeneralInformation.GetDeliveryDateValue();
            var oldComment = puchaseOrderGeneralInformation.GetComment();
            var oldStatus = puchaseOrderGeneralInformation.GetStatus();

            // test pour Activated
            Assert.IsTrue(puchaseOrderGeneralInformation.GetActivated(), "activated pas true");
            puchaseOrderGeneralInformation.SetActivated(false);
            PurchaseOrderItem itemTab = puchaseOrderGeneralInformation.ClickOnItemsTab();
            itemTab.ClickOnGeneralInformation();
            Assert.IsFalse(puchaseOrderGeneralInformation.GetActivated(), "activated pas false");
            puchaseOrderGeneralInformation.SetActivated(true);
            Assert.IsTrue(puchaseOrderGeneralInformation.GetActivated(), "activated pas true");

            // Modification des valeurs
            puchaseOrderGeneralInformation.DeliveryDateUpdate(DateUtils.Now.AddDays(+5));
            puchaseOrderGeneralInformation.SetComment("test comment");
            puchaseOrderGeneralInformation.SetStatus("Close");
            puchaseOrderGeneralInformation.SetActivated(true);
            purchaseOrdersPage = puchaseOrderGeneralInformation.BackToList();

            purchaseOrdersPage.Filter(PurchaseOrdersPage.FilterType.ByNumber, ID);
            purchaseOrderItemPage = purchaseOrdersPage.SelectFirstItem();
            purchaseOrderItemPage.ClickOnGeneralInformation();

            var newDeliveryDate = puchaseOrderGeneralInformation.GetDeliveryDateValue();
            var newComment = puchaseOrderGeneralInformation.GetComment();
            var newStatus = puchaseOrderGeneralInformation.GetStatus();

            //Assert
            Assert.AreNotEqual(oldDeliveryDate, newDeliveryDate, "La delivery date n'a pas été modifiée.");
            Assert.AreNotEqual(oldComment, newComment, "Le commentaire n'a pas été modifié.");
            Assert.AreNotEqual(oldStatus, newStatus, "Le statut n'a pas été modifié.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_Item_UpdateQuantity()
        {
            //Prepare
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();

            string quantity = "5";

            //Arrange
            LogInAsAdmin();

            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();

            var createPurchaseOrderPage = purchaseOrdersPage.CreateNewPurchaseOrder();
            createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+1), true);
            var purchaseOrderItemPage = createPurchaseOrderPage.Submit();

            var oldQuantity = purchaseOrderItemPage.GetQuantity();
            purchaseOrderItemPage.SelectFirstItemPo();
            purchaseOrderItemPage.AddQuantity(quantity);
            purchaseOrderItemPage.Validate();

            var newQuantity = purchaseOrderItemPage.GetQuantity();

            //Assert
            Assert.AreNotEqual(oldQuantity, newQuantity, "La quantité de l'item n'a pas été modifiée.");
            Assert.AreEqual(quantity, newQuantity, "La nouvelle quantité de l'item n'est pas égale à celle demandée.");
        }
        [Timeout(_timeout)]
        //Ajouter un item dans le purchase order - Ajouter commentaire
        [TestMethod]
        public void PU_PO_Item_AddComment()
        {
            //Test de la création d'un nouveau purchase order
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();

            string comment = "test comment";

            //Arrange
            LogInAsAdmin();

            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            var createPurchaseOrderPage = purchaseOrdersPage.CreateNewPurchaseOrder();
            createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+1), true);

            var purchaseOrderItemPage = createPurchaseOrderPage.Submit();
            purchaseOrderItemPage.SelectFirstItemPo();
            string itemName = purchaseOrderItemPage.GetFirstItemName();

            purchaseOrderItemPage.AddQuantity("3");
            purchaseOrderItemPage.AddComment(itemName, comment);
            purchaseOrderItemPage.WaitPageLoading();
            //Assert
            var newComment = purchaseOrderItemPage.GetComment(itemName);
            Assert.AreEqual(newComment, comment, "Le commentaire n'a pas été modifié.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_Item_Filter_ByNameAndNotPurchased()
        {
            //Prepare
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();

            //Arrange
            LogInAsAdmin();

            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();

            var createPurchaseOrderPage = purchaseOrdersPage.CreateNewPurchaseOrder();
            createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+1), true);
            var purchaseOrderItemPage = createPurchaseOrderPage.Submit();

            var itemSearched = purchaseOrderItemPage.GetFirstItemName();
            purchaseOrderItemPage.Filter(FilterItemType.ByName, itemSearched);
            Assert.AreEqual(itemSearched, purchaseOrderItemPage.GetFirstItemName(), String.Format(MessageErreur.FILTRE_ERRONE, "Search item by name"));

            purchaseOrderItemPage.ResetFilter();
            purchaseOrderItemPage.SelectFirstItemPo();
            purchaseOrderItemPage.AddQuantity("5");
            purchaseOrderItemPage.Filter(FilterItemType.ShowItemsNotPurchased, false);

            Assert.IsTrue(purchaseOrderItemPage.VerifyPurchased(), String.Format(MessageErreur.FILTRE_ERRONE, "Show items purchased"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_Item_Filter_ByKeyword()
        {
            //Prepare
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();
            string itemKeyword = TestContext.Properties["Item_Keyword"].ToString();

            //Arrange
            LogInAsAdmin();

            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();

            var createPurchaseOrderPage = purchaseOrdersPage.CreateNewPurchaseOrder();
            createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+1), true);
            var purchaseOrderItemPage = createPurchaseOrderPage.Submit();

            // Ajout du keyword
            purchaseOrderItemPage.SelectFirstItemPo();
            string itemName = purchaseOrderItemPage.GetFirstItemName();

            var itemPageItem = purchaseOrderItemPage.EditItem(itemName);
            var itemKeywordTab = itemPageItem.ClickOnKeywordItem();
            itemKeywordTab.AddKeyword(itemKeyword);
            itemPageItem.Close();

            purchaseOrderItemPage.PageSize("8");
            purchaseOrderItemPage.Filter(FilterItemType.ByKeyword, itemKeyword);

            Assert.IsTrue(purchaseOrderItemPage.VerifyKeyword(itemKeyword), String.Format(MessageErreur.FILTRE_ERRONE, "'Keyword'"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_Item_Filter_ByGroup()
        {
            //Prepare
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();

            //Arrange
            LogInAsAdmin();

            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            purchaseOrdersPage.ResetFilters();
            var createPurchaseOrderPage = purchaseOrdersPage.CreateNewPurchaseOrder();
            createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+1), true);
            var purchaseOrderItemPage = createPurchaseOrderPage.Submit();

            // Get item Group
            var groupName = purchaseOrderItemPage.GetFirstGroupName();

            purchaseOrderItemPage.PageSize("8");
            purchaseOrderItemPage.Filter(FilterItemType.ByGroup, groupName);

            Assert.IsTrue(purchaseOrderItemPage.VerifyGroup(groupName), String.Format(MessageErreur.FILTRE_ERRONE, "'Group'"));
        }
        [Timeout(_timeout)]
        //Exporter un Purchase order - Supplier - New Version
        [TestMethod]
        public void PU_PO_Export_FilterSupplier_NewVersion()
        {
            //Test de l'export d'un purchase order
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();

            bool newVersion = true;

            //Prepare
            var homePage = LogInAsAdmin();

            // Act
            var purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();

            purchaseOrdersPage.ClearDownloads();

            // Filtre des données par supplier
            purchaseOrdersPage.ResetFilters();
            purchaseOrdersPage.Filter(PurchaseOrdersPage.FilterType.DateFrom, DateUtils.Now.AddDays(1));
            purchaseOrdersPage.Filter(PurchaseOrdersPage.FilterType.DateTo, DateUtils.Now.AddDays(2));
            purchaseOrdersPage.Filter(PurchaseOrdersPage.FilterType.Supplier, supplier);

            // Si on a déjà plus de 50 données, on n'en recréé pas de nouvelle pour evite d'avoir un trop grand nombre de données
            if (purchaseOrdersPage.CheckTotalNumber() < 20)
            {
                var createPurchaseOrderPage = purchaseOrdersPage.CreateNewPurchaseOrder();
                createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+1), true);
                var purchaseOrderItemPage = createPurchaseOrderPage.Submit();

                purchaseOrderItemPage.SelectFirstItemPo();
                purchaseOrderItemPage.AddQuantity("5");
                purchaseOrderItemPage.Validate();
                purchaseOrdersPage = purchaseOrderItemPage.BackToList();
            }

            // Export des données
            purchaseOrdersPage.Export(newVersion);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = purchaseOrdersPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber(PURCHASE_ORDERS_EXCEL_SHEET_NAME, filePath);
            bool result = OpenXmlExcel.ReadAllDataInColumn("Supplier", PURCHASE_ORDERS_EXCEL_SHEET_NAME, filePath, supplier);

            // On vide le répertoire de téléchargement
            DeleteAllFileDownload();

            //Assert
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.IsTrue(result, MessageErreur.EXCEL_DONNEES_KO);
        }
        [Timeout(_timeout)]
        //Exporter un Purchase order - Site - New Version
        [TestMethod]
        public void PU_PO_Export_FilterSite_NewVersion()
        {
            //Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();

            bool newVersion = true;
            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            var purchaseOrderPage = homePage.GoToPurchasing_PurchaseOrdersPage();

            purchaseOrderPage.ClearDownloads();

            // Filtre des données par site
            purchaseOrderPage.ResetFilters();
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.DateFrom, DateUtils.Now.AddDays(1));
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.DateTo, DateUtils.Now.AddDays(2));
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.Site, site + " - " + site);

            if (purchaseOrderPage.CheckTotalNumber() < 20)
            {
                var createPurchaseOrderPage = purchaseOrderPage.CreateNewPurchaseOrder();
                createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+1), true);
                var purchaseOrderItemPage = createPurchaseOrderPage.Submit();

                purchaseOrderItemPage.SelectFirstItemPo();
                purchaseOrderItemPage.AddQuantity("5");
                purchaseOrderItemPage.Validate();
                purchaseOrderPage = purchaseOrderItemPage.BackToList();
            }

            // Export
            purchaseOrderPage.Export(newVersion);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = purchaseOrderPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber(PURCHASE_ORDERS_EXCEL_SHEET_NAME, filePath);
            bool result = OpenXmlExcel.ReadAllDataInColumn("Site", PURCHASE_ORDERS_EXCEL_SHEET_NAME, filePath, site);

            //Assert
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.IsTrue(result, MessageErreur.EXCEL_DONNEES_KO);
        }
        [Timeout(_timeout)]
        //Imprimer la liste des Purchase orders - New Version
        [TestMethod]
        public void PU_PO_PrintResults_PurchaseOrders_NewVersion()
        {
            //Prepare
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNamePdfBegin = "Purchase order report_-_";
            string DocFileNameZipBegin = "All_files_";
            bool newVersion = true;
            var siteCity = "Madrid";

            var homePage = LogInAsAdmin();

            //Act
            var purchaseOrderPage = homePage.GoToPurchasing_PurchaseOrdersPage();

            purchaseOrderPage.ClearDownloads();

            // Filtre des données par période
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.DateFrom, DateUtils.Now.AddDays(1));
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.DateTo, DateUtils.Now.AddDays(2));

            if (purchaseOrderPage.CheckTotalNumber() < 8)
            {
                var createPurchaseOrderPage = purchaseOrderPage.CreateNewPurchaseOrder();
                createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+1), true);
                var purchaseOrderItemPage = createPurchaseOrderPage.Submit();

                purchaseOrderItemPage.SelectFirstItemPo();
                purchaseOrderItemPage.AddQuantity("5");
                purchaseOrderItemPage.Validate();
                purchaseOrderPage = purchaseOrderItemPage.BackToList();
            }

            // Print
            var reportPage = purchaseOrderPage.PrintResults(newVersion);
            var isReportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();

            //Assert
            Assert.IsTrue(isReportGenerated, "L'application n'a pas pu générer le fichier attendu.");

            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);

            // cliquer sur All
            string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            FileInfo fi = new FileInfo(trouve);
            fi.Refresh();
            Assert.IsTrue(fi.Exists, trouve + " non généré");

            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.DateFrom, DateUtils.Now.AddDays(1));
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.DateTo, DateUtils.Now.AddDays(2));

            List<string> purchaseOrdersList = purchaseOrderPage.GetTablePONumbers();

            PdfDocument document = PdfDocument.Open(fi.FullName);
            List<string> mots = new List<string>();
            foreach (Page p in document.GetPages())
            {
                mots.AddRange(p.GetWords().Select(m => m.Text));
            }
            foreach (var purchaseOrder in purchaseOrdersList)
            {
                if (!string.IsNullOrEmpty(purchaseOrder))
                {
                    Assert.AreNotEqual(0, mots.Count(w => w.Contains(purchaseOrder)), "Le PO n° {0} non présent dans le Pdf", purchaseOrder);
                    Assert.AreNotEqual(0, mots.Count(w => w.Contains(siteCity)), "Le site du PO n° {0} non présent dans le Pdf", purchaseOrder);
                    Assert.AreNotEqual(0, mots.Count(w => w.Contains(location)), "La delivery location du PO n° {0} non présent dans le Pdf", purchaseOrder);
                }
            }
        }
        [Timeout(_timeout)]
        //Imprimer un Purchase Order
        [TestMethod]
        public void PU_PO_PrintDetail_PurchaseOrder_NewVersion()
        {
            //Prepare
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();

            bool newVersion = true;

            var homePage = LogInAsAdmin();
            //Act
            var purchaseOrderPage = homePage.GoToPurchasing_PurchaseOrdersPage();

            purchaseOrderPage.ClearDownloads();

            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.ByValidationDate, true);

            if (purchaseOrderPage.CheckTotalNumber() == 0)
            {
                var createPurchaseOrderPage = purchaseOrderPage.CreateNewPurchaseOrder();
                createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+1), true);
                var purchaseOrderItemPage = createPurchaseOrderPage.Submit();

                purchaseOrderItemPage.SelectFirstItemPo();
                purchaseOrderItemPage.AddQuantity("5");
                purchaseOrderItemPage.Validate();
                purchaseOrderPage = purchaseOrderItemPage.BackToList();
            }

            var purchaseOrderDetailPage = purchaseOrderPage.SelectFirstItem();

            // Impression du contenu du purchase order validé
            var reportPage = purchaseOrderDetailPage.PrintDetails(newVersion);
            var isReportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();

            Assert.IsTrue(isReportGenerated, "L'application n'a pas pu générer le fichier attendu.");
        }

       [Ignore] //Export WMS sur des Purchase order - New Version
        [TestMethod]
        public void PU_PO_WMSExport_PurchaseOrders_NewVersion()
        {
            //Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();

            bool newVersion = true;
            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            PurchaseOrdersPage purchaseOrderPage = homePage.GoToPurchasing_PurchaseOrdersPage();

            // Filtre des données par site
            purchaseOrderPage.ResetFilters();
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.DateFrom, DateUtils.Now.AddDays(1));
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.DateTo, DateUtils.Now.AddDays(2));
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.Site, site);

            if (purchaseOrderPage.CheckTotalNumber() < 20)
            {
                var createPurchaseOrderPage = purchaseOrderPage.CreateNewPurchaseOrder();
                createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+1), true);
                var purchaseOrderItemPage = createPurchaseOrderPage.Submit();

                purchaseOrderItemPage.SelectFirstItemPo();
                purchaseOrderItemPage.AddQuantity("5");
                purchaseOrderItemPage.Validate();
                purchaseOrderPage = purchaseOrderItemPage.BackToList();
            }

            // Export
            purchaseOrderPage.ExportWMS(newVersion);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = purchaseOrderPage.GetExportWMSFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);
        }
        [Timeout(_timeout)]
        //Envoyer des Purchase Orders par mail
        [TestMethod]
        public void PU_PO_SendPurchaseOrdersByMail()
        {
            //Prepare
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();
            //var contactEmail = "test@mail.com";

           HomePage homePage = LogInAsAdmin();      

            //Act
            var purchaseOrderPage = homePage.GoToPurchasing_PurchaseOrdersPage();

            var createPurchaseOrderPage = purchaseOrderPage.CreateNewPurchaseOrder();
            createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+1), true);
            var purchaseOrderItemPage = createPurchaseOrderPage.Submit();
            var ID = purchaseOrderItemPage.GetPurchaseOrderNumber();

            purchaseOrderItemPage.SelectFirstItemPo();
            purchaseOrderItemPage.AddQuantity("5");
            purchaseOrderItemPage.Validate();
            purchaseOrderPage = purchaseOrderItemPage.BackToList();

            purchaseOrderPage.ResetFilters();
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.ByNumber, ID);
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.DateFrom, DateUtils.Now);
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.DateTo, DateUtils.Now.AddMonths(1).AddDays(-1));

            // Envoi du purchase order créé par mail
            purchaseOrderPage.SendResultsByMail();

            purchaseOrderPage.ResetFilters();
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.ByNumber, ID);

            Assert.IsTrue(purchaseOrderPage.IsSentByMail(), "Le purchase order ne possède pas l'icône signifiant qu'il a été envoyé par mail.");

            // check if mail received
            var email = TestContext.Properties["Admin_UserName"].ToString();
            MailPage mailPage = purchaseOrderPage.RedirectToOutlookMailbox();
            mailPage.FillFields_LogToOutlookMailbox(email);

            mailPage.ClickOnSpecifiedOutlookMail("OC NEWREST - 1 - Purchase order -");

            // Vérifier que la pièce jointe est présente et qu'elle n'est pas vide
            bool isPieceJointeOK = mailPage.IsOutlookPieceJointeOK("Purchase order_" + ID + " " + site + ".pdf");

            Assert.IsTrue(isPieceJointeOK, "La pièce jointe n'est pas présente dans le mail.");

            mailPage.DeleteCurrentOutlookMail();

            mailPage.Close();
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_Filter_SearchByNumber()
        {
            //Prepare
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();
           
            //Arrange
            HomePage homePage = LogInAsAdmin();
            //Act
            var purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            var createPurchaseOrderPage = purchaseOrdersPage.CreateNewPurchaseOrder();
            createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+1), true);           
            var purchaseOrderItemPage = createPurchaseOrderPage.Submit();
            var ID = purchaseOrderItemPage.GetPurchaseOrderNumber();
            purchaseOrdersPage = purchaseOrderItemPage.BackToList();
            //Assert
            purchaseOrdersPage.ResetFilters();
            purchaseOrdersPage.Filter(PurchaseOrdersPage.FilterType.ByNumber, ID);
            Assert.AreEqual(ID, purchaseOrdersPage.GetFirstPurchaseOrderNumber(), String.Format(MessageErreur.FILTRE_ERRONE, "Search by number"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_Filter_SearchBySupplyOrderNumber()
        {
          
            //Arrange
             HomePage homePage = LogInAsAdmin();
            var supplyOrderPage = homePage.GoToPurchasing_SupplyOrderPage();
            supplyOrderPage.ResetFilter();
            supplyOrderPage.Filter(FilterType.ShowNotValidated, false);
            string supplyOrderNumber = supplyOrderPage.GetFirstSONumber();
            SupplyOrderItem soItem = supplyOrderPage.SelectFirstItem();

            //1) Etre sur une SO avec item/ quantité validée
            string itemName = soItem.GetFirstItemNameValidated();
            string itemQty = soItem.GetFirstItemQtyValidated();
            Assert.IsTrue(int.Parse(itemQty) > 0, "pas de quantité validée");

            //2) ... puis Generate Purchase Order
            //3) Choisir le supplier et la delivery date
            PurchaseOrdersPage purchaseOrderPage = soItem.GeneratePurchaseOrder();
            purchaseOrderPage.ResetFilters();
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.BySupplierOrderNumber, supplyOrderNumber);
            Assert.IsTrue(purchaseOrderPage.CheckTotalNumber()>0, String.Format(MessageErreur.FILTRE_ERRONE, "By Supplier Order Number"));







            if (!purchaseOrderPage.isPageSizeEqualsTo100())
            {
                purchaseOrderPage.PageSize("8");
                purchaseOrderPage.PageSize("100");
            }
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.BySupplierOrderNumber, supplyOrderNumber);
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.FilterShow, "Show validated only");
            var nbr1 = purchaseOrderPage.CheckTotalNumber();

            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.FilterShow, "Show not validated only");
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.BySupplierOrderNumber, supplyOrderNumber);

            var nbr2 = purchaseOrderPage.CheckTotalNumber();
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.BySupplierOrderNumber, supplyOrderNumber);
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.FilterShow, "Show all orders");
            var realNbr = purchaseOrderPage.CheckTotalNumber();

            Assert.AreEqual(nbr1 + nbr2, realNbr, String.Format(MessageErreur.FILTRE_ERRONE, "Show all orders"));

        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_Filter_BySupplier()
        {
            //Prepare
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();
            HomePage homePage = LogInAsAdmin();
            var purchaseOrderPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            purchaseOrderPage.ResetFilters();
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.Supplier, supplier);
            if (purchaseOrderPage.CheckTotalNumber() < 20)
            {
                var createPurchaseOrderPage = purchaseOrderPage.CreateNewPurchaseOrder();
                createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+1), true);
                var purchaseOrderItemPage = createPurchaseOrderPage.Submit();

                purchaseOrderItemPage.SelectFirstItemPo();
                purchaseOrderItemPage.AddQuantity("5");
                purchaseOrderItemPage.Validate();
                purchaseOrderPage = purchaseOrderItemPage.BackToList();
                purchaseOrderPage.ResetFilters();
                purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.Supplier, supplier);
            }

            if (!purchaseOrderPage.isPageSizeEqualsTo100())
            {
                purchaseOrderPage.PageSize("8");
                purchaseOrderPage.PageSize("100");
            }
            Assert.IsTrue(purchaseOrderPage.VerifySupplier(supplier), MessageErreur.FILTRE_ERRONE, "Supplier");

        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_Filter_Date()
        {
            //Prepare
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();
            //Arrange
            HomePage homePage = LogInAsAdmin();
            //Act
            var purchaseOrderPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            purchaseOrderPage.ResetFilters();
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.DateFrom, DateUtils.Now.AddDays(-1));
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.DateTo, DateUtils.Now.AddDays(+20));

            // Si on a déjà plus de 50 données, on n'en recréé pas de nouvelle pour evite d'avoir un trop grand nombre de données
            if (purchaseOrderPage.CheckTotalNumber() < 20)
            {
                var createPurchaseOrderPage = purchaseOrderPage.CreateNewPurchaseOrder();
                createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+1), true);
                var purchaseOrderItemPage = createPurchaseOrderPage.Submit();

                purchaseOrderItemPage.SelectFirstItemPo();
                purchaseOrderItemPage.AddQuantity("5");
                purchaseOrderItemPage.Validate();
                purchaseOrderPage = purchaseOrderItemPage.BackToList();
            }

            if (!purchaseOrderPage.isPageSizeEqualsTo100())
            {
                purchaseOrderPage.PageSize("8");
                purchaseOrderPage.PageSize("100");
            }

            // Assert
            Assert.IsTrue(purchaseOrderPage.IsDateRespected(DateUtils.Now.AddDays(-1).Date, DateUtils.Now.AddDays(+20).Date), String.Format(MessageErreur.FILTRE_ERRONE, "From/To"));
        }
        //Effectuer des recherches via les filtres - By Validation Date

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_Filter_By_Validation_Date()
        {
            //Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();

            bool newVersion = true;
            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var purchaseOrderPage = homePage.GoToPurchasing_PurchaseOrdersPage();

            purchaseOrderPage.ClearDownloads();

            purchaseOrderPage.ResetFilters();
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.DateFrom, DateUtils.Now.AddDays(-1));
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.DateTo, DateUtils.Now.AddDays(+20));
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.ByValidationDate, true);

            // Vérifier si la création de nouvelle commande est nécessaire
            if (purchaseOrderPage.CheckTotalNumber() < 20)
            {
                var createPurchaseOrderPage = purchaseOrderPage.CreateNewPurchaseOrder();
                createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+1), true);

                var purchaseOrderItemPage = createPurchaseOrderPage.Submit();
                purchaseOrderItemPage.SelectFirstItemPo();
                purchaseOrderItemPage.AddQuantity("5");
                purchaseOrderItemPage.Validate();

                purchaseOrderPage = purchaseOrderItemPage.BackToList();
            }

            // Exportation des résultats sous la forme d'un fichier Excel
            purchaseOrderPage.Export(newVersion);

            // Récupération du fichier téléchargé dans le répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();
            var correctDownloadedFile = purchaseOrderPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber(PURCHASE_ORDERS_EXCEL_SHEET_NAME, filePath);
            bool result = OpenXmlExcel.ReadAllDataInColumn("Is Valid", PURCHASE_ORDERS_EXCEL_SHEET_NAME, filePath, "True");

            // Assert : Vérification des résultats
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.IsTrue(result, MessageErreur.EXCEL_DONNEES_KO);
        }
    
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_Filter_ShowAllOrders()
        {
            //Arrange
            HomePage homePage = LogInAsAdmin();
            //Act
            var purchaseOrderPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            purchaseOrderPage.ResetFilters();
            if (!purchaseOrderPage.isPageSizeEqualsTo100())
            {
                purchaseOrderPage.PageSize("8");
                purchaseOrderPage.PageSize("100");
            }
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.FilterShow, "Show validated only");
            var nbr1 = purchaseOrderPage.CheckTotalNumber();
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.FilterShow, "Show not validated only");
            var nbr2 = purchaseOrderPage.CheckTotalNumber();
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.FilterShow, "Show all orders");
            var realNbr = purchaseOrderPage.CheckTotalNumber();
            //Assert
            Assert.AreEqual(nbr1 + nbr2, realNbr, String.Format(MessageErreur.FILTRE_ERRONE, "Show all orders"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_Filter_ShowNotValidatedOnly()
        {
            //Prepare
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var purchaseOrderPage = homePage.GoToPurchasing_PurchaseOrdersPage();

            purchaseOrderPage.ResetFilters();
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.FilterShow, "Show not validated only");

            if (purchaseOrderPage.CheckTotalNumber() < 20)
            {
                var createPurchaseOrderPage = purchaseOrderPage.CreateNewPurchaseOrder();
                createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+1), true);
                var purchaseOrderItemPage = createPurchaseOrderPage.Submit();
                purchaseOrderPage = purchaseOrderItemPage.BackToList();
            }

            if (!purchaseOrderPage.isPageSizeEqualsTo100())
            {
                purchaseOrderPage.PageSize("8");
                purchaseOrderPage.PageSize("100");
            }

            //Assert
            Assert.IsFalse(purchaseOrderPage.CheckValidation(false), String.Format(MessageErreur.FILTRE_ERRONE, "Show not validated only"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_Filter_ShowValidatedOnly()
        {
            //Prepare
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var purchaseOrderPage = homePage.GoToPurchasing_PurchaseOrdersPage();

            purchaseOrderPage.ResetFilters();
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.FilterShow, "Show validated only");

            if (purchaseOrderPage.CheckTotalNumber() < 20)
            {
                var createPurchaseOrderPage = purchaseOrderPage.CreateNewPurchaseOrder();
                createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+1), true);
                var purchaseOrderItemPage = createPurchaseOrderPage.Submit();
                purchaseOrderItemPage.SelectFirstItemPo();
                purchaseOrderItemPage.AddQuantity("5");
                purchaseOrderItemPage.Validate();
                purchaseOrderItemPage.BackToList();
            }

            if (!purchaseOrderPage.isPageSizeEqualsTo100())
            {
                purchaseOrderPage.PageSize("8");
                purchaseOrderPage.PageSize("100");
            }

            bool check = purchaseOrderPage.CheckValidation(true);
            //Assert
            Assert.IsTrue(check, String.Format(MessageErreur.FILTRE_ERRONE, "Show validated only"));
        }

        [Ignore]//Effectuer des recherches via les filtres - Show EDI Sent Only 16/12/2021 Ignore car manque test création EDI
        [TestMethod]
        public void PU_PO_Filter_ShowEdiSentOnly()
        {
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var purchaseOrderPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            purchaseOrderPage.ResetFilters();

            if (!purchaseOrderPage.isPageSizeEqualsTo100())
            {
                purchaseOrderPage.PageSize("8");
                purchaseOrderPage.PageSize("100");
            }

            purchaseOrderPage.ClearDateFrom();
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.FilterShow, "Show EDI sent only");

            //Assert
            Assert.IsTrue(purchaseOrderPage.IsSentByEDI(), String.Format(MessageErreur.FILTRE_ERRONE, "Show EDI sent only"));
        }

        [Ignore]//Effectuer des recherches via les filtres - Show EDI and Email Sent Only 16/12/2021 Ignore car manque test création EDI
        [TestMethod]
        public void PU_PO_Filter_ShowEdiAndEmailSentOnly()
        {
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var purchaseOrderPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            purchaseOrderPage.ResetFilters();

            if (!purchaseOrderPage.isPageSizeEqualsTo100())
            {
                purchaseOrderPage.PageSize("8");
                purchaseOrderPage.PageSize("100");
            }

            purchaseOrderPage.ClearDateFrom();
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.FilterShow, "Show EDI and Email sent only");

            //Assert
            Assert.IsTrue(purchaseOrderPage.IsSentByEDI(), String.Format(MessageErreur.FILTRE_ERRONE, "Show EDI and Email sent only"));
            Assert.IsTrue(purchaseOrderPage.IsSentByMail(), String.Format(MessageErreur.FILTRE_ERRONE, "Show EDI and Email sent only"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_Filter_ShowEmailSentOnly()
        {
            //Prepare
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();
            var email = TestContext.Properties["Admin_UserName"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var purchaseOrderPage = homePage.GoToPurchasing_PurchaseOrdersPage();

            purchaseOrderPage.ResetFilters();
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.FilterShow, "Show Email sent only");

            if (purchaseOrderPage.CheckTotalNumber() < 20)
            {
                var createPurchaseOrderPage = purchaseOrderPage.CreateNewPurchaseOrder();
                createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+1), true);
                var purchaseOrderItemPage = createPurchaseOrderPage.Submit();
                purchaseOrderItemPage.SelectFirstItemPo();
                purchaseOrderItemPage.AddQuantity("5");
                purchaseOrderItemPage.Validate();
                purchaseOrderItemPage.SendPODetailsByMail(email);

                purchaseOrderItemPage.BackToList();
            }

            if (!purchaseOrderPage.isPageSizeEqualsTo100())
            {
                purchaseOrderPage.PageSize("8");
                purchaseOrderPage.PageSize("100");
            }

            //Assert
            Assert.IsTrue(purchaseOrderPage.IsSentByMail(), String.Format(MessageErreur.FILTRE_ERRONE, "Show Email sent only"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_Filter_ShowNotSentByMailOnly()
        {
            //Prepare
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();

            //Arrange
           

           HomePage homePage = LogInAsAdmin();

            //Act
            var purchaseOrderPage = homePage.GoToPurchasing_PurchaseOrdersPage();

            purchaseOrderPage.ResetFilters();
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.FilterShow, "Show not sent by mail only");

            if (purchaseOrderPage.CheckTotalNumber() < 20)
            {
                var createPurchaseOrderPage = purchaseOrderPage.CreateNewPurchaseOrder();
                createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+1), true);
                var purchaseOrderItemPage = createPurchaseOrderPage.Submit();
                purchaseOrderItemPage.SelectFirstItemPo();
                purchaseOrderItemPage.AddQuantity("5");
                purchaseOrderItemPage.Validate();
                purchaseOrderItemPage.BackToList();
            }

            if (!purchaseOrderPage.isPageSizeEqualsTo100())
            {
                purchaseOrderPage.PageSize("8");
                purchaseOrderPage.PageSize("100");
            }

            //Assert
            Assert.IsFalse(purchaseOrderPage.IsSentByMail(), String.Format(MessageErreur.FILTRE_ERRONE, "Show not sent by mail only"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_Filter_ShowNotReceivedFilter()
        {
            //Prepare
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();

            //Arrange

            HomePage homePage = LogInAsAdmin();

            //Act
            var purchaseOrderPage = homePage.GoToPurchasing_PurchaseOrdersPage();

            purchaseOrderPage.ResetFilters();
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.FilterShow, "Show not received only");

            if (purchaseOrderPage.CheckTotalNumber() < 20)
            {
                var createPurchaseOrderPage = purchaseOrderPage.CreateNewPurchaseOrder();
                createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+1), true);
                var purchaseOrderItemPage = createPurchaseOrderPage.Submit();
                purchaseOrderItemPage.SelectFirstItemPo();
                purchaseOrderItemPage.AddQuantity("5");
                purchaseOrderItemPage.Validate();
                purchaseOrderItemPage.BackToList();
            }

            if (!purchaseOrderPage.isPageSizeEqualsTo100())
            {
                purchaseOrderPage.PageSize("8");
                purchaseOrderPage.PageSize("100");
            }
            bool check = purchaseOrderPage.IsDelivered();
            //Assert
            Assert.IsFalse(check, String.Format(MessageErreur.FILTRE_ERRONE, "Show not received only"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_Filter_ShowAll()
        {
            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var purchaseOrderPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            purchaseOrderPage.ResetFilters();

            if (!purchaseOrderPage.isPageSizeEqualsTo100())
            {
                purchaseOrderPage.PageSize("8");
                purchaseOrderPage.PageSize("100");
            }

            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.ShowActive, true);
            var nbr1 = purchaseOrderPage.CheckTotalNumber();

            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.ShowInactive, true);
            var nbr2 = purchaseOrderPage.CheckTotalNumber();

            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.ShowAll, true);
            var realNbr = purchaseOrderPage.CheckTotalNumber();

            //Assert
            Assert.AreEqual(nbr1 + nbr2, realNbr, String.Format(MessageErreur.FILTRE_ERRONE, "Show all"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_Filter_ShowActiveOnly()
        {
            //Prepare
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var purchaseOrderPage = homePage.GoToPurchasing_PurchaseOrdersPage();

            purchaseOrderPage.ResetFilters();
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.ShowActive, true);

            if (purchaseOrderPage.CheckTotalNumber() < 20)
            {
                var createPurchaseOrderPage = purchaseOrderPage.CreateNewPurchaseOrder();
                createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+1), true);
                var purchaseOrderItemPage = createPurchaseOrderPage.Submit();
                purchaseOrderItemPage.BackToList();
            }
            if (!purchaseOrderPage.isPageSizeEqualsTo100())
            {
                purchaseOrderPage.PageSize("8");
                purchaseOrderPage.PageSize("100");
            }
            //Assert
            Assert.IsTrue(purchaseOrderPage.CheckStatus(true), String.Format(MessageErreur.FILTRE_ERRONE, "Show only active"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_Filter_ShowInactiveOnly()
        {
            //Prepare
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var purchaseOrderPage = homePage.GoToPurchasing_PurchaseOrdersPage();

            purchaseOrderPage.ResetFilters();
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.ShowInactive, true);

            if (purchaseOrderPage.CheckTotalNumber() < 20)
            {
                var createPurchaseOrderPage = purchaseOrderPage.CreateNewPurchaseOrder();
                createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+1), false);
                var purchaseOrderItemPage = createPurchaseOrderPage.Submit();
                purchaseOrderItemPage.BackToList();
            }

            if (!purchaseOrderPage.isPageSizeEqualsTo100())
            {
                purchaseOrderPage.PageSize("8");
                purchaseOrderPage.PageSize("100");
            }

            //Assert
            Assert.IsFalse(purchaseOrderPage.CheckStatus(false), String.Format(MessageErreur.FILTRE_ERRONE, "Show only inactive"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_Filter_Receipt_Status_Delivered()
        {
            //Prepare
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();
            //Act
            var purchaseOrderPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            purchaseOrderPage.ResetFilters();
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.ReceiptStatus, "2");

            if (purchaseOrderPage.CheckTotalNumber() < 20)
            {
                var createPurchaseOrderPage = purchaseOrderPage.CreateNewPurchaseOrder();
                createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+1), true);
                var purchaseOrderItemPage = createPurchaseOrderPage.Submit();
                purchaseOrderItemPage.SelectFirstItemPo();
                purchaseOrderItemPage.AddQuantity("5");
                purchaseOrderItemPage.Validate();

                purchaseOrderItemPage.GenerateReceiptNote(false);
                var receiptNoteItemPage = purchaseOrderItemPage.ValidateReceiptNoteCreation();

                var receiptNoteGeneralInformationPage = receiptNoteItemPage.ClickOnGeneralInformationTab();
                purchaseOrderItemPage = receiptNoteGeneralInformationPage.ClickOnPurchaseOrderLink();

                purchaseOrderItemPage.ClickOnGeneralInformation();
                purchaseOrderItemPage.ChangeReceiptStatus("Delivered");
                purchaseOrderItemPage.BackToList();
                purchaseOrderPage.ResetFilters();
                purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.ReceiptStatus, "2");
            }

            if (!purchaseOrderPage.isPageSizeEqualsTo100())
            {
                purchaseOrderPage.PageSize("8");
                purchaseOrderPage.PageSize("100");
            }

            //Assert
            Assert.IsTrue(purchaseOrderPage.IsDelivered(), String.Format(MessageErreur.FILTRE_ERRONE, "Receipt status : Delivered"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_Filter_Receipt_Status_PartiallyDelivered()
        {
            //Prepare
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var purchaseOrderPage = homePage.GoToPurchasing_PurchaseOrdersPage();

            purchaseOrderPage.ResetFilters();
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.ReceiptStatus, "1");

            if (purchaseOrderPage.CheckTotalNumber() < 20)
            {
                var createPurchaseOrderPage = purchaseOrderPage.CreateNewPurchaseOrder();
                createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+1), true);
                var purchaseOrderItemPage = createPurchaseOrderPage.Submit();
                purchaseOrderItemPage.SelectFirstItemPo();
                purchaseOrderItemPage.AddQuantity("5");
                purchaseOrderItemPage.Validate();

                purchaseOrderItemPage.GenerateReceiptNote(false);
                var receiptNoteItemPage = purchaseOrderItemPage.ValidateReceiptNoteCreation();

                var receiptNoteGeneralInformationPage = receiptNoteItemPage.ClickOnGeneralInformationTab();
                purchaseOrderItemPage = receiptNoteGeneralInformationPage.ClickOnPurchaseOrderLink();

                purchaseOrderItemPage.ClickOnGeneralInformation();
                purchaseOrderItemPage.ChangeReceiptStatus("PartiallyDelivered");
                purchaseOrderItemPage.BackToList();

                purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.ReceiptStatus, "1");
            }

            if (!purchaseOrderPage.isPageSizeEqualsTo100())
            {
                purchaseOrderPage.PageSize("8");
                purchaseOrderPage.PageSize("100");
            }

            //Assert
            Assert.IsTrue(purchaseOrderPage.IsPartiallyDelivered(), String.Format(MessageErreur.FILTRE_ERRONE, "Receipt status : PartiallyDelivered"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_Filter_Receipt_Status_Unknown()
        {
            //Prepare
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var purchaseOrderPage = homePage.GoToPurchasing_PurchaseOrdersPage();

            purchaseOrderPage.ResetFilters();
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.ReceiptStatus, "0");

            if (purchaseOrderPage.CheckTotalNumber() < 20)
            {
                var createPurchaseOrderPage = purchaseOrderPage.CreateNewPurchaseOrder();
                createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+1), true);
                var purchaseOrderItemPage = createPurchaseOrderPage.Submit();
                purchaseOrderItemPage.SelectFirstItemPo();
                purchaseOrderItemPage.AddQuantity("5");
                purchaseOrderItemPage.Validate();

                purchaseOrderItemPage.GenerateReceiptNote(false);
                var receiptNoteItemPage = purchaseOrderItemPage.ValidateReceiptNoteCreation();

                var receiptNoteGeneralInformationPage = receiptNoteItemPage.ClickOnGeneralInformationTab();
                purchaseOrderItemPage = receiptNoteGeneralInformationPage.ClickOnPurchaseOrderLink();

                purchaseOrderItemPage.ClickOnGeneralInformation();
                purchaseOrderItemPage.ChangeReceiptStatus("Unknown");
                purchaseOrderItemPage.BackToList();

                purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.ReceiptStatus, "0");
            }

            if (!purchaseOrderPage.isPageSizeEqualsTo100())
            {
                purchaseOrderPage.PageSize("8");
                purchaseOrderPage.PageSize("100");
            }

            //Assert
            Assert.IsFalse(purchaseOrderPage.IsDelivered(), String.Format(MessageErreur.FILTRE_ERRONE, "Receipt status : Unknown"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_Filter_Receipt_Status_All()
        {
            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var purchaseOrderPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            purchaseOrderPage.ResetFilters();

            if (!purchaseOrderPage.isPageSizeEqualsTo100())
            {
                purchaseOrderPage.PageSize("8");
                purchaseOrderPage.PageSize("100");
            }

            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.ReceiptStatus, "2");
            var nbr1 = purchaseOrderPage.CheckTotalNumber();

            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.ReceiptStatus, "1");
            var nbr2 = purchaseOrderPage.CheckTotalNumber();

            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.ReceiptStatus, "0");
            var nbr3 = purchaseOrderPage.CheckTotalNumber();

            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.ReceiptStatus, "-1");
            var realNbr = purchaseOrderPage.CheckTotalNumber();

            //Assert
            Assert.AreEqual(nbr1 + nbr2 + nbr3, realNbr, String.Format(MessageErreur.FILTRE_ERRONE, "Receipt status : All"));
        }
        [Timeout(_timeout)]
        //Effectuer des recherches via les filtres - SortBy
        [TestMethod]
        public void PU_PO_Filter_SortByNumber()
        {
            //Prepare
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var purchaseOrderPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            purchaseOrderPage.ResetFilters();

            if (purchaseOrderPage.CheckTotalNumber() < 20)
            {
                var createPurchaseOrderPage = purchaseOrderPage.CreateNewPurchaseOrder();
                createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+1), true);
                var purchaseOrderItemPage = createPurchaseOrderPage.Submit();
                purchaseOrderItemPage.BackToList();
            }

            if (!purchaseOrderPage.isPageSizeEqualsTo100())
            {
                purchaseOrderPage.PageSize("8");
                purchaseOrderPage.PageSize("100");
            }

            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.SortBy, "Number");
            var isSortedByNumber = purchaseOrderPage.IsSortedByNumber();

            //Assert
            Assert.IsTrue(isSortedByNumber, String.Format(MessageErreur.FILTRE_ERRONE, "Sort by Number"));
        }

        [Timeout(_timeout)]
        //Effectuer des recherches via les filtres - SortBy
        [TestMethod]
        public void PU_PO_Filter_SortByDeliveryDate()
        {
            //Prepare
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var purchaseOrderPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            purchaseOrderPage.ResetFilters();

            if (purchaseOrderPage.CheckTotalNumber() < 20)
            {
                var createPurchaseOrderPage = purchaseOrderPage.CreateNewPurchaseOrder();
                createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+1), true);
                var purchaseOrderItemPage = createPurchaseOrderPage.Submit();
                purchaseOrderItemPage.BackToList();
            }

            if (!purchaseOrderPage.isPageSizeEqualsTo100())
            {
                purchaseOrderPage.PageSize("8");
                purchaseOrderPage.PageSize("100");
            }

            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.SortBy, "Delivery date");
            var isSortedByDate = purchaseOrderPage.IsSortedByDate();

            //Assert
            Assert.IsTrue(isSortedByDate, String.Format(MessageErreur.FILTRE_ERRONE, "Sort by Delivery date"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_Filter_Opened()
        {
            //Prepare
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            bool newVersion = true;
            //Arrange
            HomePage homePage = LogInAsAdmin();
            //Act
            var purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            purchaseOrdersPage.ResetFilters();
            purchaseOrdersPage.ClearDownloads();
            purchaseOrdersPage.Filter(PurchaseOrdersPage.FilterType.DateFrom, DateUtils.Now.AddDays(-1));
            purchaseOrdersPage.Filter(PurchaseOrdersPage.FilterType.DateTo, DateUtils.Now.AddDays(+20));
            purchaseOrdersPage.Filter(PurchaseOrdersPage.FilterType.Opened, true);
            if (purchaseOrdersPage.CheckTotalNumber() < 20)
            {
                var createPurchaseOrderPage = purchaseOrdersPage.CreateNewPurchaseOrder();
                createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+1), true);
                var purchaseOrderItemPage = createPurchaseOrderPage.Submit();

                purchaseOrderItemPage.SelectFirstItemPo();
                purchaseOrderItemPage.AddQuantity("5");
                purchaseOrderItemPage.ChangeStatus("Open");
                purchaseOrdersPage = purchaseOrderItemPage.BackToList();
            }
            // Export
            purchaseOrdersPage.Export(newVersion);
            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();
            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = purchaseOrdersPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);
            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber(PURCHASE_ORDERS_EXCEL_SHEET_NAME, filePath);
            bool result = OpenXmlExcel.ReadAllDataInColumn("Status", PURCHASE_ORDERS_EXCEL_SHEET_NAME, filePath, "OPENED");
            //Assert
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.IsTrue(result, MessageErreur.EXCEL_DONNEES_KO);
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_Filter_Closed()
        {
            //Prepare
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            bool newVersion = true;

            //Arrange

            HomePage homePage = LogInAsAdmin();

            //Act
            var purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();

            purchaseOrdersPage.ResetFilters();

            purchaseOrdersPage.ClearDownloads();

            purchaseOrdersPage.Filter(PurchaseOrdersPage.FilterType.DateFrom, DateUtils.Now.AddDays(-1));
            purchaseOrdersPage.Filter(PurchaseOrdersPage.FilterType.DateTo, DateUtils.Now.AddDays(+20));
            purchaseOrdersPage.Filter(PurchaseOrdersPage.FilterType.Closed, true);

            if (purchaseOrdersPage.CheckTotalNumber() < 20)
            {
                var createPurchaseOrderPage = purchaseOrdersPage.CreateNewPurchaseOrder();
                createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+1), true);
                var purchaseOrderItemPage = createPurchaseOrderPage.Submit();

                purchaseOrderItemPage.SelectFirstItemPo();
                purchaseOrderItemPage.AddQuantity("5");
                purchaseOrderItemPage.ClickOnGeneralInformation();
                purchaseOrderItemPage.ChangeStatus("Close");
                purchaseOrdersPage = purchaseOrderItemPage.BackToList();
            }


            // Export
            purchaseOrdersPage.Export(newVersion);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = purchaseOrdersPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber(PURCHASE_ORDERS_EXCEL_SHEET_NAME, filePath);
            bool result = OpenXmlExcel.ReadAllDataInColumn("Status", PURCHASE_ORDERS_EXCEL_SHEET_NAME, filePath, "CLOSED");

            //Assert
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.IsTrue(result, MessageErreur.EXCEL_DONNEES_KO);
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_Filter_Cancelled()
        {
            //Prepare
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            bool newVersion = true;
            //Arrange
             HomePage homePage = LogInAsAdmin();
            //Act
            var purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            purchaseOrdersPage.ResetFilters();
            purchaseOrdersPage.ClearDownloads();
            purchaseOrdersPage.Filter(PurchaseOrdersPage.FilterType.DateFrom, DateUtils.Now.AddDays(-1));
            purchaseOrdersPage.Filter(PurchaseOrdersPage.FilterType.DateTo, DateUtils.Now.AddDays(+20));
            purchaseOrdersPage.Filter(PurchaseOrdersPage.FilterType.Cancelled, true);
            
            var createPurchaseOrderPage = purchaseOrdersPage.CreateNewPurchaseOrder();
            createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+1), true);
            var purchaseOrderItemPage = createPurchaseOrderPage.Submit();
            var numberPurchaseOrder = purchaseOrderItemPage.GetPurchaseOrderNumber();
            purchaseOrderItemPage.SelectFirstItemPo();
            purchaseOrderItemPage.GetPurchaseOrderNumber();
            purchaseOrderItemPage.AddQuantity("5");
            purchaseOrderItemPage.ClickOnGeneralInformation();
            purchaseOrderItemPage.ChangeStatus("Cancelled");
             purchaseOrdersPage = purchaseOrderItemPage.BackToList();
            try
            {
                // Export
                purchaseOrdersPage.Export(newVersion);
                // On récupère les fichiers du répertoire de téléchargement
                DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
                FileInfo[] taskFiles = taskDirectory.GetFiles();
                // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
                var correctDownloadedFile = purchaseOrdersPage.GetExportExcelFile(taskFiles);
                Assert.IsNotNull(correctDownloadedFile);
                var fileName = correctDownloadedFile.Name;
                var filePath = Path.Combine(downloadsPath, fileName);
                // Exploitation du fichier Excel
                int resultNumber = OpenXmlExcel.GetExportResultNumber(PURCHASE_ORDERS_EXCEL_SHEET_NAME, filePath);
                bool result = OpenXmlExcel.ReadAllDataInColumn("Status", PURCHASE_ORDERS_EXCEL_SHEET_NAME, filePath, "CANCELLED");

                //Assert
                Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
                Assert.IsTrue(result, MessageErreur.EXCEL_DONNEES_KO);
            }
            finally
            {
                purchaseOrdersPage.Filter(PurchaseOrdersPage.FilterType.ByNumber, numberPurchaseOrder);
                purchaseOrdersPage.DeleteFirstPurchaseOrder();
            }
            
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_Filter_StatusAll()
        {
            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var purchaseOrderPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            purchaseOrderPage.ResetFilters();

            if (!purchaseOrderPage.isPageSizeEqualsTo100())
            {
                purchaseOrderPage.PageSize("8");
                purchaseOrderPage.PageSize("100");
            }

            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.Opened, true);
            var numberOpened = purchaseOrderPage.CheckTotalNumber();

            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.Closed, true);
            var numberClosed = purchaseOrderPage.CheckTotalNumber();

            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.Cancelled, true);
            var numberCancelled = purchaseOrderPage.CheckTotalNumber();

            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.StatusAll, true);
            var totalNumber = purchaseOrderPage.CheckTotalNumber();

            //Assert
            Assert.AreEqual(numberOpened + numberClosed + numberCancelled, totalNumber, String.Format(MessageErreur.FILTRE_ERRONE, "Status : All"));
        }
        [Timeout(_timeout)]
        //Effectuer des recherches via les filtres - Site
        [TestMethod]
        public void PU_PO_Filter_Site()
        {
            //Prepare
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();

            //Arrange

            HomePage homePage = LogInAsAdmin();
            //Act
            var purchaseOrderPage = homePage.GoToPurchasing_PurchaseOrdersPage();

            purchaseOrderPage.ResetFilters();
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.Site, site + " - " + site);

            if (purchaseOrderPage.CheckTotalNumber() < 20)
            {
                var createPurchaseOrderPage = purchaseOrderPage.CreateNewPurchaseOrder();
                createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+1), true);
                var purchaseOrderItemPage = createPurchaseOrderPage.Submit();
                purchaseOrderPage = purchaseOrderItemPage.BackToList();
            }

            if (!purchaseOrderPage.isPageSizeEqualsTo100())
            {
                purchaseOrderPage.PageSize("8");
                purchaseOrderPage.PageSize("100");
            }

            //Assert
            Assert.IsTrue(purchaseOrderPage.VerifySite(site), String.Format(MessageErreur.FILTRE_ERRONE, "Sites"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_GenerateRNFromPO()
        {
            //Prepare
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var decimalSeparator = homePage.GetDecimalSeparatorValue();
            //Etre sur une Purchase Order Validé

            //1.Créer un PO
            var purchaseOrderPage = homePage.GoToPurchasing_PurchaseOrdersPage();

            CreatePurchaseOrderModalPage purchaseOrder = purchaseOrderPage.CreateNewPurchaseOrder();
            purchaseOrder.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.Date, true);

            PurchaseOrderItem poItem = purchaseOrder.Submit();
            var poGeneralInfo = poItem.ClickOnGeneralInformation();
            string purchaseNumber = poGeneralInfo.getPurchaseOrderNumber();
            poItem = poGeneralInfo.ClickOnItemsTab();
            string itemName = poItem.GetFirstItemName();
            poItem.SelectFirstItemPo();
            //2.Remplir Prod qty sur la ligne d'un item
            poItem.AddQuantity("10");

            //3.Valider le PO
            poItem.Validate();

            //4.Survoler les...
            poItem.ShowExtendedMenu();
            //5.Cliquer sur generate RN
            //6.une pop'up RN s'ouvre
            //7.Cliquer sur create
            poItem.GenerateReceiptNote(true);
            ReceiptNotesItem receiptNoteItemPage = poItem.ValidateReceiptNoteCreation();

            //La RN est créé et vérifier les infos liées à la PO(items/ qtté)
            receiptNoteItemPage.ResetFilters();
            receiptNoteItemPage.Filter(ReceiptNotesItem.FilterItemType.SearchByName, itemName);
            Assert.AreEqual(itemName.Trim(), receiptNoteItemPage.GetFirstItemName().Trim(), "mauvais premier item");
            receiptNoteItemPage.SelectFirstItem();
            Assert.AreEqual(10.0, receiptNoteItemPage.GetDNQty(decimalSeparator), 0.00001, "mauvaise quantité");

            //Vérifier dans Gal info, le numéro de la PO de base
            ReceiptNotesGeneralInformation generalInfo = receiptNoteItemPage.ClickOnGeneralInformationTab();


            Assert.AreEqual(purchaseNumber, generalInfo.GetPurchaseOrderNumber());

            //Vérifier RN non validée
            ReceiptNotesPage list = generalInfo.BackToList();
            list.ResetFilter();
            list.Filter(ReceiptNotesPage.FilterType.ByPurchaseNumber, purchaseNumber);
            Assert.IsFalse(list.IsFirstLineValide(), "ligne valide");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_DeletePO()
        {
            var homePage = LogInAsAdmin();
            var purchaseOrderPage = homePage.GoToPurchasing_PurchaseOrdersPage();

            purchaseOrderPage.ResetFilters();
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.FilterShow, "Show not validated only");
            var number = purchaseOrderPage.DeleteFirstPurchaseOrder();
            var check = purchaseOrderPage.IsDeleted(number);
            Assert.IsTrue(check, "purchase order non supprimé");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_DetailDeleteItem()
        {
            var prodQty = "5";

            var homePage = LogInAsAdmin();
            var purchaseOrderPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            purchaseOrderPage.ResetFilters();
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.FilterShow, "Show not validated only");
            PurchaseOrderItem items = purchaseOrderPage.ClickFirstPurchaseOrder();
            items.Filter(FilterItemType.ShowItemsNotPurchased, true);
            var oldTotalVAT = "€ 0";//purchaseOrderPage.GetTotalVAT();
            var oldQuantity = "0";// purchaseOrderPage.GetQuantity();
            purchaseOrderPage.SetQuantity(prodQty);
            purchaseOrderPage.DeleteItem();
            var newQty = purchaseOrderPage.GetQuantity();
            var newTotalVAT = purchaseOrderPage.GetTotalVAT();
            Assert.IsTrue(purchaseOrderPage.VerifyValues(oldTotalVAT, newTotalVAT, oldQuantity, newQty), "Prod. Qty n'est pas repassé à 0 et Total w/o VAT n'a psa retrouvé la valeur avant ajout de quantité sur l'item");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_DetailEditItem()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var purchaseOrderPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            purchaseOrderPage.ResetFilters();
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.FilterShow, "Show not validated only");
            PurchaseOrderItem items = purchaseOrderPage.ClickFirstPurchaseOrder();
            items.Filter(FilterItemType.ShowItemsNotPurchased, true);
            var firstItemName = purchaseOrderPage.GetFirstItemNumber();
            purchaseOrderPage.ClickActionsButton();
            var itemGeneralinformationPage = purchaseOrderPage.ClickPen();
            var firstItemNameFromEditMenu = itemGeneralinformationPage.GetItemNameFromEditMenu();
            int indexOfOpenParenthesis = firstItemName.IndexOf('(');
            if (indexOfOpenParenthesis >= 0)
            {
                // La parenthèse ouvrante a été trouvée, nous pouvons utiliser Substring
                firstItemName = firstItemName.Substring(0, indexOfOpenParenthesis).Trim();
            }
            Assert.AreEqual(firstItemName.Trim(), firstItemNameFromEditMenu.Trim(), "le nom est différent dans le menu d'edition");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_DetailSendByMail()
        {
            var email = TestContext.Properties["Admin_UserName"].ToString();
            HomePage homePage = LogInAsAdmin();

            var purchaseOrderPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            purchaseOrderPage.ResetFilters();
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.FilterShow, "Show validated only");
            var site = purchaseOrderPage.GetFirstPurchaseOrderSite();
            var number = purchaseOrderPage.ClickFirstPurchaseOrderGetNumber();
            var btn = purchaseOrderPage.GetExtendedMenuButton();
            Actions action = new Actions(WebDriver);
            action.MoveToElement(btn).Perform();
            purchaseOrderPage.SendByEmail(email);
            // check if mail received
            MailPage mailPage = purchaseOrderPage.RedirectToOutlookMailbox();
            mailPage.CheckIfSpecifiedOutlookMailExist("Newrest - Purchase order " + number + " - " + site + " - " );
            Assert.IsTrue(mailPage.CheckIfSpecifiedOutlookMailExist("Newrest - Purchase order " + number + " - " + site + " - "), "purchase order n'a pas été envoyée par mail.");
            mailPage.DeleteCurrentOutlookMail();
            mailPage.Close();
            while (!WebDriver.Url.Contains("winrest-testauto"))
            {
                mailPage.Close();
            }
            homePage.GoToPurchasing_PurchaseOrdersPage();
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.ByNumber, number);
            Assert.IsTrue(purchaseOrderPage.IsPurchaseOrderSentByMail(), "icone enveloppe n'apparait pas dans l'index");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_DetailVerifyFooter()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var purchaseOrderPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            purchaseOrderPage.ResetFilters();
            purchaseOrderPage.ClickFirstPurchaseOrderGetNumber();
            var totalVat = purchaseOrderPage.GetTotalVAT();
            purchaseOrderPage.GoToFooterSubMenu();
            var vatFooter = purchaseOrderPage.GetTotalGrossAmountFromFooter();
            Assert.AreEqual(totalVat, vatFooter, "les montants ne sont pas correctes");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_Filter_DetailItem_BySubGroup()
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
            productionPage.AddNewSubGroup(subGrpName, subGrpCode);

            //Click first purchase order
            var purchaseOrderPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            var purchaseOrderNumber = purchaseOrderPage.GetFirstPurchaseOrderNumber();
            var purchaseOrderItems = purchaseOrderPage.ClickFirstPurchaseOrder();
            Thread.Sleep(1000);
            //get first purchase order item name
            var itemName = purchaseOrderItems.GetFirstItemName();

            //set item subgroup
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            var itemGeneralInformationPage = itemPage.ClickOnFirstItem();
            itemGeneralInformationPage.SetGroupName(group);
            itemGeneralInformationPage.SetSubgroupName(subGrpName);

            //filter by subgroupe
            purchaseOrderPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.ByNumber, purchaseOrderNumber);
            purchaseOrderItems = purchaseOrderPage.ClickFirstPurchaseOrder();
            purchaseOrderItems.Filter(FilterItemType.ByGroup, group);
            purchaseOrderItems.Filter(FilterItemType.BySubGroup, subGrpName);

            //assert
            Assert.IsTrue(purchaseOrderItems.VerifySubGroupFiltre(itemName), "erreur lors du filtrage par subgroupe");

            //changing setting sub group to inactive
            homePage.SetSubGroupFunctionValue(false);
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_Pagination()
        {
            //connect as admin
            LogInAsAdmin();

            //go to homePage
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var purchaseOrderPage = homePage.GoToPurchasing_PurchaseOrdersPage();

            int tot = purchaseOrderPage.CheckTotalNumber();

            purchaseOrderPage.PageSize("16");

            Assert.IsTrue(purchaseOrderPage.isPageSizeEqualsTo16(), "Paggination ne fonctionne pas.");
            Assert.AreEqual(purchaseOrderPage.GetTotalRowsForPagination(), tot >= 16 ? 16 : tot, "Paggination ne fonctionne pas.");

            purchaseOrderPage.PageSize("30");

            Assert.IsTrue(purchaseOrderPage.isPageSizeEqualsTo30(), "Paggination ne fonctionne pas.");
            Assert.AreEqual(purchaseOrderPage.GetTotalRowsForPagination(), tot >= 30 ? 30 : tot, "Paggination ne fonctionne pas.");

            purchaseOrderPage.PageSize("50");

            Assert.IsTrue(purchaseOrderPage.isPageSizeEqualsTo50(), "Paggination ne fonctionne pas.");
            Assert.AreEqual(purchaseOrderPage.GetTotalRowsForPagination(), tot >= 50 ? 50 : tot, "Paggination ne fonctionne pas.");

            purchaseOrderPage.PageSize("100");

            Assert.IsTrue(purchaseOrderPage.isPageSizeEqualsTo100(), "Paggination ne fonctionne pas.");
            Assert.AreEqual(purchaseOrderPage.GetTotalRowsForPagination(), tot >= 100 ? 100 : tot, "Paggination ne fonctionne pas.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_DetailChangeLine()
        {
            //Prepare
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();

            //connect as admin
            LogInAsAdmin();

            //go to homePage
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var purchaseOrderPage = homePage.GoToPurchasing_PurchaseOrdersPage();

            //1.Créer un PO
            CreatePurchaseOrderModalPage purchaseOrder = purchaseOrderPage.CreateNewPurchaseOrder();
            purchaseOrder.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.Date, true);
            PurchaseOrderItem poItem = purchaseOrder.Submit();
            poItem.SelectFirstItemPo();
            //Assert
            Assert.IsTrue(poItem.VerifyDetailChangeLine(), String.Format(MessageErreur.FILTRE_ERRONE, "ChangeLine ne fonctionne pas."));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_ApprovalApprove()
        {
            //Prepare
            var role = "Admin";
            var userFullName = TestContext.Properties["Admin_FullName"].ToString();
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();
            string numberPurchaseOrder = "";

            //connect as admin
            var homePage = LogInAsAdmin();

            //set admin as po approver
            var applicationSettings = homePage.GoToApplicationSettings();
            var appSettingsModalPage = applicationSettings.GetWhoMustApprovePOPage();
            appSettingsModalPage.SetWhoMustApprovePO(role, true);
            appSettingsModalPage.Save();
            var purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();

            try
            {
                var createPurchaseOrderPage = purchaseOrdersPage.CreateNewPurchaseOrder();
                createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+1), true);
                var purchaseOrderItemPage = createPurchaseOrderPage.Submit();
                numberPurchaseOrder = purchaseOrderItemPage.GetPurchaseOrderNumber();
                purchaseOrderItemPage.SelectFirstItemPo();
                purchaseOrderItemPage.AddQuantity("5");

                purchaseOrderItemPage.Approve();

                var generalInformationPage = purchaseOrderItemPage.ClickOnGeneralInformation();
                var approvedBy = generalInformationPage.GetApprovedBy();
                var check = approvedBy.StartsWith(userFullName);
                Assert.IsTrue(check, "L'approbation n'a pas été faite par l'utilisateur.");
            }
            finally
            {
                // remove admin as po approver
                applicationSettings = homePage.GoToApplicationSettings();
                appSettingsModalPage = applicationSettings.GetWhoMustApprovePOPage();
                appSettingsModalPage.SetWhoMustApprovePO(role, false);
                appSettingsModalPage.Save();
                // delete purchase order
                purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();
                purchaseOrdersPage.Filter(PurchaseOrdersPage.FilterType.ByNumber, numberPurchaseOrder);
                purchaseOrdersPage.DeleteFirstPurchaseOrder();

            }
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_ApprovalValidateWithoutApprove()
        {
            //Prepare
            var role = "Admin";
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();

            //connect as admin
            LogInAsAdmin();

            //go to homePage
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //set admin as po approver
            var applicationSettings = homePage.GoToApplicationSettings();
            var appSettingsModalPage = applicationSettings.GetWhoMustApprovePOPage();
            appSettingsModalPage.SetWhoMustApprovePO(role, true);
            appSettingsModalPage.Save();

            try
            {
                homePage.Navigate();
                var purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();

                var createPurchaseOrderPage = purchaseOrdersPage.CreateNewPurchaseOrder();
                createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+1), true);
                var purchaseOrderItemPage = createPurchaseOrderPage.Submit();
                purchaseOrderItemPage.SelectFirstItemPo();
                purchaseOrderItemPage.AddQuantity("5");

                var validateHasError = purchaseOrderItemPage.ValidateHasError();
                Assert.IsTrue(validateHasError, "La validation sans approbation n'a pas généré d'erreur.");
            }
            finally
            {
                // remove admin as po approver
                applicationSettings = homePage.GoToApplicationSettings();
                appSettingsModalPage = applicationSettings.GetWhoMustApprovePOPage();
                appSettingsModalPage.SetWhoMustApprovePO(role, false);
                appSettingsModalPage.Save();
            }
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_ApprovalValidateApprovedPO()
        {
            //Prepare
            var role = "Admin";
            var userFullName = TestContext.Properties["Admin_FullName"].ToString();
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();

            //go to homePage
            var homePage = LogInAsAdmin();

            //set admin as po approver
            var applicationSettings = homePage.GoToApplicationSettings();
            var appSettingsModalPage = applicationSettings.GetWhoMustApprovePOPage();
            appSettingsModalPage.SetWhoMustApprovePO(role, true);
            appSettingsModalPage.Save();

            try
            {
                var purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();
                var createPurchaseOrderPage = purchaseOrdersPage.CreateNewPurchaseOrder();
                createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+1), true);
                var purchaseOrderItemPage = createPurchaseOrderPage.Submit();
                purchaseOrderItemPage.SelectFirstItemPo();
                purchaseOrderItemPage.AddQuantity("5");

                purchaseOrderItemPage.Approve();
                purchaseOrderItemPage.Validate();

                var generalInformationPage = purchaseOrderItemPage.ClickOnGeneralInformation();
                var userValidatorName = generalInformationPage.GetUserValidator();
                Assert.AreEqual(userFullName, userValidatorName, "La validation d'un PO approuvé n'a pas été faite par l'utilisateur.");
            }
            finally
            {
                // remove admin as po approver
                applicationSettings = homePage.GoToApplicationSettings();
                appSettingsModalPage = applicationSettings.GetWhoMustApprovePOPage();
                appSettingsModalPage.SetWhoMustApprovePO(role, false);
                appSettingsModalPage.Save();
            }
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_ApprovalPrintApprovedPo()
        {
            //Prepare
            var role = "Admin";
            var userFullName = TestContext.Properties["Admin_FullName"].ToString();
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();
            // For PRINT PDF
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNamePdfBegin = "Purchase order report_-_";
            string DocFileNameZipBegin = "All_files_";
            bool newVersion = true;
            //connect as admin
            LogInAsAdmin();
            //go to homePage
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            //set admin as po approver
            var applicationSettings = homePage.GoToApplicationSettings();
            var appSettingsModalPage = applicationSettings.GetWhoMustApprovePOPage();
            appSettingsModalPage.SetWhoMustApprovePO(role, true);
            appSettingsModalPage.Save();
            try
            {
                homePage.Navigate();
                var purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();
                var createPurchaseOrderPage = purchaseOrdersPage.CreateNewPurchaseOrder();
                createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+1), true);
                var purchaseOrderItemPage = createPurchaseOrderPage.Submit();
                purchaseOrderItemPage.SelectFirstItemPo();
                purchaseOrderItemPage.AddQuantity("5");
                purchaseOrderItemPage.Approve();
                purchaseOrderItemPage.Validate();
                purchaseOrdersPage.ClearDownloads();
                var reportPage = purchaseOrdersPage.PrintResultsNew(newVersion);
                var isReportGenerated = reportPage.IsReportGenerated();
                reportPage.Close();
                //Assert
                Assert.IsTrue(isReportGenerated, "L'application n'a pas pu générer le fichier attendu.");
                reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
                // cliquer sur All
                string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
                FileInfo fi = new FileInfo(trouve);
                fi.Refresh();
                Assert.IsTrue(fi.Exists, trouve + " non généré");

                PdfDocument document = PdfDocument.Open(fi.FullName);
                List<string> mots = new List<string>();
                string[] wordsInName = userFullName.Split(' ');

                foreach (Page p in document.GetPages())
                {
                    mots.AddRange(p.GetWords().Select(m => m.Text));
                }

                Assert.AreEqual(1, mots.Count(w => w.Contains(role)), "Le Role du compte ayant approuvé n'apparaît pas le Pdf");
                Assert.AreEqual(2, mots.Count(w => w.Contains(wordsInName[0])), "Le Nom du compte ayant approuvé n'apparaît pas dans le Pdf");
                Assert.AreEqual(2, mots.Count(w => w.Contains(wordsInName[1])), "Le Nom du compte ayant approuvé n'apparaît pas dans le Pdf");
                Assert.AreEqual(2, mots.Count(w => w.Contains(wordsInName[2])), "Le Nom du compte ayant approuvé n'apparaît pas dans le Pdf");
            }
            finally
            {
                // remove admin as po approver
                applicationSettings = homePage.GoToApplicationSettings();
                appSettingsModalPage = applicationSettings.GetWhoMustApprovePOPage();
                appSettingsModalPage.SetWhoMustApprovePO(role, false);
                appSettingsModalPage.Save();
            }
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_ApprovalPrintApprovedPo2()
        {
            //Prepare
            var role = "Admin";
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();
            //connect as admin
            LogInAsAdmin();
            //go to homePage
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            //set admin as po approver
            var applicationSettings = homePage.GoToApplicationSettings();
            var appSettingsModalPage = applicationSettings.GetWhoMustApprovePOPage();
            appSettingsModalPage.SetWhoMustApprovePO(role, true);
            appSettingsModalPage.Save();
            try
            {
                homePage.Navigate();
                var purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();
                var createPurchaseOrderPage = purchaseOrdersPage.CreateNewPurchaseOrder();
                createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+1), true);
                var purchaseOrderItemPage = createPurchaseOrderPage.Submit();
                purchaseOrderItemPage.SelectFirstItemPo();
                purchaseOrderItemPage.AddQuantity("5");
                purchaseOrderItemPage.Approve();
                purchaseOrderItemPage.Validate();
                var ID_premier_PO = purchaseOrderItemPage.GetPurchaseOrderNumber();
                var generalInformationPage = purchaseOrderItemPage.ClickOnGeneralInformation();
                var approvedBy = generalInformationPage.GetApprovedBy();
                purchaseOrderItemPage.BackToList();
                createPurchaseOrderPage = purchaseOrdersPage.CreateNewPurchaseOrder();
                createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+1), true);
                purchaseOrderItemPage = createPurchaseOrderPage.Submit();
                var ID_deuxieme_PO = purchaseOrderItemPage.GetPurchaseOrderNumber();
                purchaseOrderItemPage.BackToList();
                purchaseOrderItemPage.ResetFilter();
                purchaseOrdersPage.Filter(PurchaseOrdersPage.FilterType.ByNumber, ID_premier_PO);
                purchaseOrdersPage.Filter(PurchaseOrdersPage.FilterType.tobeapprovedby, role);
                purchaseOrderItemPage.ResetFilter();
                purchaseOrdersPage.Filter(PurchaseOrdersPage.FilterType.ByNumber, ID_deuxieme_PO);

                Assert.AreNotEqual(ID_premier_PO, purchaseOrdersPage.GetFirstPurchaseOrderNumber(), "Le purchase order créé approuver et valider apparait dans le filtre To be approved by.");
                Assert.AreEqual(ID_deuxieme_PO, purchaseOrdersPage.GetFirstPurchaseOrderNumber(), "Le purchase order créé non approuver n'apparait pas dans le filtre To be approved by.");
            }
            finally
            {
                // remove admin as po approver
                applicationSettings = homePage.GoToApplicationSettings();
                appSettingsModalPage = applicationSettings.GetWhoMustApprovePOPage();
                appSettingsModalPage.SetWhoMustApprovePO(role, false);
                appSettingsModalPage.Save();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_PrintDetailBackdate()
        {
            //Prepare
            var role = "Admin";
            var userFullName = TestContext.Properties["Admin_FullName"].ToString();
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();

            //go to homePage
            var homePage = LogInAsAdmin();

            homePage.ClearDownloads();

            //set admin as po approver
            var applicationSettings = homePage.GoToApplicationSettings();
            var appSettingsModalPage = applicationSettings.GetWhoMustApprovePOPage();
            appSettingsModalPage.SetWhoMustApprovePO(role, true);
            appSettingsModalPage.Save();

            try
            {
                var purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();

                var createPurchaseOrderPage = purchaseOrdersPage.CreateNewPurchaseOrder();
                createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(-11), true);
                var purchaseOrderItemPage = createPurchaseOrderPage.Submit();
                purchaseOrderItemPage.SelectFirstItemPo();
                purchaseOrderItemPage.AddQuantity("5");

                bool validateUnsuccess = purchaseOrderItemPage.ValidateHasError();
                Assert.IsTrue(validateUnsuccess, "La validation ne doit pas etre possible");
                var printReportPage = purchaseOrderItemPage.PrintDetails(true);
                var isReportGenerated = printReportPage.IsReportGenerated();
                Assert.IsTrue(isReportGenerated, "L'impression ne doit pas etre possible");
                printReportPage.Close();

                purchaseOrderItemPage.Approve();
                purchaseOrderItemPage.Validate();
                printReportPage = purchaseOrderItemPage.PrintDetails(true);
                isReportGenerated = printReportPage.IsReportGenerated();
                Assert.IsTrue(isReportGenerated, "L'impression doit etre possible après Approve");
                printReportPage.Close();
            }
            finally
            {
                // remove admin as po approver
                applicationSettings = homePage.GoToApplicationSettings();
                appSettingsModalPage = applicationSettings.GetWhoMustApprovePOPage();
                appSettingsModalPage.SetWhoMustApprovePO(role, false);
                appSettingsModalPage.Save();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_PrintIndexBackdate()
        {
            //Prepare
            var role = "Admin";
            var userFullName = TestContext.Properties["Admin_FullName"].ToString();
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();

            //go to homePage
            var homePage = LogInAsAdmin();


            //set admin as po approver
            var applicationSettings = homePage.GoToApplicationSettings();
            var appSettingsModalPage = applicationSettings.GetWhoMustApprovePOPage();
            appSettingsModalPage.SetWhoMustApprovePO(role, true);
            appSettingsModalPage.Save();

            try
            {
                var purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();

                var createPurchaseOrderPage = purchaseOrdersPage.CreateNewPurchaseOrder();
                createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(-11), true);
                var purchaseOrderItemPage = createPurchaseOrderPage.Submit();
                purchaseOrderItemPage.SelectFirstItemPo();
                purchaseOrderItemPage.AddQuantity("5");

                bool validateUnsuccess = purchaseOrderItemPage.ValidateHasError();
                Assert.IsTrue(validateUnsuccess, "La validation ne doit pas etre possible");
                homePage.ClearDownloads();
                var printReportPage = purchaseOrderItemPage.PrintDetails(true);
                var isReportGenerated = printReportPage.IsReportGenerated();
                Assert.IsTrue(isReportGenerated, "L'impression doit etre possible avant Approve");
                printReportPage.Close();

                purchaseOrderItemPage.Approve();
                purchaseOrderItemPage.Validate();
                homePage.ClearDownloads();
                printReportPage = purchaseOrderItemPage.PrintDetails(true);
                isReportGenerated = printReportPage.IsReportGenerated();
                Assert.IsTrue(isReportGenerated, "L'impression doit etre possible après Approve");
                printReportPage.Close();
            }
            finally
            {
                // remove admin as po approver
                applicationSettings = homePage.GoToApplicationSettings();
                appSettingsModalPage = applicationSettings.GetWhoMustApprovePOPage();
                appSettingsModalPage.SetWhoMustApprovePO(role, false);
                appSettingsModalPage.Save();
            }


        }

        /// <summary>
        /// Update purchase order's first item allergens and check the state 
        /// of the allergens button in the purchase order details.
        /// </summary>
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_CheckAllergensUpdate()
        {
            #region PARAMS
            string allergen1 = "Cacahuetes/Peanuts";
            string allergen2 = "Frutos de cáscara- Macadamias/Nuts-Macadamia";
            #endregion

            #region LOG IN
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            #endregion

            #region ACT
            // Select item
            var purchasesPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            purchasesPage.ResetFilters();
            purchasesPage.Filter(PurchaseOrdersPage.FilterType.FilterShow, "Show not validated only");
            var po = purchasesPage.ClickFirstPurchaseOrder();
            var itemName = po.GetFirstItemName();
            po.SelectFirstItemPo();

            // Edit allergens
            var itemPage = po.EditItem(itemName);
            var itemIntolerancePage = itemPage.ClickOnIntolerancePage();
            itemIntolerancePage.AddAllergen(allergen1);
            itemIntolerancePage.AddAllergen(allergen2);

            // Check PO update
            purchasesPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            purchasesPage.ResetFilters();
            purchasesPage.Filter(PurchaseOrdersPage.FilterType.FilterShow, "Show not validated only");
            po = purchasesPage.ClickFirstPurchaseOrder();

            bool isIconGreen = po.IsAllergenIconGreen(itemName);
            List<string> allergensInItem = po.GetAllergens(itemName);
            #endregion

            #region ASSERT
            Assert.IsTrue(isIconGreen);
            Assert.IsTrue(allergensInItem.Contains(allergen1) && allergensInItem.Contains(allergen2));
            #endregion
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_SendMailGeneralInformation()
        {
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();
            var email = TestContext.Properties["Admin_UserName"].ToString();
            string ID = string.Empty;
           
            HomePage homePage = LogInAsAdmin();
            //Act
            var purchaseOrderPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            purchaseOrderPage.ResetFilters();
            purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.FilterShow, "Show validated only");

            if (purchaseOrderPage.CheckTotalNumber() == 0)
            {
                //var purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();
                var createPurchaseOrderPage = purchaseOrderPage.CreateNewPurchaseOrder();
                createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now, true);
                var purchaseOrderItemPage = createPurchaseOrderPage.Submit();
                ID = purchaseOrderItemPage.GetPurchaseOrderNumber();
                purchaseOrderItemPage.WaitPageLoading();
                purchaseOrderItemPage.SelectFirstItemPo();
                purchaseOrderItemPage.AddQuantity("2");
                purchaseOrderItemPage.Validate();
                purchaseOrderPage = purchaseOrderItemPage.BackToList();
                purchaseOrderPage.ResetFilters();
                purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.ByNumber, ID);

            }

            PurchaseOrderItem purchaseOrderItem = purchaseOrderPage.ClickFirstPurchaseOrder();
            purchaseOrderItem.SendByEmail(email);
            PurchaseOrderGeneralInformation purchaseOrderGeneralInformation = purchaseOrderItem.ClickOnGeneralInformation();
            var Date = purchaseOrderItem.GetDeliveryDate();
            purchaseOrderItem.ShowExtendedMenu();
            var emailSentDate = purchaseOrderGeneralInformation.GetGenenralInfoEmail(ID, Date);
            //assert
            Assert.IsTrue(emailSentDate, "La date d'envoi de l'e-mail n'est pas correcte.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_SendMailMailBody()
        {
            var userFullName = TestContext.Properties["Admin_FullName"].ToString();
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();
            var email = TestContext.Properties["Admin_UserName"].ToString();
            var message = "Email body test auto";


            HomePage homePage = LogInAsAdmin();
            var applicationSettings = homePage.GoToApplicationSettings();
            ParametersSites sitesPage = homePage.GoToParameters_Sites();
            sitesPage.ClickOnPurchaseOrderMailbodyTab();
            ParametersSitesModalPage modalCreate = sitesPage.ClickOnNewPurchaseOrdermailbodyModal();
            modalCreate.CreatePurchaseOrderMailBody("B59500843 - NEWREST ESPAÑA S.L.", "MAD", message);
            try
            {
                homePage.Navigate();
                var purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();
                var createPurchaseOrderPage = purchaseOrdersPage.CreateNewPurchaseOrder();
                createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+1), true);
                var purchaseOrderItemPage = createPurchaseOrderPage.Submit();
                var number = purchaseOrderItemPage.GetPurchaseOrderNumber();
                purchaseOrderItemPage.SelectFirstItemPo();
                purchaseOrderItemPage.AddQuantity("5");
                purchaseOrderItemPage.Validate();
                var btn = purchaseOrdersPage.GetExtendedMenuButton();
                Actions action = new Actions(WebDriver);
                action.MoveToElement(btn).Perform();
                purchaseOrdersPage.SendByEmail(email);
                // check if mail received
                MailPage mailPage = purchaseOrdersPage.RedirectToOutlookMailbox();
                mailPage.FillFields_LogToOutlookMailbox(email);
                Assert.IsTrue(mailPage.CheckIfSpecifiedOutlookMailExist("Newrest - Purchase order " + number + " - " + site + " - "), "purchase order n'a pas été envoyée par mail.");
                Assert.IsTrue(mailPage.CheckIfEmailBodyContainsText(message), "Le message n'est pas le même que celui dans le body du mail.");
                mailPage.DeleteCurrentOutlookMail();
                mailPage.Close();
            }
            finally
            {
                homePage.Navigate();
                applicationSettings = homePage.GoToApplicationSettings();
                ParametersSites sitePage = homePage.GoToParameters_Sites();
                sitePage.ClickOnPurchaseOrderMailbodyTab();
                sitePage.DeletePurchaseOrderMailbody();

            }
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_Details_UpdateStockQty()
        {
            //Prepare
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();
            string quantity = "7";
            // arrange
            var homePage = LogInAsAdmin();
            var purchaseOrderPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            var createPurchaseOrderPage = purchaseOrderPage.CreateNewPurchaseOrder();
            createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(0), true);
            var purchaseOrderItem = createPurchaseOrderPage.Submit();
            var oldQuantity = purchaseOrderItem.GetQuantity();
            purchaseOrderItem.SelectFirstItemPo();
            purchaseOrderItem.AddQuantity(quantity);
            purchaseOrderItem.Validate();
            var newQuantity = purchaseOrderItem.GetQuantity();
            //Assert
            Assert.AreNotEqual(oldQuantity, newQuantity, "La quantité de l'item n'a pas été mis à jour.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_Details_ShowReferenceItem()
        {
            //Test de la création d'un nouveau purchase order
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            //Act
            var purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            var FirstPurchaseOrderNumber = purchaseOrdersPage.GetFirstPurchaseOrderNumber();
            purchaseOrdersPage.Filter(PurchaseOrdersPage.FilterType.ByNumber, FirstPurchaseOrderNumber);
            purchaseOrdersPage.ClickFirstItemFiltred();
            string firstname = purchaseOrdersPage.GetNameOfTheItem();
            string reference;
            string itemName;
            purchaseOrdersPage.SplitNameAndReference(firstname, out reference, out itemName);
            purchaseOrdersPage.GoToItem();
            purchaseOrdersPage.FilterItem(PurchaseOrdersPage.FilterType.ByName, itemName);
            purchaseOrdersPage.ClickOnItem();
            string referencefromitem = purchaseOrdersPage.Getreference();
            Assert.AreEqual(reference, referencefromitem);
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_Update_Quantity()
        {
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            PurchaseOrdersPage purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            /* We gonna filter only the invalid purchase Orders */
            purchaseOrdersPage.Filter(PurchaseOrdersPage.FilterType.FilterShow, "Show not validated only");
            /* Clicking on the first invalid purchase order */
            purchaseOrdersPage.ClickFirstItemFiltred();
            /* We gonna changer the value of the Qty to 100 in case it is null*/
            purchaseOrdersPage.SetQuantity("100");
            string first_total = purchaseOrdersPage.GetTotalVAT();
            /* We gonna put the the value of the Qty to zero and test if it s gonna change the total VAT */
            purchaseOrdersPage.SetQuantity("0");
            string second_total = purchaseOrdersPage.GetTotalVAT();

            // Assert 
            Assert.AreNotEqual(first_total, second_total, "The quantity didn't changed the Total w/o VAT");


        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_PrintDetail_ValidatedPurchaseOrder()
        {
            var homePage = LogInAsAdmin();

            PurchaseOrdersPage purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            purchaseOrdersPage.ResetFilters();
            purchaseOrdersPage.Filter(PurchaseOrdersPage.FilterType.FilterShow, "Show validated only");
            PurchaseOrderItem purchaseOrderItem = purchaseOrdersPage.ClickFirstPurchaseOrder();
            PrintReportPage printReportPage = purchaseOrderItem.PrintDetails(true);
            var isReportGenerated = printReportPage.IsReportGenerated();
            printReportPage.Close();

            Assert.IsTrue(isReportGenerated, "L'application n'a pas pu générer le fichier PDF.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_PrintDetail_NotValidatedPurchaseOrder()
        {
            var homePage = LogInAsAdmin();

            PurchaseOrdersPage purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            purchaseOrdersPage.ResetFilters();
            purchaseOrdersPage.Filter(PurchaseOrdersPage.FilterType.FilterShow, "Show not validated only");
            PurchaseOrderItem purchaseOrderItem = purchaseOrdersPage.ClickFirstPurchaseOrder();
            purchaseOrderItem.SelectFirstItemPo();
            purchaseOrderItem.ChangeProdQuantity("0");
            PrintReportPage printReportPage = purchaseOrderItem.PrintDetails(true);
            var isReportGenerated = printReportPage.IsReportGenerated();
            printReportPage.Close();

            Assert.IsTrue(isReportGenerated, "L'application n'a pas pu générer le fichier PDF.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_PrintDetail_PackagingPurchaseOrder()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string docFileNamePdfBegin = "Purchase order report_-_";
            string docFileNameZipBegin = "All_files_";

            var homePage = LogInAsAdmin();

            var purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            purchaseOrdersPage.ResetFilters();
            purchaseOrdersPage.Filter(PurchaseOrdersPage.FilterType.FilterShow, "Show validated only");
            var purchaseOrderItem = purchaseOrdersPage.ClickFirstPurchaseOrder();
            string[] displayedPackings = purchaseOrderItem.GetPacking().Split(new[] { ' ' });
            purchaseOrderItem.ClearDownloads();
            purchaseOrderItem.ShowExtendedMenu();
            var printReportPage = purchaseOrderItem.PrintDetails(true);
            var isReportGenerated = printReportPage.IsReportGenerated();
            printReportPage.Close();

            Assert.IsTrue(isReportGenerated, "L'application n'a pas pu générer le fichier PDF.");

            printReportPage.Purge(downloadsPath, docFileNamePdfBegin, docFileNameZipBegin);
            string generatedFilePath = printReportPage.PrintAllZip(downloadsPath, docFileNamePdfBegin, docFileNameZipBegin);

            FileInfo fileInfo = new FileInfo(generatedFilePath);
            fileInfo.Refresh();
            Assert.IsTrue(fileInfo.Exists, $"{generatedFilePath} non généré");

            PdfDocument document = PdfDocument.Open(fileInfo.FullName);
            List<string> words = new List<string>();
            foreach (Page page in document.GetPages())
            {
                words.AddRange(page.GetWords().Select(word => word.Text));
            }

            foreach (string packing in displayedPackings)
            {
                bool packingExist = words.Any(word => word.Contains(packing));
                Assert.IsTrue(packingExist, $"Le packing n'a pas été trouvé dans le PDF.");
            }
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_Print_Site_Name()
        {
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string docFileNamePdfBegin = "Purchase order report_-_";
            string docFileNameZipBegin = "All_files_";


            var homePage = LogInAsAdmin();
            var purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();

            purchaseOrdersPage.ResetFilters();
            purchaseOrdersPage.Filter(PurchaseOrdersPage.FilterType.ByValidationDate, true);

            var site = purchaseOrdersPage.GetFirstPurchaseOrderSite();

            var settingsSitesPage = homePage.GoToParameters_Sites();
            settingsSitesPage.Filter(ParametersSites.FilterType.SearchSite, site);
            settingsSitesPage.ClickOnFirstSite();
            settingsSitesPage.ClickToInformations();

            var address = settingsSitesPage.GetAddress2().Replace("-", "").Trim();
            string[] addressSegments = address.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            var purchaseOrderItem = purchaseOrdersPage.SelectFirstItem();

            purchaseOrderItem.ClearDownloads();
            purchaseOrderItem.ShowExtendedMenu();

            var printReportPage = purchaseOrderItem.PrintDetails(true);
            var isReportGenerated = printReportPage.IsReportGenerated();
            printReportPage.Close();

            Assert.IsTrue(isReportGenerated, "L'application n'a pas pu générer le fichier attendu.");

            printReportPage.Purge(downloadsPath, docFileNamePdfBegin, docFileNameZipBegin);
            string generatedFilePath = printReportPage.PrintAllZip(downloadsPath, docFileNamePdfBegin, docFileNameZipBegin);

            FileInfo fileInfo = new FileInfo(generatedFilePath);
            fileInfo.Refresh();
            Assert.IsTrue(fileInfo.Exists, $"{generatedFilePath} non généré");

            PdfDocument document = PdfDocument.Open(fileInfo.FullName);
            List<string> words = new List<string>();
            foreach (Page page in document.GetPages())
            {
                words.AddRange(page.GetWords().Select(word => word.Text));
            }

            foreach (string segment in addressSegments)
            {
                bool segmentPresent = words.Any(word => word.Contains(segment));
                Assert.IsTrue(segmentPresent, $"Le segment d'adresse '{segment}' n'a pas été trouvé dans le PDF.");
            }

            bool sitePresent = words.Any(word => word.Contains(site));
            Assert.IsTrue(sitePresent, $"Le site '{site}' n'a pas été trouvé dans le PDF.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_Actived_Unchecked()
        {
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();
            string qty = "5";

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            var createPurchaseOrderModalPage = purchaseOrdersPage.CreateNewPurchaseOrder();

            createPurchaseOrderModalPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+3), true);
            var purchaseOrderItem = createPurchaseOrderModalPage.Submit();
            purchaseOrderItem.SelectFirstItemPo();
            purchaseOrderItem.AddQuantity(qty);

            Assert.IsTrue(purchaseOrderItem.ValidateIsEnabled(), "Le bouton validate est grisé.");

            purchaseOrderItem.Validate();
            purchaseOrderItem.BackToList();

            purchaseOrdersPage.CreateNewPurchaseOrder();
            createPurchaseOrderModalPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+3), false);
            createPurchaseOrderModalPage.Submit();
            purchaseOrderItem.SelectFirstItemPo();
            purchaseOrderItem.AddQuantity(qty);

            Assert.IsFalse(purchaseOrderItem.ValidateIsEnabled(), "Le bouton validate n'est pas grisé.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_ordered_Qty()
        {
            // Prepare
            var keyword = TestContext.Properties["Item_Keyword"].ToString();
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();
            string quantity = "5";
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            PurchaseOrdersPage purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            var createPurchaseOrderPage = purchaseOrdersPage.CreateNewPurchaseOrder();
            createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder("ACE", supplier, location, DateUtils.Now.AddDays(+10), true);
            //createPurchaseOrderPage.Submit();
            var purchaseOrderItemPage = createPurchaseOrderPage.Submit();
            purchaseOrderItemPage.SelectFirstItemPo();
            purchaseOrderItemPage.AddQuantity(quantity);
            purchaseOrderItemPage.Validate();
            var firstItem = purchaseOrderItemPage.getorderedQty();
            purchaseOrderItemPage.BackToList();
            purchaseOrdersPage.CreateNewPurchaseOrder();
            createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder("ACE", supplier, location, DateUtils.Now.AddDays(+10), true);
            createPurchaseOrderPage.Submit();
            purchaseOrderItemPage.SelectFirstItemPo();
            var secondItem = purchaseOrderItemPage.getorderedQty();
            Assert.AreNotEqual(0, firstItem, secondItem, "Erreur : Lors de la création de la deuxième PO, la colonne 'ordered qty' doit correspondre aux quantités déjà commandées dans la PO N°1 et ne doit pas être égale à 0.");

        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_Search_Items()
        {
            var location = TestContext.Properties["Location"].ToString();

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            PurchaseOrdersPage purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            var createPurchaseOrderPage = purchaseOrdersPage.CreateNewPurchaseOrder();

            createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder("ACE", "FRAN Y CHEMI, S.L.", location, DateUtils.Now.AddDays(+10), true);
            var purchaseOrderItemPage = createPurchaseOrderPage.Submit();

            var firstItemName = purchaseOrderItemPage.GetFirstItemName();

            purchaseOrderItemPage.Filter(FilterItemType.ByName, firstItemName);

            // Vérifier que le seul élément affiché après le filtrage correspond au nom recherché
            var filteredItemName = purchaseOrderItemPage.GetFirstItemName();
            Assert.AreEqual(firstItemName, filteredItemName, "Le filtre n'a pas fonctionné comme prévu.");

            // Vérifier qu'il n'y a qu'un seul élément dans la liste filtrée
            var filteredItemsCount = purchaseOrderItemPage.GetItemsCount();
            Assert.AreEqual(1, filteredItemsCount, "Le filtre doit renvoyer un seul élément.");

        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_SendForAperiod_MoreThanOneMonth()
        {
            //Preapare
            var location = TestContext.Properties["Location"].ToString();
            var email = TestContext.Properties["Admin_UserName"].ToString();
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            PurchaseOrdersPage purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            purchaseOrdersPage.Filter(PurchaseOrdersPage.FilterType.ByValidationDate, true);
            var numbers = purchaseOrdersPage.GetTotalPurshaseOrder();
            if (numbers == 0)
            {
                var createPurchaseOrderPage = purchaseOrdersPage.CreateNewPurchaseOrder();
                createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now, true);
                var purchaseOrderItemPage = createPurchaseOrderPage.Submit();

                purchaseOrderItemPage.WaitPageLoading();
                purchaseOrderItemPage.SelectFirstItemPo();
                purchaseOrderItemPage.AddQuantity("2");
                purchaseOrderItemPage.Validate();
                purchaseOrdersPage = purchaseOrderItemPage.BackToList();

            }
            var firstPurchaseOrderNumber = purchaseOrdersPage.GetFirstPurchaseOrderNumber();
            var Site = purchaseOrdersPage.GetFirstPurchaseOrderSite();
            purchaseOrdersPage.Filter(PurchaseOrdersPage.FilterType.ByNumber, firstPurchaseOrderNumber);
            var dateFrom = purchaseOrdersPage.GetFirstDeliveryDate().AddMonths(-1);
            var dateTo = dateFrom.AddMonths(1);
            purchaseOrdersPage.Filter(PurchaseOrdersPage.FilterType.DateFrom, dateFrom);
            purchaseOrdersPage.Filter(PurchaseOrdersPage.FilterType.DateTo, dateTo);
            purchaseOrdersPage.SendResultsByMail();
            MailPage mailPage = purchaseOrdersPage.RedirectToOutlookMailbox();
            mailPage.FillFields_LogToOutlookMailbox_MoreThanOneMonth(email);
            mailPage.ClickOnSpecifiedOutlookMail("OC NEWREST - 1 - Purchase order -");
            bool isPieceJointeOK = mailPage.IsOutlookPieceJointeOK("Purchase order_" + firstPurchaseOrderNumber + " " + Site + ".pdf");
            Assert.IsTrue(isPieceJointeOK, "La pièce jointe n'est pas présente dans le mail.");
            mailPage.DeleteCurrentOutlookMail();
            mailPage.Close();
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_CreationPODepuisSO()
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
            // Génération du Purchase Order
            supplyOrderItemPage.GeneratePurchaseOrder();
            supplyOrderPage.SelectFirstItem();
            supplyOrderItemPage.SelectFirstItem();
            supplyOrderItemPage.AddQuantityPO(itemName, "0");
            supplyOrderItemPage.ResetFilter();
            supplyOrderItemPage.Filter(FilterSupplyItemType.ByName, itemName);
            // Assertion: Vérifier que la quantité est bien mise à 0
            int zeroQuantity = supplyOrderItemPage.GetQuantityPO();
            Assert.AreEqual(0, zeroQuantity, "La quantité n'a pas été mise à 0 correctement.");
            supplyOrderItemPage.SelectFirstItem();
            supplyOrderItemPage.AddQuantityPO(itemName, "1");
            supplyOrderItemPage.Filter(FilterSupplyItemType.ByName, itemName);
            // Assertion: Vérifier que la quantité peut être changée après avoir été mise à 0
            int updatedQuantity = supplyOrderItemPage.GetQuantityPO();
            Assert.AreEqual(1, updatedQuantity, "La quantité n'a pas pu être changée après avoir été mise à 0.");
            supplyOrderItemPage.ValidateApresUpdate();
            Assert.IsTrue(supplyOrderItemPage.CheckValidation(true), String.Format("La validation de la PO a échoué alors qu'aucun bandeau rouge ne devrait être présent."));


        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_Print_details()
        {
            var location = TestContext.Properties["Location"].ToString();
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            PurchaseOrdersPage purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            purchaseOrdersPage.ResetFilters();
            var PurchaseOrderItem = purchaseOrdersPage.SelectFirstItem();
            purchaseOrdersPage.ClearDownloads();
            var reportPage = PurchaseOrderItem.PrintDetails(true);
            var isReportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();
            Assert.IsTrue(isReportGenerated, "L'application n'a pas pu générer le fichier attendu.");

        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_ModifyStatusComment()
        {
            //Prepare
            var site = TestContext.Properties["Site"].ToString();
            var supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();
            string Id = "";
            //arrange
            var homePage = LogInAsAdmin();

            try
            {
                var purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();
                // creer PO
                var createPurchaseOrderPage = purchaseOrdersPage.CreateNewPurchaseOrder();
                createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(-10), true);
                var purchaseOrderItemPage = createPurchaseOrderPage.Submit();
                Id = purchaseOrderItemPage.GetPurchaseOrderNumber();
                var generalInformationPage = purchaseOrderItemPage.ClickOnGeneralInformation();
                // Set comment and status
                generalInformationPage.SetComment("test comment");
                generalInformationPage.SetStatus("Close");
                purchaseOrdersPage = generalInformationPage.BackToList();
                purchaseOrdersPage.ResetFilters();
                // get the PO
                purchaseOrdersPage.Filter(PurchaseOrdersPage.FilterType.ByNumber, Id);
                purchaseOrdersPage.Filter(PurchaseOrdersPage.FilterType.DateFrom, DateUtils.Now.AddDays(-20));
                purchaseOrdersPage.Filter(PurchaseOrdersPage.FilterType.DateTo, DateUtils.Now.AddDays(-10));
                // check
                purchaseOrderItemPage = purchaseOrdersPage.SelectFirstItem();
                purchaseOrderItemPage.ClickOnGeneralInformation();

                var newDeliveryDate = generalInformationPage.GetDeliveryDateValue();
                var newComment = generalInformationPage.GetComment();
                var newStatus = generalInformationPage.GetStatus();

                bool ModifyStatusComment = false;

                if (newComment != string.Empty && newStatus != string.Empty)
                {
                    ModifyStatusComment = true;
                    Assert.IsTrue(ModifyStatusComment, "Il faut pouvoir modifier le statut et mettre un commentaire si la date Delivary date est inférieure à 10j de la date système (aujourd'hui).");
                    Assert.AreEqual(newComment, "test comment", "le commentaire n'a pas modifié ");
                    Assert.AreEqual(newStatus, "Close", "le status n'a pas modifié ");
                }
            }
            finally
            {
                // delete
                var purchaseOrderPage = homePage.GoToPurchasing_PurchaseOrdersPage();

                purchaseOrderPage.ResetFilters();
                purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.ByNumber, Id);
                purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.DateFrom, DateUtils.Now.AddDays(-20));
                purchaseOrderPage.Filter(PurchaseOrdersPage.FilterType.DateTo, DateUtils.Now.AddDays(-10));
                var number = purchaseOrderPage.DeleteFirstPurchaseOrder();
                Assert.IsTrue(purchaseOrderPage.IsDeleted(number), "purchase order non supprimé");

            }

        }
        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_ItemReference()
        {
            // Prepare
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string ArticleCode = "1234"; // Reference code to validate
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();
            var location = TestContext.Properties["Location"].ToString();
            string Quantity = "10";
            string itemName = itemNameToday + "-" + new Random().Next().ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var purchaseOrderPage = homePage.GoToPurchasing_ItemPage();
            purchaseOrderPage.ResetFilter();

            var itemCreateModalPage = purchaseOrderPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);

            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
            itemCreatePackagingPage.FillFieldReference(ArticleCode); 
            itemGeneralInformationPage.BackToList();

            // Act
            var purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            var createPurchaseOrderPage = purchaseOrdersPage.CreateNewPurchaseOrder();
            createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+10), true);
            var purchaseOrderItemPage = createPurchaseOrderPage.Submit();

            purchaseOrderItemPage.Filter(FilterItemType.ByName, itemName);
            purchaseOrderItemPage.SelectFirstItemPo();
            purchaseOrderItemPage.AddQuantity(Quantity);

            string actualNameValue = purchaseOrderItemPage.GetItemNameValue(); 

            bool IsReferenceVisibleName = actualNameValue.Contains(ArticleCode);

            // Assert that the 'Name' field contains the ArticleCode
            Assert.IsTrue(IsReferenceVisibleName, $"The Name field does not contain the expected ArticleCode: {ArticleCode}");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_TexteCommentaire()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string location = TestContext.Properties["Location"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            PurchaseOrdersPage purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();

            CreatePurchaseOrderModalPage createPurchaseOrderPage = purchaseOrdersPage.CreateNewPurchaseOrder();
            createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(10), true);
            PurchaseOrderItem purchaseOrderItemPage = createPurchaseOrderPage.Submit();
            string ID = purchaseOrderItemPage.GetPurchaseOrderNumber();
            purchaseOrderItemPage.SelectFirstItemPo();
            purchaseOrderItemPage.AddQuantity("2");
            purchaseOrderItemPage.Validate();
            purchaseOrdersPage = purchaseOrderItemPage.BackToList();
            
            purchaseOrdersPage.Filter(PurchaseOrdersPage.FilterType.ByNumber, ID);
            purchaseOrdersPage.SelectFirstItem();
            purchaseOrderItemPage.selectCommentItem();
            bool textSizeIs12px = purchaseOrderItemPage.checkSizeTextComment();

            //Assert
            Assert.IsTrue(textSizeIs12px, "La taille du texte du commentaire doit être égale à 12 px");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_FilterSortByNumber()
        {
            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            PurchaseOrdersPage purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            purchaseOrdersPage.ResetFilters();
            
            if (!purchaseOrdersPage.isPageSizeEqualsTo100())
            {
                purchaseOrdersPage.PageSize("8");
                purchaseOrdersPage.PageSize("100");
            }
            purchaseOrdersPage.Filter(PurchaseOrdersPage.FilterType.SortBy, "Number");

            bool isSortedByNumber = purchaseOrdersPage.IsSortedByNumber();

            //Assert
            Assert.IsTrue(isSortedByNumber, String.Format(MessageErreur.FILTRE_ERRONE, "Sort by Number"));

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PU_PO_FilterDeliveryLocation()
        {
            string site = TestContext.Properties["SiteACE"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string location = TestContext.Properties["Location"].ToString();
            string locationDelivery = "ACE-Economato";
            DateTime dateDelivery = DateUtils.Now.AddDays(10);
            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            PurchaseOrdersPage purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            CreatePurchaseOrderModalPage createPurchaseOrderPage = purchaseOrdersPage.CreateNewPurchaseOrder();
            createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, dateDelivery, true);
            PurchaseOrderItem purchaseOrderItemPage = createPurchaseOrderPage.Submit();
            purchaseOrderItemPage.BackToList();
            purchaseOrdersPage.ResetFilters();
            purchaseOrdersPage.Filter(PurchaseOrdersPage.FilterType.DeliveryLocation, locationDelivery);
            bool purchaseDeliveryLocation =purchaseOrdersPage.VerifyDeliveryLoction(location);
            Assert.IsTrue(purchaseDeliveryLocation, "erreur de filtrage By Delivery Location");

        }

    }
}