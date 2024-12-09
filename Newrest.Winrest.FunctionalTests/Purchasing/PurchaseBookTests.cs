using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Item;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.ReceiptNotes;
using Newrest.Winrest.FunctionalTests.Utils;
using System;
using System.Collections.Generic;
using System.IO;

namespace Newrest.Winrest.FunctionalTests.Purchasing
{
    [TestClass]
    public class PurchaseBook : TestBase
    {
        private const int _timeout = 600000;
        public const string SHEET1 = "Sheet 1";
        private readonly string itemNameToday = "itemPUBO-" + DateUtils.Now.ToString("dd/MM/yyyy");
        private readonly string menuNameToday = "MenuPUBO-" + DateUtils.Now.ToString("dd/MM/yyyy");
        private readonly string recipeNameToday = "RecipePUBO-" + DateUtils.Now.ToString("dd/MM/yyyy");

        /// <summary>
        /// 
        /// Mise en place du paramétrage pour la configuration Winrest 4.0 
        /// 
        /// </summary>
        [TestMethod]
        [Priority(0)]
        [Timeout(_timeout)]
        public void PU_PUBO_SetConfigWinrest4_0()
        {
            //Prepare
            var keyword = TestContext.Properties["Item_Keyword"].ToString();
            string itemKeyword = TestContext.Properties["Item_Keyword"].ToString();
            var site = TestContext.Properties["Site"].ToString();
            string delivery = TestContext.Properties["PlaceFrom"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();

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
            itemPage.ResetFilter();
            if (itemPage.CheckTotalNumber() == 0)
            {
                ItemCreateModalPage itemModal = itemPage.ItemCreatePage();
                ItemGeneralInformationPage generalInfo = itemModal.FillField_CreateNewItem("Test Item - " + new Random().Next(), group, workshop, taxType, prodUnit);
                itemPage = generalInfo.BackToList();
            }
            var itemGeneralInformation = itemPage.ClickOnFirstItem();

            var itemKeywordTab = itemGeneralInformation.ClickOnKeywordItem();

            itemKeywordTab.AddKeyword(itemKeyword);
            itemPage = itemGeneralInformation.BackToList();

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
            if (receiptNotesPage.CheckTotalNumber() == 0)
            {
                var receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
                receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, delivery));
                ReceiptNotesItem receiptNotesItemNew = receiptNotesCreateModalpage.Submit();
                receiptNotesPage = receiptNotesItemNew.BackToList();
            }

            ReceiptNotesItem receiptNotesItem = receiptNotesPage.SelectFirstReceiptNoteItem();

            Assert.IsTrue(receiptNotesItem.IsGroupDisplayActive(), "Le paramètre 'NewGroupDisplay' n'est pas activé.");

        }

        [Priority(1)]
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_PUBO_Select_All_Sites()
        {
            //Prepare
            var adminName = TestContext.Properties["Admin_UserName"].ToString();
            string userName = adminName.Substring(0, adminName.IndexOf("@"));

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var parametersUser = homePage.GoToParameters_User();
            parametersUser.SearchAndSelectUser(userName);
            parametersUser.ClickOnAffectedSite();
            parametersUser.SelectUnselectAllSites(true);

            // On vide le répertoire de téléchargement
            DeleteAllFileDownload();
        }

        //Création d'un nouveau site (avec droits)
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_PUBO_Create_New_Site_With_Permission_NewVersion()
        {
            //Prepare              
            string siteNameCode = "ZZZ" + GenerateName(6).Substring(0, 5).ToUpper();
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
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            bool newVersionPrint = true;

            // Create a new site and affect to the user
            CreateAndAffectNewSite(homePage, siteNameCode);

            try
            {
                // Create a new item and add a new packaging      
                var itemPage = homePage.GoToPurchasing_ItemPage();
                itemPage.ResetFilter();
                itemPage.ClearDownloads();
                itemPage.ClosePrintButton();

                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier, null, supplierRef);
                itemPage = itemGeneralInformationPage.BackToList();

                // Export Item File
                itemPage.Filter(ItemPage.FilterType.Search, itemName.ToString());
                itemPage.Filter(ItemPage.FilterType.Site, site);
                itemPage.Filter(ItemPage.FilterType.SupplierRef, supplierRef);
                itemPage.Export(newVersionPrint);

                DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
                FileInfo[] taskFiles = taskDirectory.GetFiles();

                var correctDownloadedFile = itemPage.GetExportExcelFile(taskFiles);
                Assert.IsNotNull(correctDownloadedFile);

                var fileName = correctDownloadedFile.Name;
                var filePath = Path.Combine(downloadsPath, fileName);

                // Modification du fichier Excel
                OpenXmlExcel.WriteDataInCell("Action", "Supplier", supplier, SHEET1, filePath, "A", CellValues.SharedString);
                OpenXmlExcel.WriteDataInCell("Site", "Supplier", supplier, SHEET1, filePath, siteNameCode, CellValues.SharedString);
                OpenXmlExcel.WriteDataInCell("Price id", "Supplier", supplier, SHEET1, filePath, "", CellValues.String);

                // Import the modified file
                WebDriver.Navigate().Refresh();
                var importPopup = itemPage.Import();
                importPopup.ImportFile(correctDownloadedFile.FullName);

                //ASSERT
                itemPage.Filter(ItemPage.FilterType.Search, itemName.ToString());
                itemPage.ClickOnFirstItem();
                Assert.IsTrue(itemGeneralInformationPage.VerifyPackagingSite(siteNameCode));
            }
            finally
            {
                DeactivateSite(homePage, siteNameCode);
            }
        }

        //Création d'un nouveau site (sans droits)
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_PUBO_Create_New_Site_Without_Permission_NewVersion()
        {
            //Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            //Prepare Create a site
            Random r = new Random();
            string number = r.Next(1, 200).ToString();
            string siteZipCode = r.Next(10000, 99999).ToString();

            string siteNameCode = "ZZZ" + GenerateName(6).Substring(0, 5).ToUpper();
            string siteAddress = number + " calle de " + GenerateName(10);

            var cities = new List<string> { "Barcelona", "Madrid", "San Sebastian", "Bilbao", "Cadiz", "Majorqua", "Albacete", "Reus", "Sevilla", "Granada", "Ourense", "Zaragoza" };
            int index = r.Next(cities.Count);
            string siteCity = cities[index];

            //Prepare create new item
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

            bool newVersionPrint = true;

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var parametersSites = homePage.GoToParameters_Sites();
            var sitesModalPage = parametersSites.ClickOnNewSite();
            sitesModalPage.FillPrincipalField_CreationNewSite(siteNameCode, siteAddress, siteZipCode, siteCity);

            try
            {
                //Act
                var itemPage = homePage.GoToPurchasing_ItemPage();
                itemPage.ResetFilter();
                if (newVersionPrint)
                {
                    itemPage.ClearDownloads();
                }

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
                Assert.IsNotNull(correctDownloadedFile);

                var fileName = correctDownloadedFile.Name;
                var filePath = Path.Combine(downloadsPath, fileName);

                //string columnName, string sheetName, string dataFileName, string value
                OpenXmlExcel.WriteDataInCell("Action", "Supplier", supplier, SHEET1, filePath, "A", CellValues.SharedString);
                OpenXmlExcel.WriteDataInCell("Site", "Supplier", supplier, SHEET1, filePath, siteNameCode, CellValues.SharedString);
                OpenXmlExcel.WriteDataInCell("Price id", "Supplier", supplier, SHEET1, filePath, "", CellValues.String);

                WebDriver.Navigate().Refresh();
                var importPopup = itemPage.Import();
                importPopup.CheckFile(correctDownloadedFile.FullName);

                //ASSERT
                Assert.IsTrue(importPopup.GetErrorMessageNoRights(siteNameCode));
                importPopup.ClosePopup();
            }
            finally
            {
                DeactivateSite(homePage, siteNameCode);
            }
        }

        //Création d'un nouvel article
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_PUBO_Create_New_Item_NewVersion()
        {
            //Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString(); // C:\ChromeDriverDownloads

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

            string itemNameChanged = itemName + " _Changed_From_My_Export_Import";

            bool newVersionPrint = true;

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

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
            itemPage.Filter(ItemPage.FilterType.Site, site);
            itemPage.Filter(ItemPage.FilterType.SupplierRef, supplierRef);

            itemPage.Export(newVersionPrint);

            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = itemPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string sheetName, string dataFileName, string value
            OpenXmlExcel.WriteDataInCell("Action", "Supplier", supplier, SHEET1, filePath, "A", CellValues.SharedString);
            OpenXmlExcel.WriteDataInCell("Item id", "Supplier", supplier, SHEET1, filePath, "", CellValues.String);
            OpenXmlExcel.WriteDataInCell("Price id", "Supplier", supplier, SHEET1, filePath, "", CellValues.String);
            OpenXmlExcel.WriteDataInCell("Item name", "Supplier", supplier, SHEET1, filePath, itemNameChanged, CellValues.SharedString);

            WebDriver.Navigate().Refresh();
            var importPopup = itemPage.Import();
            importPopup.ImportFile(correctDownloadedFile.FullName);

            //Assert
            WebDriver.Navigate().Refresh();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemNameChanged);
            Assert.AreEqual(itemNameChanged, itemPage.GetFirstItemName());
        }

        //Créer un nouvel article avec MAIN à FALSE
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_PUBO_Create_New_Item_Main_False_NewVersion()
        {
            //Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString(); // C:\ChromeDriverDownloads

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

            string itemNameChanged = itemName + " _Changed_From_My_Export_Import";

            bool newVersionPrint = true;

            //Arrange
            var homePage = LogInAsAdmin();


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
            itemPage.Filter(ItemPage.FilterType.Site, site);
            itemPage.Filter(ItemPage.FilterType.SupplierRef, supplierRef);

            itemPage.Export(newVersionPrint);


            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = itemPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string sheetName, string dataFileName, string value
            OpenXmlExcel.WriteDataInCell("Action", "Supplier", supplier, SHEET1, filePath, "A", CellValues.SharedString);
            OpenXmlExcel.WriteDataInCell("Item id", "Supplier", supplier, SHEET1, filePath, "", CellValues.String);
            OpenXmlExcel.WriteDataInCell("Price id", "Supplier", supplier, SHEET1, filePath, "", CellValues.String);
            OpenXmlExcel.WriteDataInCell("Item name", "Supplier", supplier, SHEET1, filePath, itemNameChanged, CellValues.SharedString);
            OpenXmlExcel.WriteDataInCell("Is main supplier", "Supplier", supplier, SHEET1, filePath, "FAUX", CellValues.Boolean);

            WebDriver.Navigate().Refresh();
            var importPopup = itemPage.Import();
            importPopup.ImportFile(correctDownloadedFile.FullName);

            //Assert
            WebDriver.Navigate().Refresh();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemNameChanged.ToString());
            itemPage.ClickOnFirstItem();
            bool check = itemGeneralInformationPage.IsMainChecked(supplier);
            Assert.IsTrue(check,"l'article n'est pas crée a main true");
        }

        //Désactiver l'unique packaging d'un article  qui a du stock
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_PUBO_Deactivate_Item_With_Stock_NewVersion()
        {
            //Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString(); // C:\ChromeDriverDownloads
            string itemName = "itemPUBO-" + DateUtils.Now.ToString("dd/MM/yyyy") + "-" + new Random().Next().ToString();
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string supplierRef = itemName + "_SupplierRef";
            string storageQty = "10";
            string qty = "10";
            string delivery = TestContext.Properties["PlaceFrom"].ToString();
            bool newVersionPrint = true;

            //Arrange
            var homePage = LogInAsAdmin();
            
            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.ClearDownloads();

            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier, null, supplierRef);


            var ReceiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
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
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            itemPage.Export(newVersionPrint);


            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = itemPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string sheetName, string dataFileName, string value
            OpenXmlExcel.WriteDataInCell("Action", "Supplier", supplier, SHEET1, filePath, "M", CellValues.SharedString);
            OpenXmlExcel.WriteDataInCell("Is active", "Supplier", supplier, SHEET1, filePath, "FAUX", CellValues.Boolean);
            OpenXmlExcel.WriteDataInCell("Is main supplier", "Supplier", supplier, SHEET1, filePath, "FAUX", CellValues.Boolean);

            WebDriver.Navigate().Refresh();
            var importPopup = itemPage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);
            importPopup.Import();

            bool errorMessagePackagingWithStock = importPopup.GetErrorMessagePackagingWithStock();
            //Assert
            Assert.IsTrue(errorMessagePackagingWithStock);
        }

        //Activer l'unique packaging d'un article en mettant le Main à False
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_PUBO_Activate_Packaging_Main_False_NewVersion()
        {
            //Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString(); // C:\ChromeDriverDownloads

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
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

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
            itemPage.Filter(ItemPage.FilterType.ShowOnlyActive, true);
            itemPage.Filter(ItemPage.FilterType.Search, itemName.ToString());
            itemPage.Filter(ItemPage.FilterType.Site, site);
            itemPage.Filter(ItemPage.FilterType.SupplierRef, supplierRef);

            itemPage.Export(newVersionPrint);

            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = itemPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            OpenXmlExcel.WriteDataInCell("Action", "Supplier", supplier, SHEET1, filePath, "M", CellValues.SharedString);
            OpenXmlExcel.WriteDataInCell("Is main supplier", "Supplier", supplier, SHEET1, filePath, "FAUX", CellValues.Boolean);
            OpenXmlExcel.WriteDataInCell("Is active", "Supplier", supplier, SHEET1, filePath, "VRAI", CellValues.Boolean);

            //string columnName, string sheetName, string dataFileName, string value           
            itemPage.ClickOnFirstItem();

            itemGeneralInformationPage.ClickOnDeleteItem();
            itemGeneralInformationPage.ConfirmDelete();

            Assert.IsFalse(itemGeneralInformationPage.IsPackagingVisible());

            itemGeneralInformationPage.BackToList();

            var importPopup = itemPage.Import();
            importPopup.ImportFile(correctDownloadedFile.FullName);

            WebDriver.Navigate().Refresh();
            itemPage.Filter(ItemPage.FilterType.ShowOnlyInactive, true);
            itemPage.Filter(ItemPage.FilterType.ShowOnlyActive, true);
            itemPage.Filter(ItemPage.FilterType.Search, itemName.ToString());
            itemPage.ClickOnFirstItem();

            //Assert          
            Assert.IsTrue(itemGeneralInformationPage.IsPackagingVisible());
            Assert.IsTrue(itemGeneralInformationPage.IsMainChecked(supplier));
        }

        //Importer un deuxième packaging
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_PUBO_Import_Second_Packaging_NewVersion()
        {
            //Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString(); // C:\ChromeDriverDownloads

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
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

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
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string idColumnName, string id, string sheetName, string fileName, string value
            OpenXmlExcel.WriteDataInCell("Action", "Supplier", supplier, SHEET1, filePath, "A", CellValues.SharedString);
            OpenXmlExcel.WriteDataInCell("Is main supplier", "Supplier", supplier, SHEET1, filePath, "VRAI", CellValues.Boolean);
            OpenXmlExcel.WriteDataInCell("Price id", "Supplier", supplier, SHEET1, filePath, "", CellValues.String);
            OpenXmlExcel.WriteDataInCell("Supplier", "Supplier", supplier, SHEET1, filePath, "A REFERENCIAR", CellValues.SharedString);

            WebDriver.Navigate().Refresh();
            var importPopup = itemPage.Import();
            importPopup.ImportFile(correctDownloadedFile.FullName);

            //Assert
            itemPage.Filter(ItemPage.FilterType.Search, itemName.ToString());
            itemGeneralInformationPage = itemPage.ClickOnFirstItem();
            Assert.IsFalse(itemGeneralInformationPage.IsMainChecked(supplier));
            Assert.IsTrue(itemGeneralInformationPage.IsMainChecked("A REFERENCIAR"));
        }

        //Ajouter un article avec un packaging inexistant
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_PUBO_Change_Packaging_Unit_Fail_NewVersion()
        {
            //Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString(); // C:\ChromeDriverDownloads

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

            string newPackagingUnit = "ABCDE";

            //Arrange
            var homePage = LogInAsAdmin();


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
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string idColumnName, string id, string sheetName, string fileName, string value
            OpenXmlExcel.WriteDataInCell("Action", "Supplier", supplier, SHEET1, filePath, "M", CellValues.SharedString);
            OpenXmlExcel.WriteDataInCell("Packaging", "Supplier", supplier, SHEET1, filePath, newPackagingUnit, CellValues.SharedString);

            WebDriver.Navigate().Refresh();
            var importPopup = itemPage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);
            bool check = importPopup.GetErrorMessagePackagingUnit(newPackagingUnit);
            //Assert
            Assert.IsTrue(check,"Importation reussite avec packing unit doesnt exist");
        }

        //Exporter les données de la base suivant les sites affectés
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_PUBO_Export_Items_Site_With_And_Without_Permission_NewVersion()
        {
            //Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            var adminName = TestContext.Properties["Admin_UserName"].ToString();

            //Prepare Create a site
            string number = new Random().Next(1, 200).ToString();
            string siteZipCode = new Random().Next(10000, 99999).ToString();
            string userName = adminName.Substring(0, adminName.IndexOf("@"));

            string siteNameCode = "ZZZ" + GenerateName(6).Substring(0, 5).ToUpper();
            string siteAddress = number + " calle de " + GenerateName(10);

            var cities = new List<string> { "Barcelona", "Madrid", "San Sebastian", "Bilbao", "Cadiz", "Majorqua", "Albacete", "Reus", "Sevilla", "Granada", "Ourense", "Zaragoza" };
            int index = new Random().Next(cities.Count);
            string siteCity = cities[index];

            //Prepare create new item
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
            bool newVersionPrint = true;

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();


            var parametersSites = homePage.GoToParameters_Sites();
            var sitesModalPage = parametersSites.ClickOnNewSite();
            parametersSites = sitesModalPage.FillPrincipalField_CreationNewSite(siteNameCode, siteAddress, siteZipCode, siteCity);
            parametersSites.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, siteNameCode);
            var id = parametersSites.CollectNewSiteID();

            var parametersUser = homePage.GoToParameters_User();
            parametersUser.SearchAndSelectUser(userName);
            parametersUser.ClickOnAffectedSite();
            parametersUser.GiveSiteRightsToUser(id, true, siteNameCode);

            try
            {
                //Act
                var itemPage = homePage.GoToPurchasing_ItemPage();
                itemPage.ResetFilter();

                if (newVersionPrint)
                {
                    itemPage.ClearDownloads();
                }

                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier, null, supplierRef);
                itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemCreatePackagingPage.FillField_CreateNewPackaging(siteNameCode, packagingName, storageQty, storageUnit, qty, supplier, null, supplierRef);

                parametersUser = homePage.GoToParameters_User();
                parametersUser.SearchAndSelectUser(userName);
                parametersUser.ClickOnAffectedSite();
                parametersUser.GiveSiteRightsToUser(id, false, siteNameCode);

                itemPage = homePage.GoToPurchasing_ItemPage();
                itemPage.Filter(ItemPage.FilterType.Search, itemName.ToString());
                itemPage.Export(newVersionPrint);

                DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
                FileInfo[] taskFiles = taskDirectory.GetFiles();

                // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
                var correctDownloadedFile = itemPage.GetExportExcelFile(taskFiles);
                Assert.IsNotNull(correctDownloadedFile);

                var fileName = correctDownloadedFile.Name;
                var filePath = Path.Combine(downloadsPath, fileName);

                //string columnName, string sheetName, string fileName, string value
                var resultSite = OpenXmlExcel.ReadAllDataInColumn("Site", SHEET1, filePath, site);
                var resultSiteNameCode = OpenXmlExcel.ReadAllDataInColumn("Site", SHEET1, filePath, siteNameCode);

                //ASSERT
                Assert.IsTrue(resultSite);
                Assert.IsFalse(resultSiteNameCode);
            }
            finally
            {
                DeactivateSite(homePage, siteNameCode);
            }
        }

        //Exporter les données d'1 seul site
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_PUBO_Export_Items_From_Site_NewVersion()
        {
            //Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString(); // C:\ChromeDriverDownloads
            string site = TestContext.Properties["SiteBis"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            bool newVersionPrint = true;

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.ClearDownloads();

            //Filter Proccess
            itemPage.Filter(ItemPage.FilterType.Site, site);

            //Export Proccess
            itemPage.Export(newVersionPrint);
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = itemPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun Fichier Exporté");
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber(SHEET1, filePath);
            bool resultSite = itemPage.IsAllResultsConatinsSameSite("Site", SHEET1, correctDownloadedFile.FullName, site);

            //Assert
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.IsTrue(resultSite, MessageErreur.EXCEL_DONNEES_KO);
        }

        //Exporter les données d'1 fournisseur sur plusieurs sites
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_PUBO_Export_Items_From_Sites_And_Supplier_NewVersion()
        {
            //Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString(); // C:\ChromeDriverDownloads
            string siteMAD = TestContext.Properties["Site"].ToString();
            string siteBCN = "BCN";
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string group = "PLATOS CALIENTES CONG";

            //Arrange
            var homePage = LogInAsAdmin();

            bool newVersionPrint = true;

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
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
            Assert.AreNotEqual(resultNumber, 0);
            Assert.IsTrue(resultSupplier);
        }

        //Exporter les données d'1 groupe d'article
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_PUBO_Export_Items_From_Site_And_Group_NewVersion()
        {
            //Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString(); // C:\ChromeDriverDownloads
            string site = TestContext.Properties["Site"].ToString();
            string group = TestContext.Properties["Item_Group"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

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
            Assert.AreNotEqual(resultNumber, 0);
            Assert.IsTrue(resultSite);
            Assert.IsTrue(resultGroup);
        }

        //Exporter les packaging inactifs
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_PUBO_Export_Inactive_Items_NewVersion()
        {
            //Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString(); // C:\ChromeDriverDownloads

            //Arrange
            var homePage = LogInAsAdmin();

            bool newVersionPrint = true;

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.ClearDownloads();

            itemPage.Filter(ItemPage.FilterType.ShowOnlyInactive, true);
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
            bool resultInactive = OpenXmlExcel.ReadAllDataInColumn("Is active", SHEET1, filePath, "Faux");

            //Assert
            Assert.AreNotEqual(resultNumber, 0);
            Assert.IsTrue(resultInactive);
        }

        //Exporter les données ALL SITES/FOURNISSEUR/GROUP + show inactive item avec un seul site affecté (PMI)
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_PUBO_Export_All_With_One_Site_Affected_NewVersion()
        {
            //FIXME radiobutton ne marche plus : "show inactive item" == items avec packaging + inactive, hors les items avec packaging sont forcément active
            //je met un FIXME,
            //et je met un ShowOnlyActive à la place de ShowOnlyInactive,
            //vue que l'objectif du test est surtout le cas de l'utilisateur avec une seule autorisation de site

            //Prepare
            Random random = new Random();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            var adminName = TestContext.Properties["Admin_UserName"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            string userName = adminName.Substring(0, adminName.IndexOf("@"));

            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();

            string itemName = "Test-export-site-affected-" + random.Next(0, 1000);

            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string supplierRef = itemName + "_SupplierRef";
            string storageQty = 10.ToString();
            string qty = 10.ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            bool newVersionPrint = true;

            try
            {
                var parametersUser = homePage.GoToParameters_User();
                parametersUser.SearchAndSelectUser(userName);
                parametersUser.ClickOnAffectedSite();
                parametersUser.SelectUnselectAllSites(false);
                parametersUser.SelectOneSite(site); //= MAD

                //WebDriver.Navigate().Refresh();

                //Act
                var itemPage = homePage.GoToPurchasing_ItemPage();
                itemPage.ResetFilter();
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                itemPage.Filter(ItemPage.FilterType.ShowOnlyActive, true);

                if (itemPage.CheckTotalNumber() == 0)
                {
                    ItemCreateModalPage itemCreate = itemPage.ItemCreatePage();
                    ItemGeneralInformationPage generalInfo = itemCreate.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
                    //generalInfo.SetActivated(false);
                    var itemCreatePackagingPage = generalInfo.NewPackaging();
                    generalInfo = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier, null, supplierRef);
                    itemPage = generalInfo.BackToList();
                }
                itemPage.ClearDownloads();
                itemPage.ResetFilter();
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                itemPage.Filter(ItemPage.FilterType.ShowOnlyActive, true);
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
                Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);
                Assert.IsTrue(resultSite, MessageErreur.EXCEL_DONNEES_KO);
            }
            finally
            {
                var parametersUsr = homePage.GoToParameters_User();
                parametersUsr.SearchAndSelectUser(userName);
                parametersUsr.ClickOnAffectedSite();
                parametersUsr.SelectUnselectAllSites(true);
            }
        }

        //duplication de packaging sur autre site déjà ayant le meme packaging
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_PUBO_Duplicate_Packaging_Already_Existing_NewVersion()
        {
            //Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString(); // C:\ChromeDriverDownloads

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

            string site2 = "ALC";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            bool newVersionPrint = true;

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.ClearDownloads();

            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier, null, supplierRef);
            itemGeneralInformationPage.DuplicateItem("ALC");
            itemPage = itemGeneralInformationPage.BackToList();

            itemPage.Filter(ItemPage.FilterType.Search, itemName.ToString());
            itemPage.Export(newVersionPrint);


            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = itemPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string idColumnName, string id, string sheetName, string fileName, string value
            OpenXmlExcel.WriteDataInCell("Action", "Site", site2, SHEET1, filePath, "M", CellValues.SharedString);
            OpenXmlExcel.WriteDataInCell("Site", "Site", site2, SHEET1, filePath, site, CellValues.SharedString);

            WebDriver.Navigate().Refresh();
            var importPopup = itemPage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);
            importPopup.Import();
            //Assert
            Assert.IsTrue(importPopup.GetErrorMessagePackagingAlreadyExisting());
        }

        //Désactiver l'unique packaging d'un article utilisé dans les recettes
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_PUBO_Deactivate_Packaging_From_Recipe_NewVersion()
        {
            //Prepare Item
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString(); // C:\ChromeDriverDownloads

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

            //Prepare recipe
            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            bool newVersionPrint = true;

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.ClearDownloads();

            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier, null, supplierRef);

            //Create recipe and menu
            CreateRecipe(homePage, recipeName, site, itemName.ToString(), null, null);

            itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.Filter(ItemPage.FilterType.Search, itemName.ToString());
            itemPage.Export(newVersionPrint);


            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = itemPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string sheetName, string dataFileName, string value
            OpenXmlExcel.WriteDataInCell("Action", "Supplier", supplier, SHEET1, filePath, "M", CellValues.SharedString);
            OpenXmlExcel.WriteDataInCell("Is active", "Supplier", supplier, SHEET1, filePath, "FAUX", CellValues.Boolean);
            OpenXmlExcel.WriteDataInCell("Is main supplier", "Supplier", supplier, SHEET1, filePath, "FAUX", CellValues.Boolean);

            WebDriver.Navigate().Refresh();
            var importPopup = itemPage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);
            importPopup.Import();
            //Assert
            var errorMessage = importPopup.GetErrorMessagePackagingUsedInRecipes();
            Assert.IsTrue(errorMessage);
        }

        //Changer le prix d'un article utilisé dans une recette
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_PUBO_Change_Item_Price_From_Recipe_NewVersion()
        {
            //Prepare Item
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString(); // C:\ChromeDriverDownloads
            var currency = TestContext.Properties["Currency"].ToString();
            string itemName = "itemPUBO-" + DateUtils.Now.ToString("dd/MM/yyyy") + "-" + new Random().Next().ToString();
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string supplierRef = itemName + "_SupplierRef";
            string storageQty = "10";
            string qty = "10";
            //Prepare recipe
            string recipeName = "RecipePUBO-" + DateUtils.Now.ToString("dd/MM/yyyy") + "-" + new Random().Next().ToString();
            double priceChanged = 2;
            bool newVersionPrint = true;

            //Arrange
            var homePage = LogInAsAdmin();
            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.ClearDownloads();

            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier, null, supplierRef);

            //Create recipe and menu
            double getInitPriceRecipe = CreateRecipe(homePage, recipeName, site, itemName, currency, decimalSeparatorValue);

            itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            itemPage.Export(newVersionPrint);

            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = itemPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string sheetName, string dataFileName, string value
            OpenXmlExcel.WriteDataInCell("Action", "Supplier", supplier, SHEET1, filePath, "M", CellValues.SharedString);
            OpenXmlExcel.WriteDataInCell("Packing Price", "Supplier", supplier, SHEET1, filePath, priceChanged.ToString(), CellValues.SharedString);

            WebDriver.Navigate().Refresh();
            var importPopup = itemPage.Import();
            importPopup.ImportFile(correctDownloadedFile.FullName);

            // Récupération des prix des recettes et menu
            double getChangedPriceRecipe = GetRecipePrice(homePage, recipeName, currency, decimalSeparatorValue);

            //Assert
            Assert.AreNotEqual(getInitPriceRecipe, getChangedPriceRecipe);
        }

        //Changer le MAIN d'un article utilisé dans une recette 
        //(l'article doit avoir des prix différents sur chaque MAIN)
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_PUBO_Change_Item_Main_From_Recipe_NewVersion()
        {
            //Prepare Item
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            var currency = TestContext.Properties["Currency"].ToString();
            string itemName = "itemPUBO-" + DateUtils.Now.ToString("dd/MM/yyyy") + " - " + new Random().Next().ToString();
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string supplierRef = itemName + "_SupplierRef";
            string storageQty = "2";
            string qty = "4";
            string supplier2 = "A REFERENCIAR";

            //Prepare recipe
            string recipeName = "RecipePUBO-" + DateUtils.Now.ToString("dd/MM/yyyy") + " - " + new Random().Next().ToString();
            string priceChanged = "1";

            //Arrange
            var homePage = LogInAsAdmin();

            bool newVersionPrint = true;
            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.ClearDownloads();

            // Create Item and two Packaging 
            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier, null, supplierRef);
            itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier2, null, supplierRef, priceChanged);

            //Create recipe and menu
            double getInitPriceRecipe = CreateRecipe(homePage, recipeName, site, itemName, currency, decimalSeparatorValue);

            // Go to Purchasing Item Page and Filter
            itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);

            itemPage.ClickOnPurchaseTab();

            // Export
            itemPage.Export(newVersionPrint);


            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = itemPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string sheetName, string dataFileName, string value
            OpenXmlExcel.WriteDataInCell("Action", "Supplier", supplier, SHEET1, filePath, "M", CellValues.SharedString);
            OpenXmlExcel.WriteDataInCell("Is main supplier", "Supplier", supplier, SHEET1, filePath, "FAUX", CellValues.Boolean);

            WebDriver.Navigate().Refresh();
            var importPopup = itemPage.Import();
            importPopup.ImportFile(correctDownloadedFile.FullName);

            // Récupération des prix des recettes et menu
            double getChangedPriceRecipe = GetRecipePrice(homePage, recipeName, currency, decimalSeparatorValue);
            // Assert
            Assert.AreNotEqual(getInitPriceRecipe, getChangedPriceRecipe);
        }

        //Changer le prix d'un article utilisé dans une recette qui apparait dans un menu
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_PUBO_Change_Item_Price_From_Menu_NewVersion()
        {
            //Prepare Item
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString(); // C:\ChromeDriverDownloads
            var currency = TestContext.Properties["Currency"].ToString();

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
            string storageQty = 2.ToString();
            string qty = 4.ToString();

            //Prepare recipe
            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            double priceChanged = 1;
            string menuName = menuNameToday + "-" + new Random().Next().ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            bool newVersionPrint = true;

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            //Create item
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.ClearDownloads();

            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier, null, supplierRef);

            //Create recipe and menu
            double getInitPriceRecipe = CreateRecipe(homePage, recipeName, site, itemName.ToString(), currency, decimalSeparatorValue);
            double getInitPriceMenu = CreateMenu(homePage, menuName, recipeName, site, currency, decimalSeparatorValue);

            itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.Filter(ItemPage.FilterType.Search, itemName.ToString());
            itemPage.Filter(ItemPage.FilterType.Site, site);
            itemPage.Filter(ItemPage.FilterType.SupplierRef, supplierRef);
            itemPage.Export(newVersionPrint);

            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = itemPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string sheetName, string dataFileName, string value
            OpenXmlExcel.WriteDataInCell("Action", "Supplier", supplier, SHEET1, filePath, "M", CellValues.SharedString);
            OpenXmlExcel.WriteDataInCell("Packing Price", "Supplier", supplier, SHEET1, filePath, priceChanged.ToString(), CellValues.SharedString);

            WebDriver.Navigate().Refresh();
            var importPopup = itemPage.Import();
            importPopup.ImportFile(correctDownloadedFile.FullName);

            // Récupération des prix des recettes et menu
            double getChangedPriceRecipe = GetRecipePrice(homePage, recipeName, currency, decimalSeparatorValue);
            double getChangedPriceMenu = GetMenuPrice(homePage, menuName, currency, decimalSeparatorValue);

            Assert.AreNotEqual(getInitPriceRecipe, getChangedPriceRecipe);
            Assert.AreNotEqual(getInitPriceMenu, getChangedPriceMenu);
        }

        //Changer le MAIN d'un article utilisé dans une recette qui apparait dans un menu
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_PUBO_Change_Item_Main_From_Menu_NewVersion()
        {
            //Prepare Item
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString(); // C:\ChromeDriverDownloads
            var currency = TestContext.Properties["Currency"].ToString();

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
            string storageQty = 2.ToString();
            string qty = 4.ToString();

            string supplier2 = "A REFERENCIAR";

            //Prepare recipe
            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            double priceChanged = 1;
            string menuName = menuNameToday + "-" + new Random().Next().ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            bool newVersionPrint = true;

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            //Create item
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.ClearDownloads();

            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier, null, supplierRef);
            itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier2, null, supplierRef, priceChanged.ToString());

            //Create recipe and menu
            double getInitPriceRecipe = CreateRecipe(homePage, recipeName, site, itemName.ToString(), currency, decimalSeparatorValue);
            double getInitPriceMenu = CreateMenu(homePage, menuName, recipeName, site, currency, decimalSeparatorValue);

            itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.Filter(ItemPage.FilterType.Search, itemName.ToString());
            itemPage.Filter(ItemPage.FilterType.Site, site);
            itemPage.Filter(ItemPage.FilterType.SupplierRef, supplierRef);

            itemPage.ClickOnPurchaseTab();

            itemPage.Export(newVersionPrint);

            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = itemPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string sheetName, string dataFileName, string value
            OpenXmlExcel.WriteDataInCell("Action", "Supplier", supplier, SHEET1, filePath, "M", CellValues.SharedString);
            OpenXmlExcel.WriteDataInCell("Is main supplier", "Supplier", supplier, SHEET1, filePath, "FAUX", CellValues.Boolean);

            WebDriver.Navigate().Refresh();
            var importPopup = itemPage.Import();
            importPopup.ImportFile(correctDownloadedFile.FullName);

            // Récupération des prix des recettes et menu
            double getChangedPriceRecipe = GetRecipePrice(homePage, recipeName, currency, decimalSeparatorValue);
            double getChangedPriceMenu = GetMenuPrice(homePage, menuName, currency, decimalSeparatorValue);

            Assert.AreNotEqual(getInitPriceRecipe, getChangedPriceRecipe);
            Assert.AreNotEqual(getInitPriceMenu, getChangedPriceMenu);
        }

        //Changer la storage qty ou la pack qty  d'un article utilisé
        //dans une recette qui apparait dans un menu
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_PUBO_Change_Storage_Quantity_From_Menu_NewVersion()
        {
            //Prepare Item
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString(); // C:\ChromeDriverDownloads
            var currency = TestContext.Properties["Currency"].ToString();

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
            string storageQty = 2.ToString();
            string qty = 4.ToString();

            //Prepare recipe
            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            string packagingPrice = 9.ToString();

            double newStorageQty = 6;

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            bool newVersionPrint = true;

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            //Create item
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.ClearDownloads();

            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier, null, supplierRef, packagingPrice);

            //Create recipe and menu
            double getInitPriceRecipe = CreateRecipe(homePage, recipeName, site, itemName.ToString(), currency, decimalSeparatorValue);
            double getInitPriceMenu = CreateMenu(homePage, menuName, recipeName, site, currency, decimalSeparatorValue);

            itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.Filter(ItemPage.FilterType.Search, itemName.ToString());
            itemPage.Filter(ItemPage.FilterType.Site, site);
            itemPage.Filter(ItemPage.FilterType.SupplierRef, supplierRef);
            itemPage.Export(newVersionPrint);

            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = itemPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string sheetName, string dataFileName, string value
            OpenXmlExcel.WriteDataInCell("Action", "Supplier", supplier, SHEET1, filePath, "M", CellValues.SharedString);
            OpenXmlExcel.WriteDataInCell("Storage qty", "Supplier", supplier, SHEET1, filePath, newStorageQty.ToString(), CellValues.SharedString);

            WebDriver.Navigate().Refresh();
            var importPopup = itemPage.Import();
            importPopup.ImportFile(correctDownloadedFile.FullName);

            // Récupération des prix des recettes et menu
            double getChangedPriceRecipe = GetRecipePrice(homePage, recipeName, currency, decimalSeparatorValue);
            double getChangedPriceMenu = GetMenuPrice(homePage, menuName, currency, decimalSeparatorValue);

            Assert.AreNotEqual(getInitPriceRecipe, getChangedPriceRecipe);
            Assert.AreNotEqual(getInitPriceMenu, getChangedPriceMenu);
        }

        //Créer un nouveau packaging MAIN d'un article utilisé 
        //dans une recette qui apparait dans un menu
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_PUBO_Create_New_Item_Main_From_Menu_NewVersion()
        {
            //Prepare Item
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString(); // C:\ChromeDriverDownloads
            var currency = TestContext.Properties["Currency"].ToString();

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
            string storageQty = 2.ToString();
            string qty = 4.ToString();

            string supplier2 = "A REFERENCIAR";

            //Prepare recipe
            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            double priceChanged = 1;
            string menuName = menuNameToday + "-" + new Random().Next().ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            bool newVersionPrint = true;

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            //Create item
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.ClearDownloads();

            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty.ToString(), storageUnit, qty, supplier, null, supplierRef);

            //Create recipe and menu
            double getInitPriceRecipe = CreateRecipe(homePage, recipeName, site, itemName.ToString(), currency, decimalSeparatorValue);
            double getInitPriceMenu = CreateMenu(homePage, menuName, recipeName, site, currency, decimalSeparatorValue);

            itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.Filter(ItemPage.FilterType.Search, itemName.ToString());
            itemPage.Export(newVersionPrint);

            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = itemPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            OpenXmlExcel.WriteDataInCell("Action", "Supplier", supplier, SHEET1, filePath, "A", CellValues.SharedString);
            OpenXmlExcel.WriteDataInCell("Price id", "Supplier", supplier, SHEET1, filePath, "", CellValues.String);
            OpenXmlExcel.WriteDataInCell("Packing Price", "Supplier", supplier, SHEET1, filePath, (priceChanged * 10).ToString(), CellValues.SharedString);
            OpenXmlExcel.WriteDataInCell("Supplier", "Supplier", supplier, SHEET1, filePath, supplier2, CellValues.SharedString);

            WebDriver.Navigate().Refresh();
            var importPopup = itemPage.Import();
            importPopup.ImportFile(correctDownloadedFile.FullName);

            // Récupération des prix des recettes et menu
            double getChangedPriceRecipe = GetRecipePrice(homePage, recipeName, currency, decimalSeparatorValue);
            double getChangedPriceMenu = GetMenuPrice(homePage, menuName, currency, decimalSeparatorValue);

            Assert.AreNotEqual(getInitPriceRecipe, getChangedPriceRecipe);
            Assert.AreNotEqual(getInitPriceMenu, getChangedPriceMenu);
        }

        //Changer l'Unité de Production d'un article 
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_PUBO_Change_Item_Unit_Prod_NewVersion()
        {
            //Prepare Item
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString(); // C:\ChromeDriverDownloads

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
            string storageQty = 2.ToString();
            string qty = 4.ToString();

            string newProdUnit = "WW";

            //Prepare recipe
            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            bool newVersionPrint = true;

            //Act
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.ClearDownloads();

            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier, null, supplierRef);

            //Create recipe and menu
            CreateRecipe(homePage, recipeName, site, itemName.ToString(), null, null);

            itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.Filter(ItemPage.FilterType.Search, itemName.ToString());
            itemPage.Export(newVersionPrint);

            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = itemPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string sheetName, string dataFileName, string value
            OpenXmlExcel.WriteDataInCell("Action", "Supplier", supplier, SHEET1, filePath, "M", CellValues.SharedString);
            OpenXmlExcel.WriteDataInCell("Prod. unit", "Supplier", supplier, SHEET1, filePath, newProdUnit, CellValues.SharedString);

            WebDriver.Navigate().Refresh();
            var importPopup = itemPage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);

            //Assert
            Assert.IsTrue(importPopup.GetMessageErrorProductionUnit());
        }

        //Désactiver un packaging MAIN d'un article possédant N packaging utilisé dans les recettes
        [TestMethod]
        [Timeout(_timeout)]
        public void PU_PUBO_Deactive_N_Packaging_From_Recipe_NewVersion()
        {
            //Prepare Item
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString(); // C:\ChromeDriverDownloads
            var currency = TestContext.Properties["Currency"].ToString();
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
            string storageQty = 2.ToString();
            string qty = 4.ToString();
            string supplier2 = TestContext.Properties["Supplier"].ToString();

            //Prepare recipe
            string recipeName = recipeNameToday + "-" + new Random().Next().ToString();
            string menuName = menuNameToday + "-" + new Random().Next().ToString();
            double priceChanged = Math.Round(new Random().NextDouble() * new Random().Next(1, 10), 2);

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            bool newVersionPrint = true;

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            //Create item
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.ClearDownloads();

            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier, null, supplierRef);
            itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier2, null, supplierRef, priceChanged.ToString());

            //Create recipe and menu
            double getInitPriceRecipe = CreateRecipe(homePage, recipeName, site, itemName.ToString(), currency, decimalSeparatorValue);
            double getInitPriceMenu = CreateMenu(homePage, menuName, recipeName, site, currency, decimalSeparatorValue);

            itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.Filter(ItemPage.FilterType.Search, itemName.ToString());
            itemPage.Export(newVersionPrint);

            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = itemPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string sheetName, string dataFileName, string value
            OpenXmlExcel.WriteDataInCell("Action", "Supplier", supplier, SHEET1, filePath, "M", CellValues.SharedString);
            OpenXmlExcel.WriteDataInCell("Is main supplier", "Supplier", supplier, SHEET1, filePath, "FAUX", CellValues.Boolean);
            OpenXmlExcel.WriteDataInCell("Is active", "Supplier", supplier, SHEET1, filePath, "FAUX", CellValues.Boolean);

            WebDriver.Navigate().Refresh();
            var importPopup = itemPage.Import();
            importPopup.ImportFile(correctDownloadedFile.FullName);

            // Récupération des prix des recettes et menu
            double getChangedPriceRecipe = GetRecipePrice(homePage, recipeName, currency, decimalSeparatorValue);
            double getChangedPriceMenu = GetMenuPrice(homePage, menuName, currency, decimalSeparatorValue);

            Assert.AreNotEqual(getInitPriceRecipe, getChangedPriceRecipe);
            Assert.AreNotEqual(getInitPriceMenu, getChangedPriceMenu);
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void PU_PUBO_Import_Keywords_NewVersion()
        {
            //Prepare
            string itemName = itemNameToday + "-" + new Random().Next().ToString();
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();

            string keyword1 = "keyword1";
            string keyword = "TEST_KEY";
            string keyword2 = "keyword2";
            string keyword3 = "keyword3";

            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string supplierRef = itemName + "_SupplierRef";
            string storageQty = 10.ToString();
            string qty = 10.ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            homePage.ClearDownloads();
            //Act
            //Importer le fichier avec 3 keywords sur un article
            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
            ItemCreateModalPage modal = itemPage.ItemCreatePage();
            ItemGeneralInformationPage generalInfo = modal.FillField_CreateNewItem(itemName, group, workshop, taxType, prodUnit);

            var itemCreatePackagingPage = generalInfo.NewPackaging();
            itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier, null, supplierRef);

            ItemKeywordPage keywords = generalInfo.ClickOnKeywordItem();
            keywords.AddKeyword(keyword);
            generalInfo = keywords.ClickOnGeneralInfo();
            itemPage = generalInfo.BackToList();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            itemPage.Export(true);
            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            FileInfo correctDownloadedFile = itemPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            // passer la ligne à M
            OpenXmlExcel.WriteDataInCell("Action", "Supplier", supplier, SHEET1, correctDownloadedFile.FullName, "M", CellValues.SharedString);

            // retirer les keyword
            generalInfo = itemPage.ClickOnFirstItem();
            keywords = generalInfo.ClickOnKeywordItem();
            keywords.RemoveKeyword(keyword);
            generalInfo = keywords.ClickOnGeneralInfo();
            itemPage = generalInfo.BackToList();

            // Import the modified file
            WebDriver.Navigate().Refresh();
            ItemImportPage importPopup = itemPage.Import();
            importPopup.ImportFile(correctDownloadedFile.FullName);

            //Le fichier est importer avec les 3 keywords sur l'article
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            generalInfo = itemPage.ClickOnFirstItem();
            keywords = generalInfo.ClickOnKeywordItem();
            Assert.IsTrue(keywords.IsKeywordPresent(keyword), "keyword1 non présent");
        }

        // -----------------------------------------------------------------------------------------------------------

        public void CreateAndAffectNewSite(HomePage homePage, string siteNameCode)
        {
            // Prepare
            var adminName = TestContext.Properties["Admin_UserName"].ToString();
            string userName = adminName.Substring(0, adminName.IndexOf("@"));
            bool isPermission = true;

            string number = new Random().Next(1, 200).ToString();
            string siteZipCode = new Random().Next(10000, 99999).ToString();
            string siteAddress = number + " calle de " + GenerateName(10);

            var cities = new List<string> { "Barcelona", "Madrid", "San Sebastian", "Bilbao", "Cadiz", "Majorqua", "Albacete", "Reus", "Sevilla", "Granada", "Ourense", "Zaragoza" };
            int index = new Random().Next(cities.Count);
            string siteCity = cities[index];

            // Create a new Site
            var parametersSites = homePage.GoToParameters_Sites();
            var sitesModalPage = parametersSites.ClickOnNewSite();
            sitesModalPage.FillPrincipalField_CreationNewSite(siteNameCode, siteAddress, siteZipCode, siteCity);
            parametersSites.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, siteNameCode);
            string id = parametersSites.CollectNewSiteID();

            // Affect it to the user
            var parametersUser = homePage.GoToParameters_User();
            parametersUser.SearchAndSelectUser(userName);
            parametersUser.ClickOnAffectedSite();
            parametersUser.GiveSiteRightsToUser(id, true, siteNameCode);
        }

        public void DeactivateSite(HomePage homePage, string siteNameCode)
        {
            // Find the Site to deactivate
            var parametersSites = homePage.GoToParameters_Sites();
            parametersSites.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, siteNameCode);
            parametersSites.Deactivate();
        }

        public double CreateRecipe(HomePage homePage, string recipeName, string site, string itemName, string currency, string decimalSeparatorValue)
        {
            double initPriceRecipe = 0;

            string recipeType = TestContext.Properties["RecipeType"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();
            int nbPortions = 1;

            var recipesPage = homePage.GoToMenus_Recipes();
            var recipesCreateModalPage = recipesPage.CreateNewRecipe();
            var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());

            recipeGeneralInfosPage.AddVariantWithSite(site, variant);
            var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
            recipeVariantPage.AddIngredient(itemName);

            if (currency != null && decimalSeparatorValue != null)
            {
                initPriceRecipe = recipeVariantPage.GetRecipePrice(currency, decimalSeparatorValue);
            }

            return initPriceRecipe;
        }

        public double GetRecipePrice(HomePage homePage, string recipeName, string currency, string decimalSeparatorValue)
        {
            var recipesPage = homePage.GoToMenus_Recipes();
            recipesPage.Filter(PageObjects.Menus.Recipes.RecipesPage.FilterType.SearchRecipe, recipeName);
            var recipeGeneralInfosPage = recipesPage.SelectFirstRecipe();
            var recipeVariantPage = recipeGeneralInfosPage.ClickOnFirstVariant();
            double newPriceRecipe = recipeVariantPage.GetRecipePrice(currency, decimalSeparatorValue);

            return newPriceRecipe;
        }

        public double CreateMenu(HomePage homePage, string menuName, string recipeName, string site, string currency, string decimalSeparatorValue)
        {
            string menuVariant = TestContext.Properties["MenuVariant"].ToString();

            var menusPage = homePage.GoToMenus_Menus();
            var menusCreateModalPage = menusPage.MenuCreatePage();
            var menuGeneralInformation = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now.AddDays(+1), DateUtils.Now.AddMonths(+1), site, menuVariant);
            menuGeneralInformation.AddRecipe(recipeName);

            double initPriceMenu = menusPage.GetMenuPrice(currency, decimalSeparatorValue);

            return initPriceMenu;
        }

        public double GetMenuPrice(HomePage homePage, string menuName, string currency, string decimalSeparatorValue)
        {
            var menusPage = homePage.GoToMenus_Menus();
            menusPage.Filter(PageObjects.Menus.Menus.MenusPage.FilterType.SearchMenu, menuName);
            menusPage.UnselectServicesFromMenu();
            menusPage.SelectFirstMenu();

            double newPriceMenu = menusPage.GetMenuPrice(currency, decimalSeparatorValue);

            return newPriceMenu;
        }
    }


}
