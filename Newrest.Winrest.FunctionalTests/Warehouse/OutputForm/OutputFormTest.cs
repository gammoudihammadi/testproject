using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.DeliveryRound;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Production;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Sites;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Item;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.Claims;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.Inventory;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.OutputForm;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.ReceiptNotes;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.Extensions;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Policy;
using System.Threading;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using static Newrest.Winrest.FunctionalTests.PageObjects.Shared.GenericExport;

namespace Newrest.Winrest.FunctionalTests.Warehouse
{
    [TestClass]
    public class OutputFormTest : TestBase
    {
        private const int _timeout = 600000;
        public string downloadsPath = "C:\\ChromeDriverDownloads";
        private readonly FileNamePattern fileNamePattern = new FileNamePattern
        {
            Pattern = "outputforms",
            HasDate = true
        };
        private static Random random = new Random();
        private static string ID = "";
        private static string ID1 = "";
        private static string ID2 = "";
        private static string ID3 = "";
        private static string ID4 = "";
        private static string ID5 = "";
        private static string ID6 = "";
        private readonly string OUTPUT_FORM_EXCEL_SHEET = "Output forms";
        private static DateTime fromDate = DateUtils.Now;
        private static DateTime dateBeforeFrom = DateUtils.Now.AddDays(-7);
        private static DateTime dateBetweenFromTo = DateUtils.Now.AddDays(+5);


        [TestInitialize]
        public override void TestInitialize()
        {
            base.TestInitialize();

            var testMethod = TestContext.TestName;
            switch (testMethod)
            {
                case nameof(WA_OF_Index_Export):
                    WA_OF_Index_Initialize();
                    break;

                case nameof(WA_OF_Filter_SearchByNumber):
                    WA_OF_Index_Initialize();
                    break;

                case nameof(WA_OF_Filter_SortByDate):
                    WA_OF_Index_Initialize();
                    break;
                case nameof(WA_OF_Filter_SortByNumber):
                    WA_OF_Index_Initialize();
                    break;

                case nameof(WA_OF_Filter_ShowInactive):
                    WA_OF_Index_Initialize();
                    break;

                case nameof(WA_OF_Filter_ShowActive):
                    WA_OF_Index_Initialize();
                    break;

                case nameof(WA_OF_Filter_ShowAll):
                    WA_OF_Index_Initialize();
                    break;

                case nameof(WA_OF_Filter_Sites):
                    WA_OF_Index_Initialize();
                    break;

                case nameof(WA_OF_Filter_DateFrom):
                    WA_OF_FilterDateFromTo_Initialize();
                    break;
                case nameof(WA_OF_Filter_DateTo):
                    WA_OF_FilterDateFromTo_Initialize();
                    break;

                default:
                    break;
            }
        }

        [TestCleanup]
        public override void TestCleanup()
        {
            var testMethod = TestContext.TestName;
            switch (testMethod)
            {
                case nameof(WA_OF_Index_Export):
                    WA_OF_Index_CleanUp();
                    break;

                case nameof(WA_OF_Filter_SearchByNumber):
                    WA_OF_Index_CleanUp();
                    break;

                case nameof(WA_OF_Filter_SortByDate):
                    WA_OF_Index_CleanUp();
                    break;
                case nameof(WA_OF_Filter_SortByNumber):
                    WA_OF_Index_CleanUp();
                    break;

                case nameof(WA_OF_Filter_ShowInactive):
                    WA_OF_Index_CleanUp();
                    break;

                case nameof(WA_OF_Filter_ShowActive):
                    WA_OF_Index_CleanUp();
                    break;

                case nameof(WA_OF_Filter_ShowAll):
                    WA_OF_Index_CleanUp();
                    break;

                case nameof(WA_OF_Filter_DateFrom):
                    WA_OF_Index_CleanUp();
                    break;

                case nameof(WA_OF_Filter_DateTo):
                    WA_OF_Index_CleanUp();
                    break;
                default:
                    break;
            }
            base.TestCleanup();
        }

        private void WA_OF_Index_Initialize()
        {
            string site = TestContext.Properties["Site"].ToString();

            //Creation d'un OF
            TestContext.Properties[string.Format("OFId0")] = InsertOutputForm(DateTime.Now, site);
            ID = GetOutputFormNumber((int)TestContext.Properties[string.Format("OFId0")]);

            // Log in
            HomePage homePage = LogInAsAdmin();

            //Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();

            outputFormPage.Filter(OutputFormPage.FilterType.SearchByNumber, ID);
            OutputFormItem outputFormItem = outputFormPage.SelectFirstOutputForm();
            var item1 = outputFormItem.GetItemNames()[0];
            var item2 = outputFormItem.GetItemNames()[1];
            outputFormItem.AddPhyQuantity(item1, "5");
            outputFormItem.AddPhyQuantity(item2, "10");
            outputFormItem.BackToList();
        }

        private void WA_OF_FilterDateFromTo_Initialize()
        {
            string site = TestContext.Properties["Site"].ToString();

            TestContext.Properties[string.Format("OFId1")] = InsertOutputForm(dateBeforeFrom, site);
            TestContext.Properties[string.Format("OFId2")] = InsertOutputForm(fromDate, site);
            TestContext.Properties[string.Format("OFId3")] = InsertOutputForm(dateBetweenFromTo, site);

            ID1 = GetOutputFormNumber((int)TestContext.Properties[string.Format("OFId1")]);
            ID2 = GetOutputFormNumber((int)TestContext.Properties[string.Format("OFId2")]);
            ID3 = GetOutputFormNumber((int)TestContext.Properties[string.Format("OFId3")]);

            // Log in
            HomePage homePage = LogInAsAdmin();

            //Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();

            List<string> ints = new List<string> {ID1, ID2, ID3};
            outputFormPage.AddItemsToOutputForm(ints);
        }
        
        private void WA_OF_Index_CleanUp()
        {
            List<int> OFIds = new List<int>();

            for (int i = 0; i <= 4; i++)
            {
                if (TestContext.Properties.Contains($"OFId{i}"))
                {
                    OFIds.Add((int)TestContext.Properties[$"OFId{i}"]);
                }
            }
            DeleteOutputForms(OFIds);
        }

        public void DeleteOutputForms(List<int> OFIds)
        {
            foreach(var id in OFIds)
            {
                DeleteOutputForm(id);
            }
        }
        public void FilterForExport(OutputFormPage outputFormPage, string ID)
        {
            outputFormPage.ResetFilter();
            outputFormPage.WaitPageLoading();
            outputFormPage.Filter(OutputFormPage.FilterType.DateFrom, DateUtils.Now.AddDays(-30));
            outputFormPage.Filter(OutputFormPage.FilterType.DateTo, DateUtils.Now.AddDays(1));
            outputFormPage.Filter(OutputFormPage.FilterType.SearchByNumber, ID);
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_Index_Export()
        {
            HomePage homePage = new HomePage(WebDriver, TestContext);
            OutputFormPage outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            FilterForExport(outputFormPage, ID);
            string site = outputFormPage.GetSite();
            string fromPlace = outputFormPage.GetFromPlace();
            string toPlace = outputFormPage.GetToPlace();
            string totalVat = outputFormPage.GetTotalVat();
            outputFormPage.ClearDownloads();
            DeleteAllFileDownload();
            outputFormPage.ExportResults(true);
            GenericExport genericExport = new GenericExport(WebDriver, TestContext);
            genericExport.WaitPageLoading();
            FileInfo correctDownloadedFile = genericExport.IsGenerated(fileNamePattern);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");
            bool isExcelFileCorrect = genericExport.IsExcelFileCorrect(correctDownloadedFile.Name, fileNamePattern);
            Assert.IsTrue(isExcelFileCorrect, "Le nom du fichier exporté ne correspond pas au pattern général.");
            int resultNumber = genericExport.resultNumber(correctDownloadedFile.Name, OUTPUT_FORM_EXCEL_SHEET);
            Assert.AreEqual(2, resultNumber, "Les données du fichier exporté ne correspondent pas au filtre appliqué.");
            List<string> numberList = OpenXmlExcel.GetValuesInList("Number", OUTPUT_FORM_EXCEL_SHEET, correctDownloadedFile.FullName);
            bool isNumberCorrect = numberList.All(x => x.Equals(ID));
            Assert.IsTrue(isNumberCorrect, MessageErreur.EXCEL_DONNEES_KO);
            List<string> siteList = OpenXmlExcel.GetValuesInList("Site", OUTPUT_FORM_EXCEL_SHEET, correctDownloadedFile.FullName);
            bool isSiteCorrect = siteList.All(x => x.Equals(site));
            Assert.IsTrue(isSiteCorrect, MessageErreur.EXCEL_DONNEES_KO);
            List<string> fromPlaceList = OpenXmlExcel.GetValuesInList("From", OUTPUT_FORM_EXCEL_SHEET, correctDownloadedFile.FullName);
            bool isFromPlaceCorrect = fromPlaceList.All(x => x.Equals(fromPlace));
            Assert.IsTrue(isFromPlaceCorrect, MessageErreur.EXCEL_DONNEES_KO);
            List<string> toPlaceList = OpenXmlExcel.GetValuesInList("To", OUTPUT_FORM_EXCEL_SHEET, correctDownloadedFile.FullName);
            bool isToPlaceCorrect = toPlaceList.All(x => x.Equals(toPlace));
            Assert.IsTrue(isToPlaceCorrect, MessageErreur.EXCEL_DONNEES_KO);
            List<string> totalVatList = OpenXmlExcel.GetValuesInList("Total No VAT", OUTPUT_FORM_EXCEL_SHEET, correctDownloadedFile.FullName);
            decimal sumTotalVat = totalVatList.Sum(x => decimal.Parse(x));
            bool isTotalVatCorrect = sumTotalVat.Equals(decimal.Parse(totalVat));
            Assert.IsTrue(isTotalVatCorrect, MessageErreur.EXCEL_DONNEES_KO);
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_Filter_SearchByNumber()
        {
            // Log in
            var homePage = LogInAsAdmin();

            // Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();

            outputFormPage.Filter(OutputFormPage.FilterType.SearchByNumber, ID);
            string firstOutputFormNumber = outputFormPage.GetFirstOutputFormNumber();
            //Assert
            Assert.AreEqual(ID, firstOutputFormNumber, String.Format(MessageErreur.FILTRE_ERRONE, "Search by number"));
        }

        /// <summary>
        /// 
        /// Mise en place du paramétrage pour la configuration Winrest 4.0 
        /// 
        /// </summary>
        [TestMethod]
        [Priority(0)]
        [Timeout(_timeout)]
        public void WA_OF_SetConfigWinrest4_0()
        {
            string site = TestContext.Properties["Site"].ToString();
            string placeFrom = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();
            ClearCache();

            // New version search
            homePage.SetNewVersionSearchValue(true);

            // New group display
            homePage.SetNewGroupDisplayValue(true);

            // Vérifier que c'est activé
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();

            // Create outputForm
            var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
            outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo);
            var outputFormItem = outputFormCreateModalpage.Submit();

            outputFormItem.SelectFirstItem();

            try
            {
                var itemName = outputFormItem.GetFirstItemName();
                outputFormItem.Filter(OutputFormItem.FilterItemType.SearchByName, itemName);
            }
            catch
            {
                throw new Exception("La recherche n'a pas pu être effectuée, le NewSearchMode est inactif.");
            }

            // vérifier new group display
            Assert.IsTrue(outputFormItem.IsGroupDisplayActive(), "Le paramètre 'NewGroupDisplay' n'est pas activé.");

            //Suppression du Premier OF
            outputFormItem.BackToList();
            outputFormPage.Filter(OutputFormPage.FilterType.CheckAllSites, true);
            outputFormPage.Filter(OutputFormPage.FilterType.CheckAllOfTypes, true);
            outputFormPage.DeleteFirstOutputForm();
        }

        [Priority(1)]
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_CreateItemForOutputForm()
        {
            // Prepare items
            string itemName = TestContext.Properties["Item_OutputForm"].ToString();
            string placeFrom = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxName = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            // Création items
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);

            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxName, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);

                itemPage = itemGeneralInformationPage.BackToList();
                itemPage.ResetFilter();
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
            }

            //Assert
            var firstItemName = itemPage.GetFirstItemName();
            Assert.AreEqual(itemName, firstItemName, $"L'item {itemName} n'est pas présent dans la liste des items disponibles.");

            // On vérifie que la quantité de l'item est supérieure à 0
            var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            inventoriesPage.ResetFilters();
            var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
            inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, placeFrom, true);
            var inventoryDetail = inventoryCreateModalpage.Submit();
            inventoryDetail.Filter(InventoryItem.FilterItemType.SearchByName, itemName);
            string theoricalQty = inventoryDetail.GetTheoricalQuantity();

            if (theoricalQty.Equals("0"))
            {
                inventoryDetail.SelectFirstItem();
                inventoryDetail.AddPhysicalQuantity(itemName, "10");
                var validateInventory = inventoryDetail.Validate();
                validateInventory.ValidatePartialInventory();
                inventoriesPage = inventoryDetail.BackToList();
                inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, placeFrom, true);
                inventoryDetail = inventoryCreateModalpage.Submit();
                inventoryDetail.Filter(InventoryItem.FilterItemType.SearchByName, itemName);
                theoricalQty = inventoryDetail.GetTheoricalQuantity();
            }

            //Assert
            Assert.AreNotEqual("0", theoricalQty, $"L'item {itemName} n'a pas de quantité définie.");
        }

        [Priority(2)]
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_PrepareExportSageConfig()
        {
            //Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToOFSage"].ToString();
            string placeTo = TestContext.Properties["PlaceToOFSage"].ToString();
            string itemName = TestContext.Properties["Item_OutputForm"].ToString();
            string journalOutputForm = TestContext.Properties["Journal_OF"].ToString();

            string taxName = TestContext.Properties["Item_TaxType"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            // Récupération du groupe de l'item
            string itemGroup = GetItemGroup(homePage, itemName);

            // Vérification du paramétrage
            // --> Admin Settings
            bool isAppSettingsOK = SetApplicationSettingsForSageAuto(homePage);

            // Sites From -- > Analytical plan et section
            bool isSiteFromAnalyticalPlanOK = VerifySiteAnalyticalPlanSection(homePage, siteFrom);

            // Sites To -- > Analytical plan et section
            bool isSiteToAnalyticalPlanOK = VerifySiteAnalyticalPlanSection(homePage, siteTo);

            // Sites From -- > Config place
            bool isSiteToConfigPlaceOK = VerifySiteConfigPlace(homePage, siteFrom, placeTo, siteTo);

            // Parameter - Accounting --> Item group & VAT
            bool isGroupAndVatOK = VerifyGroupAndVAT(homePage, itemGroup, taxName);

            // Parameter - Accounting --> Journal
            bool isJournalOk = VerifyAccountingJournal(homePage, siteFrom, journalOutputForm);

            // IntegrationDate
            var date = VerifyIntegrationDate(homePage);

            // Assert
            Assert.AreNotEqual("", itemGroup, "Le groupe de l'item n'a pas été récupéré.");
            Assert.IsTrue(isAppSettingsOK, "Les application settings pour TL ne sont pas configurés correctement.");
            Assert.IsTrue(isSiteFromAnalyticalPlanOK, $"La configuration des analytical plan du site {siteFrom} n'est pas effectuée.");
            Assert.IsTrue(isSiteToAnalyticalPlanOK, $"La configuration des analytical plan du site {siteTo} n'est pas effectuée.");
            Assert.IsTrue(isSiteToConfigPlaceOK, $"La configuration de la place du site {siteTo} n'est pas effectuée.");
            Assert.IsTrue(isGroupAndVatOK, "La configuration du group and VAT de l'item n'est pas effectuée.");
            Assert.IsTrue(isJournalOk, "La catégorie du accounting journal n'a pas été effectuée.");
            Assert.IsNotNull(date, "La date d'intégration est nulle.");
        }

        /*
         * Création d'un nouvel "output form"
         */
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_CreateOutputForm()
        {
            string site = TestContext.Properties["Site"].ToString();
            string placeFrom = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();

            // Log in
            var homePage = LogInAsAdmin();

            //Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();

            try
            {
                // Create
                var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
                outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo);
                String ID = outputFormCreateModalpage.GetOutputFormNumber();
                var outputFormItem = outputFormCreateModalpage.Submit();

                outputFormPage = outputFormItem.BackToList();
                outputFormPage.ResetFilter();
                outputFormPage.Filter(OutputFormPage.FilterType.SearchByNumber, ID);
                outputFormPage.Filter(OutputFormPage.FilterType.CheckAllSites, true);
                outputFormPage.Filter(OutputFormPage.FilterType.CheckAllOfTypes, true);
                string firstOutputForm = outputFormPage.GetFirstOutputFormNumber();
                //Assert
                Assert.AreEqual(ID, firstOutputForm, String.Format(MessageErreur.OBJET_NON_CREE, "L'output form"));
            }
            finally
            {
                outputFormPage.DeleteFirstOutputForm();
            }

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_CreateOutputFormFromSupplierOrder()
        {
            string site = TestContext.Properties["Site"].ToString();
            string placeFrom = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();

            // Log in
            var homePage = LogInAsAdmin();
            //Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            try
            {
                outputFormPage.ResetFilter();
                // Create
                // 1. Cliquer sur "New output form"
                var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
                // 2.Remplir le formulaire
                // 3. Cocher Create From
                outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo, true, true);
                // 4.Sélectionner à partir d'un SO
                string receiptNoteNumber = outputFormCreateModalpage.SelectSupplierOrder();

                // You must select at least one output form, receipt note or supply order.
                String ID = outputFormCreateModalpage.GetOutputFormNumber();
                // 5. Cliquer sur "create"
                var outputFormItem = outputFormCreateModalpage.Submit();

                var itemName = outputFormItem.GetFirstItemName();
                var quantity = outputFormItem.GetTheoricalQuantity();

                // Select First item and Add Physical Quantity
                outputFormItem.SelectFirstItem();
                outputFormItem.AddPhysicalQuantity(itemName, quantity);

                OutputFormGeneralInformation generalInfo = outputFormItem.ClickOnGeneralInformationTab();

                // check Supply Order Message and compare Number
                var checkSupplyOrder = generalInfo.GetSupplyOrderMessage();
                var checkSupplyOrderNumber = generalInfo.GetSupplyOrderNumber();
                // Assert
                Assert.AreEqual("From supply order(s) : " + checkSupplyOrderNumber, checkSupplyOrder, "supply order pas trouvé");
                Assert.AreEqual(receiptNoteNumber, checkSupplyOrderNumber, "Mauvais receipt note number dans GeneralInfo");
                outputFormPage = outputFormItem.BackToList();
                outputFormPage.ResetFilter();
                // Filter By Number Output Form Created and check Total Number 
                outputFormPage.Filter(OutputFormPage.FilterType.SearchByNumber, ID);
                var firstOFNumber = outputFormPage.GetFirstOutputFormNumber();
                var totalNumber = outputFormPage.CheckTotalNumber();
                //Assert
                Assert.AreEqual(ID, firstOFNumber, String.Format(MessageErreur.OBJET_NON_CREE, "L'output form"));
                Assert.AreEqual(1, totalNumber, String.Format(MessageErreur.OBJET_NON_CREE, "L'output form"));
            }
            finally
            {
                outputFormPage.DeleteFirstOutputForm();
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_CreateOutputFormFromReceiptNote()
        {
            string site = TestContext.Properties["Site"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();
            string itemName = TestContext.Properties["Item_OutputForm"].ToString();
            HomePage homePage = LogInAsAdmin();
            ReceiptNotesPage receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();
            ReceiptNotesCreateModalPage receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
            receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, deliveryPlace));
            ReceiptNotesItem receiptNotesItem = receiptNotesCreateModalpage.Submit();
            receiptNotesItem.ResetFilters();
            receiptNotesItem.Filter(ReceiptNotesItem.FilterItemType.SearchByName, itemName);
            receiptNotesItem.SelectFirstItem();
            receiptNotesItem.AddReceived(itemName, "2");
            receiptNotesItem.Refresh();
            ReceiptNotesQualityChecks qualityChecks = receiptNotesItem.ClickOnChecksTab();
            qualityChecks.DeliveryAccepted();
            if (qualityChecks.CanClickOnSecurityChecks())
            {
                qualityChecks.CanClickOnSecurityChecks();
                qualityChecks.SetSecurityChecks("No");
                qualityChecks.SetQualityChecks();
                receiptNotesItem = qualityChecks.ClickOnReceiptNoteItemTab();
            }
            else receiptNotesItem = qualityChecks.ClickOnReceiptNoteItemTab();
            receiptNotesItem.Validate();
            string rnNumber = receiptNotesItem.GetReceiptNoteNumber();
            OutputFormPage outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();
            OutputFormCreateModalPage outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
            outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, deliveryPlace, placeTo, true, true);
            outputFormCreateModalpage.SelectReceiptNote(rnNumber);
            string id = outputFormCreateModalpage.GetOutputFormNumber();
            OutputFormItem outputFormItem = outputFormCreateModalpage.Submit();
            outputFormItem.Filter(OutputFormItem.FilterItemType.ShowItemsWithPhysQty, false);
            string quantity = outputFormItem.GetTheoricalQuantity();
            outputFormItem.SelectFirstItem();
            outputFormItem.AddPhyQuantity(itemName, quantity);
            outputFormItem.Refresh();
            outputFormPage = outputFormItem.BackToList();
            outputFormPage.ResetFilter();
            outputFormPage.Filter(OutputFormPage.FilterType.CheckAllSites, true);
            outputFormPage.Filter(OutputFormPage.FilterType.CheckAllOfTypes, true);
            string firstOutputFormNumber = outputFormPage.GetFirstOutputFormNumber();
            Assert.AreEqual(id, firstOutputFormNumber, String.Format(MessageErreur.OBJET_NON_CREE, "L'output form"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_CreateOutputFormFromPurchaseOrder()
        {
            string site = TestContext.Properties["Site"].ToString();
            string placeFrom = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();
            string ID = String.Empty;
            // Log in
            var homePage = LogInAsAdmin();

            //Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();

            try
            {
                // Create
                // 1. Cliquer sur "New output form"
                var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
                // 2.Remplir le formulaire
                // 3. Cocher Create From
                outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo, true, true);
                // 4. Sélectionner à partir d'un PO
                outputFormCreateModalpage.SelectPurchaseOrder();

                // You must select at least one output form, receipt note or supply order.
                ID = outputFormCreateModalpage.GetOutputFormNumber();
                // 5. Cliquer sur "create"
                var outputFormItem = outputFormCreateModalpage.Submit();

                var itemName = outputFormItem.GetFirstItemName();
                var quantity = outputFormItem.GetTheoricalQuantity();
                // Select Item and Add Physical Quantity
                outputFormItem.SelectFirstItem();
                outputFormItem.AddPhysicalQuantity(itemName, quantity);
                outputFormPage = outputFormItem.BackToList();
                // Filter Output Form Created 
                outputFormPage.ResetFilter();
                outputFormPage.Filter(OutputFormPage.FilterType.SearchByNumber, ID);
                outputFormPage.Filter(OutputFormPage.FilterType.CheckAllSites, true);
                outputFormPage.Filter(OutputFormPage.FilterType.CheckAllOfTypes, true);
                string firstOutputFormNumber = outputFormPage.GetFirstOutputFormNumber();
                var totalNumber = outputFormPage.CheckTotalNumber();
                //Assert
                Assert.AreEqual(ID, firstOutputFormNumber, string.Format(MessageErreur.OBJET_NON_CREE, "L'output form"));
                Assert.AreEqual(1, totalNumber, "L'outout Form n'est pas Ajouter ");
            }
            finally
            {
                outputFormPage.DeleteFirstOutputForm();
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_CreateOutputFormFromOuputForm()
        {
            string site = TestContext.Properties["Site"].ToString();
            string placeFrom = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();

            // Log in
            var homePage = LogInAsAdmin();

            //Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();

            // Create
            // 1. Cliquer sur "New output form"
            var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
            // 2.Remplir le formulaire
            // 3. Cocher Create From
            outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo, true, true);
            // 4. Sélectionner à partir d'un OF
            outputFormCreateModalpage.SelectOutputForm();

            // You must select at least one output form, receipt note or supply order.
            String ID = outputFormCreateModalpage.GetOutputFormNumber();
            // 5. Cliquer sur "create"
            var outputFormItem = outputFormCreateModalpage.Submit();
            
            outputFormPage = outputFormItem.BackToList();
            outputFormPage.ResetFilter();
            outputFormPage.Filter(OutputFormPage.FilterType.SearchByNumber, ID);
            outputFormPage.Filter(OutputFormPage.FilterType.CheckAllSites, true);
            outputFormPage.Filter(OutputFormPage.FilterType.CheckAllOfTypes, true);
            string firstOutputForm = outputFormPage.GetFirstOutputFormNumber();
            //Assert
            Assert.AreEqual(ID, firstOutputForm, String.Format(MessageErreur.OBJET_NON_CREE, "L'output form"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_Filter_SortByDate()
        {
            // Log in
            var homePage = LogInAsAdmin();

            //Act
            string dateFormat = homePage.GetDateFormatPickerValue();
            OutputFormPage outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();

            int totalNumber = outputFormPage.CheckTotalNumber();
            Assert.AreNotEqual(totalNumber, 0, "Pas de OUTPUT FORM");

            if (!outputFormPage.isPageSizeEqualsTo100())
            {
                outputFormPage.PageSize("8");
                outputFormPage.PageSize("100");
            }
            var isSorted = outputFormPage.SortByFilters("DATE", dateFormat);

            //Assert
            Assert.IsTrue(isSorted, String.Format(MessageErreur.FILTRE_ERRONE, "'Sort by'"));

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_Filter_SortByNumber()
        {
            // Log in
            var homePage = LogInAsAdmin();

            //Act
            string dateFormat = homePage.GetDateFormatPickerValue();
            OutputFormPage outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();

            int totalNumber = outputFormPage.CheckTotalNumber();
            Assert.AreNotEqual(totalNumber, 0, "Pas de OUTPUT FORM");

            if (!outputFormPage.isPageSizeEqualsTo100())
            {
                outputFormPage.PageSize("8");
                outputFormPage.PageSize("100");
            }
            var isSorted = outputFormPage.SortByFilters("NUMBER", dateFormat);

            //Assert
            Assert.IsTrue(isSorted, String.Format(MessageErreur.FILTRE_ERRONE, "'Sort by'"));

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_Filter_ShowValidated()
        {
            // Log in
            var homePage = LogInAsAdmin();

            // Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();

            outputFormPage.ResetFilter();
            outputFormPage.Filter(OutputFormPage.FilterType.ShowNotValidated, false);

            if (!outputFormPage.isPageSizeEqualsTo100())
            {
                outputFormPage.PageSize("8");
                outputFormPage.PageSize("100");
            }

            bool checkShowItemsNotValidated = outputFormPage.CheckShowItemsNotValidated(false);
            //Assert
            Assert.IsFalse(checkShowItemsNotValidated, String.Format(MessageErreur.FILTRE_ERRONE, "'Show items validated'"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_Filter_ShowNotValidated()
        {
            string site = TestContext.Properties["Site"].ToString();
            string placeFrom = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();

            // Log in
            var homePage = LogInAsAdmin();

            // Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();
            outputFormPage.Filter(OutputFormPage.FilterType.ShowNotValidated, true);

            if (outputFormPage.CheckTotalNumber() < 20)
            {
                //Create a new Inventory NOT validated
                var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
                outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo, true);
                var outputFormItem = outputFormCreateModalpage.Submit();

                var itemName = outputFormItem.GetFirstItemName();
                var quantity = outputFormItem.GetTheoricalQuantity();

                outputFormItem.SelectFirstItem();
                outputFormItem.AddPhysicalQuantity(itemName, quantity);

                //Back to list page
                outputFormPage = outputFormItem.BackToList();
                outputFormPage.ResetFilter();
                outputFormPage.Filter(OutputFormPage.FilterType.ShowNotValidated, true);
                outputFormPage.Filter(OutputFormPage.FilterType.CheckAllSites, true);
                outputFormPage.Filter(OutputFormPage.FilterType.CheckAllOfTypes, true);
            }

            if (!outputFormPage.isPageSizeEqualsTo100())
            {
                outputFormPage.PageSize("8");
                outputFormPage.PageSize("100");
            }

            //Assert
            Assert.IsTrue(outputFormPage.CheckShowItemsNotValidated(true), String.Format(MessageErreur.FILTRE_ERRONE, "'Show items not validated'"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_Filter_ShowInactive()
        {
            // Log in
            var homePage = LogInAsAdmin();

            // Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();
            outputFormPage.Filter(OutputFormPage.FilterType.ShowGroup, true, OutputFormPage.TypeOfShowGroup.All);
            var totalNbreAll = outputFormPage.CheckTotalNumber();
            outputFormPage.Filter(OutputFormPage.FilterType.ShowGroup, true, OutputFormPage.TypeOfShowGroup.ActiveOnly);
            var totalNbreActiveOnly = outputFormPage.CheckTotalNumber();

            if (!outputFormPage.isPageSizeEqualsTo100())
            {
                outputFormPage.PageSize("8");
                outputFormPage.PageSize("100");
            }

            outputFormPage.Filter(OutputFormPage.FilterType.ShowGroup, true, OutputFormPage.TypeOfShowGroup.InactiveOnly);
            var totalNbreInactiveOnly = outputFormPage.CheckTotalNumber();
            //Assert
            Assert.IsFalse(outputFormPage.CheckStatus(false), String.Format(MessageErreur.FILTRE_ERRONE, "'Show only inactive'"));
            Assert.AreEqual(totalNbreAll - totalNbreActiveOnly, totalNbreInactiveOnly, String.Format(MessageErreur.FILTRE_ERRONE, "'Show only inactive'"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_Filter_ShowActive()
        {
            // Log in
            var homePage = LogInAsAdmin();

            // Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();
            outputFormPage.Filter(OutputFormPage.FilterType.ShowGroup, true, OutputFormPage.TypeOfShowGroup.All);
            var totalNbreAll = outputFormPage.CheckTotalNumber();
            outputFormPage.Filter(OutputFormPage.FilterType.ShowGroup, true, OutputFormPage.TypeOfShowGroup.InactiveOnly);
            var totalNbreInactiveOnly = outputFormPage.CheckTotalNumber();

            //Assert
            outputFormPage.Filter(OutputFormPage.FilterType.ShowGroup, true, OutputFormPage.TypeOfShowGroup.ActiveOnly);
            var totalNbreActiveOnly = outputFormPage.CheckTotalNumber();
            Assert.IsTrue(outputFormPage.CheckStatus(true), String.Format(MessageErreur.FILTRE_ERRONE, "'Show only active'"));
            Assert.AreEqual(totalNbreAll - totalNbreInactiveOnly, totalNbreActiveOnly, String.Format(MessageErreur.FILTRE_ERRONE, "'Show only active'"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_Filter_ShowAll()
        {
            // Log in
            var homePage = LogInAsAdmin();

            // Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();
            outputFormPage.Filter(OutputFormPage.FilterType.ShowGroup, true, OutputFormPage.TypeOfShowGroup.All);
            var totalNbreAll = outputFormPage.CheckTotalNumber();
            outputFormPage.Filter(OutputFormPage.FilterType.ShowGroup, true, OutputFormPage.TypeOfShowGroup.InactiveOnly);
            var totalNbreInactiveOnly = outputFormPage.CheckTotalNumber();
            outputFormPage.Filter(OutputFormPage.FilterType.ShowGroup, true, OutputFormPage.TypeOfShowGroup.ActiveOnly);
            var totalNbreActiveOnly = outputFormPage.CheckTotalNumber();

            // Assert
            Assert.AreEqual(totalNbreAll, (totalNbreInactiveOnly + totalNbreActiveOnly), String.Format(MessageErreur.FILTRE_ERRONE, "'Show all'"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_Filter_DateFrom()
        {
            // Log in
            var homePage = LogInAsAdmin();

            string dateFormat = homePage.GetDateFormatPickerValue();

            //Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();
            DateTime fromDate = DateTime.Now.Date;

            //Filter
            outputFormPage.Filter(OutputFormPage.FilterType.DateFrom, fromDate);

            outputFormPage.Filter(OutputFormPage.FilterType.SearchByNumber, ID1);
            Assert.AreEqual(outputFormPage.CheckTotalNumber(), 0, "Le filtre Date From ne fonctionne pas correctement .");

            outputFormPage.Filter(OutputFormPage.FilterType.SearchByNumber, ID2);
            Assert.AreEqual(outputFormPage.CheckTotalNumber(), 1, "Le filtre Date From ne fonctionne pas correctement .");

            outputFormPage.Filter(OutputFormPage.FilterType.SearchByNumber, ID3);
            Assert.AreEqual(outputFormPage.CheckTotalNumber(), 1, "Le filtre Date From ne fonctionne pas correctement .");
            outputFormPage.ResetFilter();
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_Filter_DateTo()
        {

            // Log in
            var homePage = LogInAsAdmin();

            string dateFormat = homePage.GetDateFormatPickerValue();

            //Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();
            DateTime DateTo = DateTime.Now.Date;

            //Filter
            outputFormPage.Filter(OutputFormPage.FilterType.DateTo, DateTo);

            outputFormPage.Filter(OutputFormPage.FilterType.SearchByNumber, ID1);
            Assert.AreEqual(outputFormPage.CheckTotalNumber(), 1, "Le filtre Date To ne fonctionne pas correctement .");

            outputFormPage.Filter(OutputFormPage.FilterType.SearchByNumber, ID2);
            Assert.AreEqual(outputFormPage.CheckTotalNumber(), 1, "Le filtre Date To ne fonctionne pas correctement .");

            outputFormPage.Filter(OutputFormPage.FilterType.SearchByNumber, ID3);
            Assert.AreEqual(outputFormPage.CheckTotalNumber(), 0, "Le filtre Date To ne fonctionne pas correctement .");
            outputFormPage.ResetFilter();
        }
        
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_Filter_Sites()
        {
            string site = TestContext.Properties["Site"].ToString();
            // Log in
            var homePage = LogInAsAdmin();

            //Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();
            outputFormPage.Filter(OutputFormPage.FilterType.CombosOptions, site, OutputFormPage.TypeOfShowGroup.NoShowGroupOption, OutputFormPage.TypeCombosOptions.Sites);



            if (!outputFormPage.isPageSizeEqualsTo100())
            {
                outputFormPage.PageSize("8");
                outputFormPage.PageSize("100");
            }
            bool verifFilterSite = outputFormPage.VerifySite(site);

            //Assert
            Assert.IsTrue(verifFilterSite, String.Format(MessageErreur.FILTRE_ERRONE, "'Sites'"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_Items_Filter_SearchByName()
        {
            string site = TestContext.Properties["SiteBis"].ToString();
            string placeFrom = TestContext.Properties["PlaceFrom"].ToString();
            string placeTo = "Produccion";

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();

            // Create a outputForm 
            var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
            outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo, true);
            var outputFormItem = outputFormCreateModalpage.Submit();

            var itemName = outputFormItem.GetFirstItemName();
            outputFormItem.Filter(OutputFormItem.FilterItemType.SearchByName, itemName);

            Assert.IsTrue(outputFormItem.VerifyName(itemName), String.Format(MessageErreur.FILTRE_ERRONE, "'Search item by name'"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_Items_Filter_SortBy()
        {
            string site = TestContext.Properties["SiteBis"].ToString();
            string placeFrom = TestContext.Properties["PlaceFrom"].ToString();
            string placeTo = "Produccion";

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();

            // Create a outputForm 
            var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
            outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo, true);
            var outputFormItem = outputFormCreateModalpage.Submit();

            // Filter Sort By Name
            if (!outputFormPage.isPageSizeEqualsTo100WidhoutTotalNumber())
            {
                outputFormPage.PageSize("8");
                outputFormPage.PageSize("100");
            }

            outputFormItem.Filter(OutputFormItem.FilterItemType.SortBy, "Name");
            Assert.IsTrue(outputFormItem.IsSortedByName(), String.Format(MessageErreur.FILTRE_ERRONE, "'Sort by name'"));
            // Filter Sort By Item Group
            outputFormItem.Filter(OutputFormItem.FilterItemType.SortBy, "Group");
            Assert.IsTrue(outputFormItem.IsSortedByItemGroup(), String.Format(MessageErreur.FILTRE_ERRONE, "'Sort by item group'"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_Items_Filter_ShowItemsWithQty()
        {
            //prepare
            string site = TestContext.Properties["Site"].ToString();
            string placeFrom = TestContext.Properties["PlaceFrom"].ToString();
            string placeTo = "Producción";
            string outputFormNumber = string.Empty;
            // arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();
            try
            {
                // Create a outputForm 
                var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
                outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo, true);
                outputFormNumber = outputFormCreateModalpage.GetOutputFormNumber();
                var outputFormItem = outputFormCreateModalpage.Submit();

                // Filters Process
                var itemName = outputFormItem.GetFirstItemName();

                // Set Physical Quantity to 1
                outputFormItem.SelectFirstItem();
                outputFormItem.AddPhysicalQuantity(itemName, "1");
                outputFormItem.Filter(OutputFormItem.FilterItemType.SearchByName, itemName);
                outputFormItem.Filter(OutputFormItem.FilterItemType.ShowItemsWithPhysQty, true);

                //Assert
                var QtyPhyIsEdited = outputFormItem.IsWithEditedPhysQty();
                Assert.IsTrue(QtyPhyIsEdited, String.Format(MessageErreur.FILTRE_ERRONE, "'Show items with edited physical qty'"));

                outputFormItem.ResetFilter();
                outputFormItem.Filter(OutputFormItem.FilterItemType.ShowItemsWithPhysQty, false);
                outputFormItem.Filter(OutputFormItem.FilterItemType.ShowItemsWithTheoQtyOnly, true);
                outputFormItem.PageSize("100");
                var isWithPositiveTheoQty = outputFormItem.IsWithPositiveTheoQty();
                Assert.IsTrue(isWithPositiveTheoQty, String.Format(MessageErreur.FILTRE_ERRONE, "'Show items with positive theorical qty'"));
            }
            finally
            {
                homePage.GoToWarehouse_OutputFormPage();
                if (!string.IsNullOrEmpty(outputFormNumber))
                {
                    outputFormPage.Filter(OutputFormPage.FilterType.SearchByNumber, outputFormNumber);
                    outputFormPage.DeleteFirstOutputForm();

                }

            }

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_Items_Filter_Group()
        {
            string site = TestContext.Properties["SiteBis"].ToString();
            string placeFrom = TestContext.Properties["PlaceFrom"].ToString();
            string placeTo = "Produccion";

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();

            // Create a outputForm 
            var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
            outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo, true);
            var outputFormItem = outputFormCreateModalpage.Submit();

            var itemGroup = outputFormItem.GetFirstItemGroup();
            outputFormItem.SelectFirstItem();

            // Filter by group
            outputFormItem.Filter(OutputFormItem.FilterItemType.ByGroup, itemGroup);

            if (!outputFormItem.isPageSizeEqualsTo100WidhoutTotalNumber())
            {
                outputFormItem.PageSize("8");
                outputFormItem.PageSize("100");
            }

            Assert.IsTrue(outputFormItem.VerifyGroup(itemGroup), String.Format(MessageErreur.FILTRE_ERRONE, "'Group'"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_Filter_OFTypes()
        {
            DateTime startdateInput = DateUtils.Now.AddDays(-61);
            DateTime enddateInput = DateUtils.Now.AddDays(-30);
            bool newVersion = true;
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            // Log in

            var homePage = LogInAsAdmin();
            // Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            // check all
            outputFormPage.ResetFilter();
            outputFormPage.Filter(OutputFormPage.FilterType.CheckAllSites, true);
            outputFormPage.Filter(OutputFormPage.FilterType.CheckAllOfTypes, true);
            int totalCheckAll = outputFormPage.CheckTotalNumber();
            // uncheck all
            outputFormPage.Filter(OutputFormPage.FilterType.CombosOptions, null, OutputFormPage.TypeOfShowGroup.NoShowGroupOption, OutputFormPage.TypeCombosOptions.OfTypes);
            Assert.AreEqual(0, outputFormPage.CheckTotalNumber(), "UncheckAll raté");
            // check Standard
            outputFormPage.Filter(OutputFormPage.FilterType.CombosOptions, "Standard", OutputFormPage.TypeOfShowGroup.NoShowGroupOption, OutputFormPage.TypeCombosOptions.OfTypes);
            Assert.AreEqual(totalCheckAll, outputFormPage.CheckTotalNumber(), "Pas tous Standard");
            DeleteAllFileDownload();
            outputFormPage.ClearDownloads();
            outputFormPage.Filter(OutputFormPage.FilterType.DateFrom, startdateInput);
            outputFormPage.Filter(OutputFormPage.FilterType.DateTo, enddateInput);
            outputFormPage.ExportResults(newVersion);
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();
            var correctDownloadedFile = outputFormPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);
        }

        /*
         * Mettre à jour la quantité physique
        */
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_UpdatePhysicalQuantity()
        {
            string site = TestContext.Properties["Site"].ToString();
            string placeFrom = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            CreateInventory(homePage, site, placeFrom);

            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();

            // Create an inactive OutputForm before the main test
            var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
            outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo);
            var outputFormItem = outputFormCreateModalpage.Submit();

            // Show only items with theorical qty
            outputFormItem.Filter(OutputFormItem.FilterItemType.ShowItemsWithTheoQtyOnly, true);

            var itemName = outputFormItem.GetFirstItemName();
            var quantity = outputFormItem.GetTheoricalQuantity();

            outputFormItem.SelectFirstItem();
            outputFormItem.AddPhysicalQuantity(itemName, quantity);

            // Change the Physical qty
            Assert.IsTrue(outputFormItem.IsUpdated(), "L'icône d'édition d'un item n'est pas disponible.");
        }

        /*
         * Mettre à jour la quantité physique en dépassant la quantité théorique
        */
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_UpdatePhysicalQuantityKO()
        {
            string site = TestContext.Properties["Site"].ToString();
            string placeFrom = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            CreateInventory(homePage, site, placeFrom);

            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();

            // Create an inactive OutputForm before the main test
            var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
            outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo);
            var outputFormItem = outputFormCreateModalpage.Submit();

            // Show only items with theorical qty
            outputFormItem.Filter(OutputFormItem.FilterItemType.ShowItemsWithTheoQtyOnly, true);
            int physicalQte = Int32.Parse(outputFormItem.GetTheoricalQuantity()) + 10;

            var itemName = outputFormItem.GetFirstItemName();
            outputFormItem.SelectFirstItem();
            outputFormItem.AddPhysicalQuantityOverload(itemName, physicalQte.ToString());

            Assert.IsTrue(outputFormItem.IsFailed(), "L'update de l'item a réussi alors qu'elle aurait du échouer due à un overload.");
        }

        /*
         * Valider un outputForm
         */
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_ValidateOutputForm()
        {
            string site = TestContext.Properties["Site"].ToString();
            string placeFrom = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();

            var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
            outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo, true);
            String ID = outputFormCreateModalpage.GetOutputFormNumber();
            var outputFormItem = outputFormCreateModalpage.Submit();

            var itemName = outputFormItem.GetFirstItemName();
            var quantity = outputFormItem.GetTheoricalQuantity();

            outputFormItem.SelectFirstItem();
            outputFormItem.AddPhysicalQuantity(itemName, quantity);

            outputFormItem.Validate();

            outputFormPage = outputFormItem.BackToList();
            outputFormPage.ResetFilter();
            outputFormPage.Filter(OutputFormPage.FilterType.SearchByNumber, ID);
            outputFormPage.Filter(OutputFormPage.FilterType.ShowNotValidated, false);
            outputFormPage.Filter(OutputFormPage.FilterType.CheckAllSites, true);
            outputFormPage.Filter(OutputFormPage.FilterType.CheckAllOfTypes, true);

            //Assert
            Assert.IsTrue(outputFormPage.CheckValidation(true), String.Format(MessageErreur.FILTRE_ERRONE, "'Show items validated'"));
        }

        /*
         * Rafraichir les donnees
         */
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_RefreshData()
        {
            string site = TestContext.Properties["Site"].ToString();
            string placeFrom = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();

            // Create an inactive OutputForm before the main test
            var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
            outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo);
            var outputFormItem = outputFormCreateModalpage.Submit();

            // Show only items with theorical qty
            outputFormItem.Filter(OutputFormItem.FilterItemType.ShowItemsWithTheoQtyOnly, true);
            var oldPhysicalQty = outputFormItem.GetPhysicalQuantity();

            Double theoricalQte = Double.Parse(outputFormItem.GetTheoricalQuantity().Replace(" ", ""));
            var itemName = outputFormItem.GetFirstItemName();

            outputFormItem.SelectFirstItem();
            outputFormItem.AddPhysicalQuantity(itemName, (theoricalQte).ToString());
            outputFormItem.Refresh();

            var newPhysicalQty = outputFormItem.GetPhysicalQuantity();

            // Assert
            Assert.AreNotEqual(oldPhysicalQty, newPhysicalQty, "Le refresh de l'output form a échoué.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_RefreshDataKO()
        {
            string site = TestContext.Properties["Site"].ToString();
            string placeFrom = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();

            // Create an inactive OutputForm before the main test
            var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
            outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo);
            var outputFormItem = outputFormCreateModalpage.Submit();

            // Show only items with theorical qty
            outputFormItem.Filter(OutputFormItem.FilterItemType.ShowItemsWithTheoQtyOnly, true);
            Double theoricalQte = Double.Parse(outputFormItem.GetTheoricalQuantity().Replace(" ", "")) + 10;
            var itemName = outputFormItem.GetFirstItemName();

            outputFormItem.SelectFirstItem();
            // au max
            Double maxPhys = Double.Parse(outputFormItem.GetTheoricalQuantity().Replace(" ", ""));
            outputFormItem.AddPhysicalQuantityOverload(itemName, maxPhys.ToString());
            // au max + 10 (donc validation rouge => pas sauvegardé)
            outputFormItem.AddPhysicalQuantityOverload(itemName, theoricalQte.ToString());

            // Refresh the page
            outputFormItem.Refresh();

            // Assert
            Assert.AreEqual(theoricalQte.ToString(), outputFormItem.GetPhysicalQuantity());
        }

        /*
         * Test d'impression des items d'un outputForm avec newVersionPrint = true
         */
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_PrintOutputFormNewVersion()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string placeFrom = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            //Output Form Report_-_437560_-_20220316141108.pdf
            string docFileNamePdfBegin = "Output Form Report_-_";
            //All_files_20220225_102148.zip
            string DocFileNameZipBegin = "All_files_";
            bool newVersionPrint = true;

            // Log in
            var homePage = LogInAsAdmin();

            //Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();
            outputFormPage.ClearDownloads();
            DeleteAllFileDownload();
            // Create New Output Form
            var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
            outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo);
            var outputFormItem = outputFormCreateModalpage.Submit();

            var itemName = outputFormItem.GetFirstItemName();
            var quantity = outputFormItem.GetTheoricalQuantity();

            // Add Physical Quantity
            outputFormItem.SelectFirstItem();
            outputFormItem.AddPhysicalQuantity(itemName, quantity);

            // Print Item Generique
            PrintReportPage reportPage = PrintItemGenerique(outputFormItem, newVersionPrint);
            OutputFormGeneralInformation generalInfo = outputFormItem.ClickOnGeneralInformationTab();
            string inventoryNumber = generalInfo.GetInventoryNumber();
            reportPage.Purge(downloadsPath, docFileNamePdfBegin, DocFileNameZipBegin);

            // cliquer sur PrintAll Zip
            string trouve = reportPage.PrintAllZip(downloadsPath, docFileNamePdfBegin, DocFileNameZipBegin);
            FileInfo fi = new FileInfo(trouve);
            fi.Refresh();
            Assert.IsTrue(fi.Exists, trouve + " non généré");

            PdfDocument document = PdfDocument.Open(fi.FullName);
            List<string> mots = new List<string>();
            foreach (Page page in document.GetPages())
            {
                IEnumerable<Word> words = page.GetWords();
                foreach (Word word in words)
                {
                    mots.Add(word.Text);
                }
            }
            // Assert
            Assert.AreNotEqual(0, mots.Count(w => w == "(" + site + ")"), "(" + site + ") non présent dans le Pdf");
            Assert.AreNotEqual(0, mots.Count(w => w == DateUtils.Now.Date.ToString("dd/MM/yyyy")), DateUtils.Now.Date.ToString("dd/MM/yyyy") + " non présent dans le Pdf");
            Assert.AreNotEqual(0, mots.Count(w => w.Contains(inventoryNumber)), inventoryNumber + " non présent dans le Pdf");
        }

        private PrintReportPage PrintItemGenerique(OutputFormItem outputFormItem, bool printVersion)
        {
            var reportPage = outputFormItem.Print(printVersion);
            var isReportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();

            //Assert  
            Assert.IsTrue(isReportGenerated, "L'application n'a pas pu générer le fichier attendu.");
            return reportPage;
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_UpdateGeneralInformations()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string placeFrom = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();

            var newComments = "A new comment";
            var newDate = DateUtils.Now;

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();

            var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
            outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo);
            var outputFormItem = outputFormCreateModalpage.Submit();

            var itemName = outputFormItem.GetFirstItemName();
            var quantity = outputFormItem.GetTheoricalQuantity();

            outputFormItem.SelectFirstItem();
            outputFormItem.AddPhysicalQuantity(itemName, quantity);

            var outputFormGeneralInformations = outputFormItem.ClickOnGeneralInformationTab();
            outputFormGeneralInformations.UpdateInformations(newComments, newDate);
            outputFormItem = outputFormGeneralInformations.ClickOnItemsTab();
            outputFormGeneralInformations = outputFormItem.ClickOnGeneralInformationTab();

            //Assert  
            Assert.AreEqual(outputFormGeneralInformations.GetComments(), newComments, "Le commentaire mis à jour dans la page General Information n'a pas été pris en compte.");
            Assert.AreEqual(outputFormGeneralInformations.GetDate(), newDate.ToString("dd/MM/yyyy"), "La date mise à jour dans la page General Information n'a pas été prise en compte.");
        }

        /*
         * Test d'impression d'OutputForm pour un résultat supérieur à 100
         */
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_PrintOutputFormOverload()
        {
            // Log in
            var homePage = LogInAsAdmin();

            //Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();

            outputFormPage.ClearDatesFilter();
            outputFormPage.WaitLoading();
            outputFormPage.Filter(OutputFormPage.FilterType.ShowGroup, true, OutputFormPage.TypeOfShowGroup.All);

            // Assert
            bool isPrint = true;
            var isPrintSuccess = outputFormPage.VerifyIsPrint();
            if (isPrintSuccess)
            {
                outputFormPage.PrintResults(true);
            }
            else
            {
                isPrint = false;
            }

            Assert.IsFalse(isPrint, "La fonctionnalité de Print du fichier n'est pas accessible.");
        }

        /*
         * Test d'impression des outputForm avec newVersionPrint = true
         */
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_PrintOutputFormListNewVersion()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string placeFrom = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();
            bool newVersionPrint = true;
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            //Output Form Report_-_437551_-_20220316133021.pdf
            string docFileNamePdfBegin = "Output Form Report_-_";
            //All_files_20220225_102148.zip
            string docFileNameZipBegin = "All_files_";
            string ID = String.Empty;
            // Log in
            var homePage = LogInAsAdmin();

            //Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();
            outputFormPage.ClearDownloads();
            DeleteAllFileDownload();
            outputFormPage.Filter(OutputFormPage.FilterType.DateFrom, DateTime.Now);
            outputFormPage.Filter(OutputFormPage.FilterType.DateTo, DateTime.Now.AddMonths(1));     
            try
            {
                // Create New Output Form
                var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
                outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo);
                ID = outputFormCreateModalpage.GetOutputFormNumber();
                var outputFormItem = outputFormCreateModalpage.Submit();

                var itemName = outputFormItem.GetFirstItemName();
                var quantity = outputFormItem.GetTheoricalQuantity();

                //Select First Item and add Physical Quantity
                outputFormItem.SelectFirstItem();
                outputFormItem.AddPhysicalQuantity(itemName, quantity);

                outputFormPage = outputFormItem.BackToList();
                outputFormPage.ResetFilter();

                // Filter Output Form Created 
                outputFormPage.Filter(OutputFormPage.FilterType.SearchByNumber, ID);
                outputFormPage.Filter(OutputFormPage.FilterType.CheckAllSites, true);
                outputFormPage.Filter(OutputFormPage.FilterType.CheckAllOfTypes, true);
                outputFormPage.Filter(OutputFormPage.FilterType.DateFrom, DateTime.Now);
                outputFormPage.Filter(OutputFormPage.FilterType.DateTo, DateTime.Now.AddMonths(1));

                // Print report Output Form
                PrintReportPage reportPage = PrintOutputFormGenerique(outputFormPage, newVersionPrint, ID);

                reportPage.Purge(downloadsPath, docFileNamePdfBegin, docFileNameZipBegin);

                // cliquer sur All
                string trouve = reportPage.PrintAllZip(downloadsPath, docFileNamePdfBegin, docFileNameZipBegin);
                FileInfo fi = new FileInfo(trouve);
                fi.Refresh();
                Assert.IsTrue(fi.Exists, trouve + " non généré");

                PdfDocument document = PdfDocument.Open(fi.FullName);
                List<string> mots = new List<string>();
                foreach (Page page in document.GetPages())
                {
                    IEnumerable<Word> words = page.GetWords();
                    foreach (Word word in words)
                    {
                        mots.Add(word.Text);
                    }
                }
                // Assert
                Assert.AreNotEqual(0, mots.Count(w => w == "(" + site + ")"), "(" + site + ") non présent dans le Pdf");
                Assert.AreNotEqual(0, mots.Count(w => w == DateUtils.Now.Date.ToString("dd/MM/yyyy")), DateUtils.Now.Date.ToString("dd/MM/yyyy") + " non présent dans le Pdf");
                Assert.AreNotEqual(0, mots.Count(w => w.Contains(ID)), ID + " non présent dans le Pdf");
                Assert.IsTrue(document.GetPage(1).GetImages().Count() >= 2 && document.GetPage(1).GetImages().Count() <= 3, "Pas deux images dans le pdf (+ éventuel logo)");

            }
            finally
            {
                // Filter and Delete Output Form Created
                outputFormPage.Filter(OutputFormPage.FilterType.SearchByNumber, ID);
                outputFormPage.DeleteFirstOutputForm();
            }
        }
        private PrintReportPage PrintOutputFormGenerique(OutputFormPage outputFormPage, bool printValue, string ID)
        {
            var reportPage = outputFormPage.PrintResults(printValue);
            var isReportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();

            Assert.IsTrue(isReportGenerated, "L'application n'a pas pu générer le fichier attendu.");

            return reportPage;
        }

        /*
         * Test d'export des outputForm avec newVersionPrint = true
         */
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_ExportOutputFormListNewVersion()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string placeFrom = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string id = string.Empty;

            // Log in
            var homePage = LogInAsAdmin();

            //Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();
            outputFormPage.ClearDownloads();
            DeleteAllFileDownload();
            try
            {
                // Create New Output Form 
                var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
                outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo);
                id = outputFormCreateModalpage.GetOutputFormNumber();
                var outputFormItem = outputFormCreateModalpage.Submit();
                // Filter Item with Theo Qty Only
                outputFormItem.Filter(OutputFormItem.FilterItemType.ShowItemsWithTheoQtyOnly, true);

                var itemName = outputFormItem.GetFirstItemName();
                var quantity = outputFormItem.GetTheoricalQuantity();
                // Select First Item and Add physical Quantity
                outputFormItem.SelectFirstItem();
                outputFormItem.AddPhysicalQuantity(itemName, quantity);
                outputFormPage = outputFormItem.BackToList();
                outputFormPage.ResetFilter();
                // Filter the Output form Created 
                outputFormPage.Filter(OutputFormPage.FilterType.SearchByNumber, id);
                outputFormPage.Filter(OutputFormPage.FilterType.DateFrom, DateUtils.Now.AddDays(-3));
                outputFormPage.Filter(OutputFormPage.FilterType.DateTo, DateUtils.Now);
                // Export Result and check Result is verified
                outputFormPage.ExportResults(true);
                GenericExport genericExport = new GenericExport(WebDriver, TestContext);
                genericExport.WaitPageLoading();
                FileInfo correctDownloadedFile = genericExport.IsGenerated(fileNamePattern);
                Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");
                bool isExcelFileCorrect = genericExport.IsExcelFileCorrect(correctDownloadedFile.Name, fileNamePattern);
                Assert.IsTrue(isExcelFileCorrect, "Le nom du fichier exporté ne correspond pas au pattern général.");
                int resultNumber = genericExport.resultNumber(correctDownloadedFile.Name, OUTPUT_FORM_EXCEL_SHEET);
                Assert.AreEqual(1, resultNumber, "Les données du fichier exporté ne correspondent pas au filtre appliqué.");
                List<string> numberList = OpenXmlExcel.GetValuesInList("Number", OUTPUT_FORM_EXCEL_SHEET, correctDownloadedFile.FullName);
                bool isNumberCorrect = numberList.All(x => x.Equals(id));
                Assert.IsTrue(isNumberCorrect, MessageErreur.EXCEL_DONNEES_KO);
                List<string> siteList = OpenXmlExcel.GetValuesInList("Site", OUTPUT_FORM_EXCEL_SHEET, correctDownloadedFile.FullName);
                bool isSiteCorrect = siteList.All(x => x.Equals(site));
                Assert.IsTrue(isSiteCorrect, MessageErreur.EXCEL_DONNEES_KO);
                List<string> fromPlaceList = OpenXmlExcel.GetValuesInList("From", OUTPUT_FORM_EXCEL_SHEET, correctDownloadedFile.FullName);
                bool isFromPlaceCorrect = fromPlaceList.All(x => x.Equals(placeFrom));
                Assert.IsTrue(isFromPlaceCorrect, MessageErreur.EXCEL_DONNEES_KO);
                List<string> toPlaceList = OpenXmlExcel.GetValuesInList("To", OUTPUT_FORM_EXCEL_SHEET, correctDownloadedFile.FullName);
                bool isToPlaceCorrect = toPlaceList.All(x => x.Equals(placeTo));
                Assert.IsTrue(isToPlaceCorrect, MessageErreur.EXCEL_DONNEES_KO);
            }
            finally
            {
                // Filter and Delete The Output Form Created
                outputFormPage.ResetFilter();
                outputFormPage.Filter(OutputFormPage.FilterType.SearchByNumber, id);
                outputFormPage.DeleteFirstOutputForm();
            }
            
        }

        private void ExportGenerique(OutputFormPage outputFormPage, bool printVersion, string downloadsPath)
        {
            outputFormPage.Filter(OutputFormPage.FilterType.DateFrom, DateUtils.Now.AddDays(-3));
            outputFormPage.Filter(OutputFormPage.FilterType.DateTo, DateUtils.Now);

            DeleteAllFileDownload();

            outputFormPage.ExportResults(printVersion);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = outputFormPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

            // Récupération du nom du fichier et construction de l'URL du fichier Excel à ouvrir   
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber(OUTPUT_FORM_EXCEL_SHEET, filePath);

            //Assert
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
        }

        [Ignore]// Test d'export WMS des outputForm avec newVersionPrint = true
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_ExportWMSOFListNewVersion()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string placeFrom = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            bool newVersionPrint = true;

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();
            outputFormPage.Filter(OutputFormPage.FilterType.DateTo, DateUtils.Now.AddDays(-1));
            outputFormPage.Filter(OutputFormPage.FilterType.DateFrom, DateUtils.Now);

            if (outputFormPage.CheckTotalNumber() < 20)
            {
                var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
                outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo);
                var outputFormItem = outputFormCreateModalpage.Submit();

                outputFormItem.Filter(OutputFormItem.FilterItemType.ShowItemsWithTheoQtyOnly, true);

                var itemName = outputFormItem.GetFirstItemName();
                var quantity = outputFormItem.GetTheoricalQuantity();

                outputFormItem.SelectFirstItem();
                outputFormItem.AddPhysicalQuantity(itemName, quantity);

                outputFormPage = outputFormItem.BackToList();

                outputFormPage.ResetFilter();
                outputFormPage.Filter(OutputFormPage.FilterType.DateTo, DateUtils.Now.AddMonths(1));
                outputFormPage.Filter(OutputFormPage.FilterType.DateFrom, DateUtils.Now);
            }

            outputFormPage.ExportWMS(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = outputFormPage.GetExportWMSFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_UpdateItemOF()
        {
            string site = TestContext.Properties["Site"].ToString();
            string placeFrom = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();
            string delivery = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            HomePage homePage = LogInAsAdmin();
            OutputFormPage outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();
            OutputFormCreateModalPage outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
            outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo);
            OutputFormItem outputFormItem = outputFormCreateModalpage.Submit();
            string outputFormNumber = outputFormItem.GetOutputFormNumber();
            if (outputFormItem.GetTheoricalQuantity() == "0")
            {
                string firstItem = outputFormItem.GetFirstItemName();
                ItemGeneralInformationPage editItem = outputFormItem.EditItem(firstItem);
                string supplier = editItem.GetPackagingSupplierBySite(site);
                homePage.Navigate();
                ReceiptNotesPage receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
                ReceiptNotesCreateModalPage receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
                receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, delivery));
                ReceiptNotesItem receiptNotesItem = receiptNotesCreateModalpage.Submit();
                receiptNotesItem.SetFirstReceivedQuantity("5");
                ReceiptNotesQualityChecks checksTab = receiptNotesItem.ClickOnChecksTab();
                checksTab.SetNotApplicable();
                checksTab.DeliveryAccepted();
                checksTab.ValidateQualityChecks();
                checksTab.WaitPageLoading();
                homePage.Navigate();
                outputFormPage = homePage.GoToWarehouse_OutputFormPage();
                outputFormPage.ResetFilter();
                outputFormPage.Filter(OutputFormPage.FilterType.SearchByNumber, outputFormNumber);
                outputFormItem = outputFormPage.ClickFirstOF();
            }
            outputFormItem.Filter(OutputFormItem.FilterItemType.ShowItemsWithTheoQtyOnly, true);
            string itemName = outputFormItem.GetFirstItemName();
            string quantity = outputFormItem.GetTheoricalQuantity();
            outputFormItem.SelectFirstItem();
            outputFormItem.AddPhysicalQuantity(itemName, quantity);
            try
            {
                ItemGeneralInformationPage itemPage = outputFormItem.EditItem(itemName);
                itemPage.WaitPageLoading();
                itemPage.WaitLoading();
                itemPage.SetName(itemName + " TEST");
                itemPage.WaitPageLoading();
                itemPage.WaitLoading();
                Assert.IsTrue(outputFormItem.ElementIsVisible(), "La page de l'item ne s'ouvre pas sur un autre onglet");
                itemPage.Close();
                outputFormItem.Refresh();
                outputFormItem.SelectFirstItem();
            }
            finally
            {
                ItemGeneralInformationPage itemPage = outputFormItem.EditItem(itemName + " TEST");
                itemPage.WaitPageLoading();
                itemPage.SetName(itemName.Replace(" TEST", ""));
                itemPage.WaitPageLoading();
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_CommentOF()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string placeFrom = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();
            string comment = "I am a comment";
            string id = string.Empty;
            string physQty = "10";
            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();
            try
            {
                //create
                var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
                outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo);
                id = outputFormCreateModalpage.GetOutputFormNumber();
                var outputFormItem = outputFormCreateModalpage.Submit();

                var itemName = outputFormItem.GetFirstItemName();
                outputFormItem.SelectFirstItem();

                //Add valid Qty to add comment
                outputFormItem.AddPhysicalQuantity(itemName, physQty);

                // Update the first item value to activate the activation menu
                outputFormItem.AddComment(itemName, comment);

                //Assert
                var updatedComment = outputFormItem.GetComment(itemName);
                Assert.AreEqual(updatedComment, comment, "L'ajout du commentaire dans l'OF a échoué.");
            }
            finally
            {
                homePage.GoToWarehouse_OutputFormPage();
                if (!string.IsNullOrEmpty(id))
                {
                    outputFormPage.Filter(OutputFormPage.FilterType.SearchByNumber, id);
                    outputFormPage.DeleteFirstOutputForm();
                }

            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_DeleteItemOF()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string placeFrom = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();

            var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
            outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo);
            var outputFormItem = outputFormCreateModalpage.Submit();

            outputFormItem.Filter(OutputFormItem.FilterItemType.ShowItemsWithTheoQtyOnly, true);

            var itemName = outputFormItem.GetFirstItemName();
            Assert.AreEqual(outputFormItem.GetFirstItemName(), itemName, "L'outputForm possède bien des items.");

            var quantity = outputFormItem.GetTheoricalQuantity();

            outputFormItem.SelectFirstItem();
            outputFormItem.AddPhysicalQuantity(itemName, quantity);

            // Delete the first item
            outputFormItem.Filter(OutputFormItem.FilterItemType.SearchByName, itemName);
            Assert.AreNotEqual(0, outputFormItem.GetPhysicalQuantity(), "Aucune quantité n'a été ajoutée à l'item.");

            outputFormItem.SelectFirstItem();
            outputFormItem.ClickOnButtonSize();
            outputFormItem.DeleteItem(itemName);
            outputFormItem.ClickOnButtonSize();
            outputFormItem.Filter(OutputFormItem.FilterItemType.SearchByName, itemName);
            Assert.AreEqual("0", outputFormItem.GetPhysicalQuantity(), "La quantité affectée à l'item n'a pas été supprimée.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_ByGroupsFilter()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string placeFrom = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();

            var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
            outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo);
            var outputFormItem = outputFormCreateModalpage.Submit();

            var itemGroup = outputFormItem.GetFirstItemGroup();
            var groupsPage = outputFormItem.ClickOnGroupsTab();

            groupsPage.FilterByGroup(itemGroup);

            Assert.IsTrue(groupsPage.VerifyFilterGroup(itemGroup), "Le filtre des groupes n'a pas fonctionné dans l'onglet 'By groups'.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_Fill_QualityChecks()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string placeFrom = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();
            int locationArobase = TestContext.Properties["Admin_UserName"].ToString().IndexOf("@");
            string userName = TestContext.Properties["Admin_UserName"].ToString().Substring(0, locationArobase);
            string id = "";
            var decimalSeparatorValue = ",";
            string temp = "5" + decimalSeparatorValue + "00";
            string phyQty1 = "10";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();
            try
            {
                //create Output Form
                var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
                outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo);
                id = outputFormCreateModalpage.GetOutputFormNumber();
                var outputFormItem = outputFormCreateModalpage.Submit();

                //Modifier les informations
                var itemName = outputFormItem.GetFirstItemName();
                outputFormItem.SelectFirstItem();
                outputFormItem.AddPhysicalQuantity(itemName, phyQty1);
                var qualityChecks = outputFormItem.ClickOnChecksTab();
                qualityChecks.DeliveryAccepted();

                // Changer l'onglet 
                outputFormItem.OpenNewTab();

                //Assert phyQty
                outputFormItem = qualityChecks.ClickOnItemsTab();
                outputFormItem.Filter(OutputFormItem.FilterItemType.ShowItemsWithPhysQty, true);
                string phyQty2 = outputFormItem.GetPhysicalQuantity();
                Assert.AreEqual(phyQty2, phyQty1, "La quantité n'a pas été mise à jour.");

                //Assert deliveryAccepted
                qualityChecks = outputFormItem.ClickOnChecksTab();
                bool deliveryAccepted = qualityChecks.IsDeliveryAccepted();
                Assert.IsTrue(deliveryAccepted, "Le statut n'a pas été mis à jour.");
            }
            finally
            {
                homePage.GoToWarehouse_OutputFormPage();
                if (!String.IsNullOrEmpty(id))
                {
                    outputFormPage.Filter(OutputFormPage.FilterType.SearchByNumber, id);
                    outputFormPage.DeleteFirstOutputForm();

                }
            }
        }

        [Ignore]
        [TestMethod]
        [Priority(3)]
        [Timeout(_timeout)]
        public void WA_OF_SageManual_Details_ExportSAGENewVersion()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string placeFrom = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceToOFSage"].ToString();
            string itemName = TestContext.Properties["Item_OutputForm"].ToString();
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string currency = TestContext.Properties["Currency"].ToString();

            bool newVersionPrint = true;

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Config export sage manuel
            homePage.SetSageAutoEnabled(site, false);

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            //Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();

            outputFormPage.ClearDownloads();

            var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
            outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo, true);
            var outputFormItem = outputFormCreateModalpage.Submit();

            outputFormItem.Filter(OutputFormItem.FilterItemType.SearchByName, itemName);
            var quantity = outputFormItem.GetTheoricalQuantity();

            outputFormItem.SelectFirstItem();
            outputFormItem.AddPhysicalQuantity(itemName, quantity);
            outputFormItem.Validate();

            var outputFormFooterPage = outputFormItem.ClickOnFooterTab();
            double montantOF = outputFormFooterPage.GetOutputFormTotalHT(currency, decimalSeparatorValue);

            outputFormItem = outputFormFooterPage.ClickOnItemsTab();

            outputFormItem.ManualExportSAGE(newVersionPrint);

            try
            {
                // On récupère les fichiers du répertoire de téléchargement
                DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
                FileInfo[] taskFiles = taskDirectory.GetFiles();

                var correctDownloadedFile = outputFormItem.GetExportSAGEFile(taskFiles);
                Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

                // Récupération du nom du fichier et construction de l'URL du fichier txt à exploiter   
                var fileName = correctDownloadedFile.Name;
                var filePath = Path.Combine(downloadsPath, fileName);

                // On n'exploite que les lignes avec contenu "général" --> "G"
                double[] contenuFichier = ExploitTextFiles.VerifySAGEFileContent(filePath, "G", decimalSeparatorValue);

                Assert.AreEqual(contenuFichier[0].ToString(), contenuFichier[1].ToString(), "Le contenu du fichier SAGE est incorrect.");
                Assert.AreEqual(montantOF.ToString(), contenuFichier[0].ToString(), "Le montant issu du fichier SAGE n'est pas égal au montant de l'Output Form défini dans l'application.");
            }
            finally
            {
                //Creer un nouvel inventaire pour rajouter une quantité à l'item
                CreateInventory(homePage, site, placeFrom, itemName);
            }
        }

        [Ignore]
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_SageManual_Details_ExportSAGEKONewVersion()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string placeFrom = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceToOFSage"].ToString();
            string itemName = TestContext.Properties["Item_OutputForm"].ToString();
            string journalOutputForm = TestContext.Properties["Journal_OF"].ToString();

            bool newVersionPrint = true;

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Config export sage manuel
            homePage.SetSageAutoEnabled(site, false);

            try
            {
                // Config journal KO pour le test
                VerifyAccountingJournal(homePage, site, "");

                //Act
                var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
                outputFormPage.ResetFilter();

                if (newVersionPrint)
                {
                    outputFormPage.ClearDownloads();
                }

                var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
                outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo, true);
                String ID = outputFormCreateModalpage.GetOutputFormNumber();
                var outputFormItem = outputFormCreateModalpage.Submit();

                outputFormItem.Filter(OutputFormItem.FilterItemType.SearchByName, itemName);
                var quantity = outputFormItem.GetTheoricalQuantity();

                outputFormItem.SelectFirstItem();
                outputFormItem.AddPhysicalQuantity(itemName, quantity);

                outputFormItem.Validate();

                outputFormItem.ManualExportSageError(newVersionPrint);

                outputFormPage = outputFormItem.BackToList();
                outputFormPage.ResetFilter();
                outputFormPage.Filter(OutputFormPage.FilterType.SearchByNumber, ID);

                Assert.IsFalse(outputFormPage.IsSentToSAGE(), "L'output form est notifiée comme envoyée vers le SAGE malgré l'erreur.");

            }
            finally
            {
                VerifyAccountingJournal(homePage, site, journalOutputForm);

                //Creer un nouvel inventaire pour rajouter une quantité à l'item
                CreateInventory(homePage, site, placeFrom, itemName);
            }
        }

        [Ignore]
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_Details_EnableSAGEExport()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string placeFrom = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceToOFSage"].ToString();

            bool newVersionPrint = true;

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Config export sage manuel
            homePage.SetSageAutoEnabled(site, false);

            //Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();

            if (newVersionPrint)
            {
                outputFormPage.ClearDownloads();
            }

            var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
            outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo, true);
            var outputFormItem = outputFormCreateModalpage.Submit();

            outputFormItem.Filter(OutputFormItem.FilterItemType.ShowItemsWithTheoQtyOnly, true);

            var itemName = outputFormItem.GetFirstItemName();
            var quantity = outputFormItem.GetTheoricalQuantity();

            outputFormItem.SelectFirstItem();
            outputFormItem.AddPhysicalQuantity(itemName, quantity);

            try
            {
                outputFormItem.Validate();

                bool isEnableOK = outputFormItem.CanClickOnEnableSAGE();
                Assert.IsFalse(isEnableOK, "Il est possible de cliquer sur la fonctionnalité 'Enable export for SAGE' "
                    + "pour un OF non envoyé au SAGE.");

                outputFormItem.ManualExportSAGE(newVersionPrint);

                bool isSAGEOK = outputFormItem.CanClickOnSAGE();
                isEnableOK = outputFormItem.CanClickOnEnableSAGE();

                Assert.IsFalse(isSAGEOK, "Il est possible de cliquer sur la fonctionnalité 'Export for SAGE' "
                    + "après avoir réalisé un export SAGE.");

                Assert.IsTrue(isEnableOK, "Il est impossible de cliquer sur la fonctionnalité 'Enable export for SAGE' "
                    + "pour un OF envoyé au SAGE.");

                outputFormItem.EnableExportForSage();

                isSAGEOK = outputFormItem.CanClickOnSAGE();
                isEnableOK = outputFormItem.CanClickOnEnableSAGE();

                Assert.IsTrue(isSAGEOK, "Il est impossible de cliquer sur la fonctionnalité 'Export for SAGE' "
                    + "après avoir cliqué un export SAGE.");

                Assert.IsFalse(isEnableOK, "Il est possible de cliquer sur la fonctionnalité 'Enable export for SAGE' "
                    + "pour un OF envoyé au SAGE.");

            }
            finally
            {
                //Creer un nouvel inventaire pour rajouter une quantité à l'item
                CreateInventory(homePage, site, placeFrom, itemName);
            }
        }

        [Ignore]
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_SageManual_Index_ExportSAGENewVersion()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string placeFrom = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceToOFSage"].ToString();
            string itemName = TestContext.Properties["Item_OutputForm"].ToString();
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string currency = TestContext.Properties["Currency"].ToString();

            bool newVersionPrint = true;

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Config export sage manuel
            homePage.SetSageAutoEnabled(site, false);

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            //Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();

            if (newVersionPrint)
            {
                outputFormPage.ClearDownloads();
            }

            var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
            outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo, true);
            String ID = outputFormCreateModalpage.GetOutputFormNumber();
            var outputFormItem = outputFormCreateModalpage.Submit();

            outputFormItem.Filter(OutputFormItem.FilterItemType.SearchByName, itemName);
            var quantity = outputFormItem.GetTheoricalQuantity();

            outputFormItem.SelectFirstItem();
            outputFormItem.AddPhysicalQuantity(itemName, quantity);

            try
            {
                outputFormItem.Validate();

                var outputFormFooterPage = outputFormItem.ClickOnFooterTab();
                double montantOF = outputFormFooterPage.GetOutputFormTotalHT(currency, decimalSeparatorValue);

                outputFormItem = outputFormFooterPage.ClickOnItemsTab();
                outputFormPage = outputFormItem.BackToList();

                outputFormPage.Filter(OutputFormPage.FilterType.SearchByNumber, ID);
                outputFormPage.Filter(OutputFormPage.FilterType.ShowNotValidated, false);
                outputFormPage.Filter(OutputFormPage.FilterType.CheckAllSites, true);
                outputFormPage.Filter(OutputFormPage.FilterType.CheckAllOfTypes, true);

                outputFormPage.ManualExportSage(newVersionPrint);
                // On récupère les fichiers du répertoire de téléchargement
                DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
                FileInfo[] taskFiles = taskDirectory.GetFiles();

                var correctDownloadedFile = outputFormPage.GetExportSAGEFile(taskFiles);
                Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

                // Récupération du nom du fichier et construction de l'URL du fichier txt à exploiter   
                var fileName = correctDownloadedFile.Name;
                var filePath = Path.Combine(downloadsPath, fileName);

                // On n'exploite que les lignes avec contenu "général" --> "G"
                double[] contenuFichier = ExploitTextFiles.VerifySAGEFileContent(filePath, "G", decimalSeparatorValue);

                Assert.AreEqual(contenuFichier[0].ToString(), contenuFichier[1].ToString(), "Le contenu du fichier SAGE est incorrect.");
                Assert.AreEqual(montantOF.ToString(), contenuFichier[0].ToString(), "Le montant issu du fichier SAGE n'est pas égal au montant de l'Output Form défini dans l'application.");
            }
            finally
            {
                //Creer un nouvel inventaire pour rajouter une quantité à l'item
                CreateInventory(homePage, site, placeFrom, itemName);
            }
        }

        [Ignore]
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_SageManual_Index_ExportSAGEKONewVersion()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string placeFrom = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceToOFSage"].ToString();
            string itemName = TestContext.Properties["Item_OutputForm"].ToString();
            string journalOutputForm = TestContext.Properties["Journal_OF"].ToString();

            bool newVersionPrint = true;

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Config export sage manuel
            homePage.SetSageAutoEnabled(site, false);

            try
            {
                // Config journal KO pour le test
                VerifyAccountingJournal(homePage, site, "");

                //Act
                var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
                outputFormPage.ResetFilter();

                if (newVersionPrint)
                {
                    outputFormPage.ClearDownloads();
                }

                var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
                outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo, true);
                String ID = outputFormCreateModalpage.GetOutputFormNumber();
                var outputFormItem = outputFormCreateModalpage.Submit();

                outputFormItem.Filter(OutputFormItem.FilterItemType.SearchByName, itemName);
                var quantity = outputFormItem.GetTheoricalQuantity();

                outputFormItem.SelectFirstItem();
                outputFormItem.AddPhysicalQuantity(itemName, quantity);
                outputFormItem.Validate();

                outputFormPage = outputFormItem.BackToList();
                outputFormPage.ResetFilter();
                outputFormPage.Filter(OutputFormPage.FilterType.SearchByNumber, ID);
                outputFormPage.Filter(OutputFormPage.FilterType.ShowNotValidated, false);
                outputFormPage.Filter(OutputFormPage.FilterType.CheckAllSites, true);
                outputFormPage.Filter(OutputFormPage.FilterType.CheckAllOfTypes, true);

                outputFormPage.ManualExportSageError(newVersionPrint);

                WebDriver.Navigate().Refresh();
                outputFormPage.Filter(OutputFormPage.FilterType.SearchByNumber, ID);
                outputFormPage.Filter(OutputFormPage.FilterType.ShowNotValidated, false);

                Assert.IsFalse(outputFormPage.IsSentToSAGE(), "L'output form est notifiée comme envoyée vers le SAGE malgré l'erreur.");

            }
            finally
            {
                VerifyAccountingJournal(homePage, site, journalOutputForm);

                //Creer un nouvel inventaire pour rajouter une quantité à l'item
                CreateInventory(homePage, site, placeFrom, itemName);
            }
        }

        [Ignore]
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_Index_EnableSAGEExport()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string placeFrom = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceToOFSage"].ToString();

            bool newVersionPrint = true;

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Config export sage manuel
            homePage.SetSageAutoEnabled(site, false);

            //Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();

            if (newVersionPrint)
            {
                outputFormPage.ClearDownloads();
            }

            var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
            outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo, true);
            String ID = outputFormCreateModalpage.GetOutputFormNumber();
            var outputFormItem = outputFormCreateModalpage.Submit();

            outputFormItem.Filter(OutputFormItem.FilterItemType.ShowItemsWithTheoQtyOnly, true);

            var itemName = outputFormItem.GetFirstItemName();
            var quantity = outputFormItem.GetTheoricalQuantity();

            outputFormItem.SelectFirstItem();
            outputFormItem.AddPhysicalQuantity(itemName, quantity);

            try
            {
                outputFormItem.Validate();

                outputFormPage = outputFormItem.BackToList();
                outputFormPage.ResetFilter();
                outputFormPage.Filter(OutputFormPage.FilterType.SearchByNumber, ID);
                outputFormPage.Filter(OutputFormPage.FilterType.ShowNotValidated, false);
                outputFormPage.Filter(OutputFormPage.FilterType.CheckAllSites, true);
                outputFormPage.Filter(OutputFormPage.FilterType.CheckAllOfTypes, true);

                outputFormPage.ManualExportSage(newVersionPrint);
                outputFormItem = outputFormPage.SelectFirstOutputForm();

                Assert.IsTrue(outputFormItem.CanClickOnEnableSAGE(), "Il n'est pas possible de cliquer sur la fonctionnalité 'Enable export for SAGE' "
                + "après avoir exporté l'OF vers SAGE depuis la page Index.");

                Assert.IsFalse(outputFormItem.CanClickOnSAGE(), "Il est possible de cliquer sur la fonctionnalité 'Export for SAGE' "
                    + "après avoir exporté l'OF vers SAGE depuis la page Index.");

                outputFormPage = outputFormItem.BackToList();
                outputFormPage.EnableExportForSage();
                outputFormItem = outputFormPage.SelectFirstOutputForm();

                Assert.IsFalse(outputFormItem.CanClickOnEnableSAGE(), "Il est possible de cliquer sur la fonctionnalité 'Enable export for SAGE' "
                + "après avoir activé la fonctionnalité depuis la page Index.");

                Assert.IsTrue(outputFormItem.CanClickOnSAGE(), "Il est impossible de cliquer sur la fonctionnalité 'Export for SAGE' "
                    + "après avoir cliqué sur 'Enable export for SAGE' depuis la page Index.");
            }
            finally
            {
                //Creer un nouvel inventaire pour rajouter une quantité à l'item
                CreateInventory(homePage, site, placeFrom, itemName);
            }
        }

        [Ignore]
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_Details_GenerateSageTxt()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string placeFrom = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceToOFSage"].ToString();
            string itemName = TestContext.Properties["Item_OutputForm"].ToString();
            string currency = TestContext.Properties["Currency"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            bool newVersionPrint = true;

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            // Config pour export SAGE auto mais pas pour le site
            homePage.SetSageAutoEnabled(site, true, "Output form", false);

            try
            {
                //Act
                var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
                outputFormPage.ResetFilter();

                if (newVersionPrint)
                {
                    outputFormPage.ClearDownloads();
                }

                var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
                outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo, true);
                String ID = outputFormCreateModalpage.GetOutputFormNumber();
                var outputFormItem = outputFormCreateModalpage.Submit();

                outputFormItem.Filter(OutputFormItem.FilterItemType.SearchByName, itemName);
                var quantity = outputFormItem.GetTheoricalQuantity();

                outputFormItem.SelectFirstItem();
                outputFormItem.AddPhysicalQuantity(itemName, quantity);

                outputFormItem.Validate();

                var outputFormFooterPage = outputFormItem.ClickOnFooterTab();
                double montantOF = outputFormFooterPage.GetOutputFormTotalHT(currency, decimalSeparatorValue);

                // Calcul du montant de la facture transmise à TL
                var inventoryAccounting = outputFormFooterPage.ClickOnAccountingTab();
                double montantGlobal = inventoryAccounting.GetOutputFormGrossAmount("G", decimalSeparatorValue);
                double montantDetaille = inventoryAccounting.GetOutputFormDetailAmount("G", decimalSeparatorValue);

                outputFormItem = inventoryAccounting.ClickOnItems();
                outputFormItem.ManualExportSAGE(newVersionPrint);

                // On récupère les fichiers du répertoire de téléchargement
                DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
                FileInfo[] taskFiles = taskDirectory.GetFiles();

                var correctDownloadedFile = outputFormItem.GetExportSAGEFile(taskFiles);
                Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

                // Récupération du nom du fichier et construction de l'URL du fichier txt à exploiter   
                var fileName = correctDownloadedFile.Name;
                var filePath = Path.Combine(downloadsPath, fileName);

                // On n'exploite que les lignes avec contenu "général" --> "G"
                double[] contenuFichier = ExploitTextFiles.VerifySAGEFileContent(filePath, "G", decimalSeparatorValue);

                outputFormPage = outputFormItem.BackToList();
                outputFormPage.ResetFilter();
                outputFormPage.Filter(OutputFormPage.FilterType.SearchByNumber, ID);
                outputFormPage.Filter(OutputFormPage.FilterType.ShowNotValidated, false);
                outputFormPage.Filter(OutputFormPage.FilterType.CheckAllSites, true);
                outputFormPage.Filter(OutputFormPage.FilterType.CheckAllOfTypes, true);

                Assert.AreEqual(contenuFichier[0].ToString(), contenuFichier[1].ToString(), "Le contenu du fichier SAGE généré est incorrect.");
                Assert.AreEqual(montantOF.ToString(), contenuFichier[0].ToString(), "Le montant issu du fichier SAGE n'est pas égal au montant de l'OF défini dans l'application.");
                Assert.AreEqual(montantGlobal.ToString(), montantDetaille.ToString(), "Les montants AmountDebit et AmountCredit de l'OF envoyée vers SAGE ne sont pas les mêmes dans l'onglet Accounting.");
                Assert.AreEqual(montantOF.ToString(), montantGlobal.ToString(), "Le montant issu du fichier SAGE n'est pas égal au montant de l'OF défini dans l'application.");
                Assert.IsTrue(outputFormPage.IsSentToSageManually(), "L'inventory n'a pas été envoyée manuellement vers le SAGE.");
            }
            finally
            {
                homePage.SetSageAutoEnabled(site, false);

                //Creer un nouvel inventaire pour rajouter une quantité à l'item
                CreateInventory(homePage, site, placeFrom, itemName);
            }
        }

        [Ignore]
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_Index_GenerateSageTxt()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string placeFrom = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceToOFSage"].ToString();
            string itemName = TestContext.Properties["Item_OutputForm"].ToString();
            string currency = TestContext.Properties["Currency"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            bool newVersionPrint = true;

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            // Config pour export SAGE auto mais pas pour le site
            homePage.SetSageAutoEnabled(site, true, "Output form", false);

            try
            {
                //Act
                //Creer un nouvel inventaire pour avoir un th qty
                CreateInventory(homePage, site, placeFrom, itemName);

                var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
                outputFormPage.ResetFilter();

                if (newVersionPrint)
                {
                    outputFormPage.ClearDownloads();
                }

                var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
                outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo, true);
                String ID = outputFormCreateModalpage.GetOutputFormNumber();
                var outputFormItem = outputFormCreateModalpage.Submit();

                outputFormItem.Filter(OutputFormItem.FilterItemType.SearchByName, itemName);
                var quantity = outputFormItem.GetTheoricalQuantity();

                outputFormItem.SelectFirstItem();
                outputFormItem.AddPhysicalQuantity(itemName, quantity);

                outputFormItem.Validate();

                var outputFormFooterPage = outputFormItem.ClickOnFooterTab();
                double montantOF = outputFormFooterPage.GetOutputFormTotalHT(currency, decimalSeparatorValue);

                // Calcul du montant de la facture transmise à TL
                var inventoryAccounting = outputFormFooterPage.ClickOnAccountingTab();
                double montantGlobal = inventoryAccounting.GetOutputFormGrossAmount("G", decimalSeparatorValue);
                double montantDetaille = inventoryAccounting.GetOutputFormDetailAmount("G", decimalSeparatorValue);

                outputFormPage = inventoryAccounting.BackToList();
                outputFormPage.ResetFilter();
                outputFormPage.Filter(OutputFormPage.FilterType.SearchByNumber, ID);
                // provoque le rond busy
                outputFormPage.Filter(OutputFormPage.FilterType.ShowNotValidated, true);
                outputFormPage.Filter(OutputFormPage.FilterType.ShowNotValidated, false);

                outputFormPage.GenerateSageTxt();

                // On récupère les fichiers du répertoire de téléchargement
                DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
                FileInfo[] taskFiles = taskDirectory.GetFiles();

                var correctDownloadedFile = outputFormItem.GetExportSAGEFile(taskFiles);
                Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

                // Récupération du nom du fichier et construction de l'URL du fichier txt à exploiter   
                var fileName = correctDownloadedFile.Name;
                var filePath = Path.Combine(downloadsPath, fileName);

                // On n'exploite que les lignes avec contenu "général" --> "G"
                double[] contenuFichier = ExploitTextFiles.VerifySAGEFileContent(filePath, "G", decimalSeparatorValue);

                Assert.AreEqual(contenuFichier[0].ToString(), contenuFichier[1].ToString(), "Le contenu du fichier SAGE généré est incorrect.");
                Assert.AreEqual(montantOF.ToString(), contenuFichier[0].ToString(), "Le montant issu du fichier SAGE n'est pas égal au montant de la RN défini dans l'application.");
                Assert.AreEqual(montantGlobal.ToString(), montantDetaille.ToString(), "Les montants AmountDebit et AmountCredit de la RN envoyée vers SAGE ne sont pas les mêmes dans l'onglet Accounting.");
                Assert.AreEqual(montantOF.ToString(), montantGlobal.ToString(), "Le montant issu du fichier SAGE n'est pas égal au montant de la RN défini dans l'application.");

                WebDriver.Navigate().Refresh();
                outputFormPage.ResetFilter();
                outputFormPage.Filter(OutputFormPage.FilterType.SearchByNumber, ID);
                // provoque le rond busy
                outputFormPage.Filter(OutputFormPage.FilterType.ShowNotValidated, true);
                outputFormPage.Filter(OutputFormPage.FilterType.ShowNotValidated, false);

                Assert.IsTrue(outputFormPage.IsSentToSageManually(), "L'inventory n'a pas été envoyée manuellement vers le SAGE.");
            }
            finally
            {
                homePage.SetSageAutoEnabled(site, false);

                //Creer un nouvel inventaire pour rajouter une quantité à l'item
                CreateInventory(homePage, site, placeFrom, itemName);
            }
        }

        [Ignore] // aucun pays n'utilise SAGE AUTO
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_SageAuto_ExportSageItemOK()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string placeFrom = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceToOFSage"].ToString();
            string itemName = TestContext.Properties["Item_OutputForm"].ToString();
            string currency = TestContext.Properties["Currency"].ToString();

            bool newVersionPrint = true;

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            // Config pour export SAGE auto mais pas pour le site
            homePage.SetSageAutoEnabled(site, true, "Output form");

            try
            {
                //Act
                var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
                outputFormPage.ResetFilter();

                if (newVersionPrint)
                {
                    outputFormPage.ClearDownloads();
                }

                var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
                outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo, true);
                var outputFormItem = outputFormCreateModalpage.Submit();

                outputFormItem.Filter(OutputFormItem.FilterItemType.SearchByName, itemName);
                var quantity = outputFormItem.GetTheoricalQuantity();

                outputFormItem.SelectFirstItem();
                outputFormItem.AddPhysicalQuantity(itemName, quantity);

                outputFormItem.Validate();

                var outputFormFooterPage = outputFormItem.ClickOnFooterTab();
                double montantOF = outputFormFooterPage.GetOutputFormTotalHT(currency, decimalSeparatorValue);

                // Calcul du montant de la facture transmise à TL
                var inventoryAccounting = outputFormFooterPage.ClickOnAccountingTab();
                double montantGlobal = inventoryAccounting.GetOutputFormGrossAmount("G", decimalSeparatorValue);
                double montantDetaille = inventoryAccounting.GetOutputFormDetailAmount("G", decimalSeparatorValue);

                Assert.AreEqual(montantGlobal.ToString(), montantDetaille.ToString(), "Les montants AmountDebit et AmountCredit de l'OF envoyée vers SAGE ne sont pas les mêmes dans l'onglet Accounting.");
                Assert.AreEqual(montantOF.ToString(), montantGlobal.ToString(), "Le montant issu du fichier SAGE n'est pas égal au montant de l'OF défini dans l'application.");

            }
            finally
            {
                homePage.SetSageAutoEnabled(site, false);

                //Creer un nouvel inventaire pour rajouter une quantité à l'item
                CreateInventory(homePage, site, placeFrom, itemName);
            }
        }

        [Ignore]
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_SageManuel_ExportSAGEItemKO_CodeJournal()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string placeFrom = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceToOFSage"].ToString();
            string itemName = TestContext.Properties["Item_OutputForm"].ToString();
            string journalOF = TestContext.Properties["Journal_OF"].ToString();

            bool newVersionPrint = true;

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Config pour export SAGE manuel mais pas pour le site
            homePage.SetSageAutoEnabled(site, false);

            try
            {
                // Parameter - Accounting --> Journal KO pour test
                VerifyAccountingJournal(homePage, site, "");

                //Act
                var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
                outputFormPage.ResetFilter();

                if (newVersionPrint)
                {
                    outputFormPage.ClearDownloads();
                }

                var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
                outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo, true);
                var outputFormItem = outputFormCreateModalpage.Submit();

                outputFormItem.Filter(OutputFormItem.FilterItemType.SearchByName, itemName);
                var quantity = outputFormItem.GetTheoricalQuantity();

                outputFormItem.SelectFirstItem();
                outputFormItem.AddPhysicalQuantity(itemName, quantity);

                outputFormItem.Validate();

                // Calcul du montant de la facture transmise à TL
                var outputFormAccounting = outputFormItem.ClickOnAccountingTab();
                string erreur = outputFormAccounting.GetErrorMessage();

                Assert.AreNotEqual("", erreur, "Aucun message d'erreur n'apparaît alors que le paramétrage relatif au code journal est KO.");
                Assert.IsTrue(erreur.Contains("no OF (Handover) journal value set"), "Le message d'erreur ne concerne pas le paramétrage Code journal.");
            }
            finally
            {
                // Parameter - Accounting --> Journal KO pour test
                VerifyAccountingJournal(homePage, site, journalOF);

                //Creer un nouvel inventaire pour rajouter une quantité à l'item
                CreateInventory(homePage, site, placeFrom, itemName);
            }
        }

        [Ignore]
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_SageManuel_ExportSAGEItemKO_SiteFromAnalyticalPlan()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string placeFrom = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceToOFSage"].ToString();
            string itemName = TestContext.Properties["Item_OutputForm"].ToString();
            string journalOutputForm = TestContext.Properties["Journal_OF"].ToString();

            bool newVersionPrint = true;

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Config pour export SAGE manuel mais pas pour le site
            homePage.SetSageAutoEnabled(site, false);

            try
            {
                bool isJournalOk = VerifyAccountingJournal(homePage, site, journalOutputForm);
                Assert.IsTrue(isJournalOk, "Pas de journal " + site);

                // Sites -- > Analytical plan et section KO pour test
                VerifySiteAnalyticalPlanSection(homePage, site, false);

                //Act
                var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
                outputFormPage.ResetFilter();

                if (newVersionPrint)
                {
                    outputFormPage.ClearDownloads();
                }

                var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
                outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo, true);
                var outputFormItem = outputFormCreateModalpage.Submit();

                outputFormItem.Filter(OutputFormItem.FilterItemType.SearchByName, itemName);
                var quantity = outputFormItem.GetTheoricalQuantity();

                outputFormItem.SelectFirstItem();
                outputFormItem.AddPhysicalQuantity(itemName, quantity);

                outputFormItem.Validate();

                // Calcul du montant de la facture transmise à TL
                var outputFormAccounting = outputFormItem.ClickOnAccountingTab();
                string erreur = outputFormAccounting.GetErrorMessage();

                Assert.AreNotEqual("", erreur, "Aucun message d'erreur n'apparaît alors que le paramétrage relatif au 'Analytic plan' du site est KO.");
                Assert.IsTrue(erreur.Contains($"Accounting analytic plan of the site {site} cannot be empty"), "Le message d'erreur [" + erreur + "] ne concerne pas le paramétrage relatif au 'Analytic plan' du site.");
            }
            finally
            {
                VerifySiteAnalyticalPlanSection(homePage, site);

                //Creer un nouvel inventaire pour rajouter une quantité à l'item
                CreateInventory(homePage, site, placeFrom, itemName);
            }
        }

        [Ignore]
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_SageManuel_ExportSAGEItemKO_SiteToAnalyticalPlan()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToOFSage"].ToString();
            string placeFrom = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceToOFSage"].ToString();
            string itemName = TestContext.Properties["Item_OutputForm"].ToString();
            string journalOutputForm = TestContext.Properties["Journal_OF"].ToString();

            bool newVersionPrint = true;

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Config pour export SAGE manuel mais pas pour le site
            homePage.SetSageAutoEnabled(site, false);

            try
            {
                bool isJournalOk = VerifyAccountingJournal(homePage, site, journalOutputForm);
                Assert.IsTrue(isJournalOk, "Pas de journal " + site);

                // Sites -- > Analytical plan et section KO pour test
                VerifySiteAnalyticalPlanSection(homePage, siteTo, false);

                //Act
                var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
                outputFormPage.ResetFilter();

                if (newVersionPrint)
                {
                    outputFormPage.ClearDownloads();
                }

                var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
                outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo, true);
                var outputFormItem = outputFormCreateModalpage.Submit();

                outputFormItem.Filter(OutputFormItem.FilterItemType.SearchByName, itemName);
                var quantity = outputFormItem.GetTheoricalQuantity();

                outputFormItem.SelectFirstItem();
                outputFormItem.AddPhysicalQuantity(itemName, quantity);

                outputFormItem.Validate();

                // Calcul du montant de la facture transmise à TL
                var outputFormAccounting = outputFormItem.ClickOnAccountingTab();
                string erreur = outputFormAccounting.GetErrorMessage();

                Assert.AreNotEqual("", erreur, "Aucun message d'erreur n'apparaît alors que le paramétrage relatif au 'Analytic plan' du site est KO.");
                Assert.IsTrue(erreur.Contains($"Accounting analytic plan of the site {siteTo} cannot be empty"), "Le message d'erreur [" + erreur + "] ne concerne pas le paramétrage relatif au 'Analytic plan' du site.");
            }
            finally
            {
                VerifySiteAnalyticalPlanSection(homePage, siteTo);

                //Creer un nouvel inventaire pour rajouter une quantité à l'item
                CreateInventory(homePage, site, placeFrom, itemName);
            }
        }

        [Ignore]//sage auto
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_SageAuto_ExportSAGEItemKO_PlaceToType()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteToOFSage"].ToString();
            string placeFrom = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceToOFSage"].ToString();
            string itemName = TestContext.Properties["Item_OutputForm"].ToString();

            bool newVersionPrint = true;

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Config pour export SAGE auto mais pas pour le site
            homePage.SetSageAutoEnabled(site, true, "Output form");

            try
            {

                // Sites -- > Analytical plan et section KO pour test
                VerifySiteConfigPlace(homePage, site, placeTo, siteTo, false);

                //Act
                var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
                outputFormPage.ResetFilter();

                if (newVersionPrint)
                {
                    outputFormPage.ClearDownloads();
                }

                var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
                outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo, true);
                var outputFormItem = outputFormCreateModalpage.Submit();

                outputFormItem.Filter(OutputFormItem.FilterItemType.SearchByName, itemName);
                var quantity = outputFormItem.GetTheoricalQuantity();

                outputFormItem.SelectFirstItem();
                outputFormItem.AddPhysicalQuantity(itemName, quantity);

                outputFormItem.Validate();

                // Calcul du montant de la facture transmise à TL
                var outputFormAccounting = outputFormItem.ClickOnAccountingTab();
                string erreur = outputFormAccounting.GetErrorMessage();

                Assert.AreNotEqual("", erreur, "Aucun message d'erreur n'apparaît alors que le paramétrage relatif au 'Analytic plan' du site est KO.");
                Assert.IsTrue(erreur.Contains("Related -To- site place is not an -Other site place- with a different site set"), "Le message d'erreur ne concerne pas le paramétrage relatif au 'place to type' du site.");
            }
            finally
            {
                VerifySiteConfigPlace(homePage, site, placeTo, siteTo);
                homePage.SetSageAutoEnabled(site, false);

                //Creer un nouvel inventaire pour rajouter une quantité à l'item
                CreateInventory(homePage, site, placeFrom, itemName);
            }
        }

        [Ignore]
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_SageManuel_ExportSAGEItemKO_NoGroupVat()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string placeFrom = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceToOFSage"].ToString();
            string itemName = TestContext.Properties["Item_OutputForm"].ToString();
            string journalOutputForm = TestContext.Properties["Journal_OF"].ToString();

            bool newVersionPrint = true;

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Config pour export SAGE manuel mais pas pour le site
            homePage.SetSageAutoEnabled(site, false);

            // Récupération du groupe de l'item
            string itemGroup = GetItemGroup(homePage, itemName);
            string vat = GetVat(homePage, itemName);

            try
            {
                bool isJournalOk = VerifyAccountingJournal(homePage, site, journalOutputForm);
                Assert.IsTrue(isJournalOk, "Pas de journal " + site);

                // Parameter - Accounting --> Group & VAT KO pour test
                VerifyGroupAndVAT(homePage, itemGroup, vat, false);

                //Act
                var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
                outputFormPage.ResetFilter();

                if (newVersionPrint)
                {
                    outputFormPage.ClearDownloads();
                }

                var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
                outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo, true);
                var outputFormItem = outputFormCreateModalpage.Submit();

                outputFormItem.Filter(OutputFormItem.FilterItemType.SearchByName, itemName);
                var quantity = outputFormItem.GetTheoricalQuantity();

                outputFormItem.SelectFirstItem();
                outputFormItem.AddPhysicalQuantity(itemName, quantity);

                outputFormItem.Validate();

                // Calcul du montant de la facture transmise à TL
                var outputFormAccounting = outputFormItem.ClickOnAccountingTab();
                string erreur = outputFormAccounting.GetErrorMessage();

                Assert.AreNotEqual("", erreur, "Aucun message d'erreur n'apparaît alors que le paramétrage relatif au 'Group & VAT' de l'item est KO.");
                Assert.IsTrue(erreur.Contains($"Outputform variation account for group {itemGroup} and VAT name {vat} is not set"), "Le message d'erreur [" + erreur + "] ne concerne pas le paramétrage relatif au 'Group & VAT' de l'item.");
            }
            finally
            {
                VerifyGroupAndVAT(homePage, itemGroup, vat);

                //Creer un nouvel inventaire pour rajouter une quantité à l'item
                CreateInventory(homePage, site, placeFrom, itemName);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_Items_Filter_Subgroup()
        {
            // Prepare items
            string site = TestContext.Properties["Site"].ToString();
            string itemName = "itemSubGroupOutputForm_" + site + "_2";
            string placeFrom = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();

            string group = TestContext.Properties["Item_Group"].ToString();
            string subgroup = "SubGroupOF";
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxName = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();

            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            try
            {
                //	Avoir activer le sous group : Parameters > Global Settings > IsSubGroupFunctionActive coché (penser à le décocher à la fin du test)
                homePage.SetSubGroupFunctionValue(true);

                //Vérifier si existe ou créer un sous group : Parameters > Production > SubGroup
                ParametersProduction productionPage = homePage.GoToParameters_ProductionPage();
                productionPage.GoToTab_SubGroup();
                productionPage.EnsureSubGroup(subgroup, group);
                homePage.Navigate();

                // Vérifier si existe ou créer un item avec le sous group crée 
                ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                itemPage.Filter(ItemPage.FilterType.Group, group);
                itemPage.Filter(ItemPage.FilterType.SubGroup, subgroup);
                if (itemPage.CheckTotalNumber() == 0)
                {
                    var itemCreateModalPage = itemPage.ItemCreatePage();
                    var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxName, prodUnit, subgroup);
                    var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                    itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);

                    itemPage = itemGeneralInformationPage.BackToList();
                    itemPage.ResetFilter();
                    itemPage.Filter(ItemPage.FilterType.Search, itemName);
                    itemPage.Filter(ItemPage.FilterType.Group, group);
                    itemPage.Filter(ItemPage.FilterType.SubGroup, subgroup);
                }
                Assert.AreEqual(itemName, itemPage.GetFirstItemName(), $"L'item {itemName} n'est pas présent dans la liste des items disponibles.");
                homePage.Navigate();

                //Créer un output form
                OutputFormPage outputFormPage = homePage.GoToWarehouse_OutputFormPage();
                OutputFormCreateModalPage outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
                outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo);
                String ID = outputFormCreateModalpage.GetOutputFormNumber();
                var outputFormItem = outputFormCreateModalpage.Submit();

                //Appliquer les filtres sur SubGroup
                outputFormItem.Filter(OutputFormItem.FilterItemType.SearchByName, itemName);
                outputFormItem.Filter(OutputFormItem.FilterItemType.ByGroup, group);
                outputFormItem.Filter(OutputFormItem.FilterItemType.BySubGroup, subgroup);
                //Vérifier que les résultats s'accordent bien au filtre appliqué
                var itemName2 = outputFormItem.GetFirstItemName();
                Assert.AreEqual(itemName, itemName2, "mauvais Item");
                var quantity = outputFormItem.GetTheoricalQuantity();

                outputFormItem.SelectFirstItem();
                outputFormItem.AddPhysicalQuantity(itemName2, quantity);

                outputFormPage = outputFormItem.BackToList();
                outputFormPage.ResetFilter();
                outputFormPage.Filter(OutputFormPage.FilterType.CheckAllSites, true);
                outputFormPage.Filter(OutputFormPage.FilterType.CheckAllOfTypes, true);
                //Assert
                Assert.AreEqual(ID, outputFormPage.GetFirstOutputFormNumber(), String.Format(MessageErreur.OBJET_NON_CREE, "L'output form"));
            }
            finally
            {
                var screenshot = WebDriver.TakeScreenshot();
                screenshot.SaveAsFile($"{TestContext.TestName}-SubGroup.png");
                TestContext.AddResultFile($"{TestContext.TestName}-SubGroup.png");

                homePage.Navigate();
                homePage.SetSubGroupFunctionValue(false);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_Items_ItemExpiryDate()
        {
            string site = "FUE";
            string placeFrom = "Economato";
            string placeTo = "Produccion";
            string itemName = "itemOutputExpiryForm_" + site;

            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxName = TestContext.Properties["Item_TaxType"].ToString();
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

            //Etre sur un output form non validé
            OutputFormPage outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.Filter(OutputFormPage.FilterType.ShowNotValidated, true);

            homePage.Navigate();
            // nouveau itemName
            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            itemPage.Filter(ItemPage.FilterType.Group, group);
            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxName, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);

                itemPage = itemGeneralInformationPage.BackToList();
                itemPage.ResetFilter();
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
            }
            Assert.AreEqual(itemName, itemPage.GetFirstItemName(), $"L'item {itemName} n'est pas présent dans la liste des items disponibles.");
            homePage.Navigate();

            // On vérifie que la quantité de l'item est supérieure à 0
            var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            inventoriesPage.ResetFilters();

            var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
            inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, placeFrom, true);
            var inventoryDetail = inventoryCreateModalpage.Submit();

            inventoryDetail.Filter(InventoryItem.FilterItemType.SearchByName, itemName);

            string theoricalQty = inventoryDetail.GetTheoricalQuantity();

            if (theoricalQty.Equals("0"))
            {

                inventoryDetail.SelectFirstItem();
                inventoryDetail.AddPhysicalQuantity(itemName, "10");

                var validateInventory = inventoryDetail.Validate();
                validateInventory.ValidatePartialInventory();

                inventoriesPage = inventoryDetail.BackToList();
                inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
                inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, placeFrom, true);
                inventoryDetail = inventoryCreateModalpage.Submit();

                inventoryDetail.Filter(InventoryItem.FilterItemType.SearchByName, itemName);

                theoricalQty = inventoryDetail.GetTheoricalQuantity();
            }
            Assert.IsTrue(long.Parse(theoricalQty) > 0, theoricalQty + " Mauvais theoricalQty");
            homePage.Navigate();
            outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();

            //1. Cliquer sur le expiry date du premier item
            outputFormPage.Filter(OutputFormPage.FilterType.ShowNotValidated, true);
            // create car site+placeFrom placeTo // MAD  FROM MAD4 TO PRODUCCIÓN
            OutputFormCreateModalPage createModalPage = outputFormPage.OutputFormCreatePage();
            createModalPage.FillField_CreatNewOutputForm(DateUtils.Now.Date, site, placeFrom, placeTo);
            string ofNumber = createModalPage.GetOutputFormNumber();
            OutputFormItem ofItem = createModalPage.Submit();
            Thread.Sleep(1000);
            ofItem.BackToList();
            outputFormPage.ResetFilter();
            outputFormPage.Filter(OutputFormPage.FilterType.SearchByNumber, ofNumber);
            outputFormPage.Filter(OutputFormPage.FilterType.CheckAllSites, true);
            outputFormPage.Filter(OutputFormPage.FilterType.CheckAllOfTypes, true);
            OutputFormItem outputForm = outputFormPage.SelectFirstOutputForm();
            outputForm.ResetFilter();
            outputForm.Filter(OutputFormItem.FilterItemType.SearchByName, itemName);
            Assert.AreEqual(itemName, outputForm.GetFirstItemName());
            if (outputForm.IsWithEditedPhysQty())
            {
                outputForm.Filter(OutputFormItem.FilterItemType.ShowItemsWithPhysQty, false);
            }
            outputForm.Filter(OutputFormItem.FilterItemType.SearchByName, itemName);
            outputForm.SelectFirstItem();
            outputForm.AddPhysicalQuantity(itemName, "10");
            // animation
            outputForm.WaitPageLoading();
            string name = outputForm.GetFirstItemName();

            OutputFormExpiry ExpiryDate = outputForm.ClickExpiryDate();

            //-vérifier le nom de l'item
            string ExpiryName = ExpiryDate.GetExpiryName();
            if (string.IsNullOrEmpty(ExpiryName))
            {
                ExpiryDate.CreateFirstRow(name);
                ExpiryName = ExpiryDate.GetExpiryName();
            }

            Assert.AreEqual(name, ExpiryName);
            //- vérifier que le total quantity est égal à la physical quantity de l'item
            string ExpiryTotalQty = ExpiryDate.GetExpiryTotalQty();
            Assert.AreEqual("10", ExpiryTotalQty, "Total quantity différent de physical quantity");

            //2. ajouter une nouvelle expiry date
            // - quantity supérieur au total quatity --> message d'erreur 'The sum of individual quantities is greater than the total output quantity.'
            ExpiryDate.CreateFirstRow("40");
            string errorMessage = "The sum of individual quantities is greater than the total output quantity.";
            Assert.IsTrue(ExpiryDate.HasMessageError(errorMessage), "Pas de message d'erreur");

            // - save-- > logo devenu vert et données sauvées bien dans la popup
            ExpiryDate.CreateFirstRow("10");
            ExpiryDate.Save();
            Thread.Sleep(2000);
            Assert.IsTrue(ExpiryDate.CheckGreenIcon(), "new : icone non verte");

            //3. recliquer sur expiry date et supprimer --> logo redevenu comme avant et plus de date dans la popup
            ExpiryDate = outputForm.ClickExpiryDate();
            ExpiryDate.CreateFirstRow(null);
            ExpiryDate.Save();

            //La valeur "Expiry date" est mise à jour et un symbole apparait à côté du nom.
            Thread.Sleep(2000);
            Assert.IsFalse(ExpiryDate.CheckGreenIcon(), "delete : icone verte");

            // FIXME on le remet à green en sortant (bug ?)
            ExpiryDate = outputForm.ClickExpiryDate();
            ExpiryDate.CreateFirstRow("10");
            ExpiryDate.Save();
        }

        [Ignore]
        [TestMethod]
        [Priority(4)]
        [Timeout(_timeout)]
        public void WA_OF_Filter_ExportedForSageManually()
        {
            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();
            outputFormPage.Filter(OutputFormPage.FilterType.CheckAllSites, true);
            outputFormPage.Filter(OutputFormPage.FilterType.CheckAllOfTypes, true);
            outputFormPage.PageSize("100");
            int nbLignes = outputFormPage.CheckTotalNumber();
            Assert.IsTrue(nbLignes > 0, "pas d'ouput form");
            //Appliquer les filtres sur exported for sage manually
            outputFormPage.Filter(OutputFormPage.FilterType.ShowGroup, true, OutputFormPage.TypeOfShowGroup.ExportedForSageManually);
            int nbLignesExported = outputFormPage.CheckTotalNumber();
            Assert.IsTrue(nbLignesExported > 0, "pas d'ouput form ExportedForSageManually");
            //Vérifier que les résultats s'accordent bien au filtre appliqué
            var listIconeDisquette = WebDriver.FindElements(By.XPath("//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[1]/span"));
            Assert.IsTrue(listIconeDisquette.Count > 0);
            foreach (var disquette in listIconeDisquette)
            {
                Assert.AreEqual("Accounted (sent to SAGE manually)", disquette.GetAttribute("title"), "Mauvaise disquette");
            }
            // nombre de disquette par rapport au nombre de lignes affichées
            if (nbLignesExported > 100)
            {
                Assert.AreEqual(100, listIconeDisquette.Count, "Manque des disquettes");
            }
            else
            {
                Assert.AreEqual(nbLignesExported, listIconeDisquette.Count, "Manque des disquettes (cas 2)");
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_DetailByGroupsFilter()
        {
            //prepare
            string site = TestContext.Properties["Site"].ToString();
            string placeFrom = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();
            string pysQty = "5";
            string outputFormID = string.Empty;

            //arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();
            try
            {
                // Create outputForm
                var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
                outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo);
                outputFormID = outputFormCreateModalpage.GetOutputFormNumber();
                var outputFormItem = outputFormCreateModalpage.Submit();
                outputFormItem.SelectFirstItem();
                string itemName = outputFormItem.GetFirstItemName();
                outputFormItem.AddPhysicalQuantity(itemName, pysQty);
                var vatAmountItemsDecimal = outputFormPage.GetVATAmountFromItemsSubMenu();
                // go to group by submenu
                outputFormPage.GoToByGroupSubMenu();
                var totalVATByGroupMenuDecimal = outputFormPage.GetVATByGroupName();
                //Assert
                Assert.AreEqual(vatAmountItemsDecimal, totalVATByGroupMenuDecimal, "total w/o VAT (Tab Items) est différent du somme total w/o VAT (Tab By groups)");
            }
            finally
            {
                homePage.GoToWarehouse_OutputFormPage();
                if (!string.IsNullOrEmpty(outputFormID))
                {
                    outputFormPage.Filter(OutputFormPage.FilterType.SearchByNumber, outputFormID);
                    outputFormPage.DeleteFirstOutputForm();
                }
            }

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_DetailVerifyFooter()
        {
            string qty = "5";
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();
            outputFormPage.Filter(OutputFormPage.FilterType.ShowNotValidated, true);
            outputFormPage.Filter(OutputFormPage.FilterType.CheckAllSites, true);
            outputFormPage.Filter(OutputFormPage.FilterType.CheckAllOfTypes, true);
            OutputFormItem items = outputFormPage.ClickFirstOF();
            items.SelectFirstItem();
            string itemName = items.GetFirstItemName();
            items.AddPhysicalQuantity(itemName, qty);
            var vatAmountItemsDecimal = outputFormPage.GetVATAmountFromItemsSubMenu();
            // go to footer by submenu
            outputFormPage.GoToFooterSubMenu();
            var vatAmountFooterDecimal = outputFormPage.GetVATAmountFromFooterSubMenu();
            Assert.AreEqual(vatAmountItemsDecimal, vatAmountFooterDecimal, "total w/o VAT (Tab Items) est différent du somme total w/o VAT (Tab Footer)");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_CreateItemForOutputFormErrorMessage()
        {
            // Prepare
            string itemName = "BRANDY TORRES 10 AÑOS MINI";
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxName = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();

            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = "MIGUEL TORRES, S.A.";
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();

            // la place où on a posé notre stock
            string placeFrom = "Economato";
            string placeTo = TestContext.Properties["PlaceTo"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();
            // Act
            var decimalSeparator = homePage.GetDecimalSeparatorValue();
            // Créer un item pour les tests output form
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);

            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxName, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);

                itemPage = itemGeneralInformationPage.BackToList();
                itemPage.ResetFilter();
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
            }
            ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
            ItemStoragePage storage = generalInfo.ClickOnStorage();
            // un inventory sinon storage est vide.
            int stock = 0;
            if (storage.isElementVisible(By.XPath("//*[@id='table-itemDetailsStorage']/tbody/tr[*]/td[contains(text(),'MAD')]/../td[4]")))
            {
                var stockSite = storage.WaitForElementIsVisible(By.XPath("//*[@id='table-itemDetailsStorage']/tbody/tr[*]/td[contains(text(),'MAD')]/../td[4]"));
                //FIXME mauvais Place (Enconomico et non MAD4)
                stock = int.Parse(stockSite.Text.Replace(" ", ""));
            }

            // on vide le stock
            OutputFormPage outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            OutputFormCreateModalPage createOF = outputFormPage.OutputFormCreatePage();
            createOF.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo, true);
            OutputFormItem itemsOF = createOF.Submit();
            //faire le checks avant de sélectionner l'item + Cliquer sur validate afin d'avoir le second message d'erreur.
            OutputFormQualityChecks checksOF = itemsOF.ClickOnChecksTab();
            checksOF.DeliveryAccepted();
            itemsOF = checksOF.ClickOnItemsTab();
            itemsOF.Filter(OutputFormItem.FilterItemType.SearchByName, itemName);
            itemsOF.SelectFirstItem();
            itemsOF.AddPhysicalQuantity(itemName, stock.ToString());
            itemsOF.Validate();

            //1. Créer un inventaire sur le site, ajouter 150 phys.qtty sur un item(brandy mango sur ACE, supplier EMICELA, S.A.)
            InventoriesPage inventory = itemPage.GoToWarehouse_InventoriesPage();
            InventoryCreateModalPage inventoryCreate = inventory.InventoryCreatePage();
            inventoryCreate.FillField_CreateNewInventory(DateUtils.Now, site, placeFrom);
            InventoryItem inventoryItems = inventoryCreate.Submit();
            inventoryItems.Filter(InventoryItem.FilterItemType.SearchByName, itemName);
            //le packaging qty (fois packaging unit) s'additionne au physQty
            inventoryItems.SelectFirstItem();
            inventoryItems.AddPhysicalQuantity(itemName, "150");
            inventoryItems.AddPhysicalPackagingQuantity(itemName, "0");
            InventoryValidationModalPage validateModal = inventoryItems.Validate();
            validateModal.ValidatePartialInventory();

            //2. Créer une RN sur un site et un supplier avec 100 Received sur un item, valider les checks et valider la RN
            ReceiptNotesPage receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            ReceiptNotesCreateModalPage receiptNotesCreate = receiptNotesPage.ReceiptNotesCreatePage();
            receiptNotesCreate.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, placeFrom));
            ReceiptNotesItem receiptNotesItem = receiptNotesCreate.Submit();
            var receiptNoteNumber = receiptNotesItem.GetReceiptNoteNumber();
            receiptNotesItem.Filter(ReceiptNotesItem.FilterItemType.SearchByName, itemName);
            receiptNotesItem.SelectFirstItem();
            receiptNotesItem.AddReceived(itemName, "100");
            ReceiptNotesQualityChecks qualityChecks = receiptNotesItem.ClickOnChecksTab();
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

            //3. Créer un premier OF sur même site, physical quantity à 200 et valider
            outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            createOF = outputFormPage.OutputFormCreatePage();
            createOF.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo, true);
            itemsOF = createOF.Submit();
            //faire le checks avant de sélectionner l'item + Cliquer sur validate afin d'avoir le second message d'erreur.
            checksOF = itemsOF.ClickOnChecksTab();
            checksOF.DeliveryAccepted();
            itemsOF = checksOF.ClickOnItemsTab();
            itemsOF.Filter(OutputFormItem.FilterItemType.SearchByName, itemName);
            itemsOF.SelectFirstItem();
            itemsOF.AddPhysicalQuantity(itemName, 200.ToString());
            itemsOF.Validate();

            //4.Créer un deuxièmre OF, cette fois à partir de la RN(Generate Output Form) créée précédemment où physical quantity 100 contre theroical qty à 50, la valider
            receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();
            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ByNumber, receiptNoteNumber);
            receiptNotesItem = receiptNotesPage.SelectFirstReceiptNoteItem();
            ReceiptNoteToOuputForm outputFormModal = receiptNotesItem.GenerateOutputForm();
            outputFormModal.Fill(placeFrom, placeTo, true);
            OutputFormItem outputForm = outputFormModal.Create();

            //error message on ligne
            outputForm.Filter(OutputFormItem.FilterItemType.ShowItemsWithPhysQty, true);
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");
            double phQty = double.Parse(outputForm.GetPhysicalQuantity(), ci);
            double thQty = double.Parse(outputForm.GetTheoricalQuantity(), ci);
            Assert.AreNotEqual(0.0, phQty, 0.001);
            Assert.AreNotEqual(0.0, thQty, 0.001);
            Console.WriteLine("phQty : " + phQty + " - thQty : " + thQty);

            var theroricalQuantity = outputForm.GetTheoricalQuantity();
            Assert.IsTrue(phQty > thQty, theroricalQuantity);
            bool isFailed = outputForm.IsFailed();
            Assert.IsTrue(isFailed, "Pas de message d'erreur alors que phQty > thQty");

            try
            {
                outputForm.Validate();
                Assert.Fail("Can validate");
            }
            catch
            {
                Console.WriteLine("cannot validate");
                Assert.IsTrue(outputForm.ValidationFailed(), "Lors de la validation de l'OF, pas de message d'erreur alors que phQty > thQty");
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_CreateOutputFormSeveralPackagings()
        {
            // Prepare
            string itemName = "Item-" + DateUtils.Now.ToString("yyyy-MM-dd") + "-" + new Random().Next();
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxName = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();

            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();


            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            //1.Créer un item
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);

            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName, group, workshop, taxName, prodUnit);

                //2.Rajouter 3 Packaging
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging("ACE", "KG", storageQty, storageUnit, qty, supplier);

                itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging("ACE", "Caja", storageQty, storageUnit, qty, supplier);

                itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging("ACE", "Garrafa", storageQty, storageUnit, qty, supplier);

                itemPage = itemGeneralInformationPage.BackToList();
                itemPage.ResetFilter();
                itemPage.Filter(ItemPage.FilterType.Search, itemName);
                Assert.AreEqual(1, itemPage.CheckTotalNumber(), "item créé non retrouvé");
            }

            //3.Faire un inventaire sur les 3 Packaging
            InventoriesPage inventoryPage = itemPage.GoToWarehouse_InventoriesPage();
            InventoryCreateModalPage inventoryModal = inventoryPage.InventoryCreatePage();
            inventoryModal.FillField_CreateNewInventory(DateUtils.Now, "ACE", "Economato");
            InventoryItem inventory = inventoryModal.Submit();

            inventory.Filter(InventoryItem.FilterItemType.SearchByName, itemName);
            inventory.SelectFirstItem();
            inventory.SetPhysPackagingQty("10", "5", "3");
            InventoryValidationModalPage validateModal = inventory.Validate();
            validateModal.ValidateTotalInventory();

            //4.Faire un Output Form
            OutputFormPage ofPage = inventory.GoToWarehouse_OutputFormPage();
            OutputFormCreateModalPage ofModal = ofPage.OutputFormCreatePage();
            ofModal.FillField_CreatNewOutputForm(DateUtils.Now, "ACE", "Economato", "Produccion");
            OutputFormItem ofItems = ofModal.Submit();
            ofItems.ResetFilter();
            ofItems.Filter(OutputFormItem.FilterItemType.SearchByName, itemName);
            //5.Vérifier qu'il n'y est qu'une seule ligne sur l'Output Form
            ofItems.CheckTotalNumber();
            Assert.AreEqual(1, ofItems.CheckTotalNumber(), "Plusieurs lignes sur l'output form");
            ofItems.SelectFirstItem();
            ofItems.SetPhysPackagingQty("1", "2", "3");
            ofItems.SelectFirstItem();
            ofItems.AddPhysicalQuantity(itemName, "1");
            ofItems.Validate();
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_DeleteOutputForm()
        {
            string site = TestContext.Properties["Site"].ToString();
            string placeFrom = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();

            // Log in
            var homePage = LogInAsAdmin();

            //Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();

            var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
            outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo, true);
            String ID = outputFormCreateModalpage.GetOutputFormNumber();
            var outputFormItem = outputFormCreateModalpage.Submit();
            outputFormItem.BackToList();
            outputFormPage.Filter(OutputFormPage.FilterType.CheckAllSites, true);
            outputFormPage.Filter(OutputFormPage.FilterType.CheckAllOfTypes, true);
            var totalNumberAfterCreate = outputFormPage.CheckTotalNumber();
            outputFormPage.PageUp();
            outputFormPage.Filter(OutputFormPage.FilterType.SearchByNumber, ID);
            outputFormPage.DeleteFirstOutputForm();
            outputFormPage.ResetFilter();
            outputFormPage.Filter(OutputFormPage.FilterType.CheckAllSites, true);
            outputFormPage.Filter(OutputFormPage.FilterType.CheckAllOfTypes, true);
            var totalNumberAfterDelete = outputFormPage.CheckTotalNumber();
            Assert.AreEqual(totalNumberAfterDelete, totalNumberAfterCreate - 1, "La suppression ne fonctionne pas.");
            homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.Filter(OutputFormPage.FilterType.SearchByNumber, ID);
            Assert.AreEqual(outputFormPage.CheckTotalNumber(), 0, "La suppression ne fonctionne pas.");
            Assert.IsFalse(outputFormPage.isElementVisible(By.Id("page-size-selector")), "La suppression ne fonctionne pas.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_ChangePageSize()
        {
            string site = TestContext.Properties["Site"].ToString();
            string placeFrom = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();

            // Log in
            var homePage = LogInAsAdmin();

            //Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();

            var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
            outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo, true);
            var outputFormItem = outputFormCreateModalpage.Submit();

            outputFormItem.WaitPageLoading();
            outputFormItem.PageSize("16");
            outputFormItem.WaitPageLoading();
            var checkTotalNumberForFIlter16 = outputFormItem.CheckTotalNumber();
            Assert.AreEqual(checkTotalNumberForFIlter16, 16, MessageErreur.SIZE_PAGE_KO);
            outputFormItem.PageSize("30");
            outputFormItem.WaitPageLoading();
            var checkTotalNumberForFIlter30 = outputFormItem.CheckTotalNumber();
            Assert.AreEqual(checkTotalNumberForFIlter30, 30, MessageErreur.SIZE_PAGE_KO);
            outputFormItem.PageSize("8");
            outputFormItem.WaitPageLoading();
            var checkTotalNumberForFIlter8 = outputFormItem.CheckTotalNumber();
            Assert.AreEqual(checkTotalNumberForFIlter8, 8, MessageErreur.SIZE_PAGE_KO);

            Assert.IsTrue(outputFormItem.VerifyChangePage(4), MessageErreur.SIZE_PAGE_KO);

        }
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_CreateOutputFormWithPriceZero()
        {
            string site = TestContext.Properties["Site"].ToString();
            string placeFrom = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();

            // Log in
            var homePage = LogInAsAdmin();

            //Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();

            var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
            outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo, true, false);
            var outputFormItem = outputFormCreateModalpage.Submit();
            outputFormItem.SelectFirstItem();
            bool isPriceVisible = outputFormItem.isPriceVisible();
            if (outputFormItem.isPackagingQtyZero())
            {
                Assert.IsTrue(isPriceVisible, "Le prix de l'unité de mon Item n'est pas affiché.");
            }
            Assert.IsTrue(isPriceVisible, "Le prix de l'unité de mon Item n'est pas affiché.");

        }
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_UpdateTheoricalQuantity()
        {
            string site = TestContext.Properties["Site"].ToString();
            string placeFrom = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();

            // Log in
            var homePage = LogInAsAdmin();

            //Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();

            var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
            outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo, true, false);
            String outputFormID = outputFormCreateModalpage.GetOutputFormNumber();
            var outputFormItem = outputFormCreateModalpage.Submit();
            outputFormItem.SelectFirstItem();
            var itemName = outputFormItem.GetFirstItemName();
            var priceItemOrigine = outputFormItem.GetPrice(itemName);
            var theoQtyOrigine = outputFormItem.GetTheoricalQuantity();

            ItemGeneralInformationPage itemGeneralInformationPage = outputFormItem.EditItem(itemName);
            var supplierItem = itemGeneralInformationPage.GetFirstPackagingSupplier();

            var recipeNotePage = homePage.GoToWarehouse_ReceiptNotesPage();
            ReceiptNotesCreateModalPage receiptNotesCreateModalPage = recipeNotePage.ReceiptNotesCreatePage();
            receiptNotesCreateModalPage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateTime.Today, site, supplierItem, placeTo));
            ReceiptNotesItem receiptNotesItem = receiptNotesCreateModalPage.Submit();
            receiptNotesItem.Filter(ReceiptNotesItem.FilterItemType.SearchByName, itemName);
            receiptNotesItem.SetFirstReceivedQuantity("12");

            InventoriesPage inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            InventoryCreateModalPage inventoryCreateModalPage = inventoriesPage.InventoryCreatePage();
            inventoryCreateModalPage.FillField_CreateNewInventory(DateTime.Today, site, placeFrom);
            InventoryItem inventoryItem = inventoryCreateModalPage.Submit();
            inventoryItem.Filter(InventoryItem.FilterItemType.SearchByName, itemName);
            inventoryItem.SelectFirstItem();
            inventoryItem.AddPhysicalQuantity(itemName, "12");
            inventoryItem.SetPrice(itemName, "5");

            outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.Filter(OutputFormPage.FilterType.SearchByNumber, outputFormID);
            outputFormItem = outputFormPage.SelectFirstOutputForm();
            outputFormItem.SelectFirstItem();
            Assert.AreEqual(priceItemOrigine, outputFormItem.GetPrice(itemName), "Le price s'actualise.");
            Assert.AreEqual(theoQtyOrigine, outputFormItem.GetTheoricalQuantity(), "La quantité théorique s'actualise");

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_PrintIndexBackdate()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string placeFrom = TestContext.Properties["OutputFormPlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();

            // Log in
            var homePage = LogInAsAdmin();

            //Act
            var outputFormPage = homePage.GoToWarehouse_OutputFormPage();
            outputFormPage.ResetFilter();
            DeleteAllFileDownload();
            outputFormPage.ClearDownloads();
            var outputFormCreateModalpage = outputFormPage.OutputFormCreatePage();
            outputFormCreateModalpage.FillField_CreatNewOutputForm(DateUtils.Now, site, placeFrom, placeTo, true, false);
            String outputFormID = outputFormCreateModalpage.GetOutputFormNumber();

            var outputFormItem = outputFormCreateModalpage.Submit();
            outputFormItem.SelectFirstItem();

            var itemName = outputFormItem.GetFirstItemName();
            var quantity = outputFormItem.GetTheoricalQuantity();

            outputFormItem.AddPhysicalQuantity(itemName, quantity);

            outputFormItem.Validate();

            outputFormPage = outputFormItem.BackToList();
            outputFormPage.ResetFilter();

            outputFormPage.Filter(OutputFormPage.FilterType.SearchByNumber, outputFormID);
            outputFormPage.Filter(OutputFormPage.FilterType.DateFrom, DateUtils.Now);
            outputFormPage.Filter(OutputFormPage.FilterType.DateTo, DateUtils.Now.AddMonths(1));
            PrintReportPage printReportPage = outputFormPage.PrintResults(true);

            var isReportGenerated = printReportPage.IsReportGenerated();
            printReportPage.Close();

            Assert.IsTrue(isReportGenerated, "L'impression ne doit pas etre possible");
        }
        // _________________________________________ Utilitaire _______________________________________________

        public void CreateInventory(HomePage homePage, string site, string placeFrom, string itemName)
        {
            var inventoriesPage = homePage.GoToWarehouse_InventoriesPage();
            inventoriesPage.ResetFilters();

            // Create a new Inventory
            var inventoryCreateModalpage = inventoriesPage.InventoryCreatePage();
            inventoryCreateModalpage.FillField_CreateNewInventory(DateUtils.Now, site, placeFrom, true);
            var inventoryDetailPage = inventoryCreateModalpage.Submit();

            inventoryDetailPage.Filter(InventoryItem.FilterItemType.SearchByName, itemName);
            inventoryDetailPage.SelectFirstItem();
            inventoryDetailPage.AddPhysicalQuantity(itemName, "10");

            //validate 
            var validateInventory = inventoryDetailPage.Validate();
            validateInventory.ValidatePartialInventory();
        }

        private string GetItemGroup(HomePage homePage, string itemName)
        {
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);

            var itemGeneralInfo = itemPage.ClickOnFirstItem();

            return itemGeneralInfo.GetGroupName();
        }

        private string GetVat(HomePage homePage, string itemName)
        {
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);

            var itemGeneralInfo = itemPage.ClickOnFirstItem();

            return itemGeneralInfo.GetVatName();
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
                settingsSitesPage.WaitLoading();

                settingsSitesPage.ClickToInformations();
                settingsSitesPage.SetAnalyticPlan(analyticalPlan);
                settingsSitesPage.SetAnalyticSection(analyticalSection);
                settingsSitesPage.WaitLoading();
            }
            catch
            {
                return false;
            }

            return true;
        }


        private bool VerifySiteConfigPlace(HomePage homePage, string site, string placeTo, string siteTo, bool isOK = true)
        {

            var placeType = isOK ? ParametersSites.PlaceType.OtherSite : ParametersSites.PlaceType.Production;
            try
            {
                // Act
                var settingsSitesPage = homePage.GoToParameters_Sites();
                settingsSitesPage.Filter(ParametersSites.FilterType.SearchSite, site);
                settingsSitesPage.ClickOnFirstSite();

                settingsSitesPage.ClickToOrganization();

                if (!settingsSitesPage.IsOrganizationPresent(placeTo))
                {
                    settingsSitesPage.CreateNewOrganization(placeTo);
                }

                settingsSitesPage.SearchOrganization(placeTo);
                settingsSitesPage.EditOrganization(placeType, siteTo, placeTo);
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

                accountingParametersPage.SearchGroup(group, vat);
                accountingParametersPage.EditInventoryAccounts(account, exoAccount, inventoryAccount, inventoryVariationAccount);
            }
            catch
            {
                return false;
            }

            return true;
        }

        private bool VerifyAccountingJournal(HomePage homePage, string site, string journalOutputForm)
        {
            try
            {
                var accountingJournalPage = homePage.GoToParameters_AccountingPage();
                accountingJournalPage.GoToTab_Journal();
                accountingJournalPage.EditJournal(site, null, null, null, null, journalOutputForm);
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

        public void CreateInventory(HomePage homePage, string site, string place)
        {
            var invPage = homePage.GoToWarehouse_InventoriesPage();
            var invModal = invPage.InventoryCreatePage();
            invModal.FillField_CreateNewInventory(DateUtils.Now, site, place);
            InventoryItem invItems = invModal.Submit();
            var invItemName = invItems.GetFirstItemName();
            invItems.SelectFirstItem();
            invItems.AddPhysicalQuantity(invItemName, "10");
            InventoryValidationModalPage modalValidate = invItems.Validate();
            modalValidate.ValidatePartialInventory();

        }

        /// <summary>
        /// Update output form's first item allergens and check the state 
        /// of the allergens button in the output form details.
        /// </summary>
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_OF_CheckAllergensUpdate()
        {
            string allergen1 = "Cacahuetes/Peanuts";
            string allergen2 = "Frutos de cáscara- Macadamias/Nuts-Macadamia";
            HomePage homePage = LogInAsAdmin();
            OutputFormPage ofPage = homePage.GoToWarehouse_OutputFormPage();
            ofPage.ResetFilter();
            ofPage.Filter(OutputFormPage.FilterType.ShowNotValidated, true);
            string outputFormNumber = ofPage.GetFirstOutputFormNumber();
            ofPage.Filter(OutputFormPage.FilterType.SearchByNumber, outputFormNumber);
            OutputFormItem of = ofPage.SelectFirstOutputForm();
            string itemName = of.GetFirstItemName();
            of.ResetFilter();
            of.Filter(OutputFormItem.FilterItemType.SearchByName, itemName);
            of.SelectFirstItem();
            of.AddPhysicalQuantity(itemName, "1");
            of.Refresh();
            of.ResetFilter();
            of.Filter(OutputFormItem.FilterItemType.SearchByName, itemName);
            of.SelectFirstItem();
            ItemGeneralInformationPage itemPage = of.EditItem(itemName);
            ItemIntolerancePage itemIntolerancePage = itemPage.ClickOnIntolerancePage();
            itemIntolerancePage.AddAllergen(allergen1);
            itemIntolerancePage.AddAllergen(allergen2);
            ofPage = homePage.GoToWarehouse_OutputFormPage();
            ofPage.ResetFilter();
            ofPage.Filter(OutputFormPage.FilterType.ShowNotValidated, true);
            ofPage.Filter(OutputFormPage.FilterType.SearchByNumber, outputFormNumber);
            of = ofPage.SelectFirstOutputForm();
            of.ResetFilter();
            of.Filter(OutputFormItem.FilterItemType.SearchByName, itemName);
            bool isIconGreen = of.IsAllergenIconGreen(itemName);
            List<string> allergensInItem = of.GetAllergens(itemName);
            bool isEmpty = allergensInItem.Count() == 0;
            Assert.IsFalse(isEmpty, $"No allergens for the selected item {itemName}");
            bool containsAllergen1 = allergensInItem.Contains(allergen1);
            bool containsAllergen2 = allergensInItem.Contains(allergen2);
            Assert.IsTrue(isIconGreen, "L'icon n'est pas vert!");
            Assert.IsTrue(containsAllergen1 && containsAllergen2, "Allergens n'ont pas été ajoutés");
        }

        /// ////////////////////////////////////////////////////////////////Methods//////////////////////////////////////////////
        private int InsertOutputForm(DateTime date, string site)
        {
            string query = @"
    DECLARE @nextNumber INT;
    DECLARE @nextNumberString NVARCHAR(10);

    -- Calculer le prochain Number
    SELECT @nextNumber = ISNULL(MAX(Number), 0) + 1 FROM OutputForms;

    -- Convertir le Number en NumberString avec un format de 6 chiffres
    SET @nextNumberString = FORMAT(@nextNumber, '000000');

    DECLARE @siteId INT;
    SELECT TOP 1 @siteId = Id FROM sites WHERE Name LIKE @site;

    DECLARE @placeFromId INT;
    SELECT TOP 1 @placeFromId = Id FROM SitePlaces WHERE SiteId = @siteId;

    DECLARE @placeToId INT;
    SELECT TOP 1 @placeToId = Id FROM SitePlaces WHERE SiteId = @siteId;

    INSERT INTO OutputForms 
    (
        Number, 
        NumberString, 
        Date, 
        CreationDate, 
        IsManual, 
        IsActive, 
        IsValid, 
        FromSitePlaceId, 
        ToSitePlaceId, 
        IsSentToSage, 
        SageState, 
        ExportedSageManuallyState, 
        Type, 
        Origin, 
        IsSentToWMS
    )
    VALUES 
    (
        @nextNumber,
        @nextNumberString,
        @date, 
        GETDATE(), -- Insertion de la date de création (date courante)
        1,
        1,
        0,
        @placeFromId,
        @placeToId,
        0,
        0,
        0,
        0,
        0,
        0
    );

    -- Retourner l'ID de l'enregistrement nouvellement inséré
    SELECT SCOPE_IDENTITY();
";

            return ExecuteAndGetInt(
                query,
                new KeyValuePair<string, object>("date", date),
                new KeyValuePair<string, object>("site", site)
            );
        }
        private void DeleteOutputForm(int OutputFormId)
        {
            string query = @"
            DELETE FROM OutputForms 
                    WHERE Id = @OutputFormId;";

            ExecuteNonQuery(query, new KeyValuePair<string, object>("OutputFormId", OutputFormId));
        }

        private string GetOutputFormNumber(int OutputFormId)
        {
            string query = @"
            SELECT NumberString FROM OutputForms 
            WHERE Id LIKE @OutputFormId";

            return ExecuteAndGetString(query, new KeyValuePair<string, object>("OutputFormId", OutputFormId));
        }

        private void VerifyOutputForm(OutputFormPage outputFormPage, string number, bool shouldExist, string errorMessage)
        {
            //outputFormPage.Filter(OutputFormPage.FilterType.SearchByNumber, number);
            //if (shouldExist)
            //{
            //    Assert.IsTrue(outputFormPage.VerifyDRNamesExist(deliveryRoundName), errorMessage);
            //}
            //else
            //{
            //    Assert.IsFalse(outputFormPage.VerifyDRNamesExist(deliveryRoundName), errorMessage);
            //}
        }
    }
}
