using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Accounting.SupplierInvoices;
using Newrest.Winrest.FunctionalTests.PageObjects.Admin;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Delivery;
using Newrest.Winrest.FunctionalTests.PageObjects.Jobs.ScheduledJobs;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Recipes;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Production;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Sites;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.Claims;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.OutputForm;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.ReceiptNotes;
using Newrest.Winrest.FunctionalTests.Utils;
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
using static Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.PurchaseOrderItem;


namespace Newrest.Winrest.FunctionalTests.Warehouse
{
    [TestClass]
    public class ReceiptNotesTest : TestBase
    {

        private const int _timeout = 600000;
        private readonly string RECEIPT_NOTES_EXCEL_SHEET = "Receipt notes";

        /// <summary>
        /// 
        /// Mise en place du paramétrage pour la configuration Winrest 4.0 
        /// 
        /// </summary>
		[TestMethod]
        [Priority(0)]
        [Timeout(_timeout)]
        public void WA_RN_SetConfigWinrest4_0()
        {
            // Prepare
            var site = TestContext.Properties["Site"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();
            ClearCache();

            // New version search
            homePage.SetNewVersionSearchValue(true);

            // New version keyword
            homePage.SetNewVersionKeywordValue(true);

            // New VersionItemDetails
            //homePage.SetVersionItemDetailValue(true);

            // New group display
            homePage.SetNewGroupDisplayValue(true);

            // Vérifier que c'est activé
            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();
            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ShowNotValidated, true);
            var receiptNotesItem = receiptNotesPage.SelectFirstReceiptNoteItem();

            // Vérifier New version search
            try
            {
                var itemName = receiptNotesItem.GetFirstItemName();
                receiptNotesItem.Filter(ReceiptNotesItem.FilterItemType.SearchByName, itemName);
            }
            catch
            {
                throw new Exception("La recherche a pu être effectuée, le NewSearchMode est actif.");
            }

            // vérifier new keyword search
            try
            {
                receiptNotesItem.ResetFilters();
                receiptNotesItem.Filter(ReceiptNotesItem.FilterItemType.SearchByKeyword, "TEST_KEY");

                //Show items n'est pas activé par default
                receiptNotesItem.Filter(ReceiptNotesItem.FilterItemType.ShowNotReceived, true);
            }
            catch
            {
                throw new Exception("La recherche par keyword n'a pas pu être effectuée, le NewKeywordMode est inactif.");
            }


            // vérifier new group display
            receiptNotesItem.ResetFilters();
            Assert.IsTrue(receiptNotesItem.IsGroupDisplayActive(), "Le paramètre 'NewGroupDisplay' n'est pas activé.");

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

            // Vérifier New version supplier invoice
            var supplierInvoicePage = homePage.GoToAccounting_SupplierInvoices();
            Assert.IsFalse(supplierInvoicePage.IsInvoiceAmountWithoutTaxPresent(), "Le paramètre 'NewSupplierInvoiceVersion' n'est pas activé.");
        }

        [Priority(1)]
		[TestMethod]
        [Timeout(_timeout)]
        public void WA_RN_CreateItemForReceiptNotes()
        {
            // Prepare items
            string receiptNoteItem = TestContext.Properties["Item_ReceiptNote"].ToString();

            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();

            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.Filter(ItemPage.FilterType.Search, receiptNoteItem);

            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(receiptNoteItem, group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
                itemPage = itemGeneralInformationPage.BackToList();

                itemPage.Filter(ItemPage.FilterType.Search, receiptNoteItem.ToString());
            }
            Assert.AreEqual(receiptNoteItem, itemPage.GetFirstItemName(), $"L'item {receiptNoteItem} n'est pas présent dans la liste des items disponibles.");
        }

        [Priority(2)]
		[TestMethod]
        [Timeout(_timeout)]
        public void WA_RN_PrepareExportSageConfig()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string itemName = TestContext.Properties["Item_ReceiptNote"].ToString();
            string journalRN = TestContext.Properties["Journal_RN"].ToString();

            string taxName = TestContext.Properties["Item_TaxType"].ToString();
            string taxType = "VAT";

            //Arrange
            var homePage = LogInAsAdmin();

            // Récupération du groupe de l'item
            string itemGroup = GetItemGroup(homePage, itemName);

            // Vérification du paramétrage
            // --> Admin Settings
            bool isAppSettingsOK = SetApplicationSettingsForSageAuto(homePage);

            // Sites -- > Analytical plan et section
            bool isAnalyticalPlanOK = VerifySiteAnalyticalPlanSection(homePage, site);

            // Parameter - Purchasing --> VAT
            bool isPurchasingVATOK = VerifyPurchasingVAT(homePage, taxName, taxType);

            // Parameter - Accounting --> Service categories & VAT
            bool isGroupAndVatOK = VerifyGroupAndVAT(homePage, itemGroup, taxName);

            // Parameter - Accounting --> Journal
            bool isJournalOk = VerifyAccountingJournal(homePage, site, journalRN);

            // IntegrationDate
            var date = VerifyIntegrationDate(homePage);

            // Supplier
            bool isSupplierOK = VerifySupplier(homePage, site, supplier);

            // Assert
            Assert.AreNotEqual("", itemGroup, "Le groupe de l'item n'a pas été récupéré.");
            Assert.IsTrue(isAppSettingsOK, "Les application settings pour TL ne sont pas configurés correctement.");
            Assert.IsTrue(isAnalyticalPlanOK, "La configuration des analytical plan du site n'est pas effectuée.");
            Assert.IsTrue(isPurchasingVATOK, "La configuration des purchasing VAT n'est pas effectuée.");
            Assert.IsTrue(isGroupAndVatOK, "La configuration du group and VAT de l'item n'est pas effectuée.");
            Assert.IsTrue(isJournalOk, "La catégorie du accounting journal n'a pas été effectuée.");
            Assert.IsNotNull(date, "La date d'intégration est nulle.");
            Assert.IsTrue(isSupplierOK, "La configuration du supplier n'a pas été effectuée.");
        }
   
		[TestMethod]
        [Timeout(_timeout)]
        public void WA_RN_CancelValidatedReceiptNote()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();
            HomePage homePage = LogInAsAdmin();
            ReceiptNotesPage receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            ReceiptNotesCreateModalPage receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
            receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, deliveryPlace));
            ReceiptNotesItem receiptNotesItem = receiptNotesCreateModalpage.Submit();
            string itemName = receiptNotesItem.GetFirstItemName();
            receiptNotesItem.ResetFilters();
            receiptNotesItem.Filter(ReceiptNotesItem.FilterItemType.SearchByName, itemName);
            receiptNotesItem.SelectFirstItem();
            receiptNotesItem.AddReceived(itemName, "2");
            receiptNotesItem.Refresh();
            ReceiptNotesQualityChecks qualityChecks = receiptNotesItem.ClickOnChecksTab();
            qualityChecks.SetQualityChecks();
            qualityChecks.DeliveryAccepted();
            receiptNotesItem.Validate();
            ReceiptNotesItem cancellationRN = receiptNotesItem.CancelReceiptNote();
            cancellationRN.Close();
            receiptNotesItem.ClickOnGeneralInformationTab();
            bool isLinkExists = receiptNotesItem.ExistsLinkToCancelledRN();
            Assert.IsTrue(isLinkExists, "The link to the cancellation RN doesn't exist.");
        }

        /*
         * Création d'une Receipt Note classique sans PO
         */
        [Priority(3)]
		[TestMethod]
        [Timeout(_timeout)]
        public void WA_RN_CreateNewReceiptNote()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();
            string itemName = TestContext.Properties["Item_ReceiptNote"].ToString();

            HomePage homePage = LogInAsAdmin();
            ReceiptNotesPage receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();
            ReceiptNotesCreateModalPage receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
            receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, deliveryPlace));
            ReceiptNotesItem receiptNotesItem = receiptNotesCreateModalpage.Submit();
            string receiptNoteNumber = receiptNotesItem.GetReceiptNoteNumber();
            receiptNotesItem.ResetFilters();
            receiptNotesItem.Filter(ReceiptNotesItem.FilterItemType.SearchByName, itemName);
            receiptNotesItem.SelectFirstItem();
            receiptNotesItem.AddReceived(itemName, "2");
            receiptNotesItem.Refresh();
            receiptNotesPage = receiptNotesItem.BackToList();
            receiptNotesPage.ResetFilter();
            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ByNumber, receiptNoteNumber);
            string firstReceiptNoteNumber = receiptNotesPage.GetFirstReceiptNoteNumber();
            Assert.AreEqual(firstReceiptNoteNumber, receiptNoteNumber, string.Format(MessageErreur.OBJET_NON_CREE, "La receipt note"));
        }

        [Priority(4)]
		[TestMethod]
        [Timeout(_timeout)]
        public void WA_RN_Items_Filters_SearchByKeyword()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();
            string itemKeyword = TestContext.Properties["Item_Keyword"].ToString();

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            homePage.SetNewVersionKeywordValue(true);

            //Act            
            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();

            // Create
            var receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
            receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, placeTo));

            var receiptNotesItem = receiptNotesCreateModalpage.Submit();

            var itemName = receiptNotesItem.GetFirstItemName();


            // Filter by keyword
            receiptNotesItem.ResetFilters();
            receiptNotesItem.Filter(ReceiptNotesItem.FilterItemType.SearchByKeyword, itemKeyword);
            receiptNotesItem.PageSize("8");

            var isKeywordOK = false;
            var errorMessageKeyword = "";
            try
            {
                isKeywordOK = receiptNotesItem.VerifyKeyword(itemKeyword);
            }
            catch (Exception e)
            {
                errorMessageKeyword = e.Message;
            }

            Assert.IsTrue(isKeywordOK, errorMessageKeyword);
        }

        [Priority(5)]
		[TestMethod]
        [Timeout(_timeout)]
        public void WA_RN_CheckExistenceRNInvoiced()
        {
            string invoicedRN_DN = "InvoicedRN";
            DateTime dateFrom = DateTime.ParseExact("01/01/2024", "dd/MM/yyyy", CultureInfo.InvariantCulture);

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var rnPage = homePage.GoToWarehouse_ReceiptNotesPage();
            rnPage.ResetFilter();
            rnPage.Filter(ReceiptNotesPage.FilterType.ByDeliveryNumber, invoicedRN_DN);
            rnPage.Filter(ReceiptNotesPage.FilterType.DateFrom, dateFrom);

            if (rnPage.AreAllRnInvoiced())
            {
                Assert.IsTrue(true, "");
                return;
            }

            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();
            string itemName = TestContext.Properties["Item_ReceiptNote"].ToString();

            // Create
            var receiptNotesCreateModalpage = rnPage.ReceiptNotesCreatePage();
            receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, deliveryPlace) { DeliveryNumber = invoicedRN_DN });
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
            var siDetailPage = receiptNotesItem.GenerateSupplierInvoice("SOFromInvoicedRN");
            siDetailPage.ValidateSupplierInvoice();

            Assert.IsTrue(true, "");
        }

		[TestMethod]
        [Timeout(_timeout)]
        public void WA_RN_Index_MultipleFilterChecks()
        {
            List<string> listDetectedErrors = new List<string>();

            // Log in
            var homePage = LogInAsAdmin();
            var dateFormat = homePage.GetDateFormatPickerValue();

            //Go to RN index   
            ReceiptNotesPage receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.WaitPageLoading();

            //Check pagination
            this.WA_RN_index_pagination(receiptNotesPage, listDetectedErrors);

            //Check filtering by RN number - ends test early if no RN found
            this.WA_RN_Filter_SearchByNumberFilter(receiptNotesPage, listDetectedErrors);

            //For all the other test, we want to check on the maximum page size
            if (receiptNotesPage.isPageSizeEqualsTo100() == false)
            {
                receiptNotesPage.PageSize("100");
            }

            //Check filtering by supplier
            this.WA_RN_Filter_Suppliers(receiptNotesPage, listDetectedErrors);

            //Check date filtering
            this.WA_RN_Filter_Date(receiptNotesPage, dateFormat, listDetectedErrors);

            //Check status validated, not validated and all RN
            this.WA_RN_Filter_CheckAll_Validated_NotValidated(receiptNotesPage, listDetectedErrors);

            //Check status active, not active, all
            this.WA_RN_Filter_CheckAll_Active_NotActive(receiptNotesPage, listDetectedErrors);

            //Check sort by
            this.WA_RN_Filter_SortBy(receiptNotesPage, listDetectedErrors, dateFormat);

            //Check site filtering
            this.WA_RN_Filter_Site(receiptNotesPage, listDetectedErrors);

            if (listDetectedErrors.Any())
            {
                Assert.IsTrue(false, string.Join(Environment.NewLine, listDetectedErrors));
            }
        }

        /*
         * Création d'une nouvelle Receipt Note depuis un Purchase Order
        */
		[TestMethod]
        [Timeout(_timeout)]
        public void WA_RN_CreateReceiptNotesFromPurchase()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();
            string itemName = TestContext.Properties["Item_ReceiptNote"].ToString();

            DateTime datePurchaseOrder = DateUtils.Now;

            // Log in
            var homePage = LogInAsAdmin();

            var decimalSeparator = homePage.GetDecimalSeparatorValue();
            // Create purchase order
            string purchaseOrderNumber = CreatePurchaseOrder(homePage, site, supplier, datePurchaseOrder, itemName);
            PurchaseOrderItem poItem = new PurchaseOrderItem(WebDriver, TestContext);
            double poPrice = poItem.GetTotalSum();
            //Act           
            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();

            var receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
            receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, deliveryPlace) { CreateFromPO = true, PONumber = purchaseOrderNumber, PODate = datePurchaseOrder });
            var receiptNotesItem = receiptNotesCreateModalpage.Submit();

            string receiptNoteNumber = receiptNotesItem.GetReceiptNoteNumber();

            Assert.AreEqual(poPrice, receiptNotesItem.GetDNTotalSum(decimalSeparator), "Mauvais tarif");

            Assert.AreEqual(itemName, receiptNotesItem.GetFirstItemNameText(), "La receipt note créée ne possède pas les mêmes items que la purchase order source.");

            receiptNotesPage = receiptNotesItem.BackToList();
            receiptNotesPage.ResetFilter();

            //Assert
            Assert.AreEqual(receiptNotesPage.GetFirstReceiptNoteNumber(), receiptNoteNumber, String.Format(MessageErreur.OBJET_NON_CREE, "La receipt note"));
        }

        /*
         * Création d'une nouvelle Receipt Note depuis un Purchase Order partiellement reçu
        */
		[TestMethod]
        [Timeout(_timeout)]
        public void WA_RN_CreateReceiptNotesFromPurchasePartielle()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string itemName = TestContext.Properties["Item_ReceiptNote"].ToString();

            // Log in
            var homePage = LogInAsAdmin();

            // Création Purchase Order
            string idReceiptNote = CreatePurchaseOrderPartielle(homePage, site, supplier, "PartiallyDelivered", itemName);

            //Act          
            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();
            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ByNumber, idReceiptNote);

            //Assert
            Assert.AreEqual(receiptNotesPage.GetFirstReceiptNoteNumber(), idReceiptNote, String.Format(MessageErreur.OBJET_NON_CREE, "La receipt note"));
        }

        /// <summary>
        /// Test sur le filtrage par RN number. Il doit y avoir des résultats RN déjà présents.
        /// </summary>
        /// <param name="receiptNotesPage">La page d'index RN</param>
        /// <param name="listErrors">La liste des erreurs détectées des précédents tests.</param>
        private void WA_RN_Filter_SearchByNumberFilter(ReceiptNotesPage receiptNotesPage, List<string> listErrors)
        {
            try
            {
                receiptNotesPage.ResetFilter();
                string rnNumber = receiptNotesPage.GetFirstReceiptNoteNumber();

                if (string.IsNullOrEmpty(rnNumber))
                {
                    listErrors.Add("Il n'existe pas de RN sur lesquels faire les tests.");
                    Assert.IsTrue(false, string.Join(Environment.NewLine, listErrors));
                }

                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ByNumber, rnNumber);

                int nbShowedResults = receiptNotesPage.GetNumberOfShowedResults();
                if (nbShowedResults > 1)
                {
                    listErrors.Add("Il y a plus de 1 résultat pour un filtre sur un 'RN number'.");
                }

                string rnNumberAfterFilter = receiptNotesPage.GetFirstReceiptNoteNumber();
                if (rnNumber.Equals(rnNumberAfterFilter) == false)
                {
                    listErrors.Add(String.Format(MessageErreur.FILTRE_ERRONE, "RN number"));
                }
            }
            catch (Exception ex)
            {
                listErrors.Add("Erreur dans WA_RN_Filter_SearchByNumberFilter : " + ex.Message);
            }
        }

        /*
         * Test filtrage : Par purchase number
         */
        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_Filter_SearchByPurchaseNumberFilter()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string placeTo = TestContext.Properties["PlaceFrom"].ToString();
            string itemName = TestContext.Properties["Item_ReceiptNote"].ToString();
            string downloadPath = TestContext.Properties["DownloadsPath"].ToString();
            bool newVersionPrint = true;
            HomePage homePage = LogInAsAdmin();
            string purchaseOrderNumber = CreatePurchaseOrder(homePage, site, supplier, DateUtils.Now.AddDays(-1), itemName);
            ReceiptNotesPage receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();
            receiptNotesPage.ClearDownloads();
            ReceiptNotesCreateModalPage receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
            receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, placeTo) { CreateFromPO = true, PONumber = purchaseOrderNumber, PODate = DateUtils.Now.AddDays(-1) });
            ReceiptNotesItem receiptNotesItem = receiptNotesCreateModalpage.Submit();
            receiptNotesItem.SelectFirstItem();
            receiptNotesItem.AddReicedQuantity(itemName, "2");
            receiptNotesItem.Refresh();
            receiptNotesPage = receiptNotesItem.BackToList();
            receiptNotesPage.ResetFilter();
            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ByPurchaseNumber, purchaseOrderNumber);
            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.DateFrom, DateUtils.Now);
            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.DateTo, DateUtils.Now.AddDays(1));
            receiptNotesPage.ExportExcelFile(newVersionPrint);
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();
            FileInfo correctDownloadedFile = receiptNotesPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");
            string fileName = correctDownloadedFile.Name;
            string filePath = Path.Combine(downloadPath, fileName);
            int resultNumber = OpenXmlExcel.GetExportResultNumber(RECEIPT_NOTES_EXCEL_SHEET, filePath);
            bool result = OpenXmlExcel.ReadAllDataInColumn("Purchase Order", RECEIPT_NOTES_EXCEL_SHEET, filePath, purchaseOrderNumber);
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.IsTrue(result, MessageErreur.EXCEL_DONNEES_KO);
        }

        /*
         * Test filtrage : Par delivery number
         */
        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_Filter_SearchByDeliveryNumberFilter()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string deliveryPlace = TestContext.Properties["PlaceTo"].ToString();
            string downloadPath = TestContext.Properties["DownloadsPath"].ToString();
            bool newVersionPrint = true;

            // Log in
            var homePage = LogInAsAdmin();
            //Act
            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            try
            {
                receiptNotesPage.ResetFilter();
                DeleteAllFileDownload();
                receiptNotesPage.ClearDownloads();

                // Create
                var receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
                receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, deliveryPlace));
                var deliveryNumber = receiptNotesCreateModalpage.GetDeliveryOrderNumber();
                var deliveryDate = receiptNotesCreateModalpage.GetDeliveryDate();
                var receiptNotesItem = receiptNotesCreateModalpage.Submit();

                // Get ReceiptNote Number and Item Name
                var receiptNoteNumber = receiptNotesItem.GetReceiptNoteNumber();
                var itemName = receiptNotesItem.GetFirstItemName();
                // Add Received
                receiptNotesItem.SelectFirstItem();
                receiptNotesItem.AddReceived(itemName, "2");
                receiptNotesItem.Refresh();
                receiptNotesPage = receiptNotesItem.BackToList();
                receiptNotesPage.ResetFilter();
                // Filter Receipt Note Created 
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ByDeliveryNumber, deliveryNumber);
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.DateFrom, DateUtils.Now);
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.DateTo, DateUtils.Now.AddDays(1));

                // Get data from Receipt note Created 
                var deliveryLocation = receiptNotesPage.GetDeliveryLocation();
                var totalVat = receiptNotesPage.GetTotalVat();

                // On exporte les résultats sous la forme d'un fichier Excel (dont on récupère le nom)
                receiptNotesPage.ClearDownloads();
                receiptNotesPage.ExportExcelFile(newVersionPrint);

                // On récupère les fichiers du répertoire de téléchargement
                DirectoryInfo taskDirectory = new DirectoryInfo(downloadPath);
                FileInfo[] taskFiles = taskDirectory.GetFiles();

                // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
                var correctDownloadedFile = receiptNotesPage.GetExportExcelFile(taskFiles);
                Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

                // Récupération du nom du fichier et construction de l'URL du fichier Excel à ouvrir   
                var fileName = correctDownloadedFile.Name;
                var filePath = Path.Combine(downloadPath, fileName);

                // Exploitation du fichier Excel
                int resultNumber = OpenXmlExcel.GetExportResultNumber(RECEIPT_NOTES_EXCEL_SHEET, filePath);

                bool verifyDeliveryNumber = OpenXmlExcel.ReadAllDataInColumn("Delivery N°", RECEIPT_NOTES_EXCEL_SHEET, filePath, deliveryNumber);
                bool verifyReceiptNoteNumber = OpenXmlExcel.ReadAllDataInColumn("Number", RECEIPT_NOTES_EXCEL_SHEET, filePath, receiptNoteNumber);
                bool verifySite = OpenXmlExcel.ReadAllDataInColumn("Site", RECEIPT_NOTES_EXCEL_SHEET, filePath, site);
                bool verifySupplier = OpenXmlExcel.ReadAllDataInColumn("Supplier", RECEIPT_NOTES_EXCEL_SHEET, filePath, supplier);
                bool verifyDeliveryLocation = OpenXmlExcel.ReadAllDataInColumn("Site Place", RECEIPT_NOTES_EXCEL_SHEET, filePath, deliveryLocation);
                bool verifyDeliveryDate = OpenXmlExcel.ReadAllDataInColumn("Delivery Date", RECEIPT_NOTES_EXCEL_SHEET, filePath, deliveryDate);
                bool verifyTotalVat = OpenXmlExcel.ReadAllDataInColumn("Total No Vat", RECEIPT_NOTES_EXCEL_SHEET, filePath, totalVat);

                //Assert
                Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);
                Assert.IsTrue(verifyDeliveryNumber, MessageErreur.EXCEL_DONNEES_KO);
                Assert.IsTrue(verifyReceiptNoteNumber, MessageErreur.EXCEL_DONNEES_KO);
                Assert.IsTrue(verifySite, MessageErreur.EXCEL_DONNEES_KO);
                Assert.IsTrue(verifySupplier, MessageErreur.EXCEL_DONNEES_KO);
                Assert.IsTrue(verifyDeliveryLocation, MessageErreur.EXCEL_DONNEES_KO);
                Assert.IsTrue(verifyDeliveryDate, MessageErreur.EXCEL_DONNEES_KO);
                Assert.IsTrue(verifyTotalVat, MessageErreur.EXCEL_DONNEES_KO);
            }
            finally
            {
                // Delete Receipt Note Created
                receiptNotesPage.DeleteReceiptNote();
                receiptNotesPage.ResetFilter();
            }
        }

        /// <summary>
        /// Test sur le filtrage par Suppliers.
        /// </summary>
        /// <param name="receiptNotesPage">La page d'index RN</param>
        /// <param name="listErrors">La liste des erreurs détectées des précédents tests.</param>
        private void WA_RN_Filter_Suppliers(ReceiptNotesPage receiptNotesPage, List<string> listErrors)
        {
            try
            {
                string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();

                receiptNotesPage.ResetFilter();
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.Supplier, supplier);

                if (receiptNotesPage.VerifySupplier(supplier) == false)
                {
                    listErrors.Add(String.Format(MessageErreur.FILTRE_ERRONE, "'Supplier'"));
                }
            }
            catch (Exception ex)
            {
                listErrors.Add("Erreur dans WA_RN_Filter_Suppliers : " + ex.Message);
            }
        }

        /// <summary>
        /// Test sur le filtrage par date
        /// </summary>
        /// <param name="receiptNotesPage">La page d'index RN</param>
        /// <param name="dateFormat">Le format de date renseigné dans Winrest</param>
        /// <param name="listErrors">La liste des erreurs détectées des précédents tests.</param>
        private void WA_RN_Filter_Date(ReceiptNotesPage receiptNotesPage, string dateFormat, List<string> listErrors)
        {
            try
            {
                // Variables to initialize
                DateTime fromDate = DateUtils.Now.AddMonths(-1);
                DateTime toDate = DateUtils.Now.AddDays(7);

                //Filter on the date 
                receiptNotesPage.ResetFilter();
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.DateFrom, fromDate);
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.DateTo, toDate);

                if (receiptNotesPage.isPageSizeEqualsTo100() == false)
                {
                    receiptNotesPage.PageSize("100");
                }

                if (receiptNotesPage.IsDateRespected(fromDate, toDate, dateFormat) == false)
                {
                    listErrors.Add(String.Format(MessageErreur.FILTRE_ERRONE, "'From/To'"));
                }
            }
            catch (Exception ex)
            {
                listErrors.Add("Erreur dans WA_RN_Filter_Date : " + ex.Message);
            }
        }

        private void WA_RN_Filter_CheckAll_Validated_NotValidated(ReceiptNotesPage receiptNotesPage, List<string> listErrors)
        {
            try
            {
                receiptNotesPage.ResetFilter();

                //Show not validated
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ShowNotValidated, true);
                int notValidated = receiptNotesPage.CheckTotalNumber();

                if (receiptNotesPage.CheckValidation(false) == false)
                { listErrors.Add(String.Format(MessageErreur.FILTRE_ERRONE, "'Show not validated only'")); }

                //Show validated
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ShowValidatedOnly, true);
                int validated = receiptNotesPage.CheckTotalNumber();

                if (receiptNotesPage.CheckValidation(true) == false)
                { listErrors.Add(String.Format(MessageErreur.FILTRE_ERRONE, "'Show validated only'")); }

                //Show all
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ShowAllReceipts, true);
                int allResults = receiptNotesPage.CheckTotalNumber();

                if (allResults != notValidated + validated)
                { listErrors.Add(String.Format(MessageErreur.FILTRE_ERRONE, "'Show all receipts'")); }
            }
            catch (Exception ex)
            {
                listErrors.Add("Erreur dans WA_RN_Filter_CheckAll_Validated_NotValidated : " + ex.Message);
            }
        }

        /**
         * Test filtrage : Affichage des Receipt Note validées mais non facturées
         **/
        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_Filter_Show_Validated_and_not_invoiced()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();

            // Log in
            var homePage = LogInAsAdmin();

            //Act           
            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();
            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ShowValidatedNotInvoiced, true);

            if (receiptNotesPage.CheckTotalNumber() < 20)
            {
                var receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
                receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, deliveryPlace));
                var receiptNotesItem = receiptNotesCreateModalpage.Submit();

                var itemName = receiptNotesItem.GetFirstItemName();
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

                receiptNotesPage = receiptNotesItem.BackToList();
                receiptNotesPage.ResetFilter();
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ShowValidatedNotInvoiced, true);
            }

            if (!receiptNotesPage.isPageSizeEqualsTo100())
            {
                receiptNotesPage.PageSize("8");
                receiptNotesPage.PageSize("100");
            }

            bool checkValidation = receiptNotesPage.CheckValidation(true);
            bool areAllReceiptNoteNotInvoiced = receiptNotesPage.AreAllRnNotInvoiced();

            //Assert
            Assert.IsTrue(checkValidation, String.Format(MessageErreur.FILTRE_ERRONE, "'Show validated only'"));
            Assert.IsTrue(areAllReceiptNoteNotInvoiced, String.Format(MessageErreur.FILTRE_ERRONE, "'Show validated and not invoiced only'"));
        }

        /**
         * Test filtrage : Affichage des Receipt Note validées partiellement facturées
         **/
        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_Filter_Show_Validated_partial_invoiced()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();
            string supplierInvoiceNumber = new Random().Next().ToString();
            HomePage homePage = LogInAsAdmin();
            ReceiptNotesPage receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();
            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ShowValidatedPartialInvoiced, true);
            if (receiptNotesPage.CheckTotalNumber() < 20)
            {
                ReceiptNotesCreateModalPage receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
                receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, deliveryPlace));
                ReceiptNotesItem receiptNotesItem = receiptNotesCreateModalpage.Submit();
                string itemName = receiptNotesItem.GetFirstItemName();
                receiptNotesItem.SelectFirstItem();
                receiptNotesItem.AddReceived(itemName, "10");
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
                SupplierInvoicesItem supplierInvoiceItems = receiptNotesItem.GenerateSupplierInvoice(supplierInvoiceNumber);
                supplierInvoiceItems.SelectFirstItem();
                supplierInvoiceItems.SetItemQuantity(itemName, "1");
                supplierInvoiceItems.ValidateSupplierInvoice();
                supplierInvoiceItems.Close();
                receiptNotesPage = receiptNotesItem.BackToList();
                receiptNotesPage.ResetFilter();
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ShowValidatedPartialInvoiced, true);
            }
            if (!receiptNotesPage.isPageSizeEqualsTo100())
            {
                receiptNotesPage.PageSize("8");
                receiptNotesPage.PageSize("100");
            }
            bool isValidated = receiptNotesPage.CheckValidation(true);
            bool isPartiallyInvoiced = receiptNotesPage.IsPartiallyInvoiced();
            Assert.IsTrue(isValidated, String.Format(MessageErreur.FILTRE_ERRONE, "'Show validated only'"));
            Assert.IsTrue(isPartiallyInvoiced, String.Format(MessageErreur.FILTRE_ERRONE, "'Show validated and partially invoiced only'"));
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_Filter_Show_ReceiptNotes_With_Claims()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();

            // Log in
            var homePage = LogInAsAdmin();

            //Act            
            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();
            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ShowWithClaim, true);

            if (receiptNotesPage.CheckTotalNumber() < 20)
            {
                // Create
                var receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
                receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, deliveryPlace));

                var receiptNotesItem = receiptNotesCreateModalpage.Submit();

                // Update the first item value to activate the activation menu
                var itemName = receiptNotesItem.GetFirstItemName();
                receiptNotesItem.SelectFirstItem();
                receiptNotesItem.AddReceived(itemName, "2");

                receiptNotesItem.Refresh();
                receiptNotesItem.SelectItem(itemName);

                try
                {
                    receiptNotesItem.CheckClaim(itemName);
                }
                catch
                {
                    var editClaimForm = receiptNotesItem.EditClaimForm(itemName);
                    editClaimForm.SetClaimFormForClaimedV3();
                }

                var qualityChecks = receiptNotesItem.ClickOnChecksTab();
                qualityChecks.DeliveryPartial();

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

                var claimsItem = receiptNotesItem.ValidateClaim();
                claimsItem.Close();

                receiptNotesPage = receiptNotesItem.BackToList();
                receiptNotesPage.ResetFilter();
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ShowWithClaim, true);
            }

            if (!receiptNotesPage.isPageSizeEqualsTo100())
            {
                receiptNotesPage.PageSize("8");
                receiptNotesPage.PageSize("100");
            }

            //Assert
            Assert.IsTrue(receiptNotesPage.IsWithClaim(), String.Format(MessageErreur.FILTRE_ERRONE, "'Show receipt notes with claims'"));
        }

        [Ignore]
        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_Filter_Show_ExportedForSageManually()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();
            string itemName = TestContext.Properties["Item_ReceiptNote"].ToString();

            bool newVersionPrint = true;

            // Log in
            var homePage = LogInAsAdmin();

            // Config pour export SAGE auto mais pas pour le site
            homePage.SetSageAutoEnabled(site, true, "Due to invoice receipt note", false);

            try
            {
                //Act
                var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
                receiptNotesPage.ResetFilter();
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ShowExportedForSageManually, true);

                if (receiptNotesPage.CheckTotalNumber() < 20)
                {
                    // Create
                    var receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
                    receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, deliveryPlace));
                    var receiptNotesItem = receiptNotesCreateModalpage.Submit();

                    receiptNotesItem.Filter(ReceiptNotesItem.FilterItemType.SearchByName, itemName);

                    // Update the first item value to activate the activation menu
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

                    receiptNotesItem.ExportResultsForSage(newVersionPrint);

                    receiptNotesPage = receiptNotesItem.BackToList();
                    receiptNotesPage.ResetFilter();
                    receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ShowExportedForSageManually, true);
                }

                if (!receiptNotesPage.isPageSizeEqualsTo100())
                {
                    //receiptNotesPage.PageSize("8");
                    receiptNotesPage.PageSize("100");
                }

                //Assert
                Assert.IsTrue(receiptNotesPage.IsSentToSageManually(), String.Format(MessageErreur.FILTRE_ERRONE, "'Exported for sage manually'"));

            }
            finally
            {
                homePage.SetSageAutoEnabled(site, false);
            }
        }

        /// <summary>
        /// Teste les statuts "validés", "non validés" et le statut "tout"
        /// </summary>
        /// <param name="receiptNotesPage">La page d'index RN</param>
        /// <param name="listErrors">La liste des erreurs détectées des précédents tests.</param>
        private void WA_RN_Filter_CheckAll_Active_NotActive(ReceiptNotesPage receiptNotesPage, List<string> listErrors)
        {
            try
            {
                receiptNotesPage.ResetFilter();

                //Check active status
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ShowActive, true);
                int active = receiptNotesPage.CheckTotalNumber();

                if (receiptNotesPage.CheckStatus(true) == false)
                { listErrors.Add(String.Format(MessageErreur.FILTRE_ERRONE, "'Show activated only'")); }

                //Check inactive status
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ShowInactive, true);
                int inactive = receiptNotesPage.CheckTotalNumber();

                if (receiptNotesPage.CheckStatus(true) == false)
                { listErrors.Add(String.Format(MessageErreur.FILTRE_ERRONE, "'Show inactivated only'")); }

                //Check all
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ShowAll, true);
                int allResults = receiptNotesPage.CheckTotalNumber();

                if (allResults != inactive + active)
                { listErrors.Add(String.Format(MessageErreur.FILTRE_ERRONE, "'Show all'")); }
            }
            catch (Exception ex)
            {
                listErrors.Add("Erreur dans WA_RN_Filter_CheckAll_Active_NotActive : " + ex.Message);
            }
        }

        /// <summary>
        /// Test le classement du "sort by"
        /// </summary>
        /// <param name="receiptNotesPage">La page d'index RN</param>
        /// <param name="listErrors">La liste des erreurs détectées des précédents tests.</param>
        private void WA_RN_Filter_SortBy(ReceiptNotesPage receiptNotesPage, List<string> listErrors, string dateFormat)
        {
            try
            {
                receiptNotesPage.ResetFilter();

                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.SortBy, "NUMBER");
                if (receiptNotesPage.IsSortedByNumber() == false)
                { listErrors.Add(String.Format(MessageErreur.FILTRE_ERRONE, "'Sort by number'")); }

                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.SortBy, "DATE");
                if (receiptNotesPage.IsSortedByDate(dateFormat))
                { String.Format(MessageErreur.FILTRE_ERRONE, "'Sort by date'"); }
            }
            catch (Exception ex)
            {
                listErrors.Add("Erreur dans WA_RN_Filter_SortBy : " + ex.Message);
            }
        }

        /**
         * Test filtrage : Affichage des Receipt Note Ouvertes
         **/
        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_Filter_Show_Opened()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();
            string downloadPath = TestContext.Properties["DownloadsPath"].ToString();

            bool newVersionPrint = true;

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();

            receiptNotesPage.ClearDownloads();

            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.Opened, true);
            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.DateFrom, DateUtils.Now.AddMonths(-1));
            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.DateTo, DateUtils.Now);

            if (receiptNotesPage.CheckTotalNumber() < 20)
            {
                var receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
                receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, deliveryPlace));
                var receiptNotesItem = receiptNotesCreateModalpage.Submit();

                var itemName = receiptNotesItem.GetFirstItemName();
                receiptNotesItem.SelectFirstItem();
                receiptNotesItem.AddReceived(itemName, "2");
                receiptNotesItem.Refresh();

                receiptNotesPage = receiptNotesItem.BackToList();
                receiptNotesPage.ResetFilter();
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.Opened, true);
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.DateFrom, DateUtils.Now.AddMonths(-1));
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.DateTo, DateUtils.Now);
            }

            // On exporte les résultats sous la forme d'un fichier Excel (dont on récupère le nom)
            receiptNotesPage.ExportExcelFile(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = receiptNotesPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

            // Récupération du nom du fichier et construction de l'URL du fichier Excel à ouvrir   
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadPath, fileName);

            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber(RECEIPT_NOTES_EXCEL_SHEET, filePath);
            bool result = OpenXmlExcel.ReadAllDataInColumn("Status", RECEIPT_NOTES_EXCEL_SHEET, filePath, "Opened");

            //Assert
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.IsTrue(result, MessageErreur.EXCEL_DONNEES_KO);
        }

        /**
         * Test filtrage : Affichage des Receipt Note Cloturées
         **/
        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_Filter_Show_Closed()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();
            string downloadPath = TestContext.Properties["DownloadsPath"].ToString();

            bool newVersionPrint = true;

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();

            receiptNotesPage.ClearDownloads();

            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.Closed, true);
            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.DateFrom, DateUtils.Now.AddMonths(-1));
            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.DateTo, DateUtils.Now);

            if (receiptNotesPage.CheckTotalNumber() < 20)
            {
                var receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
                receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, deliveryPlace));
                var receiptNotesItem = receiptNotesCreateModalpage.Submit();

                var itemName = receiptNotesItem.GetFirstItemName();
                receiptNotesItem.SelectFirstItem();
                receiptNotesItem.AddReceived(itemName, "2");
                receiptNotesItem.Refresh();

                var receiptNotesGeneralInformationPage = receiptNotesItem.ClickOnGeneralInformationTab();
                receiptNotesGeneralInformationPage.SetStatus("Closed");
                receiptNotesPage = receiptNotesGeneralInformationPage.BackToList();
                receiptNotesPage.ResetFilter();
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.Closed, true);
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.DateFrom, DateUtils.Now.AddMonths(-1));
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.DateTo, DateUtils.Now);
            }

            // On exporte les résultats sous la forme d'un fichier Excel (dont on récupère le nom)
            receiptNotesPage.ExportExcelFile(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = receiptNotesPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

            // Récupération du nom du fichier et construction de l'URL du fichier Excel à ouvrir   
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadPath, fileName);

            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber(RECEIPT_NOTES_EXCEL_SHEET, filePath);
            bool result = OpenXmlExcel.ReadAllDataInColumn("Status", RECEIPT_NOTES_EXCEL_SHEET, filePath, "Closed");

            //Assert
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.IsTrue(result, MessageErreur.EXCEL_DONNEES_KO);
        }

        /**
         * Test filtrage : Affichage de l'ensemble des Receipt Note
         **/
        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_Filter_Show_AllStatus()
        {
            // Log in
            var homePage = LogInAsAdmin();

            //Act
            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();

            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.All, true);
            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.DateFrom, DateUtils.Now.AddMonths(-1));
            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.DateTo, DateUtils.Now);
            var numberAllStatus = receiptNotesPage.CheckTotalNumber();

            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.Opened, true);
            var numberOpened = receiptNotesPage.CheckTotalNumber();

            receiptNotesPage.ShowBtnStatus();

            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.Closed, true);
            var numberClosed = receiptNotesPage.CheckTotalNumber();

            Assert.AreEqual(numberAllStatus, numberOpened + numberClosed, "Erreur de filtrage All Statut !");

        }

        /// <summary>
        /// Test le filtrage par sites
        /// </summary>
        /// <param name="receiptNotesPage">La page d'index RN</param>
        /// <param name="listErrors">La liste des erreurs détectées des précédents tests.</param>
        private void WA_RN_Filter_Site(ReceiptNotesPage receiptNotesPage, List<string> listErrors)
        {
            try
            {
                string site = TestContext.Properties["Site"].ToString();
                receiptNotesPage.ResetFilter();
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.Site, site);

                if (receiptNotesPage.VerifySite(site) == false)
                { listErrors.Add(String.Format(MessageErreur.FILTRE_ERRONE, "'Sites'")); }
            }
            catch (Exception ex)
            {
                listErrors.Add("Erreur dans WA_RN_Filter_Site : " + ex.Message);
            }
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_Filter_SitePlaces()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();
            string downloadPath = TestContext.Properties["DownloadsPath"].ToString();

            bool newVersionPrint = true;

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();

            receiptNotesPage.ClearDownloads();

            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.DateFrom, DateUtils.Now.AddDays(-3));
            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.DateTo, DateUtils.Now);
            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.SitePlaces, site + "-" + deliveryPlace);

            if (receiptNotesPage.CheckTotalNumber() < 20)
            {
                var receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
                receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, deliveryPlace));
                var receiptNotesItem = receiptNotesCreateModalpage.Submit();

                var itemName = receiptNotesItem.GetFirstItemName();
                receiptNotesItem.SelectFirstItem();
                receiptNotesItem.AddReceived(itemName, "2");
                receiptNotesItem.Refresh();

                receiptNotesPage = receiptNotesItem.BackToList();

                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.DateFrom, DateUtils.Now.AddDays(-3));
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.DateTo, DateUtils.Now);
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.SitePlaces, site + "-" + deliveryPlace);
            }

            // On exporte les résultats sous la forme d'un fichier Excel (dont on récupère le nom)
            receiptNotesPage.ExportExcelFile(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = receiptNotesPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

            // Récupération du nom du fichier et construction de l'URL du fichier Excel à ouvrir   
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadPath, fileName);

            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber(RECEIPT_NOTES_EXCEL_SHEET, filePath);
            bool result = OpenXmlExcel.ReadAllDataInColumn("Site Place", RECEIPT_NOTES_EXCEL_SHEET, filePath, deliveryPlace);

            //Assert
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.IsTrue(result, String.Format(MessageErreur.FILTRE_ERRONE, "'SitePlaces'"));
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_Items_Filters_SearchByName()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();

            // Log in
            var homePage = LogInAsAdmin();

            //Act            
            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();

            // Create
            var receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
            receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, placeTo));

            var receiptNotesItem = receiptNotesCreateModalpage.Submit();

            var itemName = receiptNotesItem.GetFirstItemName();

            // Filter by item name 
            receiptNotesItem.Filter(ReceiptNotesItem.FilterItemType.SearchByName, itemName);
            Assert.IsTrue(receiptNotesItem.VerifyName(itemName), String.Format(MessageErreur.FILTRE_ERRONE, "'Search item by name'"));
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_Items_Filters_ByGroup()
        {
            //prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();
            string receiptNoteNumber = String.Empty;
            //arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Affectation de la valeur de newVersionprint = true
            homePage.SetNewVersionKeywordValue(true);

            //Act            
            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();
            try
            {
                // Create
                var receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
                receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, placeTo));

                var receiptNotesItem = receiptNotesCreateModalpage.Submit();
                receiptNoteNumber = receiptNotesItem.GetReceiptNoteNumber();
                var itemName = receiptNotesItem.GetFirstItemName();
                // Add KeyWord to the item
                receiptNotesItem.SelectFirstItem();
                var itemPageItem = receiptNotesItem.EditItem(itemName);
                var itemGroup = itemPageItem.GetGroupName();
                itemPageItem.Close();

                // Filter by group
                receiptNotesItem.PageSize("8");
                receiptNotesItem.Filter(ReceiptNotesItem.FilterItemType.ByGroup, itemGroup);
                var verifySameGroup = receiptNotesItem.VerifyGroup(itemGroup);
                Assert.IsTrue(verifySameGroup, String.Format(MessageErreur.FILTRE_ERRONE, "'Sort by Group'"));
                receiptNotesItem.BackToList();

            }
            finally
            {
                receiptNotesPage.ResetFilter();
                if (!String.IsNullOrEmpty(receiptNoteNumber))
                {
                    receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ByNumber, receiptNoteNumber);
                    receiptNotesPage.DeleteReceiptNote();
                    receiptNotesPage.ResetFilter();
                }
            }

        }

        /*
         * Test de changement de la valeur de Received pour une Receipt Note
         */
        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_ReceivedReceiptNote()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();
            string receivedValueTest = "2";
            int temperatureValue = 15;
            // Log in

            var homePage = LogInAsAdmin();

            //Act            
            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();

            // Create
            var receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
            receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, deliveryPlace));

            var receiptNotesItem = receiptNotesCreateModalpage.Submit();

            var itemName = receiptNotesItem.GetFirstItemName();
            receiptNotesItem.SelectFirstItem();
            string oldReceived = receiptNotesItem.GetReceived();
            receiptNotesItem.AddTemperature(temperatureValue);
            receiptNotesItem.AddReceived(itemName, receivedValueTest);
            receiptNotesItem.Refresh();

            string newReceived = receiptNotesItem.GetReceived();

            Assert.AreNotEqual(oldReceived, newReceived, "La valeur de 'received' n'a pas été mise à jour.");
            Assert.IsTrue(receiptNotesItem.IsIconDisplayed(), "Le symbole n'est pas apparait à côté du nom..");
        }

        /**
         * Test d'ajout d'un commentaire sur un item d'une Receipt Note
         **/
        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_CommentReceiptNoteItem()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();
            string newComment = "I am a comment";
            string itemName = TestContext.Properties["Item_ReceiptNote"].ToString();
            HomePage homePage = LogInAsAdmin();
            ReceiptNotesPage receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();
            ReceiptNotesCreateModalPage receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
            receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, deliveryPlace));
            ReceiptNotesItem receiptNotesItem = receiptNotesCreateModalpage.Submit();
            receiptNotesItem.Filter(ReceiptNotesItem.FilterItemType.SearchByName, itemName);
            receiptNotesItem.SelectFirstItem();
            receiptNotesItem.AddReceived(itemName, "2");
            receiptNotesItem.Refresh();
            receiptNotesItem.SelectFirstItem();
            receiptNotesItem.AddComment(newComment);
            receiptNotesItem.Refresh();
            receiptNotesItem.SelectFirstItem();
            string comment = receiptNotesItem.GetComment(itemName);
            Assert.AreEqual(comment, newComment, "L'ajout du commentaire dans la receipt note a échoué.");
        }

        /**
         * Test d'ajout d'un des quality checks
         **/
        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_Fill_QualityChecks()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();
            var temp = -15;
            int locationArobase = TestContext.Properties["Admin_UserName"].ToString().IndexOf("@");
            string userName = TestContext.Properties["Admin_UserName"].ToString().Substring(0, locationArobase);

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act            
            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();

            // Create
            var receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
            receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, deliveryPlace) { IsActivated = false });

            var receiptNotesItem = receiptNotesCreateModalpage.Submit();
            var qualityChecks = receiptNotesItem.ClickOnChecksTab();

            // Tests
            Assert.IsTrue(String.IsNullOrEmpty(qualityChecks.GetVerifiedBy()));

            qualityChecks.SetFrozenTemperature(temp.ToString());
            qualityChecks.DeliveryAccepted();
            Assert.AreEqual(qualityChecks.GetVerifiedBy(), userName, "La receipt note n'a pas passé les quality checks.");
        }

        /*
         * Création d'une nouvelle Receipt Note Et validation de celle-ci
        */
        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_ValidateReceiptNote()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();

            // Log in
            var homePage = LogInAsAdmin();

            //Act           
            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();

            // Create
            var receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
            receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, deliveryPlace));
            var receiptNotesItem = receiptNotesCreateModalpage.Submit();
            receiptNotesItem.ResetFilters();

            // Update the first item value to activate the activation menu
            var itemName = receiptNotesItem.GetFirstItemName();
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

            // Assert
            Assert.IsFalse(receiptNotesItem.IsModifiableReceivedValue(itemName), "Les valeurs des items sont modifiables malgré la validation de la receipt note.");
        }

        /*
         * Génération d'une claim à partir d'une receipt note
        */
        [Timeout(_timeout)]
        [TestMethod]
        public void WA_RN_GenerateClaimWithReceiptNote()
        {
            //Prepare
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["Supplier"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();
            string itemName = "Item-" + DateUtils.Now.ToString("dd/MM/yyyy") + "-" + new Random().Next().ToString();
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();

            //Act
            HomePage homePage = LogInAsAdmin();
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();

            // Create New Item
            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);

            //Create New Packaging
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
            itemGeneralInformationPage.BackToList();

            // Go to warehouse ReceiptNote
            ReceiptNotesPage receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();

            //Create New ReceiptNote
            ReceiptNotesCreateModalPage receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
            receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, deliveryPlace));
            ReceiptNotesItem receiptNotesItem = receiptNotesCreateModalpage.Submit();
            receiptNotesItem.ResetFilters();
            receiptNotesItem.Filter(ReceiptNotesItem.FilterItemType.SearchByName, itemName);
            receiptNotesItem.SelectFirstItem();

            // AddReceived
            receiptNotesItem.AddReicedQuantity(itemName, "2");
            receiptNotesItem.Refresh();
            string receivedItems = receiptNotesItem.GetReceived();

            // Click Boutton Megaphone et ajouter Sanction Amount
            receiptNotesItem.ClickBtnMegaphone();
            receiptNotesItem.AddSanctionAmount("100");

            string itemNameClaims = receiptNotesItem.GetItemNameClaimsText();
            string receivedClaims = receiptNotesItem.GetReceivedClaimsText();

            //Assert 
            Assert.AreEqual(itemName, itemNameClaims, "Les Item name ne sont pas le meme !!");
            Assert.AreEqual(receivedItems, receivedClaims, "La Quantity Received ne sont pas le meme !!");

            // Allez dans ChecksTab et ajoutez les paramètres nécessaires. 
            ReceiptNotesQualityChecks qualityChecks = receiptNotesItem.ClickOnChecksTab();
            qualityChecks.DeliveryPartial();
            if (qualityChecks.CanClickOnSecurityChecks())
            {
                qualityChecks.CanClickOnSecurityChecks();
                qualityChecks.SetSecurityChecks("No");
                qualityChecks.SetQualityChecks();
            }

            // Allez dans Item Sous ReceiptNote et valider la Receipt. 
            receiptNotesItem = qualityChecks.ClickOnReceiptNoteItemTab();
            ClaimsItem claimsItem = receiptNotesItem.ValidateClaim();

            // Allez dans General Information Sous ReceiptNote 
            ClaimsGeneralInformation claimsGeneralInformation = claimsItem.ClickOnGeneralInformation();
            string claimNumber = claimsGeneralInformation.GetClaimNumber();
            claimsItem.Close();

            // Go to warehouse ClaimsPage
            ClaimsPage claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();
            claimsPage.Filter(ClaimsPage.FilterType.ByNumber, claimNumber);
            string firstId = claimsPage.GetFirstID();

            //Assert
            Assert.AreEqual(claimNumber, firstId, "La claim " + claimNumber + " n'a pas été créé à partir de la receipt note.");
        }

        /*
         * Test de rafraichissement d'une nouvelle Receipt Note 
        */
        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_RefreshReceiptNote()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();
            string receiptNoteNumber = string.Empty;
            HomePage homePage = LogInAsAdmin();
            ReceiptNotesPage receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();
            try
            {
                // Create New Receipt Note 
                ReceiptNotesCreateModalPage receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
                receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, deliveryPlace) { IsActivated = false });
                ReceiptNotesItem receiptNotesItem = receiptNotesCreateModalpage.Submit();
                receiptNoteNumber = receiptNotesItem.GetReceiptNoteNumber();
                // Filter and Select Item  
                string itemName = receiptNotesItem.GetFirstItemName();
                receiptNotesItem.ResetFilters();
                receiptNotesItem.Filter(ReceiptNotesItem.FilterItemType.SearchByName, itemName);
                string initialReceived = receiptNotesItem.GetReceived();
                double received = double.Parse(initialReceived) + 2;
                receiptNotesItem.SelectFirstItem();
                // Add Received and refresh
                receiptNotesItem.AddReceived(itemName, received.ToString());
                receiptNotesItem.Refresh();
                string updatedRecevied = receiptNotesItem.GetReceived();
                // Assert 
                Assert.AreNotEqual(initialReceived, updatedRecevied, "Le refresh n'a pas fonctionné pour les données saisies.");
                receiptNotesItem.BackToList();
            }
            finally
            {
                // Filter Receipt Note Created And Delete
                receiptNotesPage.ResetFilter();
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ByNumber, receiptNoteNumber);
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ShowAll, true);
                receiptNotesPage.DeleteReceiptNote();
            }
            
        }

        /*
         * Test d'impression d'une nouvelle Receipt Note  avec newVersionPrint = true
        */
        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_Details_PrintReceiptNoteNewVersion()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            //Receipt Note Report_-_437569_-_20220316145020.pdf
            string DocFileNamePdfBegin = "Receipt Note Report_-_";
            //All_files_20220225_102148.zip
            string DocFileNameZipBegin = "All_files_";
            bool newVersionPrint = true;

            // Log in
            var homePage = LogInAsAdmin();

            //Act           
            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();

            receiptNotesPage.ClearDownloads();
            DeleteAllFileDownload();
            var receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
            receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, deliveryPlace));
            var receiptNotesItem = receiptNotesCreateModalpage.Submit();
            // Update the first item value to activate the activation menu
            var itemName = receiptNotesItem.GetFirstItemName();
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
            string newReceived = receiptNotesItem.GetReceived();
            receiptNotesItem.Validate();

            PrintReportPage reportPage = PrintGenerique(receiptNotesItem, newVersionPrint);
            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);

            ReceiptNotesGeneralInformation generalInfo = receiptNotesItem.ClickOnGeneralInformationTab();
            string claimNumber = generalInfo.GetReceiptNoteNumber();
            string deliveryOrderNumber = generalInfo.GetReceiptNoteDeliveryOrderNumber();

            // cliquer sur All
            string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
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
            Assert.AreNotEqual(0, mots.Count(w => w == site + "(" + site + ")"), site + "(" + site + ") non présent dans le Pdf");
            Assert.AreNotEqual(0, mots.Count(w => w == newReceived ), newReceived ," non présent dans le Pdf");
            Assert.AreNotEqual(0, mots.Count(w => w == DateUtils.Now.Date.ToString("dd/MM/yyyy")), DateUtils.Now.Date.ToString("dd/MM/yyyy") + " non présent dans le Pdf");
            Assert.AreNotEqual(0, mots.Count(w => w == claimNumber), claimNumber + " non présent dans le Pdf");
            Assert.AreNotEqual(0, mots.Count(w => w == deliveryOrderNumber), deliveryOrderNumber + " non présent dans le Pdf");
        }

        private PrintReportPage PrintGenerique(ReceiptNotesItem receiptNotesItem, bool printValue)
        {
            var reportPage = receiptNotesItem.PrintReceiptNote(printValue);
            var isReportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();

            if (printValue)
            {
                receiptNotesItem.ClickPrintButton();
            }

            Assert.IsTrue(isReportGenerated, "L'application n'a pas pu générer le fichier attendu.");

            return reportPage;
        }

        /*
         * Test d'impression d'une nouvelle Receipt Note depuis l'index avec newVersionPrint = true
        */
        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_Index_PrintReceiptNoteResultsNewVersion()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            //Receipt Note Report_-_437579_-_20220316154719.pdf
            string DocFileNamePdfBegin = "Receipt Note Report_-_";
            //All_files_20220225_102148.zip
            string DocFileNameZipBegin = "All_files_";
            bool newVersionPrint = true;

            // Log in
            var homePage = LogInAsAdmin();
            //Act           
            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();
            DeleteAllFileDownload();
            receiptNotesPage.ClearDownloads();
            // Create Receipt Note 
            var receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
            receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, deliveryPlace));
            var receiptNotesItem = receiptNotesCreateModalpage.Submit();
            string receiptNoteNumber = receiptNotesItem.GetReceiptNoteNumber();

            // Update the first item value to activate the activation menu
            var itemName = receiptNotesItem.GetFirstItemName();
            // Select Item and add Rdceived
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
            // Validate Receipt Note
            receiptNotesItem.Validate();

            ReceiptNotesGeneralInformation generalInfo = receiptNotesItem.ClickOnGeneralInformationTab();
            string claimNumber = generalInfo.GetReceiptNoteNumber();
            string deliveryOrderNumber = generalInfo.GetReceiptNoteDeliveryOrderNumber();
            receiptNotesPage = generalInfo.BackToList();
            receiptNotesPage.ResetFilter();
            //Filter Receipt note Created
            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ByNumber, receiptNoteNumber);
            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.DateFrom, DateUtils.Now);
            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.DateTo, DateUtils.Now.AddMonths(1));

            // Print Data of Receipt Note Created 
            var reportPage = receiptNotesPage.PrintReceiptNoteResults(newVersionPrint);
            var isReportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();

            // Assert check if Report Generated
            Assert.IsTrue(isReportGenerated, "L'application n'a pas pu générer le fichier attendu.");
            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);


            // cliquer sur All
            string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
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
            Assert.AreNotEqual(0, mots.Count(w => w == site + "(" + site + ")"), site + "(" + site + ") non présent dans le Pdf");
            Assert.AreNotEqual(0, mots.Count(w => w == DateUtils.Now.Date.ToString("dd/MM/yyyy")), DateUtils.Now.Date.ToString("dd/MM/yyyy") + " non présent dans le Pdf");
            Assert.AreNotEqual(0, mots.Count(w => w == claimNumber), claimNumber + " non présent dans le Pdf");
            Assert.AreNotEqual(0, mots.Count(w => w == deliveryOrderNumber), deliveryOrderNumber + " non présent dans le Pdf");
        }

        //_____________________________________FIN RECEIPTS NOTES PRINT_________________________________________________

        //_____________________________________RECEIPTS NOTES GENERAL INFOS_________________________________________________

        /*
         * Test d'affichage des informations générales d'une Receipt Note 
        */
        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_GlobalInformationsReceiptNote()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();

            // Log in
            var homePage = LogInAsAdmin();

            //Act            
            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();

            // Create
            var receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
            receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, deliveryPlace) { IsActivated = false });

            var receiptNotesItem = receiptNotesCreateModalpage.Submit();

            string receiptNoteNumber = receiptNotesItem.GetReceiptNoteNumber();

            //1. Cliquer sur l'onglet "general information"
            var receiptNoteGeneralInfo = receiptNotesItem.ClickOnGeneralInformationTab();

            // Tests
            Assert.AreEqual(receiptNoteNumber, receiptNoteGeneralInfo.GetReceiptNoteNumber(), "L'ID de la receipt note analysée ne correspond à celui de la receipt note créée");

            //2.Désactiver / activer la RN
            receiptNoteGeneralInfo.SetGeneralInfoActivateState(false);
            Assert.IsFalse(receiptNoteGeneralInfo.GetGeneralInfoActivateState(), "La receipt note n'est pas inactive dans les informations générales.");

            receiptNoteGeneralInfo.SetGeneralInfoActivateState(true);
            Assert.IsTrue(receiptNoteGeneralInfo.GetGeneralInfoActivateState(), "La receipt note n'est pas activée dans les informations générales.");

            //3. Modifier la date, le Delivery order number, le comment
            Random rnd = new Random();
            var deliveryOrderNumber = rnd.Next().ToString();
            receiptNoteGeneralInfo.Fill(DateUtils.Now, deliveryOrderNumber, "Blabla");

            //4. Changer d'onglet et vérifier que les infos ont bien été modifiées
            WebDriver.Navigate().Refresh();
            receiptNotesItem = receiptNoteGeneralInfo.ClickOnItemsTab();

            receiptNotesItem.ClickOnGeneralInformationTab();

            Assert.AreEqual(deliveryOrderNumber, receiptNoteGeneralInfo.GetDeliveryOrderNumber());
            Assert.AreEqual(DateUtils.Now.ToString("dd/MM/yyyy"), receiptNoteGeneralInfo.GetDate());
            Assert.AreEqual("Blabla", receiptNoteGeneralInfo.GetComment());
        }

        //_____________________________________FIN RECEIPTS NOTES GENERAL INFOS_________________________________________________

        //_____________________________________RECEIPTS NOTES EXPORT____________________________________________________________

        /*
         * Export de la liste des résultats d'une recherche de Receipt Note avec newVersionPrint = true
        */
        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_ExportReceiptNotesListNewVersion()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            bool newVersionPrint = true;

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act           
            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();

            receiptNotesPage.ClearDownloads();

            if (receiptNotesPage.CheckTotalNumber() < 20)
            {
                var receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
                receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, deliveryPlace));
                var receiptNotesItem = receiptNotesCreateModalpage.Submit();

                var itemName = receiptNotesItem.GetFirstItemName();
                receiptNotesItem.SelectFirstItem();
                receiptNotesItem.AddReceived(itemName, "2");
                receiptNotesItem.Refresh();

                receiptNotesPage = receiptNotesItem.BackToList();
                receiptNotesPage.ResetFilter();
            }

            // Lancement de l'export avec la première valeur de printValue
            ExportGenerique(receiptNotesPage, newVersionPrint, downloadsPath);
        }

        private void ExportGenerique(ReceiptNotesPage receiptNotesPage, bool printValue, string downloadsPath)
        {
            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.DateFrom, DateUtils.Now.AddMonths(-1));
            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.DateTo, DateUtils.Now);


            DeleteAllFileDownload();

            // On exporte les résultats sous la forme d'un fichier Excel (dont on récupère le nom)
            receiptNotesPage.ExportExcelFile(printValue);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = receiptNotesPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

            // Récupération du nom du fichier et construction de l'URL du fichier Excel à ouvrir   
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber(RECEIPT_NOTES_EXCEL_SHEET, filePath);

            //Assert
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);
        }

        private void ExportTopReceivedGenerique(ReceiptNotesPage receiptNotesPage, bool printValue, string downloadsPath, string decimalSeparator, string site)
        {
            int nbTopReceived = 10;

            // On exporte les résultats sous la forme d'un fichier Excel (dont on récupère le nom)
            // Export du fichier au format Excel
            receiptNotesPage.ExportTopReceivedExcelFile(printValue, downloadsPath, site, nbTopReceived.ToString());

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = receiptNotesPage.GetTopReceivedExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

            // Récupération du nom du fichier et construction de l'URL du fichier Excel à ouvrir   
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber("Top received", filePath);
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);

            List<string> totalList = OpenXmlExcel.GetValuesInList("Total", "Top received", filePath, decimalSeparator);
            Assert.IsTrue(IsDecreasedList(totalList, decimalSeparator), "Les données du fichier TopReceived exporté ne sont pas triées.");
        }

        private bool IsDecreasedList(List<string> list, string decimalSeparator)
        {
            bool valueBool = true;
            Double oldValue;

            // Récupération du type de séparateur (, ou . selon les pays) "." pour Excel en patch, "," pour excel en local
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            Console.WriteLine("###############A");
            for (int j = 0; j < list.Count; j++)
            {
                Console.WriteLine(list[j]);
            }
            Console.WriteLine("###############B");

            if (list != null)
            {

                oldValue = Double.Parse(list[0], ci);

                for (int i = 1; i < list.Count; i++)
                {

                    //915437599999995 OK
                    //5.49262560000001E+16 KO
                    double value = Double.Parse(list[i], ci);
                    if (oldValue < value)
                    {
                        valueBool = false;
                    }

                    oldValue = value;
                }
            }
            else
            {
                valueBool = false;
            }

            return valueBool;
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_ExportMatchingRNNewVersion()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();
            string currency = TestContext.Properties["Currency"].ToString();
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            bool newVersionPrint = true;

            // Log in
            var homePage = LogInAsAdmin();
            //Act           
            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();

            receiptNotesPage.ClearDownloads();
            DeleteAllFileDownload();
            var receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
            receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, deliveryPlace));
            var receiptNotesItem = receiptNotesCreateModalpage.Submit();
            string rnNumber = receiptNotesItem.GetReceiptNoteNumber();

            var itemName = receiptNotesItem.GetFirstItemName();
            receiptNotesItem.SelectFirstItem();
            receiptNotesItem.AddReceived(itemName, "2");
            receiptNotesItem.Refresh();

            receiptNotesPage = receiptNotesItem.BackToList();
            receiptNotesPage.ResetFilter();
            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.DateTo, DateUtils.Now);
            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.DateFrom, DateUtils.Now.AddDays(-1));
            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ByNumber, rnNumber);

            // On exporte les résultats sous la forme d'un fichier Excel (dont on récupère le nom)
            receiptNotesPage.ExportMatchingFile(newVersionPrint, downloadsPath);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = receiptNotesPage.GetExportMatchingFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

            // Récupération du nom du fichier et construction de l'URL du fichier Excel à ouvrir   
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber("Matching", filePath);

            //Assert
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);

            List<string> keys = OpenXmlExcel.GetValuesInList("RN number", "Matching", filePath);
            List<string> values = OpenXmlExcel.GetValuesInList("RN total no VAT", "Matching", filePath);

            receiptNotesPage.PageSize("1000");

            Dictionary<string, string> table = receiptNotesPage.GetRNPrices(currency);
            foreach (var key in keys)
            {
                Assert.IsTrue(table.ContainsKey(key), "Clé " + key + " dans le fichier exporté mais pas dans la table");
                Assert.AreEqual(table[key], values[keys.IndexOf(key)], "Valeur " + table[key] + " dans le fichier exporté mais pas dans le tableau");
            }

            //FIXME
            //Assert.AreEqual(resultNumber, receiptNotesPage.CheckTotalNumber(),"Filtres non appliqués");
        }

        [Ignore]
        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_SageManual_Details_ExportForSage_NewVersion()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();
            string itemName = TestContext.Properties["Item_ReceiptNote"].ToString();
            string currency = TestContext.Properties["Currency"].ToString();
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            bool newVersionPrint = true;

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Config pour export SAGE manuel
            homePage.SetSageAutoEnabled(site, false);

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            //Act
            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();

            receiptNotesPage.ClearDownloads();

            var receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
            receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, deliveryPlace));
            var receiptNotesItem = receiptNotesCreateModalpage.Submit();
            receiptNotesItem.Filter(ReceiptNotesItem.FilterItemType.SearchByName, itemName);

            // Update the first item value to activate the activation menu
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

            var receiptNoteFooterPage = receiptNotesItem.ClickOnFooterTab();
            double montantRN = receiptNoteFooterPage.GetReceiptNoteTotal(currency, decimalSeparatorValue);

            receiptNotesItem = receiptNoteFooterPage.ClickOnItemsTab();

            receiptNotesItem.ExportResultsForSage(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = receiptNotesPage.GetExportForSageFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

            // Récupération du nom du fichier et construction de l'URL du fichier txt à exploiter   
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // On n'exploite que les lignes avec contenu "général" --> "G"
            double[] contenuFichier = ExploitTextFiles.VerifySAGEFileContent(filePath, "G", ",");

            Assert.AreEqual(contenuFichier[0].ToString(), contenuFichier[1].ToString(), "Le contenu du fichier SAGE est incorrect.");

            // Remarque : pour les RN, le montant issu du fichier SAGE est égal au double du montant de la RN
            Assert.AreEqual(montantRN.ToString(), contenuFichier[0].ToString(), "Le montant issu du fichier SAGE n'est pas égal au montant de la Receipt Note défini dans l'application.");
        }

        [Ignore]
        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_SageManuel_Details_EnableSAGEExport()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();
            string itemName = TestContext.Properties["Item_ReceiptNote"].ToString();

            bool newVersionPrint = true;

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Config pour export SAGE manuel
            homePage.SetSageAutoEnabled(site, false);

            homePage.GetDecimalSeparatorValue();

            //Act
            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();

            receiptNotesPage.ClearDownloads();

            var receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
            receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, deliveryPlace));
            var receiptNotesItem = receiptNotesCreateModalpage.Submit();
            receiptNotesItem.Filter(ReceiptNotesItem.FilterItemType.SearchByName, itemName);

            // Update the first item value to activate the activation menu
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

            receiptNotesItem.ExportResultsForSage(newVersionPrint);

            Assert.IsTrue(receiptNotesItem.CanClickOnEnableSAGE(), "Il n'est pas possible de cliquer sur la fonctionnalité 'Enable export for SAGE' "
                + "pour un supplier invoice envoyé vers le SAGE.");

            receiptNotesItem.ClickOnEnableSAGE();

            Assert.IsTrue(receiptNotesItem.CanClickOnSAGE(), "Il est impossible de cliquer sur la fonctionnalité 'Export for SAGE' "
                + "après avoir cliqué sur 'Enable export for SAGE'.");

            Assert.IsFalse(receiptNotesItem.CanClickOnEnableSAGE(), "Il est possible de cliquer sur la fonctionnalité 'Enable export for SAGE' "
                + "pour un supplier invoice à envoyer vers le SAGE.");
        }

        [Ignore] //sage auto
        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_Index_GenerateSageTxt()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();
            string itemName = TestContext.Properties["Item_ReceiptNote"].ToString();
            var currency = TestContext.Properties["Currency"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            // Config pour export SAGE auto mais pas pour le site
            homePage.SetSageAutoEnabled(site, false);

            try
            {
                //Act
                var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
                receiptNotesPage.ResetFilter();

                homePage.ClearDownloads();

                // Create
                var receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
                receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now.AddDays(-40), site, supplier, deliveryPlace));
                var receiptNotesItem = receiptNotesCreateModalpage.Submit();

                string receiptNoteNumber = receiptNotesItem.GetReceiptNoteNumber();

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

                var receiptNotesFooter = receiptNotesItem.ClickOnFooterTab();
                double montantRN = receiptNotesFooter.GetReceiptNoteTotal(currency, decimalSeparatorValue);

                // Calcul du montant de la facture transmise à TL
                var receiptNoteAccounting = receiptNotesFooter.ClickOnAccounting();
                double montantGlobal = receiptNoteAccounting.GetReceiptNoteGrossAmount("G", decimalSeparatorValue);
                double montantDetaille = receiptNoteAccounting.GetReceiptNoteDetailAmount("G", decimalSeparatorValue);

                receiptNotesPage = receiptNotesItem.BackToList();
                receiptNotesPage.ResetFilter();
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ByNumber, receiptNoteNumber);
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.DateTo, DateUtils.Now.AddDays(-40));

                receiptNotesPage.ExportResultsForSage(true);

                // On récupère les fichiers du répertoire de téléchargement
                DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
                FileInfo[] taskFiles = taskDirectory.GetFiles();

                var correctDownloadedFile = receiptNotesPage.GetExportForSageFile(taskFiles);
                Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

                // Récupération du nom du fichier et construction de l'URL du fichier txt à exploiter   
                var fileName = correctDownloadedFile.Name;
                var filePath = Path.Combine(downloadsPath, fileName);

                // On n'exploite que les lignes avec contenu "général" --> "G"
                double[] contenuFichier = ExploitTextFiles.VerifySAGEFileContent(filePath, "G", decimalSeparatorValue);

                WebDriver.Navigate().Refresh();
                Assert.IsTrue(receiptNotesPage.IsSentToSageManually(), "La RN n'a pas été envoyée manuellement vers le SAGE.");
                Assert.AreEqual(contenuFichier[0].ToString(), contenuFichier[1].ToString(), "Le contenu du fichier SAGE généré est incorrect.");
                Assert.AreEqual(montantRN.ToString(), contenuFichier[0].ToString(), "Le montant issu du fichier SAGE n'est pas égal au montant de la RN défini dans l'application.");
                Assert.AreEqual(montantGlobal.ToString(), montantDetaille.ToString(), "Les montants AmountDebit et AmountCredit de la RN envoyée vers SAGE ne sont pas les mêmes dans l'onglet Accounting.");
                Assert.AreEqual(montantRN.ToString(), montantGlobal.ToString(), "Le montant issu du fichier SAGE n'est pas égal au montant de la RN défini dans l'application.");

            }
            finally
            {
                homePage.SetSageAutoEnabled(site, false);
            }
        }

        [Ignore]
        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_SageManuel_Details_ExportTXT()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();
            string itemName = TestContext.Properties["Item_ReceiptNote"].ToString();
            var currency = TestContext.Properties["Currency"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            bool newVersionPrint = true;

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            // Config pour export SAGE manuel
            homePage.SetSageAutoEnabled(site, false);

            //Act
            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();

            homePage.ClearDownloads();

            // Create
            var receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
            receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, deliveryPlace));
            var receiptNotesItem = receiptNotesCreateModalpage.Submit();

            string receiptNoteNumber = receiptNotesItem.GetReceiptNoteNumber();

            receiptNotesItem.Filter(ReceiptNotesItem.FilterItemType.SearchByName, itemName);
            receiptNotesItem.SelectFirstItem();
            receiptNotesItem.AddReceived(itemName, "2");
            var qualityChecks = receiptNotesItem.ClickOnChecksTab();
            qualityChecks.DeliveryAccepted();

            if (qualityChecks.CanClickOnSecurityChecks())
            {
                qualityChecks.CanClickOnSecurityChecks();
                qualityChecks.SetSecurityChecks("Yes");
                qualityChecks.SetQualityChecks();
                receiptNotesItem = qualityChecks.ClickOnReceiptNoteItemTab();
            }
            else
            {
                receiptNotesItem = qualityChecks.ClickOnReceiptNoteItemTab();
            }

            receiptNotesItem.Validate();

            var receiptNotesFooter = receiptNotesItem.ClickOnFooterTab();
            double montantRN = receiptNotesFooter.GetReceiptNoteTotal(currency, decimalSeparatorValue);

            // Calcul du montant de la facture transmise à TL
            var receiptNoteAccounting = receiptNotesFooter.ClickOnAccounting();
            double montantGlobal = receiptNoteAccounting.GetReceiptNoteGrossAmount("G", decimalSeparatorValue);
            double montantDetaille = receiptNoteAccounting.GetReceiptNoteDetailAmount("G", decimalSeparatorValue);

            receiptNotesItem = receiptNoteAccounting.ClickOnItems();
            receiptNotesItem.ExportResultsForSage(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            var correctDownloadedFile = receiptNotesPage.GetExportForSageFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

            // Récupération du nom du fichier et construction de l'URL du fichier txt à exploiter   
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // On n'exploite que les lignes avec contenu "général" --> "G"
            double[] contenuFichier = ExploitTextFiles.VerifySAGEFileContent(filePath, "G", ",");

            receiptNotesPage = receiptNotesItem.BackToList();
            receiptNotesPage.ResetFilter();
            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ByNumber, receiptNoteNumber);

            Assert.IsTrue(receiptNotesPage.IsSentToSageManually(), "La RN n'a pas été envoyée manuellement vers le SAGE.");
            Assert.AreEqual(contenuFichier[0].ToString(), contenuFichier[1].ToString(), "Le contenu du fichier SAGE généré est incorrect.");
            Assert.AreEqual((montantRN).ToString(), contenuFichier[0].ToString(), "Le montant issu du fichier SAGE n'est pas égal au montant de la RN défini dans l'application.");
            Assert.AreEqual(montantGlobal.ToString(), montantDetaille.ToString(), "Les montants AmountDebit et AmountCredit de la RN envoyée vers SAGE ne sont pas les mêmes dans l'onglet Accounting.");
            Assert.AreEqual(montantRN.ToString(), montantGlobal.ToString(), "Le montant issu du fichier SAGE n'est pas égal au montant de la RN défini dans l'application.");
        }

        [Ignore] // aucun pays n'utilise SAGE AUTO
        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_SageAuto_ExportSageItemOK()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();
            string itemName = TestContext.Properties["Item_ReceiptNote"].ToString();
            var currency = TestContext.Properties["Currency"].ToString();

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            // Config pour export SAGE auto mais pas pour le site
            homePage.SetSageAutoEnabled(site, true, "Due to invoice receipt note");

            try
            {
                //Act
                var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
                receiptNotesPage.ResetFilter();

                homePage.ClearDownloads();

                // Create
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.Supplier, supplier);
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ShowValidatedNotInvoiced, true);
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.Site, site);
                ReceiptNotesItem receiptNotesItem = null;

                if (receiptNotesPage.CheckTotalNumber() == 0)
                {
                    var receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
                    receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, deliveryPlace));
                    receiptNotesItem = receiptNotesCreateModalpage.Submit();

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
                }
                else
                {
                    receiptNotesItem = receiptNotesPage.SelectFirstReceiptNoteItem();
                }

                var receiptNotesFooter = receiptNotesItem.ClickOnFooterTab();
                double montantRN = receiptNotesFooter.GetReceiptNoteTotal(currency, decimalSeparatorValue);

                // Calcul du montant de la facture transmise à TL
                var receiptNoteAccounting = receiptNotesFooter.ClickOnAccounting();
                double montantGlobal = receiptNoteAccounting.GetReceiptNoteGrossAmount("G", decimalSeparatorValue);
                double montantDetaille = receiptNoteAccounting.GetReceiptNoteDetailAmount("G", decimalSeparatorValue);

                Assert.AreEqual(montantGlobal.ToString(), montantDetaille.ToString(), "Les montants AmountDebit et AmountCredit de la RN envoyée vers SAGE ne sont pas les mêmes dans l'onglet Accounting.");
                Assert.AreEqual(montantRN.ToString(), montantGlobal.ToString(), "Le montant issu du fichier SAGE n'est pas égal au montant de la RN défini dans l'application.");

            }
            finally
            {
                homePage.SetSageAutoEnabled(site, false);
            }
        }

        [Ignore]
        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_SageManuel_ExportSAGEItemKO_Supplier()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();
            string itemName = TestContext.Properties["Item_ReceiptNote"].ToString();

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Config pour export SAGE manuel
            homePage.SetSageAutoEnabled(site, false);

            try
            {

                // Supplier --> KO pour test
                VerifySupplier(homePage, site, supplier, false);

                //Act
                var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
                receiptNotesPage.ResetFilter();

                // Create
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.Supplier, supplier);
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ShowValidatedNotInvoiced, true);
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.Site, site);
                ReceiptNotesItem receiptNotesItem = null;

                if (receiptNotesPage.CheckTotalNumber() == 0)
                {
                    var receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
                    receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, deliveryPlace));
                    receiptNotesItem = receiptNotesCreateModalpage.Submit();

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
                }
                else
                {
                    receiptNotesItem = receiptNotesPage.SelectFirstReceiptNoteItem();
                }

                // Vérifie qu'un message d'erreur est présent dans l'onglet Accounting
                var supplierInvoiceAccounting = receiptNotesItem.ClickOnAccountingTab();
                string erreur = supplierInvoiceAccounting.GetErrorMessage();

                Assert.AreNotEqual("", erreur, "Aucun message d'erreur n'apparaît alors que le paramétrage relatif au supplier est KO.");
                Assert.IsTrue(erreur.Contains($"Due to invoice account Id for supplier {supplier} is missing"), "Le message d'erreur ne concerne pas le paramétrage Supplier.");

            }
            finally
            {
                VerifySupplier(homePage, site, supplier);
            }
        }

        [Ignore]
        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_SageManuel_ExportSAGEItemKO_CodeJournal()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();
            string itemName = TestContext.Properties["Item_ReceiptNote"].ToString();
            string journalRN = TestContext.Properties["Journal_RN"].ToString();

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
                var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
                receiptNotesPage.ResetFilter();

                // Create
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.Supplier, supplier);
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ShowValidatedNotInvoiced, true);
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.Site, site);
                ReceiptNotesItem receiptNotesItem = null;

                if (receiptNotesPage.CheckTotalNumber() == 0)
                {
                    var receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
                    receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, deliveryPlace));
                    receiptNotesItem = receiptNotesCreateModalpage.Submit();

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
                }
                else
                {
                    receiptNotesItem = receiptNotesPage.SelectFirstReceiptNoteItem();
                }

                // Vérifie qu'un message d'erreur est présent dans l'onglet Accounting
                var supplierInvoiceAccounting = receiptNotesItem.ClickOnAccountingTab();
                string erreur = supplierInvoiceAccounting.GetErrorMessage();

                Assert.AreNotEqual("", erreur, "Aucun message d'erreur n'apparaît alors que le paramétrage relatif au code journal est KO.");
                Assert.IsTrue(erreur.Contains("no Due to invoice journal value set"), "Le message d'erreur ne concerne pas le paramétrage Code journal.");

            }
            finally
            {
                // Parameter - Accounting --> Journal KO pour test
                VerifyAccountingJournal(homePage, site, journalRN);
            }
        }

        [Ignore]
        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_SageManuel_ExportSAGEItemKO_VATPurchaseCode()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();
            string itemName = TestContext.Properties["Item_ReceiptNote"].ToString();

            string taxName = TestContext.Properties["Item_TaxType"].ToString();
            string taxType = "VAT";
            string journalRN = TestContext.Properties["Journal_RN"].ToString();

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            VerifyAccountingJournal(homePage, site, journalRN);

            // Config pour export SAGE manuel
            homePage.SetSageAutoEnabled(site, false);

            try
            {

                // Parameter - Purchasing --> VAT KO pour test
                VerifyPurchasingVAT(homePage, taxName, taxType, false);

                //Act
                var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
                receiptNotesPage.ResetFilter();

                // Create
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.Supplier, supplier);
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ShowValidatedNotInvoiced, true);
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.Site, site);
                ReceiptNotesItem receiptNotesItem = null;

                var receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
                receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, deliveryPlace));
                receiptNotesItem = receiptNotesCreateModalpage.Submit();

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

                //receiptNotesItem = receiptNotesPage.SelectFirstReceiptNoteItem();

                // Vérifie qu'un message d'erreur est présent dans l'onglet Accounting
                var supplierInvoiceAccounting = receiptNotesItem.ClickOnAccountingTab();
                string erreur = supplierInvoiceAccounting.GetErrorMessage();

                Assert.AreNotEqual("", erreur, "Aucun message d'erreur n'apparaît alors que le paramétrage relatif au 'Purchase code' de la VAT est KO.");
                Assert.IsTrue(erreur.Contains($"Due to invoice accounting code of the tax {taxName} for site {site} cannot be empty"), "Le message d'erreur ne concerne pas le paramétrage relatif au 'Purchase code' de la VAT.");

            }
            finally
            {
                VerifyPurchasingVAT(homePage, taxName, taxType);
            }
        }

        [Ignore]
        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_SageManuel_ExportSAGEItemKO_SiteAnalyticalPlan()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();
            string itemName = TestContext.Properties["Item_ReceiptNote"].ToString();
            string journalRN = TestContext.Properties["Journal_RN"].ToString();

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Config pour export SAGE manuel
            homePage.SetSageAutoEnabled(site, false);

            VerifyAccountingJournal(homePage, site, journalRN);

            try
            {

                // Sites -- > Analytical plan et section KO pour test
                VerifySiteAnalyticalPlanSection(homePage, site, false);

                //Act
                var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
                receiptNotesPage.ResetFilter();

                // Create
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.Supplier, supplier);
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ShowValidatedNotInvoiced, true);
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.Site, site);
                ReceiptNotesItem receiptNotesItem = null;

                if (receiptNotesPage.CheckTotalNumber() == 0)
                {
                    var receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
                    receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, deliveryPlace));
                    receiptNotesItem = receiptNotesCreateModalpage.Submit();

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
                }
                else
                {
                    receiptNotesItem = receiptNotesPage.SelectFirstReceiptNoteItem();
                }

                // Vérifie qu'un message d'erreur est présent dans l'onglet Accounting
                var supplierInvoiceAccounting = receiptNotesItem.ClickOnAccountingTab();
                string erreur = supplierInvoiceAccounting.GetErrorMessage();

                Assert.AreNotEqual("", erreur, "Aucun message d'erreur n'apparaît alors que le paramétrage relatif au 'Analytic plan' du site est KO.");
                Assert.IsTrue(erreur.Contains($"Accounting analytic plan of the site {site} cannot be empty"), "Le message d'erreur ne concerne pas le paramétrage relatif au 'Analytic plan' du site.");

            }
            finally
            {
                VerifySiteAnalyticalPlanSection(homePage, site);
            }
        }

        [Ignore]
        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_SageManuel_ExportSAGEItemKO_NoGroupVat()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();
            string itemName = TestContext.Properties["Item_ReceiptNote"].ToString();
            string taxName = TestContext.Properties["Item_TaxType"].ToString();
            string journalRN = TestContext.Properties["Journal_RN"].ToString();

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Config pour export SAGE manuel
            homePage.SetSageAutoEnabled(site, false);

            VerifyAccountingJournal(homePage, site, journalRN);

            // Récupération du groupe de l'item
            string itemGroup = GetItemGroup(homePage, itemName);

            try
            {
                // Parameter - Accounting --> Group & VAT KO pour test
                VerifyGroupAndVAT(homePage, itemGroup, taxName, false);

                //Act
                var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
                receiptNotesPage.ResetFilter();

                // Create
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.Supplier, supplier);
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ShowValidatedNotInvoiced, true);
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.Site, site);
                ReceiptNotesItem receiptNotesItem = null;


                var receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
                receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, deliveryPlace));
                receiptNotesItem = receiptNotesCreateModalpage.Submit();

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

                //receiptNotesItem = receiptNotesPage.SelectFirstReceiptNoteItem();

                // Vérifie qu'un message d'erreur est présent dans l'onglet Accounting
                var supplierInvoiceAccounting = receiptNotesItem.ClickOnAccountingTab();
                string erreur = supplierInvoiceAccounting.GetErrorMessage();

                Assert.AreNotEqual("", erreur, "Aucun message d'erreur n'apparaît alors que le paramétrage relatif au 'Group & VAT' de l'item est KO.");
                Assert.IsTrue(erreur.Contains($"no related ItemGroupTax for ItemGroup {itemGroup} and TaxType {taxName}"), "Le message d'erreur ne concerne pas le paramétrage relatif au 'Group & VAT' de l'item.");
            }
            finally
            {
                VerifyGroupAndVAT(homePage, itemGroup, taxName);
            }
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_Details_GenerateOutputForm()
        {

            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string placeFrom = TestContext.Properties["PlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceToPartielle"].ToString();
            HomePage homePage = LogInAsAdmin();
            ReceiptNotesPage receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();
            ReceiptNotesCreateModalPage receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
            receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, placeTo));
            ReceiptNotesItem receiptNotesItem = receiptNotesCreateModalpage.Submit();
            string receiptNoteNumber = receiptNotesItem.GetReceiptNoteNumber();
            string itemName = receiptNotesItem.GetFirstItemName();
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
            }
            receiptNotesItem = qualityChecks.ClickOnReceiptNoteItemTab();
            receiptNotesItem.ResetFilters();
            receiptNotesItem.Filter(ReceiptNotesItem.FilterItemType.SearchByName, itemName);
            string priceRNText = receiptNotesItem.GetPrice();
            receiptNotesItem.Validate();
            Thread.Sleep(2000);
            ReceiptNoteToOuputForm modalOutputForm = receiptNotesItem.GenerateOutputForm();
            modalOutputForm.Fill(placeFrom, placeTo, true);
            OutputFormItem outputFormItem = modalOutputForm.Create();
            outputFormItem.ResetFilter();
            outputFormItem.Filter(OutputFormItem.FilterItemType.ShowItemsWithPhysQty, false);
            outputFormItem.Filter(OutputFormItem.FilterItemType.SearchByName, itemName);
            string firstName = outputFormItem.VerifFirstName(itemName);
            Assert.AreEqual(itemName, firstName);
            outputFormItem.SelectItem(itemName);
            string priceOFText = outputFormItem.GetPrice(itemName);
            Assert.AreEqual(priceOFText, priceRNText.Substring(0, priceOFText.Length), "les tarifs sont différents");
            OutputFormGeneralInformation outputFormGeneralInfo = outputFormItem.ClickOnGeneralInformationTab();
            Assert.AreEqual(site, outputFormGeneralInfo.GetSite(), "mauvais site");
            Assert.AreEqual(placeFrom, outputFormGeneralInfo.GetPlaceFrom(), "mauvais placeFrom");
            Assert.AreEqual(placeTo, outputFormGeneralInfo.GetPlaceTo(), "mauvais placeTo");
        }

        [DeploymentItem("Resources\\pizza.png")]
        [DeploymentItem("Resources\\pizza_petite.png")]
        [DeploymentItem("chromedriver.exe")]
        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_Items_MultipleChecks()
        {
            List<string> listDetectedErrors = new List<string>();
            HomePage homePage = LogInAsAdmin();

            string decimalSeparator = homePage.GetDecimalSeparatorValue();

            ReceiptNotesItem receiptNotesItem = this.CreateOneRN(homePage);

            this.WA_RN_Items_UpdateReceivedQuantity(receiptNotesItem, decimalSeparator, listDetectedErrors);

            this.WA_RN_Claims_EditClaim(receiptNotesItem, decimalSeparator, listDetectedErrors);
            //After filling the claim, we're still in the claim tab so click back to Item tab
            receiptNotesItem = receiptNotesItem.ClickOnItemsTab();

            this.WA_RN_Items_AddTemperatureAndDuplicate(receiptNotesItem, listDetectedErrors);

            this.WA_RN_Items_VerifyTotalDN(receiptNotesItem, decimalSeparator, listDetectedErrors);

            this.WA_RN_Items_ItemExpiryDate(receiptNotesItem, listDetectedErrors);

            //La suppression d'item doit idéalement intervenir une fois que l'ensemble des statuts a été vérifiée.
            //Ainsi, on peut vérifier que la couleur des statut repasse à "non renseigné".
            this.WA_RN_Items_DeleteItemDetails(receiptNotesItem, listDetectedErrors);

            this.WA_RN_Items_EditItem(receiptNotesItem, listDetectedErrors);

            this.WA_RN_Footer_Verify(receiptNotesItem, decimalSeparator, listDetectedErrors);

            if (listDetectedErrors.Any())
                Assert.IsTrue(false, string.Join(Environment.NewLine, listDetectedErrors));
        }

        private void WA_RN_Claims_EditClaim(ReceiptNotesItem receiptNotesItem, string decimalSeparator, List<string> errors)
        {
            try
            {
                //On suppose qu'il existe déjà un item
                receiptNotesItem.ResetFilters();
                receiptNotesItem.SelectFirstItem();

                //Ajout de la claim
                ReceiptNotesEditClaim editClaim = receiptNotesItem.AddClaim();
                editClaim.Fill();

                //On arrive sur l'onglet claim
                ReceiptNotesClaims receiptNotesClaims = editClaim.Save();
                //5) Cliquer sur edit
                editClaim = receiptNotesClaims.EditClaim();

                editClaim.SetDecreaseStock("5");
                editClaim.UploadPicture("pizza.png");
                editClaim.SetComment("comment");
                editClaim.SetSanctionAmount("55,5");
                editClaim.SetVAT("2-General");
                receiptNotesClaims = editClaim.Save();

                //Vérifier que les informations sont bien modifiées

                if (receiptNotesClaims.GetDecrStock() == false)
                { errors.Add("Decrease stock non coché."); }

                if (receiptNotesClaims.GetDecrQty() != "5")
                { errors.Add("Decrease quantity non rempli."); }

                if (receiptNotesItem.FormCommentGreen() == false)
                { errors.Add("Commentaire dans l'onglet 'claim' non renseigné en vert."); }

                if (receiptNotesItem.FormPictureGreen() == false)
                { errors.Add("Icone dans l'onglet 'claim' non vert"); }
            }
            catch (Exception ex)
            {
                errors.Add(ex.Message);
            }
        }

        private ReceiptNotesItem CreateOneRN(HomePage homePage)
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string placeTo = TestContext.Properties["PlaceToPartielle"].ToString();

            //Act         
            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();

            //1. Entrer dans une RN non validée
            var receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
            receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, placeTo));
            return receiptNotesCreateModalpage.Submit();
        }

        /// <summary>
        /// Checks the update of received quantites and the associates calculations
        /// </summary>
        /// <param name="receiptNotesItem">The RN detail page</param>
        /// <param name="decimalSeparator">The decimal separator</param>
        /// <param name="errors">Detected errors</param>
        private void WA_RN_Items_UpdateReceivedQuantity(ReceiptNotesItem receiptNotesItem, string decimalSeparator, List<string> errors)
        {
            try
            {
                receiptNotesItem.ResetFilters();
                string itemName = receiptNotesItem.extractItemName(receiptNotesItem.GetFirstItemName());

                receiptNotesItem.SelectFirstItem();

                //Sur l'onglet items, rentrer dans la colonne "Received" les quantités
                receiptNotesItem.AddReceived(itemName, "2");
                receiptNotesItem.WaitPageLoading();

                //Vérifier que les données calculer se mettent à jour et soient ok
                double intDNPrice = receiptNotesItem.GetDNPrice(decimalSeparator);
                double intDNQty = receiptNotesItem.GetDNQty(decimalSeparator);
                double intDNTotal = receiptNotesItem.GetDNTotal(decimalSeparator);

                double roundedQty = Math.Round(intDNQty, 4);
                if (roundedQty != 2)
                { errors.Add($"Cas 1 - Mauvais arrondi de chiffre : l'attendu est 2 mais on trouve {roundedQty}"); }

                double calculatedTotal = Math.Round(intDNQty * intDNPrice, 4);
                if (calculatedTotal != intDNTotal)
                { errors.Add($"Cas 1 - Mauvais calcul ou arrondi du total : l'attendu est {calculatedTotal} mais on trouve {intDNTotal}"); }

                //Ajout d'un second item
                string itemName2 = receiptNotesItem.extractItemName(receiptNotesItem.GetItemName(2));
                receiptNotesItem.SelectItemRow(2);
                receiptNotesItem.AddReceived(itemName2, "33");

                //Afficher les items reçus at modifier l'item
                receiptNotesItem.Filter(ReceiptNotesItem.FilterItemType.ShowNotReceived, false);
                receiptNotesItem.WaitPageLoading();

                string currentFirstItemName = receiptNotesItem.extractItemName(receiptNotesItem.GetFirstItemName());
                receiptNotesItem.SelectFirstItem();
                receiptNotesItem.AddReceived(currentFirstItemName, "5");

                //Vérifier que les données calculer se mettent à jour et soient ok
                intDNPrice = receiptNotesItem.GetDNPrice(decimalSeparator);
                intDNQty = receiptNotesItem.GetDNQty(decimalSeparator);
                intDNTotal = receiptNotesItem.GetDNTotal(decimalSeparator);

                roundedQty = Math.Round(intDNQty, 4);
                if (roundedQty != 5)
                { errors.Add($"Cas 2 - Mauvais arrondi de chiffre : l'attendu est 5 mais on trouve {roundedQty}"); }

                calculatedTotal = Math.Round(intDNQty * intDNPrice, 4);
                if (calculatedTotal != intDNTotal)
                { errors.Add($"Cas 2 - Mauvais calcul ou arrondi du total : l'attendu est {calculatedTotal} mais on trouve {intDNTotal}"); }
            }
            catch (Exception ex)
            {
                errors.Add("Erreur dans WA_RN_Items_UpdateReceivedQuantity : " + ex.Message);
            }
        }

        /// <summary>
        /// Checks the change of temperature
        /// </summary>
        /// <param name="receiptNotesItem">The RN detail page</param>
        /// <param name="errors">Detected errors</param>
        private void WA_RN_Items_AddTemperatureAndDuplicate(ReceiptNotesItem receiptNotesItem, List<string> errors)
        {
            try
            {
                receiptNotesItem.ResetFilters();
                receiptNotesItem.SelectFirstItem();

                if (receiptNotesItem.TemperatureGreen())
                { errors.Add("On a une icône verte de présence T° alors qu'elle n'a pas encore été ajoutée."); }
                else
                {
                    receiptNotesItem.TemperatureIconClick(-5);

                    if (receiptNotesItem.TemperatureGreen() == false)
                    { errors.Add("La T° a été renseignée mais l'icône n'est pas verte."); }

                    receiptNotesItem.TemperatureIconClick(-5, true);

                    string groupName = receiptNotesItem.GetFirstGroupName();
                    receiptNotesItem.Filter(ReceiptNotesItem.FilterItemType.ByGroup, groupName);

                    if (receiptNotesItem.TemperatureAllGreen() == false)
                    { errors.Add("Les icônes du groupe d'item ne sont pas toutes vertes pour signaler une T° renseignée."); }
                }
            }
            catch (Exception ex)
            {
                errors.Add("Erreur dans WA_RN_Items_AddTemperatureAndDuplicate : " + ex.Message);
            }
        }

        /// <summary>
        /// Checks the calculations of DN total
        /// </summary>
        /// <param name="receiptNotesItem">The RN detail page</param>
        /// <param name="decimalSeparator">The decimal separator</param>
        /// <param name="errors">Detected errors</param>
        private void WA_RN_Items_VerifyTotalDN(ReceiptNotesItem receiptNotesItem, string decimalSeparator, List<string> errors)
        {
            try
            {
                receiptNotesItem.ResetFilters();

                string itemName1 = receiptNotesItem.GetFirstItemName();
                string itemName2 = receiptNotesItem.GetItemName(2);

                receiptNotesItem.Filter(ReceiptNotesItem.FilterItemType.SearchByName, itemName1);
                receiptNotesItem.SelectItem(itemName1);
                receiptNotesItem.AddReceived(itemName1, "10");
                double DNPrice1 = receiptNotesItem.GetDNPrice(decimalSeparator);
                double DNQty1 = receiptNotesItem.GetDNQty(decimalSeparator);
                double DNTotal1 = receiptNotesItem.GetDNTotal(decimalSeparator);

                double calculatedDNTotal1 = Math.Round(DNPrice1 * DNQty1, 4);
                if (calculatedDNTotal1 != DNTotal1)
                { errors.Add($"Cas 1 - Mauvais calcul ou arrondi du total : l'attendu est {calculatedDNTotal1} mais on trouve {DNTotal1}"); }

                receiptNotesItem.Filter(ReceiptNotesItem.FilterItemType.SearchByName, itemName2);
                receiptNotesItem.SelectItem(itemName2);
                receiptNotesItem.AddReceived(itemName2, "12");
                double DNPrice2 = receiptNotesItem.GetDNPrice(decimalSeparator);
                double DNQty2 = receiptNotesItem.GetDNQty(decimalSeparator);
                double DNTotal2 = receiptNotesItem.GetDNTotal(decimalSeparator);

                double calculatedDNTotal2 = Math.Round(DNPrice2 * DNQty2, 4);
                Assert.AreEqual(calculatedDNTotal2, DNTotal2, 0.0001, $"Cas 2 - Mauvais calcul ou arrondi du total : l'attendu est {calculatedDNTotal2} mais on trouve {DNTotal2}");
                //if (calculatedDNTotal2 != DNTotal2)
                //{ errors.Add($"Cas 2 - Mauvais calcul ou arrondi du total : l'attendu est {calculatedDNTotal2} mais on trouve {DNTotal2}"); }

                receiptNotesItem.ResetFilters();
                double DNTotalSum = receiptNotesItem.GetDNTotalSum(decimalSeparator);
                double calculatedDNTotalSum = Math.Round(calculatedDNTotal1 + calculatedDNTotal2, 4);

                if (calculatedDNTotalSum != DNTotalSum)
                { errors.Add($"Cas 3 - Mauvais calcul ou arrondi du total de tous les DN : l'attendu est {calculatedDNTotalSum} mais on trouve {DNTotalSum}"); }
            }
            catch (Exception ex)
            {
                errors.Add("Erreur dans WA_RN_Items_VerifyTotalDN : " + ex.Message);
            }
        }

        /// <summary>
        /// Checks expiry date
        /// </summary>
        /// <param name="receiptNotesItem">The RN detail page</param>
        /// <param name="errors">Detected errors</param>
        private void WA_RN_Items_ItemExpiryDate(ReceiptNotesItem receiptNotesItem, List<string> errors)
        {
            try
            {
                DateTime expiryDateTest = DateUtils.Now.AddDays(15);
                receiptNotesItem.SelectFirstItem();

                ReceiptNoteExpiry expiryDate = receiptNotesItem.ShowFirstExpiryDate();
                //Thread.Sleep(2000);
                expiryDate.ModifyFirstExpiryDate("3", expiryDateTest);
                expiryDate.AddExpiryDate("2", expiryDateTest.AddDays(2));
                receiptNotesItem = expiryDate.Save();

                if (receiptNotesItem.FormExpiryGreen() == false)
                { errors.Add("Absence d'icone verte pour l'expiration de date"); }

                receiptNotesItem.Refresh();
                receiptNotesItem.SelectFirstItem();

                //Tester la suppression
                expiryDate = receiptNotesItem.ShowFirstExpiryDate();
                //Thread.Sleep(2000);
                expiryDate.DeleteAllExpiryDate();
                expiryDate.Save();

                if (receiptNotesItem.FormExpiryGreen())
                { errors.Add("Icone verte pour l'expiration de date alors que toutes les dates ont été supprimées"); }
            }
            catch (Exception ex)
            {
                errors.Add("Erreur dans WA_RN_Items_ItemExpiryDate : " + ex.Message);
            }
        }

        private void WA_RN_Items_DeleteItemDetails(ReceiptNotesItem receiptNotesItem, List<string> errors)
        {
            try
            {
            receiptNotesItem.ResetFilters();
                string itemName = receiptNotesItem.GetFirstItemName();
                receiptNotesItem.SelectFirstItem();

                //3. Cliquer sur ... puis sur l'icone poubelle
                receiptNotesItem.DeleteItem();
                receiptNotesItem.WaitPageLoading();

                receiptNotesItem.Filter(ReceiptNotesItem.FilterItemType.ShowNotReceived, true);
                receiptNotesItem.Filter(ReceiptNotesItem.FilterItemType.SearchByName, itemName);
                receiptNotesItem.WaitPageLoading();

                receiptNotesItem.SelectFirstItem();

                if (receiptNotesItem.FormClaimGreen())
                { errors.Add("L'item supprimé a conservé son marqueur de claim renseigné."); }

                if (receiptNotesItem.FormTemperatureGreen())
                { errors.Add("L'item supprimé a conservé son marqueur de T° renseigné."); }

                if (receiptNotesItem.FormClaimGreen())
                { errors.Add("L'item supprimé a conservé son marqueur de date d'expiration renseigné."); }
            }
            catch (Exception ex)
            {
                errors.Add("Erreur dans WA_RN_Items_DeleteItemDetails : " + ex.Message);
            }
        }

        private void WA_RN_Items_EditItem(ReceiptNotesItem receiptNotesItem, List<string> errors)
        {
            try
            {
                receiptNotesItem.ResetFilters();

                // PLATO PLASTICO NEUTRAL 150x90 (PLG00014)
                string itemName = receiptNotesItem.extractItemName(receiptNotesItem.GetFirstItemName());
                receiptNotesItem.SelectFirstItem();
                ItemGeneralInformationPage itemPage = receiptNotesItem.EditItem(itemName);

                //Vérifier que l'on ouvre le bon item
                string itemNameGeneralInfo = itemPage.GetItemName();
                if (itemName.Equals(itemNameGeneralInfo) == false)
                { errors.Add("L'édition d'item dans le détail RN n'atterit pas sur la bonne page d'item."); }

                itemPage.Close();
            }
            catch (Exception ex)
            {
                errors.Add("Erreur dans WA_RN_Items_EditItem : " + ex.Message);
            }
        }

        private void WA_RN_Footer_Verify(ReceiptNotesItem receiptNotesItem, string decimalSeparator, List<string> errors)
        {
            try
            {
                receiptNotesItem.ResetFilters();
                string itemName = receiptNotesItem.GetFirstItemName();
                receiptNotesItem.SelectFirstItem();
                double DNTotal = receiptNotesItem.GetDNTotal(decimalSeparator);

                //Récupérer le nom de la taxe
                ItemGeneralInformationPage itemGeneralInfo = receiptNotesItem.EditItem(itemName);
                string itemTaxName = itemGeneralInfo.GetVatName();
                itemGeneralInfo.Close();

                //Aller sur l'onglet footer
                ReceiptNotesFooterPage footer = receiptNotesItem.ClickOnFooterTab();
                string footerTaxName = footer.GetTaxName();
                double taxBaseAmount = footer.GetTaxBaseAmount(decimalSeparator);
                double VATRate = footer.GetVATRate(decimalSeparator);
                double VATAmount = footer.GetVATAmount(decimalSeparator);
                double TotalReceiptNote = footer.GetTotalReceiptNote(decimalSeparator);

                if (itemTaxName.Equals(footerTaxName) == false)
                { errors.Add($"Le nom de la taxe dans le footer '{footerTaxName}' ne correspond pas avec celui de l'item '{itemTaxName}'."); }

                double calculatedVatAmount = Math.Round((VATRate / 100) * taxBaseAmount, 4);
                if (calculatedVatAmount != VATAmount)
                { errors.Add($"Le calcul de la taxe attendue est {calculatedVatAmount} mais celui montrée est {VATAmount}"); }

                if (DNTotal != taxBaseAmount)
                { errors.Add("Le total DN n'est pas correct sur le footer."); }

                if (TotalReceiptNote != Math.Round(calculatedVatAmount + taxBaseAmount, 4))
                { errors.Add("Le total du footer n'est pas correct."); }
            }
            catch (Exception ex)
            {
                errors.Add("Erreur dans WA_RN_Footer_Verify : " + ex.Message);
            }
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_Claims_DeleteClaim()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string placeTo = TestContext.Properties["PlaceToPartielle"].ToString();

            // Log in
            var homePage = LogInAsAdmin();

            //Act         
            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();

            //1) Créer une RN
            //Cliquer sur "new Receipt Note"
            var receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
            //Remplir le formulaire
            receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, placeTo));
            //Cliquer sur "Create"
            ReceiptNotesItem receiptNotesItem = receiptNotesCreateModalpage.Submit();

            //2) Ajouter un item
            string itemName = receiptNotesItem.GetFirstItemName();
            receiptNotesItem.SelectFirstItem();
            receiptNotesItem.AddReceived(itemName, "2");

            //3) Ajouter une claim(icone mégaphone), remplir les champs et save
            ReceiptNotesEditClaim editClaim = receiptNotesItem.AddClaim();
            editClaim.Fill();
            //4) On arrive sur l'onglet claim (des RN)
            ReceiptNotesClaims receiptNotesClaims = editClaim.Save();
            Assert.AreEqual(1, receiptNotesClaims.CheckTableSize(), "mauvais nombre de Claim");
            receiptNotesClaims.DeleteClaim();
            Assert.AreEqual(0, receiptNotesClaims.CheckTableSize(), "Claim existant");
            receiptNotesItem = receiptNotesClaims.ClickOnItemsTab();
            receiptNotesItem.SelectFirstItem();
            Assert.IsFalse(receiptNotesItem.FormMegaphoneGreen(), "claim verte");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_Items_Filters_BySubGroup()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string itemName = "itemSubGroupReceiptNotes_" + site + "_2";

            string group = TestContext.Properties["Item_Group"].ToString();
            string subgroup = "SubGroupRN";
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxName = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string deliveryPlace = "Economato";

            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            try
            {
                //	Activation du filtre
                // Aller dans les Global Setting
                // Cocher IsSubGroupFunctionActive
                homePage.SetSubGroupFunctionValue(true);

                //Vérifier si existe ou créer un sous group : Parameters > Production > SubGroup
                ParametersProduction productionPage = homePage.GoToParameters_ProductionPage();
                productionPage.GoToTab_SubGroup();
                var tailleTableau = WebDriver.FindElements(By.XPath("//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[2]"));
                bool trouve = false;
                foreach (var ligne in tailleTableau)
                {
                    if (ligne.Text == subgroup)
                    {
                        trouve = true;
                    }
                }
                if (!trouve)
                {
                    var add = productionPage.WaitForElementIsVisible(By.XPath("//*[@id=\"tabContentParameters\"]/div[1]/a"));
                    add.Click();
                    var code = productionPage.WaitForElementIsVisible(By.Id("first"));
                    code.SetValue(PageBase.ControlType.TextBox, subgroup);
                    var name = productionPage.WaitForElementIsVisible(By.Id("Name"));
                    name.SetValue(PageBase.ControlType.TextBox, subgroup);
                    var groupSelect = productionPage.WaitForElementIsVisible(By.XPath("//*[@id=\"ItemSubGroupModal\"]/div/div/div/div/form/div[2]/div[3]/div/select"));
                    groupSelect.SetValue(PageBase.ControlType.DropDownList, group);
                    var save = productionPage.WaitForElementIsVisible(By.Id("btnSave"));
                    save.Click();
                }
                homePage.Navigate();

                // Vérifier si existe ou créer un item avec le sous group crée 
                ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
                itemPage.ResetFilter();
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

                //1.Aller sur le detail d'une RN non validé
                var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
                // créer une RN avec le bon Supplier
                ReceiptNotesCreateModalPage newReceiptNote = receiptNotesPage.ReceiptNotesCreatePage();
                newReceiptNote.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now.Date, site, supplier, deliveryPlace));
                //2.Onglet items
                ReceiptNotesItem receiptNote = newReceiptNote.Submit();

                //3.Appliquer le filtre Sub Group
                receiptNote.ResetFilters();
                receiptNote.Filter(ReceiptNotesItem.FilterItemType.SearchByName, itemName);
                receiptNote.Filter(ReceiptNotesItem.FilterItemType.ByGroup, group);
                receiptNote.Filter(ReceiptNotesItem.FilterItemType.BySubGroup, subgroup);
                //les résultats s'accordent avec le filtre
                Assert.AreEqual(itemName, receiptNote.GetFirstItemName(), "pas de résultat");

                //Assert.IsTrue(receiptNote.CheckTotalNumber() > 0, "Item " + itemName + " non trouvé via SubGroup " + subgroup);
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

        [Ignore]
		[TestMethod]
        [Timeout(_timeout)]
        public void WA_RN_Filter_Show_SentToSageAndInErrorOnly()
        {
            // Log in
            var homePage = LogInAsAdmin();

            //Act         
            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            //Effectuer des recherches via les filtres
            receiptNotesPage.ResetFilter();

            //Appliquer les filtres sur
            //Show sent to SAGE and in error only
            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ShowSentToSAGEAndInErrorOnly, true);
            // patch update ReceiptNotes set SageState=-1 where NumberString='139928'
            // dev   update ReceiptNotes set SageState=-1 where NumberString='140292'
            // + Validate (sinon pas de disquette rouge)
            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.DateFrom, new DateTime(2022, 7, 1));
            //Vérifier que les résultats s'accordent bien au filtre appliqué
            Assert.IsTrue(receiptNotesPage.CheckTotalNumber() > 0, "Pas de SentToSAGE & InErrorOnly");
            //disquette rouge, envoyé à TL, poussé à cegid, en error

            var disquetteRouge = receiptNotesPage.isDisquetteRouge();
            var warningDisquetteRouge = receiptNotesPage.isWarningDisquetteRouge();
            Assert.IsTrue(disquetteRouge , "La disquette n'a pas un couleur Rouge");
            Assert.AreEqual(warningDisquetteRouge, "Warning : sage piece have been sent to WinrestTL but return in error after sage push with setting ExportSageToETL activated");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_CreateReceiptNotesFromTwoPurchaseOrders()
        {
            // Prepare
            string itemName1 = "CHOCOLATE";
            string itemName2 = "CARAMEL";
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            var location = TestContext.Properties["Location"].ToString();

            // Log in
            var homePage = LogInAsAdmin();

            var decimalSeparator = homePage.GetDecimalSeparatorValue();
            //Act         
            //Créer 2 PO
            PurchaseOrdersPage poPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            CreatePurchaseOrderModalPage createPO = poPage.CreateNewPurchaseOrder();
            createPO.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(10), true);
            PurchaseOrderItem poItems = createPO.Submit();
            poItems.Filter(PurchaseOrderItem.FilterItemType.ByName, itemName1);
            poItems.SelectFirstItemPo();
            poItems.AddQuantity("2");
            if (poItems.IsDev())
            {
                poItems.Approve();
            }
            poItems.Validate();
            PurchaseOrderGeneralInformation generalInfo = poItems.ClickOnGeneralInformation();
            string poNumber1 = generalInfo.getPurchaseOrderNumber();
            poPage = generalInfo.BackToList();

            createPO = poPage.CreateNewPurchaseOrder();
            createPO.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(10), true);
            poItems = createPO.Submit();
            poItems.Filter(PurchaseOrderItem.FilterItemType.ByName, itemName2);
            poItems.SelectFirstItemPo();
            poItems.AddQuantity("2");
            if (poItems.IsDev())
            {
                poItems.Approve();
            }
            poItems.Validate();
            generalInfo = poItems.ClickOnGeneralInformation();
            string poNumber2 = generalInfo.getPurchaseOrderNumber();
            poPage = generalInfo.BackToList();
            homePage.Navigate();

            //1.Cliquer sur "new Receipt Note"
            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            ReceiptNotesCreateModalPage createPage = receiptNotesPage.ReceiptNotesCreatePage();
            //2.Remplir le formulaire et choisir les 2 PO
            createPage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now.AddDays(10), site, supplier, deliveryPlace) { CreateFromPO = true, PONumber = poNumber1, PODate = DateUtils.Now.AddDays(10) });
            createPage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now.AddDays(10), site, supplier, deliveryPlace) { CreateFromPO = true, PONumber = poNumber2, PODate = DateUtils.Now.AddDays(10) });
            //3.Cliquer sur "Create" Vérifier que les 2 OF soient bien prises en compte
            ReceiptNotesItem items = createPage.Submit();
            ReceiptNotesGeneralInformation generalInfoRN = items.ClickOnGeneralInformationTab();
            Assert.AreEqual(poNumber1 + " " + poNumber2, generalInfoRN.GetPurchaseOrderNumbers(), "on n'a pas 2 OF");
            //La liste des articles ainsi que les quantités sont les mêmes que ceux du PO
            items = generalInfoRN.ClickOnItemsTab();
            items.Filter(ReceiptNotesItem.FilterItemType.SearchByName, itemName1);
            Assert.IsTrue(items.GetFirstItemName().Contains(itemName1));
            items.SelectFirstItem();
            Assert.AreEqual(2.0, items.GetDNQty(decimalSeparator), "Mauvaise quantité");
            items.Filter(ReceiptNotesItem.FilterItemType.SearchByName, itemName2);
            Assert.IsTrue(items.GetFirstItemName().Contains(itemName2));
            items.SelectFirstItem();
            Assert.AreEqual(2.0, items.GetDNQty(decimalSeparator), "Mauvaise quantité - cas 2");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_Detail_Items_Filters_ShowItemsNotReceived()
        {
            //Prepare
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();

            // Log in
            var homePage = LogInAsAdmin(); 

            //Act
            //Etre sur le detail d'une RN
            //1. Etre sur une RN non validée
            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            ReceiptNotesCreateModalPage createPage = receiptNotesPage.ReceiptNotesCreatePage();
            createPage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, deliveryPlace));
            //2.Onglet items
            ReceiptNotesItem receiptNoteItems = createPage.Submit();
            //3.Appliquer le filtre Show items not received
            receiptNoteItems.Filter(ReceiptNotesItem.FilterItemType.ShowNotReceived, true);
            //les données sont filtrées
            string itemName = receiptNoteItems.GetFirstItemName();

            receiptNoteItems.SetFirstReceivedQuantity("2");

            receiptNoteItems.Filter(ReceiptNotesItem.FilterItemType.ShowNotReceived, false);
            Assert.AreEqual(receiptNoteItems.GetFirstItemName(), itemName);

            receiptNoteItems.Filter(ReceiptNotesItem.FilterItemType.ShowNotReceived, true);
            Assert.AreEqual(receiptNoteItems.GetItem(itemName), itemName);

            receiptNoteItems.SetFirstReceivedQuantity("0");

            receiptNoteItems.Filter(ReceiptNotesItem.FilterItemType.ShowNotReceived, false);
            //Ticket 12294 removed (low)
            //var itemInList = WebDriver.FindElement(By.XPath("//*[@id='itemForm_0']/div[2]/div[3]/span[contains(text(),'" + itemName + "')]"));
            //Assert.IsNull(itemInList, "item présent avec ShowNotReceived=false et item.Received=0");

            receiptNoteItems.Filter(ReceiptNotesItem.FilterItemType.ShowNotReceived, true);
            Assert.AreEqual(receiptNoteItems.GetItem(itemName), itemName);
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_CloseReceiptNotes()
        {
            //Prepare
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string rnNumber = string.Empty;
            // Log in
            var homePage = LogInAsAdmin();
            // Create New Receipt Note
            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();
            try
            {
                ReceiptNotesCreateModalPage createPage = receiptNotesPage.ReceiptNotesCreatePage();
                createPage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, deliveryPlace));
                ReceiptNotesItem receiptNoteItems = createPage.Submit();
                rnNumber = receiptNoteItems.GetReceiptNoteNumber();

                //Etre sur l'index des RN et Avoir des données disponible sur l'index des RN
                ReceiptNotesPage rnPage = receiptNoteItems.BackToList();
                rnPage.ResetFilter();
                rnPage.Filter(ReceiptNotesPage.FilterType.ByNumber, rnNumber);
                rnPage.Filter(ReceiptNotesPage.FilterType.Opened, true);
                Assert.AreEqual(1, rnPage.CheckTotalNumber());
                rnPage.ResetFilter();
                //1.Survoler les...
                //2. Choisir la RN a clôturé sur la pop'up afficher
                rnPage.CloseRNNumber(rnNumber);

                rnPage.ResetFilter();
                rnPage.Filter(ReceiptNotesPage.FilterType.ByNumber, rnNumber);
                rnPage.Filter(ReceiptNotesPage.FilterType.Closed, true);

                receiptNoteItems = rnPage.SelectFirstReceiptNoteItem();
                ReceiptNotesGeneralInformation generalInfo = receiptNoteItems.ClickOnGeneralInformationTab();
                string rnNumberGeneralInfo = generalInfo.GetReceiptNoteNumber();
                Assert.AreEqual(rnNumber, rnNumberGeneralInfo);
                string status = generalInfo.GetStatus();
                Assert.AreEqual("Closed", status, "Status non Closed");
                generalInfo.BackToList();
            }
            finally
            {
                // Filter Receipt Note Created and Delete
                receiptNotesPage.ResetFilter();
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ByNumber, rnNumber);
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.Closed, true);
                receiptNotesPage.DeleteReceiptNote();

            }
            
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_Backdate_ReceiptNote_Creation()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();

            // Log in
            var homePage = LogInAsAdmin();

            string oldBackDating = null; // 31
            // Le but est de tester que le setting Backdate et de voir les messages d'erreurs apparaître les possibles régressions.
            ApplicationSettingsPage applicationSettingPage = null;
            applicationSettingPage = homePage.GoToApplicationSettings();
            try
            {
                //Attention, le temps du test seulement, mettre le setting Backdate à 3. (à travers global settings)
                var appSettingsModalPage = applicationSettingPage.GetBackDatingPage();
                oldBackDating = appSettingsModalPage.GetBackDating();
                appSettingsModalPage.SetBackDating("10");
                appSettingsModalPage.Save();

                //Act           
                var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
                receiptNotesPage.ResetFilter();

                // Create
                var receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
                //0  => -1 KO
                //1  => -2 KO
                //10 => -11 KO
                receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now.AddDays(-11), site, supplier, deliveryPlace));
                var receiptNotesItem = receiptNotesCreateModalpage.Submit();
                Assert.IsTrue(receiptNotesCreateModalpage.isBackDatingValidateur(), "pas d'erreur BackDating");
                String receiptNoteNumber = receiptNotesItem.GetReceiptNoteNumber();

            }
            finally
            {
                //A la fin du test remettre le setting sur la valeur antérieur.
                applicationSettingPage = homePage.GoToApplicationSettings();
                var appSettingsModalPage = applicationSettingPage.GetBackDatingPage();
                // 3 sur testauto4dev
                appSettingsModalPage.SetBackDating(oldBackDating);
                appSettingsModalPage.Save();
            }
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_Backdate_ReceiptNote_GlobalSettings_Setting()
        {
            // Prepare
            string itemName1 = "CHOCOLATE";
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string location = TestContext.Properties["Location"].ToString();
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();
            string commentaire = "texte commentaire";
            string commentaire2 = "testcommentaire";
            string placeTo = TestContext.Properties["PlaceToPartielle"].ToString();
            string deliveryNumber2 = new Random().Next(1000000, 9999999).ToString();

            // Log in
            var homePage = LogInAsAdmin();

            string oldBackDating = null; // 31
            // Le but est de tester que le setting Backdate et de voir les messages d'erreurs apparaître les possibles régressions.
            ApplicationSettingsPage applicationSettingPage = null;
            applicationSettingPage = homePage.GoToApplicationSettings();

            try
            {
                //Avoir le temps du test le setting sur 3 le temps du test.
                var appSettingsModalPage = applicationSettingPage.GetBackDatingPage();
                oldBackDating = appSettingsModalPage.GetBackDating();
                appSettingsModalPage.SetBackDating("10");
                appSettingsModalPage.Save();
                //Act

                // Avoir une PO validée pour le supplier et le site place avec des items actifs
                PurchaseOrdersPage poPage = homePage.GoToPurchasing_PurchaseOrdersPage();
                CreatePurchaseOrderModalPage createPO = poPage.CreateNewPurchaseOrder();
                createPO.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now, true);
                PurchaseOrderItem poItems = createPO.Submit();
                string poNumber = poItems.GetPurchaseOrderNumber();
                poItems.Filter(PurchaseOrderItem.FilterItemType.ByName, itemName1);
                poItems.SelectFirstItemPo();
                poItems.AddQuantity("2");
                if (poItems.IsDev())
                {
                    poItems.Approve();
                }
                poItems.Validate();

                //Créer une RN depuis l'onglet RN pour le supplier, la delivery place, le delivery order "132465987",
                //la date à J-2, mettre un commentaire "texte commentaire" , choisir created from et mettre la PO.
                ReceiptNotesPage rnPage = homePage.GoToWarehouse_ReceiptNotesPage();
                ReceiptNotesCreateModalPage rnCreate = rnPage.ReceiptNotesCreatePage();
                //0  => -1 KO
                //1  => -2 KO
                //10 => -11 KO
                rnCreate.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now.AddDays(-10), site, supplier, deliveryPlace) { CreateFromPO = true, PONumber = poNumber, PODate = DateUtils.Now.AddDays(-2), Commentary = commentaire });
                string rnDelivery = rnCreate.GetDeliveryOrderNumber();
                ReceiptNotesItem receiptNotesItem = rnCreate.Submit();
                string numberToCopy = receiptNotesItem.GetReceiptNoteNumber();

                //Ouvrir la RN,
                ReceiptNotesGeneralInformation rnGeneralInfo = receiptNotesItem.ClickOnGeneralInformationTab();
                //modifier la delivery Order à "987654",
                rnGeneralInfo.SetReceiptNoteDeliveryOrderNumber(deliveryNumber2);
                //changer la delivery date à J-1
                rnGeneralInfo.SetDeliveryDate(DateUtils.Now.AddDays(-1));
                //ajouter un commentaire "testcommentaire",
                rnGeneralInfo.SetComment(commentaire2);
                rnGeneralInfo.PageUp();

                receiptNotesItem = rnGeneralInfo.ClickOnItemsTab();
                rnGeneralInfo = receiptNotesItem.ClickOnGeneralInformationTab();
                Assert.AreEqual(commentaire2, rnGeneralInfo.GetComment());

                // changer à l'onglet item,
                receiptNotesItem = rnGeneralInfo.ClickOnItemsTab();
                // Mettre les Physical quantity égale à la PO et DN
                receiptNotesItem.Filter(ReceiptNotesItem.FilterItemType.SearchByName, itemName1);
                receiptNotesItem.SetFirstReceivedQuantity("2");
                // Remplir les checks
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

                receiptNotesItem = qualityChecks.ClickOnReceiptNoteItemTab();
                //FIXME si on crée une Claim, alors la Receipt Note doit être validé
                receiptNotesItem.Validate();

                // Ajouter une claim
                ClaimsPage claim = homePage.GoToWarehouse_ClaimsPage();
                ClaimsCreateModalPage claimModal = claim.ClaimsCreatePage();
                claimModal.FillField_CreatNewClaims(DateUtils.Now, site, supplier, placeTo, true, "Receipt Note", numberToCopy);
                ClaimsItem claimItem = claimModal.Submit();
                // Changer de sous onglet(exemple item puis claim puis re -general information)
                ClaimsGeneralInformation claimGeneralInfo = claimItem.ClickOnGeneralInformation();
                claimItem = claimGeneralInfo.ClickOnItems();
                claimGeneralInfo = claimItem.ClickOnGeneralInformation();
                // Ne pas valider (FIXME)

                // Dans la RN non validée, bien avoir la delivery order à "987654"
                Assert.AreEqual(deliveryNumber2, claimGeneralInfo.GetRelatedReceiptNoteDeliveryOrderNumber(), "Mauvais delivery order cas 1");
                rnPage = homePage.GoToWarehouse_ReceiptNotesPage();
                rnPage.ResetFilter();
                rnPage.Filter(ReceiptNotesPage.FilterType.ByDeliveryNumber, deliveryNumber2);
                ReceiptNotesItem rnItem = rnPage.SelectFirstReceiptNoteItem();
                rnGeneralInfo = rnItem.ClickOnGeneralInformationTab();
                Assert.AreEqual(deliveryNumber2, rnGeneralInfo.GetDeliveryOrderNumber(), "Mauvais delivery cas 2");
                // les commentaires,
                Assert.AreEqual(commentaire2, rnGeneralInfo.GetComment(), "Mauvais commentaire");
                // et la delivery date à J-1
                Assert.AreEqual(DateUtils.Now.AddDays(-1).ToString("dd/MM/yyyy"), DateTime.Parse(rnGeneralInfo.GetDateValidated()).ToString("dd/MM/yyyy"), "Mauvaise date");
                //*/span[contains(text(),'From purchase order :')]/a
                Assert.AreEqual(poNumber, rnGeneralInfo.GetPurchaseOrderNumber(), "Pas de purchase order dans general info");
                //*/span[contains(text(),'From claim  :')]/a
                Assert.AreEqual(poNumber, rnGeneralInfo.GetClaimNumber(), "Pas de claim dans general info");
                // FIXME Valider la RN sans avoir de message d'erreur sur le backdate et avoir les valeurs,
                // déjà fait.

                // delivery order à "987654" les commentaires, la delivery date à J-1, les quantités indiqués
                // et la claim

            }
            finally
            {
                //A la fin du test remettre le setting sur la valeur antérieur.
                applicationSettingPage = homePage.GoToApplicationSettings();
                var appSettingsModalPage = applicationSettingPage.GetBackDatingPage();
                // 3 sur testauto4dev
                appSettingsModalPage.SetBackDating(oldBackDating);
                appSettingsModalPage.Save();
            }
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_Backdate_Claims()
        {
            //Prepare
            string itemName1 = "CHOCOLATE";
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string location = TestContext.Properties["Location"].ToString();
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();
            string placeTo = TestContext.Properties["PlaceToPartielle"].ToString();
            string deliveryNumber2 = new Random().Next(1000000, 9999999).ToString();
            string commentaire = "commentaire1";
            var decreasestock = true;
            var claimtype = "Non compliant expiration date";
            var newcomment = "commentaireMotifClaim";
            var sanctionamount = "10";
            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            string oldBackDating = null; // 31
            // Le but est de tester que le setting Backdate et de voir les messages d'erreurs apparaître les possibles régressions.
            ApplicationSettingsPage applicationSettingPage = null;
            applicationSettingPage = homePage.GoToApplicationSettings();
            try
            {
                //Avoir une RN validée sur le site
                //Le setting backdate sur 0, le temps du test.
                var appSettingsModalPage = applicationSettingPage.GetBackDatingPage();
                oldBackDating = appSettingsModalPage.GetBackDating();
                appSettingsModalPage.SetBackDating("10");
                appSettingsModalPage.Save();

                //Act
                ReceiptNotesPage rnPage = homePage.GoToWarehouse_ReceiptNotesPage();
                ReceiptNotesCreateModalPage rnCreate = rnPage.ReceiptNotesCreatePage();
                rnCreate.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, deliveryPlace));
                string rnDelivery = rnCreate.GetDeliveryOrderNumber();
                ReceiptNotesItem receiptNotesItem = rnCreate.Submit();
                // Mettre les Physical quantity égale à la PO et DN
                receiptNotesItem.Filter(ReceiptNotesItem.FilterItemType.SearchByName, itemName1);
                receiptNotesItem.SetFirstReceivedQuantity("12");
                // Remplir les checks
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
                receiptNotesItem = qualityChecks.ClickOnReceiptNoteItemTab();
                receiptNotesItem.Validate();
                string numberToCopy = receiptNotesItem.GetReceiptNoteNumber();

                //Depuis Claims,
                //Ouvrir une claims sur la RN validée.
                //Choisir une date à J-1                
                ClaimsPage claim = homePage.GoToWarehouse_ClaimsPage();
                ClaimsCreateModalPage claimModal = claim.ClaimsCreatePage();
                //0  => -1 KO
                //1  => -2 KO
                //10 => -11 KO
                claimModal.FillField_CreatNewClaims(DateUtils.Now.AddDays(-11), site, supplier, placeTo, true, "Receipt Note", numberToCopy);

                //Créer et avoir le message d'erreur bloquant
                ClaimsItem claimItem = claimModal.Submit();
                //Avoir un message d'erreur au moment de choisir la date J-1 et cliquer sur le bouton "create"
                Assert.AreEqual("You are backdating your document, this is not allowed. Contact your Winrest Referent", claimModal.ValidationError(), "Mauvais message d'erreur validation");

                //Changer la date pour J - J et créer
                //Mettre la date sur J-J et pouvoir créer la claims avec le bouton "create"
                //0  => -1 KO
                //1  => -2 KO
                //10 => -11 KO
                claimModal.FillField_CreatNewClaims(DateUtils.Now.AddDays(-10), site, supplier, placeTo, true, "Receipt Note", numberToCopy);
                claimItem = claimModal.Submit();

                //Dans la claims changer la delivery order à "456123" et ajouter un texte de commentaire "commentaire1"
                ClaimsGeneralInformation claimGeneralInfo = claimItem.ClickOnGeneralInformation();
                claimGeneralInfo.SetDeliveryOrderNumber(deliveryNumber2);
                claimGeneralInfo.SetComment(commentaire);
                claimGeneralInfo.PageUp();

                //Se déplacer dans le sous onglet item,

                claimItem = claimGeneralInfo.ClickOnItems();
                //ajouter une quantité,
                string itemName = claimItem.GetFirstItemName();
                claimItem.SelectFirstItem();
                claimItem.SetQuantity("10");
                //mettre un motif
                ClaimEditClaimForm claimEdit = claimItem.EditClaimForm(itemName);
                claimEdit.EditClaim(decreasestock, claimtype, newcomment, sanctionamount);
                //Sans avoir validé la claim et s'être déplacé sur l'autre sous onglet item, avoir le delivery order sur 456123" et le commentaire "commentaire1" apparent.
                claimGeneralInfo = claimItem.ClickOnGeneralInformation();
                Assert.AreEqual(deliveryNumber2, claimGeneralInfo.GetDeliveryOrderNumber(), "Mauvais Delivery Order Number");
                Assert.AreEqual(commentaire, claimGeneralInfo.GetClaimComment(), "Mauvais commentaire");
                claimGeneralInfo.ClickOnItems();
                //Valider la claim
                //Valider la claim sans être bloqué
                claimItem.Validate_CLAIM();


            }
            finally
            {
                //A la fin du test remettre le setting sur la valeur antérieur.
                applicationSettingPage = homePage.GoToApplicationSettings();
                var appSettingsModalPage = applicationSettingPage.GetBackDatingPage();
                // 3 sur testauto4dev
                appSettingsModalPage.SetBackDating(oldBackDating);
                appSettingsModalPage.Save();
            }
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_DeleteReceiptNotes()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();

            // Log in
            var homePage = LogInAsAdmin();

            //Act           
            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();

            // Create
            var receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
            receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, deliveryPlace));
            var receiptNotesItem = receiptNotesCreateModalpage.Submit();
            var recipeNoteNumber = receiptNotesItem.GetReceiptNoteNumber();
            receiptNotesItem.BackToList();
            var totalNumberAfterCreate = receiptNotesPage.CheckTotalNumber();
            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ByNumber, recipeNoteNumber);
            //Delete
            receiptNotesPage.DeleteReceiptNote();
            receiptNotesPage.ResetFilter();
            var totalNumberAfterDelete = receiptNotesPage.CheckTotalNumber();
            //Assert
            Assert.AreEqual(totalNumberAfterDelete, totalNumberAfterCreate - 1, "La suppression ne fonctionne pas.");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_DetailChangeLine()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierPo"].ToString();
            string deliveryPlace = "Economato";

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            ReceiptNotesPage rnPage = homePage.GoToWarehouse_ReceiptNotesPage();
            ReceiptNotesCreateModalPage rnCreate = rnPage.ReceiptNotesCreatePage();
            rnCreate.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, deliveryPlace));
            string rnDelivery = rnCreate.GetDeliveryOrderNumber();
            ReceiptNotesItem receiptNotesItem = rnCreate.Submit();
            receiptNotesItem.SetFirstReceivedQuantity("5");
            Assert.IsTrue(receiptNotesItem.VerifyDetailChangeLine(), "Le curseur ne descend pas sur la line du bas");

        }

        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_index_IconeNotInvoiced()
        {
            // Login
            var homePage = LogInAsAdmin();

            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();
            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ShowValidatedNotInvoiced, true);
            bool checkNotInvoicedIcones = receiptNotesPage.CheckNotInvoicedIcones();
            Assert.IsTrue(checkNotInvoicedIcones, "L'icone Not invoiced n'apparaît pas");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_index_IconeInvoiced()
        {
            string invoicedRN_DN = "InvoicedRN";
            DateTime dateFrom = DateTime.ParseExact("01/01/2024", "dd/MM/yyyy", CultureInfo.InvariantCulture);

            // Login
            var homePage = LogInAsAdmin();

            var rnPage = homePage.GoToWarehouse_ReceiptNotesPage();
            rnPage.ResetFilter();

            rnPage.Filter(ReceiptNotesPage.FilterType.ByDeliveryNumber, invoicedRN_DN);
            rnPage.Filter(ReceiptNotesPage.FilterType.DateFrom, dateFrom);

            bool areAllRNInvoiced = rnPage.AreAllRnInvoiced();
            Assert.IsTrue(areAllRNInvoiced, "L'icone Invoiced n'apparaît pas");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_index_IconePartiallyInvoiced()
        {
            // Login 
            var homePage = LogInAsAdmin();

            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();
            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ShowValidatedPartialInvoiced, true);
            bool checkPartiallyInvoicedIcones = receiptNotesPage.CheckPartiallyInvoicedIcones();
            Assert.IsTrue(checkPartiallyInvoicedIcones, "L'icone Partially invoiced n'apparaît pas");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_index_IconeClaim()
        {
            // Log In
            var homePage = LogInAsAdmin();
            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();
            receiptNotesPage.PageSize("30");
            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ShowWithClaim, true);
            bool checkRnWithClaimsIcones = receiptNotesPage.CheckRnWithClaimsIcones();
            Assert.IsTrue(checkRnWithClaimsIcones, "L'icone Claim n'apparaît pas");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_Index_Filter_ShowValidatedAndNotSentToSAGE_CEGID()
        {
            var homePage = LogInAsAdmin();

            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();

            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ShowValidatedAndNotSentToSage, true);
            receiptNotesPage.PageSize("100");

            // var receiotNotesResults = receiptNotesPage.CheckTotalNumber();
            bool areAllRnValidatedAndNoCegid = receiptNotesPage.AreAllRnValidateAndNoSage();
            Assert.IsTrue(areAllRnValidatedAndNoCegid, "Some RN rows in the filter result are either not validated or with a CEGID icon. It should be validated and no CEGID icon.");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_Footer_Verify()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string delivery = TestContext.Properties["PlaceFrom"].ToString();
           
            var homePage = LogInAsAdmin();
            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();
            ReceiptNotesCreateModalPage receiptNotesCreateModalPage = receiptNotesPage.ReceiptNotesCreatePage();
            receiptNotesCreateModalPage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, delivery));
            var receiptNotesItem = receiptNotesCreateModalPage.Submit();
            var itemName = receiptNotesItem.GetFirstItemName();
            receiptNotesItem.SelectFirstItem();
            receiptNotesItem.AddReceived(itemName, "4");
            ReceiptNotesFooterPage receiptNotesFooterPage = receiptNotesItem.ClickOnFooterTab();
            var taxName = receiptNotesFooterPage.GetTaxName();
            var taxValue = receiptNotesFooterPage.GetVATRate(",");
            var taxBaseAmount = receiptNotesFooterPage.GetTaxBaseAmount(",");
            var totalReceiptNote = receiptNotesFooterPage.GetTotalReceiptNote(",");
            var vatAmount = receiptNotesFooterPage.GetVATAmount(",");
            var itemTabPage = receiptNotesFooterPage.ClickOnItemsTab();
            var total = taxBaseAmount + vatAmount;
            itemTabPage.SelectFirstItem();
            var totalRn = itemTabPage.GetDNTotalSum(",");
            itemTabPage.OpenModalToEditFirstItem();
            receiptNotesItem.SelectFirstItem();
            ItemGeneralInformationPage itemGeneralInformationPage = receiptNotesItem.EditItem(itemName);
            Assert.AreEqual(taxName, itemGeneralInformationPage.GetVatName(), "Assurez-vous que le nom de la taxe est identique.");
            Assert.AreEqual(taxBaseAmount, totalRn, "le total RN sur Items doit etre égal à TAX BASE AMOUNT dans footer");
            Assert.AreEqual(total, totalReceiptNote, 0.00001, "Vérifier le total amount (tax base amount + VAT amount)");
            ParametersSites settingsSitePage = homePage.GoToParameters_Sites();
            settingsSitePage.Filter(ParametersSites.FilterType.SearchSite, site);
            settingsSitePage.ClickOnFirstSite();
            settingsSitePage.ClickToInformations();
            var countrySite = settingsSitePage.GetCountrySelected();
            ParametersPurchasing purchasingPage = homePage.GoToParameters_PurchasingPage();
            purchasingPage.GoToTab_VAT();
            purchasingPage.SearchVAT(taxName, "VAT");
            var taxValueCountry = double.Parse(purchasingPage.GetTaxValueByCountry(countrySite));
            Assert.AreEqual(taxValueCountry, taxValue, 0.00001, "Assurez-vous que la taxe est identique selon  country.");
        }

        [DeploymentItem("Resources\\pizza.png")]
        [DeploymentItem("Resources\\pizza_petite.png")]
        [DeploymentItem("chromedriver.exe")]       
		[TestMethod]
        [Timeout(_timeout)]
        public void WA_RN_Claims_EditClaim()
        {
            //prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string placeTo = TestContext.Properties["PlaceToPartielle"].ToString();
            string typeTaxUpdate = TestContext.Properties["TaxTypeSupplierInvoicesExportSage"].ToString();
            string commentUpdate = "comment updated";
            string sanctionAmountUpdate = "111,22";
            string decreaseQTYUpdate = "55,44";
            string pictureUpdate = "pizza.png";
            string receipeNumber = "";
            string currency = TestContext.Properties["Currency"].ToString();

            // arrange

            var homePage = LogInAsAdmin();

            //Act         
            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();
            try
            {
                //1) Créer une RN
                //Cliquer sur "new Receipt Note"
                var receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
                //Remplir le formulaire
                receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, placeTo));
                //Cliquer sur "Create"
                ReceiptNotesItem receiptNotesItem = receiptNotesCreateModalpage.Submit();
                receipeNumber = receiptNotesItem.GetReceiptNoteNumber();
                //2) Ajouter un item
                string itemName = receiptNotesItem.GetFirstItemName();
                receiptNotesItem.SelectFirstItem();
                receiptNotesItem.AddReceived(itemName, "2");

                //3) Ajouter une claim(icone mégaphone), remplir les champs et save
                ReceiptNotesEditClaim editClaim = receiptNotesItem.AddClaim();
                editClaim.Fill();
                //4) On arrive sur l'onglet claim (des RN)
                ReceiptNotesClaims receiptNotesClaims = editClaim.Save();
                ReceiptNotesEditClaim editClaimReceipeNote = receiptNotesClaims.EditClaim();
                editClaim.SetDecreaseStock(decreaseQTYUpdate);
                editClaimReceipeNote.SetComment(commentUpdate);
                editClaimReceipeNote.UploadPicture(pictureUpdate);
                editClaimReceipeNote.SetSanctionAmount(sanctionAmountUpdate);
                editClaimReceipeNote.SetVAT(typeTaxUpdate);
                var decreaseStockUpdate = editClaimReceipeNote.GetDecreaseStock();
                editClaimReceipeNote.Save();
                CultureInfo culture = new CultureInfo("fr-FR");
                //Assert
                double sanctionAmount = receiptNotesClaims.getSanctionAmount(",", currency);
                Assert.AreEqual(Double.Parse(sanctionAmountUpdate, culture), sanctionAmount, "Vérifier que le sanction amount a bien été modifié");
                double decrQty = Double.Parse(receiptNotesClaims.GetDecrQty(), culture);
                Assert.AreEqual(Double.Parse(decreaseQTYUpdate, culture), decrQty, "Vérifier que le decrease quantité a bien été modifié");
                bool commentCkech = receiptNotesClaims.CommentChecker();
                Assert.AreEqual(!String.IsNullOrEmpty(commentUpdate), commentCkech, "Vérifier que le commentaire a bien été modifié");
                bool pictureCheck = receiptNotesClaims.PictureChecker();
                Assert.AreEqual(!String.IsNullOrEmpty(pictureUpdate), pictureCheck, "Vérifier que le photo a bien été modifié");
                bool decrStock = receiptNotesClaims.GetDecrStock();
                Assert.AreEqual(decreaseStockUpdate, decrStock, "Vérifier que le decrease stock a bien été modifié");
            }
            finally
            {
                homePage.GoToWarehouse_ReceiptNotesPage();
                if (!String.IsNullOrEmpty(receipeNumber))
                {
                    receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ByNumber, receipeNumber);
                    receiptNotesPage.DeleteReceiptNote();
                }
            }

        }
        /// <summary>
        /// Test la pagination dans l'index des RN
        /// </summary>
        /// <param name="receiptNotesPage">La page d'index RN</param>
        /// <param name="listErrors">La liste des erreurs détectées des précédents tests.</param>
        private void WA_RN_index_pagination(ReceiptNotesPage receiptNotesPage, List<string> listErrors)
        {
            //try
            //{
            receiptNotesPage.PageSize("8");
            int nbRows = receiptNotesPage.GetNumberOfShowedResults();
            if (nbRows != 8)
            { listErrors.Add($"On attend 8 lignes mais seulement {nbRows} sont affichés."); }

            receiptNotesPage.PageSize("16");
            nbRows = receiptNotesPage.GetNumberOfShowedResults();
            if (nbRows != 16)
            { listErrors.Add($"On attend 16 lignes mais seulement {nbRows} sont affichés."); }

            receiptNotesPage.PageSize("30");
            nbRows = receiptNotesPage.GetNumberOfShowedResults();
            if (nbRows != 30)
            { listErrors.Add($"On attend 30 lignes mais seulement {nbRows} sont affichés."); }

            receiptNotesPage.PageSize("50");
            nbRows = receiptNotesPage.GetNumberOfShowedResults();
            if (nbRows != 50)
            { listErrors.Add($"On attend 50 lignes mais seulement {nbRows} sont affichés."); }

            receiptNotesPage.PageSize("100");
            nbRows = receiptNotesPage.GetNumberOfShowedResults();
            if (nbRows != 100)
            { listErrors.Add($"On attend 100 lignes mais seulement {nbRows} sont affichés."); }

            receiptNotesPage.PageSize("8");
            bool isPageChangeOK = receiptNotesPage.VerifyChangePage(2);
            Assert.IsTrue(isPageChangeOK, "Le changement de page a rencontré un problème.");
            //}
            //catch (Exception ex)
            //{
            //    listErrors.Add("Erreur dans WA_RN_index_pagination : " + ex.Message);
            //    throw;
            //}
        }

        // _______________________________________ Utilitaire ______________________________________________

        private string CreatePurchaseOrder(HomePage homePage, string site, string supplier, DateTime datePurchaseOrder, string itemName)
        {
            string location = TestContext.Properties["Location"].ToString();
            PurchaseOrdersPage purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            CreatePurchaseOrderModalPage createPurchaseOrderPage = purchaseOrdersPage.CreateNewPurchaseOrder();
            createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, datePurchaseOrder, true);
            PurchaseOrderItem purchaseOrderItemPage = createPurchaseOrderPage.Submit();
            string ID = purchaseOrderItemPage.GetPurchaseOrderNumber();
            purchaseOrderItemPage.ResetFilter();
            purchaseOrderItemPage.Filter(FilterItemType.ByName, itemName);
            purchaseOrderItemPage.SelectFirstItemPo();
            purchaseOrderItemPage.AddQuantity("5");
            if (purchaseOrderItemPage.IsDev()) purchaseOrderItemPage.Approve();
            purchaseOrderItemPage.Validate();
            return ID;
        }

        private string CreatePurchaseOrderPartielle(HomePage homePage, string site, string supplier, string receiptStatus, string itemName)
        {
            // Prepare
            var location = TestContext.Properties["Location"].ToString();

            // Création Purchase Order
            var purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            var createPurchaseOrderPage = purchaseOrdersPage.CreateNewPurchaseOrder();

            createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateUtils.Now.AddDays(+1), true);
            var purchaseOrderItemPage = createPurchaseOrderPage.Submit();

            purchaseOrderItemPage.Filter(FilterItemType.ByName, itemName);
            purchaseOrderItemPage.SelectFirstItemPo();
            purchaseOrderItemPage.AddQuantity("5");
            if (purchaseOrderItemPage.IsDev())
            {
                purchaseOrderItemPage.Approve();
            }
            purchaseOrderItemPage.Validate();

            var idReceiptNote = purchaseOrderItemPage.GenerateReceiptNote(true);
            var receiptNoteItemPage = purchaseOrderItemPage.ValidateReceiptNoteCreation();
            var receiptNoteGeneralInformationPage = receiptNoteItemPage.ClickOnGeneralInformationTab();
            idReceiptNote = receiptNoteGeneralInformationPage.GetReceiptNoteNumber();
            var newpurchaseOrderItemPage = receiptNoteGeneralInformationPage.ClickOnPurchaseOrderLink();
            newpurchaseOrderItemPage.ClickOnGeneralInformation();
            newpurchaseOrderItemPage.ChangeReceiptStatus(receiptStatus);

            return idReceiptNote;
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
                appSettingsModalPage = applicationSettings.GetWinrestExportTLSageDbOverloadPage();
                appSettingsModalPage.SetWinrestExportTLSageDbOverload(versionBDD);
                applicationSettings = appSettingsModalPage.Save();

                // Override countryCode
                appSettingsModalPage = applicationSettings.GetWinrestExportTLSageCountryCodeOverloadPage();
                appSettingsModalPage.SetWinrestExportTLSageCountryCodeOverload(environnement);
                appSettingsModalPage.Save();
            }
            catch
            {
                return false;
            }

            return true;
        }

        private bool VerifySiteAnalyticalPlanSection(HomePage homePage, string site, bool isOK = true)
        {
            string analyticalPlan = "1";
            string analyticalSection = "314";
            string dueToInvoicePlan = isOK ? "10" : "";
            string dueToInvoiceSection = isOK ? "500" : "";

            try
            {
                var settingsSitesPage = homePage.GoToParameters_Sites();
                settingsSitesPage.Filter(ParametersSites.FilterType.SearchSite, site);
                settingsSitesPage.ClickOnFirstSite();

                settingsSitesPage.ClickToInformations();
                settingsSitesPage.SetAnalyticPlan(analyticalPlan);
                settingsSitesPage.SetAnalyticSection(analyticalSection);
                settingsSitesPage.SetDueToInvoiceAnalyticPlan(dueToInvoicePlan);
                settingsSitesPage.SetDueToInvoiceAnalyticSection(dueToInvoiceSection);
            }
            catch
            {
                return false;
            }

            return true;
        }

        private bool VerifyPurchasingVAT(HomePage homePage, string taxName, string taxType, bool isOK = true)
        {
            // Prepare
            string purchaseCode = "AS10";
            string purchaseAccount = "47205001";
            string salesCode = "AR10";
            string salesAccount = "47205001";
            string dueToInvoiceCode = isOK ? "AP10" : "";
            string dueToInvoiceAccount = isOK ? "47205001" : "";
            string taxValue = "21";

            try
            {
                // Act
                var settingsPurchasingPage = homePage.GoToParameters_PurchasingPage();
                settingsPurchasingPage.GoToTab_VAT();
                if (!settingsPurchasingPage.IsTaxPresent(taxName, taxType))
                {
                    settingsPurchasingPage.CreateNewVAT(taxName, taxType, taxValue);
                }

                settingsPurchasingPage.SearchVAT(taxName, taxType);
                settingsPurchasingPage.EditVATAccountForSpain(purchaseCode, purchaseAccount, salesCode, salesAccount, dueToInvoiceCode, dueToInvoiceAccount);
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

            try
            {
                // Act
                var accountingParametersPage = homePage.GoToParameters_AccountingPage();
                accountingParametersPage.GoToTab_GroupVats();

                if (!isOK && accountingParametersPage.IsGroupAndTaxPresent(group, vat))
                {
                    accountingParametersPage.DeleteGroup(group, vat);
                }
                else
                {
                    if (!accountingParametersPage.IsGroupAndTaxPresent(group, vat))
                    {
                        accountingParametersPage.CreateNewGroup(group, vat);
                    }

                    accountingParametersPage.SearchGroup(group, vat);
                    accountingParametersPage.EditInventoryAccounts(account, exoAccount);
                }
            }
            catch
            {
                return false;
            }

            return true;
        }

        private bool VerifyAccountingJournal(HomePage homePage, string site, string journalRN)
        {
            try
            {
                var accountingJournalPage = homePage.GoToParameters_AccountingPage();
                accountingJournalPage.GoToTab_Journal();
                accountingJournalPage.EditJournal(site, null, null, journalRN);
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
            //Decommenté pour la nouvelle version 31/08/2021 Cécile //accountingParametersPage.GoToTab_AccountSettings()
            accountingParametersPage.GoToTab_MonthlyClosingDays();

            //Decommenté pour la nouvelle version 31/08/2021 Cécile //return accountingParametersPage.GetSageClosureDayIndex()
            return accountingParametersPage.GetSageClosureMonthIndex();
        }

        private bool VerifySupplier(HomePage homePage, string site, string supplier, bool isOK = true)
        {
            // Prepare
            string thirdId = "10";
            string accountingId = "10";
            string thirdIdDTI = isOK ? "10" : "";
            string accountingIdDTI = isOK ? "10" : "";

            try
            {
                // Act
                var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
                suppliersPage.Filter(SuppliersPage.FilterType.Search, supplier);
                var supplierItem = suppliersPage.SelectFirstItem();
                supplierItem.ClickOnAccountingTab();

                if (!supplierItem.IsAccountingPresent(site))
                {
                    supplierItem.CreateNewAccounting(thirdId, accountingId, thirdIdDTI, accountingIdDTI);
                }

                supplierItem.SearchAccounting(site);
                supplierItem.EditAccounting(thirdId, accountingId, thirdIdDTI, accountingIdDTI);
            }
            catch
            {
                return false;
            }

            return true;
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_NewRNBackdate()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();

            //Arrange

            LogInAsAdmin();

            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Ouvrir le menu
            ReceiptNotesPage pageReceiptNotes = homePage.GoToWarehouse_ReceiptNotesPage();

            // Remplir le formulaire
            var receiptNotesCreateModalpage = pageReceiptNotes.ReceiptNotesCreatePage();
            receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now.AddDays(-11), site, supplier, deliveryPlace));

            // Cliquer sur "Create"
            ReceiptNotesItem receiptNotesItem = receiptNotesCreateModalpage.Submit();
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_index_Filter_WaitinginCegidSAS()
        {
            //Arrange
            var homePage = LogInAsAdmin();
            // Ouvrir le menu
            ReceiptNotesPage pageReceiptNotes = homePage.GoToWarehouse_ReceiptNotesPage();
            pageReceiptNotes.ResetFilter();
            var defaultResult = pageReceiptNotes.GetTotalRecipeNotes();
            pageReceiptNotes.Filter(ReceiptNotesPage.FilterType.ExportedWithReverseGenerated, true);
            var resultAfterFilter = pageReceiptNotes.GetTotalRecipeNotes();
            Assert.AreNotEqual(int.Parse(defaultResult), int.Parse(resultAfterFilter), "Le filter par Exported With Reverse Generated ne fonctionne pas");
        }


        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_index_Filter_ExportedWithProvisionGenerated()
        {
            var homePage = LogInAsAdmin();
            ReceiptNotesPage receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();
            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ShowAll, true);
            int nombreReceiptsShowAll = receiptNotesPage.CheckTotalNumber();
            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ExportedWithReverseGenerated, true);
            int nombreReceiptsExportedWithReverseGenerated = receiptNotesPage.CheckTotalNumber();
            Assert.AreNotEqual(nombreReceiptsShowAll, nombreReceiptsExportedWithReverseGenerated);
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_index_Filter_ExportedWithReverseGenerated()
        {
            var homePage = LogInAsAdmin();
            ReceiptNotesPage receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();
            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ShowAll, true);
            int nombreReceiptsShowAll = receiptNotesPage.CheckTotalNumber();
            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ExportedWithReverseGenerated, true);
            int nombreReceiptsExportedWithReverseGenerated = receiptNotesPage.CheckTotalNumber();
            Assert.AreNotEqual(nombreReceiptsShowAll, nombreReceiptsExportedWithReverseGenerated);
        }
        
        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_index_pagination()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();

            List<string> listDetectedErrors = new List<string>();
            //Login
            var homePage = LogInAsAdmin();

            ReceiptNotesPage receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();
            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.Supplier, supplier);
            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.Site, site);
            if (receiptNotesPage.CheckTotalNumber() > 0)
            {
                receiptNotesPage.PageSize("8");
                var totalPages = (int)Math.Ceiling((double)receiptNotesPage.CheckTotalNumber() / 8);
                var nbpageresult = receiptNotesPage.GetNumberOfPageResults();
                Assert.AreEqual(totalPages, nbpageresult, "Le changement de page a rencontré un problème.");
                if (nbpageresult > 1)
                {
                    for (int i = 1; i < nbpageresult; i++)
                    {
                        receiptNotesPage.GetPageResults(i + 1);
                        var nbRowspage = receiptNotesPage.GetNumberOfShowedResults();
                        Assert.AreNotEqual(0, nbRowspage, "problème de naviguer de page en page");
                    }
                }
                if (receiptNotesPage.CheckTotalNumber() >= 100)
                {
                    receiptNotesPage.GetPageResults(1);
                    this.WA_RN_index_pagination(receiptNotesPage, listDetectedErrors);
                }
            }
            if (listDetectedErrors.Any())
            {
                Assert.IsTrue(false, string.Join(Environment.NewLine, listDetectedErrors));
            }
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_index_ResetFilter()
        {
            var homePage = LogInAsAdmin();
            ReceiptNotesPage receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();
            Assert.IsTrue(receiptNotesPage.ResetFilterCheck(), "Les filtres ne sont pas réinitialisés");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_index_price()
        {
            //prepare
            string currency = TestContext.Properties["Currency"].ToString();
            string decimalSeparator = ",";
            //arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            ReceiptNotesPage receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();
            var number = receiptNotesPage.GetFirstReceiptNoteNumber();
            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ByNumber, number);
            var priceIndex = receiptNotesPage.GetIndexPrice(decimalSeparator, currency);
            var receiptNoteDetails = receiptNotesPage.ClickFirstReceiptNote();
            var priceRN = receiptNoteDetails.GetRNTotal(decimalSeparator, currency);
            Assert.AreEqual(priceIndex, priceRN, "Les totales ne sont pas égaux ");
        }

        /// <summary>
        /// Update receipt's first item allergens and check the state 
        /// of the allergens button in the receipt details.
        /// </summary>
        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_CheckAllergensUpdate()
        {
            string allergen1 = "Cacahuetes/Peanuts";
            string allergen2 = "Frutos de cáscara- Macadamias/Nuts-Macadamia";
            HomePage homePage = LogInAsAdmin();
            ReceiptNotesPage receiptsPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptsPage.ResetFilter();
            receiptsPage.Filter(ReceiptNotesPage.FilterType.ShowNotValidated, true);
            receiptsPage.Filter(ReceiptNotesPage.FilterType.ByNumber, receiptsPage.GetFirstRecipeNoteNumber());
            ReceiptNotesItem receiptsItem = receiptsPage.SelectFirstReceiptNoteItem();
            string itemName = receiptsItem.GetFirstItemName();
            receiptsItem.ResetFilters();
            receiptsItem.Filter(ReceiptNotesItem.FilterItemType.SearchByName, itemName);
            receiptsItem.SelectFirstItem();
            receiptsItem.AddReceived(itemName, "2");
            receiptsItem.Refresh();
            receiptsItem.ResetFilters();
            receiptsItem.Filter(ReceiptNotesItem.FilterItemType.SearchByName, itemName);
            receiptsItem.SelectFirstItem();
            ItemGeneralInformationPage itemPage = receiptsItem.EditItemGeneralInformation(itemName);
            ItemIntolerancePage itemIntolerancePage = itemPage.ClickOnIntolerancePage();
            itemIntolerancePage.AddAllergen(allergen1);
            itemIntolerancePage.AddAllergen(allergen2);
            receiptsPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptsPage.ResetFilter();
            receiptsPage.Filter(ReceiptNotesPage.FilterType.ShowNotValidated, true);
            receiptsPage.Filter(ReceiptNotesPage.FilterType.ByNumber, receiptsPage.GetFirstRecipeNoteNumber());
            receiptsItem = receiptsPage.SelectFirstReceiptNoteItem();
            receiptsItem.ResetFilters();
            receiptsItem.Filter(ReceiptNotesItem.FilterItemType.SearchByName, itemName);
            bool isIconGreen = receiptsItem.IsAllergenIconGreen(itemName);
            List<string> allergensInItem = receiptsItem.GetAllergens(itemName);
            bool containsAllergen1 = allergensInItem.Contains(allergen1);
            bool containsAllergen2 = allergensInItem.Contains(allergen2);
            Assert.IsTrue(isIconGreen, "L'icon n'est pas vert!");
            Assert.IsTrue(containsAllergen1 && containsAllergen2, "Allergens n'ont pas été ajoutés");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void WA_RN_index_Filter_ExportedForSageOrCegidAutomatically()
        {
            // Log in
            var homePage = LogInAsAdmin();
            //Act           
            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();
            var defaultResult = receiptNotesPage.GetTotalRecipeNotes();
            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ExportedForSageOrCegidAutomatically, true);
            var resultAfterFilter = receiptNotesPage.GetTotalRecipeNotes();
            Assert.AreNotEqual(int.Parse(defaultResult), int.Parse(resultAfterFilter), "Le filter par Exported For SageOrCegid Automatically ne fonctionne pas");
        }
    }
}
