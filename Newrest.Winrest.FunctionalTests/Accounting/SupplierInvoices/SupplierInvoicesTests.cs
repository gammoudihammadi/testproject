using DocumentFormat.OpenXml.VariantTypes;
using GemBox.Email.Mime;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Accounting.SupplierInvoices;
using Newrest.Winrest.FunctionalTests.PageObjects.Interim.InterimReceptions;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.GlobalSettings;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Production;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Sites;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.ProductionCO;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Suppliers;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.Claims;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.ReceiptNotes;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.Extensions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Policy;
using System.Threading;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using static Newrest.Winrest.FunctionalTests.PageObjects.Accounting.SupplierInvoices.SupplierInvoicesPage;
using static Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.PurchaseOrderItem;

namespace Newrest.Winrest.FunctionalTests.Accounting
{
    [TestClass]
    public class SupplierInvoicesTests : TestBase
    {
        private static Random random = new Random();
        private const int _timeout = 600000;
        private readonly string SUPPLIER_INVOICE_EXCEL_SHEET = "Invoices";
        private string random_supplier_name = "supplierName-" + random.Next().ToString();
        private string supplierType = "MonSupplierType";
        private string itemName = "supplier-Item" + new Random().Next().ToString();
        private string supplierInvoiceNb = random.Next().ToString();

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
                case nameof(AC_SI_Create_NewCreditNote):
                    TestInitialize_Create_SupplierType();
                    TestInitialize_Create_Supplier_With_Item_Packaging();
                    TestInitialize_Create_Supplier_Invoice_CreditNoteActivated();
                    break;

                case nameof(AC_SI_NewSI_Activated):
                    TestInitialize_Create_SupplierType();
                    TestInitialize_Create_Supplier_With_Item_Packaging();
                    TestInitialize_Create_Supplier_Invoice();
                    break;

                case nameof(AC_SI_Filter_ShowTransformedIntoCustomerInvoice):
                    TestInitialize_Create_SupplierType();
                    TestInitialize_Create_Supplier_With_Item_Packaging();
                    TestInitialize_Create_Supplier_Invoice_ValidWithItem();
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
        public void TestInitialize_Create_Supplier_With_Item_Packaging()
        {
            //Arrange
            HomePage homePage = LogInAsAdmin();
            //Act
            SuppliersPage suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            SupplierCreateModalPage supplierCreateModalpage = suppliersPage.SupplierCreatePage();
            supplierCreateModalpage.FillField_CreatNewSupplier(random_supplier_name, true, false);
            SupplierItem supplierItem = supplierCreateModalpage.Submit();
            SupplierGeneralInfoTab suppliergeneralInfo = supplierItem.ClickOnGeneralInformationTab();
            suppliergeneralInfo.SetSupplierType(supplierType);
            CreateNewItem(supplierItem, random_supplier_name);
            suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, random_supplier_name);
            var displayedSupplier = suppliersPage.GetFirstSupplierName();
            Assert.AreEqual(random_supplier_name, displayedSupplier, "Le supplier créé n'est pas présent dans la liste.");
        }
        public void TestInitialize_Create_Supplier_Invoice_CreditNoteActivated()
        {

            string site = TestContext.Properties["Site"].ToString();
            string qty = "5";

            HomePage homePage = LogInAsAdmin();
            //Act
            //Create invoice
            SupplierInvoicesPage supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();

            // Create
            SupplierInvoicesCreateModalPage supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
            supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoicesCreditNote(supplierInvoiceNb, DateUtils.Now, site, random_supplier_name, true, false);
            SupplierInvoicesItem supplierItem = supplierInvoicesCreateModalpage.Submit();
            supplierItem.ShowBtnPlus();
            string priceSiBeforeItemAdded = supplierItem.GetSI_PackPriceAdded();
            supplierItem.AddNewItemSiPriceAuto(itemName, qty);
            string priceSiAfterItemAdded = supplierItem.GetSI_PackPriceAdded();
            supplierItem.SubmitBtn();
            supplierItem.BackToList();
        }
        public void TestInitialize_Create_Supplier_Invoice()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();


            //arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();

            // Create
            var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
            supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, random_supplier_name, true, false);
            var supplierItem = supplierInvoicesCreateModalpage.Submit();

            supplierInvoicesPage = supplierItem.BackToList();
            supplierInvoicesPage.ResetFilter();
            supplierInvoicesPage.Filter(SupplierInvoicesPage.FilterType.ByNumber, supplierInvoiceNb);
            //Assert
            string FirstInvoiceNumber = supplierInvoicesPage.GetFirstInvoiceNumber();
            Assert.AreEqual(supplierInvoiceNb, FirstInvoiceNumber, "La supply invoice créée n'apparaît pas dans la liste des supply invoices disponibles.");

        }


        public void TestInitialize_Create_Supplier_Invoice_ValidWithItem()
        {

            string site = TestContext.Properties["Site"].ToString();
            string qty = "5";
            HomePage homePage = LogInAsAdmin();
            //Act
            //Create invoice
            SupplierInvoicesPage supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();

            // Create
            SupplierInvoicesCreateModalPage supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
            supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoicesCreditNote(supplierInvoiceNb, DateUtils.Now, site, random_supplier_name, true, false);
            SupplierInvoicesItem supplierItem = supplierInvoicesCreateModalpage.Submit();
            supplierItem.ShowBtnPlus();
            supplierItem.AddNewItemSiPriceAuto(itemName, qty);
            supplierItem.SubmitBtn();
            supplierItem.ValidateSupplierInvoice();
            supplierItem.BackToList();
        }
        private void CreateNewItem(SupplierItem supplierItem, string supplierName)
        {
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();

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


        [TestMethod]
        [Priority(0)]
        [Timeout(_timeout)]
        public void AC_SI_SetConfigWinrest4_0()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string delivery = TestContext.Properties["PlaceFrom"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            ClearCache();

            // New version search
            homePage.SetNewVersionSearchValue(true);

            // New group display
            homePage.SetNewGroupDisplayValue(true);

            // Vérifier New version search
            try
            {
                var itemPage = homePage.GoToPurchasing_ItemPage();
                var firstItem = itemPage.GetFirstItemName();
                itemPage.Filter(ItemPage.FilterType.Search, firstItem);
            }
            catch
            {
                throw new Exception("La recherche n'a pas pu être effectuée, le NewSearchMode est inactif.");
            }

            // Vérifier New version supplier invoice
            var supplierInvoicePage = homePage.GoToAccounting_SupplierInvoices();
            Assert.IsFalse(supplierInvoicePage.IsInvoiceAmountWithoutTaxPresent(), "Le paramètre 'NewSupplierInvoiceVersion' n'est pas activé.");

            // Vérifier new group display
            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();

            var receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
            receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, delivery));
            var receiptNotesItem = receiptNotesCreateModalpage.Submit();

            string ID = receiptNotesItem.GetReceiptNoteNumber();

            try
            {
                Assert.IsTrue(receiptNotesItem.IsGroupDisplayActive(), "Le paramètre 'NewGroupDisplay' n'est pas activé.");
            }
            finally
            {
                receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
                receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ByNumber, ID);
                receiptNotesPage.DeleteReceiptNote();
            }


        }

        [Priority(1)]
        [TestMethod]
        [Timeout(_timeout)]
        public void AC_SI_CreateSupplierForSupplierInvoice()
        {
            //Prepare
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, supplier);

            if (suppliersPage.CheckTotalNumber() == 0)
            {
                var supplierCreateModalpage = suppliersPage.SupplierCreatePage();
                supplierCreateModalpage.FillField_CreatNewSupplier(supplier, true, false);
                var supplierItem = supplierCreateModalpage.Submit();
                suppliersPage = supplierItem.BackToList();
            }

            suppliersPage.ResetFilter();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, supplier);

            // Assert
            Assert.AreEqual(supplier, suppliersPage.GetFirstSupplierName(), "Le supplier créé n'est pas présent dans la liste.");
        }

        [Priority(2)]
        [TestMethod]
        [Timeout(_timeout)]
        public void AC_SI_CreateItemsForSupplierInvoice()
        {
            //Prepare groups
            string supplierInvoiceGroup = TestContext.Properties["SupplierInvoiceItemGroup"].ToString();
            string supplierInvoiceGroupCode = TestContext.Properties["SupplierInvoiceItemGroupCode"].ToString();

            string supplierInvoiceGroup2 = TestContext.Properties["SupplierInvoiceItemGroupBis"].ToString();
            string supplierInvoiceGroupCode2 = TestContext.Properties["SupplierInvoiceItemGroupCodeBis"].ToString();

            // Prepare items
            string supplierInvoiceItem = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string supplierInvoiceItem2 = TestContext.Properties["Item_SupplierInvoiceBis"].ToString();
            string supplierInvoiceItem3 = TestContext.Properties["Item_SupplierInvoiceGrp"].ToString();

            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string storageUnit = "KG";
            string supplierRef = supplierInvoiceItem + "_SupplierRef";
            string storageQty = 10.ToString();
            string qty = 10.ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Créer groupe
            var productionPage = homePage.GoToParameters_ProductionPage();
            productionPage.GoToTab_Group();

            if (!productionPage.IsGroupPresent(supplierInvoiceGroup))
            {
                productionPage.AddNewGroup(supplierInvoiceGroup, supplierInvoiceGroupCode);
            }
            Assert.IsTrue(productionPage.IsGroupPresent(supplierInvoiceGroup), "Le groupe " + supplierInvoiceGroup + " n'a pas été créé.");

            if (!productionPage.IsGroupPresent(supplierInvoiceGroup2))
            {
                productionPage.AddNewGroup(supplierInvoiceGroup2, supplierInvoiceGroupCode2);
            }
            Assert.IsTrue(productionPage.IsGroupPresent(supplierInvoiceGroup2), "Le groupe " + supplierInvoiceGroup2 + " n'a pas été créé.");

            // Création items
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.Filter(ItemPage.FilterType.Search, supplierInvoiceItem);

            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(supplierInvoiceItem, supplierInvoiceGroup, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier, supplierRef);
                itemPage = itemGeneralInformationPage.BackToList();

                itemPage.Filter(ItemPage.FilterType.Search, supplierInvoiceItem.ToString());
            }
            Assert.AreEqual(supplierInvoiceItem, itemPage.GetFirstItemName(), "L'item " + supplierInvoiceItem + " n'est pas présent dans la liste des items disponibles.");

            // Item2
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, supplierInvoiceItem2);

            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(supplierInvoiceItem2, supplierInvoiceGroup2, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier, supplierRef);
                itemPage = itemGeneralInformationPage.BackToList();

                itemPage.Filter(ItemPage.FilterType.Search, supplierInvoiceItem2.ToString());
            }
            Assert.AreEqual(supplierInvoiceItem2, itemPage.GetFirstItemName(), "L'item " + supplierInvoiceItem2 + " n'est pas présent dans la liste des items disponibles.");

            // Item3
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, supplierInvoiceItem3);

            if (itemPage.CheckTotalNumber() == 0)
            {
                var itemCreateModalPage = itemPage.ItemCreatePage();
                var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(supplierInvoiceItem3, supplierInvoiceGroup, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemGeneralInformationPage = itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier, supplierRef, "20");
                itemPage = itemGeneralInformationPage.BackToList();

                itemPage.Filter(ItemPage.FilterType.Search, supplierInvoiceItem3.ToString());
            }
            Assert.AreEqual(supplierInvoiceItem3, itemPage.GetFirstItemName(), "L'item " + supplierInvoiceItem3 + " n'est pas présent dans la liste des items disponibles.");
        }

        [Priority(3)]
        [TestMethod]
        [Timeout(_timeout)]
        public void AC_SI_ConfigureTaxForGroups()
        {
            //Prepare groups
            string groupTax = TestContext.Properties["SupplierInvoiceManualTax"].ToString().Replace(" ", "-");
            string groupDeviation = TestContext.Properties["SupplierInvoiceDeviation"].ToString();
            string taxType = TestContext.Properties["TaxTypeSupplierInvoicesExportSage"].ToString();

            string account = "60105100";
            string exoAccount = "60105100";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Configurer le groupe pour l'export SAGE et le site
            var accountingParametersPage = homePage.GoToParameters_AccountingPage();
            accountingParametersPage.GoToTab_GroupVats();

            if (!accountingParametersPage.IsGroupAndTaxPresent(groupTax, taxType))
            {
                accountingParametersPage.CreateNewGroup(groupTax, taxType);
            }

            accountingParametersPage.SearchGroup(groupTax, taxType);
            accountingParametersPage.EditInventoryAccounts(account, exoAccount);

            if (!accountingParametersPage.IsGroupAndTaxPresent(groupDeviation, taxType))
            {
                accountingParametersPage.CreateNewGroup(groupDeviation, taxType);
            }

            accountingParametersPage.SearchGroup(groupDeviation, taxType);
            accountingParametersPage.EditInventoryAccounts(account, exoAccount);

            Assert.IsTrue(accountingParametersPage.IsGroupAndTaxPresent(groupTax, taxType), "Le groupe " + groupTax + " n'est pas configuré.");
            Assert.IsTrue(accountingParametersPage.IsGroupAndTaxPresent(groupDeviation, taxType), "Le groupe " + groupDeviation + " n'est pas configuré.");
        }

        [Priority(4)]
        [TestMethod]
        [Timeout(_timeout)]
        public void AC_SI_PrepareExportSageAutoConfig()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string taxName = TestContext.Properties["TaxTypeSupplierInvoicesExportSage"].ToString();
            string taxType = "VAT";
            string journalSI = TestContext.Properties["Journal_SI"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Récupération du groupe de l'item
            string itemGroup = GetItemGroup(homePage, itemName);

            // Vérification du paramétrage
            // --> Admin Settings
            bool isAppSettingsOK = SetApplicationSettingsForSageAuto(homePage);

            // Sites -- > Analytical plan et section
            bool isAnalyticalPlanOK = VerifySiteAnalyticalPlanSection(homePage, site);

            // Sites --> Contact sage supplier invoice
            bool isMailSageOK = VerifySupplierInvoiceSageContact(homePage, site);

            // Parameter - Purchasing --> VAT
            bool isPurchasingVATOK = VerifyPurchasingVAT(homePage, taxName, taxType);

            // Parameter - Accounting --> Service categories & VAT
            bool isGroupAndVatOK = VerifyGroupAndVAT(homePage, itemGroup, taxName);

            // Parameter - Accounting --> Journal
            bool isJournalOk = VerifyAccountingJournal(homePage, site, journalSI);

            // Parameter - Accounting --> Integration Date
            //
            DateTime date = VerifyIntegrationDate(homePage);

            // Customer
            bool isSupplierOK = VerifySupplier(homePage, site, supplier);

            // Assert
            Assert.AreNotEqual("", itemGroup, "Le groupe de l'item n'a pas été récupéré.");
            Assert.IsTrue(isAppSettingsOK, "Les application settings pour TL ne sont pas configurés correctement.");
            Assert.IsTrue(isAnalyticalPlanOK, "La configuration des analytical plan du site n'est pas effectuée.");
            Assert.IsTrue(isMailSageOK, $"Aucun mail n'est configuré pour le site {site} en cas d'erreur Sage.");
            Assert.IsTrue(isPurchasingVATOK, "La configuration des purchasing VAT n'est pas effectuée.");
            Assert.IsTrue(isGroupAndVatOK, "La configuration du group and VAT de l'item n'est pas effectuée.");
            Assert.IsTrue(isJournalOk, "La catégorie du accounting journal n'a pas été effectuée.");
            Assert.IsNotNull(date, "La date d'intégration est nulle.");
            Assert.IsTrue(isSupplierOK, "La configuration du supplier n'a pas été effectuée.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_CreateSupplierInvoice()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();

            //arrange
            HomePage homePage = LogInAsAdmin();
            try
            {
                //Act
                var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();

                // Create
                var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
                var supplierItem = supplierInvoicesCreateModalpage.Submit();

                supplierInvoicesPage = supplierItem.BackToList();
                supplierInvoicesPage.ResetFilter();
                supplierInvoicesPage.Filter(SupplierInvoicesPage.FilterType.ByNumber, supplierInvoiceNb);
                //Assert
                Assert.AreEqual(supplierInvoiceNb, supplierInvoicesPage.GetFirstInvoiceNumber(), "La supply invoice créée n'apparaît pas dans la liste des supply invoices disponibles.");
            }
            finally
            {
                homePage.Navigate();
                var suppliersInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                suppliersInvoicesPage.ResetFilter();
                suppliersInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
                suppliersInvoicesPage.WaitPageLoading();
                suppliersInvoicesPage.DeleteFirstSI();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_DeleteSupplierInvoice()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            // Log in
            HomePage homePage = LogInAsAdmin();
            //Act
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();

            // Create
            var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
            supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
            var supplierItem = supplierInvoicesCreateModalpage.Submit();
            supplierInvoicesPage = supplierItem.BackToList();
            supplierInvoicesPage.ResetFilter();

            //Assert
            string firstInvoicenb = supplierInvoicesPage.GetFirstInvoiceNumber();
            Assert.AreEqual(supplierInvoiceNb, firstInvoicenb, "La supply invoice créée n'apparaît pas dans la liste des supply invoices disponibles.");

            supplierInvoicesPage.Filter(SupplierInvoicesPage.FilterType.ByNumber, supplierInvoiceNb);
            supplierInvoicesPage.WaitPageLoading();
            supplierInvoicesPage.DeleteFirstSI();
            //Assert
            int nbItems = supplierItem.GetNumberOfItems();
            Assert.AreEqual(0, nbItems, "La supplier invoice n'a pas été supprimée.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_CreateSupplierInvoiceFromReceiptNote()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["Place"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();

            // Log in

            var homePage = LogInAsAdmin();

            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();

            // 1. Create a receipt note
            var receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
            receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, place));
            var receiptNotesItem = receiptNotesCreateModalpage.Submit();

            string receiptNoteNumber = receiptNotesItem.GetReceiptNoteNumber();

            receiptNotesItem.Filter(ReceiptNotesItem.FilterItemType.SearchByName, itemName);
            receiptNotesItem.SelectItem(itemName);
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

            // 2. Create supplier invoice
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            try
            {
                var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, true);
                supplierInvoicesCreateModalpage.CreateSIFromRN(receiptNoteNumber);
                var supplierItem = supplierInvoicesCreateModalpage.Submit();
                var numberOfItem = supplierItem.GetNumberOfItems();
                Assert.AreEqual(1, numberOfItem, "La supplier invoice créée à partir d'une RN n'a pas le même nombre d'items renseignés que la RN.");
                var isFiltered = supplierItem.IsItemFiltered(itemName);
                Assert.IsTrue(isFiltered, "La supplier invoice n'a pas d'item relié à la RN utilisée pour sa création.");

                supplierInvoicesPage = supplierItem.BackToList();
                supplierInvoicesPage.ResetFilter();

                //Assert
                var displayedInvoiceNumber = supplierInvoicesPage.GetFirstInvoiceNumber();
                Assert.AreEqual(supplierInvoiceNb, displayedInvoiceNumber, "La supply invoice créée n'apparaît pas dans la liste des supply invoices disponibles.");
            }

            finally
            {
                supplierInvoicesPage.DeleteFirstSI();
            }

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_CreateSupplierInvoiceFromPurchaseOrder()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["Place"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // 1. Create purchase order
            var purchaseOrderPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            purchaseOrderPage.ResetFilters();

            var createPurchaseOrderPage = purchaseOrderPage.CreateNewPurchaseOrder();
            createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, place, DateUtils.Now.AddDays(+1), true);
            var purchaseOrderItemPage = createPurchaseOrderPage.Submit();
            var purchaseOrderNumber = purchaseOrderItemPage.GetPurchaseOrderNumber();

            purchaseOrderItemPage.Filter(FilterItemType.ByName, itemName);
            purchaseOrderItemPage.SelectFirstItemPo();
            purchaseOrderItemPage.AddQuantity("5");
            purchaseOrderItemPage.Validate();

            // 2. Create receipt note
            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();

            var receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
            receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, place) { CreateFromPO = true, PONumber = purchaseOrderNumber, PODate = DateUtils.Now });
            receiptNotesCreateModalpage.Submit();

            // 3. Create supplier invoice
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();

            var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
            supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, true);
            supplierInvoicesCreateModalpage.CreateSIFromPO(purchaseOrderNumber);
            var supplierItem = supplierInvoicesCreateModalpage.Submit();

            Assert.AreEqual(1, supplierItem.GetNumberOfItems(), "La supplier invoice créée à partir d'un PO n'a pas le même nombre d'items renseignés que le PO.");
            Assert.IsTrue(supplierItem.IsItemFiltered(itemName), "La supplier invoice n'a pas d'item relié au PO utilisée pour sa création.");

            supplierInvoicesPage = supplierItem.BackToList();
            supplierInvoicesPage.ResetFilter();
            supplierInvoicesPage.Filter(FilterType.ShowEDI, false);

            //Assert
            Assert.AreEqual(supplierInvoiceNb, supplierInvoicesPage.GetFirstInvoiceNumber(), "La supply invoice créée n'apparaît pas dans la liste des supply invoices disponibles.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_CreateSupplierInvoiceFromCreditNote()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // 1. Create Claim
            var idClaim = CreateClaim(homePage, site, supplier, itemName);

            // 2. Create supplier invoice
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();

            var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
            supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, true);
            supplierInvoicesCreateModalpage.CreateSIFromClaim(idClaim);
            supplierInvoicesCreateModalpage.WaitLoading();
            var supplierItem = supplierInvoicesCreateModalpage.Submit();

            Assert.AreEqual(1, supplierItem.GetNumberOfItems(), "La supplier invoice créée à partir d'une Claim n'a pas le même nombre d'items renseignés que la Claim.");
            Assert.IsTrue(supplierItem.IsItemFiltered(itemName), "La supplier invoice n'a pas d'item relié à la Claim utilisée pour sa création.");

            supplierInvoicesPage = supplierItem.BackToList();
            supplierInvoicesPage.ResetFilter();
            supplierInvoicesPage.Filter(FilterType.ShowEDI, false);
            //Assert
            Assert.AreEqual(supplierInvoiceNb, supplierInvoicesPage.GetFirstInvoiceNumber(), "La supply invoice créée n'apparaît pas dans la liste des supply invoices disponibles.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_CreateSupplierInvoiceTestError()
        {
            //Prepare
            string messageErreur = "Le message d'erreur concernant {0} ne s'est pas affiché.";

            // Log in
            HomePage homePage = LogInAsAdmin();

            //Act
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            // Create
            var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
            supplierInvoicesCreateModalpage.Submit();

            bool errorInvoiceNumber = supplierInvoicesCreateModalpage.ErrorMessageInvoiceNumberRequired();
            bool msgError = supplierInvoicesCreateModalpage.ErrorMessageSupplierRequired();
            //Assert
            Assert.IsTrue(errorInvoiceNumber, String.Format(messageErreur, "l'invoice number"));
            Assert.IsTrue(msgError, String.Format(messageErreur, "le supplier"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_CreateSeveralSupplierInvoicesFromOneReceiptNote()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string supplierInvoiceNbBis = supplierInvoiceNb + "Bis".ToString();
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["Place"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string itemName2 = TestContext.Properties["Item_SupplierInvoiceBis"].ToString();

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();

            // 1. Create a receipt note with 2 items
            var receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
            receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, place));
            var receiptNotesItem = receiptNotesCreateModalpage.Submit();

            string receiptNoteNumber = receiptNotesItem.GetReceiptNoteNumber();

            receiptNotesItem.Filter(ReceiptNotesItem.FilterItemType.SearchByName, itemName);
            receiptNotesItem.SelectItem(itemName);
            receiptNotesItem.AddReceived(itemName, "10");

            receiptNotesItem.Filter(ReceiptNotesItem.FilterItemType.SearchByName, itemName2);
            receiptNotesItem.SelectItem(itemName2);
            receiptNotesItem.AddReceived(itemName2, "10");

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

            WebDriver.Navigate().Refresh();
            receiptNotesItem.Validate();

            // 2. Create first supplier invoice
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
            supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, true);
            supplierInvoicesCreateModalpage.CreateSIFromRN(receiptNoteNumber);
            var supplierItem = supplierInvoicesCreateModalpage.Submit();

            Assert.AreEqual(2, supplierItem.GetNumberOfItems(), "La supplier invoice créée à partir de la RN n'a pas le même nombre d'items renseignés que la RN.");
            Assert.IsTrue(supplierItem.IsItemFiltered(itemName) && supplierItem.IsItemFiltered(itemName2), "Les items présents dans la SI ne sont pas ceux de la RN.");

            supplierItem.SelectItem(itemName);
            supplierItem.SetItemQuantity(itemName, "0");

            supplierItem.ValidateSupplierInvoice();

            Assert.AreEqual("0", supplierItem.GetItemQuantity(itemName), "La quantité de l'item " + itemName + " n'a pas été mise à 0.");
            Assert.AreNotEqual("0", supplierItem.GetItemQuantity(itemName2), "La quantité de l'item " + itemName2 + " a été mise à 0 à la validation.");

            supplierInvoicesPage = supplierItem.BackToList();
            supplierInvoicesPage.ResetFilter();
            Assert.AreEqual(supplierInvoiceNb, supplierInvoicesPage.GetFirstInvoiceNumber(), "La supply invoice créée n'apparaît pas dans la liste des supply invoices disponibles.");

            // 3. Create second supplier invoice from the same RN
            supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
            supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNbBis, DateUtils.Now, site, supplier, true, true);
            supplierInvoicesCreateModalpage.CreateSIFromRN(receiptNoteNumber);
            supplierItem = supplierInvoicesCreateModalpage.Submit();

            Assert.AreEqual(2, supplierItem.GetNumberOfItems(), "La supplier invoice créée à partir de la RN n'a pas le même nombre d'items renseignés que la RN.");
            Assert.IsTrue(supplierItem.IsItemFiltered(itemName) && supplierItem.IsItemFiltered(itemName2), "Les items présents dans la SI ne sont pas ceux de la RN.");

            Assert.AreEqual("0", supplierItem.GetItemQuantity(itemName2), "La quantité de l'item " + itemName2 + " n'a pas égale à 0 malgré qué sa facture ait déjà été payée dans une autre SI.");
            Assert.AreNotEqual("0", supplierItem.GetItemQuantity(itemName), "La quantité de l'item " + itemName2 + " est égale à 0 alors que sa facture n'a pas été payée.");
        }

        //_____________________________________FIN CREATE SUPPLY INVOICE_____________________________________

        //_____________________________________FILTRES SUPPLY INVOICE________________________________________

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_Filter_SearchByInvoiceNumber()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string quantity = "2";
            // Log in
            HomePage homePage = LogInAsAdmin();
            //Act
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            try
            {
                supplierInvoicesPage.ResetFilter();
                var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
                var supplierItem = supplierInvoicesCreateModalpage.Submit();
                // Ajout d'une nouvelle ligne
                supplierItem.AddNewItem(itemName, quantity);
                supplierInvoicesPage = supplierItem.BackToList();
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
                //Assert
                Assert.AreEqual(1, supplierInvoicesPage.CheckTotalNumber(), String.Format(MessageErreur.FILTRE_ERRONE, "'By invoice number'"));
            }
            finally
            {
                supplierInvoicesPage.DeleteFirstSI();
            }

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_Filter_KeepFilterValue()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();

            string quantity = "2";

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();

            if (supplierInvoicesPage.CheckTotalNumber() == 0)
            {
                var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
                var supplierItemPage = supplierInvoicesCreateModalpage.Submit();

                // Ajout d'une nouvelle ligne
                supplierItemPage.AddNewItem(itemName, quantity);

                supplierInvoicesPage = supplierItemPage.BackToList();
                supplierInvoicesPage.ResetFilter();
            }
            else
            {
                supplierInvoiceNb = supplierInvoicesPage.GetFirstInvoiceNumber();
            }

            supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);

            string oldFilterValue = supplierInvoicesPage.GetSearchFilterValue();
            var result = supplierInvoicesPage.VerifyDataExist();
            Assert.IsTrue(result, "there is no supplier invoices");
            var supplierItem = supplierInvoicesPage.SelectFirstSupplierInvoice();
            supplierInvoicesPage = supplierItem.BackToList();

            //Assert
            Assert.AreEqual(oldFilterValue, supplierInvoicesPage.GetSearchFilterValue(), "La valeur du filtre n'a pas été conservée après avoir été sur la page Details.");

            WebDriver.Navigate().Refresh();
            Assert.AreEqual(oldFilterValue, supplierInvoicesPage.GetSearchFilterValue(), "La valeur du filtre n'a pas été conservée après avoir rechargé la page.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_Filter_BySupplier()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();

            string quantity = "2";

            // Log in
            HomePage homePage = LogInAsAdmin();

            //Act
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();
            supplierInvoicesPage.Filter(SupplierInvoicesPage.FilterType.Suppliers, supplier);

            if (supplierInvoicesPage.CheckTotalNumber() < 20)
            {
                // Create
                var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
                var supplierItem = supplierInvoicesCreateModalpage.Submit();

                supplierItem.AddNewItem(itemName, quantity);

                supplierInvoicesPage = supplierItem.BackToList();
                supplierInvoicesPage.ResetFilter();
                supplierInvoicesPage.Filter(SupplierInvoicesPage.FilterType.Suppliers, supplier);
            }

            if (!supplierInvoicesPage.isPageSizeEqualsTo100())
            {
                supplierInvoicesPage.PageSize("8");
                supplierInvoicesPage.PageSize("100");
            }
            bool isVerifySupplier = supplierInvoicesPage.VerifySupplier(supplier);
            //Assert
            Assert.IsTrue(isVerifySupplier, String.Format(MessageErreur.FILTRE_ERRONE, "'By supplier'"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_Filter_SortByInvoiceDate()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string quantity = "2";
            DateTime date = DateUtils.Now;
            bool isActivate = true;
            bool creatForm = false;
            int pageSize = 100;
            string filter = "DATE";

            // Log in
            HomePage homePage = LogInAsAdmin();

            // On récupère le format de date utilisé dans les date picker
            string dateFormat = homePage.GetDateFormatPickerValue();

            //Act
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();

            if (supplierInvoicesPage.CheckTotalNumber() < 20)
            {
                // Create
                SupplierInvoicesCreateModalPage supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, date, site, supplier, isActivate, creatForm);
                SupplierInvoicesItem supplierItem = supplierInvoicesCreateModalpage.Submit();

                supplierItem.AddNewItem(itemName, quantity);
                supplierInvoicesPage = supplierItem.BackToList();
            }

            if (!supplierInvoicesPage.isPageSizeEqualsTo100())
            {
                supplierInvoicesPage.PageSize("8");
                supplierInvoicesPage.PageSize("100");
            }
            bool isPageEqualTo100 = supplierInvoicesPage.isPageSizeEqualsTo100();
            Assert.IsTrue(isPageEqualTo100, "page size not set to 100");
            // Filter Sort by Date
            supplierInvoicesPage.Filter(FilterType.SortBy, filter);
            bool isSortedByDate = supplierInvoicesPage.IsSortedByDate(dateFormat, pageSize);
            // Assert
            Assert.IsTrue(isSortedByDate, String.Format(MessageErreur.FILTRE_ERRONE, "'Sort by invoice date'"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_Filter_SortByInvoiceNumber()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string quantity = "2";
            DateTime date = DateUtils.Now;
            bool isActivate = true;
            bool creatForm = false;
            int pageSize = 100;
            string filter = "NUMBER";

            // Log in
            HomePage homePage = LogInAsAdmin();

            // On récupère le format de date utilisé dans les date picker
            string dateFormat = homePage.GetDateFormatPickerValue();

            //Act
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();

            if (supplierInvoicesPage.CheckTotalNumber() < 20)
            {
                // Create
                SupplierInvoicesCreateModalPage supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, date, site, supplier, isActivate, creatForm);
                SupplierInvoicesItem supplierItem = supplierInvoicesCreateModalpage.Submit();

                supplierItem.AddNewItem(itemName, quantity);
                supplierInvoicesPage = supplierItem.BackToList();
            }

            if (!supplierInvoicesPage.isPageSizeEqualsTo100())
            {
                supplierInvoicesPage.PageSize("8");
                supplierInvoicesPage.PageSize("100");
            }
            bool isPageEqualTo100 = supplierInvoicesPage.isPageSizeEqualsTo100();
            Assert.IsTrue(isPageEqualTo100, "page size not set to 100");
            // Filter Sort by  number
            supplierInvoicesPage.Filter(FilterType.SortBy, filter);
            bool isSortedByNumber = supplierInvoicesPage.IsSortedByNumber(pageSize);

            // Assert
            Assert.IsTrue(isSortedByNumber, String.Format(MessageErreur.FILTRE_ERRONE, "'Sort by invoice number'"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_Filter_SortById()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string quantity = "2";
            DateTime date = DateUtils.Now;
            bool isActivate = true;
            bool creatForm = false;
            int pageSize = 100;
            string filter = "ID";

            // Log in
            HomePage homePage = LogInAsAdmin();

            // On récupère le format de date utilisé dans les date picker
            string dateFormat = homePage.GetDateFormatPickerValue();

            //Act
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();

            if (supplierInvoicesPage.CheckTotalNumber() < 20)
            {
                // Create
                SupplierInvoicesCreateModalPage supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, date, site, supplier, isActivate, creatForm);
                SupplierInvoicesItem supplierItem = supplierInvoicesCreateModalpage.Submit();

                supplierItem.AddNewItem(itemName, quantity);
                supplierInvoicesPage = supplierItem.BackToList();
            }

            if (!supplierInvoicesPage.isPageSizeEqualsTo100())
            {
                supplierInvoicesPage.PageSize("8");
                supplierInvoicesPage.PageSize("100");
            }
            bool isPageEqualTo100 = supplierInvoicesPage.isPageSizeEqualsTo100();
            Assert.IsTrue(isPageEqualTo100, "page size not set to 100");
            // Filter Sort by  id
            supplierInvoicesPage.Filter(FilterType.SortBy, filter);
            bool isSortedById = supplierInvoicesPage.IsSortedById(pageSize);
            // Assert
            Assert.IsTrue(isSortedById, String.Format(MessageErreur.FILTRE_ERRONE, "'Sort by id'"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_Filter_DateFrom()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb1 = "SupplierInvoice1 - " + rnd.Next().ToString();
            string supplierInvoiceNb2 = "SupplierInvoice2 - " + rnd.Next().ToString();
            string supplierInvoiceNb3 = "SupplierInvoice3 - " + rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();

            try
            {
                //Create supplier Invoice1
                var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb1, DateUtils.Now, site, supplier, true, false);
                var supplierItem = supplierInvoicesCreateModalpage.Submit();
                supplierItem.BackToList();
                string idNumber1 = supplierInvoicesPage.GetFirstSIIdNumber();

                //Create supplier Invoice2
                supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb2, DateUtils.Now.AddDays(5), site, supplier, true, false);
                supplierItem = supplierInvoicesCreateModalpage.Submit();
                supplierItem.BackToList();
                string idNumber2 = supplierInvoicesPage.GetFirstSIIdNumber();

                //Create supplier Invoice3
                supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb3, DateUtils.Now.AddDays(-5), site, supplier, true, false);
                supplierItem = supplierInvoicesCreateModalpage.Submit();
                supplierItem.BackToList();
                string idNumber3 = supplierInvoicesPage.GetFirstSIIdNumber();

                //Appliquer le filtrage sur Date
                supplierInvoicesPage.Filter(FilterType.DateFrom, DateUtils.Now);
                supplierInvoicesPage.Filter(FilterType.DateTo, DateUtils.Now.AddDays(8));

                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb1);
                int totalNumber = supplierInvoicesPage.CheckTotalNumber();
                //Assert
                Assert.AreEqual(totalNumber, 1, "Le filtre DateFrom ne fonctionne pas correctement.");

                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb2);
                totalNumber = supplierInvoicesPage.CheckTotalNumber();
                Assert.AreEqual(totalNumber, 1, "Le filtre DateFrom ne fonctionne pas correctement.");

                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb3);
                totalNumber = supplierInvoicesPage.CheckTotalNumber();
                Assert.AreEqual(totalNumber, 0, "Le filtre DateFrom ne fonctionne pas correctement.");
            }
            finally
            {
                supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                supplierInvoicesPage.ResetFilter();
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb1);
                supplierInvoicesPage.DeleteFirstSI();

                supplierInvoicesPage.ResetFilter();
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb2);
                supplierInvoicesPage.DeleteFirstSI();

                supplierInvoicesPage.ResetFilter();
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb3);
                supplierInvoicesPage.DeleteFirstSI();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_Filter_DateTo()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb1 = "SupplierInvoice1 - " + rnd.Next().ToString();
            string supplierInvoiceNb2 = "SupplierInvoice2 - " + rnd.Next().ToString();
            string supplierInvoiceNb3 = "SupplierInvoice3 - " + rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();

            try
            {
                //Create supplier Invoice1
                var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb1, DateUtils.Now, site, supplier, true, false);
                var supplierItem = supplierInvoicesCreateModalpage.Submit();
                supplierItem.BackToList();
                string idNumber1 = supplierInvoicesPage.GetFirstSIIdNumber();

                //Create supplier Invoice2
                supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb2, DateUtils.Now.AddDays(5), site, supplier, true, false);
                supplierItem = supplierInvoicesCreateModalpage.Submit();
                supplierItem.BackToList();
                string idNumber2 = supplierInvoicesPage.GetFirstSIIdNumber();

                //Create supplier Invoice3
                supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb3, DateUtils.Now.AddDays(-5), site, supplier, true, false);
                supplierItem = supplierInvoicesCreateModalpage.Submit();
                supplierItem.BackToList();
                string idNumber3 = supplierInvoicesPage.GetFirstSIIdNumber();

                //Appliquer le filtrage sur Date
                supplierInvoicesPage.Filter(FilterType.DateFrom, DateUtils.Now.AddDays(-8));
                supplierInvoicesPage.Filter(FilterType.DateTo, DateUtils.Now);

                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb1);
                int totalNumber = supplierInvoicesPage.CheckTotalNumber();
                //Assert
                Assert.AreEqual(totalNumber, 1, "Le filtre DateTo ne fonctionne pas correctement.");

                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb2);
                totalNumber = supplierInvoicesPage.CheckTotalNumber();
                Assert.AreEqual(totalNumber, 0, "Le filtre DateTo ne fonctionne pas correctement.");

                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb3);
                totalNumber = supplierInvoicesPage.CheckTotalNumber();
                Assert.AreEqual(totalNumber, 1, "Le filtre DateTo ne fonctionne pas correctement.");
            }
            finally
            {
                supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                supplierInvoicesPage.ResetFilter();
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb1);
                supplierInvoicesPage.DeleteFirstSI();

                supplierInvoicesPage.ResetFilter();
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb2);
                supplierInvoicesPage.DeleteFirstSI();

                supplierInvoicesPage.ResetFilter();
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb3);
                supplierInvoicesPage.DeleteFirstSI();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_Filter_ShowAll()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();

            string quantity = "2";

            // Log in
            HomePage homePage = LogInAsAdmin();
            //Act
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();
            supplierInvoicesPage.Filter(FilterType.ShowAll, true);

            if (supplierInvoicesPage.CheckTotalNumber() < 20)
            {
                // Create
                var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, false, false);
                var supplierItem = supplierInvoicesCreateModalpage.Submit();

                supplierItem.AddNewItem(itemName, quantity);
                supplierInvoicesPage = supplierItem.BackToList();
            }

            // Search for the created inventory
            supplierInvoicesPage.Filter(FilterType.ShowOnlyInactive, true);
            int nbInactive = supplierInvoicesPage.CheckTotalNumber();

            supplierInvoicesPage.Filter(FilterType.ShowOnlyActive, true);
            int nbActive = supplierInvoicesPage.CheckTotalNumber();

            supplierInvoicesPage.Filter(FilterType.ShowAll, true);
            int nbTotal = supplierInvoicesPage.CheckTotalNumber();

            //Assert
            int ActtiveAndInactiveTota = nbActive + nbInactive;
            Assert.AreEqual(nbTotal, ActtiveAndInactiveTota, String.Format(MessageErreur.FILTRE_ERRONE, "'Show all'"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_Filter_ShowActive()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string quantity = "2";

            // Log in
            HomePage homePage = LogInAsAdmin();

            //Act
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();
            supplierInvoicesPage.Filter(FilterType.ShowOnlyActive, true);
            if (supplierInvoicesPage.CheckTotalNumber() < 20)
            {
                // Create
                var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
                var supplierItem = supplierInvoicesCreateModalpage.Submit();

                supplierItem.AddNewItem(itemName, quantity);
                supplierInvoicesPage = supplierItem.BackToList();
                supplierInvoicesPage.ResetFilter();
                supplierInvoicesPage.Filter(FilterType.ShowOnlyActive, true);
            }
            if (!supplierInvoicesPage.isPageSizeEqualsTo100())
            {
                supplierInvoicesPage.PageSize("8");
                supplierInvoicesPage.PageSize("100");
            }
            //Assert
            bool isShowActiveOk = supplierInvoicesPage.CheckStatus(true);
            Assert.IsTrue(isShowActiveOk, String.Format(MessageErreur.FILTRE_ERRONE, "'Show only active'"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_Filter_ShowInactive()
        {
            // Log in
            HomePage homePage = LogInAsAdmin();

            //Act
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();
            // Récupération du nombre de factures avec différents filtres
            supplierInvoicesPage.Filter(FilterType.ShowAll, true);
            var numberAllInvoice = supplierInvoicesPage.CheckTotalNumber();

            supplierInvoicesPage.Filter(FilterType.ShowOnlyActive, true);
            var numberActiveInvoice = supplierInvoicesPage.CheckTotalNumber();

            supplierInvoicesPage.Filter(FilterType.ShowOnlyInactive, true);
            var numberInvoiceInactive = supplierInvoicesPage.CheckTotalNumber();

            // Assertion
            var activeAndInactive = numberActiveInvoice + numberInvoiceInactive;
            Assert.AreEqual(numberAllInvoice, activeAndInactive, String.Format(MessageErreur.FILTRE_ERRONE, "La somme des factures actives et inactives ne correspond pas au total des factures"));
            var checkStatus = supplierInvoicesPage.CheckStatus(false);
            Assert.IsFalse(checkStatus, String.Format(MessageErreur.FILTRE_ERRONE, "'Show only inactive'"));
        }


        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_Filter_ShowImportedWithEDI()
        {
            //Prepare
            DateTime dateFrom = DateTime.Parse("01/01/2024");

            // Log in
            HomePage homePage = LogInAsAdmin();


            //Act
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();
            supplierInvoicesPage.Filter(FilterType.ShowEDI, true);
            supplierInvoicesPage.Filter(FilterType.DateFrom, dateFrom);
            var listSupplierInvoice = supplierInvoicesPage.GetAllInvoiceNumber().ToList();

            //assert
            Assert.IsTrue(supplierInvoicesPage.CheckIfEdiInvoice(listSupplierInvoice), "problem, there is supplier Invoice not from import EDI");
        }

        //[Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_Filter_ShowAllInvoices()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string quantity = "2";

            // Log in
            HomePage homePage = LogInAsAdmin();
            //Act
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();
            var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
            supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
            var supplierItem = supplierInvoicesCreateModalpage.Submit();

            supplierItem.AddNewItem(itemName, quantity);
            supplierInvoicesPage = supplierItem.BackToList();
            supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
            supplierInvoicesPage.Filter(FilterType.ShowNotValidatedOnly, true);
            //Assert
            var notValidatedOnly1 = supplierInvoicesPage.CheckTotalNumber();
            Assert.AreEqual(1, notValidatedOnly1, "Le supplier invoice non validé créé n'apparaît pas dans la liste des supplier invoices non validés.");

            supplierInvoicesPage.Filter(FilterType.ShowSentToSageOnly, true);
            var sentToSageOnly = supplierInvoicesPage.CheckTotalNumber();
            Assert.AreEqual(0, sentToSageOnly, "Le supplier invoice non envoyé vers SAGE apparaît dans la liste des supplier invoices envoyées vers SAGE.");

            supplierInvoicesPage.Filter(FilterType.ShowAllInvoices, true);
            var allInvoices1 = supplierInvoicesPage.CheckTotalNumber();
            Assert.AreEqual(1, allInvoices1, String.Format(MessageErreur.FILTRE_ERRONE, "'Show all invoices'"));

            supplierItem = supplierInvoicesPage.SelectFirstSupplierInvoice();
            supplierItem.ValidateSupplierInvoice();
            supplierInvoicesPage = supplierItem.BackToList();
            supplierInvoicesPage.ResetFilter();

            //Assert
            supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
            supplierInvoicesPage.Filter(FilterType.ShowNotValidatedOnly, true);
            supplierInvoicesPage.WaitPageLoading();
            var notValidatedOnly2 = supplierInvoicesPage.CheckTotalNumber();
            Assert.AreEqual(0, notValidatedOnly2, "Le supplier invoice validé créé apparaît dans la liste des supplier invoices non validés.");

            supplierInvoicesPage.Filter(FilterType.ShowValidatedNotSentSage, true);
            var showValidatedNotSentSage = supplierInvoicesPage.CheckTotalNumber();
            Assert.AreEqual(1, showValidatedNotSentSage, "Le supplier invoice validé créé n'apparaît pas dans la liste des supplier invoices validés.");

            supplierInvoicesPage.Filter(FilterType.ShowAllInvoices, true);
            var allInvoices2 = supplierInvoicesPage.CheckTotalNumber();
            Assert.AreEqual(1, allInvoices2, String.Format(MessageErreur.FILTRE_ERRONE, "'Show all invoices'"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_Filter_ShowNotValidated()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();

            string quantity = "2";

            // Log in
            HomePage homePage = LogInAsAdmin();
            //Act
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();
            supplierInvoicesPage.Filter(FilterType.ShowNotValidatedOnly, true);

            if (supplierInvoicesPage.CheckTotalNumber() < 20)
            {
                // Create
                var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
                var supplierItem = supplierInvoicesCreateModalpage.Submit();
                supplierItem.AddNewItem(itemName, quantity);

                supplierInvoicesPage = supplierItem.BackToList();
                supplierInvoicesPage.ResetFilter();
                supplierInvoicesPage.Filter(FilterType.ShowNotValidatedOnly, true);
            }

            if (!supplierInvoicesPage.isPageSizeEqualsTo100())
            {
                supplierInvoicesPage.PageSize("8");
                supplierInvoicesPage.PageSize("100");
            }

            //Assert
            Assert.IsTrue(supplierInvoicesPage.checkNotValidateOnly(), String.Format(MessageErreur.FILTRE_ERRONE, "'Show not validated only'"));
        }

        [Ignore]
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_Filter_ShowSenttoSAGE()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string taxType = TestContext.Properties["TaxTypeSupplierInvoices"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string journalSI = TestContext.Properties["Journal_SI"].ToString();

            string quantity = "2";

            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            //Supplier Invoices List Report_-_446332_-_20220620125426.pdf
            string DocFileNamePdfBegin = "Supplier Invoices List Report_-_";
            //All_files_20220225_102148.zip
            string DocFileNameZipBegin = "All_files_";

            var newVersionPrint = true;

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var decimalSeparator = homePage.GetDecimalSeparatorValue();

            homePage.SetSageAutoEnabled(site, false);

            //Act
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();

            supplierInvoicesPage.ClearDownloads();

            supplierInvoicesPage.Filter(FilterType.ShowSentToSageOnly, true);

            if (supplierInvoicesPage.CheckTotalNumber() < 20)
            {
                // Manipulation pour permettre export SAGE 
                var accountingParametersPage = homePage.GoToParameters_AccountingPage();
                accountingParametersPage.GoToTab_Journal();
                accountingParametersPage.EditJournal(site, null, journalSI);

                // Create supplier invoice
                supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                supplierInvoicesPage.ResetFilter();

                var SupplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                SupplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
                var supplierItem = SupplierInvoicesCreateModalpage.Submit();

                supplierItem.AddNewItem(itemName, quantity, taxType);
                supplierItem.ValidateSupplierInvoice();

                // Export vers SAGE
                supplierItem.ManualExportSage(newVersionPrint);
                supplierInvoicesPage = supplierItem.BackToList();
                supplierInvoicesPage.ResetFilter();
                supplierInvoicesPage.Filter(FilterType.ShowSentToSageOnly, true);
            }

            if (!supplierInvoicesPage.isPageSizeEqualsTo100())
            {
                supplierInvoicesPage.PageSize("8");
                supplierInvoicesPage.PageSize("100");
            }

            //Assert
            Assert.IsTrue(supplierInvoicesPage.IsSentToSAGE(), String.Format(MessageErreur.FILTRE_ERRONE, "'Show sent to SAGE only'"));

            //Print/Export Pdf
            supplierInvoicesPage.ClearDownloads();
            supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();
            supplierInvoicesPage.Filter(FilterType.ShowSentToSageOnly, true);
            supplierInvoicesPage.PageSize("100");

            //print
            PrintReportPage reportPage = supplierInvoicesPage.PrintSupplierInvoices(true);
            var isReportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();
            Assert.IsTrue(isReportGenerated, "pas de document Print généré");
            string directory = TestContext.Properties["DownloadsPath"].ToString();
            foreach (string f in Directory.GetFiles(directory))
            {
                //supplier-invoices 2022-06-22 12-27-06.xlsx
                if (f.Contains("supplier-invoices") && (f.EndsWith(".xlsx") || f.EndsWith(".xslx")))
                {
                    File.Delete(f);
                }
            }
            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);

            //export
            supplierInvoicesPage.ExportExcelFile(true);
            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();
            FileInfo trouveXLSX = supplierInvoicesPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(trouveXLSX, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

            // zip
            // cliquer sur All
            string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin, false);
            FileInfo filePdf = new FileInfo(trouve);
            filePdf.Refresh();
            Assert.IsTrue(filePdf.Exists, trouve + " non généré");

            //Vérifier que le fichier Print/Export n'est pas vide et que les données correspondent
            supplierInvoicesPage.CheckExport(trouveXLSX, decimalSeparator);
            supplierInvoicesPage.CheckPrint(filePdf);
        }

        [Ignore]//sage auto
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_Filter_SentToSAGEInErrorOnly()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string taxName = TestContext.Properties["TaxTypeSupplierInvoicesExportSage"].ToString();
            string journalSI = TestContext.Properties["Journal_SI"].ToString();

            string quantity = "2";

            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            //Supplier Invoices List Report_-_446332_-_20220620125426.pdf
            string DocFileNamePdfBegin = "Supplier Invoices List Report_-_";
            //All_files_20220225_102148.zip
            string DocFileNameZipBegin = "All_files_";

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var decimalSeparator = homePage.GetDecimalSeparatorValue();

            try
            {
                // 1. Config Export Sage auto pour créer la facture
                homePage.SetSageAutoEnabled(site, true, "Supplier Invoice");

                // Parameter - Accounting --> Journal
                VerifyAccountingJournal(homePage, site, "");

                //Act
                var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                supplierInvoicesPage.ClearDownloads();
                supplierInvoicesPage.ResetFilter();
                supplierInvoicesPage.Filter(FilterType.ShowSentToSageErrorOnly, true);

                if (supplierInvoicesPage.CheckTotalNumber() < 20)
                {
                    // Create
                    var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                    supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
                    var supplierItem = supplierInvoicesCreateModalpage.Submit();

                    supplierItem.AddNewItem(itemName, quantity, taxName);
                    supplierItem.ValidateSupplierInvoice();

                    var supplierInvoiceAccounting = supplierItem.ClickOnAccounting();
                    Assert.AreNotEqual("", supplierInvoiceAccounting.GetErrorMessage(), "Le code journal est manquant mais aucun message d'erreur n'est présent dans l'onglet Accounting.");

                    supplierInvoicesPage = supplierInvoiceAccounting.BackToList();
                    supplierInvoicesPage.ResetFilter();
                    supplierInvoicesPage.Filter(FilterType.ShowSentToSageErrorOnly, true);
                }

                if (!supplierInvoicesPage.isPageSizeEqualsTo100())
                {
                    supplierInvoicesPage.PageSize("8");
                    supplierInvoicesPage.PageSize("100");
                }

                Assert.IsTrue(supplierInvoicesPage.IsSentToSAGEInErrorOnly(), String.Format(MessageErreur.FILTRE_ERRONE, "'Show sent to SAGE and in error only'"));
                supplierInvoicesPage.PageSize("100");

                //print
                PrintReportPage reportPage = supplierInvoicesPage.PrintSupplierInvoices(true);
                var isReportGenerated = reportPage.IsReportGenerated();
                reportPage.Close();
                Assert.IsTrue(isReportGenerated, "pas de document Print généré");

                // purger le dossier Download
                string directory = TestContext.Properties["DownloadsPath"].ToString();
                foreach (string f in Directory.GetFiles(directory))
                {
                    //supplier-invoices 2022-06-22 12-27-06.xlsx
                    if (f.Contains("supplier-invoices") && (f.EndsWith(".xlsx") || f.EndsWith(".xlsx")))
                    {
                        File.Delete(f);
                    }
                }
                reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
                // remove popup print service
                WebDriver.Navigate().Refresh();
                supplierInvoicesPage.ResetFilter();
                supplierInvoicesPage.Filter(FilterType.ShowSentToSageErrorOnly, true);
                supplierInvoicesPage.PageSize("100");

                //export
                supplierInvoicesPage.ExportExcelFile(true);
                // On récupère les fichiers du répertoire de téléchargement
                DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
                FileInfo[] taskFiles = taskDirectory.GetFiles();
                FileInfo trouveXLSX = supplierInvoicesPage.GetExportExcelFile(taskFiles);
                Assert.IsNotNull(trouveXLSX, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

                //zip
                // le xlsx dans le All_files*.zip a un nom bizarre
                string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin, false);
                FileInfo trouvePdf = new FileInfo(trouve);

                //Vérifier que le fichier Print/Export n'est pas vide et que les données correspondent
                supplierInvoicesPage.CheckExport(trouveXLSX, decimalSeparator);
                supplierInvoicesPage.CheckPrint(trouvePdf);
            }
            finally
            {
                VerifyAccountingJournal(homePage, site, journalSI);
                homePage.SetSageAutoEnabled(site, false);
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_Filter_ShowOnTLWaitingForSAGEPush()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string taxName = TestContext.Properties["TaxTypeSupplierInvoicesExportSage"].ToString();

            string quantity = "2";

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Config Export Sage Auto
            homePage.SetSageAutoEnabled(site, true, "Supplier Invoice");

            try
            {
                //Act
                var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                supplierInvoicesPage.ResetFilter();
                supplierInvoicesPage.Filter(FilterType.ShowOnTLWaitingForSagePush, true);

                if (supplierInvoicesPage.CheckTotalNumber() < 20)
                {
                    // Create
                    var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                    supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
                    var supplierItem = supplierInvoicesCreateModalpage.Submit();

                    supplierItem.AddNewItem(itemName, quantity, taxName);
                    supplierItem.ValidateSupplierInvoice();

                    supplierInvoicesPage = supplierItem.BackToList();
                    supplierInvoicesPage.ResetFilter();
                    supplierInvoicesPage.Filter(FilterType.ShowOnTLWaitingForSagePush, true);
                }

                if (!supplierInvoicesPage.isPageSizeEqualsTo100())
                {
                    supplierInvoicesPage.PageSize("8");
                    supplierInvoicesPage.PageSize("100");
                }

                Assert.IsTrue(supplierInvoicesPage.IsWaitingForSAGEPush(), String.Format(MessageErreur.FILTRE_ERRONE, "'Show on TL waiting for SAGE push'"));
            }
            finally
            {
                homePage.SetSageAutoEnabled(site, false);
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_Filter_ShowValidatedNotSentToSage()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string taxType = TestContext.Properties["TaxTypeSupplierInvoices"].ToString();

            string quantity = "2";

            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            //Supplier Invoices List Report_-_446332_-_20220620125426.pdf
            string DocFileNamePdfBegin = "Supplier Invoices List Report_-_";
            //All_files_20220225_102148.zip
            string DocFileNameZipBegin = "All_files_";

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var decimalSeparator = homePage.GetDecimalSeparatorValue();

            //Act
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ClearDownloads();
            supplierInvoicesPage.ResetFilter();
            supplierInvoicesPage.Filter(FilterType.ShowValidatedNotSentSage, true);

            if (supplierInvoicesPage.CheckTotalNumber() < 20)
            {
                // Create
                var SupplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                SupplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
                var supplierItem = SupplierInvoicesCreateModalpage.Submit();

                supplierItem.AddNewItem(itemName, quantity, taxType);
                supplierItem.ValidateSupplierInvoice();
                supplierInvoicesPage = supplierItem.BackToList();
                supplierInvoicesPage.ResetFilter();
                supplierInvoicesPage.Filter(FilterType.ShowValidatedNotSentSage, true);
            }

            if (!supplierInvoicesPage.isPageSizeEqualsTo100())
            {
                supplierInvoicesPage.PageSize("8");
                supplierInvoicesPage.PageSize("100");
            }

            //Assert
            Assert.IsTrue(supplierInvoicesPage.CheckValidation(true), String.Format(MessageErreur.FILTRE_ERRONE, "'Show validated'"));
            Assert.IsFalse(supplierInvoicesPage.IsSentToSAGE(), String.Format(MessageErreur.FILTRE_ERRONE, "'Show sent to SAGE'"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_Filter_ShowVerifiedOnly()
        {
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();

            string quantity = "2";

            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            //Supplier Invoices List Report_-_446332_-_20220620125426.pdf
            string DocFileNamePdfBegin = "Supplier Invoices List Report_-_";
            //All_files_20220225_102148.zip
            string DocFileNameZipBegin = "All_files_";


            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var decimalSeparator = homePage.GetDecimalSeparatorValue();

            //Act
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();

            // Create
            var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
            supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
            var supplierItem = supplierInvoicesCreateModalpage.Submit();

            supplierItem.AddNewItem(itemName, quantity);
            supplierInvoicesPage = supplierItem.BackToList();

            supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
            supplierInvoicesPage.Filter(FilterType.ShowVerifiedOnly, true);

            Assert.AreEqual(0, supplierInvoicesPage.CheckTotalNumber(), "La supplier invoice n'est pas vérifiée mais apparaît dans le filtre 'Verified only'.");

            WebDriver.Navigate().Refresh();
            supplierInvoicesPage.ResetFilter();
            supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);

            supplierItem = supplierInvoicesPage.SelectFirstSupplierInvoice();
            supplierItem.SetVerified();

            //Assert
            Assert.IsTrue(supplierItem.IsVerified(), "L'item n'est pas passé au statut 'Verified'");

            supplierInvoicesPage = supplierItem.BackToList();

            supplierInvoicesPage.ClearDownloads();

            //Appliquer les filtres sur show verified only
            supplierInvoicesPage.ResetFilter();
            supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
            supplierInvoicesPage.Filter(FilterType.ShowVerifiedOnly, true);
            // Vérifier que les résultats s'accordent bien au filtre appliqué
            Assert.AreEqual(supplierInvoiceNb, supplierInvoicesPage.GetFirstInvoiceNumber(), String.Format(MessageErreur.FILTRE_ERRONE, "'Show verified only'"));

            //Print/Export Pdf
            WebDriver.Navigate().Refresh();
            supplierInvoicesPage.ResetFilter();
            supplierInvoicesPage.Filter(FilterType.ShowVerifiedOnly, true);
            supplierInvoicesPage.PageSize("100");

            //print
            PrintReportPage reportPage = supplierInvoicesPage.PrintSupplierInvoices(true);
            var isReportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();
            Assert.IsTrue(isReportGenerated, "pas de document Print généré");
            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);

            // remove popup print service + scroll
            //WebDriver.Navigate().Refresh();
            supplierInvoicesPage.ClosePrintButton();
            supplierInvoicesPage.PageUp();
            supplierInvoicesPage.ResetFilter();
            supplierInvoicesPage.Filter(FilterType.ShowVerifiedOnly, true);
            supplierInvoicesPage.PageSize("100");

            //export
            string directory = TestContext.Properties["DownloadsPath"].ToString();
            foreach (string f in Directory.GetFiles(directory))
            {
                //supplier-invoices 2022-06-22 12-27-06.xlsx
                if (f.Contains("supplier-invoices") && (f.EndsWith(".xlsx") || f.EndsWith(".xslx")))
                {
                    File.Delete(f);
                }
            }
            supplierInvoicesPage.ExportExcelFile(true);
            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();
            FileInfo trouveXLSX = supplierInvoicesPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(trouveXLSX, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

            // zip
            // cliquer sur All
            string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin, false);
            FileInfo filePdf = new FileInfo(trouve);
            filePdf.Refresh();
            Assert.IsTrue(filePdf.Exists, trouve + " non généré");

            //Vérifier que le fichier Print/Export n'est pas vide et que les données correspondent
            supplierInvoicesPage.CheckExport(trouveXLSX, decimalSeparator);
            supplierInvoicesPage.CheckPrint(filePdf);

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_Filter_ShowSupplierWithClaim()
        {
            string DocFileNamePdfBegin = "Supplier Invoices List Report_-_";
            //All_files_20220225_102148.zip
            string DocFileNameZipBegin = "All_files_";
            //ce filtre ne fonctionne que sur les SI où la claim a été créée dedans (mégaphone rouge) et pas les SI créée d'une claim
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var decimalSeparator = homePage.GetDecimalSeparatorValue();

            //Act
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ClearDownloads();
            supplierInvoicesPage.ResetFilter();
            supplierInvoicesPage.Filter(FilterType.ShowSupplierInvoiceWithClaim, true);
            if (!supplierInvoicesPage.isPageSizeEqualsTo100())
            {
                supplierInvoicesPage.PageSize("100");
            }

            Assert.IsTrue(supplierInvoicesPage.IsWithClaim(), String.Format(MessageErreur.FILTRE_ERRONE, "'Show supplier invoice with claim'"));
            var supplierswithclaimsnumber = supplierInvoicesPage.GetInvoiceSupplierNumbers().Count();
            //clean download directory
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            foreach (var file in taskDirectory.GetFiles())
            {
                file.Delete();
            }
            // Export Excel

            supplierInvoicesPage.ExportExcelFile(true);
            FileInfo[] taskFiles = taskDirectory.GetFiles();
            FileInfo fileExcel = supplierInvoicesPage.GetExportExcelFile(taskFiles);
            Assert.IsTrue(fileExcel.Exists, "Fichier Excel non trouvé");
            var numbersInExcel = OpenXmlExcel.GetValuesInList("Id", "Invoices", fileExcel.FullName).Count;
            Assert.AreEqual(numbersInExcel, supplierswithclaimsnumber, "le nombre des suppliers with claim dans le fichier excel est ");


            // Print Pdf
            supplierInvoicesPage.ClearDownloads();
            supplierInvoicesPage.Filter(FilterType.ShowSupplierInvoiceWithClaim, true);
            if (!supplierInvoicesPage.isPageSizeEqualsTo100())
            {
                supplierInvoicesPage.PageSize("100");
            }
            PrintReportPage reportPage = supplierInvoicesPage.PrintSupplierInvoices(true);
            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            var isReportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();
            Assert.IsTrue(isReportGenerated, "pas de document Print généré");
            //
            // cliquer sur All
            string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin, false);
            FileInfo filePdf = new FileInfo(trouve);
            filePdf.Refresh();
            PdfDocument document = PdfDocument.Open(filePdf.FullName);
            int nbLignes = 0;
            foreach (Page page in document.GetPages())
            {
                IEnumerable<Word> words = page.GetWords();
                nbLignes += words.Count(w => w.Text == "Piece");
            }
            // moins le titre de la table
            nbLignes -= 1;
            Assert.AreEqual(nbLignes, supplierswithclaimsnumber, "le nombre des suppliers with claim dans le fichier pdf est différent du nombre a l'affichage");
            supplierInvoicesPage.CheckExport(fileExcel, decimalSeparator);
            supplierInvoicesPage.CheckPrint(filePdf);
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_CreateClaimFromSupplierInvoice()
        {

            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string quantity = "2";
            string dnPrice = "80";
            string dnQuantity = "1";

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var decimalSeparator = homePage.GetDecimalSeparatorValue();

            //Act
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();
            supplierInvoicesPage.Filter(FilterType.ShowSupplierInvoiceWithClaim, true);


            // 1.Create SI
            var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
            supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
            var supplierItem = supplierInvoicesCreateModalpage.Submit();


            supplierItem.AddNewItem(itemName, quantity);
            supplierItem.SelectItem(itemName);
            string priceSI = supplierItem.GetSIPrice();
            supplierItem.SetDNPrice(dnPrice, itemName);
            supplierItem.SetDNQuantity(dnQuantity, itemName);

            // 2.Add Claim from SI item
            var newClaimFromSI = supplierItem.EditClaimForm(itemName);
            newClaimFromSI.SetClaimFormForClaimedV3();

            //-Vérifier la liaison créé dans General Information de la SI avec le numéro de la claim
            SupplierInvoicesGeneralInformation generalInformation = supplierItem.ClickOnGeneralInformation();
            var claimNumber = generalInformation.GetRelatedClaimNumber();

            var claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();
            claimsPage.Filter(ClaimsPage.FilterType.ByNumber, claimNumber);
            var claimsItem = claimsPage.SelectFirstClaim();

            // 3.On arrive sur l'onglet Claim (dans les RN) et on vérifie les infos (nom de l'item, quantités...)
            Assert.AreEqual(itemName, claimsItem.GetFirstItemName(), "Mauvais claim Name");
            Assert.AreEqual(quantity, claimsItem.GetQuantity(), "Mauvais claim Quantity");
            Assert.AreEqual("€ " + dnPrice + decimalSeparator + "0000", claimsItem.GetDNPrice(), "Mauvais claim DN Price");
            Assert.AreEqual(dnQuantity, claimsItem.GetDNQuantity(), "Mauvais claim DN Quantity");
            claimsItem.SelectFirstItem();
            Assert.AreEqual(priceSI, claimsItem.GetPrice(), "Mauvais claim SI Price");
            // Check comment in claim
            var claimsGeneralInformation = claimsItem.ClickOnGeneralInformation();
            //- Vérifier la liaison créé dans General Information de la Claim avec le numéro de la SI
            Assert.AreEqual(supplierInvoiceNb, claimsGeneralInformation.GetClaimSupplierInvoice());
            Assert.AreEqual($"Created automatically to contain claims for Supplier Invoice [{supplierInvoiceNb}]", claimsGeneralInformation.GetClaimComment(), "La création de la claim depuis la SI n'a pas été commentée dans la claim.");
            // 4.Validate claim
            claimsItem.Validate();

            // 5.Validate SI
            homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();
            supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
            supplierInvoicesPage.SelectFirstSupplierInvoice();
            supplierItem.ValidateSupplierInvoice();
            supplierInvoicesPage = supplierItem.BackToList();

            //- Vérifier l'icone mégaphone sur la SI au niveau de la liste des SI
            Assert.IsTrue(supplierInvoicesPage.IsWithClaim(), String.Format(MessageErreur.FILTRE_ERRONE, "'Show supplier invoice with claim'"));
        }

        private string CreateClaim(HomePage homePage, string site, string supplier, string itemName)
        {
            string placeTo = TestContext.Properties["PlaceTo"].ToString();

            //Act
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();

            // Create
            var claimsCreateModalpage = claimsPage.ClaimsCreatePage();
            claimsCreateModalpage.FillField_CreatNewClaims(DateUtils.Now, site, supplier, placeTo, true);
            String id = claimsCreateModalpage.GetClaimId();
            var claimsItem = claimsCreateModalpage.Submit();

            claimsItem.Filter(PageObjects.Warehouse.Claims.ClaimsItem.FilterItemType.SearchByName, itemName);
            claimsItem.SelectItem(itemName);
            claimsItem.AddQuantityAndType(itemName, 2);

            //Suite a la modif Claim
            var editClaimForm = claimsItem.EditClaimForm(itemName);
            editClaimForm.SetClaimFormForClaimedV3();

            claimsItem.Validate();

            return id;
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_Filter_ExportedForSageManually()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string taxName = TestContext.Properties["TaxTypeSupplierInvoicesExportSage"].ToString();

            string quantity = "2";

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            try
            {
                // 1. Config Export Sage auto pour créer la facture mais pas pour le site MAD
                homePage.SetSageAutoEnabled(site, true, "Supplier Invoice", false);

                //Act
                var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                supplierInvoicesPage.ResetFilter();

                // Create
                var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
                var supplierItem = supplierInvoicesCreateModalpage.Submit();

                supplierItem.AddNewItem(itemName, quantity, taxName);
                supplierItem.ValidateSupplierInvoice();
                supplierInvoicesPage = supplierItem.BackToList();

                supplierInvoicesPage.ResetFilter();
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
                supplierInvoicesPage.Filter(FilterType.ExportedForSageManually, true);

                Assert.AreEqual(0, supplierInvoicesPage.CheckTotalNumber(), "La supplier invoice créée apparaît dans le résultat du filtre 'Exported for SAGE manually'" +
                    " alors qu'elle n'a pas été envoyée vers le SAGE.");

                WebDriver.Navigate().Refresh();
                supplierInvoicesPage.ResetFilter();
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
                supplierItem = supplierInvoicesPage.SelectFirstSupplierInvoice();

                supplierItem.ManualExportSage(true);
                supplierInvoicesPage = supplierItem.BackToList();

                supplierInvoicesPage.ResetFilter();
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
                supplierInvoicesPage.Filter(FilterType.ExportedForSageManually, true);

                //Assert
                Assert.AreEqual(1, supplierInvoicesPage.CheckTotalNumber(), String.Format(MessageErreur.FILTRE_ERRONE, "Exported for sage manually"));
            }
            finally
            {
                homePage.SetSageAutoEnabled(site, false);
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_Filter_Sites()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string quantity = "2";

            // Log in
            var homePage = LogInAsAdmin();

            //Act
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();
            supplierInvoicesPage.Filter(FilterType.Site, site);

            if (supplierInvoicesPage.CheckTotalNumber() < 20)
            {
                // Create
                var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
                var supplierItem = supplierInvoicesCreateModalpage.Submit();
                supplierItem.AddNewItem(itemName, quantity);

                supplierInvoicesPage = supplierItem.BackToList();
                supplierInvoicesPage.ResetFilter();
                supplierInvoicesPage.Filter(FilterType.Site, site);
            }

            //Assert
            bool isAfiltreSiteOk = supplierInvoicesPage.VerifySite(site);
            Assert.IsTrue(isAfiltreSiteOk, String.Format(MessageErreur.FILTRE_ERRONE, "Site"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_Filter_SitePlaces()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string place = TestContext.Properties["Place"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();

            string quantity = "2";
            string sitePlace = site + "-" + place;

            // Log in
            var homePage = LogInAsAdmin();

            //Act
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();

            // Create
            try
            {
                var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false, place);
                var supplierItem = supplierInvoicesCreateModalpage.Submit();

                supplierItem.AddNewItem(itemName, quantity);

                supplierInvoicesPage = supplierItem.BackToList();
                supplierInvoicesPage.ResetFilter();

                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
                supplierInvoicesPage.Filter(FilterType.SitePlaces, sitePlace);

                Assert.AreEqual(supplierInvoiceNb, supplierInvoicesPage.GetFirstInvoiceNumber(), "La supplier invoice n'a pas été créée.");
                Assert.AreEqual(1, supplierInvoicesPage.CheckTotalNumber(), String.Format(MessageErreur.FILTRE_ERRONE, "SitePlaces"));
            }
            finally
            {
                homePage.Navigate();
                supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
                supplierInvoicesPage.DeleteFirstSI();

            }

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_AddItemRow()
        {
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string quantity = "2";
            string itemName2 = TestContext.Properties["Item_SupplierInvoiceGrp"].ToString();
            string taxType2 = TestContext.Properties["TaxTypeSupplierInvoicesExportSage"].ToString();

            // Log in

            var homePage = LogInAsAdmin();

            //Act
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();

            try
            {
                var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
                var supplierItem = supplierInvoicesCreateModalpage.Submit();

                Assert.AreEqual(0, supplierItem.GetNumberOfItems(), "Il y a déjà un item dans le supplier invoice créé.");

                // Ajout d'une nouvelle ligne
                supplierItem.AddNewItem(itemName, quantity);
                Assert.AreEqual(1, supplierItem.GetNumberOfItems(), "L'ajout d'un item au supplier invoice a échoué.");
                Assert.IsTrue(supplierItem.IsItemAdded(itemName), "L'item {0} n'est pas présent dans le tableau.", itemName);
                Assert.AreEqual(quantity, supplierItem.GetItemQuantity(itemName), "La quantité de l'item {0} ajouté dans le tableau n'est pas correct.", itemName);
                Assert.AreEqual(taxType, supplierItem.GetItemVATRate(itemName), "La tax type/VAT rate de l'item {0} ajouté dans le tableau n'est pas correct.", itemName);

            }
            finally
            {
                homePage.Navigate();
                supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
                supplierInvoicesPage.DeleteFirstSI();
            }

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_AddItemRowTestError()
        {
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();

            // Log in
            var homePage = LogInAsAdmin();

            //Act
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();
            try
            {
                // Create
                var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
                var supplierItem = supplierInvoicesCreateModalpage.Submit();

                supplierItem.AddNewItemError();

                Assert.IsTrue(supplierItem.ErrorMessageItemRequired(), "Le message d'erreur concernant le nom de l'item n'est pas affiché dans la pop-up d'ajout d'item.");
                Assert.IsTrue(supplierItem.ErrorMessageTaxRequired(), "Le message d'erreur concernant la taxe à ajouter n'est pas affiché dans la pop-up d'ajout d'item.");
                supplierItem.CloseNewItemModal();

                //Assert
                Assert.AreEqual(0, supplierItem.GetNumberOfItems(), "L'item n'a pas été ajouté au supplier invoice.");
            }

            finally
            {
                homePage.Navigate();
                supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
                supplierInvoicesPage.DeleteFirstSI();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_Items_Filters_PurchaseOrderNumber()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemPO = TestContext.Properties["SupplierInvoiceItem"].ToString();
            string otherItem = TestContext.Properties["Item_SupplierInvoice"].ToString();

            string quantity = "2";

            // Log in
            var homePage = LogInAsAdmin();

            try
            {
                var infosRNetPO = CreateReceiptNotesFromPurchaseOrder(homePage, site, supplier, itemPO);

                //Act
                var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                supplierInvoicesPage.ResetFilter();

                var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, true);
                supplierInvoicesCreateModalpage.CreateSIFromRN(infosRNetPO.RNNumber);
                var supplierItem = supplierInvoicesCreateModalpage.Submit();

                // Add another item
                supplierItem.AddNewItem(otherItem, quantity);

                // Filter by PONumber
                supplierItem.Filter(SupplierInvoicesItem.FilterItemType.ByPurchaseOrderNumber, infosRNetPO.PONumber);

                //Assert
                Assert.IsTrue(supplierItem.IsItemFiltered(itemPO), String.Format(MessageErreur.FILTRE_ERRONE, "'By PO number'"));
                supplierItem.ResetFilter();
            }
            finally
            {
                var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
                supplierInvoicesPage.DeleteFirstSI();

            }

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_Items_Filters_ReceiptNoteNumber()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemPO = TestContext.Properties["SupplierInvoiceItem"].ToString();
            string otherItem = TestContext.Properties["Item_SupplierInvoice"].ToString();

            string quantity = "2";

            // Log in
            var homePage = LogInAsAdmin();

            try
            {
                var infosRNetPO = CreateReceiptNotesFromPurchaseOrder(homePage, site, supplier, itemPO);

                //Act
                var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                supplierInvoicesPage.ResetFilter();

                var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, true);
                supplierInvoicesCreateModalpage.CreateSIFromRN(infosRNetPO.RNNumber);
                var supplierItem = supplierInvoicesCreateModalpage.Submit();

                // Add another item
                supplierItem.AddNewItem(otherItem, quantity);

                // Filter by RNNumber
                supplierItem.Filter(SupplierInvoicesItem.FilterItemType.ByReceiptNoteNumber, infosRNetPO.RNNumber);
                //Assert
                Assert.IsTrue(supplierItem.IsItemFiltered(itemPO), String.Format(MessageErreur.FILTRE_ERRONE, "'By Receipt note number'"));
            }
            finally
            {
                var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
                supplierInvoicesPage.DeleteFirstSI();

            }

        }

        private InfosPOetRN CreateReceiptNotesFromPurchaseOrder(HomePage homePage, string site, string supplier, string itemName)
        {
            string placeTo = TestContext.Properties["PlaceFrom"].ToString();

            InfosPOetRN infos = new InfosPOetRN();

            // Création Purchase Order
            var purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            var createPurchaseOrderPage = purchaseOrdersPage.CreateNewPurchaseOrder();

            createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, placeTo, DateUtils.Now, true);
            var purchaseOrderItemPage = createPurchaseOrderPage.Submit();
            infos.PONumber = purchaseOrderItemPage.GetPurchaseOrderNumber();

            purchaseOrderItemPage.Filter(PurchaseOrderItem.FilterItemType.ByName, itemName);
            purchaseOrderItemPage.SelectFirstItemPo();
            purchaseOrderItemPage.AddQuantity("5");

            purchaseOrderItemPage.Validate();

            // Creation ReceiptNote
            var ReceiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();

            var receiptNotesCreateModalpage = ReceiptNotesPage.ReceiptNotesCreatePage();
            receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, placeTo) { CreateFromPO = true, PONumber = infos.PONumber, PODate = DateUtils.Now });
            var receiptNotesItem = receiptNotesCreateModalpage.Submit();

            receiptNotesItem.SelectFirstItem();
            receiptNotesItem.AddReceived_receipt(itemName, "10");
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
            infos.RNNumber = receiptNotesItem.GetReceiptNoteNumber();
            receiptNotesItem.Validate();

            return infos;
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_Items_Filters_ByGroup()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName1 = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string itemName2 = TestContext.Properties["Item_SupplierInvoiceBis"].ToString();

            string quantity = "2";

            // Log in

            var homePage = LogInAsAdmin();

            // Récupération des groupes des items
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.Filter(ItemPage.FilterType.Search, itemName1);

            var itemGeneralInformationPage = itemPage.ClickOnFirstItem();
            string group1 = itemGeneralInformationPage.GetGroupName();

            itemPage = itemGeneralInformationPage.BackToList();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName2);

            itemGeneralInformationPage = itemPage.ClickOnFirstItem();
            string group2 = itemGeneralInformationPage.GetGroupName();

            // Création supplier invoice
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();
            try
            {
                var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
                var supplierItem = supplierInvoicesCreateModalpage.Submit();

                // Add item
                supplierItem.AddNewItem(itemName1, quantity);
                supplierItem.AddNewItem(itemName2, quantity);
                var numberOfItem = supplierItem.GetNumberOfItems();
                Assert.AreEqual(2, numberOfItem, "L'ajout des items au supplier invoice a échoué.");

                // Filter by group
                supplierItem.Filter(SupplierInvoicesItem.FilterItemType.ByGroup, group1);
                numberOfItem = supplierItem.GetNumberOfItems();
                var isfilteredByGroup = supplierItem.IsFilteredByGroup(group1);
                Assert.AreEqual(1, numberOfItem, "Le filtre des items par group n'a pas fonctionné.");
                Assert.IsTrue(isfilteredByGroup, String.Format(MessageErreur.FILTRE_ERRONE, "'group'"));

                supplierItem.ResetFilter();
                supplierItem.Filter(SupplierInvoicesItem.FilterItemType.ByGroup, group2);
                numberOfItem = supplierItem.GetNumberOfItems();
                isfilteredByGroup = supplierItem.IsFilteredByGroup(group2);
                Assert.AreEqual(1, numberOfItem, "Le filtre des items par group n'a pas fonctionné.");
                Assert.IsTrue(isfilteredByGroup, String.Format(MessageErreur.FILTRE_ERRONE, "'group'"));
            }
            finally
            {
                homePage.Navigate();
                supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
                supplierInvoicesPage.DeleteFirstSI();
            }

        }

        //_____________________________________FIN SUPPLY INVOICE ITEMS _____________________________________
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_VerifySupplierInvoice()
        {
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();

            string quantity = "2";

            // Log in

            var homePage = LogInAsAdmin();

            //Act
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();
            try
            {
                var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
                var supplierItem = supplierInvoicesCreateModalpage.Submit();

                supplierItem.AddNewItem(itemName, quantity);
                supplierItem.SetVerified();

                //Assert
                Assert.IsTrue(supplierItem.IsVerified(), "L'item n'est pas passé au statut 'Verified'");

                supplierInvoicesPage = supplierItem.BackToList();
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
                supplierInvoicesPage.Filter(FilterType.ShowVerifiedOnly, true);

                Assert.AreEqual(supplierInvoiceNb, supplierInvoicesPage.GetFirstInvoiceNumber(), "La supplier invoice au statut 'Verified' n'est pas visible dans la liste.");
            }
            finally
            {
                homePage.Navigate();
                supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
                supplierInvoicesPage.WaitPageLoading();
                supplierInvoicesPage.DeleteFirstSI();
            }

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_ValidateSupplierInvoice()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string quantity = "2";
            // Log in
            HomePage homePage = LogInAsAdmin();
            //Act
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();
            // Create
            var SupplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
            SupplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
            var supplierItem = SupplierInvoicesCreateModalpage.Submit();
            supplierItem.AddNewItem(itemName, quantity);
            supplierItem.ValidateSupplierInvoice();
            supplierInvoicesPage = supplierItem.BackToList();
            supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
            supplierInvoicesPage.Filter(FilterType.ShowValidatedNotSentSage, true);
            //Assert
            Assert.AreEqual(supplierInvoiceNb, supplierInvoicesPage.GetFirstInvoiceNumber(), String.Format(MessageErreur.FILTRE_ERRONE, "'By invoice number'"));
            Assert.IsTrue(supplierInvoicesPage.CheckValidation(true), "la supplier invoice n'a pas été validée.");
        }

        //Mettre à jour les informations générales
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_UpdateGlobalInformations()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();

            string quantity = "2";
            string newInvoiceNumber = new Random().Next(10000, 500000).ToString();

            // Log in
            var homePage = LogInAsAdmin();

            //Act
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();

            var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
            supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
            var supplierItem = supplierInvoicesCreateModalpage.Submit();
            supplierItem.AddNewItem(itemName, quantity);

            var supplierItemGeneralInfo = supplierItem.ClickOnGeneralInformation();
            // invoice number
            supplierItemGeneralInfo.SetSupplierInvoiceNumber(newInvoiceNumber);

            // site (lecture seule)
            Assert.AreEqual(site, supplierItemGeneralInfo.GetSite(), "mauvais site");
            // Supplier (lecture seule)
            Assert.AreEqual(supplier, supplierItemGeneralInfo.GetSupplier(), "mauvais supplier");

            // Site Place (lecture seule)
            Assert.AreEqual("None", supplierItemGeneralInfo.GetPlace(), "mauvais supplier");

            // Invoice Date
            Assert.AreEqual(DateUtils.Now.ToString("dd/MM/yyyy"), supplierItemGeneralInfo.GetDate(), "mauvais supplier");

            // comment

            string initComment = supplierItemGeneralInfo.GetComments();
            supplierItemGeneralInfo.WaitPageLoading();
            supplierItemGeneralInfo.SetComments("This is my new comment.");
            supplierItemGeneralInfo.WaitPageLoading();
            supplierItem = supplierItemGeneralInfo.ClickOnItems();
            supplierItemGeneralInfo = supplierItem.ClickOnGeneralInformation();


            Assert.AreNotEqual(supplierInvoiceNb, supplierItemGeneralInfo.GetSupplierInvoiceNumber(), "La modification de l'invoice number n'a pas été prise en compte.");
            Assert.AreNotEqual(initComment, supplierItemGeneralInfo.GetComments(), "La modification du commentaire n'a pas été prise en compte.");

            supplierInvoicesPage = supplierItemGeneralInfo.BackToList();
            supplierInvoicesPage.ResetFilter();
            supplierInvoicesPage.Filter(FilterType.ByNumber, newInvoiceNumber);

            supplierInvoicesPage.Filter(FilterType.ShowOnlyActive, true);
            Assert.AreEqual(1, supplierInvoicesPage.CheckTotalNumber());
            supplierInvoicesPage.Filter(FilterType.ShowOnlyInactive, true);
            Assert.AreEqual(0, supplierInvoicesPage.CheckTotalNumber());
            supplierInvoicesPage.Filter(FilterType.ShowAll, true);
            Assert.AreEqual(1, supplierInvoicesPage.CheckTotalNumber());

            SupplierInvoicesItem firstSupplier = supplierInvoicesPage.SelectFirstSupplierInvoice();
            supplierItemGeneralInfo = firstSupplier.ClickOnGeneralInformation();

            supplierItemGeneralInfo.SetActive(false);

            supplierItem = supplierItemGeneralInfo.ClickOnItems();
            supplierItemGeneralInfo = supplierItem.ClickOnGeneralInformation();


            supplierInvoicesPage = supplierItemGeneralInfo.BackToList();
            supplierInvoicesPage.ResetFilter();
            supplierInvoicesPage.Filter(FilterType.ByNumber, newInvoiceNumber);

            supplierInvoicesPage.Filter(FilterType.ShowOnlyActive, true);
            Assert.AreEqual(0, supplierInvoicesPage.CheckTotalNumber());
            supplierInvoicesPage.Filter(FilterType.ShowOnlyInactive, true);
            Assert.AreEqual(1, supplierInvoicesPage.CheckTotalNumber());
            supplierInvoicesPage.Filter(FilterType.ShowAll, true);
            Assert.AreEqual(1, supplierInvoicesPage.CheckTotalNumber());

            Assert.AreEqual(newInvoiceNumber, supplierInvoicesPage.GetFirstInvoiceNumber(), "L'invoice number recherchée n'a pas été trouvée.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_RefreshSupplierInvoice()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();

            string quantity = "2";

            // Log in

            var homePage = LogInAsAdmin();

            //Act
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();

            try
            {
                var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
                var supplierItem = supplierInvoicesCreateModalpage.Submit();
                supplierItem.AddNewItem(itemName, quantity);
                var numberOfItemBeforeRefresh = supplierItem.GetNumberOfItems();

                supplierItem.Refresh();

                var numberOfItemAfterRefresh = supplierItem.GetNumberOfItems();
                var quantityAfterRefresh = supplierItem.GetfirstQuantity();
                //Assert
                Assert.AreEqual(numberOfItemBeforeRefresh, numberOfItemAfterRefresh, "La fonctionnalité 'Refresh' n'a pas fonctionné, les données ne sont pas conservées.");
                Assert.AreEqual(quantityAfterRefresh, quantity, "La fonctionnalité 'Refresh' n'a pas fonctionné, la quantité n'a pas changé.");
            }
            finally
            {
                homePage.Navigate();
                supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
                supplierInvoicesPage.DeleteFirstSI();
            }

        }

        [Ignore]
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_SageManual_ExportSAGEItemsKONewVersion()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string taxType = TestContext.Properties["TaxTypeSupplierInvoicesExportSage"].ToString();
            string journalSI = TestContext.Properties["Journal_SI"].ToString();

            string quantity = "2";
            bool newVersionPrint = true;

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            homePage.SetSageAutoEnabled(site, false);

            // Désactivation du code journal pour le test
            VerifyAccountingJournal(homePage, site, "");

            try
            {
                //Act
                var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();

                if (newVersionPrint)
                {
                    supplierInvoicesPage.ClearDownloads();
                }

                // Create
                var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
                var supplierItem = supplierInvoicesCreateModalpage.Submit();

                supplierItem.AddNewItem(itemName, quantity, taxType);
                supplierItem.ValidateSupplierInvoice();

                string errorMessage = supplierItem.ManualExportSageError(newVersionPrint, true);

                Assert.IsTrue(errorMessage.Contains("journal value set"), "Le message d'erreur n'est pas celui attendu.");

                supplierInvoicesPage = supplierItem.BackToList();

                supplierInvoicesPage.ResetFilter();
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);

                Assert.IsFalse(supplierInvoicesPage.IsSentToSAGE(), "La supplier invoice a été envoyée vers le SAGE alors quela config n'est pas correcte.");
            }
            finally
            {
                // Remise en place du code journal
                VerifyAccountingJournal(homePage, site, journalSI);
            }
        }

        [Ignore]
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_SageManual_ExportSAGEItemsNewVersion()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            var currency = TestContext.Properties["Currency"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string taxType = TestContext.Properties["TaxTypeSupplierInvoicesExportSage"].ToString();
            string journalSI = TestContext.Properties["Journal_SI"].ToString();

            string quantity = "2";
            bool newVersionPrint = true;

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            homePage.SetSageAutoEnabled(site, false);

            VerifyAccountingJournal(homePage, site, journalSI);

            //Act
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();

            supplierInvoicesPage.ClearDownloads();

            // Create
            var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
            supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
            var supplierItem = supplierInvoicesCreateModalpage.Submit();

            supplierItem.AddNewItem(itemName, quantity, taxType);
            supplierItem.ValidateSupplierInvoice();

            var supplierInvoiceFooter = supplierItem.ClickOnFooter();
            double montantInvoice = supplierInvoiceFooter.GetTotalSupplierInvoice(currency, decimalSeparatorValue);
            supplierItem = supplierInvoiceFooter.ClickOnItems();

            ExportSAGEItemGenerique(supplierItem, newVersionPrint, decimalSeparatorValue, montantInvoice);
        }
        private void ExportSAGEItemGenerique(SupplierInvoicesItem supplierItem, bool newVersionPrint, string decimalSeparatorValue, double montantInvoice)
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            supplierItem.ManualExportSage(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            var correctDownloadedFile = supplierItem.GetExportSAGEFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

            // Récupération du nom du fichier et construction de l'URL du fichier txt à exploiter   
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // On n'exploite que les lignes avec contenu "général" --> "G"
            double[] contenuFichier = ExploitTextFiles.VerifySAGEFileContent(filePath, "G", decimalSeparatorValue);

            Assert.AreEqual(contenuFichier[0].ToString(), contenuFichier[1].ToString(), "Le contenu du fichier SAGE est incorrect.");

            // Remarque : pour les Supplier Invoices, le montant issu du fichier SAGE est égal au montant TTC de la supplier invoice
            Assert.AreEqual(montantInvoice.ToString(), contenuFichier[0].ToString(), "Le montant issu du fichier SAGE n'est pas égal au montant de la supplier invoice défini dans l'application.");
        }

        [Ignore]
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_SageManual_Details_EnableSAGEExport()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string journalSI = TestContext.Properties["Journal_SI"].ToString();

            string quantity = "2";
            bool newVersionPrint = true;

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            homePage.SetSageAutoEnabled(site, false);

            // Manipulation pour permettre export SAGE 
            var accountingParametersPage = homePage.GoToParameters_AccountingPage();
            accountingParametersPage.GoToTab_Journal();
            accountingParametersPage.EditJournal(site, null, journalSI);

            //Act
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();

            // Create
            var SupplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
            SupplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
            var supplierItem = SupplierInvoicesCreateModalpage.Submit();

            supplierItem.AddNewItem(itemName, quantity);
            supplierItem.ValidateSupplierInvoice();

            supplierItem.ManualExportSage(newVersionPrint);

            Assert.IsTrue(supplierItem.CanClickOnEnableSAGE(), "Il n'est pas possible de cliquer sur la fonctionnalité 'Enable export for SAGE' "
                + "pour un supplier invoice envoyé vers le SAGE.");

            supplierItem.ClickOnEnableSAGE();

            Assert.IsTrue(supplierItem.CanClickOnSAGE(), "Il est impossible de cliquer sur la fonctionnalité 'Export for SAGE' "
                + "après avoir cliqué sur 'Enable export for SAGE'.");

            Assert.IsFalse(supplierItem.CanClickOnEnableSAGE(), "Il est possible de cliquer sur la fonctionnalité 'Enable export for SAGE' "
                + "pour un supplier invoice à envoyer vers le SAGE.");
        }

        [Ignore]
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_SageManual_ExportSAGEWithIntegrationDate()
        {
            //Prepare 
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string journalSI = TestContext.Properties["Journal_SI"].ToString();
            var currency = TestContext.Properties["Currency"].ToString();

            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            string quantity = "2";

            string firstDayOfmonth = new DateTime(DateUtils.Now.Year, DateUtils.Now.Month, 1).ToString("ddMMyy");

            bool newVersionPrint = true;

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            homePage.SetSageAutoEnabled(site, false);

            // Manipulation pour permettre export SAGE 
            var accountingParametersPage = homePage.GoToParameters_AccountingPage();
            accountingParametersPage.GoToTab_Journal();
            accountingParametersPage.EditJournal(site, null, journalSI);

            //Act
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();

            supplierInvoicesPage.ClearDownloads();

            // Create
            var SupplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
            SupplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now.AddMonths(-1), site, supplier, true, false);
            var supplierItem = SupplierInvoicesCreateModalpage.Submit();

            supplierItem.AddNewItem(itemName, quantity);
            supplierItem.ValidateSupplierInvoice();

            var supplierInvoiceFooter = supplierItem.ClickOnFooter();
            double montantInvoice = supplierInvoiceFooter.GetTotalSupplierInvoice(currency, decimalSeparatorValue);

            supplierItem = supplierInvoiceFooter.ClickOnItems();
            supplierItem.ManualExportSage(newVersionPrint, true);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            var correctDownloadedFile = supplierItem.GetExportSAGEFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

            // Récupération du nom du fichier et construction de l'URL du fichier txt à exploiter   
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // On vérifie que la date d'intégration est respectée dans le fichier exporté pour les lignes en G
            var listeDates = ExploitTextFiles.VerifySAGEFileIntegrationDate(filePath, "G");
            double[] contenuFichier = ExploitTextFiles.VerifySAGEFileContent(filePath, "G", decimalSeparatorValue);

            Assert.AreEqual(contenuFichier[0].ToString(), contenuFichier[1].ToString(), "Le contenu du fichier SAGE est incorrect.");
            Assert.AreEqual(montantInvoice.ToString(), contenuFichier[0].ToString(), "Le montant issu du fichier SAGE n'est pas égal au montant de la supplier invoice défini dans l'application.");

            Assert.AreNotEqual(0, listeDates.Count, "Le fichier SAGE ne contient aucune date d'intégration.");
            Assert.AreEqual(1, listeDates.Count, "Il y a plusieurs dates d'intégration dans le contenu du fichier SAGE.");
            Assert.AreEqual(firstDayOfmonth, listeDates[0], "La date d'intégration définie dans le fichier SAGE ne correspond pas à celle renseignée lors de l'export.");
        }

        [Ignore] // aucun pays n'utilise SAGE AUTO
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_SageAuto_ExportSAGEWithIntegrationDate()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();

            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            var currency = TestContext.Properties["Currency"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();

            string quantity = "2";
            DateTime invoiceDate = DateUtils.Now.AddMonths(-1);

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();
            var dateFormat = homePage.GetDateFormatPickerValue();

            // Config Export Sage Auto
            homePage.SetSageAutoEnabled(site, true, "Supplier Invoice");

            try
            {
                // Parameter - Accounting --> Integration Date
                var integrationDate = VerifyIntegrationDate(homePage);

                // Calcul vraie dateIntegration
                if (DateTime.Compare(integrationDate.Date, DateUtils.Now.Date) < 0)
                {
                    integrationDate = new DateTime(DateUtils.Now.Year, DateUtils.Now.Month, 1);
                }
                else
                {
                    integrationDate = invoiceDate;
                }

                // Create supplier invoice
                var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                supplierInvoicesPage.ResetFilter();
                var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, invoiceDate, site, supplier, true, false);
                var supplierItem = supplierInvoicesCreateModalpage.Submit();

                supplierItem.AddNewItem(itemName, quantity);
                supplierItem.ValidateSupplierInvoice();

                // Récupération du montant TTC de la facture
                var supplierInvoiceFooter = supplierItem.ClickOnFooter();
                double montantInvoice = supplierInvoiceFooter.GetTotalSupplierInvoice(currency, decimalSeparatorValue);

                // Calcul du montant de la facture transmise à TL
                var supplierInvoiceAccounting = supplierInvoiceFooter.ClickOnAccounting();

                double montantFacture = supplierInvoiceAccounting.GetInvoiceGrossAmount("G", decimalSeparatorValue);
                double montantDetaille = supplierInvoiceAccounting.GetInvoiceDetailAmount("G", decimalSeparatorValue);
                List<DateTime> dateIntegrationSAGE = supplierInvoiceAccounting.GetInvoiceIntegrationDate("G", dateFormat);

                // Retour à la page d'accueil pour vérifier que la facture est partie vers TL
                supplierInvoicesPage = supplierInvoiceAccounting.BackToList();
                supplierInvoicesPage.ResetFilter();
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
                supplierInvoicesPage.Filter(FilterType.DateFrom, DateUtils.Now.AddMonths(-2));
                supplierInvoicesPage.Filter(FilterType.ShowOnTLWaitingForSagePush, true);

                Assert.AreEqual(montantFacture, montantDetaille, "Les montants AmountDebit et AmountCredit de la facture SAGE ne sont pas les mêmes.");
                Assert.AreEqual(montantInvoice, montantFacture, "Le montant issu du fichier SAGE n'est pas égal au montant de la supplier invoice défini dans l'application.");

                Assert.AreNotEqual(0, dateIntegrationSAGE.Count, "Le fichier SAGE ne contient aucune date d'intégration.");
                Assert.AreEqual(1, dateIntegrationSAGE.Count, "Il y a plusieurs dates d'intégration dans le contenu du fichier SAGE.");
                Assert.AreEqual(integrationDate.Date, dateIntegrationSAGE[0].Date, "La date d'intégration définie dans le fichier SAGE ne correspond pas à celle renseignée lors de l'export.");

                Assert.AreEqual(1, supplierInvoicesPage.CheckTotalNumber(), "L'export SAGE Auto de la supplier invoice n'a pas été envoyé vers le SAGE.");
            }
            finally
            {
                homePage.SetSageAutoEnabled(site, false);
            }
        }

        [Ignore]// aucun pays n'utilise SAGE AUTO
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_SageAuto_ExportSAGEItemOK()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            var currency = TestContext.Properties["Currency"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string taxName = TestContext.Properties["TaxTypeSupplierInvoicesExportSage"].ToString();

            string quantity = "2";

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            // Config Export Sage Auto
            homePage.SetSageAutoEnabled(site, true, "Supplier Invoice");

            try
            {
                //Act
                var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                supplierInvoicesPage.ResetFilter();

                // Create
                var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
                var supplierItem = supplierInvoicesCreateModalpage.Submit();

                supplierItem.AddNewItem(itemName, quantity, taxName);
                supplierItem.ValidateSupplierInvoice();

                // Récupération du montant TTC de la facture
                var supplierInvoiceFooter = supplierItem.ClickOnFooter();
                double montantInvoice = supplierInvoiceFooter.GetTotalSupplierInvoice(currency, decimalSeparatorValue);

                // Calcul du montant de la facture transmise à TL
                var supplierInvoiceAccounting = supplierInvoiceFooter.ClickOnAccounting();

                double montantFacture = supplierInvoiceAccounting.GetInvoiceGrossAmount("G", decimalSeparatorValue);
                double montantDetaille = supplierInvoiceAccounting.GetInvoiceDetailAmount("G", decimalSeparatorValue);

                // Retour à la page d'accueil pour vérifier que la facture est partie vers TL
                supplierInvoicesPage = supplierInvoiceAccounting.BackToList();
                supplierInvoicesPage.ResetFilter();
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
                supplierInvoicesPage.Filter(FilterType.ShowOnTLWaitingForSagePush, true);

                Assert.AreEqual(montantFacture, montantDetaille, "Les montants AmountDebit et AmountCredit de la facture SAGE ne sont pas les mêmes.");
                Assert.AreEqual(montantInvoice, montantFacture, "Le montant issu du fichier SAGE n'est pas égal au montant de la supplier invoice défini dans l'application.");
                Assert.AreEqual(1, supplierInvoicesPage.CheckTotalNumber(), "L'export SAGE Auto de la supplier invoice n'a pas été envoyé vers le SAGE.");
            }
            finally
            {
                homePage.SetSageAutoEnabled(site, false);
            }
        }

        [Ignore]
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_SageManuel_ExportSAGEItemKO_Supplier()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string taxName = TestContext.Properties["TaxTypeSupplierInvoicesExportSage"].ToString();

            string quantity = "2";

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Config Export Sage Manuel
            homePage.SetSageAutoEnabled(site, false);

            try
            {
                // Supplier --> KO pour test
                VerifySupplier(homePage, site, supplier, false);

                //Act
                var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                supplierInvoicesPage.ResetFilter();

                // Create
                var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
                var supplierItem = supplierInvoicesCreateModalpage.Submit();

                supplierItem.AddNewItem(itemName, quantity, taxName);
                supplierItem.ValidateSupplierInvoice();

                // Vérifie qu'un message d'erreur est présent dans l'onglet Accounting
                var supplierInvoiceAccounting = supplierItem.ClickOnAccounting();
                string erreur = supplierInvoiceAccounting.GetErrorMessage();

                Assert.AreNotEqual("", erreur, "Aucun message d'erreur n'apparaît alors que le paramétrage relatif au supplier est KO.");
                Assert.IsTrue(erreur.Contains("Accounting_SupplierIdAccouQntingMissing"), "Le message d'erreur ne concerne pas le paramétrage Supplier.");
            }
            finally
            {
                VerifySupplier(homePage, site, supplier);
            }
        }

        [Ignore]
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_SageAuto_ExportSAGEItemKO_CodeJournal()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string taxName = TestContext.Properties["TaxTypeSupplierInvoicesExportSage"].ToString();
            string journalSI = TestContext.Properties["Journal_SI"].ToString();

            string quantity = "2";

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Config Export Sage Auto
            homePage.SetSageAutoEnabled(site, true, "Supplier Invoice");

            try
            {

                // Parameter - Accounting --> Journal KO pour test
                VerifyAccountingJournal(homePage, site, "");

                //Act
                var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                supplierInvoicesPage.ResetFilter();

                // Create
                var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
                var supplierItem = supplierInvoicesCreateModalpage.Submit();

                supplierItem.AddNewItem(itemName, quantity, taxName);
                supplierItem.ValidateSupplierInvoice();

                // Vérifie qu'un message d'erreur est présent dans l'onglet Accounting
                var supplierInvoiceAccounting = supplierItem.ClickOnAccounting();
                string erreur = supplierInvoiceAccounting.GetErrorMessage();

                Assert.AreNotEqual("", erreur, "Aucun message d'erreur n'apparaît alors que le paramétrage relatif au code journal est KO.");
                Assert.IsTrue(erreur.Contains("Supplier invoice journal value set") || erreur.Contains("NOACCOUNTINGJOURNALVALUESET"), "Le message d'erreur ne concerne pas le paramétrage Code journal.");

                // Retour à la page d'accueil pour vérifier que la facture est partie vers TL
                supplierInvoicesPage = supplierInvoiceAccounting.BackToList();
                supplierInvoicesPage.ResetFilter();
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
                supplierInvoicesPage.Filter(FilterType.ShowSentToSageErrorOnly, true);

                Assert.AreEqual(1, supplierInvoicesPage.CheckTotalNumber(), "L'export SAGE Auto de la supplier invoice n'a pas été envoyé vers le SAGE.");
                Assert.IsTrue(supplierInvoicesPage.IsSentToSAGEInErrorOnly(), string.Format(MessageErreur.FILTRE_ERRONE, "Show sent to SAGE and in error only"));

                // On vérifie qu'un mail a été envoyé
                CheckErrorMailSentToUser(homePage);
            }
            finally
            {
                homePage.SetSageAutoEnabled(site, false);
                VerifyAccountingJournal(homePage, site, journalSI);
            }
        }

        [Ignore]
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_SageManuel_ExportSAGEItemKO_SiteAnalyticalPlan()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string taxName = TestContext.Properties["TaxTypeSupplierInvoicesExportSage"].ToString();

            string quantity = "2";

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Config Export Sage Manuel
            homePage.SetSageAutoEnabled(site, false);

            try
            {
                // Sites -- > Analytical plan et section KO pour test
                VerifySiteAnalyticalPlanSection(homePage, site, false);

                //Act
                var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                supplierInvoicesPage.ResetFilter();

                // Create
                var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
                var supplierItem = supplierInvoicesCreateModalpage.Submit();

                supplierItem.AddNewItem(itemName, quantity, taxName);
                supplierItem.ValidateSupplierInvoice();

                // Vérifie qu'un message d'erreur est présent dans l'onglet Accounting
                var supplierInvoiceAccounting = supplierItem.ClickOnAccounting();
                string erreur = supplierInvoiceAccounting.GetErrorMessage();

                Assert.AreNotEqual("", erreur, "Aucun message d'erreur n'apparaît alors que le paramétrage relatif au 'Analytic plan' du site est KO.");
                Assert.IsTrue(erreur.Contains("Accounting analytic plan of the site"), "Le message d'erreur ne concerne pas le paramétrage relatif au 'Analytic plan' du site.");
            }
            finally
            {
                VerifySiteAnalyticalPlanSection(homePage, site);
            }
        }

        [Ignore]
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_SageManuel_ExportSAGEItemKO_NoGroupVat()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string taxName = TestContext.Properties["TaxTypeSupplierInvoicesExportSage"].ToString();

            string quantity = "2";

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Récupération du groupe de l'item
            string itemGroup = GetItemGroup(homePage, itemName);

            // Config Export Sage Manuel
            homePage.SetSageAutoEnabled(site, false);

            bool isJournalOk = VerifyAccountingJournal(homePage, site, "TOTO_codeJournal");
            Assert.IsTrue(isJournalOk, "journal pour " + site + " inexistant");
            try
            {
                // Parameter - Accounting --> Group & VAT KO pour test
                VerifyGroupAndVAT(homePage, itemGroup, taxName, false);

                //Act
                var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                supplierInvoicesPage.ResetFilter();

                // Create
                var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
                var supplierItem = supplierInvoicesCreateModalpage.Submit();

                supplierItem.AddNewItem(itemName, quantity, taxName);
                supplierItem.ValidateSupplierInvoice();

                // Vérifie qu'un message d'erreur est présent dans l'onglet Accounting
                var supplierInvoiceAccounting = supplierItem.ClickOnAccounting();
                string erreur = supplierInvoiceAccounting.GetErrorMessage();

                Assert.AreNotEqual("", erreur, "Aucun message d'erreur n'apparaît alors que le paramétrage relatif au 'Group & VAT' de l'item est KO.");
                Assert.IsTrue(erreur.Contains("no related ItemGroupTax for ItemGroup"), "Le message d'erreur ne concerne pas le paramétrage relatif au 'Group & VAT' de l'item.");
            }
            finally
            {
                VerifyGroupAndVAT(homePage, itemGroup, taxName);
            }
        }

        //_____________________________________ FIN SUPPLIER INVOICE ITEMS _________________________________________

        //Valider les résultats de la liste (selon les critères de recherche)
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_ValidateAllSupplierInvoices()
        {
            // prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string quantity = "2";

            // Log in
            HomePage homePage = LogInAsAdmin();

            //Act
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();
            supplierInvoicesPage.Filter(FilterType.ShowNotValidatedOnly, true);

            if (supplierInvoicesPage.CheckTotalNumber() < 10)
            {
                // Create
                var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
                var supplierItem = supplierInvoicesCreateModalpage.Submit();

                supplierItem.AddNewItem(itemName, quantity);

                supplierInvoicesPage = supplierItem.BackToList();
                supplierInvoicesPage.ResetFilter();
                supplierInvoicesPage.Filter(FilterType.ShowNotValidatedOnly, true);
            }

            if (!supplierInvoicesPage.isPageSizeEqualsTo100())
            {
                supplierInvoicesPage.PageSize("8");
                supplierInvoicesPage.PageSize("100");
            }

            Assert.AreNotEqual(0, supplierInvoicesPage.CheckTotalNumber(), "Il n'y a pas de supplier invoices non validées dans la liste."
                 + " Impossible de tester la fonctionnalité de validation des résultats.");

            Assert.IsTrue(supplierInvoicesPage.checkNotValidateOnly(), String.Format(MessageErreur.FILTRE_ERRONE, "'Show not validated only'"));

            // On valide l'ensemble des resultats
            supplierInvoicesPage.PageUp();
            supplierInvoicesPage.ValidateResults();

            WebDriver.Navigate().Refresh();
            supplierInvoicesPage.ResetFilter();
            supplierInvoicesPage.Filter(FilterType.ShowNotValidatedOnly, true);

            Assert.AreEqual(0, supplierInvoicesPage.CheckTotalNumber(), "Il reste des supplier invoices non validées malgré l'utilisation "
                 + "de la fonctionnalité de validation des résultats.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_PrintResultsNewVersion()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string currency = TestContext.Properties["Currency"].ToString();
            string logoNewRestBase64Begin = "/9j/4QAYRX";
            int logoNewRestLength = 9454;

            string quantity = "2";
            bool newVersionPrint = true;

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var decimalSeparator = homePage.GetDecimalSeparatorValue();
            //Act
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();

            homePage.ClearDownloads();

            supplierInvoicesPage.Filter(FilterType.ShowValidatedNotSentSage, true);

            if (supplierInvoicesPage.CheckTotalNumber() == 0)
            {
                // Create
                var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
                var supplierItem = supplierInvoicesCreateModalpage.Submit();

                supplierItem.AddNewItem(itemName, quantity);
                supplierItem.ValidateSupplierInvoice();

                supplierInvoicesPage = supplierItem.BackToList();
                supplierInvoicesPage.ResetFilter();
            }
            else
            {
                supplierInvoiceNb = supplierInvoicesPage.GetFirstInvoiceNumber();
            }

            string trouve = PrintPdfGenerique(supplierInvoicesPage, supplierInvoiceNb, newVersionPrint);
            // pourquoi ?
            supplierInvoicesPage.Filter(FilterType.ShowValidatedNotSentSage, true);

            supplierInvoicesPage.ClosePrintButton();

            FileInfo filePdf = new FileInfo(trouve);
            filePdf.Refresh();
            Assert.IsTrue(filePdf.Exists, trouve + " non généré");

            //Vérifier que les résultats s'accordent bien au filtre appliqué

            // Exploitation résultat Pdf
            PdfDocument document = PdfDocument.Open(filePdf.FullName);
            List<string> mots = new List<string>();
            List<IPdfImage> images = new List<IPdfImage>();
            foreach (Page page in document.GetPages())
            {
                IEnumerable<Word> words = page.GetWords();
                foreach (Word word in words)
                {
                    mots.Add(word.Text);
                }
                foreach (IPdfImage image in page.GetImages())
                {
                    images.Add(image);
                }
            }

            Assert.IsTrue(mots.Contains(supplierInvoiceNb), "Pas de supplier Invoice Nb " + supplierInvoiceNb);
            supplierInvoicesPage.CheckFirstIndexSupplierInvoices(mots, currency, decimalSeparator);

            Assert.AreEqual(1, images.Count, "Pas d'image " + images.Count);
            // logo NewRest
            string logo_newRest_PDF = Convert.ToBase64String(images[0].RawBytes.ToArray<byte>()).Substring(0, 10);
            Assert.AreEqual(logoNewRestBase64Begin, logo_newRest_PDF, "Logo NewRest manquant - mauvaise image");
            Assert.AreEqual(logoNewRestLength, images[0].RawBytes.Count, "Logo NewRest manquant - mauvaise taille");

            // dans longet Accounting : on a des trucs
        }

        private string PrintPdfGenerique(SupplierInvoicesPage supplierInvoicesPage, string supplierInvoiceNumber, bool printValue)
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNamePdfBegin = "Supplier Invoices List Report_-_";
            string DocFileNameZipBegin = "All_files_";

            WebDriver.Navigate().Refresh();
            supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNumber);

            // Lancement du Print au format PDF
            var reportPage = supplierInvoicesPage.PrintSupplierInvoices(printValue);
            var isReportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();
            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);

            //Assert
            Assert.IsTrue(isReportGenerated, "L'application n'a pas pu générer le fichier attendu.");
            return trouve;
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_ExportResultsNewVersion()
        {
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();

            string quantity = "2";
            bool newVersionPrint = true;

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();

            supplierInvoicesPage.ClearDownloads();

            supplierInvoicesPage.Filter(FilterType.ShowValidatedNotSentSage, true);

            if (supplierInvoicesPage.CheckTotalNumber() == 0)
            {
                // Create
                var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
                var supplierItem = supplierInvoicesCreateModalpage.Submit();

                supplierItem.AddNewItem(itemName, quantity);
                supplierItem.ValidateSupplierInvoice();

                supplierInvoicesPage = supplierItem.BackToList();
                supplierInvoicesPage.ResetFilter();
            }
            else
            {
                supplierInvoiceNb = supplierInvoicesPage.GetFirstInvoiceNumber();
            }

            var excelFile = ExportGenerique(supplierInvoicesPage, supplierInvoiceNb, newVersionPrint, downloadsPath);

            //Assert
            int resultNumber = OpenXmlExcel.GetExportResultNumber(SUPPLIER_INVOICE_EXCEL_SHEET, excelFile.FullName);
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);

            //onglet Invoices
            List<string> invoiceNbList = OpenXmlExcel.GetValuesInList("Invoice Number", "Invoices", excelFile.FullName);
            Assert.IsTrue(invoiceNbList.Contains(supplierInvoiceNb));
            int offset = invoiceNbList.IndexOf(supplierInvoiceNb);
            Assert.AreEqual(supplierInvoicesPage.GetFirstSIIdNumber(), OpenXmlExcel.GetValuesInList("Id", "Invoices", excelFile.FullName)[offset]);
            //Type ?
            Assert.AreEqual(supplierInvoicesPage.GetFirstSISite(), OpenXmlExcel.GetValuesInList("Site", "Invoices", excelFile.FullName)[offset]);
            Assert.AreEqual(supplierInvoicesPage.GetFirstSIDate(), OpenXmlExcel.GetValuesInList("Date", "Invoices", excelFile.FullName)[offset]);
            Assert.AreEqual(supplierInvoicesPage.GetFirstSISupplier(), OpenXmlExcel.GetValuesInList("Supplier", "Invoices", excelFile.FullName)[offset]);
            Assert.IsTrue(supplierInvoicesPage.GetFirstSIInvoiceAmountToBePaid().Contains(OpenXmlExcel.GetValuesInList("Invoiced Amount excl tax", "Invoices", excelFile.FullName)[offset]));
            Assert.IsTrue(supplierInvoicesPage.GetFirstSIInvoiceTotalInclTaxes().Contains(OpenXmlExcel.GetValuesInList("Invoiced Amount incl. tax", "Invoices", excelFile.FullName)[offset]));

            //onglet Supplier invoices detail
            List<string> supplierInvoiceDetailIDList = OpenXmlExcel.GetValuesInList("Invoice Number", "Invoices", excelFile.FullName);
            Assert.IsTrue(supplierInvoiceDetailIDList.Contains(supplierInvoiceNb));
            offset = supplierInvoiceDetailIDList.IndexOf(supplierInvoiceNb);
            Assert.AreEqual(supplierInvoicesPage.GetFirstSIIdNumber(), OpenXmlExcel.GetValuesInList("Id", "Supplier invoices detail", excelFile.FullName)[offset]);
            //Type ?
            Assert.AreEqual(supplierInvoicesPage.GetFirstSISite(), OpenXmlExcel.GetValuesInList("Site", "Supplier invoices detail", excelFile.FullName)[offset]);
            Assert.AreEqual(supplierInvoicesPage.GetFirstSIDate(), OpenXmlExcel.GetValuesInList("Date", "Supplier invoices detail", excelFile.FullName)[offset]);
            Assert.AreEqual(supplierInvoicesPage.GetFirstSISupplier(), OpenXmlExcel.GetValuesInList("Supplier", "Supplier invoices detail", excelFile.FullName)[offset]);
            Assert.IsTrue(supplierInvoicesPage.GetFirstSIInvoiceAmountToBePaid().Contains(OpenXmlExcel.GetValuesInList("Invoiced Amount excl tax", "Supplier invoices detail", excelFile.FullName)[offset]));
            Assert.IsTrue(supplierInvoicesPage.GetFirstSIInvoiceTotalInclTaxes().Contains(OpenXmlExcel.GetValuesInList("Invoiced Amount incl. tax", "Supplier invoices detail", excelFile.FullName)[offset]));

            // On vide le répertoire de téléchargement
            DeleteAllFileDownload();
        }

        private FileInfo ExportGenerique(SupplierInvoicesPage supplierInvoicesPage, string supplierInvoiceNb, bool printValue, string downloadsPath)
        {
            WebDriver.Navigate().Refresh();
            supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);

            // On exporte les résultats sous la forme d'un fichier Excel (dont on récupère le nom)
            // Export des données au format Excel
            supplierInvoicesPage.ExportExcelFile(printValue);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            var correctDownloadedFile = supplierInvoicesPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

            return correctDownloadedFile;
        }

        [Ignore]
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_SageManual_Index_EnableSAGEExport()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string journalSI = TestContext.Properties["Journal_SI"].ToString();

            string quantity = "2";
            bool newVersionPrint = true;

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            homePage.SetSageAutoEnabled(site, false);

            // Manipulation pour permettre export SAGE 
            var accountingParametersPage = homePage.GoToParameters_AccountingPage();
            accountingParametersPage.GoToTab_Journal();
            accountingParametersPage.EditJournal(site, null, journalSI);

            //Act
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();

            supplierInvoicesPage.ClearDownloads();

            // Create
            var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
            supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
            var supplierItem = supplierInvoicesCreateModalpage.Submit();

            supplierItem.AddNewItem(itemName, quantity);
            supplierItem.ValidateSupplierInvoice();

            supplierInvoicesPage = supplierItem.BackToList();
            supplierInvoicesPage.ResetFilter();
            supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);

            supplierInvoicesPage.ManualExportSage(newVersionPrint);
            supplierItem = supplierInvoicesPage.SelectFirstSupplierInvoice();

            Assert.IsTrue(supplierItem.CanClickOnEnableSAGE(), "Il n'est pas possible de cliquer sur la fonctionnalité 'Enable export for SAGE' "
                + "après avoir exporté la supplier invoice vers SAGE depuis la page Index.");

            Assert.IsFalse(supplierItem.CanClickOnSAGE(), "Il est possible de cliquer sur la fonctionnalité 'Export for SAGE' "
                + "après avoir exporté la supplier invoice vers SAGE depuis la page Index.");

            supplierInvoicesPage = supplierItem.BackToList();
            supplierInvoicesPage.EnableExportForSage();
            supplierItem = supplierInvoicesPage.SelectFirstSupplierInvoice();

            Assert.IsFalse(supplierItem.CanClickOnEnableSAGE(), "Il est possible de cliquer sur la fonctionnalité 'Enable export for SAGE' "
                + "après avoir activé la fonctionnalité depuis la page Index.");

            Assert.IsTrue(supplierItem.CanClickOnSAGE(), "Il est impossible de cliquer sur la fonctionnalité 'Export for SAGE' "
                + "après avoir cliqué sur 'Enable export for SAGE' depuis la page Index.");
        }

        [Ignore]
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_SageManual_ExportSAGEResultsKONewVersion()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string taxType = TestContext.Properties["TaxTypeSupplierInvoicesExportSage"].ToString();
            string journalSI = TestContext.Properties["Journal_SI"].ToString();

            string quantity = "2";
            bool newVersionPrint = true;

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            homePage.SetSageAutoEnabled(site, false);

            // Désactivation du code journal pour le test
            VerifyAccountingJournal(homePage, site, "");

            try
            {
                //Act
                var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                supplierInvoicesPage.ResetFilter();

                if (newVersionPrint)
                {
                    supplierInvoicesPage.ClearDownloads();
                }

                supplierInvoicesPage.Filter(FilterType.ShowValidatedNotSentSage, true);

                if (supplierInvoicesPage.CheckTotalNumber() == 0)
                {
                    // Create
                    var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                    supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
                    var supplierItem = supplierInvoicesCreateModalpage.Submit();

                    supplierItem.AddNewItem(itemName, quantity, taxType);
                    supplierItem.ValidateSupplierInvoice();
                    supplierInvoicesPage = supplierItem.BackToList();
                }
                else
                {
                    supplierInvoiceNb = supplierInvoicesPage.GetFirstInvoiceNumber();
                }

                WebDriver.Navigate().Refresh();
                supplierInvoicesPage.ResetFilter();
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);

                // export SAGE en erreur
                string errorMessage = supplierInvoicesPage.ManualExportSageError(newVersionPrint, true);
                Assert.IsTrue(errorMessage.Contains("journal value set"), "Le message d'erreur n'est pas celui attendu.");

                WebDriver.Navigate().Refresh();
                supplierInvoicesPage.ResetFilter();
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);

                Assert.IsFalse(supplierInvoicesPage.IsSentToSAGE(), "La supplier invoice a été envoyée vers le SAGE alors que la config est incomplète.");
            }
            finally
            {
                // Remise en place du code journal
                VerifyAccountingJournal(homePage, site, journalSI);
            }
        }

        [Ignore]
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_SageManual_ExportSAGEResultsNewVersion()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string taxType = TestContext.Properties["TaxTypeSupplierInvoicesExportSage"].ToString();
            var currency = TestContext.Properties["Currency"].ToString();

            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            string quantity = "2";
            bool newVersionPrint = true;

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            homePage.SetSageAutoEnabled(site, false);

            Assert.IsTrue(VerifyAccountingJournal(homePage, site, "888"));

            //Act
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();

            supplierInvoicesPage.ClearDownloads();

            supplierInvoicesPage.Filter(FilterType.ShowValidatedNotSentSage, true);

            double montantInvoice;
            if (supplierInvoicesPage.CheckTotalNumber() == 0)
            {
                // Create
                var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
                var supplierItem = supplierInvoicesCreateModalpage.Submit();

                supplierItem.AddNewItem(itemName, quantity, taxType);
                supplierItem.ValidateSupplierInvoice();

                var supplierInvoiceFooter = supplierItem.ClickOnFooter();
                montantInvoice = supplierInvoiceFooter.GetTotalSupplierInvoice(currency, decimalSeparatorValue);

                supplierInvoicesPage = supplierItem.BackToList();
            }
            else
            {
                supplierInvoiceNb = supplierInvoicesPage.GetFirstInvoiceNumber();
                var supplierItem = supplierInvoicesPage.SelectFirstSupplierInvoice();
                var supplierInvoiceFooter = supplierItem.ClickOnFooter();
                montantInvoice = supplierInvoiceFooter.GetTotalSupplierInvoice(currency, decimalSeparatorValue);

                supplierInvoicesPage = supplierItem.BackToList();
            }

            supplierInvoicesPage.ResetFilter();
            supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);

            // Export SAGE
            supplierInvoicesPage.ManualExportSage(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            var correctDownloadedFile = supplierInvoicesPage.GetExportSAGEFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

            // Récupération du nom du fichier et construction de l'URL du fichier txt à exploiter   
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // On n'exploite que les lignes avec contenu "général" --> "G"
            double[] contenuFichier = ExploitTextFiles.VerifySAGEFileContent(filePath, "G", decimalSeparatorValue);

            Assert.AreEqual(contenuFichier[0].ToString(), contenuFichier[1].ToString(), "Le contenu du fichier SAGE est incorrect.");

            // Remarque : pour les Supplier Invoices, le montant issu du fichier SAGE est égal au montant TTC de la supplier invoice
            Assert.AreEqual(montantInvoice.ToString(), contenuFichier[0].ToString(), "Le montant issu du fichier SAGE n'est pas égal au montant de la supplier invoice défini dans l'application.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_SendAutoToSAGE()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            var currency = TestContext.Properties["Currency"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string taxName = TestContext.Properties["TaxTypeSupplierInvoicesExportSage"].ToString();
            string journalSI = TestContext.Properties["Journal_SI"].ToString();

            string quantity = "2";

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            try
            {
                // Parameter - Accounting --> Journal
                VerifyAccountingJournal(homePage, site, "");

                // 1. Config Export Sage auto pour créer la facture
                homePage.SetSageAutoEnabled(site, true, "Supplier Invoice");

                //Act
                var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                supplierInvoicesPage.ResetFilter();

                // Create
                var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
                var supplierItem = supplierInvoicesCreateModalpage.Submit();

                supplierItem.AddNewItem(itemName, quantity, taxName);
                supplierItem.ValidateSupplierInvoice();

                // Récupération du montant TTC de la facture
                var supplierInvoiceFooter = supplierItem.ClickOnFooter();
                double montantInvoice = supplierInvoiceFooter.GetTotalSupplierInvoice(currency, decimalSeparatorValue);

                var supplierInvoiceAccounting = supplierInvoiceFooter.ClickOnAccounting();
                Assert.AreNotEqual("", supplierInvoiceAccounting.GetErrorMessage(), "Le code journal est manquant mais aucun message d'erreur n'est présent dans l'onglet Accounting.");

                supplierInvoicesPage = supplierInvoiceAccounting.BackToList();
                supplierInvoicesPage.ResetFilter();
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
                supplierInvoicesPage.Filter(FilterType.ShowSentToSageErrorOnly, true);

                Assert.AreEqual(1, supplierInvoicesPage.CheckTotalNumber(), "La supplier invoice n'est pas au statut 'SentToSageInerror' malgré l'échec de l'envoi vers SAGE.");

                // On vérifie qu'un mail a été envoyé
                CheckErrorMailSentToUser(homePage);

                // 2. On remet en place le code journal pour les SI
                VerifyAccountingJournal(homePage, site, journalSI);

                supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                supplierInvoicesPage.ResetFilter();
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);

                supplierInvoicesPage.SelectFirstSupplierInvoice();

                // Calcul du montant de la facture transmise à TL
                supplierInvoiceAccounting = supplierInvoiceFooter.ClickOnAccounting();
                double montantFacture = supplierInvoiceAccounting.GetInvoiceGrossAmount("G", decimalSeparatorValue);
                double montantDetaille = supplierInvoiceAccounting.GetInvoiceDetailAmount("G", decimalSeparatorValue);

                supplierInvoicesPage = supplierInvoiceAccounting.BackToList();
                supplierInvoicesPage.ResetFilter();
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);

                // Renvoi de la facture vers TL
                supplierInvoicesPage.SendAutoToSage();

                WebDriver.Navigate().Refresh();
                supplierInvoicesPage.ResetFilter();
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
                supplierInvoicesPage.Filter(FilterType.ShowOnTLWaitingForSagePush, true);

                Assert.AreEqual(1, supplierInvoicesPage.CheckTotalNumber(), "L'export SAGE Auto de la supplier invoice n'a pas été envoyé vers le SAGE.");
                Assert.AreEqual(montantFacture, montantDetaille, "Les montants AmountDebit et AmountCredit de la facture SAGE ne sont pas les mêmes.");
                Assert.AreEqual(montantInvoice, montantFacture, "Le montant issu du fichier SAGE n'est pas égal au montant de la supplier invoice défini dans l'application.");
            }
            finally
            {
                homePage.SetSageAutoEnabled(site, false);
                VerifyAccountingJournal(homePage, site, journalSI);
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_Index_GenerateSAGETxt()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            var currency = TestContext.Properties["Currency"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string taxName = TestContext.Properties["TaxTypeSupplierInvoicesExportSage"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            string journalSI = TestContext.Properties["Journal_SI"].ToString();
            string quantity = "2";

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            VerifyAccountingJournal(homePage, site, journalSI);

            try
            {
                // 1. Config Export Sage auto pour créer la facture mais pas pour le site MAD
                homePage.SetSageAutoEnabled(site, true, "Supplier Invoice", false);

                //Act
                var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                supplierInvoicesPage.ResetFilter();

                homePage.ClearDownloads();

                // Create
                var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
                var supplierItem = supplierInvoicesCreateModalpage.Submit();

                supplierItem.AddNewItem(itemName, quantity, taxName);
                supplierItem.ValidateSupplierInvoice();

                // Récupération du montant TTC de la facture
                var supplierInvoiceFooter = supplierItem.ClickOnFooter();
                double montantInvoice = supplierInvoiceFooter.GetTotalSupplierInvoice(currency, decimalSeparatorValue);

                // Calcul du montant de la facture transmise à TL
                var supplierInvoiceAccounting = supplierInvoiceFooter.ClickOnAccounting();
                double montantFacture = supplierInvoiceAccounting.GetInvoiceGrossAmount("G", decimalSeparatorValue);
                double montantDetaille = supplierInvoiceAccounting.GetInvoiceDetailAmount("G", decimalSeparatorValue);

                supplierInvoicesPage = supplierInvoiceAccounting.BackToList();
                supplierInvoicesPage.ResetFilter();
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);

                // Renvoi de la facture vers TL
                supplierInvoicesPage.GenerateSageTxt();
                supplierInvoicesPage.ClosePrintButton();

                // On récupère les fichiers du répertoire de téléchargement
                DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
                FileInfo[] taskFiles = taskDirectory.GetFiles();

                var correctDownloadedFile = supplierInvoicesPage.GetExportSAGEFile(taskFiles);
                Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

                // Récupération du nom du fichier et construction de l'URL du fichier txt à exploiter   
                var fileName = correctDownloadedFile.Name;
                var filePath = Path.Combine(downloadsPath, fileName);

                // On n'exploite que les lignes avec contenu "général" --> "G"
                double[] contenuFichier = ExploitTextFiles.VerifySAGEFileContent(filePath, "G", decimalSeparatorValue);


                Assert.AreEqual(contenuFichier[0].ToString(), contenuFichier[1].ToString(), "Le contenu du fichier SAGE généré est incorrect.");
                Assert.AreEqual(montantInvoice.ToString(), contenuFichier[0].ToString(), "Le montant issu du fichier SAGE n'est pas égal au montant de la supplier invoice défini dans l'application.");
                Assert.AreEqual(montantFacture, montantDetaille, "Les montants AmountDebit et AmountCredit de la facture SAGE ne sont pas les mêmes dans l'onglet Accounting.");
                Assert.AreEqual(montantInvoice, montantFacture, "Le montant issu du fichier SAGE n'est pas égal au montant de la supplier invoice défini dans l'application.");
            }
            finally
            {
                var screenshot = WebDriver.TakeScreenshot();
                screenshot.SaveAsFile($"{TestContext.TestName}-GenerateSAGETxt.png");
                TestContext.AddResultFile($"{TestContext.TestName}-GenerateSAGETxt.png");

                homePage.SetSageAutoEnabled(site, false);
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_RemoveItem()
        {
            // Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string currency = TestContext.Properties["Currency"].ToString();
            string NullQty = "0";
            string quantity = "2";

            //arrange
            HomePage homePage = LogInAsAdmin();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();
            try
            {
                //Act
                var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                supplierInvoicesPage.ResetFilter();
                // Create
                var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
                var supplierItem = supplierInvoicesCreateModalpage.Submit();
                // Ajout d'une nouvelle ligne
                supplierItem.AddNewItem(itemName, quantity);
                var nbItems = supplierItem.GetNumberOfItems();
                Assert.AreEqual(1, nbItems, "L'ajout d'item au supplier invoice a échoué.");
                var supplierInvoiceFooter = supplierItem.ClickOnFooter();
                var nbTax = supplierInvoiceFooter.GetNumberOfTaxBaseAmount();
                var nbGroup = supplierInvoiceFooter.GetNumberOfGroup();

                //Assert
                Assert.AreEqual(1, nbTax, "Aucune ligne tax base amount n'a été ajoutée pour la supplier invoice.");
                Assert.AreEqual(2, nbGroup, "Aucune ligne group n'a été ajoutée pour la supplier invoice.");

                supplierItem = supplierInvoiceFooter.ClickOnItems();
                supplierItem.SelectFirstItem();
                supplierItem.SetItemQuantity(itemName, NullQty);

                supplierInvoiceFooter = supplierItem.ClickOnFooter();
                var TaxBaseAmount = supplierInvoiceFooter.GetNumberOfTaxBaseAmount();
                var nbOfGroup = supplierInvoiceFooter.GetNumberOfGroup();
                var totalSI = supplierInvoiceFooter.GetTotalSupplierInvoice(currency, decimalSeparatorValue);
                Assert.AreEqual(0, TaxBaseAmount, "La ligne tax base amount n'a pas été supprimée suite à la suppression de l'item.");
                Assert.AreEqual(1, nbOfGroup, "La ligne group n'a pas été supprimée suite à la suppression de l'item.");
                Assert.AreEqual(0, totalSI, "Le total de la supplier invoice n'est pas nul alors qu'elle ne contient plus d'items.");
            }
            finally
            {
                homePage.Navigate();
                var suppliersInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                suppliersInvoicesPage.ResetFilter();
                suppliersInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
                suppliersInvoicesPage.WaitPageLoading();
                suppliersInvoicesPage.DeleteFirstSI();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_ModifyTaxTypeForSupplierInvoice()
        {
            // Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string currency = TestContext.Properties["Currency"].ToString();
            string taxType = TestContext.Properties["TaxTypeSupplierInvoicesExportSage"].ToString();
            string quantity = "2";

            //arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();
            // Récupération des groupes des items
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            var itemGeneralInformationPage = itemPage.ClickOnFirstItem();
            string group = itemGeneralInformationPage.GetGroupName();

            try
            {
                var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                supplierInvoicesPage.ResetFilter();
                // Create
                var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
                var supplierItem = supplierInvoicesCreateModalpage.Submit();
                // Ajout d'une nouvelle ligne
                supplierItem.AddNewItem(itemName, quantity);
                var supplierInvoiceFooter = supplierItem.ClickOnFooter();
                double montantInvoiceInit = supplierInvoiceFooter.GetTotalSupplierInvoice(currency, decimalSeparatorValue);
                supplierItem = supplierInvoiceFooter.ClickOnItems();
                supplierItem.SelectFirstItem();
                supplierItem.SetVatRate(taxType);
                supplierItem.WaitPageLoading();
                supplierInvoiceFooter = supplierItem.ClickOnFooter();
                var montantSI = supplierInvoiceFooter.GetTotalSupplierInvoice(currency, decimalSeparatorValue);
                Assert.AreNotEqual(montantInvoiceInit, montantSI, "Le total de la supplier invoice n'a pas été modifié malgré le changement du type de taxe de l'item.");

                double newMontant = supplierInvoiceFooter.GetTaxBaseAmountForGroup(group, currency, decimalSeparatorValue) + supplierInvoiceFooter.GetVatAmountForGroup(group, currency, decimalSeparatorValue);
                Assert.AreEqual(newMontant, montantSI, "Le total de la supplier invoice n'est pas correct.");
            }
            finally
            {
                homePage.Navigate();
                var suppliersInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                suppliersInvoicesPage.ResetFilter();
                suppliersInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
                suppliersInvoicesPage.WaitPageLoading();
                suppliersInvoicesPage.DeleteFirstSI();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_VerifyInvoiceAmountByGroup()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName1 = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string itemName2 = TestContext.Properties["Item_SupplierInvoiceBis"].ToString();
            string currency = TestContext.Properties["Currency"].ToString();

            string quantity = "2";

            // Log in

            var homePage = LogInAsAdmin();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            // Récupération des groupes des items
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.Filter(ItemPage.FilterType.Search, itemName1);

            var itemGeneralInformationPage = itemPage.ClickOnFirstItem();
            string group1 = itemGeneralInformationPage.GetGroupName();

            itemPage = itemGeneralInformationPage.BackToList();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName2);

            itemGeneralInformationPage = itemPage.ClickOnFirstItem();
            string group2 = itemGeneralInformationPage.GetGroupName();

            // Création supplier invoice
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();

            try
            {
                var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
                var supplierItem = supplierInvoicesCreateModalpage.Submit();

                // Add item
                supplierItem.AddNewItem(itemName1, quantity);
                supplierItem.AddNewItem(itemName2, quantity);

                // Filter by group
                supplierItem.Filter(SupplierInvoicesItem.FilterItemType.ByGroup, group1);
                double montantGroup1 = supplierItem.GetItemTotalPrice(itemName1, currency, decimalSeparatorValue);

                supplierItem.ResetFilter();
                supplierItem.Filter(SupplierInvoicesItem.FilterItemType.ByGroup, group2);
                double montantGroup2 = supplierItem.GetItemTotalPrice(itemName2, currency, decimalSeparatorValue);

                var supplierInvoiceFooter = supplierItem.ClickOnFooter();
                Assert.AreEqual(3, supplierInvoiceFooter.GetNumberOfGroup(), "Les groupes associés au items ne sontpas visibles dans le footer de l'invoice.");

                double montantGroup1Footer = supplierInvoiceFooter.GetTaxBaseAmountForGroup(group1, currency, decimalSeparatorValue) + supplierInvoiceFooter.GetVatAmountForGroup(group1, currency, decimalSeparatorValue);
                double montantGroup2Footer = supplierInvoiceFooter.GetTaxBaseAmountForGroup(group2, currency, decimalSeparatorValue) + supplierInvoiceFooter.GetVatAmountForGroup(group2, currency, decimalSeparatorValue);
                var totalNumberOfGroups = supplierInvoiceFooter.GetNumberOfGroup();
                Assert.AreEqual(montantGroup1Footer, montantGroup1, "Le montant de la supplier invoice pour le groupe " + group1 + " n'est pas correct.");
                Assert.AreEqual(montantGroup2Footer, montantGroup2, "Le montant de la supplier invoice pour le groupe " + group2 + " n'est pas correct.");
                Assert.AreEqual(montantGroup1Footer + montantGroup2Footer, supplierInvoiceFooter.GetTotalSupplierInvoice(currency, decimalSeparatorValue), "Le total de la supplier invoice n'est pas correct.");
                Assert.AreEqual(3, totalNumberOfGroups, "Les Lignes de Groups ne sont pas séparés");

            }
            finally
            {
                homePage.Navigate();
                supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
                supplierInvoicesPage.DeleteFirstSI();
            }

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_VerifySameTaxAndGroupAmount()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName1 = TestContext.Properties["Item_SupplierInvoice"].ToString(); //group qatar
            string itemName2 = TestContext.Properties["Item_SupplierInvoiceGrp"].ToString(); //group qatar
            string currency = TestContext.Properties["Currency"].ToString();

            string quantity = "2";

            // Log in

            var homePage = LogInAsAdmin();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            // Récupération des groupes des items
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.Filter(ItemPage.FilterType.Search, itemName1);

            var itemGeneralInformationPage = itemPage.ClickOnFirstItem();
            string group1 = itemGeneralInformationPage.GetGroupName();

            itemPage = itemGeneralInformationPage.BackToList();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName2);

            itemGeneralInformationPage = itemPage.ClickOnFirstItem();
            string group2 = itemGeneralInformationPage.GetGroupName();
            //Assert
            Assert.AreEqual(group1, group2, "Les groupes des 2 items ne sont pas identiques.");

            // Création supplier invoice
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();
            try
            {
                var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
                var supplierItem = supplierInvoicesCreateModalpage.Submit();

                // Add item
                supplierItem.AddNewItem(itemName1, quantity);
                supplierItem.AddNewItem(itemName2, quantity);

                // Calcul montant total supplier invoice
                //total invoce witout VAT
                double montantInvoicewoVAT = supplierItem.GetInvoiceTotalPrice(currency, decimalSeparatorValue);
                //total VAT
                var totalVATAmount = supplierItem.GetItemVATAmount(itemName1) + supplierItem.GetItemVATAmount(itemName2);
                //total invoice
                double montantInvoice = montantInvoicewoVAT + totalVATAmount;

                //Footer
                var supplierInvoiceFooter = supplierItem.ClickOnFooter();
                var totalNumberOfGroups = supplierInvoiceFooter.GetNumberOfGroup();
                //Assert
                Assert.AreEqual(2, totalNumberOfGroups, "Il y a plusieurs lignes pour la même association groupe/taxe.");

                //Assert
                double montantGroupFooter = supplierInvoiceFooter.GetTaxBaseAmountForGroup(group1, currency, decimalSeparatorValue) + supplierInvoiceFooter.GetVatAmountForGroup(group1, currency, decimalSeparatorValue);
                Assert.AreEqual(montantGroupFooter, montantInvoice, "Le montant de la supplier invoice pour le groupe " + group1 + " n'est pas correct.");
                Assert.AreEqual(montantInvoice, supplierInvoiceFooter.GetTotalSupplierInvoice(currency, decimalSeparatorValue), "Le total de la supplier invoice n'est pas correct.");
            }

            finally
            {
                //Delete Invoice 
                homePage.Navigate();
                supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                supplierInvoicesPage.ResetFilter();
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
                supplierInvoicesPage.WaitPageLoading();
                supplierInvoicesPage.DeleteFirstSI();

            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_AddNewTaxToSupplierInvoice()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName1 = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string taxType = TestContext.Properties["TaxTypeSupplierInvoicesExportSage"].ToString();
            string currency = TestContext.Properties["Currency"].ToString();

            string quantity = "2";
            string taxAmount = "100";

            string groupTax = TestContext.Properties["SupplierInvoiceManualTax"].ToString();

            // Log in

            var homePage = LogInAsAdmin();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            // Création supplier invoice
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();

            try
            {
                var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
                var supplierItem = supplierInvoicesCreateModalpage.Submit();

                supplierItem.AddNewItem(itemName1, quantity);

                // vérification des infos dans l'onglet Footer
                var supplierInvoiceFooter = supplierItem.ClickOnFooter();
                double montantInvoiceInit = supplierInvoiceFooter.GetTotalSupplierInvoice(currency, decimalSeparatorValue);
                var numberGroup = supplierInvoiceFooter.GetNumberOfGroup();
                Assert.AreEqual(2, numberGroup, "Les informations du groupe n'ont pas été ajoutées pour la supplier invoice.");

                // Ajout de la nouvelle taxe dans le footer
                supplierInvoiceFooter.AddNewTaxToSupplierInvoice(taxType, taxAmount);
                var taxName = supplierInvoiceFooter.GetGroupNameForTaxName(taxType);
                double montantInvoiceFinal = supplierInvoiceFooter.GetTotalSupplierInvoice(currency, decimalSeparatorValue);

                Assert.AreNotEqual(montantInvoiceInit, montantInvoiceFinal, "Le montant de la supplier invoice n'a pas été modifiée malgré l'ajout d'une taxe.");
                var numberNumberTax = supplierInvoiceFooter.GetNumberOfTaxBaseAmount();
                Assert.AreEqual(2, numberNumberTax, "La tax n'a pas été ajoutée dans le tableau 'TAX BASE AMOUNT'");
                Assert.IsTrue(supplierInvoiceFooter.IsAmountPresent(currency + " " + taxAmount), "la ligne correspondant à l'ajout de la taxe n'a pas été ajoutée dans le tableau 'TAX BASE AMOUNT'");

                // Ajout de la ligne au tableau GROUP
                numberGroup = supplierInvoiceFooter.GetNumberOfGroup();
                Assert.AreEqual(3, numberGroup, "La tax n'a pas été ajoutée dans le tableau 'GROUP'");
                Assert.AreEqual(taxAmount, supplierInvoiceFooter.GetTaxBaseAmountForGroup(groupTax, currency, decimalSeparatorValue).ToString(), "La ligne " + groupTax + " n'a pas été ajoutée dans le tableau 'GROUP' ou son montant n'est pas correct.");

                // Vérification infos de l'onglet Items
                supplierItem = supplierInvoiceFooter.ClickOnItems();
                // Calcul montant total supplier invoice
                //total invoce witout VAT
                double montantInvoicewoVAT = supplierItem.GetInvoiceTotalPrice(currency, decimalSeparatorValue);
                //total VAT
                var totalVATAmount = supplierItem.GetItemVATAmount(taxName);
                //total invoice
                double montantInvoice = montantInvoicewoVAT + totalVATAmount;
                var numberItem = supplierItem.GetNumberOfItems();
                Assert.AreEqual(2, numberItem, "La ligne d'ajout de taxe n'a pas été ajoutée dans l'onglet Items");
                var isItemFiltred = supplierItem.IsItemFiltered(groupTax);
                Assert.IsTrue(isItemFiltred, "La ligne " + groupTax + " n'est pas visible dans la page Items.");
                Assert.AreEqual(montantInvoiceFinal, montantInvoice, "Les montants de la supplier invoice sont différents entre les onglets Footer et Items.");
            }
            finally
            {
                homePage.Navigate();
                supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
                supplierInvoicesPage.DeleteFirstSI();
            }

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_RemoveTaxFromSupplierInvoice()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName1 = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string taxType = TestContext.Properties["TaxTypeSupplierInvoicesExportSage"].ToString();
            string currency = TestContext.Properties["Currency"].ToString();

            string quantity = "2";
            string taxAmount = "100";

            string groupTax = TestContext.Properties["SupplierInvoiceManualTax"].ToString();

            // Log in

            var homePage = LogInAsAdmin();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            // Création supplier invoice
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();

            try
            {
                var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
                var supplierItem = supplierInvoicesCreateModalpage.Submit();

                supplierItem.AddNewItem(itemName1, quantity);

                // vérification des infos dans l'onglet Footer
                var supplierInvoiceFooter = supplierItem.ClickOnFooter();
                double montantInvoiceInit = supplierInvoiceFooter.GetTotalSupplierInvoice(currency, decimalSeparatorValue);

                // Ajout de la nouvelle taxe dans le footer
                supplierInvoiceFooter.AddNewTaxToSupplierInvoice(taxType, taxAmount);
                var taxName = supplierInvoiceFooter.GetGroupNameForTaxName(taxType);
                double montantInvoiceAvecTax = supplierInvoiceFooter.GetTotalSupplierInvoice(currency, decimalSeparatorValue);

                // Vérification infos de l'onglet Items
                supplierItem = supplierInvoiceFooter.ClickOnItems();
                // Calcul montant total supplier invoice
                //total invoce witout VAT
                double montantInvoicewoVAT = supplierItem.GetInvoiceTotalPrice(currency, decimalSeparatorValue);
                //total VAT
                var totalVATAmount = supplierItem.GetItemVATAmount(taxName);
                //total invoice
                double montantInvoice = montantInvoicewoVAT + totalVATAmount;

                Assert.AreEqual(2, supplierItem.GetNumberOfItems(), "La ligne d'ajout de taxe n'a pas été ajoutée dans l'onglet Items");
                Assert.IsTrue(supplierItem.IsItemFiltered(groupTax), "La ligne " + groupTax + " n'est pas visible dans la page Items.");
                Assert.AreEqual(montantInvoiceAvecTax, montantInvoice, "Les montants de la supplier invoice sont différents entre les onglets Footer et Items.");

                // Suppression de la taxe ajoutée
                supplierInvoiceFooter = supplierItem.ClickOnFooter();
                Assert.IsTrue(supplierInvoiceFooter.DeleteTaxInTaxAmountTable(currency + " " + taxAmount), "La taxe ne peut pas être supprimée car elle n'est pas visible.");
                Assert.IsFalse(supplierInvoiceFooter.IsAmountPresent(currency + " " + taxAmount), "la ligne correspondant à la taxe ajoutée n'a pas été supprimée dans le tableau 'TAX BASE AMOUNT'");

                double montantInvoiceFinal = supplierInvoiceFooter.GetTotalSupplierInvoice(currency, decimalSeparatorValue);
                Assert.AreNotEqual(montantInvoiceAvecTax, montantInvoiceFinal, "le montant de la supplier invoice n'a pas été modifié suite à la suppression de la taxe.");
                Assert.AreEqual(montantInvoiceInit, montantInvoiceFinal, "le montant de la supplier invoice avec la taxe supprimée n'est pas égal au montant initial (sans taxe).");
            }
            finally
            {
                homePage.Navigate();
                supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
                supplierInvoicesPage.WaitPageLoading();
                supplierInvoicesPage.DeleteFirstSI();
            }

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_ModifyAmountOfSupplierInvoice()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName1 = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string currency = TestContext.Properties["Currency"].ToString();

            string quantity = "2";
            string groupDeviation = TestContext.Properties["SupplierInvoiceDeviation"].ToString();
            double newAmount = 500;

            // Log in

            var homePage = LogInAsAdmin();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            // Création supplier invoice
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();

            try
            {
                var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
                var supplierItem = supplierInvoicesCreateModalpage.Submit();

                supplierItem.AddNewItem(itemName1, quantity);
                double initAmount = supplierItem.GetItemTaxBaseAmount(currency, decimalSeparatorValue);

                // vérification des infos dans l'onglet Footer
                var supplierInvoiceFooter = supplierItem.ClickOnFooter();
                double montantInvoiceInit = supplierInvoiceFooter.GetTotalSupplierInvoice(currency, decimalSeparatorValue);

                // Modification du montant de la taxBase
                supplierInvoiceFooter.ModifyAmount(currency + " " + initAmount.ToString(), newAmount.ToString());
                double montantInvoiceModifiee = supplierInvoiceFooter.GetTotalSupplierInvoice(currency, decimalSeparatorValue);

                Assert.AreNotEqual(montantInvoiceInit, montantInvoiceModifiee, "La modification du total d l'invoice n'a pas été reportée suite à la modification de la valeur de la tax base.");

                // Ajout de la ligne au tableau GROUP
                var TotAmount = newAmount - initAmount;
                var TaxBaseAmountForGoup = supplierInvoiceFooter.GetTaxBaseAmountForGroup(groupDeviation, currency, decimalSeparatorValue);
                Assert.AreEqual(TotAmount, TaxBaseAmountForGoup, "La ligne " + groupDeviation + " n'a pas été ajoutée dans le tableau 'GROUP' ou son montant n'est pas correct.");

                // Vérification infos de l'onglet Items
                supplierItem = supplierInvoiceFooter.ClickOnItems();
                var ItemNumber = supplierItem.GetNumberOfItems();
                var IsItemFiltered = supplierItem.IsItemFiltered(groupDeviation);
                var TotalPrice = supplierItem.GetInvoiceTotalPrice(currency, decimalSeparatorValue);
                Assert.AreEqual(2, ItemNumber, "La ligne de modification du montant n'a pas été ajoutée dans l'onglet Items");
                Assert.IsTrue(IsItemFiltered, "La ligne " + groupDeviation + " n'est pas visible dans la page Items.");
                Assert.AreEqual(montantInvoiceModifiee, TotalPrice, "Les montants de la supplier invoice sont différents entre les onglets Footer et Items.");
            }
            finally
            {
                homePage.Navigate();
                supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
                supplierInvoicesPage.DeleteFirstSI();
            }

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_UpdateValidatedSupplierInvoice()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName1 = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string taxType = TestContext.Properties["TaxTypeSupplierInvoicesExportSage"].ToString();
            string currency = TestContext.Properties["Currency"].ToString();

            string quantity = "2";
            double newAmount = 500;

            // Log in

            var homePage = LogInAsAdmin();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            // Création supplier invoice
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();

            var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
            supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
            var supplierItem = supplierInvoicesCreateModalpage.Submit();

            supplierItem.AddNewItem(itemName1, quantity);
            double initAmount = supplierItem.GetItemTaxBaseAmount(currency, decimalSeparatorValue);
            supplierItem.ValidateSupplierInvoice();

            // vérification des infos dans l'onglet Footer
            var supplierInvoiceFooter = supplierItem.ClickOnFooter();
            var isModifyAmount = supplierInvoiceFooter.IsModifyAmount(currency + " " + initAmount.ToString(), newAmount.ToString());
            var isNewTaxAdded = supplierInvoiceFooter.IsAddNewTaxToSupplierInvoice(taxType, newAmount.ToString());
            Assert.IsFalse(isModifyAmount, "La modification du montant a pu être réalisée sur une supplier invoice validée.");
            Assert.IsFalse(isNewTaxAdded, "L'ajout d'une taxe a pu être effectué sur une supplier invoice validée.");

        }

        [Ignore]
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_SageManual_ExportSAGEUpdatedSupplierInvoice()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName1 = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string taxType = TestContext.Properties["TaxTypeSupplierInvoicesExportSage"].ToString();
            string currency = TestContext.Properties["Currency"].ToString();
            string journalSI = TestContext.Properties["Journal_SI"].ToString();

            string quantity = "2";
            double newAmount = 500;
            string taxAmount = "50";

            bool newVersionPrint = true;

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            homePage.SetSageAutoEnabled(site, false);

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            // Config pour export SAGE
            var accountingParametersPage = homePage.GoToParameters_AccountingPage();
            accountingParametersPage.GoToTab_Journal();
            accountingParametersPage.EditJournal(site, null, journalSI);

            // Création supplier invoice
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();

            var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
            supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
            var supplierItem = supplierInvoicesCreateModalpage.Submit();

            supplierItem.AddNewItem(itemName1, quantity, taxType);
            double initAmount = supplierItem.GetItemTaxBaseAmount(currency, decimalSeparatorValue);

            // Modification des infos dans l'onglet Footer
            var supplierInvoiceFooter = supplierItem.ClickOnFooter();
            supplierInvoiceFooter.ModifyAmount(currency + " " + initAmount.ToString(), newAmount.ToString());
            supplierInvoiceFooter.AddNewTaxToSupplierInvoice(taxType, taxAmount);

            double montantInvoice = supplierInvoiceFooter.GetTotalSupplierInvoice(currency, decimalSeparatorValue);

            supplierItem = supplierInvoiceFooter.ClickOnItems();
            supplierItem.ValidateSupplierInvoice();

            ExportSAGEItemGenerique(supplierItem, newVersionPrint, decimalSeparatorValue, montantInvoice);
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_VerifyTotalDN()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string quantity = "2";
            int dnPrice = 80;
            int dnQuantity = 2;
            string itemName2 = TestContext.Properties["Item_SupplierInvoiceBis"].ToString();
            string quantity2 = "1";
            int dnPrice2 = 10;
            int dnQuantity2 = 3;
            var currency = TestContext.Properties["Currency"].ToString();

            // Log in
            var homePage = LogInAsAdmin();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();
            //Act
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            try
            {
                // Create SI
                var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
                var supplierItem = supplierInvoicesCreateModalpage.Submit();

                // Add new item
                supplierItem.AddNewItem(itemName, quantity);
                supplierItem.SelectItem(itemName);
                supplierItem.SetDNPrice(dnPrice.ToString(), itemName);
                supplierItem.SetDNQuantity(dnQuantity.ToString(), itemName);

                // Add second item
                supplierItem.AddNewItem(itemName2, quantity2);
                supplierItem.SelectItem(itemName2);
                supplierItem.SetDNPrice(dnPrice2.ToString(), itemName2);
                supplierItem.SetDNQuantity(dnQuantity2.ToString(), itemName2);
                var SommeDn = (dnPrice * dnQuantity) + (dnPrice2 * dnQuantity2);
                var TotalDN = supplierItem.GetInvoiceTotalDN(currency, decimalSeparatorValue);
                //Assert
                Assert.AreEqual(SommeDn, TotalDN, "Le montant 'total DN' n'est pas correct.");
            }
            finally
            {
                homePage.Navigate();
                supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
                supplierInvoicesPage.DeleteFirstSI();

            }
        }

        //[Ignore]
        [Timeout(_timeout)]
        [TestMethod]
        public void Ac_SI_CreateSIFromInterimReception()
        {
            string supplierType = "MonSupplierType8";
            string supplierName = "INTERIM";
            string itemName = "supplier-" + new Random().Next().ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();

            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string storageUnit = "KG";
            string supplierRef = itemName + "_SupplierRef";
            string storageQty = 2.32.ToString();

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var decimalSeparator = homePage.GetDecimalSeparatorValue();
            ParametersGlobalSettings globalSettings = homePage.GoToParameters_GlobalSettings();

            ParametersPurchasing purchasing = globalSettings.GoToParameters_PurchasingPage();
            purchasing.GoToTab_Supplier();

            bool IsExistingType = purchasing.CheckForExistingType(supplierType);

            if (!IsExistingType)
            {
                purchasing.CreateNewType(supplierType);
            }

            SuppliersPage supplier = purchasing.GoToPurchasing_SuppliersPage();

            supplier.Filter(SuppliersPage.FilterType.Search, supplierName);
            SupplierItem item;
            if (supplier.CheckTotalNumber() == 0)
            {
                SupplierCreateModalPage create = supplier.SupplierCreatePage();
                create.FillField_CreatNewSupplier(supplierName);
                item = create.Submit();
                item.GetItemsTab();
                var itemCreatePage = item.ItemCreatePage();
                var itemGeneralInformationPage = itemCreatePage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
                var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
                itemCreatePackagingPage.FillField_CreateNewPackaging(
                    site, packagingName, 2.32.ToString(), storageUnit, 4.2.ToString(), supplierName, supplierRef, storageQty);

                supplier = itemCreatePackagingPage.GoToPurchasing_SuppliersPage();
                supplier.Filter(SuppliersPage.FilterType.Search, supplierName);
                item = supplier.SelectFirstItem();
            }
            else
            {
                item = supplier.SelectFirstItem();
            }

            var supplierNumber = item.CheckForSupplierExist(supplierType);
            Assert.IsTrue(supplierNumber.Length > 0, "Pas de supplierNumber");

            item.ClickOnItemsTab();
            var itemNames = item.GetItemNames();
            var result = itemNames.Any(name => name.Contains("ItemForInterim"));
            Assert.IsTrue(result, "itemName non récupéré");
            InterimOrderPage interimOrderPage = homePage.GoToInterim_Orders();
            interimOrderPage.CreateNewInterimOrder(site, supplierName);

            interimOrderPage.Filter(InterimOrderPage.FilterType.Search, "ItemForInterim");
            var dataExist = interimOrderPage.CheckDataExist();
            Assert.IsTrue(dataExist, "no data");

            interimOrderPage.ChangeQuatity();
            InterimReceptionsPage interimReceptionPage = homePage.GoToInterim_Receptions();
            interimReceptionPage.CreateWithNewQuantity(site, supplierName);
            interimReceptionPage.WaitPageLoading();
            interimReceptionPage.WaitForLoad();

            var interimReceptionNumber = interimReceptionPage.CheckNumberReception(decimalSeparator);

            SupplierInvoicesPage supplyInvoice = interimReceptionPage.GoToAccounting_SupplierInvoices();
            SupplierInvoicesCreateModalPage supplyOrderModal = supplyInvoice.SupplierInvoicesCreatePage();
            supplyOrderModal.CreateSIFromIR(site, supplierName, interimReceptionNumber, supplierNumber);
            SupplierInvoicesItem itemsTab = supplyOrderModal.Submit();
            SupplierInvoicesGeneralInformation generalInfoTab = itemsTab.ClickOnGeneralInformation();
            string invoiceNumber = generalInfoTab.GetSupplierInvoiceNumber();
            SupplierInvoicesPage supplierInvoicesList = generalInfoTab.BackToList();
            supplierInvoicesList.Filter(FilterType.ByNumber, invoiceNumber);
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_Filter_ShowNotVerifiedOnly()
        {
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            //Supplier Invoices List Report_-_445651_-_20220603094153.pdf
            string DocFileNamePdfBegin = "Supplier Invoices List Report_-_";
            //All_files_20220225_102148.zip
            string DocFileNameZipBegin = "All_files_";

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var decimalSeparator = homePage.GetDecimalSeparatorValue();

            //Act
            SupplierInvoicesPage supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();

            //Avoir des SI not verified (juste créer)
            SupplierInvoicesCreateModalPage modalSupplier = supplierInvoicesPage.SupplierInvoicesCreatePage();
            modalSupplier.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now.Date, site, supplier);
            SupplierInvoicesItem items = modalSupplier.Submit();

            supplierInvoicesPage = items.BackToList();
            supplierInvoicesPage.ClearDownloads();

            //Appliquer les filtres sur 'Show not verified only'
            supplierInvoicesPage.ResetFilter();
            supplierInvoicesPage.Filter(FilterType.ShowAll, true);
            supplierInvoicesPage.Filter(FilterType.ShowAllInvoices, true);
            supplierInvoicesPage.Filter(FilterType.ShowVerifiedOnly, true);
            supplierInvoicesPage.Filter(FilterType.ShowNotVerifiedOnly, true);
            int nbNotVerified = supplierInvoicesPage.CheckTotalNumber();

            //Print/ Export

            // Export Excel
            supplierInvoicesPage.ExportExcelFile(true);
            supplierInvoicesPage.ClosePrintButton();
            //supplier-invoices 2022-06-03 12-09-20.xlsx
            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();
            FileInfo fileExcel = supplierInvoicesPage.GetExportExcelFile(taskFiles);
            Assert.IsTrue(fileExcel.Exists, "Fichier Excel non trouvé");

            // Print Pdf
            PrintReportPage reportPage = supplierInvoicesPage.PrintSupplierInvoices(true);
            var isReportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();
            Assert.IsTrue(isReportGenerated, "pas de document Print généré");
            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            // cliquer sur All
            string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin, false);
            FileInfo filePdf = new FileInfo(trouve);
            filePdf.Refresh();
            Assert.IsTrue(filePdf.Exists, trouve + " non généré");

            //Vérifier que les résultats s'accordent bien au filtre appliqué

            // Exploitation résultat Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber(SUPPLIER_INVOICE_EXCEL_SHEET, fileExcel.FullName);
            Assert.AreEqual(nbNotVerified, resultNumber, "Mauvais nombre de lignes dans le fichier Excel");

            // Exploitation résultat Pdf
            PdfDocument document = PdfDocument.Open(filePdf.FullName);
            int nbLignes = 0;
            foreach (Page page in document.GetPages())
            {
                IEnumerable<Word> words = page.GetWords();
                nbLignes += words.Count(w => w.Text == "Piece");
            }
            // moins le titre de la table
            nbLignes -= 1;
            Assert.AreEqual(nbNotVerified, nbLignes, "Mauvais nombre de lignes dans le fichier Pdf");

            //Vérifier que le fichier Print/ Export n'est pas vide et que les données correspondent
            supplierInvoicesPage.ResetFilter();
            supplierInvoicesPage.Filter(FilterType.ShowAll, true);
            supplierInvoicesPage.Filter(FilterType.ShowAllInvoices, true);
            supplierInvoicesPage.Filter(FilterType.ShowNotVerifiedOnly, true);
            supplierInvoicesPage.PageSize("100");

            supplierInvoicesPage.CheckExport(fileExcel, decimalSeparator);
            supplierInvoicesPage.CheckPrint(filePdf);
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_Detail_DeleteItem()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string qantity = "3";
            HomePage homePage = LogInAsAdmin();
            SupplierInvoicesPage supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            SupplierInvoicesCreateModalPage modalSupplier = supplierInvoicesPage.SupplierInvoicesCreatePage();
            modalSupplier.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now.Date, site, supplier);
            SupplierInvoicesItem items = modalSupplier.Submit();
            supplierInvoicesPage = items.BackToList();
            supplierInvoicesPage.ClearDownloads();
            try
            {
                //Appliquer les filtres sur 'Show not validated only'
                supplierInvoicesPage.ResetFilter();
                supplierInvoicesPage.Filter(FilterType.ShowNotValidatedOnly, true);
                supplierInvoicesPage.ClickOnFirstline();
                items.AddNewItem(itemName, qantity);
                if (items.IsItemAdded(itemName))
                {
                    items.ClickFirstItem();
                    items.DeleteItem(itemName);
                }
                //assert that item deleted
                Assert.IsTrue(items.IsItemDeleted(itemName), "L'item n'a pas été supprimé");
            }
            finally
            {
                items.BackToList();
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
                supplierInvoicesPage.DeleteFirstSI();
            }

        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_Filter_PieceIdFrom()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            Random rnd = new Random();

            string supplierInvoiceNb = rnd.Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            SupplierInvoicesPage supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();
            try
            {
                SupplierInvoicesCreateModalPage supplierCreateModalPage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierCreateModalPage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now.Date, site, supplier);
                var supplierinvoiceItemPage = supplierCreateModalPage.Submit();
                supplierinvoiceItemPage.BackToList();
                string idNumber = supplierInvoicesPage.GetFirstSIIdNumber();

                int IdNumberAfter = int.Parse(idNumber) + 1;
                int IdNumberBefore = int.Parse(idNumber) - 1;
                //filter by pieceidfrom
                supplierInvoicesPage.Filter(FilterType.PieceIdFrom, idNumber);
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
                var totalNumber = supplierInvoicesPage.CheckTotalNumber();
                //Assert
                Assert.AreEqual(totalNumber, 1, "Le filtre by piece id from ne fonctionne pas correctement");

                supplierInvoicesPage.ResetFilter();
                supplierInvoicesPage.Filter(FilterType.PieceIdFrom, IdNumberBefore.ToString());
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
                totalNumber = supplierInvoicesPage.CheckTotalNumber();
                //Assert
                Assert.AreEqual(totalNumber, 1, "Le filtre by piece id from ne fonctionne pas correctement");

                supplierInvoicesPage.ResetFilter();
                supplierInvoicesPage.Filter(FilterType.PieceIdFrom, IdNumberAfter.ToString());
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
                totalNumber = supplierInvoicesPage.CheckTotalNumber();
                //Assert
                Assert.AreEqual(totalNumber, 0, "Le filtre by piece id from ne fonctionne pas correctement");
            }
            finally
            {
                supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                supplierInvoicesPage.ResetFilter();
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
                supplierInvoicesPage.DeleteFirstSI();
            }

        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_Filter_PieceIdTo()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            Random rnd = new Random();

            string supplierInvoiceNb = rnd.Next().ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            SupplierInvoicesPage supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();
            try
            {
                SupplierInvoicesCreateModalPage supplierCreateModalPage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierCreateModalPage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now.Date, site, supplier);
                var supplierinvoiceItemPage = supplierCreateModalPage.Submit();
                supplierinvoiceItemPage.BackToList();
                string idNumber = supplierInvoicesPage.GetFirstSIIdNumber();

                int IdNumberAfter = int.Parse(idNumber) + 1;
                int IdNumberBefore = int.Parse(idNumber) - 1;
                //filter by pieceidfrom
                supplierInvoicesPage.Filter(FilterType.PieceIdTo, idNumber);
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
                var totalNumber = supplierInvoicesPage.CheckTotalNumber();
                //Assert
                Assert.AreEqual(totalNumber, 1, "Le filtre by piece id from ne fonctionne pas correctement");

                supplierInvoicesPage.ResetFilter();
                supplierInvoicesPage.Filter(FilterType.PieceIdTo, IdNumberBefore.ToString());
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
                totalNumber = supplierInvoicesPage.CheckTotalNumber();
                //Assert
                Assert.AreEqual(totalNumber, 0, "Le filtre by piece id from ne fonctionne pas correctement");

                supplierInvoicesPage.ResetFilter();
                supplierInvoicesPage.Filter(FilterType.PieceIdTo, IdNumberAfter.ToString());
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
                totalNumber = supplierInvoicesPage.CheckTotalNumber();
                //Assert
                Assert.AreEqual(totalNumber, 1, "Le filtre by piece id from ne fonctionne pas correctement");
            }
            finally
            {
                supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                supplierInvoicesPage.ResetFilter();
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
                supplierInvoicesPage.DeleteFirstSI();
            }

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_Filter_ShowOnlyCN()
        {
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            HomePage homePage = LogInAsAdmin();
            SupplierInvoicesPage supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();
            supplierInvoicesPage.Filter(FilterType.ShowOnlyCN, true);
            supplierInvoicesPage.PageSize("100");
            var supplierInvoicesTypes = supplierInvoicesPage.GetInvoiceSupplierTypes();

            if (supplierInvoicesTypes.Any())
            {
                Assert.IsTrue(supplierInvoicesPage.CheckSupplierInvoicesTypes(supplierInvoicesTypes));
            }
            else
            {
                //create SI
                SupplierInvoicesCreateModalPage modalSupplier = supplierInvoicesPage.SupplierInvoicesCreatePage();
                modalSupplier.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now.Date, site, supplier);
                var supplierinvoicestypes = supplierInvoicesPage.GetInvoiceSupplierTypes();
                Assert.IsTrue(supplierInvoicesPage.CheckSupplierInvoicesTypes(supplierinvoicestypes));
            }

        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_Claims_DeleteClaim()
        {
            //Prepare
            //supplier invoice
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            //item
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string quantity = "2";
            string comment = "so filthy";

            LogInAsAdmin();
            var homepage = new HomePage(WebDriver, TestContext);
            var supplierInvoices = homepage.GoToAccounting_SupplierInvoices();
            var supplierInvoicesCreatePage = supplierInvoices.SupplierInvoicesCreatePage();
            supplierInvoicesCreatePage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
            var supplierinvoiceItemPage = supplierInvoicesCreatePage.Submit();
            supplierinvoiceItemPage.AddNewItem(itemName, quantity);
            supplierinvoiceItemPage.AddFirstSupplierInvoiceDetailClaim(comment);
            supplierinvoiceItemPage.ClickOnGeneralInformationTab();
            var claimNumber = supplierinvoiceItemPage.GetClaimNumber();
            supplierinvoiceItemPage.ClickonClaims();
            supplierinvoiceItemPage.DeleteClaim();
            Assert.IsTrue(!supplierinvoiceItemPage.ClaimExist(claimNumber), "Claim non supprimée");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_Claims_EditClaim()
        {
            //Prepare
            //SI
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            //item
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string quantity = "2";
            //claim
            string comment = "so filthy";
            //claim new values
            string newcomment = "stinky";
            var decreasestock = true;
            var claimtype = "Non compliant expiration date";
            string sanctionamount = "3";

            //arrange
            HomePage homePage = LogInAsAdmin();
            try
            {

                //Act
                var suppliersInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                var supplierInvoicesCreateModalPage = suppliersInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalPage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateTime.Now, site, supplier);
                var supplierInvoiceItems = supplierInvoicesCreateModalPage.Submit();
                supplierInvoiceItems.AddNewItem(itemName, quantity);
                supplierInvoiceItems.AddFirstSupplierInvoiceDetailClaim(comment);
                var claimEditClaimForm = supplierInvoiceItems.EditFirstClaim();
                claimEditClaimForm.EditClaim(decreasestock, claimtype, newcomment, sanctionamount);
                var IsCommentGreenOk = supplierInvoiceItems.IsCommentGreen();
                supplierInvoiceItems.EditFirstClaim();

                // Assert
                bool IsEditClaimOk = suppliersInvoicesPage.VerifyClaimCHanged(decreasestock, claimtype, newcomment, sanctionamount);
                Assert.IsTrue(IsCommentGreenOk, "Icone comment doit être coloré en vert");
                Assert.IsTrue(IsEditClaimOk, "erreur de mise a jour du claim");
            }
            finally
            {
                homePage.Navigate();
                var suppliersInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                suppliersInvoicesPage.ResetFilter();
                suppliersInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
                suppliersInvoicesPage.WaitPageLoading();
                suppliersInvoicesPage.DeleteFirstSI();
            }

        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_DetailItem_BySubGroup()
        {
            //Prepare data
            string subGrpName = "subgrpname";
            string subGrpCode = "subgrpcode";
            string group = "A REFERENCIA";
            string qty = "5";
            string itemName = "AEA CHML 2ESC-MACARRONES BOLOÑESA MAD";
            var r = new Random();
            string supplierNumber = r.Next().ToString();
            //packaging of itemName
            string site = "MAD";
            string supplier = "AIR CPU, S.L.";
            //login as admin

            var homePage = LogInAsAdmin();
            // activate sub fgp functionality 
            homePage.SetSubGroupFunctionValue(true);
            //create new subgroup
            ParametersProduction productionPage = homePage.GoToParameters_ProductionPage();
            productionPage.GoToTab_SubGroup();
            if (!productionPage.IsGroupPresent(subGrpName))
                productionPage.AddNewSubGroup(subGrpName, subGrpCode);
            var itemPage = homePage.GoToPurchasing_ItemPage();
            //set item subgroup
            itemPage.Filter(ItemPage.FilterType.Search, itemName);
            var itemGeneralInformationPage = itemPage.ClickOnFirstItem();
            itemGeneralInformationPage.SetGroupName(group);
            itemGeneralInformationPage.SetSubgroupName(subGrpName);
            //go to suppleir invoice page
            var supplierInvoicePage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicePage.ResetFilter();
            try
            {

                SupplierInvoicesCreateModalPage modal = supplierInvoicePage.SupplierInvoicesCreatePage();
                modal.FillField_CreateNewSupplierInvoices(supplierNumber, DateUtils.Now, site, supplier);
                SupplierInvoicesItem supplierInvoiceItem = modal.Submit();
                supplierInvoiceItem.AddNewItem(itemName, qty);
                supplierInvoiceItem.Filter(SupplierInvoicesItem.FilterItemType.ByGroup, group);
                supplierInvoiceItem.Filter(SupplierInvoicesItem.FilterItemType.BySubGrp, subGrpName);
                Assert.IsTrue(supplierInvoiceItem.VerifyItemExist(itemName), "erreur de filtrage par subgroupe");
            }
            finally
            {
                homePage.Navigate();
                supplierInvoicePage = homePage.GoToAccounting_SupplierInvoices();
                supplierInvoicePage.Filter(FilterType.ByNumber, supplierNumber);
                supplierInvoicePage.WaitPageLoading();
                supplierInvoicePage.DeleteFirstSI();
            }

        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_DeleteSupplierInvoiceIndex()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();

            // Log in
            HomePage homePage = LogInAsAdmin();

            //Act
            SupplierInvoicesPage supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();

            // Create
            var SupplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
            SupplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
            var supplierItem = SupplierInvoicesCreateModalpage.Submit();
            supplierInvoicesPage = supplierItem.BackToList();
            var totalNumberAfterCreate = supplierInvoicesPage.CheckTotalNumber();

            //Delete invoice
            supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
            supplierInvoicesPage.WaitPageLoading();
            supplierInvoicesPage.DeleteFirstSI();
            supplierInvoicesPage.ResetFilter();
            supplierInvoicesPage.Filter(FilterType.ShowEDI, false);
            var totalNumberAfterDelete = supplierInvoicesPage.CheckTotalNumber();
            Assert.AreEqual(totalNumberAfterCreate - 1, totalNumberAfterDelete, "La suppression de SI ne fonctionne pas.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_Modification_Price_Qty()
        {
            //Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();

            string quantity = rnd.Next(1, 12).ToString();
            string new_qty = rnd.Next(1, 16).ToString();

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            SupplierInvoicesPage supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();
            var SupplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
            SupplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
            var supplierItem = SupplierInvoicesCreateModalpage.Submit();

            supplierItem.AddNewItem(itemName, quantity);
            supplierItem.SelectItem(itemName);

            /* Update the qty */
            supplierItem.SetSI_Qty(new_qty);

            WebDriver.Navigate().Refresh();
            supplierItem.SelectItem(itemName);
            string qtySI = supplierItem.GetSI_Qty();
            /*Get the price SI*/
            string priceSI = supplierItem.GetSI_PackPrice();

            /* parsing */
            decimal price = decimal.Parse(priceSI);
            decimal qty = decimal.Parse(qtySI);
            decimal qtySI_beforeRefresh_num = decimal.Parse(new_qty);

            Assert.AreEqual(qty, qtySI_beforeRefresh_num, "The qty after and before update are not the same ");

            decimal totalCost = price * qty;

            string totalCostString = totalCost.ToString();
            string TotalSI = supplierItem.GetSI_Total_Vat().Replace("€", "").Trim();

            Assert.AreEqual(totalCostString, TotalSI, "The total operation SI price * SI qty is faulse  ");

            /*Delete the supplier invoice created for test */
            supplierItem.BackToList();
            supplierInvoicesPage.DeleteFirstSI();



        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_NOT_Validated()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            SupplierInvoicesPage supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();

            //Appliquer le filtre  'Show not validated only'
            supplierInvoicesPage.Filter(FilterType.ShowNotValidatedOnly, true);
            // Act
            supplierInvoicesPage.ClickOnFirstline();
            Thread.Sleep(1000);
            // Assert - Vérifier la visibilité des "..."
            supplierInvoicesPage.VerifyDotsArePhysicallyToRightOfClaim();

        }

        // _______________________________________________ Utilitaire _____________________________________________
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

        private bool VerifySupplierInvoiceSageContact(HomePage homePage, string site)
        {
            // Prepare
            string mail = TestContext.Properties["Admin_UserName"].ToString();
            string userName = mail.Substring(0, mail.IndexOf("@"));

            // Act
            var settingsSitesPage = homePage.GoToParameters_Sites();
            settingsSitesPage.Filter(ParametersSites.FilterType.SearchSite, site);
            settingsSitesPage.ClickOnFirstSite();

            settingsSitesPage.ClickToContacts();
            if (String.IsNullOrEmpty(settingsSitesPage.GetSupplierInvoiceSageManager()))
            {
                settingsSitesPage.SetSupplierInvoiceSageManager(userName, mail);
            }

            return !String.IsNullOrEmpty(settingsSitesPage.GetSupplierInvoiceSageManager());
        }

        private bool VerifyPurchasingVAT(HomePage homePage, string taxName, string taxType, bool isOK = true)
        {
            // Prepare
            string purchaseCode = isOK ? "AS05" : "";
            string purchaseAccount = isOK ? "47205001" : "";
            string salesCode = "AR10";
            string salesAccount = "47205001";
            string taxValue = "21";
            string RNCode = "10";
            string RNAccount = "10";

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
                settingsPurchasingPage.EditVATAccountForSpain(purchaseCode, purchaseAccount, salesCode, salesAccount, RNCode, RNAccount);
            }
            catch
            {
                return false;
            }

            return true;
        }

        private bool VerifySupplier(HomePage homePage, string site, string supplier, bool isOK = true)
        {
            // Prepare
            string thirdId = isOK ? "10" : "";
            string accountingId = isOK ? "10" : "";
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
                supplierItem.EditAccounting(thirdId, accountingId);
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

        public DateTime VerifyIntegrationDate(HomePage homePage)
        {
            // Act
            var accountingParametersPage = homePage.GoToParameters_AccountingPage();
            accountingParametersPage.GoToTab_MonthlyClosingDays();
            return accountingParametersPage.GetSageClosureMonthIndex();
        }

        private string GetItemGroup(HomePage homePage, string itemName)
        {
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, itemName);

            var itemGeneralInfo = itemPage.ClickOnFirstItem();

            return itemGeneralInfo.GetGroupName();
        }

        private bool VerifyAccountingJournal(HomePage homePage, string site, string journalSI)
        {
            try
            {
                var accountingJournalPage = homePage.GoToParameters_AccountingPage();
                accountingJournalPage.GoToTab_Journal();
                accountingJournalPage.EditJournal(site, null, journalSI);
            }
            catch
            {
                return false;
            }

            return true;
        }

        private void CheckErrorMailSentToUser(HomePage homePage)
        {
            // Prepare
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["MailBox_UserPassword"].ToString();

            bool mailboxTab = false;

            var mailPage = new MailPage(WebDriver, TestContext);

            try
            {
                // Act
                mailPage = homePage.RedirectToOutlookMailbox();
                mailboxTab = true;

                mailPage.FillFields_LogToOutlookMailbox(userMail);

                // Recherche du mail envoyé
                WebDriver.Navigate().Refresh();
                bool isMailFound = mailPage.CheckIfSpecifiedOutlookMailExist("Export supplier invoice to sage - Error report");
                Assert.IsTrue(isMailFound, "Aucun mail n'a été envoyé pour cette erreur.");

                // Suppression du mail et déconnexion
                mailPage.DeleteCurrentOutlookMail();
                mailPage.Close();
                mailboxTab = false;
            }

            finally
            {
                if (mailboxTab)
                {
                    mailPage.Close();
                }
            }
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_Reinvoice_in_customer_Invoice()
        {
            //Arrange
            HomePage homePage = LogInAsAdmin();
            var supplierInvoicePage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicePage.Filter(FilterType.NottransformedIntoCustomerInvoices, true);
            supplierInvoicePage.ClickOnReinvoiceIncustomerInvoice();
            supplierInvoicePage.ClickOnVerify();
            supplierInvoicePage.ClickOnGenerate();
            Assert.IsTrue(supplierInvoicePage.IsVisible(), "La génération n'est pas  effectuée");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_TVA_SumError()
        {
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string currency = TestContext.Properties["Currency"].ToString();
            string groupTax = TestContext.Properties["SupplierInvoiceManualTax"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string taxType = TestContext.Properties["TaxTypeSupplierInvoicesExportSage"].ToString();
            double newAmount = 100;
            string quantity = "10";

            var homePage = LogInAsAdmin();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();
            try
            {
                var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
                var supplierItem = supplierInvoicesCreateModalpage.Submit();
                supplierItem.AddNewItem(itemName, quantity, taxType);
                double initAmount = supplierItem.GetItemTaxBaseAmount(currency, decimalSeparatorValue);
                var supplierInvoiceFooter = supplierItem.ClickOnFooter();
                var totalVatAmountInit = supplierInvoiceFooter.GetTotalVATAmount(currency, decimalSeparatorValue);
                supplierInvoiceFooter.ModifyTAXBaseAmount(currency + " " + initAmount.ToString(), newAmount.ToString());
                var totalVatAmountModified = supplierInvoiceFooter.GetTotalVATAmount(currency, decimalSeparatorValue);
                var deviationValue = supplierInvoiceFooter.GetVATAmountDeviation(currency, decimalSeparatorValue);
                var isDeviationRowDisplayed = supplierInvoiceFooter.IsDeviationRowDisplayed();
                var TotalVATCount = totalVatAmountInit + deviationValue;
                Assert.IsTrue(isDeviationRowDisplayed, "La ligne de DEVIATION n'est pas affichée.");
                Assert.AreEqual(TotalVATCount, totalVatAmountModified, "Le total du montant de la TVA n'est pas correct et ne prend pas en compte les déviations négatives.");
                Assert.IsTrue(deviationValue < 0, "La valeur de la déviation n'est pas négative.");

            }
            finally
            {
                homePage.Navigate();
                supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
                supplierInvoicesPage.DeleteFirstSI();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_AddItem_SIPrice()
        {
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string qty = "5";
            string itemName = "AEA CHML 2ESC-MACARRONES BOLOÑESA MAD";

            var homePage = LogInAsAdmin();

            //Act
            var supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();
            try
            {
                SupplierInvoicesCreateModalPage modal = supplierInvoicesPage.SupplierInvoicesCreatePage();
                modal.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, false, false);
                SupplierInvoicesItem items = modal.Submit();

                items.ShowBtnPlus();
                string priceSiBeforeItemAdded = items.GetSI_PackPriceAdded();
                items.AddNewItemSiPriceAuto(itemName, qty);
                string priceSiAfterItemAdded = items.GetSI_PackPriceAdded();
                items.SubmitBtn();

                items.ResetFilter();
                Assert.AreNotEqual(priceSiBeforeItemAdded, priceSiAfterItemAdded, "Le Si Price n'est pas ajouté automatiquement !!.");

            }

            finally
            {
                homePage.Navigate();
                supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                supplierInvoicesPage.Filter(FilterType.ShowOnlyInactive, true);
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
                supplierInvoicesPage.DeleteFirstSI();

            }


        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_SuppInvoice_MultipleValidation()
        {
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            HomePage homePage = LogInAsAdmin();
            SupplierInvoicesPage supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();
            supplierInvoicesPage.Filter(FilterType.ShowNotValidatedOnly, true);
            if (supplierInvoicesPage.CheckTotalNumber() == 0)
            {
                SupplierInvoicesCreateModalPage supplierInvoiceseModalPage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoiceseModalPage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
                SupplierInvoicesItem supplierInvoicesItem1 = supplierInvoiceseModalPage.Submit();
                supplierInvoicesItem1.BackToList();
            }

            SupplierInvoicesItem supplierInvoicesItem = supplierInvoicesPage.SelectFirstSupplierInvoice();
            supplierInvoicesItem.ValidateSupplierInvoice();
            SupplierInvoicesGeneralInformation supplierInvoicesGeneralInformation = supplierInvoicesItem.ClickOnGeneralInformation();
            var invoiceNumber = supplierInvoicesGeneralInformation.GetSupplierInvoiceNumber();
            supplierInvoicesItem.BackToList();
            supplierInvoicesPage.Filter(FilterType.ShowNotValidatedOnly, true);
            supplierInvoicesPage.Filter(FilterType.ByNumber, invoiceNumber);
            var afterValidation = supplierInvoicesPage.CheckTotalNumber();
            Assert.AreEqual(0, afterValidation, " La facture fournisseur n'est pas validé");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_TotalFooter_TotalAccounting()
        {
            var currency = TestContext.Properties["Currency"].ToString();


            var homePage = LogInAsAdmin();
            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            SupplierInvoicesPage supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();

            supplierInvoicesPage.ResetFilter();
            supplierInvoicesPage.Filter(FilterType.ShowValidatedNotSentSage, true);

            // Select the first supplier invoice from the filtered list
            SupplierInvoicesItem supplierInvoicesItem = supplierInvoicesPage.SelectFirstSupplierInvoice();

            // Go to the Footer tab and get the total payment amount
            var supplierInvoiceFooter = supplierInvoicesItem.ClickOnFooter();
            double totalPaymentFooter = supplierInvoiceFooter.GetTotalPayment(currency, decimalSeparatorValue);

            // Go to the Accounting tab to retrieve the total "Amount CreditInLocal"
            var supplierInvoiceAccounting = supplierInvoiceFooter.ClickOnAccounting();
            double totalCreditInLocal = supplierInvoiceAccounting.GetTotalAmountCreditInLocal(decimalSeparatorValue);

            // Compare the two amounts and assert they are equal
            Assert.AreEqual(totalPaymentFooter, totalCreditInLocal, "Total Payment in Footer does not match Total Accounting Amount CreditInLocal.");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_Access()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);

            SupplierInvoicesPage supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            bool isAccessOK = supplierInvoicesPage.AccessPage();

            Assert.IsTrue(isAccessOK, "Page inaccessible");

        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_NegativeQuantityItem()
        {
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            try
            {

                SupplierInvoicesPage supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                supplierInvoicesPage.Filter(FilterType.ShowNotValidatedOnly, true);
                var supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
                supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, false, false);
                var supplierItem = supplierInvoicesCreateModalpage.Submit();
                string itemName = supplierItem.IsDev() ? "00061_UAL_SELTZER SEAGRAMSdddd" : "ETH_POLLO PECHUGA SALSA DEMIGLACE PURE PATATA";
                supplierItem.AddNewItem(itemName, "20");
                string quantity = supplierItem.GetfirstQuantity();
                supplierItem.ClickFirstItem();
                supplierItem.SetItemQuantity(itemName, "-1");
                quantity = supplierItem.GetQuantityInput();
                Assert.AreNotEqual(quantity, -1, "La quantité inseré est negative");
            }
            finally
            {
                SupplierInvoicesPage supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
                supplierInvoicesPage.ResetFilter();
                supplierInvoicesPage.Filter(FilterType.ShowNotValidatedOnly, true);
                supplierInvoicesPage.Filter(FilterType.ByNumber, supplierInvoiceNb);
                supplierInvoicesPage.Filter(FilterType.ShowAll, true);
                supplierInvoicesPage.DeleteFirstSI();
            }
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_AffichageValidated()
        {

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);

            SupplierInvoicesPage supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();

            supplierInvoicesPage.ResetFilter();

            supplierInvoicesPage.Filter(FilterType.ShowValidatedNotSentSage, true);

            string firstInvoiceNumber = supplierInvoicesPage.GetFirstInvoiceNumber();
            string firstInvoiceSite = supplierInvoicesPage.GetFirstSISite();

            supplierInvoicesPage.ResetFilter();

            // Appliquer le filtre par numéro et site
            supplierInvoicesPage.Filter(FilterType.ByNumber, firstInvoiceNumber);
            supplierInvoicesPage.Filter(FilterType.ShowValidatedNotSentSage, true);
            supplierInvoicesPage.Filter(FilterType.Site, firstInvoiceSite);

            bool isCorrectInvoiceNumber = supplierInvoicesPage.VerifyInvoiceNumber(firstInvoiceNumber);
            Assert.IsTrue(isCorrectInvoiceNumber, $"Le numéro de facture {firstInvoiceNumber} n'est pas affiché correctement.");

            supplierInvoicesPage.ResetFilter();
            supplierInvoicesPage.Filter(FilterType.ShowValidatedNotSentSage, true);
            supplierInvoicesPage.Filter(FilterType.Site, firstInvoiceSite);

            // Vérifier que la facture correcte est toujours présente après avoir sélectionné tous les sites
            bool isInvoiceNumberCorrect = supplierInvoicesPage.VerifyInvoiceNumber(firstInvoiceNumber);
            Assert.IsTrue(isCorrectInvoiceNumber, $"Le numéro de facture {firstInvoiceNumber} n'est pas affiché correctement.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_AddItemForRN()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ReceiptNotesSupplier"].ToString();
            string deliveryPlace = TestContext.Properties["PlaceFrom"].ToString();
            // Prepare
            string group = TestContext.Properties["Item_Group"].ToString();
            string workshop = TestContext.Properties["Item_Workshop"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string supplierInvoiceNumber = new Random().Next().ToString();
            string qty = 10.ToString();
            string itemName = "ItemForRN";
            string quantity = "2";
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            // Arrange

            var homePage = LogInAsAdmin();
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

            ReceiptNotesPage receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.ResetFilter();
            ReceiptNotesCreateModalPage receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
            receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, deliveryPlace));
            ReceiptNotesItem receiptNotesItem = receiptNotesCreateModalpage.Submit();
            var ID = receiptNotesItem.GetReceiptNoteNumber();
            receiptNotesItem.Filter(ReceiptNotesItem.FilterItemType.SearchByName, itemName);
            receiptNotesItem.SetFirstReceivedQuantity("10");
            receiptNotesItem.Refresh();
            ReceiptNotesQualityChecks receiptNotesQualityChecks = receiptNotesItem.ClickOnChecksTab();
            receiptNotesQualityChecks.SetRefrigeratedVehicleTemperature("5");
            receiptNotesQualityChecks.SetFrozenTemperature("5");
            receiptNotesQualityChecks.SetQualityChecks();
            receiptNotesQualityChecks.SetQAcceptance();
            receiptNotesItem.Validate();
            SupplierInvoicesItem supplierInvoiceItems = receiptNotesItem.GenerateSupplierInvoice(supplierInvoiceNumber);
            supplierInvoiceItems.AddNewItemForRN(itemName, quantity);
            supplierInvoiceItems.CheckRN();
            //assert 
            var RN = supplierInvoiceItems.IsRNChecked();
            Assert.IsTrue(RN, "Receipt Note creation failed.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_Footer_Refresh()
        {
            //Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            string itemName = "itemNameToday" + "-" + new Random().Next().ToString();

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
            string vatrate = "4-IGIC INCREMENTADO";
            // supplier invoice
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string quantity = "2";
            string comment = "so filthy";
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            var itemCreateModalPage = itemPage.ItemCreatePage();
            var itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(itemName.ToString(), group, workshop, taxType, prodUnit);
            var itemCreatePackagingPage = itemGeneralInformationPage.NewPackaging();
            itemCreatePackagingPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier, null, supplierRef);
            itemPage = itemGeneralInformationPage.BackToList();

            var homepage = new HomePage(WebDriver, TestContext);
            var supplierInvoices = homepage.GoToAccounting_SupplierInvoices();
            var supplierInvoicesCreatePage = supplierInvoices.SupplierInvoicesCreatePage();
            supplierInvoicesCreatePage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
            var supplierinvoiceItemPage = supplierInvoicesCreatePage.Submit();
            //VAT -- Footer 
            supplierinvoiceItemPage.AddNewItem(itemName, quantity);
            var supplierInvoiceFooter = supplierinvoiceItemPage.ClickOnFooter();
            var avantModification = supplierInvoiceFooter.GetVatRate();
            supplierInvoiceFooter.ClickOnItems();
            supplierinvoiceItemPage.SelectFirstItem();
            supplierinvoiceItemPage.SetVatRate(vatrate);
            supplierinvoiceItemPage.ClickOnFooter();
            supplierinvoiceItemPage.Refresh();
            var apresModification = supplierInvoiceFooter.GetVatRate();
            Assert.AreNotEqual(avantModification, apresModification, $"Le taux de TVA n'a pas été mis à jour. Avant: {avantModification}, Après: {apresModification}");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_FilterShow_Verified()
        {
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string quantity = "2";

            HomePage homePage = LogInAsAdmin();
            //Act
            SupplierInvoicesPage supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();

            // Create
            SupplierInvoicesCreateModalPage supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
            supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
            SupplierInvoicesItem supplierItem = supplierInvoicesCreateModalpage.Submit();
            supplierItem.AddNewItem(itemName, quantity);
            supplierItem.SetVerified();
            supplierInvoicesPage = supplierItem.BackToList();
            supplierInvoicesPage.ResetFilter();
            supplierInvoicesPage.Filter(FilterType.ShowVerifiedOnly, true);

            bool itemNotVerifiedShow = supplierItem.ItemIsVerifiedShow();
            Assert.IsFalse(itemNotVerifiedShow, "Il faut avoir  que des supplier invoice verified.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_FilterShow_NotVerified()
        {
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next().ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["SupplierInvoice"].ToString();
            string itemName = TestContext.Properties["Item_SupplierInvoice"].ToString();
            string quantity = "2";

            HomePage homePage = LogInAsAdmin();
            //Act
            SupplierInvoicesPage supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();

            // Create
            SupplierInvoicesCreateModalPage supplierInvoicesCreateModalpage = supplierInvoicesPage.SupplierInvoicesCreatePage();
            supplierInvoicesCreateModalpage.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier, true, false);
            SupplierInvoicesItem supplierItem = supplierInvoicesCreateModalpage.Submit();
            supplierInvoicesPage = supplierItem.BackToList();
            supplierInvoicesPage.ResetFilter();
            supplierInvoicesPage.Filter(FilterType.ShowNotVerifiedOnly, true);

            bool itemVerifiedShow = supplierItem.ItemIsNotVerifiedShow();
            Assert.IsFalse(itemVerifiedShow, "Il faut avoir  que des supplier invoice not verified.");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_Create_NewCreditNote()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();

            //Login
            HomePage homePage = LogInAsAdmin();

            //Act
            //Verify is Site active 
            ParametersSites siteParameterPage = homePage.GoToParameters_Sites();
            siteParameterPage.Filter(ParametersSites.FilterType.SearchSite, site);
            siteParameterPage.CheckIfFirstSiteIsActive();

            //Create invoice
            SupplierInvoicesPage supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();
            supplierInvoicesPage.Filter(SupplierInvoicesPage.FilterType.ByNumber, supplierInvoiceNb);
            SupplierInvoicesItem supplierItem = supplierInvoicesPage.SelectFirstSupplierInvoice();
            supplierItem.ValidateSupplierInvoice();
            supplierItem.BackToList();
            supplierInvoicesPage.ResetFilter();
            supplierInvoicesPage.Filter(SupplierInvoicesPage.FilterType.ByNumber, supplierInvoiceNb);

            //Assert
            string FirstInvoiceNumber = supplierInvoicesPage.GetFirstInvoiceNumber();
            Assert.AreEqual(supplierInvoiceNb, FirstInvoiceNumber, "La facture Credit note créée n'apparaît pas dans la liste des supply invoices disponibles.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_NewSI_Activated()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();

            //Login
            HomePage homePage = LogInAsAdmin();

            //Act
            //Verify is Site active 
            ParametersSites siteParameterPage = homePage.GoToParameters_Sites();
            siteParameterPage.Filter(ParametersSites.FilterType.SearchSite, site);
            siteParameterPage.CheckIfFirstSiteIsActive();

            //Search  supplier invoice 
            SupplierInvoicesPage supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();
            supplierInvoicesPage.Filter(SupplierInvoicesPage.FilterType.ByNumber, supplierInvoiceNb);
            SupplierInvoicesItem supplierItem = supplierInvoicesPage.SelectFirstSupplierInvoice();
            supplierItem.ClickOnGeneralInformation();

            //Assert
            var isActivated = supplierItem.IsSupplierInvoiceActivated();
            Assert.IsTrue(isActivated, "la case Activated sous supplier invoice ne soit pas toujours cochée");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_SI_Filter_ShowTransformedIntoCustomerInvoice()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();

            //Login
            HomePage homePage = LogInAsAdmin();

            //Act
            //Verify is Site active 
            ParametersSites siteParameterPage = homePage.GoToParameters_Sites();
            siteParameterPage.Filter(ParametersSites.FilterType.SearchSite, site);
            var isActive = siteParameterPage.CheckIfFirstSiteIsActive();
            Assert.IsTrue(isActive, "Le Site n'est pas Active");

            //Create invoice
            SupplierInvoicesPage supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            supplierInvoicesPage.ResetFilter();
            supplierInvoicesPage.Filter(SupplierInvoicesPage.FilterType.ByNumber, supplierInvoiceNb);
            SupplierInvoicesItem supplierItem = supplierInvoicesPage.SelectFirstSupplierInvoice();
            supplierItem.ReinvoiceInCustomerInvoice();
            supplierItem.BackToList();
            supplierInvoicesPage.ResetFilter();
            supplierInvoicesPage.Filter(SupplierInvoicesPage.FilterType.ByNumber, supplierInvoiceNb);
            supplierInvoicesPage.Filter(SupplierInvoicesPage.FilterType.TransformedIntoCustomerInvoices, true);

            //Assert
            var nbOfItems = supplierInvoicesPage.CheckTotalNumber();
            string FirstInvoiceNumber = supplierInvoicesPage.GetFirstInvoiceNumber();
            Assert.AreEqual(nbOfItems, 1, "Le Supplier Invoice n'a pas été transformé on customer invoice.");
            Assert.AreEqual(supplierInvoiceNb, FirstInvoiceNumber, "Le Supplier Invoice n'a pas été transformé on customer invoice.");
        }
    }
}
