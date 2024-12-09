using DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml.Wordprocessing;
using iText.StyledXmlParser.Jsoup.Nodes;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Accounting.SupplierInvoices;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.Claims;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.Inventory;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.ReceiptNotes;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Media.Media3D;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using static Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.Claims.ClaimsPage;

namespace Newrest.Winrest.FunctionalTests.Warehouse
{
    [TestClass]
    public class ClaimsTest : TestBase
    {
        private const int _timeout = 600000;
        /// <summary>
        /// 
        /// Mise en place du paramétrage pour la configuration Winrest 4.0 
        /// 
        /// </summary>
        [TestMethod]
        [Priority(0)]
        [Timeout(_timeout)]
        public void WA_CLAI_SetConfigWinrest4_0()
        {
            //Arrange
            var homePage = LogInAsAdmin();
            ClearCache();

            // New version search
            homePage.SetNewVersionSearchValue(true);

            // New version keyword
            homePage.SetNewVersionKeywordValue(true);

            // New group display
            homePage.SetNewGroupDisplayValue(true);

            // Vérifier que c'est activé
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();
            claimsPage.Filter(ClaimsPage.FilterType.ShowNotValidated, true);
            var claimsItem = claimsPage.SelectFirstClaim();

            // Vérifier New version search
            try
            {
                var itemName = claimsItem.GetFirstItemName();
                claimsItem.Filter(ClaimsItem.FilterItemType.SearchByName, itemName);
            }
            catch
            {
                throw new Exception("La recherche a pu être effectuée, le NewSearchMode est actif.");
            }

            // vérifier new keyword search
            try
            {
                claimsItem.ResetFilters();
                claimsItem.Filter(ClaimsItem.FilterItemType.SearchByKeyword, "TEST_KEY");
            }
            catch
            {
                throw new Exception("La recherche par keyword n'a pas pu être effectuée, le NewKeywordMode est inactif.");
            }

            // vérifier new group display
            Assert.IsTrue(claimsItem.IsGroupDisplayActive(), "Le paramètre 'NewGroupDisplay' n'est pas activé.");

            // Vérifier New version supplier invoice
            var supplierInvoicePage = homePage.GoToAccounting_SupplierInvoices();
            Assert.IsFalse(supplierInvoicePage.IsInvoiceAmountWithoutTaxPresent(), "Le paramètre 'NewSupplierInvoiceVersion' n'est pas activé.");
        }

        //Verification du mail pour SendResultByMail
        [TestMethod]
        [Priority(1)]
        [Timeout(_timeout)]
        public void WA_CLAI_CheckMailAdressForSendByResult()
        {
            var email = TestContext.Properties["Admin_UserName"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();

            // Log in
            var homePage = LogInAsAdmin();

            //Act
            var suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();

            suppliersPage.Filter(SuppliersPage.FilterType.Search, supplier);
            var supplierItemPage = suppliersPage.SelectFirstItem();
            supplierItemPage.GoToContactTab();

            supplierItemPage.EditContactForClaim(email);
        }

        //_________________________________________CREATE_CLAIM____________________________________________________________

        /*
         * Création d'une nouvelle Claims note
        */
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_CreateclaimsItem()
        {
            //prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();
            string claimNumber = String.Empty;

            //arrange
            var homePage = LogInAsAdmin();

            //Act
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();
            try
            {
                // Create
                var claimsCreateModalpage = claimsPage.ClaimsCreatePage();
                claimsCreateModalpage.FillField_CreatNewClaims(DateUtils.Now, site, supplier, placeTo, true);
                claimNumber = claimsCreateModalpage.GetClaimId();
                var claimsItem = claimsCreateModalpage.Submit();
                var itemName = claimsItem.GetFirstItemName();

                // Update the first item value to activate the activation menu
                claimsItem.SelectFirstItem();
                claimsItem.AddPrice(itemName, 10);
                claimsItem.AddQuantityAndType(itemName, 2);

                //Suite a la modif Claim
                var editClaimForm = claimsItem.EditClaimForm(itemName);
                editClaimForm.SetClaimFormForClaimedV3();

                claimsPage = claimsItem.BackToList();
                claimsPage.Filter(ClaimsPage.FilterType.ByNumber, claimNumber);
                string claimNumberFiltred = claimsPage.GetFirstID();
                var totalNumberClaim = claimsPage.CheckTotalNumber();
                //Assert
                Assert.AreEqual(claimNumberFiltred, claimNumber, String.Format(MessageErreur.OBJET_NON_CREE, "La claim"));
                Assert.AreEqual(1, totalNumberClaim, String.Format(MessageErreur.OBJET_NON_CREE, "La claim"));

            }
            finally
            {
                if (!String.IsNullOrEmpty(claimNumber))
                {
                    claimsPage.DeleteClaim(claimNumber);
                }
                claimsPage.ResetFilter();
            }
        }

        /*
         * Création d'une nouvelle Claim à partir d'une claim
        */
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_CreateClaimFromClaim()
        {
            //prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string delivery = TestContext.Properties["PlaceFrom"].ToString();
            int itemQuantity = 2;
            string sanctionAmount = "5";

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();

            // Create 1st claim
            var claimsCreateModalpage = claimsPage.ClaimsCreatePage();
            claimsCreateModalpage.FillField_CreatNewClaims(DateUtils.Now, site, supplier, delivery, true);
            string ID = claimsCreateModalpage.GetClaimId();
            var claimsItem = claimsCreateModalpage.Submit();

            // Update the first item value to activate the activation menu
            var itemName = claimsItem.GetFirstItemName();
            claimsItem.SelectFirstItem();
            claimsItem.AddQuantityAndType(itemName, itemQuantity);
            claimsItem.AddSanction(sanctionAmount);

            //Suite a la modif Claim
            var editClaimForm = claimsItem.EditClaimForm(itemName);
            editClaimForm.SetClaimFormForClaimedV3();

            claimsItem.Validate();
            var totalClaim1 = claimsItem.GetTotalClaim();
            var totalsanctionClaim1 = claimsItem.GetTotalSanction();

            claimsPage = claimsItem.BackToList();
            claimsPage.ResetFilter();

            // Create 2nd claim
            claimsCreateModalpage = claimsPage.ClaimsCreatePage();
            claimsCreateModalpage.FillField_CreatNewClaims(DateUtils.Now, site, supplier, delivery, true, "Claim", ID);
            ID = claimsCreateModalpage.GetClaimId();
            claimsItem = claimsCreateModalpage.Submit();
            claimsPage = claimsItem.BackToList();
            claimsPage.ResetFilter();

            //Assert
            var id = claimsPage.GetFirstID();
            Assert.AreEqual(id, ID, String.Format(MessageErreur.OBJET_NON_CREE, "La claim"));

            claimsItem = claimsPage.SelectFirstClaim();
            string firstItemName = claimsItem.GetFirstItemName();
            Assert.AreEqual(itemName, firstItemName, "L'item de la nouvelle claim n'est pas identique à celui de la claim initiale");
            string quantity = claimsItem.GetQuantity();
            Assert.AreEqual(itemQuantity.ToString(), quantity, "La quantité de l'item de la nouvelle claim n'est pas identique à celle de l'item de la claim initiale");
            string totalClaim = claimsItem.GetTotalClaim();
            Assert.AreEqual(totalClaim1, totalClaim, "Le total claim de la nouvelle claim n'est pas identique à celui de la claim initiale");
            string totalSanction = claimsItem.GetTotalSanction();
            Assert.AreEqual(totalsanctionClaim1, totalSanction, "Le total sanction de la nouvelle claim n'est pas identique à celui de la claim initiale");

            var claimsGeneralInformation = claimsItem.ClickOnGeneralInformation();
            string claimSite = claimsGeneralInformation.GetClaimSite();
            Assert.AreEqual(site, claimSite, "Le site de la nouvelle claim n'est pas identique à celui de la claim initiale");
            string claimSupplier = claimsGeneralInformation.GetClaimSupplier();
            Assert.AreEqual(supplier, claimSupplier, "Le supplier de la nouvelle claim n'est pas identique à celui de la claim initiale");
            string claimDeliveryLocation = claimsGeneralInformation.GetClaimDeliveryLocation();
            Assert.AreEqual(delivery, claimDeliveryLocation, "La delivery location de la nouvelle claim n'est pas identique à celui de la claim initiale");

            claimsPage = claimsItem.BackToList();
            claimsPage.Filter(ClaimsPage.FilterType.ByNumber, ID);

            Assert.IsFalse(claimsPage.CheckValidation(false), "La claim créée à partir d'un claim est validée, alors qu'elle ne le devrait pas");
        }

        /*
         * Création d'une nouvelle Claim à partir d'une Purchase Order
        */
        [Ignore]//en standby ticket 12539
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_CreateClaimFromPurchaseOrder()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string delivery = TestContext.Properties["PlaceFrom"].ToString();

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            // Create Purchase Order
            var purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();

            var createPurchaseOrderPage = purchaseOrdersPage.CreateNewPurchaseOrder();
            createPurchaseOrderPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, delivery, DateUtils.Now.AddDays(+10), true);
            var purchaseOrderItemPage = createPurchaseOrderPage.Submit();
            var ID = purchaseOrderItemPage.GetPurchaseOrderNumber();
            purchaseOrderItemPage.SelectFirstItemPo();
            purchaseOrderItemPage.AddQuantity("2");
            purchaseOrderItemPage.Validate();

            // Create claim
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();
            var claimsCreateModalpage = claimsPage.ClaimsCreatePage();
            claimsCreateModalpage.FillField_CreatNewClaims(DateUtils.Now, site, supplier, delivery, true, "Purchase Order", ID);
            var claimID = claimsCreateModalpage.GetClaimId();
            var claimsItem = claimsCreateModalpage.Submit();
            claimsPage = claimsItem.BackToList();

            //Assert
            Assert.AreEqual(claimsPage.GetFirstID(), claimID, String.Format(MessageErreur.OBJET_NON_CREE, "La claim"));
            //test à finir : vérifier les infos dans la nouvelle claim si identique à la première
        }

        /*
         * Création d'une nouvelle Claims à partir d'une receipt note
        */
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_CreateClaimsFromReceiptNote()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string delivery = TestContext.Properties["PlaceFrom"].ToString();

            // Log in
            var homePage = LogInAsAdmin();

            string receiptNoteID = CreateReceiptNote(homePage, site, supplier, delivery);

            //Act
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();

            // Create
            var claimsCreateModalpage = claimsPage.ClaimsCreatePage();
            claimsCreateModalpage.FillField_CreatNewClaims(DateUtils.Now, site, supplier, delivery, true, "Receipt Note", receiptNoteID);
            String ID = claimsCreateModalpage.GetClaimId();
            var claimsItem = claimsCreateModalpage.Submit();
            var itemName = claimsItem.GetFirstItemName();

            // Update the first item value to activate the activation menu
            claimsItem.SelectFirstItem();
            claimsItem.AddQuantityAndType(itemName, 2);

            var editClaimForm = claimsItem.EditClaimForm(itemName);
            editClaimForm.SetClaimFormForClaimedV3();

            ClaimsGeneralInformation generalInfo = claimsItem.ClickOnGeneralInformation();
            string numberOfRN = generalInfo.GetRelatedReceiptNoteNumber();
            string link = generalInfo.GetRelatedReceiptNoteNumberLink();
            Assert.AreNotEqual(link , String.Empty);
            Assert.AreEqual(numberOfRN, receiptNoteID, "Le numéro de la Receipt Note ne correspond pas à l'ID de la Receipt Note précédemment créé.");
            claimsPage = claimsItem.BackToList();
            claimsPage.ResetFilter();
            claimsPage.Filter(FilterType.ByNumber, ID);
            //Assert
            Assert.AreEqual(claimsPage.GetFirstID(), ID, String.Format(MessageErreur.OBJET_NON_CREE, "La claim"));
            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();
            receiptNotesPage.Filter(ReceiptNotesPage.FilterType.ByNumber, receiptNoteID);
            Assert.IsTrue(receiptNotesPage.IsWithClaim(), String.Format(MessageErreur.FILTRE_ERRONE, "'Show receipt notes with claims'"));
        }

        private string CreateReceiptNote(HomePage homePage, string site, string supplier, string delivery)
        {
            // Create a ReceiptNote
            var receiptNotesPage = homePage.GoToWarehouse_ReceiptNotesPage();

            // Create
            var receiptNotesCreateModalpage = receiptNotesPage.ReceiptNotesCreatePage();
            receiptNotesCreateModalpage.FillField_CreatNewReceiptNotes(new RNCreationParameters(DateUtils.Now, site, supplier, delivery));

            var receiptNotesItem = receiptNotesCreateModalpage.Submit();

            string receiptNoteID = receiptNotesItem.GetReceiptNoteNumber();

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

            return receiptNoteID;
        }

        //_________________________________________FIN CREATE_CLAIM________________________________________________________

        //_________________________________________DELETE_CLAIM________________________________________________________
        /*
         * Supprimer claim non validée
        */
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_DeleteClaim()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();
            string ID = string.Empty;
            // Log in
            var homePage = LogInAsAdmin();

            //Act
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();

            try
            {
                // Create
                var claimsCreateModalpage = claimsPage.ClaimsCreatePage();
                claimsCreateModalpage.FillField_CreatNewClaims(DateUtils.Now, site, supplier, placeTo, true);
                ID = claimsCreateModalpage.GetClaimId();
                var claimsItem = claimsCreateModalpage.Submit();

                var itemName = claimsItem.GetFirstItemName();

                // Update the first item value to activate the activation menu
                claimsItem.SelectFirstItem();
                claimsItem.AddPrice(itemName, 10);
                claimsItem.AddQuantityAndType(itemName, 2);

                //Suite a la modif Claim
                var editClaimForm = claimsItem.EditClaimForm(itemName);
                editClaimForm.SetClaimFormForClaimedV3();
                claimsPage = claimsItem.BackToList();
                claimsPage.Filter(ClaimsPage.FilterType.ByNumber, ID);
                string firstID = claimsPage.GetFirstID();
                //Assert
                Assert.AreEqual(firstID, ID, string.Format(MessageErreur.OBJET_NON_CREE, "La claim"));

            }
            finally
            {
                //Delete
                claimsPage.DeleteClaim(ID);
                string firstIDAfterDelete = claimsPage.GetFirstID();
                Assert.AreNotEqual(firstIDAfterDelete, ID, string.Format(MessageErreur.OBJET_NON_SUPPRIME, "La claim"));
            }
        }
        //_________________________________________FIN DELETE_CLAIM________________________________________________________

        //_________________________________________FILTERS TESTING_________________________________________________________

        /* 
         * Test filtrage : Par numéro
         */
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_Filter_SearchByNumberFilter()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();
            string number = string.Empty;
            //Login
            var homePage = LogInAsAdmin();

            //Act
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();

            try
            {
                // Create
                var claimsCreateModalpage = claimsPage.ClaimsCreatePage();
                claimsCreateModalpage.FillField_CreatNewClaims(DateUtils.Now, site, supplier, placeTo, true);
                number = claimsCreateModalpage.GetClaimId();
                var claimsItem = claimsCreateModalpage.Submit();

                // Update the first item value to activate the activation menu
                var itemName = claimsItem.GetFirstItemName();
                claimsItem.SelectFirstItem();
                claimsItem.AddQuantityAndType(itemName, 2);


                //Suite a la modif Claim
                var editClaimForm = claimsItem.EditClaimForm(itemName);
                editClaimForm.SetClaimFormForClaimedV3();

                claimsPage = claimsItem.BackToList();

                claimsPage.Filter(ClaimsPage.FilterType.ByNumber, number);
                string firstID = claimsPage.GetFirstID();
                //Assert
                Assert.AreEqual(number, firstID, String.Format(MessageErreur.FILTRE_ERRONE, "By number"));
            }
            finally
            {
                claimsPage.ResetFilter();
                claimsPage.Filter(ClaimsPage.FilterType.ByNumber, number);
                claimsPage.DeleteClaim(number);
            }
            
        }

        /* 
         * Test filtrage : Par supplier
         */
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_Filter_Supplier()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();
            HomePage homePage = LogInAsAdmin();
            ClaimsPage claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();
            claimsPage.Filter(ClaimsPage.FilterType.Suppliers, supplier);
            if (claimsPage.CheckTotalNumber() < 20)
            {
                ClaimsCreateModalPage claimsCreateModalpage = claimsPage.ClaimsCreatePage();
                claimsCreateModalpage.FillField_CreatNewClaims(DateUtils.Now, site, supplier, placeTo, true);
                ClaimsItem claimsItem = claimsCreateModalpage.Submit();
                string itemName = claimsItem.GetFirstItemName();
                claimsItem.SelectFirstItem();
                ClaimEditClaimForm editClaimForm = claimsItem.EditClaimForm(itemName);
                editClaimForm.SetClaimFormForClaimedV3();
                claimsPage = claimsItem.BackToList();
                claimsPage.ResetFilter();
                claimsPage.Filter(ClaimsPage.FilterType.Suppliers, supplier);
            }
            if (!claimsPage.isPageSizeEqualsTo100())
            {
                claimsPage.PageSize("8");
                claimsPage.PageSize("100");
            }
            bool isVerified = claimsPage.VerifySupplier(supplier);
            Assert.IsTrue(isVerified, string.Format(MessageErreur.FILTRE_ERRONE, "Supplier"));
        }

    
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_Filter_DateFrom()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();
            DateTime fromDate = DateTime.Now.Date;
            DateTime fromDateInf = fromDate.AddDays(-1);
            DateTime fromDateSup = fromDate.AddDays(2);

            // Log in
            var homePage = LogInAsAdmin();
            var dateFormat = homePage.GetDateFormatPickerValue();

            // Act
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();
            claimsPage.Filter(ClaimsPage.FilterType.DateFrom, fromDate);
            var claims_DateInf = "";
            var claims_DateSup = "";
            var claims_Date = "";

            try
            {
                // Create claim 1
                var claimsCreateModalpage = claimsPage.ClaimsCreatePage();
                claims_Date = claimsCreateModalpage.GetIdClaim();
                claimsCreateModalpage.FillField_CreatNewClaims(fromDate, site, supplier, placeTo, true);
                var claimsItem_Date = claimsCreateModalpage.Submit();

                claimsPage = claimsItem_Date.BackToList();
                claimsPage.ResetFilter();

                // Create claim 2
                var claimsCreateModalpage1 = claimsPage.ClaimsCreatePage();
                 claims_DateInf = claimsCreateModalpage1.GetIdClaim();
                claimsCreateModalpage1.FillField_CreatNewClaims(fromDateInf, site, supplier, placeTo, true);
                var claimsItem_DateInf = claimsCreateModalpage1.Submit();

                claimsPage = claimsItem_DateInf.BackToList();
                claimsPage.ResetFilter();

                // Create claim 3
                var claimsCreateModalpage2 = claimsPage.ClaimsCreatePage();
                 claims_DateSup = claimsCreateModalpage2.GetIdClaim();
                claimsCreateModalpage2.FillField_CreatNewClaims(fromDateSup, site, supplier, placeTo, true);
                var claimsItem_DateSup = claimsCreateModalpage2.Submit();

                claimsPage = claimsItem_DateSup.BackToList();
                claimsPage.ResetFilter();

                // Validate filtering by DateFrom
                claimsPage.Filter(ClaimsPage.FilterType.ByNumber, claims_Date);
                claimsPage.Filter(ClaimsPage.FilterType.DateFrom, fromDate);
                string dateToCompare1String = claimsPage.GetDate();
                DateTime dateToCompare1 = DateTime.Parse(dateToCompare1String);
                Assert.AreEqual(fromDate.Date, dateToCompare1.Date, "Le filtre Date From ne fonctionne pas correctement .");
                claimsPage.ResetFilter();

                // Validate filtering by Date inf
                claimsPage.Filter(ClaimsPage.FilterType.ByNumber, claims_DateInf);
                claimsPage.Filter(ClaimsPage.FilterType.DateFrom, fromDateInf);
                string dateToCompare2String = claimsPage.GetDate();
                DateTime dateToCompare2 = DateTime.Parse(dateToCompare2String);
                Assert.AreEqual(fromDateInf.Date, dateToCompare2.Date, "Le filtre Date From ne fonctionne pas correctement.");
                claimsPage.ResetFilter();

                // Validate filtering by Date sup
                claimsPage.Filter(ClaimsPage.FilterType.ByNumber, claims_DateSup);
                claimsPage.Filter(ClaimsPage.FilterType.DateFrom, fromDateSup);
                string dateToCompare3String = claimsPage.GetDate();
                DateTime dateToCompare3 = DateTime.Parse(dateToCompare3String);
                Assert.AreEqual(fromDateSup.Date, dateToCompare3.Date, "Le filtre Date From ne fonctionne pas correctement ");
            }
            finally
            {
                claimsPage.ResetFilter();
                claimsPage.Filter(ClaimsPage.FilterType.ByNumber, claims_Date);
                claimsPage.DeleteClaim(claims_Date);
                claimsPage.ResetFilter();
                claimsPage.Filter(ClaimsPage.FilterType.ByNumber, claims_DateInf);
                claimsPage.DeleteClaim(claims_DateInf);
                claimsPage.ResetFilter();
                claimsPage.Filter(ClaimsPage.FilterType.ByNumber, claims_DateSup);
                claimsPage.DeleteClaim(claims_DateSup);


            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_Filter_DateTo()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();
     

            // Log in
            var homePage = LogInAsAdmin();
            var dateFormat = homePage.GetDateFormatPickerValue();

            // Act
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();
            var claims_DateInf = "";
            var claims_DateSup = "";
            var claims_Date = "";
            var dateString = claimsPage.GetDateTo();
            DateTime toDate = DateTime.Parse(dateString);
            DateTime toDateInf = toDate.AddDays(-1);
            DateTime toDateSup = toDate.AddDays(2);
            claimsPage.ResetFilter();
            claimsPage.Filter(ClaimsPage.FilterType.DateFrom, toDate);
            try
            {
                // Create claim 1
                var claimsCreateModalpage = claimsPage.ClaimsCreatePage();
                claims_Date = claimsCreateModalpage.GetIdClaim();
                claimsCreateModalpage.FillField_CreatNewClaims(toDate, site, supplier, placeTo, true);
                var claimsItem_Date = claimsCreateModalpage.Submit();

                claimsPage = claimsItem_Date.BackToList();
                claimsPage.ResetFilter();

                // Create claim 2
                var claimsCreateModalpage1 = claimsPage.ClaimsCreatePage();
                claims_DateInf = claimsCreateModalpage1.GetIdClaim();
                claimsCreateModalpage1.FillField_CreatNewClaims(toDateInf, site, supplier, placeTo, true);
                var claimsItem_DateInf = claimsCreateModalpage1.Submit();

                claimsPage = claimsItem_DateInf.BackToList();
                claimsPage.ResetFilter();

                // Create claim 3
                var claimsCreateModalpage2 = claimsPage.ClaimsCreatePage();
                claims_DateSup = claimsCreateModalpage2.GetIdClaim();
                claimsCreateModalpage2.FillField_CreatNewClaims(toDateSup, site, supplier, placeTo, true);
                var claimsItem_DateSup = claimsCreateModalpage2.Submit();

                claimsPage = claimsItem_DateSup.BackToList();
                claimsPage.ResetFilter();

                // Validate filtering by DateFrom
                claimsPage.Filter(ClaimsPage.FilterType.ByNumber, claims_Date);

                claimsPage.Filter(ClaimsPage.FilterType.DateTo, toDate);
                string dateToCompare1String = claimsPage.GetDate();


                DateTime dateToCompare1 = DateTime.Parse(dateToCompare1String);

                Assert.AreEqual(toDate.Date, dateToCompare1.Date, "Le filtre Date To ne fonctionne pas correctement .");
                claimsPage.ResetFilter();

                // Validate filtering by DateFrom
                claimsPage.Filter(ClaimsPage.FilterType.ByNumber, claims_DateInf);

                claimsPage.Filter(ClaimsPage.FilterType.DateTo, toDateInf);
                string dateToCompare2String = claimsPage.GetDate();


                DateTime dateToCompare2 = DateTime.Parse(dateToCompare2String);

                Assert.AreEqual(toDateInf.Date, dateToCompare2.Date, "Le filtre Date To ne fonctionne pas correctement.");
                claimsPage.ResetFilter();

                // Validate filtering by DateFrom
                claimsPage.Filter(ClaimsPage.FilterType.ByNumber, claims_DateSup);

                claimsPage.Filter(ClaimsPage.FilterType.DateTo, toDateSup);
                string dateToCompare3String = claimsPage.GetDate();


                DateTime dateToCompare3 = DateTime.Parse(dateToCompare3String);

                Assert.AreEqual(toDateSup.Date, dateToCompare3.Date, "Le filtre Date To ne fonctionne pas correctement ");
            }
            finally
            {
                claimsPage.WaitPageLoading();
                claimsPage.Filter(ClaimsPage.FilterType.ByNumber, claims_Date);
                claimsPage.WaitPageLoading();
                claimsPage.DeleteClaim(claims_Date);
                claimsPage.WaitPageLoading();
                claimsPage.Filter(ClaimsPage.FilterType.ByNumber, claims_DateInf);
                claimsPage.WaitPageLoading();
                claimsPage.DeleteClaim(claims_DateInf);
                claimsPage.WaitPageLoading();
                claimsPage.Filter(ClaimsPage.FilterType.ByNumber, claims_DateSup);
                claimsPage.WaitPageLoading();
                claimsPage.DeleteClaim(claims_DateSup);


            }
        }
        /**
         * Test filtrage : Affichage de tous les Claims
         **/
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_Filter_ShowAllClaims()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();
            HomePage homePage = LogInAsAdmin();
            ClaimsPage claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();
            if (claimsPage.CheckTotalNumber() < 20)
            {
                ClaimsCreateModalPage claimsCreateModalpage = claimsPage.ClaimsCreatePage();
                claimsCreateModalpage.FillField_CreatNewClaims(DateUtils.Now, site, supplier, placeTo, true);
                ClaimsItem claimsItem = claimsCreateModalpage.Submit();
                string itemName = claimsItem.GetFirstItemName();
                claimsItem.SelectFirstItem();
                ClaimEditClaimForm editClaimForm = claimsItem.EditClaimForm(itemName);
                editClaimForm.SetClaimFormForClaimedV3();
                claimsPage = claimsItem.BackToList();
                claimsPage.ResetFilter();
            }
            claimsPage.Filter(ClaimsPage.FilterType.ShowValidatedOnly, true);
            int totalValidUnvalidClaimsNumber = claimsPage.CheckTotalNumber();
            claimsPage.Filter(ClaimsPage.FilterType.ShowNotValidated, true);
            totalValidUnvalidClaimsNumber += claimsPage.CheckTotalNumber();
            claimsPage.Filter(ClaimsPage.FilterType.ShowAllClaims, true);
            int totalClaimsNumber = claimsPage.CheckTotalNumber();
            Assert.AreEqual(totalClaimsNumber, totalValidUnvalidClaimsNumber, string.Format(MessageErreur.FILTRE_ERRONE, "Show all claims"));
        }

        /**
         * Test filtrage : Affichage des Receipt Note non validées
         **/
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_Filter_ShowNotValidatedOnly()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();
            HomePage homePage = LogInAsAdmin();
            ClaimsPage claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();
            claimsPage.Filter(ClaimsPage.FilterType.ShowNotValidated, true);
            if (claimsPage.CheckTotalNumber() < 20)
            {
                ClaimsCreateModalPage claimsCreateModalpage = claimsPage.ClaimsCreatePage();
                claimsCreateModalpage.FillField_CreatNewClaims(DateUtils.Now, site, supplier, placeTo, true);
                ClaimsItem claimsItem = claimsCreateModalpage.Submit();
                string itemName = claimsItem.GetFirstItemName();
                claimsItem.SelectFirstItem();
                ClaimEditClaimForm editClaimForm = claimsItem.EditClaimForm(itemName);
                editClaimForm.SetClaimFormForClaimedV3();
                claimsPage = claimsItem.BackToList();
                claimsPage.ResetFilter();
                claimsPage.Filter(ClaimsPage.FilterType.ShowNotValidated, true);
            }
            if (!claimsPage.isPageSizeEqualsTo100())
            {
                claimsPage.PageSize("8");
                claimsPage.PageSize("100");
            }
            bool isChecked = claimsPage.CheckValidation(false);
            Assert.IsFalse(isChecked, string.Format(MessageErreur.FILTRE_ERRONE, "'Show not validated only'"));
        }

        /**
         * Test filtrage : Afficahge des Claims validées
         **/
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_Filter_ShowValidatedOnly()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();
            HomePage homePage = LogInAsAdmin();
            ClaimsPage claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();
            claimsPage.Filter(ClaimsPage.FilterType.ShowValidatedOnly, true);
            if (claimsPage.CheckTotalNumber() < 20)
            {
                ClaimsCreateModalPage claimsCreateModalpage = claimsPage.ClaimsCreatePage();
                claimsCreateModalpage.FillField_CreatNewClaims(DateUtils.Now, site, supplier, placeTo, true);
                ClaimsItem claimsItem = claimsCreateModalpage.Submit();
                string itemName = claimsItem.GetFirstItemName();
                claimsItem.SelectFirstItem();
                claimsItem.AddQuantityAndType(itemName, 2);
                ClaimEditClaimForm editClaimForm = claimsItem.EditClaimForm(itemName);
                editClaimForm.SetClaimFormForClaimedV3();
                claimsItem.Validate();
                claimsPage = claimsItem.BackToList();
                claimsPage.ResetFilter();
                claimsPage.Filter(ClaimsPage.FilterType.ShowValidatedOnly, true);
            }
            if (!claimsPage.isPageSizeEqualsTo100())
            {
                claimsPage.PageSize("8");
                claimsPage.PageSize("100");
            }
            bool isChecked = claimsPage.CheckValidation(true);
            Assert.IsTrue(isChecked, string.Format(MessageErreur.FILTRE_ERRONE, "'Show validated only'"));
        }

        /**
         * Test filtrage : Affichage des Claims partiellements facturées
         **/
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_Filter_ShowValidatedPartial()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string delivery = TestContext.Properties["PlaceFrom"].ToString();
            string supplierInvoiceNumber = new Random().Next().ToString();
            HomePage homePage = LogInAsAdmin();
            ClaimsPage claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();
            claimsPage.Filter(ClaimsPage.FilterType.ShowValidatedPartialInvoiced, true);
            if (claimsPage.CheckTotalNumber() < 20)
            {
                ClaimsCreateModalPage claimsCreateModalpage = claimsPage.ClaimsCreatePage();
                claimsCreateModalpage.FillField_CreatNewClaims(DateUtils.Now, site, supplier, delivery, true);
                ClaimsItem claimsItem = claimsCreateModalpage.Submit();
                string itemName = claimsItem.GetFirstItemName();
                claimsItem.SelectFirstItem();
                claimsItem.AddQuantityAndType(itemName, 2);
                claimsItem.AddPrice(itemName, 20);
                ClaimEditClaimForm editClaimForm = claimsItem.EditClaimForm(itemName);
                editClaimForm.SetClaimFormForClaimedV3();
                claimsItem.Validate();
                SupplierInvoicesItem supplierInvoiceItems = claimsItem.GenerateSupplierInvoice(supplierInvoiceNumber);
                supplierInvoiceItems.SelectFirstItem();
                supplierInvoiceItems.SetItemQuantity(itemName, "10");
                supplierInvoiceItems.Close();
                claimsPage = claimsItem.BackToList();
                claimsPage.ResetFilter();
                claimsPage.Filter(ClaimsPage.FilterType.ShowValidatedPartialInvoiced, true);
            }
            if (!claimsPage.isPageSizeEqualsTo100())
            {
                claimsPage.PageSize("8");
                claimsPage.PageSize("100");
            }
            bool isChecked = claimsPage.CheckValidation(true);
            Assert.IsTrue(isChecked, string.Format(MessageErreur.FILTRE_ERRONE, "'Show validated only'"));
            bool isPartiallyInvoiced = claimsPage.IsPartiallyInvoiced();
            Assert.IsTrue(isPartiallyInvoiced, string.Format(MessageErreur.FILTRE_ERRONE, "'Show validated and partially invoiced only'"));
        }

        /**
         * Test filtrage : Affichage de tous les Claims (actifs + inactifs)
         **/
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_Filter_ShowAll()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();

            if (claimsPage.CheckTotalNumber() < 20)
            {
                // Create
                var claimsCreateModalpage = claimsPage.ClaimsCreatePage();
                claimsCreateModalpage.FillField_CreatNewClaims(DateUtils.Now, site, supplier, placeTo, true);
                var claimsItem = claimsCreateModalpage.Submit();

                // Update the first item value to activate the activation menu
                var itemName = claimsItem.GetFirstItemName();
                claimsItem.SelectFirstItem();

                //Suite a la modif Claim
                var editClaimForm = claimsItem.EditClaimForm(itemName);
                editClaimForm.SetClaimFormForClaimedV3();

                claimsPage = claimsItem.BackToList();
                claimsPage.ResetFilter();
            }

            claimsPage.Filter(ClaimsPage.FilterType.ShowActive, true);
            var activeClaimsNumber = claimsPage.CheckTotalNumber();

            claimsPage.Filter(ClaimsPage.FilterType.ShowAll, true);
            var totalClaimsNumber = claimsPage.CheckTotalNumber();

            claimsPage.Filter(ClaimsPage.FilterType.ShowInactive, true);
            var inactiveClaimsNumber = claimsPage.CheckTotalNumber();

            var totaleActiveIncative = activeClaimsNumber + inactiveClaimsNumber;

            // Assert
            Assert.IsTrue(totalClaimsNumber - totaleActiveIncative == 0, String.Format(MessageErreur.FILTRE_ERRONE, "'Show all'"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_Filter_ShowActiveOnly()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();
            HomePage homePage = LogInAsAdmin();
            ClaimsPage claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();
            claimsPage.Filter(ClaimsPage.FilterType.ShowActive, true);
            if (claimsPage.CheckTotalNumber() < 20)
            {
                ClaimsCreateModalPage claimsCreateModalpage = claimsPage.ClaimsCreatePage();
                claimsCreateModalpage.FillField_CreatNewClaims(DateUtils.Now, site, supplier, placeTo, true);
                ClaimsItem claimsItem = claimsCreateModalpage.Submit();
                string itemName = claimsItem.GetFirstItemName();
                claimsItem.SelectFirstItem();
                ClaimEditClaimForm editClaimForm = claimsItem.EditClaimForm(itemName);
                editClaimForm.SetClaimFormForClaimedV3();
                claimsPage = claimsItem.BackToList();
                claimsPage.ResetFilter();
                claimsPage.Filter(ClaimsPage.FilterType.ShowActive, true);
            }
            if (!claimsPage.isPageSizeEqualsTo100())
            {
                claimsPage.PageSize("8");
                claimsPage.PageSize("100");
            }
            bool isChecked = claimsPage.CheckStatus(true);
            Assert.IsTrue(isChecked, string.Format(MessageErreur.FILTRE_ERRONE, "'Show only active'"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_Filter_ShowInactiveOnly()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();
            HomePage homePage = LogInAsAdmin();
            ClaimsPage claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();
            claimsPage.Filter(ClaimsPage.FilterType.ShowInactive, true);
            if (claimsPage.CheckTotalNumber() < 20)
            {
                ClaimsCreateModalPage claimsCreateModalpage = claimsPage.ClaimsCreatePage();
                claimsCreateModalpage.FillField_CreatNewClaims(DateUtils.Now, site, supplier, placeTo, false);
                ClaimsItem claimsItem = claimsCreateModalpage.Submit();
                string itemName = claimsItem.GetFirstItemName();
                claimsItem.SelectFirstItem();
                claimsItem.AddQuantityAndType(itemName, 2);
                ClaimEditClaimForm editClaimForm = claimsItem.EditClaimForm(itemName);
                editClaimForm.SetClaimFormForClaimedV3();
                claimsPage = claimsItem.BackToList();
                claimsPage.ResetFilter();
                claimsPage.Filter(ClaimsPage.FilterType.ShowInactive, true);
            }
            if (!claimsPage.isPageSizeEqualsTo100())
            {
                claimsPage.PageSize("8");
                claimsPage.PageSize("100");
            }
            bool isChecked = claimsPage.CheckStatus(false);
            Assert.IsFalse(isChecked, string.Format(MessageErreur.FILTRE_ERRONE, "'Show only inactive'"));
        }

        /*
         * Test filtrage : Par numéro et par DeliveryDate
         */
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_Filter_SortByNumber()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();
            string ID = string.Empty;
            //Arrange
            var homePage = LogInAsAdmin();

            var dateFormat = homePage.GetDateFormatPickerValue();

            //Act
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();
            int totalNumber = claimsPage.CheckTotalNumber();
            try
            {
                if (totalNumber < 20)
                {
                    // Create
                    var claimsCreateModalpage = claimsPage.ClaimsCreatePage();
                    claimsCreateModalpage.FillField_CreatNewClaims(DateUtils.Now, site, supplier, placeTo, true);
                    var claimsItem = claimsCreateModalpage.Submit();

                    // Update the first item value to activate the activation menu
                    var itemName = claimsItem.GetFirstItemName();
                    claimsItem.SelectFirstItem();

                    //Suite a la modif Claim
                    var editClaimForm = claimsItem.EditClaimForm(itemName);
                    editClaimForm.SetClaimFormForClaimedV3();

                    claimsPage = claimsItem.BackToList();
                    ID =claimsPage.GetFirstClaimNumber();
                    claimsPage.ResetFilter();
                }

                if (!claimsPage.isPageSizeEqualsTo100())
                {
                    claimsPage.PageSize("8");
                    claimsPage.PageSize("100");
                }

                claimsPage.Filter(ClaimsPage.FilterType.SortBy, "NUMBER");
                var isSortedByNumber = claimsPage.IsSortedByNumber();

                //Assert
                Assert.IsTrue(isSortedByNumber, String.Format(MessageErreur.FILTRE_ERRONE, "'Sort by number'"));
            }
            finally
            {
                if(totalNumber < 20)
                {
                    claimsPage = homePage.GoToWarehouse_ClaimsPage();
                    claimsPage.Filter(FilterType.ByNumber, ID);
                    claimsPage.DeleteClaim(ID);
                }
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_Filter_SortByDeliveryDate()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();
            string ID = string.Empty;
            //Arrange
            var homePage = LogInAsAdmin();

            var dateFormat = homePage.GetDateFormatPickerValue();

            //Act
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();
            int totalNumber = claimsPage.CheckTotalNumber();
            try
            {
                if (totalNumber < 20)
                {
                    // Create
                    var claimsCreateModalpage = claimsPage.ClaimsCreatePage();
                    claimsCreateModalpage.FillField_CreatNewClaims(DateUtils.Now, site, supplier, placeTo, true);
                    var claimsItem = claimsCreateModalpage.Submit();

                    // Update the first item value to activate the activation menu
                    var itemName = claimsItem.GetFirstItemName();
                    claimsItem.SelectFirstItem();

                    //Suite a la modif Claim
                    var editClaimForm = claimsItem.EditClaimForm(itemName);
                    editClaimForm.SetClaimFormForClaimedV3();

                    claimsPage = claimsItem.BackToList();
                    ID = claimsPage.GetFirstClaimNumber();
                    claimsPage.ResetFilter();
                }

                if (!claimsPage.isPageSizeEqualsTo100())
                {
                    claimsPage.PageSize("8");
                    claimsPage.PageSize("100");
                }

                claimsPage.Filter(ClaimsPage.FilterType.SortBy, "DATE");
                var isSortedByDate = claimsPage.IsSortedByDate(dateFormat);

                //Assert
                Assert.IsTrue(isSortedByDate, String.Format(MessageErreur.FILTRE_ERRONE, "'Sort by delivery date'"));
            }
            finally
            {
                if (totalNumber < 20)
                {
                    claimsPage = homePage.GoToWarehouse_ClaimsPage();
                    claimsPage.Filter(FilterType.ByNumber, ID);
                    claimsPage.DeleteClaim(ID);
                }
            }
        }

        /*
         * Test filtrage : Ouvert
         */
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_Filter_Opened()
        {
            HomePage homePage = LogInAsAdmin();
            ClaimsPage claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();
            int result = claimsPage.CheckTotalNumber();
            claimsPage.Filter(ClaimsPage.FilterType.DateFrom, DateUtils.Now.AddMonths(-1));
            claimsPage.Filter(ClaimsPage.FilterType.DateTo, DateUtils.Now);
            claimsPage.Filter(ClaimsPage.FilterType.Opened, true);
            int resultAfterFilterOpened = claimsPage.CheckTotalNumber();
            Assert.AreNotEqual(result, resultAfterFilterOpened, "Erreur de filtrage Statut Opened !");
        }

        /*
         * Test filtrage : Closed
         */
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_Filter_Closed()
        {
            HomePage homePage = LogInAsAdmin();
            ClaimsPage claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();
            int result = claimsPage.CheckTotalNumber();
            claimsPage.Filter(ClaimsPage.FilterType.DateTo, DateUtils.Now);
            claimsPage.Filter(ClaimsPage.FilterType.DateFrom, DateUtils.Now.AddMonths(-1));
            claimsPage.Filter(ClaimsPage.FilterType.Closed, true);
            int resultAfterFilterOpened = claimsPage.CheckTotalNumber();
            Assert.AreNotEqual(result, resultAfterFilterOpened, "Erreur de filtrage Statut Closed !");
        }

        /**
         * Test filtrage : Affichage de l'ensemble des Claims
         **/
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_Filter_Show_AllStatus()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();
            HomePage homePage = LogInAsAdmin();
            ClaimsPage claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();
            claimsPage.Filter(ClaimsPage.FilterType.All, true);
            if (claimsPage.CheckTotalNumber() < 20)
            {
                ClaimsCreateModalPage claimsCreateModalpage = claimsPage.ClaimsCreatePage();
                claimsCreateModalpage.FillField_CreatNewClaims(DateUtils.Now, site, supplier, placeTo, true);
                ClaimsItem claimsItem = claimsCreateModalpage.Submit();
                string itemName = claimsItem.GetFirstItemName();
                claimsItem.SelectFirstItem();
                ClaimEditClaimForm editClaimForm = claimsItem.EditClaimForm(itemName);
                editClaimForm.SetClaimFormForClaimedV3();
                claimsPage = claimsItem.BackToList();
                claimsPage.ResetFilter();
            }
            claimsPage.Filter(ClaimsPage.FilterType.Opened, true);
            int totalOpenedClosedClaimsNumber = claimsPage.CheckTotalNumber();
            claimsPage.Filter(ClaimsPage.FilterType.Closed, true);
            totalOpenedClosedClaimsNumber += claimsPage.CheckTotalNumber();
            claimsPage.Filter(ClaimsPage.FilterType.All, true);
            int totalClaimsNumber = claimsPage.CheckTotalNumber();
            Assert.AreEqual(totalClaimsNumber, totalOpenedClosedClaimsNumber, String.Format(MessageErreur.FILTRE_ERRONE, "'Show all status'"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_Filter_Sites()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();

            HomePage homePage = LogInAsAdmin();
            ClaimsPage claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();
            claimsPage.Filter(ClaimsPage.FilterType.Site, site);
            if (claimsPage.CheckTotalNumber() < 20)
            {
                ClaimsCreateModalPage claimsCreateModalpage = claimsPage.ClaimsCreatePage();
                claimsCreateModalpage.FillField_CreatNewClaims(DateUtils.Now, site, supplier, placeTo, true);
                ClaimsItem claimsItem = claimsCreateModalpage.Submit();
                string itemName = claimsItem.GetFirstItemName();
                claimsItem.SelectFirstItem();
                ClaimEditClaimForm editClaimForm = claimsItem.EditClaimForm(itemName);
                editClaimForm.SetClaimFormForClaimedV3();
                claimsPage = claimsItem.BackToList();
                claimsPage.ResetFilter();
                claimsPage.Filter(ClaimsPage.FilterType.Site, site);
            }
            if (!claimsPage.isPageSizeEqualsTo100())
            {
                claimsPage.PageSize("8");
                claimsPage.PageSize("100");
            }
            bool isVerified = claimsPage.VerifySite(site);
            Assert.IsTrue(isVerified, string.Format(MessageErreur.FILTRE_ERRONE, "'Sites'"));
        }


        //_________________________________________FIN FILTERS TESTING_________________________________________________________

        //_________________________________________CLAIMS ITEM_________________________________________________________

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_Items_Filters_SearchByName()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();
            string ID = string.Empty;
            // Log in
            var homePage = LogInAsAdmin();

            //Act
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();

            try
            {
                // Create
                var claimsCreateModalpage = claimsPage.ClaimsCreatePage();
                claimsCreateModalpage.FillField_CreatNewClaims(DateUtils.Now, site, supplier, placeTo, true);
                ID = claimsCreateModalpage.GetClaimId();
                var claimsItem = claimsCreateModalpage.Submit();

                // Filter by item name / ref
                var itemName = claimsItem.GetFirstItemName();

                claimsItem.Filter(ClaimsItem.FilterItemType.SearchByName, itemName);
                Assert.IsTrue(claimsItem.VerifyName(itemName), MessageErreur.FILTRE_ERRONE, "'Search item by name'");
                claimsItem.BackToList();
            }
            finally
            {
                // Filter and Delete Claim Created
                claimsPage.ResetFilter();
                claimsPage.Filter(ClaimsPage.FilterType.ShowNotValidated, true);
                claimsPage.Filter(ClaimsPage.FilterType.ByNumber, ID);
                claimsPage.DeleteClaim(ID);
            }
           



        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_Items_Filters_SearchByKeyword()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();
            string itemKeyword = TestContext.Properties["Item_Keyword"].ToString();
            HomePage homePage = LogInAsAdmin();
            homePage.SetNewVersionKeywordValue(true);
            ClaimsPage claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();
            ClaimsCreateModalPage claimsCreateModalpage = claimsPage.ClaimsCreatePage();
            claimsCreateModalpage.FillField_CreatNewClaims(DateUtils.Now, site, supplier, placeTo, true);
            ClaimsItem claimsItem = claimsCreateModalpage.Submit();
            string itemName = claimsItem.GetFirstItemName();
            claimsItem.SelectFirstItem();
            claimsItem.AddPrice(itemName, 10);
            claimsItem.AddQuantityAndType(itemName, 2);
            ClaimEditClaimForm editClaimForm = claimsItem.EditClaimForm(itemName);
            editClaimForm.SetClaimFormForClaimedV3();
            claimsItem.SelectFirstItem();
            ItemGeneralInformationPage itemPageItem = claimsItem.EditItemGeneralInformation(itemName);
            PageObjects.Purchasing.Item.ItemKeywordPage keywordTab = itemPageItem.ClickOnKeywordItem();
            keywordTab.AddKeyword(itemKeyword);
            itemPageItem.Close();
            claimsItem.PageSize("8");
            claimsItem.Filter(ClaimsItem.FilterItemType.SearchByKeyword, itemKeyword);
            bool isKeywordOK = false;
            string errorMessageKeyword = "";
            try
            {
                isKeywordOK = claimsItem.VerifyKeyword(itemKeyword);
            }
            catch (Exception e)
            {
                errorMessageKeyword = e.Message;
            }

            Assert.IsTrue(isKeywordOK, errorMessageKeyword);
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_Items_Filters_ShowNotClaimed()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();
            HomePage homePage = LogInAsAdmin();
            ClaimsPage claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();
            ClaimsCreateModalPage claimsCreateModalpage = claimsPage.ClaimsCreatePage();
            claimsCreateModalpage.FillField_CreatNewClaims(DateUtils.Now, site, supplier, placeTo, true);
            ClaimsItem claimsItem = claimsCreateModalpage.Submit();
            claimsItem.SelectFirstItem();
            claimsItem.Filter(ClaimsItem.FilterItemType.ShowNotClaimed, true);
            if (!claimsItem.isPageSizeEqualsTo100WidhoutTotalNumber())
            {
                claimsItem.PageSize("8");
                claimsItem.PageSize("100");
            }
            bool isClaimed = claimsItem.IsClaimed();
            Assert.IsTrue(isClaimed, string.Format(MessageErreur.FILTRE_ERRONE, "'Show items not claimed'"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_Items_Filters_ByGroup()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();
            HomePage homePage = LogInAsAdmin();
            homePage.SetNewVersionKeywordValue(true);
            ClaimsPage claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();
            ClaimsCreateModalPage claimsCreateModalpage = claimsPage.ClaimsCreatePage();
            claimsCreateModalpage.FillField_CreatNewClaims(DateUtils.Now, site, supplier, placeTo, true);
            ClaimsItem claimsItem = claimsCreateModalpage.Submit();
            string itemName = claimsItem.GetFirstItemName();
            claimsItem.SelectFirstItem();
            ItemGeneralInformationPage itemPageItem = claimsItem.EditItemGeneralInformation(itemName);
            string itemGroup = itemPageItem.GetGroupName();
            itemPageItem.Close();
            claimsItem.PageSize("8");
            claimsItem.Filter(ClaimsItem.FilterItemType.ByGroup, itemGroup);
            bool isVerified = claimsItem.VerifyGroup(itemGroup);
            Assert.IsTrue(isVerified, String.Format(MessageErreur.FILTRE_ERRONE, "'Group'"));
            bool isSorted = claimsItem.IsSortedByName();
            Assert.IsTrue(isSorted, "Les items du groupe ne sont pas triés par ordre alphabétique.");
        }

        /*
         * Création d'une nouvelle Claim et modification de celle-ci
        */
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_UpdateValueClaims()
        {
            //prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string delivery = TestContext.Properties["PlaceFrom"].ToString();
            string NumberOfCreatedClaim = String.Empty;

            //arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();
            try
            {
                // Create
                var claimsCreateModalpage = claimsPage.ClaimsCreatePage();
                claimsCreateModalpage.FillField_CreatNewClaims(DateUtils.Now, site, supplier, delivery, true);
                NumberOfCreatedClaim = claimsCreateModalpage.GetClaimId();
                var claimsItem = claimsCreateModalpage.Submit();

                // Update the first item value to activate the activation menu
                var itemName = claimsItem.GetFirstItemName();
                var startPrice = claimsItem.GetPrice();
                claimsItem.SelectFirstItem();
                claimsItem.AddPrice(itemName, 10);
                claimsItem.AddQuantityAndType(itemName, 2);

                //Suite a la modif Claim
                var editClaimForm = claimsItem.EditClaimForm(itemName);
                editClaimForm.SetClaimFormForClaimedV3();

                //Assert
                Assert.AreNotEqual(startPrice, claimsItem.GetPrice(), "L'update de l'item a échoué.");
            }
            finally
            {
                claimsPage.BackToList();
                if (!String.IsNullOrEmpty(NumberOfCreatedClaim))
                {
                    claimsPage.Filter(ClaimsPage.FilterType.ByNumber, NumberOfCreatedClaim);
                    claimsPage.DeleteClaim(NumberOfCreatedClaim);
                }

                claimsPage.ResetFilter();
            }
        }

        /*
         * Mise à jour direct d'un item lié a une claim
        */
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_UpdateItemClaims()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string delivery = TestContext.Properties["PlaceFrom"].ToString();

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();

            // Create
            var claimsCreateModalpage = claimsPage.ClaimsCreatePage();
            claimsCreateModalpage.FillField_CreatNewClaims(DateUtils.Now, site, supplier, delivery, true);
            var claimsItem = claimsCreateModalpage.Submit();

            // Update the first item value
            var itemName = claimsItem.GetFirstItemName();
            claimsItem.SelectFirstItem();

            try
            {
                var ClaimItemDetailPage = claimsItem.EditItemGeneralInformation(itemName);
                ClaimItemDetailPage.SetName(itemName + " TEST");
                ClaimItemDetailPage.Close();
                claimsItem.Refresh();

                //Test
                Assert.AreNotEqual(itemName, claimsItem.GetFirstItemName(), "L'update de la claim a échoué.");
            }
            finally
            {
                claimsItem.SelectFirstItem();
                var claimItemDetailPage = claimsItem.EditItemGeneralInformation(itemName + " TEST");
                while (itemName.EndsWith(" TEST"))
                {
                    itemName = itemName.Substring(0, itemName.Length - " TEST".Length);
                    Console.WriteLine(itemName);
                }
                claimItemDetailPage.SetName(itemName);
            }
        }

        /*
         * Mise à jour photo d'une claim
        */
        [TestMethod]
        [DeploymentItem("PageObjects\\Warehouse\\Claims\\test.jpg")]
        [DeploymentItem("chromedriver.exe")]
        [Timeout(_timeout)]
        public void WA_CLAI_UpdatePictureClaims()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string delivery = TestContext.Properties["PlaceFrom"].ToString();

            var imagePath = TestContext.TestDeploymentDir + "\\test.jpg";

            Assert.IsTrue(new FileInfo(imagePath).Exists, "image test.jpg non trouvé");

            // Log in
            var homePage = LogInAsAdmin();

            //Act
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();

            // Create
            var claimsCreateModalpage = claimsPage.ClaimsCreatePage();
            claimsCreateModalpage.FillField_CreatNewClaims(DateUtils.Now, site, supplier, delivery, true);
            var claimsItem = claimsCreateModalpage.Submit();

            // Update the first item value to activate the activation menu
            var itemName = claimsItem.GetFirstItemName();
            claimsItem.SelectFirstItem();
            claimsItem.AddPrice(itemName, 10);
            claimsItem.AddQuantityAndType(itemName, 2);

            //
            var editClaimForm = claimsItem.EditClaimForm(itemName);
            editClaimForm.SetClaimFormForClaimedV3(imagePath);
            claimsItem.SelectFirstItem();
            claimsItem.EditClaimForm(itemName);

            //Assert
            Assert.IsTrue(editClaimForm.IsPictureAdded(), MessageErreur.ERREUR_UPLOAD_PICTURE);
        }

        /*
         * Suppression d'une Claims note
        */
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_RemoveclaimsItem()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();
            // Log in
            var homePage = LogInAsAdmin();
            var decimalSeparator = homePage.GetDecimalSeparatorValue();

            //Act
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();

            // Create
            var claimsCreateModalpage = claimsPage.ClaimsCreatePage();
            claimsCreateModalpage.FillField_CreatNewClaims(DateUtils.Now, site, supplier, placeTo, true);
            var claimsItem = claimsCreateModalpage.Submit();

            var itemName = claimsItem.GetFirstItemName();

            claimsItem.SelectFirstItem();
            claimsItem.AddPrice(itemName, 10);
            claimsItem.AddQuantityAndType(itemName, 2);

            //Suite a la modif Claim
            var editClaimForm = claimsItem.EditClaimForm(itemName);
            editClaimForm.SetClaimFormForClaimedV3();

            // Récupération des valeurs de l'item
            var oldPrice = claimsItem.GetPrice().Substring(2, claimsItem.GetPrice().IndexOf(decimalSeparator) - 2);
            var oldType = claimsItem.GetClaimType();
            var oldQty = claimsItem.GetQuantity();

            Assert.AreEqual("10", oldPrice);
            try
            {
                Assert.AreEqual("Non-compliant delivered <b>quantity</b>", oldType);
            }
            catch
            {
                Assert.AreEqual("Non-compliant delivered quantity", oldType);
            }
            Assert.AreEqual("2", oldQty);

            // Suppression de l'item
            claimsItem.SelectFirstItem();
            claimsItem.DeleteItem(itemName);

            claimsItem.Filter(ClaimsItem.FilterItemType.SearchByName, itemName);


            var newPrice = claimsItem.GetPrice().Substring(2, claimsItem.GetPrice().IndexOf(decimalSeparator) - 2);
            var newType = claimsItem.GetClaimType();
            var newQty = claimsItem.GetQuantity();

            Assert.AreNotEqual(newPrice, oldPrice);
            Assert.AreEqual("", newType);
            Assert.AreEqual("", newQty);
        }

        /**
         * Test d'ajout d'un commentaire sur une Claims
         **/
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_CommentClaims()
        {
            // Create a Reiceipt Note in first 

            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string delivery = TestContext.Properties["PlaceFrom"].ToString();
            string comment = "I am a comment";
            string ID = string.Empty;
            // Log in
            var homePage = LogInAsAdmin();

            //Act
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();

            try
            {
                // Create
                var claimsCreateModalpage = claimsPage.ClaimsCreatePage();
                claimsCreateModalpage.FillField_CreatNewClaims(DateUtils.Now, site, supplier, delivery, true);
                ID = claimsCreateModalpage.GetClaimId();
                var claimsItem = claimsCreateModalpage.Submit();
                //claimsItem

                // Update the first item value to activate the activation menu
                var itemName = claimsItem.GetFirstItemName();
                claimsItem.SelectFirstItem();
                claimsItem.AddQuantityAndType(itemName, 2);

                //Suite a la modif Claim
                var editClaimForm = claimsItem.EditClaimForm(itemName);
                editClaimForm.SetClaimFormForClaimedV3();

                claimsItem.SelectFirstItem();
                claimsItem.AddComment(itemName, comment);

                // Make the final test
                string commentFromEditMenu = claimsItem.GetCommentFromEditMenu();
                Assert.AreEqual(commentFromEditMenu, comment, "Le commentaire n'est pas affiché dans l'input Comment lors de l'édition de la réclamation..");

                //Check comment in popup comment
                claimsItem.SelectFirstItem();
                string commentAdded = claimsItem.GetComment(itemName);
                Assert.AreEqual(commentAdded, comment, "Le commentaire n'est pas affiché dans l'input Comment de la popup Comment.");
                claimsItem.CloseModalComment();
                claimsItem.BackToList();
            }
            finally
            {
                claimsPage.ResetFilter();
                claimsPage.Filter(ClaimsPage.FilterType.ShowNotValidated, true);
                claimsPage.Filter(ClaimsPage.FilterType.ByNumber, ID);
                claimsPage.DeleteClaim(ID);
            }
        }

        /*
         * Test du menu d'export (print, send by mail, Generate Supplier invoice) d'une Claim avec newVersionPrint = true
         */
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_ClaimsOptionMenuNewVersion()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string delivery = TestContext.Properties["PlaceFrom"].ToString();

            Random rnd = new Random();
            var supplierInvoiceNumber = rnd.Next().ToString();

            bool newVersionPrint = true;

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();

            claimsPage.ClearDownloads();

            // Create
            var claimsCreateModalpage = claimsPage.ClaimsCreatePage();
            claimsCreateModalpage.FillField_CreatNewClaims(DateUtils.Now, site, supplier, delivery, true);
            String ID = claimsCreateModalpage.GetClaimId();
            var claimsItem = claimsCreateModalpage.Submit();

            // Update the first item value to activate the activation menu
            var itemName = claimsItem.GetFirstItemName();
            claimsItem.SelectFirstItem();
            claimsItem.AddQuantityAndType(itemName, 2);

            //Suite a la modif Claim
            var editClaimForm = claimsItem.EditClaimForm(itemName);
            editClaimForm.SetClaimFormForClaimedV3();

            claimsItem.Validate();

            // Impression de la liste des items
            var reportPage = claimsItem.PrintClaimItems(newVersionPrint);
            var isReportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();
            Assert.IsTrue(isReportGenerated, "L'application n'a pas pu générer le fichier attendu.");

            claimsItem.SendByMail();

            //Mail assert
            claimsPage = claimsItem.BackToList();
            claimsPage.Filter(ClaimsPage.FilterType.ByNumber, ID);

            bool isSentByMail = claimsPage.IsSentByMail();
            Assert.IsTrue(isSentByMail, "La claim n'a pas été envoyée par mail.");
            claimsItem = claimsPage.SelectFirstClaim();

            //Generate supplier invoice
            var supplierInvoiceItems = claimsItem.GenerateSupplierInvoice(supplierInvoiceNumber);
            supplierInvoiceItems.ValidateSupplierInvoice();

            var supplierInvoicePage = supplierInvoiceItems.BackToList();
            supplierInvoicePage.Filter(SupplierInvoicesPage.FilterType.ByNumber, supplierInvoiceNumber);

            var supplierNumber = supplierInvoicePage.GetFirstInvoiceNumber();
            supplierInvoicePage.Close();

            Assert.AreEqual(supplierNumber, supplierInvoiceNumber, "Erreur lors du processus de création du supplier invoice associé à la claim.");
        }

        /*
         * Création d'une nouvelle Claim et validation de celle-ci
        */
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_ValidateClaims()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string delivery = TestContext.Properties["PlaceFrom"].ToString();
            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();

            // Create
            var claimsCreateModalpage = claimsPage.ClaimsCreatePage();
            claimsCreateModalpage.FillField_CreatNewClaims(DateUtils.Now, site, supplier, delivery, true);
            String ID = claimsCreateModalpage.GetClaimId();
            var claimsItem = claimsCreateModalpage.Submit();

            // Update the first item value to activate the activation menu
            var itemName = claimsItem.GetFirstItemName();
            claimsItem.SelectFirstItem();
            claimsItem.AddQuantityAndType(itemName, 2);

            //Suite a la modif Claim
            var editClaimForm = claimsItem.EditClaimForm(itemName);
            editClaimForm.SetClaimFormForClaimedV3();
            claimsItem.Validate();

            claimsPage = claimsItem.BackToList();
            claimsPage.ResetFilter();
            claimsPage.Filter(ClaimsPage.FilterType.ByNumber, ID);
            var checkValidation = claimsPage.CheckValidation(true);
            Assert.IsTrue(checkValidation, String.Format(MessageErreur.FILTRE_ERRONE, "'Show validated only'"));
        }

        //_________________________________________FIN CLAIMS ITEM_________________________________________________________

        //_________________________________________CLAIMS EXPORT_________________________________________________________

        /*
         * Export de la liste des résultats d'une recherche de Claims avec newVersionPrint = true
        */
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_ExportClaimsListNewVersion()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string delivery = TestContext.Properties["PlaceFrom"].ToString();
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            bool newVersionPrint = true;

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var decimalSeparator = homePage.GetDecimalSeparatorValue();

            //Act
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();

            claimsPage.ClearDownloads();

            if (claimsPage.CheckTotalNumber() < 20)
            {
                // Create
                var claimsCreateModalpage = claimsPage.ClaimsCreatePage();
                claimsCreateModalpage.FillField_CreatNewClaims(DateUtils.Now, site, supplier, delivery, true);
                var claimsItem = claimsCreateModalpage.Submit();

                var itemName = claimsItem.GetFirstItemName();
                claimsItem.SelectFirstItem();

                //Suite a la modif Claim
                var editClaimForm = claimsItem.EditClaimForm(itemName);
                editClaimForm.SetClaimFormForClaimedV3();

                claimsPage = claimsItem.BackToList();
                claimsPage.ResetFilter();
            }

            // Lancement de l'export avec la première valeur de printValue
            ExportGenerique(claimsPage, newVersionPrint, downloadsPath, decimalSeparator);
        }

        private void ExportGenerique(ClaimsPage claimsPage, bool printValue, string downloadsPath, string decimalSeparator)
        {
            claimsPage.Filter(ClaimsPage.FilterType.ShowNotValidated, true);
            var firstDate = claimsPage.GetFirstDate() != null ? DateTime.ParseExact(claimsPage.GetFirstDate(), "dd/MM/yyyy", CultureInfo.InvariantCulture) : DateTime.Now;

            claimsPage.Filter(ClaimsPage.FilterType.DateFrom, firstDate);
            claimsPage.Filter(ClaimsPage.FilterType.DateTo, firstDate.AddMonths(1));
            //FIXME non testé
            // ajouter un claim au premier claim
            ClaimsItem firstClaim = claimsPage.SelectFirstClaim();
            // Update the first item value to activate the activation menu
            var itemName = firstClaim.GetFirstItemName();
            firstClaim.SelectFirstItem();
            //Suite a la modif Claim
            var editClaimForm = firstClaim.EditClaimForm(itemName);
            editClaimForm.SetClaimFormForClaimedV3();
            claimsPage = firstClaim.BackToList();

            DeleteAllFileDownload();

            // On exporte les résultats sous la forme d'un fichier Excel (dont on récupère le nom)
            // Export du fichier au format Excel
            claimsPage.ExportExcelFile(printValue, downloadsPath);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = claimsPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

            // Récupération du nom du fichier et construction de l'URL du fichier Excel à ouvrir   
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber("Claims", filePath);

            //Assert
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);

            firstClaim = claimsPage.SelectFirstClaim();
            ClaimsGeneralInformation generalInfo = firstClaim.ClickOnGeneralInformation();

            // general info
            string claimNumber = generalInfo.GetClaimNumber();
            string claimComment = generalInfo.GetClaimComment();
            string claimSite = generalInfo.GetClaimSite();
            string claimSupplier = generalInfo.GetClaimSupplier();
            string claimDeliveryLocation = generalInfo.GetClaimDeliveryLocation();
            string claimDeliveryOrderNumber = generalInfo.GetDeliveryOrderNumber();
            string claimStatus = generalInfo.GetStatus();

            int offset = OpenXmlExcel.GetValuesInList("Number", "Claims", filePath).IndexOf(claimNumber);
            Assert.AreNotEqual(-1, offset, "ID " + claimNumber + " non trouvé dans la colonne Number");

            string number = OpenXmlExcel.GetValuesInList("Number", "Claims", filePath)[offset];
            string comment = OpenXmlExcel.GetValuesInList("Comment", "Claims", filePath)[offset];
            string site = OpenXmlExcel.GetValuesInList("Site", "Claims", filePath)[offset];
            string supplier = OpenXmlExcel.GetValuesInList("Supplier", "Claims", filePath)[offset];
            string deliveryLocation = OpenXmlExcel.GetValuesInList("Site Place", "Claims", filePath)[offset];
            string deliveryOrderNumber = OpenXmlExcel.GetValuesInList("Delivery N°", "Claims", filePath)[offset];
            string status = OpenXmlExcel.GetValuesInList("Status", "Claims", filePath)[offset];
            Assert.AreEqual(claimNumber, number, "XLS mauvais Number");
            Assert.AreEqual(claimComment, comment, "XLS mauvais Comment");
            Assert.AreEqual(claimSite, site, "XLS mauvais site");
            Assert.AreEqual(claimSupplier, supplier, "XLS mauvais supplier");
            Assert.AreEqual(claimDeliveryLocation, deliveryLocation, "XLS mauvais supplier");
            Assert.AreEqual(claimDeliveryOrderNumber, deliveryOrderNumber, "XLS mauvais delivery order number");
            if (claimStatus == "1")
            {
                Assert.AreEqual("Opened", status, "XLS mauvais status");
            }

            ClaimsItem claimsItem = generalInfo.ClickOnItems();
            //items
            string claimGroupName = claimsItem.GetFirstGroupName() + " ";
            string claimName = claimsItem.GetFirstItemName();
            string claimPackaging = claimsItem.GetPackaging();
            string claimType = claimsItem.GetClaimType();
            string claimPriceReceived = claimsItem.GetPrice(); // Pack Price
            string claimReceivedQty = claimsItem.GetQuantity(); // Received Quantity
            string claimPriceDN = claimsItem.GetDNPrice();
            string claimQtyDN = claimsItem.GetDNQuantity();

            string groupName = OpenXmlExcel.GetValuesInList("Group Name", "Claims", filePath)[offset].Trim();
            string name = OpenXmlExcel.GetValuesInList("Name", "Claims", filePath)[offset];
            string packaging = OpenXmlExcel.GetValuesInList("Packaging Unit", "Claims", filePath)[offset];
            string type = OpenXmlExcel.GetValuesInList("Claim Name", "Claims", filePath)[offset].Replace("\n", "\r\n");
            string priceReceived = OpenXmlExcel.GetValuesInList("Received Price", "Claims", filePath)[offset];
            string priceDN = OpenXmlExcel.GetValuesInList("DN Price", "Claims", filePath)[offset];
            string qtyReceived = OpenXmlExcel.GetValuesInList("Received Quantity", "Claims", filePath)[offset];
            string qtyReceivedRelated = OpenXmlExcel.GetValuesInList("Related RN received quantity", "Claims", filePath)[offset];
            // Claim from RNs
            string qtyDN = OpenXmlExcel.GetValuesInList("DN Quantity", "Claims", filePath)[offset];

            Assert.IsTrue(claimGroupName.Trim().ToUpper() == groupName.Trim().ToUpper(), "XLS mauvais group");
            Assert.IsTrue(claimName.Trim().ToUpper() == name.Trim().ToUpper(), "XLS mauvais name");
            // Attendu : <Caja 1000 UD>, Réel : <Caja 1000 UD 1 UD>. XLS mauvais packaging
            // "bolsa 1Kg 10 KG 10 UD" "bolsa 1Kg 10 KG (10 UD)"
            Assert.IsTrue(packaging.Contains(claimPackaging.Replace("(", "").Replace(")", "")), "XLS mauvais packaging");
            Assert.IsTrue(type.Contains(claimType), "XLS mauvais type");
            Assert.AreEqual(ArrangeTarif(claimPriceReceived, decimalSeparator), ArrangeTarif(priceReceived, decimalSeparator), "XLS mauvais price");
            //int.Parse(qty) + int.Parse(qty2) ?
            if (int.Parse(qtyReceived) > 0)
            {
                Assert.AreEqual(int.Parse(claimReceivedQty), int.Parse(qtyReceived), "XLS mauvais qty");
            }
            else if (int.Parse(qtyReceivedRelated) > 0)
            {
                Assert.AreEqual(int.Parse(claimReceivedQty), int.Parse(qtyReceivedRelated), "XLS mauvais qty");
            }
            else
            {
                Assert.AreEqual(0, int.Parse(claimReceivedQty), "XLS mauvais qty");
            }
            Assert.AreEqual(ArrangeTarif(claimPriceDN, decimalSeparator), ArrangeTarif(priceDN, decimalSeparator), "XLS mauvais DN Price");
            Assert.AreEqual(int.Parse(claimQtyDN == "" ? "0" : claimQtyDN), int.Parse(qtyDN == "" ? "0" : qtyDN), "XLS mauvais DN qty");

            ClaimsFooter footer = claimsItem.ClickOnFooter();
            // footer
            double claimTotalNoVat = footer.GetTotalNoVat();
            string claimVat = footer.GetVat();
            string claimTotalWithVat = footer.GetTotalWithVat();
            string totalNoVat = OpenXmlExcel.GetValuesInList("Total No VAT", "Claims", filePath)[offset];
            string vat = OpenXmlExcel.GetValuesInList("VAT", "Claims", filePath)[offset];
            string totalWithVat = OpenXmlExcel.GetValuesInList("Total with VAT", "Claims", filePath)[offset];
            var totalNoValue = ArrangeTarif(totalNoVat, decimalSeparator);
            Assert.AreEqual(claimTotalNoVat, totalNoValue, "XLS mauvais total no VAT");
            var claimVatValue = ArrangeTarif(claimVat, decimalSeparator);
            var vatValue = ArrangeTarif(vat, decimalSeparator);
            Assert.AreEqual(claimVatValue, vatValue, "XLS mauvais VAT");

            var claimTotalWithVatValue = ArrangeTarif(claimTotalWithVat, decimalSeparator);
            var totalWithVatValue = ArrangeTarif(totalWithVat, decimalSeparator);
            Assert.AreEqual(claimTotalWithVatValue, totalWithVatValue, "XLS mauvais total with VAT");
        }

        /*
         * Test d'impression d'une nouvelle Claim avec newVersionPrint = true
        */
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_PrintClaimsNewVersion()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string delivery = TestContext.Properties["PlaceFrom"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            //Receipt Note Report_-_437542_-_20220316093948.pdf
            string DocFileNamePdfBegin = "Receipt Note Report_-_";
            //All_files_20220225_102148.zip
            string DocFileNameZipBegin = "All_files_";

            // Log in
            var homePage = LogInAsAdmin();

            //Act
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();

            claimsPage.ClearDownloads();
            DeleteAllFileDownload();

            if (claimsPage.CheckTotalNumber() < 20)
            {
                // Create
                var claimsCreateModalpage = claimsPage.ClaimsCreatePage();
                claimsCreateModalpage.FillField_CreatNewClaims(DateUtils.Now, site, supplier, delivery, true);
                var claimsItem = claimsCreateModalpage.Submit();

                // Update the first item value to activate the activation menu
                var itemName = claimsItem.GetFirstItemName();
                claimsItem.SelectFirstItem();

                //Suite a la modif Claim
                var editClaimForm = claimsItem.EditClaimForm(itemName);
                editClaimForm.SetClaimFormForClaimedV3();

                claimsPage = claimsItem.BackToList();
                claimsPage.ResetFilter();
            }
            var firstDate = claimsPage.GetFirstDate() != null ? DateTime.ParseExact(claimsPage.GetFirstDate(), "dd/MM/yyyy", CultureInfo.InvariantCulture) : DateTime.Now;

            claimsPage.Filter(ClaimsPage.FilterType.DateFrom, firstDate);
            claimsPage.Filter(ClaimsPage.FilterType.DateTo, firstDate.AddMonths(1));
            var number = claimsPage.GetFirstID();
            if (!String.IsNullOrEmpty(number))
            {
                claimsPage.Filter(ClaimsPage.FilterType.ByNumber, number);
            }

            var printReportPage = claimsPage.PrintClaims();
            var isReportGenerated = printReportPage.IsReportGenerated();
            printReportPage.Close();


            //Assert
            Assert.IsTrue(isReportGenerated, "L'application n'a pas pu générer le fichier attendu.");

            printReportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);

            // cliquer sur All
            string trouve = printReportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            FileInfo fi = new FileInfo(trouve);
            fi.Refresh();
            Assert.IsTrue(fi.Exists, trouve + " non généré");

            ClaimsItem claimsFirstItem = claimsPage.SelectFirstClaim();
            ClaimsGeneralInformation generalInfo = claimsFirstItem.ClickOnGeneralInformation();
            string claimNumber = generalInfo.GetClaimNumber();
            string deliveryOrderNumber = generalInfo.GetDeliveryOrderNumber();
            string claimDate = generalInfo.GetClaimDate();

            PdfDocument document = PdfDocument.Open(fi.FullName);
            Page page1 = document.GetPage(1);
            IEnumerable<Word> words = page1.GetWords();
            var isContainSite = words.Where(w => w.Text.Contains(site + "(" + site + ")")).Count();
            Assert.AreEqual(1, isContainSite, site + "(" + site + ") non présent dans le Pdf");
            var numberOfDate = words.Count(w => w.Text == claimDate);
            Assert.AreEqual(1, numberOfDate, claimDate + " non présent dans le Pdf");
            var numberOfClaim = words.Count(w => w.Text == claimNumber);
            Assert.AreEqual(1, numberOfClaim, claimNumber + " non présent dans le Pdf");
            var numberOfdeliveryOrderNumber = words.Count(w => w.Text == deliveryOrderNumber);
            Assert.AreEqual(1, numberOfdeliveryOrderNumber, deliveryOrderNumber + " non présent dans le Pdf");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_SendResultsByMail()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string delivery = TestContext.Properties["PlaceFrom"].ToString();
            string userEmail = TestContext.Properties["Admin_UserName"].ToString();

            // Log in
            var homePage = LogInAsAdmin();

            //Act
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();

            // Create
            var claimsCreateModalpage = claimsPage.ClaimsCreatePage();
            claimsCreateModalpage.FillField_CreatNewClaims(DateUtils.Now, site, supplier, delivery, true);
            String ID = claimsCreateModalpage.GetClaimId();
            var claimsItem = claimsCreateModalpage.Submit();

            // Update the first item value to activate the activation menu
            var itemName = claimsItem.GetFirstItemName();
            claimsItem.SelectFirstItem();
            claimsItem.AddQuantityAndType(itemName, 2);


            ////Suite modif Validation Claim
            var editClaimForm = claimsItem.EditClaimForm(itemName);
            editClaimForm.SetClaimFormForClaimedV3();

            claimsItem.Validate();

            claimsPage = claimsItem.BackToList();
            claimsPage.ResetFilter();
            claimsPage.Filter(ClaimsPage.FilterType.ByNumber, ID);
            claimsPage.Filter(ClaimsPage.FilterType.DateFrom, DateUtils.Now);
            claimsPage.Filter(ClaimsPage.FilterType.DateTo, DateUtils.Now.AddMonths(1));

            bool isSent = claimsPage.SendByMail();
            Assert.IsTrue(isSent, "L'envoi par mail a été effectué.");
            MailPage mailPage = claimsPage.RedirectToOutlookMailbox();

            mailPage.FillFields_LogToOutlookMailbox(userEmail);

            bool checkIfSpecifiedOutlookMailExist = mailPage.CheckIfSpecifiedOutlookMailExist("OC NEWREST - 1 Claim note" + " " + site);
            Assert.IsTrue(checkIfSpecifiedOutlookMailExist, "La claim n'a pas été envoyée par mail.");
            mailPage.DeleteCurrentOutlookMail();
            WebDriver.Navigate().Refresh();
            mailPage.Close();

            bool checkEnveloppe = claimsPage.CheckEnveloppe(ID);
            Assert.IsTrue(checkEnveloppe, "pas d'enveloppe");
        }

        /*
         * Test d'impression des items d'une Claim avec newVersionPrint = true
        */
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_PrintClaimItemsNewVersion()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string delivery = TestContext.Properties["PlaceFrom"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            //Claim Note Report_-_438018_-_20220316082719.pdf
            string DocFileNamePdfBegin = "Claim Note Report_-_";
            //All_files_20220225_102148.zip
            string DocFileNameZipBegin = "All_files_";

            bool newVersionPrint = true;

            // Log in
            var homePage = LogInAsAdmin();

            //Act
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();

            claimsPage.ClearDownloads();

            // Create
            var claimsCreateModalpage = claimsPage.ClaimsCreatePage();
            claimsCreateModalpage.FillField_CreatNewClaims(DateUtils.Now, site, supplier, delivery, true);
            var claimsItem = claimsCreateModalpage.Submit();

            ClaimsGeneralInformation generalInfo = claimsItem.ClickOnGeneralInformation();
            string claimNumber = generalInfo.GetClaimNumber();
            string deliveryOrderNumber = generalInfo.GetDeliveryOrderNumber();
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);

            //clean directory
            foreach (var file in taskDirectory.GetFiles())
            {
                file.Delete();
            }
            claimsPage.ClearDownloads();
            DeleteAllFileDownload();

            // Impression de la liste des items
            var reportPage = claimsItem.PrintClaimItems(newVersionPrint);
            var isReportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();

            Assert.IsTrue(isReportGenerated, "L'application n'a pas pu générer le fichier attendu.");

            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);

            // cliquer sur All
            string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            FileInfo fi = new FileInfo(trouve);
            fi.Refresh();
            Assert.IsTrue(fi.Exists, trouve + " non généré");

            PdfDocument document = PdfDocument.Open(fi.FullName);
            Page page1 = document.GetPage(1);
            IEnumerable<Word> words = page1.GetWords();
            Assert.AreEqual(1, words.Where(w => w.Text.Contains(site + "(" + site + ")")).Count(), site + "(" + site + ") non présent dans le Pdf");
            Assert.AreEqual(1, words.Count(w => w.Text == DateUtils.Now.Date.ToString("dd/MM/yyyy")), DateUtils.Now.Date.ToString("dd/MM/yyyy") + " non présent dans le Pdf");
            Assert.AreEqual(1, words.Count(w => w.Text == claimNumber), claimNumber + " non présent dans le Pdf");
            Assert.AreEqual(1, words.Count(w => w.Text == deliveryOrderNumber), deliveryOrderNumber + " non présent dans le Pdf");

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_Items_DecreaseStockAndQuantity()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string delivery = TestContext.Properties["PlaceFrom"].ToString();
            string ID = string.Empty;
            // Log in
            var homePage = LogInAsAdmin();

            //Act
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();

            try
            {
                // Create
                // Etre sur un claim non validée
                var claimsCreateModalpage = claimsPage.ClaimsCreatePage();
                claimsCreateModalpage.FillField_CreatNewClaims(DateUtils.Now, site, supplier, delivery, true);
                ID = claimsCreateModalpage.GetClaimId();
                ClaimsItem claimsItem = claimsCreateModalpage.Submit();

                string itemName = claimsItem.GetFirstItemName();
                claimsItem.SelectFirstItem();
                //1. Ajouter une quantité sur un item puis cocher decr. stock et ajouter une quantité sur Decreased Qty
                claimsItem.AddQuantityAndType(itemName, 4);

                //2. AJouter une claim et vérifier la decr. stock et la quantité sur Decreased Qty
                ClaimEditClaimForm claimForm = claimsItem.EditClaimForm(itemName);
                // Non-compliant delivered deselectionne decr
                claimForm.FillClaimFormCheckbox("Non compliant expiration date");
                claimsItem = claimForm.Save();
                claimsItem.SelectFirstItem();
                claimsItem.ClickDecrStockCheckBox();
                claimsItem.SetDecrQty("2");
                Assert.IsTrue(claimsItem.IsDecrStock(), "mauvaise coche");
                Assert.AreEqual("2", claimsItem.GetDecrQty());

                //3 Modifier la decr. stock et la quantité sur Decreased Qty dans la claim
                claimsItem.AddQuantityAndType(itemName, 5);
                claimsItem.SetDecrQty("3");

                claimsItem.WaitPageLoading();

                //4. Vérifier les nouvelles valeurs sur l'item
                ClaimsGeneralInformation generalInfo = claimsItem.ClickOnGeneralInformation();
                claimsItem = generalInfo.ClickOnItems();
                Assert.AreEqual(itemName, claimsItem.GetFirstItemName());
                claimsItem.SelectFirstItem();
                Assert.IsTrue(claimsItem.IsDecrStock(), "mauvaise coche");
                Assert.AreEqual("3", claimsItem.GetDecrQty(), "mauvaise desc qty");
                Assert.AreEqual("5", claimsItem.GetFormQty(), "mauvaise qty");
                claimsItem.BackToList();
            }
            finally
            {
                claimsPage.ResetFilter();
                claimsPage.Filter(ClaimsPage.FilterType.ByNumber, ID);
                claimsPage.Filter(ClaimsPage.FilterType.ShowAll, true);
                claimsPage.DeleteClaim(ID);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_Items_EditItem()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string delivery = TestContext.Properties["PlaceFrom"].ToString();

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();

            //1.Créer une claim
            var claimsCreateModalpage = claimsPage.ClaimsCreatePage();
            claimsCreateModalpage.FillField_CreatNewClaims(DateUtils.Now, site, supplier, delivery, true);
            ClaimsItem claimsItem = claimsCreateModalpage.Submit();
            string itemName = claimsItem.GetFirstItemName();
            //2. Cliquer sur un item
            claimsItem.SelectFirstItem();
            //3.Cliquer sur... puis sur l'icone crayon
            ItemGeneralInformationPage itemGeneralInfo = claimsItem.EditItem();

            //Vérifier que l'on ouvre le bon item
            Assert.AreEqual(itemName, itemGeneralInfo.GetItemName());
            itemGeneralInfo.Close();
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_Items_VerifyClaimAmount()
        {
            HomePage homePage = LogInAsAdmin();
            string decimalSeparator = homePage.GetDecimalSeparatorValue();
            ClaimsPage claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();
            claimsPage.Filter(ClaimsPage.FilterType.ShowNotValidated, true);
            ClaimsItem claimItems = claimsPage.SelectFirstClaim();
            string itemName = claimItems.GetFirstItemName();
            claimItems.SelectFirstItem();
            claimItems.AddQuantityAndType(itemName, 4);
            claimItems.Refresh();
            double price = ArrangeTarif(claimItems.GetPrice(), decimalSeparator);
            double quantity = ArrangeTarif(claimItems.GetQuantity(), decimalSeparator);
            double priceDN = ArrangeTarif(claimItems.GetDNPrice(), decimalSeparator);
            double quantityDN = ArrangeTarif(claimItems.GetDNQuantity(), decimalSeparator);
            double total = ArrangeTarif(claimItems.GetClaimTotalAmount(), decimalSeparator);
            Assert.AreEqual(4.0d, quantity);
            Assert.AreEqual((priceDN * quantityDN) - (price * quantity), total, 0.0001, "Mauvais total claim");
            claimItems.SelectFirstItem();
            ClaimEditClaimForm editClaim = claimItems.EditClaimForm(itemName);
            double amount = ArrangeTarif(editClaim.WaitForElementIsVisible(By.Id("spanClaimAmount")).Text, decimalSeparator);
            Assert.AreEqual(total, amount, "Mauvais claim amount");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_Items_VerifySanctionAmount()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var decimalSeparator = homePage.GetDecimalSeparatorValue();

            //Act
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();
            // Etre sur un claim non validée

            //claimsPage.Filter(ClaimsPage.FilterType.ShowNotValidated, true);
            //ClaimsItem claimItems = claimsPage.SelectFirstClaim();
            ClaimsCreateModalPage newClaimModal = claimsPage.ClaimsCreatePage();
            newClaimModal.FillField_CreatNewClaims(DateUtils.Now, site, supplier, placeTo, true);
            String ID = newClaimModal.GetClaimId();
            ClaimsItem claimItems = newClaimModal.Submit();


            string itemName = claimItems.GetFirstItemName();

            //1. Ajouter une quantité sur un item puis une sanction amount
            claimItems.SelectFirstItem();
            claimItems.AddQuantityAndType(itemName, 4);
            string sanction = "6" + decimalSeparator + "4";
            claimItems.AddSanction(sanction);

            //2. AJouter une claim et vérifier la sanction amount
            ClaimEditClaimForm editClaim = claimItems.EditClaimForm(itemName);
            editClaim.FillClaimFormCheckbox("Non compliant expiration date");
            var sanctionClaim = editClaim.WaitForElementIsVisible(By.Id("SanctionAmount"));
            Assert.AreEqual(sanction, sanctionClaim.GetAttribute("value"));
            var selectVAT = editClaim.WaitForElementIsVisible(By.Id("TaxTypeId"));
            var selectVATOption = new SelectElement(selectVAT);
            Assert.AreEqual("1-Reducido", selectVATOption.SelectedOption.Text);
            //3 Modifier la sanction amount dans la claim
            sanctionClaim.Clear();
            string sanctionMaj = "12" + decimalSeparator + "5";
            sanctionClaim.SendKeys(sanctionMaj);
            claimItems = editClaim.Save();
            claimItems.SelectFirstItem();

            //4.Vérifier la nouvelle valeur sur l'item
            string sanctionUpdate = claimItems.GetSanction();
            Assert.AreEqual(sanctionMaj, sanctionUpdate);

            //5. Vérifier également le total sanction
            string sanctionTotalUpdate = claimItems.GetTotalSanction();
            claimItems.Refresh();
            // 1-Reducido 10%
            Assert.AreEqual(ArrangeTarif(sanctionUpdate, decimalSeparator) + (10.0d * ArrangeTarif(sanctionUpdate, decimalSeparator)) / 100.0d, ArrangeTarif(sanctionTotalUpdate, decimalSeparator));

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_Items_VerifyTotalDNQuantity()
        {
            //prepare
            double DNPrice1 = 4.0d;
            double priceItem = 3.3d;
            double DNPrice2 = 2.0d;
            double qtyDNItem1 = 4.0d;
            double qtyItem = 3.0d;
            double qtyDNItem2 = 2.0d;
            double DNQuantity1Edited = 8.0d;
            double DNQuantity2Edited = 9.0d;
            var site = TestContext.Properties["Site"].ToString();
            var date = DateTime.Now;
            var supplier = TestContext.Properties["Supplier"].ToString();
            var placeTo = TestContext.Properties["InventoryPlace"].ToString();
            // Log in
            var homePage = LogInAsAdmin();
            var decimalSeparator = homePage.GetDecimalSeparatorValue();
            string numberOfCreatedClaim = String.Empty;
            int offset = 0;
            //Act
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();
            try
            {
                // creé une claim non validé
                ClaimsCreateModalPage claimNotValidates = claimsPage.ClaimsCreatePage();
                claimNotValidates.FillField_CreatNewClaims(date, site, supplier, placeTo);
                numberOfCreatedClaim = claimNotValidates.GetClaimId();
                ClaimsItem claimItems = claimNotValidates.Submit();
                claimItems.Filter(ClaimsItem.FilterItemType.ShowNotClaimed, true);
                var itemName1 = claimItems.GetItemForClaim(1);
                var itemName2 = claimItems.GetItemForClaim(2);
                // 1. Ajouter 2 items, ajouter DN quantity et price
                claimItems.Filter(ClaimsItem.FilterItemType.SearchByName, itemName1);
                string itemName = claimItems.GetFirstItemName();
                claimItems.SelectFirstItem();
                Assert.AreEqual(itemName, itemName1, "item  non trouvé (cas 1)");
                claimItems.SetPrice(priceItem.ToString(), offset);
                claimItems.SetDNPrice(DNPrice1.ToString(), offset);
                claimItems.SetQuantity(qtyItem.ToString(), offset);
                claimItems.SetDNQuantity(qtyDNItem1.ToString(), offset);
                claimItems.WaitPageLoading();
                claimItems.Refresh();
                claimItems.Filter(ClaimsItem.FilterItemType.SearchByName, itemName2);
                itemName = claimItems.GetFirstItemName();
                claimItems.SelectFirstItem();
                Assert.AreEqual(itemName, itemName2, "item non trouvé (cas 2)");
                claimItems.SetPrice(priceItem.ToString(), offset);
                claimItems.SetDNPrice(DNPrice2.ToString(), offset);
                claimItems.SetQuantity(qtyItem.ToString(), offset);
                claimItems.SetDNQuantity(qtyDNItem2.ToString(), offset);
                claimItems.WaitPageLoading();
                claimItems.ResetFilters();
                claimItems.Filter(ClaimsItem.FilterItemType.ShowNotClaimed, false);
                claimItems.Refresh();
                var claimTotalAmount1BeforeEdit = ArrangeTarif(claimItems.GetClaimTotalAmount(0), decimalSeparator);
                // update the first item
                claimItems.SelectFirstItem();
                claimItems.SetDNQuantity(DNQuantity1Edited.ToString(), offset);
                claimItems.WaitPageLoading();
                claimItems.Refresh();
                double qtyDN1 = ArrangeTarif(claimItems.GetDNQuantity(0), decimalSeparator);
                double priceDN1 = ArrangeTarif(claimItems.GetDNPrice(2), decimalSeparator);
                double qty1 = ArrangeTarif(claimItems.GetQuantity(0), decimalSeparator);
                double price1 = ArrangeTarif(claimItems.GetPrice(2), decimalSeparator);
                double total1 = qtyDN1 * priceDN1 - qty1 * price1;
                var claimTotalAmount1AfterEdit = ArrangeTarif(claimItems.GetClaimTotalAmount(0), decimalSeparator);
                Assert.AreEqual(total1, claimTotalAmount1AfterEdit, 0.0005d, "mauvais total claim 1");
                Assert.AreEqual(qtyDN1, DNQuantity1Edited, "mauvais DNQuantity 1");
                Assert.AreNotEqual(claimTotalAmount1BeforeEdit, claimTotalAmount1AfterEdit, "mauvais total claim of item1 after edit");
                // update the second item
                var claimTotalAmount2BeforeEdit = ArrangeTarif(claimItems.GetClaimTotalAmount(0), decimalSeparator);
                claimItems.SelectSecondItem();
                offset = 1;
                claimItems.SetDNQuantity(DNQuantity2Edited.ToString(), offset);
                claimItems.WaitPageLoading();
                claimItems.Refresh();
                double qtyDN2 = ArrangeTarif(claimItems.GetDNQuantity(1), decimalSeparator);
                double priceDN2 = ArrangeTarif(claimItems.GetDNPrice(3), decimalSeparator);
                double qty2 = ArrangeTarif(claimItems.GetQuantity(1), decimalSeparator);
                double price2 = ArrangeTarif(claimItems.GetPrice(3), decimalSeparator);
                double total2 = qtyDN2 * priceDN2 - qty2 * price2;
                var claimTotalAmount2AfterEd = ArrangeTarif(claimItems.GetClaimTotalAmount(1), decimalSeparator);
                Assert.AreEqual(total2, claimTotalAmount2AfterEd, 0.0005d, "mauvais total claim 2");
                Assert.AreEqual(qtyDN2, DNQuantity2Edited, "mauvais DNQuantity 2");
                Assert.AreNotEqual(claimTotalAmount2BeforeEdit, claimTotalAmount2AfterEd, "mauvais total claim of item2 after edit");
                //check the total
                var total = total1 + total2;
                var totalClaim = ArrangeTarif(claimItems.GetTotalClaim(), decimalSeparator);
                Assert.AreEqual(total, totalClaim, 0.0005d, "le total n'est pas correct");
            }
            finally
            {
                claimsPage.BackToList();
                claimsPage.ResetFilter();
                claimsPage.Filter(ClaimsPage.FilterType.ShowNotValidated, true);
                claimsPage.Filter(ClaimsPage.FilterType.ByNumber, numberOfCreatedClaim);
                claimsPage.DeleteClaim(numberOfCreatedClaim);
            }
        }
        
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_Items_VerifyTotalDNPrice()
        {
            //prepare
            double DNPrice1 = 4.0d;
            double priceItem = 3.3d;
            double DNPrice2 = 2.0d;
            double qtyDNItem1 = 4.0d;
            double qtyItem = 3.0d;
            double qtyDNItem2 = 2.0d;
            double DNPrice1Edited = 8.0d;
            double DNPrice2Edited = 9.0d;
            var site = TestContext.Properties["Site"].ToString();
            var date = DateTime.Now;
            var supplier = TestContext.Properties["Supplier"].ToString(); 
            var placeTo = TestContext.Properties["InventoryPlace"].ToString();
            // Log in
            var homePage = LogInAsAdmin();
            var decimalSeparator = homePage.GetDecimalSeparatorValue();
            string numberOfCreatedClaim = String.Empty;
            int offset = 0;
            //Act
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();
            try
            {
                // creé une claim non validé
                ClaimsCreateModalPage claimNotValidates = claimsPage.ClaimsCreatePage();
                claimNotValidates.FillField_CreatNewClaims(date, site, supplier, placeTo);
                numberOfCreatedClaim = claimNotValidates.GetClaimId();
                ClaimsItem claimItems = claimNotValidates.Submit();
                claimItems.Filter(ClaimsItem.FilterItemType.ShowNotClaimed, true);
                var itemName1 = claimItems.GetItemForClaim(1);
                var itemName2 = claimItems.GetItemForClaim(2);
                // 1. Ajouter 2 items, ajouter DN quantity et price
                claimItems.Filter(ClaimsItem.FilterItemType.SearchByName, itemName1);
                string itemName = claimItems.GetFirstItemName();
                claimItems.SelectFirstItem();
                Assert.AreEqual(itemName, itemName1, "item  non trouvé (cas 1)");
                claimItems.SetPrice(priceItem.ToString(), offset);
                claimItems.SetDNPrice(DNPrice1.ToString(), offset);
                claimItems.SetQuantity(qtyItem.ToString(), offset);
                claimItems.SetDNQuantity(qtyDNItem1.ToString(), offset);
                claimItems.WaitPageLoading();
                claimItems.Refresh();
                claimItems.Filter(ClaimsItem.FilterItemType.SearchByName, itemName2);
                itemName = claimItems.GetFirstItemName();
                claimItems.SelectFirstItem();
                Assert.AreEqual(itemName, itemName2, "item non trouvé (cas 2)");
                claimItems.SetPrice(priceItem.ToString(), offset);
                claimItems.SetDNPrice(DNPrice2.ToString(), offset);
                claimItems.SetQuantity(qtyItem.ToString(), offset);
                claimItems.SetDNQuantity(qtyDNItem2.ToString(), offset);
                claimItems.WaitPageLoading();
                claimItems.ResetFilters();
                claimItems.Filter(ClaimsItem.FilterItemType.ShowNotClaimed, false);
                claimItems.Refresh();
                var claimTotalAmount1BeforeEdit = ArrangeTarif(claimItems.GetClaimTotalAmount(0), decimalSeparator);
                // update the first item
                claimItems.SelectFirstItem();
                claimItems.SetDNPrice(DNPrice1Edited.ToString(), offset);
                claimItems.WaitPageLoading();
                claimItems.Refresh();
                double qtyDN1 = ArrangeTarif(claimItems.GetDNQuantity(0), decimalSeparator);
                double priceDN1 = ArrangeTarif(claimItems.GetDNPrice(2), decimalSeparator);
                double qty1 = ArrangeTarif(claimItems.GetQuantity(0), decimalSeparator);
                double price1 = ArrangeTarif(claimItems.GetPrice(2), decimalSeparator);
                double total1 = qtyDN1 * priceDN1 - qty1 * price1;
                var claimTotalAmount1AfterEdit = ArrangeTarif(claimItems.GetClaimTotalAmount(0), decimalSeparator);
                Assert.AreEqual(total1, claimTotalAmount1AfterEdit, 0.0005d, "mauvais total claim 1");
                Assert.AreEqual(priceDN1, DNPrice1Edited, "mauvais DNPrice 1");
                Assert.AreNotEqual(claimTotalAmount1BeforeEdit, claimTotalAmount1AfterEdit, "mauvais total claim of item1 after edit");
                // update the second item
                var claimTotalAmount2BeforeEdit = ArrangeTarif(claimItems.GetClaimTotalAmount(0), decimalSeparator);
                claimItems.WaitPageLoading();
                claimItems.SelectSecondItem();
                offset = 1;
                claimItems.SetDNPrice(DNPrice2Edited.ToString(), offset);
                claimItems.WaitPageLoading();
                claimItems.Refresh();
                double qtyDN2 = ArrangeTarif(claimItems.GetDNQuantity(1), decimalSeparator);
                double priceDN2 = ArrangeTarif(claimItems.GetDNPrice(3), decimalSeparator);
                double qty2 = ArrangeTarif(claimItems.GetQuantity(1), decimalSeparator);
                double price2 = ArrangeTarif(claimItems.GetPrice(3), decimalSeparator);
                double total2 = qtyDN2 * priceDN2 - qty2 * price2;
                var claimTotalAmount2AfterEd = ArrangeTarif(claimItems.GetClaimTotalAmount(1), decimalSeparator);
                Assert.AreEqual(total2, claimTotalAmount2AfterEd, 0.0005d, "mauvais total claim 2");
                Assert.AreEqual(priceDN2, DNPrice2Edited, "mauvais DNPrice 2");
                Assert.AreNotEqual(claimTotalAmount2BeforeEdit, claimTotalAmount2AfterEd, "mauvais total claim of item2 after edit");
                var total = total1 + total2;
                var totalClaim = ArrangeTarif(claimItems.GetTotalClaim(), decimalSeparator);
                Assert.AreEqual(total, totalClaim, 0.0005d, "le total n'est pas correct");
            }
            finally
            {
                // delete the claim
                claimsPage.BackToList();
                claimsPage.ResetFilter();
                claimsPage.Filter(ClaimsPage.FilterType.ShowNotValidated, true);
                claimsPage.Filter(ClaimsPage.FilterType.ByNumber, numberOfCreatedClaim);
                claimsPage.DeleteClaim(numberOfCreatedClaim);
            }

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_Reset_Filter()
        {
            //prepare
            object value = null;
            // Login and Go to warehouse Claims
            HomePage homePage = LogInAsAdmin();
            ClaimsPage claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();

            // filter with Show all and Reset filter
            claimsPage.Filter(ClaimsPage.FilterType.ShowAll, true);
            int claimsNumberShowAll = claimsPage.CheckTotalNumber();
            bool isNotEqual = claimsNumberShowAll != 0;
            Assert.IsTrue(isNotEqual, MessageErreur.FILTRE_ERRONE);
            claimsPage.ResetFilter();

            // filter with Show active and Reset filter
            claimsPage.Filter(ClaimsPage.FilterType.ShowActive, true);
            int claimsNumberShowActive = claimsPage.CheckTotalNumber();
            isNotEqual = claimsNumberShowAll != claimsNumberShowActive;
            Assert.IsTrue(isNotEqual, MessageErreur.FILTRE_ERRONE);
            claimsPage.ResetFilter();

            // filter with Show inactive and Reset filter
            claimsPage.Filter(ClaimsPage.FilterType.ShowInactive, true);
            int claimsNumberShowInactive = claimsPage.CheckTotalNumber();
            isNotEqual = claimsNumberShowAll != claimsNumberShowInactive;
            Assert.IsTrue(isNotEqual, MessageErreur.FILTRE_ERRONE);
            claimsPage.ResetFilter();

            // filter with Date From and Reset filter
            claimsPage.Filter(ClaimsPage.FilterType.DateFrom, DateTime.Now.AddDays(5));
            int claimsNumberDateFrom = claimsPage.CheckTotalNumber();
            isNotEqual = claimsNumberShowAll != claimsNumberDateFrom;
            Assert.IsTrue(isNotEqual, MessageErreur.FILTRE_ERRONE);
            claimsPage.ResetFilter();

            // filter with Date To and Reset filter
            claimsPage.Filter(ClaimsPage.FilterType.DateTo, DateTime.Now.AddDays(-5));
            int claimsNumberDateTo = claimsPage.CheckTotalNumber();
            isNotEqual = claimsNumberShowAll != claimsNumberDateTo;
            Assert.IsTrue(isNotEqual, MessageErreur.FILTRE_ERRONE);
            claimsPage.ResetFilter();

            // filter with Site and Reset filter
            claimsPage.Filter(ClaimsPage.FilterType.Site, TestContext.Properties["Site"].ToString());
            int claimsNumberSite = claimsPage.CheckTotalNumber();
            isNotEqual = claimsNumberShowAll != claimsNumberSite;
            Assert.IsTrue(isNotEqual, MessageErreur.FILTRE_ERRONE);
            claimsPage.ResetFilter();

            // filter with Show all Claims and Reset filter
            claimsPage.Filter(ClaimsPage.FilterType.ShowAllClaims, true);
            value = claimsPage.GetFilterValue(ClaimsPage.FilterType.ShowAllClaims);
            Assert.AreEqual(true, value, "ResetFilter ShowAll");
            claimsPage.ResetFilter();

            // filter with Show not Validated and Reset filter
            claimsPage.Filter(ClaimsPage.FilterType.ShowNotValidated, true);
            value = claimsPage.GetFilterValue(ClaimsPage.FilterType.ShowNotValidated);
            Assert.AreEqual(true, value, "filtre Show Not Validated non reset");
            claimsPage.ResetFilter();

            // filter with Show Validated Only and Reset filter
            claimsPage.Filter(ClaimsPage.FilterType.ShowValidatedOnly, true);
            value = claimsPage.GetFilterValue(ClaimsPage.FilterType.ShowValidatedOnly);
            Assert.AreEqual(true, value, "filtre Show  Validated non reset");
            claimsPage.ResetFilter();

            // filter with Show Validated Partial Invoiced and Reset filter
            claimsPage.Filter(ClaimsPage.FilterType.ShowValidatedPartialInvoiced, true);
            value = claimsPage.GetFilterValue(ClaimsPage.FilterType.ShowValidatedPartialInvoiced);
            Assert.AreEqual(true, value, "filtre Show Validated Partial non reset");
            claimsPage.ResetFilter();

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_Filter_StatusAll()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();
            string itemNumClosed = "";
            string itemNumOpened = "";
            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            //Etre sur l'index des Claim et avoir des données disponibles
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();

            claimsPage.Filter(ClaimsPage.FilterType.All, true);
            claimsPage.WaitPageLoading();
            int nbAll = claimsPage.CheckTotalNumber();
            claimsPage.ResetFilter();

            claimsPage.Filter(ClaimsPage.FilterType.Closed, true);
            claimsPage.WaitPageLoading();
            int nbClosed = claimsPage.CheckTotalNumber();

            claimsPage.Filter(ClaimsPage.FilterType.Opened, true);
            claimsPage.WaitPageLoading();
            int nbOpened = claimsPage.CheckTotalNumber();

            try
            {
                if (nbClosed < 1)
                {
                    var claimCreatePage = claimsPage.ClaimsCreatePage();
                    claimCreatePage.FillField_CreatNewClaims(DateTime.Now, site, supplier, placeTo, true);
                    ClaimsItem claimsItem = claimCreatePage.Submit();
                    var generalInformationPage = claimsItem.ClickOnGeneralInformation();
                    itemNumClosed = generalInformationPage.GetClaimNumber();
                    generalInformationPage.SetStatus("Closed");
                    generalInformationPage.BackToList();

                }
                if (nbOpened < 1)
                {
                    var claimCreatePage = claimsPage.ClaimsCreatePage();
                    claimCreatePage.FillField_CreatNewClaims(DateTime.Now, site, supplier, placeTo, true);
                    ClaimsItem claimsItem = claimCreatePage.Submit();
                    var generalInformationPage = claimsItem.ClickOnGeneralInformation();
                    itemNumOpened = generalInformationPage.GetClaimNumber();
                    generalInformationPage.SetStatus("Opened");
                    generalInformationPage.BackToList();
                }

                // Appliquer le filtre Statut ALL
                claimsPage.Filter(ClaimsPage.FilterType.All, true);
                claimsPage.WaitPageLoading();
                nbAll = claimsPage.CheckTotalNumber();

                //Vérifier que l'addition des filtre Closed et Opened correspondent au Filtre ALL
                claimsPage.Filter(ClaimsPage.FilterType.Opened, true);
                claimsPage.WaitPageLoading();
                var allNbOpened = claimsPage.CheckTotalNumber();

                claimsPage.Filter(ClaimsPage.FilterType.Closed, true);
                claimsPage.WaitPageLoading();
                var allNbClosed = claimsPage.CheckTotalNumber();
                Assert.AreEqual(nbAll, allNbClosed + allNbOpened, "Problème filtre All!=Opened+Closed");
            }
            finally
            {
                claimsPage.Filter(ClaimsPage.FilterType.All, true);
                if (nbClosed < 1)
                {
                    claimsPage.DeleteClaim(itemNumClosed);    

                }
                if (nbOpened < 1)
                {
                    claimsPage.DeleteClaim(itemNumOpened);
                }
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_Filter_StatusClosed()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();
            string itemNumClosed = "";
            string itemNumOpened = "";
            string closedStatus = "Closed";
            string openedStatus = "Opened";
            int toCheck0 = 0;
            int toCheck1 = 1;
            DateTime date = DateTime.Now;
            bool isActivate = true;

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            //Etre sur l'index des Claim et avoir des données disponibles
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();

            try
            {
                //Create claim closed
                var claimCreatePage = claimsPage.ClaimsCreatePage();
                claimCreatePage.FillField_CreatNewClaims(date, site, supplier, placeTo, isActivate);
                ClaimsItem claimsItem = claimCreatePage.Submit();
                var generalInformationPage = claimsItem.ClickOnGeneralInformation();
                itemNumClosed = generalInformationPage.GetClaimNumber();
                generalInformationPage.SetStatus(closedStatus);
                generalInformationPage.WaitPageLoading();
                generalInformationPage.BackToList();

                //create claim opened
                claimCreatePage = claimsPage.ClaimsCreatePage();
                claimCreatePage.FillField_CreatNewClaims(date, site, supplier, placeTo, isActivate);
                claimsItem = claimCreatePage.Submit();
                generalInformationPage = claimsItem.ClickOnGeneralInformation();
                itemNumOpened = generalInformationPage.GetClaimNumber();
                generalInformationPage.SetStatus(openedStatus);
                generalInformationPage.BackToList();

                //Filtrage closed
                claimsPage.Filter(ClaimsPage.FilterType.Closed, true);
                claimsPage.Filter(ClaimsPage.FilterType.ByNumber, itemNumClosed);
                claimsPage.WaitPageLoading();
                var nbClosed = claimsPage.CheckTotalNumber();
                Assert.AreEqual(nbClosed, toCheck1, "Problème avec le filtre Closed");

                claimsPage.Filter(ClaimsPage.FilterType.Closed, true);
                claimsPage.Filter(ClaimsPage.FilterType.ByNumber, itemNumOpened);
                claimsPage.WaitPageLoading();
                var nbOpened = claimsPage.CheckTotalNumber();
                Assert.AreEqual(nbOpened, toCheck0, "Problème avec le filtre Closed");

            }
            finally
            {
                //delete
                claimsPage.Filter(ClaimsPage.FilterType.All, true);
                claimsPage.Filter(ClaimsPage.FilterType.ByNumber, itemNumClosed);
                claimsPage.WaitPageLoading();
                claimsPage.DeleteClaim(itemNumClosed);
                claimsPage.Filter(ClaimsPage.FilterType.ByNumber, itemNumOpened);
                claimsPage.WaitPageLoading();
                claimsPage.DeleteClaim(itemNumOpened);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_Filter_StatusOpened()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();
            string itemNumClosed = "";
            string itemNumOpened = "";
            string closedStatus = "Closed";
            string openedStatus = "Opened";
            int toCheck0 = 0;
            int toCheck1 = 1;
            DateTime date = DateTime.Now;
            bool isActivate = true;

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            //Etre sur l'index des Claim et avoir des données disponibles
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();

            try
            {
                //Create claim closed
                var claimCreatePage = claimsPage.ClaimsCreatePage();
                claimCreatePage.FillField_CreatNewClaims(date, site, supplier, placeTo, isActivate);
                ClaimsItem claimsItem = claimCreatePage.Submit();
                var generalInformationPage = claimsItem.ClickOnGeneralInformation();
                itemNumClosed = generalInformationPage.GetClaimNumber();
                generalInformationPage.SetStatus(closedStatus);
                generalInformationPage.WaitPageLoading();
                generalInformationPage.BackToList();

                //create claim opened
                claimCreatePage = claimsPage.ClaimsCreatePage();
                claimCreatePage.FillField_CreatNewClaims(date, site, supplier, placeTo, isActivate);
                claimsItem = claimCreatePage.Submit();
                generalInformationPage = claimsItem.ClickOnGeneralInformation();
                itemNumOpened = generalInformationPage.GetClaimNumber();
                generalInformationPage.SetStatus(openedStatus);
                generalInformationPage.BackToList();

                //Filtrage Opened
                claimsPage.Filter(ClaimsPage.FilterType.Opened, true);
                claimsPage.Filter(ClaimsPage.FilterType.ByNumber, itemNumClosed);
                claimsPage.WaitPageLoading();
                var nbClosed = claimsPage.CheckTotalNumber();
                Assert.AreEqual(nbClosed, toCheck0, "Problème avec le filtre Opened");

                claimsPage.Filter(ClaimsPage.FilterType.Opened, true);
                claimsPage.Filter(ClaimsPage.FilterType.ByNumber, itemNumOpened);
                claimsPage.WaitPageLoading();
                var nbOpened = claimsPage.CheckTotalNumber();
                Assert.AreEqual(nbOpened, toCheck1, "Problème avec le filtre Opened");

            }
            finally
            {
                //delete
                claimsPage.Filter(ClaimsPage.FilterType.All, true);
                claimsPage.Filter(ClaimsPage.FilterType.ByNumber, itemNumClosed);
                claimsPage.WaitPageLoading();
                claimsPage.DeleteClaim(itemNumClosed);
                claimsPage.Filter(ClaimsPage.FilterType.ByNumber, itemNumOpened);
                claimsPage.WaitPageLoading();
                claimsPage.DeleteClaim(itemNumOpened);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_Filter_Sites_Places()
        {
            string place = TestContext.Properties["Place"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string sitePlace = site + "-" + place;
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string placeFrom = TestContext.Properties["PlaceFrom"].ToString();
            string numberOfCreatedClaim = String.Empty;
            DateTime date1 = DateUtils.Now;

            var homePage = LogInAsAdmin();

            //Act
            //Etre sur l'index et avoir des données disponible
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();

            try
            {
                // creé un claim non validé
                ClaimsCreateModalPage newClaim = claimsPage.ClaimsCreatePage();
                newClaim.FillField_CreatNewClaims(date1, site, supplier, placeFrom);
                numberOfCreatedClaim = newClaim.GetClaimId();
                ClaimsItem claimItem = newClaim.Submit();

                claimItem.BackToList();
                //Appliquer les filtres sur sites Places
                claimsPage.Filter(ClaimsPage.FilterType.ByNumber, numberOfCreatedClaim);
                claimsPage.Filter(ClaimsPage.FilterType.SitePlaces, sitePlace);

                //Vérifier que les résultats s'accordent bien au filtre appliqué
                ClaimsItem itemsPage = claimsPage.SelectFirstClaim();
                ClaimsGeneralInformation generalInfo = itemsPage.ClickOnGeneralInformation();
                string claimDeliveryLocation = generalInfo.GetClaimDeliveryLocation();
                string claimSite = generalInfo.GetClaimSite();
                Assert.AreEqual(place, claimDeliveryLocation, "Mauvaise place " + place);
                Assert.AreEqual(site, claimSite, "Mauvais site " + site);
                claimItem.BackToList();
            }
            finally
            {
                claimsPage.ResetFilter();
                claimsPage.Filter(ClaimsPage.FilterType.ByNumber, numberOfCreatedClaim);
                claimsPage.DeleteClaim(numberOfCreatedClaim);
            }

            
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_DetailsModifyGeneralInfo_Actived()
        {
            //prepare
            bool activate = false;
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();
            string numberOfCreatedClaim = String.Empty;
            DateTime date1 = DateUtils.Now;

            //arrange
            var homePage = LogInAsAdmin();

            //Act
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();
            try
            {
                // creé un claim non validé
                ClaimsCreateModalPage newClaim = claimsPage.ClaimsCreatePage();
                newClaim.FillField_CreatNewClaims(date1, site, supplier, placeTo);
                numberOfCreatedClaim = newClaim.GetClaimId();
                ClaimsItem claimItem = newClaim.Submit();
                claimItem.ClickOnGeneralInformation();
                claimItem.SetActivateFromGeneralInformation(activate);
                claimItem.WaitPageLoading();
                claimItem.ClickItemsSubMenu();
                claimItem.ClickGeneralInformationSubMenu();
                bool checkIsActivate = claimItem.GetActivateFromGeneralInformation();
                Assert.IsFalse(checkIsActivate, "claim n'a pas été modifié");
            }
            finally
            {
                // supprimer le claim
                claimsPage.BackToList();
                claimsPage.ResetFilter();
                if (!String.IsNullOrEmpty(numberOfCreatedClaim))
                {
                    claimsPage.Filter(ClaimsPage.FilterType.ShowNotValidated, true);
                    claimsPage.Filter(ClaimsPage.FilterType.ShowInactive, true);
                    claimsPage.Filter(ClaimsPage.FilterType.ByNumber, numberOfCreatedClaim);
                    claimsPage.DeleteClaim(numberOfCreatedClaim);
                }
            }

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_DetailsModifyGeneralInfo_Comment()
        {
            //prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();
            string numberOfCreatedClaim = String.Empty;
            DateTime date1 = DateUtils.Now;
            string comment = " this claim is not activated";

            //arrange
            var homePage = LogInAsAdmin();

            //Act
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();
            try
            {
                // creé un claim non validé
                ClaimsCreateModalPage newClaim = claimsPage.ClaimsCreatePage();
                newClaim.FillField_CreatNewClaims(date1, site, supplier, placeTo);
                numberOfCreatedClaim = newClaim.GetClaimId();
                ClaimsItem claimItem = newClaim.Submit();
                claimItem.ClickOnGeneralInformation();
                claimItem.SetCommentFromGeneralInformation(comment);
                claimItem.WaitPageLoading();
                claimItem.ClickItemsSubMenu();
                claimItem.ClickGeneralInformationSubMenu();
                string checkComment = claimItem.GetCommentFromGeneralInformation();
                Assert.AreEqual(checkComment,comment , "claim n'a pas été modifié");
            }
            finally
            {
                // supprimer le claim
                claimsPage.BackToList();
                claimsPage.ResetFilter();
                if (!String.IsNullOrEmpty(numberOfCreatedClaim))
                {
                    claimsPage.Filter(ClaimsPage.FilterType.ShowNotValidated, true);
                    claimsPage.Filter(ClaimsPage.FilterType.ShowActive, true);
                    claimsPage.Filter(ClaimsPage.FilterType.ByNumber, numberOfCreatedClaim);
                    claimsPage.DeleteClaim(numberOfCreatedClaim);
                }
            }
        }
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_DetailsModifyGeneralInfo_Status()
        {
            //prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();
            string numberOfCreatedClaim = String.Empty;
            DateTime date1 = DateUtils.Now;
            string choice2 = "2";
            string status = "Closed";
            //arrange
            var homePage = LogInAsAdmin();

            //Act
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();
            try
            {
                // creé un claim non validé
                ClaimsCreateModalPage newClaim = claimsPage.ClaimsCreatePage();
                newClaim.FillField_CreatNewClaims(date1, site, supplier, placeTo);
                numberOfCreatedClaim = newClaim.GetClaimId();
                ClaimsItem claimItem = newClaim.Submit();
                claimItem.ClickOnGeneralInformation();
                claimItem.SetStatusFromGeneralInformation(choice2);
                claimItem.WaitPageLoading();
                claimItem.ClickItemsSubMenu();
                claimItem.ClickGeneralInformationSubMenu();
                string statusCheck = claimItem.GetStatusFromGeneralInformation();
                Assert.AreEqual(statusCheck, status, "claim n'a pas été modifié");
            }
            finally
            {
                // supprimer le claim
                claimsPage.BackToList();
                claimsPage.ResetFilter();
                if (!String.IsNullOrEmpty(numberOfCreatedClaim))
                {
                    claimsPage.Filter(ClaimsPage.FilterType.ShowNotValidated, true);
                    claimsPage.Filter(ClaimsPage.FilterType.ShowActive, true);
                    claimsPage.Filter(ClaimsPage.FilterType.ByNumber, numberOfCreatedClaim);
                    claimsPage.DeleteClaim(numberOfCreatedClaim);
                }
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_DetailsVerifyFooter()
        {
            var sanctionAmount = "55.5";
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var decimalSeparator = homePage.GetDecimalSeparatorValue();
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();
            claimsPage.Filter(ClaimsPage.FilterType.ShowNotValidated, true);
            claimsPage.ClickFirstClaim();
            var culture = decimalSeparator == "," ? CultureInfo.CreateSpecificCulture("fr-FR") : CultureInfo.CreateSpecificCulture("en-US");
            var totalSanctionsItems = claimsPage.GetTotalSactionsFromItems();
            decimal.TryParse(totalSanctionsItems, NumberStyles.AllowCurrencySymbol | NumberStyles.Number, culture, out decimal result);

            if (result == 0 || String.IsNullOrEmpty(totalSanctionsItems))
            {
                // add sanction to the claim
                claimsPage.ClickFirstItem();
                claimsPage.AddSanctionToItem(sanctionAmount);
            }

            string totalClaimItems = claimsPage.GetTotalClaimFromItems();
            string newtotalSanctionsItems = claimsPage.GetTotalSactionsFromItems();
            claimsPage.ClickFooterSubMenu();
            string totalGrossAmountFooter = claimsPage.GetTotalGrossAmountFromFooter();
            string totalSanctionsFooter = claimsPage.GetSanctionsFromFooter();
            string vatSanctionsFooter = claimsPage.GetSanctionVATFooter();
            // parse string values to decimal

            decimal.TryParse(totalClaimItems, NumberStyles.AllowCurrencySymbol | NumberStyles.Number, culture, out decimal totalClaimsItemsDecimal);
            decimal.TryParse(newtotalSanctionsItems, NumberStyles.AllowCurrencySymbol | NumberStyles.Number, culture, out decimal totalSanctionsItemsDecimal);
            decimal.TryParse(totalGrossAmountFooter, NumberStyles.AllowCurrencySymbol | NumberStyles.Number, culture, out decimal totalGrossAmountFooterDecimal);
            decimal.TryParse(totalSanctionsFooter, NumberStyles.AllowCurrencySymbol | NumberStyles.Number, culture, out decimal totalSanctionsFooterDecimal);
            decimal.TryParse(vatSanctionsFooter, NumberStyles.AllowCurrencySymbol | NumberStyles.Number, culture, out decimal varSanctionsFooterDecimal);
            var sanctionsTotalFooter = totalSanctionsFooterDecimal + varSanctionsFooterDecimal;
            //compare sanction item to sanction footer and claim items to claim footer values
            Assert.AreEqual(totalClaimsItemsDecimal, totalGrossAmountFooterDecimal, "les données ne sont pas correctes (cas 1)");
            Assert.AreEqual(totalSanctionsItemsDecimal, sanctionsTotalFooter, "les données ne sont pas correctes (cas 2)");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_CreateClaimFromSupplierInvoice()
        {
            // Prepare
            Random rnd = new Random();
            string supplierInvoiceNb = rnd.Next(100000000, 999999999).ToString();
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string itemName = TestContext.Properties["Item_ReceiptNote"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();
            string currency = TestContext.Properties["Currency"].ToString();

            // Log in
            var homePage = LogInAsAdmin();

            var decimalSeparator = homePage.GetDecimalSeparatorValue();

            // Act
            SupplierInvoicesPage supplierInvoicesPage = homePage.GoToAccounting_SupplierInvoices();
            //Créer un Supplier Invoice, ajouter une quantité sur un item avec SI qty, SI price et Tax et valider
            //(cf AC_SI_CreateSupplierInvoice)
            SupplierInvoicesCreateModalPage modalSI = supplierInvoicesPage.SupplierInvoicesCreatePage();
            modalSI.FillField_CreateNewSupplierInvoices(supplierInvoiceNb, DateUtils.Now, site, supplier);
            SupplierInvoicesItem itemsSI = modalSI.Submit();
            itemsSI.AddNewItem(itemName, "45");
            double totalPriceSI = itemsSI.GetInvoiceTotalPrice(currency, decimalSeparator);
            itemsSI.ValidateSupplierInvoice();

            //1. Cliquer sur + puis "New claim'
            ClaimsPage claimPage = itemsSI.GoToWarehouse_ClaimsPage();
            ClaimsCreateModalPage claimModal = claimPage.ClaimsCreatePage();
            //2. Remplir le formulaire avec le même site, supplier que la SI et cohcer "Created from SI"
            claimModal.FillField_CreatNewClaims(DateUtils.Now, site, supplier, placeTo, true, "Supplier Invoice", supplierInvoiceNb);
            ClaimsItem claimItems = claimModal.Submit();
            double qtySI = int.Parse(claimItems.GetQuantity().Replace(" ", ""));
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");
            double priceSI = double.Parse(claimItems.GetPrice().Replace("€", "").Replace(" ", ""), ci);
            ClaimsGeneralInformation generalInfo = claimItems.ClickOnGeneralInformation();
            string claimSupplierInvoice = generalInfo.GetClaimSupplierInvoice();
            Assert.AreEqual(supplierInvoiceNb, claimSupplierInvoice);
            //3. Sélectionner la SI précédemment créée / Vérifier le Total Amount correspondant à la SI sur la popup de création
            Assert.AreEqual(totalPriceSI, qtySI * priceSI, 0.0001, "Le total Amount SI ne correspond pas au total Amount de Claim");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_DetailsModifyChecks()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            //Créer une claim
            ClaimsPage claimPage = homePage.GoToWarehouse_ClaimsPage();
            ClaimsCreateModalPage claimModal = claimPage.ClaimsCreatePage();
            claimModal.FillField_CreatNewClaims(DateUtils.Now, site, supplier, placeTo, true);
            ClaimsItem claimItems = claimModal.Submit();

            //1.Aller sur l'onglet checks
            ClaimsChecks claimCkecks = claimItems.ClickOnChecks();
            //2.Vérifier que les checks sont à Yes
            claimCkecks.CheckAllYes();
            //3.Modifier les checks à No
            claimCkecks.SetAllNo();
            //4.Changer d'onglet et revenir sur l'onglet Checks
            claimItems = claimCkecks.ClickOnItems();
            claimItems.ClickOnChecks();
            //5. Vérifier que les Checks ont bien modifié
            claimCkecks.CheckAllNo();
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_CreateClaimFromRNKO()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();
            string receiptNoteNumber = "123456";
            // Log in
            var homePage = LogInAsAdmin();

            // Act
            //Créer une claim
            ClaimsPage claimPage = homePage.GoToWarehouse_ClaimsPage();
            ClaimsCreateModalPage claimModal = claimPage.ClaimsCreatePage();
            claimModal.FillField_CreatNewClaims(DateUtils.Now, site, supplier, placeTo, true, "Receipt Note", null);
            // numero bidon pour réduire la fenêtres
            claimModal.FillFieldReceiptNoteSearch(receiptNoteNumber);
            claimModal.Submit();

            Assert.AreEqual("You must select at least one item to copy.", claimModal.ValidationError());
        }
        
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_DetailChangeLine()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();
            string qtyReceived = "1112";
            string ID = string.Empty;
            // Log in
            var homePage = LogInAsAdmin();

            // Act
            //Créer une claim
            ClaimsPage claimPage = homePage.GoToWarehouse_ClaimsPage();
            claimPage.ResetFilter();
            try
            {
                ClaimsCreateModalPage claimModal = claimPage.ClaimsCreatePage();
                claimModal.FillField_CreatNewClaims(DateUtils.Now, site, supplier, placeTo, false, null, null);
                ID = claimModal.GetClaimId();
                ClaimsItem claimsItem = claimModal.Submit();
                bool verifDetailChangeLineIsTrue = claimsItem.VerifyDetailChangeLine(qtyReceived);
                Assert.IsTrue(verifDetailChangeLineIsTrue, String.Format(MessageErreur.FILTRE_ERRONE, "ChangeLine ne fonctionne pas."));
                claimsItem.BackToList();
            }
            finally
            {
                claimPage.ResetFilter();
                claimPage.Filter(ClaimsPage.FilterType.ByNumber, ID);
                claimPage.Filter(ClaimsPage.FilterType.ShowAll, true);
                claimPage.DeleteClaim(ID);
            }
            
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_PrintDetailBackdate()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string delivery = TestContext.Properties["PlaceFrom"].ToString();

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();

            // Create
            var claimsCreateModalpage = claimsPage.ClaimsCreatePage();
            claimsCreateModalpage.FillField_CreatNewClaims(DateUtils.Now, site, supplier, delivery, true);
            String ID = claimsCreateModalpage.GetClaimId();
            var claimsItem = claimsCreateModalpage.Submit();

            // Update the first item value to activate the activation menu
            var itemName = claimsItem.GetFirstItemName();
            claimsItem.SelectFirstItem();
            claimsItem.AddQuantityAndType(itemName, 2);


            claimsItem.Validate();
            PrintReportPage printClaimItems = claimsItem.PrintClaimItems(true);
            var isReportGenerated = printClaimItems.IsReportGenerated();
            printClaimItems.Close();

            Assert.IsTrue(isReportGenerated, "L'impression ne doit pas etre possible");

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_PrintIndexBackdate()
        {
            //Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string delivery = TestContext.Properties["PlaceFrom"].ToString();

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();

            // Create
            var claimsCreateModalpage = claimsPage.ClaimsCreatePage();
            claimsCreateModalpage.FillField_CreatNewClaims(DateUtils.Now, site, supplier, delivery, true);
            String ID = claimsCreateModalpage.GetClaimId();
            var claimsItem = claimsCreateModalpage.Submit();

            // Update the first item value to activate the activation menu
            var itemName = claimsItem.GetFirstItemName();
            claimsItem.SelectFirstItem();
            claimsItem.AddQuantityAndType(itemName, 2);



            claimsItem.Validate();

            claimsPage = claimsItem.BackToList();
            claimsPage.ResetFilter();
            claimsPage.Filter(ClaimsPage.FilterType.ByNumber, ID);
            claimsPage.Filter(ClaimsPage.FilterType.DateFrom, DateUtils.Now);
            claimsPage.Filter(ClaimsPage.FilterType.DateTo, DateUtils.Now.AddMonths(1));

            PrintReportPage printReportPage = claimsItem.PrintClaimResults(true);

            var isReportGenerated = printReportPage.IsReportGenerated();
            printReportPage.Close();

            Assert.IsTrue(isReportGenerated, "L'impression ne doit pas etre possible");

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_FooterCurrencies()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();

            // Log in
            var homePage = LogInAsAdmin();

            //Act
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();
            claimsPage.Filter(FilterType.Site, site);
            claimsPage.PageUp();
            claimsPage.Filter(FilterType.Suppliers, supplier);
            claimsPage.ClickFirstClaim();
            claimsPage.ClickFooterSubMenu();
            string localCurrency = claimsPage.GetLocalCurrency();
            string supplierCurrency = claimsPage.GetSupplierCurrency();
            homePage.Navigate();
            var settingsPage = homePage.GoToApplicationSettings();
            var parametreSitesPage = settingsPage.GoToParameters_Sites();
            parametreSitesPage.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, site);
            parametreSitesPage.ClickOnFirstSite();
            string sitePageCurrency = parametreSitesPage.GetCurrency();
            Assert.IsTrue(localCurrency.Contains(sitePageCurrency) || sitePageCurrency.Contains(localCurrency), "local currency et site currency ne sont pas le même");

            homePage.Navigate();
            var supplierPage = homePage.GoToPurchasing_SuppliersPage();
            supplierPage.Filter(SuppliersPage.FilterType.Search, supplier);
            var supplierItem = supplierPage.ClickAndGoFirstSupplier();
            string supplierPageCurrency = supplierItem.GetCurrency();

            Assert.IsTrue(supplierCurrency.Contains(supplierPageCurrency) || supplierPageCurrency.Contains(supplierCurrency), "supplier currency et purchasing supplier currency ne sont pas le même");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_Refresh()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();
            string ID = string.Empty;
            // Log in
            var homePage = LogInAsAdmin();

            //Act
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();
            try
            {
                var createClaimPage = claimsPage.ClaimsCreatePage();
                createClaimPage.FillField_CreatNewClaims(DateUtils.Now, site, supplier, placeTo, true);
                ID = createClaimPage.GetClaimId();
                var claimsItem = createClaimPage.Submit();

                //claimsPage.PageUp();
                var currentRecievedQuantity = claimsPage.GetRecievedQuantity().Replace(" ", "");
                var currentDecreasedQuantity = claimsPage.GetDecreasedQuantity().Replace(" ", "");

                var currentRecievedQuantityInt = currentRecievedQuantity == null || currentRecievedQuantity == string.Empty ? 0 : double.Parse(currentRecievedQuantity) + 5;
                var currentDecreasedQuantityInt = currentDecreasedQuantity == null || currentDecreasedQuantity == string.Empty ? 0 : double.Parse(currentDecreasedQuantity) + 5;
                claimsPage.ClickFirstItem();

                
                claimsPage.SetRecievedQuantity(currentRecievedQuantityInt.ToString());
                claimsPage.SetDecreasedQuantity(currentDecreasedQuantityInt.ToString());
                claimsPage.Refresh();
                claimsPage.BackToList();
                claimsPage.Filter(FilterType.ByNumber, ID);
                claimsPage.ClickFirstClaim();
                var newRecievedQuantity = claimsPage.GetRecievedQuantity().Replace(" ", "");
                var newDecreasedQuantity = claimsPage.GetDecreasedQuantity().Replace(" ", "");
                Assert.AreNotEqual(currentRecievedQuantity, newRecievedQuantity, "La quantité Recieved n'est pas modifiée");
                Assert.AreNotEqual(currentDecreasedQuantity, newDecreasedQuantity, "La quantité Decreased  n'est pas modifiéé");
            }
            finally
            {
                // Filter and Delete Claim Created
                claimsPage.BackToList();
                claimsPage.ResetFilter();
                claimsPage.Filter(ClaimsPage.FilterType.ShowNotValidated, true);
                claimsPage.Filter(ClaimsPage.FilterType.ByNumber, ID);
                claimsPage.DeleteClaim(ID);
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_GenerateSI()
        {
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();
            string invoiceNumber = new Random().Next().ToString();

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();
            var createClaimPage = claimsPage.ClaimsCreatePage();
            createClaimPage.FillField_CreatNewClaims(DateUtils.Now, site, supplier, placeTo, true);
            string ID = createClaimPage.GetClaimId();
            var claimsItem = createClaimPage.Submit();
            claimsItem.SelectFirstItem();
            claimsItem.ClickClaim();
            claimsItem.FillClaim("Test Comment");

            claimsItem.Validate();

            var supplierInvoicesPage = claimsPage.GenerateSI(invoiceNumber);

            supplierInvoicesPage.ClickGeneralInformationTab();
            var supplierInvoicesNumber = supplierInvoicesPage.GetInvoiceNumber();
            var supplierInvoicesSite = supplierInvoicesPage.GetInvoiceSite();
            var supplierInvoicesSupplier = supplierInvoicesPage.GetInvoiceSupplier();

            Assert.IsTrue(site.Equals(supplierInvoicesSite) && supplier.Equals(supplierInvoicesSupplier) && invoiceNumber.Equals(supplierInvoicesNumber), "les données ne sont pas le même.");

        }

        /// <summary>
        /// Update claim's first item allergens and check the state 
        /// of the allergens button in the claim details.
        /// </summary>
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_CheckAllergensUpdate()
        {
            string allergen1 = "Cacahuetes/Peanuts";
            string allergen2 = "Frutos de cáscara- Macadamias/Nuts-Macadamia";
            HomePage homePage = LogInAsAdmin();
            ClaimsPage claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();
            claimsPage.Filter(ClaimsPage.FilterType.ShowNotValidated, true);
            string claimNumber = claimsPage.GetFirstClaimNumber();
            claimsPage.Filter(ClaimsPage.FilterType.ByNumber, claimNumber);
            ClaimsItem claimsItem = claimsPage.SelectFirstClaim();
            string itemName = claimsItem.GetFirstItemName();
            claimsItem.ResetFilters();
            claimsItem.Filter(ClaimsItem.FilterItemType.SearchByName, itemName);
            claimsItem.SelectFirstItem();
            claimsItem.SetReceviedQuantity("10");
            claimsItem.Refresh();
            claimsItem.ResetFilters();
            claimsItem.Filter(ClaimsItem.FilterItemType.SearchByName, itemName);
            claimsItem.SelectFirstItem();
            ItemGeneralInformationPage itemPage = claimsItem.EditItemGeneralInformation(itemName);
            ItemIntolerancePage itemIntolerancePage = itemPage.ClickOnIntolerancePage();
            itemIntolerancePage.AddAllergen(allergen1);
            itemIntolerancePage.AddAllergen(allergen2);
            claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();
            claimsPage.Filter(ClaimsPage.FilterType.ShowNotValidated, true);
            claimsPage.Filter(ClaimsPage.FilterType.ByNumber, claimNumber);
            claimsItem = claimsPage.SelectFirstClaim();
            claimsItem.ResetFilters();
            claimsItem.Filter(ClaimsItem.FilterItemType.SearchByName, itemName);
            bool isIconGreen = claimsItem.IsAllergenIconGreen(itemName);
            List<string> allergensInItem = claimsItem.GetAllergens(itemName);
            bool containsAllergen1 = allergensInItem.Contains(allergen1);
            bool containsAllergen2 = allergensInItem.Contains(allergen2);
            Assert.IsTrue(isIconGreen, "L'icon n'est pas vert!");
            Assert.IsTrue(containsAllergen1 && containsAllergen2, "Allergens n'ont pas été ajoutés");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_Send_Claim()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string supplier = TestContext.Properties["ClaimsSupplier"].ToString();
            string placeTo = TestContext.Properties["PlaceTo"].ToString();
            string mail = TestContext.Properties["Admin_UserName"].ToString();
            string ID = string.Empty;
            ClaimsItem claimsItem = null;
            // Log in
            var homePage = LogInAsAdmin();

            //Act
            var claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();
            claimsPage.Filter(FilterType.ShowValidatedOnly, true);
            claimsPage.Filter(FilterType.Site, site);
            if (claimsPage.CheckTotalNumber() == 0)
            {

                var createClaimPage = claimsPage.ClaimsCreatePage();
                createClaimPage.FillField_CreatNewClaims(DateUtils.Now, site, supplier, placeTo, true);
                ID = createClaimPage.GetClaimId();
                claimsItem = createClaimPage.Submit();

                claimsItem.SelectFirstItem();
                claimsItem.ClickClaim();
                claimsItem.FillClaim("Test Comment");
                claimsItem.Validate();

            }
            else
            {
                ID = claimsPage.GetFirstClaimNumber();
            }

            var settingsPage = homePage.GoToApplicationSettings();
            var parametreSitesPage = settingsPage.GoToParameters_Sites();
            parametreSitesPage.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, site);
            parametreSitesPage.ClickOnFirstSite();
            parametreSitesPage.WaitPageLoading();
            parametreSitesPage.ClickToContacts();

            string AdressManager = parametreSitesPage.GetClaimManagerMail();

            homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();
            claimsPage.Filter(FilterType.ByNumber, ID);

            claimsItem = claimsPage.SelectFirstClaim();
            bool isMailSendOK = claimsItem.VerifyAdressAndSendByMail(mail, AdressManager);

            //Mail assert
            Assert.IsTrue(isMailSendOK, "L'email n'est pas été envoyée");

        }
        [TestMethod]
        [Timeout(_timeout)]       
        public void WA_CLAI_Items_VerifyPackPrice()
        {
            HomePage homePage = LogInAsAdmin();
            string decimalSeparator = homePage.GetDecimalSeparatorValue();
            ClaimsPage claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();
            claimsPage.Filter(ClaimsPage.FilterType.ShowNotValidated, true);
            ClaimsItem claimItems = claimsPage.SelectFirstClaim();
            string itemName = claimItems.GetFirstItemName();
            claimItems.SelectFirstItem();
            claimItems.AddQuantityAndType(itemName, 10);
            claimItems.AddPrice(itemName,10);
            claimItems.Refresh();
            double price = ArrangeTarif(claimItems.GetPrice(), decimalSeparator);
            double quantity = ArrangeTarif(claimItems.GetQuantity(), decimalSeparator);
            double  total = ArrangeTarif(claimItems.GetClaimTotalAmount(), decimalSeparator);
            double priceDN = ArrangeTarif(claimItems.GetDNPrice(), decimalSeparator);
            double quantityDN = ArrangeTarif(claimItems.GetDNQuantity(), decimalSeparator);
            Assert.AreEqual(10, price, 0.0001, "Mauvais price update");
            Assert.AreEqual((priceDN * quantityDN) - (price * quantity), total, 0.0001, "Mauvais total claim");
            claimItems.SelectFirstItem();
            ClaimEditClaimForm editClaim = claimItems.EditClaimForm(itemName);
            double amount = ArrangeTarif(editClaim.WaitForElementIsVisible(By.Id("spanClaimAmount")).Text, decimalSeparator);
            Assert.AreEqual(total, amount, "Mauvais claim amount");
        }
       
        [TestMethod]
        [Timeout(_timeout)]
        public void WA_CLAI_Items_VerifyReceivedQty()
        {
            //Arrange
            HomePage homePage = LogInAsAdmin();
            //Act
            string decimalSeparator = homePage.GetDecimalSeparatorValue();
            ClaimsPage claimsPage = homePage.GoToWarehouse_ClaimsPage();
            claimsPage.ResetFilter();
            claimsPage.Filter(ClaimsPage.FilterType.ShowNotValidated, true);
            ClaimsItem claimItems = claimsPage.SelectFirstClaim();
            string itemName = claimItems.GetFirstItemName();
            claimItems.SelectFirstItem();
            claimItems.AddQuantityAndType(itemName, 3);
            claimItems.Refresh();
            double price = ArrangeTarif(claimItems.GetPrice(), decimalSeparator);
            double quantity = ArrangeTarif(claimItems.GetQuantity(), decimalSeparator);
            double total = ArrangeTarif(claimItems.GetClaimTotalAmount(), decimalSeparator);
            double priceDN = ArrangeTarif(claimItems.GetDNPrice(), decimalSeparator);
            double quantityDN = ArrangeTarif(claimItems.GetDNQuantity(), decimalSeparator);
            //Assert
            Assert.AreEqual(3, quantity, 0.0001, "Mauvais quantity update");
            Assert.AreEqual((priceDN * quantityDN) - (price * quantity), total, 0.0001, "Mauvais total claim");
            claimItems.SelectFirstItem();
            ClaimEditClaimForm editClaim = claimItems.EditClaimForm(itemName);
            double amount = ArrangeTarif(editClaim.WaitForElementIsVisible(By.Id("spanClaimAmount")).Text, decimalSeparator);
            Assert.AreEqual(total, amount, "Mauvais claim amount");
        }
    }
}