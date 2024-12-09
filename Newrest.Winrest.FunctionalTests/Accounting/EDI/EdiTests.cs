using Limilabs.Mail.Fluent;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Accounting.Edi;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Sites;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.CustomerOrder;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Item;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Suppliers;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.Claims;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Policy;

namespace Newrest.Winrest.FunctionalTests.Accounting
{
    [TestClass]
    public class EdiTests : TestBase
    {
        private const int _timeout = 600000;

        [TestMethod]
        [Priority(0)]
        [Timeout(_timeout)]
        public void AC_EDI_Export_Ci()
        {
            var timeStringCode = DateTime.Now.ToString("ddHHmmssff");

            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            bool newVersionPrint = true;
            HomePage homePage= LogInAsAdmin();

            //create data to export
            //create customer chorus
            var customerPage = homePage.GoToCustomers_CustomerPage();
            customerPage.CustomerCreatePage();
            customerPage.CreateCustomerChorus(timeStringCode);
            customerPage.WaitLoading();

            // create service with that customer
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ServiceCreatePage();
            servicePage.CreateNewService(timeStringCode);
            //create invoice
            var invoicePage = homePage.GoToAccounting_InvoicesPage();
            invoicePage.ManualInvoiceCreatePage();
            invoicePage.CreateManualInvoiceWithCustomer(timeStringCode);
            invoicePage.SendToEdi();

            //export data
            homePage.GoToWinrestHome();
            var ediPage = homePage.GoToAccounting_EdiPage();
            var ci = ediPage.ClickOnCustomerInvoicesTab();
            ci.ResetFilters();
            ediPage.ClearDownloads();
            DeleteAllFileDownload();
            ci.ExportExcelCI(newVersionPrint);
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            var correctDownloadedFile = ci.GetExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

        }

        [TestMethod]
        [Priority(1)]
        [Timeout(_timeout)]
        public void AC_EDI_Export_Po()
        {
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string codeEdi = new Random().Next().ToString();
            string idEdi = new Random().Next(1000, 5000).ToString();
            string quantity = new Random().Next(1, 9).ToString();
            string baseSupplierName = "SupplierForEdiPo";
            DateTime now = DateTime.Now;
            string formattedDateTime = now.ToString("yyyyMMdd_HHmmss");
            string supplier = $"{baseSupplierName}_{formattedDateTime}";
            string site = TestContext.Properties["SiteACE"].ToString();
            string location = "Produccion";
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();
            LogInAsAdmin();
            HomePage homePage = new HomePage(WebDriver, TestContext);
            DeleteAllFileDownload();
            EdiPage ediPage = homePage.GoToAccounting_EdiPage();
            ediPage.ClearDownloads();
            EdiPurchaseOrdersPage ediPurchaseOrdersPage = ediPage.ClickOnPurchaseOrdersTab();
            int PurchaseOrderBefore = ediPurchaseOrdersPage.CheckTotalNumber();
            ParametersSites parametersSites = homePage.GoToParameters_Sites();
            parametersSites.ClickOnFirstSite();
            string ediCode = parametersSites.GetEdiGNLCode();
            if (ediCode == "") parametersSites.SetEdiGNLCode(codeEdi);
            SuppliersPage suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            /* creation */
            SupplierCreateModalPage supplierCreateModalPage = suppliersPage.SupplierCreatePage();
            supplierCreateModalPage.FillField_CreatNewSupplier(supplier, true);
            SupplierGeneralInfoTab supplierInfo = supplierCreateModalPage.SubmitToGeneralInfo();
            SupplierDeliveriesTab supplierDeliveriesTab = supplierInfo.GoToSupplierDeliveries();
            supplierDeliveriesTab.SetAmountDeliveredSites(site, "5");
            supplierDeliveriesTab.SetShippingCost(site, "5");
            supplierDeliveriesTab.SetPrepaDelay(site, "5");
            supplierDeliveriesTab.SetDeliveryWeekDays();
            SupplierEDITab supplierEDITab = supplierDeliveriesTab.GoToEdiTab();
            supplierEDITab.SetIdEDI(idEdi);
            supplierEDITab.SetPurchaseOrderFileFormatToXML();
            supplierEDITab.SelectASiteInOrderEDI(site);
            if (ediCode == "") supplierEDITab.SetEdiGNLCode(codeEdi);
            else supplierEDITab.SetEdiGNLCode(ediCode);
            ItemPage itemPage = supplierEDITab.GoToPurchasing_ItemPage();
            ItemGeneralInformationPage itemGeneralInformationPage = itemPage.ClickOnFirstItem();
            //itemGeneralInformationPage.SetActivated(true);
            ItemCreateNewPackagingModalPage itemCreateNewPackagingModalPage = itemGeneralInformationPage.NewPackaging();
            itemCreateNewPackagingModalPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);
            PurchaseOrdersPage purchaseOrdersPage = homePage.GoToPurchasing_PurchaseOrdersPage();
            CreatePurchaseOrderModalPage createPurchaseOrderModalPage = purchaseOrdersPage.CreateNewPurchaseOrder();
            createPurchaseOrderModalPage.FillPrincipalField_CreationNewPurchaseOrder(site, supplier, location, DateTime.Today, true);
            PurchaseOrderItem purchaseOrderItem = createPurchaseOrderModalPage.Submit();
            purchaseOrderItem.SelectFirstItemPo();
            purchaseOrderItem.AddQuantity(quantity);
            purchaseOrderItem.Validate();
            purchaseOrderItem.SendToEdi();
            purchaseOrderItem.GoToWinrestHome();
            ediPage = homePage.GoToAccounting_EdiPage();
            ediPurchaseOrdersPage = ediPage.ClickOnPurchaseOrdersTab();
            int PurchaseOrderAfter = ediPurchaseOrdersPage.CheckTotalNumber();
            bool isGreater = PurchaseOrderAfter > PurchaseOrderBefore;
            Assert.IsTrue(isGreater, "Le Accounting (Purchase Order) n'est pas créé");
            ediPurchaseOrdersPage.ResetFilters();
            ediPurchaseOrdersPage.ExportExcelPO(true);
            //en comment //pas de données, pas de print, en attendant d'implémenter données
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();
            FileInfo correctDownloadedFile = ediPurchaseOrdersPage.GetExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_EDI_ResetFilter()
        {
            //Prepare

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            EdiPage edi = homePage.GoToAccounting_EdiPage();
            EdiSupplierInvoicesPage ediSupplierInvoicesPage = edi.ClickOnSupplierAccountingsTab();
            object value = null;

            ediSupplierInvoicesPage.ResetFilters();
            // remplir les valeurs
            ediSupplierInvoicesPage.Filter(EdiSupplierInvoicesPage.FilterType.Search, "test");
            value = ediSupplierInvoicesPage.GetFilterValue(EdiSupplierInvoicesPage.FilterType.Search);
            Assert.AreEqual("test", value, "ResetFilter Search 'test'");
            ediSupplierInvoicesPage.Filter(EdiSupplierInvoicesPage.FilterType.Supplier, "AIR CPU, S.L.");
            value = ediSupplierInvoicesPage.GetFilterValue(EdiSupplierInvoicesPage.FilterType.Supplier);
            Assert.AreEqual("AIR CPU, S.L.", value, "ResetFilter Supplier 'AIR CPU, S.L.'");
            ediSupplierInvoicesPage.Filter(EdiSupplierInvoicesPage.FilterType.From, DateUtils.Now.AddMonths(-2));
            value = ediSupplierInvoicesPage.GetFilterValue(EdiSupplierInvoicesPage.FilterType.From);
            Assert.AreEqual(DateUtils.Now.AddMonths(-2).ToString("dd/MM/yyyy"), value, "ResetFilter From '" + DateUtils.Now.AddMonths(-2).ToString("dd/MM/yyyy") + "'");
            ediSupplierInvoicesPage.Filter(EdiSupplierInvoicesPage.FilterType.To, DateUtils.Now.AddMonths(1));
            value = ediSupplierInvoicesPage.GetFilterValue(EdiSupplierInvoicesPage.FilterType.To);
            Assert.AreEqual(DateUtils.Now.AddMonths(1).ToString("dd/MM/yyyy"), value, "ResetFilter To '" + DateUtils.Now.AddMonths(1).ToString("dd/MM/yyyy") + "'");
            ediSupplierInvoicesPage.Filter(EdiSupplierInvoicesPage.FilterType.Sites, "MAD - MAD");
            value = ediSupplierInvoicesPage.GetFilterValue(EdiSupplierInvoicesPage.FilterType.Sites);
            Assert.IsTrue(((string)value).StartsWith("1 of "), "ResetFilter Sites 'MAD - MAD'");
            ediSupplierInvoicesPage.Filter(EdiSupplierInvoicesPage.FilterType.Status, "SI created and validated");
            value = ediSupplierInvoicesPage.GetFilterValue(EdiSupplierInvoicesPage.FilterType.Status);
            Assert.IsTrue(((string)value).StartsWith("1 of "), "ResetFilter Status 'SI created and validated'");
            value = ediSupplierInvoicesPage.GetFilterValue(EdiSupplierInvoicesPage.FilterType.ShowOnlyCN);
            Assert.AreEqual(true, value, "ResetFilter ShowOnlyCN");

            ediSupplierInvoicesPage.ResetFilters();
            // valeurs après reset
            value = ediSupplierInvoicesPage.GetFilterValue(EdiSupplierInvoicesPage.FilterType.Search);
            Assert.AreEqual("", value, "ResetFilter Search ''");
            value = ediSupplierInvoicesPage.GetFilterValue(EdiSupplierInvoicesPage.FilterType.Supplier);
            Assert.AreEqual("ALL", value, "ResetFilter Supplier ''");
            value = ediSupplierInvoicesPage.GetFilterValue(EdiSupplierInvoicesPage.FilterType.From);
            Assert.AreEqual(DateUtils.Now.AddMonths(-1).ToString("dd/MM/yyyy"), value, "ResetFilter From '" + DateUtils.Now.AddMonths(-1).ToString("dd/MM/yyyy") + "'");
            value = ediSupplierInvoicesPage.GetFilterValue(EdiSupplierInvoicesPage.FilterType.To);
            Assert.AreEqual(DateUtils.Now.ToString("dd/MM/yyyy"), value, "ResetFilter To '" + DateUtils.Now.ToString("dd/MM/yyyy") + "'");
            value = ediSupplierInvoicesPage.GetFilterValue(EdiSupplierInvoicesPage.FilterType.Sites);
            Assert.IsFalse(((string)value).StartsWith("1 of "), "ResetFilter Sites 'all'");
            value = ediSupplierInvoicesPage.GetFilterValue(EdiSupplierInvoicesPage.FilterType.Status);
            Assert.IsFalse(((string)value).StartsWith("1 of "), "ResetFilter Status 'all'");
            value = ediSupplierInvoicesPage.GetFilterValue(EdiSupplierInvoicesPage.FilterType.ShowAll);
            Assert.AreEqual(true, value, "ResetFilter ShowAll");



            // goodies

            EdiSupplierInvoicesPage si = edi.ClickOnSupplierAccountingsTab();
            si.ResetFilters();
            EdiPurchaseOrdersPage po = edi.ClickOnPurchaseOrdersTab();
            po.ResetFilters();
            EdiCustomerInvoicesPage ci = edi.ClickOnCustomerInvoicesTab();
            ci.ResetFilters();
            EdiCustomerOrderPage co = edi.ClickOnCustomerOrderTab();
            co.ResetFilters();
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_EDI_ResetFilter_Po()
        {
            //Prepare
            object value = null;
           
            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            EdiPage edi = homePage.GoToAccounting_EdiPage();
            EdiPurchaseOrdersPage po = edi.ClickOnPurchaseOrdersTab();
            po.ResetFilters();
            // remplir les valeurs
            po.Filter(EdiPurchaseOrdersPage.FilterEdi.Search, "test");
            value = po.GetFilterValue(EdiPurchaseOrdersPage.FilterEdi.Search);
            Assert.AreEqual("test", value, "ResetFilter Search 'test'");
            po.Filter(EdiPurchaseOrdersPage.FilterEdi.Supplier, "AIR CPU, S.L.");
            value = po.GetFilterValue(EdiPurchaseOrdersPage.FilterEdi.Supplier);
            Assert.AreEqual("AIR CPU, S.L.", value, "ResetFilter Supplier 'AIR CPU, S.L.'");
            po.Filter(EdiPurchaseOrdersPage.FilterEdi.From, DateUtils.Now.AddMonths(-2));
            value = po.GetFilterValue(EdiPurchaseOrdersPage.FilterEdi.From);
            Assert.AreEqual(DateUtils.Now.AddMonths(-2).ToString("dd/MM/yyyy"), value, "ResetFilter From '" + DateUtils.Now.AddMonths(-2).ToString("dd/MM/yyyy") + "'");
            po.Filter(EdiPurchaseOrdersPage.FilterEdi.To, DateUtils.Now.AddMonths(1));
            value = po.GetFilterValue(EdiPurchaseOrdersPage.FilterEdi.To);
            Assert.AreEqual(DateUtils.Now.AddMonths(1).ToString("dd/MM/yyyy"), value, "ResetFilter To '" + DateUtils.Now.AddMonths(1).ToString("dd/MM/yyyy") + "'");
            po.Filter(EdiPurchaseOrdersPage.FilterEdi.Sites, "MAD - MAD");
            value = po.GetFilterValue(EdiPurchaseOrdersPage.FilterEdi.Sites);
            Assert.IsTrue(((string)value).StartsWith("1 of "), "ResetFilter Sites 'MAD - MAD'");
            po.Filter(EdiPurchaseOrdersPage.FilterEdi.Status, "Received");
            value = po.GetFilterValue(EdiPurchaseOrdersPage.FilterEdi.Status);

            po.ResetFilters();
            // valeurs après reset
            value = po.GetFilterValue(EdiPurchaseOrdersPage.FilterEdi.Search);
            Assert.AreEqual("", value, "ResetFilter Search ''");
            value = po.GetFilterValue(EdiPurchaseOrdersPage.FilterEdi.Supplier);
            Assert.AreEqual("ALL", value, "ResetFilter Supplier ''");
            value = po.GetFilterValue(EdiPurchaseOrdersPage.FilterEdi.From);
            Assert.AreEqual(DateUtils.Now.AddMonths(-1).ToString("dd/MM/yyyy"), value, "ResetFilter From '" + DateUtils.Now.AddMonths(-1).ToString("dd/MM/yyyy") + "'");
            value = po.GetFilterValue(EdiPurchaseOrdersPage.FilterEdi.To);
            Assert.AreEqual(DateUtils.Now.ToString("dd/MM/yyyy"), value, "ResetFilter To '" + DateUtils.Now.ToString("dd/MM/yyyy") + "'");
            value = po.GetFilterValue(EdiPurchaseOrdersPage.FilterEdi.Sites);
            Assert.IsFalse(((string)value).StartsWith("1 of "), "ResetFilter Sites 'all'");
            value = po.GetFilterValue(EdiPurchaseOrdersPage.FilterEdi.Status);
            Assert.IsFalse(((string)value).StartsWith("1 of "), "ResetFilter Status 'all'");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_EDI_ResetFilter_Ci()
        {
            //Prepare
            object value = null;
            
            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            EdiPage edi = homePage.GoToAccounting_EdiPage();
            EdiCustomerInvoicesPage ci = edi.ClickOnCustomerInvoicesTab();
            ci.ResetFilters();
            // remplir les valeurs
            ci.Filter(EdiCustomerInvoicesPage.FilterEdi.Search, "123");
            value = ci.GetFilterValue(EdiCustomerInvoicesPage.FilterEdi.Search);
            //Assert
            Assert.AreEqual("123", value, "ResetFilter Search '123'");
            ci.Filter(EdiCustomerInvoicesPage.FilterEdi.Customer, "BRIT AIR");
            value = ci.GetFilterValue(EdiCustomerInvoicesPage.FilterEdi.Customer);
            //Assert
            Assert.AreEqual("BRIT AIR", value, "ResetFilter Supplier 'BRIT AIR'");
            ci.Filter(EdiCustomerInvoicesPage.FilterEdi.From, DateUtils.Now.AddMonths(-2));
            value = ci.GetFilterValue(EdiCustomerInvoicesPage.FilterEdi.From);
            //Assert
            Assert.AreEqual(DateUtils.Now.AddMonths(-2).ToString("dd/MM/yyyy"), value, "ResetFilter From '" + DateUtils.Now.AddMonths(-2).ToString("dd/MM/yyyy") + "'");
            ci.Filter(EdiCustomerInvoicesPage.FilterEdi.To, DateUtils.Now.AddMonths(1));
            value = ci.GetFilterValue(EdiCustomerInvoicesPage.FilterEdi.To);
            //Assert
            Assert.AreEqual(DateUtils.Now.AddMonths(1).ToString("dd/MM/yyyy"), value, "ResetFilter To '" + DateUtils.Now.AddMonths(1).ToString("dd/MM/yyyy") + "'");
            ci.Filter(EdiCustomerInvoicesPage.FilterEdi.Sites, "MAD - MAD");
            value = ci.GetFilterValue(EdiCustomerInvoicesPage.FilterEdi.Sites);
            //Assert
            Assert.IsTrue(((string)value).StartsWith("1 of "), "ResetFilter Sites 'MAD - MAD'");
            ci.Filter(EdiCustomerInvoicesPage.FilterEdi.Status, "Received");
            value = ci.GetFilterValue(EdiCustomerInvoicesPage.FilterEdi.Status);

            ci.ResetFilters();
            // valeurs après reset
            value = ci.GetFilterValue(EdiCustomerInvoicesPage.FilterEdi.Search);
            //Assert
            Assert.AreEqual("", value, "ResetFilter Search ''");
            value = ci.GetFilterValue(EdiCustomerInvoicesPage.FilterEdi.Customer);
            //Assert
            Assert.AreEqual("ALL", value, "ResetFilter Supplier ''");
            value = ci.GetFilterValue(EdiCustomerInvoicesPage.FilterEdi.From);
            //Assert
            Assert.AreEqual(DateUtils.Now.AddMonths(-1).ToString("dd/MM/yyyy"), value, "ResetFilter From '" + DateUtils.Now.AddMonths(-1).ToString("dd/MM/yyyy") + "'");
            value = ci.GetFilterValue(EdiCustomerInvoicesPage.FilterEdi.To);
            //Assert
            Assert.AreEqual(DateUtils.Now.ToString("dd/MM/yyyy"), value, "ResetFilter To '" + DateUtils.Now.ToString("dd/MM/yyyy") + "'");
            value = ci.GetFilterValue(EdiCustomerInvoicesPage.FilterEdi.Sites);
            //Assert
            Assert.IsFalse(((string)value).StartsWith("1 of "), "ResetFilter Sites 'all'");
            value = ci.GetFilterValue(EdiCustomerInvoicesPage.FilterEdi.Status);
            //Assert
            Assert.IsFalse(((string)value).StartsWith("1 of "), "ResetFilter Status 'all'");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_EDI_ResetFilter_Co()
        {
            //Prepare

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            EdiPage edi = homePage.GoToAccounting_EdiPage();
            object value = null;

            EdiCustomerOrderPage co = edi.ClickOnCustomerOrderTab();
            co.ResetFilters();
            // remplir les valeurs
            co.Filter(EdiCustomerOrderPage.FilterEdi.SearchCustomerCode, "test");
            value = co.GetFilterValue(EdiCustomerOrderPage.FilterEdi.SearchCustomerCode);
            Assert.AreEqual("test", value, "ResetFilter Search 'test'");
            co.Filter(EdiCustomerOrderPage.FilterEdi.SearchCONumber, "test2");
            value = co.GetFilterValue(EdiCustomerOrderPage.FilterEdi.SearchCONumber);
            Assert.AreEqual("test2", value, "ResetFilter Supplier 'test2'");
            co.Filter(EdiCustomerOrderPage.FilterEdi.From, DateUtils.Now.AddMonths(-2));
            value = co.GetFilterValue(EdiCustomerOrderPage.FilterEdi.From);
            Assert.AreEqual(DateUtils.Now.AddMonths(-2).ToString("dd/MM/yyyy"), value, "ResetFilter From '" + DateUtils.Now.AddMonths(-2).ToString("dd/MM/yyyy") + "'");
            co.Filter(EdiCustomerOrderPage.FilterEdi.To, DateUtils.Now.AddMonths(1));
            value = co.GetFilterValue(EdiCustomerOrderPage.FilterEdi.To);
            Assert.AreEqual(DateUtils.Now.AddMonths(1).ToString("dd/MM/yyyy"), value, "ResetFilter To '" + DateUtils.Now.AddMonths(1).ToString("dd/MM/yyyy") + "'");
            co.Filter(EdiCustomerOrderPage.FilterEdi.Sites, "MAD - MAD");
            value = co.GetFilterValue(EdiCustomerOrderPage.FilterEdi.Sites);
            Assert.IsTrue(((string)value).StartsWith("1 of "), "ResetFilter Sites 'MAD - MAD'");
            co.Filter(EdiCustomerOrderPage.FilterEdi.Import, true);
            value = co.GetFilterValue(EdiCustomerOrderPage.FilterEdi.Import);

            co.ResetFilters();
            // valeurs après reset
            value = co.GetFilterValue(EdiCustomerOrderPage.FilterEdi.SearchCustomerCode);
            Assert.AreEqual("", value, "ResetFilter Search ''");
            value = co.GetFilterValue(EdiCustomerOrderPage.FilterEdi.SearchCONumber);
            Assert.AreEqual("", value, "ResetFilter Supplier ''");
            value = co.GetFilterValue(EdiCustomerOrderPage.FilterEdi.From);
            Assert.AreEqual(DateUtils.Now.ToString("dd/MM/yyyy"), value, "ResetFilter From '" + DateUtils.Now.ToString("dd/MM/yyyy") + "'");
            value = co.GetFilterValue(EdiCustomerOrderPage.FilterEdi.To);
            Assert.AreEqual(DateUtils.Now.ToString("dd/MM/yyyy"), value, "ResetFilter To '" + DateUtils.Now.ToString("dd/MM/yyyy") + "'");
            value = co.GetFilterValue(EdiCustomerOrderPage.FilterEdi.Sites);
            Assert.IsFalse(((string)value).StartsWith("1 of "), "ResetFilter Sites 'all'");
            value = co.GetFilterValue(EdiCustomerOrderPage.FilterEdi.ShowAll);
            Assert.AreEqual(true, value, "ResetFilter ShowAll");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_EDI_Export()
        {
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            homePage.ClearDownloads();
            EdiPage edi = homePage.GoToAccounting_EdiPage();
            EdiSupplierInvoicesPage ediSupplierInvoicesPage = edi.ClickOnSupplierAccountingsTab();

            ediSupplierInvoicesPage.ResetFilters();
            var idsGrid = edi.GetListIds();
            //clear download directory 
            foreach (var file in new DirectoryInfo(downloadsPath).GetFiles())
            {
                file.Delete();
            }
            //export
            edi.Export();

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();


            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = edi.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);
            var idsExcel = OpenXmlExcel.GetValuesInList("Id", "Supplier Invoice imports", filePath);
            //verifier les données 
            var verify = edi.VerifyExcel(idsGrid, idsExcel);
            Assert.IsTrue(verify, "les donées du grid et du fichier Excel ne sont pas identiques");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_EDI_Import()
        {
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            LogInAsAdmin();
            HomePage homePage = new HomePage(WebDriver, TestContext);
            homePage.ClearDownloads();
            EdiPage edi = homePage.GoToAccounting_EdiPage();
            EdiSupplierInvoicesPage ediSupplierInvoicesPage = edi.ClickOnSupplierAccountingsTab();
            int numberRawsListSupplierInvoices = ediSupplierInvoicesPage.GetNumberSupplierInvoicesEdi();
            edi.Export();
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();
            FileInfo correctDownloadedFile = edi.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);
            string fileName = correctDownloadedFile.Name;
            string filePath = Path.Combine(downloadsPath, fileName);
            edi.Import(filePath);
            bool isAddedFileVerif = ediSupplierInvoicesPage.IsAddedFileVerif(numberRawsListSupplierInvoices);
            Assert.IsFalse(isAddedFileVerif, "le fichier d'export .xlsx a pu être importé alors qu'il ne devrait pas.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_EDI_Filter_From()
        {
            var from = DateTime.Now.AddDays(-2);
            from = new DateTime(from.Year, from.Month, from.Day);
            var to = DateTime.Now.AddDays(2);
            to = new DateTime(to.Year, to.Month, to.Day);

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            EdiPage edi = homePage.GoToAccounting_EdiPage();
            EdiSupplierInvoicesPage ediSupplierInvoicesPage = edi.ClickOnSupplierAccountingsTab();
            ediSupplierInvoicesPage.Filter(EdiSupplierInvoicesPage.FilterType.From, from);
            ediSupplierInvoicesPage.Filter(EdiSupplierInvoicesPage.FilterType.To, to);
            ediSupplierInvoicesPage.PageSize("100");

            var datesSupplierInvoicesAfterFiltre = ediSupplierInvoicesPage.GetDatesListSupplierInvoices();
            Assert.IsTrue(ediSupplierInvoicesPage.VerifyDatesSupplierInvoicesInFilter(datesSupplierInvoicesAfterFiltre, from, to), "Filter failed");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_EDI_Filter_To()
        {
            var to = DateTime.Now.AddDays(2);
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            EdiPage ediPage = homePage.GoToAccounting_EdiPage();
            EdiSupplierInvoicesPage ediSupplierInvoicesPage = ediPage.ClickOnSupplierAccountingsTab();
            ediSupplierInvoicesPage.PageSize("100");
            ediPage.Filter(EdiPage.FilterEdi.To, to);
            var datesSupplierInvoicesAfterFiltre = ediSupplierInvoicesPage.GetDatesListSupplierInvoices();
            Assert.IsTrue(ediSupplierInvoicesPage.VerifySupplierInvoicesToDate(datesSupplierInvoicesAfterFiltre, to), "Filter failed");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_EDI_Filter_Search_Si()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            EdiPage ediPage = homePage.GoToAccounting_EdiPage();
            EdiSupplierInvoicesPage ediSupplierInvoicesPage = ediPage.ClickOnSupplierAccountingsTab();
            ediSupplierInvoicesPage.PageSize("100");
            var totalNumberSIEdi = ediSupplierInvoicesPage.CheckTotalNumber();
            ediSupplierInvoicesPage.Filter(EdiSupplierInvoicesPage.FilterType.Search, "test");
            var totalNumberSIEdiAfterSearch = ediSupplierInvoicesPage.CheckTotalNumber();
            Assert.AreNotEqual(totalNumberSIEdiAfterSearch, totalNumberSIEdi, "Search filter ne fonctionne pas.");
            ediSupplierInvoicesPage.ResetFilters();
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_EDI_Filter_BySupplier()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            EdiPage ediPage = homePage.GoToAccounting_EdiPage();
            EdiSupplierInvoicesPage ediSupplierInvoicesPage = ediPage.ClickOnSupplierAccountingsTab();
            ediSupplierInvoicesPage.PageSize("100");
            ediSupplierInvoicesPage.Filter(EdiSupplierInvoicesPage.FilterType.From, DateTime.Today.AddMonths(-4));
            ediSupplierInvoicesPage.Filter(EdiSupplierInvoicesPage.FilterType.Supplier, "ALL");
            var totalNumberSIEdiAllSuppliers = ediSupplierInvoicesPage.CheckTotalNumber();
            ediSupplierInvoicesPage.Filter(EdiSupplierInvoicesPage.FilterType.Supplier, "A REFERENCIAR");
            var totalNumberSIEdiAfterFilter = ediSupplierInvoicesPage.CheckTotalNumber();
            Assert.AreNotEqual(totalNumberSIEdiAfterFilter, totalNumberSIEdiAllSuppliers, "Supplier filter ne fonctionne pas.");
            ediSupplierInvoicesPage.ResetFilters();
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_EDI_Filter_Sites()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            EdiPage ediPage = homePage.GoToAccounting_EdiPage();
            EdiSupplierInvoicesPage ediSupplierInvoicesPage = ediPage.ClickOnSupplierAccountingsTab();
            ediSupplierInvoicesPage.PageSize("100");
            ediSupplierInvoicesPage.Filter(EdiSupplierInvoicesPage.FilterType.From, DateTime.Today.AddMonths(-4));
            ediSupplierInvoicesPage.Filter(EdiSupplierInvoicesPage.FilterType.To, DateTime.Today);
            var totalNumberSIEdi = ediSupplierInvoicesPage.CheckTotalNumber();
            ediSupplierInvoicesPage.Filter(EdiSupplierInvoicesPage.FilterType.Sites, "ACE");
            var totalNumberSIEdiAfterFilter = ediSupplierInvoicesPage.CheckTotalNumber();
            Assert.AreNotEqual(totalNumberSIEdiAfterFilter, totalNumberSIEdi, "Sites filter ne fonctionne pas.");
            ediSupplierInvoicesPage.ResetFilters();
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_EDI_Filter_Statuts()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            EdiPage ediPage = homePage.GoToAccounting_EdiPage();
            EdiSupplierInvoicesPage ediSupplierInvoicesPage = ediPage.ClickOnSupplierAccountingsTab();
            ediSupplierInvoicesPage.PageSize("100");
            ediSupplierInvoicesPage.Filter(EdiSupplierInvoicesPage.FilterType.From, DateTime.Today.AddMonths(-4));
            ediSupplierInvoicesPage.Filter(EdiSupplierInvoicesPage.FilterType.To, DateTime.Today);
            var totalNumberSIEdi = ediSupplierInvoicesPage.CheckTotalNumber();
            ediSupplierInvoicesPage.Filter(EdiSupplierInvoicesPage.FilterType.Status, "Matching error File / Winrest");
            var totalNumberSIEdiAfterFilter = ediSupplierInvoicesPage.CheckTotalNumber();
            Assert.AreNotEqual(totalNumberSIEdiAfterFilter, totalNumberSIEdi, "Statuts filter ne fonctionne pas.");
            ediSupplierInvoicesPage.ResetFilters();
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_EDI_Filter_From_Po()
        {
            //Prepare 
            var from = DateTime.Now.AddMonths(-1);
            from = new DateTime(from.Year, from.Month, from.Day);
            var to = DateTime.Now.AddDays(2);
            to = new DateTime(to.Year, to.Month, to.Day);
            
            //Arrange
            HomePage homePage = LogInAsAdmin();
            
            //Act
            EdiPage ediPage = homePage.GoToAccounting_EdiPage();
            var ediPurchaseOrdersPage = ediPage.ClickOnPurchaseOrdersTab();
            ediPurchaseOrdersPage.Filter(EdiPurchaseOrdersPage.FilterEdi.From, from);
            ediPurchaseOrdersPage.Filter(EdiPurchaseOrdersPage.FilterEdi.To, to);
            ediPurchaseOrdersPage.PageSize("100");
            
            //Assert
            var datesPurchaseOrdersAfterFiltre = ediPurchaseOrdersPage.GetDatesListPurchaseOrders();
            bool verifFilterDate = ediPurchaseOrdersPage.VerifyDatesPurchaseOrdersInFilter(datesPurchaseOrdersAfterFiltre, from, to); 
            Assert.IsTrue(verifFilterDate, "Filter failed");

        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_EDI_Filter_To_Po()
        {
            //Prepare 
            var to = DateTime.Now.AddDays(2);
            to = new DateTime(to.Year, to.Month, to.Day);

            //Arrange
            HomePage homePage = LogInAsAdmin();
          
            //Act
            EdiPage ediPage = homePage.GoToAccounting_EdiPage();
            var ediPurchaseOrdersPage = ediPage.ClickOnPurchaseOrdersTab();          
            ediPurchaseOrdersPage.PageSize("100");
            ediPurchaseOrdersPage.Filter(EdiPurchaseOrdersPage.FilterEdi.To, to);
            var datesPurchaseOrdersAfterFilter = ediPurchaseOrdersPage.GetDatesListPurchaseOrders();
            
            //Assert
            bool isVerifyPurchaseOrdersOk = ediPurchaseOrdersPage.VerifyPurchaseOrdersToDate(datesPurchaseOrdersAfterFilter, to); 
            Assert.IsTrue(isVerifyPurchaseOrdersOk, "Filter Date to ne fonctionne pas");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_EDI_Filter_FromTo_Co()
        {
            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            EdiPage ediPage = homePage.GoToAccounting_EdiPage();
            var ediCustomerOrdersPage = ediPage.ClickOnCustomerOrderTab();
            var from = DateTime.Now.AddMonths(-2);
            from = new DateTime(from.Year, from.Month, from.Day);
            var to = DateTime.Now.AddDays(2);
            to = new DateTime(to.Year, to.Month, to.Day);

            ediCustomerOrdersPage.Filter(EdiCustomerOrderPage.FilterEdi.From, from);
            ediCustomerOrdersPage.Filter(EdiCustomerOrderPage.FilterEdi.To, to);
            ediCustomerOrdersPage.PageSize("100");

            var datesCustomerOrdersAfterFiltre = ediCustomerOrdersPage.GetDatesListCustomerOrders();

            //Assert
            bool verifyDatesCustomerOrdersInFilter = ediCustomerOrdersPage.VerifyDatesCustomerOrdersInFilter(datesCustomerOrdersAfterFiltre, from, to);
            Assert.IsTrue(verifyDatesCustomerOrdersInFilter, "Filter failed");

        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_EDI_Filter_From_Ci()
        {
            //Prepare 
            DateTime dateFrom = DateTime.Today.AddMonths(-1);

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            EdiPage ediPage = homePage.GoToAccounting_EdiPage();
            EdiCustomerInvoicesPage ediCustomerInvoicesPage = ediPage.ClickOnCustomerInvoicesTab();
            ediCustomerInvoicesPage.Filter(EdiCustomerInvoicesPage.FilterEdi.From, dateFrom);
            ediCustomerInvoicesPage.PageSize("100");
            var datesCustomerInvoicesFiltre = ediCustomerInvoicesPage.GetDatesListCustomerInvoices();
           
            //Assert
            bool verifyDatesCustomerInvoicesInFilter = ediCustomerInvoicesPage.VerifyDatesCustomerInvoicesFilterFrom(datesCustomerInvoicesFiltre, dateFrom);
            Assert.IsTrue(verifyDatesCustomerInvoicesInFilter, "Filter From failed");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_EDI_Filter_To_Ci()
        {
            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            EdiPage ediPage = homePage.GoToAccounting_EdiPage();
            EdiCustomerInvoicesPage ediCustomerInvoicesPage = ediPage.ClickOnCustomerInvoicesTab();

            var to = DateTime.Today;
            ediCustomerInvoicesPage.Filter(EdiCustomerInvoicesPage.FilterEdi.To, to);
            ediCustomerInvoicesPage.PageSize("100");

            var datesCustomerInvoicesFiltre = ediCustomerInvoicesPage.GetDatesListCustomerInvoices();
            //Assert
            bool verifyDatesCustomerInvoicesInFilter = ediCustomerInvoicesPage.VerifyDatesCustomerInvoicesFilterTo(datesCustomerInvoicesFiltre, to);
            Assert.IsTrue(verifyDatesCustomerInvoicesInFilter, "Filter To failed");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_EDI_Filter_Search_Ci()
        {
            //Arrange
            HomePage homePage = LogInAsAdmin();
           
            //Act
            var ediPage = homePage.GoToAccounting_EdiPage();
            var ci = ediPage.ClickOnCustomerInvoicesTab();
            ci.ResetFilters();
            var invoiceNumber = ci.GetFirstCIInvoiceNumber();
           
            //Assert
            ci.Filter(EdiCustomerInvoicesPage.FilterEdi.Search, invoiceNumber);
            string firstNumber = ci.GetFirstCIInvoiceNumber(); 
            Assert.AreEqual(invoiceNumber, firstNumber, "Failed Filter Search");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_EDI_Filter_Search_Po()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            EdiPage ediPage = homePage.GoToAccounting_EdiPage();
            EdiPurchaseOrdersPage ediPurchaseOrdersPage = ediPage.ClickOnPurchaseOrdersTab();
            ediPurchaseOrdersPage.ResetFilters();
            ediPurchaseOrdersPage.Filter(EdiPurchaseOrdersPage.FilterEdi.From, DateTime.Today.AddMonths(-4));
            //en comment pas de données
            //string PoNumber = ediPurchaseOrdersPage.GetfirstNumberPo();
            //ediPurchaseOrdersPage.Filter(EdiPurchaseOrdersPage.FilterEdi.Search, PoNumber);
            //ediPurchaseOrdersPage.WaitPageLoading();
            //var PoNumberAfterSearch = ediPurchaseOrdersPage.GetfirstNumberPo();
            //Assert.AreEqual(PoNumberAfterSearch, PoNumber, "Search filter ne fonctionne pas.");
            //ediPurchaseOrdersPage.ResetFilters();
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_EDI_Filter_Site_Co()
        {
            //Prepare 
            string site = TestContext.Properties["SiteACE"].ToString();
            
            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            EdiPage ediPage = homePage.GoToAccounting_EdiPage();
            EdiCustomerOrderPage ediCustomerOrderPage = ediPage.ClickOnCustomerOrderTab();
            ediCustomerOrderPage.PageSize("100");
            var totalNumberSIEdi = ediCustomerOrderPage.CheckTotalNumber();        
            ediCustomerOrderPage.Filter(EdiCustomerOrderPage.FilterEdi.Sites, site);
            var totalNumberSIEdiAfter = ediCustomerOrderPage.CheckTotalNumber();
            //pas de données, à revoir quand données
            //Assert
            Assert.AreNotEqual(totalNumberSIEdi, totalNumberSIEdiAfter, "Filter par site ne fonctionne pas.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_EDI_Filter_Sites_Ci()
        {
            //Prepare 
            HomePage homePage = LogInAsAdmin();

            //Arrange
            var ediPage = homePage.GoToAccounting_EdiPage();
            var ci = ediPage.ClickOnCustomerInvoicesTab();
            ci.ResetFilters();
            var siteName = ci.GetFirstCISiteName();
            ci.Filter(EdiCustomerInvoicesPage.FilterEdi.Sites, siteName);
            ci.WaitPageLoading();
            var sitesNames = ci.GetListSiteName();
            //Assert
            bool isSiteExists = sitesNames.Any(name => name.Equals(siteName)); 
            Assert.IsTrue(isSiteExists, "Soucis Sites filter ");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_EDI_Filter_ByCustomer()
        {
            //Prepare 
            string all = "ALL";
            string customer = "BRIT AIR";

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            EdiPage ediPage = homePage.GoToAccounting_EdiPage();
            EdiCustomerInvoicesPage ediCustomerInvoicesPage = ediPage.ClickOnCustomerInvoicesTab();
            ediCustomerInvoicesPage.ResetFilters();
            ediCustomerInvoicesPage.PageSize("100");
            ediCustomerInvoicesPage.Filter(EdiCustomerInvoicesPage.FilterEdi.Customer, all);
            var totalNumberCIEdiAllCustomers = ediCustomerInvoicesPage.CheckTotalNumber();
            ediCustomerInvoicesPage.Filter(EdiCustomerInvoicesPage.FilterEdi.Customer, customer);
            ediCustomerInvoicesPage.WaitPageLoading();
            var totalNumberCIEdiAfterFilter = ediCustomerInvoicesPage.CheckTotalNumber();

            //Assert
            Assert.AreNotEqual(totalNumberCIEdiAfterFilter, totalNumberCIEdiAllCustomers, "Customer filter for Customer Invoices ne fonctionne pas.");


        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_EDI_Filter_Sites_Po()
        {
            //Prepare 
            string siteName = TestContext.Properties["Site"].ToString();
           
            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            EdiPage ediPage = homePage.GoToAccounting_EdiPage();
            EdiPurchaseOrdersPage ediPurchaseOrdersPage = ediPage.ClickOnPurchaseOrdersTab();
            ediPurchaseOrdersPage.ResetFilters();

            ediPurchaseOrdersPage.Filter(EdiPurchaseOrdersPage.FilterEdi.Sites, siteName);
            if (ediPurchaseOrdersPage.CheckTotalNumber() > 0)
            {
                ediPurchaseOrdersPage.PageSize("100");
                var sitesNames = ediPurchaseOrdersPage.GetSiteNamesListPurchaseOrders();
                Assert.IsFalse(sitesNames.Any(name => !name.Equals(siteName)), "Soucis Sites filter ");
            }
            ediPurchaseOrdersPage.ResetFilters();
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_EDI_Filter_Statuts_Ci()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            var ediPage = homePage.GoToAccounting_EdiPage();
            var ci = ediPage.ClickOnCustomerInvoicesTab();
            ci.ResetFilters();
            var status = ci.GetFirstCIStatus();
            ci.Filter(EdiCustomerInvoicesPage.FilterEdi.Status, status);
            ci.PageSize("100");
            var statuslist = ci.GetListStatus();
            Assert.IsTrue(statuslist.Any(name => name.Equals(status)), "Soucis Statuts filter ");
            ci.ResetFilters();
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_EDI_Filter_ShowAll_Sionly_CNonly()
        {
            {
                LogInAsAdmin();
                var homePage = new HomePage(WebDriver, TestContext);
                EdiPage ediPage = homePage.GoToAccounting_EdiPage();
                EdiSupplierInvoicesPage ediSupplierInvoicesPage = ediPage.ClickOnSupplierAccountingsTab();
                ediSupplierInvoicesPage.ResetFilters();
                ediSupplierInvoicesPage.PageSize("100");

                ediSupplierInvoicesPage.Filter(EdiSupplierInvoicesPage.FilterType.ShowAll, true);
                var totalNumberSIEdiAfterSearchAll = ediSupplierInvoicesPage.CheckTotalNumber();
                ediSupplierInvoicesPage.Filter(EdiSupplierInvoicesPage.FilterType.ShowOnlyCN, true);
                var totalNumberSIEdiAfterSearchCn = ediSupplierInvoicesPage.CheckTotalNumber();
                Assert.AreNotEqual(totalNumberSIEdiAfterSearchAll, totalNumberSIEdiAfterSearchCn, "Show CN only filter ne fonctionne pas.");
                ediSupplierInvoicesPage.Filter(EdiSupplierInvoicesPage.FilterType.ShowAll, true);
                ediSupplierInvoicesPage.Filter(EdiSupplierInvoicesPage.FilterType.ShowOnlySI, true);
                var totalNumberSIEdiAfterSearchSI = ediSupplierInvoicesPage.CheckTotalNumber();
                Assert.AreNotEqual(totalNumberSIEdiAfterSearchAll, totalNumberSIEdiAfterSearchSI, "Show SI only filter ne fonctionne pas.");
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_EDI_Filter_BySupplier_Po()
        {
            {
                //Prepare 
                string supplierFilter = "AGUACANA, S.A.";

                //Arrange
                HomePage homePage = LogInAsAdmin();
            
                //Act
                EdiPage ediPage = homePage.GoToAccounting_EdiPage();
                EdiPurchaseOrdersPage ediPurchaseOrdersPage = ediPage.ClickOnPurchaseOrdersTab();
                ediPurchaseOrdersPage.Filter(EdiPurchaseOrdersPage.FilterEdi.Supplier, supplierFilter);
                var supplierResult = ediPurchaseOrdersPage.GetFilterValue(EdiPurchaseOrdersPage.FilterEdi.Supplier);
              
                //Assert
                Assert.AreEqual(supplierFilter, supplierResult.ToString(), "Les données ne sont pas filtré par supplier.");
            }

        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_EDI_Filter_Statuts_Po()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            // Act
            EdiPage ediPage = homePage.GoToAccounting_EdiPage();
            EdiPurchaseOrdersPage ediPurchaseOrdersPage = ediPage.ClickOnPurchaseOrdersTab();
            ediPurchaseOrdersPage.ResetFilters();
            string status = "all";
            var totalNumberPREDIBeforeSearchAll = ediPurchaseOrdersPage.CheckTotalNumber();
            ediPurchaseOrdersPage.Filter(EdiPurchaseOrdersPage.FilterEdi.Status, status);
            var statusResult = ediPurchaseOrdersPage.GetFilterValue(EdiPurchaseOrdersPage.FilterEdi.Status);
            Assert.AreNotEqual(statusResult, totalNumberPREDIBeforeSearchAll.ToString());

        }


        [Timeout(_timeout)]
        [TestMethod]
        public void AC_EDI_MessageIndex()
        {
            var from = DateTime.Now.AddMonths(-1);
            from = new DateTime(from.Year, from.Month, from.Day);
            var to = DateTime.Now.AddDays(2);
            to = new DateTime(to.Year, to.Month, to.Day);

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            EdiPage edi = homePage.GoToAccounting_EdiPage();
            EdiSupplierInvoicesPage ediSupplierInvoicesPage = edi.ClickOnSupplierAccountingsTab();
            ediSupplierInvoicesPage.Filter(EdiSupplierInvoicesPage.FilterType.From, from);
            ediSupplierInvoicesPage.Filter(EdiSupplierInvoicesPage.FilterType.To, to);
            ediSupplierInvoicesPage.PageSize("100");

            bool IsStatusAfficheOK = ediSupplierInvoicesPage.IsListeStatusAffiche();
            Assert.IsTrue(IsStatusAfficheOK, "Les messages status ne s'affiche pas dans l'index des EDIs.");

            var settingsPage = homePage.GoToApplicationSettings();
            var paramPage = settingsPage.GoToParameters_GlobalSettings();
            var translation = paramPage.ClickOnTranslationTab();
            paramPage.Filter(PageObjects.Parameters.GlobalSettings.ParametersGlobalSettings.FilterEdi.culture, "English");
            paramPage.Filter(PageObjects.Parameters.GlobalSettings.ParametersGlobalSettings.FilterEdi.module, "Accounting");

            bool IsListAfficheOK = paramPage.IsListeAffiche();
            Assert.IsTrue(IsListAfficheOK, "Les messages ne sont pas bien paramétrés depuis l'onglet Traductions.");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_EDI_Export_Statut()
        {
            var from = DateTime.Now.AddMonths(-1);
            from = new DateTime(from.Year, from.Month, from.Day);
            var to = DateTime.Now;
            to = new DateTime(to.Year, to.Month, to.Day);
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            EdiPage edi = homePage.GoToAccounting_EdiPage();
            edi.ResetFilters();
            EdiSupplierInvoicesPage ediSupplierInvoicesPage = edi.ClickOnSupplierAccountingsTab();
            ediSupplierInvoicesPage.Filter(EdiSupplierInvoicesPage.FilterType.From, from);
            ediSupplierInvoicesPage.Filter(EdiSupplierInvoicesPage.FilterType.To, to);
            ediSupplierInvoicesPage.PageSize("100");

            bool IsStatusAfficheOK = ediSupplierInvoicesPage.IsListeStatusAffiche();
            Assert.IsTrue(IsStatusAfficheOK, "Les messages status ne s'affiche pas dans l'index des EDIs.");

            var coloneStatutGrid = edi.GetListStatuts();

            edi.ClearDownloads();
            //clear download directory 
            foreach (var file in new DirectoryInfo(downloadsPath).GetFiles())
            {
                file.Delete();
            }
            //export
            edi.Export();

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = edi.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);
            var statutExcel = OpenXmlExcel.GetValuesInList("Info", "Supplier Invoice imports", filePath);

            //verifier les données 
            var verify = edi.VerifyStatutExcel(coloneStatutGrid, statutExcel);
            Assert.IsTrue(verify, "les donées du grid et du fichier Excel ne sont pas identiques");


        }


        [Timeout(_timeout)]
        [TestMethod]
        public void AC_EDI_VisuelIHM()
        {
            var from = DateTime.Now.AddMonths(-1);
            from = new DateTime(from.Year, from.Month, from.Day);
            var to = DateTime.Now.AddDays(2);
            to = new DateTime(to.Year, to.Month, to.Day);
            string format = "Format";
            string fileName = "File name";
            string fileParsing = "File parsing";

            //arrange
            HomePage homePage = LogInAsAdmin();

            EdiPage edi = homePage.GoToAccounting_EdiPage();
            EdiSupplierInvoicesPage ediSupplierInvoicesPage = edi.ClickOnSupplierAccountingsTab();
            ediSupplierInvoicesPage.ResetFilters();
            ediSupplierInvoicesPage.Filter(EdiSupplierInvoicesPage.FilterType.From, from);
            ediSupplierInvoicesPage.Filter(EdiSupplierInvoicesPage.FilterType.To, to);
            var firstEditDetails = ediSupplierInvoicesPage.ClickOnFirstItem();
            bool IsRetourligneFormatOK = firstEditDetails.IsRetourligneFormatEDI(format, fileName, fileParsing);
            Assert.IsTrue(IsRetourligneFormatOK, "EDIFormat, File name et Related invoice(s) doivent être chacune sur une ligne différente.");

        }


    }
}
