using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.Utils;
using System;
using Newrest.Winrest.FunctionalTests.PageObjects.Interim.InterimReceptions;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using System.Security.Policy;
using iText.StyledXmlParser.Jsoup.Helper;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.Flight;
using System.Drawing.Printing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Sites;
using DocumentFormat.OpenXml.Presentation;
using Newrest.Winrest.FunctionalTests.PageObjects.Jobs.ScheduledJobs;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.ReceiptNotes;
using Microsoft.VisualBasic.Devices;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Delivery;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Bibliography;
using iText.Kernel.Geom;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Item;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Suppliers;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Threading;
using OpenQA.Selenium.Support.UI;
using Newrest.Winrest.FunctionalTests.PageObjects.Accounting.SupplierInvoices;
using Newrest.Winrest.FunctionalTests.PageObjects.Interim.InterimOrders;
using System.Threading;
using Microsoft.VisualBasic;
using Newrest.Winrest.FunctionalTests.PageObjects.Interim.InterimOrders;
using DocumentFormat.OpenXml.Drawing;

namespace Newrest.Winrest.FunctionalTests.Interim.InterimReception
{

    [TestClass]
    public class InterimReceptionTest : TestBase
    {

        private const int _timeout = 600000;
        //_________________________________________CREATE_INTERIM____________________________________________________________

		[TestMethod]
        [Priority(0)]
        [Timeout(_timeout)]
        public void INT_RECEP_index_CreateData()
        {
            string site = TestContext.Properties["SiteACE"].ToString();
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string group = "A REFERENCIA";
            string workshop = TestContext.Properties["Workshop1"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string location = "Produccion";
            string item = "ItemForInterim";
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string storageUnit = "KG";
            string storageQty = 10.ToString();
            string qty = 10.ToString();

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            SuppliersPage suppliersPage = homePage.GoToPurchasing_SuppliersPage();
            suppliersPage.ResetFilter();
            suppliersPage.Filter(SuppliersPage.FilterType.Search, supplier);
            if (suppliersPage.CheckTotalNumber() == 0)
            {
                SupplierCreateModalPage supplierCreateModalPage = suppliersPage.SupplierCreatePage();
                supplierCreateModalPage.FillField_CreatNewSupplier(supplier, true);
                SupplierGeneralInfoTab supplierInfo = supplierCreateModalPage.SubmitToGeneralInfo();
                supplierInfo.SetSite();
                supplierInfo.SetSupplierType("InterimSupplierType");
                supplierInfo.PageUp();
                SupplierItem supplierItem = supplierInfo.ClickOnItemsTab();
                ItemCreateModalPage itemCreateModalPage = supplierItem.ItemCreatePage();
                ItemGeneralInformationPage itemGeneralInformationPage = itemCreateModalPage.FillField_CreateNewItem(item, group, workshop, taxType, prodUnit);
                ItemCreateNewPackagingModalPage itemCreateNewPackagingModalPage = itemGeneralInformationPage.NewPackaging();
                itemCreateNewPackagingModalPage.FillField_CreateNewPackaging(site, packagingName, storageQty, storageUnit, qty, supplier);

            }
            InterimReceptionsPage interimPage = homePage.GoToInterim_Receptions();
            InterimReceptionsCreateModalPage modalcreateInterim = interimPage.CreateNewInterim();
            modalcreateInterim.FillField_CreatNewIntermin(DateTime.Now, site, supplier, location);
            var InterimId = modalcreateInterim.GetInterimId();
            InterimReceptionsItem InterimItem = modalcreateInterim.Submit();
            interimPage = InterimItem.BackToList();

            //Assert
            Assert.AreEqual(interimPage.GetFirstID(), InterimId, String.Format(MessageErreur.OBJET_NON_CREE, "L'Interim"));
        }
        /*
         * Création d'une nouvelle Interim
        */
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_RECEP_index_AddNew()
        {
            //Prepare
            string site = TestContext.Properties["SiteACE"].ToString();
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string location = "Produccion";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            InterimReceptionsPage interimPage = homePage.GoToInterim_Receptions();
            InterimReceptionsCreateModalPage modalcreateInterim = interimPage.CreateNewInterim();
            modalcreateInterim.FillField_CreatNewIntermin(DateTime.Now, site, supplier, location);
            var InterimId = modalcreateInterim.GetInterimId();
            InterimReceptionsItem InterimItem = modalcreateInterim.Submit();
            interimPage = InterimItem.BackToList();

            //Assert
            Assert.AreEqual(interimPage.GetFirstID(), InterimId, String.Format(MessageErreur.OBJET_NON_CREE, "L'Interim"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void INT_RECEP_index_AddNewCopyFromOrderPrefill()
        {
            //Prepare
            string site = TestContext.Properties["SiteACE"].ToString();
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string location = "Produccion";
            int index = 0;

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            InterimReceptionsPage interimPage = homePage.GoToInterim_Receptions();
            InterimReceptionsCreateModalPage modalcreateInterim = interimPage.CreateNewInterim();
            modalcreateInterim.FillField_CreatNewIntermin(DateTime.Now, site, supplier, location);
            modalcreateInterim.fillCreateInterimReceptionAndRecivedFrom();
            InterimReceptionsItem InterimItem = modalcreateInterim.Submit();
            Thread.Sleep(2000);
            var orderedBefore = InterimItem.GetDeliveredQty(index);
            var recivedBefore = InterimItem.GetDeliveredQty(index);
            var receptionNumber = InterimItem.GetReceptionNumber();
            interimPage = InterimItem.BackToList();
            interimPage.Filter(InterimReceptionsPage.FilterType.ShowAll, true);
            interimPage.Filter(InterimReceptionsPage.FilterType.Bynumber, receptionNumber);
            interimPage.SelectInterimReceptionsItem(1);
            var orderedAfter = InterimItem.GetDeliveredQty(index);
            var recivedAfter = InterimItem.GetDeliveredQty(index);
            //Assert
            Assert.AreEqual(orderedBefore, orderedAfter, "La quantité ordered ne correspond pas à la quantité attendue");
            Assert.AreEqual(recivedBefore, recivedAfter, "La quantité recived ne correspond pas à la quantité attendue");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void INT_RECEP_index_AddNewCopyFromReceptionSelect()
        {
            //Prepare
            string site = TestContext.Properties["SiteACE"].ToString();
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string location = "Produccion";
            string qte = "5";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            InterimReceptionsPage interimPage = homePage.GoToInterim_Receptions();
            InterimReceptionsCreateModalPage modalcreateInterim = interimPage.CreateNewInterim();
            modalcreateInterim.FillField_CreatNewIntermin(DateTime.Now, site, supplier, location);
            InterimReceptionsItem InterimItem = modalcreateInterim.Submit();
            string randQty = new Random().Next(2, 10).ToString();
            InterimItem.ClickOnItem();
            InterimItem.SetQty(qte);
            InterimItem.Validate();
            string receptionNumber = InterimItem.GetReceptionNumber();
            string totalVAT = InterimItem.showTotalPriceToCopy();
            InterimItem.BackToList();
            interimPage.CreateNewInterim();
            modalcreateInterim.FillField_CreatNewIntermin(DateTime.Now, site, supplier, location);  
            modalcreateInterim.ClickOnCreateInterimReceptionFrom();
            modalcreateInterim.toAninterimReception();
            modalcreateInterim.SearchForInterimOrderNumber(receptionNumber);
            modalcreateInterim.Submit();
            string totalWOvatitem = InterimItem.showTotalPriceToCopy();

            //Assert
            Assert.AreEqual(totalVAT, totalWOvatitem);
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void INT_RECEP_index_AddNewCopyFromReceptionSelects()
        {
            //Prepare
            string site = TestContext.Properties["SiteACE"].ToString();
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string location = "Produccion";
            string qte = "5";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            InterimReceptionsPage interimPage = homePage.GoToInterim_Receptions();
            InterimReceptionsCreateModalPage modalcreateInterim = interimPage.CreateNewInterim();
            /* first */
            modalcreateInterim.FillField_CreatNewIntermin(DateTime.Now, site, supplier, location);
            InterimReceptionsItem InterimItem = modalcreateInterim.Submit();
            string randQty = new Random().Next(2, 10).ToString();
            InterimItem.ClickOnItem();
            InterimItem.SetQty(qte);
            float totalVAT1 = InterimItem.showTotalPriceToCopy2();
            InterimItem.Validate();
            string receptionNumber1 = InterimItem.GetReceptionNumber();
            InterimItem.BackToList();

            /*second */
            interimPage.CreateNewInterim();
            modalcreateInterim.FillField_CreatNewIntermin(DateTime.Now, site, supplier, location);
            modalcreateInterim.Submit();
            string randQty2 = new Random().Next(2, 10).ToString();
            InterimItem.ClickOnItem();
            InterimItem.SetQty(qte);
            //Thread.Sleep(2000);
            InterimItem.Validate();
            string receptionNumber2 = InterimItem.GetReceptionNumber();
            float totalVAT2 = InterimItem.showTotalPriceToCopy2();
            InterimItem.BackToList();
            /* sum*/
            float totalForBothInterimOrder = totalVAT1 + totalVAT2;
            //1 + 2 
            interimPage.CreateNewInterim();
            modalcreateInterim.FillField_CreatNewIntermin(DateTime.Now, site, supplier, location);

            modalcreateInterim.ClickOnCreateInterimReceptionFrom();
            modalcreateInterim.toAninterimReception();
            modalcreateInterim.SearchForInterimOrderNumber(receptionNumber1);
            modalcreateInterim.SearchForInterimSecondNumber(receptionNumber2);
            modalcreateInterim.selectSecond();
            modalcreateInterim.Submit();
            float totalWOvatitem = InterimItem.GetTotalWOVATinNumbers();

            //Assert
            Assert.AreEqual(totalForBothInterimOrder, totalWOvatitem, "the total VAT are not equal ");
        }
        /*
         * Création d'une nouvelle Interim Validator
        */
        
		[TestMethod]
        [Timeout(_timeout)]
        public void INT_RECEP_index_AddNewValidator()
        {
            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var interimPage = homePage.GoToInterim_Receptions();
            var modalcreateInterim = interimPage.CreateNewInterim();
            var InterimItem = modalcreateInterim.Submit();
            //Assert
            Assert.IsTrue(!interimPage.CheckValidator(), "Les validators n'apparaissent pas!");

        }
        //_________________________________________FIN CREATE_INTERIM________________________________________________________

        //_________________________________________DELETE_INTERIM________________________________________________________
        /*
         * Supprimer Interim non validée
        */
        
		[TestMethod]
        [Timeout(_timeout)]
        public void INT_RECEP_index_Delete()
        {
            string site = TestContext.Properties["SiteACE"].ToString();
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string location = "Produccion";
            // Log in
            var homePage = LogInAsAdmin();
            //Act
            var interimPage = homePage.GoToInterim_Receptions();
            // Create
            var InterimCreateModalpage = interimPage.CreateNewInterim();
            InterimCreateModalpage.FillField_CreatNewIntermin(DateUtils.Now, site, supplier, location, true);
            var InterimId = InterimCreateModalpage.GetInterimId();
            var InterimItem = InterimCreateModalpage.Submit();

            interimPage = InterimItem.BackToList();
            //Assert
            Assert.AreEqual(interimPage.GetFirstID(), InterimId, string.Format(MessageErreur.OBJET_NON_CREE, "La Interim"));

            //Delete
            interimPage.Filter(InterimReceptionsPage.FilterType.ShowNotValidatedOnly, true);
            interimPage.DeleteInterim(InterimId);
            Assert.AreNotEqual(interimPage.GetFirstID(), InterimId, string.Format(MessageErreur.OBJET_NON_SUPPRIME, "La Interim"));
        }
        //_________________________________________FIN DELETE_INTERIM________________________________________________________


        
		[TestMethod]
        [Timeout(_timeout)]
        public void INT_RECEP_details_Pagination()
        {
            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            var receptionsPage = homePage.GoToInterim_Receptions();
            receptionsPage.PageSize("8");
            Assert.IsTrue(receptionsPage.GetNameProvidersList().Count <= 8, "Paggination ne fonctionne pas..");
            receptionsPage.PageSize("16");
            Assert.IsTrue(receptionsPage.GetNameProvidersList().Count <= 16, "Paggination ne fonctionne pas..");
            receptionsPage.PageSize("30");
            Assert.IsTrue(receptionsPage.GetNameProvidersList().Count <= 30, "Paggination ne fonctionne pas..");
            receptionsPage.PageSize("50");
            Assert.IsTrue(receptionsPage.GetNameProvidersList().Count <= 50, "Paggination ne fonctionne pas..");
            receptionsPage.PageSize("100");
            Assert.IsTrue(receptionsPage.GetNameProvidersList().Count <= 100, "Paggination ne fonctionne pas..");
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_RECEP_Details_ResetFilter()
        {
            //Prepare
            object value = null;
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            InterimReceptionsPage interimReceptionsPage = homePage.GoToInterim_Receptions();
            interimReceptionsPage.ResetFilters();
            var firstNumber = interimReceptionsPage.GetFirstInterimReceptionsNumber();
            var numberDefaultListSiteFilter = interimReceptionsPage.GetNumberSelectedSiteFilter();
            var numberDefaultListSupplierFilter = interimReceptionsPage.GetNumberSelectedSupplierFilter();

            // remplir les valeurs
            interimReceptionsPage.Filter(InterimReceptionsPage.FilterType.ByNumber, firstNumber);
            interimReceptionsPage.Filter(InterimReceptionsPage.FilterType.Suppliers, supplier);
            interimReceptionsPage.Filter(InterimReceptionsPage.FilterType.DateFrom, DateUtils.Now.AddDays(-15));
            interimReceptionsPage.Filter(InterimReceptionsPage.FilterType.DateTo, DateUtils.Now.AddDays(15));
            interimReceptionsPage.Filter(InterimReceptionsPage.FilterType.ShowAllReceptions, true);
            interimReceptionsPage.Filter(InterimReceptionsPage.FilterType.ShowNotValidatedOnly, true);
            interimReceptionsPage.Filter(InterimReceptionsPage.FilterType.ShowValidatedOnly, true);
            interimReceptionsPage.Filter(InterimReceptionsPage.FilterType.ShowAll, true);
            interimReceptionsPage.Filter(InterimReceptionsPage.FilterType.ShowOnlyInactive, true);
            interimReceptionsPage.Filter(InterimReceptionsPage.FilterType.ShowOnlyActive, true);
            interimReceptionsPage.Filter(InterimReceptionsPage.FilterType.ShowAllInterimReception, true);
            interimReceptionsPage.Filter(InterimReceptionsPage.FilterType.ShowClosedInterimReception, true);
            interimReceptionsPage.Filter(InterimReceptionsPage.FilterType.ShowOpenedInterimReception, true);
            interimReceptionsPage.ScrollUntilSitesFilterIsVisible();
            interimReceptionsPage.Filter(InterimReceptionsPage.FilterType.Site, site);

            //Cliquer sur Reset Filtrer
            interimReceptionsPage.ScrollUntilResetFilterIsVisible();
            interimReceptionsPage.ResetFilters();

            //valeurs après reset
            value = interimReceptionsPage.GetFilterValue(InterimReceptionsPage.FilterType.ByNumber);
            Assert.AreEqual("", value, "ResetFilter ByNumber ''");

            var numberSelectedSiteFilter = interimReceptionsPage.GetNumberSelectedSiteFilter();
            Assert.AreEqual(numberSelectedSiteFilter, numberDefaultListSiteFilter, "ResetFilter Site ''");

            var numberSelectedSupplierFilter = interimReceptionsPage.GetNumberSelectedSupplierFilter();
            Assert.AreEqual(numberSelectedSupplierFilter, numberDefaultListSupplierFilter, "ResetFilter supprier ''");

            value = interimReceptionsPage.GetFilterValue(InterimReceptionsPage.FilterType.ShowAll);
            Assert.AreEqual(false, value, "ResetFilter ShowAll");
            value = interimReceptionsPage.GetFilterValue(InterimReceptionsPage.FilterType.ShowOnlyActive);
            Assert.AreEqual(true, value, "ResetFilter ShowOnlyActive");
            value = interimReceptionsPage.GetFilterValue(InterimReceptionsPage.FilterType.ShowOnlyInactive);
            Assert.AreEqual(false, value, "ResetFilter ShowOnlyInactive");
            value = interimReceptionsPage.GetFilterValue(InterimReceptionsPage.FilterType.ShowAllReceptions);
            Assert.AreEqual(true, value, "ResetFilter ShowAll");
            value = interimReceptionsPage.GetFilterValue(InterimReceptionsPage.FilterType.ShowNotValidatedOnly);
            Assert.AreEqual(false, value, "ResetFilter ShowOnlyActive");
            value = interimReceptionsPage.GetFilterValue(InterimReceptionsPage.FilterType.ShowValidatedOnly);
            Assert.AreEqual(false, value, "ResetFilter ShowOnlyInactive");
            value = interimReceptionsPage.GetFilterValue(InterimReceptionsPage.FilterType.ShowAllInterimReception);
            Assert.AreEqual(true, value, "ResetFilter ShowOnlyInactive");
            value = interimReceptionsPage.GetFilterValue(InterimReceptionsPage.FilterType.ShowOpenedInterimReception);
            Assert.AreEqual(false, value, "ResetFilter ShowOnlyInactive");
            value = interimReceptionsPage.GetFilterValue(InterimReceptionsPage.FilterType.ShowClosedInterimReception);
            Assert.AreEqual(false, value, "ResetFilter ShowOnlyInactive");

        }

        [Timeout(_timeout)]
		[TestMethod]
        public void INT_RECEP_index_Export()
        { //Prepare 
            string site = TestContext.Properties["SiteACE"].ToString();
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string workshop = TestContext.Properties["Workshop1"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string location = "Produccion";
            string packagingName = TestContext.Properties["Item_PackagingName"].ToString();
            string storageQty = 10.ToString();
            string qty = 10.ToString();

            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            bool newVersion = true;
            // Arrange

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var receptionsPage = homePage.GoToInterim_Receptions();
            DeleteAllFileDownload();
            if (receptionsPage.CheckTotalNumberInterim() == 0)
            {
                InterimReceptionsCreateModalPage modalcreateInterim = receptionsPage.CreateNewInterim();
                modalcreateInterim.FillField_CreatNewIntermin(DateTime.Now, site, supplier, location);
                var InterimId = modalcreateInterim.GetInterimId();
                InterimReceptionsItem InterimItem = modalcreateInterim.Submit();
                InterimItem.BackToList();
            }
            string dateString = receptionsPage.GetFirstInterimReceptionsDate();
            DateTime startdateInput = DateTime.ParseExact(dateString, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);

            receptionsPage.Filter(InterimReceptionsPage.FilterType.DateFrom, startdateInput);
            receptionsPage.Filter(InterimReceptionsPage.FilterType.DateTo, startdateInput.AddMonths(1));
            string InterimreceptionFirstNumber = receptionsPage.GetFirstInterimReceptionsNumber();
            receptionsPage.ExportInterimReceptions(InterimReceptionsPage.ExportType.Export, newVersion);
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();
            var correctDownloadedFile = receptionsPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);
            var fileName = correctDownloadedFile.Name;
            var filePath = System.IO.Path.Combine(downloadsPath, fileName);
            int resultNumber = OpenXmlExcel.GetExportResultNumber("Interim receptions", filePath);
            var listResult = OpenXmlExcel.GetValuesInList("Number", "Interim receptions", filePath);
            //Assert
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.IsFalse(!listResult.Contains(InterimreceptionFirstNumber), MessageErreur.EXCEL_DONNEES_KO);
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_RECEP_index_filter_sites()
        {
            string site = "ACE";
            var homePage = LogInAsAdmin();
            InterimReceptionsPage interimPage = homePage.GoToInterim_Receptions();
            interimPage.ResetFilters();
            if (!interimPage.isPageSizeEqualsTo100())
            {
                interimPage.PageSize("8");
                interimPage.PageSize("100");
            }
            interimPage.Filter(InterimReceptionsPage.FilterType.Sites, site);
            bool IsSortedBySites = interimPage.IsSortedBySites();
            // Assert
            Assert.IsTrue(IsSortedBySites, MessageErreur.FILTRE_ERRONE, "Sort by site");
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_RECEP_index_filter_SortByNumber()
        {
            string sortbynumber = "NUMBER";
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            InterimReceptionsPage interimPage = homePage.GoToInterim_Receptions();
            interimPage.ResetFilters();
            interimPage.Filter(InterimReceptionsPage.FilterType.SortBy, sortbynumber);
            Assert.IsTrue(interimPage.IsSortedByNumber(), MessageErreur.FILTRE_ERRONE, "Sort by Number");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void INT_RECEP_index_filter_SortByDeliveryDate()
        {
            string sortbydeliverydate = "DATE";
            var homePage = LogInAsAdmin();
            string dateFormat = homePage.GetDateFormatPickerValue();
            InterimReceptionsPage interimPage = homePage.GoToInterim_Receptions();
            interimPage.ResetFilters();
            interimPage.Filter(InterimReceptionsPage.FilterType.SortBy, sortbydeliverydate);
            Assert.IsTrue(interimPage.IsSortedByDeliveryDate(dateFormat), MessageErreur.FILTRE_ERRONE, "Sort by Delivery Date");

        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_RECEP_index_filter_ByFromToDate()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            string dateFormat = homePage.GetDateFormatPickerValue();
            var receptionsPage = homePage.GoToInterim_Receptions();
            receptionsPage.PageSize("50");
            receptionsPage.ResetFilters();
            receptionsPage.Filter(InterimReceptionsPage.FilterType.DateFrom, DateTime.Now.AddDays(-10));
            receptionsPage.Filter(InterimReceptionsPage.FilterType.DateTo, DateTime.Now);
            Assert.IsTrue(receptionsPage.IsFromToDateRespected(DateTime.Now.AddDays(-10), DateTime.Now, dateFormat), String.Format(MessageErreur.FILTRE_ERRONE, "'From/To'"));
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_RECEP_index_filter_SearchByNumber()
        {
            //Login
            var homePage = LogInAsAdmin();
            var receptionsPage = homePage.GoToInterim_Receptions();
            string interimreceptionsNumber = receptionsPage.GetFirstInterimReceptionsNumber();
            receptionsPage.ResetFilters();
            receptionsPage.Filter(InterimReceptionsPage.FilterType.Bynumber, interimreceptionsNumber);
            bool numberFound = receptionsPage.GetAllInterinReceptionPaged().Any(item => item.Contains(interimreceptionsNumber));
            //Asssert : is Found
            Assert.IsTrue(numberFound, "Les résultats ne sont pas mis à jour en fonction de filtre number.");
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_RECEP_index_filter_Supplier()
        {
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            InterimReceptionsPage interimPage = homePage.GoToInterim_Receptions();
            interimPage.ResetFilters();
            interimPage.Filter(InterimReceptionsPage.FilterType.Suppliers, supplier);
            // Assert
            Assert.IsTrue(interimPage.IsSortedBySuppliers(), MessageErreur.FILTRE_ERRONE, "Sort by suppliers");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void INT_RECEP_index_ViewReception()
        {
            var homePage = LogInAsAdmin();
            var receptionsPage = homePage.GoToInterim_Receptions();
            var InterimId = receptionsPage.GetFirstID();
            var InterimItem = receptionsPage.ClickFirstLine();
            var receptionsPageDetail = InterimItem.GoToGeneralInformation();
            // Verify if the Detailpage is displayed
            Assert.AreEqual(receptionsPageDetail.GetInterimId(), InterimId, "Page Introuvable");

        }

        [Timeout(_timeout)]
		[TestMethod]
        public void INT_RECEP_index_Total()
        {
            // Arrange
            var homePage = LogInAsAdmin();
            InterimReceptionsPage interimReceptionsPage = homePage.GoToInterim_Receptions();
            interimReceptionsPage.ResetFilters();
            interimReceptionsPage.Filter(InterimReceptionsPage.FilterType.ShowValidatedOnly, true);
            interimReceptionsPage.PageSize("16");
            //valeurs 
            var interimRecptionTotalCount = interimReceptionsPage.CheckTotalNumber().ToString();
            var interimRecptionCounter = interimReceptionsPage.InterimRecptionCounter();

            Assert.AreEqual(interimRecptionTotalCount, interimRecptionCounter, "Le total de l'index n'est pas égal.");
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_RECEP_index_filter_ShowAll()
        {
            var homePage = LogInAsAdmin();
            InterimReceptionsPage interimreceptionsPage = homePage.GoToInterim_Receptions();
            interimreceptionsPage.ResetFilters();
            interimreceptionsPage.Filter(InterimReceptionsPage.FilterType.ShowOnlyActive, true);
            var resultsActivated = interimreceptionsPage.CheckTotalNumber();
            interimreceptionsPage.Filter(InterimReceptionsPage.FilterType.ShowOnlyInactive, true);
            var resultsInactivated = interimreceptionsPage.CheckTotalNumber();
            interimreceptionsPage.Filter(InterimReceptionsPage.FilterType.ShowAll, true);
            var resultAll = interimreceptionsPage.CheckTotalNumber();
            Assert.AreEqual(resultsActivated + resultsInactivated, resultAll, "Les résultats ne sont pas mis à jour en fonction des filtres Show All");
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_RECEP_index_filter_ShowActiveOnly()
        {
            var homePage = LogInAsAdmin();
            InterimReceptionsPage interimreceptionsPage = homePage.GoToInterim_Receptions();
            interimreceptionsPage.ResetFilters();
            interimreceptionsPage.Filter(InterimReceptionsPage.FilterType.ShowOnlyActive, true);
            var resultsActivated = interimreceptionsPage.CheckTotalNumber();
            interimreceptionsPage.Filter(InterimReceptionsPage.FilterType.ShowOnlyInactive, true);
            var resultsInactivated = interimreceptionsPage.CheckTotalNumber();
            interimreceptionsPage.Filter(InterimReceptionsPage.FilterType.ShowAll, true);
            var resultAll = interimreceptionsPage.CheckTotalNumber();
            //Assert
            Assert.AreEqual(resultAll - resultsInactivated, resultsActivated, "Les résultats ne sont pas mis à jour en fonction des filtres Show active only");
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_RECEP_Details_filter_Group()
        {
            //Prepare
            string supplierInterim = TestContext.Properties["SupplierForInterim"].ToString();
            string group = "A REFERENCIA";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            var receptionsPage = homePage.GoToInterim_Receptions();
            receptionsPage.Filter(InterimReceptionsPage.FilterType.Suppliers, supplierInterim);
            InterimReceptionsItem interimReceptionsItem = receptionsPage.GoToInterimReceptionItem();
            interimReceptionsItem.Filter(InterimReceptionsItem.FilterType.Groups, group);

            //assert
            Assert.IsTrue(interimReceptionsItem.RowsNumber() >= 1, "Les lignes ayant un groupe correspondant à la recherche n'apparaitre pas.");
        }


        [Timeout(_timeout)]
		[TestMethod]
        public void INT_RECEP_index_AddNewComment()
        {
            //Prepare
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            string location = "Produccion";
            string comment = "This is my comment";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            InterimReceptionsPage interimPage = homePage.GoToInterim_Receptions();
            InterimReceptionsCreateModalPage modalcreateInterim = interimPage.CreateNewInterim();
            modalcreateInterim.FillField_CreatNewInterminAddComment(DateTime.Now, site, supplier, location, comment);
            InterimReceptionsItem InterimItem = modalcreateInterim.Submit();
            InterimReceptionsGeneralInformation interimReceptionsGeneralInformation = InterimItem.GoToGeneralInformation();
            string savedComment = interimReceptionsGeneralInformation.GetComment();
            //Assert
            Assert.AreEqual(comment, savedComment, "Le commentaire enregistré ne correspond pas au commentaire initial.");

        }

        [Timeout(_timeout)]
		[TestMethod]
        public void INT_RECEP_index_filter_ShowNotValidatedOnly()
        {
            var homePage = LogInAsAdmin();
            var receptionsPage = homePage.GoToInterim_ReceptionsModified();
            receptionsPage.PageSize("100");
            receptionsPage.Filter(InterimReceptionsPage.FilterType.ShowNotValidatedOnly, true);
            Assert.IsFalse(receptionsPage.CheckValidation(false), String.Format(MessageErreur.FILTRE_ERRONE, "'Show not validated only'"));
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void INT_RECEP_index_filter_ShowValidatedOnly()
        {
            var homePage = LogInAsAdmin();
            var receptionsPage = homePage.GoToInterim_ReceptionsModified();
            receptionsPage.PageSize("100");
            receptionsPage.Filter(InterimReceptionsPage.FilterType.ShowValidatedOnly, true);
            Assert.IsTrue(receptionsPage.CheckValidation(true), String.Format(MessageErreur.FILTRE_ERRONE, "'Show validated only'"));
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void INT_RECEP_Details_filter_Keyword()
        {
            //Prepare
            string keyword = TestContext.Properties["Item_Keyword"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            var interimReceptionPage = homePage.GoToInterim_Receptions();
            interimReceptionPage.ResetFilters();
            InterimReceptionsItem interimReceptionsItem = interimReceptionPage.GoToInterimReceptionItem();
            var itemName = interimReceptionsItem.GetItemName();

            var purchasingItemPage = homePage.GoToPurchasing_ItemPage();
            purchasingItemPage.Filter(ItemPage.FilterType.Search, itemName);
            ItemGeneralInformationPage generalInfo = purchasingItemPage.ClickOnFirstItem();

            ItemKeywordPage itemKeywordPage = generalInfo.ClickOnKeywordItem();
            try
            {
                if (!itemKeywordPage.IsKeywordAdded(keyword))
                {
                    itemKeywordPage.AddKeyword(keyword);
                }
                interimReceptionPage = homePage.GoToInterim_Receptions();
                interimReceptionPage.ResetFilters();
                interimReceptionsItem = interimReceptionPage.GoToInterimReceptionItem();
                var rowsItem = interimReceptionsItem.GetNombreRowItem();
                interimReceptionsItem.Filter(InterimReceptionsItem.FilterType.Keyword, keyword);
                var rowsItemAfterFilter = interimReceptionsItem.GetNombreRowItem();
                Assert.AreEqual(rowsItem, rowsItemAfterFilter, "Les résultats ne sont pas mis à jour en fonction des filtres KEYWORDS");
            }
            finally
            {
                purchasingItemPage = homePage.GoToPurchasing_ItemPage();
                purchasingItemPage.Filter(ItemPage.FilterType.Search, itemName);
                ItemGeneralInformationPage generalInfos = purchasingItemPage.ClickOnFirstItem();

                itemKeywordPage = generalInfo.ClickOnKeywordItem();
                itemKeywordPage.RemoveKeyword(keyword);
            }
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void INT_RECEP_index_filter_Opened()
        {
            //Login
            var homePage = LogInAsAdmin();
            InterimReceptionsPage interimReceptionsPage = homePage.GoToInterim_Receptions();
            interimReceptionsPage.ResetFilters();
            interimReceptionsPage.Filter(InterimReceptionsPage.FilterType.ShowOpenedInterimReception, true);
            var resultsOpened = interimReceptionsPage.CheckTotalNumber();
            interimReceptionsPage.Filter(InterimReceptionsPage.FilterType.ShowClosedInterimReception, true);
            var resultsClosed = interimReceptionsPage.CheckTotalNumber();
            interimReceptionsPage.Filter(InterimReceptionsPage.FilterType.ShowAllInterimReception, true);
            var resultAll = interimReceptionsPage.CheckTotalNumber();
            //assert : opened = all - closed
            Assert.AreEqual(resultAll - resultsClosed, resultsOpened, "Les résultats ne sont pas mis à jour en fonction des filtres Status Opened");

        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_RECEP_index_filter_closed()
        {
            //login
            var homePage = LogInAsAdmin();
            InterimReceptionsPage interimReceptionsPage = homePage.GoToInterim_Receptions();
            interimReceptionsPage.ResetFilters();
            interimReceptionsPage.Filter(InterimReceptionsPage.FilterType.ShowOpenedInterimReception, true);
            var resultsOpened = interimReceptionsPage.CheckTotalNumber();
            interimReceptionsPage.Filter(InterimReceptionsPage.FilterType.ShowClosedInterimReception, true);
            var resultsClosed = interimReceptionsPage.CheckTotalNumber();
            interimReceptionsPage.Filter(InterimReceptionsPage.FilterType.ShowAllInterimReception, true);
            var resultAll = interimReceptionsPage.CheckTotalNumber();
            //Assert : closed = all - opened
            Assert.AreEqual(resultAll - resultsOpened, resultsClosed, "Les résultats ne sont pas mis à jour en fonction des filtres Status closed");
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_RECEP_index_filter_ShowAllReceptions()
        {
            var homePage = LogInAsAdmin();
            InterimReceptionsPage interimreceptionsPage = homePage.GoToInterim_Receptions();
            interimreceptionsPage.ResetFilters();
            interimreceptionsPage.Filter(InterimReceptionsPage.FilterType.ShowValidatedOnly, true);
            var resultsValidated = interimreceptionsPage.CheckTotalNumber();
            interimreceptionsPage.Filter(InterimReceptionsPage.FilterType.ShowNotValidatedOnly, true);
            var resultsInvalidated = interimreceptionsPage.CheckTotalNumber();
            interimreceptionsPage.Filter(InterimReceptionsPage.FilterType.ShowAllReceptions, true);
            var resultAll = interimreceptionsPage.CheckTotalNumber();
            Assert.AreEqual(resultsValidated + resultsInvalidated, resultAll, "Les résultats ne sont pas mis à jour en fonction des filtres Show All");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void INT_RECEP_index_filter_ALL()
        {
            //login
            var homePage = LogInAsAdmin();
            InterimReceptionsPage interimReceptionsPage = homePage.GoToInterim_Receptions();
            interimReceptionsPage.ResetFilters();
            interimReceptionsPage.Filter(InterimReceptionsPage.FilterType.ShowOpenedInterimReception, true);
            var resultsOpened = interimReceptionsPage.CheckTotalNumber();
            interimReceptionsPage.Filter(InterimReceptionsPage.FilterType.ShowClosedInterimReception, true);
            var resultsClosed = interimReceptionsPage.CheckTotalNumber();
            interimReceptionsPage.Filter(InterimReceptionsPage.FilterType.ShowAllInterimReception, true);
            var resultAll = interimReceptionsPage.CheckTotalNumber();
            //Assert : All = closed + open 
            Assert.AreEqual(resultsOpened + resultsClosed, resultAll, "Les résultats ne sont pas mis à jour en fonction des filtres Status All");
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_RECEP_index_filter_ResetFilter()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            // Act
            InterimReceptionsPage interimreceptionsPage = homePage.GoToInterim_Receptions();
            var defaultShowreceptions = interimreceptionsPage.GetShowFilterSelected();
            var defaultSuppliersSelected = interimreceptionsPage.GetSelectedSuppliersToFilter();
            var defaulshowActiveSelected = interimreceptionsPage.GetShowFilterActiveSelected();
            var defaultStatusDisplayed = interimreceptionsPage.GetStatusFilterSelected();
            interimreceptionsPage.ScrollUntilSitesFilterIsVisible();
            var defaultSitesDisplayed = interimreceptionsPage.GetSelectedSitesToFilter();
            interimreceptionsPage.Filter(InterimReceptionsPage.FilterType.Suppliers, "MonSupplier8");
            interimreceptionsPage.Filter(InterimReceptionsPage.FilterType.ShowValidatedOnly, true);
            interimreceptionsPage.Filter(InterimReceptionsPage.FilterType.ShowAll, true);
            interimreceptionsPage.Filter(InterimReceptionsPage.FilterType.Opened, true);
            interimreceptionsPage.Filter(InterimReceptionsPage.FilterType.Sites, "ACE - ACE");
            interimreceptionsPage.PageUp();
            interimreceptionsPage.ResetFilters();
            var resultShowreceptionsAfterReset = interimreceptionsPage.GetShowFilterSelected();
            var resultSuppliersSelectedAfterReset = interimreceptionsPage.GetSelectedSuppliersToFilter();
            var resultshowActiveSelectedAfterReset = interimreceptionsPage.GetShowFilterActiveSelected();
            var resultStatusDisplayedAfterReset = interimreceptionsPage.GetStatusFilterSelected();
            interimreceptionsPage.ScrollUntilSitesFilterIsVisible();
            var resultSitesDisplayedAfterReset = interimreceptionsPage.GetSelectedSitesToFilter();
            Assert.AreEqual(defaultShowreceptions, resultShowreceptionsAfterReset, "Le filter Show receptions ne remet pas");
            Assert.AreEqual(defaultSuppliersSelected.Count, resultSuppliersSelectedAfterReset.Count, "Le filter SUPPLIERS ne remet pas");
            Assert.AreEqual(defaulshowActiveSelected, resultshowActiveSelectedAfterReset, "Le filter Show all ne remet pas");
            Assert.AreEqual(defaultStatusDisplayed, resultStatusDisplayedAfterReset, "Le filter Show STATUS ne remet pas");
            Assert.AreEqual(defaultSitesDisplayed.Count, resultSitesDisplayedAfterReset.Count, "Le filter SITES ne remet pas");
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_RECEP_index_ExportMoreThan1Month()
        {
            DateTime startdateInput = DateUtils.Now.AddDays(-91);
            DateTime enddateInput = DateUtils.Now.AddDays(-30);
            var homePage = LogInAsAdmin();
            var receptionsPage = homePage.GoToInterim_Receptions();
            receptionsPage.PageSize("8");
            DeleteAllFileDownload();
            receptionsPage.ClearDownloads();
            receptionsPage.Filter(InterimReceptionsPage.FilterType.DateFrom, startdateInput);
            receptionsPage.Filter(InterimReceptionsPage.FilterType.DateTo, enddateInput);
            string Interimreceptions = receptionsPage.GetFirstInterimReceptionsNumber();
            receptionsPage.Export();
            var errorMessage = receptionsPage.GetExportErrorMsg();
            Assert.AreEqual("Please choose only one month.", errorMessage, "The expected error message did not appear.");
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_RECEP_index_PrintWithoutPrices()
        {
            DateTime startdateInput = DateUtils.Now.AddDays(-5);
            DateTime enddateInput = DateUtils.Now.AddDays(10);
            bool versionprint = true;
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNamePdfBegin = "Interim Reception Report_-_";
            string DocFileNameZipBegin = "All_files_";
            string status = string.Empty;
            string referenceitem = string.Empty;
            string dateitem = string.Empty;
            string packagingitem = string.Empty;
            string orderqtyitem = string.Empty;
            string deliveryitem = string.Empty;
            string totalvatitem = string.Empty;
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            InterimReceptionsPage receptionsPage = homePage.GoToInterim_Receptions();
            receptionsPage.ResetFilters();
            receptionsPage.Filter(InterimReceptionsPage.FilterType.DateFrom, startdateInput);
            receptionsPage.Filter(InterimReceptionsPage.FilterType.DateTo, enddateInput);
            receptionsPage.Filter(InterimReceptionsPage.FilterType.Showvalidatedonly, true);
            if (receptionsPage.CheckTotalNumber() > 0)
            {
                receptionsPage.ClearDownloads();
                var InterimrecipId = receptionsPage.GetFirstID();
                receptionsPage.Filter(InterimReceptionsPage.FilterType.Bynumber, InterimrecipId);
                //Print Report
                var reportPage = receptionsPage.PrintReport(versionprint);
                var isGenerated = reportPage.IsReportGenerated();
                Assert.IsTrue(isGenerated, "Interim Receptions n'a pas été générer.");
                reportPage.Close();
                //download pdf file
                reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
                string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
                FileInfo fichier = new FileInfo(trouve);
                fichier.Refresh();
                //Assert.IsTrue(filePdfNoInclude.Exists, trouveNoInclude + " non généré");
                PdfDocument document = PdfDocument.Open(fichier.FullName);
                List<string> mots = new List<string>();
                var nbpage = document.GetPages();
                foreach (Page p in document.GetPages())
                {
                    mots.AddRange(p.GetWords().Select(m => m.Text));
                }
                string ch = String.Join(" ", mots);
                receptionsPage.ResetFilters();
                receptionsPage.Filter(InterimReceptionsPage.FilterType.Bynumber, InterimrecipId);
                var InterimrecipItem = receptionsPage.ClickFirstLine();
                var nbrowitem = InterimrecipItem.GetNombreRowItem();
                if (nbrowitem > 1)
                {
                    for (int i = 0; i < nbrowitem / 2; i++)
                    {
                        referenceitem = InterimrecipItem.GetReference(i);
                        dateitem = InterimrecipItem.GetDateItem(i);
                        packagingitem = InterimrecipItem.GetPackaging(i);
                        orderqtyitem = InterimrecipItem.GetOrderedQty(i);
                        deliveryitem = InterimrecipItem.GetDeliveredQty(i);
                        totalvatitem = InterimrecipItem.GetTotalVAT(i);

                        if (!string.IsNullOrEmpty(referenceitem))
                        {
                            Assert.IsTrue(ch.Replace(" ", "").Contains(referenceitem.Replace(" ", "")), " information Referenceitem n est pas présents dans le PDF");
                        }
                        if (!string.IsNullOrEmpty(dateitem))
                        {
                            DateTime date = DateTime.ParseExact(dateitem, "yyyy-MM-dd", null);
                            string formattedDate = date.ToString("dd/MM/yyyy");
                            Assert.IsTrue(mots.Contains(formattedDate), " information Date Item n est pas présents dans le PDF");
                        }
                        if (!string.IsNullOrEmpty(packagingitem))
                        {
                            Assert.IsTrue(ch.Contains(packagingitem), " information Packaging Item n est pas présents dans le PDF");
                        }
                        if (!string.IsNullOrEmpty(orderqtyitem))
                        {
                            Assert.IsTrue(mots.Contains(orderqtyitem), " information Order qty Item n est pas présents dans le PDF");
                        }
                        else
                        {
                            orderqtyitem = "0";
                            Assert.IsTrue(mots.Contains(orderqtyitem), " information Order qty Item n est pas présents dans le PDF");
                        }
                        if (!string.IsNullOrEmpty(deliveryitem))
                        {
                            Assert.IsTrue(mots.Contains(deliveryitem), " information Delivery Item n est pas présents dans le PDF");
                        }
                    }
                }
                var receptionsPageDetail = InterimrecipItem.GoToGeneralInformation();
                var interimrecipId = receptionsPageDetail.GetInterimId();
                var site = receptionsPageDetail.GetSite();
                var supplier = receptionsPageDetail.GetSupplier();
                var place = receptionsPageDetail.GetDelivery_location();
                var deliveryordernumber = receptionsPageDetail.GetDelivery_order_Number();
                var deliverydate = receptionsPageDetail.GetDelivery_Date();
                var validatedby = receptionsPageDetail.GetValidated_By();
                receptionsPage = InterimrecipItem.BackToList();
                var settingsSitesPage = homePage.GoToParameters_Sites();
                settingsSitesPage.Filter(ParametersSites.FilterType.SearchSite, site);
                settingsSitesPage.ClickOnFirstSite();
                settingsSitesPage.ClickToInformations();
                var adress = settingsSitesPage.GetAddress();
                var adress2 = settingsSitesPage.GetAddress2();
                var city = settingsSitesPage.GetCity();
                Assert.IsTrue(mots.Contains(InterimrecipId), " Number Iterim Reception n est pas présents dans le PDF");
                if (!string.IsNullOrEmpty(site))
                {
                    site = site + "(" + site + ")";
                    Assert.IsTrue(mots.Contains(site), " information Site n est pas présents dans le PDF");
                }
                Assert.IsTrue(mots.Contains(place), " information Place n est pas présents dans le PDF");
                Assert.IsTrue(mots.Contains(supplier), " information Supplier n est pas présents dans le PDF");
                Assert.IsTrue(ch.Contains(adress), " information Adresse n est pas présents dans le PDF");
                Assert.IsTrue(ch.Contains(adress2), " information Adresse n est pas présents dans le PDF");
                Assert.IsTrue(ch.Contains(city), " information City n est pas présents dans le PDF");
                Assert.IsTrue(mots.Contains(deliverydate), " information delivery date n est pas présents dans le PDF");
                if (string.IsNullOrEmpty(validatedby))
                {
                    status = "Not validated";
                    Assert.IsTrue(ch.Contains(status), " information Status n est pas présents dans le PDF");
                }
                else
                {
                    status = "Validated by " + validatedby;
                    Assert.IsTrue(ch.Replace(" ", "").Contains(status.Replace(" ", "")), " information Status n est pas présents dans le PDF");
                }
                Assert.IsTrue(ch.Contains(deliveryordernumber), " information delivery orde rnumber n est pas présents dans le PDF");
            }
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_RECEP_Details_filter_SearchByNameRef()
        {
            //Prepare 
            string notFoundItem = "Testifthefilterworks";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            InterimReceptionsPage interimreceptionsPagee = homePage.GoToInterim_Receptions();
            interimreceptionsPagee.Filter(InterimReceptionsPage.FilterType.ShowValidatedOnly, true);
            string interimreceptionsNumber = interimreceptionsPagee.GetFirstInterimReceptionsNumber();
            InterimReceptionsItem interimReceptionsItem = interimreceptionsPagee.GoToInterimReceptionItem();
            string interimReceptionsNameRef = interimReceptionsItem.GetFirstInterimReceptionsNameRef();
            interimReceptionsItem.Filter(InterimReceptionsItem.FilterType.Name, interimReceptionsNameRef);
            var ordersList = interimReceptionsItem.GetItemsForInterimFiltred();
            // Assert that the interimReceptionsNameRef is found in the filtered list
            Assert.IsTrue(ordersList != null && ordersList.Contains(interimReceptionsNameRef), "The filtered list does not contain the expected name reference.");
            // Assert that no items are found with the 'Testifthefilterworks' to ensure the filter is working correctly. 
            interimReceptionsItem.Filter(InterimReceptionsItem.FilterType.Name, notFoundItem);
            var ordersListNotfound = interimReceptionsItem.GetItemsForInterimFiltred();
            Assert.IsTrue(ordersListNotfound == null || ordersListNotfound.Count == 0, "The filtered list should be empty for the 'Testifthefilterworks' value.");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void INT_RECEP_Index_PrintWithPrices()
        {
            DateTime startdateInput = DateUtils.Now.AddDays(-5);
            DateTime enddateInput = DateUtils.Now.AddDays(10);
            bool versionprint = true;
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNamePdfBegin = "Interim Reception Report_-_";
            string DocFileNameZipBegin = "All_files_";
            string status = string.Empty;
            string referenceitem = string.Empty;
            string dateitem = string.Empty;
            string packagingitem = string.Empty;
            string packagingitemprice = string.Empty;
            string orderqtyitem = string.Empty;
            string deliveryitem = string.Empty;
            string totalvatitem = string.Empty;
            string totalWOvatitem = string.Empty;
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            InterimReceptionsPage receptionsPage = homePage.GoToInterim_Receptions();
            receptionsPage.ResetFilters();
            receptionsPage.Filter(InterimReceptionsPage.FilterType.DateFrom, startdateInput);
            receptionsPage.Filter(InterimReceptionsPage.FilterType.DateTo, enddateInput);
            receptionsPage.Filter(InterimReceptionsPage.FilterType.Showvalidatedonly, true);
            if (receptionsPage.CheckTotalNumber() > 0)
            {
                receptionsPage.ClearDownloads();
                var InterimrecipId = receptionsPage.GetFirstID();
                receptionsPage.Filter(InterimReceptionsPage.FilterType.Bynumber, InterimrecipId);
                //Print Report
                var reportPage = receptionsPage.PrintReport(versionprint, true);
                var isGenerated = reportPage.IsReportGenerated();
                Assert.IsTrue(isGenerated, "Interim Receptions Print n'été pas générer.");
                reportPage.Close();
                //download pdf file
                reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
                string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
                FileInfo fichier = new FileInfo(trouve);
                fichier.Refresh();
                //Assert.IsTrue(filePdfNoInclude.Exists, trouveNoInclude + " non généré");
                PdfDocument document = PdfDocument.Open(fichier.FullName);
                List<string> mots = new List<string>();
                var nbpage = document.GetPages();
                foreach (Page p in document.GetPages())
                {
                    mots.AddRange(p.GetWords().Select(m => m.Text));
                }
                string ch = String.Join(" ", mots);
                receptionsPage.ResetFilters();
                receptionsPage.Filter(InterimReceptionsPage.FilterType.Bynumber, InterimrecipId);
                var InterimrecipItem = receptionsPage.ClickFirstLine();
                var nbrowitem = InterimrecipItem.GetNombreRowItem();
                if (nbrowitem > 1)
                {
                    for (int i = 0; i < nbrowitem / 2; i++)
                    {
                        referenceitem = InterimrecipItem.GetReference(i);
                        dateitem = InterimrecipItem.GetDateItem(i);
                        packagingitem = InterimrecipItem.GetPackaging(i);
                        packagingitemprice = InterimrecipItem.GetPackagingPrice(i);
                        orderqtyitem = InterimrecipItem.GetOrderedQty(i);
                        deliveryitem = InterimrecipItem.GetDeliveredQty(i);
                        totalvatitem = InterimrecipItem.GetTotalVAT(i);
                        if (!string.IsNullOrEmpty(referenceitem))
                        {
                            Assert.IsTrue(mots.Contains(referenceitem) || ch.Replace(" ", "").Contains(referenceitem.Replace(" ", "")), " information Referenceitem n est pas présents dans le PDF");
                        }
                        if (!string.IsNullOrEmpty(dateitem))
                        {
                            DateTime date = DateTime.ParseExact(dateitem, "yyyy-MM-dd", null);
                            string formattedDate = date.ToString("dd/MM/yyyy");
                            Assert.IsTrue(mots.Contains(formattedDate), " information Date Item n est pas présents dans le PDF");
                        }
                        if (!string.IsNullOrEmpty(packagingitem))
                        {
                            Assert.IsTrue(ch.Contains(packagingitem), " information Packaging Item n est pas présents dans le PDF");
                        }
                        if (!string.IsNullOrEmpty(packagingitemprice))
                        {
                            if (packagingitemprice.Equals("Free")) packagingitemprice = "€ 0,0000";

                            Assert.IsTrue(ch.Contains(packagingitemprice), " information Packaging Item n est pas présents dans le PDF");
                        }
                        if (!string.IsNullOrEmpty(orderqtyitem))
                        {
                            Assert.IsTrue(mots.Contains(orderqtyitem), " information Order qty Item n est pas présents dans le PDF");
                        }
                        else
                        {
                            orderqtyitem = "0";
                            Assert.IsTrue(mots.Contains(orderqtyitem), " information Order qty Item n est pas présents dans le PDF");
                        }
                        if (!string.IsNullOrEmpty(deliveryitem))
                        {
                            Assert.IsTrue(mots.Contains(deliveryitem), " information Delivery Item n est pas présents dans le PDF");
                        }
                        if (!string.IsNullOrEmpty(totalWOvatitem))
                        {
                            Assert.IsTrue(mots.Contains(deliveryitem), " information Delivery Item n est pas présents dans le PDF");
                        }
                        Assert.IsTrue(ch.Replace(" ", "").Contains("VATRate"), " information VAT Rate n est pas présents dans le PDF");
                    }
                    totalWOvatitem = InterimrecipItem.GetTotalWOVAT();
                    Assert.IsTrue(ch.Contains(totalWOvatitem), " information Total w/o VAT n est pas présents dans le PDF");
                }
                var receptionsPageDetail = InterimrecipItem.GoToGeneralInformation();
                var interimrecipId = receptionsPageDetail.GetInterimId();
                var site = receptionsPageDetail.GetSite();
                var supplier = receptionsPageDetail.GetSupplier();
                var place = receptionsPageDetail.GetDelivery_location();
                var deliveryordernumber = receptionsPageDetail.GetDelivery_order_Number();
                var deliverydate = receptionsPageDetail.GetDelivery_Date();
                var validatedby = receptionsPageDetail.GetValidated_By();
                receptionsPage = InterimrecipItem.BackToList();
                var settingsSitesPage = homePage.GoToParameters_Sites();
                settingsSitesPage.Filter(ParametersSites.FilterType.SearchSite, site);
                settingsSitesPage.ClickOnFirstSite();
                settingsSitesPage.ClickToInformations();
                var adress = settingsSitesPage.GetAddress();
                var adress2 = settingsSitesPage.GetAddress2();
                var city = settingsSitesPage.GetCity();
                Assert.IsTrue(mots.Contains(InterimrecipId), " Number Iterim Reception n est pas présents dans le PDF");
                if (!string.IsNullOrEmpty(site))
                {
                    site = site + "(" + site + ")";
                    Assert.IsTrue(mots.Contains(site), " information Site n est pas présents dans le PDF");
                }
                Assert.IsTrue(mots.Contains(place), " information Place n est pas présents dans le PDF");
                Assert.IsTrue(mots.Contains(supplier), " information Supplier n est pas présents dans le PDF");
                Assert.IsTrue(ch.Contains(adress), " information Adresse n est pas présents dans le PDF");
                Assert.IsTrue(ch.Contains(adress2), " information Adresse n est pas présents dans le PDF");
                Assert.IsTrue(ch.Contains(city), " information City n est pas présents dans le PDF");
                Assert.IsTrue(mots.Contains(deliverydate), " information delivery date n est pas présents dans le PDF");

                if (string.IsNullOrEmpty(validatedby))
                {
                    status = "Not validated";
                    Assert.IsTrue(ch.Contains(status), " information Status n est pas présents dans le PDF");
                }
                else
                {
                    status = "Validated by " + validatedby;
                    Assert.IsTrue(ch.Replace(" ", "").Contains(status.Replace(" ", "")), " information Status n est pas présents dans le PDF");
                }
                Assert.IsTrue(ch.Contains(deliveryordernumber), " information delivery orde rnumber n est pas présents dans le PDF");
            }
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_RECEP_details_ItemEdit()
        {
            //Prepare
            string qty = new Random().Next(100, 5000).ToString();
            
            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            InterimReceptionsPage interimreceptionsPage = homePage.GoToInterim_Receptions();
            interimreceptionsPage.Filter(InterimReceptionsPage.FilterType.Shownotvalidatedonly, true);
            InterimReceptionsItem interimItem = interimreceptionsPage.ClickFirstlineinterimNotValidated();
            interimItem.ClickOnItem();
            interimItem.SetQty(qty);

            //Assert
            bool isIconVisible = interimItem.IsVisible();
            Assert.IsTrue(isIconVisible, "la quantité n'est pas été modifié");
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_RECEP_Details_filter_SubGroup()
        {
            //Prepare 
            string subgrpname = "subgrpname";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            var receptionsPage = homePage.GoToInterim_ReceptionsModified();

            receptionsPage.PageSize("100");
            var InterimRecptionCounter = receptionsPage.InterimRecptionCounter();
            Random random = new Random();
            int randomIndex = random.Next(1, Int32.Parse(InterimRecptionCounter));
            var InterimReceptionsItem = receptionsPage.SelectInterimReceptionsItem(1);
            var ItemGeneralInforamtion = InterimReceptionsItem.EditItem();
            InterimReceptionsItem.Go_To_New_Navigate();
            ItemGeneralInforamtion.SetSubgroupName(subgrpname);
            var name = ItemGeneralInforamtion.GetItemNameFromEditMenu();
            InterimReceptionsItem.Go_To_Old_Navigate();
            InterimReceptionsItem.Filter(InterimReceptionsItem.FilterType.SubGroups, subgrpname);

            //Assert
            Assert.IsTrue(InterimReceptionsItem.CheckNameExistSInList(name), "filtrage erreur");
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_RECEP_details_ItemLink()
        {
            var homePage = LogInAsAdmin();
            InterimReceptionsPage interimreceptionsPage = homePage.GoToInterim_Receptions();
            interimreceptionsPage.Filter(InterimReceptionsPage.FilterType.Shownotvalidatedonly, true);
            InterimReceptionsItem interimItem = interimreceptionsPage.ClickFirstlineinterimNotValidated();
            interimItem.ClickOnItem();
            interimItem.ClickOnCrayon();
            Assert.IsTrue(interimItem.ElementIsVisible(), "La page ne souvre pas");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void INT_RECEP_details_DeleteItem()
        {
            //arrange
            var homePage = LogInAsAdmin();

            //act
            var receptionsPage = homePage.GoToInterim_ReceptionsModified();
            receptionsPage.Filter(InterimReceptionsPage.FilterType.ShowNotValidatedOnly, true);
            var InterimReceptionsItem = receptionsPage.ClickFirstLine();
            InterimReceptionsItem.SetReceived("1000");
            InterimReceptionsItem.DeleteItem();
            var quantite = InterimReceptionsItem.GetQuanity();

            //assert
            Assert.AreEqual(quantite, 0.000, "La quantite n'est pas remis à 0");


        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_RECEP_index_Pagination()
        {
            List<int> nombre_pagination = new List<int> { 8, 16, 30, 50, 100 };
            var homePage = LogInAsAdmin();
            InterimReceptionsPage receptionsPage = homePage.GoToInterim_Receptions();
            receptionsPage.ResetFilters();
            var totalNumberLigne = receptionsPage.CheckTotalNumber();
            int display_pagination = 0;
            if (receptionsPage.CheckTotalNumber() >= nombre_pagination[display_pagination])
            {
                do
                {
                    receptionsPage.PageSize("" + nombre_pagination[display_pagination]);
                    var totalPages = (int)Math.Ceiling((double)totalNumberLigne / nombre_pagination[display_pagination]);
                    var nbpageresult = receptionsPage.GetNumberOfPageResults();
                    Assert.AreEqual(totalPages, nbpageresult, "Le Chargement du Total Page rencontre un problème.");
                    var AllNbLigne = totalNumberLigne;
                    if (nbpageresult >= 1)
                    {
                        for (int i = 0; i < nbpageresult; i++)
                        {
                            Assert.IsTrue(receptionsPage.IsChangementPage(i + 1), "Le changement de page ne fonctionne pas.");
                            receptionsPage.GetPageResults(i + 1);
                            var nbRowspage = receptionsPage.GetNumberOfShowedResults();
                            Assert.AreEqual(nbRowspage, AllNbLigne >= nombre_pagination[display_pagination] ? nombre_pagination[display_pagination] : AllNbLigne, "Le changement du nombre de ligne affichées ne fonctionne pas.");
                            AllNbLigne = AllNbLigne - nbRowspage;
                        }
                    }
                    display_pagination++;
                } while (display_pagination < 5);
            }
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_RECEP_index_AddNewCopyFromOrderNumber()
        {
            //Prepare
            string number = "001201";
            string site = TestContext.Properties["SiteACE"].ToString();
            string comment = "";
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string placeTo = "Produccion";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            InterimReceptionsPage receptionsPage = homePage.GoToInterim_Receptions();
            receptionsPage.ResetFilters();
            InterimReceptionsCreateModalPage modalcreateInterimReception = receptionsPage.CreateNewInterim();
            modalcreateInterimReception.FillField_CreatNewInterminAddComment(DateTime.Now, site, supplier, placeTo, comment);
            modalcreateInterimReception.ClickOnCreateInterimReceptionFrom();
            modalcreateInterimReception.SetNumber(number);

            //Assert
            Assert.IsTrue(modalcreateInterimReception.IsExistingFilteredNumber(), "Le nombre n'est pas été trouvé");

        }


        [Timeout(_timeout)]
		[TestMethod]
        public void INT_RECEP_index_AddNewCopyFromOrderDate()
        {
            //Prepare
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            string location = "Produccion";
            string comment = "This is my comment";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            var receptionsPage = homePage.GoToInterim_ReceptionsModified();

            InterimReceptionsCreateModalPage modalcreateInterim = receptionsPage.CreateNewInterim();
            modalcreateInterim.FillField_CreatNewInterminAddCommentReceptionFromOn(DateTime.Now, site, supplier, location, comment);
            modalcreateInterim.Filter(InterimReceptionsCreateModalPage.FilterType.DateFrom, DateUtils.Now.AddDays(-15));
            modalcreateInterim.Filter(InterimReceptionsCreateModalPage.FilterType.DateTo, DateUtils.Now.AddDays(15));
            modalcreateInterim.PageSize("100");
            bool testFiltering = modalcreateInterim.CheckDateFiltringWorking(DateUtils.Now.AddDays(-15), DateUtils.Now.AddDays(15));

            //assert
            Assert.IsTrue(testFiltering, " Filtering not working right");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void INT_RECEP_index_AddNewCopyFromReceptionNumber()
        {
            //Prepare
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            string location = "Produccion";
            string comment = "This is my comment";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            InterimReceptionsPage receptionsPage = homePage.GoToInterim_Receptions();

            InterimReceptionsCreateModalPage modalcreateInterim = receptionsPage.CreateNewInterim();

            IWebElement button = receptionsPage.GetAddNewInterimReceptionButton();
            button.Click();
            modalcreateInterim.FillField_CreatNewInterminAddCommentReceptionFromOn(DateTime.Now, site, supplier, location, comment);

            IWebElement numberElement = receptionsPage.GetExtractedNumberElement();
            string extractedNumber = numberElement.Text;

            IWebElement filterInput = receptionsPage.GetFilterInput();
            filterInput.Clear();
            filterInput.SendKeys(extractedNumber);

            Thread.Sleep(3000);

            WebDriverWait wait = new WebDriverWait(WebDriver, TimeSpan.FromSeconds(10));
            wait.Until(d => receptionsPage.GetFilteredRows().Count == 2);

            IList<IWebElement> filteredRows = receptionsPage.GetFilteredRows();
            Assert.AreEqual(2, filteredRows.Count, "Le tableau devrait contenir deux lignes après le filtrage.");

            IWebElement firstRow = filteredRows[0];
            IWebElement cell = receptionsPage.GetFirstRowCell();
            string cellValue = cell.Text;

            Assert.AreEqual(extractedNumber, cellValue, "La première ligne affichée ne contient pas la valeur filtrée attendue.");
        }




        [Timeout(_timeout)]
		[TestMethod]
        public void INT_RECEP_index_AddNewCopyFromReception()
        {
            //Prepare
            string site = TestContext.Properties["SiteACE"].ToString();
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string location = "Produccion";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            InterimReceptionsPage interimReceptionsPage = homePage.GoToInterim_Receptions();
            interimReceptionsPage.ResetFilters();
            InterimReceptionsCreateModalPage createModalPage = interimReceptionsPage.CreateNewInterim();
            createModalPage.CreateInterimReceptionFrom();
            createModalPage.ToInterimReception();
            createModalPage.FillField_CreatNewIntermin(DateTime.Now, site, supplier, location, true, true);
            var isInterimReceptionList = createModalPage.IsInterimReceptionList();
            var isInterimReceptionValid = createModalPage.IsInterimReceptionValid();

            //Assert
            Assert.IsTrue(isInterimReceptionList, "La liste de Interim Reception n'est pas visible.");
            Assert.IsTrue(isInterimReceptionValid, "Tous les éléments de la liste de Interim Reception ne sont pas validés.");
        }
        [Timeout(_timeout)]
		[TestMethod]

        public void INT_RECEP_index_AddNewFromReceptionPagination()
        {
            //Prepare
            string site = TestContext.Properties["SiteACE"].ToString();
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string location = "Produccion";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            InterimReceptionsPage interimReceptionsPage = homePage.GoToInterim_Receptions();
            interimReceptionsPage.ResetFilters();
            InterimReceptionsCreateModalPage createModalPage = interimReceptionsPage.CreateNewInterim();
            createModalPage.CreateInterimReceptionFrom();
            createModalPage.ToInterimReception();
            createModalPage.FillField_CreatNewIntermin(DateTime.Now, site, supplier, location, true, true);
            var isInterimReceptionList = createModalPage.IsInterimReceptionList();
            string firstnumber = createModalPage.GetFirstNumberOfFirstLineIR();
            createModalPage.PaginateToSecond();
            string firstnumberAfterPagination = createModalPage.GetFirstNumberOfFirstLineIR();

            //Assert
            Assert.AreNotEqual(firstnumber, firstnumberAfterPagination, "On ne peut pas naviguer de page en page");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void INT_RECEP_index_AddNewCopyFromReceptionDate()
        {
            string site = TestContext.Properties["SiteACE"].ToString();
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string location = "Produccion";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            string dateFormat = homePage.GetDateFormatPickerValue();
            InterimReceptionsPage interimReceptionsPage = homePage.GoToInterim_Receptions();
            interimReceptionsPage.ResetFilters();
            InterimReceptionsCreateModalPage createModalPage = interimReceptionsPage.CreateNewInterim();
            createModalPage.CreateInterimReceptionFrom();
            createModalPage.ToInterimReception();
            createModalPage.FillField_CreatNewIntermin(DateTime.Now, site, supplier, location, true, true);
            createModalPage.Filter(InterimReceptionsCreateModalPage.FilterType.DateFrom, DateTime.Now.AddDays(-10));
            createModalPage.Filter(InterimReceptionsCreateModalPage.FilterType.DateTo, DateTime.Now);
            createModalPage.SetPageSize("100");

            //Assert
            Assert.IsTrue(createModalPage.IsFromToDateRespected(DateTime.Now.AddDays(-10), DateTime.Now, dateFormat), String.Format(MessageErreur.FILTRE_ERRONE, "'From/To'"));
        }


        [Timeout(_timeout)]
		[TestMethod]
        public void INT_RECEP_index_AddNewCopyFromOrderSelect()
        {
            //prepare
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            string location = "Produccion";
            string comment = "This is my comment";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            /* We create a valide Interim Order */
            InterimOrdersPage interimOrdersPage = homePage.GoToInterim_InterimOrders();
            InterimOrdersCreateModalPage modalcreateInterimOrder = interimOrdersPage.CreateNewInterimOrder();
            modalcreateInterimOrder.FillField_CreatNewInterminOrder(DateTime.Now, site, supplier, location);
            InterimOrdersItem interimItemOrder = modalcreateInterimOrder.Submit();
            string randQty = new Random().Next(2, 10).ToString();
            interimItemOrder.SetQty(randQty);

            float totalVAT = interimItemOrder.GetTotal();
            interimItemOrder.Validate();

            string OrderNumber = interimItemOrder.ReturnInterimOrderNumber();
            interimItemOrder.BackToList();

            /* We create an interim reception from the interim order previously created */
            InterimReceptionsPage interimReceptionPage = homePage.GoToInterim_Receptions();
            InterimReceptionsCreateModalPage modalcreateInterim = interimReceptionPage.CreateNewInterim();
            modalcreateInterim.FillField_CreatNewInterminAddComment(DateTime.Now, site, supplier, location, comment);
            modalcreateInterim.ClickOnCreateInterimReceptionFrom();
            modalcreateInterim.SearchForInterimOrderNumberAndChoseFirst(OrderNumber);
            InterimReceptionsItem interimItem = modalcreateInterim.Submit();
            float totalWOvatitem = interimItem.GetTotalWOVATinNumbers();
        
            string interimReceptionNumber = interimItem.ReturnInterimReceptionNumber();

            /* Delete the created Interim reception for test */
            interimItem.BackToList();
            interimReceptionPage.DeleteInterim(interimReceptionNumber);

            //Assert
            Assert.AreEqual(totalVAT, totalWOvatitem, "L'intérim reception n'est pas crée");

        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_RECEP_index_AddNewCopyFromReceptionNoPrefill()
        {
            //Prepare
            string site = TestContext.Properties["SiteACE"].ToString();
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string location = "Produccion";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            InterimReceptionsPage interimReceptionsPage = homePage.GoToInterim_Receptions();
            interimReceptionsPage.ResetFilters();
            InterimReceptionsCreateModalPage createModalPage = interimReceptionsPage.CreateNewInterim();
            createModalPage.FillField_CreatNewIntermin(DateTime.Now, site, supplier, location, true, true);
            createModalPage.UnchekPrefillReceivedQuantities();
            var interimOrderToCopy = createModalPage.GetFirstNumberOfFirstLine();
            createModalPage.SelectFirstInterimOrder();
            InterimReceptionsItem interimReceptionsItem = createModalPage.Submit();
            var orderedInterimReceptionsItem = interimReceptionsItem.GetOrderedInterimReceptionItem(); // ya pas de ordered 
            var receivedInterimReceptionsItem = interimReceptionsItem.GetReceivedInterimReceptionItem();
            InterimOrdersPage interimOrdersPage = interimReceptionsItem.GoToReceptionOrder();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.SearchByNumber, interimOrderToCopy);
            InterimOrdersItem interimOrdersItem = interimOrdersPage.GoToInterim_InterimOrdersItem();
            var prodQtyInterimOrderItem = interimReceptionsItem.GetProdQtyInterimOrderItem();

            //Assert
            Assert.AreEqual(orderedInterimReceptionsItem, prodQtyInterimOrderItem, "Les quantités dans \"Ordered\" ne correspondent pas aux quantités commandées dans l'interim order.");
            Assert.AreEqual(receivedInterimReceptionsItem, 0, "Les quantités \"received\" ne doivent pas être pré-remplies.");
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_RECEP_index_AddNewCopyFromOrderSelects()
        {
            //Prepare
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            string location = "Produccion";
            string comment = "This is my comment";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            /* We create a valide Interim Order */
            InterimOrdersPage interimOrdersPage = homePage.GoToInterim_InterimOrders();

            /* Create the first Interim Order */
            InterimOrdersCreateModalPage modalcreateInterimOrder = interimOrdersPage.CreateNewInterimOrder();
            modalcreateInterimOrder.FillField_CreatNewInterminOrder(DateTime.Now, site, supplier, location);
            InterimOrdersItem interimItemOrder = modalcreateInterimOrder.Submit();
            string randQty = new Random().Next(2, 10).ToString();
            interimItemOrder.SetQty(randQty);
            Thread.Sleep(2000);

            float totalVAT1 = interimItemOrder.GetTotal();
            interimItemOrder.Validate();

            string OrderNumber1 = interimItemOrder.ReturnInterimOrderNumber();
            interimItemOrder.BackToList();

            /* We create the second Interim order */

            interimOrdersPage.CreateNewInterimOrder();
            modalcreateInterimOrder.FillField_CreatNewInterminOrder(DateTime.Now, site, supplier, location);
            modalcreateInterimOrder.Submit();
            string randQty2 = new Random().Next(2, 10).ToString();
            interimItemOrder.SetQty(randQty2);
            Thread.Sleep(2000);

            float totalVAT2 = interimItemOrder.GetTotal();
            interimItemOrder.Validate();

            string OrderNumber2 = interimItemOrder.ReturnInterimOrderNumber();
            interimItemOrder.BackToList();

            /* we sum the total VAT */

            float totalForBothInterimOrder = totalVAT1 + totalVAT2;

            /* We create an interim reception from the interim order previously created */
            InterimReceptionsPage interimReceptionPage = homePage.GoToInterim_Receptions();
            InterimReceptionsCreateModalPage modalcreateInterim = interimReceptionPage.CreateNewInterim();
            modalcreateInterim.FillField_CreatNewInterminAddComment(DateTime.Now, site, supplier, location, comment);
            modalcreateInterim.ClickOnCreateInterimReceptionFrom();
            modalcreateInterim.SearchForInterimOrderNumberAndChoseFirst(OrderNumber1);
            Thread.Sleep(3000);
            modalcreateInterim.SearchForInterimOrderNumberAndChoseFirst(OrderNumber2);
            Thread.Sleep(3000);
            InterimReceptionsItem interimItem = modalcreateInterim.Submit();

            float totalWOvatitem = interimItem.GetTotalWOVATinNumbers();

            string interimReceptionNumber = interimItem.ReturnInterimReceptionNumber();

            /* Delete the created Interim reception for test */
            interimItem.BackToList();
            interimReceptionPage.DeleteInterim(interimReceptionNumber);

            //Assert
            Assert.AreEqual(totalForBothInterimOrder, totalWOvatitem , "the total VAT are not equal ");

        }


        [Timeout(_timeout)]
		[TestMethod]
        public void INT_RECEP_Supplier_list()
        {
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            string location = "Produccion";
            string comment = "This is my comment";

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            /* We create an interim reception from the interim order previously created */
            InterimReceptionsPage interimReceptionPage = homePage.GoToInterim_Receptions();
            interimReceptionPage.WaitLoading();
            InterimReceptionsCreateModalPage modalcreateInterim = interimReceptionPage.CreateNewInterim();
            modalcreateInterim.FillField_CreatNewInterminWithoutOrderNumber(DateTime.Now, site, supplier, location);
            InterimReceptionsItem InterimItem = modalcreateInterim.Submit();

            bool isVerifiedSupplierAndList = modalcreateInterim.VerifiedSupplierAndList(supplier);
            //Assert
            Assert.IsTrue(isVerifiedSupplierAndList, "le fournisseur doit rester sélectionnable et la liste est non vide tant qu'il y a des fournisseurs.");


        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_RECEP_index_AddNewPagination()
        {
            string site = TestContext.Properties["SiteACE"].ToString();
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string location = "Produccion";

            //arrange
            HomePage homePage = LogInAsAdmin();
            //act
            InterimReceptionsPage interimPage = homePage.GoToInterim_Receptions();

            InterimReceptionsCreateModalPage modalcreateInterim = interimPage.CreateNewInterim();
            modalcreateInterim.FillField_CreatNewIntermin(DateTime.Now, site, supplier, location);
            modalcreateInterim.CreateInterimFrom();

            modalcreateInterim.ValidatePagination("8", 8);
            modalcreateInterim.ValidatePagination("16", 16);
            modalcreateInterim.ValidatePagination("30", 30);
            modalcreateInterim.ValidatePagination("50", 50);
            modalcreateInterim.ValidatePagination("100", 100, true);
        }



    }
}