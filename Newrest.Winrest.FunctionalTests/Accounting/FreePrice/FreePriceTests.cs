using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Accounting.FreePrice;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.Claims;
using Newrest.Winrest.FunctionalTests.Utils;
using System;
using System.IO;
using System.Linq;
using System.Security.Policy;

namespace Newrest.Winrest.FunctionalTests.Accounting
{
    [TestClass]
    public class FreePriceTests : TestBase
    {

        private const int _timeout = 600000;

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_FRPR_Filter_Search()
        {
            //Prepare 
            var nameInputSearch = "€20.00 Phone Card";

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            FreePricePage freePricePage = homePage.GoToAccounting_FreePricePage();
            freePricePage.ResetFilters();
            var avant = freePricePage.GetNumberInHeaderFreePrices();
            freePricePage.Filter(FreePricePage.FilterType.Search, nameInputSearch);
            var apres = freePricePage.GetNumberInHeaderFreePrices();

            //Assert
            Assert.AreNotEqual(avant, apres, "Filter Name, Foreign Name ne fonctionne pas.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_FRPR_Filter_ShowAll()
        {
            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            FreePricePage freePricePage = homePage.GoToAccounting_FreePricePage();
            freePricePage.ResetFilters();

            freePricePage.Filter(FreePricePage.FilterType.ShowOnlyActive, true);
            var resultsActivated = freePricePage.CheckTotalNumber();

            freePricePage.Filter(FreePricePage.FilterType.ShowOnlyInactive, true);
            var resultsInactivated = freePricePage.CheckTotalNumber();

            freePricePage.Filter(FreePricePage.FilterType.ShowAll, true);
            var  resultAll = freePricePage.CheckTotalNumber();
         
            //Assert
            Assert.AreEqual(resultsActivated + resultsInactivated, resultAll, "Les résultats ne sont pas mis à jour en fonction des filtres Show All");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_FRPR_Filter_Site()
        { 
            // Prepare      
            string site = TestContext.Properties["SiteACE"].ToString();
            
            //Arrange
            HomePage homePage = LogInAsAdmin();
            
            //Act
            FreePricePage freePricePage = homePage.GoToAccounting_FreePricePage();
            freePricePage.ResetFilters();
            freePricePage.Filter(FreePricePage.FilterType.Sites, site);
            var listSitesInGrid = freePricePage.GetSites();
           
            //Assert
            bool isFilterSiteOK = freePricePage.VerifyFilterSite(listSitesInGrid, site); 
            Assert.IsTrue(isFilterSiteOK, "Filter site is failed.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_FRPR_Filter_Customer()
        {
            //Prepare 
            var customer = "AVT EU UNIPERSSOAL LDA";

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            FreePricePage freePricePage = homePage.GoToAccounting_FreePricePage();
            freePricePage.Filter(FreePricePage.FilterType.Customer, customer);
            var listCustomersInGrid = freePricePage.GetCustomers();

            //Assert
            bool isVerifyFilterCustomerOk = freePricePage.VerifyFilterCustomer(listCustomersInGrid, customer);
            Assert.IsTrue(isVerifyFilterCustomerOk, "Filter customer is failed.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_FRPR_ResetFilter()
        {
            object value = null;
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            FreePricePage freePricePage = homePage.GoToAccounting_FreePricePage();

            freePricePage.ResetFilters();

            var numberDefaultListSiteFilter = freePricePage.GetNumberSelectedSiteFilter();
            var numberDefaultListCustomerFilter = freePricePage.GetNumberSelectedCustomerFilter();

            freePricePage.Filter(FreePricePage.FilterType.Search, "CAPT 1ST SERVICE");
            freePricePage.Filter(FreePricePage.FilterType.Sites, "BCN");
            freePricePage.Filter(FreePricePage.FilterType.Customer, "ASIANA AIRLINES INC");
            freePricePage.Filter(FreePricePage.FilterType.ShowAll, true);
            freePricePage.Filter(FreePricePage.FilterType.ShowOnlyActive, true);
            freePricePage.Filter(FreePricePage.FilterType.ShowOnlyInactive, true);

            freePricePage.ResetFilters();

            value = freePricePage.GetFilterValue(FreePricePage.FilterType.Search);
            Assert.AreEqual("", value, "ResetFilter Search ''");
            var numberSelectedSiteFilter = freePricePage.GetNumberSelectedSiteFilter();
            Assert.AreEqual(numberSelectedSiteFilter, numberDefaultListSiteFilter, "ResetFilter Site ''");
            var numberSelectedCustomerFilter = freePricePage.GetNumberSelectedCustomerFilter();
            Assert.AreEqual(numberSelectedCustomerFilter, numberDefaultListCustomerFilter, "ResetFilter Customer ''");
            value = freePricePage.GetFilterValue(FreePricePage.FilterType.ShowAll);
            Assert.AreEqual(true, value, "ResetFilter ShowAll");
            value = freePricePage.GetFilterValue(FreePricePage.FilterType.ShowOnlyActive);
            Assert.AreEqual(false, value, "ResetFilter ShowOnlyActive");
            value = freePricePage.GetFilterValue(FreePricePage.FilterType.ShowOnlyInactive);
            Assert.AreEqual(false, value, "ResetFilter ShowOnlyInactive");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_FRPR_Filter_ShowOnlyActive()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            FreePricePage freePricePage = homePage.GoToAccounting_FreePricePage();
            freePricePage.Filter(FreePricePage.FilterType.ShowOnlyActive, true);
            //verify for the first item 
            var freePriceDetailPage = freePricePage.SelectFirstFreePrice();
            Assert.IsTrue(freePriceDetailPage.VerifyIsActiveFilter(), "erreur de filtrage Show Only Active");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_FRPR_Filter_ShowOnlyInactive()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            FreePricePage freePricePage = homePage.GoToAccounting_FreePricePage();
            freePricePage.Filter(FreePricePage.FilterType.ShowOnlyInactive, true);
            //verify for the first item 
            var freePriceDetailPage = freePricePage.SelectFirstFreePrice();
            Assert.IsFalse(freePriceDetailPage.VerifyIsActiveFilter(), "erreur de filtrage Show Only InActive");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_FRPR_ModifyDetail()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            FreePricePage freePricePage = homePage.GoToAccounting_FreePricePage();
            var nameFirstFreePrice = freePricePage.GetFirstFreePriceName();
            FreePriceDetailsPage freePriceDetailsPage = freePricePage.SelectFirstFreePrice();

            freePriceDetailsPage.EditDetailFreePrice("TEST AUTO", "11", "Airport Tax", "Corte", false);
            freePriceDetailsPage.BackToList();

            freePricePage.Filter(FreePricePage.FilterType.Search, nameFirstFreePrice);

            freePriceDetailsPage = freePricePage.SelectFirstFreePrice();
            Assert.IsTrue(freePriceDetailsPage.VerifyDataEdited("TEST AUTO", "11", "Airport Tax", "Corte", false), "Les données ne sont pas modifiés.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_FRPR_Filter_SortByName()
        {
            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            FreePricePage freePricePage = homePage.GoToAccounting_FreePricePage();
            freePricePage.ResetFilters();
            freePricePage.Filter(FreePricePage.FilterType.ShowOnlyActive, true);
            freePricePage.Filter(FreePricePage.FilterType.SortByName, null, null);
            freePricePage.PageSizeEqualsTo100();

            //Assert
            bool isSortedByNameOK = freePricePage.IsSortedByName();
            Assert.IsTrue(isSortedByNameOK, String.Format(MessageErreur.FILTRE_ERRONE, "'Sort by name'"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_FRPR_Filter_SortById()
        {
            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            FreePricePage freePricePage = homePage.GoToAccounting_FreePricePage();
            freePricePage.ResetFilters();
            freePricePage.Filter(FreePricePage.FilterType.ShowOnlyActive, true);
            freePricePage.Filter(FreePricePage.FilterType.SortById, null, null);
           
            //Assert
            bool isSortByIdOK = freePricePage.VerifySortById();
            Assert.IsTrue(isSortByIdOK, "Les données ne sont pas filtré par Id.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_FRPR_Filter_SortByIdDescending()
        {
            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            FreePricePage freePricePage = homePage.GoToAccounting_FreePricePage();
            freePricePage.ResetFilters();
            freePricePage.Filter(FreePricePage.FilterType.ShowOnlyActive, true);
            freePricePage.Filter(FreePricePage.FilterType.SortByIdDescending, null, null);

            //Assert
            bool isSortByIdDescendingOK = freePricePage.VerifySortByIdDescending();
            Assert.IsTrue(isSortByIdDescendingOK, "erreur de sorting by id descending");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_FRPR_Export()
        {
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            homePage.ClearDownloads();

            FreePricePage freePricePage = homePage.GoToAccounting_FreePricePage();

            freePricePage.ResetFilters();

            var nbrfreePrice = freePricePage.CheckTotalNumber();
            //clear download directory 
            foreach (var file in new DirectoryInfo(downloadsPath).GetFiles())
            {
                file.Delete();
            }
            //export
            freePricePage.Export();

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();


            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = freePricePage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);
            var resultNumber = OpenXmlExcel.GetExportResultNumber("FreePrices", filePath);
            //verifier les données 
            Assert.AreEqual(resultNumber, nbrfreePrice, "les donées du grid et du fichier Excel ne sont pas identiques");
        }


        [Timeout(_timeout)]
        [TestMethod]
        public void AC_FRPR_Massivedelete()
        {

            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string workshop = TestContext.Properties["InvoiceWorkshop"].ToString();

            string namePrice = "Free Price For Massive Delete " + new Random().Next().ToString();
            string qtyValue = "100";
            string sellingPrice = "2";


            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Create a new freePrice
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
            var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customer, true);

            // On créé une nouvelle ligne FreePrice
            var createFreePrice = invoiceDetails.AddFreePrice();
            createFreePrice.FillField_CreatNewFreePrice(namePrice, qtyValue, sellingPrice, workshop);
            invoiceDetails = createFreePrice.ValidateForInvoice();

            FreePricePage freePricePage = homePage.GoToAccounting_FreePricePage();
            freePricePage.ResetFilters();
            freePricePage.Filter(FreePricePage.FilterType.Search, namePrice);
            var firstCustomerName = freePricePage.GetFirstCustomer();
            var firstSiteName = freePricePage.GetFirstSite();
            var status = "Active free price";
            freePricePage.MassiveDelete();
            freePricePage.DeleteFreePrice(true, true, namePrice, firstSiteName, firstCustomerName, status);
            freePricePage.Filter(FreePricePage.FilterType.Search, namePrice);
            Assert.IsTrue(freePricePage.CheckTotalNumber() == 0);

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_FRPR_Import()
        {
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            homePage.ClearDownloads();
            FreePricePage freePricePage = homePage.GoToAccounting_FreePricePage();
            freePricePage.ResetFilters();

            var nbrfreePrice = freePricePage.CheckTotalNumber();
            //clear download directory 
            foreach (var file in new DirectoryInfo(downloadsPath).GetFiles())
            {
                file.Delete();
            }
            //export
            freePricePage.Export();

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = freePricePage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);
            var resultNumber = OpenXmlExcel.GetExportResultNumber("FreePrices", filePath);
            //verifier les données 
            Assert.AreEqual(resultNumber, nbrfreePrice, "les donées du grid et du fichier Excel ne sont pas identiques");

            freePricePage.ClickImport();
            bool isImported = freePricePage.ImportFile(correctDownloadedFile.FullName);
            Assert.IsTrue(isImported, "Erreur lors de l'import du fichier.");

        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_FRPR_ExportFilterSite()
        {
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            var customer = "AIR EUROPA LINEAS AEREAS,S.A.U";
            var site = TestContext.Properties["SiteACE"].ToString();
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            homePage.ClearDownloads();
            DeleteAllFileDownload();
            FreePricePage freePricePage = homePage.GoToAccounting_FreePricePage();
            freePricePage.ResetFilters();
            freePricePage.Filter(FreePricePage.FilterType.Customer, customer);
            freePricePage.Filter(FreePricePage.FilterType.Sites, site);
            var nbrfreePrice = freePricePage.CheckTotalNumber();
            freePricePage.Export();
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();
            var correctDownloadedFile = freePricePage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);
            var resultNumber = OpenXmlExcel.GetExportResultNumber("FreePrices", filePath);
            var listResult = OpenXmlExcel.GetValuesInList("Site Name", "FreePrices", filePath);
            Assert.AreEqual(resultNumber, nbrfreePrice, "les donées du grid et du fichier Excel ne sont pas identiques");
            Assert.IsTrue(listResult.Any(item => item.Equals(site)), MessageErreur.EXCEL_DONNEES_KO);

        }
    }
}
