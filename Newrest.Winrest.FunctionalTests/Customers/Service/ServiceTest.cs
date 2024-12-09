using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Customer;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Delivery;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.LoadingPlans;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Datasheet;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Sites;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.User;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using static Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service.ServiceMassiveDeleteModalPage;

namespace Newrest.Winrest.FunctionalTests.Customers
{

    [TestClass]
    public class ServiceTest : TestBase
    {
        private const int _timeout = 600000; 
        private readonly string serviceNameToday = "Service-" + DateUtils.Now.ToString("dd/MM/yyyy");
        private const string SERVICES_EXCEL_SHEET_NAME = "Services Prices";
        private const string SERVICES_PRINT_SHEET_NAME = "Services Calendar";
        public const string SHEET1 = "Services Prices";
        public const string CUSTOMER = "$$ - CAT Genérico";
        //_______________________________________________________________CREATE_____________________________________________________________

        [TestInitialize]
        public override void TestInitialize()
        {
            base.TestInitialize();

            var testMethod = TestContext.TestName;
            switch (testMethod)
            {
                case nameof(CU_SERV_Import_With_A_ActionCategoryManyLine):
                case nameof(CU_SERV_Import_With_A_ActionCategoryManyLine_Error):
                case nameof(CU_SERV_Import_With_A_ActionDefaultLoadingModeNotExist):
                case nameof(CU_SERV_Import_With_M_ActionDefaultLoadingModeError):
                case nameof(CU_SERV_Import_With_A_ActionIDCustomerSrvEmpty):
                case nameof(CU_SERV_Import_With_A_ActionIDCustomerSrvExist):

                    CU_SERV_Import_CreateService_TestInitialize();
                    break;
                case nameof(CU_SERV_Import_With_A_ActionAddDatasheet):
                    CU_SERV_Import_With_A_ActionAddDatasheet_TestInitialize();
                    break;
                case nameof(CU_SERV_Import_With_M_ActionCustomerSrv):
                    CU_SERV_Import_With_M_ActionCustomerSrv_TestInitialize();
                    break;
                default:
                    break;
            }
        }

        [TestCleanup]
        public override void TestCleanup()
        {
            base.TestCleanup();

            var testMethod = TestContext.TestName;
            switch (testMethod)
            {
                case nameof(CU_SERV_Import_With_A_ActionCategoryManyLine):
                case nameof(CU_SERV_Import_With_A_ActionCategoryManyLine_Error):
                case nameof(CU_SERV_Import_With_A_ActionDefaultLoadingModeNotExist):
                case nameof(CU_SERV_Import_With_A_ActionIDCustomerSrvEmpty):
                case nameof(CU_SERV_Import_With_A_ActionIDCustomerSrvExist):
                case nameof(CU_SERV_Import_With_M_ActionCustomerSrv):
                    CU_SERV_Import_CreateService_TestCleanup();
                    break;
                case nameof(CU_SERV_Import_With_A_ActionAddDatasheet):
                    CU_SERV_Import_With_A_ActionAddDatasheet_TestCleanup();
                    break;

                default:
                    break;
            }
        }

        private void CU_SERV_Import_CreateService_TestInitialize()
        {
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);

            string query = @"
declare @categoryId int;
select top 1 @categoryId = Id from ServiceCategories order by name;
insert into services (Name, Code, CategoryId, ProductionName, IsActive, IsProduced, IsInvoiced) 
    values (@serviceName,@serviceCode,@categoryId,@serviceProduction, 1, 0, 0);
select scope_identity();";

            TestContext.Properties["ServiceName"] = serviceName;
            TestContext.Properties["ServiceId"] = ExecuteAndGetInt(
                query,
                new KeyValuePair<string, object>("serviceName", serviceName),
                new KeyValuePair<string, object>("serviceCode", serviceCode),
                new KeyValuePair<string, object>("serviceProduction", serviceProduction));

            var customerName = TestContext.Properties["CustomerLP"] as string;
            query = "select Code from Customers where Name = @customerName";
            TestContext.Properties["CustomerCode"] = ExecuteAndGetString(query, new KeyValuePair<string, object>("customerName", customerName));
        }
        private void CU_SERV_Import_CreateService_TestCleanup()
        {
            int serviceId = (int)TestContext.Properties["ServiceId"];

            string query = @"
delete ServiceToCustomerToSites where ServiceId = @serviceId;
delete Services where Id = @serviceId;";

            ExecuteNonQuery(query, new KeyValuePair<string, object>("serviceId", serviceId));
        }

        private void CU_SERV_Import_With_A_ActionAddDatasheet_TestInitialize()
        {
            var customerName = TestContext.Properties["CustomerLP"] as string;
            var query = "select Code from Customers where Name = @customerName";
            TestContext.Properties["CustomerCode"] = ExecuteAndGetString(query, new KeyValuePair<string, object>("customerName", customerName));
        }
        private void CU_SERV_Import_With_A_ActionAddDatasheet_TestCleanup()
        {
            string newService = TestContext.Properties["NewServiceName"] as string;

            string query = @"
delete ServiceToCustomerToSites where ServiceId in (select sv.Id from Services sv where sv.Name = @newServiceName);
delete Services where Name = @newServiceName;";

            ExecuteNonQuery(query, new KeyValuePair<string, object>("newServiceName", newService));

        }
        private void CU_SERV_Import_With_M_ActionCustomerSrv_TestInitialize()
        {
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string site = TestContext.Properties["SiteACE"].ToString();
            string customerSrv = "srv";
            string taxName = TestContext.Properties["InvoiceTaxName"].ToString();


            string query = @"
      declare @categoryId int;
      select top 1 @categoryId = Id from ServiceCategories order by name;
      insert into services (Name, Code, CategoryId, ProductionName, IsActive ,ServiceTypeId,  IsProduced, IsInvoiced) 
          values (@serviceName,@serviceCode,@categoryId,@serviceProduction, 1, '1', 0, 0);
      select scope_identity();";

            TestContext.Properties["ServiceName"] = serviceName;
            TestContext.Properties["ServiceId"] = ExecuteAndGetInt(
                query,
                new KeyValuePair<string, object>("serviceName", serviceName),
                new KeyValuePair<string, object>("serviceCode", serviceCode),
                new KeyValuePair<string, object>("serviceProduction", serviceProduction));

            var customerName = TestContext.Properties["CustomerLP"] as string;
            query = "select Code from Customers where Name = @customerName";
            TestContext.Properties["CustomerCode"] = ExecuteAndGetString(query, new KeyValuePair<string, object>("customerName", customerName));

            TestContext.Properties["ServiceDetailId"] = InsertServiceDetail(site, customerName, taxName, (int)TestContext.Properties["ServiceId"], customerSrv);

        }
        private int InsertServiceDetail(string site, string customer, string taxType, int serviceId, string customerSrv = null)
        {
            string query = @"
            DECLARE @siteId INT;
            SELECT TOP 1 @siteId = Id FROM Sites WHERE Name LIKE @site;
 
            DECLARE @customerId INT;
            SELECT TOP 1 @customerId = Id FROM Customers WHERE Name LIKE @customer;
 
            DECLARE @taxTypeId INT;
            SELECT TOP 1 @taxTypeId = Id FROM TaxTypes WHERE Name LIKE @taxType;
 
            Insert into [dbo].[ServiceToCustomerToSites] (
            [ServiceId]
              ,[SiteId]
              ,[CustomerId]
                ,[CustomerSrv]
              ,[BeginDate]
              ,[EndDate]
              ,[TaxTypeId]
              ,[Method]
              ,[IsActive]
              ,[SellingUnitId]
              ,[DefaultMode]
              ,[Value])
 
              values
              (
              @serviceId,
              @siteId,
              @customerId,
                @customerSrv,
              '2024-10-23 00:00:00.000',
              '2024-10-30 00:00:00.000',
              @taxTypeId,
              '10',
              1,
              '1',
              '10',
              '1'
              );
            select scope_identity();";

            return ExecuteAndGetInt(
                query,
                new KeyValuePair<string, object>("taxType", taxType),
                new KeyValuePair<string, object>("customer", customer),
                new KeyValuePair<string, object>("serviceId", serviceId),
                new KeyValuePair<string, object>("site", site),
                new KeyValuePair<string, object>("customerSrv", customerSrv));
        }
        private string GetCustomerSrv(int servicedetailId)
        {
            string query = @"
            Select CustomerSrv FROM [dbo].[ServiceToCustomerToSites] 
            WHERE Id = @servicedetailId;";

            return ExecuteAndGetString(
                query,
                new KeyValuePair<string, object>("servicedetailId", servicedetailId));
        }

        [Priority(0)]
		[TestMethod]
        [Timeout(_timeout)]
        public void CU_SERV_CreateCustomerForService()
        {
            //Prepare
            string customerName = "customerForservice";
            string deliveryName = "deliveryForService";
            string customerIcao = "Icao4Test" + DateTime.Today.ToShortDateString();
            string customerTypeInflight = TestContext.Properties["CustomerType2"].ToString();
            string contactName = "contact1";
            string mail = "test@test.fr";
            string phone = "+216 123 456 789";
            string siteACE = TestContext.Properties["SiteACE"].ToString();
            string serviceName = "ServiceForTest";
            DateTime fromDate = DateUtils.Now.AddDays(-10);
            DateTime toDate = DateUtils.Now.AddDays(10);
            string dataSheetName = "datasheetForService";
            string guestType = "NONE";
            //Arrange
            HomePage homePage = LogInAsAdmin();

            //act
            CustomerPage customerPage = homePage.GoToCustomers_CustomerPage();
            customerPage.ResetFilters();
            customerPage.Filter(CustomerPage.FilterType.Search, customerName);
            if (customerPage.CheckTotalNumber() == 0)
            {
                CustomerCreateModalPage customerCreateModalPage = customerPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customerName, customerIcao, customerTypeInflight);
                CustomerGeneralInformationPage customerGeneralInformationPage = customerCreateModalPage.Create();
                CustomerContactsPage customerContactsPage = customerGeneralInformationPage.GoToContactsPage();
                customerContactsPage.CreateContact(contactName, mail, phone, siteACE);
                customerPage = customerContactsPage.BackToList();
                customerPage.ResetFilters();
                customerPage.Filter(CustomerPage.FilterType.Search, customerName);
            }
            var firstCustomerName = customerPage.GetFirstCustomerName();
            Assert.AreEqual(customerName, firstCustomerName, "Le customer n'a pas été créé.");
            DatasheetPage dataSheetPage = homePage.GoToMenus_Datasheet();
            dataSheetPage.ResetFilter();
            dataSheetPage.Filter(DatasheetPage.FilterType.DatasheetName, dataSheetName);
            if (dataSheetPage.CheckTotalNumber() == 0)
            {
                DatasheetCreateModalPage datasheetCreatePage = dataSheetPage.CreateNewDatasheet();
                var datasheetDetailPage = datasheetCreatePage.FillField_CreateNewDatasheet(dataSheetName, guestType, siteACE);
                datasheetDetailPage.BackToList();
                dataSheetPage.ResetFilter();
                dataSheetPage.Filter(DatasheetPage.FilterType.DatasheetName, dataSheetName);
            }
            var firstDatasheetName = dataSheetPage.GetFirstDatasheetName();
            Assert.AreEqual(dataSheetName, firstDatasheetName, "Le Datasheet n'a pas été créé.");
            ServicePage servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);

            if (servicePage.CheckTotalNumber() == 0)
            {
                ServiceCreateModalPage ServiceCreatePage = servicePage.ServiceCreatePage();
                ServiceCreatePage.FillFields_CreateServiceModalPage(serviceName);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = ServiceCreatePage.Create();
                ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteACE, customerName, fromDate, toDate, null, dataSheetName);
                servicePage = pricePage.BackToList();
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            }
            var firstServiceName = servicePage.GetFirstServiceName().Contains(serviceName);
            Assert.IsTrue(firstServiceName, "Le service n'a pas été créé.");

            DeliveryPage deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.ResetFilter();
            deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName);

            if (deliveryPage.CheckTotalNumber() == 0)
            {
                DeliveryCreateModalPage deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, customerName, siteACE, true);
                DeliveryLoadingPage deliveryLoadingPage = deliveryCreateModalPage.Create();
                deliveryLoadingPage.AddService(serviceName);
                deliveryPage = deliveryLoadingPage.BackToList();
                deliveryPage.ResetFilter();
                deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName);
            }

            //assert
            var firstDeliveryName = deliveryPage.GetFirstDeliveryName().Contains(deliveryName);
            Assert.IsTrue(deliveryPage.GetFirstDeliveryName().Contains(deliveryName), "Le service n'a pas été créé.");
        }
        //Filter Search
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Filter_Search()
        {
            // Prepare
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);

            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();

            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();
            servicePage = serviceGeneralInformationsPage.BackToList();

            //Assert
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);

            servicePage.UncheckAllSites();
            servicePage.UncheckAllCustomers();

            bool isServiceNameCorrecte = servicePage.GetFirstServiceName().Contains(serviceName);
            Assert.IsTrue(isServiceNameCorrecte, MessageErreur.FILTRE_ERRONE, "Search");
        }

        //Filter Sort by
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Filter_Sort_By()
        {
            // Prepare
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();

            if (servicePage.CheckTotalNumber() < 20)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                servicePage = serviceGeneralInformationsPage.BackToList();
                servicePage.ResetFilters();
            }

            if (!servicePage.isPageSizeEqualsTo100())
            {
                servicePage.PageSize("8");
                servicePage.PageSize("100");
            }

            servicePage.Filter(ServicePage.FilterType.SortBy, "Name");
            var isSortedByName = servicePage.IsSortedByName();

            servicePage.Filter(ServicePage.FilterType.SortBy, "Category");
            var isSortedByCategory = servicePage.IsSortedByCategory();

            //Assert
            Assert.IsTrue(isSortedByName, String.Format(MessageErreur.FILTRE_ERRONE, "Sort by name"));
            Assert.IsTrue(isSortedByCategory, String.Format(MessageErreur.FILTRE_ERRONE, "Sort by category"));
        }

        //Filter Show All
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Filter_Show_All()
        {
            // Prepare
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);

            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();

            if (servicePage.CheckTotalNumber() == 0)
            {
                var ServiceCreateModalPage = servicePage.ServiceCreatePage();
                ServiceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
                var serviceGeneralInformationsPage = ServiceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10));

                servicePage = pricePage.BackToList();
                servicePage.ResetFilters();
            }

            servicePage.ResetFilters();

            servicePage.Filter(ServicePage.FilterType.ShowOnlyActive, true);
            var nbr1 = servicePage.CheckTotalNumber();

            servicePage.Filter(ServicePage.FilterType.ShowOnlyInactive, true);
            var nbr2 = servicePage.CheckTotalNumber();

            servicePage.Filter(ServicePage.FilterType.ShowAll, true);
            var realNbr = servicePage.CheckTotalNumber();

            //Assert
            Assert.AreEqual(nbr1 + nbr2, realNbr, String.Format(MessageErreur.FILTRE_ERRONE, "Show all"));
        }

        //Filter Show active
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Filter_Show_Active()
        {
            // Prepare
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);

            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.ShowOnlyActive, true);

            if (servicePage.CheckTotalNumber() < 20)
            {
                // Create
                var ServiceCreateModalPage = servicePage.ServiceCreatePage();
                ServiceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
                var serviceGeneralInformationsPage = ServiceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10));
                servicePage = pricePage.BackToList();

                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.ShowOnlyActive, true);
            }

            if (!servicePage.isPageSizeEqualsTo100())
            {
                servicePage.PageSize("8");
                servicePage.PageSize("100");
            }

            bool statusCheck = servicePage.CheckStatus(true);
            Assert.IsTrue(statusCheck, string.Format(MessageErreur.FILTRE_ERRONE, "Show only active"));
        }

        //Filter Show Inactive
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Filter_Show_Inactive()
        {
            // Prepare
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);

            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.ShowOnlyInactive, true);

            if (servicePage.CheckTotalNumber() < 20)
            {
                // Create
                var ServiceCreateModalPage = servicePage.ServiceCreatePage();
                ServiceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
                var serviceGeneralInformationsPage = ServiceCreateModalPage.Create();
                serviceGeneralInformationsPage.SetActive(false);

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10));
                servicePage = pricePage.BackToList();
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.ShowOnlyInactive, true);
            }

            if (!servicePage.isPageSizeEqualsTo100())
            {
                servicePage.PageSize("8");
                servicePage.PageSize("100");
            }

            Assert.IsFalse(servicePage.CheckStatus(false), string.Format(MessageErreur.FILTRE_ERRONE, "Show only inactive"));
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Filter_Show_Generic()
        {
            // Prepare
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);

            bool newVersionPrint = true;

            var homePage = LogInAsAdmin();

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();

            servicePage.ResetFilters();
            servicePage.ClearDownloads();

            servicePage.Filter(ServicePage.FilterType.ShowGenericServices, true);
          
            var ServiceCreateModalPage = servicePage.ServiceCreatePage();
            ServiceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = ServiceCreateModalPage.Create();
            serviceGeneralInformationsPage.SetGeneric();
            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10));
            servicePage = pricePage.BackToList();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.ShowGenericServices, true);
          
            servicePage.Export(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber(SERVICES_EXCEL_SHEET_NAME, filePath);
            bool result = OpenXmlExcel.ReadAllDataInColumn("Is Generic", SERVICES_EXCEL_SHEET_NAME, filePath, "VRAI");

            //Assert
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.IsTrue(result, string.Format(MessageErreur.FILTRE_ERRONE, "Show only generic services"));
        }

        // Filter price expired
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Filter_Show_Price_Expired()
        {
            // Prepare
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);

            DateTime fromDate = DateUtils.Now.AddDays(-10);
            DateTime toDate = DateUtils.Now.AddDays(-5);

            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.ShowServicesWithExpiredPrice, true);
          
            var ServiceCreateModalPage = servicePage.ServiceCreatePage();
            ServiceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = ServiceCreateModalPage.Create();

            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, fromDate, toDate);
            servicePage = pricePage.BackToList();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.ShowServicesWithExpiredPrice, true);
         
            if (!servicePage.isPageSizeEqualsTo100())
            {
                servicePage.PageSize("8");
                servicePage.PageSize("100");
            }

            bool IsPriceActive = servicePage.IsPriceActive();
            //Assert
            Assert.IsFalse(IsPriceActive, MessageErreur.FILTRE_ERRONE, "Show only active services with expired prices");
        }

        // Create price
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Filter_Show_Valid_Price_On()
        {
            // Prepare
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);

            DateTime fromDate = DateUtils.Now;
            DateTime toDate = DateUtils.Now.AddDays(2);

            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.ShowServicesWithValidPrice, true);

            if (servicePage.CheckTotalNumber() < 20)
            {
                // Create
                var ServiceCreateModalPage = servicePage.ServiceCreatePage();
                ServiceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
                var serviceGeneralInformationsPage = ServiceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, fromDate, toDate);
                servicePage = pricePage.BackToList();

                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.ShowServicesWithValidPrice, true);
            }

            //Assert
            try
            {
                servicePage.PageSize("100");
                Assert.IsTrue(servicePage.IsPriceActive(), MessageErreur.FILTRE_ERRONE, "Show service with valid price");
            }
            finally
            {
                servicePage.PageSize("8");
            }
        }

        //Filter Show Customer
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Filter_Customer()
        {
            // Prepare
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string customer = TestContext.Properties["Customer"].ToString();
            string customerCode = TestContext.Properties["CustomerCode"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            DateTime fromDate = DateUtils.Now;
            DateTime toDate = DateUtils.Now.AddDays(10);
            string nomCustomer = customerCode + " - " + customer;

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Customers, nomCustomer);
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);

            if (servicePage.CheckTotalNumber() == 0)
            {
                // Create
                var ServiceCreateModalPage = servicePage.ServiceCreatePage();
                ServiceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
                var serviceGeneralInformationsPage = ServiceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                priceModalPage.FillFields_CustomerPrice(site, customer, fromDate, toDate);
                servicePage = serviceGeneralInformationsPage.BackToList();

                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Customers, nomCustomer);
            }

            if (!servicePage.isPageSizeEqualsTo100())
            {
                servicePage.PageSize("8");
                servicePage.PageSize("100");
            }
                //Assert
                bool isCustomerVerified = servicePage.VerifyCustomer(customerCode); 
                Assert.IsTrue(isCustomerVerified, MessageErreur.FILTRE_ERRONE, "Les résultats ne concordent pas avec le filtre utilisé");
        }

        //Filter Show Site
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Filter_Site()
        {
            // Prepare
            string customer = TestContext.Properties["CustomerLPFlight"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);

            DateTime fromDate = DateUtils.Now;
            DateTime toDate = DateUtils.Now.AddDays(10);

            //Log in
            var homePage = LogInAsAdmin();

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            try
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, fromDate, toDate);
                homePage.GoToCustomers_ServicePage();

                servicePage.ResetFilters();

                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
                servicePage.Filter(ServicePage.FilterType.Sites, site);
                var filterSite = servicePage.CheckTotalNumber();

                // Assert
                Assert.AreEqual(filterSite, 1, "The number of services displayed does not match the expected filter criteria.");
            }
            finally
            {
                servicePage = homePage.GoToCustomers_ServicePage();
                var serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();

            }
        }

        //Filter Show Category
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Filter_Category()
        {
            // Prepare
            string customer = TestContext.Properties["CustomerLPFlight"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string category = TestContext.Properties["CategoryService"].ToString();
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            DateTime fromDate = DateUtils.Now;
            DateTime toDate = DateUtils.Now.AddDays(10);

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Categories, category);
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);

            if (servicePage.CheckTotalNumber() == 0)
            {
                // Create
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, fromDate, toDate);
                servicePage = pricePage.BackToList();

                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Categories, category);
            }

            if (!servicePage.isPageSizeEqualsTo100())
            {
                servicePage.PageSize("8");
                servicePage.PageSize("100");
            }

            //Assert
            bool isCategoryVerified = servicePage.VerifyCategory(category);
            Assert.IsTrue(isCategoryVerified, MessageErreur.FILTRE_ERRONE, "Les résultats ne concordent pas avec le filtre utilisé.");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Create_New_Service()
        {
            // Prepare
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);

            //Log in
            var homePage = LogInAsAdmin();

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();

            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

            bool isGeneralInformationPageLoaded = servicePage.IsGeneralInformationServiceLoaded();
            servicePage = serviceGeneralInformationsPage.BackToList();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);

            servicePage.UncheckAllCustomers();
            servicePage.UncheckAllSites();
            bool serviceNameVerif = servicePage.GetFirstServiceName().Contains(serviceName);

            //Assert
            Assert.IsTrue(isGeneralInformationPageLoaded, "Failed to load the General information Service page.");
            Assert.IsTrue(serviceNameVerif, "Le service n'a pas été créé.");
        }

        // Create a service without serviceName
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Create_New_Service_without_ServiceName()
        {
            // Prepare
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);

            //Log in
            var homePage = LogInAsAdmin();

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();

            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage("", serviceCode, serviceProduction);
            serviceCreateModalPage.Create();

            bool errorMessage = serviceCreateModalPage.GetError();
            //Assert
            Assert.IsTrue(errorMessage, "Le service a pu être créé malgré qu'il n'ait pas de nom.");
        }

        // Create a service already exist
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Create_New_Service_Already_Exist()
        {
            // Prepare          
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);

            //Log in
            var homePage = LogInAsAdmin();

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();

            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();
            servicePage = serviceGeneralInformationsPage.BackToList();

            serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            serviceCreateModalPage.Create();

            bool errorMessage = serviceCreateModalPage.GetError();
            //Assert
            Assert.IsTrue(errorMessage, "Le service a pu être créé malgré qu'il existe déjà un service du même nom.");
        }

        // Create price
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Add_Price()
        {
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            // Prepare
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();

            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10));
            pricePage.UnfoldAll();

            //Assert
            var isPriceVisible= pricePage.IsPriceVisible();
            Assert.IsTrue(isPriceVisible, "Le prix n'a pas été créé pour le service.");
        }

        // Create price already Exists
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Add_Price_Already_Exist()
        {
            // Prepare
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);

            string error = "A price already exists for this period";
            DateTime fromDate = DateUtils.Now;
            DateTime toDate = DateUtils.Now.AddDays(10);

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();

            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, fromDate, toDate);

            priceModalPage = pricePage.AddNewCustomerPrice();
            priceModalPage.FillFields_CustomerPrice(site, customer, fromDate, toDate);

            //Assert
            var errorPrice= priceModalPage.GetErrorPriceAlreadyExist();
            Assert.AreEqual(error, errorPrice, "Le prix a pu être ajouté malgré le fait qu'il existe déjà.");
        }

        // Show price detail
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Show_Price_Detail()
        {
            // Prepare
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();

            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10));

            pricePage.UnfoldAll();
            Assert.IsTrue(pricePage.IsUnfoldAll(), "Les détails des prix ne sont pas affichés.");
        }

        // hide price detail
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Hide_Price_Detail()
        {
            // Prepare
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();

            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10));

            pricePage.UnfoldAll();
            Assert.IsTrue(pricePage.IsUnfoldAll(), "Les détails des prix ne sont pas affichés.");

            pricePage.FoldAll();
            Assert.IsTrue(pricePage.IsFoldAll(), "Les détails des prix ne sont pas masqués.");
        }

        // Duplicate price 
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Duplicate_Price()
        {
            // Prepare
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string site1 = TestContext.Properties["SiteLP"].ToString();

            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();

            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10));

            pricePage.DuplicateAllPrices(site1, customer);

            //Assert
            var isPriceDuplicated= pricePage.IsPriceDuplicated("", site1);
            Assert.IsTrue(isPriceDuplicated, "Le prix n'a pas été dupliqué.");
        }

        // Show price detail and modification
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Modification_Price_Detail()
        {
            // Prepare
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string customer1 = TestContext.Properties["CustomerLPFlight"].ToString();
            string site1 = TestContext.Properties["SitePriceList"].ToString();

            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);

            DateTime fromDate = DateUtils.Now;
            DateTime toDate = DateUtils.Now.AddDays(10);

            DateTime fromDate1 = DateUtils.Now.AddDays(11);
            DateTime toDate1 = DateUtils.Now.AddDays(20);

            string dateFromUpdate = fromDate1.ToString("yyyy-MM-dd");
            string dateToUpdate = toDate1.ToString("yyyy-MM-dd");

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();

            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, fromDate, toDate);

            pricePage.UnfoldAll();
            priceModalPage = pricePage.EditFirstPrice(site, "");
            pricePage = priceModalPage.FillFields_EditCustomerPrice(site1, customer1, fromDate1, toDate1);

            WebDriver.Navigate().Refresh();
            pricePage.UnfoldAll();

            var priceName = pricePage.GetPriceName();
            var priceDateFrom = pricePage.GetPriceDateFrom();
            var priceDateTo = pricePage.GetPriceDateTo();

            // Assert
            Assert.IsTrue(priceName.Contains(customer1), "Le prix n'a pas été modifié (customer).");
            Assert.IsTrue(priceName.Contains(site1), "Le prix n'a pas été modifié (site).");
            Assert.IsTrue(priceDateFrom.Contains(dateFromUpdate), "Le prix n'a pas été modifié (dateFrom).");
            Assert.IsTrue(priceDateTo.Contains(dateToUpdate), "Le prix n'a pas été modifié (dateTo).");
        }

        // Delete price 
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Delete_Price()
        {
            // Prepare
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);

            DateTime fromDate = DateUtils.Now;
            DateTime toDate = DateUtils.Now.AddDays(10);

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();

            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, fromDate, toDate);

            pricePage.UnfoldAll();
            pricePage.DeleteFirstPrice();

            Assert.IsFalse(pricePage.IsPriceVisible(), "Le prix n'a pas été supprimé pour le service.");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Menu()
        {
            // Prepare
            string customer = TestContext.Properties["MenuServiceCustomer"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();

            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);

            DateTime fromDate = DateUtils.Now;
            DateTime toDate = DateUtils.Now.AddDays(10);

            string menuName = GenerateName(10);
            DateTime startDate = DateUtils.Now;
            DateTime endDate = DateUtils.Now.AddDays(10);

            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();

            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            priceModalPage.FillFields_CustomerPrice(site, customer, fromDate, toDate);
            //Create menu
            var menuPage = homePage.GoToMenus_Menus();
            var menuCreateModalPage = menuPage.MenuCreatePage();
            menuCreateModalPage.FillField_CreateNewMenu(menuName, startDate, endDate, site, variant, serviceName);

            servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);

            pricePage = servicePage.ClickOnFirstService();
            var serviceMenuPage = pricePage.GoToMenusTab();

            string NameMenu = serviceMenuPage.GetNameMenu();
            //Assert
            Assert.AreEqual(NameMenu, menuName, "Le menu n'a pas été associé au service.");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Delivery()
        {
            // Prepare
            string customer = TestContext.Properties["MenuServiceCustomer"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);

            DateTime fromDate = DateUtils.Now;
            DateTime toDate = DateUtils.Now.AddDays(10);

            string deliveryName = GenerateName(10);

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();

            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            priceModalPage.FillFields_CustomerPrice(site, customer, fromDate, toDate);

            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
            deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, customer, site, true);
            var loadingPage = deliveryCreateModalPage.Create();
            loadingPage.AddService(serviceName);

            servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);

            pricePage = servicePage.ClickOnFirstService();
            var serviceDeliveryPage = pricePage.GoToDeliveryTab();

            //Assert
            string DeliveryNameNew = serviceDeliveryPage.GetDeliveryName();
            Assert.AreEqual(DeliveryNameNew, deliveryName, "Le deliveray n'a pas été associé au service.");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Loading_Plan()
        {
            // Prepare
            string site = TestContext.Properties["SiteACE"].ToString();
            var aircraft = TestContext.Properties["Aircraft"].ToString();
            var route = TestContext.Properties["RouteLP"].ToString();

            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);

            DateTime fromDate = DateUtils.Now.AddMonths(-1);
            DateTime toDate = DateUtils.Now.AddMonths(3);

            string loadingPlanName = GenerateName(10);
            string guestName = "YC";

            // Arrange
            HomePage homePage= LogInAsAdmin();
        

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            try
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                priceModalPage.FillFields_CustomerPrice(site, CUSTOMER, fromDate, toDate);

                // Act
                var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, CUSTOMER, route, aircraft, site);
                var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();

                loadingPlanDetailsPage.ClickAddGuestBtn();
                loadingPlanDetailsPage.SelectGuest(guestName);
                loadingPlanDetailsPage.ClickCreateGuestBtn();

                loadingPlanDetailsPage.ClickGuestBtn();
                loadingPlanDetailsPage.AddServiceBtn();
                loadingPlanDetailsPage.AddNewService(serviceName);
                servicePage = homePage.GoToCustomers_ServicePage();
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
                pricePage = servicePage.ClickOnFirstService();

                var serviceLoadingPage = pricePage.GoToLoadingPlanTab();
                var loadingPlanName1 = serviceLoadingPage.GetLoadingPlanName();

                //Assert
                Assert.AreEqual(loadingPlanName1, loadingPlanName, "Le loading plan n'a pas été relié au service.");
            }
            finally
            {
                servicePage = homePage.GoToCustomers_ServicePage();
                var serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();
                servicePage.ResetFilters();
            }
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Unfold_All()
        {
            // Prepare
            string customer = TestContext.Properties["CustomerLPFlight"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);

            DateTime fromDate = DateUtils.Now;
            DateTime toDate = DateUtils.Now.AddDays(10);

            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();

            if (servicePage.CheckTotalNumber() < 20)
            {
                var ServiceCreateModalPage = servicePage.ServiceCreatePage();
                ServiceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
                var serviceGeneralInformationsPage = ServiceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, fromDate, toDate);
                servicePage = pricePage.BackToList();
                servicePage.ResetFilters();
            }

            servicePage.UnfoldAll();

            //Assert
            Assert.IsTrue(servicePage.IsUnfoldAll(), "Les données ne sont pas affichées en détail.");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_fold_All()
        {
            // Prepare
            string customer = TestContext.Properties["CustomerLPFlight"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);

            DateTime fromDate = DateUtils.Now;
            DateTime toDate = DateUtils.Now.AddDays(10);

            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();

            if (servicePage.CheckTotalNumber() < 20)
            {
                var ServiceCreateModalPage = servicePage.ServiceCreatePage();
                ServiceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
                var serviceGeneralInformationsPage = ServiceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, fromDate, toDate);
                servicePage = pricePage.BackToList();
                servicePage.ResetFilters();
            }

            servicePage.UnfoldAll();
            Assert.IsTrue(servicePage.IsUnfoldAll(), "Les données ne sont pas affichées en détail.");

            servicePage.FoldAll();
            Assert.IsTrue(servicePage.IsFoldAll(), "Les données ne sont pas masquées.");
        }

        //Export Prices
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Export_Prices_NEW_VERSION()
        {
            // Prepare
            string customer = TestContext.Properties["CustomerLPFlight"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);

            bool newVersionPrint = true;

            HomePage homePage = LogInAsAdmin();

            // Act
            homePage.ClearDownloads();

            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();



            var ServiceCreateModalPage = servicePage.ServiceCreatePage();
            ServiceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = ServiceCreateModalPage.Create();

            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10));
            servicePage = pricePage.BackToList();
            servicePage.ResetFilters();

            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.Export(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber(SERVICES_EXCEL_SHEET_NAME, filePath);

            //Assert
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);

            correctDownloadedFile.MoveTo(Path.Combine(downloadsPath, fileName + "_first.xlsx"));

            servicePage.ClearDownloads();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Customers, "0004C - AIR CPU SL");
            servicePage.Filter(ServicePage.FilterType.ShowOnlyActive, true);
            servicePage.PageSize("100");
            servicePage.PageUp();

            servicePage.Export(newVersionPrint);
            // On récupère les fichiers du répertoire de téléchargement
            taskDirectory = new DirectoryInfo(downloadsPath);
            taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            Assert.IsNotNull(correctDownloadedFile);

            servicePage.CheckExport(correctDownloadedFile);
        }

        // Print Cycle Calendars
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Print_Cycle_Calendars_NEW_VERSION()
        {
            // Prepare
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);

            bool newVersionPrint = true;

            HomePage homePage=LogInAsAdmin();
         
            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();

            servicePage.ClearDownloads();

            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(2));
            servicePage = pricePage.BackToList();
            servicePage.ResetFilters();

            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.CLearDownloadsFolder(downloadsPath);
            servicePage.PrintCyclesForCalendar(newVersionPrint, DateUtils.Now, DateUtils.Now.AddDays(2));

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();
            FileInfo correctDownloadedFile = taskFiles.FirstOrDefault();
            correctDownloadedFile = servicePage.GetExcelFile(taskFiles, false);

            // Vérification que le fichier a été trouvé
            Assert.IsNotNull(correctDownloadedFile, "Le fichier correct n'a pas été trouvé.");

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber(SERVICES_PRINT_SHEET_NAME, filePath);

            //Assert
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);
        }

        //[Ignore]
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_Action_Add_Price()
        {

            // Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            bool newVersionPrint = true;

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();

            servicePage.ClearDownloads();

            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10));
            //verification on price creation 
            pricePage.UnfoldAll();
            if (pricePage.PriceIsVisible() == false)
            {
                pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10));
                pricePage.UnfoldAll();
            }
            servicePage = serviceGeneralInformationsPage.BackToList();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);

            var servivePricePage = servicePage.ClickOnFirstService();
            servivePricePage.UnfoldAll();
            int initialItemPrices = servivePricePage.GetPricesItem();
            pricePage.BackToList();

            servicePage.Export(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string sheetName, string fileName, string value
            OpenXmlExcel.WriteDataInColumn("Action", SHEET1, filePath, "A", CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("From", SHEET1, filePath, DateUtils.Now.AddDays(11).ToString("dd/MM/yyyy"), CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("To", SHEET1, filePath, DateUtils.Now.AddDays(20).ToString("dd/MM/yyyy"), CellValues.SharedString);

            WebDriver.Navigate().Refresh();
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);
            importPopup.Import();
            importPopup.WaitPageLoading();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);

            servicePage.ClickOnFirstService();
            servivePricePage.UnfoldAll();
            int NewItemPrices = servivePricePage.GetPricesItem();

            //ASSERT
            Assert.AreNotEqual(initialItemPrices, NewItemPrices, "La valeur n'a pas été modifiée par l'import.");

        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_Action_Update_Price()
        {

            // Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            bool newVersionPrint = true;

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();

            servicePage.ClearDownloads();
            try
            {
                servicePage.ResetFilters();
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10), "10");
                pricePage.UnfoldAll();
                servicePage = serviceGeneralInformationsPage.BackToList();



                var servivePricePage = servicePage.ClickOnFirstService();
                servivePricePage.UnfoldAll();

                var initialItemPrices = servivePricePage.GetPrice();

                pricePage.BackToList();

                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
                servicePage.Filter(ServicePage.FilterType.ShowAll, true);

                servicePage.Export(newVersionPrint);

                // On récupère les fichiers du répertoire de téléchargement
                DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
                FileInfo[] taskFiles = taskDirectory.GetFiles();

                // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
                var correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
                Assert.IsNotNull(correctDownloadedFile);

                var fileName = correctDownloadedFile.Name;
                var filePath = Path.Combine(downloadsPath, fileName);

                //string columnName, string sheetName, string fileName, string value
                OpenXmlExcel.WriteDataInColumn("Action", SHEET1, filePath, "M", CellValues.SharedString);
                OpenXmlExcel.WriteDataInColumn("Price", SHEET1, filePath, "50", CellValues.Number);

                WebDriver.Navigate().Refresh();
                var importPopup = servicePage.Import();
                importPopup.CheckFile(correctDownloadedFile.FullName);
                importPopup.Import();

                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
                servicePage.Filter(ServicePage.FilterType.ShowAll, true);

                servicePage.ClickOnFirstService();
                servivePricePage.UnfoldAll();
                servivePricePage.WaitPageLoading();
                var NewItemPrices = servivePricePage.GetPrice();

                //ASSERT
                Assert.AreNotEqual(initialItemPrices, NewItemPrices, "La valeur n'a pas été modifiée par l'import.");
            }
            finally
            {
                servicePage = homePage.GoToCustomers_ServicePage();
                var serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();
                servicePage.ResetFilters();
            }
        }


        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_Action_Error()
        {

            // Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            bool newVersionPrint = true;

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();

            servicePage.ClearDownloads();

            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10), "10");
            pricePage.UnfoldAll();
            servicePage = serviceGeneralInformationsPage.BackToList();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);

            var servivePricePage = servicePage.ClickOnFirstService();
            servivePricePage.UnfoldAll();

            servivePricePage.GetPrice();
            pricePage.BackToList();

            servicePage.Export(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string sheetName, string fileName, string value
            OpenXmlExcel.WriteDataInColumn("Action", SHEET1, filePath, "", CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Service", SHEET1, filePath, "", CellValues.Number);
            OpenXmlExcel.WriteDataInColumn("ServiceId", SHEET1, filePath, "", CellValues.Number);

            WebDriver.Navigate().Refresh();
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);
            var isNotError = importPopup.ImportWithFail();

            ////ASSERT
            Assert.IsFalse(isNotError, "l'import est executé malgré que le fichier est corrompu.");

        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_With_X_Action()
        {
            // Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            bool newVersionPrint = true;

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();

            servicePage.ClearDownloads();

            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10), "10");
            pricePage.UnfoldAll();
            servicePage = serviceGeneralInformationsPage.BackToList();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);

            var servivePricePage = servicePage.ClickOnFirstService();
            servivePricePage.UnfoldAll();

            var initialItemPrices = servivePricePage.GetPrice();
            pricePage.BackToList();

            servicePage.Export(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string sheetName, string fileName, string value
            OpenXmlExcel.WriteDataInColumn("Action", SHEET1, filePath, "X", CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Price", SHEET1, filePath, "50", CellValues.Number);

            WebDriver.Navigate().Refresh();
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);
            bool isImport = importPopup.ImportWithFail();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);

            servicePage.ClickOnFirstService();
            servivePricePage.UnfoldAll();
            var NewItemPrices = servivePricePage.GetPrice();

            ////ASSERT
            Assert.AreEqual(initialItemPrices.Replace("Price: ", ""), NewItemPrices.Replace("Price: ", ""), "La valeur n'a pas été modifiée par l'import.");
            Assert.IsFalse(isImport, "Le message d'erreur ne s'est pas affiché.");

        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_With_T_Action()
        {
            // Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            bool newVersionPrint = true;

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();

            servicePage.ClearDownloads();

            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10), "10");
            pricePage.UnfoldAll();
            servicePage = serviceGeneralInformationsPage.BackToList();

            //clear downloads directory
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            foreach (var file in taskDirectory.GetFiles())
            {
                file.Delete();
            }


            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);


            servicePage.Export(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string sheetName, string fileName, string value
            OpenXmlExcel.WriteDataInColumn("Action", SHEET1, filePath, "T", CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Price", SHEET1, filePath, "50", CellValues.Number);

            WebDriver.Navigate().Refresh();
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);
            bool isImport = importPopup.ImportWithFail();
            Assert.IsFalse(isImport, "Le message d'erreur ne s'est pas affiché.");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_With_A_ActionCustomerCode_Error()
        {
            // Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            bool newVersionPrint = true;

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();

            servicePage.ClearDownloads();

            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();


            var pricePage = serviceGeneralInformationsPage.GoToPricePage();

            var priceModalPage = pricePage.AddNewCustomerPrice();
            pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10), "10");

            servicePage = serviceGeneralInformationsPage.BackToList();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);

            var servivePricePage = servicePage.ClickOnFirstService();
            servivePricePage.UnfoldAll();

            servivePricePage.GetPrice();
            pricePage.BackToList();

            servicePage.Export(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string sheetName, string fileName, string value
            OpenXmlExcel.WriteDataInColumn("Action", SHEET1, filePath, "A", CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Customer Code", SHEET1, filePath, "", CellValues.SharedString);

            WebDriver.Navigate().Refresh();
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);
            bool isImport = importPopup.ImportModal();

            Assert.IsFalse(isImport, "Le message d'erreur ne s'est pas affiché.");
        }


        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_With_A_ActionCustomerCode_NotValideCodeError()
        {
            // Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            bool newVersionPrint = true;

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();

            servicePage.ClearDownloads();

            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();


            var pricePage = serviceGeneralInformationsPage.GoToPricePage();

            var priceModalPage = pricePage.AddNewCustomerPrice();
            pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10), "10");

            servicePage = serviceGeneralInformationsPage.BackToList();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);

            var servivePricePage = servicePage.ClickOnFirstService();
            servivePricePage.UnfoldAll();

            servivePricePage.GetPrice();
            pricePage.BackToList();

            servicePage.Export(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string sheetName, string fileName, string value
            OpenXmlExcel.WriteDataInColumn("Action", SHEET1, filePath, "A", CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Customer Code", SHEET1, filePath, "TOTO", CellValues.SharedString);

            WebDriver.Navigate().Refresh();
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);
            bool isImport = importPopup.ImportModal();

            Assert.IsFalse(isImport, "Le message d'erreur ne s'est pas affiché.");
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_With_A_ActionCustomerSrv()
        {

            // Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            bool newVersionPrint = true;
            var servicePage = homePage.GoToCustomers_ServicePage();
            try
            {
                // Act
                servicePage.ClearDownloads();

                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10), "10", null, "srv");
                pricePage.UnfoldAll();

                pricePage.ClickFirstItem();
                var initialCustomerSrv = pricePage.GetCustomerSrvFromInput();
                homePage.GoToCustomers_ServicePage();
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
                servicePage.Filter(ServicePage.FilterType.ShowAll, true);
                servicePage.Export(newVersionPrint);

                // On récupère les fichiers du répertoire de téléchargement
                DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
                FileInfo[] taskFiles = taskDirectory.GetFiles();

                // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
                var correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
                Assert.IsNotNull(correctDownloadedFile);

                var fileName = correctDownloadedFile.Name;
                var filePath = Path.Combine(downloadsPath, fileName);

                //string columnName, string sheetName, string fileName, string value
                OpenXmlExcel.WriteDataInColumn("Action", SHEET1, filePath, "A", CellValues.SharedString);
                servicePage.WaitPageLoading();
                OpenXmlExcel.WriteDataInColumn("Customer srv", SHEET1, filePath, "", CellValues.SharedString);
                servicePage.WaitPageLoading();
                OpenXmlExcel.WriteDataInColumn("Site Code", SHEET1, filePath, "ACE", CellValues.SharedString);


                //WebDriver.Navigate().Refresh();
                var importPopup = servicePage.Import();
                importPopup.CheckFile(correctDownloadedFile.FullName);
                importPopup.Import();

                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
                servicePage.Filter(ServicePage.FilterType.ShowAll, true);

                var servivePricePage = servicePage.ClickOnFirstService();
                servivePricePage.UnfoldAll(); ;
                servivePricePage.ClickFirstItem();
                var newCustomerSrv = servivePricePage.GetCustomerSrvFromInput();

                //ASSERT
                Assert.AreNotEqual(initialCustomerSrv, newCustomerSrv, "L'Import a échoué.");

            }
            finally
            {
                servicePage = homePage.GoToCustomers_ServicePage();
                var serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();

            }

        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_With_A_ActionIDCustomerSrvEmpty()
        {


            // Arrange
            LogInAsAdmin();

            // Act
            string newService = GenerateName(8);
            var customerCode = TestContext.Properties["CustomerCode"] as string;
            var servicePage = ServicePage.NavigateTo(WebDriver, TestContext);
            TestContext.Properties["NewServiceName"] = newService;
            // On récupère le fichier dans le répertoire d'exécution
            var path = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase).Substring(6);
            var filePath = path + "\\Customers\\Service\\Test_CU_SERV_Import_With_A_ActionIDCustomerSrvvEmpty.xlsx";

            // On met le nom du nouveau service dans le fichier à importer, ainsi que le nom du datasheet
            OpenXmlExcel.WriteDataInColumn("Service", SHEET1, filePath, newService, CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Customer Code", SHEET1, filePath, customerCode, CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("ID Customer srv", SHEET1, filePath, "", CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Site Code", SHEET1, filePath, "ACE", CellValues.SharedString);

            var importPopup = servicePage.Import();
            importPopup.CheckFile(filePath);
            bool isImport = importPopup.Import();

            // query use for assert :
            var query = "select 1 from ServiceToCustomerToSites where serviceId in (select sv.Id from Services sv where sv.Name = @serviceName )";
            var result = ExecuteAndGetInt(query, new KeyValuePair<string, object>("serviceName", newService));
            //ASSERT
            Assert.AreEqual(1, result, "L'Import a échoué.");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_With_A_ActionIDCustomerSrvExist()
        {
            LogInAsAdmin();

            // Act
            string newService = GenerateName(8);
            var customerCode = TestContext.Properties["CustomerCode"] as string;
            var servicePage = ServicePage.NavigateTo(WebDriver, TestContext);
            TestContext.Properties["NewServiceName"] = newService;
            // On récupère le fichier dans le répertoire d'exécution
            var path = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase).Substring(6);
            var filePath = path + "\\Customers\\Service\\Test_CU_SERV_Import_With_A_ActionIDCustomerSrvExist.xlsx";

            // On met le nom du nouveau service dans le fichier à importer, ainsi que le nom du datasheet
            OpenXmlExcel.WriteDataInColumn("Service", SHEET1, filePath, newService, CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Customer Code", SHEET1, filePath, customerCode, CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("ID Customer srv", SHEET1, filePath, "", CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Site Code", SHEET1, filePath, "ACE", CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Action", SHEET1, filePath, "A", CellValues.SharedString);

            var importPopup = servicePage.Import();
            importPopup.CheckFile(filePath);
            bool isImport = importPopup.Import();

            // query use for assert :
            var query = "select 1 from ServiceToCustomerToSites where serviceId in (select sv.Id from Services sv where sv.Name = @serviceName )";
            var result = ExecuteAndGetInt(query, new KeyValuePair<string, object>("serviceName", newService));
            //ASSERT
            Assert.AreEqual(1, result, "L'Import a échoué.");


        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_With_M_ActionCustomerSrv()
        {

            LogInAsAdmin();

            //// Act
            ////string newService = GenerateName(8);
            var servicePage = ServicePage.NavigateTo(WebDriver, TestContext);
            var newService = TestContext.Properties["ServiceName"].ToString();
            int serviceIDDetail = int.Parse(TestContext.Properties["ServiceDetailId"].ToString());
            int serviceID = int.Parse(TestContext.Properties["ServiceId"].ToString());
            // On récupère le fichier dans le répertoire d'exécution

            var path = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase).Substring(6);
            var filePath = path + "\\Customers\\Service\\Test_CU_SERV_Import_With_M_Action_CustomerSrv.xlsx";
            //Get value from database

            var initialCustomerSrv = GetCustomerSrv(serviceIDDetail);

            OpenXmlExcel.WriteDataInColumn("ServiceId", SHEET1, filePath, serviceID.ToString(), CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("PriceId", SHEET1, filePath, serviceIDDetail.ToString(), CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Service", SHEET1, filePath, newService, CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Action", SHEET1, filePath, "M", CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Customer srv", SHEET1, filePath, "newSrv", CellValues.SharedString);


            servicePage.WaitPageLoading();
            var importPopup = servicePage.Import();
            importPopup.CheckFile(filePath);
            importPopup.Import();


            string newCustomerSrv = GetCustomerSrv(serviceIDDetail);

            //ASSERT
            Assert.AreNotEqual(initialCustomerSrv, newCustomerSrv, "L'Import a échoué.");


        }
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_With_M_ActionDatasheet()
        {

            // Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string datasheetName1 = TestContext.Properties["DatasheetName1"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            bool newVersionPrint = true;

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();

            servicePage.ClearDownloads();

            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10), "10", datasheetName1, "srv");
            pricePage.UnfoldAll();
            servicePage = serviceGeneralInformationsPage.BackToList();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);

            var servivePricePage = servicePage.ClickOnFirstService();
            servivePricePage.UnfoldAll();

            var initialDatasheet = servivePricePage.GetDatasheet();
            pricePage.BackToList();

            servicePage.Export(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string sheetName, string fileName, string value
            OpenXmlExcel.WriteDataInColumn("Action", SHEET1, filePath, "M", CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Datasheet", SHEET1, filePath, "1900 ALMUERZO CALIENTE BC GFML-ALMFCGFML DIC'14 (MAD)", CellValues.SharedString);



            WebDriver.Navigate().Refresh();
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);
            importPopup.Import();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);

            servicePage.ClickOnFirstService();
            servivePricePage.UnfoldAll();
            var newDatasheet = servivePricePage.GetDatasheet();

            //ASSERT
            Assert.AreNotEqual(initialDatasheet, newDatasheet, "L'Import a échoué.");
        }


        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_With_M_ActionDatasheet_Erase()
        {

            // Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string datasheetName1 = TestContext.Properties["DatasheetName1"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            bool newVersionPrint = true;

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            try
            {

                servicePage.ClearDownloads();

                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10), "10", datasheetName1, "srv");
                pricePage.UnfoldAll();
                servicePage = serviceGeneralInformationsPage.BackToList();

                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
                servicePage.Filter(ServicePage.FilterType.ShowAll, true);

                var servivePricePage = servicePage.ClickOnFirstService();
                servivePricePage.UnfoldAll();
                pricePage.ClickFirstItem();
                var initialDatasheet = servivePricePage.GetDatasheetValue();
                homePage.GoToCustomers_ServicePage();
                if (servicePage.CheckTotalNumber() >= 1)
                {
                    servicePage.Filter(ServicePage.FilterType.Search, serviceName);
                    servicePage.Filter(ServicePage.FilterType.ShowAll, true);
                }

                servicePage.Export(newVersionPrint);

                // On récupère les fichiers du répertoire de téléchargement
                DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
                FileInfo[] taskFiles = taskDirectory.GetFiles();

                // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
                var correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
                Assert.IsNotNull(correctDownloadedFile);

                var fileName = correctDownloadedFile.Name;
                var filePath = Path.Combine(downloadsPath, fileName);

                //string columnName, string sheetName, string fileName, string value
                OpenXmlExcel.WriteDataInColumn("Action", SHEET1, filePath, "M", CellValues.SharedString);
                OpenXmlExcel.WriteDataInColumn("Datasheet", SHEET1, filePath, "", CellValues.SharedString);



                WebDriver.Navigate().Refresh();
                var importPopup = servicePage.Import();
                importPopup.CheckFile(correctDownloadedFile.FullName);
                importPopup.Import();

                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
                servicePage.Filter(ServicePage.FilterType.ShowAll, true);

                servicePage.ClickOnFirstService();
                servivePricePage.UnfoldAll();
                pricePage.ClickFirstItem();
                var newDatasheet = servivePricePage.GetDatasheetValue();

                //ASSERT
                Assert.AreNotEqual(initialDatasheet, newDatasheet, "L'Import a échoué.");
                Assert.IsTrue(String.IsNullOrEmpty(newDatasheet), "L'Import a échoué et n'a pas supprimé la datasheet.");
            }
            finally
            {
                servicePage = homePage.GoToCustomers_ServicePage();
                var serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();
                servicePage.ResetFilters();
            }
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_With_M_ActionDatasheetError()
        {

            // Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string datasheetName1 = TestContext.Properties["DatasheetName1"].ToString();
            string datasheet = "datasheet" + DateTime.Now.ToString("dd/MM/yyyy");
            //arrange

            HomePage homePage = LogInAsAdmin();


            bool newVersionPrint = true;

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();

            servicePage.ClearDownloads();

            var serviceCreateModalPage = servicePage.ServiceCreate();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10), "10", datasheetName1, "srv");
            pricePage.UnfoldAll();
            servicePage = serviceGeneralInformationsPage.BackToList();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);

            var servivePricePage = servicePage.ClickOnFirstService();
            servivePricePage.UnfoldAll();

            pricePage.BackToList();

            servicePage.Export(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string sheetName, string fileName, string value
            OpenXmlExcel.WriteDataInColumn("Action", SHEET1, filePath, "M", CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Datasheet", SHEET1, filePath, datasheet, CellValues.SharedString);



            WebDriver.Navigate().Refresh();
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);
            var isImport = importPopup.ImportModal();

            //ASSERT
            Assert.IsFalse(isImport, "L'Import s'est faite malgrès une datasheet inexistante.");
        }


        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_With_A_ActionNoDatasheet()
        {

            // Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string datasheetName1 = TestContext.Properties["DatasheetName1"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            bool newVersionPrint = true;

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();

            servicePage.ClearDownloads();

            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10), "10", datasheetName1, "srv");
            pricePage.UnfoldAll();
            servicePage = serviceGeneralInformationsPage.BackToList();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);

            var servivePricePage = servicePage.ClickOnFirstService();
            servivePricePage.UnfoldAll();

            var initialDatasheet = servivePricePage.GetDatasheet();
            pricePage.BackToList();

            servicePage.Export(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string sheetName, string fileName, string value
            OpenXmlExcel.WriteDataInColumn("Action", SHEET1, filePath, "A", CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Datasheet", SHEET1, filePath, "", CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Site Code", SHEET1, filePath, "ACE", CellValues.SharedString);


            WebDriver.Navigate().Refresh();
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);
            importPopup.Import();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);

            servicePage.ClickOnFirstService();
            servivePricePage.UnfoldAll();
            var newDatasheet = servivePricePage.GetDatasheet();

            //ASSERT
            Assert.AreNotEqual(initialDatasheet, newDatasheet, "L'Import a échoué.");
        }


        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_With_A_ActionAddDatasheet()
        {
            // Arrange
            LogInAsAdmin();

            // Act
            string newService = GenerateName(8);
            var customerCode = TestContext.Properties["CustomerCode"] as string;
            var servicePage = ServicePage.NavigateTo(WebDriver, TestContext);
            TestContext.Properties["NewServiceName"] = newService;
            // On récupère le fichier dans le répertoire d'exécution
            var path = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase).Substring(6);
            var filePath = path + "\\Customers\\Service\\Test_CU_SERV_Import_With_A_ActionAddDatasheet.xlsx";

            var datashhetName = "1900 ALMUERZO CALIENTE BC GFML-ALMFCGFML DIC'14 (MAD)";

            // On met le nom du nouveau service dans le fichier à importer, ainsi que le nom du datasheet
            OpenXmlExcel.WriteDataInColumn("Service", SHEET1, filePath, newService, CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Customer Code", SHEET1, filePath, customerCode, CellValues.SharedString);

            var importPopup = servicePage.Import();
            importPopup.CheckFile(filePath);
            bool isImport = importPopup.Import();

            // query use for assert :
            var query = "select 1 from ServiceToCustomerToSites where serviceId in (select sv.Id from Services sv where sv.Name = @serviceName) and DatasheetId in (select ds.Id from DataSheets ds where ds.Name = @datasheetName)";
            var result = ExecuteAndGetInt(query, new KeyValuePair<string, object>("serviceName", newService), new KeyValuePair<string, object>("datasheetName", datashhetName));
            //ASSERT
            Assert.AreEqual(1, result, "L'Import a échoué.");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_With_A_ActionAddDatasheetNotExist()
        {
            // Arrange
            LogInAsAdmin();

            // Act
            string newService = GenerateName(8);
            string notExistdatasheet = Guid.NewGuid().ToString();
            var customerCode = TestContext.Properties["CustomerCode"] as string;
            var servicePage = ServicePage.NavigateTo(WebDriver, TestContext);
            TestContext.Properties["NewServiceName"] = newService;
            // On récupère le fichier dans le répertoire d'exécution
            var path = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase).Substring(6);
            var filePath = path + "\\Customers\\Service\\Test_CU_SERV_Import_With_A_ActionAddDatasheetNotExist.xlsx";

            // On met le nom du nouveau service dans le fichier à importer, ainsi que le nom du datasheet
            OpenXmlExcel.WriteDataInColumn("Service", SHEET1, filePath, newService, CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Customer Code", SHEET1, filePath, customerCode, CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Datasheet", SHEET1, filePath, notExistdatasheet, CellValues.SharedString);

            var importPopup = servicePage.Import();
            importPopup.CheckFile(filePath);
            var isImport = importPopup.Import();

            //ASSERT
            Assert.IsFalse(isImport, "L'Import s'est fait malgrès une datasheet inexistante en base.");
        }


        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_With_M_ActionCustomerCode()
        {
            // Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            bool newVersionPrint = true;

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();

            servicePage.ClearDownloads();

            var serviceCreateModalPage = servicePage.ServiceCreate();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();


            var pricePage = serviceGeneralInformationsPage.GoToPricePage();

            var priceModalPage = pricePage.AddNewCustomerPrice();
            pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10), "10");

            servicePage = serviceGeneralInformationsPage.BackToList();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);

            var servivePricePage = servicePage.ClickOnFirstService();
            servivePricePage.UnfoldAll();

            servivePricePage.GetPrice();
            pricePage.BackToList();

            servicePage.Export(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string sheetName, string fileName, string value
            OpenXmlExcel.WriteDataInColumn("Action", SHEET1, filePath, "M", CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Customer Code", SHEET1, filePath, "TVS", CellValues.SharedString);

            WebDriver.Navigate().Refresh();
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);
            bool isImport = importPopup.ImportModal();

            Assert.IsTrue(isImport, "Le message d'erreur s'est affiché.");
        }



        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_With_M_ActionCategoryManyLine_Error()
        {
            // Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            DeleteAllFileDownload();
            homePage.ClearDownloads();
            bool newVersionPrint = true;

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();

            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

            var pricePage = serviceGeneralInformationsPage.GoToPricePage();

            var priceModalPage = pricePage.AddNewCustomerPrice();
            pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10), "10");

            pricePage.AddNewCustomerPrice();
            priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.AddDays(11), DateUtils.Now.AddDays(20), "10");
            pricePage.UnfoldAll();

            pricePage.AddNewCustomerPrice();
            priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.AddDays(21), DateUtils.Now.AddDays(30), "10");

            servicePage = serviceGeneralInformationsPage.BackToList();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);

            servicePage.Export(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string sheetName, string fileName, string value
            OpenXmlExcel.WriteDataInColumn("Action", SHEET1, filePath, "M", CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Category", SHEET1, filePath, "TOTO", CellValues.SharedString);


            WebDriver.Navigate().Refresh();
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);
            bool isImport = importPopup.ImportWithFail();

            Assert.IsFalse(isImport, "Le message d'erreur ne s'est pas affiché.");
        }


        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_With_M_ActionCategoryManyLine()
        {
            // Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            Random rnd = new Random((int)DateUtils.Now.Ticks);
            string serviceName = rnd.Next().ToString();
            string serviceCode = rnd.Next().ToString();
            string serviceProduction = GenerateName(4);
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            bool newVersionPrint = true;

            // Act
            ServicePage servicePage = homePage.GoToCustomers_ServicePage();

            ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();


            ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();

            var priceModalPage = pricePage.AddNewCustomerPrice();
            pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10), "10");


            pricePage.AddNewCustomerPrice();
            priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.AddDays(11), DateUtils.Now.AddDays(20), "10");
            pricePage.UnfoldAll();

            pricePage.AddNewCustomerPrice();
            priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.AddDays(21), DateUtils.Now.AddDays(30), "10");

            servicePage = serviceGeneralInformationsPage.BackToList();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);

            var servivePricePage = servicePage.ClickOnFirstService();
            servivePricePage.UnfoldAll();

            pricePage.BackToList();

            servicePage.Export(newVersionPrint);



            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string sheetName, string fileName, string value
            OpenXmlExcel.WriteDataInColumn("Action", SHEET1, filePath, "M", CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Category", SHEET1, filePath, "BEBIDAS CALIENTES", CellValues.SharedString);


            WebDriver.Navigate().Refresh();
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);
            bool isImport = importPopup.Import();

            Assert.IsTrue(isImport, "Le message d'erreur ne s'est pas affiché.");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_With_A_ActionCategoryManyLine()
        {
            LogInAsAdmin();

            // préparation du fichier à importer avec le service créé dans les prérequis
            var serviceId = (int)TestContext.Properties["ServiceId"];
            var serviceName = TestContext.Properties["ServiceName"] as string;
            var customerCode = TestContext.Properties["CustomerCode"] as string;
            var path = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase).Substring(6);
            var filePath = path + "\\Customers\\Service\\Test_CU_SERV_Import_With_A_ActionCategoryManyLine.xlsx";
            OpenXmlExcel.WriteDataInColumn("ServiceId", SHEET1, filePath, serviceId.ToString(), CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Service", SHEET1, filePath, serviceName, CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Customer Code", SHEET1, filePath, customerCode, CellValues.SharedString);

            // déroulement du test
            // Naviguer vers la page des Services
            var servicePage = ServicePage.NavigateTo(WebDriver, TestContext);
            // Cliquer sur Importer
            var importPopup = servicePage.Import();
            // Renseigner et vérifier le fichier à importer
            importPopup.CheckFile(filePath);
            // Exécuter l'import
            bool isImport = importPopup.Import();

            // Assertion de vérification d'import réussi
            Assert.IsTrue(isImport, "L'import a échoué");
            // Récupération en base de données du nombre de lignes de prix importées
            var pricesCount = ExecuteAndGetInt("select count(*) from [dbo].[ServiceToCustomerToSites] where serviceid = @serviceId", new KeyValuePair<string, object>("serviceId", (int)TestContext.Properties["ServiceId"]));
            // Assertion pour vérifier qu'il y a bien 3 lignes importées
            Assert.AreEqual(pricesCount, 3);
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_With_A_ActionCategoryManyLine_Error()
        {
            // Arrange
            LogInAsAdmin();

            // Act
            var serviceId = (int)TestContext.Properties["ServiceId"];
            var serviceName = TestContext.Properties["ServiceName"] as string;
            var customerCode = TestContext.Properties["CustomerCode"] as string;
            var servicePage = ServicePage.NavigateTo(WebDriver, TestContext);

            // On récupère le fichier dans le répertoire d'exécution
            var path = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase).Substring(6);
            var filePath = path + "\\Customers\\Service\\Test_CU_SERV_Import_With_A_ActionCategoryManyLine_Error.xlsx";

            // On change la valeur des colonnes ServiceId et Service avec les données du service créé dans le TestInitialize
            OpenXmlExcel.WriteDataInColumn("ServiceId", SHEET1, filePath, serviceId.ToString(), CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Service", SHEET1, filePath, serviceName, CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Customer Code", SHEET1, filePath, customerCode, CellValues.SharedString);

            var importPopup = servicePage.Import();
            importPopup.CheckFile(filePath);
            bool isImport = importPopup.ImportWithFail();

            Assert.IsFalse(isImport, "L'import est Ok malgrés un fichier corrompu.");
        }


        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_With_A_ActionDefaultLoadingModeNotExist()
        {
            // Arrange
            LogInAsAdmin();

            // Act
            var serviceId = (int)TestContext.Properties["ServiceId"];
            var serviceName = TestContext.Properties["ServiceName"] as string;
            var customerCode = TestContext.Properties["CustomerCode"] as string;
            var servicePage = ServicePage.NavigateTo(WebDriver, TestContext);

            // On récupère le fichier dans le répertoire d'exécution
            var path = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase).Substring(6);
            var filePath = path + "\\Customers\\Service\\Test_CU_SERV_Import_With_A_ActionDefaultLoadingModeNotExist.xlsx";

            // On change la valeur des colonnes ServiceId et Service avec les données du service créé dans le TestInitialize
            OpenXmlExcel.WriteDataInColumn("ServiceId", SHEET1, filePath, serviceId.ToString(), CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Service", SHEET1, filePath, serviceName, CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Customer Code", SHEET1, filePath, customerCode, CellValues.SharedString);

            var importPopup = servicePage.Import();
            importPopup.CheckFile(filePath);
            var isImport = importPopup.ImportModal();

            //ASSERT
            Assert.IsFalse(isImport, "L'Import s'est faite malgrès le Default loading mode vide.");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_With_M_ActionDefaultLoadingMode()
        {

            // Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string datasheetName1 = TestContext.Properties["DatasheetName1"].ToString();


            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            bool newVersionPrint = true;

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();

            servicePage.ClearDownloads();

            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10), "10", datasheetName1, "srv");
            pricePage.UnfoldAll();
            servicePage = serviceGeneralInformationsPage.BackToList();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);

            var servicePricePage = servicePage.ClickOnFirstService();
            servicePricePage.UnfoldAll();

            var editPriceModal = servicePricePage.EditFirstPrice(null, null);
            var oldDefaultLoadingMode = editPriceModal.GetDefaultMode();
            homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);
            servicePage.Export(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string sheetName, string fileName, string value
            OpenXmlExcel.WriteDataInColumn("Action", SHEET1, filePath, "M", CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Default loading mode", SHEET1, filePath, "Standard", CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Default loading round method", SHEET1, filePath, "", CellValues.SharedString);

            WebDriver.Navigate().Refresh();
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);
            var isImport = importPopup.Import();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);

            servicePage.ClickOnFirstService();
            servicePricePage.UnfoldAll();
            servicePricePage.EditFirstPrice(null, null);
            var newDefaultLoadingMode = editPriceModal.GetDefaultMode();

            //ASSERT
            Assert.IsTrue(isImport, "L'Import a échoué.");
            Assert.AreNotEqual(newDefaultLoadingMode, oldDefaultLoadingMode, "Le default loading Mode ne s'est pas mis a jour.");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_With_M_ActionDefaultLoadingModeError()
        {

            // Arrange
            LogInAsAdmin();

            //Act
            string newService = GenerateName(8);
            var customerCode = TestContext.Properties["CustomerCode"] as string;
            var servicePage = ServicePage.NavigateTo(WebDriver, TestContext);
            TestContext.Properties["NewServiceName"] = newService;
            // On récupère le fichier dans le répertoire d'exécution
            var path = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase).Substring(6);
            var filePath = path + "\\Customers\\Service\\Test_CU_SERV_Import_With_M_ActionDefaultLoadingModeError.xlsx";




            //string columnName, string sheetName, string fileName, string value
            // On met le nom du nouveau service dans le fichier à importer, ainsi que le nom du datasheet
            OpenXmlExcel.WriteDataInColumn("Service", SHEET1, filePath, newService, CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Action", SHEET1, filePath, "M", CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Default loading mode", SHEET1, filePath, "Stantard", CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Default loading round method", SHEET1, filePath, "", CellValues.SharedString);
            var importPopup = servicePage.Import();
            importPopup.CheckFile(filePath);
            bool isImport = importPopup.Import();


            //ASSERT
            var query = "select count(*) from ServiceToCustomerToSites where serviceId in (select sv.Id from Services sv where sv.Name = @serviceName)";
            var result = ExecuteAndGetInt(query, new KeyValuePair<string, object>("serviceName", newService));
            Assert.AreEqual(result, 0, "L'Import a été effectué malgrès la mauvaise valeur.");
        }


        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_With_A_ActionDefaultLoadingMode()
        {

            // Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string datasheetName1 = TestContext.Properties["DatasheetName1"].ToString();
            string newService = GenerateName(8);

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            bool newVersionPrint = true;

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();

            servicePage.ClearDownloads();

            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10), "10", datasheetName1, "srv");
            pricePage.UnfoldAll();
            servicePage = serviceGeneralInformationsPage.BackToList();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);

            var servivePricePage = servicePage.ClickOnFirstService();
            servivePricePage.UnfoldAll();

            servivePricePage.GetDatasheet();
            pricePage.BackToList();

            servicePage.Export(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string sheetName, string fileName, string value
            OpenXmlExcel.WriteDataInColumn("Action", SHEET1, filePath, "A", CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Default loading mode", SHEET1, filePath, "Standard", CellValues.SharedString);

            OpenXmlExcel.WriteDataInColumn("Service", SHEET1, filePath, newService, CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("ServiceId", SHEET1, filePath, "", CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("PriceId", SHEET1, filePath, "", CellValues.SharedString);

            WebDriver.Navigate().Refresh();
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);
            var isImport = importPopup.Import();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, newService);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);

            var servicePricePage = servicePage.ClickOnFirstService();
            servicePricePage.UnfoldAll();
            var editPriceModal = servicePricePage.EditFirstPrice(null, null);
            var newDefaultLoadingMode = editPriceModal.GetDefaultMode();

            //ASSERT
            Assert.IsTrue(isImport, "L'Import a échoué.");
            Assert.AreEqual("STD", newDefaultLoadingMode, "Le default loading Mode ne s'est pas mis a jour.");
        }


        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_With_M_ActionDefaultLoadingValueEmptyError()
        {

            // Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string datasheetName1 = TestContext.Properties["DatasheetName1"].ToString();


            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            bool newVersionPrint = true;

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();

            servicePage.ClearDownloads();

            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10), "10", datasheetName1, "srv");
            pricePage.UnfoldAll();
            servicePage = serviceGeneralInformationsPage.BackToList();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);

            var servicePricePage = servicePage.ClickOnFirstService();
            servicePricePage.UnfoldAll();

            var editPriceModal = servicePricePage.EditFirstPrice(null, null);
            editPriceModal.GetDefaultModeValue();
            servicePricePage.BackToList();

            servicePage.Export(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string sheetName, string fileName, string value
            OpenXmlExcel.WriteDataInColumn("Action", SHEET1, filePath, "M", CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Default loading value", SHEET1, filePath, "", CellValues.SharedString);


            WebDriver.Navigate().Refresh();
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);
            var isImport = importPopup.ImportModal();

            //ASSERT
            Assert.IsFalse(isImport, "L'Import a échoué.");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_SERV_Import_With_M_ActionDefaultLoadingValue()
        {

            // Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string datasheetName1 = TestContext.Properties["DatasheetName1"].ToString();


            // Arrange
            HomePage homePage = LogInAsAdmin();

            bool newVersionPrint = true;

            // Act
            ServicePage servicePage = homePage.GoToCustomers_ServicePage();

            servicePage.ClearDownloads();

            ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

            ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10), "10", datasheetName1, "srv");
            pricePage.UnfoldAll();
            servicePage = serviceGeneralInformationsPage.BackToList();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);

            var servicePricePage = servicePage.ClickOnFirstService();
            servicePricePage.UnfoldAll();

            var editPriceModal = servicePricePage.EditFirstPrice(null, null);
            var oldDefaultLoadingModeValue = editPriceModal.GetDefaultModeValue();
            homePage.GoToCustomers_ServicePage();
            if (servicePage.CheckTotalNumber() >= 1)
            {
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
                servicePage.Filter(ServicePage.FilterType.ShowAll, true);
            }
            servicePage.Export(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string sheetName, string fileName, string value
            OpenXmlExcel.WriteDataInColumn("Action", SHEET1, filePath, "M", CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Default loading value", SHEET1, filePath, "200", CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Default loading round method", SHEET1, filePath, "", CellValues.SharedString);

            WebDriver.Navigate().Refresh();
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);
            var isImport = importPopup.Import();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);

            servicePage.ClickOnFirstService();
            servicePricePage.UnfoldAll();
            servicePricePage.EditFirstPrice(null, null);
            var newDefaultLoadingModeValue = editPriceModal.GetDefaultModeValue();

            //ASSERT
            Assert.IsTrue(isImport, "L'Import a échoué.");
            Assert.AreNotEqual(newDefaultLoadingModeValue, oldDefaultLoadingModeValue, "Le default loading Mode ne s'est pas mis a jour.");
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_With_M_ActionPriceValue()
        {
            // Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string datasheetName1 = TestContext.Properties["DatasheetName1"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            bool newVersionPrint = true;

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            try
            {
                servicePage.ClearDownloads();

                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10), null, datasheetName1, "srv");
                servicePage = serviceGeneralInformationsPage.BackToList();

                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
                servicePage.Filter(ServicePage.FilterType.ShowAll, true);

                var servicePricePage = servicePage.ClickOnFirstService();
                servicePricePage.UnfoldAll();

                // Récupérer le modal et son ancien prix
                servicePricePage.ClickFirstItem();
                var oldPrice = servicePricePage.GetPrice();

                homePage.GoToCustomers_ServicePage();

                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
                servicePage.Filter(ServicePage.FilterType.ShowAll, true);
                servicePage.Export(newVersionPrint);

                // On récupère les fichiers du répertoire de téléchargement
                DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
                FileInfo[] taskFiles = taskDirectory.GetFiles();

                // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
                var correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
                Assert.IsNotNull(correctDownloadedFile);

                var fileName = correctDownloadedFile.Name;
                var filePath = Path.Combine(downloadsPath, fileName);

                //string columnName, string sheetName, string fileName, string value
                servicePage.WaitLoading();
                OpenXmlExcel.WriteDataInColumn("Action", SHEET1, filePath, "M", CellValues.SharedString);
                servicePage.WaitLoading();
                OpenXmlExcel.WriteDataInColumn("Price", SHEET1, filePath, "20", CellValues.String);

                WebDriver.Navigate().Refresh();

                // Importer le fichier modifié
                var importPopup = servicePage.Import();
                importPopup.CheckFile(correctDownloadedFile.FullName);
                var isImport = importPopup.Import();

                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
                servicePage.Filter(ServicePage.FilterType.ShowAll, true);

                servicePage.ClickOnFirstService();
                servicePricePage.UnfoldAll();

                // Récupérer à nouveau le modal et le nouveau prix après modification
                servicePricePage.ClickFirstItem();
                var newPrice = servicePricePage.GetPrice();

                // ASSERTIONS
                Assert.IsTrue(isImport, "L'Import a échoué.");
                Assert.AreNotEqual(newPrice, oldPrice, "Le price ne s'est pas mis à jour.");
            }
            finally
            {
                servicePage = homePage.GoToCustomers_ServicePage();
                var serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();
                servicePage.ResetFilters();
            }
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_With_M_ActionPriceAnotherValue()
        {

            // Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            Random rnd = new Random((int)DateUtils.Now.Ticks);
            string serviceName = rnd.Next().ToString();
            string serviceCode = rnd.Next().ToString();
            string serviceProduction = GenerateName(4);
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string datasheetName1 = TestContext.Properties["DatasheetName1"].ToString();


            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            bool newVersionPrint = true;

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();

            servicePage.ClearDownloads();

            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10), null, datasheetName1, "srv");
            servicePage = serviceGeneralInformationsPage.BackToList();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);

            servicePage.Export(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string sheetName, string fileName, string value
            OpenXmlExcel.WriteDataInColumn("Action", SHEET1, filePath, "M", CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Price", SHEET1, filePath, "fhdf", CellValues.String);

            WebDriver.Navigate().Refresh();
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);
            var isImport = importPopup.ImportModal();

            //ASSERT
            Assert.IsFalse(isImport, "L'Import a été éffectué malgrès la mauvaise valeur du price.");
        }


        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_With_M_ActionPriceEmptyValue()
        {

            // Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string datasheetName1 = TestContext.Properties["DatasheetName1"].ToString();


            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            bool newVersionPrint = true;

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();

            servicePage.ClearDownloads();

            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10), null, datasheetName1, "srv");
            servicePage = serviceGeneralInformationsPage.BackToList();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);

            servicePage.Export(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string sheetName, string fileName, string value
            OpenXmlExcel.WriteDataInColumn("Action", SHEET1, filePath, "M", CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Price", SHEET1, filePath, "", CellValues.String);

            WebDriver.Navigate().Refresh();
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);
            var isImport = importPopup.ImportModal();

            //ASSERT
            Assert.IsFalse(isImport, "L'Import a été éffectué malgrèes la mauvaise valeur du price.");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_With_M_ActionDefaultLoadingValueError()
        {

            // Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string datasheetName1 = TestContext.Properties["DatasheetName1"].ToString();


            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            bool newVersionPrint = true;

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();

            servicePage.ClearDownloads();

            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10), "10", datasheetName1, "srv");
            pricePage.UnfoldAll();
            servicePage = serviceGeneralInformationsPage.BackToList();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);

            var servicePricePage = servicePage.ClickOnFirstService();
            servicePricePage.UnfoldAll();

            var editPriceModal = servicePricePage.EditFirstPrice(null, null);
            editPriceModal.GetDefaultModeValue();
            servicePricePage.BackToList();

            servicePage.Export(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string sheetName, string fileName, string value
            OpenXmlExcel.WriteDataInColumn("Action", SHEET1, filePath, "M", CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Default loading value", SHEET1, filePath, "dfgd", CellValues.SharedString);


            WebDriver.Navigate().Refresh();
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);
            var isImport = importPopup.ImportModal();

            //ASSERT
            Assert.IsFalse(isImport, "L'Import a échoué.");

        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_With_M_ActionFromEmptyValueError()
        {

            // Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string datasheetName1 = TestContext.Properties["DatasheetName1"].ToString();


            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            bool newVersionPrint = true;

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();

            servicePage.ClearDownloads();

            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10), "10", datasheetName1, "srv");
            servicePage = serviceGeneralInformationsPage.BackToList();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);

            servicePage.Export(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string sheetName, string fileName, string value
            OpenXmlExcel.WriteDataInColumn("Action", SHEET1, filePath, "M", CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("From", SHEET1, filePath, "", CellValues.Date);

            WebDriver.Navigate().Refresh();
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);
            var isImport = importPopup.ImportModal();

            //ASSERT
            Assert.IsFalse(isImport, "L'Import a été éffectué malgrès la date de départ vide.");
        }


        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_With_M_ActionIsGenericEmptyValueError()
        {

            // Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string datasheetName1 = TestContext.Properties["DatasheetName1"].ToString();


            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            bool newVersionPrint = true;

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();

            servicePage.ClearDownloads();

            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10), "10", datasheetName1, "srv");
            pricePage.UnfoldAll();
            servicePage = serviceGeneralInformationsPage.BackToList();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);

            servicePage.Export(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string sheetName, string fileName, string value
            OpenXmlExcel.WriteDataInColumn("Action", SHEET1, filePath, "M", CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Is Generic", SHEET1, filePath, "", CellValues.String);


            WebDriver.Navigate().Refresh();
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);
            var isImport = importPopup.ImportModal();

            //ASSERT
            Assert.IsFalse(isImport, "L'Import a été éffectué malgrès le Is Generic vide.");

        }


        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_With_M_ActionIsInvoicedEmptyValueError()
        {

            // Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string datasheetName1 = TestContext.Properties["DatasheetName1"].ToString();


            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            bool newVersionPrint = true;

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();

            servicePage.ClearDownloads();

            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10), "10", datasheetName1, "srv");
            servicePage = serviceGeneralInformationsPage.BackToList();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);

            servicePage.Export(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string sheetName, string fileName, string value
            OpenXmlExcel.WriteDataInColumn("Action", SHEET1, filePath, "M", CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Is Invoiced", SHEET1, filePath, "", CellValues.String);

            WebDriver.Navigate().Refresh();
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);
            var isImport = importPopup.ImportModal();

            //ASSERT
            Assert.IsFalse(isImport, "L'Import a été éffectué malgrès le Is Generic vide.");

        }


        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_With_M_ActionIsProducedEmptyValueError()
        {

            // Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string datasheetName1 = TestContext.Properties["DatasheetName1"].ToString();


            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            bool newVersionPrint = true;

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();

            servicePage.ClearDownloads();

            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10), "10", datasheetName1, "srv");
            servicePage = serviceGeneralInformationsPage.BackToList();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);

            servicePage.Export(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string sheetName, string fileName, string value
            OpenXmlExcel.WriteDataInColumn("Action", SHEET1, filePath, "M", CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Is Produced", SHEET1, filePath, "", CellValues.String);

            WebDriver.Navigate().Refresh();
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);
            var isImport = importPopup.ImportModal();

            //ASSERT
            Assert.IsFalse(isImport, "L'Import a été éffectué malgrès le Is Generic vide.");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_With_M_ActionIsGenericAnotherValueError()
        {

            // Prepare
            // Répertoire du fichire à importer
            var path = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase).Substring(6);
            var excelPath = path + "\\PageObjects\\Customers\\Service\\isGeneric.xlsx";

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();

            servicePage.ClearDownloads();

            WebDriver.Navigate().Refresh();

            var isImport = servicePage.ImportCustomersExcelFile(excelPath);

            //ASSERT
            Assert.IsFalse(isImport, "L'Import a été éffectué malgrès le Is Generic avec une mauvaise valeur.");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_With_M_ActionIsProducedAnotherValueError()
        {

            // Répertoire du fichire à importer
            var path = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase).Substring(6);
            var excelPath = path + "\\PageObjects\\Customers\\Service\\isProduced.xlsx";

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();

            servicePage.ClearDownloads();

            WebDriver.Navigate().Refresh();
            var isImport = servicePage.ImportCustomersExcelFile(excelPath);

            //ASSERT
            Assert.IsFalse(isImport, "L'Import a été éffectué malgrès le Is Generic avec une mauvaise valeur.");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_With_M_ActionIsInvoicedAnotherValueError()
        {

            // Prepare
            // Répertoire du fichire à importer
            var path = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase).Substring(6);
            var excelPath = path + "\\PageObjects\\Customers\\Service\\isInvoiced.xlsx";

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();

            servicePage.ClearDownloads();

            WebDriver.Navigate().Refresh();
            var isImport = servicePage.ImportCustomersExcelFile(excelPath);

            //ASSERT
            Assert.IsFalse(isImport, "L'Import a été éffectué malgrès le Is Generic avec une mauvaise valeur.");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_With_M_ActionIsGenericToTrue()
        {

            // Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string datasheetName1 = TestContext.Properties["DatasheetName1"].ToString();


            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            bool newVersionPrint = true;

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();

            servicePage.ClearDownloads();

            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10), "10", datasheetName1, "srv");
            servicePage = serviceGeneralInformationsPage.BackToList();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);

            servicePage.Export(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string sheetName, string fileName, string value
            OpenXmlExcel.WriteDataInColumn("Action", SHEET1, filePath, "M", CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Is Generic", SHEET1, filePath, "VRAI", CellValues.Boolean);

            WebDriver.Navigate().Refresh();
            var importPopup = servicePage.Import();

            importPopup.CheckFile(correctDownloadedFile.FullName);
            importPopup.ImportModal();

            var servicePagedetail = servicePage.ClickOnFirstService();
            var generalInformationPage = servicePagedetail.ClickOnGeneralInformationTab();
            var isGeneric = generalInformationPage.IsGeneric();
            Assert.IsTrue(isGeneric, "L'Import a échoué.");

        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_With_M_ActionIsInvoicedToTrue()
        {
            // Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string datasheetName1 = TestContext.Properties["DatasheetName1"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            bool newVersionPrint = true;

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();

            servicePage.ClearDownloads();

            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10), "10", datasheetName1, "srv");
            servicePage = serviceGeneralInformationsPage.BackToList();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);

            servicePage.Export(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string sheetName, string fileName, string value
            OpenXmlExcel.WriteDataInColumn("Action", SHEET1, filePath, "M", CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Is Invoiced", SHEET1, filePath, "VRAI", CellValues.Boolean);

            WebDriver.Navigate().Refresh();
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);
            importPopup.ImportModal();

            var servicePagedetail = servicePage.ClickOnFirstService();
            var generalInformationPage = servicePagedetail.ClickOnGeneralInformationTab();

            //ASSERT
            Assert.IsTrue(generalInformationPage.IsInvoiced(), "L'Import a échoué.");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_With_M_ActionIsProducedToTrue()
        {

            // Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string datasheetName1 = TestContext.Properties["DatasheetName1"].ToString();


            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            bool newVersionPrint = true;

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();

            servicePage.ClearDownloads();

            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10), "10", datasheetName1, "srv");
            servicePage = serviceGeneralInformationsPage.BackToList();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);

            servicePage.Export(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string sheetName, string fileName, string value
            OpenXmlExcel.WriteDataInColumn("Action", SHEET1, filePath, "M", CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("Is Produced", SHEET1, filePath, "VRAI", CellValues.Boolean);

            WebDriver.Navigate().Refresh();
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);
            importPopup.ImportModal();

            var servicePagedetail = servicePage.ClickOnFirstService();
            var generalInformationPage = servicePagedetail.ClickOnGeneralInformationTab();

            //ASSERT
            Assert.IsTrue(generalInformationPage.IsProduced(), "L'Import a échoué.");

        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_With_M_ActionFromDateSuperiorToError()
        {

            // Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string datasheetName1 = TestContext.Properties["DatasheetName1"].ToString();
            DateTime dateChange = DateUtils.Now.AddDays(11);
            string dateValue = dateChange.ToOADate().ToString(CultureInfo.InvariantCulture);

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            bool newVersionPrint = true;

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();

            servicePage.ClearDownloads();

            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10), "10", datasheetName1, "srv");
            servicePage = serviceGeneralInformationsPage.BackToList();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);

            servicePage.Export(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string sheetName, string fileName, string value
            OpenXmlExcel.WriteDataInColumn("Action", SHEET1, filePath, "M", CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("From", SHEET1, filePath, dateValue, CellValues.Date);

            WebDriver.Navigate().Refresh();
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);
            var isImport = importPopup.ImportWithFail();

            //ASSERT
            Assert.IsFalse(isImport, "L'Import a été effectué malgrès une date de depart superieur à la date fin.");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_With_M_ActionDateValueToError()
        {
            // Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string datasheetName1 = TestContext.Properties["DatasheetName1"].ToString();


            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            bool newVersionPrint = true;

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();

            servicePage.ClearDownloads();

            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice(); priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10), "10", datasheetName1, "srv");
            servicePage = serviceGeneralInformationsPage.BackToList();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);

            servicePage.Export(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //string columnName, string sheetName, string fileName, string value
            OpenXmlExcel.WriteDataInColumn("Action", SHEET1, filePath, "M", CellValues.SharedString);
            OpenXmlExcel.WriteDataInColumn("From", SHEET1, filePath, "Toto", CellValues.Date);

            WebDriver.Navigate().Refresh();
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);
            var isImport = importPopup.ImportModal();

            //ASSERT
            Assert.IsFalse(isImport, "L'Import a été effectué malgrès une mauvaise valeur.");
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_IsActive()
        {
            // Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string datasheetName1 = TestContext.Properties["DatasheetName1"].ToString();

            // Ensuring consistent date format
            DateTime fromDate = DateTime.Now.AddDays(-6);
            DateTime toDate = DateTime.Now.AddDays(-4);

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            bool newVersionPrint = true;

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            try
            {
                servicePage.ClearDownloads();

                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();

                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, fromDate, toDate, "10", datasheetName1, "srv");
                pricePage.UnfoldAll();
                homePage.GoToCustomers_ServicePage();

                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);

                servicePage.Export(newVersionPrint);
                // On récupère les fichiers du répertoire de téléchargement
                DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
                FileInfo[] taskFiles = taskDirectory.GetFiles();
                // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
                var correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
                Assert.IsNotNull(correctDownloadedFile);

                var fileName = correctDownloadedFile.Name;
                var filePath = Path.Combine(downloadsPath, fileName);
                //string columnName, string sheetName, string fileName, string value
                OpenXmlExcel.WriteDataInColumn("Action", SHEET1, filePath, "M", CellValues.SharedString);
                if (OpenXmlExcel.ReadAllDataInColumn("Is Active", SHEET1, filePath, "VRAI", true))
                {
                    OpenXmlExcel.WriteDataInColumn("Is Active", SHEET1, filePath, "FAUX", CellValues.Boolean);
                    OpenXmlExcel.WriteDataInColumn("Price is Active", SHEET1, filePath, "FAUX", CellValues.Boolean);
                    OpenXmlExcel.WriteDataInColumn("Datasheet", SHEET1, filePath, null, CellValues.String);
                }
                else
                {
                    OpenXmlExcel.WriteDataInColumn("Is Active", SHEET1, filePath, "VRAI", CellValues.Boolean);
                }

                WebDriver.Navigate().Refresh();
                var importPopup = servicePage.Import();
                importPopup.CheckFile(correctDownloadedFile.FullName);
                var isImport = importPopup.ImportModal();

                // ASSERT
                Assert.IsTrue(isImport, "L'Import a échoué.");

                // Check if the service is inactive
                WebDriver.Navigate().Refresh();
                servicePage.Filter(ServicePage.FilterType.ShowOnlyInactive, true);
                Assert.IsTrue(servicePage.CheckTotalNumber() >= 1, "Le service n'as pas été incativé via l'import");
            }
            finally
            {
                servicePage = homePage.GoToCustomers_ServicePage();
                var serviceMassiveDelete = servicePage.ClickMassiveDelete();

                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.From, DateUtils.Now.AddDays(-6).ToString("dd/MM/yyyy"));
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.To, DateUtils.Now.AddDays(-4).ToString("dd/MM/yyyy"));

                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();
                servicePage.ResetFilters();
            }
        }


        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_ImportAddAndUpdate()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var servicePage = homePage.GoToCustomers_ServicePage();


            bool newVersionPrint = true;
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            if (newVersionPrint)
            {
                servicePage.ClearDownloads();
            }

            // chercher ou sinon créer un nouveau Service, ajouter un Price
            string serviceName = "AUTO_TEST_SERVICE_A2M1_AddAndUpdate";
            //*[@id="ServiceFilter"]
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            int counterService1 = servicePage.TableCount();
            if (counterService1 == 0)
            {
                ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.Date, DateUtils.Now.Date.AddDays(7));
                servicePage = pricePage.BackToList();
                // provoque la bonne recherche
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.ShowAll, true);
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            }


            // exporter (en oubliant pas l'imprimante en haut à droite)
            servicePage.Export(newVersionPrint);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            FileInfo correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            //Prices_20220104_0159.xlsx
            Assert.IsNotNull(correctDownloadedFile, "Fichier en entré non trouvé");

            int resultNumber = OpenXmlExcel.GetExportResultNumber("Services Prices", correctDownloadedFile.FullName);
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);
            List<string> services = OpenXmlExcel.GetValuesInList("Service", "Services Prices", correctDownloadedFile.FullName);
            int trouve = -1;
            for (int row = 0; row < services.Count; row++)
            {
                if (services[row].Contains(serviceName))
                {
                    trouve = row;
                    break;
                }
            }
            Assert.IsTrue(trouve >= 0, "service AUTO_TEST_SERVICE non trouvé");

            // Ligne type M : date To modifié
            OpenXmlExcel.WriteDataInColumn("Action", "Services Prices", correctDownloadedFile.FullName, "M", CellValues.SharedString, true);
            DateTime From = DateUtils.Now.Date;
            DateTime To = From.AddYears(5);
            // Rows 3 : Date begin can't be modified. The price is currently active.
            OpenXmlExcel.WriteDataInColumn("To", "Services Prices", correctDownloadedFile.FullName, Convert.ToInt32(To.ToOADate()).ToString(), CellValues.Date, true);

            // 1ere ligne type A : date From To modifié, sur autre plage de date sans chevauchement
            DateTime From2 = To.AddDays(1);
            DateTime To2 = From2.AddDays(10);
            // copie de la ligne d'en haut (inséré en première position)
            OpenXmlExcel.DuplicateFirstLine("Services Prices", correctDownloadedFile.FullName);
            OpenXmlExcel.WriteDataInColumn("Action", "Services Prices", correctDownloadedFile.FullName, "A", CellValues.SharedString, true);
            OpenXmlExcel.WriteDataInColumn("PriceId", "Services Prices", correctDownloadedFile.FullName, null, CellValues.Number, true);
            OpenXmlExcel.WriteDataInColumn("From", "Services Prices", correctDownloadedFile.FullName, Convert.ToInt32(From2.ToOADate()).ToString(), CellValues.Date, true);
            OpenXmlExcel.WriteDataInColumn("To", "Services Prices", correctDownloadedFile.FullName, Convert.ToInt32(To2.ToOADate()).ToString(), CellValues.Date, true);

            // 2ème ligne A
            DateTime From3 = To2.AddDays(1);
            DateTime To3 = From3.AddDays(10);
            // copie de la ligne d'en haut (inséré en première position)
            OpenXmlExcel.DuplicateFirstLine("Services Prices", correctDownloadedFile.FullName);
            OpenXmlExcel.WriteDataInColumn("From", "Services Prices", correctDownloadedFile.FullName, Convert.ToInt32(From3.ToOADate()).ToString(), CellValues.Date, true);
            OpenXmlExcel.WriteDataInColumn("To", "Services Prices", correctDownloadedFile.FullName, Convert.ToInt32(To3.ToOADate()).ToString(), CellValues.Date, true);

            //1.Je charge le fichier dans WR
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);
            //2.Le message suivant apparait selon les cas
            //Verification step done successfully.
            //To add : X line
            //To update : X line
            Assert.AreEqual("Verification step done successfully.", importPopup.CheckFilePopupMessage());
            Assert.AreEqual("2 lines", importPopup.CheckFilePopupToAdd());
            Assert.AreEqual("1 line", importPopup.CheckFilePopupToUpdate());
            //Sur la ligne to add on doit voir apparaitre le nombre de lignes sur lesquelles on a mis M dans le champs Action.

            //3.Je clique sur le bouton import
            //Import of prices done successfully.
            var isImport = importPopup.Import();

            // permet de rejouer le test
            ServicePricePage servicePricePage = servicePage.ClickOnFirstService();
            servicePricePage.ToggleFirstPrice();
            servicePricePage.DeleteLastPrice();
            servicePricePage.ToggleFirstPrice();
            servicePricePage.DeleteLastPrice();

            Assert.IsTrue(isImport, "import raté");
        }

        [Ignore]//le besoin fonctionnel n'est pas clair
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_ImportServiceID()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var servicePage = homePage.GoToCustomers_ServicePage();

            bool newVersionPrint = true;
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);

            servicePage.ClearDownloads();

            // chercher ou sinon créer un nouveau Service, ajouter un Price
            string serviceName1 = "AUTO_TEST_SERVICE_M1_ServiceId_1";
            string serviceName2 = "AUTO_TEST_SERVICE_M1_ServiceId_2";
            string serviceNames = "AUTO_TEST_SERVICE_M1_ServiceId";
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);
            servicePage.Filter(ServicePage.FilterType.Search, serviceNames);

            servicePage.Filter(ServicePage.FilterType.Search, serviceName1);
            int counterService1 = servicePage.TableCount();
            if (counterService1 == 0)
            {
                ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName1);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.Date.AddDays(-7), DateUtils.Now.Date.AddDays(7));
                servicePage = pricePage.BackToList();
            }

            servicePage.Filter(ServicePage.FilterType.Search, serviceName2);
            int counterService2 = servicePage.TableCount();
            if (counterService2 == 0)
            {
                ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName2);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.Date.AddDays(-7), DateUtils.Now.Date.AddDays(7));
                servicePage = pricePage.BackToList();
            }
            // on liste les deux services
            servicePage.Filter(ServicePage.FilterType.Search, serviceNames);

            // exporter (en oubliant pas l'imprimante en haut à droite)
            servicePage.Export(newVersionPrint);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            FileInfo correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            //Prices_20220104_0159.xlsx
            Assert.IsNotNull(correctDownloadedFile, "Fichier en entré non trouvé");

            int resultNumber = OpenXmlExcel.GetExportResultNumber("Services Prices", correctDownloadedFile.FullName);
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);

            OpenXmlExcel.DuplicateFirstLine("Services Prices", correctDownloadedFile.FullName);
            OpenXmlExcel.WriteDataInColumn("Action", "Services Prices", correctDownloadedFile.FullName, "M", CellValues.SharedString, true);

            OpenXmlExcel.GetValuesInList("Service", "Services Prices", correctDownloadedFile.FullName);
            if (OpenXmlExcel.ReadAllDataInColumn("Service", "Services Prices", correctDownloadedFile.FullName, serviceName1, true))
            {
                OpenXmlExcel.WriteDataInColumn("Service", "Services Prices", correctDownloadedFile.FullName, serviceName2, CellValues.SharedString, true);
            }
            else
            {
                OpenXmlExcel.WriteDataInColumn("Service", "Services Prices", correctDownloadedFile.FullName, serviceName1, CellValues.SharedString, true);
            }

            DateTime From = DateUtils.Now.Date;
            DateTime To = From.AddYears(3);
            // Rows 3 : Date begin can't be modified. The price is currently active.
            OpenXmlExcel.WriteDataInColumn("To", "Services Prices", correctDownloadedFile.FullName, Convert.ToInt32(To.ToOADate()).ToString(), CellValues.Date, true);


            //1.Je charge le fichier dans WR
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);
            //2.Le message suivant apparait
            //Verification step done successfully.
            //To add: X line
            //To update : X line
            Assert.AreEqual("Verification step done successfully.", importPopup.CheckFilePopupMessage());
            Assert.AreEqual("0 line", importPopup.CheckFilePopupToAdd());
            Assert.AreEqual("1 line", importPopup.CheckFilePopupToUpdate());
            //Sur la ligne to update on doit voir apparaitre le nombre de lignes sur lesquelles on a mis M dans le champs Action.

            //3.Je clique sur le bouton import
            var isImport = importPopup.Import();


            //Import can't be performed. Some error(s) have been detected while extracting data from file :
            //Rows 2 : Service with name 'AUTO_TEST_SERVICE2' already exists.
            Assert.IsTrue(isImport, "import raté");

            //Resultat attendu : Le service est mis à jour, mais le Service ID n'est pas pris en compte
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_ImportServiceIDPriceID()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var servicePage = homePage.GoToCustomers_ServicePage();


            bool newVersionPrint = true;
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);

            servicePage.ClearDownloads();

            // chercher ou sinon créer un nouveau Service, ajouter un Price
            string serviceName = "AUTO_TEST_SERVICE_A1_ServiceIDPriceID";
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            int counterService1 = servicePage.TableCount();
            if (counterService1 == 0)
            {
                ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.Date.AddDays(-7), DateUtils.Now.Date.AddDays(7));
                servicePage = pricePage.BackToList();
                // provoque la bonne recherche
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            }


            // exporter (en oubliant pas l'imprimante en haut à droite)
            servicePage.Export(newVersionPrint);
            taskDirectory.Refresh();
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            FileInfo correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            //Prices_20220104_0159.xlsx
            Assert.IsNotNull(correctDownloadedFile, "Fichier en entré non trouvé");

            int resultNumber = OpenXmlExcel.GetExportResultNumber("Services Prices", correctDownloadedFile.FullName);
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);
            OpenXmlExcel.DuplicateFirstLine("Services Prices", correctDownloadedFile.FullName);

            OpenXmlExcel.WriteDataInColumn("Action", "Services Prices", correctDownloadedFile.FullName, "A", CellValues.SharedString, true);
            // nouveau service
            OpenXmlExcel.WriteDataInColumn("Service", "Services Prices", correctDownloadedFile.FullName, serviceName + "_" + DateUtils.Now.ToString("yyyy-MM-dd_HH-mm-ss"), CellValues.SharedString, true);
            // nouveau serviceId à null
            OpenXmlExcel.WriteDataInColumn("ServiceId", "Services Prices", correctDownloadedFile.FullName, null, CellValues.Number, true);

            //1.Je charge le fichier dans WR
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);

            //2.Le message suivant apparait
            //Verification step done successfully.
            //To add: X line
            //To update : X line
            Assert.AreEqual("Verification step done successfully.", importPopup.CheckFilePopupMessage());
            Assert.AreEqual("1 line", importPopup.CheckFilePopupToAdd());
            Assert.AreEqual("0 line", importPopup.CheckFilePopupToUpdate());
            //Sur la ligne to Add on doit voir apparaitre le nombre de lignes sur lesquelles on a mis A dans le champs Action.

            //3.Je clique sur le bouton import
            var isImport = importPopup.Import();

            Assert.IsTrue(isImport, "import raté");
            // Resultat attendu : Le service est mis à jour, mais le Service ID n'est pas pris en compte
            // Resultat attendu : Le service est créé, mais le Price ID n'est pas pris en compte ???

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);

            // Contrer le "Prices_20220107_0151 (1).xlsx"


            // re-exporter (en oubliant pas l'imprimante en haut à droite)
            servicePage.Export(newVersionPrint);
            taskDirectory.Refresh();
            taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            //Prices_20220104_0159.xlsx
            Assert.IsNotNull(correctDownloadedFile, "Fichier en entré non trouvé");
            string secondTime = correctDownloadedFile.FullName;
            secondTime = secondTime.Replace(".xlsx", " (1).xlsx");
            FileInfo secondTimeFileInfo = new FileInfo(secondTime);
            if (secondTimeFileInfo.Exists)
            {
                correctDownloadedFile = secondTimeFileInfo;
            }

            resultNumber = OpenXmlExcel.GetExportResultNumber("Services Prices", correctDownloadedFile.FullName);
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);
            List<string> columnPriceId = OpenXmlExcel.GetValuesInList("PriceId", "Services Prices", correctDownloadedFile.FullName);
            List<string> checkDoublons = new List<string>();
            foreach (string priceId in columnPriceId)
            {
                if (checkDoublons.Contains(priceId))
                {
                    Assert.Fail("Pas nouveau PriceId");
                }
                else
                {
                    checkDoublons.Add(priceId);
                }
            }

        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_ImportSiteCode()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var servicePage = homePage.GoToCustomers_ServicePage();


            bool newVersionPrint = true;
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);

            servicePage.ClearDownloads();

            // chercher ou sinon créer un nouveau Service, ajouter un Price
            string serviceName = "AUTO_TEST_SERVICE_M1_SiteCode";
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            int counterService1 = servicePage.TableCount();
            if (counterService1 == 0)
            {
                ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.Date.AddDays(-7), DateUtils.Now.Date.AddDays(7));
                servicePage = pricePage.BackToList();
                // provoque la bonne recherche
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            }


            // exporter (en oubliant pas l'imprimante en haut à droite)
            servicePage.Export(newVersionPrint);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            FileInfo correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            //Prices_20220104_0159.xlsx
            Assert.IsNotNull(correctDownloadedFile, "Fichier en entré non trouvé");

            int resultNumber = OpenXmlExcel.GetExportResultNumber("Services Prices", correctDownloadedFile.FullName);
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);

            OpenXmlExcel.WriteDataInColumn("Action", "Services Prices", correctDownloadedFile.FullName, "M", CellValues.SharedString, true);
            // nouveau SiteCode ACE à la place de MAD
            OpenXmlExcel.WriteDataInColumn("Site Code", "Services Prices", correctDownloadedFile.FullName, "ACE", CellValues.SharedString, true);

            //1.Je charge le fichier dans WR
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);

            //2.Le message suivant apparait
            //Verification step done successfully.
            //To add: X line
            //To update : X line

            //Sur la ligne to update on doit voir apparaitre le nombre de lignes sur lesquelles on a mis M dans le champs Action.
            //3.Je clique sur le bouton import
            var isImport = importPopup.ImportWithFail();

            Assert.IsFalse(isImport, "import réussi malgrès changement de SiteCode");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_ImportSiteCodeService()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var servicePage = homePage.GoToCustomers_ServicePage();


            bool newVersionPrint = true;
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);

            servicePage.ClearDownloads();

            // chercher ou sinon créer un nouveau Service, ajouter un Price
            string serviceName = "AUTO_TEST_SERVICE_A1_SiteCodeService";
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            int counterService1 = servicePage.TableCount();
            if (counterService1 == 0)
            {
                ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.Date.AddDays(-7), DateUtils.Now.Date.AddDays(7));
                servicePage = pricePage.BackToList();
                // provoque la bonne recherche
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            }


            // exporter (en oubliant pas l'imprimante en haut à droite)
            servicePage.Export(newVersionPrint);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            FileInfo correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            //Prices_20220104_0159.xlsx
            Assert.IsNotNull(correctDownloadedFile, "Fichier en entré non trouvé");

            int resultNumber = OpenXmlExcel.GetExportResultNumber("Services Prices", correctDownloadedFile.FullName);
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);

            OpenXmlExcel.WriteDataInColumn("Action", "Services Prices", correctDownloadedFile.FullName, "A", CellValues.SharedString, true);
            // ACE à la place de MAD
            OpenXmlExcel.WriteDataInColumn("Site Code", "Services Prices", correctDownloadedFile.FullName, "ACE", CellValues.SharedString, true);
            // nouveau service
            OpenXmlExcel.WriteDataInColumn("PriceId", "Services Prices", correctDownloadedFile.FullName, null, CellValues.Number, true);
            OpenXmlExcel.WriteDataInColumn("ServiceId", "Services Prices", correctDownloadedFile.FullName, null, CellValues.Number, true);
            OpenXmlExcel.WriteDataInColumn("Service", "Services Prices", correctDownloadedFile.FullName, serviceName + "_" + DateUtils.Now.ToString("yyyy-MM-dd_HH-mm-ss"), CellValues.SharedString, true);

            //1.Je charge le fichier dans WR
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);

            //2.Le message suivant apparait
            //Verification step done successfully.
            //To add: X line
            //To update : X line
            Assert.AreEqual("Verification step done successfully.", importPopup.CheckFilePopupMessage());
            Assert.AreEqual("1 line", importPopup.CheckFilePopupToAdd());
            Assert.AreEqual("0 line", importPopup.CheckFilePopupToUpdate());

            //Sur la ligne to add on doit voir apparaitre le nombre de lignes sur lesquelles on a mis A dans le champs Action.
            //3.Je clique sur le bouton import
            var isImport = importPopup.Import();

            Assert.IsTrue(isImport, "import raté");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_ImportUpdateServiceID()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();
            var servicePage = homePage.GoToCustomers_ServicePage();


            bool newVersionPrint = true;
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);

            servicePage.ClearDownloads();

            // chercher ou sinon créer un nouveau Service, ajouter un Price
            string serviceName = "AUTO_TEST_SERVICE_M1_UpdateServiceID";
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            int counterService1 = servicePage.TableCount();
            if (counterService1 == 0)
            {
                ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.Date.AddDays(-7), DateUtils.Now.Date.AddDays(7));
                servicePage = pricePage.BackToList();
                // provoque la bonne recherche
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            }


            // exporter (en oubliant pas l'imprimante en haut à droite)
            servicePage.Export(newVersionPrint);
            taskDirectory.Refresh();
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            FileInfo correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            //Prices_20220104_0159.xlsx
            Assert.IsNotNull(correctDownloadedFile, "Fichier en entré non trouvé");

            int resultNumber = OpenXmlExcel.GetExportResultNumber("Services Prices", correctDownloadedFile.FullName);
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);

            // nombre de cas avant l'opération
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            int counterServiceAvant = servicePage.TableCount();

            OpenXmlExcel.WriteDataInColumn("Action", "Services Prices", correctDownloadedFile.FullName, "M", CellValues.SharedString, true);
            OpenXmlExcel.WriteDataInColumn("Service", "Services Prices", correctDownloadedFile.FullName, serviceName + "_" + DateUtils.Now.ToString("yyyy-MM-dd_HH-mm-ss"), CellValues.SharedString, true);

            //1.Je charge le fichier dans WR
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);

            //2.Le message suivant apparait
            //Verification step done successfully.
            //To add: X line
            //To update : X line
            Assert.AreEqual("Verification step done successfully.", importPopup.CheckFilePopupMessage());
            Assert.AreEqual("0 line", importPopup.CheckFilePopupToAdd());
            Assert.AreEqual("1 line", importPopup.CheckFilePopupToUpdate());
            //Sur la ligne to update on doit voir apparaitre le nombre de lignes sur lesquelles on a mis M dans le champs Action.

            //3.Je clique sur le bouton import
            var isImport = importPopup.Import();

            Assert.IsTrue(isImport, "import raté");

            // nombre de cas avant l'opération
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            int counterServiceApres = servicePage.TableCount();

            Assert.AreEqual(counterServiceAvant, counterServiceApres, "Mauvais nombre de services");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_IsActive_IsEmpty()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var servicePage = homePage.GoToCustomers_ServicePage();


            bool newVersionPrint = true;
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);

            servicePage.ClearDownloads();

            // chercher ou sinon créer un nouveau Service, ajouter un Price
            string serviceName = "AUTO_TEST_SERVICE_M1A1_Is_Active_IsEmpty";
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            int counterService1 = servicePage.TableCount();
            if (counterService1 == 0)
            {
                ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.Date.AddDays(-7), DateUtils.Now.Date.AddDays(7));
                servicePage = pricePage.BackToList();
                // provoque la bonne recherche
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            }


            // exporter (en oubliant pas l'imprimante en haut à droite)
            servicePage.Export(newVersionPrint);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            FileInfo correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            //Prices_20220104_0159.xlsx
            Assert.IsNotNull(correctDownloadedFile, "Fichier en entré non trouvé");

            int resultNumber = OpenXmlExcel.GetExportResultNumber("Services Prices", correctDownloadedFile.FullName);
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);

            // Modif
            OpenXmlExcel.WriteDataInColumn("Action", "Services Prices", correctDownloadedFile.FullName, "M", CellValues.SharedString, true);
            // Boolean null => VRAI
            OpenXmlExcel.WriteDataInColumn("Is Active", "Services Prices", correctDownloadedFile.FullName, "", CellValues.String, true);
            // Ajout
            OpenXmlExcel.DuplicateFirstLine("Services Prices", correctDownloadedFile.FullName);
            OpenXmlExcel.WriteDataInColumn("Action", "Services Prices", correctDownloadedFile.FullName, "A", CellValues.SharedString, true);
            OpenXmlExcel.WriteDataInColumn("ServiceId", "Services Prices", correctDownloadedFile.FullName, null, CellValues.Number, true);
            OpenXmlExcel.WriteDataInColumn("Service", "Services Prices", correctDownloadedFile.FullName, serviceName + "_" + DateUtils.Now.ToString("yyyy-MM-dd_HH-mm-ss"), CellValues.SharedString, true);

            //1.Je charge le fichier dans WR
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);

            //3.Je clique sur le bouton Cancel
            var isImport = importPopup.ImportWithFail2();
            Assert.IsFalse(isImport, "import avec succes alors que Is Active est vide");

        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_ImportRenameAlreadyExistsServiceID()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var servicePage = homePage.GoToCustomers_ServicePage();

            bool newVersionPrint = true;
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);

            servicePage.ClearDownloads();

            // chercher ou sinon créer un nouveau Service, ajouter un Price
            string serviceName1 = "AUTO_TEST_SERVICE_M1_RenameAlreadyExistesServiceID_1";
            string serviceName2 = "AUTO_TEST_SERVICE_M1_RenameAlreadyExistesServiceID_2";
            string serviceNames = "AUTO_TEST_SERVICE_M1_RenameAlreadyExistesServiceID";
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);
            servicePage.Filter(ServicePage.FilterType.Search, serviceNames);

            servicePage.Filter(ServicePage.FilterType.Search, serviceName1);
            int counterService1 = servicePage.TableCount();
            if (counterService1 == 0)
            {
                ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName1);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.Date.AddDays(-7), DateUtils.Now.Date.AddDays(7));
                servicePage = pricePage.BackToList();
            }

            servicePage.Filter(ServicePage.FilterType.Search, serviceName2);
            int counterService2 = servicePage.TableCount();
            if (counterService2 == 0)
            {
                ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName2);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.Date.AddDays(-7), DateUtils.Now.Date.AddDays(7));
                servicePage = pricePage.BackToList();
            }
            // on liste les deux services
            servicePage.Filter(ServicePage.FilterType.Search, serviceNames);

            // exporter (en oubliant pas l'imprimante en haut à droite)
            servicePage.Export(newVersionPrint);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            FileInfo correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            Assert.IsNotNull(correctDownloadedFile, "Fichier en entré non trouvé");

            int resultNumber = OpenXmlExcel.GetExportResultNumber("Services Prices", correctDownloadedFile.FullName);
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);

            // Je mets à jour le nom du service avec un nom déjà existant
            OpenXmlExcel.WriteDataInColumn("Action", "Services Prices", correctDownloadedFile.FullName, "M", CellValues.SharedString, true);
            if (OpenXmlExcel.ReadAllDataInColumn("Service", "Services Prices", correctDownloadedFile.FullName, serviceName1, true))
            {
                OpenXmlExcel.WriteDataInColumn("Service", "Services Prices", correctDownloadedFile.FullName, serviceName2, CellValues.SharedString, true);
            }
            else
            {
                OpenXmlExcel.WriteDataInColumn("Service", "Services Prices", correctDownloadedFile.FullName, serviceName1, CellValues.SharedString, true);
            }

            //1.Je charge le fichier dans WR
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);

            //3.Je clique sur le bouton import
            var isImport = importPopup.ImportWithFail2();
            Assert.IsFalse(isImport, "import avec succes alors que Is Active est vide");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_ImportNewServiceID()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var servicePage = homePage.GoToCustomers_ServicePage();


            bool newVersionPrint = true;
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);

            servicePage.ClearDownloads();

            // chercher ou sinon créer un nouveau Service, ajouter un Price
            string serviceName1 = "AUTO_TEST_SERVICE_A1_NewServiceId";
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);
            servicePage.Filter(ServicePage.FilterType.Search, serviceName1);
            int counterService1 = servicePage.TableCount();
            if (counterService1 == 0)
            {
                ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName1);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.Date.AddDays(-7), DateUtils.Now.Date.AddDays(7));
                servicePage = pricePage.BackToList();
                counterService1 = 1;
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);
            servicePage.Filter(ServicePage.FilterType.Search, serviceName1);

            // exporter (en oubliant pas l'imprimante en haut à droite)
            servicePage.Export(newVersionPrint);
            taskDirectory.Refresh();
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            FileInfo correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            //Prices_20220104_0159.xlsx
            Assert.IsNotNull(correctDownloadedFile, "Fichier en entré non trouvé");

            int resultNumber = OpenXmlExcel.GetExportResultNumber("Services Prices", correctDownloadedFile.FullName);
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);

            // Ajout
            OpenXmlExcel.DuplicateFirstLine("Services Prices", correctDownloadedFile.FullName);
            OpenXmlExcel.WriteDataInColumn("Action", "Services Prices", correctDownloadedFile.FullName, "A", CellValues.SharedString, true);
            OpenXmlExcel.WriteDataInColumn("ServiceId", "Services Prices", correctDownloadedFile.FullName, null, CellValues.Number, true);
            OpenXmlExcel.WriteDataInColumn("Service", "Services Prices", correctDownloadedFile.FullName, serviceName1 + "_" + DateUtils.Now.ToString("yyyy-MM-dd_HH-mm-ss"), CellValues.SharedString, true);

            //1.Je charge le fichier dans WR
            servicePage.ClosePrintButton();
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);

            //2.Le message suivant apparait
            //Verification step done successfully.
            //To add: X line
            //To update : X line
            Assert.AreEqual("Verification step done successfully.", importPopup.CheckFilePopupMessage());
            Assert.AreEqual("1 line", importPopup.CheckFilePopupToAdd());
            Assert.AreEqual("0 line", importPopup.CheckFilePopupToUpdate());
            //Sur la ligne to add on doit voir apparaitre le nombre de lignes sur lesquelles on a mis A dans le champs Action.
            //3.Je clique sur le bouton import
            var isImport = importPopup.Import();

            Assert.IsTrue(isImport, "import raté");

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);
            servicePage.Filter(ServicePage.FilterType.Search, serviceName1);
            int counterService1Resultat = servicePage.TableCount();

            Assert.AreEqual(counterService1 + 1, counterService1Resultat, "Pas de nouvelle ligne service");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_ImportBadDateTo()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var servicePage = homePage.GoToCustomers_ServicePage();


            bool newVersionPrint = true;
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);

            servicePage.ClearDownloads();

            // chercher ou sinon créer un nouveau Service, ajouter un Price
            string serviceName1 = "AUTO_TEST_SERVICE_A1_BadDateTo";
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);
            servicePage.Filter(ServicePage.FilterType.Search, serviceName1);
            int counterService1 = servicePage.TableCount();
            if (counterService1 == 0)
            {
                ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName1);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.Date.AddDays(-7), DateUtils.Now.Date.AddDays(7));
                servicePage = pricePage.BackToList();
            }

            servicePage.Filter(ServicePage.FilterType.Search, serviceName1 + "BLAH");
            servicePage.Filter(ServicePage.FilterType.Search, serviceName1);

            // exporter (en oubliant pas l'imprimante en haut à droite)
            servicePage.Export(newVersionPrint);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            FileInfo correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            //Prices_20220104_0159.xlsx
            Assert.IsNotNull(correctDownloadedFile, "Fichier en entré non trouvé");

            int resultNumber = OpenXmlExcel.GetExportResultNumber("Services Prices", correctDownloadedFile.FullName);
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);

            OpenXmlExcel.WriteDataInColumn("Action", "Services Prices", correctDownloadedFile.FullName, "A", CellValues.SharedString, true);
            OpenXmlExcel.WriteDataInColumn("To", "Services Prices", correctDownloadedFile.FullName, "HELLO", CellValues.String, true);

            //1.Je charge le fichier dans WR
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);

            //3.Je clique sur le bouton import
            var isImport = importPopup.ImportWithFail2();
            Assert.IsFalse(isImport, "import avec succes alors que To est corrompu");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_ImportBadDateFromTo()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var servicePage = homePage.GoToCustomers_ServicePage();

            bool newVersionPrint = true;
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);

            servicePage.ClearDownloads();

            // chercher ou sinon créer un nouveau Service, ajouter un Price
            string serviceName1 = "AUTO_TEST_SERVICE_A1_BadDateFromTo";
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);
            servicePage.Filter(ServicePage.FilterType.Search, serviceName1);
            int counterService1 = servicePage.TableCount();
            if (counterService1 == 0)
            {
                ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName1);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.Date.AddDays(-7), DateUtils.Now.Date.AddDays(7));
                servicePage = pricePage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName1);

            // exporter (en oubliant pas l'imprimante en haut à droite)
            servicePage.Export(newVersionPrint);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            FileInfo correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            //Prices_20220104_0159.xlsx
            Assert.IsNotNull(correctDownloadedFile, "Fichier en entré non trouvé");

            int resultNumber = OpenXmlExcel.GetExportResultNumber("Services Prices", correctDownloadedFile.FullName);
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);

            // Je mets un date inférieur à From
            OpenXmlExcel.WriteDataInColumn("Action", "Services Prices", correctDownloadedFile.FullName, "A", CellValues.SharedString, true);
            DateTime To = DateUtils.Now.Date;
            OpenXmlExcel.WriteDataInColumn("To", "Services Prices", correctDownloadedFile.FullName, Convert.ToInt32(To.AddDays(-14).ToOADate()).ToString(), CellValues.Date, true);

            //1.Je charge le fichier dans WR
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);

            //2.Le message suivant apparait
            //Verification step done successfully.
            //To add: X line
            //To update : X line
            Assert.AreEqual("Verification step done successfully.", importPopup.CheckFilePopupMessage());
            Assert.AreEqual("1 line", importPopup.CheckFilePopupToAdd());
            Assert.AreEqual("0 line", importPopup.CheckFilePopupToUpdate());
            //Sur la ligne to update on doit voir apparaitre le nombre de lignes sur lesquelles on a mis M dans le champs Action ou sur to add le nombre de ligne sur lesquelles on a mis A

            //3.Je clique sur le bouton import
            var isImport = importPopup.ImportWithFail();
            Assert.IsFalse(isImport, "import avec succes alors que To est avant From");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_ImportPriceMethodScale()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var servicePage = homePage.GoToCustomers_ServicePage();


            bool newVersionPrint = true;
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);

            servicePage.ClearDownloads();

            // chercher ou sinon créer un nouveau Service, ajouter un Price
            string serviceName1 = "AUTO_TEST_SERVICE_A1_PriceMethodScale";
            string datasheet = "1900 ALMUERZO CALIENTE BC CHML-HMFCCHML DIC'14 (MAD)";
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);
            servicePage.Filter(ServicePage.FilterType.Search, serviceName1);
            int counterService1 = servicePage.TableCount();
            if (counterService1 == 0)
            {
                ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName1);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.Date.AddDays(-7), DateUtils.Now.Date.AddDays(7), "10", datasheet);
                servicePage = pricePage.BackToList();
            }

            servicePage.Filter(ServicePage.FilterType.Search, serviceName1 + "BLAH");
            servicePage.Filter(ServicePage.FilterType.Search, serviceName1);

            // exporter (en oubliant pas l'imprimante en haut à droite)
            servicePage.Export(newVersionPrint);
            taskDirectory.Refresh();
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            FileInfo correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            //Prices_20220104_0159.xlsx
            Assert.IsNotNull(correctDownloadedFile, "Fichier en entré non trouvé");

            int resultNumber = OpenXmlExcel.GetExportResultNumber("Services Prices", correctDownloadedFile.FullName);
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);
            //Pour qu'un service soit ajouté, il faut que les champs suivants soient resnseignés:
            //.Scale Nb Person(s)
            //. Default Loading Mode
            //.les Champs de l'onglet Default Scale
            //De plus, il faut avoir une ligne par Scale.


            OpenXmlExcel.WriteDataInColumn("Action", "Services Prices", correctDownloadedFile.FullName, "A", CellValues.SharedString, true);
            // CellValues.Number sort une phrase "Is Invoiced" donc faut écraser via CellValues.SharedString
            // Number KO
            OpenXmlExcel.WriteDataInColumn("Scale Nb person(s)", "Services Prices", correctDownloadedFile.FullName, "10", CellValues.Number, true);
            OpenXmlExcel.WriteDataInColumn("Price", "Services Prices", correctDownloadedFile.FullName, "10.50", CellValues.Number, true);
            OpenXmlExcel.WriteDataInColumn("Price Method", "Services Prices", correctDownloadedFile.FullName, "Scale", CellValues.SharedString, true);
            OpenXmlExcel.WriteDataInColumn("Service", "Services Prices", correctDownloadedFile.FullName, serviceName1 + "_" + DateUtils.Now.ToString("yyyy-MM-dd_HH-mm-ss"), CellValues.SharedString, true);
            OpenXmlExcel.WriteDataInColumn("ServiceId", "Services Prices", correctDownloadedFile.FullName, null, CellValues.Number, true);

            OpenXmlExcel.DuplicateFirstLine("Services Prices", correctDownloadedFile.FullName);
            OpenXmlExcel.WriteDataInColumn("Scale Nb person(s)", "Services Prices", correctDownloadedFile.FullName, "15", CellValues.Number, true);
            OpenXmlExcel.WriteDataInColumn("Price", "Services Prices", correctDownloadedFile.FullName, "12.50", CellValues.Number, true);

            //1.Je charge le fichier dans WR
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);

            //2.Le message suivant apparait selon les cas
            //Verification step done successfully.
            //To add : X line
            //To update : X line
            Assert.AreEqual("Verification step done successfully.", importPopup.CheckFilePopupMessage());
            Assert.AreEqual("2 lines", importPopup.CheckFilePopupToAdd());
            Assert.AreEqual("0 line", importPopup.CheckFilePopupToUpdate());
            //Sur la ligne to add on doit voir apparaitre le nombre de lignes sur lesquelles on a mis A dans le champs Action.

            //3.Je clique sur le bouton import
            var isImport = importPopup.Import();

            //Service AUTO_TEST_SERVICE_A1_PriceMethodScale, Customer TVS, Site MAD
            //Row 2 : The 03/01/2022 > 17/01/2022 price is in conflict with the 03/01/2022 > 17/01/2022 price.
            Assert.IsTrue(isImport, "import raté");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_ImportPriceMethodFixed()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var servicePage = homePage.GoToCustomers_ServicePage();


            bool newVersionPrint = true;
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);

            servicePage.ClearDownloads();

            // chercher ou sinon créer un nouveau Service, ajouter un Price
            string serviceName1 = "AUTO_TEST_SERVICE_A1_PriceMethodFixed_1";
            string datasheet = "1900 ALMUERZO CALIENTE BC CHML-HMFCCHML DIC'14 (MAD)";
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);
            servicePage.Filter(ServicePage.FilterType.Search, serviceName1);
            int counterService1 = servicePage.TableCount();
            DateTime From = DateUtils.Now.Date.AddDays(-7);
            DateTime To = DateUtils.Now.Date.AddDays(7);

            if (counterService1 == 0)
            {
                ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName1);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, From, To, "10", datasheet);
                servicePage = pricePage.BackToList();
            }

            servicePage.Filter(ServicePage.FilterType.Search, serviceName1 + "BLAH");
            servicePage.Filter(ServicePage.FilterType.Search, serviceName1);

            // exporter (en oubliant pas l'imprimante en haut à droite)
            servicePage.Export(newVersionPrint);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            FileInfo correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            //Prices_20220104_0159.xlsx
            Assert.IsNotNull(correctDownloadedFile, "Fichier en entré non trouvé");

            int resultNumber = OpenXmlExcel.GetExportResultNumber("Services Prices", correctDownloadedFile.FullName);
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);

            OpenXmlExcel.WriteDataInColumn("Action", "Services Prices", correctDownloadedFile.FullName, "A", CellValues.SharedString, true);
            //Je mets Fixed comme Price Method
            OpenXmlExcel.WriteDataInColumn("Price Method", "Services Prices", correctDownloadedFile.FullName, "Fixed", CellValues.SharedString, true);
            OpenXmlExcel.WriteDataInColumn("Price", "Services Prices", correctDownloadedFile.FullName, "10.50", CellValues.SharedString, true);
            OpenXmlExcel.WriteDataInColumn("From", "Services Prices", correctDownloadedFile.FullName, Convert.ToInt32(To.AddDays(1).ToOADate()).ToString(), CellValues.Date, true);
            OpenXmlExcel.WriteDataInColumn("To", "Services Prices", correctDownloadedFile.FullName, Convert.ToInt32(To.AddDays(2).ToOADate()).ToString(), CellValues.Date, true);


            //1. Je charge le fichier dans WR
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);

            //2.Le message suivant
            //Verification step done successfully.
            //To add: X line
            //To update : X line
            Assert.AreEqual("Verification step done successfully.", importPopup.CheckFilePopupMessage());
            Assert.AreEqual("1 line", importPopup.CheckFilePopupToAdd());
            Assert.AreEqual("0 line", importPopup.CheckFilePopupToUpdate());
            //Sur la ligne to update on doit voir apparaitre le nombre de lignes sur lesquelles on a mis M dans le champs Action ou sur to add le nombre de ligne sur lesquelles on a mis A

            //3.Je clique sur le bouton import
            var isImport = importPopup.Import();

            //Import has been cancelled. Some error(s) have occurred during the process:
            //Service AUTO_TEST_SERVICE_A1_PriceMethodScaleuh222, Customer TVS, Site MAD
            //Row 2 : The 04/01/2022 > 18/01/2022 price is in conflict with the 04/01/2022 > 18/01/2022 price.
            Assert.IsTrue(isImport, "import raté");

            // permet de rejouer le test
            ServicePricePage servicePricePage = servicePage.ClickOnFirstService();
            servicePricePage.ToggleFirstPrice();
            servicePricePage.DeleteLastPrice();
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_ImportPriceMethodCycle()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var servicePage = homePage.GoToCustomers_ServicePage();


            bool newVersionPrint = true;
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);

            servicePage.ClearDownloads();

            // chercher ou sinon créer un nouveau Service, ajouter un Price
            string serviceName1 = "AUTO_TEST_SERVICE_A1_PriceMethodCycle";
            string datasheet = "1900 ALMUERZO CALIENTE BC CHML-HMFCCHML DIC'14 (MAD)";
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);
            servicePage.Filter(ServicePage.FilterType.Search, serviceName1);
            int counterService1 = servicePage.TableCount();
            DateTime From = DateUtils.Now.Date.AddDays(-7);
            DateTime To = DateUtils.Now.Date.AddDays(7);

            if (counterService1 == 0)
            {
                ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName1);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, From, To, "10", datasheet);
                servicePage = pricePage.BackToList();
            }

            servicePage.Filter(ServicePage.FilterType.Search, serviceName1 + "BLAH");
            servicePage.Filter(ServicePage.FilterType.Search, serviceName1);

            // exporter (en oubliant pas l'imprimante en haut à droite)
            servicePage.Export(newVersionPrint);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            FileInfo correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            //Prices_20220104_0159.xlsx
            Assert.IsNotNull(correctDownloadedFile, "Fichier en entré non trouvé");

            int resultNumber = OpenXmlExcel.GetExportResultNumber("Services Prices", correctDownloadedFile.FullName);
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);

            //Pour qu'un service en Cycle soit ajouté, il faut que les champs suivants soient renseignés:
            //.Price Method
            //.CycleMode
            //.NbCycles
            //.From
            //.To
            //.Cycle
            //.Cycle Duration
            //. Price
            //.Cycle Start At
            //.Cycle Start Duration
            //.Cycle Order
            OpenXmlExcel.WriteDataInColumn("Action", "Services Prices", correctDownloadedFile.FullName, "A", CellValues.SharedString, true);
            OpenXmlExcel.WriteDataInColumn("Service", "Services Prices", correctDownloadedFile.FullName, serviceName1 + "_" + DateUtils.Now.ToString("yyyy-MM-dd_HH-mm-ss"), CellValues.SharedString, true);
            OpenXmlExcel.WriteDataInColumn("ServiceId", "Services Prices", correctDownloadedFile.FullName, null, CellValues.Number, true);

            OpenXmlExcel.WriteDataInColumn("Price Method", "Services Prices", correctDownloadedFile.FullName, "Cycle", CellValues.SharedString, true);
            OpenXmlExcel.WriteDataInColumn("CycleMode", "Services Prices", correctDownloadedFile.FullName, "Continuous", CellValues.SharedString, true);
            OpenXmlExcel.WriteDataInColumn("NbCycles", "Services Prices", correctDownloadedFile.FullName, "2", CellValues.Number, true);
            //.From
            //.To
            OpenXmlExcel.WriteDataInColumn("Cycle", "Services Prices", correctDownloadedFile.FullName, "Cycle 2", CellValues.SharedString, true);
            OpenXmlExcel.WriteDataInColumn("Cycle duration", "Services Prices", correctDownloadedFile.FullName, "1", CellValues.Number, true);
            OpenXmlExcel.WriteDataInColumn("Price", "Services Prices", correctDownloadedFile.FullName, "20", CellValues.Number, true);
            OpenXmlExcel.WriteDataInColumn("Cycle start at", "Services Prices", correctDownloadedFile.FullName, "1", CellValues.Number, true);
            OpenXmlExcel.WriteDataInColumn("Cycle start duration", "Services Prices", correctDownloadedFile.FullName, "5", CellValues.Number, true);
            OpenXmlExcel.WriteDataInColumn("Cycle order", "Services Prices", correctDownloadedFile.FullName, "2", CellValues.Number, true);

            OpenXmlExcel.DuplicateFirstLine("Services Prices", correctDownloadedFile.FullName);

            OpenXmlExcel.WriteDataInColumn("Cycle", "Services Prices", correctDownloadedFile.FullName, "Cycle 1", CellValues.SharedString, true);
            OpenXmlExcel.WriteDataInColumn("Price", "Services Prices", correctDownloadedFile.FullName, "10", CellValues.Number, true);
            OpenXmlExcel.WriteDataInColumn("Cycle order", "Services Prices", correctDownloadedFile.FullName, "1", CellValues.Number, true);

            //1.Je charge le fichier dans WR
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);

            //2.Le message suivant apparait
            //Verification step done successfully.
            //To add: X line
            //To update : X line
            Assert.AreEqual("Verification step done successfully.", importPopup.CheckFilePopupMessage());
            Assert.AreEqual("2 lines", importPopup.CheckFilePopupToAdd());
            Assert.AreEqual("0 line", importPopup.CheckFilePopupToUpdate());
            //Sur la ligne to update on doit voir apparaitre le nombre de lignes sur lesquelles on a mis M dans le champs Action.

            //3.Je clique sur le bouton import
            var isImport = importPopup.Import();

            Assert.IsTrue(isImport, "import raté");

            servicePage.Filter(ServicePage.FilterType.Search, serviceName1 + "BLAH");
            servicePage.Filter(ServicePage.FilterType.Search, serviceName1);

            // exporter (en oubliant pas l'impri
            // mante en haut à droite)
            servicePage.Export(newVersionPrint);
            taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);

            //Prices_20220104_0159.xlsx
            Assert.IsNotNull(correctDownloadedFile, "Fichier en entré non trouvé");
            string secondTime = correctDownloadedFile.FullName;
            secondTime = secondTime.Replace(".xlsx", " (1).xlsx");
            FileInfo secondTimeFileInfo = new FileInfo(secondTime);
            if (secondTimeFileInfo.Exists)
            {
                correctDownloadedFile = secondTimeFileInfo;
            }

            resultNumber = OpenXmlExcel.GetExportResultNumber("Services Prices", correctDownloadedFile.FullName);
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);
            OpenXmlExcel.DuplicateLastLine("Services Prices", correctDownloadedFile.FullName);
            OpenXmlExcel.WriteDataInColumn("Action", "Services Prices", correctDownloadedFile.FullName, "M", CellValues.SharedString, true);
            OpenXmlExcel.WriteDataInColumn("Price", "Services Prices", correctDownloadedFile.FullName, "51.0", CellValues.Number, true);
            OpenXmlExcel.WriteDataInColumn("Cycle", "Services Prices", correctDownloadedFile.FullName, "Cycle 2", CellValues.SharedString, true);
            OpenXmlExcel.WriteDataInColumn("Cycle order", "Services Prices", correctDownloadedFile.FullName, "2", CellValues.Number, true);
            OpenXmlExcel.DuplicateLastLine("Services Prices", correctDownloadedFile.FullName);
            OpenXmlExcel.WriteDataInColumn("Action", "Services Prices", correctDownloadedFile.FullName, "M", CellValues.SharedString, true);
            OpenXmlExcel.WriteDataInColumn("Price", "Services Prices", correctDownloadedFile.FullName, "61.0", CellValues.Number, true);
            OpenXmlExcel.WriteDataInColumn("Cycle", "Services Prices", correctDownloadedFile.FullName, "Cycle 1", CellValues.SharedString, true);
            OpenXmlExcel.WriteDataInColumn("Cycle order", "Services Prices", correctDownloadedFile.FullName, "1", CellValues.Number, true);

            //1.Je charge le fichier dans WR
            importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);

            //2.Le message suivant apparait
            //Verification step done successfully.
            //To add: X line
            //To update : X line
            Assert.AreEqual("Verification step done successfully.", importPopup.CheckFilePopupMessage());
            Assert.AreEqual("0 line", importPopup.CheckFilePopupToAdd());
            Assert.AreEqual("2 lines", importPopup.CheckFilePopupToUpdate());
            //Sur la ligne to update on doit voir apparaitre le nombre de lignes sur lesquelles on a mis M dans le champs Action.

            //3.Je clique sur le bouton import
            isImport = importPopup.Import();

            Assert.IsTrue(isImport, "import raté");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_ImportBadSiteCode()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var servicePage = homePage.GoToCustomers_ServicePage();


            bool newVersionPrint = true;
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);

            servicePage.ClearDownloads();

            // chercher ou sinon créer un nouveau Service, ajouter un Price
            string serviceName1 = "AUTO_TEST_SERVICE_A1_BadSiteCode";
            string datasheet = null;
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);
            servicePage.Filter(ServicePage.FilterType.Search, serviceName1);
            int counterService1 = servicePage.TableCount();
            DateTime From = DateUtils.Now.Date.AddDays(-7);
            DateTime To = DateUtils.Now.Date.AddDays(7);

            if (counterService1 == 0)
            {
                ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName1);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, From, To, "10", datasheet);
                servicePage = pricePage.BackToList();
            }

            servicePage.Filter(ServicePage.FilterType.Search, serviceName1 + "BLAH");
            servicePage.Filter(ServicePage.FilterType.Search, serviceName1);

            // exporter (en oubliant pas l'imprimante en haut à droite)
            servicePage.Export(newVersionPrint);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            FileInfo correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            //Prices_20220104_0159.xlsx
            Assert.IsNotNull(correctDownloadedFile, "Fichier en entré non trouvé");

            int resultNumber = OpenXmlExcel.GetExportResultNumber("Services Prices", correctDownloadedFile.FullName);
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);

            // J'ajoute un SiteCode non existant
            OpenXmlExcel.WriteDataInColumn("Action", "Services Prices", correctDownloadedFile.FullName, "A", CellValues.SharedString, true);
            OpenXmlExcel.WriteDataInColumn("Site Code", "Services Prices", correctDownloadedFile.FullName, "PIRLOUIT", CellValues.SharedString, true);

            //1.Je charge le fichier dans WR
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);

            //2.Un message d'erreur apparait en précisant la ligne de l'erreur et la raison de l'erreur.
            //3.Je clique sur le bouton Cancel
            var isImport = importPopup.ImportWithFail2();

            Assert.IsFalse(isImport, "import réussi malgrès nouveau Site Code inexistant");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_ImportBadPriceMethod()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();

            bool newVersionPrint = true;
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);

            servicePage.ClearDownloads();

            // chercher ou sinon créer un nouveau Service, ajouter un Price
            string serviceName1 = "AUTO_TEST_SERVICE_A1_BadPriceMethod";
            string datasheet = null;
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);
            servicePage.Filter(ServicePage.FilterType.Search, serviceName1);
            int counterService1 = servicePage.TableCount();
            DateTime From = DateUtils.Now.Date.AddDays(-7);
            DateTime To = DateUtils.Now.Date.AddDays(7);

            if (counterService1 == 0)
            {
                ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName1);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, From, To, "10", datasheet);
                servicePage = pricePage.BackToList();
            }

            servicePage.Filter(ServicePage.FilterType.Search, serviceName1 + "BLAH");
            servicePage.Filter(ServicePage.FilterType.Search, serviceName1);

            // exporter (en oubliant pas l'imprimante en haut à droite)
            servicePage.Export(newVersionPrint);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            FileInfo correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            //Prices_20220104_0159.xlsx
            Assert.IsNotNull(correctDownloadedFile, "Fichier en entré non trouvé");

            int resultNumber = OpenXmlExcel.GetExportResultNumber("Services Prices", correctDownloadedFile.FullName);
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);

            // Je mets un Price Method non valide
            OpenXmlExcel.WriteDataInColumn("Action", "Services Prices", correctDownloadedFile.FullName, "A", CellValues.SharedString, true);
            OpenXmlExcel.WriteDataInColumn("Price Method", "Services Prices", correctDownloadedFile.FullName, "PIRLOUIT", CellValues.SharedString, true);

            //1.Je charge le fichier dans WR
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);

            //2.Un message d'erreur apparait en précisant la ligne de l'erreur et la raison de l'erreur.
            //3.Je clique sur le bouton Cancel
            //Import can't be performed. Some error(s) have been detected while extracting data from file :
            //Row[2] : Price method name[PIRLOUIT] unknown.Expected values : Fixed,Scale,Cycle
            var isImport = importPopup.ImportWithFail2();
            Assert.IsFalse(isImport, "import réussi malgrès mauvais Price Method");

        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_ImportEmptyDateTo()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var servicePage = homePage.GoToCustomers_ServicePage();


            bool newVersionPrint = true;
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);

            servicePage.ClearDownloads();

            // chercher ou sinon créer un nouveau Service, ajouter un Price
            string serviceName1 = "AUTO_TEST_SERVICE_A1_EmptyDateTo";
            string datasheet = null;
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);
            servicePage.Filter(ServicePage.FilterType.Search, serviceName1);
            int counterService1 = servicePage.TableCount();
            DateTime From = DateUtils.Now.Date.AddDays(-7);
            DateTime To = DateUtils.Now.Date.AddDays(7);

            if (counterService1 == 0)
            {
                ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName1);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, From, To, "10", datasheet);
                servicePage = pricePage.BackToList();
            }

            servicePage.Filter(ServicePage.FilterType.Search, serviceName1 + "BLAH");
            servicePage.Filter(ServicePage.FilterType.Search, serviceName1);

            // exporter (en oubliant pas l'imprimante en haut à droite)
            servicePage.Export(newVersionPrint);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            FileInfo correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            //Prices_20220104_0159.xlsx
            Assert.IsNotNull(correctDownloadedFile, "Fichier en entré non trouvé");

            int resultNumber = OpenXmlExcel.GetExportResultNumber("Services Prices", correctDownloadedFile.FullName);
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);

            // Je laisse le champ vide.
            OpenXmlExcel.WriteDataInColumn("Action", "Services Prices", correctDownloadedFile.FullName, "A", CellValues.SharedString, true);
            OpenXmlExcel.WriteDataInColumn("To", "Services Prices", correctDownloadedFile.FullName, "", CellValues.SharedString, true);

            //1.Je charge le fichier dans WR
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);

            //2.Un message d'erreur apparait en précisant la ligne de l'erreur et la raison de l'erreur.
            //3.Je clique sur le bouton Cancel"
            var isImport = importPopup.ImportWithFail();
            Assert.IsFalse(isImport, "import réussi malgrès mauvais Price Method");

        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_ImportDateTo()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var servicePage = homePage.GoToCustomers_ServicePage();


            bool newVersionPrint = true;
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);

            servicePage.ClearDownloads();

            // chercher ou sinon créer un nouveau Service, ajouter un Price
            string serviceName1 = "AUTO_TEST_SERVICE_M1_DateTo";
            string datasheet = null;
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.CustomersCheckAll, true);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);
            servicePage.Filter(ServicePage.FilterType.Search, serviceName1);
            int counterService1 = servicePage.TableCount();
            DateTime From = DateUtils.Now.Date.AddDays(-7);
            DateTime To = DateUtils.Now.Date.AddDays(7);

            if (counterService1 == 0)
            {
                ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName1);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, From, To, "10", datasheet);
                servicePage = pricePage.BackToList();
            }

            servicePage.Filter(ServicePage.FilterType.Search, serviceName1 + "BLAH");
            servicePage.Filter(ServicePage.FilterType.Search, serviceName1);

            // exporter (en oubliant pas l'imprimante en haut à droite)
            servicePage.Export(newVersionPrint);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            FileInfo correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            //Prices_20220104_0159.xlsx
            Assert.IsNotNull(correctDownloadedFile, "Fichier en entré non trouvé");

            int resultNumber = OpenXmlExcel.GetExportResultNumber("Services Prices", correctDownloadedFile.FullName);
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);

            // chaque jour ce test évolue de 2 jours
            OpenXmlExcel.WriteDataInColumn("Action", "Services Prices", correctDownloadedFile.FullName, "M", CellValues.SharedString, true);
            OpenXmlExcel.WriteDataInColumn("To", "Services Prices", correctDownloadedFile.FullName, Convert.ToInt32(To.AddDays(2).ToOADate()).ToString(), CellValues.Date, true);

            //1.Je charge le fichier dans WR
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);

            //2.Le message suivant apparait selon les cas
            //Verification step done successfully.
            //To add : X line
            //To update : X line
            Assert.AreEqual("Verification step done successfully.", importPopup.CheckFilePopupMessage());
            Assert.AreEqual("0 line", importPopup.CheckFilePopupToAdd());
            Assert.AreEqual("1 line", importPopup.CheckFilePopupToUpdate());
            //Sur la ligne to add on doit voir apparaitre le nombre de lignes sur lesquelles on a mis M dans le champs Action.

            //3.Je clique sur le bouton import
            var isImport = importPopup.Import();
            Assert.IsTrue(isImport, "import raté");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_ImportBadVAT()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var servicePage = homePage.GoToCustomers_ServicePage();


            bool newVersionPrint = true;
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);

            servicePage.ClearDownloads();

            // chercher ou sinon créer un nouveau Service, ajouter un Price
            string serviceName1 = "AUTO_TEST_SERVICE_M1_BadVAT";
            string datasheet = null;
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);
            servicePage.Filter(ServicePage.FilterType.Search, serviceName1);
            int counterService1 = servicePage.TableCount();
            DateTime From = DateUtils.Now.Date.AddDays(-7);
            DateTime To = DateUtils.Now.Date.AddDays(7);

            if (counterService1 == 0)
            {
                ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName1);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, From, To, "10", datasheet);
                servicePage = pricePage.BackToList();
            }

            servicePage.Filter(ServicePage.FilterType.Search, serviceName1 + "BLAH");
            servicePage.Filter(ServicePage.FilterType.Search, serviceName1);

            // exporter (en oubliant pas l'imprimante en haut à droite)
            servicePage.Export(newVersionPrint);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            FileInfo correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            //Prices_20220104_0159.xlsx
            Assert.IsNotNull(correctDownloadedFile, "Fichier en entré non trouvé");

            int resultNumber = OpenXmlExcel.GetExportResultNumber("Services Prices", correctDownloadedFile.FullName);
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);

            //Je mets une mauvaise valeur
            OpenXmlExcel.WriteDataInColumn("Action", "Services Prices", correctDownloadedFile.FullName, "M", CellValues.SharedString, true);
            OpenXmlExcel.WriteDataInColumn("VAT", "Services Prices", correctDownloadedFile.FullName, "PIRLOUIT", CellValues.SharedString, true);

            //1.Je charge le fichier dans WR
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);

            //2.Un message d'erreur apparait en précisant la ligne de l'erreur et la raison de l'erreur.
            //3.Je clique sur le bouton Cancel
            //Import can't be performed. Some error(s) have been detected while extracting data from file :
            //Rows 2 : Tax Type named[PIRLOUIT] not found in database.
            var isImport = importPopup.ImportWithFail2();
            Assert.IsFalse(isImport, "import réussi malgrès mauvais VAT");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_ImportEmptyVAT()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var servicePage = homePage.GoToCustomers_ServicePage();


            bool newVersionPrint = true;
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);

            servicePage.ClearDownloads();

            // chercher ou sinon créer un nouveau Service, ajouter un Price
            string serviceName1 = "AUTO_TEST_SERVICE_M1_EmptyVAT";
            string datasheet = null;
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.CustomersCheckAll, true);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);
            servicePage.Filter(ServicePage.FilterType.Search, serviceName1);
            int counterService1 = servicePage.TableCount();
            DateTime From = DateUtils.Now.Date.AddDays(-7);
            DateTime To = DateUtils.Now.Date.AddDays(7);

            if (counterService1 == 0)
            {
                ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName1);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, From, To, "10", datasheet);
                servicePage = pricePage.BackToList();
            }

            servicePage.Filter(ServicePage.FilterType.Search, serviceName1 + "BLAH");
            servicePage.Filter(ServicePage.FilterType.Search, serviceName1);

            // exporter (en oubliant pas l'imprimante en haut à droite)
            servicePage.Export(newVersionPrint);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            FileInfo correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            //Prices_20220104_0159.xlsx
            Assert.IsNotNull(correctDownloadedFile, "Fichier en entré non trouvé");

            int resultNumber = OpenXmlExcel.GetExportResultNumber("Services Prices", correctDownloadedFile.FullName);
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);

            //Je mets une mauvaise valeur
            OpenXmlExcel.WriteDataInColumn("Action", "Services Prices", correctDownloadedFile.FullName, "M", CellValues.SharedString, true);
            OpenXmlExcel.WriteDataInColumn("VAT", "Services Prices", correctDownloadedFile.FullName, "", CellValues.SharedString, true);

            //1.Je charge le fichier dans WR
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);

            //2.Un message d'erreur apparait en précisant la ligne de l'erreur et la raison de l'erreur.
            //3.Je clique sur le bouton Cancel
            //Import can't be performed. Some error(s) have been detected while extracting data from file :
            //Rows 2 : Tax Type named[] not found in database.
            var isImport = importPopup.ImportWithFail2();
            Assert.IsFalse(isImport, "import réussi malgrès VAT empty");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_ImportEmptySiteCode()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var servicePage = homePage.GoToCustomers_ServicePage();


            bool newVersionPrint = true;
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);

            servicePage.ClearDownloads();

            // chercher ou sinon créer un nouveau Service, ajouter un Price
            string serviceName1 = "AUTO_TEST_SERVICE_M1_EmptySiteCode";
            string datasheet = null;
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);
            servicePage.Filter(ServicePage.FilterType.Search, serviceName1);
            int counterService1 = servicePage.TableCount();
            DateTime From = DateUtils.Now.Date.AddDays(-7);
            DateTime To = DateUtils.Now.Date.AddDays(7);

            if (counterService1 == 0)
            {
                ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName1);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, From, To, "10", datasheet);
                servicePage = pricePage.BackToList();
            }

            servicePage.Filter(ServicePage.FilterType.Search, serviceName1 + "BLAH");
            servicePage.Filter(ServicePage.FilterType.Search, serviceName1);

            // exporter (en oubliant pas l'imprimante en haut à droite)
            servicePage.Export(newVersionPrint);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            FileInfo correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            //Prices_20220104_0159.xlsx
            Assert.IsNotNull(correctDownloadedFile, "Fichier en entré non trouvé");

            int resultNumber = OpenXmlExcel.GetExportResultNumber("Services Prices", correctDownloadedFile.FullName);
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);

            //Je mets une mauvaise valeur
            OpenXmlExcel.WriteDataInColumn("Action", "Services Prices", correctDownloadedFile.FullName, "M", CellValues.SharedString, true);
            OpenXmlExcel.WriteDataInColumn("Site Code", "Services Prices", correctDownloadedFile.FullName, "", CellValues.SharedString, true);

            //1.Je charge le fichier dans WR
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);

            //2.Un message d'erreur apparait en précisant la ligne de l'erreur et la raison de l'erreur.
            //3.Je clique sur le bouton Cancel
            //Import can't be performed. Some error(s) have been detected while extracting data from file :
            //Cannot import as some forbidden values have been modified
            //Row[2] has a modified site.
            //Rows 2 : Site code[] not found in database.
            var isImport = importPopup.ImportWithFail2();
            Assert.IsFalse(isImport, "import réussi malgrès VAT empty");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_ImportEmptyServiceID()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var servicePage = homePage.GoToCustomers_ServicePage();


            bool newVersionPrint = true;
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);

            servicePage.ClearDownloads();

            // chercher ou sinon créer un nouveau Service, ajouter un Price
            string serviceName1 = "AUTO_TEST_SERVICE_M1_EmptyServiceId";
            string datasheet = null;
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Sites, site);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);
            servicePage.Filter(ServicePage.FilterType.Search, serviceName1);
            int counterService1 = servicePage.TableCount();
            DateTime From = DateUtils.Now.Date.AddDays(-7);
            DateTime To = DateUtils.Now.Date.AddDays(7);

            if (counterService1 == 0)
            {
                ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName1);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, From, To, "10", datasheet);
                servicePage = pricePage.BackToList();
            }

            servicePage.Filter(ServicePage.FilterType.Search, serviceName1 + "BLAH");
            servicePage.Filter(ServicePage.FilterType.Search, serviceName1);

            // exporter (en oubliant pas l'imprimante en haut à droite)
            servicePage.Export(newVersionPrint);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            FileInfo correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            //Prices_20220104_0159.xlsx
            Assert.IsNotNull(correctDownloadedFile, "Fichier en entré non trouvé");

            int resultNumber = OpenXmlExcel.GetExportResultNumber("Services Prices", correctDownloadedFile.FullName);
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);

            //Je laisse le champs vide
            OpenXmlExcel.WriteDataInColumn("Action", "Services Prices", correctDownloadedFile.FullName, "M", CellValues.SharedString, true);
            OpenXmlExcel.WriteDataInColumn("ServiceId", "Services Prices", correctDownloadedFile.FullName, null, CellValues.Number, true);

            //1.Je charge le fichier dans WR
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);
            Assert.AreEqual("Import can't be performed. Some error(s) have been detected while extracting data from file :", importPopup.CheckFilePopupMessageError());

            var isImport = importPopup.ImportWithFail();
            Assert.IsFalse(isImport, "import réussi malgrès Service empty");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_ImportEmptyPriceID()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var servicePage = homePage.GoToCustomers_ServicePage();


            bool newVersionPrint = true;
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);

            servicePage.ClearDownloads();

            // chercher ou sinon créer un nouveau Service, ajouter un Price
            string serviceName1 = "AUTO_TEST_SERVICE_M1_EmptyPriceId";
            string datasheet = null;
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);
            servicePage.Filter(ServicePage.FilterType.Search, serviceName1);
            int counterService1 = servicePage.TableCount();
            DateTime From = DateUtils.Now.Date.AddDays(-7);
            DateTime To = DateUtils.Now.Date.AddDays(7);

            if (counterService1 == 0)
            {
                ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName1);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, From, To, "10", datasheet);
                servicePage = pricePage.BackToList();
            }

            servicePage.Filter(ServicePage.FilterType.Search, serviceName1 + "BLAH");
            servicePage.Filter(ServicePage.FilterType.Search, serviceName1);

            // exporter (en oubliant pas l'imprimante en haut à droite)
            servicePage.Export(newVersionPrint);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            FileInfo correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            //Prices_20220104_0159.xlsx
            Assert.IsNotNull(correctDownloadedFile, "Fichier en entré non trouvé");

            int resultNumber = OpenXmlExcel.GetExportResultNumber("Services Prices", correctDownloadedFile.FullName);
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);

            //Je laisse le champs vide
            OpenXmlExcel.WriteDataInColumn("Action", "Services Prices", correctDownloadedFile.FullName, "M", CellValues.SharedString, true);
            OpenXmlExcel.WriteDataInColumn("PriceId", "Services Prices", correctDownloadedFile.FullName, null, CellValues.Number, true);
            OpenXmlExcel.DuplicateFirstLine("Services Prices", correctDownloadedFile.FullName);
            OpenXmlExcel.WriteDataInColumn("Action", "Services Prices", correctDownloadedFile.FullName, "A", CellValues.SharedString, true);

            //1.Je charge le fichier dans WR.
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);

            //2.Un message d'erreur s'affiche.
            //3.Je clique sur Cancel
            //Import can't be performed. Some error(s) have been detected while extracting data from file :
            //Import can't be performed. Some error(s) have been detected while extracting data from file :
            //Row[3] : PriceId column must be specified for a modify (M) operation.
            var isImport = importPopup.ImportWithFail2();
            Assert.IsFalse(isImport, "import réussi malgrès PriceID empty");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_IsActive_WrongValue()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var servicePage = homePage.GoToCustomers_ServicePage();


            bool newVersionPrint = true;
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);

            servicePage.ClearDownloads();

            // chercher ou sinon créer un nouveau Service, ajouter un Price
            string serviceName = "AUTO_TEST_SERVICE_M1A1_Is_WrongValue";
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            int counterService1 = servicePage.TableCount();
            if (counterService1 == 0)
            {
                ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.Date.AddDays(-7), DateUtils.Now.Date.AddDays(7));
                servicePage = pricePage.BackToList();
                // provoque la bonne recherche
                servicePage.Filter(ServicePage.FilterType.Search, serviceName + "BLAH");
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            }


            // exporter (en oubliant pas l'imprimante en haut à droite)
            servicePage.Export(newVersionPrint);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            FileInfo correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            //Prices_20220104_0159.xlsx
            Assert.IsNotNull(correctDownloadedFile, "Fichier en entré non trouvé");

            int resultNumber = OpenXmlExcel.GetExportResultNumber("Services Prices", correctDownloadedFile.FullName);
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);

            // Modif
            OpenXmlExcel.WriteDataInColumn("Action", "Services Prices", correctDownloadedFile.FullName, "M", CellValues.SharedString, true);
            // Boolean null => VRAI
            OpenXmlExcel.WriteDataInColumn("Is Active", "Services Prices", correctDownloadedFile.FullName, "PIRLOUIT", CellValues.SharedString, true);
            // Ajout
            OpenXmlExcel.DuplicateFirstLine("Services Prices", correctDownloadedFile.FullName);
            OpenXmlExcel.WriteDataInColumn("Action", "Services Prices", correctDownloadedFile.FullName, "A", CellValues.SharedString, true);
            OpenXmlExcel.WriteDataInColumn("ServiceId", "Services Prices", correctDownloadedFile.FullName, null, CellValues.Number, true);
            OpenXmlExcel.WriteDataInColumn("Service", "Services Prices", correctDownloadedFile.FullName, serviceName + "_" + DateUtils.Now.ToString("yyyy-MM-dd_HH-mm-ss"), CellValues.SharedString, true);

            //1.Je charge le fichier dans WR
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);

            //3.Je clique sur le bouton Cancel
            var isImport = importPopup.ImportWithFail2();
            Assert.IsFalse(isImport, "import avec succes alors que Is Active est mauvais");

        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_ImportDeleteService()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var servicePage = homePage.GoToCustomers_ServicePage();


            bool newVersionPrint = true;
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);

            servicePage.ClearDownloads();

            // chercher ou sinon créer un nouveau Service, ajouter un Price
            string serviceName = "AUTO_TEST_SERVICE_M1_DeleteService";
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            if (servicePage.CheckTotalNumber() == 0)
            {
                ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.Date.AddDays(-7), DateUtils.Now.Date.AddDays(7));
                servicePage = pricePage.BackToList();
                // provoque la bonne recherche
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.ShowAll, true);
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            }


            // exporter (en oubliant pas l'imprimante en haut à droite)
            servicePage.Export(newVersionPrint);
            taskDirectory.Refresh();
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            FileInfo correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            //Prices_20220104_0159.xlsx
            Assert.IsNotNull(correctDownloadedFile, "Fichier en entré non trouvé");

            int resultNumber = OpenXmlExcel.GetExportResultNumber("Services Prices", correctDownloadedFile.FullName);
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);

            OpenXmlExcel.WriteDataInColumn("Action", "Services Prices", correctDownloadedFile.FullName, "M", CellValues.SharedString, true);
            OpenXmlExcel.WriteDataInColumn("Service", "Services Prices", correctDownloadedFile.FullName, "", CellValues.SharedString, true);

            //1.Je charge le fichier dans WR
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);

            //2.Le message suivant apparait
            //Verification step done successfully.
            //To add: X line
            //To update : X line
            Assert.AreEqual("Verification step done successfully.", importPopup.CheckFilePopupMessage());
            Assert.AreEqual("0 line", importPopup.CheckFilePopupToAdd());
            Assert.AreEqual("1 line", importPopup.CheckFilePopupToUpdate());
            //Sur la ligne to update on doit voir apparaitre le nombre de lignes sur lesquelles on a mis M dans le champs Action.

            //3.Je clique sur le bouton import
            var isImport = importPopup.ImportWithFail();
            Assert.IsFalse(isImport, "import avec succes alors que Nom Service est vide");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_ImportDeletePrice()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string siteBis = TestContext.Properties["SiteBis"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var servicePage = homePage.GoToCustomers_ServicePage();


            bool newVersionPrint = true;
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);

            servicePage.ClearDownloads();

            // chercher ou sinon créer un nouveau Service, ajouter un /*/*Price*/*/
            string serviceName = "AUTO_TEST_SERVICE_M1_DeletePrice";
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            int counterService1 = servicePage.TableCount();
            if (counterService1 == 0)
            {
                servicePage.Filter(ServicePage.FilterType.CustomersUncheck, true);
                servicePage.Filter(ServicePage.FilterType.SitesUncheck, true);
                servicePage.Filter(ServicePage.FilterType.CategoriesUncheck, true);
            }
            counterService1 = servicePage.TableCount();
            if (counterService1 == 0)
            {
                servicePage.Filter(ServicePage.FilterType.CustomersCheckAll, true);
                servicePage.Filter(ServicePage.FilterType.SitesCheckAll, true);
                servicePage.Filter(ServicePage.FilterType.CategoriesCheckAll, true);
            }
            counterService1 = servicePage.TableCount();
            if (counterService1 == 0)
            {
                ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.Date.AddDays(-7), DateUtils.Now.Date.AddDays(7));
                // a price already exists for this period
                //priceModalPage = pricePage.AddNewCustomerPrice();
                //pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.Date.AddDays(-7), DateUtils.Now.Date.AddDays(7));
                servicePage = pricePage.BackToList();
                // provoque la bonne recherche
                servicePage.Filter(ServicePage.FilterType.Search, serviceName + "BLAH");
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            }
            else
            {
                ServicePricePage pricePage = servicePage.ClickOnFirstService();
                int nbPrices = WebDriver.FindElements(By.XPath("//*/table[@class='service-table']/tbody/tr")).Count;
                for (int p = 0; p < nbPrices; p++)
                {
                    pricePage.UnfoldAll();
                    pricePage.DeleteFirstPrice();
                }
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.Date.AddDays(-7), DateUtils.Now.Date.AddDays(7), "10");
                priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteBis, customer, DateUtils.Now.Date.AddDays(-7), DateUtils.Now.Date.AddDays(7), "20");
                servicePage = pricePage.BackToList();
                // provoque la bonne recherche
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            }


            // exporter (en oubliant pas l'imprimante en haut à droite)
            servicePage.Export(newVersionPrint);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            FileInfo correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            //Prices_20220104_0159.xlsx
            Assert.IsNotNull(correctDownloadedFile, "Fichier en entré non trouvé");

            int resultNumber = OpenXmlExcel.GetExportResultNumber("Services Prices", correctDownloadedFile.FullName);
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);

            OpenXmlExcel.WriteDataInColumn("Action", "Services Prices", correctDownloadedFile.FullName, "D", CellValues.SharedString, true);

            for (int phase = 3; phase < 4; phase++)
            {
                switch (phase)
                {
                    case 0:
                        //IsActive à TRUE
                        OpenXmlExcel.WriteDataInColumn("Is Active", "Services Prices", correctDownloadedFile.FullName, "VRAI", CellValues.Boolean, true);
                        //Price Is active à false
                        OpenXmlExcel.WriteDataInColumn("Price is Active", "Services Prices", correctDownloadedFile.FullName, "FAUX", CellValues.Boolean, true);
                        //date échéance à passé
                        string From = ((int)DateUtils.Now.AddDays(-60).ToOADate()).ToString();
                        string To = ((int)DateUtils.Now.AddDays(-30).ToOADate()).ToString();
                        OpenXmlExcel.WriteDataInColumn("From", "Services Prices", correctDownloadedFile.FullName, From, CellValues.Date, true);
                        OpenXmlExcel.WriteDataInColumn("To", "Services Prices", correctDownloadedFile.FullName, To, CellValues.Date, true);
                        break;
                    case 1:
                        //IsActive à false
                        OpenXmlExcel.WriteDataInColumn("Is Active", "Services Prices", correctDownloadedFile.FullName, "FAUX", CellValues.Boolean, true);
                        //Price Is active à TRUE
                        OpenXmlExcel.WriteDataInColumn("Price is Active", "Services Prices", correctDownloadedFile.FullName, "VRAI", CellValues.Boolean, true);
                        //date échéance à passé
                        string From2 = ((int)DateUtils.Now.AddDays(-60).ToOADate()).ToString();
                        string To2 = ((int)DateUtils.Now.AddDays(-30).ToOADate()).ToString();
                        OpenXmlExcel.WriteDataInColumn("From", "Services Prices", correctDownloadedFile.FullName, From2, CellValues.Date, true);
                        OpenXmlExcel.WriteDataInColumn("To", "Services Prices", correctDownloadedFile.FullName, To2, CellValues.Date, true);
                        break;
                    case 2:
                        //IsActive à false
                        OpenXmlExcel.WriteDataInColumn("Is Active", "Services Prices", correctDownloadedFile.FullName, "FAUX", CellValues.Boolean, true);
                        //Price Is active à false
                        OpenXmlExcel.WriteDataInColumn("Price is Active", "Services Prices", correctDownloadedFile.FullName, "FAUX", CellValues.Boolean, true);
                        //date échéance à PRESENT
                        string From3 = ((int)DateUtils.Now.AddDays(-7).ToOADate()).ToString();
                        string To3 = ((int)DateUtils.Now.AddDays(7).ToOADate()).ToString();
                        OpenXmlExcel.WriteDataInColumn("From", "Services Prices", correctDownloadedFile.FullName, From3, CellValues.Date, true);
                        OpenXmlExcel.WriteDataInColumn("To", "Services Prices", correctDownloadedFile.FullName, To3, CellValues.Date, true);
                        break;
                    case 3:
                        //IsActive à false
                        OpenXmlExcel.WriteDataInColumn("Is Active", "Services Prices", correctDownloadedFile.FullName, "FAUX", CellValues.Boolean, true);
                        //Price Is active à false
                        OpenXmlExcel.WriteDataInColumn("Price is Active", "Services Prices", correctDownloadedFile.FullName, "FAUX", CellValues.Boolean, true);
                        //date échéance à passé
                        string From4 = ((int)DateUtils.Now.AddDays(-60).ToOADate()).ToString();
                        string To4 = ((int)DateUtils.Now.AddDays(-30).ToOADate()).ToString();
                        OpenXmlExcel.WriteDataInColumn("From", "Services Prices", correctDownloadedFile.FullName, From4, CellValues.Date, true);
                        OpenXmlExcel.WriteDataInColumn("To", "Services Prices", correctDownloadedFile.FullName, To4, CellValues.Date, true);
                        break;
                }

                //1.Je charge le fichier dans WR
                var importPopup = servicePage.Import();
                importPopup.CheckFile(correctDownloadedFile.FullName);

                //2.Le message suivant apparait
                //Verification step done successfully.
                //To add: X line
                //To update : X line
                Assert.AreEqual("Verification step done successfully.", importPopup.CheckFilePopupMessage());
                Assert.AreEqual("0 line", importPopup.CheckFilePopupToAdd());
                Assert.AreEqual("0 line", importPopup.CheckFilePopupToUpdate());
                bool zap = false;
                if (phase < 3)
                {
                    //FIXME normalement 1 line pour le résumé
                    Assert.AreEqual("0 line", importPopup.CheckFilePopupToDelete());
                    zap = true; // phase 1
                }
                else
                {
                    Assert.AreEqual("1 line", importPopup.CheckFilePopupToDelete());
                }
                //Sur la ligne to update on doit voir apparaitre le nombre de lignes sur lesquelles on a mis M dans le champs Action.

                //3.Je clique sur le bouton import
                var isImport = importPopup.ImportWithFail();
                if (phase < 3 && !zap)
                {
                    string[] message = {"suppression d'un Price avec IsActive à true",
                                        "suppression d'un Price avec Price Is active à true",
                                        "suppression d'un Price avec date échéance à présent"
                                        };
                    Assert.IsFalse(isImport, message[phase]);
                }
                else
                {
                    Assert.IsTrue(isImport, "suppression d'un Price");
                }


            }
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            ServicePricePage servicePricePage = servicePage.ClickOnFirstService();
            int nbPricesApres = WebDriver.FindElements(By.XPath("//*/table[@class='service-table']/tbody/tr")).Count;
            Assert.AreEqual(1, nbPricesApres, "Lignes Price non supprimée(s)");


        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_ImportAlreadyExistsService()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var servicePage = homePage.GoToCustomers_ServicePage();


            bool newVersionPrint = true;
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);

            servicePage.ClearDownloads();

            // chercher ou sinon créer un nouveau Service, ajouter un Price
            string serviceName = "AUTO_TEST_SERVICE_A1_AlreadyExistsService";
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            int counterService1 = servicePage.TableCount();
            if (counterService1 == 0)
            {
                ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.Date.AddDays(-7), DateUtils.Now.Date.AddDays(7));
                servicePage = pricePage.BackToList();
                // provoque la bonne recherche
                servicePage.Filter(ServicePage.FilterType.Search, serviceName + "BLAH");
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            }


            // exporter (en oubliant pas l'imprimante en haut à droite)
            servicePage.Export(newVersionPrint);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            FileInfo correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            //Prices_20220104_0159.xlsx
            Assert.IsNotNull(correctDownloadedFile, "Fichier en entré non trouvé");

            int resultNumber = OpenXmlExcel.GetExportResultNumber("Services Prices", correctDownloadedFile.FullName);
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);

            OpenXmlExcel.WriteDataInColumn("Action", "Services Prices", correctDownloadedFile.FullName, "A", CellValues.SharedString, true);
            OpenXmlExcel.WriteDataInColumn("ServiceId", "Services Prices", correctDownloadedFile.FullName, null, CellValues.Number, true);
            OpenXmlExcel.WriteDataInColumn("Service", "Services Prices", correctDownloadedFile.FullName, serviceName, CellValues.SharedString, true);

            //1.Je charge le fichier dans WR
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);

            //2.Un message d'erreur apparait en précisant la ligne de l'erreur et la raison de l'erreur.
            //Import can't be performed. Some error(s) have been detected while extracting data from file :
            //Rows 2 : Cannot create a new service named[AUTO_TEST_SERVICE_A1_AlreadyExistsService] because there is already a service named like this.

            //3.Je clique sur le bouton Cancel
            var isImport = importPopup.ImportWithFail2();
            Assert.IsFalse(isImport, "import avec succes alors que Nom Service existe déjà");

        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_ImportAlreadyExistsPriceID()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var servicePage = homePage.GoToCustomers_ServicePage();


            bool newVersionPrint = true;
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);

            servicePage.ClearDownloads();

            // chercher ou sinon créer un nouveau Service, ajouter un Price
            string serviceName = "AUTO_TEST_SERVICE_A1_AlreadyExistsPrice";
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            int counterService1 = servicePage.TableCount();
            if (counterService1 == 0)
            {
                ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.Date.AddDays(-7), DateUtils.Now.Date.AddDays(7));
                servicePage = pricePage.BackToList();
                // provoque la bonne recherche
                servicePage.Filter(ServicePage.FilterType.Search, serviceName + "BLAH");
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            }


            // exporter (en oubliant pas l'imprimante en haut à droite)
            servicePage.Export(newVersionPrint);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            FileInfo correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            //Prices_20220104_0159.xlsx
            Assert.IsNotNull(correctDownloadedFile, "Fichier en entré non trouvé");

            int resultNumber = OpenXmlExcel.GetExportResultNumber("Services Prices", correctDownloadedFile.FullName);
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);

            OpenXmlExcel.WriteDataInColumn("Action", "Services Prices", correctDownloadedFile.FullName, "A", CellValues.SharedString, true);
            OpenXmlExcel.WriteDataInColumn("Price", "Services Prices", correctDownloadedFile.FullName, "12.50", CellValues.Number, true);

            //1.Je charge le fichier dans WR.
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);
            Assert.AreEqual("Verification step done successfully.", importPopup.CheckFilePopupMessage());
            Assert.AreEqual("1 line", importPopup.CheckFilePopupToAdd());
            Assert.AreEqual("0 line", importPopup.CheckFilePopupToUpdate());
            importPopup.CheckFileClickImportFile();

            //2.Un message d'erreur s'affiche.
            Assert.AreEqual("Import has been cancelled. Some error(s) have occurred during the process:", importPopup.CheckFilePopupMessageError());
            //Import has been cancelled. Some error(s) have occurred during the process:
            //Service AUTO_TEST_SERVICE_A1_AlreadyExistsPrice, Customer TVS, Site MAD
            //Row 2 : The 05/01/2022 > 19/01/2022 price is in conflict with the 05/01/2022 > 19/01/2022 price.

            //3.Je clique sur Cancel
            var isImport = importPopup.ImportWithFail();
            Assert.IsFalse(isImport, "import avec succes alors que PriceID existe déjà");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_ImportAddUpdateServicePriceID()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var servicePage = homePage.GoToCustomers_ServicePage();


            bool newVersionPrint = true;
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);

            servicePage.ClearDownloads();

            // chercher ou sinon créer un nouveau Service, ajouter un Price
            string serviceName = "AUTO_TEST_SERVICE_A1M1_AddUpdateServicePriceID";
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            int counterService1 = servicePage.TableCount();
            if (counterService1 == 0)
            {
                ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.Date.AddDays(-7), DateUtils.Now.Date.AddDays(7));
                servicePage = pricePage.BackToList();
                // provoque la bonne recherche
                servicePage.Filter(ServicePage.FilterType.Search, serviceName + "BLAH");
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            }


            // exporter (en oubliant pas l'imprimante en haut à droite)
            servicePage.Export(newVersionPrint);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            FileInfo correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            //Prices_20220104_0159.xlsx
            Assert.IsNotNull(correctDownloadedFile, "Fichier en entré non trouvé");

            int resultNumber = OpenXmlExcel.GetExportResultNumber("Services Prices", correctDownloadedFile.FullName);
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);

            string PriceIdNumber = "123";
            OpenXmlExcel.WriteDataInColumn("Action", "Services Prices", correctDownloadedFile.FullName, "A", CellValues.SharedString, true);
            OpenXmlExcel.WriteDataInColumn("PriceId", "Services Prices", correctDownloadedFile.FullName, PriceIdNumber, CellValues.SharedString, true);
            OpenXmlExcel.WriteDataInColumn("Price", "Services Prices", correctDownloadedFile.FullName, "12.50", CellValues.Number, true);
            // nouveau service
            OpenXmlExcel.WriteDataInColumn("Service", "Services Prices", correctDownloadedFile.FullName, serviceName + "_" + DateUtils.Now.ToString("yyyy-MM-dd_HH-mm-ss"), CellValues.SharedString, true);
            // nouveau serviceId à null
            OpenXmlExcel.WriteDataInColumn("ServiceId", "Services Prices", correctDownloadedFile.FullName, null, CellValues.Number, true);

            //1.Je charge le fichier dans WR
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);

            //2.Le message suivant apparait
            //Verification step done successfully.
            //To add: X line
            //To update : X line
            Assert.AreEqual("Verification step done successfully.", importPopup.CheckFilePopupMessage());
            Assert.AreEqual("1 line", importPopup.CheckFilePopupToAdd());
            Assert.AreEqual("0 line", importPopup.CheckFilePopupToUpdate());
            //Sur la ligne to add on doit voir apparaitre le nombre de lignes sur lesquelles on a mis A dans le champs Action.

            //3.Je clique sur le bouton import
            var isImport = importPopup.Import();
            Assert.IsTrue(isImport, "import raté");

            servicePage.Filter(ServicePage.FilterType.Search, serviceName + "BLAH");
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);

            // Contrer le "Prices_20220107_0151 (1).xlsx"
            // re-exporter (en oubliant pas l'imprimante en haut à droite)
            servicePage.Export(newVersionPrint);
            taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            //Prices_20220104_0159.xlsx
            Assert.IsNotNull(correctDownloadedFile, "Fichier en entré non trouvé");
            string secondTime = correctDownloadedFile.FullName;
            secondTime = secondTime.Replace(".xlsx", " (1).xlsx");
            FileInfo secondTimeFileInfo = new FileInfo(secondTime);
            if (secondTimeFileInfo.Exists)
            {
                correctDownloadedFile = secondTimeFileInfo;
            }

            resultNumber = OpenXmlExcel.GetExportResultNumber("Services Prices", correctDownloadedFile.FullName);
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);
            List<string> columnPriceId = OpenXmlExcel.GetValuesInList("PriceId", "Services Prices", correctDownloadedFile.FullName);
            List<string> checkDoublons = new List<string>();
            foreach (string priceId in columnPriceId)
            {
                if (PriceIdNumber == priceId)
                {
                    Assert.Fail("PriceId non regénéré");
                }
                else if (checkDoublons.Contains(priceId))
                {
                    Assert.Fail("Pas nouveau PriceId");
                }
                else
                {
                    checkDoublons.Add(priceId);
                }
            }
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_With_M_ActionPriceMethodEmptyValue()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var servicePage = homePage.GoToCustomers_ServicePage();


            bool newVersionPrint = true;
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);

            servicePage.ClearDownloads();

            // chercher ou sinon créer un nouveau Service, ajouter un Price
            string serviceName1 = "AUTO_TEST_SERVICE_A1_PriceMethodEmptyValue";
            string datasheet = null;
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);
            servicePage.Filter(ServicePage.FilterType.Search, serviceName1);
            int counterService1 = servicePage.TableCount();
            DateTime From = DateUtils.Now.Date.AddDays(-7);
            DateTime To = DateUtils.Now.Date.AddDays(7);

            if (counterService1 == 0)
            {
                ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName1);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, From, To, "10", datasheet);
                servicePage = pricePage.BackToList();
            }

            servicePage.Filter(ServicePage.FilterType.Search, serviceName1 + "BLAH");
            servicePage.Filter(ServicePage.FilterType.Search, serviceName1);

            // exporter (en oubliant pas l'imprimante en haut à droite)
            servicePage.Export(newVersionPrint);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            FileInfo correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            //Prices_20220104_0159.xlsx
            Assert.IsNotNull(correctDownloadedFile, "Fichier en entré non trouvé");

            int resultNumber = OpenXmlExcel.GetExportResultNumber("Services Prices", correctDownloadedFile.FullName);
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);

            // Je mets un Price Method non valide
            OpenXmlExcel.WriteDataInColumn("Action", "Services Prices", correctDownloadedFile.FullName, "M", CellValues.SharedString, true);
            OpenXmlExcel.WriteDataInColumn("Price Method", "Services Prices", correctDownloadedFile.FullName, "", CellValues.SharedString, true);

            //1.Je charge le fichier dans WR
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);

            //2.Un message d'erreur apparait en précisant la ligne de l'erreur et la raison de l'erreur.
            //3.Je clique sur le bouton Cancel
            //Import can't be performed. Some error(s) have been detected while extracting data from file :
            //Row [2] : Price method name [] unknown. Expected values : Fixed,Scale,Cycle
            var isImport = importPopup.ImportWithFail2();
            Assert.IsFalse(isImport, "import réussi malgrès Price Method vide");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Import_With_A_ActionCategoryManyLine_ValueEmptyError()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var servicePage = homePage.GoToCustomers_ServicePage();


            bool newVersionPrint = true;
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);

            servicePage.ClearDownloads();

            // chercher ou sinon créer un nouveau Service, ajouter un Price
            string serviceName1 = "AUTO_TEST_SERVICE_A1_CategoryManyLineValueEmptyError";
            string datasheet = null;
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);
            servicePage.Filter(ServicePage.FilterType.Search, serviceName1);
            int counterService1 = servicePage.TableCount();
            DateTime From = DateUtils.Now.Date.AddDays(-7);
            DateTime To = DateUtils.Now.Date.AddDays(7);

            if (counterService1 == 0)
            {
                ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName1);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, From, To, "10", datasheet);
                servicePage = pricePage.BackToList();
            }

            servicePage.Filter(ServicePage.FilterType.Search, serviceName1 + "BLAH");
            servicePage.Filter(ServicePage.FilterType.Search, serviceName1);

            // exporter (en oubliant pas l'imprimante en haut à droite)
            servicePage.Export(newVersionPrint);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            FileInfo correctDownloadedFile = servicePage.GetExcelFile(taskFiles, true);
            //Prices_20220104_0159.xlsx
            Assert.IsNotNull(correctDownloadedFile, "Fichier en entré non trouvé");

            int resultNumber = OpenXmlExcel.GetExportResultNumber("Services Prices", correctDownloadedFile.FullName);
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);

            OpenXmlExcel.WriteDataInColumn("Action", "Services Prices", correctDownloadedFile.FullName, "A", CellValues.SharedString, true);
            OpenXmlExcel.WriteDataInColumn("Service", "Services Prices", correctDownloadedFile.FullName, serviceName1 + "_" + DateUtils.Now.ToString("yyyy-MM-dd_HH-mm-ss"), CellValues.SharedString, true);
            // nouveau serviceId à null
            OpenXmlExcel.WriteDataInColumn("ServiceId", "Services Prices", correctDownloadedFile.FullName, null, CellValues.Number, true);
            OpenXmlExcel.WriteDataInColumn("Category", "Services Prices", correctDownloadedFile.FullName, "", CellValues.SharedString, true);

            //1. Je charge le fichier dans WR
            var importPopup = servicePage.Import();
            importPopup.CheckFile(correctDownloadedFile.FullName);

            //2.Le message suivant apparait
            //Verification step done successfully.
            //To add: X line
            //To update : X line
            //Sur la ligne to add on doit voir apparaitre le nombre de lignes sur lesquelles on a mis A dans le champs Action.

            //3.Je clique sur le bouton import
            //Import can't be performed. Some error(s) have been detected while extracting data from file :
            //Rows 2 : service category with code[] not found in database.
            var isImport = importPopup.ImportWithFail2();
            Assert.IsFalse(isImport, "import réussi malgrès Category vide");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_CreateServiceBOB()
        {
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            // chercher ou sinon créer un nouveau Service, ajouter un Price
            string serviceName1 = "AUTO_TEST_SERVICE_CreateServiceBob";
            string datasheet = "1900 ALMUERZO CALIENTE BC CHML-HMFCCHML DIC'14 (MAD)";
            string category = "BUY ON BOARD";
            string guestType = "BOB";

            //Log in
            var homePage = LogInAsAdmin();

            var servicePage = homePage.GoToCustomers_ServicePage();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);
            servicePage.Filter(ServicePage.FilterType.Search, serviceName1);
            int counterService1 = servicePage.TableCount();
            DateTime From = DateUtils.Now.Date.AddDays(-2);
            DateTime To = DateUtils.Now.Date.AddDays(2);
            if (counterService1 == 0)
            {
                ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName1, null, null, category, guestType);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, From, To, "10", datasheet, null, null);
                servicePage = pricePage.BackToList();
            }
            servicePage.Filter(ServicePage.FilterType.Search, serviceName1);
            ServicePricePage servicePricePage = servicePage.ClickOnFirstService();
            ServiceGeneralInformationPage generalInfo = servicePricePage.ClickOnGeneralInformationTab();
            var categoryPath = generalInfo.WaitForElementIsVisible(By.Id("CategoryId"));
            var categorySelectElement = new SelectElement(categoryPath);
            Assert.AreEqual(category, categorySelectElement.SelectedOption.Text, "Mauvaise relecture category");
            var guestTypePath = generalInfo.WaitForElementIsVisible(By.Id("GuestTypeId"));
            var guestTypeSelectElement = new SelectElement(guestTypePath);
            Assert.AreEqual(guestType, guestTypeSelectElement.SelectedOption.Text, "Mauvaise relecture guest type");
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Massivedelete_select()
        {
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string serviceName = "AUTO_TEST_SERVICE_" + DateTime.Now.ToString("dd_MM_yyyy-") + new Random().Next().ToString();
            string serviceName1 = serviceName + "_1";
            string serviceName2 = serviceName + "_2";
            string datasheet = "1900 ALMUERZO CALIENTE BC CHML-HMFCCHML DIC'14 (MAD)";
            string category = "BUY ON BOARD";
            string guestType = "BOB";

            // Arrange
            var homePage = LogInAsAdmin();

            var servicePage = homePage.GoToCustomers_ServicePage();

            // Create 2 services
            ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName1, null, null, category, guestType);
            ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
            ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
            ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
            pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.AddDays(-1), DateUtils.Now, "10", datasheet, null, null);

            servicePage = serviceGeneralInformationsPage.BackToList();

            serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName2, null, null, category, guestType);
            serviceGeneralInformationsPage = serviceCreateModalPage.Create();
            pricePage = serviceGeneralInformationsPage.GoToPricePage();
            priceModalPage = pricePage.AddNewCustomerPrice();
            pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.AddDays(-1), DateUtils.Now, "10", datasheet, null, null);

            //Find 2 services created
            servicePage = pricePage.BackToList();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            var total = servicePage.CheckTotalNumber();
            if (total == 2)
            {
                servicePage.ResetFilters();
                // Delete 2 services one by one
                //Delete service1
                var serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.From, DateUtils.Now.AddMonths(-1).ToString("dd/MM/yyyy"));
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.To, DateUtils.Now.AddMonths(1).ToString("dd/MM/yyyy"));
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
                total = servicePage.CheckTotalNumber();
                Assert.IsTrue(total == 1, "La suppression d'un service ne fonctionne pas.");
                //Delete service2
                serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.From, DateUtils.Now.AddMonths(-1).ToString("dd/MM/yyyy"));
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.To, DateUtils.Now.AddMonths(1).ToString("dd/MM/yyyy"));
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
                total = servicePage.CheckTotalNumber();
                Assert.IsTrue(total == 0, "La suppression d'un service ne fonctionne pas.");
                servicePage.ResetFilters();
            }
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Massivedelete_selectall()
        {
            // Arrange
            var homePage = LogInAsAdmin();

            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            var serviceMassiveDelete = servicePage.ClickMassiveDelete();
            serviceMassiveDelete.ClickSearchButton();
            serviceMassiveDelete.PageSize("8");
            int totalPageNumber = serviceMassiveDelete.CheckTotalPageNumber();
            if (totalPageNumber >= 2)
            {
                serviceMassiveDelete.ClickSelectAllButton();
                int selectcount = serviceMassiveDelete.CheckTotalSelectCount();
                Assert.IsTrue(serviceMassiveDelete.VerifySelectedAllByPage(1), "Erreur de Select All,Le nombre de lignes cochées et l'affichage du nombre de sélections ne sont pas égaux.");
                Assert.IsTrue(serviceMassiveDelete.VerifySelectedAllByPage(2), "Erreur de Select All,Le nombre de lignes cochées et l'affichage du nombre de sélections ne sont pas égaux.");
                Assert.IsTrue(serviceMassiveDelete.VerifySelectedAllByPage(totalPageNumber), "Erreur de Select All,Le nombre de lignes cochées et l'affichage du nombre de sélections ne sont pas égaux.");
                int TotalSelectedService = serviceMassiveDelete.CheckTotalNumberByPage(1) * (totalPageNumber - 1) + serviceMassiveDelete.CheckTotalNumberByPage(totalPageNumber);
                Assert.IsTrue(selectcount == TotalSelectedService, "Erreur de Select All,Le nombre de lignes cochées et l'affichage du nombre de sélections ne sont pas égaux.");
            }
            else
            {
                servicePage = serviceMassiveDelete.Cancel();
                // create 10 service for test by 2 page with page Size 8
                string customer = TestContext.Properties["CustomerLP"].ToString();
                string site = TestContext.Properties["Site"].ToString();
                string serviceName = "AUTO_TEST_SERVICE_" + DateTime.Now.ToString("dd_MM_yyyy-") + new Random().Next().ToString();
                string datasheet = "1900 ALMUERZO CALIENTE BC CHML-HMFCCHML DIC'14 (MAD)";
                string category = "BUY ON BOARD";
                string guestType = "BOB";
                List<string> ServiceToDeleteList = new List<string>();

                for (int i = 1; i <= 10; i++)
                {
                    var service = serviceName + "_" + i;
                    // Create  service i
                    ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                    serviceCreateModalPage.FillFields_CreateServiceModalPage(service, null, null, category, guestType);
                    ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                    ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                    ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                    pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.AddDays(-1), DateUtils.Now, "10", datasheet, null, null);
                    ServiceToDeleteList.Add(service);
                    servicePage = serviceGeneralInformationsPage.BackToList();
                }
                serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.PageSize("8");

                serviceMassiveDelete.ClickSelectAllButton();
                int selectcount = serviceMassiveDelete.CheckTotalSelectCount();
                totalPageNumber = serviceMassiveDelete.CheckTotalPageNumber();
                Assert.IsTrue(serviceMassiveDelete.VerifySelectedAllByPage(1), "Erreur de Select All,Le nombre de lignes cochées et l'affichage du nombre de sélections ne sont pas égaux.");
                Assert.IsTrue(serviceMassiveDelete.VerifySelectedAllByPage(2), "Erreur de Select All,Le nombre de lignes cochées et l'affichage du nombre de sélections ne sont pas égaux.");
                Assert.IsTrue(serviceMassiveDelete.VerifySelectedAllByPage(totalPageNumber), "Erreur de Select All,Le nombre de lignes cochées et l'affichage du nombre de sélections ne sont pas égaux.");
                int TotalSelectedService = serviceMassiveDelete.CheckTotalNumberByPage(1) * (totalPageNumber - 1) + serviceMassiveDelete.CheckTotalNumberByPage(totalPageNumber);
                Assert.IsTrue(selectcount == TotalSelectedService, "Erreur de Select All,Le nombre de lignes cochées et l'affichage du nombre de sélections ne sont pas égaux.");
                // Delete 10 services created 
                servicePage = serviceMassiveDelete.Cancel();
                ServiceToDeleteList.ForEach(service =>
                {
                    serviceMassiveDelete = servicePage.ClickMassiveDelete();
                    serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, service);
                    serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.From, DateUtils.Now.AddMonths(-1).ToString("dd/MM/yyyy"));
                    serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.To, DateUtils.Now.AddMonths(1).ToString("dd/MM/yyyy"));
                    serviceMassiveDelete.ClickSearchButton();
                    serviceMassiveDelete.DeleteFirstService();
                });
            }
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Massivedelete_searchByDate()
        {
            var from = DateTime.Today.AddMonths(-1);
            var to = DateTime.Today.AddMonths(1);
            // Arrange
            var homePage = LogInAsAdmin();

            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            var serviceMassiveDelete = servicePage.ClickMassiveDelete();
            serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.From, from.ToString("dd/MM/yyyy"));
            serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.To, to.ToString("dd/MM/yyyy"));
            serviceMassiveDelete.ClickSearchButton();
            serviceMassiveDelete.PageSize("100");
            int totalpage = serviceMassiveDelete.CheckTotalPageNumber();
            if (totalpage > 0)
            {
                Assert.IsTrue(serviceMassiveDelete.VerifyFromTo(from, to), "erreur de filtrage par date dates out of range");
                serviceMassiveDelete.GoToPage(totalpage);
                Assert.IsTrue(serviceMassiveDelete.VerifyFromTo(from, to), "erreur de filtrage par date dates out of range");
            }
            else
            {
                // Create 2 services for test

                servicePage = serviceMassiveDelete.Cancel();
                string customer = TestContext.Properties["CustomerLP"].ToString();
                string site = TestContext.Properties["Site"].ToString();
                string serviceName = "AUTO_TEST_SERVICE_" + DateTime.Now.ToString("dd_MM_yyyy-") + new Random().Next().ToString();
                string serviceName1 = serviceName + "_1";
                string serviceName2 = serviceName + "_2";
                string datasheet = "1900 ALMUERZO CALIENTE BC CHML-HMFCCHML DIC'14 (MAD)";
                string category = "BUY ON BOARD";
                string guestType = "BOB";


                ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName1, null, null, category, guestType);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.AddDays(-1), DateUtils.Now, "10", datasheet, null, null);

                servicePage = serviceGeneralInformationsPage.BackToList();

                serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName2, null, null, category, guestType);
                serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                pricePage = serviceGeneralInformationsPage.GoToPricePage();
                priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.AddDays(-1), DateUtils.Now, "10", datasheet, null, null);
                servicePage = serviceGeneralInformationsPage.BackToList();

                serviceMassiveDelete = servicePage.ClickMassiveDelete();

                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.From, from.ToString("dd/MM/yyyy"));
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.To, to.ToString("dd/MM/yyyy"));
                serviceMassiveDelete.ClickSearchButton();
                bool result = serviceMassiveDelete.VerifyFromTo(from, to);
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName1);
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();
                serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName2);
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.From, from.ToString("dd/MM/yyyy"));
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.To, to.ToString("dd/MM/yyyy"));
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();
                Assert.IsTrue(result, "erreur de filtrage par date dates out of range");
            }
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Massivedelete_sitesearch()
        {
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();
            var servicePage = homePage.GoToCustomers_ServicePage();

            // Ouvrir la pop-up de suppression massive et appliquer un filtre par site
            var serviceMassiveDelete = servicePage.ClickMassiveDelete();
            serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.Site, site);
            serviceMassiveDelete.ClickSearchButton();
            serviceMassiveDelete.PageSize("100");
            int totalpage = serviceMassiveDelete.CheckTotalPageNumber();

            // Vérifier si des services sont trouvés pour le site spécifié
            if (totalpage > 0)
            {
                // Vérifier que tous les services trouvés correspondent au site sélectionné
                Assert.IsTrue(serviceMassiveDelete.VerifySiteSearch(site), "Erreur de filtrage par site");

                // Naviguer jusqu'à la dernière page et revérifier que les services appartiennent au site
                serviceMassiveDelete.GoToPage(totalpage);
                Assert.IsTrue(serviceMassiveDelete.VerifySiteSearch(site), "Erreur de filtrage par site");
            }
            else
            {
                // Aucun service trouvé : créer deux services de test
                var from = DateTime.Today.AddMonths(-1); // Date de début pour le filtre
                var to = DateTime.Today.AddMonths(1); // Date de fin pour le filtre
                servicePage = serviceMassiveDelete.Cancel();

                // Récupérer les informations pour la création de services de test
                string customer = TestContext.Properties["CustomerLP"].ToString();
                string serviceName = "AUTO_TEST_SERVICE_" + DateTime.Now.ToString("dd_MM_yyyy-") + new Random().Next().ToString();
                string serviceName1 = serviceName + "_1";
                string serviceName2 = serviceName + "_2";
                string datasheet = "1900 ALMUERZO CALIENTE BC CHML-HMFCCHML DIC'14 (MAD)";
                string category = "BUY ON BOARD";
                string guestType = "BOB";

                // Créer le premier service de test
                ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName1, null, null, category, guestType);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.AddDays(-1), DateUtils.Now, "10", datasheet, null, null);
                servicePage = serviceGeneralInformationsPage.BackToList();

                // Créer le deuxième service de test
                serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName2, null, null, category, guestType);
                serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                pricePage = serviceGeneralInformationsPage.GoToPricePage();
                priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.AddDays(-1), DateUtils.Now, "10", datasheet, null, null);
                servicePage = serviceGeneralInformationsPage.BackToList();

                // Ouvrir la pop-up de suppression massive et effectuer une recherche pour les nouveaux services créés
                serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.From, from.ToString("dd/MM/yyyy"));
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.To, to.ToString("dd/MM/yyyy"));
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.Site, site);
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.PageSize("100");

                // Vérifier que les services créés sont bien filtrés par site
                bool result = serviceMassiveDelete.VerifySiteSearch(site);

                // Supprimer le premier service de test en utilisant le filtre par nom
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName1);
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();

                // Supprimer le deuxième service de test en utilisant le filtre par nom et les dates
                serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName2);
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.From, from.ToString("dd/MM/yyyy"));
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.To, to.ToString("dd/MM/yyyy"));
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();

                // Validation finale pour s'assurer que le filtre fonctionne comme prévu
                Assert.IsTrue(result, "Erreur de filtrage par site");
            }
        }


        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Massivedelete_inactivecustomercheck()
        {
            // Arrange
            var homePage = LogInAsAdmin();

            var servicePage = homePage.GoToCustomers_ServicePage();

            ServiceMassiveDeleteModalPage massiveDeleteModal = servicePage.ClickMassiveDelete();

            massiveDeleteModal.ShowInactiveCustomers();

            massiveDeleteModal.Filter(ServiceMassiveDeleteFilterType.Customer, "Inactive - ");

            massiveDeleteModal.ClickOnCustomerFilter();

            bool IsChecked = massiveDeleteModal.IsInactive();

            Assert.IsTrue(IsChecked, "Le show inactive customer ne fonctionne pas.");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Massivedelete_inactivesitescheck()
        {
            // Arrange
            var homePage = LogInAsAdmin();

            var servicePage = homePage.GoToCustomers_ServicePage();

            ServiceMassiveDeleteModalPage massiveDeleteModal = servicePage.ClickMassiveDelete();

            massiveDeleteModal.ShowInactiveSites();

            massiveDeleteModal.Filter(ServiceMassiveDeleteFilterType.Site, "Inactive - ");

            massiveDeleteModal.ClickOnSiteFilter();

            Assert.IsTrue(massiveDeleteModal.IsInactive(), "Le show inactive site ne fonctionne pas.");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Massivedelete_namesearch()
        {
            string serviceName = "AUTO_TEST_SERVICE_A2M1_AddAndUpdate";

            // Arrange
            var homePage = LogInAsAdmin();

            var servicePage = homePage.GoToCustomers_ServicePage();

            ServiceMassiveDeleteModalPage massiveDeleteModal = servicePage.ClickMassiveDelete();

            massiveDeleteModal.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);

            massiveDeleteModal.ClickSearchButton();

            bool ServiceNameSearch = massiveDeleteModal.VerifyServiceNameSearch(serviceName);

            Assert.IsTrue(ServiceNameSearch, "La recherche par service name ne fonctionne pas.");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Massivedelete_inactivesitessearch()
        {
            string site = "MAD";
            string inactive_site = "Inactive - MAD";

            // Arrange
            var homePage = LogInAsAdmin();

            ParametersSites siteParameterPage = homePage.GoToParameters_Sites();
            siteParameterPage.Filter(ParametersSites.FilterType.SearchSite, site);
            siteParameterPage.DeactivateFirstSiteInList();

            homePage.Navigate();
            ParametersUser userParameter = homePage.GoToParameters_User();
            userParameter.SearchAndSelectUser("wr.testauto");
            userParameter.ClickOnAffectedSite();
            userParameter.ActivateSite(site);
            try
            {
                homePage.Navigate();
                var servicePage = homePage.GoToCustomers_ServicePage();
                var serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.ShowInactiveSites();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.Site, inactive_site);
                serviceMassiveDelete.ClickSearchButton();
                int totalpage = 0;
                var isData = serviceMassiveDelete.CheckData();
                if (isData)
                {
                    serviceMassiveDelete.PageSize("100");
                    totalpage = serviceMassiveDelete.CheckTotalPageNumber();
                }

                if (totalpage > 0)
                {
                    Assert.IsTrue(serviceMassiveDelete.VerifySiteSearch(site), "erreur de filtrage par site");
                    serviceMassiveDelete.GoToPage(totalpage);
                    Assert.IsTrue(serviceMassiveDelete.VerifySiteSearch(site), "erreur de filtrage par site");
                }
                else
                {
                    // Create 2 services for test
                    var from = DateTime.Today.AddMonths(-1);
                    var to = DateTime.Today.AddMonths(1);
                    //servicePage = serviceMassiveDelete.Cancel();
                    string customer = TestContext.Properties["CustomerLP"].ToString();
                    string serviceName = "AUTO_TEST_SERVICE_" + DateTime.Now.ToString("dd_MM_yyyy-") + new Random().Next().ToString();
                    string serviceName1 = serviceName + "_1";
                    string serviceName2 = serviceName + "_2";
                    string datasheet = "1900 ALMUERZO CALIENTE BC CHML-HMFCCHML DIC'14 (MAD)";
                    string category = "BUY ON BOARD";
                    string guestType = "BOB";

                    homePage.Navigate();
                    siteParameterPage = homePage.GoToParameters_Sites();
                    siteParameterPage.Filter(ParametersSites.FilterType.SearchSite, site);
                    siteParameterPage.ActivateFirstSiteInList();

                    homePage.Navigate();
                    servicePage = homePage.GoToCustomers_ServicePage();
                    ServiceCreateModalPage serviceCreateModalPage = servicePage.ServiceCreatePage();
                    serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName1, null, null, category, guestType);
                    ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                    ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                    ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                    priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.AddDays(-1), DateUtils.Now, "10", datasheet, null, null);

                    servicePage = serviceGeneralInformationsPage.BackToList();

                    serviceCreateModalPage = servicePage.ServiceCreatePage();
                    serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName2, null, null, category, guestType);
                    serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                    pricePage = serviceGeneralInformationsPage.GoToPricePage();
                    priceModalPage = pricePage.AddNewCustomerPrice();
                    priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.AddDays(-1), DateUtils.Now, "10", datasheet, null, null);


                    homePage.Navigate();
                    siteParameterPage = homePage.GoToParameters_Sites();
                    siteParameterPage.Filter(ParametersSites.FilterType.SearchSite, site);
                    siteParameterPage.Deactivate();

                    homePage.Navigate();
                    userParameter = homePage.GoToParameters_User();
                    userParameter.SearchAndSelectUser("wr.testauto");
                    userParameter.ClickOnAffectedSite();
                    userParameter.SelectOneSite(site);

                    homePage.Navigate();
                    servicePage = homePage.GoToCustomers_ServicePage();
                    serviceMassiveDelete = servicePage.ClickMassiveDelete();
                    serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.From, from.ToString("dd/MM/yyyy"));
                    serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.To, to.ToString("dd/MM/yyyy"));
                    serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.Site, inactive_site);
                    serviceMassiveDelete.ClickSearchButton();
                    serviceMassiveDelete.PageSize("100");
                    bool result = serviceMassiveDelete.VerifySiteSearch(site);
                    serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName1);
                    serviceMassiveDelete.ClickSearchButton();
                    serviceMassiveDelete.DeleteFirstService();
                    serviceMassiveDelete = servicePage.ClickMassiveDelete();
                    serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName2);
                    serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.From, from.ToString("dd/MM/yyyy"));
                    serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.To, to.ToString("dd/MM/yyyy"));
                    serviceMassiveDelete.ClickSearchButton();
                    serviceMassiveDelete.DeleteFirstService();
                    Assert.IsTrue(result, "erreur de filtrage par site");
                }
            }
            catch
            {
                homePage.Navigate();
                siteParameterPage = homePage.GoToParameters_Sites();
                siteParameterPage.Filter(ParametersSites.FilterType.SearchSite, site);
                siteParameterPage.ActivateFirstSiteInList();

                homePage.Navigate();
                userParameter = homePage.GoToParameters_User();
                userParameter.SearchAndSelectUser("wr.testauto");
                userParameter.ClickOnAffectedSite();
                userParameter.ActivateSite(site);
            }
            homePage.Navigate();
            siteParameterPage = homePage.GoToParameters_Sites();
            siteParameterPage.Filter(ParametersSites.FilterType.SearchSite, site);
            siteParameterPage.ActivateFirstSiteInList();

            homePage.Navigate();
            userParameter = homePage.GoToParameters_User();
            userParameter.SearchAndSelectUser("wr.testauto");
            userParameter.ClickOnAffectedSite();
            userParameter.ActivateSite(site);
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Massivedelete_OrganizeByCustomer()
        {
            //Log in
            var homePage = LogInAsAdmin();

            ServicePage servicePage = homePage.GoToCustomers_ServicePage();
            ServiceMassiveDeleteModalPage massiveDeleteModal = servicePage.ClickMassiveDelete();

            massiveDeleteModal.ClickSearchButton();

            massiveDeleteModal.ClickOnCustomerHeader();
            var firstComparison = massiveDeleteModal.VerifyCustomerCodeSort(SortType.Ascending);

            massiveDeleteModal.ClickOnCustomerHeader();
            var secondComparison = massiveDeleteModal.VerifyCustomerCodeSort(SortType.Descending);

            // Assert
            Assert.IsTrue(firstComparison, " Les Customers ne sont pas ordonnés ( de A à Z )");
            Assert.IsTrue(secondComparison, " Les Customers ne sont pas ordonnés ( de Z à A )");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Massivedelete_link()
        {
            //Log in
            var homePage = LogInAsAdmin();

            ServicePage servicePage = homePage.GoToCustomers_ServicePage();
            ServiceMassiveDeleteModalPage massiveDeleteModal = servicePage.ClickMassiveDelete();

            massiveDeleteModal.ClickSearchButton();

            string serviceName = massiveDeleteModal.GetFirstServiceName();
            Assert.IsTrue(serviceName != null, "Échec de la récupération du nom du service : aucun service trouvé pour la suppression massive.");

            bool linkIsClicked = massiveDeleteModal.ClickOnResultLinkAndCheckServiceName(serviceName);
            Assert.IsTrue(linkIsClicked, $"Le lien vers le service '{serviceName}' n'a pas ouvert la page dans un nouvel onglet.");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Massivedelete_pagination()
        {
            //Log in
            var homePage = LogInAsAdmin();

            ServicePage servicePage = homePage.GoToCustomers_ServicePage();
            ServiceMassiveDeleteModalPage massiveDeleteModal = servicePage.ClickMassiveDelete();

            massiveDeleteModal.ClickSearchButton();
            massiveDeleteModal.ClickSelectAllButton();

            int tot = massiveDeleteModal.CheckTotalSelectCount();

            massiveDeleteModal.PageSize("16");
            Assert.IsTrue(massiveDeleteModal.IsPageSizeEqualsTo("16"), "Pagination ne fonctionne pas.");
            Assert.AreEqual(massiveDeleteModal.GetTotalRowsForPagination(), tot >= 16 ? 16 : tot, "Pagination ne fonctionne pas.");

            massiveDeleteModal.PageSize("30");
            Assert.IsTrue(massiveDeleteModal.IsPageSizeEqualsTo("30"), "Pagination ne fonctionne pas.");
            Assert.AreEqual(massiveDeleteModal.GetTotalRowsForPagination(), tot >= 30 ? 30 : tot, "Pagination ne fonctionne pas.");

            massiveDeleteModal.PageSize("50");
            Assert.IsTrue(massiveDeleteModal.IsPageSizeEqualsTo("50"), "Pagination ne fonctionne pas.");
            Assert.AreEqual(massiveDeleteModal.GetTotalRowsForPagination(), tot >= 50 ? 50 : tot, "Pagination ne fonctionne pas.");

            massiveDeleteModal.PageSize("100");
            Assert.IsTrue(massiveDeleteModal.IsPageSizeEqualsTo("100"), "Pagination ne fonctionne pas.");
            Assert.AreEqual(massiveDeleteModal.GetTotalRowsForPagination(), tot >= 100 ? 100 : tot, "Pagination ne fonctionne pas.");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Massivedelete_customersearch()
        {
            var customerToCheckValue = "SMARTWINGS, A.S. (TVS)";
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            var servicePage = homePage.GoToCustomers_ServicePage();
            var serviceMassiveDeleteModalPage = servicePage.ClickMassiveDelete();
            serviceMassiveDeleteModalPage.Filter(ServiceMassiveDeleteFilterType.Customer, customerToCheckValue);
            serviceMassiveDeleteModalPage.ClickSearchButton();
            var result = serviceMassiveDeleteModalPage.VerifyCustomerSearch(customerToCheckValue, "active");
            Assert.IsTrue(result, "search by customer is not working");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Massivedelete()
        {
            Random rnd = new Random();
            var serviceCode = rnd.Next(100000, 1000000).ToString();
            string serviceName = "service" + serviceCode;

            //login
            var homePage = LogInAsAdmin();

            // Création d'un nouveau service
            var servicePage = homePage.GoToCustomers_ServicePage();
            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode);
            serviceCreateModalPage.Create();

            servicePage = homePage.GoToCustomers_ServicePage();
            var serviceMassiveDeleteModalPage = servicePage.ClickMassiveDelete();

            serviceMassiveDeleteModalPage.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
            serviceMassiveDeleteModalPage.CheckAllCustomer();
            serviceMassiveDeleteModalPage.CheckAllSites();
            serviceMassiveDeleteModalPage.CheckAllStatus();
            serviceMassiveDeleteModalPage.ClickSearchButton();

            serviceMassiveDeleteModalPage.DeleteFirstService();

            // Vérification que le service a été supprimé
            servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            var result = servicePage.VerifyServiceExist(serviceName);

            // Assert : Vérifier que le service a bien été supprimé
            Assert.IsTrue(result, "Le service n'a pas été supprimé.");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Massivedelete_OrganizeByDate()
        {
            //Log in
            var homePage = LogInAsAdmin();

            var servicePage = homePage.GoToCustomers_ServicePage();
            var serviceMassiveDeleteModalPage = servicePage.ClickMassiveDelete();
            serviceMassiveDeleteModalPage.ClickSearchButton();
            //verification from filter
            serviceMassiveDeleteModalPage.ClickFromFilter();
            var resultFrom = serviceMassiveDeleteModalPage.VerifyFromDatesASC();
            Assert.IsTrue(resultFrom, "erreur de tri par date from");
            //verification to filter
            serviceMassiveDeleteModalPage.ClickToFilter();
            var resultTo = serviceMassiveDeleteModalPage.VerifyToDatesASC();
            Assert.IsTrue(resultTo, "erreur de tri par date to");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Massivedelete_inactivecustomersearch()
        {
            string customer = "SMARTWINGS, A.S. (TVS)";

            //Log in
            var homePage = LogInAsAdmin();
            CustomerPage customerPage = homePage.GoToCustomers_CustomerPage();
            customerPage.Filter(CustomerPage.FilterType.ShowAll, true);
            customerPage.Filter(CustomerPage.FilterType.Search, customer);
            CustomerGeneralInformationPage customerDetails = customerPage.SelectFirstCustomer();

            try
            {
                customerDetails.SetCustomerInactive();
                var servicePage = homePage.GoToCustomers_ServicePage();
                ServiceMassiveDeleteModalPage massiveDeleteModal = servicePage.ClickMassiveDelete();
                massiveDeleteModal.ShowInactiveCustomers();
                massiveDeleteModal.Filter(ServiceMassiveDeleteFilterType.Customer, customer);
                massiveDeleteModal.ClickSearchButton();
                int totalpage = 0;
                var isData = massiveDeleteModal.CheckData();
                if (isData)
                {
                    massiveDeleteModal.PageSize("100");
                    totalpage = massiveDeleteModal.CheckTotalPageNumber();
                }

                if (totalpage > 0)
                {
                    bool IscustomerInactive1 = massiveDeleteModal.VerifyCustomerSearch(customer, "Inactive");
                    Assert.IsTrue(IscustomerInactive1, "erreur de filtrage par site");
                    massiveDeleteModal.GoToPage(totalpage);
                    bool IscustomerInactive2 = massiveDeleteModal.VerifyCustomerSearch(customer, "Inactive");
                    Assert.IsTrue(IscustomerInactive2, "erreur de filtrage par site");
                }

            }
            finally
            {
                customerPage = homePage.GoToCustomers_CustomerPage();
                customerPage.Filter(CustomerPage.FilterType.ShowAll, true);
                customerPage.Filter(CustomerPage.FilterType.Search, customer);
                customerDetails = customerPage.SelectFirstCustomer();
                customerDetails.SetCustomerActive();
            }

        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Massivedelete_organizebysite()
        {
            //Log in
            var homePage = LogInAsAdmin();

            ServicePage servicePage = homePage.GoToCustomers_ServicePage();
            var massiveDelete = servicePage.ClickMassiveDelete();

            massiveDelete.SortBySite();
            bool IsSiteOrganized = massiveDelete.IsSortedBySite();

            Assert.IsTrue(IsSiteOrganized, "Le sort by site n'est pas fonctionnel");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_SERV_Massivedelete_status_activeService()
        {
            string customer = TestContext.Properties["CustomerLP"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string serviceName = "AUTO_TEST_SERVICE_status_activeService";
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string datasheetName1 = TestContext.Properties["DatasheetName1"].ToString();

           


            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);
            if (servicePage.CheckTotalNumber() == 0)
            { 
             var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();
            serviceGeneralInformationsPage.GoToCustomers_ServicePage();
            }
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);
            var servicePageCount = servicePage.TableCount();
            ServiceMassiveDeleteModalPage massiveDeleteModal = servicePage.ClickMassiveDelete();
            massiveDeleteModal.WaitPageLoading();
            massiveDeleteModal.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
            massiveDeleteModal.Filter(ServiceMassiveDeleteFilterType.Status, "Without price");
            massiveDeleteModal.Filter(ServiceMassiveDeleteFilterType.Customer, customer);
            massiveDeleteModal.Filter(ServiceMassiveDeleteFilterType.Site, site);
            massiveDeleteModal.WaitPageLoading();
            //massiveDeleteModal.CheckOnlyActiveServicesStatus();
            var serviceList = massiveDeleteModal.FillAndGetDistinctServiceList();
            massiveDeleteModal.ClickSelectAllButton();
            massiveDeleteModal.DeleteService();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);
            var serviceListCount = servicePage.TableCount();

            Assert.AreNotEqual(servicePageCount, serviceListCount, "Massivedelete activeService ne fonctionne pas.");
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Massivedelete_organizebyService()
        {
            var homePage = LogInAsAdmin();

            ServicePage servicePage = homePage.GoToCustomers_ServicePage();
            var massiveDelete = servicePage.ClickMassiveDelete();

            massiveDelete.SortByService();

            bool IsServiceOrganised = massiveDelete.IsSortedByService();

            Assert.IsTrue(IsServiceOrganised, "Le sort by service ne fonctionne pas");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Price_NewPriceScale()
        {
            List<string> value_cscale_mode = new List<string> { "1", "2", "3", "4", "5", "6", "7", "8" };

            // Arrange
            var homePage = LogInAsAdmin();

            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.ShowOnlyActive, true);
            var servicePricePage = servicePage.ClickOnFirstService();
            var serviceCreatePriceModalPage = servicePricePage.AddNewCustomerPrice();
            var serviceLodingScalemode = serviceCreatePriceModalPage.SetModeScale();
            serviceLodingScalemode.SetValueForFirstScaleMode(value_cscale_mode[0], value_cscale_mode[1]);
            serviceLodingScalemode.AddManyScaleMode(3, value_cscale_mode);
            serviceCreatePriceModalPage = serviceLodingScalemode.Save_Scale_Mode();
            serviceLodingScalemode = serviceCreatePriceModalPage.ClickEditModeScale();

            bool sequence = value_cscale_mode.SequenceEqual(serviceLodingScalemode.GetAllValueScaleMode());
            Assert.IsTrue(sequence, "Les scales paramétrés n'apparaissent pas");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_GeneralInfoSPML()
        {
            // Prepare
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            //Arrange
            HomePage homePage= LogInAsAdmin();

            //Act
            try
            {
                // Navigate to Customer Service Page
                CustomerPage customerPage = homePage.GoToCustomers_CustomerPage();
                ServicePage servicePage = customerPage.GoToCustomers_ServicePage();
                servicePage.ResetFilters();

                // Create a new Service
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                // Add a Customer Price
                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10), "10", null, "srv");
                pricePage.UnfoldAll();

                // Navigate back to Service List
                servicePage = serviceGeneralInformationsPage.BackToList();

                // Select the newly created service and validate SPML selection
                var serviceliste = servicePage.SelectFirstRow();
                serviceliste.SelectGeneralInformation();
                serviceliste.SelectISSPML();
                serviceliste.SelectPrice();

                // Assert that the popup forces SPML selection
                Assert.IsTrue(serviceliste.IsVisible(), "The popup is displayed, requiring SPML selection before proceeding.");
            }
            finally
            {
                // Clean up by deleting the created service
                var servicePage = homePage.GoToCustomers_ServicePage();
                var serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();
                servicePage.ResetFilters();
            }
        }


        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_GeneralInfoProducedNo()
        {
            var serviceName = "ServiceForTest";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            ClearCache();
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);

            var servicePrice = servicePage.SelectFirstRowInCustmerService();
            var serviceGeneralInfo = servicePrice.ClickOnGeneralInformationTab();
            //Update
            var typeService = serviceGeneralInfo.GetServiceType();
            if (typeService == "Food")
            {
                serviceGeneralInfo.SetNotProduced();

            }
            //Assert
            bool isValidator = !serviceGeneralInfo.CheckValidator();
            Assert.IsTrue(isValidator, "Les validators n'apparaissent pas!");

        }
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Price_NewPricePourcentage()
        {
            string labelround = "Round";
            // Arrange
            var homePage = LogInAsAdmin();

            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.ShowOnlyActive, true);
            var servicePricePage = servicePage.ClickOnFirstService();
            var serviceCreatePriceModalPage = servicePricePage.AddNewCustomerPrice();
            serviceCreatePriceModalPage.SetPourcentageMode();
            Assert.IsTrue(serviceCreatePriceModalPage.Vrifylabelroundexist(labelround)
                && serviceCreatePriceModalPage.VerifyOptionsRoundExist(), "Les trois options supplémentaire n'apparaissent pas");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_LoadingPlan2()
        {
            //prepare
            string serviceName = "12279 APP, J, CARROT SOUP ETDE";

            //arrange
            HomePage homePage = LogInAsAdmin();

            // Act 
            CustomerPage customerPage = homePage.GoToCustomers_CustomerPage();
            ServicePage servicePage = customerPage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            var serviceliste = servicePage.SelectFirstRowInCustmerService();
            var serviceLoadingPlanTab = serviceliste.GoToLoadingPlanTab();
            var loadingPlanName = serviceLoadingPlanTab.GetLoadingPlanName();
            LoadingPlansGeneralInformationsPage flightLoadingPlanPage = serviceLoadingPlanTab.SelectFirstLoadingPlan();

            // Assert 
            var flightLoadingPlanName = flightLoadingPlanPage.GetFlightLoadingPlanName();
            Assert.AreEqual(loadingPlanName, flightLoadingPlanName, "La fenêtre ne se met pas à jour correctement dans le module Flight/Loading Plan.");
        }


        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_ResetFilter()
        {
            string servicename = "Service";
            string categoryname = "BEBIDAS";
            string customersname = "$$ - CAT Genérico";

            HomePage homePage = LogInAsAdmin();


            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();



            // Get the default filter values

            var ServicefirstNameBeforeFiltrage = servicePage.GetFirstServiceName();
            // Apply various filters
            servicePage.Filter(ServicePage.FilterType.Search, servicename);
            servicePage.Filter(ServicePage.FilterType.Categories, categoryname);
            servicePage.Filter(ServicePage.FilterType.Customers, customersname);

            var ServicefirstNameAfterFiltrage = servicePage.GetFirstServiceName();

            // Reset filters
            servicePage.ResetFilters();
            var ServicefirstNameAfterRester = servicePage.GetFirstServiceName();

            Assert.IsTrue(ServicefirstNameBeforeFiltrage != ServicefirstNameAfterFiltrage && ServicefirstNameBeforeFiltrage == ServicefirstNameAfterRester, "Filtrage sont reset");


        }

		[TestMethod]
        [Priority(2)]
        [Timeout(_timeout)]
        public void CU_SERV_Price_AddNewPrice()
        {
            string serviceName = "ServiceForTest";
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customerAT = TestContext.Properties["InvoiceCustomerAirportTax"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();
           
            try
            {
                var servicePage = homePage.GoToCustomers_ServicePage();
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
                ServicePricePage servicePricePage = servicePage.ClickOnFirstService();
                var numberOfPriceInHeader = servicePage.GetPricesItem();
                var serviceCreatePriceModalPage = servicePricePage.AddNewCustomerPrice();

                servicePricePage = serviceCreatePriceModalPage.FillFields_CustomerPrice(site, customerAT, DateUtils.Now.AddMonths(-2), DateUtils.Now.AddMonths(+4));

                var numberOfPriceInHeaderAfterCreate = servicePage.GetPricesItem();
                servicePage = servicePricePage.BackToList();
                Assert.AreNotEqual((numberOfPriceInHeaderAfterCreate), (numberOfPriceInHeader), "Price n'est pas ajoutée");
            }
            finally
            {
                var servicePage = homePage.GoToCustomers_ServicePage();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
                ServicePricePage servicePricePage = servicePage.ClickOnFirstService();
                var namePrice = servicePricePage.GetLastPriceName();
                servicePricePage.TogglePrice();
                servicePricePage.DeletePrice();
            }
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Price_Datasheet()
        {
            //Prepare
            string serviceName = "ServiceForTestDATASHEET";
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            DateTime fromDate = DateUtils.Now.AddDays(-10);
            DateTime toDate = DateUtils.Now.AddDays(10);
            string customer = TestContext.Properties["CustomerLpCart"].ToString();
            string datasheetNameDefault = "AAL 1787 BEV, TOMATO JUICE, CAN, 12 EA solo manipulacion (MAD/BCN)";

            //Arrange
            HomePage homePage= LogInAsAdmin();
            ServicePage servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            try
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, fromDate, toDate, "10", datasheetNameDefault);
                pricePage.ToggleLastPrice();
                pricePage.ClickPictoDataSheet();
                var windowIsOpened = pricePage.VerifyIfNewWindowIsOpened();
                //Assert
                Assert.IsTrue(windowIsOpened);
                pricePage.Go_To_New_Navigate();
            }
            finally
            {
                servicePage = homePage.GoToCustomers_ServicePage();
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
                var serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();
                servicePage.ResetFilters();
            }
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Price_NewPriceMethodScale()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string datasheetName1 = TestContext.Properties["DatasheetName1"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            ClearCache();
            ServicePage servicePage = homePage.GoToCustomers_ServicePage();


            try
            {
                DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);

                servicePage.ClearDownloads();
                servicePage.ResetFilters();
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10), "10", datasheetName1, "srv");


                ServiceCreatePriceModalPage serviceCreatePriceModal = pricePage.AddNewCustomerPrice();
                serviceCreatePriceModal.SetMethodWithoutSAVE("Scale", 4, "10");
                serviceCreatePriceModal.DeleteScaleWithoutSave(2);
                var scaleList = serviceCreatePriceModal.GetAllScalePrice();
                pricePage = serviceCreatePriceModal.Save();
                pricePage.UnfoldAll();
                ServiceCreatePriceModalPage serviceEditPriceModal = pricePage.ServiceEitPriceModal(1);
                var scaleListt = serviceEditPriceModal.GetAllScalePrice();
                serviceEditPriceModal.Close();
                pricePage.BackToList();
                servicePage.Filter(ServicePage.FilterType.ShowAll, true);
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
                Assert.IsTrue(scaleList.Count >= 2, "Les scales ne sont pas présents.");

            }
            finally
            {

                if (servicePage.CheckTotalNumber() >= 1)
                {
                    ServicePricePage servicePricePage = servicePage.ClickOnFirstService();
                    servicePricePage.DeletePriceForService();
                    servicePricePage.BackToList();
                    ServiceMassiveDeleteModalPage massiveDeleteModal = servicePage.ClickMassiveDelete();
                    massiveDeleteModal.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
                    massiveDeleteModal.Filter(ServiceMassiveDeleteFilterType.Customer, customer);
                    massiveDeleteModal.Filter(ServiceMassiveDeleteFilterType.Site, site);
                    massiveDeleteModal.Filter(ServiceMassiveDeleteFilterType.Status, "Without price");
                    massiveDeleteModal.ClickSearchButton();
                    massiveDeleteModal.DeleteServiceForPrice();
                }

            }
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_TotalServiceIndex()
        {
            string serviceName = "SERV_TotalServiceIndex";
            string customerName = "customerForservice";
            string customerTypeInflight = TestContext.Properties["CustomerType2"].ToString();
            string siteACE = TestContext.Properties["SiteACE"].ToString();
            DateTime fromDate = DateUtils.Now.AddDays(-10);
            DateTime toDate = DateUtils.Now.AddDays(10);
            string dataSheetName = "datasheetForService";
            // Arrange
            HomePage homePage = LogInAsAdmin();
            //Act

            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            var totalServicesDisplayedBeforeFiltrage = servicePage.TableCount();
            try
            { 
                ServiceCreateModalPage ServiceCreatePage = servicePage.ServiceCreatePage();
                ServiceCreatePage.FillFields_CreateServiceModalPage(serviceName);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = ServiceCreatePage.Create();
                ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteACE, customerName, fromDate, toDate, null, dataSheetName);
                servicePage = pricePage.BackToList();
                var totalServicesDisplayedafterFiltrage = servicePage.TableCount();
                //Assert
                Assert.AreNotEqual(totalServicesDisplayedBeforeFiltrage, totalServicesDisplayedafterFiltrage, "le filtrage n'est");
            }
            finally
            {
                homePage.GoToCustomers_ServicePage();
                ServiceMassiveDeleteModalPage massiveDeleteModal = servicePage.ClickMassiveDelete();
                massiveDeleteModal.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
                massiveDeleteModal.ClickSearchButton();
                massiveDeleteModal.DeleteServiceForPrice();

            }

        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_FoodPackets()
        {
            //prepare
            string site = TestContext.Properties["SiteACE"].ToString();
            var packet = "ATLANTIC";
            var customer = "UNITED AIRLINES, INC. SUC EN ESPAÑA";
            string serviceName = "ServiceForTest";

            //arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            var servicepricePage = servicePage.SelectFirstRowInCustmerService();

            var foodpacketpage = servicepricePage.GoToFoodPacketsTab();
            try
            {
                ServiceCreateFoodPacketModal createfoodpackets = foodpacketpage.CreateFoodPackets();
                createfoodpackets.FillServiceFoodPacket(site, packet, customer);

                //Assrt
                bool IsVisible = createfoodpackets.IsVisible();
                Assert.IsTrue(IsVisible, "La liste des menus liés au service est présente");
            }
            finally
            {
                foodpacketpage = servicepricePage.GoToFoodPacketsTab();
                foodpacketpage.DeleteRow();

            }
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_FoodPacketsCreation()
        {
            //prepare
            string site = TestContext.Properties["SiteACE"].ToString();
            var packet = "ATLANTIC";
            var customer = "UNITED AIRLINES, INC. SUC EN ESPAÑA";
            string serviceName = "ServiceForTest";

            //arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            var servicepricePage = servicePage.SelectFirstRowInCustmerService();
            var foodpacketpage = servicepricePage.GoToFoodPacketsTab();
            var number = foodpacketpage.GetNumberOfPacketFoodsInService();

            try
            {
                foodpacketpage.DeleteRows(number);
                var numberFoodsPacketsB = foodpacketpage.GetNumberOfPacketFoodsInService();
                ServiceCreateFoodPacketModal createfoodpackets = foodpacketpage.CreateFoodPackets();
                createfoodpackets.FillServiceFoodPacket(site, packet, customer);
                var numberOfFoodsPacketsA = foodpacketpage.GetNumberOfPacketFoodsInService();
                Assert.AreEqual(numberOfFoodsPacketsA, numberFoodsPacketsB + 1, "Food Packet n'est pas ajouté");
            }
            finally
            {
                foodpacketpage = servicepricePage.GoToFoodPacketsTab();
                foodpacketpage.DeleteRow();
            }
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_FoodPacketsDelete()
        {
            //prepare
            string site = TestContext.Properties["SiteACE"].ToString();
            var packet = "ATLANTIC";
            var customer = "UNITED AIRLINES INC SUC EN ESPAÑA TEST";

            //arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            string serviceName = "ServiceForTest";
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            var servicepricePage = servicePage.SelectFirstRowInCustmerService();
            var foodpacketpage = servicepricePage.GoToFoodPacketsTab();

            //Delete processing
            var numberOfFoodsPackets = foodpacketpage.GetNumberOfPacketFoodsInService();
            if (numberOfFoodsPackets == 0)
            {
                ServiceCreateFoodPacketModal createfoodpackets = foodpacketpage.CreateFoodPackets();
                createfoodpackets.FillServiceFoodPacket(site, packet, customer);
            }
            var numberBeforeDelete = foodpacketpage.GetNumberOfPacketFoodsInService();
            foodpacketpage.DeleteRow();

            //Assert
            var numberAfterDelete = foodpacketpage.GetNumberOfPacketFoodsInService();
            Assert.AreEqual(numberAfterDelete, numberBeforeDelete - 1, "Food Packet n'est pas supprimé");

        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_FoodPacketsModification()
        {
            //prepare 
            string site = TestContext.Properties["SiteACE"].ToString();
            var packet = "ATLANTIC BEVERAGES UPB";
            var customer = "UNITED AIRLINES INC SUC EN ESPAÑA TEST";
            string serviceName = "ServiceForTest";

            //arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            var servicepricePage = servicePage.SelectFirstRowInCustmerService();
            var foodpacketpage = servicepricePage.GoToFoodPacketsTab();

            //Modification processing
            try
            {
                var numberOfFoodsPackets = foodpacketpage.GetNumberOfPacketFoodsInService();
                ServiceCreateFoodPacketModal createfoodpackets = foodpacketpage.CreateFoodPackets();
                createfoodpackets.FillServiceFoodPacket(site, packet, customer);
                var packetValue = foodpacketpage.GetPacket();
                ServiceCreateFoodPacketModal foodPacketModal = foodpacketpage.EditItem();
                foodPacketModal.DeletePacket();
                foodPacketModal.FillPacketServiceFoodPacket(packet);

                //Assert
                var packetValueAfterModif = foodpacketpage.GetPacket();
                Assert.AreEqual(packetValue, packetValueAfterModif, "Le food packet n'est pas modifié ou un problème au niveau de suppression de packet");

            }
            finally
            {
                foodpacketpage = servicepricePage.GoToFoodPacketsTab();
                foodpacketpage.DeleteRow();
            }


        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Price_Calendrier()
        {
            string serviceName = "ServiceTestingCalendrier";
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            DateTime fromDate = DateUtils.Now.AddDays(-10);
            DateTime toDate = DateUtils.Now.AddDays(10);
            string customer = TestContext.Properties["CustomerLpCart"].ToString();
            string datasheetNameDefault = "AAL 1787 BEV, TOMATO JUICE, CAN, 12 EA solo manipulacion (MAD/BCN)";
            string price = "10";

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            ServicePage servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            try
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, fromDate, toDate, "10", datasheetNameDefault);
                pricePage.UnfoldAll();

                var mode = pricePage.GetMode();
                if (pricePage.GetMode() != "Cycle")
                {
                    pricePage.ClickFirstItem();
                    pricePage.ClickEditModeCycle();
                    pricePage.SetPrice(price);
                }
                pricePage.OpenModalCalender();
                bool verifyIfModalIsOpened = pricePage.VerifyIfModalIsOpened();
                Assert.IsTrue(verifyIfModalIsOpened, "Calendrier n'est pas fonctionnel ");
            }
            finally
            {
                servicePage = homePage.GoToCustomers_ServicePage();
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
                var serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();
                servicePage.ResetFilters();
            }

        }


        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Price_DuplicatePrice()
        {
            //Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string datasheetName1 = TestContext.Properties["DatasheetName1"].ToString();
            //Arrange
            LogInAsAdmin();

            //Act
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            ServicePage servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10), "10", datasheetName1, "srv");
            servicePage = serviceGeneralInformationsPage.BackToList();
            servicePage.ResetFilters();
            try
            {
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
                var numberPricebeforeDuplication = servicePage.GetActivePricesFirstService();
                ServicePricePage servicePricePage = servicePage.ClickOnFirstService();
                servicePricePage.DuplicatePrice();
                servicePricePage.UpdateSiteWhenDuplicatingPrice(site);
                servicePricePage.ConfirmDuplicatePrice();
                servicePricePage.BackToList();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
                var numberPriceAfterDuplication = servicePage.GetActivePricesFirstService();
                //Assert
                Assert.AreNotEqual(numberPriceAfterDuplication, numberPricebeforeDuplication, "Le prix n'a pas été dupliqué!");
            }
            finally
            {
                if (servicePage.CheckTotalNumber() >= 1)
                {
                    ServicePricePage servicePricePage = servicePage.ClickOnFirstService();
                    servicePricePage.DeletePriceForService();
                    servicePricePage.BackToList();
                    ServiceMassiveDeleteModalPage massiveDeleteModal = servicePage.ClickMassiveDelete();
                    massiveDeleteModal.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
                    massiveDeleteModal.Filter(ServiceMassiveDeleteFilterType.Customer, customer);
                    massiveDeleteModal.Filter(ServiceMassiveDeleteFilterType.Site, site);
                    massiveDeleteModal.Filter(ServiceMassiveDeleteFilterType.Status, "Without price");
                    massiveDeleteModal.ClickSearchButton();
                    massiveDeleteModal.DeleteServiceForPrice();
                }
            }
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Price_DuplicatePriceError()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            ServicePage servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            ServicePricePage servicePricePage = servicePage.ClickOnFirstService();
            servicePricePage.DuplicatePrice();
            servicePricePage.ConfirmDuplicatePrice();
            bool isShowed = servicePricePage.CheckErrorMessage();
            Assert.IsTrue(isShowed, "Le mesage d'erreur ne s'affiche pas");
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Delivery2()
        {

            // Prepare
            string serviceName = "ServiceForTest";
            string deliveryName = "deliveryForService";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            ServicePage servicePage = homePage.GoToCustomers_ServicePage();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);

            ServicePricePage servivePricePage = servicePage.ClickOnFirstService();
            ServiceDeliveryPage serviceDeliveryPage = servivePricePage.GoToDeliveryTab();
            var isDeliveryExist = serviceDeliveryPage.CheckDeliveriesExist();
            if (isDeliveryExist)
            {
                serviceDeliveryPage.ClickCrayon();
            }
            else
            {
                DeliveryPage deliveryPage = homePage.GoToCustomers_DeliveryPage();
                deliveryPage.ResetFilter();
                deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName);
                DeliveryLoadingPage deliveryLoadingPage = deliveryPage.ClickOnFirstDelivery();
                deliveryLoadingPage.AddService(serviceName);
                servicePage = homePage.GoToCustomers_ServicePage();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
                servivePricePage = servicePage.ClickOnFirstService();
                serviceDeliveryPage = servivePricePage.GoToDeliveryTab();
                serviceDeliveryPage.ClickCrayon();

            }
            var windowIsOpened = servicePage.VerifyIfNewWindowIsOpened();
            Assert.IsTrue(windowIsOpened);

        }
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Price_NewPriceMethodCycle()
        {
            // Prepare
            List<int> value_scale_mode = new List<int> { 1, 2, 3, 4, 5, 6, 7, 8 };
            List<string> list_parametre_prix = new List<string> { "1", "2", "3", "4", "5", "6", "7", "8", "1", "4", "6", "8", "10", "12", "14", "16" };
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            DateTime fromDate = DateUtils.Now;
            DateTime toDate = DateUtils.Now.AddDays(10);
            ServicePricePageEditScaleForCycle editscalecycle = new ServicePricePageEditScaleForCycle(WebDriver, TestContext);
            List<string> tab_value_after_save = new List<string>();

            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();

            // Créer un nouveau service
            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();

            pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, fromDate, toDate);
            pricePage.UnfoldAll();

            var editPriceModal = pricePage.EditFirstPrice(null, null);
            editPriceModal.SetMethodeCycle();

            editPriceModal.SetCycleMode();

            editPriceModal.SetNbCycle("2");

            int coef_row = 1;
            for (int i = 1; i <= editPriceModal.NbCycles(); i++)
            {
                int nb_row = 3;
                var datasheetName = TestContext.Properties["Needs_DatasheetName" + i].ToString();
                editPriceModal.SetTextDataSheet(i, datasheetName);

                editscalecycle = editPriceModal.GetEditScaleCycle(i);
                editscalecycle.SetValueForFirstScaleMode((value_scale_mode[0]).ToString(), (value_scale_mode[1] * coef_row).ToString());

                editscalecycle.AddManyScaleMode(nb_row, value_scale_mode, coef_row);
                editPriceModal = editscalecycle.Save_Scale_Mode();
                coef_row++;
            }

            pricePage = editPriceModal.Save_Scale_Mode();

            var editPriceModal_after_save = pricePage.EditFirstPrice(null, null);

            for (int i = 1; i <= editPriceModal_after_save.NbCycles(); i++)
            {
                editPriceModal_after_save.SelectMethodeScaleCycle(i);
                editscalecycle = editPriceModal_after_save.GetEditScaleCycleRow(i);
                editscalecycle.GetAllValueScaleMode(ref tab_value_after_save);
                editscalecycle.CloseEditModeCycle();
            }

            Assert.IsTrue(list_parametre_prix.SequenceEqual(tab_value_after_save), "Les paramétrages ne sont pas présents lors de la Rouvrir la ligne de prix.");
        }


        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Price_NewPriceMethodCycle2()
        {
            // Prepare
            List<int> value_scale_mode = new List<int> { 1, 2, 3, 4, 5, 6, 7, 8 };
            List<string> list_parametre_prix = new List<string> { "1", "2", "3", "4", "5", "6", "7", "8", "1", "4", "6", "8", "10", "12", "14", "16" };
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            DateTime fromDate = DateUtils.Now;
            DateTime toDate = DateUtils.Now.AddDays(10);
            ServicePricePageEditScaleForCycle editscalecycle = new ServicePricePageEditScaleForCycle(WebDriver, TestContext);
            List<string> tab_value_after_save = new List<string>();

            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();
            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, fromDate, toDate);
            pricePage.UnfoldAll();
            var editPriceModal = pricePage.EditFirstPrice(null, null);
            editPriceModal.SetMethodeCycle();
            editPriceModal.SetCycleMode_EndOfMonth();
            editPriceModal.SetNbCycle("2");
            int coef_row = 1;
            for (int i = 1; i <= editPriceModal.NbCycles(); i++)
            {
                int nb_row = 3;
                var datasheetName = TestContext.Properties["Needs_DatasheetName" + i].ToString();
                editPriceModal.SetTextDataSheet(i, datasheetName);
                editPriceModal.SetFirstDuration("1");
                editscalecycle = editPriceModal.GetEditScaleCycle(i);
                editscalecycle.SetValueForFirstScaleMode((value_scale_mode[0]).ToString(), (value_scale_mode[1] * coef_row).ToString());
                editscalecycle.AddManyScaleMode(nb_row, value_scale_mode, coef_row);
                editPriceModal = editscalecycle.Save_Scale_Mode();
                coef_row++;
            }
            pricePage = editPriceModal.Save_Scale_Mode();
            var editPriceModal_after_save = pricePage.EditFirstPrice(null, null);
            //recuperation des données sauvgarder
            for (int i = 1; i <= editPriceModal.NbCycles(); i++)
            {
                editPriceModal_after_save.SelectMethodeScaleCycle(i);
                editscalecycle = editPriceModal_after_save.GetEditScaleCycleRow(i);
                editscalecycle.GetAllValueScaleMode(ref tab_value_after_save);
                editscalecycle.CloseEditModeCycle();
            }
            Assert.IsTrue(list_parametre_prix.SequenceEqual(tab_value_after_save), "Les paramétrages ne sont pas présents lors de la Rouvrir la ligne de prix.");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_Price_NewPriceMethodCycle3()
        {
            // Prepare
            List<int> value_scale_mode = new List<int> { 1, 2, 3, 4, 5, 6, 7, 8 };
            List<string> list_parametre_prix = new List<string> { "1", "2", "3", "4", "5", "6", "7", "8", "1", "4", "6", "8", "10", "12", "14", "16" };
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            DateTime fromDate = DateUtils.Now;
            DateTime toDate = DateUtils.Now.AddDays(10);
            ServicePricePageEditScaleForCycle editscalecycle = new ServicePricePageEditScaleForCycle(WebDriver, TestContext);
            List<string> tab_value_after_save = new List<string>();

            // Arrange
            var homePage = LogInAsAdmin();
            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();
            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, fromDate, toDate);
            pricePage.UnfoldAll();
            var editPriceModal = pricePage.EditFirstPrice(null, null);
            editPriceModal.SetMethodeCycle();
            editPriceModal.SetCycleMode_FullMonth();
            editPriceModal.SetNbCycle("2");
            int coef_row = 1;
            for (int i = 1; i <= editPriceModal.NbCycles(); i++)
            {
                int nb_row = 3;
                var datasheetName = TestContext.Properties["Needs_DatasheetName" + i].ToString();
                editPriceModal.SetTextDataSheet(i, datasheetName);

                editscalecycle = editPriceModal.GetEditScaleCycle(i);
                editscalecycle.SetValueForFirstScaleMode((value_scale_mode[0]).ToString(), (value_scale_mode[1] * coef_row).ToString());
                editscalecycle.AddManyScaleMode(nb_row, value_scale_mode, coef_row);
                editPriceModal = editscalecycle.Save_Scale_Mode();
                coef_row++;
            }
            pricePage = editPriceModal.Save_Scale_Mode();
            var editPriceModal_after_save = pricePage.EditFirstPrice(null, null);
            //recuperation des données sauvgarder
            for (int i = 1; i <= editPriceModal.NbCycles(); i++)
            {
                editPriceModal_after_save.SelectMethodeScaleCycle(i);
                editscalecycle = editPriceModal_after_save.GetEditScaleCycleRow(i);
                editscalecycle.GetAllValueScaleMode(ref tab_value_after_save);
                editscalecycle.CloseEditModeCycle();
            }
            Assert.IsTrue(list_parametre_prix.SequenceEqual(tab_value_after_save), "Les paramétrages ne sont pas présents lors de la Rouvrir la ligne de prix.");
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_AffichageNomServiceLong()
        {
            // Prepare          
            string serviceName = serviceNameToday + "-" + new Random().Next(1000, 9999).ToString()
                    + "-" + Guid.NewGuid().ToString()
                    + "-" + DateTime.Now.Ticks.ToString()
                    + "-" + new Random().Next(100000, 999999).ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);

            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();
            servicePage = serviceGeneralInformationsPage.BackToList();
            //Assert
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.UncheckAllSites();
            servicePage.UncheckAllCustomers();

            bool serviceNameVerified = servicePage.GetFirstServiceName().Contains(serviceName);
            Assert.IsTrue(serviceNameVerified, $"Le nom du service '{serviceName}' apparaît tronqué dans la liste des services.");
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_AccessDetail()
        {
            // Prepare
            string customer = TestContext.Properties["MenuServiceCustomer"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            string variant = TestContext.Properties["MenuVariant"].ToString();

            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);

            DateTime fromDate = DateUtils.Now;
            DateTime toDate = DateUtils.Now.AddDays(10);

            string menuName = GenerateName(10);
            DateTime startDate = DateUtils.Now;
            DateTime endDate = DateUtils.Now.AddDays(10);

            // Arrange
            HomePage homePage= LogInAsAdmin();

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();

            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            priceModalPage.FillFields_CustomerPrice(site, customer, fromDate, toDate);

            //Create menu
            var menuPage = homePage.GoToMenus_Menus();
            var menuCreateModalPage = menuPage.MenuCreatePage();
            menuCreateModalPage.FillField_CreateNewMenu(menuName, startDate, endDate, site, variant, serviceName);

            servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);

            pricePage = servicePage.ClickOnFirstService();
            bool IsDetailPageLoaded = servicePage.IsDetailPageServiceLoaded();

            // Assert - Verify that the detail page is loaded
            Assert.IsTrue(IsDetailPageLoaded, "Failed to load the Service Detail page.");

            // Assert: Verify the URL contains the expected path
            string expectedUrlPart = "Customers/Service/Detail";
            Assert.IsTrue(WebDriver.Url.Contains(expectedUrlPart), $"The URL does not contain the expected path: {expectedUrlPart}");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_SERV_AffichageIHM()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.PageSize("16");
            var resultHeadService = servicePage.CheckServiceHeadTableIHM();
            var resultHeadCategorie = servicePage.CheckCategoryHeadTableIHM();
            var resultHeadActivePrice = servicePage.CheckActivePriceHeadTableIHM();
            var resultHeadMenu = servicePage.CheckMenuHeadTableIHM();
            var resultHeadDelivery = servicePage.CheckDeliveryHeadTableIHM();
            var resultHeadLoadingPlan = servicePage.CheckLoadingPlanHeadTableIHM();

            Assert.IsFalse(resultHeadService, "décalage dans la colonne Service");
            Assert.IsFalse(resultHeadCategorie, "décalage dans la colonne Categorie");
            Assert.IsFalse(resultHeadActivePrice, "décalage dans la colonne Active Price");
            Assert.IsFalse(resultHeadMenu, "décalage dans la colonne Menu");
            Assert.IsFalse(resultHeadDelivery, "décalage dans la colonne Delivery");
            Assert.IsFalse(resultHeadLoadingPlan, "décalage dans la colonne Loading Plan");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_SERV_AddPrice_Methode()
        {
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string customer = TestContext.Properties["CustomerLP"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            List<string> listLabelMethodeFixed = new List<string> { "Datasheet", "Invoice price" };
            List<string> listLabelMethodeScale = new List<string> { "Datasheet", "Nb of persons", "Total Price", "Unit Price", "New scale" };
            List<string> listLabelMethodeCycle = new List<string> { "Cycle mode", "Nb cycles", "Start at", "with duration", "Name", "Duration (days)", "Datasheet", "Method", "Price" };

            var homePage = LogInAsAdmin();

            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();

            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);

            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();
            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();

            try
            {
                priceModalPage.SetMethodeFixed();
                var allLabelMethodeFixed = priceModalPage.GetAll_LabelMethodeFixed();
                bool labelsFixedCorrect = allLabelMethodeFixed.All(label => listLabelMethodeFixed.Contains(label));

                Assert.IsTrue(
                    allLabelMethodeFixed.Count() == listLabelMethodeFixed.Count && labelsFixedCorrect,
                    $"Les champs correspondants à la méthode Fixed n'apparaissent pas correctement. Champs attendus : {string.Join(", ", listLabelMethodeFixed)}. Champs affichés : {string.Join(", ", allLabelMethodeFixed)}"
                );

                priceModalPage.SetMethodscale();
                var allLabelMethodeScale = priceModalPage.GetAll_LabelMethodeScale();
                bool labelsScaleCorrect = allLabelMethodeScale.All(label => listLabelMethodeScale.Contains(label));

                Assert.IsTrue(
                    allLabelMethodeScale.Count() == listLabelMethodeScale.Count && labelsScaleCorrect,
                    $"Les champs correspondants à la méthode Scale n'apparaissent pas correctement. Champs attendus : {string.Join(", ", listLabelMethodeScale)}. Champs affichés : {string.Join(", ", allLabelMethodeScale)}"
                );

                priceModalPage.SetMethodeCycle();
                var allLabelMethodeCycle = priceModalPage.GetAll_LabelMethodeCycle();
                bool labelsCycleCorrect = allLabelMethodeCycle.All(label => listLabelMethodeCycle.Contains(label));

                Assert.IsTrue(
                    allLabelMethodeCycle.Count() == listLabelMethodeCycle.Count && labelsCycleCorrect,
                    $"Les champs correspondants à la méthode Cycle n'apparaissent pas correctement. Champs attendus : {string.Join(", ", listLabelMethodeCycle)}. Champs affichés : {string.Join(", ", allLabelMethodeCycle)}"
                );

                priceModalPage.SetMethodeFixed();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(10));

            }
            finally
            {
                // Nettoyage : Suppression du service
                servicePage = homePage.GoToCustomers_ServicePage();
                var serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_SERV_GeneralInfo_Category()
        {
            //Prepare
            string serviceName = new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string category = TestContext.Properties["InvoiceService"].ToString();
            string status = "Without price";

            //Act
            HomePage homePage = LogInAsAdmin();
            ServicePage servicePage = homePage.GoToCustomers_ServicePage();
            try
            {
                ServiceCreateModalPage ServiceCreatePage = servicePage.ServiceCreatePage();
                ServiceCreatePage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = ServiceCreatePage.Create();
                string categoryFromIG = serviceGeneralInformationsPage.GetCategory();
                Assert.AreNotEqual(category, categoryFromIG, "Le champ 'Category' doit etre different ");
                serviceGeneralInformationsPage.SetCategory(category);
                serviceGeneralInformationsPage.GoToPricePage();
                serviceGeneralInformationsPage.SelectGeneralInformation();
                categoryFromIG = serviceGeneralInformationsPage.GetCategory();
                serviceGeneralInformationsPage.BackToList();
                Assert.AreEqual(category, categoryFromIG, "La modification du champ 'Category' doit etre prise en compte");
            }
            finally
            {
                //Delete service
                ServiceMassiveDeleteModalPage serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.Status, status);
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_SERV_GeneralInfo_ProductionName()
        {
            //Prepare
            string serviceName = new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string serviceProduction2 = GenerateName(6);
            string category = TestContext.Properties["InvoiceService"].ToString();
            string status = "Without price";

            //Act
            HomePage homePage = LogInAsAdmin();
            ServicePage servicePage = homePage.GoToCustomers_ServicePage();
            try
            {
                ServiceCreateModalPage ServiceCreatePage = servicePage.ServiceCreatePage();
                ServiceCreatePage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = ServiceCreatePage.Create();
                Assert.AreNotEqual(serviceProduction, serviceProduction2, "Le champ 'Production name' doit etre different ");
                serviceGeneralInformationsPage.SetProductionName(serviceProduction2);
                serviceGeneralInformationsPage.GoToPricePage();
                serviceGeneralInformationsPage.SelectGeneralInformation();
                string productionNameFromIG = serviceGeneralInformationsPage.GetProductionName();
                serviceGeneralInformationsPage.BackToList();
                Assert.AreEqual(serviceProduction2, productionNameFromIG, "La modification du champ 'Production name' doit etre prise en compte");
            }
            finally
            {
                //Delete service
                ServiceMassiveDeleteModalPage serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.Status, status);
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_SERV_GeneralInfo_ServiceCode()
        {
            //Prepare
            string serviceName = new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceCode2 = serviceCode + 2;
            string serviceProduction = GenerateName(4);
            string status = "Without price";

            //Act
            HomePage homePage = LogInAsAdmin();
            ServicePage servicePage = homePage.GoToCustomers_ServicePage();
            try
            {
                ServiceCreateModalPage ServiceCreatePage = servicePage.ServiceCreatePage();
                ServiceCreatePage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = ServiceCreatePage.Create();
                Assert.AreNotEqual(serviceCode, serviceCode2, $"Les champs ServiceCode: {serviceCode} et ServiceCode:{serviceCode2} ne sont pas differents.");
                serviceGeneralInformationsPage.SetServiceCode(serviceCode2);
                serviceGeneralInformationsPage.GoToPricePage();
                serviceGeneralInformationsPage.SelectGeneralInformation();
                string serviceCodeFromIG = serviceGeneralInformationsPage.GetServiceCode();
                Assert.AreEqual(serviceCode2, serviceCodeFromIG, "La modification du champ 'Service Code' n'est pas prise en compte");
                serviceGeneralInformationsPage.BackToList();
            }
            finally
            {
                //Delete service
                ServiceMassiveDeleteModalPage serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.Status, status);
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_SERV_GeneralInfo_GuestType()
        {
            //Prepare
            string guestType1 = TestContext.Properties["GuestTypeTrolley"].ToString();
            string guestType2 = TestContext.Properties["GuestTypeTrolley1"].ToString();
            string serviceName = new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string status = "Without price";

            //Act
            HomePage homePage = LogInAsAdmin();
            ServicePage servicePage = homePage.GoToCustomers_ServicePage();
            try
            {
                ServiceCreateModalPage ServiceCreatePage = servicePage.ServiceCreatePage();
                ServiceCreatePage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction,null, guestType1);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = ServiceCreatePage.Create();
                Assert.AreNotEqual(guestType1, guestType2, "Le champ 'Guest Type' doit etre different ");
                serviceGeneralInformationsPage.SetGuestType(guestType2);
                serviceGeneralInformationsPage.GoToPricePage();
                serviceGeneralInformationsPage.SelectGeneralInformation();
                string guestTypeFromIG = serviceGeneralInformationsPage.GetGuestType();
                serviceGeneralInformationsPage.BackToList();
                Assert.AreEqual(guestType2, guestTypeFromIG, "La modification du champ 'Guest Type' doit etre prise en compte");
            }
            finally
            {
                //Delete service
                ServiceMassiveDeleteModalPage serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.Status, status);
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_SERV_GeneralInfo_ServiceName()
        {
            //Prepare
            string serviceName = DateUtils.Now + new Random().Next().ToString();
            string serviceNewName = $"{serviceName} Updated";
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string status = "Without price";

            //Act
            HomePage homePage = LogInAsAdmin();
            ServicePage servicePage = homePage.GoToCustomers_ServicePage();
            try
            {
                ServiceCreateModalPage ServiceCreatePage = servicePage.ServiceCreatePage();
                ServiceCreatePage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = ServiceCreatePage.Create();
                Assert.AreNotEqual(serviceName, serviceNewName, $"Les champs ServiceName: {serviceName} et ServiceName:{serviceNewName} ne sont pas differents.");
                serviceGeneralInformationsPage.SetServiceName(serviceNewName);
                serviceGeneralInformationsPage.GoToPricePage();
                serviceGeneralInformationsPage.SelectGeneralInformation();
                string serviceNameFromIG = serviceGeneralInformationsPage.GetServiceName();
                Assert.AreEqual(serviceNewName, serviceNameFromIG, "La modification du champ 'Service Name' n'est pas prise en compte");
                serviceGeneralInformationsPage.BackToList();
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.ShowAll, true);
                servicePage.Filter(ServicePage.FilterType.Search, serviceNewName);
                bool isUpdated = servicePage.TableCount() == 1;
                Assert.IsTrue(isUpdated, "La modification du champ 'Service Name' n'est pas prise en compte");
            }
            finally
            {
                //Delete service
                ServiceMassiveDeleteModalPage serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceNewName);
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.Status, status);
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_SERV_GeneralInfo_ServiceType()
        {
            //Prepare
            string serviceType1 = TestContext.Properties["Production_ServiceType"].ToString();
            string serviceType2 = "Dotation";
            string serviceName = new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string status = "Without price";

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            ServicePage servicePage = homePage.GoToCustomers_ServicePage();
            try
            {
                ServiceCreateModalPage ServiceCreatePage = servicePage.ServiceCreatePage();
                ServiceCreatePage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction,null,null,serviceType1);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = ServiceCreatePage.Create();
                Assert.AreNotEqual(serviceType1, serviceType2, $"Les champs ServiceType: {serviceType1} et ServiceType:{serviceType2} ne sont pas differents.");
                serviceGeneralInformationsPage.SetServiceType(serviceType2);
                serviceGeneralInformationsPage.GoToPricePage();
                serviceGeneralInformationsPage.SelectGeneralInformation();
                string serviceTypeFromIG = serviceGeneralInformationsPage.GetServiceType();
                //Assert
                Assert.AreEqual(serviceType2, serviceTypeFromIG, "La modification du champ 'Service Type' n'est pas prise en compte");
                serviceGeneralInformationsPage.BackToList();
            }
            finally
            {
                //Delete service
                ServiceMassiveDeleteModalPage serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.Status, status);
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_SERV_GeneralInfo_ExternalId()
        {
            //Prepare
            string serviceName = new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string externalId = new Random().Next().ToString();
            string status = "Without price";

            //Act
            HomePage homePage = LogInAsAdmin();
            ServicePage servicePage = homePage.GoToCustomers_ServicePage();
            try
            {
                ServiceCreateModalPage ServiceCreatePage = servicePage.ServiceCreatePage();
                ServiceCreatePage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = ServiceCreatePage.Create();
                serviceGeneralInformationsPage.SetExternalId(externalId);
                serviceGeneralInformationsPage.GoToPricePage();
                serviceGeneralInformationsPage.SelectGeneralInformation();
                string externalIdFromIG = serviceGeneralInformationsPage.GetExternalId();
                serviceGeneralInformationsPage.BackToList();
                Assert.AreEqual(externalId, externalIdFromIG, "La modification du champ 'Service Type' doit etre prise en compte");
            }
            finally
            {
                //Delete service
                ServiceMassiveDeleteModalPage serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.Status, status);
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_SERV_GeneralInfo_TemplateId()
        {
            //Prepare
            string serviceName = new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string templateId = new Random().Next().ToString();
            string status = "Without price";

            //Arrange 
            HomePage homePage = LogInAsAdmin();

            //Act
            ServicePage servicePage = homePage.GoToCustomers_ServicePage();
            try
            {
                //Create service
                ServiceCreateModalPage ServiceCreatePage = servicePage.ServiceCreatePage();
                ServiceCreatePage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
                ServiceGeneralInformationPage serviceGeneralInformationsPage = ServiceCreatePage.Create();
                //update templateID
                serviceGeneralInformationsPage.SetTemplateId(templateId);
                serviceGeneralInformationsPage.GoToPricePage();
                serviceGeneralInformationsPage.SelectGeneralInformation();
                //get  templateID updated
                string templateIdFromIG = serviceGeneralInformationsPage.GetTemplateId();
                serviceGeneralInformationsPage.BackToList();
                //Assert
                Assert.AreEqual(templateId, templateIdFromIG, "La modification du champ 'Template Id' doit etre prise en compte");
            }
            finally
            {
                //Delete service
                ServiceMassiveDeleteModalPage serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.Status, status);
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();
            }
        }
    }
}
