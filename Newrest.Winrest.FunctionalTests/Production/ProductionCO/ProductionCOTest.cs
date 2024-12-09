using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Datasheet;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Sites;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.CustomerOrder;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.ProductionCO;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using FilterType = Newrest.Winrest.FunctionalTests.PageObjects.Customers.Customer.CustomerPage.FilterType;
using System.Globalization;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using static Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service.ServiceMassiveDeleteModalPage;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.OutputForm;
using Newrest.Winrest.FunctionalTests.PageObject.Parameters.Customer;

namespace Newrest.Winrest.FunctionalTests.Production
{
    [TestClass]
    public class ProductionCOTest : TestBase
    {
        private const int Timeout = 600000;
        private const string PCO_SITE = "CPU";
        private const string PCO_CUSTOMER_TYPE = "CPU";
        private const string PCO_CUSTOMER = "Customer Production CO";
        private const string PCO_CUSTOMER_CODE = "CPCO";
        private const string PCO_SERVICE = "Service Production CO";
        private const string PCO_SERVICEWITHDATASHEET = "Service Production CO With DataSheet";
        private const string PCO_SERVICE_VALIDE = "Service Production CO validé";
        private const string PCO_SERVICE_CATEGORY = "PLATOS CPU";
        private const string DATASHEETNAME = "DataSheetProdCO";
        private const string PCO_SERVICEWITHDATASHEET_VALIDE = "Service Production CO With DataSheet validé";
        string CUSTOMER_ORDER_WITH_DATASHEET;

        DateTime DateFrom = DateTime.Today;
        DateTime DateTo = DateTime.Today.AddMonths(5);

        [Priority(0)]
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PCO_PrepareDatasCustomerAndService()
        {
            // Prepare
            string customer = PCO_CUSTOMER;
            string customerCode = PCO_CUSTOMER_CODE;
            string customerType = PCO_CUSTOMER_TYPE;
            string site = PCO_SITE;
            string serviceName = PCO_SERVICE;
            string serviceName2 = PCO_SERVICEWITHDATASHEET;
            string serviceCategory = PCO_SERVICE_CATEGORY;
            var adminName = TestContext.Properties["Admin_UserName"].ToString();
            string userName = adminName.Substring(0, adminName.IndexOf("@"));

            string guestType = TestContext.Properties["DatasheetGuestType"].ToString();

            string datasheetName = DATASHEETNAME;
            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            ClearCache();

            // Create a new Site
            var parametersSites = homePage.GoToParameters_Sites();
            parametersSites.Filter(ParametersSites.FilterType.SearchSite, site);
            bool isSite = parametersSites.isSiteExists(site);

            if (!isSite)
            {
                var sitesModalPage = parametersSites.ClickOnNewSite();
                sitesModalPage.FillPrincipalField_CreationNewSite(site, "-", "-", "-", site);

                parametersSites.Filter(ParametersSites.FilterType.SearchSite, site);
                string id = parametersSites.CollectNewSiteID();
                // Affect it to the user
                var parametersUser = homePage.GoToParameters_User();
                parametersUser.SearchAndSelectUser(userName);
                parametersUser.ClickOnAffectedSite();
                parametersUser.GiveSiteRightsToUser(id, true, site);
            }

            //Parameters Customer
            //type of customer CPU
            var parametersCustomer = homePage.GoToParameters_CustomerPage();
            bool isCustomerType = parametersCustomer.isTypeOfCustomerExist(customerType);
            if (!isCustomerType)
            {
                parametersCustomer.AddNewTypeOfCustomer(customerType);
            }

            //service category families CPU
            parametersCustomer.GoToTab_ServiceCategoryFamilies();
            if (!parametersCustomer.isServiceCategoryFamiliesExist(site))
            {
                parametersCustomer.AddNewServiceCategoryFamilies(customerType);
            }

            //category Platos CPU
            parametersCustomer.GoToTab_Category();
            if (!parametersCustomer.isCategoryExist(serviceCategory))
            {
                parametersCustomer.AddNewService();
                parametersCustomer.AddService(serviceCategory);
            }

            // Act customer
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customer);

            if (customersPage.CheckTotalNumber() == 0)
            {
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customer, customerCode, customerType);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();
                customersPage.ResetFilters();
                customersPage.Filter(FilterType.Search, customer);
            }

            //Assert
            Assert.AreEqual(customer, customersPage.GetFirstCustomerName(), "Le customer n'est pas présent.");
            //Act datasheet for service with datasheet

            var datasheetPage = homePage.GoToMenus_Datasheet();
            datasheetPage.Filter(DatasheetPage.FilterType.DatasheetName, datasheetName);
            if (datasheetPage.CheckTotalNumber() == 0)
            {
                var datasheetCreateModalPage = datasheetPage.CreateNewDatasheet();
                var datasheetDetailPage = datasheetCreateModalPage.FillField_CreateNewDatasheet(datasheetName, guestType, site);
                datasheetPage = datasheetDetailPage.BackToList();
            }

            // Act service
            var servicesPage = homePage.GoToCustomers_ServicePage();
            servicesPage.ResetFilters();
            servicesPage.Filter(ServicePage.FilterType.Search, serviceName);

            if (servicesPage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicesPage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, new Random().Next().ToString(), GenerateName(4), serviceCategory);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.AddDays(-30), DateUtils.Now.AddDays(+30));
                servicesPage = pricePage.BackToList();
            }
            else
            {
                var pricePage = servicesPage.ClickOnFirstService();
                pricePage.SearchPriceForCustomer(customer, site, DateUtils.Now.AddDays(-30), DateUtils.Now.AddDays(+30));
                servicesPage = pricePage.BackToList();
            }

            servicesPage.ResetFilters();
            servicesPage.Filter(ServicePage.FilterType.Search, serviceName);

            Assert.IsTrue(servicesPage.GetFirstServiceName().Contains(serviceName), "Le service n'a pas été créé.");

            servicesPage.ResetFilters();
            servicesPage.Filter(ServicePage.FilterType.Search, serviceName2);

            if (servicesPage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicesPage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName2, new Random().Next().ToString(), GenerateName(4), serviceCategory);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.AddDays(-30), DateUtils.Now.AddDays(+30), null, datasheetName);

                servicesPage = pricePage.BackToList();
            }
            else
            {
                var pricePage = servicesPage.ClickOnFirstService();
                pricePage.SearchPriceForCustomer(customer, site, DateUtils.Now.AddDays(-30), DateUtils.Now.AddDays(+30));
                servicesPage = pricePage.BackToList();
            }

            servicesPage.ResetFilters();
            servicesPage.Filter(ServicePage.FilterType.Search, serviceName2);

            Assert.IsTrue(servicesPage.GetFirstServiceName().Contains(serviceName2), "Le service n'a pas été créé.");
        }

        [TestMethod]
        [Priority(1)]
        [Timeout(Timeout)]
        public void PR_PCO_NewProductionCOSearchCustomerOrder()
        {
            // Prepare
            string customer = PCO_CUSTOMER;
            string site = PCO_SITE;
            string serviceName = PCO_SERVICE;
            string serviceName2 = PCO_SERVICEWITHDATASHEET;
            string delivery = new Random().Next().ToString();
            string category = PCO_SERVICE_CATEGORY;
            List<String> service = new List<string>
            {
                serviceName,
                serviceName2
            };
            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Prepare : créer
            // *un customer order non validé
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.ResetFilter();
            if (CUSTOMER_ORDER_WITH_DATASHEET != null && CUSTOMER_ORDER_WITH_DATASHEET != string.Empty)
            {
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, CUSTOMER_ORDER_WITH_DATASHEET);
                customerOrderPage.DeleteCustomerOrder();
                customerOrderPage.ResetFilter();
            }
            var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
            customerOrderCreateModalPage.FillField_CreatNewCustomerOrderWOUTflight(site, customer, delivery, DateTime.Now.AddDays(6));
            var customerOrderDetailPage = customerOrderCreateModalPage.Submit();
            customerOrderDetailPage.AddNewMiltipleItemWithCategory(service, "10", category);
            var generalInfo = customerOrderDetailPage.ClickOnGeneralInformationTab();
            CUSTOMER_ORDER_WITH_DATASHEET = generalInfo.GetOrderNumber();
            customerOrderDetailPage.BackToList();

            var name = customerOrderPage.GetCustomerOrderNumber();

            //Assert
            Assert.IsTrue(name.Contains(CUSTOMER_ORDER_WITH_DATASHEET), "Le customer order n'a pas été créé.");

            //Act
            var productionCOPage = homePage.GoToProduction_ProductionCOPage();
            var productionCOCreateModalPage = productionCOPage.ProductionCOCreatePage();
            productionCOCreateModalPage.FillField_NewProductionCOSearch(site, DateFrom, DateTo);
            var checkServiceName = productionCOCreateModalPage.CheckServiceToProduce(serviceName);
            Assert.IsTrue(checkServiceName, "Le customer order non validé ne remonte pas le service à produire.");
            var checkServiceName2 = productionCOCreateModalPage.CheckServiceToProduce(serviceName2);
            Assert.IsTrue(checkServiceName2, "Le customer order non validé ne remonte pas le service à produire.");

        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PCO_NewProductionCOSearchValidatedCustomerOrder()
        {
            // Prepare
            string customer = PCO_CUSTOMER;
            string site = PCO_SITE;
            string serviceName = PCO_SERVICE_VALIDE;
            string delivery = new Random().Next().ToString();
            string category = PCO_SERVICE_CATEGORY;
            string serviceCategory = PCO_SERVICE_CATEGORY;

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act service
            var servicesPage = homePage.GoToCustomers_ServicePage();
            servicesPage.ResetFilters();
            servicesPage.Filter(ServicePage.FilterType.Search, serviceName);

            if (servicesPage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicesPage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, new Random().Next().ToString(), GenerateName(4), serviceCategory);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.AddDays(-30), DateUtils.Now.AddDays(+30));
                servicesPage = pricePage.BackToList();
            }
            else
            {
                var pricePage = servicesPage.ClickOnFirstService();
                pricePage.SearchPriceForCustomer(customer, site, DateUtils.Now.AddDays(-30), DateUtils.Now.AddDays(+30));
                servicesPage = pricePage.BackToList();
            }

            servicesPage.ResetFilters();
            servicesPage.Filter(ServicePage.FilterType.Search, serviceName);

            Assert.IsTrue(servicesPage.GetFirstServiceName().Contains(serviceName), "Le service n'a pas été créé.");

            // Prepare : créer un customer order validé
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
            customerOrderCreateModalPage.FillField_CreatNewCustomerOrderWOUTflight(site, customer, delivery, DateUtils.Now.AddDays(6));
            var customerOrderDetailPage = customerOrderCreateModalPage.Submit();
            customerOrderDetailPage.AddNewItemWithCategory(serviceName, "10", category);
            customerOrderDetailPage.ValidateCustomerOrder();

            var generalInfo = customerOrderDetailPage.ClickOnGeneralInformationTab();
            var customerOrderNumber = generalInfo.GetOrderNumber();
            customerOrderDetailPage.BackToList();

            var name = customerOrderPage.GetCustomerOrderNumber();

            //Assert
            Assert.IsTrue(name.Contains(customerOrderNumber), "Le customer order n'a pas été créé.");

            //Act
            var productionCOPage = homePage.GoToProduction_ProductionCOPage();
            var productionCOCreateModalPage = productionCOPage.ProductionCOCreatePage();
            productionCOCreateModalPage.FillField_NewProductionCOSearch(site, DateUtils.Now.AddDays(-3), DateUtils.Now.AddMonths(5));
            Assert.IsFalse(productionCOCreateModalPage.CheckServiceToProduce(serviceName), "Le customer order validé remonte le service à produire, alors qu'il ne devrait pas.");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PCO_SearchService()
        {
            //Prepare
            string customerOrderNumber = null;
            string site = PCO_SITE;
            string serviceName = "ServiceForProduction";
            string customer = "CPCO - Customer Production CO";
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            //Arrange
            var homePage = LogInAsAdmin();
            // Act
            try
            {
                var servicePage = homePage.GoToCustomers_ServicePage();
                servicePage.ResetFilters();


                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction, PCO_SERVICE_CATEGORY);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                var servicePricePage = serviceGeneralInformationsPage.GoToPricePage();

                var serviceCreatePriceModalPage = servicePricePage.AddNewCustomerPrice();

                servicePricePage = serviceCreatePriceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.AddMonths(-3), DateUtils.Now.AddMonths(+6));

                servicePage = serviceGeneralInformationsPage.BackToList();



                var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
                var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrderWOUTflight(site, "CPCO", "123", DateUtils.Now.AddMonths(3));
                var customerOrderDetailPage = customerOrderCreateModalPage.Submit();
                customerOrderDetailPage.AddNewItemWithCategory(serviceName, "10", PCO_SERVICE_CATEGORY);
                var generalInfo = customerOrderDetailPage.ClickOnGeneralInformationTab();
                customerOrderNumber = generalInfo.GetOrderNumber();
                customerOrderDetailPage.BackToList();


                var productionCOPage = homePage.GoToProduction_ProductionCOPage();
                //homePage.GoToProduction_ProductionCOPageModified();
                var productionCOCreateModalPage = productionCOPage.ProductionCOCreatePage();
                productionCOCreateModalPage.FillField_NewProductionCOSearch(site, DateUtils.Now.AddMonths(-2), DateUtils.Now.AddMonths(5));
                productionCOCreateModalPage.WaitPageLoading();
                var IsServicesExist = productionCOCreateModalPage.VerifierPresenceService(serviceName);
                productionCOCreateModalPage.WaitPageLoading();

                Assert.IsTrue(IsServicesExist, "La Pop-up n' afficher pas les services prévus.");
                productionCOCreateModalPage.CancelBtn();
            }
            finally
            {
                homePage.Navigate();

                CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrderNumber);


                customerOrderPage.DeleteCustomerOrder();
                var servicePage = homePage.GoToCustomers_ServicePage();
                var serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();





            }


        }


        //_________________________________________CREATE_DATE_ProductionCO____________________________________________________________

        /*
         * Création d'une nouvelle Date ProductionCO
        */

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PCO_AddNewDate()
        {
            HomePage homePage = LogInAsAdmin();

            ProductionCOPage ProductionPage = homePage.GoToProduction_ProductionCOPage();
            ProductionCOCreateModalPage modalcreateProductionCo = ProductionPage.ProductionCOCreatePage();
            int numberRouwsBeforeAdd = modalcreateProductionCo.CheckTotalNumber();
            modalcreateProductionCo.FillField_NewProductionCOSearch(PCO_SITE, DateUtils.Now.AddDays(-3), DateUtils.Now.AddMonths(6));
            modalcreateProductionCo.FillField_NewDate(DateTime.Now);
            int numberRouwsAfterAdd = modalcreateProductionCo.CheckTotalNumber();
            bool isGreater = numberRouwsAfterAdd > numberRouwsBeforeAdd;
            Assert.IsTrue(isGreater, "Le productionCO New Date n'a pas été créé.");
        }
        //_________________________________________FIN CREATE_DATE_ProductionCO________________________________________________________

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PCO_OrderByDate()
        {
            DateTime dateFrom = DateTime.Now.AddDays(-6);
            string orderbydate = "ProductionDate";

            //Arrange
            HomePage homePage = LogInAsAdmin();

            string dateFormat = homePage.GetDateFormatPickerValue();
            ProductionCOPage ProductionPage = homePage.GoToProduction_ProductionCOPage();
            ProductionPage.ResetFilter();
            ProductionPage.Filter(ProductionCOPage.FilterType.DateFrom, dateFrom);
            ProductionPage.Filter(ProductionCOPage.FilterType.SortBy, orderbydate);
            ProductionPage.PageSize("100");
            bool sortedByProductionDate = ProductionPage.SortedByProductionDate(dateFormat);

            //Assert
            Assert.IsTrue(sortedByProductionDate, MessageErreur.FILTRE_ERRONE, "Sort by Production Date");

        }

        //_________________________________________FIN CREATE_DATE_ProductionCO________________________________________________________

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PCO_ServiceBlockIfNoDatasheet()
        {
            string site = PCO_SITE;

            var homePage = LogInAsAdmin();

            ProductionCOPage ProductionPage = homePage.GoToProduction_ProductionCOPage();
            ProductionCOCreateModalPage modalcreateProductionCo = ProductionPage.ProductionCOCreatePage();
            modalcreateProductionCo.Filter(ProductionCOCreateModalPage.FilterType.Site, site);
            modalcreateProductionCo.Filter(ProductionCOCreateModalPage.FilterType.DateFrom, DateFrom);
            modalcreateProductionCo.Filter(ProductionCOCreateModalPage.FilterType.DateTo, DateTo);
            modalcreateProductionCo.Search();
            var isServicehidden = modalcreateProductionCo.isServicehidden();
            Assert.IsTrue(isServicehidden, "le service est associé à une Datasheet");
        }


        //_________________________________________FILTRER_BY_DATE_ProductionCO____________________________________________________________

        /*
         * filtrag par Date ProductionCO
        */
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PCO_FilterByProdDate()
        {
            Random random = new Random();
            string site = PCO_SITE;
            int minDays = 5;
            int maxDays = 20;

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            ProductionCOPage ProductionPage = homePage.GoToProduction_ProductionCOPage();

            ProductionPage.Filter(ProductionCOPage.FilterType.Sites, site);
            ProductionPage.Filter(ProductionCOPage.FilterType.DateFrom, DateTime.Now.AddDays(-10));
            List<string> Dates = ProductionPage.GetAllDates();
            List<string> distinctDates = Dates.Distinct().Take(3).ToList();

            foreach (string date in distinctDates)
            {
                int randomDays = random.Next(minDays, maxDays + 1);
                ProductionPage.Filter(ProductionCOPage.FilterType.DateFrom, DateTime.ParseExact(date, "dd/MM/yyyy", CultureInfo.InvariantCulture));
                ProductionPage.Filter(ProductionCOPage.FilterType.DateTo, DateTime.Now.AddDays(randomDays));
                var numberRows = ProductionPage.CheckTotalNumber();

                Assert.IsTrue(numberRows > 0, "Ce Plage ne contient aucune date correspond à la  date de production.");
            }
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PCO_LinktoDatasheetNewProd()
        {
            string DatasheetsTitle = "Datasheet : DataSheetProdCO - Winrest";
            string site = PCO_SITE;
            var homePage = LogInAsAdmin();
            var productionCOPage = homePage.GoToProduction_ProductionCOPage();
            var productionCOCreateModalPage = productionCOPage.ProductionCOCreatePage();
            productionCOCreateModalPage.FillField_NewProductionCOSearch(site, DateUtils.Now.AddDays(-3), DateUtils.Now.AddMonths(5));
            productionCOCreateModalPage.AddNewPRCOWithQty(DateTime.Now, true, PCO_SERVICEWITHDATASHEET);
            productionCOCreateModalPage.SetNewQtyForNewDate("10");
            productionCOCreateModalPage.Submit();
            ProductionCODetailsPage productionCODetailsPage = productionCOPage.SelectFirstRow();
            bool isRedirectedToDatasheetPage = productionCOCreateModalPage.VerifyRedirectionToDatasheetPage(DatasheetsTitle);
            Assert.IsTrue(isRedirectedToDatasheetPage, "Apres Click sur le picto il n a pas de redirection de la page qui nous amène à la fiche technique correspondante");

        }

        //_________________________________________FIN FILTRER_BY_DATE_ProductionCO________________________________________________________

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PCO_OrderByNumber()
        {
            DateTime dateFrom = DateTime.Now.AddDays(-6);
            string number = "Number";
            //Arrange
            HomePage homePage = LogInAsAdmin();

            ProductionCOPage ProductionPage = homePage.GoToProduction_ProductionCOPage();
            ProductionPage.ResetFilter();
            ProductionPage.Filter(ProductionCOPage.FilterType.DateFrom, dateFrom);         
            ProductionPage.Filter(ProductionCOPage.FilterType.SortBy, number);
            ProductionPage.PageSize("100");
            bool verifyByNumber = ProductionPage.IsSortedByNumber();

            //Assert 
            Assert.IsTrue(verifyByNumber, MessageErreur.FILTRE_ERRONE, "Les nombres ne sont pas triés par ordre croissant.");
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PCO_CreateNewFirstPopUp()
        {

            var homePage = LogInAsAdmin();

            ProductionCOPage ProductionPage = homePage.GoToProduction_ProductionCOPage();
            ProductionCOCreateModalPage modalcreateProductionCo = ProductionPage.ProductionCOCreatePage();
            modalcreateProductionCo.Filter(ProductionCOCreateModalPage.FilterType.Site, PCO_SITE);
            modalcreateProductionCo.Filter(ProductionCOCreateModalPage.FilterType.DateFrom, DateFrom);
            modalcreateProductionCo.Filter(ProductionCOCreateModalPage.FilterType.DateTo, DateTo);
            modalcreateProductionCo.Search();
            Assert.IsTrue(modalcreateProductionCo.isServicehidden(), "le service n'est pas associé à une Datasheet");
            modalcreateProductionCo.AddNewPRCOWithQty(DateTime.Now, true, PCO_SERVICEWITHDATASHEET);
            string expeditionDateStr = modalcreateProductionCo.GetDateExpedition();

            // Attempt to parse the date string and assert that parsing was successful and the date is in range
            Assert.IsTrue(modalcreateProductionCo.VerifyDateExpedition(expeditionDateStr, DateFrom, DateTo),
                 $"The expedition date {expeditionDateStr} is not valid or not between {DateFrom.ToShortDateString()} and {DateTo.ToShortDateString()}.");
        }


        //_________________________________________FILTRER_BY_SITE_ProductionCO____________________________________________________________

        /*
         * filtrag par site ProductionCO
        */
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PCO_FilterBySite()
        {

            string site = "MAD";

            //Arrange
            HomePage homePage = LogInAsAdmin();
            ProductionCOPage ProductionPage = homePage.GoToProduction_ProductionCOPage();

            //1ere filrage 
            ProductionPage.Filter(ProductionCOPage.FilterType.Sites, PCO_SITE);


            //2éme filrage 
            ProductionPage.Filter(ProductionCOPage.FilterType.Sites, site);


            var numberRows = ProductionPage.CheckTotalNumber();

            Assert.AreEqual(0, numberRows, "Il ne faut pas avoir de production pour ce site: " + site);


        }

        //_________________________________________FILTRER_BY_DATE_ProductionCO________________________________________________________

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PCO_FilterByFromDate()
        {
            //prepare
            DateTime productionDate = DateTime.Today;
            DateTime dateEqualOfProductionDate = productionDate;
            DateTime dateSuperieurOfProductionDate = DateTime.Today.AddDays(10);
            DateTime dateInfirieurOfProductionDate = DateTime.Today.AddDays(-10);
            bool isActive = true;
            string quantity = "10";
            string prodCoNumber = "";

            // arrange
            var homePage = LogInAsAdmin();
            ProductionCOPage productionCOPage = homePage.GoToProduction_ProductionCOPage();
            // act
            try
            {
                // create procductionCo
                ProductionCOCreateModalPage modalcreateProductionCo = productionCOPage.ProductionCOCreatePage();
                modalcreateProductionCo.FillField_NewProductionCOSearch(PCO_SITE, DateFrom, DateTo);
                modalcreateProductionCo.AddNewPRCOWithQty(productionDate, isActive, PCO_SERVICEWITHDATASHEET);
                modalcreateProductionCo.GetPlannedQty();
                modalcreateProductionCo.GetDiffQty();
                modalcreateProductionCo.SetNewQtyForNewDate(quantity);
                modalcreateProductionCo.Submit();
                // get the first productionCo
                ProductionCODetailsPage firstProductionCo = productionCOPage.SelectFirstRow();
                prodCoNumber = firstProductionCo.GetProductionCONumber();
                //set productionDate
                firstProductionCo.ApplyProductionDate(productionDate);
                //apply filters
                productionCOPage.Filter(ProductionCOPage.FilterType.Search, prodCoNumber);
                productionCOPage.Filter(ProductionCOPage.FilterType.DateFrom, dateEqualOfProductionDate);
                var numberOfProduct = productionCOPage.GetCountResult();
                Assert.AreEqual(numberOfProduct, 1, "le filtre qui s'applique sur DateFrom ne s'applique pas");

                productionCOPage.Filter(ProductionCOPage.FilterType.DateFrom, dateSuperieurOfProductionDate);
                numberOfProduct = productionCOPage.GetCountResult();
                Assert.AreEqual(numberOfProduct, 0, "le filtre qui s'applique sur DateFrom ne s'applique pas");

                productionCOPage.Filter(ProductionCOPage.FilterType.DateFrom, dateInfirieurOfProductionDate);
                numberOfProduct = productionCOPage.GetCountResult();
                Assert.AreEqual(numberOfProduct, 1, "le filtre qui s'applique sur DateFrom ne s'applique pas");
            }
            finally
            {
                productionCOPage.ResetFilter();
                productionCOPage.Filter(ProductionCOPage.FilterType.Search, prodCoNumber);
                productionCOPage.deleteProductionCO();
            }
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PCO_FilterByToDate()
        {
            //prepare
            DateTime productionDate = DateTime.Today;
            DateTime dateEqualOfProductionDate = productionDate;
            DateTime dateSuperieurOfProductionDate = DateTime.Today.AddDays(10);
            DateTime dateInfirieurOfProductionDate = DateTime.Today.AddDays(-10);
            bool isActive = true;
            string quantity = "10";
            string prodCoNumber = "";

            // arrange
            var homePage = LogInAsAdmin();
            ProductionCOPage productionCOPage = homePage.GoToProduction_ProductionCOPage();
            // act
            try
            {
                // create procductionCo
                ProductionCOCreateModalPage modalcreateProductionCo = productionCOPage.ProductionCOCreatePage();
                modalcreateProductionCo.FillField_NewProductionCOSearch(PCO_SITE, DateFrom, DateTo);
                modalcreateProductionCo.AddNewPRCOWithQty(productionDate, isActive, PCO_SERVICEWITHDATASHEET);
                modalcreateProductionCo.GetPlannedQty();
                modalcreateProductionCo.GetDiffQty();
                modalcreateProductionCo.SetNewQtyForNewDate(quantity);
                modalcreateProductionCo.Submit();
                // get the first productionCo
                ProductionCODetailsPage firstProductionCo = productionCOPage.SelectFirstRow();
                prodCoNumber = firstProductionCo.GetProductionCONumber();
                //set productionDate
                firstProductionCo.ApplyProductionDate(productionDate);
                //apply filters
                productionCOPage.Filter(ProductionCOPage.FilterType.Search, prodCoNumber);
                productionCOPage.Filter(ProductionCOPage.FilterType.DateTo, dateEqualOfProductionDate);
                var numberOfProduct = productionCOPage.GetCountResult();
                Assert.AreEqual(numberOfProduct, 1, "le filtre qui s'applique sur dateTo ne s'applique pas");

                productionCOPage.Filter(ProductionCOPage.FilterType.DateTo, dateSuperieurOfProductionDate);
                numberOfProduct = productionCOPage.GetCountResult();
                Assert.AreEqual(numberOfProduct, 1, "le filtre qui s'applique sur dateTo ne s'applique pas");

                productionCOPage.Filter(ProductionCOPage.FilterType.DateTo, dateInfirieurOfProductionDate);
                numberOfProduct = productionCOPage.GetCountResult();
                Assert.AreEqual(numberOfProduct, 0, "le filtre qui s'applique sur dateTo ne s'applique pas");
            }
            finally
            {
                productionCOPage.ResetFilter();
                productionCOPage.Filter(ProductionCOPage.FilterType.Search, prodCoNumber);
                productionCOPage.deleteProductionCO();
            }
        }

        //_________________________________________PR_PCO_FilterByProdTime________________________________________________________

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PCO_FilterByStartTime()
        {
            string productionTime = "03:00:00";
            string timeEqaualWithProductionTime = "03:00";
            string timeSuperieurOfProductionTime = "04:00";
            string timeInferieurOfProductionTime = "02:00";
            string quantity = "10";
            string prodCoNumber = "";
            DateTime date = DateTime.Now;
            bool isActivate = true;
            //Arrange
            var homePage = LogInAsAdmin();
            ProductionCOPage productionPage = homePage.GoToProduction_ProductionCOPage();
            productionPage.ResetFilter();
            string choice = "AM";
            try
            {
                // create productionCO
                ProductionCOCreateModalPage modalcreateProductionCo = productionPage.ProductionCOCreatePage();
                modalcreateProductionCo.FillField_NewProductionCOSearch(PCO_SITE, DateFrom, DateTo);
                modalcreateProductionCo.AddNewPRCOWithQty(date, isActivate, PCO_SERVICEWITHDATASHEET);
                modalcreateProductionCo.GetPlannedQty();
                modalcreateProductionCo.GetDiffQty();
                modalcreateProductionCo.SetNewQtyForNewDate(quantity);
                modalcreateProductionCo.Submit();
                //Get First Production CO 
                ProductionCODetailsPage firstProductionCo = productionPage.SelectFirstRow();
                prodCoNumber = firstProductionCo.GetProductionCONumber();
                // set the productionTime
                firstProductionCo.ApplyProductionTime(productionTime, choice);
                // Apply the filters
                productionPage.Filter(ProductionCOPage.FilterType.Search, prodCoNumber);
                productionPage.Filter(ProductionCOPage.FilterType.DateFrom, date);
                productionPage.Filter(ProductionCOPage.FilterType.DateTo, date);
                productionPage.WaitPageLoading();
                productionPage.Filter(ProductionCOPage.FilterType.StartTime, timeEqaualWithProductionTime, choice);
                productionPage.WaitPageLoading();
                var numberOfProduct = productionPage.GetCountResult();
                Assert.AreEqual(numberOfProduct, 1, "le filtre qui s'applique sur stratTime ne marche pas");

                productionPage.Filter(ProductionCOPage.FilterType.StartTime, timeSuperieurOfProductionTime, choice);
                productionPage.WaitPageLoading();
                numberOfProduct = productionPage.GetCountResult();
                Assert.AreEqual(numberOfProduct, 0, "le filtre qui s'applique sur stratTime ne marche pas");

                productionPage.Filter(ProductionCOPage.FilterType.StartTime, timeInferieurOfProductionTime, choice);
                productionPage.WaitPageLoading();
                numberOfProduct = productionPage.GetCountResult();
                Assert.AreEqual(numberOfProduct, 1, "le filtre qui s'applique sur stratTime ne marche pas");
            }
            finally
            {
                productionPage.ResetFilter();
                productionPage.Filter(ProductionCOPage.FilterType.Search, prodCoNumber);
                productionPage.deleteProductionCO();
            }

        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PCO_FilterByEndTime()
        {
            string productionTime = "03:00:00";
            string timeEqaualWithProductionTime = "03:00";
            string timeSuperieurOfProductionTime = "04:00";
            string timeInferieurOfProductionTime = "02:00";
            string quantity = "10";
            string prodCoNumber = "";
            DateTime date = DateTime.Now;
            DateTime dateFrom = DateTime.Now;
            DateTime dateTo = DateTime.Now;
            bool isActivate = true;
            string choice = "PM";
            //Arrange
            var homePage = LogInAsAdmin();
            ProductionCOPage productionPage = homePage.GoToProduction_ProductionCOPage();
            productionPage.ResetFilter();
            try
            {
                // create productionCO
                ProductionCOCreateModalPage modalcreateProductionCo = productionPage.ProductionCOCreatePage();
                modalcreateProductionCo.FillField_NewProductionCOSearch(PCO_SITE, DateFrom, DateTo);
                modalcreateProductionCo.AddNewPRCOWithQty(date, isActivate, PCO_SERVICEWITHDATASHEET);
                modalcreateProductionCo.GetPlannedQty();
                modalcreateProductionCo.GetDiffQty();
                modalcreateProductionCo.SetNewQtyForNewDate(quantity);
                modalcreateProductionCo.Submit();
                //Get First Production CO 
                ProductionCODetailsPage firstProductionCo = productionPage.SelectFirstRow();
                prodCoNumber = firstProductionCo.GetProductionCONumber();
                // set the productionTime
                firstProductionCo.ApplyProductionTime(productionTime, choice);
                // Apply the filters
                productionPage.Filter(ProductionCOPage.FilterType.Search, prodCoNumber);
                productionPage.Filter(ProductionCOPage.FilterType.DateFrom,dateFrom);
                productionPage.Filter(ProductionCOPage.FilterType.DateTo,dateTo);
                productionPage.WaitPageLoading();
                productionPage.Filter(ProductionCOPage.FilterType.EndTime, timeEqaualWithProductionTime, choice);
                productionPage.WaitPageLoading();
                var numberOfProduct = productionPage.GetCountResult();
                Assert.AreEqual(numberOfProduct, 1, "le filtre qui s'applique sur endTime ne marche pas");

                productionPage.Filter(ProductionCOPage.FilterType.EndTime, timeSuperieurOfProductionTime, choice);
                productionPage.WaitPageLoading();
                numberOfProduct = productionPage.GetCountResult();
                Assert.AreEqual(numberOfProduct, 1, "le filtre qui s'applique sur endTime ne marche pas");

                productionPage.Filter(ProductionCOPage.FilterType.EndTime, timeInferieurOfProductionTime, choice);
                productionPage.WaitPageLoading();
                numberOfProduct = productionPage.GetCountResult();
                Assert.AreEqual(numberOfProduct, 0, "le filtre qui s'applique sur endTime ne marche pas");
            }
            finally
            {
                productionPage.ResetFilter();
                productionPage.Filter(ProductionCOPage.FilterType.Search, prodCoNumber);
                productionPage.deleteProductionCO();
            }

        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PCO_IndexSearch()
        {

            string ServiceName = "Service Production CO With DataSheet";
            string site = "CPU";


            //Arrange
            var homePage = LogInAsAdmin();
            ProductionCOPage ProductionPage = homePage.GoToProduction_ProductionCOPage();
            ProductionPage.ResetFilter();
            var numberRouws = ProductionPage.GetCountResult();
            while (numberRouws < 2)
            {
                ProductionCOCreateModalPage modalcreateProductionCo = ProductionPage.ProductionCOCreatePage();
                modalcreateProductionCo.FillField_NewProductionCOSearch(site, DateFrom, DateTo);
                modalcreateProductionCo.checkFirstService();
                modalcreateProductionCo.createNewDate(DateTime.Now);
                modalcreateProductionCo.Submit();
                ProductionPage.ResetFilter();
                numberRouws = ProductionPage.GetCountResult();

            }
            ProductionPage.ResetFilter();
            ProductionCODetailsPage ProductionCODetailsPage = ProductionPage.SelectFirstRow();
            var prodNumber = ProductionCODetailsPage.GetProductionCONumber();
            ProductionCODetailsPage.CloseViewDetails();
            ProductionPage.Filter(ProductionCOPage.FilterType.Search, prodNumber);
            bool verifyProdNumber = ProductionPage.IsProdNumber(prodNumber);
            Assert.IsTrue(verifyProdNumber, "Le numéro de production n'a pas été trouvé dans les résultats");

            ProductionPage.ResetFilter();
            ProductionPage.Filter(ProductionCOPage.FilterType.Search, ServiceName);
            bool verifyServiceName = ProductionPage.CheckServices(ServiceName);
            Assert.IsTrue(verifyServiceName, "Le ServiceName non validé ne remonte pas le service à produire.");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PCO_NewPRCOAddQty()
        {
            string quantity = "10";
            DateTime dateFrom = DateUtils.Now.AddDays(-3);
            DateTime dateTo = DateUtils.Now.AddMonths(5);
            //Arrange
            var homePage = LogInAsAdmin();
            ProductionCOPage ProductionPage = homePage.GoToProduction_ProductionCOPage();
            ProductionCOCreateModalPage modalcreateProductionCo = ProductionPage.ProductionCOCreatePage();
            modalcreateProductionCo.FillField_NewProductionCOSearch(PCO_SITE, dateFrom, dateTo);
            modalcreateProductionCo.AddNewPRCOWithQty(DateTime.Now, true, PCO_SERVICEWITHDATASHEET);
            var plannedqtybefore = modalcreateProductionCo.GetPlannedQty();
            var diffqtybefore = modalcreateProductionCo.GetDiffQty();
            modalcreateProductionCo.SetNewQtyForNewDate(quantity);
            var plannedqtyafter = modalcreateProductionCo.GetPlannedQty();
            var diffqtyafter = modalcreateProductionCo.GetDiffQty();
            modalcreateProductionCo.ClickDuplicateNewDate();
            modalcreateProductionCo.ClickSupprimerNewDate();
            Assert.IsTrue((plannedqtybefore + 10) == plannedqtyafter, "le compteur Planned qty ne s'affect pas.");
            Assert.IsTrue((diffqtybefore + 10 == diffqtyafter), " la diff qty ne s'affect pas.");
            Assert.IsTrue(modalcreateProductionCo.VerifyDuplicateInputDate() && modalcreateProductionCo.VerifyDuplicateInputQuantity(quantity), " La ligne de la date + qty il n'a pas été duplique");
            bool verifySupprimerDuplicate = modalcreateProductionCo.VerifySupprimerDuplicate();
            Assert.IsTrue(verifySupprimerDuplicate, "La ligne date + qty il n'a pas été supprimé.");

        }
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PCO_CreateNew()
        {
            string quantity = "10";
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            ProductionCOPage ProductionPage = homePage.GoToProduction_ProductionCOPage();
            var numberRouwsBeforeAdd = ProductionPage.CheckTotalNumber();
            ProductionCOCreateModalPage modalcreateProductionCo = ProductionPage.ProductionCOCreatePage();
            modalcreateProductionCo.Filter(ProductionCOCreateModalPage.FilterType.Site, PCO_SITE);
            modalcreateProductionCo.Filter(ProductionCOCreateModalPage.FilterType.DateFrom, DateFrom);
            modalcreateProductionCo.Filter(ProductionCOCreateModalPage.FilterType.DateTo, DateTo);
            modalcreateProductionCo.Search();
            Assert.IsTrue(modalcreateProductionCo.isServicehidden(), "le service n'est pas associé à une Datasheet");
            modalcreateProductionCo.AddNewPRCOWithQty(DateTime.Now, true, PCO_SERVICEWITHDATASHEET);
            var plannedqtybefore = modalcreateProductionCo.GetPlannedQty();
            var diffqtybefore = modalcreateProductionCo.GetDiffQty();
            modalcreateProductionCo.SetNewQtyForNewDate(quantity);
            var plannedqtyafter = modalcreateProductionCo.GetPlannedQty();
            var diffqtyafter = modalcreateProductionCo.GetDiffQty();
            modalcreateProductionCo.ClickDuplicateNewDate();
            modalcreateProductionCo.ClickSupprimerNewDate();
            Assert.IsTrue((plannedqtybefore + 10) == plannedqtyafter, "le compteur Planned qty ne s'affect pas.");
            Assert.IsTrue((diffqtybefore + 10 == diffqtyafter), " la diff qty ne s'affect pas.");
            Assert.IsTrue(modalcreateProductionCo.VerifyDuplicateInputDate() && modalcreateProductionCo.VerifyDuplicateInputQuantity(quantity), " La ligne de la date + qty il n'a pas été duplique");
            modalcreateProductionCo.Submit();
            var numberRouwsAfterAdd = modalcreateProductionCo.CheckTotalNumber();
            Assert.AreEqual(numberRouwsBeforeAdd + 1, numberRouwsAfterAdd, "Le productionCO n'a pas été créé.");

        }
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PCO_CounterTitleIndex()
        {

            DateTime DateFrom = DateTime.Today.AddDays(-1);

            // Arrange
            HomePage homePage = LogInAsAdmin();

            ProductionCOPage ProductionPage = homePage.GoToProduction_ProductionCOPage();
            //Vérifier le compteur initial de la rubrique "Production CO"
            var initialCounter = ProductionPage.GetCountResult();
            var initialNumberOfRows = ProductionPage.CheckTotalNumber();
            Assert.AreEqual(initialCounter, initialNumberOfRows, "Le compteur initial de la rubrique 'Production CO' ne correspond pas au nombre de productions dans l'index.");
            ProductionPage.Filter(ProductionCOPage.FilterType.DateFrom, DateFrom);
            var initialCounter1 = ProductionPage.GetCountResult();
            var initialNumberOfRows1 = ProductionPage.CheckTotalNumber();
            Assert.AreEqual(initialCounter1, initialNumberOfRows1, "Le compteur apres le filtre de la rubrique 'Production CO' ne correspond pas au nombre de productions dans l'index.");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PCO_CreateNewPCO()
        {
            string site = PCO_SITE;
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            ProductionCOPage ProductionPage = homePage.GoToProduction_ProductionCOPage();
            var numberRouwsBeforeAdd = ProductionPage.CheckTotalNumber();
            ProductionCOCreateModalPage modalcreateProductionCo = ProductionPage.ProductionCOCreatePage();
            modalcreateProductionCo.FillField_NewProductionCOSearch(site, DateFrom, DateTo);
            modalcreateProductionCo.checkFirstService();
            modalcreateProductionCo.createNewDate(DateTime.Now);
            modalcreateProductionCo.Submit();
            var numberRouwsAfterAdd = modalcreateProductionCo.CheckTotalNumber();
            Assert.AreEqual(numberRouwsBeforeAdd + 1, numberRouwsAfterAdd, "Le système ne devrait créer qu'une seule production par clic sur 'OK'.");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PCO_Access()
        {
            // Arrange
            HomePage homePage = LogInAsAdmin();
            ProductionCOPage ProductionPage = homePage.GoToProduction_ProductionCOPage();

            bool isAccessOK = ProductionPage.AccessPage();

            //Assert
            Assert.IsTrue(isAccessOK, "Page inaccessible");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PCO_PrintIcon()
        {
            string site = PCO_SITE;
            LogInAsAdmin();

            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var ProductionPage = homePage.GoToProduction_ProductionCOPage();
            try
            {
                ProductionCOCreateModalPage modalcreateProductionCo = ProductionPage.ProductionCOCreatePage();
                modalcreateProductionCo.FillField_NewProductionCOSearch(site, DateFrom, DateTo);
                modalcreateProductionCo.checkFirstService();
                modalcreateProductionCo.createNewDate(DateTime.Now);
                modalcreateProductionCo.Submit();
                modalcreateProductionCo.ClearDownloads();
                // Vérification et interactions avec les éléments de la page
                ProductionPage.ResetFilter();
                ProductionPage.VerifyPrintIconExists();
                ProductionPage.ClickPrintIcon();
                ProductionPage.SelectLanguageAndPrint("English");
                ProductionPage.ClickTopRightPrinterIcon();

                // Ouvrir le dernier fichier PDF
                string pdfUrl = ProductionPage.OpenLatestPDF();

                // Assertion pour vérifier si le PDF est ouvert
                Assert.IsTrue(!string.IsNullOrEmpty(pdfUrl), "Le PDF n'a pas été ouvert.");
            }
            finally
            {
                ProductionPage.Close();
                ProductionPage.ResetFilter();
                ProductionPage.deleteProductionCO();
            }
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PCO_PopUpStyleAddNew()
        {
            // Prepare  
            string site = PCO_SITE;

            // Arrange
            LogInAsAdmin();
            HomePage homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            //Act
            ProductionCOPage ProductionPage = homePage.GoToProduction_ProductionCOPage();
            ProductionCOCreateModalPage modalcreateProductionCo = ProductionPage.ProductionCOCreatePage();

            modalcreateProductionCo.FillField_NewProductionCOSearch(site, DateUtils.Now.AddDays(-3), DateUtils.Now.AddMonths(6));
            modalcreateProductionCo.OpenProductionCOPopUp(DateTime.Now);

            //Affichage style popup	
            bool isokstyle = modalcreateProductionCo.IsStyleAffiche();

            // Assert
            Assert.IsTrue(isokstyle, "La pop up qui s'ouvre après le search doit avoir un style");

        }
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PCO_COValideeInvisible()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomerOrder"].ToString();
            string itemName = TestContext.Properties["NameCustomerOrder"].ToString();
            string site = TestContext.Properties["SiteCPU"].ToString();
            // Arrange
            LogInAsAdmin();
            HomePage homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            //Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
            customerOrderCreateModalPage.FillField_CreatNewCustomerOrder_WithoutFlight(site, customerName, DateTime.Now.AddDays(10));
            var customerOrderItem = customerOrderCreateModalPage.Submit();
            customerOrderItem.AddItemDetails(itemName, "20");
            customerOrderItem.ValidateCO();
            string orderName = customerOrderItem.GetOrderName();
            string orderCategory = customerOrderItem.GetOrderCategory();
            int orderQuantity = customerOrderItem.GetOrderQuantity();
            ProductionCOPage ProductionPage = customerOrderCreateModalPage.GoToProduction_ProductionCOPage();
            var productionCOCreateModalPage = ProductionPage.ProductionCOCreatePage();
            productionCOCreateModalPage.FillField_NewProductionCOSearch(site, DateUtils.Now.AddDays(9), DateUtils.Now.AddDays(12));
            bool isOrderPresent = productionCOCreateModalPage.IsCustomerOrderPresent(orderName, orderCategory, orderQuantity);
            Assert.IsFalse(isOrderPresent, "The customer order should not be visible in the Production CO list after it has been validated.");
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PCO_ServiceListCategory()
        {
            //Prepare
            string site = TestContext.Properties["SiteCPU"].ToString();
            string serviceCategory = PCO_SERVICE_CATEGORY;
            string customer = "PCO_Customer" + "-" + new Random().Next().ToString();
            string customerCode = PCO_CUSTOMER_CODE;
            string customerType = PCO_CUSTOMER_TYPE;
            string serviceName = "ServiceCategory" + "-" + new Random().Next(10, 99).ToString();
            string deleteFrom = DateUtils.Now.AddMonths(-1).ToString("dd/MM/yyyy");
            string deleteTo = DateUtils.Now.AddMonths(3).ToString("dd/MM/yyyy");
            string customerOrderNumber = null;

            //Arrange
            var homePage = LogInAsAdmin();

            //act
            ParametersSites siteParameterPage = homePage.GoToParameters_Sites();
            siteParameterPage.Filter(ParametersSites.FilterType.SearchSite, site);
            siteParameterPage.CheckIfFirstSiteIsActive();
            //category Platos CPU		
            ParametersCustomer parametersCustomer = homePage.GoToParameters_CustomerPage();
            parametersCustomer.GoToTab_Category();
            if (parametersCustomer.isCategoryExist(serviceCategory))
            {
                if (!parametersCustomer.isCategoryProduct(serviceCategory))
                {
                    parametersCustomer.EditCategory(serviceCategory, true, false);

                }
            }
            try
            {
                // Add customer
                var customersPage = homePage.GoToCustomers_CustomerPage();
                customersPage.ResetFilters();
                customersPage.Filter(FilterType.Search, customer);

                if (customersPage.CheckTotalNumber() == 0)
                {
                    var customerCreateModalPage = customersPage.CustomerCreatePage();
                    customerCreateModalPage.FillFields_CreateCustomerModalPage(customer, customerCode + new Random().Next().ToString(), customerType);
                    var customerGeneralInformationsPage = customerCreateModalPage.Create();
                    customersPage = customerGeneralInformationsPage.BackToList();
                    customersPage.ResetFilters();
                    customersPage.Filter(FilterType.Search, customer);
                }
                //Assert
                Assert.AreEqual(customer, customersPage.GetFirstCustomerName(), "Le customer n'est pas présent.");

                // Act service
                var servicesPage = homePage.GoToCustomers_ServicePage();
                servicesPage.ResetFilters();
                servicesPage.Filter(ServicePage.FilterType.Search, serviceName);

                if (servicesPage.CheckTotalNumber() == 0)
                {
                    var serviceCreateModalPage = servicesPage.ServiceCreatePage();
                    serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, new Random().Next().ToString(), GenerateName(4), serviceCategory);
                    var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                    var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                    var priceModalPage = pricePage.AddNewCustomerPrice();
                    priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.AddDays(-30), DateUtils.Now.AddMonths(2));
                    servicesPage = pricePage.BackToList();
                }
                //Add customer Order 
                var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
                var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrderWOUTflight(site, customer, new Random().Next().ToString(), DateUtils.Now.AddMonths(2));
                var customerOrderDetailPage = customerOrderCreateModalPage.Submit();
                customerOrderDetailPage.AddNewItemWithCategory(serviceName, "10", serviceCategory);
          
                var generalInfo = customerOrderDetailPage.ClickOnGeneralInformationTab();
                customerOrderNumber = generalInfo.GetOrderNumber();

                //Create Production CO 
                var productionCOPage = homePage.GoToProduction_ProductionCOPage();
                //homePage.GoToProduction_ProductionCOPageModified();
                var productionCOCreateModalPage = productionCOPage.ProductionCOCreatePage();
                productionCOCreateModalPage.FillField_NewProductionCOSearch(site, DateUtils.Now.AddMonths(-2), DateUtils.Now.AddMonths(5));
                productionCOCreateModalPage.WaitPageLoading();
                var IsServicesExist = productionCOCreateModalPage.VerifierPresenceService(serviceName);
                productionCOCreateModalPage.WaitPageLoading();
                //Assert 
                Assert.IsTrue(IsServicesExist, "seuls les services classés dans la catégorie is product doivent être listés dans cette section.");
            }

            finally

            {
                //Delete Custmer Order 
                CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
                customerOrderPage.ResetFilter(); 
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrderNumber);
                customerOrderPage.Filter(CustomerOrderPage.FilterType.DateTo, DateUtils.Now.AddMonths(2));
                customerOrderPage.WaitPageLoading(); 
                customerOrderPage.DeleteCustomerOrder();
                //Delete service
                ServicePage servicePage = homePage.GoToCustomers_ServicePage();
                var serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.From, deleteFrom);
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.To, deleteTo);
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();

            }
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PCO_GenerateOutputForm()
        {
            DateTime dateFrom = DateTime.Now.AddMonths(-12);
            //Arrange
            var homePage = LogInAsAdmin();

            ProductionCOPage productionPage = homePage.GoToProduction_ProductionCOPage();
            productionPage.Filter(ProductionCOPage.FilterType.DateFrom, dateFrom);
            productionPage.Filter(ProductionCOPage.FilterType.DateTo, DateTo);
            ProductionCOOutputFormPage productionCOOutputFormPage = productionPage.GenerateNewOutputFrom();
            string numberInitial = productionCOOutputFormPage.GetOutputFormNumber();
            productionCOOutputFormPage.SetFromPlace("CPU_Economato");
            productionCOOutputFormPage.SetToPlace("CPU_Produccion");
            productionCOOutputFormPage.Generate();
            var numberAfterGeneration = productionCOOutputFormPage.VerifyIsOutputFormGenerated(numberInitial);

            //Assert
            Assert.IsTrue(numberAfterGeneration, "Générer une Output Form sans avoir de blocage.");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_PCO_DoublonPCO_GenerationOF()
        {
            //Prepare
            DateTime dateFrom = DateTime.Now.AddMonths(-12);

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //act
            ProductionCOPage productionPage = homePage.GoToProduction_ProductionCOPage();
            productionPage.Filter(ProductionCOPage.FilterType.DateFrom, dateFrom);
            productionPage.Filter(ProductionCOPage.FilterType.DateTo, DateTo);
            ProductionCOOutputFormPage productionCOOutputFormPage = productionPage.GenerateNewOutputFrom();
            string numberInitial = productionCOOutputFormPage.GetPCONumber();
            productionCOOutputFormPage.SetFromPlace("CPU_Economato");
            productionCOOutputFormPage.SetToPlace("CPU_Produccion");
            productionCOOutputFormPage.Generate();
            OutputFormGeneralInformation outputFormGeneralInformation = productionCOOutputFormPage.ClickOnGeneralInformationTab();
            ProductionCOPage productionCOPage = outputFormGeneralInformation.ClickNumber();
            int resultCount = productionPage.CheckTotalNumber();

            //Assert 
            Assert.AreEqual(1, resultCount, "Il ne devrait y avoir qu'un seul bon de sortie avec ce numéro. Un doublon a été trouvé.");
             
        }

        
    }
    }