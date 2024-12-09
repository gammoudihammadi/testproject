using FluentAssertions.Execution;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Customer;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Delivery;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.DeliveryRound;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.PriceList;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Sites;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.User;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Item;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.OutputForm;
using Newrest.Winrest.FunctionalTests.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Policy;
using static Newrest.Winrest.FunctionalTests.PageObjects.Customers.Delivery.DeliveryMassiveDeletePage;
using static Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service.ServiceMassiveDeleteModalPage;

namespace Newrest.Winrest.FunctionalTests.Customers
{
    [TestClass]
    public class DeliveryTest : TestBase
    {
        private const int _timeout = 600000;
        private const string DELIVERY = "Deliveries";
        // UTC+2 = GMT France (+1 +2)
        readonly string deliveryNameToday = "delivery-" + DateUtils.Now.ToString("dd/MM/yyyy");
        readonly string serviceNameToday = "ServiceDeliveryTest-" + DateUtils.Now.ToString("dd/MM/yyyy");
        readonly string serviceNameDuplicateToday = "ServiceDeliveryDuplicate-" + DateUtils.Now.ToString("dd/MM/yyyy");

        //Massive Delete datas
        readonly string customerMassiveDelete1 = "CustomerMassiveDelete1";
        readonly string customeMassiveDelete1Code = "MS1";
        readonly string serviceMassiveDelete1 = "ServiceMassiveDelete1";
        readonly string deliveryMassiveDelete1 = "DeliveryMassiveDelete1";
        readonly string deliveryRoundMassiveDelete1 = "DR MD1";

        readonly string customerMassiveDelete2 = "CustomerMassiveDelete2";
        readonly string customeMassiveDelete2Code = "MS2";
        readonly string serviceMassiveDelete2 = "ServiceMassiveDelete2";
        readonly string deliveryMassiveDelete2 = "DeliveryMassiveDelete2";
        readonly string deliveryRoundMassiveDelete2 = "DR MD2";


        string deliveryToday = "";
        string deliveryToday1 = "";
        string deliveryToday2 = "";
        string serviceNameToday1 = "";

        [TestInitialize]
        public override void TestInitialize()
        {
            base.TestInitialize();

            var testMethod = TestContext.TestName;
            switch (testMethod)
            {
                case nameof(CU_DELI_SearchSortByName):
                    TestInitialize_CreateDeliveriesForSortBy();
                    break;
                case nameof(CU_DELI_SearchSortByCustomer):
                    TestInitialize_CreateDeliveriesForSortBy();
                    break; 
                case nameof(CU_DELI_SearchText):
                    TestInitialize_CreateDeliveriesForSortBy();
                    break;
                case nameof(CU_DELI_MassiveDelete):
                    TestInitialize_CreateServiceToDelivery();
                    TestInitialize_CreateDeliveryWithService();
                    break;
                default:
                    break;
            }
        }

        private void TestInitialize_CreateDelivery()
        {
            //Prepare
            Random random = new Random();
            string deliveryCustomer = "A.F.B. LANZAROTE";
            string deliverySite = TestContext.Properties["Site"].ToString();
            DateTime date = DateUtils.Now;

            deliveryToday = "Delivery-" + DateUtils.Now.ToString("dd/MM/yyyy") + "-" + random.Next().ToString();

            TestContext.Properties[string.Format("deliveryId")] = InsertDelivery(deliveryToday, date, deliverySite, deliveryCustomer);
            Assert.IsNotNull(DeliveryExist((int)TestContext.Properties[string.Format("deliveryId")]), "Le delivery n'est pas crée.");
        }

        private void TestInitialize_CreateDeliveryWithService()
        {
            //Prepare
            Random random = new Random();
            string deliveryCustomer = "A.F.B. LANZAROTE";
            string deliverySite = TestContext.Properties["Site"].ToString();
            DateTime date = DateUtils.Now;

            deliveryToday = "Delivery-" + DateUtils.Now.ToString("dd/MM/yyyy") + "-" + random.Next().ToString();

            TestContext.Properties[string.Format("deliveryId")] = InsertDelivery(deliveryToday, date, deliverySite, deliveryCustomer);
            Assert.IsNotNull(DeliveryExist((int)TestContext.Properties[string.Format("deliveryId")]), "Le delivery n'est pas crée.");

            TestContext.Properties[string.Format("deliveryWithServiceId")] = InsertDeliveryToService((int)TestContext.Properties[string.Format("deliveryId")], serviceNameToday1);

        }

        private void TestInitialize_CreateServiceToDelivery()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string deliveryCustomer = "A.F.B. LANZAROTE";
            string serviceName = "ServiceDeliveryTest-" + DateUtils.Now.ToString("dd/MM/yyyy") + "_" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceCategorie = TestContext.Properties["Prodman_Needs_ServiceCategory1"].ToString();
            serviceNameToday1 = serviceName;
            DateTime fromDate = DateUtils.Now.AddDays(-10);
            DateTime toDate = DateUtils.Now.AddDays(10);

            TestContext.Properties["ServiceId"] = InsertService(serviceName, serviceCode, serviceName, serviceCategorie);
            TestContext.Properties["ServiceCustomerId"] = InsertServiceToCustomerToSites((int)TestContext.Properties["ServiceId"], deliveryCustomer, site, fromDate, toDate);

            Assert.IsNotNull(ServiceExist((int)TestContext.Properties[string.Format("ServiceId")]), "Le service n'est pas crée.");
        }

        private void TestInitialize_CreateDeliveriesForSortBy()
        {
            //Prepare
            Random random = new Random();
            string deliveryCustomer = "A.F.B. LANZAROTE";
            string deliverySite = TestContext.Properties["Site"].ToString();
            DateTime date = DateUtils.Now;
            deliveryToday1 = "Delivery1-" + DateUtils.Now.ToString("dd/MM/yyyy") + "-" + random.Next().ToString();
            deliveryToday2 = "Delivery2-" + DateUtils.Now.ToString("dd/MM/yyyy") + "-" + random.Next().ToString();

            TestContext.Properties[string.Format("deliveryId1")] = InsertDelivery(deliveryToday1, date, deliverySite, deliveryCustomer);
            Assert.IsNotNull(DeliveryExist((int)TestContext.Properties[string.Format("deliveryId1")]), "Le delivery n'est pas crée.");

            TestContext.Properties[string.Format("deliveryId2")] = InsertDelivery(deliveryToday2, date, deliverySite, deliveryCustomer);
            Assert.IsNotNull(DeliveryExist((int)TestContext.Properties[string.Format("deliveryId2")]), "Le delivery n'est pas crée.");
        }

        [TestCleanup]
        public override void TestCleanup()
        {
            var testMethod = TestContext.TestName;
            switch (testMethod)
            {
                case nameof(CU_DELI_SearchSortByName):
                    TestCleanUp_Deliveries();
                    break;
                case nameof(CU_DELI_SearchSortByCustomer):
                    TestCleanUp_Deliveries();
                    break;
                case nameof(CU_DELI_SearchText):
                    TestCleanUp_Deliveries();
                    break;

                default:
                    break;
            }
        }

        public void TestCleanup_DeleteDelivery()
        {
            DeleteDelivery((int)TestContext.Properties[string.Format("deliveryId")]);

            Assert.IsNull(DeliveryExist((int)TestContext.Properties[string.Format("deliveryId")]), "Le delivery n'est pas supprimé.");
        }

        [Priority(0)]
        [TestMethod]
        [Timeout(_timeout)]
        public void CU_DELI_CreateServiceForDelivery()
        {
            // Prepare
            string deliveryCustomer = "A.F.B. LANZAROTE";
            string deliverySite = TestContext.Properties["Site"].ToString();

            DateTime fromDate = DateUtils.Now;
            DateTime toDate = DateUtils.Now.AddDays(10);

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Création du service
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceNameToday);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceNameToday);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(deliverySite, deliveryCustomer, fromDate, toDate);
                servicePage = pricePage.BackToList();
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceNameToday);
            }

            //Assert
            Assert.IsTrue(servicePage.GetFirstServiceName().Contains(serviceNameToday), "Le service n'a pas été créé.");
        }

        [Priority(1)]
        [TestMethod]
        [Timeout(_timeout)]
        public void CU_DELI_CreateCustomersForDelivery()
        {
            // Prepare
            Random rnd = new Random();
            string customer1 = "DeliveryCustomer - Colectividades";
            string customer2 = "DeliveryCustomer - Financiero";
            string deliveryCustomer = "A.F.B. LANZAROTE";
            string deliveryCustomer2 = TestContext.Properties["CustomerDelivery2"].ToString();

            string customerType1 = TestContext.Properties["CustomerType2"].ToString();
            string customerType2 = TestContext.Properties["CustomerType3"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Création des customers
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, deliveryCustomer);

            if (customersPage.CheckTotalNumber() == 0)
            {
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(deliveryCustomer, rnd.Next().ToString(), customerType1);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customerGeneralInformationsPage.BackToList();
            }

            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, deliveryCustomer2);

            if (customersPage.CheckTotalNumber() == 0)
            {
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(deliveryCustomer2, rnd.Next().ToString(), customerType1);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();
            }

            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, customer1);

            if (customersPage.CheckTotalNumber() == 0)
            {
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customer1, rnd.Next().ToString(), customerType1);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();
            }

            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, customer2);

            if (customersPage.CheckTotalNumber() == 0)
            {
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customer2, rnd.Next().ToString(), customerType2);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();
            }

            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, deliveryCustomer);
            Assert.AreEqual(customersPage.GetFirstCustomerName(), deliveryCustomer, $"Le customer {deliveryCustomer} de type {customerType1} n'est pas présent dans la liste des customers.");

            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, deliveryCustomer2);
            Assert.AreEqual(customersPage.GetFirstCustomerName(), deliveryCustomer2, $"Le customer {deliveryCustomer2} de type {customerType1} n'est pas présent dans la liste des customers.");

            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, customer1);
            Assert.AreEqual(customersPage.GetFirstCustomerName(), customer1, $"Le customer {customer1} de type {customerType1} n'est pas présent dans la liste des customers.");

            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, customer2);
            Assert.AreEqual(customersPage.GetFirstCustomerName(), customer2, $"Le customer {customer2} de type {customerType2} n'est pas présent dans la liste des customers.");
        }

        // Créer un nouveau point de livraison

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_DELI_CreateNewDelivery()
        {
            // Prepare
            Random rnd = new Random();
            string deliveryCustomer = "A.F.B. LANZAROTE";
            string deliverySite = TestContext.Properties["Site"].ToString();

            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();

            //arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName);
            int initDeliveryNumber = deliveryPage.CheckTotalNumber();

            if (initDeliveryNumber == 0)
            {
                var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, deliveryCustomer, deliverySite, true);
                var deliveryLoadingPage = deliveryCreateModalPage.Create();
                deliveryPage = deliveryLoadingPage.BackToList();

            }
            deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName);

            // Assert
            Assert.AreEqual(deliveryPage.CheckTotalNumber(), initDeliveryNumber + 1, "Le delivery n'a pas été créé.");
        }

        // Créer un nouveau point de livraison avec champs obligatoire
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_DELI_CreateNewDelivery_Without_Code()
        {
            // Prepare
            string deliveryCustomer = "A.F.B. LANZAROTE";
            string deliverySite = TestContext.Properties["Site"].ToString();
            string deliveryName = "";
            string error = "Code/Name is required";

            //arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
            deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, deliveryCustomer, deliverySite, true);
            deliveryCreateModalPage.Create();

            //Assert
            string errorMsg = deliveryCreateModalPage.GetErrorMessage();
            Assert.AreEqual(error, errorMsg, "Aucune erreur n'est apparue la nouvelle livraison n'est pas créée.");
        }

        // Créer un nouveau point de livraison avec champs obligatoire
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_DELI_CreateNewDelivery_Already_Exist()
        {
            // Prepare
            Random rnd = new Random();
            string deliveryCustomer = "A.F.B. LANZAROTE";
            string deliverySite = TestContext.Properties["Site"].ToString();
            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();
            string error = "Name already exists for the same site and customer";

            //arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();

            var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
            deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, deliveryCustomer, deliverySite, true);
            var deliveryLoadingPage = deliveryCreateModalPage.Create();
            deliveryPage = deliveryLoadingPage.BackToList();

            // On recréé le même delivery
            deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
            deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, deliveryCustomer, deliverySite, true);
            deliveryCreateModalPage.Create();

            //Assert
            string msgError = deliveryCreateModalPage.GetErrorMessage();
            Assert.AreEqual(error, msgError, "Aucune erreur n'est apparue alors qu'un autre delivery avec le même nom existe déjà.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void CU_DELI_General_Information_Modification()
        {
            // Prepare
            Random rnd = new Random();
            string deliveryCustomer = "A.F.B. LANZAROTE";
            string deliverySite = TestContext.Properties["Site"].ToString();
            string deliveryCustomerUpdated = TestContext.Properties["CustomerDelivery2"].ToString();
            string deliverySiteUpdated = TestContext.Properties["SiteLP"].ToString();
            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();
            string nameUpdateValue = deliveryName + " updated";
            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();

            var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
            deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, deliveryCustomer, deliverySite, true);
            var deliveryLoadingPage = deliveryCreateModalPage.Create();
            deliveryPage = deliveryLoadingPage.BackToList();

            deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName);
            deliveryLoadingPage = deliveryPage.ClickOnFirstDelivery();

            ///Update
            var deliveryGeneralInfo = deliveryLoadingPage.ClickOnGeneralInformation();
            deliveryGeneralInfo.Update_General_Information(nameUpdateValue, deliveryCustomerUpdated, deliverySiteUpdated);
            deliveryPage = deliveryGeneralInfo.BackToList();

            deliveryPage.Filter(DeliveryPage.FilterType.Search, nameUpdateValue);
            var totalDelvNumber = deliveryPage.CheckTotalNumber();
            Assert.AreEqual(1, totalDelvNumber, "Le delivery avec le nom mis à jour n'a pas été trouvé.");

            deliveryLoadingPage = deliveryPage.ClickOnFirstDelivery();

            deliveryGeneralInfo = deliveryLoadingPage.ClickOnGeneralInformation();
            bool isUpdated = deliveryGeneralInfo.IsUpdated(nameUpdateValue, deliveryCustomerUpdated, deliverySiteUpdated);

            //Assert
            Assert.IsTrue(isUpdated, "Les informations du delivery n'ont pas été mises à jour.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_DELI_General_Information_Desactived()
        {
            // Prepare
            Random rnd = new Random();
            string deliveryCustomer = "A.F.B. LANZAROTE";
            string deliverySite = TestContext.Properties["Site"].ToString();

            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();

            var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
            deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, deliveryCustomer, deliverySite, true);
            var deliveryLoadingPage = deliveryCreateModalPage.Create();
            deliveryPage = deliveryLoadingPage.BackToList();

            deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName);
            deliveryLoadingPage = deliveryPage.ClickOnFirstDelivery();

            //Update
            var deliveryGeneralInfo = deliveryLoadingPage.ClickOnGeneralInformation();
            deliveryGeneralInfo.SetActive(false);
            deliveryPage = deliveryGeneralInfo.BackToList();

            deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName);
            deliveryPage.Filter(DeliveryPage.FilterType.ShowInactive, true);

            //Assert
            Assert.AreEqual(1, deliveryPage.CheckTotalNumber(), "Le delivery n'est pas visible en tant qu'inactif.");
            var firstDeliveryName = deliveryPage.GetFirstDeliveryName();
            Assert.AreEqual(deliveryName, firstDeliveryName, "Le delivery n'existe pas dans la liste des delivery.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void CU_DELI_General_Information_DesactivedAndDelete()
        {
            // Prepare
            Random rnd = new Random();
            string deliveryCustomer = "A.F.B. LANZAROTE";
            string deliverySite = TestContext.Properties["Site"].ToString();

            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();


            // Act
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();

            var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
            deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, deliveryCustomer, deliverySite, true);
            var deliveryLoadingPage = deliveryCreateModalPage.Create();
            deliveryPage = deliveryLoadingPage.BackToList();

            deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName);
            deliveryLoadingPage = deliveryPage.ClickOnFirstDelivery();

            //Update
            var deliveryGeneralInfo = deliveryLoadingPage.ClickOnGeneralInformation();
            deliveryGeneralInfo.SetActive(false);
            deliveryGeneralInfo.DeleteDelivery();

            deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName);
            deliveryPage.Filter(DeliveryPage.FilterType.ShowAll, true);

            //Assert
            Assert.AreEqual(0, deliveryPage.CheckTotalNumber(), "Le delivery n'a pas été supprimé.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_DELI_General_Information_DesactivedAndDeleteWithService()
        {
            // Prepare
            Random rnd = new Random();
            string deliveryCustomer = "A.F.B. LANZAROTE";
            string deliverySite = TestContext.Properties["Site"].ToString();

            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();
            string qty = "10";

            string error = "You cannot deactivate because there are 1 services in this delivery.";

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Création du delivery
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();

            var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
            deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, deliveryCustomer, deliverySite, true);
            var deliveryLoadingPage = deliveryCreateModalPage.Create();

            deliveryLoadingPage.AddService(serviceNameToday);
            deliveryLoadingPage.AddQuantity(qty);
            deliveryPage = deliveryLoadingPage.BackToList();

            deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName);
            deliveryLoadingPage = deliveryPage.ClickOnFirstDelivery();

            var deliveryGeneralInfo = deliveryLoadingPage.ClickOnGeneralInformation();
            deliveryGeneralInfo.SetActive(false);

            //Assert
            Assert.IsTrue(deliveryGeneralInfo.GetDeleteErrorMessage(error), "Aucune erreur n'apparaît alors qu'on tente de supprimer un delivery avec service ou le message est incorrect.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void CU_DELI_General_Information_Modification_AlreadyDeliveryName()
        {
            // Prepare
            Random rnd = new Random();
            string deliveryCustomer = "A.F.B. LANZAROTE";
            string deliverySite = TestContext.Properties["Site"].ToString();

            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();
            string deliveryNameSecond = deliveryName + " second";

            string error = "Name already exists for the same site and customer";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();

            var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
            deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, deliveryCustomer, deliverySite, true);
            var deliveryLoadingPage = deliveryCreateModalPage.Create();
            deliveryPage = deliveryLoadingPage.BackToList();

            deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
            deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryNameSecond, deliveryCustomer, deliverySite, true);
            deliveryLoadingPage = deliveryCreateModalPage.Create();
            deliveryPage = deliveryLoadingPage.BackToList();

            deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName);
            deliveryLoadingPage = deliveryPage.ClickOnFirstDelivery();

            //Update
            var deliveryGeneralInfo = deliveryLoadingPage.ClickOnGeneralInformation();
            deliveryGeneralInfo.Update_General_Information(deliveryNameSecond);
            var isError = deliveryGeneralInfo.IsError(error);
            //Assert
            Assert.IsTrue(isError, "Aucune erreur n'apparaît alors qu'on renommer un delivery avec un nom déjà existant pour le même site et customer.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_DELI_Add_Service()
        {
            // Prepare
            Random rnd = new Random();
            string deliverySite = TestContext.Properties["Site"].ToString();
            string deliveryCustomer = "A.F.B. LANZAROTE";
            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();
            string qty = "100";

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            // Création du delivery
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();

            var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
            deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, deliveryCustomer, deliverySite, true);
            var deliveryLoadingPage = deliveryCreateModalPage.Create();

            deliveryLoadingPage.AddService(serviceNameToday);
            deliveryLoadingPage.AddQuantity(qty);
            deliveryPage = deliveryLoadingPage.BackToList();

            deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName);
            deliveryPage.UnfoldAll();

            //Assert
            Assert.IsTrue(deliveryPage.GetFirstServiceName().Contains(serviceNameToday), "Le service n'a pas été ajouté au delivery.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_DELI_Change_Quantity()
        {
            // Prepare
            Random rnd = new Random();
            string deliverySite = TestContext.Properties["Site"].ToString();
            string deliveryCustomer = "A.F.B. LANZAROTE";
            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();
            string qty = "100";
            string qtyUpdated = "500";

            DateTime fromDate = DateUtils.Now;
            DateTime toDate = DateUtils.Now.AddDays(10);

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Création du delivery
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();

            var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
            deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, deliveryCustomer, deliverySite, true);
            var deliveryLoadingPage = deliveryCreateModalPage.Create();

            //Création du service pour gérer le problème de prérequis non exécuté

            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceNameToday);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceNameToday);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(deliverySite, deliveryCustomer, fromDate, toDate);
                servicePage = pricePage.BackToList();
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceNameToday);
            }
            homePage.GoToCustomers_DeliveryPage();
            deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName);
            deliveryPage.ClickOnFirstDelivery();
            deliveryLoadingPage.AddService(serviceNameToday);

            deliveryLoadingPage.AddQuantity(qty);
            deliveryPage = deliveryLoadingPage.BackToList();

            deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName);
            deliveryPage.UnfoldAll();
            var qtyReturn = deliveryPage.GetServiceQuantityMonday();

            Assert.IsTrue(qtyReturn.Contains(qty), "La quantité associée au delivery n'est pas égale à celle du service.");

            deliveryPage.FoldAll();
            deliveryLoadingPage = deliveryPage.ClickOnFirstDelivery();
            deliveryLoadingPage.AddQuantity(qtyUpdated);

            deliveryPage = deliveryLoadingPage.BackToList();

            deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName);
            deliveryPage.UnfoldAll();
            var qtyReturnUpdated = deliveryPage.GetServiceQuantityMonday();

            //Assert           
            Assert.IsTrue(qtyReturnUpdated.Contains(qtyUpdated), "La mise à jour de la quantité n'a pas été réalisée.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_DELI_Delete_Service()
        {
            // Prepare
            Random rnd = new Random();
            string deliverySite = TestContext.Properties["Site"].ToString();
            string deliveryCustomer = "A.F.B. LANZAROTE";
            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            // Création du delivery
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();

            var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
            deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, deliveryCustomer, deliverySite, true);
            var deliveryLoadingPage = deliveryCreateModalPage.Create();

            deliveryLoadingPage.AddService(serviceNameToday);
            deliveryPage = deliveryLoadingPage.BackToList();

            deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName);
            deliveryLoadingPage = deliveryPage.ClickOnFirstDelivery();
            deliveryLoadingPage.DeleteService();

            //Assert
            Assert.IsFalse(deliveryLoadingPage.IsServiceVisible(), "Le service n'a pas été supprimé.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_DELI_Add_Packaging()
        {
            // Prepare
            Random rnd = new Random();
            string deliverySite = TestContext.Properties["Site"].ToString();
            string deliveryCustomer = "A.F.B. LANZAROTE";
            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();

            string packagingMethod = "Individual / Multi-Portion";
            string packagingName = "Multiporcion";
            string nbPortions = "2";
            DateTime fromDate = DateUtils.Now;
            DateTime toDate = DateUtils.Now.AddDays(10);

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceNameToday);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceNameToday);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(deliverySite, deliveryCustomer, fromDate, toDate);
                servicePage = pricePage.BackToList();
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceNameToday);
            }

            // Création du delivery
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();

            var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
            deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, deliveryCustomer, deliverySite, true);
            var deliveryLoadingPage = deliveryCreateModalPage.Create();

            deliveryLoadingPage.AddService(serviceNameToday);

            deliveryLoadingPage.AddPackaging();
            deliveryLoadingPage.FillField_FoodPackaging(packagingMethod, packagingName, false, "0", "0", nbPortions);

            //Assert
            Assert.AreEqual(deliveryLoadingPage.GetFirstPackagingName(), packagingName, "Le packaging n'a pas été ajouté au service du delivery.");
            Assert.AreEqual(deliveryLoadingPage.GetFirstPackagingMethod(), packagingMethod, "La méthode de packaging n'a pas été ajoutée au service du delivery.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void CU_DELI_Packaging_Update()
        {
            // Prepare
            Random rnd = new Random();
            string deliveryCustomer = "A.F.B. LANZAROTE";
            string deliverySite = TestContext.Properties["Site"].ToString();

            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();

            string packagingMethod = "Individual / Multi-Portion";

            string packagingName = "Multiporcion";
            string packagingNameUpdate = "Individual";

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Création du delivery
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();

            var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
            deliveryPage.WaitPageLoading();
            deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, deliveryCustomer, deliverySite, true);
            var deliveryLoadingPage = deliveryCreateModalPage.Create();

            deliveryLoadingPage.AddService(serviceNameToday);

            deliveryLoadingPage.AddPackaging();
            deliveryLoadingPage.FillFields_FoodPackaging(packagingName, false, "0", "0", "2");

            Assert.AreEqual(deliveryLoadingPage.GetFirstPackagingName(), packagingName, "Le packaging n'a pas été ajouté au service du delivery.");
            Assert.AreEqual(deliveryLoadingPage.GetFirstPackagingMethod(), packagingMethod, "La méthode de packaging n'a pas été ajoutée au service du delivery.");

            deliveryLoadingPage.UpdatePackagingFood(packagingNameUpdate, true);

            //Assert
            var firstPackagingName = deliveryLoadingPage.GetFirstPackagingName();
            Assert.AreEqual(firstPackagingName, packagingNameUpdate, "Le nom du packaging n'a pas été mis à jour.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void CU_DELI_Packaging_Duplicate()
        {
            // Prepare
            Random rnd = new Random();
            string deliveryCustomer = "A.F.B. LANZAROTE";
            string deliverySite = TestContext.Properties["Site"].ToString();

            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();
            string deliveryNameForDuplicate = deliveryNameToday + "-" + rnd.Next().ToString() + " For duplicate";

            string serviceNameForDuplicate = serviceNameDuplicateToday + rnd.Next().ToString();
            string serviceCode = rnd.Next().ToString();
            string serviceProduction = GenerateName(4);

            DateTime fromDate = DateUtils.Now;
            DateTime toDate = DateUtils.Now.AddDays(10);

            string packagingMethod = "Individual / Multi-Portion";
            string packagingName = "Multiporcion";

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Création du service
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();

            // Create service For Duplicate 
            var serviceCreateModalPage2 = servicePage.ServiceCreatePage();
            serviceCreateModalPage2.FillFields_CreateServiceModalPage(serviceNameForDuplicate, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage2 = serviceCreateModalPage2.Create();

            var pricePage = serviceGeneralInformationsPage2.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            priceModalPage.FillFields_CustomerPrice(deliverySite, deliveryCustomer, fromDate, toDate);

            // Create delivery
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
            deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, deliveryCustomer, deliverySite, true);
            var deliveryLoadingPage = deliveryCreateModalPage.Create();

            deliveryLoadingPage.AddService(serviceNameToday);
            deliveryPage = deliveryLoadingPage.BackToList();

            // Create delivery 2
            deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
            deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryNameForDuplicate, deliveryCustomer, deliverySite, true);
            deliveryLoadingPage = deliveryCreateModalPage.Create();

            deliveryLoadingPage.AddService(serviceNameToday);
            deliveryLoadingPage.AddPackaging();
            deliveryLoadingPage.FillFields_FoodPackaging(packagingName, false, "0", "0", "2");

            //Assert
            var firstPackagingName = deliveryLoadingPage.GetFirstPackagingName();
            Assert.AreEqual(firstPackagingName, packagingName, "Le packaging n'a pas été ajouté au service du delivery.");
            var firstPackagingMethod = deliveryLoadingPage.GetFirstPackagingMethod();
            Assert.AreEqual(firstPackagingMethod, packagingMethod, "La méthode de packaging n'a pas été ajoutée au service du delivery.");

            deliveryLoadingPage.DuplicatePackaging(deliveryName);
            deliveryPage = deliveryLoadingPage.BackToList();

            deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName);
            deliveryLoadingPage = deliveryPage.ClickOnFirstDelivery();

            deliveryLoadingPage.AddPackaging();
            var firstPackagingName1 = deliveryLoadingPage.GetFirstPackagingName();
            Assert.AreEqual(firstPackagingName1, packagingName, "Le packaging n'a pas été dupliqué pour le service du delivery.");
            var firstPackagingMethod1 = deliveryLoadingPage.GetFirstPackagingMethod();
            Assert.AreEqual(firstPackagingMethod1, packagingMethod, "La méthode de packaging n'a pas été dupliquée pour le service du delivery.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void CU_DELI_Delete_Packaging()
        {
            // Prepare
            Random rnd = new Random();
            string deliveryCustomer = "A.F.B. LANZAROTE";
            string deliverySite = TestContext.Properties["Site"].ToString();

            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();

            string packagingMethod = "Individual / Multi-Portion";
            string packagingName = "Multiporcion";
            string nbPortions = "2";
            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Création du delivery
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();

            var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
            deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, deliveryCustomer, deliverySite, true);
            var deliveryLoadingPage = deliveryCreateModalPage.Create();

            deliveryLoadingPage.AddService(serviceNameToday);

            deliveryLoadingPage.AddPackaging();
            deliveryLoadingPage.FillField_FoodPackaging(packagingMethod, packagingName, false, "0", "0", nbPortions);

            //Assert
            var firstpakingname = deliveryLoadingPage.GetFirstPackagingName();
            Assert.AreEqual(firstpakingname, packagingName, "Le packaging n'a pas été ajouté au service du delivery.");
            var firstpackingmethod = deliveryLoadingPage.GetFirstPackagingMethod();
            Assert.AreEqual(firstpackingmethod, packagingMethod, "La méthode de packaging n'a pas été ajoutée au service du delivery.");

            deliveryLoadingPage.DeletePackaging();

            //Assert
            var result = deliveryLoadingPage.IsPackagingVisible();
            Assert.IsFalse(result, "Le packaging n'a pas été supprimé.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void CU_DELI_Fold_All()
        {
            // Prepare
            Random rnd = new Random();
            string deliveryCustomer = "A.F.B. LANZAROTE";
            string deliverySite = TestContext.Properties["Site"].ToString();

            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();
            DateTime fromDate = DateUtils.Now;
            DateTime toDate = DateUtils.Now.AddDays(10);

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Création du delivery
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();

            var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
            deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, deliveryCustomer, deliverySite, true);
            var deliveryLoadingPage = deliveryCreateModalPage.Create();

            //Création du service pour gérer le problème de prérequis non exécuté

            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceNameToday);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceNameToday);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(deliverySite, deliveryCustomer, fromDate, toDate);
                servicePage = pricePage.BackToList();
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceNameToday);
            }
            homePage.GoToCustomers_DeliveryPage();
            deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName);
            deliveryPage.ClickOnFirstDelivery();
            deliveryLoadingPage.AddService(serviceNameToday);
            deliveryPage = deliveryLoadingPage.BackToList();

            deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName);
            deliveryPage.UnfoldAll();
            Assert.IsTrue(deliveryPage.IsUnfoldAll(), "Les détails du delivery ne sont pas affichés.");

            deliveryPage.FoldAll();
            Assert.IsTrue(deliveryPage.IsFoldAll(), "Les détails du delivery ne sont pas masqués.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_DELI_Unfold_All()
        {
            // Prepare
            Random rnd = new Random();
            string deliveryCustomer = "A.F.B. LANZAROTE";
            string deliverySite = TestContext.Properties["Site"].ToString();

            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();
            DateTime fromDate = DateUtils.Now;
            DateTime toDate = DateUtils.Now.AddDays(10);

            //arrange
            HomePage homePage = LogInAsAdmin();

            // Création du delivery
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();

            var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
            deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, deliveryCustomer, deliverySite, true);
            var deliveryLoadingPage = deliveryCreateModalPage.Create();
            //Création du service pour gérer le problème de prérequis non exécuté

            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceNameToday);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceNameToday);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(deliverySite, deliveryCustomer, fromDate, toDate);
                servicePage = pricePage.BackToList();
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceNameToday);
            }
            homePage.GoToCustomers_DeliveryPage();
            deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName);
            deliveryPage.ClickOnFirstDelivery();

            deliveryLoadingPage.AddService(serviceNameToday);
            deliveryPage = deliveryLoadingPage.BackToList();

            deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName);
            deliveryPage.UnfoldAll();

            bool isOkAfficheDetails = deliveryPage.IsUnfoldAll();
            Assert.IsTrue(isOkAfficheDetails, "Les détails du delivery ne sont pas affichés.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_DELI_Filter_SortBy()
        {
            // Prepare
            Random rnd = new Random();
            string deliveryCustomer = "A.F.B. LANZAROTE";
            string deliverySite = TestContext.Properties["Site"].ToString();

            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.ResetFilter();

            if (deliveryPage.CheckTotalNumber() < 50)
            {
                var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, deliveryCustomer, deliverySite, true);
                var deliveryLoadingPage = deliveryCreateModalPage.Create();
                deliveryPage = deliveryLoadingPage.BackToList();
                deliveryPage.ResetFilter();
            }

            if (!deliveryPage.isPageSizeEqualsTo100())
            {
                deliveryPage.PageSize("8");
                deliveryPage.PageSize("100");
            }

            deliveryPage.Filter(DeliveryPage.FilterType.SortBy, "Name");
            bool isSortedByName = deliveryPage.IsSortedByName();


            deliveryPage.Filter(DeliveryPage.FilterType.SortBy, "Customer");
            bool isSortedByCustomer = deliveryPage.IsSortedByCustomer();

            // Assert
            Assert.IsTrue(isSortedByName, String.Format(MessageErreur.FILTRE_ERRONE, "'Sort by name'"));
            Assert.IsTrue(isSortedByCustomer, String.Format(MessageErreur.FILTRE_ERRONE, "'Sort by customer'"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_DELI_Filter_Show_All()
        {
            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.ResetFilter();

            deliveryPage.Filter(DeliveryPage.FilterType.ShowAll, true);
            var totalNbreAll = deliveryPage.CheckTotalNumber();

            deliveryPage.Filter(DeliveryPage.FilterType.ShowInactive, true);
            var totalNbreInactiveOnly = deliveryPage.CheckTotalNumber();

            deliveryPage.Filter(DeliveryPage.FilterType.ShowActive, true);
            var totalNbreActiveOnly = deliveryPage.CheckTotalNumber();

            // Assert
            Assert.AreEqual(totalNbreAll, (totalNbreInactiveOnly + totalNbreActiveOnly), String.Format(MessageErreur.FILTRE_ERRONE, "'Show all'"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_DELI_Filter_Show_OnlyActive()
        {
            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.ResetFilter();
            deliveryPage.Filter(DeliveryPage.FilterType.ShowAll, true);
            var totalNbreAll = deliveryPage.CheckTotalNumber();

            deliveryPage.Filter(DeliveryPage.FilterType.ShowInactive, true);
            var totalNbreInactiveOnly = deliveryPage.CheckTotalNumber();

            if (!deliveryPage.isPageSizeEqualsTo100())
            {
                deliveryPage.PageSize("8");
                deliveryPage.PageSize("100");
            }
            deliveryPage.Filter(DeliveryPage.FilterType.ShowActive, true);
            var totalNbreActiveOnly = deliveryPage.CheckTotalNumber();

            //Assert
            Assert.AreEqual(totalNbreAll - totalNbreInactiveOnly, totalNbreActiveOnly, String.Format(MessageErreur.FILTRE_ERRONE, "'Show only active'"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_DELI_Filter_Show_OnlyInActive()
        {
            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.ResetFilter();
            deliveryPage.Filter(DeliveryPage.FilterType.ShowAll, true);
            var totalNbreAll = deliveryPage.CheckTotalNumber();

            deliveryPage.Filter(DeliveryPage.FilterType.ShowActive, true);
            var totalNbreActiveOnly = deliveryPage.CheckTotalNumber();

            if (!deliveryPage.isPageSizeEqualsTo100())
            {
                deliveryPage.PageSize("8");
                deliveryPage.PageSize("100");
            }
            deliveryPage.Filter(DeliveryPage.FilterType.ShowInactive, true);
            var totalNbreInactiveOnly = deliveryPage.CheckTotalNumber();

            //Assert
            Assert.AreEqual(totalNbreAll - totalNbreActiveOnly, totalNbreInactiveOnly, String.Format(MessageErreur.FILTRE_ERRONE, "'Show only inactive'"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void CU_DELI_Filter_Customer()
        {
            // Prepare
            Random rnd = new Random();
            string deliveryCustomer = "A.F.B. LANZAROTE";
            string deliverySite = TestContext.Properties["Site"].ToString();

            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.ResetFilter();
            deliveryPage.Filter(DeliveryPage.FilterType.Customers, deliveryCustomer);

            if (deliveryPage.CheckTotalNumber() < 20)
            {
                var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, deliveryCustomer, deliverySite, true);
                var deliveryLoadingPage = deliveryCreateModalPage.Create();
                deliveryPage = deliveryLoadingPage.BackToList();
                deliveryPage.ResetFilter();
                deliveryPage.Filter(DeliveryPage.FilterType.Customers, deliveryCustomer);
            }

            if (!deliveryPage.isPageSizeEqualsTo100())
            {
                deliveryPage.PageSize("8");
                deliveryPage.PageSize("100");
            }
            var verifyCustomer = deliveryPage.VerifyCustomer(deliveryCustomer);
            Assert.IsTrue(deliveryPage.VerifyCustomer(deliveryCustomer), MessageErreur.FILTRE_ERRONE, "Customers");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void CU_DELI_Filter_CustomerType()
        {
            // Prepare
            Random rnd = new Random();
            string deliverySite = TestContext.Properties["Site"].ToString();
            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();
            string deliveryName1 = deliveryNameToday + "-" + rnd.Next().ToString();
            string customer1 = "DeliveryCustomer - Colectividades";
            string customer2 = "DeliveryCustomer - Financiero";
            string customerType1 = TestContext.Properties["CustomerType2"].ToString();
            string customerType2 = TestContext.Properties["CustomerType3"].ToString();


            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.ResetFilter();

            var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
            deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, customer1, deliverySite, true);
            var deliveryLoadingPage = deliveryCreateModalPage.Create();
            deliveryPage = deliveryLoadingPage.BackToList();

            deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
            deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName1, customer2, deliverySite, true);
            deliveryLoadingPage = deliveryCreateModalPage.Create();
            deliveryPage = deliveryLoadingPage.BackToList();

            deliveryPage.ResetFilter();
            deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName);
            deliveryPage.Filter(DeliveryPage.FilterType.CustomerTypes, customerType1);
            var deliveryNameForCustomerType1 = deliveryPage.GetFirstDeliveryName();

            // Assert         
            Assert.AreEqual(1, deliveryPage.CheckTotalNumber(), String.Format(MessageErreur.FILTRE_ERRONE, "Customer type"));
            Assert.AreEqual(deliveryName, deliveryNameForCustomerType1, String.Format(MessageErreur.FILTRE_ERRONE, "Search"));

            deliveryPage.ResetFilter();
            deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName1);
            deliveryPage.Filter(DeliveryPage.FilterType.CustomerTypes, customerType2);
            var deliveryNameForCustomerType2 = deliveryPage.GetFirstDeliveryName();

            Assert.AreEqual(1, deliveryPage.CheckTotalNumber(), String.Format(MessageErreur.FILTRE_ERRONE, "Customer type"));
            Assert.AreEqual(deliveryName1, deliveryNameForCustomerType2, String.Format(MessageErreur.FILTRE_ERRONE, "Search"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void CU_DELI_Filter_Site()
        {
            // Prepare
            Random rnd = new Random();
            string deliveryCustomer = "A.F.B. LANZAROTE";
            string deliverySite = TestContext.Properties["Site"].ToString();

            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.ResetFilter();
            deliveryPage.Filter(DeliveryPage.FilterType.Sites, deliverySite);

            if (deliveryPage.CheckTotalNumber() < 20)
            {
                var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, deliveryCustomer, deliverySite, true);
                var deliveryLoadingPage = deliveryCreateModalPage.Create();
                deliveryPage = deliveryLoadingPage.BackToList();
                deliveryPage.ResetFilter();
                deliveryPage.Filter(DeliveryPage.FilterType.Sites, deliverySite);
            }

            if (!deliveryPage.isPageSizeEqualsTo100())
            {
                deliveryPage.PageSize("8");
                deliveryPage.PageSize("100");
            }
            var verifySite = deliveryPage.VerifySite(deliverySite);
            Assert.IsTrue(verifySite, MessageErreur.FILTRE_ERRONE, "Sites");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_DELI_Export_New_Version()
        {
            // Prepare
            Random rnd = new Random();
            string deliveryCustomer = "A.F.B. LANZAROTE";
            string deliverySite = TestContext.Properties["Site"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();
            bool newVersionPrint = true;

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.ResetFilter();

            deliveryPage.ClearDownloads();

            if (deliveryPage.CheckTotalNumber() == 0)
            {
                var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, deliveryCustomer, deliverySite, true);
                var deliveryLoadingPage = deliveryCreateModalPage.Create();
                deliveryPage = deliveryLoadingPage.BackToList();
            }

            // Act
            deliveryPage.ExportDelivery(newVersionPrint);

            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            var correctDownloadedFile = deliveryPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber(DELIVERY, filePath);

            //Assert
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void CU_DELI_Export_New_Version_Filter()
        {
            // Prepare
            Random rnd = new Random();
            string deliveryCustomer = "A.F.B. LANZAROTE";
            string deliverySite = TestContext.Properties["Site"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();
            bool newVersionPrint = true;

            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.ResetFilter();

            deliveryPage.ClearDownloads();

            if (deliveryPage.CheckTotalNumber() == 0)
            {
                var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, deliveryCustomer, deliverySite, true);
                var deliveryLoadingPage = deliveryCreateModalPage.Create();
                deliveryPage = deliveryLoadingPage.BackToList();
            }
            else
            {
                deliveryName = deliveryPage.GetFirstDeliveryName();
            }

            deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName);
            deliveryPage.ExportDelivery(newVersionPrint);

            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            var correctDownloadedFile = deliveryPage.GetExportExcelFile(taskFiles, deliveryName);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber(DELIVERY, filePath);
            bool result = OpenXmlExcel.ReadAllDataInColumn("Code/Name", DELIVERY, filePath, deliveryName);

            //Assert
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.IsTrue(result, MessageErreur.EXCEL_DONNEES_KO);
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void CU_DELI_Massive_Delete_Datas()
        {
            // Prepare
            string customerType1 = TestContext.Properties["Production_CustomerType1"].ToString();
            string serviceCategorie = TestContext.Properties["Production_Service1"].ToString();
            string serviceType = TestContext.Properties["Production_ServiceType"].ToString();
            string siteACE = TestContext.Properties["Production_Site1"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            //Create CustomerMassiveDelete1
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, customerMassiveDelete1);

            if (customersPage.CheckTotalNumber() == 0)
            {
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customerMassiveDelete1, customeMassiveDelete1Code, customerType1);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();

                customersPage.ResetFilters();
                customersPage.Filter(CustomerPage.FilterType.Search, customerMassiveDelete1);
            }

            Assert.IsTrue(customersPage.GetFirstCustomerName().Contains(customerMassiveDelete1), "Le customer customerMassiveDelete1 n'a pas été créé.");

            //Create CustomerMassiveDelete2
            customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, customerMassiveDelete2);

            if (customersPage.CheckTotalNumber() == 0)
            {
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customerMassiveDelete2, customeMassiveDelete2Code, customerType1);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();

                customersPage.ResetFilters();
                customersPage.Filter(CustomerPage.FilterType.Search, customerMassiveDelete2);
            }

            Assert.IsTrue(customersPage.GetFirstCustomerName().Contains(customerMassiveDelete2), "Le customer customerMassiveDelete2 n'a pas été créé.");

            var servicePage = homePage.GoToCustomers_ServicePage();
            //Create ServiceMassiveDelete1
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceMassiveDelete1);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceMassiveDelete1, null, null, serviceCategorie, null, serviceType);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                serviceGeneralInformationsPage.SetProduced(true);

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteACE, customerMassiveDelete1, DateUtils.Now.AddDays(-20), DateUtils.Now.AddMonths(2));
                servicePage = pricePage.BackToList();
            }
            else
            {
                var pricePage = servicePage.ClickOnFirstService();
                pricePage.ToggleFirstPrice();
                var serviceCreatePriceModalPage = pricePage.EditFirstPrice(siteACE, customerMassiveDelete1);
                try
                {
                    serviceCreatePriceModalPage.EditPriceDates(DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2));
                }
                catch
                {
                    serviceCreatePriceModalPage.Close();
                }

                var serviceGeneralInformationsPage = pricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                servicePage = serviceGeneralInformationsPage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceMassiveDelete1);

            Assert.IsTrue(servicePage.GetFirstServiceName().Contains(serviceMassiveDelete1), "Le service " + serviceMassiveDelete1 + " n'existe pas.");

            //Create ServiceMassiveDelete2
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceMassiveDelete2);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceMassiveDelete2, null, null, serviceCategorie, null, serviceType);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                serviceGeneralInformationsPage.SetProduced(true);

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteACE, customerMassiveDelete2, DateUtils.Now.AddDays(-20), DateUtils.Now.AddMonths(2));
                servicePage = pricePage.BackToList();
            }
            else
            {
                var pricePage = servicePage.ClickOnFirstService();
                pricePage.ToggleFirstPrice();
                var serviceCreatePriceModalPage = pricePage.EditFirstPrice(siteACE, customerMassiveDelete2);
                try
                {
                    serviceCreatePriceModalPage.EditPriceDates(DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2));
                }
                catch
                {
                    serviceCreatePriceModalPage.Close();
                }

                var serviceGeneralInformationsPage = pricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                servicePage = serviceGeneralInformationsPage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceMassiveDelete2);

            Assert.IsTrue(servicePage.GetFirstServiceName().Contains(serviceMassiveDelete2), "Le service " + serviceMassiveDelete2 + " n'existe pas.");


            //Create DeliveryMassiveDelete1
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, deliveryMassiveDelete1);

            if (deliveryPage.CheckTotalNumber() == 0)
            {
                var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryMassiveDelete1, customerMassiveDelete1, siteACE, true);
                var deliveryLoadingPage = deliveryCreateModalPage.Create();

                deliveryLoadingPage.AddService(serviceMassiveDelete1);

                deliveryPage = deliveryLoadingPage.BackToList();
            }

            //Create DeliveryMassiveDelete2
            deliveryPage.ResetFilter();
            deliveryPage.Filter(PageObjects.Customers.Delivery.DeliveryPage.FilterType.Search, deliveryMassiveDelete2);

            if (deliveryPage.CheckTotalNumber() == 0)
            {
                var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryMassiveDelete2, customerMassiveDelete2, siteACE, true);
                var deliveryLoadingPage = deliveryCreateModalPage.Create();

                deliveryLoadingPage.AddService(serviceMassiveDelete2);

                deliveryPage = deliveryLoadingPage.BackToList();
            }

            //Create DeliveryRoundMassiveDelete1
            var deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            deliveryRoundPage.ResetFilter();

            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRoundMassiveDelete1);
            if (deliveryRoundPage.CheckTotalNumber() == 0)
            {
                var deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
                var deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRoundForAllDays(deliveryRoundMassiveDelete1, siteACE, DateUtils.Now, DateUtils.Now.AddDays(+31));

                var deliveryRoundDeliveriesPage = deliveryRoundGeneralInfoPage.ClickOnDeliveryTab();
                deliveryRoundDeliveriesPage.AddDelivery(deliveryMassiveDelete1);
                deliveryRoundPage = deliveryRoundDeliveriesPage.BackToList();
                deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRoundMassiveDelete1);
            }
            else
            {
                var deliveryRoundDeliveryPage = deliveryRoundPage.SelectFirstDeliveryRound();
                var deliveryRoundGeneralInformationPage = deliveryRoundDeliveryPage.ClickOnGeneralInfoTab();
                deliveryRoundGeneralInformationPage.SetDeliveryRoundEndDate(DateUtils.Now.AddDays(+31));
                deliveryRoundGeneralInformationPage.BackToList();
            }

            //Assert
            Assert.AreEqual(deliveryRoundMassiveDelete1, deliveryRoundPage.GetFirstDeliveryRound(), "Le delivery round deliveryRoundMassiveDelete1 n'a pas été créé.");

            //Create DeliveryRoundMassiveDelete2
            deliveryRoundPage.ResetFilter();

            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRoundMassiveDelete2);
            if (deliveryRoundPage.CheckTotalNumber() == 0)
            {
                var deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
                var deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRoundForAllDays(deliveryRoundMassiveDelete2, siteACE, DateUtils.Now, DateUtils.Now.AddDays(+31));

                var deliveryRoundDeliveriesPage = deliveryRoundGeneralInfoPage.ClickOnDeliveryTab();
                deliveryRoundDeliveriesPage.AddDelivery(deliveryMassiveDelete2);
                deliveryRoundPage = deliveryRoundDeliveriesPage.BackToList();
                deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRoundMassiveDelete2);
            }
            else
            {
                var deliveryRoundDeliveryPage = deliveryRoundPage.SelectFirstDeliveryRound();
                var deliveryRoundGeneralInformationPage = deliveryRoundDeliveryPage.ClickOnGeneralInfoTab();
                deliveryRoundGeneralInformationPage.SetDeliveryRoundEndDate(DateUtils.Now.AddDays(+31));
                deliveryRoundGeneralInformationPage.BackToList();
            }

            //Assert
            Assert.AreEqual(deliveryRoundMassiveDelete2, deliveryRoundPage.GetFirstDeliveryRound(), "Le delivery round deliveryRoundMassiveDelete2 n'a pas été créé.");

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void CU_DELI_Massivedelete_customersearch()
        {
            // Prepare
            string deliveryCustomer = TestContext.Properties["CustomerDelivery"].ToString();
            string deliveryCustomerCode = TestContext.Properties["CustomerDeliveryCode"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();


            // Act
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            var deliveryMassiveDeletePage = deliveryPage.MassiveDelete();

            deliveryMassiveDeletePage.SelectCustomer(deliveryCustomer);
            deliveryMassiveDeletePage.ClickSearchButton();
            var verifyCustomerFilter = deliveryMassiveDeletePage.VerifyCustomerFilter(deliveryCustomerCode);

            // Assert
            Assert.IsTrue(verifyCustomerFilter, "Les clients ne sont pas les bons.");

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void CU_DELI_Massivedelete_organizebycustomer()
        {
            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            var deliveryMassiveDeletePage = deliveryPage.MassiveDelete();


            deliveryMassiveDeletePage.ClickSearchButton();
            deliveryMassiveDeletePage.ClickCustomerCodeButton();

            // Assert
            var verifyCustomerCodeSortById = deliveryMassiveDeletePage.VerifyCustomerCodeSortById();
            Assert.IsTrue(verifyCustomerCodeSortById, " Les clients ne sont pas ordonnés");

            deliveryMassiveDeletePage.ClickCustomerCodeButton();
            var verifyCustomer = deliveryMassiveDeletePage.VerifyCustomerCodeSortByIdDescending();
            Assert.IsTrue(verifyCustomer, " Les clients ne sont pas ordonnés");

        }
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_DELI_Massivedelete_organizebyname()
        {
            // Arrange
            HomePage homePage = LogInAsAdmin();


            // Act
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            var deliveryMassiveDeletePage = deliveryPage.MassiveDelete();

            deliveryMassiveDeletePage.ClickSearchButton();
            deliveryMassiveDeletePage.ClickDeliveryButton();

            // Assert
            bool verifyDeliverySortById = deliveryMassiveDeletePage.VerifyDeliverySortById();
            Assert.IsTrue(verifyDeliverySortById, " Les delivery ne sont pas ordonnés");


            deliveryMassiveDeletePage.ClickDeliveryButton();
            bool verifyDeliverySortByIdDescending = deliveryMassiveDeletePage.VerifyDeliverySortByIdDescending();
            Assert.IsTrue(verifyDeliverySortByIdDescending, " Les delivery ne sont pas ordonnés");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void CU_DELI_Massivedelete_organizebysite()
        {
            // Prepare

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            var deliveryMassiveDeletePage = deliveryPage.MassiveDelete();


            deliveryMassiveDeletePage.ClickSearchButton();
            deliveryMassiveDeletePage.ClickSiteButton();

            // Assert
            Assert.IsTrue(deliveryMassiveDeletePage.VerifySiteSortById(), " Les delivery ne sont pas ordonnés par sites");

            deliveryMassiveDeletePage.ClickDeliveryButton();
            Assert.IsTrue(deliveryMassiveDeletePage.VerifySiteSortByIdDescending(), " Les delivery ne sont pas ordonnés par sites");

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void CU_DELI_Massivedelete_organizebydelrounds()
        {
            // Prepare

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            var deliveryMassiveDeletePage = deliveryPage.MassiveDelete();


            deliveryMassiveDeletePage.ClickSearchButton();
            deliveryMassiveDeletePage.ClickDelRoundButton();

            // Assert
            bool verifyDelRoundSortById = deliveryMassiveDeletePage.VerifyDelRoundSortById();
            Assert.IsTrue(verifyDelRoundSortById, " Les delivery ne sont pas ordonnés par delivery round");

            deliveryMassiveDeletePage.ClickDeliveryButton();
            bool verifyDelRoundSortByIdDescending = deliveryMassiveDeletePage.VerifyDelRoundSortByIdDescending();
            Assert.IsTrue(verifyDelRoundSortByIdDescending, " Les delivery ne sont pas ordonnés delivery round");

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void CU_DELI_Massivedelete_organizebyvalidationdate()
        {
            // Prepare

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            var deliveryMassiveDeletePage = deliveryPage.MassiveDelete();


            deliveryMassiveDeletePage.ClickSearchButton();
            deliveryMassiveDeletePage.ClickValitionDateButton();

            // Assert
            Assert.IsTrue(deliveryMassiveDeletePage.VerifyValitionDateSortById(), " Les delivery ne sont pas ordonnés par date de validation");

            deliveryMassiveDeletePage.ClickDeliveryButton();
            Assert.IsTrue(deliveryMassiveDeletePage.VerifyValitionDateSortByIdDescending(), " Les delivery ne sont pas ordonnés par date de validation");

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void CU_DELI_Massivedelete_organizebyservice()
        {
            // Prepare

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            var deliveryMassiveDeletePage = deliveryPage.MassiveDelete();


            deliveryMassiveDeletePage.ClickSearchButton();
            deliveryMassiveDeletePage.ClickServiceButton();

            // Assert
            Assert.IsTrue(deliveryMassiveDeletePage.VerifyServiceSortById(), " Les delivery ne sont pas ordonnés par service");

            deliveryMassiveDeletePage.ClickServiceButton();
            Assert.IsTrue(deliveryMassiveDeletePage.VerifyServiceSortByIdDescending(), " Les delivery ne sont pas ordonnés par service");

        }
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_DELI_Massivedelete_inactivecustomersearch()

        {   // Prepare
            Random rnd = new Random();
            string customerType1 = TestContext.Properties["CustomerType2"].ToString();
            String customerInactive = "Customer Delete" + "-" + rnd.Next(1000, 9999).ToString();

            //arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            //create an inactive customer  
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, customerInactive);
            customersPage.Filter(CustomerPage.FilterType.ShowInactive, true);

            if (customersPage.CheckTotalNumber() == 0)
            {
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerInactiveModalPage(customerInactive, rnd.Next().ToString(), customerType1);

                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();

                customersPage.ResetFilters();
                customersPage.Filter(CustomerPage.FilterType.ShowInactive, true);
                customersPage.Filter(CustomerPage.FilterType.Search, customerInactive);
            }

            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            var deliveryMassiveDeletePage = deliveryPage.MassiveDelete();
            deliveryMassiveDeletePage.SelectInactiveCustomer();
            deliveryMassiveDeletePage.SelectCustomer(customerInactive);
            deliveryMassiveDeletePage.ClickOnCustomerFilter();
            bool inactiveIsOK = deliveryMassiveDeletePage.IsInactive();

            //Assert
            Assert.IsTrue(inactiveIsOK, "Le show inactive site ne fonctionne pas.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void CU_DELI_Massivedelete_sitesearch()
        {
            // Prepare
            string site = TestContext.Properties["SiteACE"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            var deliveryMassiveDeletePage = deliveryPage.MassiveDelete();
            deliveryMassiveDeletePage.SelectSites(site);
            deliveryMassiveDeletePage.ClickSearchButton();
            var VerifySearchSite = deliveryMassiveDeletePage.VerifySearchSite(site);
            // Assert
            Assert.IsTrue(VerifySearchSite, "Les search sur le site ne fonction pas.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void CU_DELI_Massivedelete_namesearch()
        {
            // Prepare
            Random rnd = new Random();
            string deliveryCustomer = "A.F.B. LANZAROTE";
            string deliverySite = TestContext.Properties["SiteMAD"].ToString();

            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();
            DateTime fromDate = DateUtils.Now;
            DateTime toDate = DateUtils.Now.AddDays(10);

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            // Création du service
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceNameToday);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceNameToday);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(deliverySite, deliveryCustomer, fromDate, toDate);
                servicePage = pricePage.BackToList();
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceNameToday);
            }

            Assert.IsTrue(servicePage.GetFirstServiceName().Contains(serviceNameToday), "Le service n'a pas été créé.");

            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
            deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, deliveryCustomer, deliverySite, true);
            var deliveryLoadingPage = deliveryCreateModalPage.Create();
            deliveryLoadingPage.AddService(serviceNameToday);
            deliveryPage = deliveryLoadingPage.BackToList();

            var deliveryMassiveDeletePage = deliveryPage.MassiveDelete();
            deliveryMassiveDeletePage.SearchDeliveryName(deliveryName);
            deliveryMassiveDeletePage.ClickSearchButton();
            var verifyDeliveryName = deliveryMassiveDeletePage.VerifyDeliveryName(deliveryName);
            // Assert
            Assert.IsTrue(verifyDeliveryName, "Les search sur le delivery name ne fonction pas.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_DELI_Massivedelete_link()
        {
            // Prepare
            Random rnd = new Random();
            string customerType1 = TestContext.Properties["CustomerType2"].ToString();
            string deliveryCustomer = "Customer Massive Delete Link";
            string deliverySite = TestContext.Properties["SiteMAD"].ToString();
            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            string noService = "No service";
            DeliveryMassiveDeletePage deliveryMassiveDeletePage;

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            //create customer  
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, deliveryCustomer);

            if (customersPage.CheckTotalNumber() == 0)
            {
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(deliveryCustomer, rnd.Next().ToString(), customerType1);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();

            }

            //create delivery 
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName);

            if (deliveryPage.CheckTotalNumber() == 0)
            {

                var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, deliveryCustomer, site, true);
                var deliveryLoadingPage = deliveryCreateModalPage.Create();
                deliveryLoadingPage.BackToList();

            }
            deliveryMassiveDeletePage = deliveryPage.MassiveDelete();
            deliveryMassiveDeletePage.SelectCustomer(deliveryCustomer);
            deliveryMassiveDeletePage.SelectSites(site);
            deliveryMassiveDeletePage.SearchDeliveryName(deliveryName);
            deliveryMassiveDeletePage.SelectStatus(noService);

            deliveryMassiveDeletePage.ClickSearchButton();
            var searchDeliveryName = deliveryMassiveDeletePage.GetFirstDeliveryName();
            var deliveryLoadingPage1 = deliveryMassiveDeletePage.ClickDeliveryDetail();
            bool isNewTabAdded = deliveryMassiveDeletePage.GetTotalTabs();

            // Assert
            Assert.IsTrue(isNewTabAdded, "Aucun Onglet n'a été ouvert");
            Assert.AreEqual(searchDeliveryName.ToLower().Trim(), deliveryLoadingPage1.GetDeliveryName().ToLower().Trim(), "La page detail delivery n'est pas visible.");
            deliveryLoadingPage1.Close();
            deliveryMassiveDeletePage.SelectAll();
            deliveryMassiveDeletePage.ClickDeleteButton();
            deliveryMassiveDeletePage.ClickConfirmDeleteButton();

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void CU_DELI_Massivedelete_selectall()
        {

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            var deliveryMassiveDeletePage = deliveryPage.MassiveDelete();
            deliveryMassiveDeletePage.SelectAllCustomer();
            deliveryMassiveDeletePage.SelectAllSites();
            deliveryMassiveDeletePage.ClickSearchButton();
            deliveryMassiveDeletePage.SelectAll();
            var selectedDelivery = deliveryMassiveDeletePage.GetSelectedDelivery();

            // Assert
            Assert.AreNotEqual(0, selectedDelivery, "Le Select All du delivery ne fonctionne pas.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void CU_DELI_Massivedelete_pagination()
        {
            //Arrange
            HomePage homePage = LogInAsAdmin();

            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            var deliveryMassiveDeletePage = deliveryPage.MassiveDelete();
            deliveryMassiveDeletePage.SelectAllCustomer();
            deliveryMassiveDeletePage.SelectAllSites();
            deliveryMassiveDeletePage.ClickSearchButton();
            deliveryMassiveDeletePage.SelectAll();
            int tot = deliveryMassiveDeletePage.CheckTotalNumberDataSheet();

            deliveryMassiveDeletePage.PageSizeMassiveDeletePopup("16");

            Assert.IsTrue(deliveryMassiveDeletePage.isPageSizeEqualsTo16(), "Paggination ne fonctionne pas.");
            Assert.AreEqual(deliveryMassiveDeletePage.GetTotalRowsForPagination(), tot >= 16 ? 16 : tot, "Paggination ne fonctionne pas.");

            deliveryMassiveDeletePage.PageSizeMassiveDeletePopup("30");

            Assert.IsTrue(deliveryMassiveDeletePage.isPageSizeEqualsTo30(), "Paggination ne fonctionne pas.");
            Assert.AreEqual(deliveryMassiveDeletePage.GetTotalRowsForPagination(), tot >= 30 ? 30 : tot, "Paggination ne fonctionne pas.");

            deliveryMassiveDeletePage.PageSizeMassiveDeletePopup("50");

            Assert.IsTrue(deliveryMassiveDeletePage.isPageSizeEqualsTo50(), "Paggination ne fonctionne pas.");
            Assert.AreEqual(deliveryMassiveDeletePage.GetTotalRowsForPagination(), tot >= 50 ? 50 : tot, "Paggination ne fonctionne pas.");

            deliveryMassiveDeletePage.PageSizeMassiveDeletePopup("100");

            Assert.IsTrue(deliveryMassiveDeletePage.isPageSizeEqualsTo100(), "Paggination ne fonctionne pas.");
            Assert.AreEqual(deliveryMassiveDeletePage.GetTotalRowsForPagination(), tot >= 100 ? 100 : tot, "Paggination ne fonctionne pas.");

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void CU_DELI_Massivedelete_status_activedelivery()
        {
            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            var deliveryMassiveDeletePage = deliveryPage.MassiveDelete();
            deliveryMassiveDeletePage.SelectStatus(MassiveDeleteStatus.ActiveDelivery.ToString());
            deliveryMassiveDeletePage.ClickSearchButton();

            // Assert
            int rowNumber = 0;
            DeliveryMassiveDeleteRowResult delResult = deliveryMassiveDeletePage.GetRowResultInfo(rowNumber);

            while (delResult != null)
            {
                var result = delResult.IsDeliveryInactive;
                Assert.IsFalse(result, $"Le delivery {delResult.DeliveryName} est inactif alors que le filtre est sur 'delivery active'.");
                rowNumber++;
                delResult = deliveryMassiveDeletePage.GetRowResultInfo(rowNumber);
            }
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_DELI_Massivedelete_status_inactivesites()
        {
            // Prepare
            Random rnd = new Random();
            string deliveryCustomer = "A.F.B. LANZAROTE";
            string deliverySite = TestContext.Properties["SiteMAD"].ToString();
            string deliveryName = deliveryNameToday + "-" + rnd.Next(1000, 9999).ToString();
            string customerType3 = TestContext.Properties["CustomerType3"].ToString();
            string inactiveSiteName = GenerateRandomString(16);
            DateTime fromDate = DateUtils.Now;
            DateTime toDate = DateUtils.Now.AddDays(10);
            string onlyInactiveSites = "Only inactive sites";
            string noService = "No service";


            //arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var parametersSites = homePage.GoToParameters_Sites();
            parametersSites.Filter(ParametersSites.FilterType.SearchSite, inactiveSiteName);
            if (!parametersSites.IsSiteExisting(inactiveSiteName))
            {
                CreateAndAffectNewSite(homePage, parametersSites, inactiveSiteName);
            }

            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, deliveryCustomer);

            if (customersPage.CheckTotalNumber() == 0)
            {
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerInactiveModalPage(deliveryCustomer, rnd.Next().ToString(), customerType3);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();
            }

            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName);

            if (deliveryPage.CheckTotalNumber() == 0)
            {
                deliveryPage = homePage.GoToCustomers_DeliveryPage();
                var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, deliveryCustomer, inactiveSiteName, true);
                var deliveryLoadingPage = deliveryCreateModalPage.Create();
            }


            deliveryPage = homePage.GoToCustomers_DeliveryPage();
            parametersSites = homePage.GoToParameters_Sites();
            parametersSites.Filter(ParametersSites.FilterType.SearchSite, inactiveSiteName);
            parametersSites.DisActivateSite();
            deliveryPage = homePage.GoToCustomers_DeliveryPage();
            DeliveryMassiveDeletePage massiveDeleteModal = deliveryPage.MassiveDelete();

            massiveDeleteModal.SearchDeliveryName(deliveryName);
            massiveDeleteModal.SelectStatus(noService);
            massiveDeleteModal.SelectStatus(onlyInactiveSites);
            massiveDeleteModal.ClickSearchButton();

            // Assert
            massiveDeleteModal.WaitLoading();
            massiveDeleteModal.SelectAll();
            int rowNumberResult = massiveDeleteModal.CheckTotalNumberDataSheet();
            massiveDeleteModal.SelectStatus(onlyInactiveSites);
            massiveDeleteModal.ClickSearchButton();
            massiveDeleteModal.WaitLoading();
            massiveDeleteModal.UnSelectAll();
            int rowNumberResult2 = massiveDeleteModal.CheckTotalNumberDataSheet();

            Assert.AreNotEqual(rowNumberResult, rowNumberResult2, "Seuls les delivery des sites inactifs doivent ressortir");

        }
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_DELI_Massivedelete_status_inactivecustomers()
        {
            // Prepare
            Random rnd = new Random();
            string customerType1 = TestContext.Properties["CustomerType2"].ToString();
            String customerInactive = "Customer Delete1" + "-" + rnd.Next(1000, 9999).ToString();
            String customerInactive2 = "Customer Delete2" + "-" + rnd.Next(1000, 9999).ToString();
            string deliveryName = deliveryNameToday + "-" + rnd.Next(1000, 9999).ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            string noService = "No service";

            //arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            //create an inactive customer  
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, customerInactive);
            customersPage.Filter(CustomerPage.FilterType.ShowInactive, true);

            if (customersPage.CheckTotalNumber() == 0)
            {
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerInactiveModalPage(customerInactive, rnd.Next().ToString(), customerType1);

                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();

            }
            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, customerInactive2);

            if (customersPage.CheckTotalNumber() == 0)
            {
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customerInactive2, rnd.Next().ToString(), customerType1);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customerGeneralInformationsPage.BackToList();

            }

            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName);

            if (deliveryPage.CheckTotalNumber() == 0)
            {

                var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, customerInactive2, site, true);
                var deliveryLoadingPage = deliveryCreateModalPage.Create();

            }

            //pdate customer inactive
            customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.Filter(CustomerPage.FilterType.Search, customerInactive2);
            CustomerGeneralInformationPage customerDetails = customersPage.SelectFirstCustomer();
            customerDetails.SetCustomerInactive();

            deliveryPage = homePage.GoToCustomers_DeliveryPage();
            DeliveryMassiveDeletePage massiveDeleteModal = deliveryPage.MassiveDelete();

            massiveDeleteModal.ShowInactiveCustomers();
            massiveDeleteModal.SearchDeliveryName(deliveryName);
            massiveDeleteModal.SelectStatus(noService);
            massiveDeleteModal.ClickSearchButton();
            // Assert
            massiveDeleteModal.WaitLoading();
            massiveDeleteModal.SelectAll();
            int rowNumberResult = massiveDeleteModal.CheckTotalNumberDataSheet();

            massiveDeleteModal.ShowInactiveCustomers();
            massiveDeleteModal.ClickSearchButton();
            massiveDeleteModal.WaitLoading();
            massiveDeleteModal.UnSelectAll();
            int rowNumberResult2 = massiveDeleteModal.CheckTotalNumberDataSheet();

            Assert.AreNotEqual(rowNumberResult, rowNumberResult2, "Seuls les delivery des customers inactifs doivent ressortir.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_DELI_Massivedelete_status_noservices()
        {
            // Prepare
            Random rnd = new Random();
            string customerType1 = TestContext.Properties["CustomerType2"].ToString();
            string customer = "Customer Delete" + "-" + rnd.Next(1000, 9999).ToString();
            string deliveryName = deliveryNameToday + "-" + rnd.Next(1000, 9999).ToString();
            string site = TestContext.Properties["SiteACE"].ToString();

            //arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            //create an inactive customer  
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, customer);

            if (customersPage.CheckTotalNumber() == 0)
            {
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customer, rnd.Next().ToString(), customerType1);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customerGeneralInformationsPage.BackToList();

            }

            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName);

            if (deliveryPage.CheckTotalNumber() == 0)
            {

                var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, customer, site, true);
                var deliveryLoadingPage = deliveryCreateModalPage.Create();
                deliveryLoadingPage.BackToList();
            }

            DeliveryMassiveDeletePage massiveDeleteModal = deliveryPage.MassiveDelete();
            massiveDeleteModal.SearchDeliveryName(deliveryName);
            massiveDeleteModal.SelectCustomer(customer);
            massiveDeleteModal.SelectSites(site);
            massiveDeleteModal.SelectStatus(MassiveDeleteStatus.NoService.ToString());
            massiveDeleteModal.ClickSearchButton();

            // Assert
            massiveDeleteModal.WaitLoading();
            massiveDeleteModal.SelectAll();
            int rowNumberResult = massiveDeleteModal.CheckTotalNumberDataSheet();
            Assert.AreEqual(rowNumberResult, 1, "Seuls les delivery sans services doivent ressortir.");
            massiveDeleteModal.ClickDeleteButton();
            massiveDeleteModal.ClickConfirmDeleteButton();
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void CU_DELI_Massivedelete_status_novalidationdate()
        {
            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            var deliveryMassiveDeletePage = deliveryPage.MassiveDelete();
            deliveryMassiveDeletePage.SelectStatus(MassiveDeleteStatus.NoValidationDate.ToString());
            deliveryMassiveDeletePage.ClickSearchButton();

            // Assert
            int rowNumber = 0;
            DeliveryMassiveDeleteRowResult delResult = deliveryMassiveDeletePage.GetRowResultInfo(rowNumber);

            while (delResult != null)
            {
                var result = delResult.DateValidation;
                Assert.IsNull(result, $"Le delivery {delResult.DeliveryName} est avec un date de validation alors que le filtre est sur 'no validation date'.");
                rowNumber++;
                delResult = deliveryMassiveDeletePage.GetRowResultInfo(rowNumber);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void CU_DELI_Massivedelete_inactivesitescheck()
        {
            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            DeliveryMassiveDeletePage massiveDeleteModal = deliveryPage.MassiveDelete();
            massiveDeleteModal.ShowInactiveSites();
            massiveDeleteModal.SelectSites("Inactive - ");
            massiveDeleteModal.ClickOnSiteFilter();
            var isInactive = massiveDeleteModal.IsInactive();
            //Assert
            Assert.IsTrue(isInactive, "Le show inactive site ne fonctionne pas.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_DELI_Massivedelete_inactivecustomercheck()
        {
            // Prepare
            Random rnd = new Random();
            string customerType1 = TestContext.Properties["CustomerType2"].ToString();
            String customerInactive = "Customer Delete1" + "-" + rnd.Next(1000, 9999).ToString();
            String customerInactive2 = "Customer Delete2" + "-" + rnd.Next(1000, 9999).ToString();
            string deliveryName = deliveryNameToday + "-" + rnd.Next(1000, 9999).ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            string noService = "No service";

            //arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            //create an inactive customer  
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, customerInactive);
            customersPage.Filter(CustomerPage.FilterType.ShowInactive, true);

            if (customersPage.CheckTotalNumber() == 0)
            {
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerInactiveModalPage(customerInactive, rnd.Next().ToString(), customerType1);

                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();

            }
            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, customerInactive2);

            if (customersPage.CheckTotalNumber() == 0)
            {
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customerInactive2, rnd.Next().ToString(), customerType1);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customerGeneralInformationsPage.BackToList();

            }

            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName);

            if (deliveryPage.CheckTotalNumber() == 0)
            {

                var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, customerInactive2, site, true);
                var deliveryLoadingPage = deliveryCreateModalPage.Create();

            }

            //pdate customer inactive
            customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.Filter(CustomerPage.FilterType.Search, customerInactive2);
            CustomerGeneralInformationPage customerDetails = customersPage.SelectFirstCustomer();
            customerDetails.SetCustomerInactive();

            deliveryPage = homePage.GoToCustomers_DeliveryPage();
            DeliveryMassiveDeletePage massiveDeleteModal = deliveryPage.MassiveDelete();

            massiveDeleteModal.ShowInactiveCustomers();
            massiveDeleteModal.SelectAllInactiveCustomers();
            massiveDeleteModal.SearchDeliveryName(deliveryName);
            massiveDeleteModal.SelectStatus(noService);

            massiveDeleteModal.ClickSearchButton();
            // Assert
            massiveDeleteModal.WaitLoading();
            massiveDeleteModal.SelectAll();
            int rowNumberResult = massiveDeleteModal.CheckTotalNumberDataSheet();

            massiveDeleteModal.ShowInactiveCustomers();
            massiveDeleteModal.ClickSearchButton();
            massiveDeleteModal.WaitLoading();
            massiveDeleteModal.UnSelectAll();
            int rowNumberResult2 = massiveDeleteModal.CheckTotalNumberDataSheet();

            Assert.AreNotEqual(rowNumberResult, rowNumberResult2, "Seuls les delivery des customer inactifs sélectionnés doivent ressortir.");

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void CU_DELI_Massivedelete_inactivesitessearch()
        {
            //Prepare
            string site = "MAD";
            string inactive_site = "Inactive - MAD";
            string inactive_site1 = "Inactive - TRE3 EVE VLC";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
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
                var deliveryPage = homePage.GoToCustomers_DeliveryPage();
                var deliveryMassiveDelete = deliveryPage.MassiveDelete();
                deliveryMassiveDelete.ShowInactiveSites();
                deliveryMassiveDelete.WaitPageLoading();
                deliveryMassiveDelete.SelectMultipleInactiveSites(inactive_site1, inactive_site);
                deliveryMassiveDelete.WaitPageLoading();
                deliveryMassiveDelete.ClickSearchButton();
                deliveryMassiveDelete.WaitPageLoading();

                deliveryMassiveDelete.SelectAll();
                deliveryMassiveDelete.CheckTotalSelectCount();
                var totalNumberBeforeDelete = deliveryMassiveDelete.CheckTotalSelectCount();
                deliveryMassiveDelete.UnSelectAll();
                deliveryMassiveDelete.WaitPageLoading();
                deliveryMassiveDelete.Cancel();
                deliveryPage.MassiveDelete();
                deliveryMassiveDelete.ShowInactiveSites();
                deliveryMassiveDelete.SelectSites(inactive_site);
                deliveryMassiveDelete.ClickSearchButton();
                deliveryMassiveDelete.SelectAll();
                var totalNumberAfterDelete = deliveryMassiveDelete.CheckTotalSelectCount();
                Assert.IsTrue(totalNumberAfterDelete <= totalNumberBeforeDelete, "Seuls les delivaries des sites inactifs sélectionnés doivent ressortir.");
            }
            finally
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
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void CU_DELI_ShowStyleService()
        {
            // Prepare
            Random rnd = new Random();
            DateTime fromDate = DateUtils.Now;
            DateTime toDate = DateUtils.Now.AddDays(10);
            DateTime toDateExpirer = DateUtils.Now;
            string deliveryCustomer = "A.F.B. LANZAROTE";
            string deliverySite = TestContext.Properties["Site"].ToString();
            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();
            string servicenameAdd = "ServiceTestForDelivery-" + DateUtils.Now.ToString("dd/MM/yyyy") + rnd.Next().ToString();
            string servicenameExpirerAdd = "ServiceTestForDeliveryExpirer-" + DateUtils.Now.ToString("dd/MM/yyyy") + rnd.Next().ToString();
            string normal_color = "rgb(51, 51, 51)";
            string normal_color_orange = "rgb(255, 153, 0)";

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Création du service
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, servicenameAdd);
            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(servicenameAdd);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(deliverySite, deliveryCustomer, fromDate, toDate);
                servicePage = pricePage.BackToList();
                servicePage.ResetFilters();
            }
            servicePage.Filter(ServicePage.FilterType.Search, servicenameExpirerAdd);
            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(servicenameExpirerAdd);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(deliverySite, deliveryCustomer, fromDate, toDateExpirer);
            }
            // Création du delivery
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
            deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, deliveryCustomer, deliverySite, true);
            var deliveryLoadingPage = deliveryCreateModalPage.Create();
            try
            {
                deliveryLoadingPage.AddService(servicenameAdd);
                var color_service = deliveryLoadingPage.GetFirstStyleService();
                var verifier_color_service = color_service.Equals(normal_color);
                // Assert
                Assert.IsTrue(verifier_color_service, "le service n est pas expirer est son couleur affiché est ne est pas noir.");
                deliveryLoadingPage.AddService(servicenameExpirerAdd);
                var color_service_expirer = deliveryLoadingPage.GetFirstStyleService();
                var verifier_color_service_expirer = color_service_expirer.Equals(normal_color_orange);
                // Assert
                Assert.IsTrue(verifier_color_service_expirer, "le service expirer mais sont couleur ne est pas oronger.");
            }
            finally
            {
                deliveryPage = homePage.GoToCustomers_DeliveryPage();
                deliveryPage.ResetFilter();
                deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName);
                if (deliveryPage.CheckTotalNumber() > 0)
                {

                    var deliveryMassiveDeletePage = deliveryPage.MassiveDelete();
                    deliveryMassiveDeletePage.SearchDeliveryName(deliveryName);
                    deliveryMassiveDeletePage.ClickSearchButton();


                    deliveryMassiveDeletePage.SelectAll();
                    deliveryMassiveDeletePage.ClickDeleteButton();
                    deliveryMassiveDeletePage.ClickConfirmDeleteButton();
                }
                servicePage = homePage.GoToCustomers_ServicePage();
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, servicenameAdd);
                if (servicePage.CheckTotalNumber() > 0)
                {
                    var serviceMassiveDelete = servicePage.ClickMassiveDelete();
                    serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, servicenameAdd);
                    serviceMassiveDelete.ClickSearchButton();
                    serviceMassiveDelete.DeleteFirstService();
                }
                servicePage.Filter(ServicePage.FilterType.Search, servicenameExpirerAdd);
                if (servicePage.CheckTotalNumber() > 0)
                {
                    var serviceMassiveDeletesecond = servicePage.ClickMassiveDelete();
                    serviceMassiveDeletesecond.Filter(ServiceMassiveDeleteFilterType.ServiceName, servicenameExpirerAdd);
                    serviceMassiveDeletesecond.ClickSearchButton();
                    serviceMassiveDeletesecond.DeleteFirstService();
                }
            }
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_DELI_AddServicePrixInactif()
        {
            //prepare
            Random rnd = new Random();
            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();
            string deliveryCustomer = "A.F.A. LANZAROTE";
            string site = TestContext.Properties["SiteACE"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string itemName = TestContext.Properties["InvoiceService"].ToString();
            string servicename = "Test" + new Random().Next();
            DeliveryLoadingPage deliveryLoadingPage = null;

            //arrange
            HomePage homePage = LogInAsAdmin();

            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName);

            if (deliveryPage.CheckTotalNumber() == 0)
            {
                var servicesPage = homePage.GoToCustomers_ServicePage();
                var ServiceCreateModalPage = servicesPage.ServiceCreatePage();
                ServiceCreateModalPage.FillFields_CreateServiceModalPage(servicename);
                var serviceGeneralInformationsPage = ServiceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                priceModalPage.FillFields_CustomerPrice(site, deliveryCustomer, DateUtils.Now.AddDays(-30), DateUtils.Now.AddDays(-30), "500");

                deliveryPage = homePage.GoToCustomers_DeliveryPage();
                var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, deliveryCustomer, site, true);
                deliveryLoadingPage = deliveryCreateModalPage.Create();
                deliveryLoadingPage.AddService(servicename);
            }
            bool AddedService = deliveryLoadingPage.VerifyAddedService();
            Assert.IsTrue(AddedService, "Service n'est pas ajouté");
            deliveryLoadingPage.BackToList();
            //delete delivery
            var deliveryMassiveDeletePage = deliveryPage.MassiveDelete();
            deliveryMassiveDeletePage.SelectSites(site);
            deliveryMassiveDeletePage.SearchDeliveryName(deliveryName);
            deliveryMassiveDeletePage.ClickSearchButton();
            deliveryMassiveDeletePage.SelectAll();
            deliveryMassiveDeletePage.ClickDeleteButton();
            deliveryMassiveDeletePage.ClickConfirmDeleteButton();
        }
        public void CreateAndAffectNewSite(HomePage homePage, ParametersSites parametersSites, string siteNameCode)
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
            var sitesModalPage = parametersSites.ClickOnNewSite();
            sitesModalPage.FillPrincipalField_CreationNewSite(siteNameCode, siteAddress, siteZipCode, siteCity);
            string id = parametersSites.CollectNewSiteID();

            // Affect it to the user
            var parametersUser = homePage.GoToParameters_User();
            parametersUser.SearchAndSelectUser(userName);
            parametersUser.ClickOnAffectedSite();
            parametersUser.GiveSiteRightsToUser(id, true, siteNameCode);
        }
        public static string GenerateRandomString(int length)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
            Random random = new Random();
            char[] stringChars = new char[length];

            for (int i = 0; i < length; i++)
            {
                stringChars[i] = chars[random.Next(chars.Length)];
            }

            return new string(stringChars);
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_DELI_SearchSortByName()
        {
            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.ResetFilter();

            //Sor By Name
            deliveryPage.Filter(DeliveryPage.FilterType.SortBy, "Name");
            bool isSortedByName = deliveryPage.IsSortedByName();

            //Assert
            Assert.IsTrue(isSortedByName, "Les données ne dont pas Sorted By Name");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_DELI_SearchSortByCustomer()
        {
            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.ResetFilter();

            //Sor By Name
            deliveryPage.Filter(DeliveryPage.FilterType.SortBy, "Customer");
            bool isSortedByCustomer = deliveryPage.IsSortedByCustomer();

            //Assert
            Assert.IsTrue(isSortedByCustomer, "Les données ne dont pas Sorted By Customer");
        } 
        
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_DELI_SearchText()
        {
            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.ResetFilter();

            //Search Text
            deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryToday1);
            bool deliveryNameExist = deliveryPage.DeliveryNameExist(deliveryToday1);

            //Assert
            Assert.IsTrue(deliveryNameExist, "Le delivery1 n'apparait pas avec Search Text");

            deliveryNameExist = deliveryPage.DeliveryNameExist(deliveryToday2);
            Assert.IsFalse(deliveryNameExist, "Le delivery2 apparait avec Search Text");
        }

        private int InsertDelivery(string deliveryNameForDR, DateTime date, string deliverySite, string deliveryCustomer)
        {
            Random random = new Random();
            int cst = random.Next(1, 1000);
            string query = @"
        DECLARE @siteId INT;
        SELECT TOP 1 @siteId = Id FROM sites WHERE Name LIKE @deliverySite;
        DECLARE @customerId INT;
        SELECT TOP 1 @customerId = Id FROM customers WHERE Name LIKE @deliveryCustomer;
        INSERT INTO FlightDeliveries 
            (Name, IsActive, SiteId, CustomerId, DeliveryTime, HoursBeforeAccess, AllowedModificationsPercent, HoursBeforeModifications, Method, CustomerPortalBlockAccessType) 
        VALUES 
            (@deliveryNameForDR, 1, @siteId, @customerId, '00:00:00.0000000', 1, 1, 1, 1, 1);
 
        SELECT SCOPE_IDENTITY();";

            return ExecuteAndGetInt(
                query,
                new KeyValuePair<string, object>("deliveryNameForDR", deliveryNameForDR),
                new KeyValuePair<string, object>("deliveryCustomer", deliveryCustomer),
                new KeyValuePair<string, object>("deliverySite", deliverySite));
        }

        private int InsertDeliveryToService(int deliveryId, string service)
        {
            Random random = new Random();
            int cst = random.Next(1, 1000);
            string query = @"
            DECLARE @serviceId INT;
            SELECT TOP 1 @serviceId = Id FROM Services WHERE Name LIKE @service;

            Insert into [FlightDeliveryToServices] (
                [FlightDeliveryId]
              ,[ServiceId]
              ,[Monday]
              ,[Tuesday]
              ,[Wednesday]
              ,[Thursday]
              ,[Friday]
              ,[Saturday]
              ,[Sunday])
	          values (
	          @deliveryId,
	          @serviceId,
	          '0',
	          '0',
	          '0',
	          '0',
	          '0',
	          '0',
	          '0'
	          );
            SELECT SCOPE_IDENTITY();";

            return ExecuteAndGetInt(
                query,
                new KeyValuePair<string, object>("deliveryId", deliveryId),
                new KeyValuePair<string, object>("service", service));
        }

        private void DeleteDelivery(int deliveryId)
        {
            string query = @"
    DELETE FROM FlightDeliveries 
            WHERE Id = @deliveryId;";

            TestContext.Properties["deliveryId"] = ExecuteAndGetInt(
                query,
                new KeyValuePair<string, object>("deliveryId", deliveryId));
        }

        public void TestCleanUp_Deliveries()
        {
            List<int> Deliveries = new List<int>();

            for (int i = 1; i <= 2; i++)
            {
                if (TestContext.Properties.Contains($"deliveryId{i}"))
                {
                    Deliveries.Add((int)TestContext.Properties[$"deliveryId{i}"]);
                }
            }
            DeleteDeliveries(Deliveries);
        }
        private void DeleteDeliveries(List<int> deliveries)
        {
            foreach (var deliveryId in deliveries)
            {
                DeleteDelivery(deliveryId);
            }
        }

        private int? DeliveryExist(int id)
        {
            string query = @"
             select Id from FlightDeliveries where Id = @id

              SELECT SCOPE_IDENTITY();
            "
            ;
            int? result = ExecuteAndGetInt(
            query,
            new KeyValuePair<string, object>("id", id)
            );
            return result == 0 ? (int?)null : result;
        }

        private int InsertService(string serviceName, string serviceCode, string serviceProduction, string categorie)
        {
            string query = @"
      declare @categoryId int;
      select top 1 @categoryId = Id from ServiceCategories where name = @categorie;

      insert into services (Name, Code, CategoryId, ProductionName, IsActive, IsProduced, IsInvoiced) 
          values (@serviceName,@serviceCode,@categoryId,@serviceProduction, 1, 0, 0);
      select scope_identity();";

            return ExecuteAndGetInt(
                query,
                new KeyValuePair<string, object>("serviceName", serviceName),
                new KeyValuePair<string, object>("serviceCode", serviceCode),
                new KeyValuePair<string, object>("serviceProduction", serviceProduction),
                new KeyValuePair<string, object>("categorie", categorie));
        }

        private int InsertServiceToCustomerToSites(int serviceId, string customer, string site, DateTime startDate, DateTime endDate)
        {
            string query = @"
      declare @customerId int;
      select top 1 @customerId = Id from Customers where name = @customer;

      declare @siteId int;
      select top 1 @siteId = Id from Sites where name = @site;

      Insert into [ServiceToCustomerToSites] (
       [ServiceId]
        ,[SiteId]
        ,[CustomerId]
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
        @startDate,
        @endDate,
        '2',
        '0',
        1,
        '1',
        '1',
        '1'
        );
      select scope_identity();";

            return ExecuteAndGetInt(
                query,
                new KeyValuePair<string, object>("serviceId", serviceId),
                new KeyValuePair<string, object>("customer", customer),
                new KeyValuePair<string, object>("site", site),
                new KeyValuePair<string, object>("startDate", startDate),
                new KeyValuePair<string, object>("endDate", endDate));
        }

        private int? ServiceExist(int id)
        {
            string query = @"
             select Id from Services where Id = @id

              SELECT SCOPE_IDENTITY();
            "
            ;
            int? result = ExecuteAndGetInt(
            query,
            new KeyValuePair<string, object>("id", id)
            );
            return result == 0 ? (int?)null : result;
        }


        [Timeout(_timeout)]
        [TestMethod]
        public void CU_DELI_MassiveDelete()
        {
            //Prepare
            Random random = new Random();
            string deliveryCustomer = "A.F.B. LANZAROTE";
            string deliverySite = TestContext.Properties["Site"].ToString();
            string messageConfirmation = "All selected deliveries were deleted.";
            
            // Arrange
            HomePage homePage = LogInAsAdmin();
            try
            {
                // Act
                var deliveryPage = homePage.GoToCustomers_DeliveryPage();
                var deliveryMassiveDeletePage = deliveryPage.MassiveDelete();

                deliveryMassiveDeletePage.SearchDeliveryName(deliveryToday);
                deliveryMassiveDeletePage.SelectSites(deliverySite);
                deliveryMassiveDeletePage.SelectCustomer(deliveryCustomer);

                deliveryMassiveDeletePage.ClickSearchButton();
                deliveryMassiveDeletePage.SelectAll();
                deliveryMassiveDeletePage.ClickDeleteButton();
                deliveryMassiveDeletePage.ClickConfirmDeleteButton();
                Assert.IsTrue(messageConfirmation.Equals(deliveryMassiveDeletePage.GetMessageConfirmationMassive()), "Le Delivery n'est pas effacée !");

                homePage.GoToCustomers_DeliveryPage();
                deliveryPage.ResetFilter();
                deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryToday);
                int numberOfDelivery1 = deliveryPage.CheckTotalNumber();
                Assert.AreEqual(0, numberOfDelivery1, "Le Delivery n'est pas effacée !");
            }
            finally
            {
                var servicePage = homePage.GoToCustomers_ServicePage();
                var serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceNameToday1);
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]

        public void CU_DELI_ResetFilter()
        {
            //Prepare
            string deliveryCustomer = "A.F.B. LANZAROTE";
            string deliverySite = TestContext.Properties["Site"].ToString();
            string customerType = TestContext.Properties["CustomerType2"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();

            //value of filters before filter
            var searchTextBeforeFilter = deliveryPage.GetSearchFilterText();
            var NameFilterBeforeFilter = deliveryPage.GetNameFilterText();
            var OnlyBoolBeforeFilter = deliveryPage.GetactiveOnlyBool();
            var customerNumberBeforeFilter = deliveryPage.GetcustomerNumber();
            var customertypeNumberBeforeFilter = deliveryPage.GetcustomertypeNumber();
            var siteSelectedBeforeFilter = deliveryPage.GetSiteSelectedNumber();

            //Filter 
            deliveryPage.Filter(DeliveryPage.FilterType.SortBy, "Customer");
            deliveryPage.Filter(DeliveryPage.FilterType.Customers, deliveryCustomer);
            deliveryPage.Filter(DeliveryPage.FilterType.CustomerTypes, customerType);
            deliveryPage.Filter(DeliveryPage.FilterType.Sites, deliverySite);
            deliveryPage.Filter(DeliveryPage.FilterType.ShowAll, true);
      

            //Reset
            deliveryPage.ResetFilter();

            //value of filters after rest filters
            var searchTextAfterFilter = deliveryPage.GetSearchFilterText();
            var NameFilterAfterFilter = deliveryPage.GetNameFilterText();
            var OnlyBoolAfterFilter = deliveryPage.GetactiveOnlyBool();
            var customerNumberAfterFilter = deliveryPage.GetcustomerNumber();
            var customertypeNumberAfterFilter = deliveryPage.GetcustomertypeNumber();
            var siteSelectedAfterFilter = deliveryPage.GetSiteSelectedNumber();

            //Assert
            Assert.AreEqual(searchTextBeforeFilter, searchTextAfterFilter, "Le filtre search n'a pas été réinitialisé.");
            Assert.AreEqual(NameFilterBeforeFilter, NameFilterAfterFilter, "Le filtre name n'a pas été réinitialisé.");
            Assert.AreEqual(OnlyBoolBeforeFilter, OnlyBoolAfterFilter, "Le filtre active n'a pas été réinitialisé.");
            Assert.AreEqual(customerNumberBeforeFilter, customerNumberAfterFilter, "Le filtre customer n'a pas été réinitialisé.");
            Assert.AreEqual(customertypeNumberBeforeFilter, customertypeNumberAfterFilter, "Le filtre customer type n'a pas été réinitialisé.");
            Assert.AreEqual(siteSelectedBeforeFilter, siteSelectedAfterFilter, "Le filtre site n'a pas été réinitialisé.");
        }
    }
}


