using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObject.Parameters.Customer;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Accounting.FreePrice;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Customer;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Reinvoice;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.Flight;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Flights;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Sites;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.CustomerOrder;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Security.Policy;
using System.Threading;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using static Newrest.Winrest.FunctionalTests.PageObjects.Customers.Customer.CustomerPage;

namespace Newrest.Winrest.FunctionalTests.Production
{
    [TestClass]
    public class CustomerOrderTest : TestBase
    {

        private const int Timeout = 600000;
        private const string CUSTOMER_ORDER = "Sheet 1";
        private string CustomerOrderNumber;



        [TestInitialize]
        public override void TestInitialize()
        {
            base.TestInitialize();

            var testMethod = TestContext.TestName;
            switch (testMethod)
            {
                case nameof(PR_CO_Filter_PreOrders_Yes):
                    TestInitialize_CreateCustomerOrder_IsPreOrder();
                    TestInitialize_CreateCustomerOrder_IsNotPreOrder();
                    break;
                case nameof(PR_CO_Filter_PreOrders_No):
                    TestInitialize_CreateCustomerOrder_IsPreOrder();
                    TestInitialize_CreateCustomerOrder_IsNotPreOrder();
                    break;

                case nameof(PR_CO_Filter_IsOnlyDelivered):
                    TestInitialize_CreateCustomerOrder_IsDelivered();
                    TestInitialize_CreateCustomerOrder_IsNotDelivered();
                    break;
                case nameof(PR_CO_Filter_IsOnlyNotDelivered):
                    TestInitialize_CreateCustomerOrder_IsDelivered();
                    TestInitialize_CreateCustomerOrder_IsNotDelivered();
                    break;
                case nameof(PR_CO_Filter_IsAllDelivered):
                    TestInitialize_CreateCustomerOrder_IsDelivered();
                    TestInitialize_CreateCustomerOrder_IsNotDelivered();
                    break;
                case nameof(PR_CO_Filter_IsAllVerified):
                    TestInitialize_CreateCustomerOrder_IsNotValidated_NotVerified();
                    TestInitialize_CreateCustomerOrder_IsNotValidated_Verified();
                    break;
                case nameof(PR_CO_Filter_IsOnlyVerified):
                    TestInitialize_CreateCustomerOrder_IsNotValidated_NotVerified();
                    TestInitialize_CreateCustomerOrder_IsNotValidated_Verified();
                    break;
                case nameof(PR_CO_Filter_IsOnlyNotVerified):
                    TestInitialize_CreateCustomerOrder_IsNotValidated_NotVerified();
                    TestInitialize_CreateCustomerOrder_IsNotValidated_Verified();
                    break;
                case nameof(PR_CO_ChangeSellingUnit):
                    TestInitialize_CreateCustomerOrder_Inflight();
                    break;
                case nameof(PR_CO_DeleteLinewithZeroQty):
                    TestInitialize_CreateCustomerOrder_IsNotValidated_NotVerified();
                    break;
                case nameof(PR_CO_GeneralInformationAddDelivComment):
                    TestInitialize_CreateCustomerOrder_IsNotValidated_NotVerified();
                    break;           
            }
        }

        [TestCleanup]
        public override void TestCleanup()
        {
            var testMethod = TestContext.TestName;
            switch (testMethod)
            {
                case nameof(PR_CO_Filter_PreOrders_Yes):
                    TestCleanUp_CustomerOrders();
                    break;
                case nameof(PR_CO_Filter_PreOrders_No):
                    TestCleanUp_CustomerOrders();
                    break;
                case nameof(PR_CO_Filter_IsOnlyDelivered):
                    TestCleanUp_CustomerOrders();
                    break;
                case nameof(PR_CO_Filter_IsOnlyNotDelivered):
                    TestCleanUp_CustomerOrders();
                    break;
                case nameof(PR_CO_Filter_IsAllDelivered):
                    TestCleanUp_CustomerOrders();
                    break;
                case nameof(PR_CO_Filter_IsAllVerified):
                    TestCleanUp_CustomerOrders();
                    break;
                case nameof(PR_CO_Filter_IsOnlyVerified):
                    TestCleanUp_CustomerOrders();
                    break;
                case nameof(PR_CO_Filter_IsOnlyNotVerified):
                    TestCleanUp_CustomerOrders();
                    break;
                case nameof(PR_CO_DeleteLinewithZeroQty):
                    TestCleanUp_CustomerOrders();
                    break;
                case nameof(PR_CO_GeneralInformationAddDelivComment):
                    TestCleanUp_CustomerOrders();
                    break;

            }
        }

        public void TestCleanUp_CustomerOrders()
        {
            List<int> CustomerOrders = new List<int>();

            for (int i = 1; i <= 12; i++)
            {
                if (TestContext.Properties.Contains($"CustomerOrderId{i}"))
                {
                    CustomerOrders.Add((int)TestContext.Properties[$"CustomerOrderId{i}"]);
                }
            }
            DeleteCustomerOrders(CustomerOrders);
        }

        private void DeleteCustomerOrders(List<int> customerOrders)
        {
            foreach (var customerOrderId in customerOrders)
            {
                DeleteCustomerOrder(customerOrderId);
            }
        }

        /// <summary>
        /// /////////////////////////////////////////////////Tests initialize Methods////////////////////////////////////////////
        public void TestInitialize_CreateCustomerOrder_IsPreOrder()
        {
            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();

            TestContext.Properties[string.Format("CustomerOrderId1")] = InsertCustomerOrder(site, customer, DateTime.Today, 0, 1, 0, 0, 1);
            Assert.IsNotNull(CustomerOrderExist((int)TestContext.Properties[string.Format("CustomerOrderId1")]), "Le customer ordre n'est pas crée.");
        }

        public void TestInitialize_CreateCustomerOrder_IsNotPreOrder()
        {
            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();

            TestContext.Properties[string.Format("CustomerOrderId2")] = InsertCustomerOrder(site, customer, DateTime.Today, 0, 0, 0, 0, 0);
            Assert.IsNotNull(CustomerOrderExist((int)TestContext.Properties[string.Format("CustomerOrderId2")]), "Le customer ordre n'est pas crée.");
        }

        public void TestInitialize_CreateCustomerOrder_Validated()
        {
            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();

            TestContext.Properties[string.Format("CustomerOrderId3")] = InsertCustomerOrder(site, customer, DateTime.Today, 0, 1);
            Assert.IsNotNull(CustomerOrderExist((int)TestContext.Properties[string.Format("CustomerOrderId3")]), "Le customer ordre n'est pas crée.");
        }

        public void TestInitialize_CreateCustomerOrder_NotValidated()
        {
            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();

            TestContext.Properties[string.Format("CustomerOrderId4")] = InsertCustomerOrder(site, customer, DateTime.Today, 0, 0);
            Assert.IsNotNull(CustomerOrderExist((int)TestContext.Properties[string.Format("CustomerOrderId4")]), "Le customer ordre n'est pas crée.");
        }

        public void TestInitialize_CreateCustomerOrder_IsDelivered()
        {
            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();

            TestContext.Properties[string.Format("CustomerOrderId5")] = InsertCustomerOrder(site, customer, DateTime.Today, 1);
            Assert.IsNotNull(CustomerOrderExist((int)TestContext.Properties[string.Format("CustomerOrderId5")]), "Le customer ordre n'est pas crée.");
        }

        public void TestInitialize_CreateCustomerOrder_IsNotDelivered()
        {
            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();

            TestContext.Properties[string.Format("CustomerOrderId6")] = InsertCustomerOrder(site, customer, DateTime.Today, 0);
            Assert.IsNotNull(CustomerOrderExist((int)TestContext.Properties[string.Format("CustomerOrderId6")]), "Le customer ordre n'est pas crée.");
        }

        public void TestInitialize_CreateCustomerOrder_IsValidated_Verified()
        {
            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();

            TestContext.Properties[string.Format("CustomerOrderId7")] = InsertCustomerOrder(site, customer, DateTime.Today, 0, 1, 1);
            Assert.IsNotNull(CustomerOrderExist((int)TestContext.Properties[string.Format("CustomerOrderId7")]), "Le customer ordre n'est pas crée.");
        }

        public void TestInitialize_CreateCustomerOrder_IsNotValidated_Verified()
        {
            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();

            TestContext.Properties[string.Format("CustomerOrderId8")] = InsertCustomerOrder(site, customer, DateTime.Today, 0, 0, 1);
            Assert.IsNotNull(CustomerOrderExist((int)TestContext.Properties[string.Format("CustomerOrderId8")]), "Le customer ordre n'est pas crée.");
        }

        public void TestInitialize_CreateCustomerOrder_IsNotValidated_NotVerified()
        {
            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();

            TestContext.Properties[string.Format("CustomerOrderId9")] = InsertCustomerOrder(site, customer, DateTime.Today, 0, 0, 0);
            Assert.IsNotNull(CustomerOrderExist((int)TestContext.Properties[string.Format("CustomerOrderId9")]), "Le customer ordre n'est pas crée.");

        }

        public void TestInitialize_CreateCustomerOrder_IsValidated_NotVerified()
        {
            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();

            TestContext.Properties[string.Format("CustomerOrderId10")] = InsertCustomerOrder(site, customer, DateTime.Today, 0, 1, 0);
            Assert.IsNotNull(CustomerOrderExist((int)TestContext.Properties[string.Format("CustomerOrderId10")]), "Le customer ordre n'est pas crée.");

        }

        public void TestInitialize_CreateCustomerOrder_Inflight()
        {
            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();

            TestContext.Properties[string.Format("CustomerOrderId11")] = InsertInflightCustomerOrder(site, customer, DateTime.Today, null, aircraft);
            Assert.IsNotNull(CustomerOrderExist((int)TestContext.Properties[string.Format("CustomerOrderId11")]), "Le customer ordre n'est pas crée.");
        }
        /// </summary>
        /// <summary>
        /// /////////////////////////////////////////////////Tests Methods////////////////////////////////////////////
        /// </summary>
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_Filter_PreOrders_Yes()
        {
            // Prepare
            string customerOrderPreOrderId = TestContext.Properties[string.Format("CustomerOrderId1")].ToString();
            string customerOrderNotPreOrderId = TestContext.Properties[string.Format("CustomerOrderId2")].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.PreOrderYes, true);
            customerOrderPage.PageUpCO();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrderPreOrderId);
            string firstCustomerOrder = customerOrderPage.GetCustomerOrderNumber().Split(' ')[1];

            //Assert
            Assert.AreEqual(firstCustomerOrder, customerOrderPreOrderId, MessageErreur.FILTRE_ERRONE, "IsPreOrder");

            customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrderNotPreOrderId);
            int customerOrdersNbr = customerOrderPage.CheckTotalNumber();

            //Assert
            Assert.AreEqual(customerOrdersNbr, 0, MessageErreur.FILTRE_ERRONE, "IsPreOrder");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_Filter_PreOrders_No()
        {
            // Prepare
            string customerOrderPreOrderId = TestContext.Properties[string.Format("CustomerOrderId1")].ToString();
            string customerOrderNotPreOrderId = TestContext.Properties[string.Format("CustomerOrderId2")].ToString();
            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.PreOrderNo, true);
            customerOrderPage.PageUpCO();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrderNotPreOrderId);
            int preOrdersNbr = customerOrderPage.CheckTotalNumber();
            string firstCustomerOrder = customerOrderPage.GetCustomerOrderNumber().Split(' ')[1];
            //Assert
            Assert.AreEqual(firstCustomerOrder, customerOrderNotPreOrderId, MessageErreur.FILTRE_ERRONE, "IsNotPreOrder");


            customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrderPreOrderId);
            int customerOrdersNbr = customerOrderPage.CheckTotalNumber();

            //Assert
            Assert.AreEqual(customerOrdersNbr, 0, MessageErreur.FILTRE_ERRONE, "IsNotPreOrder");
        }
        [Priority(0)]
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_VerifyCustomerAndServiceForCustomerOrder()
        {
            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string customerCode = TestContext.Properties["InvoiceCustomerCode"].ToString();
            string customerType = TestContext.Properties["CustomerType"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string itemName = TestContext.Properties["InvoiceService"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            ClearCache();

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

            // Act service
            var servicesPage = homePage.GoToCustomers_ServicePage();
            servicesPage.ResetFilters();
            servicesPage.Filter(ServicePage.FilterType.Search, itemName);

            if (servicesPage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicesPage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(itemName, new Random().Next().ToString(), GenerateName(4));
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
            servicesPage.Filter(ServicePage.FilterType.Search, itemName);

            Assert.IsTrue(servicesPage.GetFirstServiceName().Contains(itemName), "Le service n'a pas été créé.");
        }

        [TestMethod]
        [Priority(1)]
        [Timeout(Timeout)]
        public void PR_CO_CreateCustomerWithAirportTax()
        {
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string customerIcao = TestContext.Properties["InvoiceCustomerCode"].ToString();
            string customerType = TestContext.Properties["InvoiceCustomerType"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.Filter(CustomerPage.FilterType.Search, customer);

            if (customersPage.CheckTotalNumber() == 0)
            {
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customer, customerIcao, customerType);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();

                customerGeneralInformationsPage.ActivateAirportTax();

                customersPage = customerGeneralInformationsPage.BackToList();
                customersPage.ResetFilters();
                customersPage.Filter(CustomerPage.FilterType.Search, customer);
            }
            else
            {
                var customerGeneralInformationsPage = customersPage.SelectFirstCustomer();
                customerGeneralInformationsPage.ActivateAirportTax();

                customersPage = customerGeneralInformationsPage.BackToList();
                customersPage.ResetFilters();
                customersPage.Filter(CustomerPage.FilterType.Search, customer);

            }

            var customerName = customersPage.GetFirstCustomerName();

            Assert.AreEqual(customer, customerName, "Le customer n'a pas été créé.");
        }

        [TestMethod]
        [Priority(2)]
        [Timeout(Timeout)]
        public void PR_CO_CreateServiceWithAirportTax()
        {
            string serviceName = TestContext.Properties["InvoiceService"].ToString();
            string serviceCategory = TestContext.Properties["InvoiceServiceCategory"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);

            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);

            if (servicePage.CheckTotalNumber() == 0)
            {
                // Create
                var ServiceCreateModalPage = servicePage.ServiceCreatePage();
                ServiceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction, serviceCategory);
                var serviceGeneralInformationsPage = ServiceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.AddMonths(-1), DateUtils.Now.AddMonths(+3));
                servicePage = pricePage.BackToList();
            }
            else
            {
                var servicePricePage = servicePage.ClickOnFirstService();
                var serviceGeneralInformationPage = servicePricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationPage.SetCategory(serviceCategory);

                var pricePage = serviceGeneralInformationPage.GoToPricePage();
                pricePage.SearchPriceForCustomer(site, customer, DateUtils.Now.AddMonths(-1), DateUtils.Now.AddMonths(+3));
                servicePage = serviceGeneralInformationPage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);

            //Assert           
            Assert.IsTrue(servicePage.GetFirstServiceName().Contains(serviceName), String.Format(MessageErreur.FILTRE_ERRONE, "Search"));
        }

        //_______________________________________________________________CREATE_____________________________________________________________

        // Créer un nouveau point de livraison
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_CreateNewCustomerOrder()
        {
            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();


            // Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
            customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customer, aircraft);
            var customerOrderDetailPage = customerOrderCreateModalPage.Submit();

            var generalInfo = customerOrderDetailPage.ClickOnGeneralInformationTab();
            var customerOrderNumber = generalInfo.GetOrderNumber();
            customerOrderDetailPage.BackToList();

            var name = customerOrderPage.GetCustomerOrderNumber();

            //Assert
            Assert.IsTrue(name.Contains(customerOrderNumber), "Le customer order n'a pas été créé.");
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_Filter_Search()
        {
            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
            customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customer, aircraft);
            var customerOrderItem = customerOrderCreateModalPage.Submit();

            var generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
            var customerOrderNumber = generalInfo.GetOrderNumber();
            customerOrderPage = generalInfo.BackToList();

            customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrderNumber);
            var name = customerOrderPage.GetCustomerOrderNumber();

            //Assert
            Assert.IsTrue(name.Contains(customerOrderNumber), String.Format(MessageErreur.FILTRE_ERRONE, "Search"));
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_Filter_SortBy_OrderDate()
        {
            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            string dateFormat = homePage.GetDateFormatPickerValue();

            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.ResetFilter();

            if (customerOrderPage.CheckTotalNumber() < 20)
            {
                var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customer, aircraft);
                var customerOrderItem = customerOrderCreateModalPage.Submit();
                customerOrderPage = customerOrderItem.BackToList();
                customerOrderPage.ResetFilter();
            }

            if (!customerOrderPage.isPageSizeEqualsTo100())
            {
                customerOrderPage.PageSize("8");
                customerOrderPage.PageSize("100");
            }

            customerOrderPage.Filter(CustomerOrderPage.FilterType.SortBy, "OrderDate");
            var isSortedByOrderDate = customerOrderPage.IsSortedByOrderDate(dateFormat);

            //Assert
            Assert.IsTrue(isSortedByOrderDate, MessageErreur.FILTRE_ERRONE, "Sort by order date");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_Filter_SortBy_Id()
        {
            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.ResetFilter();

            if (customerOrderPage.CheckTotalNumber() < 20)
            {
                var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customer, aircraft);
                var customerOrderItem = customerOrderCreateModalPage.Submit();
                customerOrderPage = customerOrderItem.BackToList();
                customerOrderPage.ResetFilter();
            }

            if (!customerOrderPage.isPageSizeEqualsTo100())
            {
                customerOrderPage.PageSize("8");
                customerOrderPage.PageSize("100");
            }

            customerOrderPage.Filter(CustomerOrderPage.FilterType.SortBy, "Id");
            var isSortedById = customerOrderPage.IsSortedById();

            //Assert
            Assert.IsTrue(isSortedById, MessageErreur.FILTRE_ERRONE, "Sort by id");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_Filter_SortBy_CustomerName()
        {
            // Prepare
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();

            string customerSortBy = "CustomerSortBy" + new Random().Next().ToString();
            string customerCode = new Random().Next().ToString();
            string customerType = TestContext.Properties["CustomerType"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //création d'un autre customer pour populer la base
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();

            var customerCreateModalPage = customersPage.CustomerCreatePage();
            customerCreateModalPage.FillFields_CreateCustomerModalPage(customerSortBy, customerCode, customerType);
            customerCreateModalPage.Create();

            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.ResetFilter();

            var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
            customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerSortBy, aircraft);
            var customerOrderItem = customerOrderCreateModalPage.Submit();
            customerOrderPage = customerOrderItem.BackToList();
            customerOrderPage.ResetFilter();

            if (!customerOrderPage.isPageSizeEqualsTo100())
            {
                customerOrderPage.PageSize("8");
                customerOrderPage.PageSize("100");
            }

            customerOrderPage.Filter(CustomerOrderPage.FilterType.SortBy, "Customer");
            var isSortedByCustomer = customerOrderPage.IsSortedByCustomer();

            //Assert
            Assert.IsTrue(isSortedByCustomer, MessageErreur.FILTRE_ERRONE, "Sort by customer");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_Filter_Site()
        {
            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string itemName = TestContext.Properties["InvoiceService"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            bool newVersionPrint = true;

            // Arrange
            HomePage homePage = LogInAsAdmin();


            // Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.ResetFilter();

            customerOrderPage.ClearDownloads();

            customerOrderPage.ClickInvoiceSteps();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.InvoiceValidate, true);
            customerOrderPage.Filter(CustomerOrderPage.FilterType.Site, site);

            if (customerOrderPage.CheckTotalNumber() < 20)
            {
                // Create
                var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customer, aircraft);
                var customerOrderItem = customerOrderCreateModalPage.Submit();

                customerOrderItem.AddNewItem(itemName, "10");
                customerOrderItem.ValidateCustomerOrder();
                customerOrderItem.BackToList();

                customerOrderPage.ResetFilter();
                customerOrderPage.ClickInvoiceSteps();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.InvoiceValidate, true);
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Site, site);
            }

            var customerOrdersList = customerOrderPage.GetCustomerOrdersIdList();
            bool isGoodSite = false;
            foreach (var customerOrder in customerOrdersList)
            {
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrder);
                var customerOrderItem = customerOrderPage.SelectFirstCustomerOrder();

                var customerOrderGeneralInformationPage = customerOrderItem.ClickOnGeneralInformationTab();
                var customerOrderSite = customerOrderGeneralInformationPage.GetSite();

                if (customerOrderSite != site)
                {
                    isGoodSite = false;
                    break;
                }
                else
                {
                    isGoodSite = true;
                    customerOrderGeneralInformationPage.BackToList();
                }
            }

            Assert.IsTrue(isGoodSite, String.Format(MessageErreur.FILTRE_ERRONE, "Site"));
            customerOrderPage.ClearDownloads();
            DeleteAllFileDownload();
            //Check excel export
            customerOrderPage.ExportExcel(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = customerOrderPage.GetExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

            // Récupération du nom du fichier et construction de l'URL du fichier Excel à ouvrir   
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber(CUSTOMER_ORDER, filePath);
            bool result = OpenXmlExcel.ReadAllDataInColumn("Site", CUSTOMER_ORDER, filePath, site);
            //WebDriver.Close();
            WebDriver.SwitchTo().Window(WebDriver.WindowHandles[0]);

            //Assert
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
            var customerOrderCount = customerOrderPage.CheckTotalNumber().Equals(resultNumber);
            Assert.IsTrue(customerOrderCount, "Le nombre de customer orders affichés ne correspond pas au nombre de lignes du fichier excel exporté");
            Assert.IsTrue(result, String.Format(MessageErreur.FILTRE_ERRONE, "Site"));
        }

        //[Ignore]
        [TestMethod]

        [Timeout(Timeout)]
        public void PR_CO_Filter_ShowAll_Order()
        {
            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string itemName = TestContext.Properties["InvoiceService"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();


            // Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.ResetFilter();

            customerOrderPage.ClickInvoiceSteps();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.ShowAllOrders, true);

            if (customerOrderPage.CheckTotalNumber() < 20)
            {
                // Create
                var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customer, aircraft);
                var customerOrderItem = customerOrderCreateModalPage.Submit();
                customerOrderItem.AddNewItem(itemName, "10");
                customerOrderPage = customerOrderItem.BackToList();

                customerOrderPage.ResetFilter();
                customerOrderPage.ClickInvoiceSteps();
            }

            if (!customerOrderPage.isPageSizeEqualsTo100())
            {
                customerOrderPage.PageSize("8");
                customerOrderPage.PageSize("100");
            }

            customerOrderPage.Filter(CustomerOrderPage.FilterType.InvoiceInProgress, true);
            int nbInProgress = customerOrderPage.CheckTotalNumber();

            WebDriver.Navigate().Refresh();
            customerOrderPage.ResetFilter();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.ShowNotInvoicedOnly, true);
            int nbNotInvoiced = customerOrderPage.CheckTotalNumber();

            customerOrderPage.Filter(CustomerOrderPage.FilterType.ShowInvoicedOnly, true);
            int nbInvoiced = customerOrderPage.CheckTotalNumber();

            customerOrderPage.Filter(CustomerOrderPage.FilterType.ShowAllOrders, true);
            int nbTotal = customerOrderPage.CheckTotalNumber();

            Assert.AreEqual(nbTotal, nbInProgress + nbNotInvoiced + nbInvoiced, String.Format(MessageErreur.FILTRE_ERRONE, "Show all orders"));
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_Filter_Show_Invoiced()
        {
            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string itemName = TestContext.Properties["InvoiceService"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.ResetFilter();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.ShowInvoicedOnly, true);

            if (customerOrderPage.CheckTotalNumber() < 20)
            {
                var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customer, aircraft);
                var customerOrderItem = customerOrderCreateModalPage.Submit();

                customerOrderItem.AddNewItem(itemName, "10");
                customerOrderItem.ValidateCustomerOrder();
                var generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
                var orderNumb = generalInfo.GetOrderNumber();

                var invoicesPage = homePage.GoToAccounting_InvoicesPage();
                var invoiceCreateAutoModalpage = invoicesPage.AutoInvoiceCreatePage();
                invoiceCreateAutoModalpage.FillField_CreateNewAutoInvoice(customer, site, CustomerPickMethod.AllCustomerOrdersInSelectedPeriod);
                var invoiceDetails = invoiceCreateAutoModalpage.Submit();
                // on s'est perdu entre Invoice Details et Invoice index
                homePage.GoToWinrestHome();
                invoicesPage = homePage.GoToAccounting_InvoicesPage();
                invoicesPage.SelectFirstInvoice();
                invoiceDetails.Validate();

                homePage.GoToProduction_CustomerOrderPage();
                customerOrderPage.ResetFilter();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.ShowInvoicedOnly, true);
            }

            if (!customerOrderPage.isPageSizeEqualsTo100())
            {
                customerOrderPage.PageSize("8");
                customerOrderPage.PageSize("100");
            }

            // Assert
            Assert.IsTrue(customerOrderPage.IsInvoiced(), String.Format(MessageErreur.FILTRE_ERRONE, "Show invoiced only"));
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_Filter_Show_NotInvoiced()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string itemName = TestContext.Properties["InvoiceService"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.ResetFilter();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.ShowNotInvoicedOnly, true);

            if (customerOrderPage.CheckTotalNumber() < 20)
            {
                var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
                var customerOrderItem = customerOrderCreateModalPage.Submit();

                customerOrderItem.AddNewItem(itemName, "10");
                customerOrderItem.ValidateCustomerOrder();
                customerOrderItem.BackToList();

                customerOrderPage.ResetFilter();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.ShowNotInvoicedOnly, true);
            }

            if (!customerOrderPage.isPageSizeEqualsTo100())
            {
                customerOrderPage.PageSize("8");
                customerOrderPage.PageSize("100");
            }

            // Assert
            bool customerOrderIsNotInvoiced = customerOrderPage.IsNotInvoiced();
            Assert.IsTrue(customerOrderIsNotInvoiced, String.Format(MessageErreur.FILTRE_ERRONE, "Show not invoiced only"));
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_Filter_Show_Status_All()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string itemName = TestContext.Properties["InvoiceService"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();

            // Create
            var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
            customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
            var customerOrderItem = customerOrderCreateModalPage.Submit();
            customerOrderItem.AddNewItem(itemName, "10");
            var generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
            generalInfo.ChangeStatus("Opened");
            var orderNumb = generalInfo.GetOrderNumber();
            customerOrderPage = customerOrderItem.BackToList();

            customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, orderNumb);
            customerOrderPage.ClickStatus();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.StatusClosed, true);
            Assert.AreEqual(0, customerOrderPage.CheckTotalNumber(), "Le customer order créé apparaît en tant que Closed");

            customerOrderPage.Filter(CustomerOrderPage.FilterType.StatusAll, true);
            Assert.AreEqual(1, customerOrderPage.CheckTotalNumber(), String.Format(MessageErreur.FILTRE_ERRONE, "Status ALL"));

            customerOrderItem = customerOrderPage.SelectFirstCustomerOrder();
            generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
            generalInfo.ChangeStatus("Closed");
            customerOrderPage = customerOrderItem.BackToList();

            customerOrderPage.ResetFilter();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, orderNumb);
            customerOrderPage.ClickStatus();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.StatusClosed, true);
            Assert.AreEqual(1, customerOrderPage.CheckTotalNumber(), "Le customer order créé n'apparaît pas en tant que Closed");

            customerOrderPage.Filter(CustomerOrderPage.FilterType.StatusAll, true);
            Assert.AreEqual(1, customerOrderPage.CheckTotalNumber(), String.Format(MessageErreur.FILTRE_ERRONE, "Status ALL"));
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_Filter_Show_Status_Opened()
        {
            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string itemName = TestContext.Properties["InvoiceService"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            bool newVersionPrint = true;

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.ResetFilter();

            customerOrderPage.ClearDownloads();

            // Create
            var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
            customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customer, aircraft);
            var customerOrderItem = customerOrderCreateModalPage.Submit();
            customerOrderItem.AddNewItem(itemName, "10");
            var generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
            generalInfo.ChangeStatus("Opened");
            var orderNumb = generalInfo.GetOrderNumber();
            customerOrderPage = customerOrderItem.BackToList();

            customerOrderPage.ResetFilter();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, orderNumb);
            customerOrderPage.ClickStatus();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.StatusOpened, true);
            Assert.AreEqual(1, customerOrderPage.CheckTotalNumber(), String.Format(MessageErreur.FILTRE_ERRONE, "Opened"));

            // Reset des filtres pour vérifier Export
            customerOrderPage.ResetFilter();
            customerOrderPage.ClickStatus();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.StatusOpened, true);

            customerOrderPage.ExportExcel(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = customerOrderPage.GetExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

            // Récupération du nom du fichier et construction de l'URL du fichier Excel à ouvrir   
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber(CUSTOMER_ORDER, filePath);
            bool result = OpenXmlExcel.ReadAllDataInColumn("Status", CUSTOMER_ORDER, filePath, "Opened");

            //Assert
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.IsTrue(result, String.Format(MessageErreur.FILTRE_ERRONE, "Export Opened"));
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_Filter_Show_Status_Closed()
        {
            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string itemName = TestContext.Properties["InvoiceService"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            bool newVersionPrint = true;

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.ResetFilter();

            customerOrderPage.ClearDownloads();

            // Create
            var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
            customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customer, aircraft);
            var customerOrderItem = customerOrderCreateModalPage.Submit();
            customerOrderItem.AddNewItem(itemName, "10");
            var generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
            generalInfo.ChangeStatus("Closed");
            var orderNumb = generalInfo.GetOrderNumber();
            customerOrderPage = customerOrderItem.BackToList();

            customerOrderPage.ResetFilter();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, orderNumb);
            customerOrderPage.ClickStatus();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.StatusClosed, true);
            Assert.AreEqual(1, customerOrderPage.CheckTotalNumber(), String.Format(MessageErreur.FILTRE_ERRONE, "Closed"));

            // Reset des filtres pour vérifier Export
            customerOrderPage.ResetFilter();
            customerOrderPage.ClickStatus();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.StatusClosed, true);

            customerOrderPage.ExportExcel(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = customerOrderPage.GetExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

            // Récupération du nom du fichier et construction de l'URL du fichier Excel à ouvrir   
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber(CUSTOMER_ORDER, filePath);
            bool result = OpenXmlExcel.ReadAllDataInColumn("Status", CUSTOMER_ORDER, filePath, "Closed");

            //Assert
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.IsTrue(result, String.Format(MessageErreur.FILTRE_ERRONE, "Export Closed"));
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_Filter_Show_InvoiceStep_All()
        {
            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string itemName = TestContext.Properties["InvoiceService"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.ResetFilter();

            customerOrderPage.ClickInvoiceSteps();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.InvoiceAll, true);

            if (customerOrderPage.CheckTotalNumber() < 50)
            {
                // Create
                var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customer, aircraft);
                var customerOrderItem = customerOrderCreateModalPage.Submit();
                customerOrderItem.AddNewItem(itemName, "10");
                customerOrderPage = customerOrderItem.BackToList();

                customerOrderPage.ResetFilter();
                customerOrderPage.ClickInvoiceSteps();
            }

            if (!customerOrderPage.isPageSizeEqualsTo100())
            {
                customerOrderPage.PageSize("8");
                customerOrderPage.PageSize("100");
            }

            customerOrderPage.Filter(CustomerOrderPage.FilterType.InvoiceInProgress, true);
            int nbInProgress = customerOrderPage.CheckTotalNumber();

            customerOrderPage.Filter(CustomerOrderPage.FilterType.InvoiceValidate, true);
            int nbValidate = customerOrderPage.CheckTotalNumber();

            customerOrderPage.Filter(CustomerOrderPage.FilterType.InvoiceInvoiced, true);
            int nbInvoiced = customerOrderPage.CheckTotalNumber();

            customerOrderPage.Filter(CustomerOrderPage.FilterType.InvoiceAll, true);
            int nbTotal = customerOrderPage.CheckTotalNumber();

            Assert.AreEqual(nbTotal, nbInProgress + nbValidate + nbInvoiced, String.Format(MessageErreur.FILTRE_ERRONE, "Invoice steps : All"));
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_Filter_Show_InvoiceStep_InProgress()
        {
            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string itemName = TestContext.Properties["InvoiceService"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            bool newVersionPrint = true;

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.ResetFilter();

            customerOrderPage.ClearDownloads();

            customerOrderPage.ClickInvoiceSteps();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.InvoiceInProgress, true);

            if (customerOrderPage.CheckTotalNumber() < 50)
            {
                // Create
                var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customer, aircraft);
                var customerOrderItem = customerOrderCreateModalPage.Submit();
                customerOrderItem.AddNewItem(itemName, "10");
                customerOrderPage = customerOrderItem.BackToList();

                customerOrderPage.ResetFilter();

                customerOrderPage.ClickInvoiceSteps();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.InvoiceInProgress, true);
            }

            customerOrderPage.ExportExcel(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = customerOrderPage.GetExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

            // Récupération du nom du fichier et construction de l'URL du fichier Excel à ouvrir   
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber(CUSTOMER_ORDER, filePath);
            bool result = OpenXmlExcel.ReadAllDataInColumn("Invoice step", CUSTOMER_ORDER, filePath, "InProgress");

            //Assert
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.IsTrue(result, String.Format(MessageErreur.FILTRE_ERRONE, "Invoice steps : In progress"));
        }

        [TestMethod]

        [Timeout(Timeout)]
        public void PR_CO_Filter_Show_InvoiceStep_Validate()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string itemName = TestContext.Properties["InvoiceService"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();


            // Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.ResetFilter();

            customerOrderPage.ClickInvoiceSteps();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.InvoiceValidate, true);

            if (customerOrderPage.CheckTotalNumber() < 20)
            {
                var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
                var customerOrderItem = customerOrderCreateModalPage.Submit();

                customerOrderItem.AddNewItem(itemName, "10");
                customerOrderItem.ValidateCustomerOrder();
                customerOrderItem.BackToList();

                customerOrderPage.ResetFilter();

                customerOrderPage.ClickInvoiceSteps();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.InvoiceValidate, true);
            }

            if (!customerOrderPage.isPageSizeEqualsTo100())
            {
                customerOrderPage.PageSize("8");
                customerOrderPage.PageSize("100");
            }

            // Assert
            Assert.IsTrue(customerOrderPage.CheckValidation(true), String.Format(MessageErreur.FILTRE_ERRONE, "Invoice steps : Validated"));
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_Filter_Show_InvoiceStep_Invoiced()
        {
            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string itemName = TestContext.Properties["InvoiceService"].ToString();

            //arrange
            HomePage homePage = LogInAsAdmin();

            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.ResetFilter();

            customerOrderPage.ClickInvoiceSteps();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.InvoiceInvoiced, true);

            if (customerOrderPage.CheckTotalNumber() < 20)
            {
                var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customer, aircraft);
                var customerOrderItem = customerOrderCreateModalPage.Submit();

                customerOrderItem.AddNewItem(itemName, "10");
                customerOrderItem.ValidateCustomerOrder();
                var generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
                var orderNumb = generalInfo.GetOrderNumber();

                var invoicesPage = homePage.GoToAccounting_InvoicesPage();
                var invoiceCreateAutoModalpage = invoicesPage.AutoInvoiceCreatePage();
                invoiceCreateAutoModalpage.FillField_CreateNewAutoInvoice(customer, site, CustomerPickMethod.AllCustomerOrdersInSelectedPeriod);
                invoiceCreateAutoModalpage.FillFieldSelectSomes(Int32.Parse(orderNumb));

                // on s'est perdu entre Invoice Details et Invoice index
                homePage.GoToWinrestHome();
                invoicesPage = homePage.GoToAccounting_InvoicesPage();
                var invoiceDetails = invoicesPage.SelectFirstInvoice();
                invoiceDetails.Validate();

                homePage.GoToProduction_CustomerOrderPage();
                customerOrderPage.ResetFilter();

                customerOrderPage.ClickInvoiceSteps();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.InvoiceInvoiced, true);
            }

            if (!customerOrderPage.isPageSizeEqualsTo100())
            {
                customerOrderPage.PageSize("8");
                customerOrderPage.PageSize("100");
            }
            bool customerOrderIsInvoiced = customerOrderPage.IsInvoiced();
            // Assert
            Assert.IsTrue(customerOrderIsInvoiced, String.Format(MessageErreur.FILTRE_ERRONE, "Invoice steps : invoiced"));
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_Filter_Customer()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string customerCode = TestContext.Properties["InvoiceCustomerCode"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string itemName = TestContext.Properties["InvoiceService"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();

            string customer = customerCode + " - " + customerName;

            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.ResetFilter();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.Customers, customer);

            if (customerOrderPage.CheckTotalNumber() < 20)
            {
                var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
                var customerOrderItem = customerOrderCreateModalPage.Submit();

                customerOrderItem.AddNewItem(itemName, "10");
                customerOrderItem.ValidateCustomerOrder();
                customerOrderItem.BackToList();

                customerOrderPage.ResetFilter();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Customers, customer);
            }

            if (!customerOrderPage.isPageSizeEqualsTo100())
            {
                customerOrderPage.PageSize("8");
                customerOrderPage.PageSize("100");
            }
            bool verifyCustomer = customerOrderPage.VerifyCustomer(customerName);
            //Assert
            Assert.IsTrue(verifyCustomer, string.Format(MessageErreur.FILTRE_ERRONE, "Customers"));
        }

        //[Ignore]
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_GeneralInfo_EditDelivery()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string itemName = TestContext.Properties["InvoiceService"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string category = "COLECTIVIDADES";
            string delivery = GenerateName(6) + " delivery";
            DateTime date = DateUtils.Now.AddDays(10);

            // Arrange
            HomePage homePage = LogInAsAdmin();
            string dateFormat = homePage.GetDateFormatPickerValue();
            // Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            try 
            {
                var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
                var customerOrderItem = customerOrderCreateModalPage.Submit();
                customerOrderItem.AddNewItemWithCategory(itemName, "10", category);
                customerOrderItem.ValidateCustomerOrder();
                var generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
                var siteEditable = generalInfo.IsSelectSiteEditable();
                //Assert
                Assert.IsFalse(siteEditable, "Le site est editable");
                bool deliveryEdit = generalInfo.CanEditDelivery();
                generalInfo.BackToList();
                //Assert
                Assert.IsFalse(deliveryEdit, "Le delivery peut être édité sur un customer order validé.");
                
            }
            finally
            {
                customerOrderPage.DeleteCustomerOrder();
            }
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_Update_Item()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string itemName = TestContext.Properties["InvoiceService"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string quantity = "10";
            string price = "2";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            try 
            {
                var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
                var customerOrderItem = customerOrderCreateModalPage.Submit();
                customerOrderItem.AddNewItem(itemName, "10");

                customerOrderItem.SetItemQuantity(quantity);
                customerOrderItem.SetItemPrice(price);

                var generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
                var orderN = generalInfo.GetOrderNumber();
                customerOrderItem = generalInfo.ClickOnDetailTab();

                //Assert
                var itemQuantity = customerOrderItem.GetItemQuantity();
                Assert.AreEqual(quantity, itemQuantity, "La quantité de l'item n'a pas été modifiée.");
                var itemPrice = customerOrderItem.GetItemPrice();
                Assert.AreEqual(price, itemPrice, "Le prix de l'item n'a pas été modifié.");

                generalInfo.BackToList();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, orderN);
            }
            finally
            {
                customerOrderPage.DeleteCustomerOrder();
            }
            
        }

        //[Ignore]
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_Add_Item()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string itemName = TestContext.Properties["InvoiceService"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
            customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
            var customerOrderItem = customerOrderCreateModalPage.Submit();
            customerOrderItem.AddNewItem(itemName, "10");

            //Assert
            Assert.IsTrue(customerOrderItem.IsVisible(), "L'item n'a pas été ajouté.");
        }

        [TestMethod]


        [Timeout(Timeout)]
        public void PR_CO_Delete_Item()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string itemName = TestContext.Properties["InvoiceService"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();


            // Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
            customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
            var customerOrderItem = customerOrderCreateModalPage.Submit();
            customerOrderItem.AddNewItem(itemName, "10");
            Assert.IsTrue(customerOrderItem.IsVisible(), "L'item n'a pas été ajouté.");

            customerOrderItem.DeleteItem();

            //Assert
            Assert.IsFalse(customerOrderItem.IsVisible(), "L'item n'a pas été supprimé.");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_Add_ProdComment_Item()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string itemName = TestContext.Properties["InvoiceService"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string prodComment = "comment";

            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
            customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
            var customerOrderItem = customerOrderCreateModalPage.Submit();
            customerOrderItem.AddNewItem(itemName, "10");

            bool isComment = customerOrderItem.AddProdComment(prodComment);
            Assert.IsTrue(isComment, "Un commentaire a été ajouté.");

            customerOrderItem.ClickProdComment();
            var comment =customerOrderItem.GetComment();
            Assert.AreEqual(prodComment, comment, "Le commentaire n'a pas été ajouté correctement.");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_Add_BillinComment_Item()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string itemName = TestContext.Properties["InvoiceService"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string billingComment = "comment";

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
            customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
            var customerOrderItem = customerOrderCreateModalPage.Submit();
            customerOrderItem.AddNewItem(itemName, "10");

            bool isComment = customerOrderItem.AddBillingComment(billingComment);
            Assert.IsTrue(isComment, "Un commentaire a été ajouté.");

            customerOrderItem.ClickBillingComment();
            Assert.AreEqual(billingComment, customerOrderItem.GetComment(), "Le commentaire n'a pas été ajouté correctement.");
        }

        [TestMethod]


        [Timeout(Timeout)]
        public void PR_CO_Print_CustomerOrder_NewVersion()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string itemName = TestContext.Properties["InvoiceService"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();

            bool newVersionPrint = true;

            // Arrange
            HomePage homePage = LogInAsAdmin();


            // Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();

            customerOrderPage.ClearDownloads();

            var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
            customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
            var customerOrderItem = customerOrderCreateModalPage.Submit();
            customerOrderItem.AddNewItem(itemName, "10");

            var reportPage = customerOrderItem.PrintCustomerOrder(newVersionPrint);
            var isReportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();

            //Assert
            Assert.IsTrue(isReportGenerated, "Le document PDF n'a pas pu être généré par l'application.");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_PrintProd_CustomerOrder_NewVersion()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string itemName = TestContext.Properties["InvoiceService"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();

            bool newVersionPrint = true;

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();

            customerOrderPage.ClearDownloads();
            try 
            {
                var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
                var customerOrderItem = customerOrderCreateModalPage.Submit();
                var generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
                var orderN = generalInfo.GetOrderNumber();
                var reportPage = customerOrderItem.PrintProdCustomerOrder(newVersionPrint);
                var isReportGenerated = reportPage.IsReportGenerated();
                reportPage.Close();

                //Assert
                Assert.IsTrue(isReportGenerated, "Le document PDF n'a pas pu être généré par l'application.");
                generalInfo.BackToList();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, orderN);
            }
            finally 
            {
                customerOrderPage.DeleteCustomerOrder();
            }
            
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_Validate()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string itemName = TestContext.Properties["InvoiceService"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
            customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
            var customerOrderItem = customerOrderCreateModalPage.Submit();

            customerOrderItem.AddNewItem(itemName, "10");
            customerOrderItem.ValidateCustomerOrder();

            var generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
            var orderNumb = generalInfo.GetOrderNumber();
            customerOrderItem.BackToList();
            customerOrderPage.ResetFilter();

            if (!customerOrderPage.isPageSizeEqualsTo100())
            {
                customerOrderPage.PageSize("8");
                customerOrderPage.PageSize("100");
            }

            customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, orderNumb);

            // Assert
            Assert.IsTrue(customerOrderPage.CheckValidation(true), "Le customer order n'a pas été validé.");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_Delete_Customer_Order()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string itemName = TestContext.Properties["InvoiceService"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
            customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
            var customerOrderItem = customerOrderCreateModalPage.Submit();
            customerOrderItem.AddNewItem(itemName, "10");
            var generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
            var orderNumb = generalInfo.GetOrderNumber();
            customerOrderItem.BackToList();
            customerOrderPage.ResetFilter();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, orderNumb);
            int totalNumber =customerOrderPage.CheckTotalNumber();

            Assert.AreEqual(1, totalNumber, "Le customer order n'a pas été créé.");

            customerOrderPage.DeleteCustomerOrder();
            customerOrderPage.ResetFilter();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, orderNumb);
            int totalNumber1 = customerOrderPage.CheckTotalNumber();
            //Assert
            Assert.AreEqual(0, totalNumber1, "Le customer ordre n'a pas été supprimé.");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_SendByMail()
        {
            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string itemName = TestContext.Properties["InvoiceService"].ToString();
            string userEmail = TestContext.Properties["Admin_UserName"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();


            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.ResetFilter();

            var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
            customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customer, aircraft);
            var customerOrderItem = customerOrderCreateModalPage.Submit();

            customerOrderItem.AddNewItem(itemName, "10");
            customerOrderItem.ValidateCustomerOrder();
            var generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
            var customerOrderName = generalInfo.GetOrderNumber();

            homePage.GoToProduction_CustomerOrderPage();

            customerOrderPage.ResetFilter();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrderName);
            customerOrderPage.Filter(CustomerOrderPage.FilterType.Site, site);
            customerOrderPage.Filter(CustomerOrderPage.FilterType.Customers, customer);

            customerOrderPage.SendByMail();

            WebDriver.Navigate().Refresh();
            //customerOrderPage.ResetFilter();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrderName);
            customerOrderPage.Filter(CustomerOrderPage.FilterType.Site, site);
            customerOrderPage.Filter(CustomerOrderPage.FilterType.Customers, customer);
            bool verifyMailSentByMail = customerOrderPage.IsSentByMail();
            Assert.IsTrue(verifyMailSentByMail, "Le customer order validé n'a pas été envoyé par mail.");

            //open mail box 
            MailPage mailPage = customerOrderPage.RedirectToOutlookMailbox();
            mailPage.FillFields_LogToOutlookMailbox(userEmail);
            // check if mail sent
            bool isMailSent = mailPage.CheckIfSpecifiedOutlookMailExist("Winrest - 1 order - " + site);
            Assert.IsTrue(isMailSent, "La claim n'a pas été envoyée par mail.");
            mailPage.DeleteCurrentOutlookMail();
            mailPage.Close();
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_Export_CustomerOrder_NewVersion()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string itemName = TestContext.Properties["InvoiceService"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            bool newVersionPrint = true;

            // Arrange
            HomePage homePage = LogInAsAdmin();


            DeleteAllFileDownload();

            // Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.ResetFilter();

            customerOrderPage.ClearDownloads();

            customerOrderPage.Filter(CustomerOrderPage.FilterType.Site, site);
            customerOrderPage.Filter(CustomerOrderPage.FilterType.From, DateUtils.Now.AddDays(-1));

            if (customerOrderPage.CheckTotalNumber() < 20)
            {
                var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
                var customerOrderItem = customerOrderCreateModalPage.Submit();
                customerOrderItem.AddNewItem(itemName, "10");
                customerOrderItem.ValidateCustomerOrder();
                customerOrderPage = customerOrderItem.BackToList();

                customerOrderPage.ResetFilter();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Site, site);
                customerOrderPage.Filter(CustomerOrderPage.FilterType.From, DateUtils.Now.AddDays(-1));
            }

            customerOrderPage.ExportExcel(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = customerOrderPage.GetExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

            // Récupération du nom du fichier et construction de l'URL du fichier Excel à ouvrir   
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber(CUSTOMER_ORDER, filePath);

            //Assert
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_IsSubjectToAirportTax_AddService()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string customerCode = TestContext.Properties["InvoiceCustomerCode"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string itemName = TestContext.Properties["InvoiceService"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string taxName = TestContext.Properties["InvoiceTaxName"].ToString();
            string taxType = TestContext.Properties["InvoiceTaxType"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            bool newVersionPrint = true;

            string taxValue = "15";

            string purchaseCode = "AS10";
            string purchaseAccount = "47205001";
            string salesCode = "AR10";
            string salesAccount = "47205001";
            string dueToInvoiceCode = "AP10";
            string dueToInvoiceAccount = "47205001";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //_____________________________ Mise en place des paramètres____________________________
            // Client avec AirportTax = true
            var customerPage = homePage.GoToCustomers_CustomerPage();
            customerPage.Filter(CustomerPage.FilterType.Search, customerName);
            var editPage = customerPage.SelectFirstCustomer();
            editPage.ActivateAirportTax();

            // Avoir une airport tax crée
            var purchasingPage = homePage.GoToParameters_PurchasingPage();
            purchasingPage.GoToTab_VAT();

            if (!purchasingPage.IsTaxPresent(taxName, taxType))
            {
                purchasingPage.CreateNewVAT(taxName, taxType, taxValue);
            }

            purchasingPage.SearchVAT(taxName, taxType);
            purchasingPage.EditVATAccountForSpain(purchaseCode, purchaseAccount, salesCode, salesAccount, dueToInvoiceCode, dueToInvoiceAccount);

            // Airport tax configurée pour site et client
            var accountingPage = homePage.GoToParameters_AccountingPage();
            accountingPage.GoToTab_AirportFee();
            accountingPage.SetAirportFeeSiteAndCustomer(site, customerCode, customerName, "15");

            // Catégorie de service avec AirportTax = true
            var parametersCustomerPage = homePage.GoToParameters_CustomerPage();
            parametersCustomerPage.GoToTab_Category();
            parametersCustomerPage.SetServiceWithAirportTax(itemName);

            //_____________________________ Fin Mise en place des paramètres________________________

            // Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.ResetFilter();

            customerOrderPage.ClearDownloads();

            customerOrderPage.Filter(CustomerOrderPage.FilterType.Site, site);

            var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
            customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
            var customerOrderItem = customerOrderCreateModalPage.Submit();

            customerOrderItem.AddNewItem(itemName, "10");
            customerOrderItem.ValidateCustomerOrder();

            var generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
            var customerOrderNumber = generalInfo.GetOrderNumber();
            customerOrderItem.BackToList();

            customerOrderPage.ResetFilter();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.Site, site);
            customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrderNumber);

            customerOrderPage.ExportExcel(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = customerOrderPage.GetExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

            // Récupération du nom du fichier et construction de l'URL du fichier Excel à ouvrir   
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber(CUSTOMER_ORDER, filePath);
            bool result = OpenXmlExcel.ReadAllDataInColumn("Is subject to airport tax", CUSTOMER_ORDER, filePath, "VRAI");

            //Assert
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.IsTrue(result, "L'airport tax n'est pas prise en compte pour le customer order.");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_IsSubjectToAirportTax_AddFreePrice()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string customerCode = TestContext.Properties["InvoiceCustomerCode"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string itemName = TestContext.Properties["CustomerOrderFreePrice"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string taxName = TestContext.Properties["InvoiceTaxName"].ToString();
            string taxType = TestContext.Properties["InvoiceTaxType"].ToString();
            string workshop = TestContext.Properties["InvoiceWorkshop"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            bool newVersionPrint = true;
            bool isFreePrice;

            string taxValue = "15";

            string purchaseCode = "AS10";
            string purchaseAccount = "47205001";
            string salesCode = "AR10";
            string salesAccount = "47205001";
            string dueToInvoiceCode = "AP10";
            string dueToInvoiceAccount = "47205001";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            //_____________________________ Mise en place des paramètres____________________________
            // Client avec AirportTax = true
            var customerPage = homePage.GoToCustomers_CustomerPage();
            customerPage.Filter(CustomerPage.FilterType.Search, customerName);
            var editPage = customerPage.SelectFirstCustomer();
            editPage.ActivateAirportTax();

            // Avoir une airport tax crée
            var purchasingPage = homePage.GoToParameters_PurchasingPage();
            purchasingPage.GoToTab_VAT();

            if (!purchasingPage.IsTaxPresent(taxName, taxType))
            {
                purchasingPage.CreateNewVAT(taxName, taxType, taxValue);
            }

            purchasingPage.SearchVAT(taxName, taxType);
            purchasingPage.EditVATAccountForSpain(purchaseCode, purchaseAccount, salesCode, salesAccount, dueToInvoiceCode, dueToInvoiceAccount);

            // Airport tax configurée pour site et client
            var accountingPage = homePage.GoToParameters_AccountingPage();
            accountingPage.GoToTab_AirportFee();
            accountingPage.SetAirportFeeSiteAndCustomer(site, customerCode, customerName, "15");

            // Catégorie de service avec AirportTax = true
            var parametersCustomerPage = homePage.GoToParameters_CustomerPage();
            parametersCustomerPage.GoToTab_Category();
            parametersCustomerPage.SetServiceWithAirportTax(itemName);

            //_____________________________ Fin Mise en place des paramètres________________________

            // Vérification du freePrice
            var freePricePage = homePage.GoToAccounting_FreePricePage();
            freePricePage.Filter(FreePricePage.FilterType.Search, itemName);

            if (freePricePage.CheckTotalNumber_freePrice() == 0)
            {
                isFreePrice = false;
            }
            else
            {
                var freePriceDetails = freePricePage.SelectFirstFreePrice();
                freePriceDetails.SetTaxType("Airport Tax");
                isFreePrice = true;
            }

            // Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.ResetFilter();

            customerOrderPage.ClearDownloads();

            customerOrderPage.Filter(CustomerOrderPage.FilterType.Site, site);

            var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
            customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
            var customerOrderItem = customerOrderCreateModalPage.Submit();

            if (!isFreePrice)
            {
                var freePriceModal = customerOrderItem.AddFreePrice(itemName);
                freePriceModal.FillField_CreatNewFreePrice(itemName, "10", "10", workshop);
                customerOrderItem = freePriceModal.ValidateForCustomerOrder();
                customerOrderItem.ValidateCustomerOrder();
            }
            else
            {
                customerOrderItem.AddNewItem(itemName, "10");
                customerOrderItem.ValidateCustomerOrder();
            }

            var generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
            var customerOrderNumber = generalInfo.GetOrderNumber();

            if (!isFreePrice)
            {
                freePricePage = homePage.GoToAccounting_FreePricePage();
                freePricePage.Filter(FreePricePage.FilterType.Search, itemName);
                var freePriceDetails = freePricePage.SelectFirstFreePrice();
                freePriceDetails.SetTaxType("Airport Tax");
            }

            customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.ResetFilter();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.Site, site);
            customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrderNumber);

            customerOrderPage.ExportExcel(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = customerOrderPage.GetExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

            // Récupération du nom du fichier et construction de l'URL du fichier Excel à ouvrir   
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber(CUSTOMER_ORDER, filePath);
            bool result = OpenXmlExcel.ReadAllDataInColumn("Is subject to airport tax", CUSTOMER_ORDER, filePath, "VRAI");

            //Assert
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.IsTrue(result, "L'airport tax n'est pas prise en compte pour le customer order.");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_PrintShopList_CustomerOrder()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string itemName = TestContext.Properties["InvoiceService"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();

            bool newVersionPrint = true;

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();

            customerOrderPage.ClearDownloads();

            var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
            customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
            var customerOrderItem = customerOrderCreateModalPage.Submit();
            customerOrderItem.AddNewItem(itemName, "10");

            var reportPage = customerOrderItem.PrintShopListCustomerOrder(newVersionPrint);
            var isReportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();

            //Assert
            Assert.IsTrue(isReportGenerated, "Le document PDF n'a pas pu être généré par l'application.");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_PrintCustomerOrdersList()
        {
            //print max 20 results and print results according page size

            string site = TestContext.Properties["InvoiceSite"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            var pageSizeValue = 16;

            // Arrange
            HomePage homePage = LogInAsAdmin();


            homePage.ClearDownloads();

            // Act
            //Etre sur l'index des Customer Order
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.ResetFilter();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.Site, site);
            customerOrderPage.Filter(CustomerOrderPage.FilterType.InvoiceValidate, true);
            customerOrderPage.PageSize(pageSizeValue.ToString());
            //1. Survoler les ...
            //2.Cliquer sur Print
            DeleteAllFileDownload();
            PrintReportPage printReport = customerOrderPage.Print();
            var isReportGenerated = printReport.IsReportGenerated();
            printReport.Close();
            Assert.IsTrue(isReportGenerated, "Fichier non généré");

            printReport.Purge(downloadsPath, "Orders Report_-_", "All_files_");
            string trouve = printReport.PrintAllZip(downloadsPath, "Orders Report_-_", "All_files_", false);
            FileInfo file = new FileInfo(trouve);
            //Le fichier PDF est généré
            Assert.IsTrue(file.Exists, "Fichier PDF non généré");

            //Vérifier que les données correspondent
            PdfDocument document = PdfDocument.Open(file.FullName);
            List<string> mots = new List<string>();
            foreach (Page p in document.GetPages())
            {
                foreach (var mot in p.GetWords())
                {
                    mots.Add(mot.Text);
                }
            }

            var orderNumbers = WebDriver.FindElements(By.XPath("//*/tr/td[5]"));
            int counter = 0;
            foreach (var orderNumber in orderNumbers)
            {
                if (orderNumber.Text == "")
                    continue;
                counter++;
                Assert.IsTrue(mots.Contains(orderNumber.Text.Substring(3)), "order number " + orderNumber.Text.Substring(3) + " non présent dans le PDF");
            }
            Assert.AreEqual(Math.Min(pageSizeValue, customerOrderPage.GetCustomerOrdersCount()), counter, "Différence entre les CustomerOrder number du tableau et du PDF");

            var customerNames = WebDriver.FindElements(By.XPath("//*/tr/td[7]/b"));
            counter = 0;
            foreach (var customerName in customerNames)
            {
                if (customerName.Text == "") continue;
                counter++;
                Assert.IsTrue(mots.Contains(customerName.Text.Split(' ')[0]), "Customer Name " + customerName.Text.Split(' ')[0] + " non présent dans le PDF");
            }
            Assert.AreEqual(Math.Min(pageSizeValue, customerOrderPage.GetCustomerOrdersCount()), counter, "Différence entre les CustomerName number du tableau et du PDF");
            ReadOnlyCollection<IWebElement> deliveries;
            deliveries = WebDriver.FindElements(By.XPath("//*/tr/td[11]"));

            counter = 0;
            foreach (var delivery in deliveries)
            {
                if (delivery.Text == "") continue;
                counter++;
                Assert.IsTrue(mots.Contains(delivery.Text.Trim().Substring(0, "18/07/2023".Length)), "Delivery " + delivery.Text.Trim() + " non présent dans le PDF");
            }
            Assert.AreEqual(Math.Min(pageSizeValue, customerOrderPage.GetCustomerOrdersCount()), counter, "Différence entre les Delivery du tableau et du PDF");

            ReadOnlyCollection<IWebElement> orderDates;
            orderDates = WebDriver.FindElements(By.XPath("//*/tr/td[12]"));

            counter = 0;
            foreach (var orderDate in orderDates)
            {
                if (orderDate.Text == "") continue;
                counter++;
                Assert.IsTrue(mots.Contains(orderDate.Text.Split(' ')[0]), "Expedition" + orderDate.Text.Split(' ')[0] + " non présent dans le PDF");
            }
            Assert.AreEqual(Math.Min(pageSizeValue, customerOrderPage.GetCustomerOrdersCount()), counter, "Différence entre les Order Date du tableau et du PDF");

            IReadOnlyCollection<IWebElement> prices;
            prices = WebDriver.FindElements(By.XPath("//*/tr/td[13]"));

            counter = 0;
            foreach (var price in prices)
            {
                if (price.Text == "") continue;
                counter++;
                string priceText;
                if (price.Text.StartsWith("SEK "))
                {
                    priceText = price.Text.Trim().Substring(4);
                }
                else
                {
                    priceText = price.Text.Trim().Substring(2);
                }
                if (!priceText.Contains(","))
                {
                    priceText = priceText + ",00";
                }
                else
                {
                    priceText = priceText + "000";
                }

                Assert.IsTrue(mots.Contains(priceText), "Price " + priceText + " non présent dans le PDF");
            }
            Assert.AreEqual(Math.Min(pageSizeValue, customerOrderPage.GetCustomerOrdersCount()), counter, "Différence entre les Prices du tableau et du PDF");


        }

        [Ignore] // Orders.IsPreOrder utilisé par JustEat et BoB (Arthur)
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_PrintPreOrderCustomerOrdersPackingList()
        {
            //Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.ResetFilter();
            //A développer (ce print affiche seulement les données des customers orders qui ont le PreOrder YES.)
            customerOrderPage.Filter(CustomerOrderPage.FilterType.PreOrderYes, true);
            //select *,IsPreorder from Orders where Id = 52861
            //update Orders set IsPreorder = 1 where Id = 52861
            if (customerOrderPage.CheckTotalNumber() == 0)
            {
                CustomerOrderCreateModalPage customerOrder = customerOrderPage.CustomerOrderCreatePage();
                customerOrder.FillField_CreatNewCustomerOrder(site, customer, aircraft);
                CustomerOrderItem coItem = customerOrder.Submit();
                CustomerOrderGeneralInformationPage generalInfo = coItem.ClickOnGeneralInformationTab();
                //generalInfo
                customerOrderPage = generalInfo.BackToList();
            }
            //1.Survoler les "…"
            //customerOrderPage.PrintPreOrderPackingList();
            //2.Cliquer sur Print PreOrder Packing List
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_Filter_IsOnlyDelivered()
        {
            // Prepare
            string customerOrderDeliveryId = TestContext.Properties[string.Format("CustomerOrderId5")].ToString();
            string customerOrderNotDeliveryId = TestContext.Properties[string.Format("CustomerOrderId6")].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.ResetFilter();

            customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrderDeliveryId);
            customerOrderPage.Filter(CustomerOrderPage.FilterType.IsDelivered, true);
            customerOrderPage.PageUpCO();
            string firstCustomerOrder = customerOrderPage.GetCustomerOrderNumber().Split(' ')[1];

            //Assert
            Assert.AreEqual(firstCustomerOrder, customerOrderDeliveryId, MessageErreur.FILTRE_ERRONE, "IsDelivery");

            customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrderNotDeliveryId);
            int customerOrdersNbr = customerOrderPage.CheckTotalNumber();

            //Assert
            Assert.AreEqual(customerOrdersNbr, 0, MessageErreur.FILTRE_ERRONE, "IsDelivery");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_Filter_IsOnlyNotDelivered()
        {
            // Prepare
            string customerOrderDeliveryId = TestContext.Properties[string.Format("CustomerOrderId5")].ToString();
            string customerOrderNotDeliveryId = TestContext.Properties[string.Format("CustomerOrderId6")].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.ResetFilter();

            customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrderNotDeliveryId);
            customerOrderPage.Filter(CustomerOrderPage.FilterType.IsNotDelivered, true);
            customerOrderPage.PageUpCO();
            string firstCustomerOrder = customerOrderPage.GetCustomerOrderNumber().Split(' ')[1];

            //Assert
            Assert.AreEqual(firstCustomerOrder, customerOrderNotDeliveryId, MessageErreur.FILTRE_ERRONE, "IsNotDelivery");

            customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrderDeliveryId);
            customerOrderPage.PageUpCO();
            int customerOrdersNbr = customerOrderPage.CheckTotalNumber();

            //Assert
            Assert.AreEqual(customerOrdersNbr, 0, MessageErreur.FILTRE_ERRONE, "IsNotDelivery");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_Filter_IsAllDelivered()
        {
            // Prepare
            string customerOrderDeliveryId = TestContext.Properties[string.Format("CustomerOrderId5")].ToString();
            string customerOrderNotDeliveryId = TestContext.Properties[string.Format("CustomerOrderId6")].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.ResetFilter();

            customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrderNotDeliveryId);
            customerOrderPage.Filter(CustomerOrderPage.FilterType.IsAllDelivered, true);
            customerOrderPage.PageUpCO();
            string firstCustomerOrder = customerOrderPage.GetCustomerOrderNumber().Split(' ')[1];

            //Assert
            Assert.AreEqual(firstCustomerOrder, customerOrderNotDeliveryId, MessageErreur.FILTRE_ERRONE, "IsAllDelivery");

            customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrderDeliveryId);
            firstCustomerOrder = customerOrderPage.GetCustomerOrderNumber().Split(' ')[1];

            //Assert
            Assert.AreEqual(firstCustomerOrder, customerOrderDeliveryId, MessageErreur.FILTRE_ERRONE, "IsAllDelivery");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_Filter_IsAllVerified()
        {
            // Prepare
            string customerOrderVerifiedId = TestContext.Properties[string.Format("CustomerOrderId8")].ToString();
            string customerOrderNotVerifiedId = TestContext.Properties[string.Format("CustomerOrderId9")].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.ResetFilter();

            customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrderVerifiedId);
            customerOrderPage.Filter(CustomerOrderPage.FilterType.IsAllVerified, true);
            customerOrderPage.PageUpCO();
            string firstCustomerOrder = customerOrderPage.GetCustomerOrderNumber().Split(' ')[1];

            //Assert
            Assert.AreEqual(firstCustomerOrder, customerOrderVerifiedId, MessageErreur.FILTRE_ERRONE, "IsAllVerified");

            customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrderNotVerifiedId);
            firstCustomerOrder = customerOrderPage.GetCustomerOrderNumber().Split(' ')[1];

            //Assert
            Assert.AreEqual(firstCustomerOrder, customerOrderNotVerifiedId, MessageErreur.FILTRE_ERRONE, "IsAllVerified");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_Filter_IsOnlyVerified()
        {
            // Prepare
            string customerOrderVerifiedId = TestContext.Properties[string.Format("CustomerOrderId8")].ToString();
            string customerOrderNotVerifiedId = TestContext.Properties[string.Format("CustomerOrderId9")].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.ResetFilter();

            customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrderVerifiedId);
            customerOrderPage.Filter(CustomerOrderPage.FilterType.IsVerified, true);
            customerOrderPage.PageUpCO();
            string firstCustomerOrder = customerOrderPage.GetCustomerOrderNumber().Split(' ')[1];

            //Assert
            Assert.AreEqual(firstCustomerOrder, customerOrderVerifiedId, MessageErreur.FILTRE_ERRONE, "IsVerified");

            customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrderNotVerifiedId);

            int customerOrdersNbr = customerOrderPage.CheckTotalNumber();

            //Assert
            Assert.AreEqual(customerOrdersNbr, 0, MessageErreur.FILTRE_ERRONE, "IsVerified");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_Filter_IsOnlyNotVerified()
        {
            // Prepare
            string customerOrderVerifiedId = TestContext.Properties[string.Format("CustomerOrderId8")].ToString();
            string customerOrderNotVerifiedId = TestContext.Properties[string.Format("CustomerOrderId9")].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.ResetFilter();

            customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrderNotVerifiedId);
            customerOrderPage.Filter(CustomerOrderPage.FilterType.IsNotVerified, true);
            customerOrderPage.PageUpCO();
            string firstCustomerOrder = customerOrderPage.GetCustomerOrderNumber().Split(' ')[1];

            //Assert
            Assert.AreEqual(firstCustomerOrder, customerOrderNotVerifiedId, MessageErreur.FILTRE_ERRONE, "IsVerified");

            customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrderVerifiedId);

            int customerOrdersNbr = customerOrderPage.CheckTotalNumber();

            //Assert
            Assert.AreEqual(customerOrdersNbr, 0, MessageErreur.FILTRE_ERRONE, "IsVerified");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_Filter_FlightNumber()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.ResetFilter();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.Customers, customerName);
            var customerOrderItem = customerOrderPage.SelectFirstCustomerOrder();
            var generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
            var flightnumber = generalInfo.GetFlightNumber();
            customerOrderPage = customerOrderItem.BackToList();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.FlightNumber, flightnumber);
            var totalCO = customerOrderPage.CheckTotalNumber();
            Assert.AreEqual(1, totalCO, "Filter flight number ne fonctionne pas.");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_Filter_CrewPreOrders()
        {
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.ResetFilter();
            var totalCOAll = customerOrderPage.CheckTotalNumber();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.CrewPreOrder, true);
            var totalCOYes = customerOrderPage.CheckTotalNumber();
            customerOrderPage.ResetFilter();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.CrewPreOrder, false);
            var totalCONo = customerOrderPage.CheckTotalNumber();

            Assert.AreEqual(totalCOAll, totalCOYes + totalCONo, "Filter CrewPreOrders ne fonctionne pas.");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_Filter_OrdersTypes()
        {
            //prepare 
            string name = "Test Order Type";

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.ResetFilter();
            var totalCOAll = customerOrderPage.CheckTotalNumber();
            if (!customerOrderPage.isOrderTypeExist(name))
            {
                ParametersCustomer parametersCustomer = homePage.GoToParameters_CustomerPage();
                parametersCustomer.ClickOrderTypeTab();
                parametersCustomer.AddNewOrderType(name);
                customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            }
            customerOrderPage.ResetFilter();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.OrdersTypes, name);
            var totalCOYes = customerOrderPage.CheckTotalNumber();

            Assert.AreNotEqual(totalCOAll, totalCOYes, "Filter CrewPreOrders ne fonctionne pas.");
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_index_customername()
        {
            string customer = TestContext.Properties["Customer"].ToString();
            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.ResetFilter();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.Customers, customer);
            var firstCustomerNameFromIndex = customerOrderPage.GetFirstCustomer();
            var customerOrderDetailPage = customerOrderPage.SelectFirstCustomerOrder();
            var customerOrderGeneralInformationPage = customerOrderDetailPage.ClickOnGeneralInformationTab();
            var customerNameFromGeneralInformation = customerOrderGeneralInformationPage.GetCustomerName();
            //Assert
            Assert.IsTrue(customerNameFromGeneralInformation.Contains(firstCustomerNameFromIndex), "Le customer name sur la page d'index n'est pas le même que sur la page de détails du customer order");
        }

        [TestMethod]

        [Timeout(Timeout)]
        public void PR_CO_index_delivery()
        {
            HomePage homePage = LogInAsAdmin();


            //Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            var firstCustomerOrderDelivery = customerOrderPage.GetFirstDelivery();
            var customerOrderDetailPage = customerOrderPage.SelectFirstCustomerOrder();
            var customerOrderGeneralInformationPage = customerOrderDetailPage.ClickOnGeneralInformationTab();
            var firstCustomerOrderDeliveryFromDetail = customerOrderGeneralInformationPage.GetFlightNumber();
            Assert.AreEqual(firstCustomerOrderDelivery, firstCustomerOrderDeliveryFromDetail, "delivery sur la page d'index n'est pas le même que sur la page de détails du customer order");

        }
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_Filter_FlightDateFrom()
        {
            // Prepare
            string aircraft = TestContext.Properties["Registration"].ToString();
            string customer = TestContext.Properties["CustomerLPFlight"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            DateTime date1 = DateTime.Today;
            DateTime date2 = DateTime.Today.AddMonths(+3);
            DateTime date3 = DateTime.Today.AddMonths(-3);
            DateTime dateForFilter = DateTime.Today;
            DateTime dateForDelete = DateTime.Today.AddMonths(-4);
            string flightNumber = "FlightToday"+ DateTime.Now.ToString();
            int check = 2;
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            string customerOrderNumber="";
            string customerOrderNumber2="";
            string customerOrderNumber3="";
            try
            {
                // creer customer 1
                var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
                var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder2(site, customer, aircraft,flightNumber,date1);
                var customerOrderItem = customerOrderCreateModalPage.Submit();
                var generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
                customerOrderNumber = generalInfo.GetOrderNumber();
                customerOrderPage = generalInfo.BackToList();
                // creer customer 2
                customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder2(site, customer, aircraft, flightNumber, date2);
                customerOrderItem = customerOrderCreateModalPage.Submit();
                generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
                customerOrderNumber2 = generalInfo.GetOrderNumber();
                customerOrderPage = generalInfo.BackToList();
                // creer customer 3
                customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder2(site, customer, aircraft, flightNumber, date3);
                customerOrderItem = customerOrderCreateModalPage.Submit();
                generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
                customerOrderNumber3 = generalInfo.GetOrderNumber();
                customerOrderPage = generalInfo.BackToList();
                customerOrderPage.WaitPageLoading();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.From, dateForFilter);
                customerOrderPage.Filter(CustomerOrderPage.FilterType.FlightNumber, flightNumber);
                customerOrderPage.WaitPageLoading();
                var count = customerOrderPage.GetCustomerOrdersCount();
                Assert.AreEqual(check, count,"le filtre sur datefrom ne s'applique pas");
            }
            finally
            {
                // delete customers
                CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.From, dateForDelete);
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrderNumber);
                customerOrderPage.DeleteCustomerOrder();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrderNumber2);
                customerOrderPage.DeleteCustomerOrder();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrderNumber3);
                customerOrderPage.DeleteCustomerOrder();
            
            }
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_Filter_FlightDateTomorrow()
        {
            //prepare
            string aircraft = TestContext.Properties["Registration"].ToString();
            string customer = TestContext.Properties["CustomerLPFlight"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            DateTime date1 = DateTime.Today;
            DateTime date2 = DateTime.Today.AddMonths(-5);
            DateTime dateTomorrowForCO = DateTime.Today.AddDays(1);
            string dateFlightTomorow = DateTime.Today.AddDays(1).ToString("dd/MM/yyyy");
            string flightNumber = "FlightToday" + DateTime.Now.ToString();
            var customerOrderNumber = "";
            bool tomorrowFilter = true;
            int verif = 1;
            // arrange
            var homePage = LogInAsAdmin();
            //act
            try
            {
                // creer customer 1
                var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
                var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder2(site, customer, aircraft, flightNumber, date1, null, dateTomorrowForCO);
                var customerOrderItem = customerOrderCreateModalPage.Submit();
                var generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
                customerOrderNumber = generalInfo.GetOrderNumber();
                customerOrderPage = generalInfo.BackToList();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.From, date2);
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Tomorrow, tomorrowFilter);
                bool check = customerOrderPage.verifTomorrowFilter(dateFlightTomorow);
                Assert.IsTrue(check,"le filtre ne s'applique pas correctement sur flight start time et flight endtime");
                customerOrderPage.Filter(CustomerOrderPage.FilterType.FlightNumber, flightNumber);
                int count = customerOrderPage.GetCustomerOrdersCount();
                Assert.AreEqual(count, verif,"le filtre tomorow ne s'applique pas correctement");
            }
            finally
            {
                // delete customer 
                CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.From, date2);
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrderNumber);
                customerOrderPage.DeleteCustomerOrder();
            }
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_Filter_FlightDateTo()
        {
            // Prepare
            string aircraft = TestContext.Properties["Registration"].ToString();
            string customer = TestContext.Properties["CustomerLPFlight"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            DateTime date1 = DateTime.Today;
            DateTime date2 = DateTime.Today.AddMonths(+3);
            DateTime date3 = DateTime.Today.AddMonths(-3);
            DateTime dateForFilterDateTo = DateTime.Today;
            DateTime dateForDelete = DateTime.Today.AddMonths(+4);
            DateTime dateForFilterDateFrom = DateTime.Today.AddMonths(-5);
            string flightNumber = "FlightToday" + DateTime.Now.ToString();
            int check = 2;
            string siteTo = TestContext.Properties["SiteToFlight"].ToString();
            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            string customerOrderNumber = "";
            string customerOrderNumber2 = "";
            string customerOrderNumber3 = "";
            try
            {
                // creer customer 1
                var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
                var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder2(site, customer, aircraft, flightNumber, date1);
                var customerOrderItem = customerOrderCreateModalPage.Submit();
                var generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
                customerOrderNumber = generalInfo.GetOrderNumber();
                customerOrderPage = generalInfo.BackToList();
                // creer customer 2
                customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder2(site, customer, aircraft, flightNumber, date2);
                customerOrderItem = customerOrderCreateModalPage.Submit();
                generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
                customerOrderNumber2 = generalInfo.GetOrderNumber();
                customerOrderPage = generalInfo.BackToList();
                // creer customer 3
                customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder2(site, customer, aircraft, flightNumber, date3);
                customerOrderItem = customerOrderCreateModalPage.Submit();
                generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
                customerOrderNumber3 = generalInfo.GetOrderNumber();
                customerOrderPage = generalInfo.BackToList();
                customerOrderPage.WaitPageLoading();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.From, dateForFilterDateFrom);
                customerOrderPage.Filter(CustomerOrderPage.FilterType.To, dateForFilterDateTo);
                customerOrderPage.Filter(CustomerOrderPage.FilterType.FlightNumber, flightNumber);
                customerOrderPage.WaitPageLoading();
                var count = customerOrderPage.GetCustomerOrdersCount();
                Assert.AreEqual(check, count, "le filtre sur datefrom ne s'applique pas");
            }
            finally
            {
                // delete customers
                CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.From, dateForFilterDateFrom);
                customerOrderPage.Filter(CustomerOrderPage.FilterType.To, dateForDelete);
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrderNumber);
                customerOrderPage.DeleteCustomerOrder();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrderNumber2);
                customerOrderPage.DeleteCustomerOrder();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrderNumber3);
                customerOrderPage.DeleteCustomerOrder();

            }
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_Filter_FlightDateToday()
        {
            //prepare
            string aircraft = TestContext.Properties["Registration"].ToString();
            string customer = TestContext.Properties["CustomerLPFlight"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            DateTime date1 = DateTime.Today;
            DateTime date2 = DateTime.Today.AddMonths(-5);
            DateTime dateToday = DateTime.Today;
            string dateTodayFormat = DateTime.Today.ToString("dd/MM/yyyy");
            string flightNumber = "FlightToday" + DateTime.Now.ToString();
            var customerOrderNumber = "";
            bool todayFilter = true;
            int verif = 1;
            // arrange
            var homePage = LogInAsAdmin();
            //act
            try
            {
                // creer customer 1
                var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
                var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder2(site, customer, aircraft, flightNumber, date1, null, dateToday);
                var customerOrderItem = customerOrderCreateModalPage.Submit();
                var generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
                customerOrderNumber = generalInfo.GetOrderNumber();
                customerOrderPage = generalInfo.BackToList();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.From, date2);
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Today, todayFilter);
                bool check = customerOrderPage.verifTomorrowFilter(dateTodayFormat);
                Assert.IsTrue(check, "le filtre ne s'applique pas correctement sur flight start time et flight endtime");
                customerOrderPage.Filter(CustomerOrderPage.FilterType.FlightNumber, flightNumber);
                int count = customerOrderPage.GetCustomerOrdersCount();
                Assert.AreEqual(count, verif, "le filtre tomorow ne s'applique pas correctement");
            }
            finally
            {
                // delete customer 
                CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.From, date2);
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrderNumber);
                customerOrderPage.DeleteCustomerOrder();
            }
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_messagerie_answer()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string messageTitle = "message title";
            string messageBody = "message body TestAuto";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            //Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            try
            {
                var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
                CustomerOrderItem customerOrderItem = customerOrderCreateModalPage.Submit();
                var generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
                var orderN = generalInfo.GetOrderNumber();
                CustomerOrderIMessagesPage customerOrderIMessagesPage = generalInfo.GoToMessagesTab();
                customerOrderIMessagesPage.AddNewMessage(messageTitle, messageBody);
                bool isMessageAdded = customerOrderIMessagesPage.VerifyMessageDisplayed(messageTitle);
                Assert.IsTrue(isMessageAdded, "Le message n'est pas ajouté.");
                customerOrderPage = customerOrderIMessagesPage.BackToList();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, orderN);
            }
            finally
            {
                customerOrderPage.DeleteCustomerOrder();
            }
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_messagerie_delete()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string messageTitle = "message title";
            string messageBody = "message body TestAuto";
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            //Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            try
            {
                var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
                CustomerOrderItem customerOrderItem = customerOrderCreateModalPage.Submit();
                var generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
                var orderN = generalInfo.GetOrderNumber();
                CustomerOrderIMessagesPage customerOrderIMessagesPage = generalInfo.GoToMessagesTab();
                customerOrderIMessagesPage.AddNewMessage(messageTitle, messageBody);
                Assert.IsTrue(customerOrderIMessagesPage.VerifyMessageDisplayed(messageTitle), "Le message n'est pas ajouté.");

                customerOrderIMessagesPage.DeleteMessage();
                Assert.IsFalse(customerOrderIMessagesPage.VerifyMessageAdded(messageTitle), "Le message n'est pas ajouté.");
                customerOrderPage = customerOrderIMessagesPage.BackToList();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, orderN);
            }
            finally
            {
                customerOrderPage.DeleteCustomerOrder();
            }
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_messagerie_deletemessage()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string messageTitle = "message title";
            string messageBody = "message body TestAuto";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            //Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            try
            {
                var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
                CustomerOrderItem customerOrderItem = customerOrderCreateModalPage.Submit();
                var generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
                var orderN = generalInfo.GetOrderNumber();
                CustomerOrderIMessagesPage customerOrderIMessagesPage = generalInfo.GoToMessagesTab();
                customerOrderIMessagesPage.AddNewMessage(messageTitle, messageBody);
                Assert.IsTrue(customerOrderIMessagesPage.VerifyMessageDisplayed(messageTitle), "Le message n'est pas ajouté.");
                customerOrderIMessagesPage.ViewMessage();
                customerOrderIMessagesPage.DeleteMessageDetail();
                Assert.IsTrue(customerOrderIMessagesPage.VerifyMessageDetailDeleted(), "Le message n'est pas ajouté.");
                customerOrderIMessagesPage.CloseMessageModal();
                customerOrderPage = customerOrderIMessagesPage.BackToList();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, orderN);
            }
            finally
            {
                customerOrderPage.DeleteCustomerOrder();
            }
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_index_notverified()
        {
            HomePage homePage = LogInAsAdmin();

            //Act
            CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.IsVerified, "ONLY VERIFIED");
            CustomerOrderItem co = customerOrderPage.SelectFirstCustomerOrder();
            co.DoNotVerify();
            customerOrderPage = co.BackToList();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.IsVerified, "ONLY NOT VERIFIED");
            var result = customerOrderPage.VerifyAllNotVerified();
            Assert.IsTrue(result, "filtre only not verified not working");
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_index_pagination()
        {
            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.PageSize("8");
            var count = customerOrderPage.GetCustomerOrdersIdList().Count;
            Assert.IsTrue(count <= 8, "Erreur de pagination 8");
            customerOrderPage.PageSize("16");
            count = customerOrderPage.GetCustomerOrdersIdList().Count;
            Assert.IsTrue(count <= 16, "Erreur de pagination 16");
            customerOrderPage.PageSize("30");
            count = customerOrderPage.GetCustomerOrdersIdList().Count;
            Assert.IsTrue(count <= 30, "Erreur de pagination 30");
            customerOrderPage.PageSize("50");
            count = customerOrderPage.GetCustomerOrdersIdList().Count;
            Assert.IsTrue(count <= 50, "Erreur de pagination 50");
            customerOrderPage.PageSize("100");
            count = customerOrderPage.GetCustomerOrdersIdList().Count;
            Assert.IsTrue(count <= 100, "Erreur de pagination 100");
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_index_type()
        {
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            try 
            {
                var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
                CustomerOrderItem customerOrderItem = customerOrderCreateModalPage.Submit();
                var generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
                var orderN = generalInfo.GetOrderNumber();
                generalInfo.BackToList();
                var firstCustomerOrderNumber = customerOrderPage.GetFirstOrderNumber();
                var customerOrderItemPage = customerOrderPage.SelectFirstCustomerOrder();
                var customerOrderGeneralInformationPage = customerOrderItemPage.ClickOnGeneralInformationTab();
                customerOrderPage.SetOrderType(CustomerOrderPage.CustomerOrderType.TestAuto);
                customerOrderPage.GoToProduction_CustomerOrderPage();
                var customerOrderTypeByNumber = customerOrderPage.GetCustomerOrderTypeByNumber(firstCustomerOrderNumber);

                //Assert
                Assert.AreEqual(customerOrderTypeByNumber, "testAuto", "erreur de mise a jour du Customer Order Type");
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, orderN);

            }
            finally
            {

                customerOrderPage.DeleteCustomerOrder();
            }
           
        }

        [TestMethod]


        [Timeout(Timeout)]
        public void PR_CO_index_verified()
        {
            HomePage homePage = LogInAsAdmin();

            //Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.IsVerified, "ONLY NOT VERIFIED");
            CustomerOrderItem co = customerOrderPage.SelectFirstCustomerOrder();
            co.DoVerify();
            customerOrderPage = co.BackToList();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.IsVerified, "ONLY VERIFIED");
            var result = customerOrderPage.VerifyAllVerified();
            Assert.IsTrue(result, "filtre only verified not working");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_messagerie_newmessage()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string messageTitle = "message title";
            string messageBody = "message body TestAuto";

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            //Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            try
            {
                var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
                CustomerOrderItem customerOrderItem = customerOrderCreateModalPage.Submit();
                var generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
                var orderN = generalInfo.GetOrderNumber();
                CustomerOrderIMessagesPage customerOrderIMessagesPage = generalInfo.GoToMessagesTab();
                var date = DateTime.Today.ToString("yyyy-MM-dd");
                var time = customerOrderIMessagesPage.AddNewMessage(messageTitle, messageBody);
                bool isMessageAdded = customerOrderIMessagesPage.VerifyMessageDisplayed(messageTitle);
                bool isMessageDataAdded = customerOrderIMessagesPage.VerifyMessageData(messageTitle, messageBody, date, time);

                Assert.IsTrue(isMessageAdded, "Le message n'est pas ajouté.");
                Assert.IsTrue(isMessageDataAdded, "Le message n'est pas ajouté.");
                customerOrderPage = customerOrderIMessagesPage.BackToList();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, orderN);
            }
            finally
            {
                customerOrderPage.DeleteCustomerOrder();
            }
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_messagerie_view()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string messageTitle = "message title";
            string messageBody = "message body TestAuto";
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            //Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
            customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
            CustomerOrderItem customerOrderItem = customerOrderCreateModalPage.Submit();
            var generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
            var orderN = generalInfo.GetOrderNumber();
            CustomerOrderIMessagesPage customerOrderIMessagesPage = generalInfo.GoToMessagesTab();
            var date = DateTime.Today.ToString("yyyy-MM-dd");
            var time = customerOrderIMessagesPage.AddNewMessage(messageTitle, messageBody);
            Assert.IsTrue(customerOrderIMessagesPage.VerifyMessageDisplayed(messageTitle), "Le message n'est pas ajouté.");
            customerOrderIMessagesPage.ViewMessage();
            Assert.IsTrue(customerOrderIMessagesPage.VerifyViewMessage(), "Le message n'est pas ajouté.");
            customerOrderIMessagesPage.CloseMessageModal();
            customerOrderPage = customerOrderIMessagesPage.BackToList();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, orderN);
            customerOrderPage.DeleteCustomerOrder();
        }
        [DeploymentItem("Resources\\pizza.png")]
        [DeploymentItem("Resources\\pizza_petite.png")]
        [DeploymentItem("chromedriver.exe")]
        [TestMethod]


        [Timeout(Timeout)]
        public void PR_CO_PrintLogo()
        {
            // Prepare
            string site = TestContext.Properties["InvoiceSite"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNameZipBegin = "All_files_";
            string DocFileNamePdfBegin = "Order Report";
            //Arrange
            HomePage homePage = LogInAsAdmin();

            ParametersSites SitesParams = homePage.GoToParameters_Sites();
            SitesParams.Filter(ParametersSites.FilterType.SearchSite, site);
            SitesParams.ClickOnFirstSite();
            //Configurer un logo sur un site dans les globals settings
            //Avoir des données disponibles.
            // serialize + sha1sum (filesize)
            Console.WriteLine(TestContext.TestDeploymentDir + "\\pizza.png");
            FileInfo fiUpload = new FileInfo(TestContext.TestDeploymentDir + "\\pizza.png");
            SitesParams.ClickToInformations();
            SitesParams.UploadLogo(fiUpload);
            homePage.Navigate();
            homePage.ClearDownloads();
            //Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.Site, site);
            var customerOrderItem = customerOrderPage.SelectFirstCustomerOrder();
            var generalInfoCO = customerOrderItem.ClickOnGeneralInformationTab();
            //1. cliquer sur print results
            //2. cliquer sur print
            PrintReportPage reportPage = generalInfoCO.PrintCustomerOrder(true);
            var isReportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();
            //Assert
            Assert.IsTrue(isReportGenerated, "L'application n'a pas pu générer le fichier attendu.");
            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            FileInfo fi = new FileInfo(trouve);
            fi.Refresh();
            Assert.IsTrue(fi.Exists, trouve + " non généré");
            //3.Vérifier le logo sur le PDF
            PdfDocument document = PdfDocument.Open(fi.FullName);
            Page page1 = document.GetPage(1);
            List<IPdfImage> images = page1.GetImages().ToList<IPdfImage>();
            Assert.AreEqual(1, images.Count, "Pas d'image dans le PDF");
            IPdfImage image = images[0];
            FileInfo fiPdf = new FileInfo(downloadsPath + "\\pizza_test.png");
            if (fiPdf.Exists)
            {
                fiPdf.Delete();
            }
            File.WriteAllBytes(downloadsPath + "\\pizza_test.png", image.RawBytes.ToArray<byte>());
            fiPdf.Refresh();
            Assert.IsTrue(fiPdf.Exists, "fichier non trouvé");
            FileInfo fiTest;
            fiTest = new FileInfo(TestContext.TestDeploymentDir + "\\pizza_petite.png");
            Assert.IsTrue(fiTest.Exists, "fichier non trouvé (cas 2)");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_Print_CustomerOrder()
        {
            bool newVersionPrint = true;

            //Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.ClearDownloads();
            var customerOrderItem = customerOrderPage.SelectFirstCustomerOrder();
            var reportPage = customerOrderItem.PrintProdCustomerOrderWithTypeOfReport(newVersionPrint);
            var isReportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();
            //Assert
            Assert.IsTrue(isReportGenerated, "Le document PDF n'a pas pu être généré par l'application.");
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_DeletewithMessage()
        {

            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string messageTitle = "message title";
            string messageBody = "message body TestAuto";

            // Arrange
            var homePage = LogInAsAdmin();
            // Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            try
            {
                var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
                CustomerOrderItem customerOrderItem = customerOrderCreateModalPage.Submit();
                var generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
                var orderN = generalInfo.GetOrderNumber();
                CustomerOrderIMessagesPage customerOrderIMessagesPage = generalInfo.GoToMessagesTab();
                var date = DateTime.Today.ToString("yyyy-MM-dd");
                var time = customerOrderIMessagesPage.AddNewMessage(messageTitle, messageBody);
                bool isMessageAdded = customerOrderIMessagesPage.VerifyMessageAdded(messageTitle);
                bool isMessageDataAdded = customerOrderIMessagesPage.VerifyMessageData(messageTitle, messageBody, date, time);
                Assert.IsTrue(isMessageAdded, "Le message n'est pas ajouté.");
                Assert.IsTrue(isMessageDataAdded, "Le message n'est pas ajouté.");
                customerOrderPage = customerOrderIMessagesPage.BackToList();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, orderN);
            }
            finally
            {
                customerOrderPage.DeleteCustomerOrder();
                int total_number = customerOrderPage.CheckTotalNumber_customerOrders();
                Assert.IsTrue(total_number == 0, "le customer order est existant ");
            }
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_SendyBmailPJ()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNameZipBegin = "All_files_";
            string DocFileNamePdfBegin = "Order Report";
            var email = "z.hamdi.ext@newrest.eu";

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            DeleteAllFileDownload();
            // Act
            CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            CustomerOrderItem customerOrderItem = customerOrderPage.SelectFirstCustomerOrder();
            CustomerOrderGeneralInformationPage generalInfoCO = customerOrderItem.ClickOnGeneralInformationTab();
            //1. cliquer sur print results
            //2. cliquer sur print
            PrintReportPage reportPage = generalInfoCO.PrintCustomerOrder(true);
            var isReportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();
            Assert.IsTrue(isReportGenerated, "L'application n'a pas pu générer le fichier attendu.");
            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            customerOrderItem.SendEmailPopUp();
            var numberBeforeAddingFile = customerOrderItem.GetNumberAttachementFile();
            customerOrderItem.SetEmailTo(email);
            customerOrderItem.ChooseFile(trouve);
            var numberAfterAddingFile = customerOrderItem.GetNumberAttachementFile();
            Assert.AreEqual(numberBeforeAddingFile + 1, numberAfterAddingFile, "Aucun fichier n'a été ajouté");
            customerOrderItem.SendEmailButton();

        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_CurrencyDisplay()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNameZipBegin = "All_files_";
            string DocFileNamePdfBegin = "Order Report";

            Random rnd = new Random();
            string customerType = TestContext.Properties["CustomerType"].ToString();
            string customerCurrency = "$ Dolar Americano";

            string customerName = "CustomerTestForProdCO";
            string customerCode = rnd.Next().ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string itemName = TestContext.Properties["InvoiceService"].ToString();
            bool newVersionPrint = true;


            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ClearDownloads();
            customersPage.Filter(FilterType.Search, customerName);

            // Creation Customer with Currency different €

            if (customersPage.CheckTotalNumber_display() == 0)
            {
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFieldsWithCurrency_CreateCustomerModalPage(customerName, customerCode, customerType, customerCurrency);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                var currency = customerGeneralInformationsPage.GetDevise();

                Assert.AreEqual(currency, customerCurrency, "Currency n'est pas en $.");

                customersPage = customerGeneralInformationsPage.BackToList();

                customersPage.ResetFilters();
                customersPage.Filter(FilterType.Search, customerName);

                Assert.AreEqual(customerName, customersPage.GetFirstCustomerName(), "Le customer n'a pas été créé.");
            }

            // Creation Customer Order
            CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
            customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
            var customerOrderDetailPage = customerOrderCreateModalPage.Submit();
            CustomerOrderGeneralInformationPage customerOrderGeneralInformationPage = customerOrderDetailPage.ClickOnGeneralInformationTab();
            var orderNumber = customerOrderGeneralInformationPage.GetOrderNumber();
            customerOrderDetailPage.BackToList();

            //Assert
            customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, orderNumber);
            var currencyPrice = customerOrderPage.GetFirstPriceCurrency();
            Assert.AreEqual(currencyPrice, "$", "Currency n'est pas en $.");

            CustomerOrderItem customerOrderItem = customerOrderPage.SelectFirstCustomerOrder();
            customerOrderItem.AddNewItem(itemName, "3");
            var PriceItemCurrency = customerOrderItem.GetItemPriceCurrency();
            Assert.AreEqual(PriceItemCurrency, "$", "Currency n'est pas en $.");

            var reportPage = customerOrderItem.PrintCustomerOrder(newVersionPrint);
            var isReportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();

            //Assert
            Assert.IsTrue(isReportGenerated, "Le document PDF n'a pas pu être généré par l'application.");

            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            string generatedFilePath = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);

            FileInfo fileInfo = new FileInfo(generatedFilePath);
            fileInfo.Refresh();
            Assert.IsTrue(fileInfo.Exists, $"{generatedFilePath} non généré");

            PdfDocument document = PdfDocument.Open(fileInfo.FullName);
            List<string> words = new List<string>();
            foreach (Page page in document.GetPages())
            {
                words.AddRange(page.GetWords().Select(word => word.Text));
            }

            bool CurrencyExistance = words.Any(word => word.Contains("$"));
            //Assert
            Assert.IsTrue(CurrencyExistance, "Currency $ n'a pas été trouvé dans le PDF.");

        }
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_ResearchPricelistItem()
        {
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string itemName = TestContext.Properties["InvoiceService"].ToString();
            string firstTwoLetters = new string(itemName.Take(2).ToArray());
            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
            customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
            CustomerOrderItem customerOrderItem = customerOrderCreateModalPage.Submit();
            customerOrderItem.BackToList();
            CustomerOrderItem customerOrderItemPage = customerOrderPage.SelectFirstCustomerOrder();
            customerOrderItemPage.SetNameItem(firstTwoLetters);
            var menuNameList = customerOrderItemPage.GetSuggestionsMenuNameList();

            // Assert
            Assert.IsTrue(menuNameList.All(name => name.StartsWith(firstTwoLetters)), "Dès que vous tapez les deux premières lettres de l'Item voulu, la liste affichée n'est pas bien optimisée.");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_ClosePopUpMessage()
        {
            string title = "title test";
            string body = "body test";

            var homePage = LogInAsAdmin();
            CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            CustomerOrderItem customerOrderItem = customerOrderPage.SelectFirstCustomerOrder();
            CustomerOrderIMessagesPage customerOrderIMessagesPage = customerOrderItem.GoToMessagesTab();
            customerOrderIMessagesPage.AddNewMessage(title, body);
            var messageClosedSuccesfully = customerOrderIMessagesPage.CheckIfMessageClosedSuccesfully();
            Assert.IsTrue(messageClosedSuccesfully, "close button not working right");


        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_FilterByDelivery()
        {
            // Prepare
            string customer = "a";
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            CustomerOrderPage customerOrderPage = null;

            try
            {
                // Act
                customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
                var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrderWOUTflight(site, customer, "123456789", DateTime.Today);
                var customerOrderDetailPage = customerOrderCreateModalPage.Submit();

                customerOrderDetailPage.BackToList();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, "123456789");
                var countlines = customerOrderDetailPage.CheckTotalNumber();

                // Assert
                Assert.IsTrue(countlines > 0, "Expected the count of lines to be greater than 0, but it was " + countlines);
            }
            finally
            {
                if (customerOrderPage != null)
                {
                    customerOrderPage.DeleteCustomerOrder();
                }
            }
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_OrdersFlightInformations()
        {
            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            CustomerOrderPage customerOrderPage = null;

            // Act
            customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
            customerOrderCreateModalPage.FillField_CreatNewCustomerOrderWOUTflight(site, customer, "123456789", DateTime.Today);
            var customerOrderDetailPage = customerOrderCreateModalPage.Submit();
            bool Result = customerOrderCreateModalPage.VerifierLesMessages();
            Assert.IsTrue(Result, "les messages n'apparaissent pas");



        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_DisplayApplyVAT()
        {
            // Prepare
            string customer = "ABC BEDARFSFLUG";
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            CustomerOrderPage customerOrderPage = null;

            // Act
            customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
            customerOrderCreateModalPage.FillField_CreatNewCustomerOrderWOUTflight(site, customer, "123456789", DateTime.Today);
            var customerOrderDetailPage = customerOrderCreateModalPage.Submit();
            bool Result = customerOrderCreateModalPage.verifierApplyVAT();
            Assert.IsTrue(Result, "Apply VAT n'apparait pas");



        }
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_CreateFrom()
        {
            // Préparer les données
            string CustomerOrderCustomer = "Customer Production CO";
            string CustomerOrderSite = TestContext.Properties["SiteCPU"].ToString();

            HomePage homePage = LogInAsAdmin();

            CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();

            var customerOrderCreateModalPage = customerOrderPage.CreateFromCreatePage();

            // Remplir les champs pour la création de la commande client
            customerOrderCreateModalPage.FillField_CreatFromCustomerOrder(CustomerOrderSite, CustomerOrderCustomer);

            // Utiliser l'option "Copy items from another order"
            customerOrderCreateModalPage.CopyItemsFromAnotherOrder();

            // Cocher la première ligne des résultats
            customerOrderCreateModalPage.SelectFirstOrderToDuplicate();

            // Cliquer sur le bouton "Duplicate"
            var duplicatedOrder = customerOrderCreateModalPage.Submit();

            Assert.IsTrue(duplicatedOrder.IsOnCorrectPage(), "La page après duplication n'est pas correcte.");


        }
        [TestMethod]


        [Timeout(Timeout)]
        public void PR_CO_LabelSave()
        {
            HomePage homePage = LogInAsAdmin();


            // Générer une chaîne unique basée sur la date et l'heure actuelles
            string currentDateTime = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string label1 = $"label1test_{currentDateTime}";
            string label2 = $"label2test_{currentDateTime}";
            string label3 = $"label3test_{currentDateTime}";
            string label4 = $"label4test_{currentDateTime}";
            string label5 = $"label5test_{currentDateTime}";

            // Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            var firstlinename = customerOrderPage.GetFirstLineName();
            var customerOrderDetailPage = customerOrderPage.SelectFirstCustomerOrder();
            customerOrderDetailPage.ClickOnGeneralInformationTab();
            customerOrderDetailPage.SetLabels(label1, label2, label3, label4, label5);
            customerOrderDetailPage.BackToList();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, firstlinename);
            customerOrderPage.SelectFirstCustomerOrder();
            customerOrderDetailPage.ClickOnGeneralInformationTab();

            // Récupérer les labels après l'édition
            var label1AfterEdit = customerOrderDetailPage.GetLabel1();
            var label2AfterEdit = customerOrderDetailPage.GetLabel2();
            var label3AfterEdit = customerOrderDetailPage.GetLabel3();
            var label4AfterEdit = customerOrderDetailPage.GetLabel4();
            var label5AfterEdit = customerOrderDetailPage.GetLabel5();

            // Vérifier que les labels sont correctement enregistrés
            Assert.AreEqual(label1, label1AfterEdit, "Label1 ne correspond pas après l'édition");
            Assert.AreEqual(label2, label2AfterEdit, "Label2 ne correspond pas après l'édition");
            Assert.AreEqual(label3, label3AfterEdit, "Label3 ne correspond pas après l'édition");
            Assert.AreEqual(label4, label4AfterEdit, "Label4 ne correspond pas après l'édition");
            Assert.AreEqual(label5, label5AfterEdit, "Label5 ne correspond pas après l'édition");
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_CheckServiceForCustomerOrder()
        {
            //prepare
            string customer = TestContext.Properties["CustomerLPFlight"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string itemName = TestContext.Properties["InvoiceService"].ToString();
            string servicename = "Test" + new Random().Next();

            //arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            try
            {
                var servicesPage = homePage.GoToCustomers_ServicePage();
                var ServiceCreateModalPage = servicesPage.ServiceCreatePage();
                ServiceCreateModalPage.FillFields_CreateServiceModalPageWithDesactivatedMode(servicename);
                var serviceGeneralInformationsPage = ServiceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(+30));

                var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
                var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customer, aircraft);
                var customerOrderDetailPage = customerOrderCreateModalPage.Submit();
                customerOrderDetailPage.AddNewItem("Free Price", "10");
                string FP = customerOrderDetailPage.CheckFreePriceAndS();
                Assert.AreEqual("FP", FP, "le type n'est pas FP");
                customerOrderDetailPage.AddNewItem(servicename, "10");
                customerOrderDetailPage.WaitLoading();
                var customerNumber = customerOrderDetailPage.GetNumberCustomer();
                homePage.GoToProduction_CustomerOrderPage();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerNumber);
                customerOrderPage.ClickFirstLine();
                string S = customerOrderDetailPage.CheckFreePriceAndS();
                Assert.AreEqual("S", S, "le type n'est pas S");
            }
            finally
            {
                var servicesPage = homePage.GoToCustomers_ServicePage();
                var ServiceMassiveDeleteModalPage = servicesPage.ClickMassiveDelete();
                ServiceMassiveDeleteModalPage.Filter(ServiceMassiveDeleteModalPage.ServiceMassiveDeleteFilterType.ServiceName, servicename);
                ServiceMassiveDeleteModalPage.ClickSearchButton();
                ServiceMassiveDeleteModalPage.DeleteFirstService();
            }

        }
        [TestMethod]

        [Timeout(Timeout)]
        public void PR_CO_CommentVATModification()
        {
            //prepare
            string customer = "customer" + new Random().Next();
            string customerIcao = "ICAO" + "-" + new Random().Next().ToString();
            string customerType = "Ash";
            string site = TestContext.Properties["SiteACE"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string itemName = "[Free Price]";
            string Comment = "test";
            string qty = "0";
            //arrange
            HomePage homePage = LogInAsAdmin();

            CustomerPage customerPage = homePage.GoToCustomers_CustomerPage();
            CustomerCreateModalPage customerCreateModalPage = customerPage.CustomerCreatePage();
            customerCreateModalPage.FillFields_CreateCustomerModalPage(customer, customerIcao, customerType);
            CustomerGeneralInformationPage customerGeneralInformationsPage = customerCreateModalPage.Create();
            customerGeneralInformationsPage.SetVATfree(true);
            CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            CustomerOrderCreateModalPage customerOrderCreat = customerOrderPage.CustomerOrderCreatePage();
            customerOrderCreat.FillField_CreatNewCustomerOrder(site, customer, aircraft, DateTime.Today.AddDays(40));
            CustomerOrderItem customerOrderItem = customerOrderCreat.Submit();
            customerOrderItem.AddNewItem(itemName, qty);
            CustomerOrderGeneralInformationPage customerOrderGeneralInfo = customerOrderItem.ClickOnGeneralInformationTab();
            customerOrderGeneralInfo.SetComment(Comment);
            customerOrderGeneralInfo.ClickOnDetailTab();
            bool IsVAYexl = customerOrderItem.IsVATexcl();
            Assert.IsTrue(IsVAYexl, "VAT is not Excl.");

        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_SendMailCommercialManager()
        {
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string newCommercialManagerEmail = "testauto.newrest@gmail.com";


            HomePage homePage = LogInAsAdmin();


            var settingsPage = homePage.GoToApplicationSettings();
            var parametreSitesPage = settingsPage.GoToParameters_Sites();
            parametreSitesPage.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, site);
            parametreSitesPage.ClickOnFirstSite();
            parametreSitesPage.ClickToContacts();

            var currentCommercialManagerEmail = parametreSitesPage.GetCommercialManagerMail();
            parametreSitesPage.EnsureCommercialManagerEmail(currentCommercialManagerEmail, newCommercialManagerEmail);


            var commercialManagerMail = parametreSitesPage.GetCommercialManagerMail();

            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
            customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
            var customerOrderItem = customerOrderCreateModalPage.Submit();
            customerOrderItem.SendEmailPopUp();

            bool checkCommercialManagerEmail = customerOrderItem.CheckCommercialManagerEmail(commercialManagerMail);
            Assert.IsTrue(checkCommercialManagerEmail, $"L'email du Commercial Manager '{commercialManagerMail}' n'est pas valide.");
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_ExportOrderType()
        {
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            HomePage homePage = LogInAsAdmin();

            var settingsPage = homePage.GoToApplicationSettings();
            var parametreSitesPage = settingsPage.GoToParameters_Sites();
            parametreSitesPage.ClickOnCustomer();
            parametreSitesPage.ClickOnOrderType();
            bool Verif = parametreSitesPage.VerifOrderType("Just Eat");
            if (Verif == false)
            {
                ParametersCustomer parametersCustomer = homePage.GoToParameters_CustomerPage();
                parametersCustomer.ClickOrderTypeTab();
                parametersCustomer.AddNewOrderType("Just Eat");
            }
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
            customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customer, aircraft, DateTime.Today, "Just Eat");
            CustomerOrderItem customerOrderItemPage = customerOrderCreateModalPage.Submit();
            customerOrderItemPage.AddNewItem("ItemForOrderType", new Random().Next(5, 20).ToString());
            var generalInfo = customerOrderItemPage.ClickOnGeneralInformationTab();
            var customerOrderNumber = generalInfo.GetOrderNumber();
            customerOrderItemPage.BackToList();
            customerOrderPage.ClearDownloads();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrderNumber);
            customerOrderPage.ExportExcel(true);
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = customerOrderPage.GetExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

            // Récupération du nom du fichier et construction de l'URL du fichier Excel à ouvrir   
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber(CUSTOMER_ORDER, filePath);
            bool result = OpenXmlExcel.ReadAllDataInColumn("Order Type", CUSTOMER_ORDER, filePath, "Just Eat");

            //Assert
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.IsTrue(result, String.Format(MessageErreur.FILTRE_ERRONE, "Order Type Just Eat"));
        }
        [TestMethod]

        [Timeout(Timeout)]
        public void PR_CO_ExportDonnees()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            bool newVersionPrint = true;

            //arrange
            HomePage homePage = LogInAsAdmin();

            //act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            var firstCustomerOrderNumber = customerOrderPage.GetFirstCustomerOrderNumber();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, firstCustomerOrderNumber);
            customerOrderPage.ClearDownloads();
            customerOrderPage.ExportExcel(newVersionPrint);

            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            var correctDownloadedFile = customerOrderPage.GetExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            int resultNumber = OpenXmlExcel.GetExportResultNumber(CUSTOMER_ORDER, filePath);
            bool result = OpenXmlExcel.ReadAllDataInColumn("Sheet 1", CUSTOMER_ORDER, filePath, "InProgress");

            //Assert
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);

        }
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_RechercheFiltre()
        {
            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.ResetFilter();
            var orderNumb = customerOrderPage.GetFirstCustomerOrderNumber();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, orderNumb);
            int nbTotal = customerOrderPage.CheckTotalNumber();
            Assert.AreEqual(1, nbTotal, "la liste doit afficher au moins un Customer Order");
        }
        [Timeout(Timeout)]
        [TestMethod]
        public void PR_CO_GeneralInfo_EditExpeditionDate()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            DateTime date = DateUtils.Now.AddDays(10);
            string orderN = "";

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            try
            {
                CustomerOrderCreateModalPage customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
                CustomerOrderItem customerOrderItem = customerOrderCreateModalPage.Submit();
                CustomerOrderGeneralInformationPage generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
                string expeditionDate = generalInfo.GetExpeditionDate();
                string formattedDate = date.ToString("dd/MM/yyyy");
                generalInfo.UpdateExpeditionDate(date);
                //Assert 
                Assert.AreNotEqual(expeditionDate, formattedDate, "La valeur du champ 'Expedition Date' doit etre different.");

                // On change de menu et on revient
                generalInfo.ClickOnDetailTab();
                generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
                orderN = generalInfo.GetOrderNumber();
                // Récupération des valeurs
                expeditionDate = generalInfo.GetExpeditionDate();

                //Assert
                Assert.AreEqual(expeditionDate, formattedDate, "La valeur du champ 'Expedition Date' n'a pas été modifiée.");
            }
            finally
            {
                customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
                customerOrderPage.ResetFilter();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, orderN);
                customerOrderPage.WaitPageLoading();
                customerOrderPage.DeleteCustomerOrder();
            }
        }

        [Timeout(Timeout)]
        [TestMethod]
        public void PR_CO_GeneralInfo_EditFlightNumber()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();

            // Act
            HomePage homePage = LogInAsAdmin();
            CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            try
            {
                CustomerOrderCreateModalPage customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
                string updateNumber = (customerOrderCreateModalPage.GetFlightNumber() + 3);
                CustomerOrderItem customerOrderItem = customerOrderCreateModalPage.Submit();
                CustomerOrderGeneralInformationPage generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
                generalInfo.UpdateFlightNumber(updateNumber);

                // On change de menu et on revient
                generalInfo.PageUp();
                generalInfo.ClickOnDetailTab();
                generalInfo = customerOrderItem.ClickOnGeneralInformationTab();

                // Récupération des valeurs
                string flightNumber = generalInfo.GetFlightNumber();
                generalInfo.BackToList();
                //Assert
                Assert.AreEqual(flightNumber, updateNumber, "La valeur du champ 'Flight Number' n'a pas été modifiée.");
            }
            finally
            {
                customerOrderPage.DeleteCustomerOrder();
            }
        }

        [Timeout(Timeout)]
        [TestMethod]
        public void PR_CO_GeneralInfo_EditFlightDate()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            DateTime date = DateUtils.Now.AddDays(10);

            // Act
            HomePage homePage = LogInAsAdmin();
            CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            try
            {
                CustomerOrderCreateModalPage customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
                CustomerOrderItem customerOrderItem = customerOrderCreateModalPage.Submit();
                CustomerOrderGeneralInformationPage generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
                string flightDate = generalInfo.GetFlightDate();
                generalInfo.UpdateFlightDate(date);
                Assert.AreNotEqual(flightDate, date.ToString("d"), "La valeur du champ 'Flight Date' n'a pas été modifiée.");

                // On change de menu et on revient
                generalInfo.ClickOnDetailTab();
                generalInfo = customerOrderItem.ClickOnGeneralInformationTab();

                // Récupération des valeurs
                flightDate = generalInfo.GetFlightDate();
                generalInfo.BackToList();

                //Assert
                Assert.AreEqual(flightDate, date.ToString("d"), "La valeur du champ 'Flight Date' n'a pas été modifiée.");
            }
            finally
            {
                customerOrderPage.DeleteCustomerOrder();
            }
        }

        [Timeout(Timeout)]
        [TestMethod]
        public void PR_CO_GeneralInfo_EditHour()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string hour = DateUtils.Now.ToString("HH:mm");
            // Act
            HomePage homePage = LogInAsAdmin();
            CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            try
            {
                CustomerOrderCreateModalPage customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
                CustomerOrderItem customerOrderItem = customerOrderCreateModalPage.Submit();
                CustomerOrderGeneralInformationPage generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
                string hourFromInfo = generalInfo.GetHour();
                generalInfo.UpdateHour(hour);
                Assert.AreNotEqual(hour, hourFromInfo, "La valeur du champ 'Hour' doit etre differente.");

                // On change de menu et on revient
                generalInfo.ClickOnDetailTab();
                generalInfo = customerOrderItem.ClickOnGeneralInformationTab();

                // Récupération des valeurs
                hourFromInfo = generalInfo.GetHour();
                generalInfo.BackToList();

                //Assert
                Assert.AreEqual(hour, hourFromInfo, "La valeur du champ 'Hour' n'a pas été modifiée.");
            }
            finally
            {
                customerOrderPage.DeleteCustomerOrder();
            }
        }

        [Timeout(Timeout)]
        [TestMethod]
        public void PR_CO_GeneralInfo_EditStatut()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string status = "Closed";
            // Act
            HomePage homePage = LogInAsAdmin();
            CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            try
            {
                CustomerOrderCreateModalPage customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
                CustomerOrderItem customerOrderItem = customerOrderCreateModalPage.Submit();
                CustomerOrderGeneralInformationPage generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
                string statusValue = generalInfo.GetStatus();
                generalInfo.ChangeStatus(status);
                generalInfo.PageUp();
                Assert.AreNotEqual(status, statusValue, "La valeur du champ 'Status' doit etre differente.");
                // On change de menu et on revient
                generalInfo.ClickOnDetailTab();
                generalInfo = customerOrderItem.ClickOnGeneralInformationTab();

                // Récupération des valeurs
                statusValue = generalInfo.GetStatus();
                generalInfo.BackToList();

                //Assert
                Assert.AreEqual(status, statusValue, "La valeur du champ 'Status' n'a pas été modifiée.");
            }
            finally
            {
                customerOrderPage.DeleteCustomerOrder();
            }

        }

        [Timeout(Timeout)]
        [TestMethod]
        public void PR_CO_GeneralInfo_EditAircraft()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string aircraftBis = TestContext.Properties["AircraftBis"].ToString();
            // Act
            HomePage homePage = LogInAsAdmin();
            CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            try
            {
                CustomerOrderCreateModalPage customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
                CustomerOrderItem customerOrderItem = customerOrderCreateModalPage.Submit();
                CustomerOrderGeneralInformationPage generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
                generalInfo.UpdateAircraft(aircraftBis);
                Assert.AreNotEqual(aircraft, aircraftBis, "La valeur du champ 'Aircraft' n'a pas été modifiée.");

                // On change de menu et on revient
                generalInfo.PageUp();
                generalInfo.ClickOnDetailTab();
                generalInfo = customerOrderItem.ClickOnGeneralInformationTab();

                // Récupération des valeurs
                bool isAircraftUpdeted = generalInfo.GetAirCraft(aircraftBis);
                generalInfo.BackToList();

                //Assert
                Assert.IsTrue(isAircraftUpdeted, "La valeur du champ 'Aircraft' n'a pas été modifiée.");
            }
            finally
            {
                customerOrderPage.DeleteCustomerOrder();
            }
        }

        [Timeout(Timeout)]
        [TestMethod]
        public void PR_CO_GeneralInfo_EditDestination()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string airportDest = "CBG";
            // Act
            HomePage homePage = LogInAsAdmin();
            CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            try
            {
                CustomerOrderCreateModalPage customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
                CustomerOrderItem customerOrderItem = customerOrderCreateModalPage.Submit();
                CustomerOrderGeneralInformationPage generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
                bool isAirportDest = generalInfo.GetAirportDest(airportDest);
                Assert.IsFalse(isAirportDest, "La valeur du champ 'Destination' doit etre différente.");
                generalInfo.UpdateAirportDest(airportDest);

                // On change de menu et on revient
                generalInfo.PageUp();
                generalInfo.ClickOnDetailTab();
                generalInfo = customerOrderItem.ClickOnGeneralInformationTab();

                // Récupération des valeurs
                isAirportDest = generalInfo.GetAirportDest(airportDest);
                generalInfo.BackToList();

                //Assert
                Assert.IsTrue(isAirportDest, "La valeur du champ 'Destination' n'a pas été modifiée.");
            }
            finally
            {
                customerOrderPage.DeleteCustomerOrder();
            }
        }

        [Timeout(Timeout)]
        [TestMethod]
        public void PR_CO_GeneralInfo_EditETD()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string etd = DateUtils.Now.ToString("HH:mm");
            // Act
            HomePage homePage = LogInAsAdmin();
            CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            try
            {
                CustomerOrderCreateModalPage customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
                CustomerOrderItem customerOrderItem = customerOrderCreateModalPage.Submit();
                CustomerOrderGeneralInformationPage generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
                string fromGeneralInfoEtd = generalInfo.GetETDValue();
                Assert.AreNotEqual(etd, fromGeneralInfoEtd, "La valeur du champ 'ETD' doit etre différente.");
                generalInfo.UpdateEtd(etd);

                // On change de menu et on revient
                generalInfo.PageUp();
                generalInfo.ClickOnDetailTab();
                generalInfo = customerOrderItem.ClickOnGeneralInformationTab();

                // Récupération des valeurs
                fromGeneralInfoEtd = generalInfo.GetETDValue();
                generalInfo.BackToList();

                //Assert
                Assert.AreEqual(etd, fromGeneralInfoEtd, "La valeur du champ 'ETD' n'a pas été modifiée.");
            }
            finally
            {
                customerOrderPage.DeleteCustomerOrder();
            }
        }

        [Timeout(Timeout)]
        [TestMethod]
        public void PR_CO_GeneralInfo_EditOrderType()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string orderType = "orderTest";
            // Act
            HomePage homePage = LogInAsAdmin();
            CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            try
            {
                CustomerOrderCreateModalPage customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
                CustomerOrderItem customerOrderItem = customerOrderCreateModalPage.Submit();
                CustomerOrderGeneralInformationPage generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
                bool isOrderType = generalInfo.GetOrderType(orderType);
                generalInfo.UpdateOrderType(orderType);
                Assert.IsFalse(isOrderType, "La valeur du champ 'Order Type' n'a pas été modifiée.");

                // On change de menu et on revient
                generalInfo.PageUp();
                generalInfo.ClickOnDetailTab();
                generalInfo = customerOrderItem.ClickOnGeneralInformationTab();

                // Récupération des valeurs
                isOrderType = generalInfo.GetOrderType(orderType);
                generalInfo.BackToList();

                //Assert
                Assert.IsTrue(isOrderType, "La valeur du champ 'Order Type' n'a pas été modifiée.");
            }
            finally
            {
                customerOrderPage.DeleteCustomerOrder();
            }
        }

        [Timeout(Timeout)]
        [TestMethod]
        public void PR_CO_GeneralInfo_EditOrigin()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string origin = "ABB";
            // Act
            HomePage homePage = LogInAsAdmin();
            CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            try
            {
                CustomerOrderCreateModalPage customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
                CustomerOrderItem customerOrderItem = customerOrderCreateModalPage.Submit();
                CustomerOrderGeneralInformationPage generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
                bool isOrigin = generalInfo.GetOrigin(origin);
                generalInfo.UpdateOrigin(origin);
                Assert.IsFalse(isOrigin, "La valeur du champ 'Origin' n'a pas été modifiée.");

                // On change de menu et on revient
                generalInfo.PageUp();
                generalInfo.ClickOnDetailTab();
                generalInfo = customerOrderItem.ClickOnGeneralInformationTab();

                // Récupération des valeurs
                isOrigin = generalInfo.GetOrigin(origin);
                generalInfo.BackToList();

                //Assert
                Assert.IsTrue(isOrigin, "La valeur du champ 'Origin' n'a pas été modifiée.");
            }
            finally
            {
                customerOrderPage.DeleteCustomerOrder();
            }
        }

        [Timeout(Timeout)]
        [TestMethod]
        public void PR_CO_GeneralInfo_EditRegType()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string registrationType = "New registration type";
            // Act
            HomePage homePage = LogInAsAdmin();
            CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            try
            {
                CustomerOrderCreateModalPage customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
                CustomerOrderItem customerOrderItem = customerOrderCreateModalPage.Submit();
                CustomerOrderGeneralInformationPage generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
                generalInfo.UpdateRegistrationType(registrationType);

                // On change de menu et on revient
                generalInfo.PageUp();
                generalInfo.ClickOnDetailTab();
                generalInfo = customerOrderItem.ClickOnGeneralInformationTab();

                // Récupération des valeurs
                string updatedRegistrationType = generalInfo.GetRegistrationType();
                generalInfo.BackToList();

                //Assert
                Assert.AreEqual(registrationType, updatedRegistrationType, "La valeur du champ 'Reg. Type' n'a pas été modifiée.");
            }
            finally
            {
                customerOrderPage.DeleteCustomerOrder();
            }
        }

        [Timeout(Timeout)]
        [TestMethod]
        public void PR_CO_GeneralInfo_EditSwitchLocalForeign()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            // Act
            HomePage homePage = LogInAsAdmin();
            CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            try
            {
                CustomerOrderCreateModalPage customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
                CustomerOrderItem customerOrderItem = customerOrderCreateModalPage.Submit();
                CustomerOrderGeneralInformationPage generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
                string localForeign = generalInfo.GetLocalForeignValue();
                customerOrderItem = generalInfo.UpdateLocalForeign();
                generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
                string localForeign2 = generalInfo.GetLocalForeignValue();
                Assert.AreNotEqual(localForeign, localForeign2, "La valeur du champ 'Local / Foreign name ?' n'a pas été modifiée.");
                customerOrderItem = generalInfo.UpdateLocalForeign();
                generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
                string localForeign3 = generalInfo.GetLocalForeignValue();
                Assert.AreNotEqual(localForeign3, localForeign2, "La valeur du champ 'Local / Foreign name ?' n'a pas été modifiée.");
                generalInfo.BackToList();
            }
            finally
            {
                customerOrderPage.DeleteCustomerOrder();
            }
        }

        [Timeout(Timeout)]
        [TestMethod]
        public void PR_CO_GeneralInfo_AddComment()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string comment = "A comment";
            // Act
            HomePage homePage = LogInAsAdmin();
            CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            try
            {
                CustomerOrderCreateModalPage customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
                CustomerOrderItem customerOrderItem = customerOrderCreateModalPage.Submit();
                CustomerOrderGeneralInformationPage generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
                string updatedComment = generalInfo.GetComment();
                generalInfo.SetComment(comment);
                Assert.AreNotEqual(comment, updatedComment, "La valeur du champ 'Comment' n'a pas été modifiée.");

                // On change de menu et on revient
                generalInfo.ClickOnDetailTab();
                generalInfo = customerOrderItem.ClickOnGeneralInformationTab();

                // Récupération des valeurs
                updatedComment = generalInfo.GetComment();
                generalInfo.BackToList();

                //Assert
                Assert.AreEqual(comment, updatedComment, "La valeur du champ 'Comment' n'a pas été modifiée.");
            }
            finally
            {
                customerOrderPage.DeleteCustomerOrder();
            }
        }

        [Timeout(Timeout)]
        [TestMethod]
        public void PR_CO_GeneralInfo_AddLabels()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            Random random = new Random();
            string label = "label" + random.Next(1, 6).ToString();
            string labelText = "Some text for LABEL";
            // Act
            HomePage homePage = LogInAsAdmin();
            CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            try
            {
                CustomerOrderCreateModalPage customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
                CustomerOrderItem customerOrderItem = customerOrderCreateModalPage.Submit();
                CustomerOrderGeneralInformationPage generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
                string updatedLabelText = generalInfo.GetLabel(label);
                generalInfo.SetLabel(label, labelText);
                Assert.AreNotEqual(labelText, updatedLabelText, $"La valeur du champ '{label}' n'a pas été modifiée.");

                // On change de menu et on revient
                generalInfo.PageUp();
                generalInfo.ClickOnDetailTab();
                generalInfo = customerOrderItem.ClickOnGeneralInformationTab();

                // Récupération des valeurs
                updatedLabelText = generalInfo.GetLabel(label);
                generalInfo.BackToList();

                //Assert
                Assert.AreEqual(labelText, updatedLabelText, $"La valeur du champ '{label}' n'a pas été modifiée.");
            }
            finally
            {
                customerOrderPage.DeleteCustomerOrder();
            }
        }

        [Timeout(Timeout)]
        [TestMethod]
        public void PR_CO_GeneralInfo_EditHandler()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            Random rnd = new Random();
            string handler1 = "handler " + rnd.Next(5, 100);
            string handler2 = "handler " + rnd.Next(5, 100);

            // Act
            HomePage homePage = LogInAsAdmin();
            ParametresFlights parametresFlights = homePage.GoToParametres_Flights();
            ParametersFlightHandler handlerPage = parametresFlights.GoToFlightsHandler();
            handlerPage.CreateHandlerModelPage();
            handlerPage.FillFieldsHandlerModelPage(handler1);
            handlerPage.CreateHandlerModelPageFromPopUp();
            handlerPage.FillFieldsHandlerModelPage(handler2);
            handlerPage.Save();
            CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            try
            {
                CustomerOrderCreateModalPage customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrderWithHandler(site, customerName, aircraft, handler1);
                CustomerOrderItem customerOrderItem = customerOrderCreateModalPage.Submit();
                CustomerOrderGeneralInformationPage generalInfo = customerOrderItem.ClickOnGeneralInformationTab();

                bool isHandlerName = generalInfo.GetHandler(handler1);
                Assert.IsTrue(isHandlerName, $"La valeur du champ 'Handler' n'a pas été créé.");

                generalInfo.UpdateHandler(handler2);

                // On change de menu et on revient
                generalInfo.PageUp();
                generalInfo.ClickOnDetailTab();
                generalInfo = customerOrderItem.ClickOnGeneralInformationTab();

                // Récupération des valeurs
                isHandlerName = generalInfo.GetHandler(handler2);
                generalInfo.BackToList();

                //Assert
                Assert.IsTrue(isHandlerName, $"La valeur du champ 'Handler' n'a pas été modifiée.");
            }
            finally
            {
                customerOrderPage.DeleteCustomerOrder();
                homePage.GoToParametres_Flights();
                parametresFlights.GoToFlightsHandler();
                handlerPage.DeleteAircraftByHandlerName(handler1);
                handlerPage.DeleteAircraftByHandlerName(handler2);
            }

        }

        [Timeout(Timeout)]
        [TestMethod]
        public void PR_CO_Filter_DateFrom()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            var date1 = DateUtils.Now;
            var date2 = DateUtils.Now.AddDays(10);
            var date3 = DateUtils.Now.AddDays(-3);
            string customerOrder1Name = null, customerOrder2Name = null, customerOrder3Name = null;
            // Act
            HomePage homePage = LogInAsAdmin();
            CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            try
            {
                CustomerOrderCreateModalPage customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft, date1);
                CustomerOrderItem customerOrderItem = customerOrderCreateModalPage.Submit();
                customerOrder1Name = customerOrderItem.GetNumberCustomer();
                customerOrderItem.BackToList();
                customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft, date2);
                customerOrderItem = customerOrderCreateModalPage.Submit();
                customerOrder2Name = customerOrderItem.GetNumberCustomer();
                customerOrderItem.BackToList();
                customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft, date3);
                customerOrderItem = customerOrderCreateModalPage.Submit();
                customerOrder3Name = customerOrderItem.GetNumberCustomer();
                customerOrderItem.BackToList();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.From, date1);
                customerOrderPage.CloseDatePicker();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrder1Name);
                bool isExist = customerOrderPage.CheckTotalNumber() == 1;
                Assert.IsTrue(isExist, "Le resultat ne correspond pas au 'filtre DateTo'");
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrder2Name);
                isExist = customerOrderPage.CheckTotalNumber() == 1;
                Assert.IsTrue(isExist, "Le resultat ne correspond pas au 'filtre DateTo'");
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrder3Name);
                isExist = customerOrderPage.CheckTotalNumber() == 1;
                Assert.IsFalse(isExist, "Le resultat ne correspond pas au 'filtre DateTo'");
                customerOrderPage.ResetFilter();
            }
            finally
            {
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrder1Name);
                customerOrderPage.DeleteCustomerOrder();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrder2Name);
                customerOrderPage.DeleteCustomerOrder();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrder3Name);
                customerOrderPage.DeleteCustomerOrder();
            }

        }

        [Timeout(Timeout)]
        [TestMethod]
        public void PR_CO_Filter_DateTo()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            var date1 = DateUtils.Now;
            var date2 = DateUtils.Now.AddDays(10);
            var date3 = DateUtils.Now.AddDays(-3);
            string customerOrder1Name = null, customerOrder2Name = null, customerOrder3Name = null;
            // Act
            HomePage homePage = LogInAsAdmin();
            CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            try
            {
                CustomerOrderCreateModalPage customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft,date1);
                CustomerOrderItem customerOrderItem = customerOrderCreateModalPage.Submit();
                customerOrder1Name = customerOrderItem.GetNumberCustomer();
                customerOrderItem.BackToList();
                customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft, date2);
                customerOrderItem = customerOrderCreateModalPage.Submit();
                customerOrder2Name = customerOrderItem.GetNumberCustomer();
                customerOrderItem.BackToList();
                customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft, date3);
                customerOrderItem = customerOrderCreateModalPage.Submit();
                customerOrder3Name = customerOrderItem.GetNumberCustomer();
                customerOrderItem.BackToList();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.To, date1);
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrder1Name);
                bool isExist = customerOrderPage.CheckTotalNumber() == 1;
                Assert.IsTrue(isExist, "Le resultat ne correspond pas au 'filtre DateTo'");
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrder2Name);
                isExist = customerOrderPage.CheckTotalNumber() == 1;
                Assert.IsFalse(isExist, "Le resultat ne correspond pas au 'filtre DateTo'");
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrder3Name);
                isExist = customerOrderPage.CheckTotalNumber() == 1;
                Assert.IsTrue(isExist, "Le resultat ne correspond pas au 'filtre DateTo'");
                customerOrderPage.ResetFilter();
            }
            finally
            {
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrder1Name);
                customerOrderPage.DeleteCustomerOrder();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrder2Name);
                customerOrderPage.DeleteCustomerOrder();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrder3Name);
                customerOrderPage.DeleteCustomerOrder();
            }

        }

        [Timeout(Timeout)]
        [TestMethod]
        public void PR_CO_ChangeStatusToCancelled()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string statusCancelled = "Cancelled";

            // Login 
            HomePage homePage = LogInAsAdmin();

            //Act
            ParametersSites siteParameterPage = homePage.GoToParameters_Sites();
            siteParameterPage.Filter(ParametersSites.FilterType.SearchSite, site);
            siteParameterPage.CheckIfFirstSiteIsActive();


            CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            try
            {
                var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
                var customerOrderDetailPage = customerOrderCreateModalPage.Submit();

                var generalInfo = customerOrderDetailPage.ClickOnGeneralInformationTab();
                var orderNumber = generalInfo.GetOrderNumber();
                customerOrderDetailPage.BackToList();
                customerOrderPage.ResetFilter();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, orderNumber);
                customerOrderPage.EditStatusCO(statusCancelled);
                customerOrderPage.ResetFilter();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, orderNumber);
                //Assert
                bool IsStatusCancelled = customerOrderPage.VerifyStatus(statusCancelled);
                Assert.AreNotEqual(0, customerOrderPage.CheckTotalNumber(), "Le customer order créé n'apparaît pas en tant que Cancelled");


            }
            finally
            {

                customerOrderPage.DeleteCustomerOrder();

            }

        }

        [Timeout(Timeout)]
        [TestMethod]
        public void PR_CO_ChangeStatusToOpened()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string statusOpened = "Opened";

            // Login 
            HomePage homePage = LogInAsAdmin();

            //Act
            ParametersSites siteParameterPage = homePage.GoToParameters_Sites();
            siteParameterPage.Filter(ParametersSites.FilterType.SearchSite, site);
            siteParameterPage.CheckIfFirstSiteIsActive();


            CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            try
            {
                var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
                var customerOrderDetailPage = customerOrderCreateModalPage.Submit();

                var generalInfo = customerOrderDetailPage.ClickOnGeneralInformationTab();
                var orderNumber = generalInfo.GetOrderNumber();
                customerOrderDetailPage.BackToList();
                customerOrderPage.ResetFilter();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, orderNumber);
                customerOrderPage.EditStatusCO(statusOpened);
                customerOrderPage.ResetFilter();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, orderNumber);
                customerOrderPage.ClickStatus();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.StatusOpened, true);
                //Assert
                bool IsStatusCancelled = customerOrderPage.VerifyStatus(statusOpened);
                Assert.AreNotEqual(0, customerOrderPage.CheckTotalNumber(), "Le customer order créé n'apparaît pas en tant que Opened");


            }
            finally
            {

                customerOrderPage.DeleteCustomerOrder();

            }
        }
           
            [Timeout(Timeout)]
            [TestMethod]
            public void PR_CO_FiltreDate_TO()
            {
                // Prepare
                string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
                string site = TestContext.Properties["InvoiceSite"].ToString();
                string aircraft = TestContext.Properties["Aircraft"].ToString();
                DateTime dateTo = DateTime.Now.AddDays(0);
                DateTime dateTo1 = DateTime.Now.AddDays(-5);

                // Login 
                HomePage homePage = LogInAsAdmin();

                //Act
                ParametersSites siteParameterPage = homePage.GoToParameters_Sites();
                siteParameterPage.Filter(ParametersSites.FilterType.SearchSite, site);
                siteParameterPage.CheckIfFirstSiteIsActive();


                CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            try
            {
                var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
                var customerOrderDetailPage = customerOrderCreateModalPage.Submit();
                var generalInfo = customerOrderDetailPage.ClickOnGeneralInformationTab();
                var customerOrderNumber = generalInfo.GetOrderNumber();
                customerOrderDetailPage.BackToList();
                customerOrderPage.ResetFilter();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrderNumber);
                customerOrderPage.Filter(CustomerOrderPage.FilterType.To, dateTo);
                var totalNumber = customerOrderPage.CheckTotalNumber();
                Assert.AreEqual(1, totalNumber, "le filtre date to n'est pas fonctionnel");
                customerOrderPage.ResetFilter();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrderNumber);
                customerOrderPage.Filter(CustomerOrderPage.FilterType.To, dateTo1);
                var totalNumber1 = customerOrderPage.CheckTotalNumber();
                Assert.AreEqual(0, totalNumber1, "le filtre date to n'est pas fonctionnel");
                customerOrderPage.ResetFilter();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrderNumber);
            }
            finally 
            {
                
                customerOrderPage.DeleteCustomerOrder();

            }

        }

        [Timeout(Timeout)]
        [TestMethod]
        public void PR_CO_ChangeStatusToNone()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string statusNone = "none";

            // Login 
            HomePage homePage = LogInAsAdmin();

            //Act
            ParametersSites siteParameterPage = homePage.GoToParameters_Sites();
            siteParameterPage.Filter(ParametersSites.FilterType.SearchSite, site);
            siteParameterPage.CheckIfFirstSiteIsActive();


            CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            try
            {
                var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
                var customerOrderDetailPage = customerOrderCreateModalPage.Submit();

                var generalInfo = customerOrderDetailPage.ClickOnGeneralInformationTab();
                var orderNumber = generalInfo.GetOrderNumber();
                customerOrderDetailPage.BackToList();
                customerOrderPage.ResetFilter();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, orderNumber);
                customerOrderPage.EditStatusCO(statusNone);
                customerOrderPage.ResetFilter();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, orderNumber);
                customerOrderPage.ClickStatus();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.StatusAll, true);
                //Assert
                bool IsStatusCancelled = customerOrderPage.VerifyStatus(statusNone);
                Assert.AreNotEqual(0, customerOrderPage.CheckTotalNumber(), "Le customer order créé n'apparaît pas en tant que None");


            }
            finally
            {

                customerOrderPage.DeleteCustomerOrder();

            }

        }

        [Timeout(Timeout)]
        [TestMethod]
        public void PR_CO_ChangeStatusToClose()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string statusClosed = "Closed";

            // Login 
            HomePage homePage = LogInAsAdmin();

            //Act
            ParametersSites siteParameterPage = homePage.GoToParameters_Sites();
            siteParameterPage.Filter(ParametersSites.FilterType.SearchSite, site);
            siteParameterPage.CheckIfFirstSiteIsActive();


            CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            try
            {
                var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
                var customerOrderDetailPage = customerOrderCreateModalPage.Submit();

                var generalInfo = customerOrderDetailPage.ClickOnGeneralInformationTab();
                var orderNumber = generalInfo.GetOrderNumber();
                customerOrderDetailPage.BackToList();
                customerOrderPage.ResetFilter();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, orderNumber);
                customerOrderPage.EditStatusCO(statusClosed);
                customerOrderPage.ResetFilter();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, orderNumber);
                customerOrderPage.ClickStatus();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.StatusClosed, true);
                //Assert
                bool IsStatusCancelled = customerOrderPage.VerifyStatus(statusClosed);
                Assert.AreNotEqual(0, customerOrderPage.CheckTotalNumber(), "Le customer order créé n'apparaît pas en tant que Closed");


            }
            finally
            {

                customerOrderPage.DeleteCustomerOrder();

            }

        }

        [Timeout(Timeout)]
        [TestMethod]
        public void PR_CO_FiltreDate_From()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            DateTime dateFrom = DateTime.Now.AddDays(0);
            DateTime dateFrom1 = DateTime.Now.AddDays(+5);

            // Login 
            HomePage homePage = LogInAsAdmin();

            //Act
            ParametersSites siteParameterPage = homePage.GoToParameters_Sites();
            siteParameterPage.Filter(ParametersSites.FilterType.SearchSite, site);
            siteParameterPage.CheckIfFirstSiteIsActive();


            CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            try
            {
                var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
                var customerOrderDetailPage = customerOrderCreateModalPage.Submit();
                var generalInfo = customerOrderDetailPage.ClickOnGeneralInformationTab();
                var customerOrderNumber = generalInfo.GetOrderNumber();
                customerOrderDetailPage.BackToList();
                customerOrderPage.ResetFilter();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrderNumber);
                customerOrderPage.Filter(CustomerOrderPage.FilterType.From, dateFrom);
                var totalNumber = customerOrderPage.CheckTotalNumber();
                Assert.AreEqual(1, totalNumber, "le filtre date to n'est pas fonctionnel");
                customerOrderPage.ResetFilter();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrderNumber);
                customerOrderPage.Filter(CustomerOrderPage.FilterType.From, dateFrom1);
                var totalNumber1 = customerOrderPage.CheckTotalNumber();
                Assert.AreEqual(0, totalNumber1, "le filtre date to n'est pas fonctionnel");
                customerOrderPage.ResetFilter();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrderNumber);
            }
            finally
            {

                customerOrderPage.DeleteCustomerOrder();

            }

        }

        [Timeout(Timeout)]
        [TestMethod]
        public void PR_CO_ChangePage()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string statusClosed = "Closed";

            // Login 
            HomePage homePage = LogInAsAdmin();
            CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();

            customerOrderPage.SecondPage();


            customerOrderPage.PageSize("16");
            var numberOfCO_1 = customerOrderPage.GetNumberCO();
            if (numberOfCO_1 < 16)
            {
                var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
                var customerOrderDetailPage = customerOrderCreateModalPage.Submit();

            }
            Assert.AreEqual(16, numberOfCO_1, "avoir 16 lignes de résultats");
            customerOrderPage.PageSize("30");
            var numberOfCO_2 = customerOrderPage.GetNumberCO();

            if (numberOfCO_2 < 30)
            {
                var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
                var customerOrderDetailPage = customerOrderCreateModalPage.Submit();

            }
            Assert.AreEqual(30, numberOfCO_2, "avoir 30 lignes de résultats");

            customerOrderPage.PageSize("50");
            var numberOfCO_3 = customerOrderPage.GetNumberCO();
            if (numberOfCO_3 < 50)
            {
                var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
                var customerOrderDetailPage = customerOrderCreateModalPage.Submit();

            }
            Assert.AreEqual(50, numberOfCO_3, "avoir 50 lignes de résultats");

            customerOrderPage.PageSize("100");
            var numberOfCO_4 = customerOrderPage.GetNumberCO();
            if (numberOfCO_4 < 100)
            {
                var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
                var customerOrderDetailPage = customerOrderCreateModalPage.Submit();

            }
            Assert.AreEqual(100, numberOfCO_4, "avoir 100 lignes de résultats");

        }

        [Timeout(Timeout)]
        [TestMethod]
        public void PR_CO_ChangeSellingUnit()
        {
            // Prepare
            string itemName = TestContext.Properties["InvoiceService"].ToString();
            string sellingUnit1 = "Unit";
            string sellingUnit2 = "Test Unit 1";
   
            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();

            //Customer Order not verified not validated
            string customerId = customerOrderPage.GetFirstCustomerOrderNumber();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerId);
            CustomerOrderItem customerOrderItem = customerOrderPage.ClickFirstLine();
            customerOrderItem.AddNewItem(itemName, "10");
            string sellingUnitBefore = customerOrderItem.GetSellingUnit();
            customerOrderItem.SelectSellingUnit(sellingUnit2);

            WebDriver.Navigate().Refresh();
            string sellingUnitAfter = customerOrderItem.GetSellingUnit();

            //Assert
            Assert.AreNotEqual(sellingUnitAfter, sellingUnitBefore, "Le champs SellingUnit n'a pas changé");

            //Customer Order verified not validated

            customerOrderItem.ShowExtendedMenu();
            customerOrderItem.DoVerify();

            sellingUnitBefore = customerOrderItem.GetSellingUnit();
            customerOrderItem.SelectSellingUnit(sellingUnit1);
            WebDriver.Navigate().Refresh();
            sellingUnitAfter = customerOrderItem.GetSellingUnit();

            //Assert
            Assert.AreNotEqual(sellingUnitAfter, sellingUnitBefore, "Le champs SellingUnit n'a pas changé");

            //Avoir un customer valide et vérifier
            customerOrderItem.ValidateCustomerOrder();
            bool nonInteractableSellingUnit = customerOrderItem.InteractableSellingUnit();

            Assert.IsTrue(nonInteractableSellingUnit, "Le champs SellingUnit peut êtres modifié");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_DeleteLinewithZeroQty()
        {
            string itemName = "[Free Price]";
            string qty = "0";
            string customer = TestContext.Properties[string.Format("CustomerOrderId9")].ToString();
            //Arrange
            HomePage homePage = LogInAsAdmin();
            //Act
            CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.ResetFilter();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customer);
            CustomerOrderItem customerOrderItem = customerOrderPage.SelectFirstCustomerOrder();
            customerOrderItem.DeleteAllItem();
            customerOrderItem.AddNewItem(itemName , qty);
            customerOrderItem.DeleteAllItem();
            //Assert
            bool isItemVisible = customerOrderItem.IsVisible();
            Assert.IsFalse(isItemVisible, "le Item n'est pas été supprimeé");
            // Suppression Item Aprés verification
            customerOrderItem.DoVerify();
            customerOrderItem.AddNewItem(itemName, qty);
            customerOrderItem.DeleteAllItem();
            //Assert
            isItemVisible = customerOrderItem.IsVisible();
            Assert.IsFalse(isItemVisible, "le Item n'est pas été supprimeé");
        }


        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_GeneralInformationAddDelivComment()
        {
            string Comment = "New Comment";
            bool newVersionPrint = true;
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNameZipBegin = "All_files_";
            string DocFileNamePdfBegin = "Order Report";
            List<string> words = new List<string>();
            string customer = TestContext.Properties[string.Format("CustomerOrderId9")].ToString();



            //Arrange
            HomePage homePage = LogInAsAdmin();
            //Act
            CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.ResetFilter();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customer);
            CustomerOrderItem customerOrderItem = customerOrderPage.SelectFirstCustomerOrder();
            CustomerOrderGeneralInformationPage customerOrderGeneralInformationPage = customerOrderItem.ClickOnGeneralInformationTab();
            //Ajouter Commentaire
            customerOrderGeneralInformationPage.SetComment(Comment);
            // verifier le commentaire apparait après refresh sur General information
            WebDriver.Navigate().Refresh();
            customerOrderGeneralInformationPage = customerOrderItem.ClickOnGeneralInformationTab();
            bool isCommentExist = customerOrderGeneralInformationPage.IsCommentExist(Comment);
            Assert.IsTrue(isCommentExist, "le commentaire n'est été pas ajouté");
            // printPdf
            customerOrderGeneralInformationPage.ClearDownloads();
            DeleteAllFileDownload();
            var reportPage = customerOrderGeneralInformationPage.PrintCustomerOrder(newVersionPrint);
            var isReportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();
            //Assert
            Assert.IsTrue(isReportGenerated, "Le document PDF n'a pas pu être généré par l'application.");
            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            string generatedFilePath = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            FileInfo fileInfo = new FileInfo(generatedFilePath);
            PdfDocument document = PdfDocument.Open(fileInfo.FullName);
            string pdfText = string.Join(" ", document.GetPages().SelectMany(page => page.GetWords()).Select(word => word.Text));
            isCommentExist = pdfText.Contains(Comment);
            // Assert
            Assert.IsTrue(isCommentExist, "Commentaire n'a pas été trouvé dans le PDF.");

        }


        [Timeout(Timeout)]
        [TestMethod]
        public void PR_CO_ProdComment()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string itemName = TestContext.Properties["InvoiceService"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string prodComment = "comment";
            string prodComment1 = "comment prod";
            // Login 
            HomePage homePage = LogInAsAdmin();
            CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            CustomerOrderCreateModalPage customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
            customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
            var customerOrderItem = customerOrderCreateModalPage.Submit();
            customerOrderItem.AddNewItem(itemName, "10");
            bool isComment = customerOrderItem.AddProdComment(prodComment);
            //Assert
            Assert.IsTrue(isComment, "Un commentaire a été ajouté.");
            CustomerOrderGeneralInformationPage generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
            CustomerOrderItem CustomerOrderItem = generalInfo.ClickOnDetailTab();
            customerOrderItem.ClickProdComment();
            string comment = customerOrderItem.GetComment();
            //Assert
            Assert.AreEqual(prodComment, comment, "Le commentaire n'a pas été ajouté correctement.");
            customerOrderItem.CloseCommentModal();
            generalInfo.BackToList();
            customerOrderPage.ResetFilter();
            customerOrderPage.CustomerOrderCreatePage();
            customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
            customerOrderItem = customerOrderCreateModalPage.Submit();
            customerOrderItem.AddNewItem(itemName, "20");
            customerOrderItem.DoVerify();
            bool isComment1 = customerOrderItem.AddProdComment(prodComment1);
            //Assert
            Assert.IsTrue(isComment1, "Un commentaire a été ajouté.");
            generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
            customerOrderItem = generalInfo.ClickOnDetailTab();
            customerOrderItem.ClickProdComment();
            string comment1 = customerOrderItem.GetComment();
            //Assert
            Assert.AreEqual(prodComment1, comment1, "Le commentaire n'a pas été ajouté correctement.");
        }

        [Timeout(Timeout)]
        [TestMethod]
        public void PR_CO_UnitPrice()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string itemName = TestContext.Properties["InvoiceService"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string price = "10";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.ResetFilter();
            CustomerOrderCreateModalPage customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
            customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
            var customerOrderItem = customerOrderCreateModalPage.Submit();
            customerOrderItem.AddNewItem(itemName, "10");
            customerOrderItem.SetItemPrice(price);
            WebDriver.Navigate().Refresh();
            string priceAfter=customerOrderItem.GetItemPrice();
            //Assert
            Assert.AreEqual(priceAfter, price, "La modification du prix n'a pas été prise en compte");
            customerOrderItem.BackToList();
            customerOrderPage.ResetFilter();
            CustomerOrderCreateModalPage CustomerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
            customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
            customerOrderItem = customerOrderCreateModalPage.Submit();
            customerOrderItem.DoVerify();
            customerOrderItem.AddNewItem(itemName, "20");
            customerOrderItem.SetItemPrice(price);
            WebDriver.Navigate().Refresh();
            string priceAfter1 = customerOrderItem.GetItemPrice();
            //Assert
            Assert.AreEqual(priceAfter1, price, "La modification du prix n'a pas été prise en compte");

        }

        [Timeout(Timeout)]
        [TestMethod]
        public void PR_CO_TotalPrice()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string itemName = TestContext.Properties["InvoiceService"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string totalPrice = "100";
            string totalPrice1 = "200";
            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.ResetFilter();
            CustomerOrderCreateModalPage customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
            customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
            var customerOrderItem = customerOrderCreateModalPage.Submit();
            customerOrderItem.AddNewItem(itemName, "1");
            customerOrderItem.SetItemPriceTotal(totalPrice);
            WebDriver.Navigate().Refresh();
            string priceExcl = customerOrderItem.GetPriceExcl();
            //Assert 
            Assert.AreEqual(totalPrice, priceExcl , "Le Total Price et le prix en bas de page ne sont pas synchronisés après l'actualisation");
            customerOrderItem.BackToList();
            customerOrderPage.ResetFilter();
            CustomerOrderCreateModalPage CustomerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
            customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
            customerOrderItem = customerOrderCreateModalPage.Submit();
            customerOrderItem.DoVerify();
            customerOrderItem.AddNewItem(itemName, "2");
            customerOrderItem.SetItemPriceTotal(totalPrice1);
            WebDriver.Navigate().Refresh();
            string priceExcl1 = customerOrderItem.GetPriceExcl();
            Assert.AreEqual(totalPrice1, priceExcl1, "Le Total Price et le prix en bas de page ne sont pas synchronisés après l'actualisation");

        }
        ///////////////////////Methode SQL ///////////////
        ///
        private int InsertCustomerOrder(string site, string customer, DateTime date, int isDelivered = 0, int isValid = 1, int IsVerified = 0, int statut = 0, int isPreOrder = 0)
        {
            string queryFlightDeliveries = @"
        DECLARE @siteId INT;
        SELECT TOP 1 @siteId = Id FROM sites WHERE Name LIKE @site;
 
        DECLARE @customerId INT;
        SELECT TOP 1 @customerId = Id FROM customers WHERE Name LIKE @customer;
        Insert into [dbo].[Orders] (
        [OrderDate]
          ,[SiteId]
          ,[CustomerId]
          ,[CreationDate]
          ,[InvoiceStepId]
          ,[PriceExclVAT]
          ,[PriceVAT]
          ,[LastUpdate]
          ,[IsVatFree]
          ,[IsValidated]
          ,[IsAirportTax]
          ,[Status]
          ,[IsForeignName]
          ,[IsPreorder]
          ,[COShippingDuration]
          ,[COProcessingDuration]
          ,[COLastUpdate]
          ,[IsSentToWMS],
           [IsDelivered],
           [IsVerified])
 
          values (
          @date,
          @siteId,
          @customerId,
          @date,
          '1',
          '10',
          '10',
          @date,
          0,
          @isValid,
          1,
          @statut,
          0,
          @isPreOrder,
          '0',
          '0',
          '0',
          0,
            @isDelivered,
            @IsVerified
          );
         SELECT SCOPE_IDENTITY();
        "
            ;
            return ExecuteAndGetInt(
            queryFlightDeliveries,
                new KeyValuePair<string, object>("site", site),
                new KeyValuePair<string, object>("customer", customer),
                new KeyValuePair<string, object>("isPreOrder", isPreOrder),
                new KeyValuePair<string, object>("date", date),
                new KeyValuePair<string, object>("isDelivered", isDelivered),
                new KeyValuePair<string, object>("IsVerified", IsVerified),
                new KeyValuePair<string, object>("isValid", isValid),
                new KeyValuePair<string, object>("statut", statut)
            );
        }
        private void DeleteCustomerOrder(int customerOrderId)
        {
            string query = @"
            delete from Orders where id = @customerOrderId";

            ExecuteNonQuery(
                query,
                new KeyValuePair<string, object>("customerOrderId", customerOrderId));
        }

        private int InsertInflightCustomerOrder(string site, string customer, DateTime date, string flightNumber, string aircraft, string origin = null, string destination = null, int statut = 0, int isValid = 1, int isDelivered = 0, int IsVerified = 0)
        {
            string queryFlightDeliveries = @"
        DECLARE @siteId INT;
        SELECT TOP 1 @siteId = Id FROM sites WHERE Name LIKE @site;
 
        DECLARE @customerId INT;
        SELECT TOP 1 @customerId = Id FROM customers WHERE Name LIKE @customer;

        DECLARE @aircraftId INT;
        SELECT TOP 1 @aircraftId = Id FROM aircraft WHERE Name LIKE @aircraft;

        DECLARE @airportOriginId INT;
        SELECT TOP 1 @airportOriginId = Id FROM Airports WHERE Name LIKE @origin;

        DECLARE @airportDestinId INT;
        SELECT TOP 1 @airportDestinId = Id FROM Airports WHERE Name LIKE @destination;

        Insert into [dbo].[Orders] (
        [OrderDate]
          ,[SiteId]
          ,[CustomerId]
          ,[CreationDate]
          ,[InvoiceStepId]
          ,[PriceExclVAT]
          ,[PriceVAT]
          ,[LastUpdate]
          ,[IsVatFree]
          ,[IsValidated]
          ,[IsAirportTax]
          ,[FlightNumber]
            ,[AircraftId]
      ,[DestinationAirportId]
      ,[AirportId]
          ,[Status]
          ,[IsForeignName]
          ,[IsPreorder]
          ,[COShippingDuration]
          ,[COProcessingDuration]
          ,[COLastUpdate]
          ,[IsSentToWMS],
           [IsDelivered],
           [IsVerified])
 
          values (
          @date,
          @siteId,
          @customerId,
          @date,
          '2',
          '10',
          '10',
          @date,
          0,
          @isValid,
          1,
        @flightNumber,
        @aircraftId,
        @airportDestinId,
        @airportOriginId,
          @statut,
          0,
          @isPreOrder,
          '0',
          '0',
          '0',
          0,
            @isDelivered,
            @IsVerified
          );
         SELECT SCOPE_IDENTITY();
        "
            ;
            return ExecuteAndGetInt(
            queryFlightDeliveries,
                new KeyValuePair<string, object>("site", site),
                new KeyValuePair<string, object>("customer", customer),
                new KeyValuePair<string, object>("flightNumber", flightNumber),
                new KeyValuePair<string, object>("date", date),
                new KeyValuePair<string, object>("origin", origin),
                new KeyValuePair<string, object>("aircraft", aircraft),
                new KeyValuePair<string, object>("destination", destination),
                new KeyValuePair<string, object>("isDelivered", isDelivered),
                new KeyValuePair<string, object>("IsVerified", IsVerified)
            );
        }
        [TestMethod]
        [Timeout(Timeout)]
        public void PR_CO_FilterPreOrderNO()
        {
            // Prepare
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string No = "No";
            string delivery = new Random().Next().ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            //Arrange            
            var homePage = LogInAsAdmin();
            //Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            //Creation d'un customer order 
            var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
            //customerOrderCreateModalPage.FillField_CreatNewCustomerOrderWOUTflight(site, customerName, delivery, DateUtils.Now.AddDays(6));
            customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
            var customerOrderDetailPage = customerOrderCreateModalPage.Submit();
            var generalInfo = customerOrderDetailPage.ClickOnGeneralInformationTab();
            var customerOrderNumber = generalInfo.GetOrderNumber();
            customerOrderDetailPage.BackToList();
            //Filter 
            customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrderNumber);
            customerOrderPage.Filter(CustomerOrderPage.FilterType.PreOrderNo, No);
            var filteredOrders = customerOrderPage.GetFilteredOrders();
            //Assert
            Assert.AreEqual(1, filteredOrders.Count, "Erreur : L'affichage des données ne correspond pas au filtre 'PreOrderNo' sélectionné.");
        }

        private int? CustomerOrderExist(int id)
        {
            string query = @"
             select Id from Orders where Id = @id

              SELECT SCOPE_IDENTITY();
            "
            ;
            int? result = ExecuteAndGetInt(
            query,
            new KeyValuePair<string, object>("id", id)
            );
            return result == 0 ? (int?)null : result;
        }
    }
}