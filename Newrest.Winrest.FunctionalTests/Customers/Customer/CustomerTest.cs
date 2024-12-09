using DocumentFormat.OpenXml.Spreadsheet;
using iText.Signatures.Validation.V1.Report;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Catalogs;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Customer;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.PriceList;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Sites;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.PageObjects.ToDoList.Tasks;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.IO;
using System.Web.UI.WebControls;
using static Newrest.Winrest.FunctionalTests.PageObjects.Customers.Customer.CustomerCardexNotifPage;
using static Newrest.Winrest.FunctionalTests.PageObjects.Customers.Customer.CustomerPage;
using static Newrest.Winrest.FunctionalTests.PageObjects.Customers.Customer.CustomerReinvoicePage;

namespace Newrest.Winrest.FunctionalTests.Customers
{
    [TestClass]
    public class CustomerTest : TestBase
    {
        private static Random random = new Random();
        private const int _timeout = 600000;
        string customerNameToday = $"CustomerTest-{DateUtils.Now.ToShortDateString()}";
        string customerNameTodayInactive = "";
        string customerNameTodayActive = "";
        private const string CUSTOMERS = "Customers";
        const string comment = "New comment for test";

        [TestInitialize]
        public override void TestInitialize()
        {

            base.TestInitialize();

            var testMethod = TestContext.TestName;
            switch (testMethod)
            {
                case nameof(CU_CU_CustomerIndexSearch_Showactiveonly):
                    TestInitialize_CreateActiveCustomer();
                    TestInitialize_CreateInActiveCustomer();
                    break;

                case nameof(CU_CU_TaxOffice):
                    TestInitialize_CreateActiveCustomer();
                    break;

                case nameof(CU_CU_ProcessingDuration):
                    TestInitialize_CreateActiveCustomer();
                    break;

                case nameof(CU_CU_CustomersIndexSearch_ShowInactiveOnly):
                    TestInitialize_CreateActiveCustomer();
                    TestInitialize_CreateInActiveCustomer();
                    break;

                case nameof(CU_CU_CustomerIndex_SearchShowAll):
                    TestInitialize_CreateActiveCustomer();
                    TestInitialize_CreateInActiveCustomer();
                    break;

                case nameof(CU_CU_CustomersGeneralInformation_RevisionPrice):

                    TestInitialize_CreateActiveCustomer();
                    break;

                case nameof(CU_CU_CustomerGeneralInformation_ThirdId):
                    TestInitialize_CreateActiveCustomer();
                    break;

                case nameof(CU_CU_CustomerIndex_Searchtext):
                    TestInitialize_CreateActiveCustomer();
                    break;

                case nameof(CU_CU_VATFree):
                    TestInitialize_CreateActiveCustomer();
                    break;

                case nameof(CU_CUST_Filter_Search):
                    TestInitialize_CreateActiveCustomer();
                    break;

                case nameof(CU_CUST_Filter_SortBy):
                    TestInitialize_CreateActiveCustomer();
                    break;
                case nameof(CU_CUST_Filter_ShowAll):
                    TestInitialize_CreateActiveCustomer();
                    TestInitialize_CreateInActiveCustomer();
                    break;
                case nameof(CU_CUST_Filter_ShowActive):
                    TestInitialize_CreateActiveCustomer();
                    TestInitialize_CreateInActiveCustomer();
                    break;
                case nameof(CU_CUST_Filter_ShowInactive):
                    TestInitialize_CreateActiveCustomer();
                    TestInitialize_CreateInActiveCustomer();
                    break;
                case nameof(CU_CUST_Modify_General_Informations):
                    TestInitialize_CreateActiveCustomer();
                    break;
                case nameof(CU_CUST_Modify_General_Informations_For_BuyOnBoard):
                    TestInitialize_CreateActiveCustomer();
                    break;
                case nameof(CU_CUST_Modify_BuyOnBoard):
                    TestInitialize_CreateActiveCustomer();
                    break;
                case nameof(CU_CUST_Create_New_Contact):
                    TestInitialize_CreateActiveCustomer();
                    break;
                case nameof(CU_CUST_Create_Existing_Contact):
                    TestInitialize_CreateActiveCustomer();
                    break;
                case nameof(CU_CUST_Modify_Contact):
                    TestInitialize_CreateActiveCustomer();
                    break;
                case nameof(CU_CUST_Delete_Contact):
                    TestInitialize_CreateActiveCustomer();
                    break;
                case nameof(CU_CUST_Link_Price_List_To_Customer):
                    TestInitialize_CreateActiveCustomer();
                    break;
                case nameof(CU_CUST_Link_To_Price_List):
                    TestInitialize_CreateActiveCustomer();
                    break;
                case nameof(CU_CUST_Create_New_Discount):
                    TestInitialize_CreateActiveCustomer();
                    break;
                case nameof(CU_CUST_Create_New_Increase):
                    TestInitialize_CreateActiveCustomer();
                    break;
                case nameof(CU_CUST_Modify_Discount):
                    TestInitialize_CreateActiveCustomer();
                    break;
                case nameof(CU_CUST_Delete_Discount):
                    TestInitialize_CreateActiveCustomer();
                    break;
                case nameof(CU_CUST_Create_New_Cardex_Notification):
                    TestInitialize_CreateActiveCustomer();
                    break;
                case nameof(CU_CUST_Modify_Cardex_Notification):
                    TestInitialize_CreateActiveCustomer();
                    break;
                case nameof(CU_CUST_Delete_Cardex_Notification):
                    TestInitialize_CreateActiveCustomer();
                    break;
                case nameof(CU_CUST_Search_Cardex_By_Registration):
                    TestInitialize_CreateActiveCustomer();
                    break;
                case nameof(CU_CUST_Create_New_Agreement):
                    TestInitialize_CreateActiveCustomer();
                    break;
                case nameof(CU_CUST_Create_Existing_Agreement):
                    TestInitialize_CreateActiveCustomer();
                    break;
                case nameof(CU_CUST_Modify_Agreement):
                    TestInitialize_CreateActiveCustomer();
                    break;
                case nameof(CU_CUST_Delete_Agreement):
                    TestInitialize_CreateActiveCustomer();
                    break;
                case nameof(CU_CUST_Create_New_Reinvoice):
                    TestInitialize_CreateActiveCustomer();
                    break;
                case nameof(CU_CUST_Create_Existing_Reinvoice):
                    TestInitialize_CreateActiveCustomer();
                    break;
                case nameof(CU_CUST_Create_Reinvoice_Same_From_And_To_Site):
                    TestInitialize_CreateActiveCustomer();
                    break;
                case nameof(CU_CUST_Modify_Reinvoice):
                    TestInitialize_CreateActiveCustomer();
                    break;
                case nameof(CU_CU_Modify_Generalnfo_InvoiceComment):
                    TestInitialize_CreateActiveCustomer();
                    break;
                case nameof(CU_CUST_Delete_Reinvoice):
                    TestInitialize_CreateActiveCustomer();
                    break;
                case nameof(CU_CUST_Search_Reinvoice_By_Site):
                    TestInitialize_CreateActiveCustomer();
                    break;
                case nameof(CU_CUST_Create_CustomerBob):
                    TestInitialize_CreateActiveCustomer();
                    break;
                case nameof(CU_CU_CustomersGeneralInformation_SIRET):
                    TestInitialize_CreateActiveCustomer();
                    break;
                case nameof(CU_CU_GeneralInformation_ProformaAddress):
                    TestInitialize_CreateActiveCustomer();
                    break;
                case nameof(CU_CU_Modify_Generalnfo_ICAO):
                    TestInitialize_CreateActiveCustomer();
                    break;
                case nameof(CU_CU_Modify_Generalnfo_DeliveryNoteComment):
                    TestInitialize_CreateActiveCustomer();
                    break;
                case nameof(CU_CU_GeneralInformation_Name):
                    TestInitialize_CreateActiveCustomer();
                    break;
                case nameof(CU_CU_Modify_Generalnfo_Logo):
                    TestInitialize_CreateActiveCustomer();
                    break;
                case nameof(CU_CU_Modify_Generalnfo_DeliveryNoteValorization):
                    TestInitialize_CreateActiveCustomer();
                    break;
                case nameof(CU_CU_GeneralInformation_PaymentTerm):
                    TestInitialize_CreateActiveCustomer();
                    break;
                case nameof(CU_CU_Modify_Generalnfo_EngagementNo):
                    TestInitialize_CreateActiveCustomer();
                    break;
                case nameof(CU_CUST_Modify_Generalnfo_ExternalId):
                    TestInitialize_CreateActiveCustomer();
                    break;
                case nameof(CU_CU_Modify_Generalnfo_Filigran):
                    TestInitialize_CreateActiveCustomer();
                    break;
                case nameof(CU_CU_Modify_Generalnfo_IATA):
                    TestInitialize_CreateActiveCustomer();
                    break;
                case nameof(CU_CUST_Modify_Generalnfo_LastUpdate):
                    TestInitialize_CreateActiveCustomer();
                    break;

                default:
                    break;
            }
        }

        private void TestInitialize_CreateActiveCustomer()
        {
            //Prepare
            Random random = new Random();
            string customerType = TestContext.Properties["CustomerType"].ToString();
            string customerCode = random.Next().ToString();
            customerNameTodayActive = "CustomerActive-" + customerType + DateUtils.Now.ToString("dd/MM/yyyy") + "-" + random.Next().ToString();

            TestContext.Properties[string.Format("CustomerId")] =  InsertCustomer(customerCode, customerNameTodayActive, customerType);
            int i = (int)TestContext.Properties[string.Format("CustomerId")];
            Assert.IsNotNull(CustomerExist((int)TestContext.Properties[string.Format("CustomerId")]), "Le customer n'est pas crée.");
        }
        private void TestInitialize_CreateInActiveCustomer()
        {
            //Prepare
            Random random = new Random();
            string customerType = TestContext.Properties["CustomerType"].ToString();
            string customerCode = random.Next().ToString();
            customerNameTodayInactive = "CustomerInactive-" + DateUtils.Now.ToString("dd/MM/yyyy") + "-" + random.Next().ToString();

            TestContext.Properties[string.Format("CustomerId")] = InsertCustomer(customerCode, customerNameTodayInactive, customerType, 0);
            Assert.IsNotNull(CustomerExist((int)TestContext.Properties[string.Format("CustomerId")]), "Le customer n'est pas crée.");

        }

        [TestCleanup]
        public override void TestCleanup()
        {
            var testMethod = TestContext.TestName;
            switch (testMethod)
            {
                case nameof(CU_CU_CustomerIndexSearch_Showactiveonly):
                    TestCleanup_DeleteCustomers();
                    break;

                case nameof(CU_CU_TaxOffice):
                    TestCleanup_DeleteCustomers();
                    break;

                case nameof(CU_CU_ProcessingDuration):
                    TestCleanup_DeleteCustomers();
                    break;

                case nameof(CU_CU_CustomersIndexSearch_ShowInactiveOnly):
                    TestCleanup_DeleteCustomers();
                    break;

                case nameof(CU_CU_CustomerIndex_SearchShowAll):
                    TestCleanup_DeleteCustomers();
                    break;

                case nameof(CU_CU_CustomersGeneralInformation_RevisionPrice):

                    TestCleanup_DeleteCustomers();
                    break;

                case nameof(CU_CU_CustomerGeneralInformation_ThirdId):
                    TestCleanup_DeleteCustomers();
                    break;

                case nameof(CU_CU_CustomerIndex_Searchtext):
                    TestCleanup_DeleteCustomers();
                    break;

                case nameof(CU_CU_VATFree):
                    TestCleanup_DeleteCustomers();
                    break;
                case nameof(CU_CUST_Filter_Search):
                    TestCleanup_DeleteCustomers();
                    break;
                case nameof(CU_CUST_Filter_SortBy):
                    TestCleanup_DeleteCustomers();
                    break;
                case nameof(CU_CUST_Filter_ShowAll):
                    TestCleanup_DeleteCustomers();
                    break;
                case nameof(CU_CUST_Filter_ShowActive):
                    TestCleanup_DeleteCustomers();
                    break;
                case nameof(CU_CUST_Filter_ShowInactive):
                    TestCleanup_DeleteCustomers();
                    break;
                case nameof(CU_CUST_Modify_General_Informations):
                    TestCleanup_DeleteCustomers();
                    break;
                case nameof(CU_CUST_Modify_General_Informations_For_BuyOnBoard):
                    TestCleanup_DeleteCustomers();
                    break;
                case nameof(CU_CUST_Create_New_Contact):
                    TestCleanup_DeleteCustomers();
                    break;
                case nameof(CU_CUST_Create_Existing_Contact):
                    TestCleanup_DeleteCustomers();
                    break;
                case nameof(CU_CUST_Modify_Contact):
                    TestCleanup_DeleteCustomers();
                    break;
                case nameof(CU_CUST_Delete_Contact):
                    TestCleanup_DeleteCustomers();
                    break;
                case nameof(CU_CUST_Link_To_Price_List):
                    TestCleanup_DeleteCustomers();
                    break;
                case nameof(CU_CUST_Create_New_Discount):
                    TestCleanup_DeleteCustomers();
                    break;
                case nameof(CU_CUST_Create_New_Increase):
                    TestCleanup_DeleteCustomers();
                    break;
                case nameof(CU_CUST_Modify_Discount):
                    TestCleanup_DeleteCustomers();
                    break;
                case nameof(CU_CUST_Delete_Discount):
                    TestCleanup_DeleteCustomers();
                    break;
                case nameof(CU_CUST_Create_New_Agreement):
                    TestCleanup_DeleteCustomers();
                    break;
                case nameof(CU_CUST_Create_Existing_Agreement):
                    TestCleanup_DeleteCustomers();
                    break;
                case nameof(CU_CUST_Modify_Agreement):
                    TestCleanup_DeleteCustomers();
                    break;
                case nameof(CU_CUST_Delete_Agreement):
                    TestCleanup_DeleteCustomers();
                    break;
                case nameof(CU_CUST_Create_New_Reinvoice):
                    TestCleanup_DeleteCustomers();
                    break;
                case nameof(CU_CUST_Create_Existing_Reinvoice):
                    TestCleanup_DeleteCustomers();
                    break;
                case nameof(CU_CUST_Create_Reinvoice_Same_From_And_To_Site):
                    TestCleanup_DeleteCustomers();
                    break;
                case nameof(CU_CUST_Modify_Reinvoice):
                    TestCleanup_DeleteCustomers();
                    break;
                case nameof(CU_CUST_Delete_Reinvoice):
                    TestCleanup_DeleteCustomers();
                    break;
                case nameof(CU_CUST_Search_Reinvoice_By_Site):
                    TestCleanup_DeleteCustomers();
                    break;
                case nameof(CU_CUST_Create_CustomerBob):
                    TestCleanup_DeleteCustomers();
                    break;
                case nameof(CU_CU_GeneralInformation_ProformaAddress):
                    TestCleanup_DeleteCustomers();
                    break;
                case nameof(CU_CU_Modify_Generalnfo_InvoiceComment):
                    TestCleanup_DeleteCustomers();
                    break;
                case nameof(CU_CU_Modify_Generalnfo_DeliveryNoteComment):
                    TestCleanup_DeleteCustomers();
                    break;
                case nameof(CU_CU_Modify_Generalnfo_ICAO):
                    TestCleanup_DeleteCustomers();
                    break;
                case nameof(CU_CU_Modify_Generalnfo_Logo):
                    TestCleanup_DeleteCustomers();
                    break;
                case nameof(CU_CU_Modify_Generalnfo_DeliveryNoteValorization):
                    TestCleanup_DeleteCustomers();
                    break;
                case nameof(CU_CU_Modify_Generalnfo_EngagementNo):
                    TestCleanup_DeleteCustomers();
                    break;
                case nameof(CU_CUST_Modify_Generalnfo_ExternalId):
                    TestCleanup_DeleteCustomers();
                    break;
                case nameof(CU_CU_Modify_Generalnfo_Filigran):
                    TestCleanup_DeleteCustomers();
                    break;
                case nameof(CU_CU_Modify_Generalnfo_IATA):
                    TestCleanup_DeleteCustomers();
                    break;
                case nameof(CU_CU_GeneralInformation_Name):
                    TestCleanup_DeleteCustomers();
                    break;
                case nameof(CU_CU_GeneralInformation_PaymentTerm):
                    TestCleanup_DeleteCustomers();
                    break;
                case nameof(CU_CUST_Modify_Generalnfo_LastUpdate):
                    TestCleanup_DeleteCustomers();
                    break;
                default:
                    break;
            }
        }

        public void TestCleanup_DeleteCustomers()
        {
            DeleteCustomer((int)TestContext.Properties[string.Format("CustomerId")]);

            Assert.IsNull(CustomerExist((int)TestContext.Properties[string.Format("CustomerId")]), "La Task n'est pas supprimé.");
        }
        // Créer un nouveau client
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Create_New_Customer()
        {
            // Prepare
            Random rnd = new Random();
            string customerType = TestContext.Properties["CustomerType"].ToString();

            string customerName = customerNameToday + "-" + rnd.Next().ToString();
            string customerCode = rnd.Next().ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var customersPage = homePage.GoToCustomers_CustomerPage();
            var customerCreateModalPage = customersPage.CustomerCreatePage();
            customerCreateModalPage.FillFields_CreateCustomerModalPage(customerName, customerCode, customerType);
            var customerGeneralInformationsPage = customerCreateModalPage.Create();
            customersPage = customerGeneralInformationsPage.BackToList();

            //Assert
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerName);
            string firstCustomerName = customersPage.GetFirstCustomerName();
            Assert.AreEqual(customerName, firstCustomerName, "Le customer n'a pas été créé.");
        }

        // Créer un nouveau client sans le champ "Name" obligatoire
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Create_New_Customer_Without_Name()
        {
            // Prepare
            string customerType = TestContext.Properties["CustomerType"].ToString();

            string customerName = String.Empty;
            string customerCode = String.Empty;

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var customersPage = homePage.GoToCustomers_CustomerPage();
            var customerCreateModalPage = customersPage.CustomerCreatePage();
            customerCreateModalPage.FillFields_CreateCustomerModalPage(customerName, customerCode, customerType);
            customerCreateModalPage.Create();

            // Assert
            Assert.IsTrue(customerCreateModalPage.IsNameErrorDisplayed(), MessageErreur.MESSAGE_ERREUR_NON_OBTENU);
            Assert.IsTrue(customerCreateModalPage.IsIcaoErrorDisplayed(), MessageErreur.MESSAGE_ERREUR_NON_OBTENU);
        }

        //_______________________________________________________________FIN_CREATE__________________________________________________________

        //_______________________________________________________________FILTRES_____________________________________________________________

        // Rechercher un client par son nom ou ICAO
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Filter_Search()
        {

            // Arrange
            HomePage homePage = LogInAsAdmin();
            // Act
            var customersPage = homePage.GoToCustomers_CustomerPage();
            //Assert
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            string firstNameCustomer = customersPage.GetFirstCustomerName();
            Assert.AreEqual(customerNameTodayActive, firstNameCustomer, MessageErreur.FILTRE_ERRONE, "Search");
        }

        // Trier les clients par leur nom
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Filter_SortBy()
        {
            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();

            if (!customersPage.isPageSizeEqualsTo100())
            {
                customersPage.PageSize("8");
                customersPage.PageSize("100");
            }

            try
            {
                // Create          
                customersPage.Filter(FilterType.SortBy, "NAME");
                bool isSortedByName = customersPage.IsSortedByName();

                customersPage.Filter(FilterType.SortBy, "CUSTOMER");
                bool isSortedByTypeOfCustomer = customersPage.IsSortedByCustomerType();

                // Assert
                Assert.IsTrue(isSortedByName, MessageErreur.FILTRE_ERRONE, "Sort by name");
                Assert.IsTrue(isSortedByTypeOfCustomer, MessageErreur.FILTRE_ERRONE, "Sort by type of customer");
            }
            finally
            {
                customersPage.PageSize("8");
            }
        }

        // Afficher tous les clients (actifs et inactifs)
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Filter_ShowAll()
        {
            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.ShowInactive, true);
            int nbInactive = customersPage.CheckTotalNumber();

            customersPage.Filter(FilterType.ShowActive, true);
            int nbActive = customersPage.CheckTotalNumber();

            customersPage.Filter(FilterType.ShowAll, true);
            int nbTotal = customersPage.CheckTotalNumber();

            //Assert
            Assert.AreEqual(nbTotal, (nbActive + nbInactive), String.Format(MessageErreur.FILTRE_ERRONE, "'Show all'"));
        }

        // Afficher tous les clients actifs
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Filter_ShowActive()
        {
            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.ShowAll, true);
            var totalNumberShowAll = customersPage.CheckTotalNumber();
            customersPage.Filter(FilterType.ShowInactive, true);
            var totalNumberShowInactive = customersPage.CheckTotalNumber();


            if (!customersPage.isPageSizeEqualsTo100())
            {
                customersPage.PageSize("8");
                customersPage.PageSize("100");
            }
            customersPage.Filter(FilterType.ShowActive, true);
            var totalNumberShowActive = customersPage.CheckTotalNumber();


            // Assert
            Assert.IsTrue(customersPage.CheckStatus(true), String.Format(MessageErreur.FILTRE_ERRONE, "'Show active'"));
            Assert.AreEqual(totalNumberShowAll - totalNumberShowInactive, totalNumberShowActive, "'Show active'");

        }

        // Afficher tous les clients inactifs
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Filter_ShowInactive()
        {
            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();

            if (!customersPage.isPageSizeEqualsTo100())
            {
                customersPage.PageSize("8");
                customersPage.PageSize("100");
            }
            customersPage.Filter(FilterType.ShowAll, true);
            var totalNumberShowAll = customersPage.CheckTotalNumber();
            customersPage.Filter(FilterType.ShowActive, true);
            var totalNumberShowActive = customersPage.CheckTotalNumber();
            customersPage.Filter(FilterType.ShowInactive, true);
            var totalNumberShowInactive = customersPage.CheckTotalNumber();

            // Assert
            Assert.IsFalse(customersPage.CheckStatus(false), String.Format(MessageErreur.FILTRE_ERRONE, "'Show inactive'"));
            Assert.AreEqual(totalNumberShowInactive, totalNumberShowAll - totalNumberShowActive, "le filtre n'est pas fonctionnel");
        }

        // Afficher tous les clients de tous les types
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Filter_AllTypesOfCustomer()
        {

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();

            // Create
            customersPage.Filter(FilterType.TypeOfCustomer, "Financiero");
            int financieroNumber = customersPage.CheckTotalNumber();

            customersPage.ResetFilters();
            customersPage.Filter(FilterType.TypeOfCustomer, "Ash");
            int ashNumber = customersPage.CheckTotalNumber();

            customersPage.ResetFilters();
            customersPage.Filter(FilterType.TypeOfCustomer, "Colectividades");
            int collectividadesNumber = customersPage.CheckTotalNumber();

            customersPage.ResetFilters();
            customersPage.Filter(FilterType.TypeOfCustomer, "Inflight");
            int inflightNumber = customersPage.CheckTotalNumber();

            customersPage.ResetFilters();
            customersPage.Filter(FilterType.TypeOfCustomer, "Resto Co");
            int restoCoNumber = customersPage.CheckTotalNumber();

            customersPage.Filter(FilterType.TypeOfCustomer, "ALL CUSTOMER TYPES");
            int totalNumber = customersPage.CheckTotalNumber();

            // Assert           
            Assert.AreEqual(totalNumber, (financieroNumber + ashNumber + collectividadesNumber + inflightNumber + restoCoNumber), MessageErreur.FILTRE_ERRONE, "All customer types");
        }

        // Afficher tous les clients du type Inflight
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Filter_ShowTypes_Inflight()
        {
            // Prepare
            Random rnd = new Random();

            string customerName = customerNameToday + "-" + rnd.Next(1000, 9000).ToString();
            string customerCode = rnd.Next().ToString();
            string customerType = "Inflight";

            // Arrange
            HomePage homePage = LogInAsAdmin();


            // Act
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.TypeOfCustomer, customerType);

            if (customersPage.CheckTotalNumber() < 20)
            {
                // Création Customer Inflight
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customerName, customerCode, customerType);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();
                customersPage.ResetFilters();
                customersPage.Filter(FilterType.TypeOfCustomer, customerType);
            }

            if (!customersPage.isPageSizeEqualsTo100())
            {
                customersPage.PageSize("8");
                customersPage.PageSize("100");
            }

            try
            {
                Assert.IsTrue(customersPage.VerifyTypeCustomer(customerType), MessageErreur.FILTRE_ERRONE, customerType);
            }
            finally
            {
                customersPage.PageSize("8");
            }
        }

        // Afficher tous les clients du type Ash
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Filter_ShowTypes_Ash()
        {
            // Prepare
            Random rnd = new Random();

            string customerName = customerNameToday + "-" + rnd.Next(1000, 9000).ToString();
            string customerCode = rnd.Next().ToString();
            string customerType = "Ash";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.TypeOfCustomer, customerType);

            if (customersPage.CheckTotalNumber() < 20)
            {
                // Création Customer Ash
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customerName, customerCode, customerType);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();
                customersPage.ResetFilters();
                customersPage.Filter(FilterType.TypeOfCustomer, customerType);
            }

            if (!customersPage.isPageSizeEqualsTo100())
            {
                customersPage.PageSize("8");
                customersPage.PageSize("100");
            }

            try
            {
                Assert.IsTrue(customersPage.VerifyTypeCustomer(customerType), MessageErreur.FILTRE_ERRONE, customerType);
            }
            finally
            {
                customersPage.PageSize("8");
            }
        }

        // Afficher tous les clients du type Colectividades
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Filter_ShowTypes_Colectividades()
        {
            // Prepare
            Random rnd = new Random();
            string customerName = customerNameToday + "-" + rnd.Next(1000, 9000).ToString();
            string customerType = "Colectividades";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.TypeOfCustomer, customerType);

            if (customersPage.CheckTotalNumber() < 20)
            {
                // Création Customer Colectividades
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customerName, rnd.Next(10000, 20000).ToString(), customerType);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();

                customersPage.ResetFilters();
                customersPage.Filter(FilterType.TypeOfCustomer, customerType);
            }

            if (!customersPage.isPageSizeEqualsTo100())
            {
                customersPage.PageSize("8");
                customersPage.PageSize("100");
            }

            try
            {
                Assert.IsTrue(customersPage.VerifyTypeCustomer(customerType), MessageErreur.FILTRE_ERRONE, customerType);
            }
            finally
            {
                customersPage.PageSize("8");
            }
        }


        // Afficher tous les clients du type Financiero
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Filter_ShowTypes_Financiero()
        {
            // Prepare
            Random rnd = new Random();

            string customerName = customerNameToday + "-" + rnd.Next(1000, 9000).ToString();
            string customerType = "Financiero";

            // Arrange
            HomePage homePage = LogInAsAdmin();


            // Act
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.TypeOfCustomer, customerType);

            if (customersPage.CheckTotalNumber() < 20)
            {
                // Création Customer Financiero
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customerName, rnd.Next(10000, 20000).ToString(), customerType);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();
                customersPage.ResetFilters();
                customersPage.Filter(FilterType.TypeOfCustomer, customerType);
            }

            if (!customersPage.isPageSizeEqualsTo100())
            {
                customersPage.PageSize("8");
                customersPage.PageSize("100");
            }

            try
            {
                Assert.IsTrue(customersPage.VerifyTypeCustomer(customerType), MessageErreur.FILTRE_ERRONE, customerType);
            }
            finally
            {
                customersPage.PageSize("8");
            }
        }

        //____________________________________________________________ FIN_FILTRES___________________________________________________________

        //________________________________________________________________EXPORT_____________________________________________________________

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Export_Results_New_Version()
        {
            // Prepare
            Random rnd = new Random();

            string customerName = customerNameToday + "-" + rnd.Next(1000, 9000).ToString();
            string customerCode = rnd.Next().ToString();
            string customerType = "Colectividades";

            bool newVersionPrint = true;

            // Log in
            HomePage homePage = LogInAsAdmin();

            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();

            customersPage.ClearDownloads();

            customersPage.Filter(FilterType.TypeOfCustomer, customerType);

            if (customersPage.CheckTotalNumber() < 20)
            {
                // Create
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customerName, customerCode, customerType);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();

                customersPage = customerGeneralInformationsPage.BackToList();
            }

            ExportGenerique(customersPage, newVersionPrint);
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Export_Filter_New_Version()
        {
            // Prepare
            Random rnd = new Random();

            string customerName = customerNameToday + "-" + rnd.Next(1000, 9000).ToString();
            string customerCode = rnd.Next().ToString();

            string customerType = TestContext.Properties["CustomerType"].ToString();

            bool newVersionPrint = true;
            DeleteAllFileDownload();
            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();

            customersPage.ClearDownloads();

            if (customersPage.CheckTotalNumber() == 0)
            {
                // Create
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customerName, customerCode, customerType);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();
                customersPage.ResetFilters();
            }
            else
            {
                customerName = customersPage.GetFirstCustomerName();
            }

            customersPage.Filter(FilterType.Search, customerName);

            // Price : onglet customer price indexes ne doit pas être vide
            CustomerGeneralInformationPage customer = customersPage.SelectFirstCustomer();
            CustomerPricePage customerPrice = customer.ClickPriceListTab();
            if (customerPrice.PriceNumber() == 0)
            {
                PriceListPage priceListPage = customerPrice.GoToCustomers_PriceListPage();
                priceListPage.Filter(PriceListPage.FilterType.Site, "MAD");

                if (priceListPage.CheckTotalNumber() > 0)
                {
                    priceListPage.DeleteAllPriceList();
                }

                var pricingCreateModalpage = priceListPage.CreateNewPricing();
                var priceListDetails = pricingCreateModalpage.FillField_CreateNewPricing("TestPrice", "MAD", DateUtils.Now.AddMonths(-1), DateUtils.Now.AddMonths(1));
                customersPage = priceListDetails.GoToCustomers_CustomerPage();

                customer = customersPage.SelectFirstCustomer();
                customerPrice = customer.ClickPriceListTab();
                customerPrice.AddPrice("MAD", "TestPrice");
            }
            customersPage = customerPrice.BackToList();

            ExportGenerique(customersPage, newVersionPrint, customerName);
        }

        private void ExportGenerique(CustomerPage customerPage, bool printValue, string customerName = null)
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            // On exporte les résultats sous la forme d'un fichier Excel (dont on récupère le nom)
            customerPage.ExportCustomer(printValue);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            var correctDownloadedFile = customerPage.GetExportExcelFile(taskFiles, customerName);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

            // Récupération du nom du fichier et construction de l'URL du fichier Excel à ouvrir   
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber(CUSTOMERS, filePath);

            //Assert
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);

            if (!string.IsNullOrEmpty(customerName))
            {
                bool result = OpenXmlExcel.ReadAllDataInColumn("Name", CUSTOMERS, filePath, customerName);
                Assert.IsTrue(result, MessageErreur.EXCEL_DONNEES_KO);
            }
        }
        //____________________________________________________________ FIN_EXPORT___________________________________________________________

        //____________________________________________________________ GENERAL_INFORMATION__________________________________________________

        // Modifier des informations présentes dans l'onglet GeneralInformation
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Modify_General_Informations()
        {
            // Prepare
            string comment = "This is a comment";
            string mailCustomer = "test@test.fr";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var customersPage = homePage.GoToCustomers_CustomerPage();

            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);

            CustomerGeneralInformationPage customerGeneralInformationsPage = customersPage.SelectFirstCustomer();
            customerGeneralInformationsPage.ModifyInformations(comment, mailCustomer);
            customersPage = customerGeneralInformationsPage.BackToList();

            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            customerGeneralInformationsPage = customersPage.SelectFirstCustomer();

            // Assert
            Assert.AreEqual(mailCustomer, customerGeneralInformationsPage.GetCustomerMail(), "Le mail n'est pas celui attendu.");
            Assert.AreEqual(comment, customerGeneralInformationsPage.GetComment(), "Le commentaire n'est pas celui attendu.");
        }


        // l'onglet BuyOnBoard
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Modify_General_Informations_For_BuyOnBoard()
        {
            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage customerGeneralInformationsPage = customersPage.SelectFirstCustomer();

            customerGeneralInformationsPage.ActiveBuyOnBoard();
            var resultat = customerGeneralInformationsPage.IsBoBVisible();
            // Assert
            Assert.IsTrue(resultat, "le buy on bord n'est pas activé");
        }


        //Modification dans l'onglet BuyOnBoard
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Modify_BuyOnBoard()
        {
            // Prepare
            Random rnd = new Random();
            string clientKey = "ClientKey";
            string clientSecret = "ClientSecret";
            string airfi = "Airfi" + rnd.Next().ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage customerGeneralInformationsPage = customersPage.SelectFirstCustomer();

            customerGeneralInformationsPage.ActiveBuyOnBoard();
            CustomerBuyOnBoardPage customerBoBPage = customerGeneralInformationsPage.ClickBobTab();
            customerBoBPage.SelectGuestValueOnBob("BOB");
            customerBoBPage.SetValueOnBob(clientKey, clientSecret, airfi);

            customerBoBPage.SwitchSmarBarAndAirfi();
            customersPage = customerBoBPage.BackToList();

            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            customersPage.WaitPageLoading();
            customerGeneralInformationsPage = customersPage.SelectFirstCustomer();
            customerGeneralInformationsPage.ActiveBuyOnBoard();
            customerGeneralInformationsPage.ActiveBuyOnBoard();

            customerBoBPage = customerGeneralInformationsPage.ClickBobTab();
            var returnValue = customerBoBPage.GetValueOnBob();

            // Assert
            Assert.IsTrue(returnValue.Contains(clientKey));
            Assert.IsTrue(returnValue.Contains(clientSecret));
            Assert.IsTrue(returnValue.Contains(airfi));
        }

        //____________________________________________________________FIN GENERAL_INFORMATION_______________________________________________

        //____________________________________________________________CONTACTS______________________________________________________________

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Create_New_Contact()
        {
            // Prepare
            string contactName = "contact1";
            string mail = "test@test.fr";
            string phone = "+34 123 456 789";
            string site = TestContext.Properties["Site"].ToString();
            string contactResult = contactName + " - " + site.ToString();
            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage customerGeneralInformationsPage = customersPage.SelectFirstCustomer();

            var contactPage = customerGeneralInformationsPage.GoToContactsPage();
            contactPage.CreateContact(contactName, mail, phone, site);

            // Assert
            var isContactDisplayed = contactPage.IsContactDisplayed();
            Assert.IsTrue(isContactDisplayed, "Le contact n'est pas affiché.");
            var contactNameResult = contactPage.GetContactName();
            Assert.AreEqual(contactResult, contactNameResult, "Le contact n'a pas le nom attendu.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Create_Existing_Contact()
        {
            // Prepare
            string contactName = "contact1";
            string mail = "test@test.fr";
            string phone = "+34 123 456 789";
            string site = TestContext.Properties["Site"].ToString();
            string contactResult = contactName + " - " + site.ToString();
            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage customerGeneralInformationsPage = customersPage.SelectFirstCustomer();

            var contactPage = customerGeneralInformationsPage.GoToContactsPage();
            contactPage.WaitPageLoading();
            contactPage.CreateContact(contactName, mail, phone, site);

            // On tente de recréer le même contact
            contactPage.CreateContact(contactName, mail, phone, site);
            var IsErrorMessageDisplayed = contactPage.IsErrorMessageDisplayed();
            Assert.IsTrue(IsErrorMessageDisplayed, "Aucun message d'erreur n'est affiché.");
            contactPage.Cancel();

            // Assert
            var contactNameResult = contactPage.GetContactName();
            Assert.AreEqual(contactResult, contactNameResult, "Le nom du contact ajouté n'est pas correcte");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Modify_Contact()
        {
            // Prepare
            string contactName = "contact1";
            string mail = "test@test.fr";
            string mailBis = "test2@test.fr";
            string phone = "+34 123 456 789";
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage customerGeneralInformationsPage = customersPage.SelectFirstCustomer();

            var contactPage = customerGeneralInformationsPage.GoToContactsPage();
            contactPage.CreateContact(contactName, mail, phone, site);
            contactPage.ClickFirstContact();

            string initMail = contactPage.GetFirstContactMail();

            // On modifie le contact
            contactPage.EditContact(mailBis);
            contactPage.ClickFirstContact();

            string newMail = contactPage.GetFirstContactMail();

            // Assert
            Assert.AreNotEqual(initMail, newMail, "Le mail n'a pas été modifié.");
            Assert.AreEqual(newMail, mailBis, "Le mail attendu ne correspond pas à celui renvoyé.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Delete_Contact()
        {
            // Prepare
            string contactName = "contact1";
            string mail = "test@test.fr";
            string phone = "+34 123 456 789";
            string site = TestContext.Properties["Site"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage customerGeneralInformationsPage = customersPage.SelectFirstCustomer();

            CustomerContactsPage contactPage = customerGeneralInformationsPage.GoToContactsPage();
            contactPage.CreateContact(contactName, mail, phone, site);
            contactPage.ClickFirstContact();

            // On modifie le contact
            contactPage.DeleteContact();
            bool isContactDisplayed = contactPage.IsContactDisplayed();

            // Assert
            Assert.IsFalse(isContactDisplayed, "Le contact n'est pas affiché.");
        }
        //____________________________________________________________FIN CONTACTS__________________________________________________________

        //____________________________________________________________PRICE LIST____________________________________________________________

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Link_Price_List_To_Customer()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();

            string name = "test-" + site;
            DateTime start = DateUtils.Now;
            DateTime end = DateUtils.Now.AddDays(1);

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var priceListGlobalPage = homePage.GoToCustomers_PriceListPage();
            GetPriceList(priceListGlobalPage, site, name, start, end);

            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage customerGeneralInformationsPage = customersPage.SelectFirstCustomer();

            CustomerPriceListPage priceListPage = customerGeneralInformationsPage.GoToPriceListPage();
            priceListPage.CreateNewPrice(name, site);

            // Assert
            bool isPriceDisplayed = priceListPage.IsPriceDisplayed();
            Assert.IsTrue(isPriceDisplayed, "Le prix n'est pas affiché.");
            string firstPriceSite = priceListPage.GetFirstPriceSite();
            Assert.AreEqual(firstPriceSite, site, "Le site associé au prix n'est pas correct.");
            string firstPriceName = priceListPage.GetFirstPriceName();
            Assert.AreEqual(firstPriceName, name, "Le nom associé au prix n'est pas correct.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Link_To_Price_List()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string name = "test-" + site;
            DateTime start = DateUtils.Now;
            DateTime end = DateUtils.Now.AddDays(1);

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var priceListGlobalPage = homePage.GoToCustomers_PriceListPage();
            GetPriceList(priceListGlobalPage, site, name, start, end);

            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage customerGeneralInformationsPage = customersPage.SelectFirstCustomer();

            try
            {
                CustomerPriceListPage priceListPage = customerGeneralInformationsPage.GoToPriceListPage();
                priceListPage.CreateNewPrice(name, site);

                PriceListDetailsPage priceListDetailPage = priceListPage.EditPriceList();
                bool isNewTabAdded = priceListDetailPage.GetTotalTabs();
                var priceName = priceListDetailPage.GetPriceName().ToLower();

                // Assert
                Assert.IsTrue(isNewTabAdded, "Aucun Onglet n'a été ouvert");
                Assert.IsTrue(priceName.Contains(name.ToLower()), "Le nom du prix n'est pas correct.");
                priceListDetailPage.Close();
            }
            finally
            {
                priceListGlobalPage = homePage.GoToCustomers_PriceListPage();
                priceListGlobalPage.Filter(PriceListPage.FilterType.Site, site);
                priceListGlobalPage.Filter(PriceListPage.FilterType.Search, name);
                int priceNumber = priceListGlobalPage.CheckTotalNumber();
                if (priceNumber > 0)
                {
                    priceListGlobalPage.DeletePricing(name);
                }
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Delete_Price_List()
        {
            // Prepare
            Random rnd = new Random();
            string customerName = customerNameToday + "-" + rnd.Next(1000, 9000).ToString();
            string customerCode = rnd.Next().ToString();
            string customerType = TestContext.Properties["CustomerType"].ToString();

            string site = TestContext.Properties["Site"].ToString();
            string name = "test-" + site;
            DateTime start = DateUtils.Now;
            DateTime end = DateUtils.Now.AddDays(1);

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var priceListGlobalPage = homePage.GoToCustomers_PriceListPage();
            // Vérifier qu'un PriceList existe pour les tests, sinon, on le créé
            GetPriceList(priceListGlobalPage, site, name, start, end);
            var customersPage = homePage.GoToCustomers_CustomerPage();

            var customerCreateModalPage = customersPage.CustomerCreatePage();
            customerCreateModalPage.FillFields_CreateCustomerModalPage(customerName, customerCode, customerType);
            var customerGeneralInformationsPage = customerCreateModalPage.Create();

            var priceListPage = customerGeneralInformationsPage.GoToPriceListPage();

            // Creation d'un New price
            priceListPage.CreateNewPrice(name, site);

            priceListPage.DeletePriceList();

            // Assert
            bool priceList = priceListPage.IsPriceDisplayed();
            Assert.IsFalse(priceList, "Le prix est affiché.");
        }


        //__________________________________________________________FIN PRICE LIST__________________________________________________________

        //____________________________________________________________DISCOUNT______________________________________________________________

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Create_New_Discount()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string serviceCategory = TestContext.Properties["ServiceCategory"].ToString();
            string discountRate = "10,000";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage customerGeneralInformationsPage = customersPage.SelectFirstCustomer();

            //Create new discount
            CustomerDiscountPage discountPage = customerGeneralInformationsPage.GoToDiscountPage();
            discountPage.CreateDiscount(site, serviceCategory, discountRate, false);

            // Assert
            bool isDiscountDisplayed = discountPage.IsDiscountDisplayed();
            Assert.IsTrue(isDiscountDisplayed, "Le discount n'est pas affiché.");
            bool isIncrease = discountPage.IsIncrease();
            Assert.IsFalse(isIncrease, "Le discount est négatif.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Create_New_Increase()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string serviceCategory = TestContext.Properties["ServiceCategory"].ToString();
            string discountRate = "10,000";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage customerGeneralInformationsPage = customersPage.SelectFirstCustomer();

            var discountPage = customerGeneralInformationsPage.GoToDiscountPage();
            discountPage.CreateDiscount(site, serviceCategory, discountRate, true);

            // Assert
            var isDiscountDisplayed = discountPage.IsDiscountDisplayed();
            Assert.IsTrue(isDiscountDisplayed, "Le discount n'est pas affiché.");
            var isIncrease = discountPage.IsIncrease();
            Assert.IsTrue(isIncrease, "Le discount est une réduction.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Modify_Discount()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string serviceCategory = TestContext.Properties["Prodman_Needs_ServiceCategory2"].ToString();
            string discountRate = "10";
            string newDiscount = "20";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage customerGeneralInformationsPage = customersPage.SelectFirstCustomer();

            CustomerDiscountPage discountPage = customerGeneralInformationsPage.GoToDiscountPage();

            discountPage.CreateDiscount(site, serviceCategory, discountRate, true);

            string intialDiscountValue = discountPage.GetFirstDiscount().Substring(0, discountPage.GetFirstDiscount().Length - 2);

            discountPage.EditDiscount(newDiscount);

            string newDiscountValue = discountPage.GetFirstDiscount().Substring(0, discountPage.GetFirstDiscount().Length - 2);

            // Assert
            Assert.AreNotEqual(newDiscountValue, intialDiscountValue, "Le discount n'a pas été modifié.");
            Assert.AreEqual(newDiscountValue, newDiscount, "Le discount n'a pas été modifié.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Delete_Discount()
        {
            // Prepare
            string site = TestContext.Properties["Site"].ToString();
            string serviceCategory = TestContext.Properties["ServiceCategory"].ToString();
            string discountRate = "10,000";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage customerGeneralInformationsPage = customersPage.SelectFirstCustomer();

            CustomerDiscountPage discountPage = customerGeneralInformationsPage.GoToDiscountPage();
            discountPage.CreateDiscount(site, serviceCategory, discountRate, true);

            // On vérife que le discount a été créé et apparaît dans la liste
            bool Discount = discountPage.IsDiscountDisplayed();
            Assert.IsTrue(Discount, "le discount n'a pas été créé.");

            // On supprime le discount
            discountPage.DeleteDiscount();
            bool DiscountDeleted = discountPage.IsDiscountDisplayed();

            // Assert : on vérifie que le discount n'apparaît plus dans la liste
            Assert.IsFalse(DiscountDeleted, "Le discount n'a pas été supprimé.");
        }

        //____________________________________________________________FIN DISCOUNT________________________________________________________

        //____________________________________________________________CARDEX______________________________________________________________

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Create_New_Cardex_Notification()
        {
            // Prepare
            string registration = TestContext.Properties["Registration"].ToString();
            string operation = "test operation";
            string invoicing = "Value of invoicing : 10,000";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage customerGeneralInformationsPage = customersPage.SelectFirstCustomer();

            CustomerCardexNotifPage cardexPage = customerGeneralInformationsPage.GoToCardexPage();
            cardexPage.CreateCardex(registration, operation, invoicing);

            // Assert
            bool IsCardexDisplayed = cardexPage.IsCardexDisplayed();
            Assert.IsTrue(IsCardexDisplayed, "Le cardex n'est pas affiché.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Modify_Cardex_Notification()
        {
            // Prepare
            string registration = TestContext.Properties["Registration"].ToString();
            string operation = "test operation";
            string invoicing = "Value of invoicing : 10,000";
            string newInvoicing = "Value of invoicing : 25,000";

            // Arrange
            HomePage homePage = LogInAsAdmin();


            // Act
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage customerGeneralInformationsPage = customersPage.SelectFirstCustomer();

            CustomerCardexNotifPage cardexPage = customerGeneralInformationsPage.GoToCardexPage();
            cardexPage.CreateCardex(registration, operation, invoicing);

            // On récupère la valeur de l'invoicing
            string initialCardexInvoicingValue = cardexPage.GetFirstInvoicing();

            // On modifie la valeur de l'invoicing
            cardexPage.EditCardex(newInvoicing);

            string newCardexInvoicingValue = cardexPage.GetFirstInvoicing();

            // Assert
            Assert.AreNotEqual(newCardexInvoicingValue, initialCardexInvoicingValue, "Le cardex n'a pas été modifié.");
            Assert.AreEqual(newCardexInvoicingValue, newInvoicing, "Le cardex n'a pas été modifié.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Delete_Cardex_Notification()
        {
            // Prepare
            string registration = TestContext.Properties["Registration"].ToString();
            string operation = "test operation";
            string invoicing = "Value of invoicing : 10,000";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage customerGeneralInformationsPage = customersPage.SelectFirstCustomer();

            CustomerCardexNotifPage cardexPage = customerGeneralInformationsPage.GoToCardexPage();
            cardexPage.CreateCardex(registration, operation, invoicing);

            // On vérifie que l'invoicing a été créé
            Assert.IsTrue(cardexPage.IsCardexDisplayed(), "Le cardex n'est pas affiché.");

            // On supprime la valeur de l'invoicing
            cardexPage.DeleteCardex();

            // Assert
            Assert.IsFalse(cardexPage.IsCardexDisplayed(), "Le cardex n'a pas été supprimé.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Search_Cardex_By_Registration()
        {
            // Prepare
            string registration = TestContext.Properties["Registration"].ToString();
            string registrationBis = "B782";
            string operation = "test operation";
            string invoicing = "Value of invoicing : 10,000";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage customerGeneralInformationsPage = customersPage.SelectFirstCustomer();

            CustomerCardexNotifPage cardexPage = customerGeneralInformationsPage.GoToCardexPage();
            cardexPage.CreateCardex(registration, operation, invoicing);

            // On vérifie que l'invoicing a été créé
            Assert.IsTrue(cardexPage.IsCardexDisplayed(), "Le cardex n'est pas affiché.");

            cardexPage.CreateCardex(registrationBis, operation, invoicing);

            // On filtre les CARDEX par Registration
            cardexPage.Filter(CardexFilterType.Search, registration);

            // Assert
            string firstRegistration = cardexPage.GetFirstRegistration();
            string firstInvoicing = cardexPage.GetFirstInvoicing();

            Assert.AreEqual(firstRegistration, registration, "La registration n'est pas correcte.");
            Assert.AreEqual(firstInvoicing, invoicing, "L'invoicing n'a pas été modifié.");
        }

        //____________________________________________________________FIN CARDEX________________________________________________________

        //____________________________________________________________Agreement number__________________________________________________

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Create_New_Agreement()
        {
            // Prepare
            string agreementNumber = new Random().Next(1000, 9000).ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage customerGeneralInformationsPage = customersPage.SelectFirstCustomer();

            //create new agreement
            CustomerAgreementPage agreementPage = customerGeneralInformationsPage.GoToAgreementPage();
            agreementPage.CreateAgreement(agreementNumber);

            // Assert
            bool isAgreementDisplayed = agreementPage.IsAgreementDisplayed();
            Assert.IsTrue(isAgreementDisplayed, "L'agreement n'est pas affiché.");
        }


        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Create_Existing_Agreement()
        {
            // Prepare
            string agreementNumber = new Random().Next(1000, 9000).ToString();
            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act       
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage customerGeneralInformationsPage = customersPage.SelectFirstCustomer();

            //create an agreement
            CustomerAgreementPage agreementPage = customerGeneralInformationsPage.GoToAgreementPage();
            agreementPage.CreateAgreement(agreementNumber);

            // On vérifie que l'agreement est bien affiché
            string firstAgreementNumber = agreementPage.GetFirstAgreementNumber();
            Assert.AreEqual(firstAgreementNumber, agreementNumber, "L'agreement number est incorrect.");

            // On tente de recréer le même agreement
            agreementPage.CreateAgreement(agreementNumber);

            // Assert
            bool isErrorMessageDisplayed = agreementPage.IsErrorMessageDisplayed();
            Assert.IsTrue(isErrorMessageDisplayed, "Le message d'erreur n'est pas affiché");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Modify_Agreement()
        {
            // Prepare
            string agreementNumber = new Random().Next(1000, 9000).ToString();
            string agreementNumberBis = agreementNumber + "bis";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage customerGeneralInformationsPage = customersPage.SelectFirstCustomer();

            // Create New Agreement
            CustomerAgreementPage agreementPage = customerGeneralInformationsPage.GoToAgreementPage();
            agreementPage.CreateAgreement(agreementNumber);
            agreementPage.ClickFirstAgreement();
            string initNumber = agreementPage.GetFirstAgreementNumber();

            // Modifiy existing Agreement
            agreementPage.EditAgreement(agreementNumberBis);
            agreementPage.ClickFirstAgreement();
            string newAgreement = agreementPage.GetFirstAgreementNumber();

            // Assert
            Assert.AreNotEqual(initNumber, newAgreement, "L'agreement n'a pas été modifié.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Delete_Agreement()
        {
            // Prepare
            string agreementNumber = new Random().Next(1000, 9000).ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage customerGeneralInformationsPage = customersPage.SelectFirstCustomer();

            CustomerAgreementPage agreementPage = customerGeneralInformationsPage.GoToAgreementPage();
            agreementPage.CreateAgreement(agreementNumber);
            agreementPage.ClickFirstAgreement();

            // On supprime l'agreement
            agreementPage.DeleteAgreement();

            // Assert
            bool isAgreementDisplayed = agreementPage.IsAgreementDisplayed();
            Assert.IsFalse(isAgreementDisplayed, "L'agreement n'a pas été supprimé");
        }

        //____________________________________________________________FIN Agreement number______________________________________________

        //____________________________________________________________Reinvoice_________________________________________________________

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Create_New_Reinvoice()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteBis"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage customerGeneralInformationsPage = customersPage.SelectFirstCustomer();

            CustomerReinvoicePage reinvoicePage = customerGeneralInformationsPage.GoToReinvoicePage();

            reinvoicePage.CreateReinvoice(siteFrom, siteTo);

            // Assert
            bool isReinvoiceDisplayed = reinvoicePage.IsReinvoiceDisplayed();
            Assert.IsTrue(isReinvoiceDisplayed, "Le reinvoice n'a pas été créé ou n'est pas visible.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Create_Existing_Reinvoice()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteBis"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage customerGeneralInformationsPage = customersPage.SelectFirstCustomer();

            CustomerReinvoicePage reinvoicePage = customerGeneralInformationsPage.GoToReinvoicePage();
            reinvoicePage.CreateReinvoice(siteFrom, siteTo);

            // On vérifie que le reinvoice a été créé
            bool isReinvoiceDisplayed = reinvoicePage.IsReinvoiceDisplayed();
            Assert.IsTrue(isReinvoiceDisplayed, "Le reinvoice n'a pas été créé ou n'est pas visible.");

            reinvoicePage.CreateReinvoice(siteFrom, siteTo);

            // Assert
            bool isErrorMessageDisplayed = reinvoicePage.IsErrorMessageDisplayed();
            Assert.IsTrue(isErrorMessageDisplayed, "Aucun message d'erreur n'est affiché.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Create_Reinvoice_Same_From_And_To_Site()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage customerGeneralInformationsPage = customersPage.SelectFirstCustomer();

            CustomerReinvoicePage reinvoicePage = customerGeneralInformationsPage.GoToReinvoicePage();
            reinvoicePage.CreateReinvoice(siteFrom, siteFrom);

            // Assert
            bool isErrorMessageDisplayed = reinvoicePage.IsErrorMessageDisplayed();
            Assert.IsTrue(isErrorMessageDisplayed, "Le message d'erreur n'est pas affiché.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Modify_Reinvoice()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteBis"].ToString();
            string autreSite = TestContext.Properties["SitePriceList"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage customerGeneralInformationsPage = customersPage.SelectFirstCustomer();

            CustomerReinvoicePage reinvoicePage = customerGeneralInformationsPage.GoToReinvoicePage();

            reinvoicePage.CreateReinvoice(siteFrom, siteTo);

            // On récupère la valeur de l'invoicing toSite
            var initialReinvoiceToSite = reinvoicePage.GetFirstReinvoiceToSite();

            // On modifie la valeur de l'invoicing
            reinvoicePage.EditReinvoice(autreSite);

            var newReinvoiceToSite = reinvoicePage.GetFirstReinvoiceToSite();

            // Assert
            Assert.AreNotEqual(newReinvoiceToSite, initialReinvoiceToSite, "Le reinvoice n'a pas été modifié.");
            Assert.AreEqual(newReinvoiceToSite, autreSite, "Le reinvoice n'a pas été modifié.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Delete_Reinvoice()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteBis"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage customerGeneralInformationsPage = customersPage.SelectFirstCustomer();

            CustomerReinvoicePage reinvoicePage = customerGeneralInformationsPage.GoToReinvoicePage();
            reinvoicePage.CreateReinvoice(siteFrom, siteTo);

            // On vérifie que l'invoicing est affiché
            Assert.IsTrue(reinvoicePage.IsReinvoiceDisplayed(), "Le reinvoice n'est pas affiché.");

            // On supprime l'invoicing
            reinvoicePage.DeleteReinvoice();

            // On vérifie que l'invoicing est effacé
            Assert.IsFalse(reinvoicePage.IsReinvoiceDisplayed(), "Le reinvoice est affiché.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Search_Reinvoice_By_Site()
        {
            // Prepare
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteBis"].ToString();
            string site2From = TestContext.Properties["SiteACE"].ToString();
            string site2To = TestContext.Properties["SiteToFlightBob"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage customerGeneralInformationsPage = customersPage.SelectFirstCustomer();

            var reinvoicePage = customerGeneralInformationsPage.GoToReinvoicePage();

            // Création d'un premier reinvoice
            reinvoicePage.CreateReinvoice(siteFrom, siteTo);

            // Création d'un second reinvoice
            reinvoicePage.CreateReinvoice(site2From, site2To);

            Assert.AreEqual(2, reinvoicePage.GetReinvoiceNumber(), "Le nombre de reinvoice est incorrect.");

            // On filtre par SiteFrom
            reinvoicePage.Filter(ReinvoiceFilterType.Search, siteFrom);

            Assert.AreEqual(1, reinvoicePage.GetReinvoiceNumber(), "Le nombre de reinvoice est incorrect.");
            Assert.AreEqual(siteFrom, reinvoicePage.GetFirstReinvoiceFromSite(), "Le site de reinvoice est incorrect.");

            // On filtre par SiteTo
            reinvoicePage.Filter(ReinvoiceFilterType.Search, site2To);
            Assert.AreEqual(1, reinvoicePage.GetReinvoiceNumber(), "Le nombre de reinvoice est incorrect.");
            Assert.AreEqual(reinvoicePage.GetFirstReinvoiceToSite(), site2To, "Le site de reinvoice est incorrect.");

        }
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Create_CustomerBob()
        {
            // Prepare
            Random rnd = new Random();
            string organizationName = $"Org-{rnd.Next().ToString()}";
            string siteACE = TestContext.Properties["SiteACE"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();
            // Act
            ParametersSites parametersSitesPage = homePage.GoToParameters_Sites();
            parametersSitesPage.Filter(ParametersSites.FilterType.SearchSite, siteACE);
            parametersSitesPage.ClickOnFirstSite();
            var site = parametersSitesPage.GetFirstSiteName();
            parametersSitesPage.ClickToOrganization();
            parametersSitesPage.CreateNewOrganization(organizationName);

            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage customer = customersPage.SelectFirstCustomer();

            if (!customer.IsBoBVisible())
            {
                customer.ActiveBuyOnBoard();
            }
            var customerBoBTab = customer.ClickBobTab();
            var resultat = customerBoBTab.OrganizationIsExist(site, organizationName);

            //Assert
            Assert.IsTrue(resultat, "la nouvelle organisation  n'apparaît pas");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CU_CustomerIndexSearch_Showactiveonly()
        {
            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            //Filter show only active customer
            customersPage.Filter(FilterType.ShowActive, true);

            //Assert
            //Search Customer Active
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            int nbResult = customersPage.CheckTotalNumber();
            Assert.AreEqual(1, nbResult, String.Format(MessageErreur.FILTRE_ERRONE, "'Show only active'"));

            //Assert
            //Search Customer Active Inactive
            customersPage.Filter(FilterType.Search, customerNameTodayInactive);
            int nbResult1 = customersPage.CheckTotalNumber();
            Assert.AreEqual(0, nbResult1, String.Format(MessageErreur.FILTRE_ERRONE, "'Show only active'"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CU_CustomersGeneralInformation_SIRET()
        {
            // Prepare
            string siret = "10";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage generalInformationPage = customersPage.SelectFirstCustomer();

            generalInformationPage.SetSiret(siret);
            WebDriver.Navigate().Refresh();

            //Récupérer la valeur du Siret
            string newSiret = generalInformationPage.GetSiret();

            //Assert
            Assert.AreEqual(newSiret, siret, " La saisie du champ Siret n'était pas effectuée correctement.");


        }

        //____________________________________________________________FIN Reinvoice_____________________________________________________

        //____________________________________________________________Utilitaire________________________________________________________

        public void GetPriceList(PriceListPage priceListPage, string site, string name, DateTime start, DateTime end)
        {
            // Act
            priceListPage.ResetFilter();
            priceListPage.Filter(PriceListPage.FilterType.Site, site);
            priceListPage.Filter(PriceListPage.FilterType.Search, name);

            if (priceListPage.CheckTotalNumber() == 0)
            {
                var priceListCreatePage = priceListPage.CreateNewPricing();
                priceListCreatePage.FillField_CreateNewPricing(name, site, start, end);
            }

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CU_TaxOffice()
        {
            int i = (int)TestContext.Properties[string.Format("CustomerId")];
            // Prepare
            string taxOffice = "10";
            // Arrange
            HomePage homePage = LogInAsAdmin();
            // Act
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage generalInformationPage = customersPage.SelectFirstCustomer();
            generalInformationPage.SetTaxOffice(taxOffice);
            WebDriver.Navigate().Refresh();
            //Récupérer la valeur du TaxOffice
            string newTaxOffice = generalInformationPage.GetTaxOffice();
            //Assert
            Assert.AreEqual(newTaxOffice, taxOffice, " La saisie du champ TaxOffice n'était pas effectuée correctement.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CU_ProcessingDuration()
        {
            // Prepare
            string processingDuration = "5";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage generalInformationPage = customersPage.SelectFirstCustomer();
            generalInformationPage.SetProcessingDuration(processingDuration);

            WebDriver.Navigate().Refresh();
            //Récupérer la valeur du ProcessingDuration
            string newProcessingDuration = generalInformationPage.GetProcessingDuration();

            //Assert
            Assert.AreEqual(newProcessingDuration, processingDuration, " La saisie du champ ProcessingDuration n'était pas effectuée correctement.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CU_CustomersIndexSearch_ShowInactiveOnly()
        {
            //Prepare
            bool value = true;

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            //Filter show only inactive customer
            customersPage.Filter(FilterType.ShowInactive, value);

            //Assert
            //Search Customer Active
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            int nbResult = customersPage.CheckTotalNumber();
            Assert.AreEqual(0, nbResult, String.Format(MessageErreur.FILTRE_ERRONE, "'Show only inactive'"));

            //Assert
            //Search Customer Active Inactive
            customersPage.Filter(FilterType.Search, customerNameTodayInactive);
            int nbResult1 = customersPage.CheckTotalNumber();
            Assert.AreEqual(1, nbResult1, String.Format(MessageErreur.FILTRE_ERRONE, "'Show only inactive'"));

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CU_CustomerIndex_SearchShowAll()
        {
            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            //Filter show all customer
            customersPage.Filter(FilterType.ShowAll, true);

            //Assert
            //Search Customer Active
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            int nbResult = customersPage.CheckTotalNumber();
            Assert.AreEqual(1, nbResult, string.Format(MessageErreur.FILTRE_ERRONE, "'Show All Customers'"));

            //Assert
            //Search Customer Inactive
            customersPage.Filter(FilterType.Search, customerNameTodayInactive);
            int nbResult1 = customersPage.CheckTotalNumber();
            Assert.AreEqual(1, nbResult1, string.Format(MessageErreur.FILTRE_ERRONE, "'Show All Customers'"));

        }


        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CU_CustomersGeneralInformation_RevisionPrice()
        {
            // Prepare
            DateTime revisionPrice = DateTime.Now;

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage generalInformationPage = customersPage.SelectFirstCustomer();
            generalInformationPage.SetRevisionPrice(revisionPrice);

            WebDriver.Navigate().Refresh();
            //Récupérer la valeur du RevisionPrice
            string newRevisionPrice = generalInformationPage.GetRevisionPrice();

            //Assert
            Assert.AreEqual(newRevisionPrice, revisionPrice.ToString("dd/MM/yyyy"), " La saisie du champ ProcessingDuration n'était pas effectuée correctement.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CU_CustomerGeneralInformation_ThirdId()
        {
            // Prepare
            string thirdId = "83C";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage generalInformationPage = customersPage.SelectFirstCustomer();
            generalInformationPage.SetThirdId(thirdId);

            WebDriver.Navigate().Refresh();
            //Récupérer la valeur du RevisionPrice
            string newThirdId = generalInformationPage.GetThirdId();

            //Assert
            Assert.AreEqual(newThirdId, thirdId, " La saisie du champ ProcessingDuration n'était pas effectuée correctement.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CU_CustomerIndex_Searchtext()
        {
            //Prepare 
            int toCheck = 0;
            // Arrange
            HomePage homePage = LogInAsAdmin();

            //Create another customer 
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            //Filter search by
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            string firstCustomerName = customersPage.GetFirstCustomerName();

            //Assert
            bool isFirstNameStartOk = firstCustomerName.Contains(customerNameTodayActive);
            Assert.IsTrue(isFirstNameStartOk, "Le customer n'a pas commencé par le même nom de domaine créé.");

            //Assert
            int nbResult = customersPage.CheckTotalNumber();
            Assert.AreNotEqual(toCheck, nbResult, String.Format(MessageErreur.FILTRE_ERRONE, "'Au moins un Customer doit remonter dans les résultats de la recherche'"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CU_VATFree()
        {
            // Prepare
            bool isVATFree = false;

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage generalInformationPage = customersPage.SelectFirstCustomer();
            isVATFree = generalInformationPage.IsVATFREE();
            if (isVATFree != false)
            {
                generalInformationPage.SetVATfree(false);
            }
            generalInformationPage.SetVATfree(true);
            WebDriver.Navigate().Refresh();
            //Récupérer la valeur du RevisionPrice
            isVATFree = generalInformationPage.IsVATFREE();
            //Assert
            Assert.IsTrue(isVATFree, "Set VATfee To Yes ne fonctionne pas correctement.");

            generalInformationPage.SetVATfree(false);
            WebDriver.Navigate().Refresh();
            //Récupérer la valeur du RevisionPrice
            isVATFree = generalInformationPage.IsVATFREE();
            //Assert
            Assert.IsFalse(isVATFree, "Set VATfee To No ne fonctionne pas correctement.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CU_GeneralInformation_ProformaAddress()
        {
            // Prepare
            bool isProFormaAddress = false;
            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage generalInformationPage = customersPage.SelectFirstCustomer();

            isProFormaAddress = generalInformationPage.IsProFormaAddress();
            if (isProFormaAddress != false)
            {
                generalInformationPage.SetProFormaAddress(false);
            }
            generalInformationPage.SetProFormaAddress(true);
            WebDriver.Navigate().Refresh();
            isProFormaAddress = generalInformationPage.IsProFormaAddress();
            //Assert
            Assert.IsTrue(isProFormaAddress, "Set Pro Forma Address To Yes ne fonctionne pas correctement.");

            generalInformationPage.SetProFormaAddress(false);
            WebDriver.Navigate().Refresh();
            isProFormaAddress = generalInformationPage.IsProFormaAddress();
            //Assert
            Assert.IsFalse(isProFormaAddress, " Set Pro Forma Address To No ne fonctionne pas correctement.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void CU_CU_Modify_Generalnfo_DeliveryNoteComment()
        {
            //Arrange 
            HomePage homePage = LogInAsAdmin();
            // Act
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage generalInformationPage = customersPage.SelectFirstCustomer();
            // modifier commentaire 
            generalInformationPage.SetCustomerDeliveryNoteComment(comment);
            generalInformationPage.UpToMenueTabs();

            generalInformationPage.GoToDiscountPage();
            generalInformationPage.GoToGeneralInformation();

            string commentFromInterface = generalInformationPage.GetCustomerDeliveryNoteComment();
            //assert 
            Assert.AreEqual(commentFromInterface, comment, $"La modifivation du champ Note Comment :{comment} n'est pas enregistrée");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CU_Modify_Generalnfo_InvoiceComment()
        {
            //Prepare
            string comment = "Customer_Comment";
            //Arranage 
            HomePage homePage = LogInAsAdmin();
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage generalInformationPage = customersPage.SelectFirstCustomer();
            //modifier commentaire
            generalInformationPage.SetComment(comment);
            generalInformationPage.UpToMenueTabs();
            CustomerPricePage customerPrice = generalInformationPage.ClickPriceListTab();
            generalInformationPage.GoToGeneralInformation();
            string _modifiedComment = generalInformationPage.GetComment();
            //Assert
            Assert.AreEqual(comment, _modifiedComment, "La modification du commentaire n'a pas été effectuée correctement.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CU_Modify_Generalnfo_ICAO()
        {

            Random random = new Random();
            int randomNumber = random.Next(100000000, 1000000000);
            string _icao = randomNumber.ToString();


            HomePage homePage = LogInAsAdmin();
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage generalInformationPage = customersPage.SelectFirstCustomer();
            var icao = generalInformationPage.GetIcao();
            generalInformationPage.ModifyICAO(_icao);
            generalInformationPage.UpToMenueTabs();
            CustomerPricePage customerPrice = generalInformationPage.ClickPriceListTab();
            generalInformationPage.GoToGeneralInformation();
            var modifiedIcao = generalInformationPage.GetIcao();
            Assert.AreNotEqual(icao, modifiedIcao, "La modification du ICAO n'a pas été effectuée correctement.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CU_GeneralInformation_Name()
        {
            // Prepare
            string nameModified = "name was updated !!" + random.Next().ToString();
            // Arrange
            HomePage homePage = LogInAsAdmin();
            // Act
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage generalInformationPage = customersPage.SelectFirstCustomer();
            generalInformationPage.SetCustomerName(nameModified);
            WebDriver.Navigate().Refresh();
            string updatedName = generalInformationPage.GetCustomerName();
            // Assert
            Assert.AreEqual(nameModified, updatedName, "Le nom du client devrait correspondre à la valeur mise à jour.");
        }
        [DeploymentItem("Resources\\logo.jpg")]
        [DeploymentItem("chromedriver.exe")]
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CU_Modify_Generalnfo_Logo()
        {
            //Arrange
            HomePage homePage = LogInAsAdmin();
            //Act 
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage generalInformationPage = customersPage.SelectFirstCustomer();
            FileInfo fiUpload = new FileInfo(TestContext.TestDeploymentDir + "\\logo.jpg");
            generalInformationPage.UploadLogo(fiUpload);
            CustomerPricePage customerPrice = generalInformationPage.ClickPriceListTab();
            generalInformationPage.GoToGeneralInformation();
            //Assert
            bool isImageExist= generalInformationPage.IsImageExist();
            Assert.IsTrue(isImageExist, "La modification de l'image n'a pas été effectuée correctement");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CU_GeneralInformation_PaymentTerm()
        {
            // Prepare
            string paymentToSet = "60 días fecha factura";
            // Arrange
            HomePage homePage = LogInAsAdmin();
            // Act
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage generalInformationPage = customersPage.SelectFirstCustomer();
            generalInformationPage.SetPaymentTerm(paymentToSet);
            generalInformationPage.BackToList();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            customersPage.SelectFirstCustomer();
            string displayedPayment = generalInformationPage.GetPaymentTerm();
            // Assert
            Assert.AreEqual(paymentToSet, displayedPayment, "Le Payment Term du client devrait correspondre à la valeur mise à jour.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void CU_CU_Modify_Generalnfo_DeliveryNoteValorization()
        {
            //Arrange 
            HomePage homePage = LogInAsAdmin();
            // Act
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);            
            CustomerGeneralInformationPage generalInformationPage = customersPage.SelectFirstCustomer();
            bool deliveryNoteValorization = generalInformationPage.IsDeliveryNoteValorization();
            //modifier deliveryNoteValorization
            generalInformationPage.SetDeliveryNoteValorization(!deliveryNoteValorization);
            generalInformationPage.GoToDiscountPage();
            generalInformationPage.GoToGeneralInformation();
            bool isUpdated = generalInformationPage.IsDeliveryNoteValorization() != deliveryNoteValorization;
            //assert 
            Assert.IsTrue(isUpdated, "La modification du Delivery Note Valorization n'a pas été effectuée correctement.");
        }

        [TestMethod()]
        [Timeout(_timeout)]
        public void CU_CU_Modify_Generalnfo_EngagementNo()
        {
            string engagementNo = random.Next().ToString();
            HomePage homePage = LogInAsAdmin();
            // Act
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage generalInformationPage = customersPage.SelectFirstCustomer();
            generalInformationPage.SetEngagementNo(engagementNo);
            generalInformationPage.UpToMenueTabs();

            generalInformationPage.GoToDiscountPage();
            generalInformationPage.GoToGeneralInformation();
            bool isUpdated = generalInformationPage.GetEngagementNo() == engagementNo;
            Assert.IsTrue(isUpdated, "La modification du Engagement No n'a pas été effectuée correctement.");
        }

        [TestMethod()]
        [Timeout(_timeout)]
        public void CU_CUST_Modify_Generalnfo_ExternalId()
        {
            string externalId = random.Next().ToString();
            HomePage homePage = LogInAsAdmin();
            // Act
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage generalInformationPage = customersPage.SelectFirstCustomer();
            generalInformationPage.SetExternalid(externalId);
            generalInformationPage.UpToMenueTabs();

            generalInformationPage.GoToDiscountPage();
            generalInformationPage.GoToGeneralInformation();
            bool isUpdated = generalInformationPage.GetExternalId() == externalId;
            Assert.IsTrue(isUpdated, "La modification du External Identifier n'a pas été effectuée correctement.");
        }

        [TestMethod()]
        [Timeout(_timeout)]
        public void CU_CU_Modify_Generalnfo_Filigran()
        {
            HomePage homePage = LogInAsAdmin();
            // Act
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage generalInformationPage = customersPage.SelectFirstCustomer();
            bool filigran = generalInformationPage.IsFiligran();
            generalInformationPage.SetFiligran(!filigran);
            generalInformationPage.UpToMenueTabs();

            generalInformationPage.GoToDiscountPage();
            generalInformationPage.GoToGeneralInformation();
            bool isUpdated = generalInformationPage.IsFiligran() != filigran;
            Assert.IsTrue(isUpdated, "La modification du Filigran n'a pas été effectuée correctement.");
        }

        [TestMethod()]
        [Timeout(_timeout)]
        public void CU_CU_Modify_Generalnfo_IATA()
        {
            string iata = random.Next().ToString();
            HomePage homePage = LogInAsAdmin();
            // Act
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage generalInformationPage = customersPage.SelectFirstCustomer();
            generalInformationPage.SetIata(iata);
            generalInformationPage.GoToDiscountPage();
            generalInformationPage.GoToGeneralInformation();
            bool isUpdated = generalInformationPage.GetIata() == iata;
            Assert.IsTrue(isUpdated, "La modification du IATA n'a pas été effectuée correctement.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_CUST_Modify_Generalnfo_LastUpdate()
        {

            Random random = new Random();
            int randomNumber = random.Next(1, 10);
            string lastUpdateDays = randomNumber.ToString();

            HomePage homePage = LogInAsAdmin();
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(FilterType.Search, customerNameTodayActive);
            CustomerGeneralInformationPage generalInformationPage = customersPage.SelectFirstCustomer();
            var lastUpdate = generalInformationPage.GetLasUpdate();
            generalInformationPage.SetProcessingDuration(lastUpdateDays);
            generalInformationPage.SetLasUpdate(lastUpdateDays);
            generalInformationPage.UpToMenueTabs();
            CustomerPricePage customerPrice = generalInformationPage.ClickPriceListTab();
            generalInformationPage.GoToGeneralInformation();
            var modifiedLastUpdate = generalInformationPage.GetLasUpdate();
            Assert.AreNotEqual(lastUpdate, modifiedLastUpdate, "La modification du LAST UPDATE n'a pas été effectuée correctement.");
        }




        ////////////////////////////////Insert Customer////////////////////////////////
        ///
        private int InsertCustomer(string code, string name, string customerType, int isActive = 1, string country = "Spain", string currency = "Euro")
        {
            string query = @"
            DECLARE @customerTypeId INT;
            SELECT TOP 1 @customerTypeId = Id FROM CustomerTypes WHERE Name LIKE @customerType;

            DECLARE @countryId INT;
            SELECT TOP 1 @countryId = Id FROM Countries WHERE Name LIKE @country;

            DECLARE @currencyId INT;
            SELECT TOP 1 @currencyId = Id FROM Currencies WHERE Name LIKE @currency;

            Insert into [dbo].[Customers] (
		        [Code]
              ,[Name]
              ,[CustomerTypeId]
              ,[CountryId]
              ,[CurrencyId]
              ,[IsVATFree]
              ,[IsValorizationDeliveryNote]
              ,[IsProFormaAddress]
              ,[IsAirportTax]
              ,[IsActive]
              ,[PaymentTypeId]
              ,[PrintLabelFormat]
              ,[COShippingDuration]
              ,[COProcessingDuration]
              ,[COLastUpdate])

	          values 
	          (
	          @code,
	          @name,
	          @customerTypeId,
	          @countryId,
	          @currencyId,
	          0,
	          0,
	          0,
	          1,
	          @isActive,
	          '1',
	          '0',
	          '0',
	          '0',
	          '0'
	          );
             SELECT SCOPE_IDENTITY();
            "
            ;
            return ExecuteAndGetInt(
            query,
                new KeyValuePair<string, object>("code", code),
                new KeyValuePair<string, object>("name", name),
                new KeyValuePair<string, object>("country", country),
                new KeyValuePair<string, object>("currency", currency),
                new KeyValuePair<string, object>("customerType", customerType),
                new KeyValuePair<string, object>("isActive", isActive)
            );
        }

        private int? CustomerExist(int id)
        {
            string query = @"
             select Id from Customers where Id = @id

              SELECT SCOPE_IDENTITY();
            "
            ;
            int? result = ExecuteAndGetInt(
            query,
            new KeyValuePair<string, object>("id", id)
            );
            return result == 0 ? (int?)null : result;
        }


        private void DeleteCustomer(int id)
        {
            string query = @"
            delete from Customers where id = @id
            "
            ;
            ExecuteNonQuery(
            query,
                new KeyValuePair<string, object>("id", id)
            );
        }



    }
}