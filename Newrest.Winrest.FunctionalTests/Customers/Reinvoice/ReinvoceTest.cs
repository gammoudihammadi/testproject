using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObject.Parameters.Customer;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Accounting.Invoice;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Customer;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Delivery;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Reinvoice;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.Flight;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.Dispatch;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using System;
using System.IO;
using System.Threading;
using System.Web;

namespace Newrest.Winrest.FunctionalTests.Customers
{
    [TestClass]
    public class ReinvoceTest : TestBase
    {
        private const int _timeout = 600000; 
        readonly string customerNameToday = "CustomerTest-" + DateUtils.Now.ToString("dd/MM/yyyy") + "-";
        readonly string serviceNameToday = "ServiceTest-" + DateUtils.Now.ToString("dd/MM/yyyy") + "-";

        private const string REINVOICES_EXCEL_SHEET_NAME = "Reinvoices";

        //Effectuer des recherches via les filtres (customer From)
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_RINV_Filter_Customers_From()
        {
            string site = TestContext.Properties["Site"].ToString();
            string customerType = TestContext.Properties["CustomerType"].ToString();

            Random rnd = new Random();

            // Prepare customer
            int rndCustomer = rnd.Next();
            int rndCustomer1 = rnd.Next();

            string customerName = customerNameToday + rndCustomer.ToString();
            string customerCode = rndCustomer.ToString();
            string customerName1 = customerNameToday + rndCustomer1.ToString();
            string customerCode1 = rndCustomer1.ToString();

            // Prepare service
            string serviceName = serviceNameToday + rndCustomer.ToString();
            string serviceName1 = serviceNameToday + rndCustomer1.ToString();

            DateTime fromDate = DateUtils.Now;
            DateTime toDate = DateUtils.Now.AddDays(10);

            // Arrange
            HomePage homePage = LogInAsAdmin();

            var reinvoicePage = homePage.GoToCustomers_ReinvoicePage();

            if (reinvoicePage.CheckTotalNumber() == 0)
            {
                // Create new customer
                var customersPage = homePage.GoToCustomers_CustomerPage();
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customerName, customerCode, customerType);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();

                customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customerName1, customerCode1, customerType);
                customerCreateModalPage.Create();

                // Create new service
                var servicePage = homePage.GoToCustomers_ServicePage();
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                priceModalPage.FillFields_CustomerPrice(site, customerName, fromDate, toDate);
                servicePage = serviceGeneralInformationsPage.BackToList();

                serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName1);
                serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                pricePage = serviceGeneralInformationsPage.GoToPricePage();
                priceModalPage = pricePage.AddNewCustomerPrice();
                priceModalPage.FillFields_CustomerPrice(site, customerName1, fromDate, toDate);

                reinvoicePage = homePage.GoToCustomers_ReinvoicePage();
                var reinvoiceCreateModalpage = reinvoicePage.CreateNewReinvoice();
                reinvoiceCreateModalpage.FillFields_CreateReinvoiceModalPage(customerName1, customerName);
                reinvoicePage = reinvoiceCreateModalpage.Create();
                reinvoicePage.ResetFilter();
            }
            else
            {
                customerName1 = reinvoicePage.GetFirstCustomerFromName();
                customerCode1 = reinvoicePage.GetFirstCustomerFromCode();
            }

            var customerFrom = customerCode1 + " - " + customerName1;

            reinvoicePage.Filter(ReinvoicePage.FilterType.SearchCustomersFrom, customerFrom);

            if (!reinvoicePage.isPageSizeEqualsTo100())
            {
                reinvoicePage.PageSize("8");
                reinvoicePage.PageSize("100");
            }

            Assert.IsTrue(reinvoicePage.VerifyCustomerFrom(customerName1), string.Format(MessageErreur.FILTRE_ERRONE, "CustomersFrom"));
        }

        //Effectuer des recherches via les filtres (customer To)
        [TestMethod]
        [Timeout(_timeout)]
        public void CU_RINV_Filter_Customers_To()
        {
            string site = TestContext.Properties["Site"].ToString();
            string customerType = TestContext.Properties["CustomerType"].ToString();

            Random rnd = new Random();

            // Prepare customer
            int rndCustomer = rnd.Next();
            int rndCustomer1 = rnd.Next();

            string customerName = customerNameToday + rndCustomer.ToString();
            string customerCode = rndCustomer.ToString();
            string customerName1 = customerNameToday + rndCustomer1.ToString();
            string customerCode1 = rndCustomer1.ToString();

            // Prepare service
            string serviceName = serviceNameToday + rndCustomer.ToString();
            string serviceName1 = serviceNameToday + rndCustomer1.ToString();

            DateTime fromDate = DateUtils.Now;
            DateTime toDate = DateUtils.Now.AddDays(10);

            // Arrange
            var homePage = LogInAsAdmin();


            var reinvoicePage = homePage.GoToCustomers_ReinvoicePage();

            if (reinvoicePage.CheckTotalNumber() == 0)
            {
                // Create new customer
                var customersPage = homePage.GoToCustomers_CustomerPage();
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customerName, customerCode, customerType);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();

                customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customerName1, customerCode1, customerType);
                customerCreateModalPage.Create();

                // Create new service
                var servicePage = homePage.GoToCustomers_ServicePage();
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customerName, fromDate, toDate);
                servicePage = pricePage.BackToList();

                serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName1);
                serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                pricePage = serviceGeneralInformationsPage.GoToPricePage();
                priceModalPage = pricePage.AddNewCustomerPrice();
                priceModalPage.FillFields_CustomerPrice(site, customerName1, fromDate, toDate);

                reinvoicePage = homePage.GoToCustomers_ReinvoicePage();
                var reinvoiceCreateModalpage = reinvoicePage.CreateNewReinvoice();
                reinvoiceCreateModalpage.FillFields_CreateReinvoiceModalPage(customerName1, customerName);
                reinvoicePage = reinvoiceCreateModalpage.Create();
            }
            else
            {
                customerName1 = reinvoicePage.GetFirstCustomerToName();
                customerCode1 = reinvoicePage.GetFirstCustomerToCode();
            }

            var customerTo = customerCode1 + " - " + customerName1.Substring(0, customerName1.Length - 3);

            reinvoicePage.Filter(ReinvoicePage.FilterType.SearchCustomersTo, customerTo);

            if (!reinvoicePage.isPageSizeEqualsTo100())
            {
                reinvoicePage.PageSize("8");
                reinvoicePage.PageSize("100");
            }

            Assert.IsTrue(reinvoicePage.VerifyCustomerTo(customerName1), string.Format(MessageErreur.FILTRE_ERRONE, "CustomersTo"));
        }

        //Créer une nouvelle refacturation
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_RINV_Create_New_Reinvoice()
        {
            string customerType = TestContext.Properties["CustomerType"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            Random rnd = new Random();

            // Prepare customer
            int rndCustomer = rnd.Next();
            int rndCustomer1 = rnd.Next();

            string customerName = customerNameToday + rndCustomer.ToString();
            string customerCode = rndCustomer.ToString();
            string customerName1 = customerNameToday + rndCustomer1.ToString();
            string customerCode1 = rndCustomer1.ToString();

            // Prepare service
            string serviceName = serviceNameToday + rndCustomer.ToString();
            string serviceName1 = serviceNameToday + rndCustomer1.ToString();

            DateTime fromDate = DateUtils.Now;
            DateTime toDate = DateUtils.Now.AddDays(10);

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act customer
            var customersPage = homePage.GoToCustomers_CustomerPage();
            var customerCreateModalPage = customersPage.CustomerCreatePage();
            customerCreateModalPage.FillFields_CreateCustomerModalPage(customerName, customerCode, customerType);
            var customerGeneralInformationsPage = customerCreateModalPage.Create();
            customersPage = customerGeneralInformationsPage.BackToList();

            customerCreateModalPage = customersPage.CustomerCreatePage();
            customerCreateModalPage.FillFields_CreateCustomerModalPage(customerName1, customerCode1, customerType);
            customerCreateModalPage.Create();
            customerGeneralInformationsPage.BackToList();

            // Act service
            var servicePage = homePage.GoToCustomers_ServicePage();
            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            pricePage = priceModalPage.FillFields_CustomerPrice(site, customerName, fromDate, toDate);
            servicePage = pricePage.BackToList();

            serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName1);
            serviceGeneralInformationsPage = serviceCreateModalPage.Create();
            pricePage = serviceGeneralInformationsPage.GoToPricePage();
            priceModalPage = pricePage.AddNewCustomerPrice();
            pricePage = priceModalPage.FillFields_CustomerPrice(site, customerName1, fromDate, toDate);
            pricePage.BackToList();


            // Act reinvoice
            var reinvoicePage = homePage.GoToCustomers_ReinvoicePage();

            try
            {
                var reinvoiceCreateModalpage = reinvoicePage.CreateNewReinvoice();
                reinvoiceCreateModalpage.FillFields_CreateReinvoiceModalPage(customerName1, customerName);
                reinvoicePage = reinvoiceCreateModalpage.Create();

                var customerFrom = customerCode1 + " - " + customerName1;
                var customerTo = customerCode + " - " + customerName;

                reinvoicePage.Filter(ReinvoicePage.FilterType.SearchCustomersFrom, customerFrom);
                reinvoicePage.Filter(ReinvoicePage.FilterType.SearchCustomersTo, customerTo);

                string CustomerFromName = reinvoicePage.GetFirstCustomerFromName(); 
                // Assert
                Assert.AreEqual(CustomerFromName, customerName1, "Le reinvoice n'a pas été créé et ajouté à la liste.");
            }
            finally
            {
                reinvoicePage.DeleteReinvoice();
            }
        }

        //Créer un export new version
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_RINV_Export_New_Version_Reinvoice()
        {
            // Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            bool newVersion = true;

            string customerType = TestContext.Properties["CustomerType"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            Random rnd = new Random();

            // Prepare customer
            int rndCustomer = rnd.Next();
            int rndCustomer1 = rnd.Next();

            string customerName = customerNameToday + rndCustomer.ToString();
            string customerCode = rndCustomer.ToString();
            string customerName1 = customerNameToday + rndCustomer1.ToString();
            string customerCode1 = rndCustomer1.ToString();

            // Prepare service
            string serviceName = serviceNameToday + rndCustomer.ToString();
            string serviceName1 = serviceNameToday + rndCustomer1.ToString();

            DateTime fromDate = DateUtils.Now;
            DateTime toDate = DateUtils.Now.AddDays(10);

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act reinvoice
            var reinvoicePage = homePage.GoToCustomers_ReinvoicePage();

            reinvoicePage.ClearDownloads();
            reinvoicePage.ResetFilter();

            if (reinvoicePage.CheckTotalNumber() == 0)
            {
                // Act customer
                var customersPage = homePage.GoToCustomers_CustomerPage();
                customersPage.ResetFilters();
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customerName, customerCode, customerType);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();

                customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customerName1, customerCode1, customerType);
                customerCreateModalPage.Create();

                // Act service
                var servicePage = homePage.GoToCustomers_ServicePage();
                servicePage.ResetFilters();
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                priceModalPage.FillFields_CustomerPrice(site, customerName, fromDate, toDate);
                servicePage = serviceGeneralInformationsPage.BackToList();

                serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName1);
                serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                pricePage = serviceGeneralInformationsPage.GoToPricePage();
                priceModalPage = pricePage.AddNewCustomerPrice();
                priceModalPage.FillFields_CustomerPrice(site, customerName1, fromDate, toDate);

                // Act reinvoice
                reinvoicePage = homePage.GoToCustomers_ReinvoicePage();
                var reinvoiceCreateModalpage = reinvoicePage.CreateNewReinvoice();
                reinvoiceCreateModalpage.FillFields_CreateReinvoiceModalPage(customerName, customerName1);
                reinvoicePage = reinvoiceCreateModalpage.Create();
            }

            var customerFromName = reinvoicePage.GetFirstCustomerFromName();
            var customerFromCode = reinvoicePage.GetFirstCustomerFromCode();

            var customerFrom = customerFromCode + " - " + customerFromName;

            var customerTo = reinvoicePage.GetFirstCustomerToCode() + " - " + reinvoicePage.GetFirstCustomerToName();

            reinvoicePage.Filter(ReinvoicePage.FilterType.SearchCustomersFrom, customerFrom);
            // plus rapide l'export ?!?
            reinvoicePage.Filter(ReinvoicePage.FilterType.SearchCustomersTo, customerTo);

            reinvoicePage.Export(newVersion);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = reinvoicePage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber(REINVOICES_EXCEL_SHEET_NAME, filePath);

            //Assert
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_RINV_Create_New_Reinvoice_With_Empty_Input()
        {
            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act reinvoice
            var reinvoicePage = homePage.GoToCustomers_ReinvoicePage();
            var reinvoiceCreateModalpage = reinvoicePage.CreateNewReinvoice();

            // Assert
            bool IsOkToCreate = reinvoiceCreateModalpage.IsOkToCreate(); 
            Assert.IsFalse(IsOkToCreate, "Le reinvoice a pu être créé sans renseigner les données obligtoires.");
        }

        //Créer une nouvelle refacturation all ready exist
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_RINV_Create_New_Reinvoice_AllReady()
        {
            string customerType = TestContext.Properties["CustomerType"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            Random rnd = new Random();

            // Prepare customer
            int rndCustomer = rnd.Next();
            int rndCustomer1 = rnd.Next();

            string customerName = customerNameToday + rndCustomer.ToString();
            string customerCode = rndCustomer.ToString();
            string customerName1 = customerNameToday + rndCustomer1.ToString();
            string customerCode1 = rndCustomer1.ToString();

            // Prepare service
            string serviceName = serviceNameToday + rndCustomer.ToString();
            string serviceName1 = serviceNameToday + rndCustomer1.ToString();

            DateTime fromDate = DateUtils.Now;
            DateTime toDate = DateUtils.Now.AddDays(10);

            // Arrange
            HomePage homePage = LogInAsAdmin();

            var customerFromName = customerName1;
            var customerToName = customerName;

            // Create new customer
            var customersPage = homePage.GoToCustomers_CustomerPage();
            var customerCreateModalPage = customersPage.CustomerCreatePage();
            customerCreateModalPage.FillFields_CreateCustomerModalPage(customerName, customerCode, customerType);
            var customerGeneralInformationsPage = customerCreateModalPage.Create();
            customersPage = customerGeneralInformationsPage.BackToList();

            customerCreateModalPage = customersPage.CustomerCreatePage();
            customerCreateModalPage.FillFields_CreateCustomerModalPage(customerName1, customerCode1, customerType);
            customerCreateModalPage.Create();

            // Create new service
            var servicePage = homePage.GoToCustomers_ServicePage();
            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();
            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            priceModalPage.FillFields_CustomerPrice(site, customerName, fromDate, toDate);
            servicePage = serviceGeneralInformationsPage.BackToList();

            serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName1);
            serviceGeneralInformationsPage = serviceCreateModalPage.Create();
            pricePage = serviceGeneralInformationsPage.GoToPricePage();
            priceModalPage = pricePage.AddNewCustomerPrice();
            priceModalPage.FillFields_CustomerPrice(site, customerName1, fromDate, toDate);

            var reinvoicePage = homePage.GoToCustomers_ReinvoicePage();

            var reinvoiceCreateModalpage = reinvoicePage.CreateNewReinvoice();
            reinvoiceCreateModalpage.FillFields_CreateReinvoiceModalPage(customerFromName, customerToName);
            reinvoicePage = reinvoiceCreateModalpage.Create();

            reinvoiceCreateModalpage = reinvoicePage.CreateNewReinvoice();
            reinvoiceCreateModalpage.FillFields_CreateReinvoiceModalPage(customerFromName, customerToName);
            // Assert
            bool IsServiceFromSelected = reinvoiceCreateModalpage.IsServiceFromSelected();
            bool IsServiceToSelected = reinvoiceCreateModalpage.IsServiceToSelected(); 

            Assert.IsFalse(IsServiceFromSelected, "Aucun service n'est sélectionné pour le customer From");
            Assert.IsFalse(IsServiceToSelected, "Aucun service n'est sélectionné pour le customer To");

            reinvoiceCreateModalpage.CloseReinvoiceCreateModal();
        }

        //Supprimer une refacturation
        [TestMethod]
        [Timeout(_timeout)]
        public void CU_RINV_Delete_Reinvoice()
        {
            string customerType = TestContext.Properties["CustomerType"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            Random rnd = new Random();

            // Prepare customer
            int rndCustomer = rnd.Next();
            int rndCustomer1 = rnd.Next();

            string customerName = customerNameToday + rndCustomer.ToString();
            string customerCode = rndCustomer.ToString();
            string customerName1 = customerNameToday + rndCustomer1.ToString();
            string customerCode1 = rndCustomer1.ToString();

            // Prepare service
            string serviceName = serviceNameToday + rndCustomer.ToString();
            string serviceName1 = serviceNameToday + rndCustomer1.ToString();

            DateTime fromDate = DateUtils.Now;
            DateTime toDate = DateUtils.Now.AddDays(10);

            // Arrange
            HomePage homePage = LogInAsAdmin();
           

            var reinvoicePage = homePage.GoToCustomers_ReinvoicePage();
            reinvoicePage.ResetFilter();

            if (reinvoicePage.CheckTotalNumber() == 0 || reinvoicePage.CheckTotalNumber() == 1)
            {
                // Act customer
                var customersPage = homePage.GoToCustomers_CustomerPage();
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customerName, customerCode, customerType);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();

                customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customerName1, customerCode1, customerType);
                customerCreateModalPage.Create();

                // Act service
                var servicePage = homePage.GoToCustomers_ServicePage();
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                priceModalPage.FillFields_CustomerPrice(site, customerName, fromDate, toDate);
                servicePage = serviceGeneralInformationsPage.BackToList();

                serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName1);
                serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                pricePage = serviceGeneralInformationsPage.GoToPricePage();
                priceModalPage = pricePage.AddNewCustomerPrice();
                priceModalPage.FillFields_CustomerPrice(site, customerName1, fromDate, toDate);

                // Act reinvoice
                reinvoicePage = homePage.GoToCustomers_ReinvoicePage();
                var reinvoiceCreateModalpage = reinvoicePage.CreateNewReinvoice();
                reinvoiceCreateModalpage.FillFields_CreateReinvoiceModalPage(customerName, customerName1);
                reinvoiceCreateModalpage.Create();
            }
            else
            {
                customerName1 = reinvoicePage.GetFirstCustomerToName();
                customerName = reinvoicePage.GetFirstCustomerFromName();

                customerCode1 = reinvoicePage.GetFirstCustomerToCode();
                customerCode = reinvoicePage.GetFirstCustomerFromCode();
            }

            var customerFrom = customerCode + " - " + customerName;
            var customerTo = customerCode1 + " - " + customerName1;

            int valueAfterCreate = reinvoicePage.CheckTotalNumber();

            reinvoicePage.Filter(ReinvoicePage.FilterType.SearchCustomersFrom, customerFrom);
            reinvoicePage.Filter(ReinvoicePage.FilterType.SearchCustomersTo, customerTo);

            // Assert
            Assert.AreEqual(reinvoicePage.GetFirstCustomerFromName(), customerName, "Le reinvoice n'est pas présent dans la liste.");
            reinvoicePage.DeleteReinvoice();

            reinvoicePage.ResetFilter();
            int valueAfterDelete = reinvoicePage.CheckTotalNumber();

            // Assert
            Assert.AreEqual((valueAfterCreate - 1), valueAfterDelete, "Le reinvoice n'a pas été supprimé de la liste.");
        }

        // Enlever les filtres de recherches
        [TestMethod]
        [Timeout(_timeout)]
        public void CU_RINV_Reset_Filter()
        {
            string site = TestContext.Properties["Site"].ToString();
            string customerType = TestContext.Properties["CustomerType"].ToString();

            Random rnd = new Random();

            // Prepare customer
            int rndCustomer = rnd.Next();
            int rndCustomer1 = rnd.Next();

            string customerName = customerNameToday + rndCustomer.ToString();
            string customerCode = rndCustomer.ToString();
            string customerName1 = customerNameToday + rndCustomer1.ToString();
            string customerCode1 = rndCustomer1.ToString();

            // Prepare service
            string serviceName = serviceNameToday + rndCustomer.ToString();
            string serviceName1 = serviceNameToday + rndCustomer1.ToString();

            DateTime fromDate = DateUtils.Now;
            DateTime toDate = DateUtils.Now.AddDays(10);

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var reinvoicePage = homePage.GoToCustomers_ReinvoicePage();

            if (reinvoicePage.CheckTotalNumber() == 0)
            {
                // Act customer
                var customersPage = homePage.GoToCustomers_CustomerPage();
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customerName, customerCode, customerType);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();

                customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customerName1, customerCode1, customerType);
                customerCreateModalPage.Create();

                // Act service
                var servicePage = homePage.GoToCustomers_ServicePage();
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                priceModalPage.FillFields_CustomerPrice(site, customerName, fromDate, toDate);
                servicePage = serviceGeneralInformationsPage.BackToList();

                serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName1);
                serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                pricePage = serviceGeneralInformationsPage.GoToPricePage();
                priceModalPage = pricePage.AddNewCustomerPrice();
                priceModalPage.FillFields_CustomerPrice(site, customerName1, fromDate, toDate);

                // Act reinvoice
                reinvoicePage = homePage.GoToCustomers_ReinvoicePage();
                var reinvoiceCreateModalpage = reinvoicePage.CreateNewReinvoice();
                reinvoiceCreateModalpage.FillFields_CreateReinvoiceModalPage(customerName, customerName1);
                reinvoiceCreateModalpage.Create();
            }
            else
            {
                customerName1 = reinvoicePage.GetFirstCustomerToName();
                customerName = reinvoicePage.GetFirstCustomerFromName();

                customerCode1 = reinvoicePage.GetFirstCustomerToCode();
                customerCode = reinvoicePage.GetFirstCustomerFromCode();
            }

            var customerFrom = customerCode + " - " + customerName;
            var customerTo = customerCode1 + " - " + customerName1;

            reinvoicePage.Filter(ReinvoicePage.FilterType.SearchCustomersFrom, customerFrom);
            reinvoicePage.Filter(ReinvoicePage.FilterType.SearchCustomersTo, customerTo);

            var numberCustFromSelected = reinvoicePage.GetNumberCustFromSelected();
            var numberCustToSelected = reinvoicePage.GetNumberCustToSelected();

            reinvoicePage.ResetFilter();

            var numberCustFromSelectedAfterReset = reinvoicePage.GetNumberCustFromSelected();
            var numberCustToSelectedAfterReset = reinvoicePage.GetNumberCustToSelected();

            // Assert
            Assert.AreNotEqual(numberCustFromSelected, numberCustFromSelectedAfterReset, "Le filtre CustomersFrom n'a pas été réinitialisé.");
            Assert.AreNotEqual(numberCustToSelected, numberCustToSelectedAfterReset, "Le filtre CustomersTo n'a pas été réinitialisé.");
        }

        //Créer une nouvelle refacturation vers un autre site
        [TestMethod]
        [Timeout(_timeout)]
        public void CU_RINV_Create_Reinvoice_Flight_Invoice()
        {
            string customerType = TestContext.Properties["CustomerType"].ToString();
            string siteMAD = TestContext.Properties["Site"].ToString();
            string siteACE = TestContext.Properties["SiteLP"].ToString();

            // Prepare customer
            string customerFromName = "CustomerTest-RFI-From";
            string customerFromCode = "RFIF";
            string customerToName = "CustomerTest-RFI-To";
            string customerToCode = "RFIT";

            // Prepare service
            string serviceFromName = "ServiceTest-RFI-From";
            string serviceToName = "ServiceTest-RFI-To";

            DateTime fromDate = DateUtils.Now.AddDays(-10);
            DateTime toDate = DateUtils.Now.AddDays(1);
            DateTime yesterdayDate = DateUtils.Now.AddDays(-1);

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act customer
            var customersPage = homePage.GoToCustomers_CustomerPage();

            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, customerFromName);
            if (customersPage.CheckTotalNumber() == 0)
            {
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customerFromName, customerFromCode, customerType);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();
            }

            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, customerToName);
            if (customersPage.CheckTotalNumber() == 0)
            {
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customerToName, customerToCode, customerType);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customerGeneralInformationsPage.BackToList();
            }

            // Act service
            var servicePage = homePage.GoToCustomers_ServicePage();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceFromName);
            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceFromName);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteACE, customerFromName, fromDate, toDate);
                servicePage = pricePage.BackToList();
            }
            else
            {
                var pricePage = servicePage.ClickOnFirstService();
                pricePage.ToggleFirstPrice();
                var serviceCreatePriceModalPage = pricePage.EditFirstPrice(siteACE, customerFromName);
                serviceCreatePriceModalPage.EditPriceDates(fromDate, toDate);

                var serviceGeneralInformationsPage = pricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                servicePage = serviceGeneralInformationsPage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceToName);
            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceToName);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteACE, customerToName, fromDate, toDate);
                priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteMAD, customerToName, fromDate, toDate);
                pricePage.BackToList();
            }
            else
            {
                var pricePage = servicePage.ClickOnFirstService();
                pricePage.ToggleFirstPrice();
                var serviceCreatePriceModalPage = pricePage.EditFirstPrice(siteACE, customerToName);
                serviceCreatePriceModalPage.EditPriceDates(fromDate, toDate);
                pricePage.TogglePrice(siteMAD, customerToName);
                serviceCreatePriceModalPage = pricePage.EditPrice(siteMAD, customerToName);
                serviceCreatePriceModalPage.EditPriceDates(fromDate, toDate);

                var serviceGeneralInformationsPage = pricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                serviceGeneralInformationsPage.BackToList();
            }

            // Act reinvoice
            //Attention, on ne peut créer une reinvoice de customer si la date TO du price est antérieur à aujourd'hui

            var reinvoicePage = homePage.GoToCustomers_ReinvoicePage();
            var customerFrom = customerFromCode + " - " + customerFromName;
            var customerTo = customerToCode + " - " + customerToName;
            reinvoicePage.Filter(ReinvoicePage.FilterType.SearchCustomersFrom, customerFrom);
            reinvoicePage.Filter(ReinvoicePage.FilterType.SearchCustomersTo, customerTo);
            if (reinvoicePage.GetFirstCustomerFromName() != customerFromName || reinvoicePage.CheckTotalNumber() == 0)
            {
                var reinvoiceCreateModalpage = reinvoicePage.CreateNewReinvoiceWithSites();
                reinvoiceCreateModalpage.FillFields_CreateReinvoiceModalPage(customerFromName, customerToName, siteMAD);
                reinvoiceCreateModalpage.Create();
            }

            servicePage = homePage.GoToCustomers_ServicePage();

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceFromName);
            if (servicePage.CheckTotalNumber() != 0)
            {
                var pricePage = servicePage.ClickOnFirstService();
                pricePage.ToggleFirstPrice();
                var serviceCreatePriceModalPage = pricePage.EditFirstPrice(siteACE, customerFromName);
                serviceCreatePriceModalPage.EditPriceDates(fromDate, yesterdayDate);

                var serviceGeneralInformationsPage = pricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                servicePage = serviceGeneralInformationsPage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceToName);
            if (servicePage.CheckTotalNumber() != 0)
            {
                var pricePage = servicePage.ClickOnFirstService();
                pricePage.ToggleFirstPrice();
                var serviceCreatePriceModalPage = pricePage.EditFirstPrice(siteACE, customerToName);
                serviceCreatePriceModalPage.EditPriceDates(fromDate, yesterdayDate);
                pricePage.TogglePrice(siteMAD, customerToName);
                serviceCreatePriceModalPage = pricePage.EditPrice(siteMAD, customerToName);
                serviceCreatePriceModalPage.EditPriceDates(fromDate, yesterdayDate);

                var serviceGeneralInformationsPage = pricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                serviceGeneralInformationsPage.BackToList();
            }

            // Flight  part
            string flightNb = CreateFlight(homePage, siteMAD, customerToName, serviceToName, yesterdayDate);

            string customerPickMethod = TestContext.Properties["AllFlightsInSelectedPeriod"].ToString();
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            invoicesPage.ResetFilters();
            var invoiceCreateAutoModalpage = invoicesPage.AutoInvoiceCreatePage();
            invoiceCreateAutoModalpage.FillField_CreateNewAutoInvoice(customerToName, siteMAD, customerPickMethod);
            var invoiceDetails = invoiceCreateAutoModalpage.Submit();
            invoiceCreateAutoModalpage.CloseWarningInvoicePopup();
            invoiceDetails = invoicesPage.SelectFirstInvoice();
            var numInvoice = invoiceDetails.GetInvoiceNumber();
            string invoiceServiceName = invoiceDetails.GetInvoiceFirstServiceName();
            string invoiceFlightDetails = invoiceDetails.GetInvoiceFlightDetails();

            invoicesPage = invoiceDetails.BackToList();
            invoicesPage.ResetFilters();

            //Asserts
            Assert.IsTrue(invoicesPage.GetFirstInvoiceID().Contains(numInvoice), "L'invoice créée n'apparaît pas dans la liste des invoices.");
            Assert.IsTrue(invoiceServiceName.Contains(serviceToName), "Le service associé à l'invoice n'est pas le bon.");
            Assert.IsTrue(invoiceFlightDetails.Contains(flightNb), "Le flight associé à l'invoice n'est pas le bon.");
        }

        
        [TestMethod]
        [Timeout(_timeout)]
        public void CU_RINV_Create_Reinvoice_Delivery_Invoice()
        {
            // Prepare
            string customerType = "Resto Co";
            string customerA = "SmithA";
            string customerCodeA = "A";
            string customerB = "SmithB";
            string customerCodeB = "B";
            string serviceA = "SmithSA";
            string serviceB = "SmithSB";
            string siteA = "MAD";
            string siteB = "ACE";
            string deliveryA = "SmithD";
            string invoiceB = "SMITHB";

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act customer
            // type Customer "Resto co"
            ParametersCustomer paramCustomer = homePage.GoToParameters_CustomerPage();
            if (!paramCustomer.isTypeOfCustomerExist(customerType))
            {
                paramCustomer.AddNewTypeOfCustomer(customerType);
            }

            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            //1. Créer 2 Customer A et B avec comme customer type "Resto Co"
            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, customerA);
            if (customersPage.CheckTotalNumber() == 0)
            {
                CustomerCreateModalPage createCustomer = customersPage.CustomerCreatePage();
                createCustomer.FillFields_CreateCustomerModalPage(customerA, customerCodeA, "Resto Co");
                CustomerGeneralInformationPage generalInfo = createCustomer.Create();
                customersPage = generalInfo.BackToList();
            }

            customersPage.Filter(CustomerPage.FilterType.Search, customerB);
            if (customersPage.CheckTotalNumber() == 0)
            {
                CustomerCreateModalPage createCustomer = customersPage.CustomerCreatePage();
                createCustomer.FillFields_CreateCustomerModalPage(customerB, customerCodeB, "Resto Co");
                CustomerGeneralInformationPage generalInfo = createCustomer.Create();
                customersPage = generalInfo.BackToList();
            }

            // Clean-up
            // supprimer l'invoiceB via customerB
            var invoicePage = homePage.GoToAccounting_InvoicesPage();
            invoicePage.ResetFilters();
            invoicePage.Filter(InvoicesPage.FilterType.Customer, customerB);
            if (invoicePage.CheckTotalNumber() > 0)
            {
                invoicePage.DeleteFirstInvoice();
            }

            // service SmithS
            //Lui ajouter 1 service avec un prix actif associé au Customer et au Site
            //2. Créer 1 Service A avec price sur site A (MAD) et Customer A, avec un invoice price
            ServicePage service = homePage.GoToCustomers_ServicePage();
            service.ResetFilters();
            service.Filter(ServicePage.FilterType.Search, serviceA);
            service.Filter(ServicePage.FilterType.ShowAll, true);
            if (service.CheckTotalNumber() == 0)
            {
                ServiceCreateModalPage serv = service.ServiceCreatePage();
                serv.FillFields_CreateServiceModalPage(serviceA);
                ServiceGeneralInformationPage serviceGeneralInfo = serv.Create();
                ServicePricePage price = serviceGeneralInfo.GoToPricePage();
                ServiceCreatePriceModalPage createPrice = price.AddNewCustomerPrice();
                price = createPrice.FillFields_CustomerPrice(siteA, customerA, DateUtils.Now.AddMonths(-1), DateUtils.Now.AddMonths(1),"100");
            }
            else
            {
                //maj dateTo
                ServicePricePage price = service.ClickOnFirstService();
                price.UnfoldAll();
                ServiceCreatePriceModalPage priceModel = price.EditFirstPrice(siteA, customerA);
                price = priceModel.EditPriceDates(DateUtils.Now.AddMonths(-1), DateUtils.Now.AddMonths(1));
            }
            //3. Créer une service B avec un price sur site B et CUstomer B, avec une invoice price (150 €) et date à jours
            service = homePage.GoToCustomers_ServicePage();
            service.ResetFilters();
            service.Filter(ServicePage.FilterType.Search, serviceB);
            service.Filter(ServicePage.FilterType.ShowAll, true);
            if (service.CheckTotalNumber() == 0)
            {
                ServiceCreateModalPage serv = service.ServiceCreatePage();
                serv.FillFields_CreateServiceModalPage(serviceB);
                ServiceGeneralInformationPage serviceGeneralInfo = serv.Create();
                ServicePricePage price = serviceGeneralInfo.GoToPricePage();
                ServiceCreatePriceModalPage createPrice = price.AddNewCustomerPrice();
                price = createPrice.FillFields_CustomerPrice(siteB, customerB, DateUtils.Now.AddMonths(-1), DateUtils.Now.AddMonths(1),"150");
            }
            else
            {
                //maj dateTo
                ServicePricePage price = service.ClickOnFirstService();
                price.UnfoldAll();
                ServiceCreatePriceModalPage priceModel = price.EditFirstPrice(siteB, customerB);
                price = priceModel.EditPriceDates(DateUtils.Now.AddMonths(-1), DateUtils.Now.AddMonths(1));

            }
            
            //4. Créer un Delivery / Choisir le Customer A / Lui ajouter 1 service avec un prix actif associé au Customer et au Site A, ajouter une quantité
            DeliveryPage delivery = homePage.GoToCustomers_DeliveryPage();
            delivery.Filter(DeliveryPage.FilterType.Search, deliveryA);
            if(delivery.CheckTotalNumber() == 0)
            {
                DeliveryCreateModalPage createDelivery = delivery.DeliveryCreatePage();
                createDelivery.FillFields_CreateDeliveryModalPage(deliveryA, customerA, siteA, true);
                DeliveryLoadingPage deliveryLoad = createDelivery.Create();
                deliveryLoad.AddService(serviceA);
                deliveryLoad.AddQuantity("10");
                delivery = deliveryLoad.BackToList();
            }
            delivery.Filter(DeliveryPage.FilterType.Customers, customerA);
            DeliveryLoadingPage deliverLoadDing = delivery.ClickOnFirstDelivery();
            deliverLoadDing.AddQuantity("10");
            deliverLoadDing.ClickOnGeneralInformation();

            //Aller dans les Dispatch / Chercher le Customer A / Valider les quantités sur les différents onglets
            //5. Aller dans les Dispatch / Chercher le Customer A / Valider les quantités Previsional quantity, Quantity to produce et Quantity to invoice
            DispatchPage dispatch = homePage.GoToProduction_DispatchPage();
            dispatch.ResetFilter();
            dispatch.Filter(DispatchPage.FilterType.Site, siteA);
            dispatch.Filter(DispatchPage.FilterType.Customers, customerA);

            try
            {
                QuantityToInvoicePage onglet3d = dispatch.ClickQuantityToInvoice();
                dispatch.UnValidateAll();
                QuantityToProducePage onglet2d = dispatch.ClickQuantityToProduce();
                dispatch.UnValidateAll();
                PrevisionalQtyPage onglet1d = dispatch.ClickPrevisonalQuantity();
                dispatch.UnValidateAll();
            }
            catch
            {
                // déjà UnValidate
                dispatch = homePage.GoToProduction_DispatchPage();
                dispatch.ResetFilter();
                dispatch.Filter(DispatchPage.FilterType.Site, siteA);
                dispatch.Filter(DispatchPage.FilterType.Customers, customerA);
            }

            dispatch.AddQuantityOnPrevisonalQuantity(serviceA, "10");
            dispatch.ValidateAll();

            QuantityToProducePage onglet2 = dispatch.ClickQuantityToProduce();
            onglet2.UpdateQuantities("10");
            onglet2.ValidateFirstDispatch();

            QuantityToInvoicePage onglet3 = dispatch.ClickQuantityToInvoice();
            onglet3.UpdateQuantities("10");
            onglet3.ValidateTheFirst();
            
            //1. Create New Reinvoice with Sites : Site A - Customer A - Service A vers Site B - Customer B -Service B
            ReinvoicePage reinvoice = homePage.GoToCustomers_ReinvoicePage();
            // filtre dynamique au chargement de la page
            reinvoice.WaitPageLoading();
            reinvoice.Filter(ReinvoicePage.FilterType.SearchCustomersFrom, customerCodeA + " - " + customerA);
            reinvoice.Filter(ReinvoicePage.FilterType.SearchCustomersTo, customerCodeB + " - " + customerB);
            if (reinvoice.CheckTotalNumber() > 0)
            {
                reinvoice.DeleteReinvoice();
            }
            ReinvoiceCreateModalPage reinvoiceModal = reinvoice.CreateNewReinvoiceWithSites();
            reinvoiceModal.FillFields_CreateReinvoiceModalPage(customerA, customerB, siteB);
            reinvoice = reinvoiceModal.Create();

            //2. Aller dans Accounting invoice auto / create invoice auto / Site A/ Select All Deliveries / Customer A / Select All ==> Create
            InvoicesPage invoices = homePage.GoToAccounting_InvoicesPage();
            invoices.ResetFilters();
            invoices.Filter(InvoicesPage.FilterType.Customer, customerA);
            int compteur = 0;
            while (invoices.CheckTotalNumber()>0 && compteur<5)
            {
                invoices.DeleteFirstInvoice();
				compteur++;

			}
			invoices.Filter(InvoicesPage.FilterType.Customer, customerB);
			compteur = 0;
			while (invoices.CheckTotalNumber() > 0 && compteur < 5)
            {
                invoices.DeleteFirstInvoice();
				compteur++;
			}
			AutoInvoiceCreateModalPage auto = invoices.AutoInvoiceCreatePage();
            auto.FillField_CreateNewAutoInvoice(customerA, siteA, CustomerPickMethod.AllDeliveriesInSelectedPeriod);
            auto.FillFieldSelectAll();
            
            //1. Vérifier que la règle de Reinvoice soit bien créer
            invoices =  homePage.GoToAccounting_InvoicesPage();
            invoices.Filter(InvoicesPage.FilterType.Customer, customerB);
            Assert.AreEqual(1,invoices.CheckTotalNumber());

            //2. Vérifier que les infos de l'invoice soit correcte Customer B / Service B / quantity / unit price B
            InvoiceDetailsPage invoice = invoices.SelectFirstInvoice();
            var titre = invoice.WaitForElementIsVisible(By.XPath("//*[@id=\"div-body\"]/div/div[1]/h1"));
            Assert.IsTrue(titre.Text.EndsWith(invoiceB));
            Assert.AreEqual(serviceB, invoice.GetInvoiceFirstServiceName());
            Assert.AreEqual(150.0, invoice.GetUnitPrice(), 0.001);
            Assert.AreEqual(10.0, invoice.GetQuantity(),0.001);
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_RINV_Create_New_ReinvoiceWithSites()
        {
            string customerType = TestContext.Properties["CustomerType"].ToString();
            string site = TestContext.Properties["Site"].ToString();

            Random rnd = new Random();

            // Prepare customer
            int rndCustomer = rnd.Next();
            int rndCustomer1 = rnd.Next();

            string customerName = customerNameToday + rndCustomer.ToString();
            string customerCode = rndCustomer.ToString();
            string customerName1 = customerNameToday + rndCustomer1.ToString();
            string customerCode1 = rndCustomer1.ToString();

            // Prepare service
            string serviceName = serviceNameToday + rndCustomer.ToString();
            string serviceName1 = serviceNameToday + rndCustomer1.ToString();

            DateTime fromDate = DateUtils.Now;
            DateTime toDate = DateUtils.Now.AddDays(10);

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act customer
            var customersPage = homePage.GoToCustomers_CustomerPage();
            var customerCreateModalPage = customersPage.CustomerCreatePage();
            customerCreateModalPage.FillFields_CreateCustomerModalPage(customerName, customerCode, customerType);
            var customerGeneralInformationsPage = customerCreateModalPage.Create();
            customersPage = customerGeneralInformationsPage.BackToList();

            customerCreateModalPage = customersPage.CustomerCreatePage();
            customerCreateModalPage.FillFields_CreateCustomerModalPage(customerName1, customerCode1, customerType);
            customerCreateModalPage.Create();
            customerGeneralInformationsPage.BackToList();

            // Act service
            var servicePage = homePage.GoToCustomers_ServicePage();
            var serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName);
            var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            pricePage = priceModalPage.FillFields_CustomerPrice(site, customerName, fromDate, toDate);
            servicePage = pricePage.BackToList();

            serviceCreateModalPage = servicePage.ServiceCreatePage();
            serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName1);
            serviceGeneralInformationsPage = serviceCreateModalPage.Create();
            pricePage = serviceGeneralInformationsPage.GoToPricePage();
            priceModalPage = pricePage.AddNewCustomerPrice();
            pricePage = priceModalPage.FillFields_CustomerPrice(site, customerName1, fromDate, toDate);
            pricePage.BackToList();


            // Act reinvoice
            var reinvoicePage = homePage.GoToCustomers_ReinvoicePage();

            try
            {

                var reinvoiceCreateModalpage = reinvoicePage.CreateNewReinvoiceWithSites();
                reinvoiceCreateModalpage.FillFields_CreateReinvoiceModalPage(customerName1, customerName, site);
                reinvoiceCreateModalpage.Create();

                var customerFrom = customerCode1 + " - " + customerName1;
                var customerTo = customerCode + " - " + customerName;

                reinvoicePage.Filter(ReinvoicePage.FilterType.SearchCustomersFrom, customerFrom);
                reinvoicePage.Filter(ReinvoicePage.FilterType.SearchCustomersTo, customerTo);

                // Assert
                string CustomerFromName = reinvoicePage.GetFirstCustomerFromName();
                Assert.AreEqual(CustomerFromName, customerName1, "Le reinvoice n'a pas été créé et ajouté à la liste.");
            }
            finally
            {
                reinvoicePage.DeleteReinvoice();
            }
        }

        //METHOD
        private string CreateFlight(HomePage homePage, string site, string customer, string service, DateTime date)
        {
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string siteTo = TestContext.Properties["SiteLP"].ToString();
            var flightNumber = new Random().Next();
            var state = "I";

            // On vérifie que la valeur de newVersionFlight est égale à 2
            //CheckNewVersionFlight(homePage);//old version

            // Create flight
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, site);

            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber.ToString(), customer, aircraft, site, siteTo, null, "00", "23", null, date);

            var editPage = flightPage.EditFirstFlight(flightNumber.ToString());
            editPage.AddGuestType();
            editPage.AddService(service);
            flightPage = editPage.CloseViewDetails();

            //contre le Invoiced "No guest in flight" + Regist.
            
            WebDriver.Navigate().Refresh();
            Thread.Sleep(2000);
            if (flightPage.IsPreval())
            {
                flightPage.SetNewState("V");
                // animation orange->vert
                Thread.Sleep(3000);
            }

            WebDriver.Navigate().Refresh();
            Thread.Sleep(2000);
            if (flightPage.IsValidated())
            {
                flightPage.SetNewState("I");
                // animation orange->vert
                Thread.Sleep(3000);
            }

            // On vérifie que le state Invoice est sélectionné
            Assert.IsTrue(flightPage.GetFirstFlightStatus("I"), "Le state 'Invoiced' n'est pas sélectionné pour le flight créé.");

            return flightPage.GetFirstFlightNumber();
        }
    }
}