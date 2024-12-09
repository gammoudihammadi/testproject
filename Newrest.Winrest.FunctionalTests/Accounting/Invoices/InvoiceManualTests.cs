using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Accounting.Invoice;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Customer;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Delivery;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.DeliveryRound;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Reinvoice;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.Flight;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Sites;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.Dispatch;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.OutputForm;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Security.Policy;
using System.Security.Principal;
using System.Web;
using System.Xml.Linq;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using static Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service.ServiceMassiveDeleteModalPage;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;


namespace Newrest.Winrest.FunctionalTests.Accounting
{
    [TestClass]
    public class InvoiceManualTests : TestBase
    {

        private const int _timeout = 600000;
        private readonly string INVOICE_EXCEL_SHEET = "Invoices";
        private static DateTime fromdate = DateUtils.Now;
        private static DateTime todate = DateUtils.Now.AddDays(+30);
        private readonly string serviceNameToday = "Service-" + DateUtils.Now.ToString("dd/MM/yyyy");
        private string manualInvoiceNumber;
        private string ID;

        [TestInitialize]
        public override void TestInitialize()
        {
            base.TestInitialize();

            var testMethod = TestContext.TestName;
            switch (testMethod)
            {
                case nameof(AC_INVO_FilterInvoiceStepAccounted):
                    TestInitialize_CreateCustomerForInvoiceManual();
                    TestInitialize_CreateManualInvoiceValidate();
                    break;

                case nameof(AC_INVO_FilterShow_ShownotsenttoSageCegidonly):
                    TestInitialize_CreateCustomerForInvoiceManual();
                    TestInitialize_CreateManualInvoiceValidate();
                    break;

                case nameof(AC_INVO_FilterShow_Showallinvoices):
                    TestInitialize_CreateCustomerForInvoiceManual();
                    TestInitialize_CreateManualInvoiceValidate();
                    break;

                case nameof(AC_INVO_FilterShow_Showmanualinvoices):
                    TestInitialize_CreateCustomerForInvoiceManual();
                    TestInitialize_CreateManualInvoiceValidate();
                    break;

                case nameof(AC_INVO_FilterShow_Shownotvalidareonly):
                    TestInitialize_CreateCustomerForInvoiceManual();
                    TestInitialize_CreateManualInvoiceNonValidate();
                    break;

                case nameof(AC_INVO_FilterInvoiceStep_ValidatedAccounted):
                    TestInitialize_CreateCustomerForInvoiceManual();
                    TestInitialize_CreateManualInvoiceNonValidate();
                    break;

                case nameof(AC_INVO_FilterInvoiceStepDraft):
                    TestInitialize_CreateCustomerForInvoiceManual();
                    TestInitialize_CreateManualInvoiceNonValidate();
                    break;
                case nameof(AC_INVO_FilterInvoiceStepAll):
                    TestInitialize_CreateCustomerForInvoiceManual();
                    TestInitialize_CreateManualInvoiceNonValidate();
                    break;

                case nameof(AC_INVO_FilterShow_Shownotsentbymailonly):
                    TestInitialize_CreateCustomerForInvoiceManual();
                    TestInitialize_CreateManualInvoiceValidate();

                    break;
                 case nameof(AC_INVO_Filter_SortById):
                    TestInitialize_CreateCustomerForInvoiceManual();
                    TestInitialize_CreateManualInvoiceValidate();

                    break;

                default:
                    break;
            }
        }
        public void TestInitialize_CreateCustomerForInvoiceManual()
        {
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string customerIcao = TestContext.Properties["InvoiceCustomerCode"].ToString();
            string customerType = TestContext.Properties["InvoiceCustomerType"].ToString();
            string serviceName = TestContext.Properties["InvoiceService"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            DateTime fromDate = DateUtils.Now.AddYears(-1);
            DateTime toDate = DateUtils.Now.AddYears(+3);
            int toCheck = 0;

            // Arrange
            HomePage homePage = LogInAsAdmin();

            ClearCache();

            // Create Customer
            CustomerPage customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, customer);

            if (customersPage.CheckTotalNumber() == toCheck)
            {
                CustomerCreateModalPage customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customer, customerIcao, customerType);
                CustomerGeneralInformationPage customerGeneralInformationsPage = customerCreateModalPage.Create();

                customersPage = customerGeneralInformationsPage.BackToList();
                customersPage.ResetFilters();
                customersPage.Filter(CustomerPage.FilterType.Search, customer);
            }
            string checkCustomer = customersPage.GetFirstCustomerName();
            Assert.AreEqual(customer, checkCustomer, "Le customer n'a pas été créé.");

            // Create Service
            ServicePage servicesPage = homePage.GoToCustomers_ServicePage();
            servicesPage.ResetFilters();
            servicesPage.Filter(ServicePage.FilterType.Search, serviceName);

            if (servicesPage.CheckTotalNumber() == toCheck)
            {
                ServiceCreateModalPage serviceModalPage = servicesPage.ServiceCreatePage();
                serviceModalPage.FillFields_CreateServiceModalPage(serviceName, "", "");
                ServiceGeneralInformationPage serviceGeneralInformationsPage = serviceModalPage.Create();

                ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
                ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, fromDate, toDate);
                servicesPage = pricePage.BackToList();
            }
            else
            {
                ServicePricePage pricePage = servicesPage.ClickOnFirstService();
                pricePage.SearchPriceForCustomer(site, customer, fromDate, toDate);
                servicesPage = pricePage.BackToList();
            }
            servicesPage.ResetFilters();
            servicesPage.Filter(ServicePage.FilterType.Search, serviceName);
            string displayedService = servicesPage.GetFirstServiceName();
            bool checkService = displayedService.Contains(serviceName);
            Assert.IsTrue(checkService, MessageErreur.FILTRE_ERRONE, "Search");
        }
        private void TestInitialize_CreateManualInvoiceValidate()
        {
            // prepare
            bool isVATForced = true;
            DateTime date = DateTime.Now;
            string quantity = "10";
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string serviceName = TestContext.Properties["InvoiceService"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            // arrange
            HomePage homePage = LogInAsAdmin();
            // act
            InvoicesPage invoicesPage = homePage.GoToAccounting_InvoicesPage();
            // créer un invoice manuelle
            ManualInvoiceCreateModalPage invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
            InvoiceDetailsPage invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, date, customer, isVATForced);
            // ajouter un service
            invoiceDetails.AddService(serviceName, quantity);
            // valider invoice manuelle 
            invoiceDetails.Validate();
            // récupérer le numèro de invoice manuelle 
            manualInvoiceNumber = invoiceDetails.GetInvoiceNumber();
            invoiceDetails.BackToList();
        }

        private void TestInitialize_CreateManualInvoiceNonValidate()
        {
            //prepare
            bool isVATForced = true;
            DateTime date = DateTime.Now;
            string quantity = "10";
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string serviceName = TestContext.Properties["InvoiceService"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            // arrange
            HomePage homePage = LogInAsAdmin();
            // act
            InvoicesPage invoicesPage = homePage.GoToAccounting_InvoicesPage();
            // créer un invoice manuelle
            ManualInvoiceCreateModalPage invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
            InvoiceDetailsPage invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, date, customer, isVATForced);
            // ajouter un service
            invoiceDetails.AddService(serviceName, quantity);
            ID = invoiceDetails.GetInvoiceNumber_details();
            // récupérer le numèro de invoice manuelle 
            manualInvoiceNumber = invoiceDetails.GetInvoiceNumber();
            invoiceDetails.BackToList();
        }

        // ________________________________________________________CREATE_INVOICE___________________________________________________

        [Priority(0)]
        [TestMethod]
        [Timeout(_timeout)]
        public void AC_INVO_CreateCustomerForInvoiceManual()
        {
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string customerIcao = TestContext.Properties["InvoiceCustomerCode"].ToString();
            string customerType = TestContext.Properties["InvoiceCustomerType"].ToString();

            string serviceName = TestContext.Properties["InvoiceService"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            ClearCache();

            // Act
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.Filter(CustomerPage.FilterType.Search, customer);

            if (customersPage.CheckTotalNumber() == 0)
            {
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customer, customerIcao, customerType);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();

                customersPage = customerGeneralInformationsPage.BackToList();
                customersPage.ResetFilters();
                customersPage.Filter(CustomerPage.FilterType.Search, customer);
            }

            //Assert
            Assert.AreEqual(customer, customersPage.GetFirstCustomerName(), "Le customer n'a pas été créé.");


            CheckService(homePage, serviceName, site, customer, DateUtils.Now.AddYears(-1), DateUtils.Now.AddYears(+3));
        }

        private void CheckService(HomePage homePage, string serviceName, string site, string customer, DateTime dateFrom, DateTime dateTo)
        {
            var servicesPage = homePage.GoToCustomers_ServicePage();
            servicesPage.Filter(ServicePage.FilterType.Search, serviceName);

            if (servicesPage.CheckTotalNumber() == 0)
            {
                var serviceModalPage = servicesPage.ServiceCreatePage();
                serviceModalPage.FillFields_CreateServiceModalPage(serviceName, "", "");
                var serviceGeneralInformationsPage = serviceModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.AddYears(-1), DateUtils.Now.AddYears(+3));
                servicesPage = pricePage.BackToList();
            }
            else
            {
                var pricePage = servicesPage.ClickOnFirstService();
                pricePage.SearchPriceForCustomer(site, customer, dateFrom, dateTo);
                servicesPage = pricePage.BackToList();
            }

            servicesPage.Filter(ServicePage.FilterType.Search, serviceName);
            Assert.IsTrue(servicesPage.GetFirstServiceName().Contains(serviceName), MessageErreur.FILTRE_ERRONE, "Search");
        }     

        [Priority(1)]
        [TestMethod]
        [Timeout(_timeout)]
        public void AC_INVO_CreateCustomerForInvoiceManualAirportTax()
        {
            string customer = TestContext.Properties["InvoiceCustomerAirportTax"].ToString();
            string customerIcao = TestContext.Properties["InvoiceCustomerCodeAirportTax"].ToString();
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

                customersPage = customerGeneralInformationsPage.BackToList();
                customersPage.ResetFilters();
                customersPage.Filter(CustomerPage.FilterType.Search, customer);
            }

            //Assert
            Assert.AreEqual(customer, customersPage.GetFirstCustomerName(), "Le customer n'a pas été créé.");
        }

        [Priority(2)]
        [TestMethod]
        [Timeout(_timeout)]
        public void AC_INVO_CreateServiceForInvoiceManualAirportTax()
        {
            string serviceName = TestContext.Properties["InvoiceServiceAirportTax"].ToString();
            string serviceCategory = TestContext.Properties["InvoiceServiceCategoryAirportTax"].ToString();
            string customer = TestContext.Properties["InvoiceCustomerAirportTax"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string serviceCode = new Random().Next().ToString();

            // Arrange
            var homePage = LogInAsAdmin();


            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);

            if (servicePage.CheckTotalNumber() == 0)
            {
                // Create
                var ServiceCreateModalPage = servicePage.ServiceCreatePage();
                ServiceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, null, serviceCategory);
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

        [Priority(3)]
        [TestMethod]
        [Timeout(_timeout)]
        public void AC_INVO_CreateCustomerWithForeignDevise()
        {
            string serviceName = TestContext.Properties["InvoiceServiceAirportTax"].ToString();
            string serviceCategory = TestContext.Properties["InvoiceServiceCategoryAirportTax"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string serviceCode = new Random().Next().ToString();

            string customerForeignDevise = TestContext.Properties["InvoiceCustomerWithForeignDevise"].ToString();
            string customerIcao = TestContext.Properties["InvoiceCustomerWithForeignDeviseIcao"].ToString();
            string customerType = TestContext.Properties["InvoiceCustomerType"].ToString();
            string foreignDevise = TestContext.Properties["InvoiceForeignDevise"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();


            // Act
            var customersPage = homePage.GoToCustomers_CustomerPage();

            customersPage.Filter(CustomerPage.FilterType.Search, customerForeignDevise);

            if (customersPage.CheckTotalNumber() == 0)
            {
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customerForeignDevise, customerIcao, customerType);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();

                customerGeneralInformationsPage.SetCurrency(foreignDevise);
                customersPage = customerGeneralInformationsPage.BackToList();

                customersPage.ResetFilters();
                customersPage.Filter(CustomerPage.FilterType.Search, customerForeignDevise);
                Assert.AreEqual(customerForeignDevise, customersPage.GetFirstCustomerName(), "Le customer n'a pas été créé.");
            }
            else
            {
                var customerGeneralInformation = customersPage.SelectFirstCustomer();
                string devise = customerGeneralInformation.GetDevise();

                if (devise != foreignDevise)
                    customerGeneralInformation.SetCurrency(foreignDevise);
            }

            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);

            if (servicePage.CheckTotalNumber() == 0)
            {
                // Create
                var ServiceCreateModalPage = servicePage.ServiceCreatePage();
                ServiceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, null, serviceCategory);
                var serviceGeneralInformationsPage = ServiceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customerForeignDevise, DateUtils.Now.AddMonths(-1), DateUtils.Now.AddMonths(+3));
                servicePage = pricePage.BackToList();
            }
            else
            {
                var servicePricePage = servicePage.ClickOnFirstService();
                servicePricePage.SearchPriceForCustomer(customerForeignDevise, site, DateUtils.Now.AddMonths(-1), DateUtils.Now.AddMonths(+3));

                servicePage = servicePricePage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);

            //Assert           
            Assert.IsTrue(servicePage.GetFirstServiceName().Contains(serviceName), String.Format(MessageErreur.FILTRE_ERRONE, "Search"));
        }

        [Priority(4)]
        [TestMethod]
        [Timeout(_timeout)]
        public void AC_INVO_PrepareExportSageConfig()
        {
            //Prepare
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string customerType = TestContext.Properties["InvoiceCustomerType"].ToString();
            string serviceName = TestContext.Properties["InvoiceService"].ToString();
            string journalInvoice = TestContext.Properties["Journal_Invoice"].ToString();

            string taxName = TestContext.Properties["TaxTypeInvoicesExportSage"].ToString();
            string taxType = "VAT";

            //Arrange
            var homePage = LogInAsAdmin();


            // Récupération de la catégorie du service
            string serviceCategory = GetServiceCategory(homePage, serviceName);

            // Vérification du paramétrage
            // --> Admin Settings
            bool isAppSettingsOK = SetApplicationSettingsForSageAuto(homePage);

            // Sites -- > Analytical plan et section
            bool isAnalyticalPlanOK = VerifySiteAnalyticalPlanSection(homePage, site);

            // Sites --> Contact sage invoice
            bool isMailSageOK = VerifyInvoiceSageContact(homePage, site);

            // Parameter - Purchasing --> VAT
            bool isPurchasingVATOK = VerifyPurchasingVAT(homePage, taxName, taxType);

            // Parameter - Accounting --> Service categories & VAT
            bool isCategoryAndVatOK = VerifyCategoryAndVAT(homePage, serviceCategory, taxName, customerType);

            // Parameter - Accounting --> Journal
            bool isJournalOk = VerifyAccountingJournal(homePage, site, journalInvoice);

            // Parameter - Accounting --> Integration Date
            DateTime date = VerifyIntegrationDate(homePage);

            // Customer
            bool isCustomerOK = VerifyCustomer(homePage, customer);

            //Site

            // Assert
            Assert.AreNotEqual("", serviceCategory, "La catégorie du service n'a pas été récupérée.");
            Assert.IsTrue(isAppSettingsOK, "Les application settings pour TL ne sont pas configurés correctement.");
            Assert.IsTrue(isAnalyticalPlanOK, "La configuration des analytical plan du site n'est pas effectuée.");
            Assert.IsTrue(isMailSageOK, $"Aucun mail n'est configuré pour le site {site} en cas d'erreur Sage.");
            Assert.IsTrue(isPurchasingVATOK, "La configuration des purchasing VAT n'est pas effectuée.");
            Assert.IsTrue(isCategoryAndVatOK, "La configuration des category and VAT n'est pas effectuée.");
            Assert.IsTrue(isJournalOk, "La catégorie du accounting journal n'a pas été effectuée.");
            Assert.IsNotNull(date, "La date d'intégration est nulle.");
            Assert.IsTrue(isCustomerOK, "La configuration du customer n'a pas été effectuée corrctement.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_CreateManualInvoice()
        {
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();


            //Act
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
            var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customer, true);
            // Récupération de l'identifiant de l'invoice créée
            invoicesPage = homePage.GoToAccounting_InvoicesPage();
            string ID = invoicesPage.GetFirstInvoiceID();
            invoicesPage.ResetFilters();
            //Assert
            Assert.IsTrue(invoicesPage.GetFirstInvoiceID().Contains(ID), "L'invoice créée n'apparaît pas dans la liste des invoices.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_CreateManualInvoiceEmpty()
        {
            //Arrange
            var homePage = LogInAsAdmin();


            //Act
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();

            //Assert
            bool canCreate = invoiceCreateModalpage.CanCreate();
            Assert.IsFalse(canCreate, "Impossible de créer cette invoice car les champs obligatoires ne sont pas renseignés.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_CreateManualInvoiceWithAirportTax()
        {
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomerAirportTax"].ToString();
            string customerCode = TestContext.Properties["InvoiceCustomerCodeAirportTax"].ToString();
            string taxName = TestContext.Properties["InvoiceTaxName"].ToString();
            string taxType = TestContext.Properties["InvoiceTaxType"].ToString();
            string serviceName = TestContext.Properties["InvoiceServiceAirportTax"].ToString();
            string currency = TestContext.Properties["Currency"].ToString();
            string taxValue = "15";

            string purchaseCode = "AS05";
            string purchaseAccount = "47205001";
            string salesCode = "AR10";
            string salesAccount = "47205001";
            string dueToInvoiceCode = "AP10";
            string dueToInvoiceAccount = "47205001";

            //Arrange
            var homePage = LogInAsAdmin();


            //_____________________________ Mise en place des paramètres____________________________
            // Client avec AirportTax = true
            var customerPage = homePage.GoToCustomers_CustomerPage();
            customerPage.Filter(CustomerPage.FilterType.Search, customer);
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
            accountingPage.SetAirportFeeSiteAndCustomer(site, customerCode, customer, taxValue);

            // Catégorie de service avec AirportTax = true
            var parametersCustomerPage = homePage.GoToParameters_CustomerPage();
            parametersCustomerPage.GoToTab_Category();
            parametersCustomerPage.SetServiceWithAirportTax(serviceName);

            //_____________________________ Fin Mise en place des paramètres________________________

            // Création d'un invoice
            //Invoice part
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
            InvoiceDetailsPage invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customer, true);
            invoicesPage = homePage.GoToAccounting_InvoicesPage();
            string ID = invoicesPage.GetFirstInvoiceID();
            invoicesPage.ClickFirstLine();

            invoiceDetails.AddService(serviceName, "10");

            var invoiceFooter = invoiceDetails.ClickOnInvoiceFooter();

            // On vérifie que la ligne Airport Tax est présente et qu'une valeur lui est associée
            var airportTaxOK = invoiceFooter.IsAirportTaxPresent(currency);
            Assert.IsTrue(airportTaxOK, "L'airport tax n'est pas présente pour l'invoice créée.");

            invoicesPage = invoiceFooter.BackToList();
            invoicesPage.ResetFilters();

            Assert.IsTrue(invoicesPage.GetFirstInvoiceID().Contains(ID), "L'invoice créée n'apparaît pas dans la liste des invoices.");
        }

        // ________________________________________________________FIN CREATE_INVOICE_______________________________________________

        // ________________________________________________________FILTER_INVOICE___________________________________________________
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_Filter_SearchByInvoiceNumber()
        {
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string serviceName = TestContext.Properties["InvoiceService"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();


            //Act
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
            var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customer, true);

            string ID = invoiceDetails.GetInvoiceNumber();
            invoiceDetails.AddService(serviceName, "10");

            // Validation de l'invoice
            invoiceDetails.Validate();

            // Récupération de l'identifiant de l'invoice créée
            string number = invoiceDetails.GetInvoiceNumber();
            invoicesPage = invoiceDetails.BackToList();
            invoicesPage.ResetFilters();
            // On filtre pour conserver l'invoice que l'on vient de créer
            invoicesPage.Filter(InvoicesPage.FilterType.ByNumber, number);

            // Assert
            var invoiceID = invoicesPage.GetFirstInvoiceID().Contains(ID);
            var invoiceNumber = invoicesPage.GetFirstInvoiceNumber().Contains(number);
            Assert.IsTrue(invoiceID, String.Format(MessageErreur.FILTRE_ERRONE, "'InvoiceNumber'"));
            Assert.IsTrue(invoiceNumber, String.Format(MessageErreur.FILTRE_ERRONE, "'InvoiceNumber'"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_Filter_Sites()
        {
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string serviceName = TestContext.Properties["InvoiceService"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            invoicesPage.ResetFilters();
            invoicesPage.Filter(InvoicesPage.FilterType.Site, site);

            if (invoicesPage.CheckTotalNumber() < 20)
            {
                // Create
                var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
                var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customer, true);

                // Ajout d'un service et retour à la page invoices
                invoiceDetails.AddService(serviceName, "10");
                invoicesPage = invoiceDetails.BackToList();
            }

            if (!invoicesPage.isPageSizeEqualsTo100())
            {
                invoicesPage.PageSize("8");
                invoicesPage.PageSize("100");
            }
            var existsite = invoicesPage.VerifySite(site);
            //Assert
            Assert.IsTrue(existsite, String.Format(MessageErreur.FILTRE_ERRONE, "'Site'"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_Filter_Customers()
        {
            //prepare
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string serviceName = TestContext.Properties["InvoiceService"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            invoicesPage.ResetFilters();
            invoicesPage.Filter(InvoicesPage.FilterType.Customer, customer);

            if (invoicesPage.CheckTotalNumber() < 20)
            {
                // Create a new invoice
                var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
                var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customer, true);

                // Ajout d'un service et retour à la page invoices
                invoiceDetails.AddService(serviceName, "10");
                invoicesPage = invoiceDetails.BackToList();
            }

            if (!invoicesPage.isPageSizeEqualsTo100())
            {
                invoicesPage.PageSize("8");
                invoicesPage.PageSize("100");
            }
            bool existcustomer = invoicesPage.VerifyCustomer(customer);

            //Assert
            Assert.IsTrue(existcustomer, String.Format(MessageErreur.FILTRE_ERRONE, "'Customer'"));
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_Filter_Date()
        {
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string serviceName = TestContext.Properties["InvoiceService"].ToString();
            List<DateTime> dateTimes = new List<DateTime> { fromdate.AddDays(-2), fromdate, GenerateRandomDate(fromdate, todate), todate, todate.AddDays(2) };
            List<string> numbersinvoice = new List<string>();
            List<bool> listOfResultInvoiceAfterSearch = new List<bool> { false, true, true, true, false };

            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            invoicesPage.ResetFilters();

            for (int i = 0; i < 5; i++)
            {
                // Create
                var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
                var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, dateTimes[i], customer, true);
                var isinvoicecreate = invoiceDetails.IsOnInvoiceDetailsPage();
                Assert.IsTrue(isinvoicecreate, "L'invoice créée n'apparaît pas dans la liste des invoices.");
                invoiceDetails.AddService(serviceName, "10");
                invoiceDetails.Validate();
                numbersinvoice.Add(invoiceDetails.GetInvoiceNumber());
                invoicesPage = invoiceDetails.BackToList();
                invoicesPage.Filter(InvoicesPage.FilterType.DateFrom, fromdate);
                invoicesPage.Filter(InvoicesPage.FilterType.DateTo, todate);
                //Assert
                VerifyInvoiceManuelDate(invoicesPage, numbersinvoice[i], listOfResultInvoiceAfterSearch[i], "le filtre sur Date From n est pas appliqué");
                invoicesPage.ResetFilters();
            }
        }

        private void VerifyInvoiceManuelDate(InvoicesPage invoicepage, string invoiceNumber, bool shouldExist, string errorMessage)
        {
            invoicepage.Filter(InvoicesPage.FilterType.ByNumber, invoiceNumber);
            //var nbDraft = invoicepage.CheckTotalNumber();

            if (shouldExist)
            {
                Assert.IsTrue(invoicepage.VerifyInvoiceNumberExist(invoiceNumber), errorMessage);
            }
            else
            {
                Assert.IsFalse(invoicepage.VerifyInvoiceNumberExist(invoiceNumber), errorMessage);
            }
        }

        public static DateTime GenerateRandomDate(DateTime fromdate, DateTime todate)
        {
            Random random = new Random();

            // Calculer l'intervalle en jours entre les deux dates
            int range = (todate - fromdate).Days;

            // Générer un nombre de jours aléatoire et l'ajouter à fromdate
            return fromdate.AddDays(random.Next(range));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_Filter_InvoiceStep_All()
        {
            //Prepare
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            invoicesPage.ResetFilters();

            if (invoicesPage.CheckTotalNumber() < 20)
            {
                // Create
                var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
                var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customer, true);

                invoicesPage = invoiceDetails.BackToList();
            }

            if (!invoicesPage.isPageSizeEqualsTo100())
            {
                invoicesPage.PageSize("8");
                invoicesPage.PageSize("100");
            }

            // Affichage des filtres InvoiceSteps s'ils sont cachés
            invoicesPage.ShowInvoiceStep();
            invoicesPage.FilterInvoiceStep(InvoicesPage.FilterInvoiceStepType.DRAFT, true);
            var nbDraft = invoicesPage.CheckTotalNumber();

            invoicesPage.FilterInvoiceStep(InvoicesPage.FilterInvoiceStepType.VALIDATED_ACCOUNTED, true);
            var nbValidated = invoicesPage.CheckTotalNumber();

            invoicesPage.FilterInvoiceStep(InvoicesPage.FilterInvoiceStepType.ALL, true);
            var nbTotal = invoicesPage.CheckTotalNumber();

            Assert.AreEqual(nbTotal, (nbDraft + nbValidated), String.Format(MessageErreur.FILTRE_ERRONE, "'ALL'"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_Filter_InvoiceStep_Draft()
        {
            //Prepare
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string serviceName = TestContext.Properties["InvoiceService"].ToString();
            string Qty = "10";

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            invoicesPage.ResetFilters();
            invoicesPage.ShowInvoiceStep();
            invoicesPage.FilterInvoiceStep(InvoicesPage.FilterInvoiceStepType.DRAFT, true);

            if (invoicesPage.CheckTotalNumber() < 20)
            {
                // Create
                var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
                var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customer, true);

                // Ajout d'un service et retour à la page invoices
                invoiceDetails.AddService(serviceName, Qty);
                invoicesPage = invoiceDetails.BackToList();
            }

            if (!invoicesPage.isPageSizeEqualsTo100())
            {
                invoicesPage.PageSize("8");
                invoicesPage.PageSize("100");
            }
            //Assert
            bool isDraftOK = invoicesPage.IsDraft();
            Assert.IsTrue(isDraftOK, String.Format(MessageErreur.FILTRE_ERRONE, "'DRAFT'"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_Filter_InvoiceStep_Validated()
        {
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string serviceName = TestContext.Properties["InvoiceService"].ToString();
            string Qty = "10";

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            invoicesPage.ResetFilters();
            invoicesPage.ShowInvoiceStep();
            invoicesPage.FilterInvoiceStep(InvoicesPage.FilterInvoiceStepType.VALIDATED_ACCOUNTED, true);

            if (invoicesPage.CheckTotalNumber() < 10)
            {
                // Create
                var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
                var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customer, true);

                // Ajout d'un service
                invoiceDetails.AddService(serviceName, Qty);
                invoiceDetails.Validate();
                invoiceDetails.ConfirmValidation();
                invoicesPage = invoiceDetails.BackToList();
            }

            if (!invoicesPage.isPageSizeEqualsTo100())
            {
                invoicesPage.PageSize("8");
                invoicesPage.PageSize("100");
            }

            // Assert
            bool isValidateOk = invoicesPage.CheckValidation(true);
            Assert.IsTrue(isValidateOk, String.Format(MessageErreur.FILTRE_ERRONE, "'VALIDATED'"));
        }



        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_Filter_InvoiceStep_Accounted()
        {
            //prepare
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string serviceName = TestContext.Properties["InvoiceService"].ToString();
            string journalInvoice = TestContext.Properties["Journal_Invoice"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            homePage.SetSageAutoEnabled(site, false);

            //Act
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            try
            {
                invoicesPage.ClearDownloads();

                invoicesPage.ResetFilters();
                invoicesPage.ShowInvoiceStep();
                invoicesPage.FilterInvoiceStep(InvoicesPage.FilterInvoiceStepType.ACCOUNTED, true);

                if (invoicesPage.CheckTotalNumber() < 20)
                {
                    // Manipulation pour permettre export SAGE 
                    var accountingParametersPage = homePage.GoToParameters_AccountingPage();
                    accountingParametersPage.GoToTab_Journal();
                    accountingParametersPage.EditJournal(site, journalInvoice);

                    // Create
                    invoicesPage = homePage.GoToAccounting_InvoicesPage();
                    var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
                    var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customer, true);

                    // Ajout d'un service
                    invoiceDetails.AddService(serviceName, "10");

                    // Validation de l'invoice
                    invoiceDetails.Validate();

                    // Export vers SAGE
                    invoiceDetails.ExportSage();

                    invoicesPage = invoiceDetails.BackToList();
                    invoicesPage.ResetFilters();
                    invoicesPage.ShowInvoiceStep();
                    invoicesPage.FilterInvoiceStep(InvoicesPage.FilterInvoiceStepType.ACCOUNTED, true);
                }

                if (!invoicesPage.isPageSizeEqualsTo100())
                {
                    invoicesPage.PageSize("8");
                    invoicesPage.PageSize("100");
                }

                // Assert
                Assert.IsTrue(invoicesPage.IsSentToSage(), String.Format(MessageErreur.FILTRE_ERRONE, "'ACCOUNTED'"));

            }
            finally
            {
                invoicesPage.ResetFilters();
            }

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_Filter_Show_All()
        {
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();


            //Act
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            invoicesPage.ResetFilters();
            invoicesPage.Filter(InvoicesPage.FilterType.DateFrom, DateUtils.Now.AddDays(-10));

            if (invoicesPage.CheckTotalNumber() < 20)
            {
                // Create
                var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
                var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customer, true);

                invoicesPage = invoiceDetails.BackToList();
                invoicesPage.Filter(InvoicesPage.FilterType.DateFrom, DateUtils.Now.AddDays(-10));
            }


            WebDriver.Navigate().Refresh();
            invoicesPage.ShowInvoiceShow();
            invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowAll, true);
            // Les ids à vider
            invoicesPage.PageSize("1000");
            List<string> idsTotal = invoicesPage.GetAllIds();
            var nbTotal = invoicesPage.CheckTotalNumber();


            WebDriver.Navigate().Refresh();
            // Affichage des filtres InvoiceSteps s'ils sont cachés
            invoicesPage.ShowInvoiceShow();
            invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowFlightInvoice, true);
            // Les ids valides
            invoicesPage.PageSize("1000");
            List<string> idsShowFlightInvoice = invoicesPage.GetAllIds();
            foreach (string id in idsShowFlightInvoice)
            {
                Assert.IsTrue(idsTotal.Contains(id), "ShowFilghtInvoice " + id + " pas dans total");
                idsTotal.Remove(id);
            }
            var nbFlight = invoicesPage.CheckTotalNumber();

            invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowDeliveryInvoice, true);
            invoicesPage.PageSize("1000");
            List<string> idsShowDeliveryInvoice = invoicesPage.GetAllIds();
            foreach (string id in idsShowDeliveryInvoice)
            {
                Assert.IsTrue(idsTotal.Contains(id), "ShowDeliveryInvoice " + id + " pas dans total");
                idsTotal.Remove(id);
            }
            var nbDelivery = invoicesPage.CheckTotalNumber();

            invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowCustomerOrderInvoice, true);
            invoicesPage.PageSize("1000");
            List<string> idsShowCustomerOrderInvoice = invoicesPage.GetAllIds();
            foreach (string id in idsShowCustomerOrderInvoice)
            {
                Assert.IsTrue(idsTotal.Contains(id), "ShowCustomerOrderInvoice " + id + " pas dans total");
                idsTotal.Remove(id);
            }
            var nbCustomerOrder = invoicesPage.CheckTotalNumber();

            invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowManualInvoice, true);
            invoicesPage.PageSize("1000");
            List<string> idsShowManualInvoice = invoicesPage.GetAllIds();
            int negatif = 0;
            foreach (string id in idsShowManualInvoice)
            {
                if (idsShowFlightInvoice.Contains(id))
                {
                    negatif++;
                }
                Assert.IsTrue(idsShowFlightInvoice.Contains(id) || idsTotal.Contains(id), "ShowManualInvoice " + id + " pas dans total");
                idsTotal.Remove(id);
            }
            var nbManual = invoicesPage.CheckTotalNumber();

            Assert.AreEqual(0, idsTotal.Count, "Ids manquants :" + String.Join(",", idsTotal));

            Assert.AreEqual(nbTotal, (nbFlight + nbDelivery + nbCustomerOrder + nbManual - negatif), String.Format(MessageErreur.FILTRE_ERRONE, "'Show All invoices (partie 1)'") + " : " + nbTotal + "==" + nbFlight + "+" + nbDelivery + "+" + nbCustomerOrder + "+" + nbManual);
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_Filter_Show_FlightInvoices()
        {
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string service = TestContext.Properties["InvoiceService"].ToString();
            string downloadPath = TestContext.Properties["DownloadsPath"].ToString();
            string customerPickMethod = TestContext.Properties["AllFlightsInSelectedPeriod"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();


            //Invoice part
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            invoicesPage.ResetFilters();

            invoicesPage.ClearDownloads();

            invoicesPage.ShowInvoiceShow();
            invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowFlightInvoice, true);
            invoicesPage.Filter(InvoicesPage.FilterType.DateTo, DateUtils.Now);
            invoicesPage.Filter(InvoicesPage.FilterType.DateFrom, DateUtils.Now.AddMonths(-1));

            if (invoicesPage.CheckTotalNumber() < 20)
            {
                // Flight  part
                CreateFlightForInvoice(homePage, site, customer, service);

                invoicesPage = homePage.GoToAccounting_InvoicesPage();
                var invoiceCreateAutoModalpage = invoicesPage.AutoInvoiceCreatePage();
                invoiceCreateAutoModalpage.FillField_CreateNewAutoInvoice(customer, site, customerPickMethod);
                var invoiceDetails = invoiceCreateAutoModalpage.Submit();
                invoicesPage.WaitPageLoading();
                invoicesPage = invoiceDetails.BackToList();
            }

            invoicesPage.ResetFilters();
            invoicesPage.ShowInvoiceShow();
            invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowFlightInvoice, true);
            invoicesPage.Filter(InvoicesPage.FilterType.DateTo, DateUtils.Now);
            invoicesPage.Filter(InvoicesPage.FilterType.DateFrom, DateUtils.Now.AddMonths(-1));

            // On exporte les résultats sous la forme d'un fichier Excel (dont on récupère le nom)
            // Export du fichier au format Excel
            invoicesPage.ExportExcelFile();

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            var correctDownloadedFile = invoicesPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern n'a été téléchargé.");

            // Récupération du nom du fichier et construction de l'URL du fichier Excel à ouvrir   
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadPath, fileName);

            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber(INVOICE_EXCEL_SHEET, filePath);
            bool result = OpenXmlExcel.ReadAllDataInColumn("Type", INVOICE_EXCEL_SHEET, filePath, "Flight");

            //Assert
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.IsTrue(result, MessageErreur.EXCEL_DONNEES_KO);
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_Filter_Show_CustomerOrderInvoices()
        {

            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string itemName = "HOT MEAL TECH CREW UAE 2 PAX"; // TestContext.Properties["InvoiceService"].ToString();
            string downloadPath = TestContext.Properties["DownloadsPath"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();


            //Invoice part
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            invoicesPage.ResetFilters();

            invoicesPage.ClearDownloads();

            invoicesPage.ShowInvoiceShow();
            invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowCustomerOrderInvoice, true);
            invoicesPage.Filter(InvoicesPage.FilterType.DateTo, DateUtils.Now);
            invoicesPage.Filter(InvoicesPage.FilterType.DateFrom, DateUtils.Now.AddMonths(-1));

            if (invoicesPage.CheckTotalNumber() < 20)
            {
                // CustomerOrder  part
                var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
                var customerOrderCreatePage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreatePage.FillField_CreatNewCustomerOrder(site, customer, aircraft);
                var customerOrderItem = customerOrderCreatePage.Submit();

                customerOrderItem.AddNewItem(itemName, "10");
                customerOrderItem.ValidateCustomerOrder();

                var generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
                var orderNumb = generalInfo.GetOrderNumber();

                invoicesPage = homePage.GoToAccounting_InvoicesPage();
                var invoiceCreateAutoModalpage = invoicesPage.AutoInvoiceCreatePage();
                invoiceCreateAutoModalpage.FillField_CreateNewAutoInvoice(customer, site, CustomerPickMethod.AllCustomerOrdersInSelectedPeriod);
                var invoiceDetails = invoiceCreateAutoModalpage.FillFieldSelectSomes(int.Parse(orderNumb));
                //var invoiceDetails = invoiceCreateAutoModalpage.Submit();

                string idInvoice = null;
                string numInvoice = null;
                try
                {
                    // un ID si non validé, un number si validé
                    idInvoice = invoiceDetails.GetInvoiceNumber();
                    invoiceDetails.Validate();
                    numInvoice = invoiceDetails.GetInvoiceNumber();
                }
                catch
                {
                    //on est dans index à la place d'invoiceDetails
                    idInvoice = invoicesPage.GetFirstInvoiceID();
                    invoiceDetails = invoicesPage.SelectFirstInvoice();
                    invoiceDetails.Validate();
                    numInvoice = invoiceDetails.GetInvoiceNumber();
                }

                invoicesPage = invoiceDetails.BackToList();
                invoicesPage.ResetFilters();
                invoicesPage.ShowInvoiceShow();
                invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowCustomerOrderInvoice, true);
                invoicesPage.Filter(InvoicesPage.FilterType.DateTo, DateUtils.Now);
                invoicesPage.Filter(InvoicesPage.FilterType.DateFrom, DateUtils.Now.AddMonths(-1));
            }

            // On exporte les résultats sous la forme d'un fichier Excel (dont on récupère le nom)
            // Export du fichier au format Excel
            invoicesPage.ExportExcelFile();

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            var correctDownloadedFile = invoicesPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier ne correspondant au pattern n'a été téléchargé.");

            // Récupération du nom du fichier et construction de l'URL du fichier Excel à ouvrir   
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadPath, fileName);

            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber(INVOICE_EXCEL_SHEET, filePath);
            bool result = OpenXmlExcel.ReadAllDataInColumn("Type", INVOICE_EXCEL_SHEET, filePath, "CustomerOrder");

            //Assert
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.IsTrue(result, MessageErreur.EXCEL_DONNEES_KO);
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_Filter_Show_ManualInvoices()
        {
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string serviceName = TestContext.Properties["InvoiceService"].ToString();
            string downloadPath = TestContext.Properties["DownloadsPath"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();


            //Act
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            invoicesPage.ResetFilters();

            invoicesPage.ClearDownloads();

            invoicesPage.ShowInvoiceShow();
            invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowManualInvoice, true);
            invoicesPage.Filter(InvoicesPage.FilterType.DateTo, DateUtils.Now);
            invoicesPage.Filter(InvoicesPage.FilterType.DateFrom, DateUtils.Now.AddDays(-1));

            if (invoicesPage.CheckTotalNumber() < 20)
            {
                // Create
                var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
                var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customer, true);

                // Ajout d'un service et retour à la page invoices
                invoiceDetails.AddService(serviceName, "10");
                invoicesPage = invoiceDetails.BackToList();
                invoicesPage.ResetFilters();
                invoicesPage.ShowInvoiceShow();
                invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowManualInvoice, true);
                invoicesPage.Filter(InvoicesPage.FilterType.DateTo, DateUtils.Now);
                invoicesPage.Filter(InvoicesPage.FilterType.DateFrom, DateUtils.Now.AddDays(-1));
            }

            // On exporte les résultats sous la forme d'un fichier Excel (dont on récupère le nom)
            // Export du fichier au format Excel
            invoicesPage.ExportExcelFile();

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            var correctDownloadedFile = invoicesPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier ne correspondant au pattern n'a été téléchargé.");

            // Récupération du nom du fichier et construction de l'URL du fichier Excel à ouvrir   
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadPath, fileName);

            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber(INVOICE_EXCEL_SHEET, filePath);
            bool result = OpenXmlExcel.ReadAllDataInColumn("Type", INVOICE_EXCEL_SHEET, filePath, "Manual");

            //Assert
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.IsTrue(result, MessageErreur.EXCEL_DONNEES_KO);
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_Filter_Show_AllInvoices()
        {
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();


            //Act
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            invoicesPage.ResetFilters();

            if (invoicesPage.CheckTotalNumber() < 20)
            {
                // Create
                var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
                var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customer, true);

                invoicesPage = invoiceDetails.BackToList();
            }

            if (!invoicesPage.isPageSizeEqualsTo100())
            {
                invoicesPage.PageSize("8");
                invoicesPage.PageSize("100");
            }

            // Affichage des filtres InvoiceSteps s'ils sont cachés
            invoicesPage.ShowInvoiceShow();
            invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowNotSentToSage, true);
            var nbNotSendToSage = invoicesPage.CheckTotalNumber();

            invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowExportedForSageManually, true);
            var nbExportedManually = invoicesPage.CheckTotalNumber();

            invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowExportedForSageAutomatically, true);
            var nbExportedAutomatically = invoicesPage.CheckTotalNumber();

            invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowAllInvoices, true);
            var nbTotal = invoicesPage.CheckTotalNumber();

            Assert.AreEqual(nbTotal, (nbNotSendToSage + nbExportedManually + nbExportedAutomatically), String.Format(MessageErreur.FILTRE_ERRONE, "'Show All invoices (partie 2)'"));//
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_Filter_Show_NotSentByMail()
        {
            //Prepare
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string serviceName = TestContext.Properties["InvoiceService"].ToString();
            string Qty = "10";

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            invoicesPage.ResetFilters();
            invoicesPage.ShowInvoiceShow();
            invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowNotSentByMailOnly, true);

            if (invoicesPage.CheckTotalNumber() < 10)
            {
                // Create
                var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
                var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customer, true);

                // Ajout d'un service et retour à la page invoices
                invoiceDetails.AddService(serviceName, Qty);
                invoicesPage = invoiceDetails.BackToList();
            }
            if (!invoicesPage.isPageSizeEqualsTo100())
            {
                invoicesPage.PageSize("8");
                invoicesPage.PageSize("100");
            }

            bool isSendMailOk = invoicesPage.IsSentByMail();
            Assert.IsFalse(isSendMailOk, String.Format(MessageErreur.FILTRE_ERRONE, "'Not Sent by mail'"));
        }

        [Ignore]
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_Filter_Show_NotSentToSage()
        {
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string serviceName = TestContext.Properties["InvoiceService"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();


            //Act
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            invoicesPage.ResetFilters();
            invoicesPage.ShowInvoiceShow();
            invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowNotSentToSage, true);

            if (invoicesPage.CheckTotalNumber() < 20)
            {
                // Create
                var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
                var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customer, true);

                // Ajout d'un service et retour à la page invoices
                invoiceDetails.AddService(serviceName, "10");
                invoicesPage = invoiceDetails.BackToList();
            }

            if (!invoicesPage.isPageSizeEqualsTo100())
            {
                invoicesPage.PageSize("8");
                invoicesPage.PageSize("100");
            }

            Assert.IsFalse(invoicesPage.IsSentToSage(), String.Format(MessageErreur.FILTRE_ERRONE, "'Not Sent to SAGE'"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_Filter_Show_NotValidated()
        {
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string serviceName = TestContext.Properties["InvoiceService"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();


            //Act
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            invoicesPage.ResetFilters();
            invoicesPage.ShowInvoiceShow();
            invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowNotValidatedOnly, true);

            if (invoicesPage.CheckTotalNumber() < 20)
            {
                // Create
                var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
                var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customer, true);

                // Ajout d'un service et retour à la page invoices
                invoiceDetails.AddService(serviceName, "10");
                invoicesPage = invoiceDetails.BackToList();
            }

            if (!invoicesPage.isPageSizeEqualsTo100())
            {
                invoicesPage.PageSize("8");
                invoicesPage.PageSize("100");
            }
            bool checkValidation = invoicesPage.CheckValidation(false);

            Assert.IsFalse(checkValidation, String.Format(MessageErreur.FILTRE_ERRONE, "'Not validated'"));
        }

        [Ignore]
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_Filter_SentToSAGEInErrorOnly()
        {
            //Prepare
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string serviceName = TestContext.Properties["InvoiceService"].ToString();
            string journalInvoice = TestContext.Properties["Journal_Invoice"].ToString();

            // Log in
            var homePage = LogInAsAdmin();


            // Config Export Sage Auto
            homePage.SetSageAutoEnabled(site, true, "Invoice");

            try
            {
                // Parameter - Accounting --> Journal KO pour le test
                VerifyAccountingJournal(homePage, site, "");

                //Act
                var invoicesPage = homePage.GoToAccounting_InvoicesPage();
                invoicesPage.ResetFilters();
                invoicesPage.ShowInvoiceShow();
                invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowSentToSageAndInError, true);

                if (invoicesPage.CheckTotalNumber() < 20)
                {
                    var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
                    var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customer, false);

                    // Ajout d'un service
                    invoiceDetails.AddService(serviceName, "10");
                    invoiceDetails.Validate();

                    var invoiceAccounting = invoiceDetails.ClickOnAccounting();
                    Assert.AreNotEqual("", invoiceAccounting.GetErrorMessage(), "Le code journal est manquant mais aucun message d'erreur n'est présent dans l'onglet Accounting.");

                    invoicesPage = invoiceAccounting.BackToList();
                    invoicesPage.ResetFilters();
                    invoicesPage.ShowInvoiceShow();
                    invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowSentToSageAndInError, true);
                }

                if (!invoicesPage.isPageSizeEqualsTo100())
                {
                    invoicesPage.PageSize("8");
                    invoicesPage.PageSize("100");
                }

                Assert.IsTrue(invoicesPage.IsSentToSAGEInErrorOnly(), String.Format(MessageErreur.FILTRE_ERRONE, "'Show sent to SAGE and in error only'"));
            }
            finally
            {
                VerifyAccountingJournal(homePage, site, journalInvoice);
                homePage.SetSageAutoEnabled(site, false);
            }
        }

        [Ignore]// aucun pays n'utilise SAGE AUTO
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_Filter_ShowOnTLWaitingForSAGEPush()
        {
            //Prepare
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string serviceName = TestContext.Properties["InvoiceService"].ToString();

            // Log in
            var homePage = LogInAsAdmin();


            // Config Export Sage Auto
            homePage.SetSageAutoEnabled(site, true, "Invoice");

            try
            {
                //Act
                var invoicesPage = homePage.GoToAccounting_InvoicesPage();
                invoicesPage.ResetFilters();
                invoicesPage.ShowInvoiceShow();
                invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowOnTLWaitingForSagePush, true);

                if (invoicesPage.CheckTotalNumber() < 20)
                {
                    var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
                    var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customer, false);

                    // Ajout d'un service
                    invoiceDetails.AddService(serviceName, "10");
                    invoiceDetails.Validate();

                    invoicesPage = invoiceDetails.BackToList();
                    invoicesPage.ResetFilters();
                    invoicesPage.ShowInvoiceShow();
                    invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowOnTLWaitingForSagePush, true);
                }

                if (invoicesPage.CheckTotalNumber() > 0 && !invoicesPage.isPageSizeEqualsTo100())
                {
                    invoicesPage.PageSize("8");
                    invoicesPage.PageSize("100");
                }

                Assert.IsTrue(invoicesPage.IsWaitingForSAGEPush(), String.Format(MessageErreur.FILTRE_ERRONE, "'Show on TL waiting for SAGE push'"));
            }
            finally
            {
                homePage.SetSageAutoEnabled(site, false);
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_Filter_ExportedForSageManually()
        {
            //Prepare
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string serviceName = TestContext.Properties["InvoiceService"].ToString();
            string journalInvoice = TestContext.Properties["Journal_Invoice"].ToString();

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            homePage.ClearDownloads();

            try
            {
                // Config pour export sage manuel
                homePage.SetSageAutoEnabled(site, false);

                var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

                // Parameter - Accounting --> Journal
                Assert.IsTrue(VerifyAccountingJournal(homePage, site, journalInvoice), "problème lors de la création du journal BCN");

                //Act
                var invoicesPage = homePage.GoToAccounting_InvoicesPage();
                invoicesPage.ResetFilters();

                // Create
                var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
                var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customer, false);

                // Ajout d'un service
                invoiceDetails.AddService(serviceName, "10");
                invoiceDetails.Validate();

                var invoiceNumber = invoiceDetails.GetInvoiceNumber();

                invoicesPage = invoiceDetails.BackToList();

                invoicesPage.ResetFilters();
                invoicesPage.Filter(InvoicesPage.FilterType.ByNumber, invoiceNumber);
                invoicesPage.ShowInvoiceShow();
                invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowExportedForSageManually, true);

                Assert.AreEqual(0, invoicesPage.CheckTotalNumber(), "L'invoice créée apparaît dans le résultat du filtre 'Exported for SAGE manually'" +
                    " alors qu'elle n'a pas été envoyée vers le SAGE.");

                WebDriver.Navigate().Refresh();
                invoicesPage.ResetFilters();
                invoicesPage.Filter(InvoicesPage.FilterType.ByNumber, invoiceNumber);
                invoiceDetails = invoicesPage.SelectFirstInvoice();

                invoiceDetails.ExportSage();
                invoicesPage = invoiceDetails.BackToList();

                invoicesPage.ResetFilters();
                invoicesPage.Filter(InvoicesPage.FilterType.ByNumber, invoiceNumber);
                invoicesPage.ShowInvoiceShow();
                invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowExportedForSageManually, true);

                //Assert
                Assert.AreEqual(1, invoicesPage.CheckTotalNumber(), String.Format(MessageErreur.FILTRE_ERRONE, "Exported for sage manually"));
            }
            finally
            {
                homePage.SetSageAutoEnabled(site, false);
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_Filter_SortBy()
        {
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string serviceName = TestContext.Properties["InvoiceService"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();


            //Act
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            invoicesPage.ResetFilters();

            if (invoicesPage.CheckTotalNumber() < 20)
            {
                // Create
                var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
                var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customer, true);

                // Ajout d'un service et retour à la page invoices
                invoiceDetails.AddService(serviceName, "10");
                invoicesPage = invoiceDetails.BackToList();
            }

            if (!invoicesPage.isPageSizeEqualsTo100())
            {
                invoicesPage.PageSize("8");
                invoicesPage.PageSize("100");
            }

            invoicesPage.Filter(InvoicesPage.FilterType.SortBy, "Name");
            var isSortedByName = invoicesPage.IsSortedByName();

            invoicesPage.Filter(InvoicesPage.FilterType.SortBy, "Number");
            var isSortedByNumber = invoicesPage.IsSortedByNumber();

            invoicesPage.Filter(InvoicesPage.FilterType.SortBy, "Id");
            var isSortedById = invoicesPage.IsSortedById();

            Assert.IsTrue(isSortedByName, String.Format(MessageErreur.FILTRE_ERRONE, "'Sort by name'"));
            Assert.IsTrue(isSortedByNumber, String.Format(MessageErreur.FILTRE_ERRONE, "'Sort by number'"));
            Assert.IsTrue(isSortedById, String.Format(MessageErreur.FILTRE_ERRONE, "'Sort by id'"));
        }

        // ________________________________________________________FIN FILTER_INVOICE_______________________________________________
        // ________________________________________________________INVOICE DETAILS__________________________________________________
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_AddService()
        {
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string serviceName = TestContext.Properties["InvoiceService"].ToString();

            string qtyValue = "10";

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
            var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customer, true);

            // Ajout d'un service
            bool isServiceAdded = invoiceDetails.AddService(serviceName, qtyValue);

            //Assert
            Assert.IsTrue(isServiceAdded, "Le service n'a pas été ajouté à l'invoice créée.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_AddFreePrice()
        {
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string workshop = TestContext.Properties["InvoiceWorkshop"].ToString();

            string namePrice = new Random().Next().ToString();
            string qtyValue = "100";
            string sellingPrice = "2";

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
            var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customer, true);

            // On créé une nouvelle ligne FreePrice
            var createFreePrice = invoiceDetails.AddFreePrice();
            createFreePrice.FillField_CreatNewFreePrice(namePrice, qtyValue, sellingPrice, workshop);
            invoiceDetails = createFreePrice.ValidateForInvoice();

            Assert.IsTrue(invoiceDetails.VerifyFreePrice(namePrice, qtyValue), "Le free price n'a pas été ajouté.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_CommentService()
        {
            //Prepare
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string qtyValue = "10";
            string comment = "I am a comment";
            string serviceName = "ServiceInvoice" + "-" + new Random().Next(10, 99).ToString();
            string serviceCategory = TestContext.Properties["CategorieBOB"].ToString();
            string deleteFrom = DateUtils.Now.AddMonths(-1).ToString("dd/MM/yyyy");
            string deleteTo = DateUtils.Now.AddMonths(3).ToString("dd/MM/yyyy");

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            try
            {
                // Act ajout  service
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
                var invoicesPage = homePage.GoToAccounting_InvoicesPage();
                var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
                var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customer, true);
                // Ajout d'un service
                invoiceDetails.AddService(serviceName, qtyValue);
                //Create a new service 
                invoiceDetails.CommentService(comment);

                // Assert
                bool isCommentAdded = invoiceDetails.VerifyComment(comment);
                Assert.IsTrue(isCommentAdded, "Le commentaire n'a pas été ajouté à l'invoice.");
            }
            finally
            {
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

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_DeleteService()
        {
            //Prepare
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string serviceName = "ServiceInvoice" + "-" + new Random().Next(10, 99).ToString();
            string qtyValue = "10";
            string serviceCategory = TestContext.Properties["CategorieBOB"].ToString();
            string deleteFrom = DateUtils.Now.AddMonths(-1).ToString("dd/MM/yyyy");
            string deleteTo = DateUtils.Now.AddMonths(3).ToString("dd/MM/yyyy");

            //Arrange
            HomePage homePage = LogInAsAdmin();
            try
            {
                // Act ajout  service
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
                //Act
                var invoicesPage = homePage.GoToAccounting_InvoicesPage();
                var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
                var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customer, true);

                // Ajout d'un service
                bool isServiceAdded = invoiceDetails.AddService(serviceName, qtyValue);
                Assert.IsTrue(isServiceAdded, "Le service n'a pas été ajouté à l'invoice créée.");

                //Assert
                bool isServiceDeleteOK = invoiceDetails.DeleteService();
                Assert.IsTrue(isServiceDeleteOK, "Le service n'a pas été supprimé de l'invoice.");
            }
            finally
            {
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

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_ValidateInvoice()
        {
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string serviceName = TestContext.Properties["InvoiceService"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
            var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customer, true);

            // Ajout d'un service
            invoiceDetails.AddService(serviceName, "10");
            invoiceDetails.Validate();

            // Récupération de l'identifiant de l'invoice créée
            string number = invoiceDetails.GetInvoiceNumber();
            invoicesPage = invoiceDetails.BackToList();
            invoicesPage.ResetFilters();
            invoicesPage.Filter(InvoicesPage.FilterType.ByNumber, number);

            // Assert
            var checkValidate = invoicesPage.CheckValidation(true);
            Assert.IsTrue(checkValidate, "L'invoice créée n'a pas été validée.");
        }

        // ________________________________________________________FIN INVOICE DETAILS______________________________________________

        // ________________________________________________________INVOICE FOOTER___________________________________________________
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_ConsultInvoiceFooter()
        {
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string serviceName = TestContext.Properties["InvoiceService"].ToString();
            string currency = TestContext.Properties["Currency"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();
            //Act
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
            var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customer, true);

            // Ajout d'un service
            invoiceDetails.AddService(serviceName, "10");

            var invoiceFooter = invoiceDetails.ClickOnInvoiceFooter();
            var totalTTC = invoiceFooter.GetTotalTTC(currency, decimalSeparatorValue);
            var localCurrency = invoiceFooter.GetLocalCurrency();

            // Assert
            var currencyinvoice = localCurrency.Contains(currency);
            Assert.IsTrue(currencyinvoice, "L'invoice footer ne contient pas la devise utilisée.");
            Assert.AreNotEqual(0, totalTTC, "Le total TTC est égal à 0");
        }

        // ________________________________________________________FIN INVOICE FOOTER_______________________________________________

        // ________________________________________________________INVOICE GENERAL INFORMATIONS_____________________________________

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_UpdateGlobalInformations()
        {
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string serviceName = TestContext.Properties["InvoiceService"].ToString();
            var comment = "I am a comment";
            string invoiceNumber = "";
            //Arrange
            var homePage = LogInAsAdmin();

            InvoicesPage invoicesPage = homePage.GoToAccounting_InvoicesPage();
            invoicesPage.ResetFilters();

            //Act
            try
            {
                var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
                var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customer, true);
                invoicesPage = homePage.GoToAccounting_InvoicesPage();
                invoiceNumber = invoicesPage.GetFirstInvoiceID().Substring(3);
                invoicesPage.ClickFirstLine();
                // Ajout d'un service
                invoiceDetails.AddService(serviceName, "10");

                // Click sur l'onglet General Informations et modification des valeurs de 'ApplyVAT' et 'Commentaire'
                var generalInformations = invoiceDetails.ClickOnGeneralInformation();

                //- comment
                Assert.AreNotEqual(generalInformations.GetComment(), comment, "Le commentaire de l'invoice est déjà égal à " + comment);
                generalInformations.SetComment(comment);
                //- date
                generalInformations.SetDate(DateUtils.Now.AddDays(-2));
                //-engagement No
                generalInformations.SetEngagementNo(436);
                //-apply VAT
                generalInformations.SetApplyVAT(false);
                //vérifier dans footer: enlève lignes TVA
                InvoiceFooterPage footerTab = generalInformations.ClickOnFooter();
                Assert.IsTrue(footerTab.CheckVAT(1), "Plus de 1 VAT");
                //footerTab.Get
                generalInformations = footerTab.ClickOnGeneralInformation();
                // Passage sur l'onglet Details et retour sur l'onglet GeneralInformations
                invoiceDetails = generalInformations.ClickOnDetails();
                generalInformations = invoiceDetails.ClickOnGeneralInformation();

                //- comment
                Assert.AreEqual(generalInformations.GetComment(), comment, "Le commentaire de l'invoice n'a pas été mis à jour.");
                //- date
                Assert.AreEqual(DateUtils.Now.AddDays(-2).ToString("dd/MM/yyyy"), generalInformations.GetDate(), "la date de l'invoice n'a pas été mis à jour.");
                //-engagement No
                Assert.AreEqual("436", generalInformations.GetEngagementNo(), "Le no d'engagement n'a pas été mise à jour.");
                //-apply VAT
                Assert.AreEqual(false, generalInformations.GetApplyVAT(), "Le Apply VAT n'a pas été mise à jour.");

                //-ajouter un comment et vérifier affichage sur index
                invoicesPage = generalInformations.BackToList();
                Assert.AreEqual(invoiceNumber, invoicesPage.GetFirstInvoiceID().Substring(3));
                var newComment = invoicesPage.CheckFirstComment();
                Assert.AreEqual(comment, newComment, "Pas d'icone commentaire dans l'index");

            }
            finally
            {

                if (!string.IsNullOrEmpty(invoicesPage.GetFirstInvoiceID().Substring(3)))
                {
                    invoicesPage.DeleteFirstInvoice();

                }
            }

        }

        // ________________________________________________________FIN INVOICE GENERAL INFORMATIONS_________________________________

        // ________________________________________________________TESTS PAGE INVOICES___________________________________________________
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_DeleteInvoice()
        {
            //Prepare
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string serviceName = "ServiceInvoice" + "-" + new Random().Next(10, 99).ToString();
            string serviceCategory = TestContext.Properties["CategorieBOB"].ToString();
            string deleteFrom = DateUtils.Now.AddMonths(-1).ToString("dd/MM/yyyy");
            string deleteTo = DateUtils.Now.AddMonths(3).ToString("dd/MM/yyyy");
            string Qty = "10";
            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            try
            {
                // Act ajout  service
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
                var invoicesPage = homePage.GoToAccounting_InvoicesPage();
                var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
                var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customer, true);

                var ID = invoiceDetails.GetInvoiceNumber();

                invoiceDetails.AddService(serviceName, Qty);
                invoicesPage = invoiceDetails.BackToList();

                invoicesPage.ResetFilters();
                invoicesPage.ShowInvoiceShow();
                invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowNotValidatedOnly, true);
                var number_invoices_invalid = invoicesPage.CheckTotalNumber();
                var isdeletedinvoice = invoicesPage.DeleteInvoiceById(ID);
                var number_invoices_invalid_after_deleted = invoicesPage.CheckTotalNumber();
                var comparenumberInvoice = number_invoices_invalid == number_invoices_invalid_after_deleted + 1;
                Assert.IsTrue(isdeletedinvoice && comparenumberInvoice, "L'invoice n'a pas été supprimée.");
            }
            finally
            {
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

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_PrintResultsPDFNewVersion()
        {
            //Prepare
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string serviceName = TestContext.Properties["InvoiceService"].ToString();
            string number = "";
            DateTime date = DateUtils.Now;
            string qty = "10";
            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            invoicesPage.ResetFilters();
            var firstDayOfMonth = new DateTime(date.Year, date.Month, 1);
            invoicesPage.Filter(InvoicesPage.FilterType.DateFrom, firstDayOfMonth);
            invoicesPage.ScrollToInvoiceStep();
            invoicesPage.ShowInvoiceStep();
            invoicesPage.FilterInvoiceStep(InvoicesPage.FilterInvoiceStepType.VALIDATED_ACCOUNTED, true);

            if (invoicesPage.CheckTotalNumber() == 0)
            {
                // Create
                var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
                var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, date, customer, true);

                invoiceDetails.AddService(serviceName, qty);
                invoiceDetails.Validate();

                // Récupération de l'identifiant de l'invoice créée
                number = invoiceDetails.GetInvoiceNumber();
                invoicesPage = invoiceDetails.BackToList();
            }
            else
            {
                var invoiceDetailsPage = invoicesPage.SelectFirstInvoice();
                number = invoiceDetailsPage.GetInvoiceNumber();
                invoicesPage = invoiceDetailsPage.BackToList();
            }

            if (!invoicesPage.isPageSizeEqualsTo100())
            {
                invoicesPage.PageSize("8");
                invoicesPage.PageSize("100");
            }

            invoicesPage.Filter(InvoicesPage.FilterType.ByNumber, number);
            //Clear
            invoicesPage.ClearDownloads();
            DeleteAllFileDownload();

            // Lancement du Print au format PDF
            var reportPage = invoicesPage.PrintInvoiceResults();
            bool isNewTabAdded = invoicesPage.GetTotalTabs();
            var isReportGenerated = reportPage.IsReportGenerated();

            reportPage.Close();

            invoicesPage.ClickPrintButton();

            //Assert
            Assert.IsTrue(isNewTabAdded, "Aucun Onglet n'a été ouvert");
            Assert.IsTrue(isReportGenerated, "L'application n'a pas pu générer le fichier attendu.");

        }


        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_PrintResultsZIPNewVersion()
        {
            //Prepare
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string serviceName = TestContext.Properties["InvoiceService"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            DateTime dateFrom = DateUtils.Now;
            DateTime dateTo = DateUtils.Now.AddMonths(1);

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            invoicesPage.ResetFilters();
            invoicesPage.Filter(InvoicesPage.FilterType.DateFrom, dateFrom);
            invoicesPage.Filter(InvoicesPage.FilterType.DateTo, dateTo);
            invoicesPage.ShowInvoiceStep();
            invoicesPage.FilterInvoiceStep(InvoicesPage.FilterInvoiceStepType.VALIDATED_ACCOUNTED, true);
            string number;

            // Create
            var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
            var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, dateFrom, customer, true);
            invoiceDetails.AddService(serviceName, "10");
            invoiceDetails.WaitPageLoading();
            invoiceDetails.Validate();

            // Récupération de l'identifiant de l'invoice créée
            number = invoiceDetails.GetInvoiceNumber();
            invoicesPage = invoiceDetails.BackToList();
            invoicesPage.ClearDownloads();
            DeleteAllFileDownload();
            ExportZipGenerique(invoicesPage, downloadsPath, number, dateFrom, dateTo);
        }

        private void ExportZipGenerique(InvoicesPage invoicesPage, string downloadsPath, string number, DateTime dateFrom, DateTime dateTo)
        {
            invoicesPage.ResetFilters();
            invoicesPage.Filter(InvoicesPage.FilterType.ByNumber, number);
            invoicesPage.Filter(InvoicesPage.FilterType.DateFrom, dateFrom);
            invoicesPage.Filter(InvoicesPage.FilterType.DateTo, dateTo);
            // On exporte les résultats sous la forme d'un fichier Zip (dont on récupère le nom)
            // Export du fichier au format Zip
            invoicesPage.ExportZipFile();

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            var correctDownloadedFile = invoicesPage.GetExportZipFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_SendInvoicesByMailNotValidated()
        {
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string serviceName = TestContext.Properties["InvoiceService"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();


            //Act
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            invoicesPage.ResetFilters();
            invoicesPage.ShowInvoiceShow();
            invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowNotValidatedOnly, true);
            invoicesPage.Filter(InvoicesPage.FilterType.DateFrom, DateUtils.Now);

            if (invoicesPage.CheckTotalNumber() < 20)
            {
                // Create
                var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
                var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customer, true);

                // Ajout d'un service
                invoiceDetails.AddService(serviceName, "10");
                invoicesPage = invoiceDetails.BackToList();
            }

            if (!invoicesPage.isPageSizeEqualsTo100())
            {
                invoicesPage.PageSize("8");
                invoicesPage.PageSize("100");
            }

            bool isSent = invoicesPage.SendByMail();

            Assert.IsFalse(isSent, "Le bouton Send est accessible malgré l'absence d'invoices validées.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_SendInvoicesByMailValidated()
        {
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string serviceName = TestContext.Properties["InvoiceService"].ToString();
            string userEmail = TestContext.Properties["Admin_UserName"].ToString();
            string number;
            string invoiceId;
            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            invoicesPage.ResetFilters();

            invoicesPage.Filter(InvoicesPage.FilterType.Customer, customer);
            DateTime date = DateUtils.Now;
            var firstDayOfMonth = new DateTime(date.Year, date.Month, 1);
            invoicesPage.Filter(InvoicesPage.FilterType.DateFrom, firstDayOfMonth);

            var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
            var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customer, true);

            // Ajout d'un service
            invoiceDetails.AddService(serviceName, "10");
            invoiceDetails.Validate();

            // Récupération de l'identifiant de l'invoice créée
            number = invoiceDetails.GetInvoiceNumber();
            invoicesPage = invoiceDetails.BackToList();
            invoiceId = invoicesPage.GetFirstInvoiceID().Split(' ')[1];

            invoicesPage.Filter(InvoicesPage.FilterType.ByNumber, number);

            bool isSent = invoicesPage.SendByMail();
            Assert.IsTrue(isSent, "L'envoi par mail a été effectué.");
            MailPage mailPage = invoicesPage.RedirectToOutlookMailbox();

            mailPage.FillFields_LogToOutlookMailbox(userEmail);

            Assert.IsTrue(mailPage.CheckIfSpecifiedOutlookMailExist("Winrest - 1 Invoice - " + site + " - " + invoiceId), "Mail Invoice n° " + invoiceId + " (quote) non trouvé");

            WebDriver.Navigate().Refresh();
            mailPage.Close();
            invoicesPage.ResetFilters();
            invoicesPage.Filter(InvoicesPage.FilterType.ByNumber, number);
            invoicesPage.ShowInvoiceShow();
            invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowNotSentToSage, true);
            invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowNotSentByMailOnly, true);

            var totalNumber = invoicesPage.CheckTotalNumber();
            Assert.AreEqual(0, totalNumber, "Les invoice n'ont pas été envoyées par mail");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_AffichageIndex()
        {
            //Prepare
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string serviceName = "ServiceInvoice" + "-" + new Random().Next(10, 99).ToString();
            string serviceCategory = TestContext.Properties["CategorieBOB"].ToString();
            string ID = string.Empty;
            DateTime dateInvoice = DateUtils.Now;
            string qty = "10";
            string deleteFrom = DateUtils.Now.AddMonths(-1).ToString("dd/MM/yyyy");
            string deleteTo = DateUtils.Now.AddMonths(3).ToString("dd/MM/yyyy");

            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            try
            {
                // Act ajout  service
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

                var invoicesPage = homePage.GoToAccounting_InvoicesPage();
                invoicesPage.ResetFilters();
                // Create
                var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
                var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, dateInvoice, customer, true);
                // Ajout d'un service
                invoiceDetails.AddService(serviceName, qty);
                invoiceDetails.WaitPageLoading();
                invoiceDetails.Validate();
                invoiceDetails.ConfirmValidation();
                ID = invoiceDetails.GetInvoiceNumber();
                invoiceDetails.WaitPageLoading();
                invoicesPage = invoiceDetails.BackToList();
                invoicesPage.ResetFilters();
                invoicesPage.Filter(InvoicesPage.FilterType.ByNumber, ID);

                //Assert
                var formattedDateInvoice = dateInvoice.ToString("dd/MM/yyyy");
                bool isVerifAffichageColOK = invoicesPage.VerifAffichageCol(ID, site, customer, formattedDateInvoice);
                Assert.IsTrue(isVerifAffichageColOK, "Les données ne correspondent pas à l'affichage de l'entête du tableau.");
            }
            finally
            {
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

        private FileInfo ExportGenerique(InvoicesPage invoicesPage, string customer)
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            // On filtre sur la date
            invoicesPage.Filter(InvoicesPage.FilterType.DateFrom, DateUtils.Now.AddDays(-1));

            // Le print ne fonctionne que s'il contient moins de 100 résultats
            if (invoicesPage.TotalNumber() > 200)
            {
                // Si on a plus de 100 résultats, on filtre par customer pour réduire le nombre
                invoicesPage.Filter(InvoicesPage.FilterType.Customer, customer);
            }

            // On exporte les résultats sous la forme d'un fichier Excel (dont on récupère le nom)
            // Export des données au format Excel
            invoicesPage.ExportExcelFile();

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            var correctDownloadedFile = invoicesPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

            return correctDownloadedFile;
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_GetCustomReportNewVersion()
        {
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string service = TestContext.Properties["InvoiceService"].ToString();
            string customerPickMethod = TestContext.Properties["AllFlightsInSelectedPeriod"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();


            // Flight  part
            var flightNumber = CreateFlightForInvoice(homePage, site, customer, service);

            // Create
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            invoicesPage.ResetFilters();

            invoicesPage.ClearDownloads();

            var invoiceCreateAutoModalpage = invoicesPage.AutoInvoiceCreatePage();
            invoiceCreateAutoModalpage.FillField_CreateNewAutoInvoice(customer, site, customerPickMethod);
            var invoiceDetails = invoiceCreateAutoModalpage.Submit();
            // on s'est perdu entre Invoice Details et Invoice index
            homePage.GoToWinrestHome();
            invoicesPage = homePage.GoToAccounting_InvoicesPage();
            invoicesPage.SelectFirstInvoice();
            invoiceDetails.Validate();

            invoicesPage = invoiceDetails.BackToList();

            var excelFile = CustomReportGenerique(invoicesPage, site);

            //Assert
            int resultNumber = OpenXmlExcel.GetExportResultNumber("Customs", excelFile.FullName);

            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);

            List<string> flightNbList = OpenXmlExcel.GetValuesInList("N° Vol", "Customs", excelFile.FullName);
            Assert.IsTrue(flightNbList.Contains(flightNumber));

            int offset = flightNbList.IndexOf(flightNumber);
            Assert.AreEqual(site, OpenXmlExcel.GetValuesInList("Site", "Customs", excelFile.FullName)[offset]);
            Assert.AreEqual(customer, OpenXmlExcel.GetValuesInList("Cie", "Customs", excelFile.FullName)[offset]);
        }

        private FileInfo CustomReportGenerique(InvoicesPage invoicesPage, string site)
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            // On vide le répertoire de téléchargement
            DeleteAllFileDownload();

            // On exporte les résultats sous la forme d'un fichier Excel (dont on récupère le nom)
            // Export des données au format Excel
            invoicesPage.ExportCustomFile(downloadsPath, site);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            var correctDownloadedFile = invoicesPage.GetExportCustomFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

            return correctDownloadedFile;
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_ValidateResults()
        {
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string serviceName = TestContext.Properties["InvoiceService"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            invoicesPage.ResetFilters();
            invoicesPage.ShowInvoiceShow();
            invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowNotValidatedOnly, true);
            invoicesPage.Filter(InvoicesPage.FilterType.DateFrom, DateUtils.Now);

            if (invoicesPage.CheckTotalNumber() < 20)
            {
                // Create
                var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
                var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customer, true);


                // Ajout d'un service et retour à la page invoices
                invoiceDetails.AddService(serviceName, "10");

                invoicesPage = invoiceDetails.BackToList();
                invoicesPage.ResetFilters();
                invoicesPage.ShowInvoiceShow();
                invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowNotValidatedOnly, true);
                invoicesPage.Filter(InvoicesPage.FilterType.DateFrom, DateUtils.Now);
            }

            if (!invoicesPage.isPageSizeEqualsTo100())
            {
                invoicesPage.PageSize("8");
                invoicesPage.PageSize("100");
            }

            // On valide l'ensemble des resultats
            invoicesPage.ValidateResults();

            WebDriver.Navigate().Refresh();
            invoicesPage.ResetFilters();
            invoicesPage.ShowInvoiceShow();
            invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowNotValidatedOnly, true);
            invoicesPage.Filter(InvoicesPage.FilterType.DateFrom, DateUtils.Now);

            // On recupère le nombre d'invoices non validées après traitement
            var newNbInvoicesNotValidated = invoicesPage.CheckTotalNumber();

            Assert.AreEqual(0, newNbInvoicesNotValidated, "Il reste des invoices non validées malgré l'utilisation " + "de la fonctionnalité de validation des résultats.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_DeleteUnvalidatedResults()
        {
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string serviceName = TestContext.Properties["InvoiceService"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            invoicesPage.ResetFilters();
            invoicesPage.ShowInvoiceShow();
            invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowNotValidatedOnly, true);

            if (invoicesPage.CheckTotalNumber() < 20)
            {
                // Create
                var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
                var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customer, true);


                // Ajout d'un service et retour à la page invoices
                invoiceDetails.AddService(serviceName, "10");
                invoicesPage = invoiceDetails.BackToList();
                invoicesPage.ResetFilters();
                invoicesPage.ShowInvoiceShow();
                invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowNotValidatedOnly, true);
            }

            if (!invoicesPage.isPageSizeEqualsTo100())
            {
                invoicesPage.PageSize("8");
                invoicesPage.PageSize("100");
            }

            // On valide l'ensemble des resultats
            invoicesPage.DeleteUnValidatedInvoices();

            WebDriver.Navigate().Refresh();

            invoicesPage.ResetFilters();
            invoicesPage.ShowInvoiceShow();
            invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowNotValidatedOnly, true);

            // On recupère le nombre d'invoices non validées après traitement
            var newNbInvoicesNotValidated = invoicesPage.CheckTotalNumber();

            Assert.AreEqual(0, newNbInvoicesNotValidated, "Il reste des invoices non validées malgré l'utilisation "
                 + "de la fonctionnalité de validation des résultats.");
        }

        // ________________________________________________________FIN TESTS PAGE INVOICES_______________________________________________

        // ________________________________________________________TESTS PAGE INVOICES ITEM______________________________________________

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_ExportItemsNewVersion()
        {
            //Prepare
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string serviceName = "ServiceInvoice" + "-" + new Random().Next(10, 99).ToString();
            string qtyValue = "10";
            string serviceCategory = TestContext.Properties["CategorieBOB"].ToString();
            string deleteFrom = DateUtils.Now.AddMonths(-1).ToString("dd/MM/yyyy");
            string deleteTo = DateUtils.Now.AddMonths(3).ToString("dd/MM/yyyy");

            // Log in
            HomePage homePage = LogInAsAdmin();
            try
            {
                // Act ajout  service
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
                //Act
                var invoicesPage = homePage.GoToAccounting_InvoicesPage();
                // Create
                var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
                var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customer, true);
                // Ajout d'un service
                invoiceDetails.AddService(serviceName, qtyValue);
                invoicesPage.ClearDownloads();
                DeleteAllFileDownload();
                // Export des données au format Excel
                invoiceDetails.ExportExcelFile();
                // On récupère les fichiers du répertoire de téléchargement
                DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
                FileInfo[] taskFiles = taskDirectory.GetFiles();
                var correctDownloadedFile = invoiceDetails.GetExportExcelFile(taskFiles);
                Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");
                // Récupération du nom du fichier et construction de l'URL du fichier Excel à ouvrir   
                var fileName = correctDownloadedFile.Name;
                var filePath = Path.Combine(downloadsPath, fileName);
                // Exploitation du fichier Excel
                int resultNumber = OpenXmlExcel.GetExportResultNumber("Invoices", filePath);

                //Assert
                Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
            }
            finally
            {
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

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_PrintItemsNewVersion()
        {
            Random rnd = new Random();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string serviceName = "ServiceInvoice" + "-" + new Random().Next(10, 99).ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string serviceCategory = TestContext.Properties["CategorieBOB"].ToString();
            string customerName = "customerInvoice" + "-" + rnd.Next().ToString();
            string customerCode = rnd.Next().ToString();
            string customerType = TestContext.Properties["CustomerType"].ToString();
            string deleteFrom = DateUtils.Now.AddMonths(-1).ToString("dd/MM/yyyy");
            string deleteTo = DateUtils.Now.AddMonths(3).ToString("dd/MM/yyyy");
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNamePdfBegin = "Print Invoice_-_";
            string DocFileNameZipBegin = "All_files_";
            string currency = TestContext.Properties["Currency"].ToString();
            long quantity = 10;
            double unitPrice = 15;
            string ID = string.Empty;
            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();
            //Create Customer 
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.Filter(CustomerPage.FilterType.Search, customerName);
            if (customersPage.CheckTotalNumber_display() == 0)
            {
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customerName, customerCode, customerType);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
            }
            //Create service  
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            if (servicePage.CheckTotalNumber() == 0)
            {
                var ServiceCreateModalPage = servicePage.ServiceCreatePage();
                ServiceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction, serviceCategory);
                var serviceGeneralInformationsPage = ServiceCreateModalPage.Create();
                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customerName, DateUtils.Now, DateUtils.Now.AddMonths(2));
                servicePage = pricePage.BackToList();
            }
            try
            {

                //Create ManualInvoice
                var invoicesPage = homePage.GoToAccounting_InvoicesPage();
                // Create
                var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
                var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customerName, true);

                ID = invoiceDetails.GetInvoiceNumber();
                // Ajout d'un service
                invoiceDetails.AddService(serviceName, "0");
                invoiceDetails.SelectFirstLine();
                invoiceDetails.Fill(null, quantity, unitPrice, null, null);

                string vatName = invoiceDetails.GetVatName();

                // Comparaison des taux VAT RATE
                var invoiceFooter = invoiceDetails.ClickOnInvoiceFooter();
                string taxName = invoiceFooter.getTaxName();
                string montantInvoice = invoiceFooter.GetTotalTTC(currency, decimalSeparatorValue).ToString();

                Assert.AreEqual(taxName, vatName, "Le taux de VAT dans le pied de facture et le VAT de service ne sont pas égaux.");

                var VATRATFromFooterInvoice = invoiceDetails.GetVATRATFromFooterInvoice().Replace("%", "").Trim();
                //var VATRATFromFooterInvoice = invoiceDetails.GetVATRATFromFooterInvoice().Replace("%", "").Trim();

                invoicesPage.ClearDownloads();
                DeleteAllFileDownload();

                //3.Cliquer sur print
                PrintReportPage reportPage = PrintItemGenerique(invoiceDetails);
                reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
                // cliquer sur All
                string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin, false);
                FileInfo fi = new FileInfo(trouve);
                fi.Refresh();
                Assert.IsTrue(fi.Exists, trouve + " non généré");

                PdfDocument document = PdfDocument.Open(fi.FullName);
                List<string> mots = new List<string>();
                foreach (Page p in document.GetPages())
                {
                    foreach (var mot in p.GetWords())
                    {
                        mots.Add(mot.Text);
                    }
                }
                invoiceDetails.WaitPageLoading();
                bool isVerifVatTotalOK = invoiceDetails.VerifVatTotal(mots, VATRATFromFooterInvoice, montantInvoice);

                Assert.IsTrue(isVerifVatTotalOK, "Pas d'affichage pour VAT RAT et Total invoice.");

            }
            finally
            {
                //Delete service
                servicePage = homePage.GoToCustomers_ServicePage();
                var serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.From, deleteFrom);
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.To, deleteTo);
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();

                //delete invoice 
                var invoicesPage = homePage.GoToAccounting_InvoicesPage();
                invoicesPage.ResetFilters();
                invoicesPage.ShowInvoiceShow();
                invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowNotValidatedOnly, true);
                invoicesPage.DeleteInvoiceById(ID);

            }
        }

        private PrintReportPage PrintItemGenerique(InvoiceDetailsPage invoiceDetails)
        {
            // Lancement du Print
            var reportPage = invoiceDetails.PrintInvoiceResults();
            var isReportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();

            invoiceDetails.ClickPrintButton();

            //Assert
            Assert.IsTrue(isReportGenerated, "L'application n'a pas pu générer le fichier attendu.");

            return reportPage;
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_SendItemsByMail()
        {
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string serviceName = TestContext.Properties["InvoiceService"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string userEmail = TestContext.Properties["Admin_UserName"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            invoicesPage.ResetFilters();
            // Create
            var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
            var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customer, true);
            // Ajout d'un service
            invoiceDetails.AddService(serviceName, "10");
            // Validation de l'invoice et envoi par mail
            invoiceDetails.Validate();
            string number = invoiceDetails.GetInvoiceNumber();

            invoiceDetails.SendByMail(userEmail);
            MailPage mailPage = invoiceDetails.RedirectToOutlookMailbox();

            mailPage.FillFields_LogToOutlookMailbox(userEmail);
            mailPage.WaitPageLoading();
            WebDriver.Navigate().Refresh();
            var checkmail = mailPage.CheckIfSpecifiedOutlookMailExist("Winrest - 1 Invoice - NEWREST ESPAÑA S.L. - " + site + " - " + DateUtils.Now.ToString("yyyyMM") + " - " + number);
            // Recherche du mail envoyé
            Assert.IsTrue(checkmail, "Mail Invoice n° " + number + " (quote) non trouvé");

            FileInfo trouve = mailPage.DownloadOutlookAttachmentPDF(downloadsPath, "Invoice_" + number + ".pdf");
            var fileexist = trouve.Exists;
            Assert.IsTrue(fileexist, "fichier joint non téléchargé");
            trouve.Delete();

            mailPage.DeleteCurrentOutlookMail();
            mailPage.Close();


            WebDriver.Navigate().Refresh();
            invoicesPage = invoiceDetails.BackToList();

            // Recherche de l'invoice validée pour voir si elle a été envoyée par mail
            invoicesPage.ResetFilters();
            invoicesPage.Filter(InvoicesPage.FilterType.ByNumber, number);

            // Assert
            var verifiermailenvoyer = invoicesPage.IsSentByMail();
            Assert.IsTrue(verifiermailenvoyer, "L'invoice créée n'a pas été envoyée par mail.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        [DeploymentItem("PageObjects\\Accounting\\Invoice\\pieceJointe.jpg")]
        [DeploymentItem("chromedriver.exe")]

        public void AC_INVO_SendItemsByMailWithAttachment()
        {
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string serviceName = TestContext.Properties["InvoiceService"].ToString();
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            var imagePath = TestContext.TestDeploymentDir + "\\pieceJointe.jpg";
            Assert.IsTrue(new FileInfo(imagePath).Exists);

            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();

            // Create
            var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
            var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customer, true);

            // Ajout d'un service
            invoiceDetails.AddService(serviceName, "10");

            // Validation de l'invoice et envoi par mail
            invoiceDetails.Validate();
            string number = invoiceDetails.GetInvoiceNumber();
            invoiceDetails.SendByMail(userMail, imagePath);
            invoiceDetails.WaitPageLoading();
            MailPage mailPage = invoiceDetails.RedirectToOutlookMailbox();
            mailPage.FillFields_LogToOutlookMailbox(userMail);
            WebDriver.Navigate().Refresh();
            // Recherche du mail envoyé
            var mailSubject = "Winrest - 1 Invoice - NEWREST ESPAÑA S.L. - " + site + " - " + DateUtils.Now.ToString("yyyyMM") + " - " + number;
            Assert.IsTrue(mailPage.CheckIfSpecifiedOutlookMailExist(mailSubject), "Mail Invoice n° " + number + " (quote) non trouvé");

            FileInfo pieceJointe = mailPage.DownloadOutlookAttachmentZIP(downloadsPath, "pieceJointe.jpg", true, mailSubject);
            Assert.IsTrue(pieceJointe.Exists, "fichier attaché pieceJointe.jpg non trouvé");
            pieceJointe.Delete();
            FileInfo invoiceJointe = mailPage.DownloadOutlookAttachmentZIP(downloadsPath, "Invoice_" + number + ".pdf", false, mailSubject);
            Assert.IsTrue(invoiceJointe.Exists, "fichier attaché Invoice_" + number + ".pdf non trouvé");
            invoiceJointe.Delete();

            mailPage.DeleteCurrentOutlookMail();
            mailPage.Close();

            WebDriver.Navigate().Refresh();
            invoicesPage = invoiceDetails.BackToList();

            // Recherche de l'invoice validée pour voir si elle a été envoyée par mail
            invoicesPage.ResetFilters();
            invoicesPage.Filter(InvoicesPage.FilterType.ByNumber, number);
            invoicesPage.Filter(InvoicesPage.FilterType.Site, site);

            bool isSentByMail = invoicesPage.IsSentByMail();
            // Assert
            Assert.IsTrue(isSentByMail, "L'invoice créée n'a pas été envoyée par mail.");
        }

        [Ignore]
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_SageManuel_Detail_EnableExportSage()
        {
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string serviceName = TestContext.Properties["InvoiceService"].ToString();
            string journalInvoice = TestContext.Properties["Journal_Invoice"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Paramétrage Export Sage manuel
            homePage.SetSageAutoEnabled(site, false);

            // Manipulation pour permettre export SAGE 
            var accountingParametersPage = homePage.GoToParameters_AccountingPage();
            accountingParametersPage.GoToTab_Journal();
            accountingParametersPage.EditJournal(site, journalInvoice);

            //Act
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();

            invoicesPage.ClearDownloads();

            // Create
            var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
            var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customer, true);


            // Ajout d'un service et validation
            invoiceDetails.AddService(serviceName, "10");
            invoiceDetails.Validate();

            Assert.IsFalse(invoiceDetails.IsEnableExportForSageEnabled(), "Il est possible de cliquer sur la fonctionnalité 'Enable export for SAGE' "
                + "pour un invoice non envoyé au SAGE.");

            // Export vers SAGE
            invoiceDetails.ExportSage();

            Assert.IsFalse(invoiceDetails.IsExportSageEnabled(), "Il est possible de cliquer sur la fonctionnalité 'Export for SAGE' "
                + "après avoir réalisé un export SAGE.");

            Assert.IsTrue(invoiceDetails.IsEnableExportForSageEnabled(), "Il est impossible de cliquer sur la fonctionnalité 'Enable export for SAGE' "
                + "pour un invoice envoyé au SAGE.");

            invoiceDetails.EnableExportForSage();

            Assert.IsTrue(invoiceDetails.IsExportSageEnabled(), "Il est impossible de cliquer sur la fonctionnalité 'Export for SAGE' "
                + "après avoir cliqué un export SAGE.");

            Assert.IsFalse(invoiceDetails.IsEnableExportForSageEnabled(), "Il est possible de cliquer sur la fonctionnalité 'Enable export for SAGE' "
                + "pour un invoice envoyé au SAGE.");

            invoiceDetails.ExportSage();

            Assert.IsTrue(invoiceDetails.IsEnableExportForSageEnabled(), "Il est impossible de cliquer sur la fonctionnalité 'Enable export for SAGE' "
                + "pour un invoice envoyé au SAGE.");
        }

        [Ignore]
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_SageManuel_Index_EnableSAGEExport()
        {
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string serviceName = TestContext.Properties["InvoiceService"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Paramétrage Export Sage manuel
            homePage.SetSageAutoEnabled(site, false);

            //Act
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();

            invoicesPage.ClearDownloads();

            var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
            var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customer, false);

            // Ajout d'un service et validation
            invoiceDetails.AddService(serviceName, "10");
            invoiceDetails.Validate();

            var invoiceNumber = invoiceDetails.GetInvoiceNumber();

            invoicesPage = invoiceDetails.BackToList();
            invoicesPage.ResetFilters();
            invoicesPage.Filter(InvoicesPage.FilterType.ByNumber, invoiceNumber);

            invoicesPage.ExportResultsToSage();
            invoiceDetails = invoicesPage.SelectFirstInvoice();

            Assert.IsTrue(invoiceDetails.IsEnableExportForSageEnabled(), "Il n'est pas possible de cliquer sur la fonctionnalité 'Enable export for SAGE' "
                + "après avoir exporté la supplier invoice vers SAGE depuis la page Index.");

            Assert.IsFalse(invoiceDetails.IsExportSageEnabled(), "Il est possible de cliquer sur la fonctionnalité 'Export for SAGE' "
                + "après avoir exporté la supplier invoice vers SAGE depuis la page Index.");

            invoicesPage = invoiceDetails.BackToList();
            invoicesPage.EnableExportForSage();
            invoiceDetails = invoicesPage.SelectFirstInvoice();

            Assert.IsFalse(invoiceDetails.IsEnableExportForSageEnabled(), "Il est possible de cliquer sur la fonctionnalité 'Enable export for SAGE' "
                + "après avoir activé la fonctionnalité depuis la page Index.");

            Assert.IsTrue(invoiceDetails.IsExportSageEnabled(), "Il est impossible de cliquer sur la fonctionnalité 'Export for SAGE' "
                + "après avoir cliqué sur 'Enable export for SAGE' depuis la page Index.");

            invoiceDetails.ExportSage();

            Assert.IsTrue(invoiceDetails.IsEnableExportForSageEnabled(), "Il n'est pas possible de cliquer sur la fonctionnalité 'Enable export for SAGE' "
                + "après avoir exporté la supplier invoice vers SAGE depuis la page Index.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_CegidAuto_SendAutoToCegid()
        {
            //Prepare
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string serviceName = TestContext.Properties["InvoiceService"].ToString();
            string currency = TestContext.Properties["Currency"].ToString();
            string journalInvoice = TestContext.Properties["Journal_Invoice"].ToString();

            var mailPage = new MailPage(WebDriver, TestContext);

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            try
            {
                // Parameter - Accounting --> Journal KO pour le test
                VerifyAccountingJournal(homePage, site, "");

                // 1. Config Export Sage auto pour créer la facture
                homePage.SetSageAutoEnabled(site, true, "Invoice");

                //Act
                var invoicesPage = homePage.GoToAccounting_InvoicesPage();
                invoicesPage.ResetFilters();

                // Create
                var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
                var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customer, false);

                // Ajout d'un service
                invoiceDetails.AddService(serviceName, "10");
                invoiceDetails.Validate();

                var invoiceNumber = invoiceDetails.GetInvoiceNumber();

                // Récupération du montant TTC de la facture
                var invoiceFooter = invoiceDetails.ClickOnInvoiceFooter();
                double montantInvoice = invoiceFooter.GetTotalTTC(currency, decimalSeparatorValue);

                var invoiceAccounting = invoiceDetails.ClickOnAccounting();
                Assert.AreNotEqual("", invoiceAccounting.GetErrorMessage(), "Le code journal est manquant mais aucun message d'erreur n'est présent dans l'onglet Accounting.");

                invoicesPage = invoiceAccounting.BackToList();
                invoicesPage.ResetFilters();
                invoicesPage.Filter(InvoicesPage.FilterType.ByNumber, invoiceNumber);
                invoicesPage.ShowInvoiceShow();
                invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowSentToSageAndInError, true);

                Assert.AreEqual(1, invoicesPage.CheckTotalNumber(), "L'invoice n'est pas au statut 'SentToSageInerror' malgré l'échec de l'envoi vers SAGE.");

                // On vérifie qu'un mail a été envoyé
                CheckErrorMailSentToUser(homePage, mailPage);

                // 2. On remet en place le code journal pour les Invoices
                VerifyAccountingJournal(homePage, site, journalInvoice);

                invoicesPage = homePage.GoToAccounting_InvoicesPage();
                invoicesPage.ResetFilters();
                invoicesPage.Filter(InvoicesPage.FilterType.ByNumber, invoiceNumber);

                invoiceDetails = invoicesPage.SelectFirstInvoice();

                // Calcul du montant de la facture transmise à TL
                invoiceAccounting = invoiceDetails.ClickOnAccounting();
                double montantFacture = invoiceAccounting.GetInvoiceGrossAmount("G", decimalSeparatorValue);
                double montantDetaille = invoiceAccounting.GetInvoiceDetailAmount("G", decimalSeparatorValue);

                invoicesPage = invoiceAccounting.BackToList();
                invoicesPage.ResetFilters();
                invoicesPage.Filter(InvoicesPage.FilterType.ByNumber, invoiceNumber);

                // Renvoi de la facture vers TL
                invoicesPage.SendAutoToSage();

                WebDriver.Navigate().Refresh();
                invoicesPage.ResetFilters();
                invoicesPage.Filter(InvoicesPage.FilterType.ByNumber, invoiceNumber);
                invoicesPage.ShowInvoiceShow();
                invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowOnTLWaitingForSagePush, true);

                Assert.AreEqual(1, invoicesPage.CheckTotalNumber(), "L'export SAGE Auto de l'invoice n'a pas été envoyé vers le SAGE.");
                Assert.AreEqual(montantFacture.ToString(), montantDetaille.ToString(), "Les montants AmountDebit et AmountCredit de la facture SAGE ne sont pas les mêmes.");
                Assert.AreEqual(montantInvoice.ToString(), montantFacture.ToString(), "Le montant issu du fichier SAGE n'est pas égal au montant de l'invoice défini dans l'application.");
            }
            finally
            {
                homePage.SetSageAutoEnabled(site, false);
                VerifyAccountingJournal(homePage, site, journalInvoice);
            }
        }

        [Ignore]// aucun pays n'utilise SAGE AUTO
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_SageAuto_Index_ExportTXT()
        {
            //Prepare
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string serviceName = TestContext.Properties["InvoiceService"].ToString();
            string currency = TestContext.Properties["Currency"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string journalInvoice = TestContext.Properties["Journal_Invoice"].ToString();

            // Log in
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            try
            {
                // 1. Config Export Sage auto
                homePage.SetSageAutoEnabled(site, true, "Invoice");

                // journal Piece Id 132881 : Site [BCN] has no accounting journal setting set.
                VerifyAccountingJournal(homePage, site, journalInvoice);

                //Act
                var invoicesPage = homePage.GoToAccounting_InvoicesPage();

                invoicesPage.ClearDownloads();

                var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
                var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customer, false);

                // Ajout d'un service
                invoiceDetails.AddService(serviceName, "10");
                invoiceDetails.Validate();

                var invoiceNumber = invoiceDetails.GetInvoiceNumber();

                // Récupération du montant TTC de la facture
                var invoiceFooter = invoiceDetails.ClickOnInvoiceFooter();
                double montantInvoice = invoiceFooter.GetTotalTTC(currency, decimalSeparatorValue);

                // Calcul du montant de la facture transmise à TL
                var invoiceAccounting = invoiceFooter.ClickOnAccounting();

                double montantFacture = invoiceAccounting.GetInvoiceGrossAmount("G", decimalSeparatorValue);
                double montantDetaille = invoiceAccounting.GetInvoiceDetailAmount("G", decimalSeparatorValue);

                // Retour à la page d'accueil pour vérifier que la facture est partie vers TL
                invoicesPage = invoiceAccounting.BackToList();
                invoicesPage.ResetFilters();
                invoicesPage.Filter(InvoicesPage.FilterType.ByNumber, invoiceNumber);

                // Renvoi de la facture vers TL
                invoicesPage.ExportResultsToSage();

                // On récupère les fichiers du répertoire de téléchargement
                DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
                FileInfo[] taskFiles = taskDirectory.GetFiles();

                var correctDownloadedFile = invoicesPage.GetExportSageFile(taskFiles);
                Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

                // Récupération du nom du fichier et construction de l'URL du fichier txt à exploiter   
                var fileName = correctDownloadedFile.Name;
                var filePath = Path.Combine(downloadsPath, fileName);

                // On n'exploite que les lignes avec contenu "général" --> "G"
                double[] contenuFichier = ExploitTextFiles.VerifySAGEFileContent(filePath, "G", decimalSeparatorValue);

                Assert.AreEqual(contenuFichier[0].ToString(), contenuFichier[1].ToString(), "Le contenu du fichier SAGE généré est incorrect.");
                Assert.AreEqual(montantInvoice.ToString(), contenuFichier[0].ToString(), "Le montant issu du fichier SAGE n'est pas égal au montant de l'invoice défini dans l'application.");
                Assert.AreEqual(montantFacture.ToString(), montantDetaille.ToString(), "Les montants AmountDebit et AmountCredit de la facture SAGE ne sont pas les mêmes dans l'onglet Accounting.");
                Assert.AreEqual(montantInvoice.ToString(), montantFacture.ToString(), "Le montant issu du fichier SAGE n'est pas égal au montant de l'invoice défini dans l'application.");
            }
            finally
            {
                homePage.Navigate();
                homePage.SetSageAutoEnabled(site, false);
            }
        }

        [Ignore]
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_Details_GenerateSAGETxt()
        {
            //Prepare
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string serviceName = TestContext.Properties["InvoiceService"].ToString();
            string currency = TestContext.Properties["Currency"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            // Log in
            var homePage = LogInAsAdmin();


            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();

            // 1. Config Export Sage manuel
            homePage.SetSageAutoEnabled(site, false);

            //Act
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();

            invoicesPage.ClearDownloads();

            var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
            var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customer, false);

            // Ajout d'un service
            invoiceDetails.AddService(serviceName, "10");
            invoiceDetails.Validate();

            // Récupération du montant TTC de la facture
            var invoiceFooter = invoiceDetails.ClickOnInvoiceFooter();
            double montantInvoice = invoiceFooter.GetTotalTTC(currency, decimalSeparatorValue);

            // Calcul du montant de la facture transmise à TL
            var invoiceAccounting = invoiceFooter.ClickOnAccounting();

            double montantFacture = invoiceAccounting.GetInvoiceGrossAmount("G", decimalSeparatorValue);
            double montantDetaille = invoiceAccounting.GetInvoiceDetailAmount("G", decimalSeparatorValue);

            invoiceDetails = invoiceAccounting.ClickOnItems();
            invoiceDetails.ExportSage();

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            var correctDownloadedFile = invoicesPage.GetExportSageFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

            // Récupération du nom du fichier et construction de l'URL du fichier txt à exploiter   
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // On n'exploite que les lignes avec contenu "général" --> "G"
            double[] contenuFichier = ExploitTextFiles.VerifySAGEFileContent(filePath, "G", decimalSeparatorValue);

            Assert.AreEqual(contenuFichier[0].ToString(), contenuFichier[1].ToString(), "Le contenu du fichier SAGE généré est incorrect.");
            Assert.AreEqual(montantInvoice.ToString(), contenuFichier[0].ToString(), "Le montant issu du fichier SAGE n'est pas égal au montant de l'invoice défini dans l'application.");
            Assert.AreEqual(montantFacture.ToString(), montantDetaille.ToString(), "Les montants AmountDebit et AmountCredit de la facture SAGE ne sont pas les mêmes dans l'onglet Accounting.");
            Assert.AreEqual(montantInvoice.ToString(), montantFacture.ToString(), "Le montant issu du fichier SAGE n'est pas égal au montant de l'invoice défini dans l'application.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_InvoiceWithForeignDevise()
        {
            //Prepare
            Random rnd = new Random();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string serviceName = "ServiceCurrency" + "-" + new Random().Next(10, 99).ToString();
            string currency = "$ Dolar Americano";
            string customerName = "customerCurrency" + "-" + rnd.Next().ToString();
            string customerCode = rnd.Next().ToString();
            string customerType = TestContext.Properties["CustomerType"].ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string serviceCategory = TestContext.Properties["CategorieBOB"].ToString();
            string deleteFrom = DateUtils.Now.AddMonths(-1).ToString("dd/MM/yyyy");
            string deleteTo = DateUtils.Now.AddMonths(3).ToString("dd/MM/yyyy");
            long quantity = 10;
            double unitPrice = 12;

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            var settingsPurchasingPage = homePage.GoToParameters_PurchasingPage();
            settingsPurchasingPage.EditCurrency(currency);
            //Create Customer 
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.Filter(CustomerPage.FilterType.Search, customerName);
            if (customersPage.CheckTotalNumber_display() == 0)
            {
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFieldsWithCurrency_CreateCustomerModalPage(customerName, customerCode, customerType, currency);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                var currencyDolar = customerGeneralInformationsPage.GetDevise();

                Assert.AreEqual(currency, currencyDolar, "Currency n'est pas en $.");
            }
            //Create service with same Customer 
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            if (servicePage.CheckTotalNumber() == 0)
            {
                var ServiceCreateModalPage = servicePage.ServiceCreatePage();
                ServiceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction, serviceCategory);
                var serviceGeneralInformationsPage = ServiceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customerName, DateUtils.Now, DateUtils.Now.AddDays(10));
            }
            try
            {
                //Create Manual Invoice 
                var invoicesPage = homePage.GoToAccounting_InvoicesPage();
                // Create
                var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
                var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customerName, true);
                // Ajout d'un service et validation
                invoiceDetails.AddService(serviceName);
                invoiceDetails.SelectFirstLine();
                invoiceDetails.Fill(null, quantity, unitPrice, null, null);

                invoiceDetails.Validate();
                invoiceDetails.WaitPageLoading();
                var invoiceFooter = invoiceDetails.ClickOnInvoiceFooter();

                string CustomerCurrency = invoiceFooter.GetCustomerCurrency();
                string LocalCurrency = invoiceFooter.GetLocalCurrency();
                //Local Total Incl Taxes = Total Incl Taxes * Exchang rate

                bool IsverifyTotalInvoice = invoiceFooter.VerifyTotalInvoice();

                Assert.AreNotEqual(CustomerCurrency, LocalCurrency, "La devise du client n'a pas été prise en compte.");
                Assert.IsTrue(IsverifyTotalInvoice, "Le Local Total Incl Taxes = Total Incl Taxes * Exchang rate.");
            }
            finally
            {
                //Delete service
                servicePage = homePage.GoToCustomers_ServicePage();
                var serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.From, deleteFrom);
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.To, deleteTo);
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_PdfTVA()
        {
            //Prepare
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string serviceName = TestContext.Properties["InvoiceService"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            //Print Invoice_-_443562_-_20220314110329.pdf
            string DocFileNamePdfBegin = "Print Invoice_-_";
            //All_files_20220314_110409.zip
            string DocFileNameZipBegin = "All_files_";
            string qty = "10";
            DateTime date = DateUtils.Now;
            bool isVATForced = true;
            string currency = "€";

            //Arrange
            var homePage = LogInAsAdmin();

            var decimalSeparator = homePage.GetDecimalSeparatorValue();

            //Act
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();

            invoicesPage.ClearDownloads();
            var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
            var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, date, customer, isVATForced);

            // Ajout d'un service
            invoiceDetails.AddService(serviceName, qty);
            invoiceDetails.WaitPageLoading();
            invoiceDetails.Validate();
            invoiceDetails.WaitPageLoading();
            InvoiceFooterPage InvoiceFooter = invoiceDetails.ClickOnInvoiceFooter();
            string[] VATrate = InvoiceFooter.GetColumnVATrates();
            double tarif = InvoiceFooter.GetTotalTTC(currency, decimalSeparator);
            NumberFormatInfo nfi = new NumberFormatInfo();
            nfi.NumberDecimalSeparator = decimalSeparator;
            string tarifString = tarif.ToString(nfi);

            //1.Cliquer sur print
            //2. Remplir les paramètres d'impression
            //3.Cliquer sur print
            invoicesPage.ClearDownloads();
            PrintReportPage reportPage = PrintItemGenerique(invoiceDetails);
            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            // cliquer sur All
            string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin, false);
            FileInfo fi = new FileInfo(trouve);
            fi.Refresh();
            Assert.IsTrue(fi.Exists, trouve + " non généré");

            //4. Vérifier la TVA sur le PDF
            //Un rapport PDF est généré selon les paramètres choisis et vérifier la TVA
            PdfDocument document = PdfDocument.Open(fi.FullName);
            Page page1 = document.GetPage(1);
            IEnumerable<Word> words = page1.GetWords();
            //Assert
            var verifyIVAIGIC = words.Count(w => w.Text == "IVA/IGIC");
            Assert.AreEqual(7, verifyIVAIGIC, "IVA/IGIC non présent dans le Pdf");
            //Assert
            var verifyVATrate0 = words.Any(w => w.Text == VATrate[0]);
            Assert.IsTrue(verifyVATrate0, VATrate[0] + " non présent dans le Pdf");
            //Assert
            var verifyVATrate1 = words.Any(w => w.Text == VATrate[1]);
            Assert.IsTrue(verifyVATrate1, VATrate[1] + " non présent dans le Pdf");
            //Assert
            var verifyVATrate2 = words.Any(w => w.Text == VATrate[2]);
            Assert.IsTrue(verifyVATrate2, VATrate[2] + " non présent dans le Pdf");
            //Assert
            var verifyTarif = words.Any(w => w.Text == tarifString);
            Assert.IsTrue(verifyTarif, tarifString + " non présent dans le Pdf");
        }

        /**
         * Imprimer les données de l'invoice
         */
        [DeploymentItem("Resources\\pizza.png")]
        [DeploymentItem("Resources\\pizza_petite-old.png")]
        [DeploymentItem("Resources\\pizza_petite.png")]
        [DeploymentItem("chromedriver.exe")]
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_DetailsPdfLogo()
        {
            //Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNamePdfBegin = "Print Invoice_-_";
            string DocFileNameZipBegin = "All_files_";
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string serviceName = "ServiceInvoice" + "-" + new Random().Next(10, 99).ToString();
            string qty = "10";
            string serviceCategory = TestContext.Properties["CategorieBOB"].ToString();
            string deleteFrom = DateUtils.Now.AddMonths(-1).ToString("dd/MM/yyyy");
            string deleteTo = DateUtils.Now.AddMonths(3).ToString("dd/MM/yyyy");

            //Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            ParametersSites SitesParams = homePage.GoToParameters_Sites();
            SitesParams.Filter(ParametersSites.FilterType.SearchSite, site);
            SitesParams.ClickOnFirstSite();
            //Configurer un logo sur un site dans les globals settings
            //Avoir des données disponibles.
            // serialize + sha1sum (filesize)
            Console.WriteLine(TestContext.TestDeploymentDir + "\\pizza.png");
            FileInfo fiUpload = new FileInfo(TestContext.TestDeploymentDir + "\\pizza.png");
            SitesParams.UploadLogo(fiUpload);

            try
            {
                // Act ajout  service
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

                InvoicesPage invoicesPage = homePage.GoToAccounting_InvoicesPage();
                ManualInvoiceCreateModalPage manualPage = invoicesPage.ManualInvoiceCreatePage();
                InvoiceDetailsPage invoiceDetails = manualPage.FillField_CreatNewManualInvoice(site, DateUtils.Now.Date, customer, false);

                invoiceDetails.AddService(serviceName, qty);
                invoiceDetails.Validate();
                invoiceDetails.ConfirmValidation();
                invoicesPage.ClearDownloads();
                DeleteAllFileDownload();

                //2. cliquer sur print
                PrintReportPage reportPage = PrintItemGenerique(invoiceDetails);
                reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
                string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin, false);
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

                // comparaison fichiers
                var md5 = MD5.Create();
                var fsTest = File.Open(fiTest.FullName, FileMode.Open);
                var fsPdf = File.Open(fiPdf.FullName, FileMode.Open);
                var file1 = System.Convert.ToBase64String(md5.ComputeHash(fsTest));
                var file2 = System.Convert.ToBase64String(md5.ComputeHash(fsPdf));
                Assert.AreEqual(file1, file2, "fichiers différents");
            }
            finally
            {
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

        /**
         * Imprimer les index des invoices selon les critères de filtre en PDF
         */
        [DeploymentItem("Resources\\pizza.png")]
        [DeploymentItem("Resources\\pizza_petite.png")]
        [DeploymentItem("Resources\\pizza_petite-old.png")]
        [DeploymentItem("chromedriver.exe")]
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_IndexPdfLogo()
        {
            // Resources\\pizza.png
            // F4
            // Copy to Ouput Directory : "Copmy always"

            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            //Print Invoice_-_443562_-_20220314110329.pdf
            string DocFileNamePdfBegin = "Print Invoice_-_";
            //All_files_20220314_110409.zip
            string DocFileNameZipBegin = "All_files_";
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            ParametersSites SitesParams = homePage.GoToParameters_Sites();
            SitesParams.Filter(ParametersSites.FilterType.SearchSite, site);
            SitesParams.ClickOnFirstSite();

            //Configurer un logo sur un site dans les globals settings
            //Avoir des données disponibles.
            // serialize + sha1sum (filesize)
            Console.WriteLine(TestContext.TestDeploymentDir + "\\pizza.png");
            FileInfo fiUpload = new FileInfo(TestContext.TestDeploymentDir + "\\pizza.png");
            SitesParams.UploadLogo(fiUpload);

            homePage.Navigate();

            //Act
            homePage.ClearDownloads();
            InvoicesPage invoicesPage = homePage.GoToAccounting_InvoicesPage();
            invoicesPage.ResetFilters();
            invoicesPage.Filter(InvoicesPage.FilterType.Site, site);
            invoicesPage.Filter(InvoicesPage.FilterType.Customer, customer);
            invoicesPage.Filter(InvoicesPage.FilterType.DateFrom, DateUtils.Now.Date.AddDays(-1));
            string invoiceNumber = invoicesPage.GetFirstInvoiceNumber().Substring(2);
            invoicesPage.Filter(InvoicesPage.FilterType.ByNumber, invoiceNumber);

            // [...]
            PrintReportPage reportPage = invoicesPage.PrintInvoiceResults();
            var isReportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();

            //Assert
            Assert.IsTrue(isReportGenerated, "L'application n'a pas pu générer le fichier attendu.");

            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);

            string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin, false);
            FileInfo fi = new FileInfo(trouve);
            fi.Refresh();
            var f1exist = fi.Exists;
            Assert.IsTrue(f1exist, trouve + " non généré");

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
            var existfilepdf = fiPdf.Exists;
            Assert.IsTrue(existfilepdf, "fichier non trouvé");

            FileInfo fiTest;
            fiTest = new FileInfo(TestContext.TestDeploymentDir + "\\pizza_petite.png");
            var existfiletest = fiTest.Exists;
            Assert.IsTrue(existfiletest, "fichier non trouvé (cas 2)");

            // comparaison fichiers
            var md5 = MD5.Create();
            var fsTest = File.Open(fiTest.FullName, FileMode.Open);
            var fsPdf = File.Open(fiPdf.FullName, FileMode.Open);
            var comparefsTest = System.Convert.ToBase64String(md5.ComputeHash(fsTest));
            var comparefsPdf = System.Convert.ToBase64String(md5.ComputeHash(fsPdf));

            Assert.AreEqual(comparefsTest, comparefsPdf, "fichiers différents");
        }

        [DeploymentItem("Resources\\pizza.png")]
        [DeploymentItem("Resources\\pizza_petite.png")]
        [DeploymentItem("Resources\\pizza_petite-old.png")]
        [DeploymentItem("chromedriver.exe")]
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_PdfDataEtLogo()
        {
            // Resources\\pizza.png
            // F4
            // Copy to Ouput Directory : "Copy always"
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            //Print Invoice_-_437933_-_20220406125011.pdf
            string DocFileNamePdfBegin = "Print Invoice_-_";
            //All_files_20220314_110409.zip
            string DocFileNameZipBegin = "All_files_";
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string service = TestContext.Properties["InvoiceService"].ToString();
            string quantity = "10";
            string currency = TestContext.Properties["Currency"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();
            var decimalSeparatorValue = homePage.GetDecimalSeparatorValue();
            homePage.ClearDownloads();
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

            //Act
            homePage.ClearDownloads();
            //Etre sur l'index des invoices
            InvoicesPage invoicesPage = homePage.GoToAccounting_InvoicesPage();

            //1. Etre sur une credtit note
            ManualCreditNoteCreatePage creditNote = invoicesPage.ManualCreditNoteCreatePage();
            creditNote.FillField_CreateCreditNote(site, customer);
            InvoiceDetailsPage detailsPage = creditNote.Create();

            // enrichir une ligne
            //4) Ajouter un item et quantité Vérifier les bonnes infos dans General Info
            detailsPage.AddService(service, quantity);

            var invoiceFooterPage = detailsPage.ClickOnInvoiceFooter();
            string totalTTC = invoiceFooterPage.GetTotalTTC(currency);
            string totalGrossAmount = invoiceFooterPage.GetTotalGrossAmount(currency);
            //désormais un Invoice n'a pas de total TTC à zéro, (confirmation lors du print)
            if (!totalGrossAmount.Contains(","))
            {
                totalGrossAmount = totalGrossAmount + ",00";
            }

            //2.Survoler les …
            creditNote.dropBtnClick();

            //3.Cliquer sur Print
            creditNote.ClickPrint();

            //Fichier PDF est généré dans la print list 
            PrintReportPage reportPage = new PrintReportPage(WebDriver, TestContext);
            reportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            WebDriver.Navigate().Refresh();

            invoicesPage.ShowBtnDownload();
            string trouve = reportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin, false);
            FileInfo fi = new FileInfo(trouve);
            fi.Refresh();
            Assert.IsTrue(fi.Exists, trouve + " non généré");

            //& vérifier le logo et données
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

            // comparaison fichiers
            var md5 = MD5.Create();
            var fsTest = File.Open(fiTest.FullName, FileMode.Open);
            var fsPdf = File.Open(fiPdf.FullName, FileMode.Open);

            var verifFile1 = System.Convert.ToBase64String(md5.ComputeHash(fsTest));
            var verifFile2 = System.Convert.ToBase64String(md5.ComputeHash(fsPdf));
            Assert.AreEqual(verifFile1, verifFile2, "fichiers différents");

            // données
            var verifSite = page1.GetWords().Count(x => x.Text == site);
            Assert.AreEqual(2, verifSite, "Le site n'est pas présent dans le pdf.");

            var verifService = page1.GetWords().Any(substring => service.Contains(substring.ToString()));
            Assert.IsTrue(verifService, "Le service ajouté n'est pas présent dans le pdf.");

            var verifQuantity = page1.GetWords().Count(x => x.Text == quantity);
            Assert.AreEqual(1, verifQuantity, "La quantité de l'item ajouté n'est pas présente dans le pdf.");

            //désormais un Invoice n'a pas de total TTC à zéro, (confirmation lors du print)
            var verifTotalGrossAmount = page1.GetWords().Count(x => x.Text.Contains(totalGrossAmount));
            Assert.AreEqual(4, verifTotalGrossAmount, "Le total sans TVA n'est pas affiché");

            var verifTotalTTC = page1.GetWords().Count(x => x.Text.Contains(totalTTC));
            Assert.AreEqual(2, verifTotalTTC, "Le total n'est pas affiché");

            var verifDate = page1.GetWords().Count(x => x.Text == DateUtils.Now.ToString("dd/MM/yyyy"));
            Assert.AreEqual(2, verifDate, "La date de l'invoice n'est pas affichée");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_Details_UpdateService()
        {
            //prepare
            //invoice
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string qtyValue = "100";
            DateTime date = DateUtils.Now.Date;
            // service
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            // price
            DateTime fromDate = DateTime.Now;
            DateTime toDate = DateUtils.Now.AddMonths(12);
            string price = "10";
            // modify service
            long quantity = 3;
            double editeddiscount = -5.5;
            double editedTotal = 30;
            string serviceCategory = "HOT MEAL YC UAE";
            string Id = "";
            double discoutCheckFormated;
            double unitPrice;
            double discountNull = 0;

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //act
            var servicePage = homePage.GoToCustomers_ServicePage();
            try
            {
                // create a service
                var serviceCreateModalPage = servicePage.ServiceCreate();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                // add price
                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, fromDate, toDate, price);
                pricePage.UnfoldAll();
                servicePage = serviceGeneralInformationsPage.BackToList();
                //creé invoice non validé
                InvoicesPage invoices = homePage.GoToAccounting_InvoicesPage();
                ManualInvoiceCreateModalPage modal = invoices.ManualInvoiceCreatePage();
                InvoiceDetailsPage detailsInvoice = modal.FillField_CreatNewManualInvoice(site, date, customer, false);
                // ajouté un service
                detailsInvoice.AddService(serviceName, qtyValue);
                // Vérifier le calcul du total avec changement discount -, et quantity
                detailsInvoice.Fill(null, quantity, null, editeddiscount, null);

                Double.TryParse(price, out unitPrice);
                double totalCalculated;
                totalCalculated = quantity * unitPrice * (1 + editeddiscount / 100);
                double total = detailsInvoice.GetTotal();
                //Assert
                Assert.AreEqual(totalCalculated, total, 0.01, "le changement de discount - ne s'applique pas correctement");
                detailsInvoice.WaitPageLoading();
                // Vérifier le calcul du total avec changement discount +, et quantity
                quantity = 10;
                editeddiscount = +5.5;
                detailsInvoice.Fill(null, quantity, null, editeddiscount, null);
                totalCalculated = quantity * unitPrice * (1 + editeddiscount / 100);
                total = detailsInvoice.GetTotal();
                //Assert
                Assert.AreEqual(totalCalculated, total, 0.01, "le changement de discount + ne s'applique pas correctement");
                // verifier calcul de unit price aprés modification de total
                detailsInvoice.WaitPageLoading();
                detailsInvoice.Fill(null, null, null, null, editedTotal);
                unitPrice = editedTotal / (quantity * (1 + editeddiscount / 100));
                var unitePriceCheck = detailsInvoice.GetUnitPrice();
                //Assert
                Assert.AreEqual(unitPrice, unitePriceCheck, 0.01, "le changement de total ne s'applique pas correctement");
                // verifier le changement apres la modification de service category
                detailsInvoice.FillServiceCategorie(serviceCategory);
                detailsInvoice.WaitPageLoading();
                string discountCheck = detailsInvoice.GetFirstServiceDiscount();
                Double.TryParse(discountCheck, out discoutCheckFormated);
                //Assert
                Assert.AreEqual(discoutCheckFormated, discountNull, "le changment sur service category ne s'apllique pas");
                Id = detailsInvoice.GetInvoiceNumber();
            }
            finally
            {
                //Delete Invoice 
                var invoicesPage = homePage.GoToAccounting_InvoicesPage();
                invoicesPage.ResetFilters();
                invoicesPage.ShowInvoiceShow();
                invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowNotValidatedOnly, true);
                invoicesPage.DeleteInvoiceById(Id);
                //Delete service
                servicePage = homePage.GoToCustomers_ServicePage();
                var serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_Details_UpdatePrices()
        {
            // prepare
            string serviceName = serviceNameToday + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string qtyValue = "100";
            DateTime date = DateUtils.Now.Date;
            DateTime fromDate = DateTime.Now;
            DateTime toDate = DateUtils.Now.AddMonths(12);
            string price = "10";
            double editedPrice = 20;
            double discount = 5;
            string Id = "";
            double quantity;

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var servicePage = homePage.GoToCustomers_ServicePage();

            try
            {
                var serviceCreateModalPage = servicePage.ServiceCreate();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                // add price
                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, fromDate, toDate, price);
                pricePage.UnfoldAll();
                servicePage = serviceGeneralInformationsPage.BackToList();
                //creé invoice non validé
                InvoicesPage invoices = homePage.GoToAccounting_InvoicesPage();
                ManualInvoiceCreateModalPage modal = invoices.ManualInvoiceCreatePage();
                InvoiceDetailsPage detailsInvoice = modal.FillField_CreatNewManualInvoice(site, date, customer, false);
                // ajouté un service
                detailsInvoice.AddService(serviceName, qtyValue);
                // modifier le price unit
                detailsInvoice.Fill(null, null, editedPrice, discount, null);
                // calcul
                double.TryParse(qtyValue, out quantity);
                double totalCalculated = quantity * editedPrice * (1 + discount / 100);
                double totalCheck = detailsInvoice.GetTotal();
                //Assert
                Assert.AreEqual(totalCalculated, totalCheck, "la modification sur unit price ne s'applique pas");
                // update prices
                detailsInvoice.UpdatePrices();
                detailsInvoice.WaitPageLoading();
                double unitPrice = detailsInvoice.GetUnitPrice();
                //Assert
                Assert.AreNotEqual(editedPrice, unitPrice, "le price unit n'a pas changé");
                string unitPriceFormated = unitPrice.ToString();
                //Assert
                Assert.AreEqual(unitPriceFormated, price, "le price unit n'a pas changé");
                Id = detailsInvoice.GetInvoiceNumber();

            }
            finally
            {
                //Delete Invoice 
                var invoicesPage = homePage.GoToAccounting_InvoicesPage();
                invoicesPage.ResetFilters();
                invoicesPage.ShowInvoiceShow();
                invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowNotValidatedOnly, true);
                invoicesPage.DeleteInvoiceById(Id);
                //Delete service
                servicePage = homePage.GoToCustomers_ServicePage();
                var serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();
            }
        }


        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_CreateCreditNoteManual()
        {
            // Prepare 
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string service = TestContext.Properties["InvoiceService"].ToString();
            var homePage = LogInAsAdmin();
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            invoicesPage.ResetFilters();

            // Act: Create manual credit note
            var creditNotePage = invoicesPage.ManualCreditNoteCreatePage();
            creditNotePage.FillField_CreateCreditNote(site, customer);
            var detailsPage = creditNotePage.Create();

            // Add service details to credit note
            detailsPage.AddService(service, "10");

            // Retrieve and verify invoice information
            var generalInfo = detailsPage.ClickOnGeneralInformation();
            string createdInvoiceId = generalInfo.GetInvoiceId();
            invoicesPage = generalInfo.BackToList();

            // Validate that the created invoice appears in the invoices list
            string listedInvoiceId = invoicesPage.GetFirstInvoiceID().Substring(3);
            Assert.AreEqual(createdInvoiceId, listedInvoiceId, "The created invoice ID does not match the listed ID.");

            // Ensure the invoice is not validated
            Assert.IsFalse(invoicesPage.IsExistingFirstInvoiceNumber(), "Invoice is validated unexpectedly.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_Filter_Show_DeliveryInvoices()
        {
            //Arrange
            var homePage = LogInAsAdmin();

            InvoicesPage invoicesPage = homePage.GoToAccounting_InvoicesPage();
            invoicesPage.ResetFilters();

            invoicesPage.ShowInvoiceShow();
            invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowAll, true);
            var totalNbreAll = invoicesPage.CheckTotalNumber();

            invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowFlightInvoice, true);
            var totalNbreFlightInvoice = invoicesPage.CheckTotalNumber();

            invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowDeliveryInvoice, true);
            var totalNbreDeliveryInvoice = invoicesPage.CheckTotalNumber();

            invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowCustomerOrderInvoice, true);
            var totalNbreCustomerOrderInvoice = invoicesPage.CheckTotalNumber();

            invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowManualInvoice, true);
            var totalNbreManualInvoice = invoicesPage.CheckTotalNumber();

            // Assert
            Assert.AreEqual(totalNbreDeliveryInvoice, totalNbreAll - (totalNbreFlightInvoice + totalNbreCustomerOrderInvoice + totalNbreManualInvoice), String.Format(MessageErreur.FILTRE_ERRONE, "'Show delivery invoices'"));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_Create_Manual_Invoice()
        {
            //Prepare
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            InvoicesPage invoicesPage = homePage.GoToAccounting_InvoicesPage();
            ManualInvoiceCreateModalPage Invoice = invoicesPage.ManualInvoiceCreatePage();
            Invoice.FillField_CreatNewManualInvoice(site, DateUtils.Now.Date, customer, false);
            Assert.IsFalse(invoicesPage.IsExistingFirstInvoiceNumber(), "Invoice Validated");

        }

        // ___________________________________________ Utilitaire _________________________________________

        private string CreateFlightForInvoice(HomePage homePage, string site, string customer, string service)
        {
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            var flightNumber = new Random().Next();
            var state = "I";

            //CheckNewVersionFlight(homePage);//old version

            // Flight  part
            var flightPage = homePage.GoToFlights_FlightPage();

            // On filtre par rapport au site sur lequel on veut créer un vol
            flightPage.Filter(FlightPage.FilterType.Sites, site);

            // Création d'un vol
            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber.ToString(), customer, aircraft, site, "ACE");

            var editPage = flightPage.EditFirstFlight(flightNumber.ToString());
            editPage.AddGuestType();
            editPage.AddService(service);
            editPage.UnsetState("P");
            editPage.SetNewState(state);
            editPage.CloseViewDetails();

            WebDriver.Navigate().Refresh();

            // On vérifie que le state Invoice est sélectionné
            Assert.IsTrue(flightPage.GetFirstFlightStatus("I"), "Le state 'Invoiced' n'est pas sélectionné pour le flight créé.");

            return flightNumber.ToString();
        }

        private string GetServiceCategory(HomePage homePage, string serviceName)
        {
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);

            var pricePage = servicePage.ClickOnFirstService();
            var serviceGeneralInfo = pricePage.ClickOnGeneralInformationTab();

            return serviceGeneralInfo.GetCategory();
        }

        private bool SetApplicationSettingsForSageAuto(HomePage homePage)
        {
            string environnement = TestContext.Properties["Winrest_Environnement"].ToString().ToUpper();

            try
            {
                var applicationSettings = homePage.GoToApplicationSettings();
                var versionBDD = applicationSettings.GetApplicationDbVersion();

                // Country code
                var appSettingsModalPage = applicationSettings.GetWinrestTLCountryCodePage();
                appSettingsModalPage.SetWinrestTLCountryCode(environnement);
                applicationSettings = appSettingsModalPage.Save();

                // BDD
                appSettingsModalPage = applicationSettings.GetWinrestExportTLSageDbOverloadPage();
                appSettingsModalPage.SetWinrestExportTLSageDbOverload(versionBDD);
                applicationSettings = appSettingsModalPage.Save();

                // Override countryCode
                appSettingsModalPage = applicationSettings.GetWinrestExportTLSageCountryCodeOverloadPage();
                appSettingsModalPage.SetWinrestExportTLSageCountryCodeOverload(environnement);
                appSettingsModalPage.Save();
            }
            catch
            {
                return false;
            }

            return true;
        }

        private bool VerifySiteAnalyticalPlanSection(HomePage homePage, string site, bool isOK = true)
        {
            string analyticalPlan = isOK ? "1" : "";
            string analyticalSection = isOK ? "314" : "";

            try
            {
                var settingsSitesPage = homePage.GoToParameters_Sites();
                settingsSitesPage.Filter(ParametersSites.FilterType.SearchSite, site);
                settingsSitesPage.ClickOnFirstSite();

                settingsSitesPage.ClickToInformations();
                settingsSitesPage.SetAnalyticPlan(analyticalPlan);
                settingsSitesPage.SetAnalyticSection(analyticalSection);
            }
            catch
            {
                return false;
            }

            return true;
        }
        private bool VerifySiteSageCegidAuto(HomePage homePage, string site)
        {
            try
            {
                var settingsSitesPage = homePage.GoToParameters_Sites();
                settingsSitesPage.Filter(ParametersSites.FilterType.SearchSite, site);
                settingsSitesPage.ClickOnFirstSite();

                settingsSitesPage.ClickToInformations();

                // Ajouter la vérification et l'action pour la checkbox
                var checkbox = settingsSitesPage.WaitForElementExists(By.XPath("//*[@id=\"FullSite_EnableSageAutoForSite\"]"));
                bool isChecked = checkbox.Selected; // Vérifie si la case est cochée

                if (!isChecked)
                {
                    checkbox.Click(); // Coche la case si elle ne l'est pas
                }

            }
            catch
            {
                return false;
            }

            return true;
        }


        private bool VerifyInvoiceSageContact(HomePage homePage, string site)
        {
            // Prepare
            string mail = TestContext.Properties["Admin_UserName"].ToString();
            string userName = mail.Substring(0, mail.IndexOf("@"));

            // Act
            var settingsSitesPage = homePage.GoToParameters_Sites();
            settingsSitesPage.Filter(ParametersSites.FilterType.SearchSite, site);
            settingsSitesPage.ClickOnFirstSite();

            settingsSitesPage.ClickToContacts();
            if (String.IsNullOrEmpty(settingsSitesPage.GetInvoiceSageManager()))
            {
                settingsSitesPage.SetInvoiceSageManager(userName, mail);
            }

            return !String.IsNullOrEmpty(settingsSitesPage.GetInvoiceSageManager());
        }

        private bool VerifyPurchasingVAT(HomePage homePage, string taxName, string taxType, bool isOK = true)
        {
            // Prepare
            string purchaseCode = "AS05";
            string purchaseAccount = "47205001";
            string salesCode = isOK ? "AR10" : "";
            string salesAccount = isOK ? "47205001" : "";
            string taxValue = "21";
            string RNCode = isOK ? "0" : "";
            string RNAccount = isOK ? "0" : "";

            try
            {
                // Act
                var settingsPurchasingPage = homePage.GoToParameters_PurchasingPage();
                settingsPurchasingPage.GoToTab_VAT();
                if (!settingsPurchasingPage.IsTaxPresent(taxName, taxType))
                {
                    settingsPurchasingPage.CreateNewVAT(taxName, taxType, taxValue);
                }

                settingsPurchasingPage.SearchVAT(taxName, taxType);
                settingsPurchasingPage.EditVATAccountForSpain(purchaseCode, purchaseAccount, salesCode, salesAccount, RNCode, RNAccount);
            }
            catch
            {
                return false;
            }

            return true;
        }

        public bool VerifyCategoryAndVAT(HomePage homePage, string category, string vat, string customerType, bool isOK = true)
        {
            // Prepare
            string account = "60105100";
            string exoAccount = "60105100";

            try
            {
                // Act
                var accountingParametersPage = homePage.GoToParameters_AccountingPage();
                accountingParametersPage.GoToTab_ServiceCategoryVats();

                if (!isOK && accountingParametersPage.IsServiceCategoryAndTaxPresent(category, vat))
                {
                    accountingParametersPage.DeleteCategory(category, vat);
                }
                else
                {
                    if (!accountingParametersPage.IsServiceCategoryAndTaxPresent(category, vat))
                    {
                        accountingParametersPage.CreateNewServiceCategory(category, vat, customerType);
                    }

                    accountingParametersPage.SearchServiceCategory(category, vat);
                    accountingParametersPage.EditInventoryAccounts(account, exoAccount);
                }
            }
            catch
            {
                return false;
            }

            return true;
        }

        public DateTime VerifyIntegrationDate(HomePage homePage)
        {
            // Act
            var accountingParametersPage = homePage.GoToParameters_AccountingPage();
            //accountingParametersPage.GoToTab_AccountSettings();//Decommenté pour la nouvelle version 31/08/2021 Cécile

            accountingParametersPage.GoToTab_MonthlyClosingDays();

            //return accountingParametersPage.GetSageClosureDayIndex();//Decommenté pour la nouvelle version 31/08/2021 Cécile

            return accountingParametersPage.GetSageClosureMonthIndex();
        }

        private bool VerifyCustomer(HomePage homePage, string customer, bool isOK = true)
        {
            // Prepare
            string thirdId = isOK ? "4120C" : "";
            string accountingId = isOK ? "4300000" : "";

            try
            {
                // Act
                var customerPage = homePage.GoToCustomers_CustomerPage();
                customerPage.Filter(CustomerPage.FilterType.Search, customer);
                var customerGeneralInfo = customerPage.SelectFirstCustomer();

                customerGeneralInfo.SetAccountingId(accountingId);
                customerGeneralInfo.SetThirdId(thirdId);
            }
            catch
            {
                return false;
            }

            return true;
        }

        private bool VerifyAccountingJournal(HomePage homePage, string site, string journalInvoice)
        {
            try
            {
                var accountingJournalPage = homePage.GoToParameters_AccountingPage();
                accountingJournalPage.GoToTab_Journal();
                accountingJournalPage.EditJournal(site, journalInvoice);
            }
            catch
            {
                return false;
            }

            return true;
        }

        private void CheckErrorMailSentToUser(HomePage homePage, MailPage mailPage)
        {
            // Prepare
            var userPassword = TestContext.Properties["MailBox_UserPassword"].ToString();
            var userMail = TestContext.Properties["Admin_UserName"].ToString();

            bool mailboxTab = false;

            try
            {
                // Act
                mailPage = homePage.RedirectToOutlookMailbox();
                mailboxTab = true;

                mailPage.FillFields_LogToOutlookMailbox(userMail);

                // Recherche du mail envoyé
                WebDriver.Navigate().Refresh();
                bool isMailFound = mailPage.CheckIfSpecifiedOutlookMailExist("Export invoice to sage - Error report");
                Assert.IsTrue(isMailFound, "Aucun mail n'a été envoyé pour cette erreur.");

                // Suppression du mail et déconnexion
                mailPage.DeleteCurrentOutlookMail();
                mailPage.Close();
                mailboxTab = false;
            }

            finally
            {
                if (mailboxTab)
                {
                    mailPage.Close();
                }
            }
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_ValidateAvecServiceScale()
        {
            // Prepare
            string customerType = "BC";
            string serviceName = "INVOICE AUTO FLIGHTT" + DateTime.Now.Day + DateTime.Now.Month + DateTime.Now.Hour + DateTime.Now.Minute + DateTime.Now.Second;
            string site = TestContext.Properties["SiteLP"].ToString();
            DateTime fromDate = DateUtils.Now.AddDays(-10);
            DateTime toDate = DateUtils.Now.AddDays(10);
            string deleteFrom = DateUtils.Now.AddMonths(-1).ToString("dd/MM/yyyy");
            string deleteTo = DateUtils.Now.AddMonths(3).ToString("dd/MM/yyyy");
            //flight
            string customerBob = TestContext.Properties["CustomerSchedule"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();
            //Act
            // Création du service
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            servicePage.Filter(ServicePage.FilterType.ShowAll, true);

            try
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, null, null, null, customerType);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                priceModalPage.FillFields_CustomerPrice(site, customerBob, fromDate, toDate);
                pricePage.UnfoldAll();
                ServiceCreatePriceModalPage serviceEditPriceModal = pricePage.ServiceEitPriceModal(1);

                //Add method Scale
                serviceEditPriceModal.SetMethodScaleWithZeroNbPersonsWithoutSAVE("Scale", 1, "2");
                pricePage = priceModalPage.Save();
                pricePage.UnfoldAll();
                pricePage.ServiceEitPriceModal(1);
                var scaleListt = serviceEditPriceModal.GetAllScalePrice();
                serviceEditPriceModal.Close();
                //search service by Name
                homePage.GoToCustomers_ServicePage();
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
                var invoicesPage = homePage.GoToAccounting_InvoicesPage();
                var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
                var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customerBob, true);
                // Ajout d'un service
                invoiceDetails.AddService(serviceName, "10");
                invoiceDetails.WaitPageLoading();
                invoiceDetails.Validate();
                invoiceDetails.ConfirmValidation();
                // Récupération de l'identifiant de l'invoice créée
                string number = invoiceDetails.GetInvoiceNumber();
                invoicesPage = invoiceDetails.BackToList();
                invoicesPage.ResetFilters();
                // On filtre pour conserver l'invoice que l'on vient de créer
                invoicesPage.Filter(InvoicesPage.FilterType.ByNumber, number);

                // Assert
                var checkValidation = invoicesPage.CheckValidation(true);
                Assert.IsTrue(checkValidation, "L'invoice créée n'a pas été validée.");
                invoiceDetails = invoicesPage.SelectFirstInvoice();
                var isValidate = invoiceDetails.IsValidate();
                Assert.IsTrue(isValidate, "L'invoice créée n'a pas été validée.");
            }
            finally
            {
                //Delete service
                servicePage = homePage.GoToCustomers_ServicePage();
                var serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.From, deleteFrom);
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.To, deleteTo);
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();
            }

        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_AirportFee_Footer_Print()
        {
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string docFileNamePdfBegin = "Print Invoice_-_";
            string docFileNameZipBegin = "All_files_";

            // Arrange
            var homePage = LogInAsAdmin();


            var accountingPage = homePage.GoToParameters_AccountingPage();
            accountingPage.GoToTab_AirportFee();
            var VATRATFromAirportFee = accountingPage.GetVATRATFromAirportFee();


            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            invoicesPage.ClearDownloads();

            // Création d'une nouvelle invoice manuelle
            var invoiceCreateModalPage = invoicesPage.ManualInvoiceCreatePage();
            var invoiceDetails = invoiceCreateModalPage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customer, true);

            // Récupérer l'identifiant de la facture créée
            string ID = invoiceDetails.GetInvoiceNumber();
            Assert.IsFalse(string.IsNullOrEmpty(ID), "Erreur : Le numéro de facture est vide ou non valide.");
            invoiceDetails.Validate();
            invoicesPage = homePage.GoToAccounting_InvoicesPage();
            invoicesPage.WaitLoading();
            invoicesPage.ResetFilters();
            var invoiceNumber = invoicesPage.GetInvoiceNumberById(ID).Replace("N°", "").Trim();
            invoicesPage.Filter(InvoicesPage.FilterType.ByNumber, invoiceNumber);

            invoicesPage.ClickFirstLine();
            // Comparaison des taux VATRAT
            invoiceDetails.ClickOnInvoiceFooter();
            var VATRATFromFooterInvoice = invoiceDetails.GetVATRATFromFooterInvoice().Replace("%", "").Trim();
            Assert.AreEqual(VATRATFromFooterInvoice, VATRATFromAirportFee,
                "Le taux de VAT Airport Fee dans le pied de facture et la page Airport Fee ne sont pas égaux.");
            //Clear
            invoicesPage.ClearDownloads();
            DeleteAllFileDownload();
            // Impression du rapport
            var reportPage = invoiceDetails.PrintInvoiceResults();
            bool isReportGenerated = reportPage.IsReportGenerated();
            Assert.IsTrue(isReportGenerated, "Échec : Le rapport n'a pas pu être généré lors de l'impression de la facture.");

            reportPage.Close();
            invoiceDetails.ClickPrintButton();
            reportPage.ClickPrintButton();
            // Vérification que le fichier généré existe
            reportPage.Purge(downloadsPath, docFileNamePdfBegin, docFileNameZipBegin);
            string generatedFilePath = reportPage.PrintAllZip(downloadsPath, docFileNamePdfBegin, docFileNameZipBegin);

            FileInfo fileInfo = new FileInfo(generatedFilePath);
            fileInfo.Refresh();
            Assert.IsTrue(fileInfo.Exists, $"Échec : Le fichier {generatedFilePath} n'a pas été généré.");

            // Lecture et vérification du contenu du fichier PDF
            PdfDocument document = PdfDocument.Open(fileInfo.FullName);
            List<string> words = new List<string>();
            foreach (Page page in document.GetPages())
            {
                words.AddRange(page.GetWords().Select(word => word.Text));
            }

            // Validation des segments du taux VAT dans le PDF
            string[] VatRat = VATRATFromFooterInvoice.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string segment in VatRat)
            {
                bool segmentPresent = words.Any(word => word.Contains(segment));
                Assert.IsTrue(segmentPresent, $"Échec : Le segment '{segment}' du taux VAT n'a pas été trouvé dans le PDF.");
            }

            bool VATRATPresent = words.Any(word => word.Contains(VATRATFromFooterInvoice));
            Assert.IsTrue(VATRATPresent, $"Échec : Le taux de TVA '{VATRATFromFooterInvoice}' n'a pas été trouvé dans le PDF.");

            // Réinitialiser les filtres des factures
            homePage.GoToAccounting_InvoicesPage();
            invoicesPage.ResetFilters();
            invoicesPage.Filter(InvoicesPage.FilterType.ByNumber, invoiceNumber);
            invoicesPage.Filter(InvoicesPage.FilterType.DateFrom, DateTime.Now);

            //Clear
            invoicesPage.ClearDownloads();
            DeleteAllFileDownload();

            // Ré-impression du rapport
            invoicesPage.PrintInvoiceResults();
            Assert.IsTrue(reportPage.IsReportGenerated(), "L'application n'a pas pu générer le fichier attendu.");
            reportPage.Close();


            // Vérification finale du fichier généré
            reportPage.Purge(downloadsPath, docFileNamePdfBegin, docFileNameZipBegin);
            string finalGeneratedFilePath = reportPage.PrintAllZip(downloadsPath, docFileNamePdfBegin, docFileNameZipBegin, false);
            FileInfo finalFileInfo = new FileInfo(finalGeneratedFilePath);
            finalFileInfo.Refresh();
            Assert.IsTrue(finalFileInfo.Exists, $"{finalGeneratedFilePath} n'a pas été généré.");

            // Lecture et validation finale du PDF
            PdfDocument finalDocument = PdfDocument.Open(finalFileInfo.FullName);
            List<string> finalWords = new List<string>();
            foreach (Page page in finalDocument.GetPages())
            {
                finalWords.AddRange(page.GetWords().Select(word => word.Text));
            }

            foreach (string segment in VatRat)
            {
                bool segmentPresent = finalWords.Any(word => word.Contains(segment));
                Assert.IsTrue(segmentPresent, $"Échec : Le segment '{segment}' du taux VAT n'a pas été trouvé dans le PDF.");
            }

            bool finalVATRATPresent = finalWords.Any(word => word.Contains(VATRATFromFooterInvoice));
            Assert.IsTrue(finalVATRATPresent, $"Échec : Le taux de TVA '{VATRATFromFooterInvoice}' n'a pas été trouvé dans le PDF.");
        }

        [Priority(0)]
        [TestMethod]
        [Timeout(_timeout)]
        public void AC_INVO_Send_Customer_Invoice()
        {
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string customerIcao = TestContext.Properties["InvoiceCustomerCode"].ToString();
            string customerType = TestContext.Properties["InvoiceCustomerType"].ToString();

            string serviceName = TestContext.Properties["InvoiceService"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            ClearCache();

            // Act
            //Act
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            var invoiceCreateModalpage = invoicesPage.ManualInvoiceCreatePage();
            var invoiceDetails = invoiceCreateModalpage.FillField_CreatNewManualInvoice(site, DateUtils.Now, customer, true);
            // si pas de service alors pas de pdf joint  au mail après validation, donc pas de mails
            if (invoiceDetails.IsDev())
            {
                invoiceDetails.AddService("HOT MEAL YC UAE", "10");
            }
            else
            {
                invoiceDetails.AddService("HOT MEAL SPML FPML BC UAE", "10");
            }

            invoiceDetails.ShowValidationMenu();

            invoiceDetails.Validate();
            invoiceDetails.SendCustomerInvoice();
            invoiceDetails.NextButton();
            bool verifpassed = invoiceDetails.VerifPassedSendCustomerIvoice();
            Assert.IsTrue(verifpassed, "Erreur au niveau Send Customer Invoice");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_FilterInvoiceStepAccounted()
        {
            // prepare
            bool value = true;
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            int toCheck0 = 0;
            int toCheck1 = 1;

            //Arrange
            HomePage homePage = LogInAsAdmin();
            InvoicesPage invoicesPage = homePage.GoToAccounting_InvoicesPage();

            //filter Invoice
            invoicesPage.ResetFilters();
            invoicesPage.Filter(InvoicesPage.FilterType.ByNumber, manualInvoiceNumber);
            invoicesPage.ShowInvoiceStep();
            invoicesPage.FilterInvoiceStep(InvoicesPage.FilterInvoiceStepType.ACCOUNTED, value);

            //Assert
            int nmbrOfInvoice = invoicesPage.CheckTotalNumber();
            Assert.AreEqual(nmbrOfInvoice, toCheck0, "La facture est affichée malgré qu'elle n'est pas exportée vers Sage/Cegid.");

            //Access the invoice
            invoicesPage.FilterInvoiceStep(InvoicesPage.FilterInvoiceStepType.ALL, value);
            InvoiceDetailsPage details = invoicesPage.ClickFirstLine();

            //Export for SAGE/CEGID
            details.ClearDownloads();
            details.ShowExtendedMenu();
            details.ExportSage();
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();
            FileInfo correctDownloadedFile = invoicesPage.GetExportSageFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");
            details.BackToList();

            //filter Invoice
            invoicesPage.ResetFilters();
            invoicesPage.Filter(InvoicesPage.FilterType.ByNumber, manualInvoiceNumber);
            invoicesPage.ShowInvoiceStep();
            invoicesPage.FilterInvoiceStep(InvoicesPage.FilterInvoiceStepType.ACCOUNTED, value);

            //Assert
            nmbrOfInvoice = invoicesPage.CheckTotalNumber();
            Assert.AreEqual(nmbrOfInvoice, toCheck1, "La facture n'est pas affichée bien qu'elle ait été exportée vers Sage/Cegid.");
            bool verifDisquette = invoicesPage.IsSentToSage();
            Assert.IsTrue(verifDisquette, "L'icône de disquette verte indiquant 'Sent to Sage manually' est absente.");
        }


        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_FilterShow_Showallinvoices()
        {
            // prepare
            bool filtre = true;
            int verif = 0;
            // arrange
            var homePage = LogInAsAdmin();
            // act
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            // filtre by number
            invoicesPage.Filter(InvoicesPage.FilterType.ByNumber, manualInvoiceNumber);
            // filtre by SHOW ALL INVOICES Not validated in SHOW
            invoicesPage.ShowInvoiceShow();
            invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowAll, filtre);
            invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowNotValidatedOnly, filtre);
            var count = invoicesPage.CheckTotalNumber();
            // assert
            Assert.AreEqual(verif, count, "le filtre show all n'est pas appliqué");
            // filtre by SHOW ALL INVOICES  in SHOW
            invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowAllInvoices, filtre);
            count = invoicesPage.CheckTotalNumber();
            verif = 1;
            // assert
            Assert.AreEqual(verif, count, "le filtre show all n'est pas appliqué");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_FilterShow_ShownotsenttoSageCegidonly()
        {
            // Récupération du site
            string site = TestContext.Properties["InvoiceSite"].ToString();

            //Arrange
            HomePage homePage = LogInAsAdmin();

            // Vérification et configuration du site pour s'assurer que l'option "Enable Sage/Cegid Auto" est activée
            bool isEnableSageCegidChecked = VerifySiteSageCegidAuto(homePage, site);

            InvoicesPage invoicesPage = homePage.GoToAccounting_InvoicesPage();

            // Application d'un filtre sur les factures en fonction de leur numéro de invoice
            invoicesPage.Filter(InvoicesPage.FilterType.ByNumber, manualInvoiceNumber);
            invoicesPage.ShowInvoiceShow();

            // Filtrage des factures pour afficher uniquement celles qui n'ont pas été envoyées à Sage/Cegid
            invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowNotSentToSage, true);

            // Récupération du numéro de la première facture visible après filtrage
            string firstInvoicedNumber = invoicesPage.GetFirstInvoiceOnlyNumber();

            // Vérification que la facture créée correspond bien au numéro attendu
            Assert.AreEqual(manualInvoiceNumber, firstInvoicedNumber, "La facture nouvellement crée n'apparait pas.");

            // Vérification que le site est correctement configuré
            Assert.IsTrue(isEnableSageCegidChecked, "La configuration le site n'est pas effectuée.");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_FilterShow_Showmanualinvoices()
        {
            // prepare
            bool filtre = true;
            int verif = 0;
            // arrange
            var homePage = LogInAsAdmin();
            // act
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            // filtre by number
            invoicesPage.Filter(InvoicesPage.FilterType.ByNumber, manualInvoiceNumber);
            // filtre by SHOW ALL INVOICES Not validated in SHOW
            invoicesPage.ShowInvoiceShow();
            invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowManualInvoice, filtre);
            invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowNotValidatedOnly, filtre);
            var count = invoicesPage.CheckTotalNumber();
            // assert
            Assert.AreEqual(verif, count, "le filtre show manualy n'est pas appliqué");
            // filtre by SHOW ALL INVOICES  in SHOW
            invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowAllInvoices, filtre);
            count = invoicesPage.CheckTotalNumber();
            verif = 1;
            // assert
            Assert.AreEqual(verif, count, "le filtre show manualy n'est pas appliqué");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_FilterShow_Shownotvalidareonly()
        {
            // prepare
            bool filtre = true;
            DateTime date = DateTime.Now;
            bool notValidated = false;
            string filterName = "'Show not validated only'";
            // arrange
            HomePage homePage = LogInAsAdmin();
            // act
            InvoicesPage invoicesPage = homePage.GoToAccounting_InvoicesPage();
            // filter by date
            invoicesPage.ResetFilters();
            invoicesPage.Filter(InvoicesPage.FilterType.DateFrom, date);
            invoicesPage.Filter(InvoicesPage.FilterType.DateTo, date);
            // filter by SHOW ALL INVOICES Not validated in SHOW
            invoicesPage.ShowInvoiceShow();
            invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowAllInvoices, filtre);
            invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowNotValidatedOnly, filtre);
            // Assert
            bool checkValidation = invoicesPage.CheckValidation(notValidated);
            Assert.IsFalse(checkValidation, String.Format(MessageErreur.FILTRE_ERRONE, filterName));
            bool checkInvoiceExist = invoicesPage.IsManualInvoiceExists(manualInvoiceNumber);
            Assert.IsTrue(checkInvoiceExist, String.Format(MessageErreur.FILTRE_ERRONE, filterName));
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_FilterInvoiceStep_ValidatedAccounted()
        {
            //Prepare
            DateTime dateFrom = DateUtils.Now.AddDays(-1);
            DateTime dateTo = DateUtils.Now;
            bool value = true;
            //Arrange
            HomePage homePage = LogInAsAdmin();
            // act
            InvoicesPage invoicesPage = homePage.GoToAccounting_InvoicesPage();
            // filtre by number
            invoicesPage.ResetFilters();
            invoicesPage.ShowInvoiceStep();
            //show validate invoice 
            invoicesPage.FilterInvoiceStep(InvoicesPage.FilterInvoiceStepType.VALIDATED_ACCOUNTED, value);
            //Assert
            bool isInvoiceExist = invoicesPage.GetFirstInvoiceID().Contains(ID);
            Assert.IsFalse(isInvoiceExist, "L'invoice créée n'apparaît pas dans la liste des invoices.");
            invoicesPage.FilterInvoiceStep(InvoicesPage.FilterInvoiceStepType.ALL, value);
            invoicesPage.Filter(InvoicesPage.FilterType.DateFrom, dateFrom);
            invoicesPage.Filter(InvoicesPage.FilterType.DateTo, dateTo);
            InvoiceDetailsPage invoiceDetails = invoicesPage.ClickFirstLine();
            //validate invoice
            invoiceDetails.Validate();
            manualInvoiceNumber = invoiceDetails.GetInvoiceNumber();
            invoiceDetails.BackToList();
            //show validate invoice 
            invoicesPage.ResetFilters();
            //filter by number
            invoicesPage.Filter(InvoicesPage.FilterType.ByNumber, manualInvoiceNumber);
            invoicesPage.ShowInvoiceStep();
            //show validate invoice 
            invoicesPage.FilterInvoiceStep(InvoicesPage.FilterInvoiceStepType.VALIDATED_ACCOUNTED, value);
            int nbResult = invoicesPage.CheckTotalNumber();
            //Assert
            Assert.AreEqual(nbResult, 1, "La facture envoie manuelle validée n'apparaît pas.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_FilterInvoiceStepDraft()
        {
            //Prepare 
            DateTime dateFrom = DateUtils.Now.AddDays(-1);
            DateTime dateTo = DateUtils.Now;
            string draft = "Draft";
            bool value = true;

            //Arrange
            HomePage homePage = LogInAsAdmin();

            // act
            InvoicesPage invoicesPage = homePage.GoToAccounting_InvoicesPage();
         
            // filter by Draft and Date
            invoicesPage.ResetFilters();
            invoicesPage.ShowInvoiceStep();
            invoicesPage.FilterInvoiceStep(InvoicesPage.FilterInvoiceStepType.DRAFT, value);
            invoicesPage.Filter(InvoicesPage.FilterType.DateFrom, dateFrom);
            invoicesPage.Filter(InvoicesPage.FilterType.DateTo, dateTo);

            //invoice Draft show
            bool isInvoiceDraftOk = invoicesPage.IsInvoiceDraft(draft); 
            bool isInvoiceExist = invoicesPage.GetFirstInvoiceID().Contains(ID);
            //Assert
            Assert.IsTrue(isInvoiceDraftOk, "L'invoice créée n'apparaît pas comme Draft.");
            Assert.IsTrue(isInvoiceExist, "L'invoice créée n'apparaît pas dans la liste des invoices.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_FilterShow_Shownotsentbymailonly()
        {
            //Arrange
            HomePage homePage = LogInAsAdmin();

            InvoicesPage invoicesPage = homePage.GoToAccounting_InvoicesPage();

            // Application d'un filtre sur les factures en fonction de leur numéro de invoice
            invoicesPage.Filter(InvoicesPage.FilterType.ByNumber, manualInvoiceNumber);
            invoicesPage.ShowInvoiceShow();

            // Filtrage des factures pour afficher uniquement celles qui n'ont pas été envoyées par mail
            invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowNotSentByMailOnly, true);

            // Récupération du numéro de la première facture visible après filtrage
            string firstInvoicedNumber = invoicesPage.GetFirstInvoiceOnlyNumber();

            //Assert
            Assert.AreEqual(manualInvoiceNumber, firstInvoicedNumber, "La facture nouvellement crée n'apparait pas.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_FilterInvoiceStepAll()
        {
            DateTime dateFrom = DateUtils.Now.AddDays(-1);
            DateTime dateTo = DateUtils.Now;
            //Arrange
            HomePage homePage = LogInAsAdmin();
            InvoicesPage invoicesPage = homePage.GoToAccounting_InvoicesPage();
            invoicesPage.ResetFilters();
            invoicesPage.ShowInvoiceStep();
            //Filter
            invoicesPage.FilterInvoiceStep(InvoicesPage.FilterInvoiceStepType.ALL,true);
            invoicesPage.Filter(InvoicesPage.FilterType.DateFrom, dateFrom);
            invoicesPage.Filter(InvoicesPage.FilterType.DateTo, dateTo);
            var invoiceDetails = invoicesPage.ClickFirstLine();
            //validate du facture
            invoiceDetails.Validate();
            manualInvoiceNumber = invoiceDetails.GetInvoiceNumber();
            invoiceDetails.BackToList();
            invoicesPage.Filter(InvoicesPage.FilterType.ByNumber, manualInvoiceNumber);
            var nbResult = invoicesPage.CheckTotalNumber();
            //Assert
            Assert.AreEqual(nbResult, 1, "La facture nouvellement créée n'apparaît pas.");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_Filter_SortById()
        {
            //Arrange
            HomePage homePage = LogInAsAdmin();

            InvoicesPage invoicesPage = homePage.GoToAccounting_InvoicesPage();
            invoicesPage.ResetFilters();

            //Filter
            invoicesPage.Filter(InvoicesPage.FilterType.SortBy, "Id");
            bool isSortedById = invoicesPage.IsSortedById();

            //Assert
            Assert.IsTrue(isSortedById, "Les invoices affichées ne sont pas filtrées par Id.");
        }

    }
}
