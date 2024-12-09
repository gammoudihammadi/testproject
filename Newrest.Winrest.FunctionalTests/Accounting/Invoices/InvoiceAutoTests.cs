using DocumentFormat.OpenXml.Drawing.Charts;
using iText.Signatures;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Accounting.FreePrice;
using Newrest.Winrest.FunctionalTests.PageObjects.Accounting.Invoice;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Customer;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Delivery;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Reinvoice;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.Flight;
using Newrest.Winrest.FunctionalTests.PageObjects.Interim.InterimOrders;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Production;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.Dispatch;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.Production;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Security.Cryptography;
using System.Security.Policy;
using System.Text.RegularExpressions;
using System.Threading;
using System.Web.Configuration;
using System.Web.UI;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using static Newrest.Winrest.FunctionalTests.PageObjects.Accounting.Invoice.InvoicesPage;
using static Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service.ServiceMassiveDeleteModalPage;

namespace Newrest.Winrest.FunctionalTests.Accounting
{
    [TestClass]
    public class InvoiceAutoTests : TestBase
    {

        private const int _timeout = 600000;
        private readonly string PURCHASE_CODE = "AS05";
        private readonly string PURCHASE_ACCOUNT = "47205001";
        private readonly string SALES_CODE = "AR10";
        private readonly string SALES_ACCOUNT = "47205001";
        private readonly string DUE_TO_INVOICE_CODE = "AP10";
        private readonly string DUE_TO_INVOICE_ACCOUNT = "47205001";
        private readonly string CUSTOMER = $"TestCustomer_COInvoiceAuto";
        private readonly string CUSTOMERICAO = DateTime.Now.Ticks.ToString().Substring(9);
        private readonly string CUSTOMERTYPE = "Colectividades";
        private readonly string SERVICECATEGORY = "COLECTIVIDADES";
        private readonly string SERVICENAME = $"TestService_COInvoiceAuto";
        private readonly string SERVICEQUANTITY = "10";
        private readonly string SITE = "BCN";
        private readonly string DELIVERY = $"TestDelivery_COInvoiceAuto";
        private readonly string serviceNameToday = "Service-" + DateUtils.Now.ToString("dd/MM/yyyy");
        private readonly string PUPUP_TITLE = "//*[@id=\"dataAlertLabel\"]";
        private readonly string TOTAL_PRICE = "//*[@id=\"dataAlertModal\"]/div/div/div[2]/div/div[3]/div/div/p/b";
        private readonly string OK_BTN = "//*[@id=\"dataAlertCancel\"]";

        private string AutoInvoiceNumber;
        private string flightNumber;

        [TestInitialize]
        public override void TestInitialize()
        {
            base.TestInitialize();

            var testMethod = TestContext.TestName;
            switch (testMethod)
            {
                case nameof(AC_INVO_FilterShow_Showflightinvoices):
                    TestInitialize_CreateCustomerForInvoiceAuto();
                    TestInitialize_CreateNewFlight();
                    TestInitialize_CreateAutoInvoiceValidate();

                    break;


                default:
                    break;
            }
        }
        public void TestInitialize_CreateCustomerForInvoiceAuto()
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

                customersPage = customerGeneralInformationsPage.BackToList();
                customersPage.ResetFilters();
                customersPage.Filter(CustomerPage.FilterType.Search, customer);
            }

            //Assert
            var checkCustomer = customersPage.GetFirstCustomerName();
            Assert.AreEqual(customer, checkCustomer, "Le customer n'a pas été créé.");
        }
        private void TestInitialize_CreateAutoInvoiceValidate()
        {

            // Initialisation des données pour le test.
            string site = TestContext.Properties["SiteLpCart"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string customerPickMethod = TestContext.Properties["AllFlightsInSelectedPeriod"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            invoicesPage.ResetFilters(); // Réinitialisation des filtres.

            // Création d'une nouvelle facture automatique.
            AutoInvoiceCreateModalPage invoiceDetails = invoicesPage.AutoInvoiceCreatePage();
            invoiceDetails.FillField_CreateNewAutoInvoice(customer, site, customerPickMethod); 
            invoiceDetails.CheckBoxSeparatedInvoices(); 

            // Validation et récupération du numéro de facture.
            invoiceDetails.Submit(); 
            invoiceDetails.Validate(); 
            AutoInvoiceNumber = invoiceDetails.GetInvoiceNumber(); 

        }
        private void TestInitialize_CreateNewFlight()
        {

           // Initialisation des données nécessaires pour le test.
            string site = TestContext.Properties["SiteLpCart"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string aircraft = TestContext.Properties["AircraftBis"].ToString();
            string siteTo = TestContext.Properties["SiteToFlightBob"].ToString();
            string serviceName = TestContext.Properties["InvoiceService"].ToString();
            flightNumber = "flight_lp" + new Random().Next(1000, 5000).ToString(); // Génération d'un numéro de vol aléatoire.


            // Arrange
            var homePage = LogInAsAdmin();

            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter(); // Réinitialisation des filtres.
            flightPage.Filter(FlightPage.FilterType.Sites, site);

            // Création d'un nouveau vol.
            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, site, siteTo);


            // Recherche et modification du vol créé.
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            var editPage = flightPage.EditFirstFlight(flightNumber);
            editPage.AddGuestType(); // Ajout d'un type d'invité.
            editPage.AddService(serviceName); // Ajout d'un service.
            editPage.ClickOnValidateAll(); // Validation des informations.
            editPage.ClickOnInvoieAll(); // Génération des factures.

            // Fermeture de la vue des détails
            flightPage = editPage.CloseViewDetails();

        }
        [TestMethod]
        [Priority(0)]
        [Timeout(_timeout)]
        public void AC_INVO_CreateCustomerForInvoiceAuto()
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

                customersPage = customerGeneralInformationsPage.BackToList();
                customersPage.ResetFilters();
                customersPage.Filter(CustomerPage.FilterType.Search, customer);
            }

            //Assert
            Assert.AreEqual(customer, customersPage.GetFirstCustomerName(), "Le customer n'a pas été créé.");
        }

        [TestMethod]
        [Priority(1)]
        [Timeout(_timeout)]
        public void AC_INVO_CreateCustomerForInvoiceAutoAirportTax()
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

        [Priority(3)]
        [TestMethod]
        [Timeout(_timeout)]
        public void AC_INVO_CreateDatasForInvoiceAutoDeliveries()
        {
            // Prepare
            string customerType = TestContext.Properties["CustomerType3"].ToString();
            string customerName = "Customer for Invoice Deliveries";
            string customerCode = "CID";
            string serviceName = "Service for Invoice Deliveries";
            string site = TestContext.Properties["SiteLP"].ToString();

            DateTime fromDate = DateUtils.Now.AddDays(-10);
            DateTime toDate = DateUtils.Now.AddDays(10);

            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var customersPage = homePage.GoToCustomers_CustomerPage();
            customersPage.ResetFilters();
            customersPage.Filter(CustomerPage.FilterType.Search, customerName);

            if (customersPage.CheckTotalNumber() == 0)
            {
                var customerCreateModalPage = customersPage.CustomerCreatePage();
                customerCreateModalPage.FillFields_CreateCustomerModalPage(customerName, customerCode, customerType);
                var customerGeneralInformationsPage = customerCreateModalPage.Create();
                customersPage = customerGeneralInformationsPage.BackToList();

                customersPage.ResetFilters();
                customersPage.Filter(CustomerPage.FilterType.Search, customerName);
            }

            Assert.AreEqual(customerName, customersPage.GetFirstCustomerName(), "Le customer n'a pas été créé.");

            // Création du service
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customerName, fromDate, toDate);
                servicePage = pricePage.BackToList();
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            }
            else
            {
                var pricePage = servicePage.ClickOnFirstService();
                pricePage.ToggleFirstPrice();
                var serviceCreatePriceModalPage = pricePage.EditFirstPrice(site, customerName);
                try
                {
                    serviceCreatePriceModalPage.EditPriceDates(fromDate, toDate);
                }
                catch
                {
                    serviceCreatePriceModalPage.Close();
                }

                var serviceGeneralInformationsPage = pricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                servicePage = serviceGeneralInformationsPage.BackToList();
            }

            //Assert
            Assert.IsTrue(servicePage.GetFirstServiceName().Contains(serviceName), "Le service n'a pas été créé.");
        }

        [TestMethod]
        [Priority(2)]
        [Timeout(_timeout)]
        public void AC_INVO_CreateServiceForInvoiceAutoAirportTax()
        {
            string serviceInvoice1 = TestContext.Properties["InvoiceServiceAirportTax"].ToString();
            string serviceInvoice2 = TestContext.Properties["InvoiceService"].ToString();
            string serviceCategory = TestContext.Properties["InvoiceServiceCategoryAirportTax"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customerAT = TestContext.Properties["InvoiceCustomerAirportTax"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();
            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();

            // -------------------------------------- service 1 -----------------------------------------------
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceInvoice1);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceInvoice1, null, null, serviceCategory);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customerAT, DateUtils.Now.AddMonths(-1), DateUtils.Now.AddMonths(+3));
                servicePage = pricePage.BackToList();
            }
            else
            {
                var servicePricePage = servicePage.ClickOnFirstService();
                servicePricePage.SearchPriceForCustomer(customerAT, site, DateUtils.Now.AddMonths(-1), DateUtils.Now.AddMonths(+3));

                var serviceGeneralInformationPage = servicePricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationPage.SetCategory(serviceCategory);
                servicePage = serviceGeneralInformationPage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceInvoice1);

            string serviceInvoiceAirportTax = servicePage.GetFirstServiceName();

            // -------------------------------------- service 2 -----------------------------------------------

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceInvoice2);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceInvoice2, null, null, serviceCategory);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, DateUtils.Now.AddMonths(-1), DateUtils.Now.AddMonths(+3));
                servicePage = pricePage.BackToList();
            }
            else
            {
                var servicePricePage = servicePage.ClickOnFirstService();
                servicePricePage.SearchPriceForCustomer(customer, site, DateUtils.Now.AddMonths(-1), DateUtils.Now.AddMonths(+3));

                var serviceGeneralInformationPage = servicePricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationPage.SetCategory(serviceCategory);
                servicePage = serviceGeneralInformationPage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceInvoice2);

            string serviceInvoice = servicePage.GetFirstServiceName();

            //Assert           
            Assert.IsTrue(serviceInvoiceAirportTax.Contains(serviceInvoice1), "Le service " + serviceInvoiceAirportTax + " n'existe pas.");
            Assert.IsTrue(serviceInvoice.Contains(serviceInvoice2), "Le service " + serviceInvoice + " n'existe pas.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_CreateInvoiceAuto()
        {
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string service = TestContext.Properties["InvoiceService"].ToString();
            string customerPickMethod = TestContext.Properties["AllFlightsInSelectedPeriod"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            // Flight  part
            CreateFlightForInvoice(homePage, site, customer, service);

            //Invoice part
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            invoicesPage.ResetFilters();
            invoicesPage.Filter(InvoicesPage.FilterType.ByNumber, "");
            var invoiceCreateAutoModalpage = invoicesPage.AutoInvoiceCreatePage();
            invoiceCreateAutoModalpage.FillField_CreateNewAutoInvoice(customer, site, customerPickMethod);
            var invoiceDetails = invoiceCreateAutoModalpage.Submit();

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
            invoicesPage.Filter(InvoicesPage.FilterType.ByNumber, numInvoice);

            Assert.IsTrue(invoicesPage.GetFirstInvoiceNumber().Contains(numInvoice), "L'invoice number (" + invoicesPage.GetFirstInvoiceNumber() + "==" + numInvoice + ") créée n'apparaît pas dans la liste des invoices.");
            Assert.IsTrue(invoicesPage.GetFirstInvoiceID().Contains(idInvoice), "L'invoice id (" + invoicesPage.GetFirstInvoiceID() + "==" + idInvoice + ") créée n'apparaît pas dans la liste des invoices.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_CreateInvoiceAutoWithAirportTax()
        {
            string customer = TestContext.Properties["InvoiceCustomerAirportTax"].ToString();
            string customerCode = TestContext.Properties["InvoiceCustomerCodeAirportTax"].ToString();
            string serviceName = TestContext.Properties["InvoiceServiceAirportTax"].ToString();
            string taxName = TestContext.Properties["InvoiceTaxName"].ToString();
            string taxType = TestContext.Properties["InvoiceTaxType"].ToString();
            string currency = TestContext.Properties["Currency"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customerPickMethod = TestContext.Properties["AllFlightsInSelectedPeriod"].ToString();
            string taxValue = "15";

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
            purchasingPage.EditVATAccountForSpain(PURCHASE_CODE, PURCHASE_ACCOUNT, SALES_CODE, SALES_ACCOUNT, DUE_TO_INVOICE_CODE, DUE_TO_INVOICE_ACCOUNT);

            // Airport tax configurée pour site et client
            var accountingPage = homePage.GoToParameters_AccountingPage();
            accountingPage.GoToTab_AirportFee();
            accountingPage.SetAirportFeeSiteAndCustomer(site, customerCode, customer, "15");

            // Catégorie de service avec AirportTax = true
            var parametersCustomerPage = homePage.GoToParameters_CustomerPage();
            parametersCustomerPage.GoToTab_Category();
            parametersCustomerPage.SetServiceWithAirportTax(serviceName);

            //_____________________________ Fin Mise en place des paramètres________________________

            // Création d'un vol
            // Flight  part
            CreateFlightForInvoice(homePage, site, customer, serviceName);

            // Création d'un invoice
            //Invoice part
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            invoicesPage.ResetFilters();
            invoicesPage.Filter(InvoicesPage.FilterType.ByNumber, "");
            var invoiceCreateAutoModalpage = invoicesPage.AutoInvoiceCreatePage();
            invoiceCreateAutoModalpage.FillField_CreateNewAutoInvoice(customerCode, site, customerPickMethod);
            var invoiceDetails = invoiceCreateAutoModalpage.Submit();

            string ID = null;
            try
            {
                ID = invoiceDetails.GetInvoiceNumber();
            }
            catch
            {
                //on est dans index à la place d'invoiceDetails
                ID = invoicesPage.GetFirstInvoiceID();
                invoiceDetails = invoicesPage.SelectFirstInvoice();
            }


            var invoiceFooter = invoiceDetails.ClickOnInvoiceFooter();

            // On vérifie que la ligne Airport Tax est présente et qu'une valeur lui est associée
            var airportTaxOK = invoiceFooter.IsAirportTaxPresent(currency);
            Assert.IsTrue(airportTaxOK, "L'airport tax n'est pas présente pour l'invoice créée.");

            invoicesPage = invoiceFooter.BackToList();

            invoicesPage.ResetFilters();

            Assert.IsTrue(invoicesPage.GetFirstInvoiceID().Contains(ID), "L'invoice créée n'apparaît pas dans la liste des invoices.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_CreateCustomerInvoiceAutoSeparate()
        {
            //Prepare
            string site = TestContext.Properties["SiteLP"].ToString();
            string customerPickMethod = "All deliveries in selected period";
            Random rnd = new Random();
            string deliveryNameToday = "Delivery-" + DateUtils.Now.ToString("dd/MM/yyyy");
            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();
            string customerName = "Customer for Invoice Deliveries";
            string serviceName = "Service for Invoice Deliveries";

            //Arrange
            var homePage = LogInAsAdmin();


            //Vaider mes quantité sur les dispatch(qty to invoice)
            try
            {
                CreateNewDispatch(homePage, deliveryName + "_bis", customerName, serviceName);

                var dispatchPage = homePage.GoToProduction_DispatchPage();
                dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryName + "_bis");
                dispatchPage.Filter(DispatchPage.FilterType.Site, site);
                dispatchPage.WaitPageLoading();
                var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
                previsionnalQty.ValidateFirstDispatch();
                var qtyToProduce = dispatchPage.ClickQuantityToProduce();
                qtyToProduce.ValidateFirstDispatch();
                var qtytoInvoice = dispatchPage.ClickQuantityToInvoice();
                qtytoInvoice.ValidateTheFirst();

                CreateNewDispatch(homePage, deliveryName, customerName, serviceName);

                dispatchPage = homePage.GoToProduction_DispatchPage();
                dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryName);
                dispatchPage.Filter(DispatchPage.FilterType.Site, site);
                dispatchPage.WaitPageLoading();
                previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
                previsionnalQty.ValidateFirstDispatch();
                qtyToProduce = dispatchPage.ClickQuantityToProduce();
                qtyToProduce.ValidateFirstDispatch();
                qtytoInvoice = dispatchPage.ClickQuantityToInvoice();
                qtytoInvoice.ValidateTheFirst();

                //Aller sur les invoice
                InvoicesPage invoicesPage = homePage.GoToAccounting_InvoicesPage();
                invoicesPage.ResetFilters();

                //1.Survoler le + et cliquer sur invoice auto
                AutoInvoiceCreateModalPage invoiceCreate = invoicesPage.AutoInvoiceCreatePage();

                //2.Sélectionner le site, date etc...
                //3.Cliquer sur Next
                //4.Faire basculer un ou plusieurs "available customers et cliquer sur Next
                invoiceCreate.FillField_CreateNewAutoInvoice(customerName, site, customerPickMethod);
                var select = invoiceCreate.WaitForElementIsVisible(By.XPath("//*[@id='SelectedEntities']/option[1]"));
                select.Click();
                invoiceCreate.WaitPageLoading();
                var select2 = invoiceCreate.WaitForElementIsVisible(By.XPath("//*[@id='SelectedEntities']/option[2]"));
                select2.Click();
                invoiceCreate.WaitPageLoading();
                var selectAcme = invoiceCreate.WaitForElementIsVisible(By.Id("SelectedEntities"));
                SelectElement se = new SelectElement(selectAcme);

                // se.SelectByText("Delivery " + deliveryName + " - " + DateUtils.Now.ToString("yyyy-MM-dd"));
                invoiceCreate.WaitPageLoading();

                //5.Cocher "Separated invoices" et cliquer sur create
                invoiceCreate.CheckBoxSeparatedInvoices();
                invoiceCreate.Submit();

                invoiceCreate.CloseNewInvoicePopup();
                invoicesPage.Filter(InvoicesPage.FilterType.Customer, customerName);
                invoicesPage.Filter(InvoicesPage.FilterType.DateFrom, DateTime.Today);
                invoicesPage.Filter(InvoicesPage.FilterType.DateTo, DateTime.Today);
                invoicesPage.ScrollToInvoiceStep();
                invoicesPage.ShowInvoiceStep();
                invoicesPage.FilterInvoiceStep(InvoicesPage.FilterInvoiceStepType.DRAFT, true);
                invoicesPage.WaitPageLoading(); 
                Assert.IsTrue(invoicesPage.LinesCountByCustomer(customerName) == 2, "Invoices not seperated");
            }
            finally
            {
                //1. Delete invoices
                InvoicesPage invoicesPage = homePage.GoToAccounting_InvoicesPage();
                invoicesPage.ResetFilters();
                invoicesPage.Filter(InvoicesPage.FilterType.Customer, customerName);
                invoicesPage.Filter(InvoicesPage.FilterType.DateFrom, DateTime.Today);
                invoicesPage.ScrollToInvoiceStep();
                invoicesPage.ShowInvoiceStep();
                invoicesPage.FilterInvoiceStep(InvoicesPage.FilterInvoiceStepType.DRAFT, true);

                while (invoicesPage.CheckTotalNumber() > 0)
                {
                    invoicesPage.DeleteFirstInvoice();
                }

                //2. Delete Dispatch (unvalidate)
                var dispatchPage = homePage.GoToProduction_DispatchPage();
                dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryName);
                dispatchPage.Filter(DispatchPage.FilterType.Site, site);
                dispatchPage.WaitPageLoading();
                dispatchPage.ClickQuantityToInvoice();
                dispatchPage.UnValidateAll();
                dispatchPage.ClickQuantityToProduce();
                dispatchPage.UnValidateAll();
                dispatchPage.ClickPrevisonalQuantity();
                dispatchPage.UnValidateAll();

                //3. Delete Service & delivery
                var deliveryPage = homePage.GoToCustomers_DeliveryPage();
                deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName);
                var deliveryLoadingPage = deliveryPage.ClickOnFirstDelivery();
                deliveryLoadingPage.DeleteService();

                var deliveryGeneralInfo = deliveryLoadingPage.ClickOnGeneralInformation();
                deliveryGeneralInfo.SetActive(false);
                deliveryGeneralInfo.DeleteDelivery();

                deliveryPage = homePage.GoToCustomers_DeliveryPage();
                deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName + "_bis");
                deliveryPage.Filter(DeliveryPage.FilterType.Sites, site);
                dispatchPage.WaitPageLoading();
                deliveryLoadingPage = deliveryPage.ClickOnFirstDelivery();
                deliveryLoadingPage.DeleteService();

                deliveryGeneralInfo = deliveryLoadingPage.ClickOnGeneralInformation();
                deliveryGeneralInfo.SetActive(false);
                deliveryGeneralInfo.WaitPageLoading();
                deliveryGeneralInfo.DeleteDelivery();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_CannotInvoiceDeliveriesAgain()
        {
            string site = TestContext.Properties["SiteLP"].ToString();
            string customerName = "Customer for Invoice Deliveries";
            string serviceName = "Service for Invoice Deliveries";
            Random rnd = new Random();
            string deliveryName = "Delivery Again" + DateUtils.Now.ToString("dd/MM/yyyy") + "-" + rnd.Next().ToString();

            // Arrange
            var homePage = LogInAsAdmin();
            //Prepare
            CreateNewDispatch(homePage, deliveryName, customerName, serviceName);
            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryName);
            dispatchPage.Filter(DispatchPage.FilterType.Site, site);

            var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            previsionnalQty.ValidateFirstDispatch();
            var qtyToProduce = dispatchPage.ClickQuantityToProduce();
            qtyToProduce.ValidateFirstDispatch();
            var qtytoInvoice = dispatchPage.ClickQuantityToInvoice();
            qtytoInvoice.ValidateTheFirst();

            //Invoice part
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            invoicesPage.ResetFilters();
            try
            {
                var autoInvoiceCreateModalPage = invoicesPage.AutoInvoiceCreatePage();
                autoInvoiceCreateModalPage.FillField_CreateNewAutoInvoice(customerName, site, CustomerPickMethod.AllDeliveriesInSelectedPeriod);
                var invoiceDetails = autoInvoiceCreateModalPage.FillFieldSelectAll();
                autoInvoiceCreateModalPage.CloseNewInvoicePopup();

                invoicesPage.ResetFilters();
                autoInvoiceCreateModalPage = invoicesPage.AutoInvoiceCreatePage();
                autoInvoiceCreateModalPage.CreateNewAutoInvoiceFirstPart(customerName, site, CustomerPickMethod.AllDeliveriesInSelectedPeriod);

                Assert.IsTrue(autoInvoiceCreateModalPage.IsSelectCustomersEmpty(customerName), "It is possible to create invoice deliveries again.");
            }
            finally
            {
                //1. Delete invoices
                invoicesPage = homePage.GoToAccounting_InvoicesPage();
                invoicesPage.ResetFilters();
                invoicesPage.Filter(InvoicesPage.FilterType.Customer, customerName);
                invoicesPage.Filter(InvoicesPage.FilterType.DateFrom, DateTime.Today);

                while (invoicesPage.CheckTotalNumber() > 0)
                {
                    invoicesPage.DeleteFirstInvoice();
                }

                //2. Delete Dispatch (unvalidate)
                dispatchPage = homePage.GoToProduction_DispatchPage();
                dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryName);
                dispatchPage.Filter(DispatchPage.FilterType.Site, site);
                dispatchPage.ClickQuantityToInvoice();
                dispatchPage.UnValidateAll();
                dispatchPage.ClickQuantityToProduce();
                dispatchPage.UnValidateAll();
                dispatchPage.ClickPrevisonalQuantity();
                dispatchPage.UnValidateAll();

                //3. Delete Service & delivery
                var deliveryPage = homePage.GoToCustomers_DeliveryPage();
                deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName);
                var deliveryLoadingPage = deliveryPage.ClickOnFirstDelivery();
                deliveryLoadingPage.DeleteService();

                var deliveryGeneralInfo = deliveryLoadingPage.ClickOnGeneralInformation();
                deliveryGeneralInfo.SetActive(false);
                deliveryGeneralInfo.DeleteDelivery();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_CanInvoiceAgainDeletedInvoiceDeliveries()
        {
            string site = TestContext.Properties["SiteLP"].ToString();
            string customerName = "Customer for Invoice Deliveries";
            string serviceName = "Service for Invoice Deliveries";
            Random rnd = new Random();
            string deliveryName = "Delivery Again" + DateUtils.Now.ToString("dd/MM/yyyy") + "-" + rnd.Next().ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            //Prepare
            CreateNewDispatch(homePage, deliveryName, customerName, serviceName);
            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryName);
            dispatchPage.Filter(DispatchPage.FilterType.Site, site);

            var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            previsionnalQty.ValidateFirstDispatch();
            var qtyToProduce = dispatchPage.ClickQuantityToProduce();
            qtyToProduce.ValidateFirstDispatch();
            var qtytoInvoice = dispatchPage.ClickQuantityToInvoice();
            qtytoInvoice.ValidateTheFirst();

            //Invoice part
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            invoicesPage.ResetFilters();

            //1.Créer puis Supprimer la facture (non validée)
            var autoInvoiceCreateModalPage = invoicesPage.AutoInvoiceCreatePage();
            autoInvoiceCreateModalPage.FillField_CreateNewAutoInvoice(customerName, site, CustomerPickMethod.AllDeliveriesInSelectedPeriod);
            var invoiceDetails = autoInvoiceCreateModalPage.FillFieldSelectAll();
            autoInvoiceCreateModalPage.CloseNewInvoicePopup();
            invoicesPage.Filter(InvoicesPage.FilterType.Customer, customerName);
            invoicesPage.Filter(InvoicesPage.FilterType.DateFrom, DateTime.Today);
            invoicesPage.ScrollToInvoiceStep();
            invoicesPage.ShowInvoiceStep();
            invoicesPage.FilterInvoiceStep(InvoicesPage.FilterInvoiceStepType.DRAFT, true);
            invoicesPage.PageSize("100");
            int number_of_lines = invoicesPage.LinesCountByCustomer(customerName);
            Assert.IsTrue(invoicesPage.LinesCountByCustomer(customerName) == 1, "Je ne peux pas facturer le dispatch");

            invoicesPage.ResetFilters();
            invoicesPage.Filter(InvoicesPage.FilterType.Customer, customerName);
            invoicesPage.Filter(InvoicesPage.FilterType.DateFrom, DateTime.Today);
            invoicesPage.ScrollToInvoiceStep();
            invoicesPage.ShowInvoiceStep();
            invoicesPage.FilterInvoiceStep(InvoicesPage.FilterInvoiceStepType.DRAFT, true);
            while (invoicesPage.CheckTotalNumber() > 0)
            {
                invoicesPage.WaitPageLoading();
                invoicesPage.DeleteFirstInvoice();
            }

            //2.Cliquer sur “invoice auto”
            //3.Dans la popup, sélectionner le site et All Deliveries In Selected Period
            //4.Sélectionner le customer ayant le Dispatch en To Invoice.
            //5.Cliquer sur Create
            try
            {
                autoInvoiceCreateModalPage = invoicesPage.AutoInvoiceCreatePage();
                autoInvoiceCreateModalPage.FillField_CreateNewAutoInvoice(customerName, site, CustomerPickMethod.AllDeliveriesInSelectedPeriod);
                invoiceDetails = autoInvoiceCreateModalPage.FillFieldSelectAll();
                autoInvoiceCreateModalPage.CloseNewInvoicePopup();

                //Assert
                Assert.IsTrue(invoicesPage.LinesCountByCustomer(customerName) == 1, "Je ne peux pas refacturer le dispatch de la facture supprimée");
            }
            finally
            {
                //1. Delete invoices
                invoicesPage = homePage.GoToAccounting_InvoicesPage();
                invoicesPage.ResetFilters();
                invoicesPage.Filter(InvoicesPage.FilterType.Customer, customerName);
                invoicesPage.Filter(InvoicesPage.FilterType.DateFrom, DateTime.Today);
                invoicesPage.ScrollToInvoiceStep();
                invoicesPage.ShowInvoiceStep();
                invoicesPage.FilterInvoiceStep(InvoicesPage.FilterInvoiceStepType.DRAFT, true);

                while (invoicesPage.CheckTotalNumber() > 0)
                {
                    invoicesPage.DeleteFirstInvoice();
                }

                //2. Delete Dispatch (unvalidate)
                dispatchPage = homePage.GoToProduction_DispatchPage();
                dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryName);
                dispatchPage.Filter(DispatchPage.FilterType.Site, site);
                dispatchPage.ClickQuantityToInvoice();
                dispatchPage.UnValidateAll();
                dispatchPage.ClickQuantityToProduce();
                dispatchPage.UnValidateAll();
                dispatchPage.ClickPrevisonalQuantity();
                dispatchPage.UnValidateAll();

                //3. Delete Service & delivery
                var deliveryPage = homePage.GoToCustomers_DeliveryPage();
                deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName);
                var deliveryLoadingPage = deliveryPage.ClickOnFirstDelivery();
                deliveryLoadingPage.DeleteService();

                var deliveryGeneralInfo = deliveryLoadingPage.ClickOnGeneralInformation();
                deliveryGeneralInfo.SetActive(false);
                deliveryGeneralInfo.WaitPageLoading(); 
                deliveryGeneralInfo.DeleteDelivery();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_CreateFlightInvoiceAutoSeperate()
        {
            //Prepare
            string flightNumber = new Random().Next().ToString();
            string siteFrom = TestContext.Properties["SiteLP"].ToString();
            string siteTo = TestContext.Properties["SiteToFlightBob"].ToString();
            string aircraft = TestContext.Properties["AircraftBis"].ToString();
            string customerBob = TestContext.Properties["CustomerSchedule"].ToString();
            string customerPickMethod = TestContext.Properties["AllFlightsInSelectedPeriod"].ToString();
            string serviceName = TestContext.Properties["ServiceNameLP"].ToString();
            string customer = TestContext.Properties["Bob_Customer"].ToString();
           
            // Log in
            var homePage = LogInAsAdmin();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
            var flightCreatePageModal = flightPage.FlightCreatePage();

            flightCreatePageModal.FillField_CreatNewFlight(flightNumber, customerBob, aircraft, siteFrom, siteTo);
            var flightDetailsPage = flightPage.EditFirstFlight(flightNumber);
            flightDetailsPage.AddGuestType();
            flightDetailsPage.AddService(serviceName);

            flightPage = homePage.GoToFlights_FlightPage();
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            flightPage.SetNewState("I");

            var invoiceAutoPage = homePage.GoToAccounting_InvoicesPage();
            invoiceAutoPage.ResetFilters();
            var lastInvoiceId = invoiceAutoPage.GetFirstInvoiceId();
            var autoInvoiceCreateModalPage = invoiceAutoPage.AutoInvoiceCreatePage();
            autoInvoiceCreateModalPage.FillField_CreateNewAutoInvoice(customer, siteFrom, customerPickMethod);
            var selectedFlightsNumber = autoInvoiceCreateModalPage.GetSelectedFlightsNumber();
            autoInvoiceCreateModalPage.CheckBoxSeparatedInvoices();
            autoInvoiceCreateModalPage.Submit();

            invoiceAutoPage = homePage.GoToAccounting_InvoicesPage();
            invoiceAutoPage.ResetFilters();
            var newLastInvoiceId = invoiceAutoPage.GetFirstInvoiceId();
            var isSortedById = invoiceAutoPage.IsSortedById();
            if (!isSortedById) invoiceAutoPage.Filter(InvoicesPage.FilterType.SortBy, "Id");
            invoiceAutoPage.WaitPageLoading(); 
               //Assert
               var InvoiceSperatedCheck = invoiceAutoPage.VerifyInvoiceSeperated(lastInvoiceId, newLastInvoiceId, selectedFlightsNumber);
            Assert.IsTrue(InvoiceSperatedCheck, "la facture n'est pas séparée");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_CannotInvoiceFlightsAgain()
        {
            string flightNumber = new Random().Next().ToString();
            string siteFrom = TestContext.Properties["SiteLP"].ToString();
            string siteTo = TestContext.Properties["SiteToFlightBob"].ToString();
            string aircraft = TestContext.Properties["AircraftBis"].ToString();
            string customerBob = TestContext.Properties["CustomerSchedule"].ToString();
            string serviceName = TestContext.Properties["ServiceNameLP"].ToString();
            // Arrange
            var homePage = LogInAsAdmin();

            //Prepare
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
            var flightCreatePageModal = flightPage.FlightCreatePage();
            flightCreatePageModal.FillField_CreatNewFlight(flightNumber, customerBob, aircraft, siteFrom, siteTo);
            var flightDetailsPage = flightPage.EditFirstFlight(flightNumber);
            flightDetailsPage.AddGuestType();
            flightDetailsPage.AddService(serviceName);
            flightPage = homePage.GoToFlights_FlightPage();
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            flightPage.SetNewState("I");

            //1. Cliquer sur “invoice auto”
            //2.Dans la popup, sélectionner le site et All Flights In Selected Period
            //3.Sélectionner le customer ayant le vol déjà facturé.
            // 4.Appuyer sur Next
            var invoicePage = homePage.GoToAccounting_InvoicesPage();
            invoicePage.ResetFilters();

            try
            {
                var autoInvoicePage = invoicePage.AutoInvoiceCreatePage();
                autoInvoicePage.FillField_CreateNewAutoInvoice(customerBob, siteFrom, CustomerPickMethod.AllFlightsInSelectedPeriod);
                autoInvoicePage.FillFieldSelectOneFlight(flightNumber);
                var invoiceDetailsPage = autoInvoicePage.Submit();

                var invoiceId = "";
                if (invoiceDetailsPage.IsOnInvoiceDetailsPage())
                {
                    // un ID si non validé, un number si validé
                    invoiceId = invoiceDetailsPage.GetInvoiceNumber();
                    invoiceDetailsPage.BackToList();
                }
                else
                {
                    //on est dans index à la place d'invoiceDetails
                    invoiceId = invoicePage.GetFirstInvoiceID();
                }

                //Sur le dernier écran de la popup, le vol précédemment facturé n’est plus facturable.
                autoInvoicePage = invoicePage.AutoInvoiceCreatePage();
                try
                {
                    autoInvoicePage.FillField_CreateNewAutoInvoice(customerBob, siteFrom, CustomerPickMethod.AllFlightsInSelectedPeriod);
                }
                catch
                {
                }
                Assert.IsFalse(autoInvoicePage.FillFieldSelectOneFlight(flightNumber), "It is possible to create invoiced flight again.");

            }
            finally
            {
                var invoicesPage = homePage.GoToAccounting_InvoicesPage();
                invoicePage.Filter(InvoicesPage.FilterType.Customer, customerBob);
                invoicePage.DeleteFirstInvoice();

                flightPage = homePage.GoToFlights_FlightPage();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                flightPage.SetNewState("P");
                flightPage.SetNewState("C");
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_CanInvoiceAgainDeletedInvoiceFlight()
        {
            string flightNumber = new Random().Next().ToString();
            string siteFrom = TestContext.Properties["SiteLP"].ToString();
            string siteTo = TestContext.Properties["SiteToFlightBob"].ToString();
            string aircraft = TestContext.Properties["AircraftBis"].ToString();
            string customerBob = TestContext.Properties["CustomerLP"].ToString();
            string serviceName = TestContext.Properties["ServiceNameLP"].ToString();
            // Arrange
            var homePage = LogInAsAdmin();

            //Prepare
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
            var flightCreatePageModal = flightPage.FlightCreatePage();
            flightCreatePageModal.FillField_CreatNewFlight(flightNumber, customerBob, aircraft, siteFrom, siteTo);
            var flightDetailsPage = flightPage.EditFirstFlight(flightNumber);
            flightDetailsPage.AddGuestType();

            flightDetailsPage.AddService(serviceName);
            flightPage = homePage.GoToFlights_FlightPage();
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            flightPage.SetNewState("I");
            var invoicePage = homePage.GoToAccounting_InvoicesPage();

            try
            {
                var autoInvoicePage = invoicePage.AutoInvoiceCreatePage();
                autoInvoicePage.FillField_CreateNewAutoInvoice(customerBob, siteFrom, CustomerPickMethod.AllFlightsInSelectedPeriod);
                autoInvoicePage.FillFieldSelectOneFlight(flightNumber);
                var invoiceDetailsPage = autoInvoicePage.Submit();

                var invoiceId = "";
                if (invoiceDetailsPage.IsOnInvoiceDetailsPage())
                {
                    // un ID si non validé, un number si validé
                    invoiceId = invoiceDetailsPage.GetInvoiceNumber();
                    invoiceDetailsPage.BackToList();
                }
                else
                {
                    //on est dans index à la place d'invoiceDetails
                    invoicePage.ResetFilters();
                    invoiceId = invoicePage.GetFirstInvoiceID();
                }

                //1.Supprimer la facture créée
                invoicePage.ResetFilters();
                invoicePage.WaitPageLoading();
                invoicePage.Filter(InvoicesPage.FilterType.Customer, customerBob);
                invoicePage.Filter(InvoicesPage.FilterType.DateFrom, DateTime.Today);
                invoicePage.ScrollToInvoiceStep();
                invoicePage.ShowInvoiceStep();
                invoicePage.FilterInvoiceStep(InvoicesPage.FilterInvoiceStepType.DRAFT, true);
                while (invoicePage.CheckTotalNumber() > 0)
                {
                    invoicePage.DeleteFirstInvoice();
                }
                autoInvoicePage = invoicePage.AutoInvoiceCreatePage();
                autoInvoicePage.FillField_CreateNewAutoInvoice(customerBob, siteFrom, CustomerPickMethod.AllFlightsInSelectedPeriod);

                Assert.IsTrue(autoInvoicePage.FillFieldSelectOneFlight(flightNumber), "It is not possible to select flight previously invoiced and invoice deleted.");
                invoiceDetailsPage = autoInvoicePage.Submit();

                invoiceId = "";
                if (invoiceDetailsPage.IsOnInvoiceDetailsPage())
                {
                    // un ID si non validé, un number si validé
                    invoiceId = invoiceDetailsPage.GetInvoiceNumber();
                    invoiceDetailsPage.BackToList();
                }
                else
                {
                    //on est dans index à la place d'invoiceDetails
                    invoiceId = invoicePage.GetFirstInvoiceID();
                }

                invoicePage = homePage.GoToAccounting_InvoicesPage();
                invoicePage.ResetFilters();
                Assert.IsTrue(invoiceId.Contains(invoicePage.GetFirstInvoiceId()), "Je ne peux pas refacturer le flight de la facture supprimée.");

            }
            finally
            {
                invoicePage = homePage.GoToAccounting_InvoicesPage();
                invoicePage.ResetFilters();
                invoicePage.Filter(InvoicesPage.FilterType.Customer, customerBob);
                invoicePage.WaitPageLoading();
                invoicePage.DeleteFirstInvoice();

                flightPage = homePage.GoToFlights_FlightPage();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                flightPage.SetNewState("P");
                flightPage.SetNewState("C");
            }
        }

        [TestMethod]
        [Priority(4)]
        [Timeout(_timeout)]
        public void AC_INVO_CreateDataForInvoiceAutoCO()
        {
            DateTime fromDate = DateUtils.Now.AddDays(-10);
            DateTime toDate = DateUtils.Now.AddDays(10);

            // Arrange
            var homePage = LogInAsAdmin();
            ClearCache();

            // Act
            //Création d’un Customer
            var customerPage = homePage.GoToCustomers_CustomerPage();
            customerPage.ResetFilters();
            customerPage.Filter(CustomerPage.FilterType.Search, CUSTOMER);

            if (customerPage.CheckTotalNumber() == 0)
            {
                var creteCustomerModal = customerPage.CustomerCreatePage();
                creteCustomerModal.FillFields_CreateCustomerModalPage(CUSTOMER, CUSTOMERICAO, CUSTOMERTYPE);
                var generalInfo = creteCustomerModal.Create();
                customerPage = generalInfo.GoToCustomers_CustomerPage();
            }
            else
            {
                var customerGeneralInfo = customerPage.SelectFirstCustomer();
                customerGeneralInfo.SetCategory(CUSTOMERTYPE);
                customerPage = customerGeneralInfo.GoToCustomers_CustomerPage(); ;
            }

            // d’un Service, ajout d’un packaging dans le service lié au Customer ou mis à jour de la date

            var servicePage = customerPage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, SERVICENAME);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(SERVICENAME, category: SERVICECATEGORY);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(SITE, CUSTOMER, fromDate, toDate);
                servicePage = pricePage.BackToList();
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, SERVICENAME);
            }
            else
            {
                var pricePage = servicePage.ClickOnFirstService();
                pricePage.ToggleFirstPrice();
                var serviceCreatePriceModalPage = pricePage.EditFirstPrice(SITE, CUSTOMER);
                serviceCreatePriceModalPage.FillFields_EditCustomerPrice(SITE, CUSTOMER, fromDate, toDate);

                var serviceGeneralInformationsPage = pricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                serviceGeneralInformationsPage.SetCategory(SERVICECATEGORY);
                servicePage = serviceGeneralInformationsPage.BackToList();
            }

            // création d'un delivery
            var deliveryPage = servicePage.GoToCustomers_DeliveryPage();
            deliveryPage.ResetFilter();
            deliveryPage.Filter(DeliveryPage.FilterType.Search, DELIVERY);
            deliveryPage.Filter(DeliveryPage.FilterType.Sites, SITE);
            deliveryPage.Filter(DeliveryPage.FilterType.Customers, CUSTOMER);

            if (deliveryPage.CheckTotalNumber() == 0)
            {
                var deliveryCreateModal = deliveryPage.DeliveryCreatePage();
                deliveryCreateModal.FillFields_CreateDeliveryModalPage(DELIVERY, CUSTOMER, SITE, true);
                var deliveryLoadingPage = deliveryCreateModal.Create();
                deliveryPage = deliveryLoadingPage.GoToCustomers_DeliveryPage();
            }

            // création d’un Customer Order (dans Production – Customer Order)
            var customerorderPage = homePage.GoToProduction_CustomerOrderPage();
            var customerOrderCreateModal = customerorderPage.CustomerOrderCreatePage();
            customerOrderCreateModal.FillField_CreatNewCustomerOrderWOUTflight(SITE, CUSTOMER, DELIVERY, DateUtils.Now.AddDays(6));
            var customerOrderDetail = customerOrderCreateModal.Submit();

            // ajout du service dans le CO, lui mettre une quantité non nulle, puis valider le CO.
            customerOrderDetail.AddNewItemWithCategory(SERVICENAME, SERVICEQUANTITY, SERVICECATEGORY);
            customerOrderDetail.ValidateCustomerOrder();
            var customerorderGeneralInfo = customerOrderDetail.ClickOnGeneralInformationTab();
            var orderNumber = customerorderGeneralInfo.GetOrderNumber();

            // Le Customer
            customerPage = customerorderGeneralInfo.GoToCustomers_CustomerPage();
            customerPage.ResetFilters();
            customerPage.Filter(CustomerPage.FilterType.Search, CUSTOMER);
            customerPage.WaitLoading();
            Assert.AreEqual(1, customerPage.CheckTotalNumber(), "Le customer a mal été crée.");

            // le Service

            servicePage = customerOrderDetail.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, SERVICENAME);
            servicePage.Filter(ServicePage.FilterType.Sites, SITE);

            Assert.AreEqual(1, servicePage.CheckTotalNumber(), "Le service a mal été crée.");

            // le packaging

            var pricepage = servicePage.ClickOnFirstService();
            var priceName = $"{SITE} / {CUSTOMER}";
            Assert.AreEqual(priceName, pricepage.GetPriceName(), "Le prix pour le service a mal été crée.");

            // et le Customer Order sont bien créés et à jour.
            customerorderPage = pricepage.GoToProduction_CustomerOrderPage();
            customerorderPage.ResetFilter();
            customerorderPage.Filter(PageObjects.Production.CustomerOrder.CustomerOrderPage.FilterType.Search, orderNumber);
            customerorderPage.Filter(PageObjects.Production.CustomerOrder.CustomerOrderPage.FilterType.From, fromDate);
            customerorderPage.Filter(PageObjects.Production.CustomerOrder.CustomerOrderPage.FilterType.To, toDate);
            Assert.AreEqual(1, customerorderPage.CheckTotalNumber(), "Le customer order a mal été crée et il est bien à jour.");
            //clean invoices
            var invoicesPage = customerorderPage.GoToAccounting_InvoicesPage();
            invoicesPage.Filter(InvoicesPage.FilterType.Customer, CUSTOMER);
            // les validated ne sont pas supprimable
            invoicesPage.ShowInvoiceShow();
            invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowNotValidatedOnly, true);
            if (invoicesPage.CheckTotalNumber() > 0)
            {
                invoicesPage.DeleteFirstInvoice();
            }
        }


        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_CreateInvoiceAutoCO()
        {
            // Arrange
            var homePage = LogInAsAdmin();

            //1.Cliquer sur “invoice auto”
            //2.Dans la popup, sélectionner le site et All Customer Orders In Selected Period
            //3.Sélectionner le customer ayant le Customer Order validé.
            //4.Cliquer sur Create
            //5.Ne pas valider la facture.
            var invoicePage = homePage.GoToAccounting_InvoicesPage();
            try
            {
                var autoInvoicePage = invoicePage.AutoInvoiceCreatePage();
                autoInvoicePage.FillField_CreateNewAutoInvoice(CUSTOMER, SITE, CustomerPickMethod.AllCustomerOrdersInSelectedPeriod);
                var invoiceDetailsPage = autoInvoicePage.FillFieldSelectAll();

                var invoiceId = "";
                if (invoiceDetailsPage.IsOnInvoiceDetailsPage())
                {
                    // un ID si non validé, un number si validé 
                    invoiceId = invoiceDetailsPage.GetInvoiceNumber();
                    invoiceDetailsPage.BackToList();
                }
                else
                {
                    //on est dans index à la place d'invoiceDetails
                    invoiceId = invoicePage.GetFirstInvoiceID();
                }

                // L'Invoice
                Thread.Sleep(2000);
                invoicePage.ResetFilters();
                invoicePage.Filter(InvoicesPage.FilterType.Site, SITE);
                invoicePage.Filter(InvoicesPage.FilterType.Customer, CUSTOMER);
                var invoiceIds = invoicePage.GetAllIds();
                if (!invoiceId.StartsWith("Id "))
                {
                    invoiceId = "Id " + invoiceId;
                }
                Assert.IsTrue(invoiceIds.Contains(invoiceId), "L'Invoice a mal été crée.");

                var invoiceDetailPage = invoicePage.SelectFirstInvoice();
                var unitPrice = invoiceDetailPage.GetUnitPrice();
                var total = invoiceDetailPage.GetTotal();
                var quantity = invoiceDetailPage.GetQuantity();

                Assert.AreEqual(total, unitPrice * quantity, "L'Invoice n'a pas le bon prix.");

                invoicePage = invoiceDetailPage.BackToList();
            }
            finally
            {
                var invoicesPage = homePage.GoToAccounting_InvoicesPage();
                invoicePage.Filter(InvoicesPage.FilterType.Customer, CUSTOMER);
                invoicePage.DeleteFirstInvoice();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_CannotInvoiceCOAgain()
        {
            // Arrange
            var homePage = LogInAsAdmin();

            //1.Cliquer sur “invoice auto”
            //2.Dans la popup, sélectionner le site et All Customer Orders In Selected Period
            //3.Sélectionner le customer ayant le Customer Order validé.
            //4.Cliquer sur Create
            //5.Ne pas valider la facture.
            InvoicesPage invoicePage = homePage.GoToAccounting_InvoicesPage();

            try
            {
                var autoInvoicePage = invoicePage.AutoInvoiceCreatePage();
                autoInvoicePage.FillField_CreateNewAutoInvoice(CUSTOMER, SITE, CustomerPickMethod.AllCustomerOrdersInSelectedPeriod);
                var invoiceDetailsPage = autoInvoicePage.FillFieldSelectAll();

                var invoiceId = "";

                invoicePage = homePage.GoToAccounting_InvoicesPage();
                // un ID si non validé, un number si validé
                invoiceId = invoicePage.GetFirstInvoiceID();
                // L'Invoice
                invoicePage.ResetFilters();
                invoicePage.Filter(InvoicesPage.FilterType.Site, SITE);
                invoicePage.Filter(InvoicesPage.FilterType.Customer, CUSTOMER);
                var invoiceIds = invoicePage.GetAllIds();
                var CreateInvoice = invoiceIds[0].Contains(invoiceId);
                Assert.IsTrue(CreateInvoice, "L'Invoice a mal été crée.");

                autoInvoicePage = invoicePage.AutoInvoiceCreatePage();
                autoInvoicePage.FillField_CreateNewAutoInvoice(CUSTOMER, SITE, CustomerPickMethod.AllCustomerOrdersInSelectedPeriod);
                var IsSelectOrders = autoInvoicePage.IsSelectOrdersEmpty();
                Assert.IsTrue(IsSelectOrders, "It is possible to create invoice customer order again.");
                //TODO
            }
            finally
            {
                var invoicesPage = homePage.GoToAccounting_InvoicesPage();
                invoicePage.Filter(InvoicesPage.FilterType.Customer, CUSTOMER);
                invoicePage.DeleteFirstInvoice();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_Invoices_Export()
        {
            //prepare
            var downloadPath = TestContext.Properties["DownloadsPath"].ToString();

            // Arrange
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string service = TestContext.Properties["InvoiceService"].ToString();
            string customerPickMethod = TestContext.Properties["AllFlightsInSelectedPeriod"].ToString();
            var startDate = DateTime.Today; // Today's date
            var endDate = new DateTime(startDate.Year, startDate.Month, DateTime.DaysInMonth(startDate.Year, startDate.Month));
            var numInvoice = "";
            // Arrange
            var homePage = LogInAsAdmin();
            homePage.ClearDownloads();

            // Flight  part
            CreateFlightForInvoice(homePage, site, customer, service);

            //Invoice part
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            invoicesPage.ResetFilters();
            var invoiceCreateAutoModalpage = invoicesPage.AutoInvoiceCreatePage();
            invoiceCreateAutoModalpage.FillField_CreateNewAutoInvoice(customer, site, customerPickMethod);
            var invoiceDetails = invoiceCreateAutoModalpage.Submit();
            try
            {
                invoiceDetails.Validate();
                numInvoice = invoiceDetails.GetInvoiceNumber();
            }

            catch
            {
                invoiceDetails = invoicesPage.SelectFirstInvoice();
                invoiceDetails.Validate();
                numInvoice = invoiceDetails.GetInvoiceNumber();
            }

            var serviceTotalNumber = invoiceDetails.ServiceTotalNumber();
            InvoiceFooterPage invoiceFooterPage = invoiceDetails.ClickOnInvoiceFooter();
            var TaxNameTotalNumber = invoiceFooterPage.TaxNameTotalNumber();
            invoicesPage = invoiceDetails.BackToList();

            invoicesPage.ResetFilters();
            invoicesPage.Filter(InvoicesPage.FilterType.ByNumber, numInvoice);
            invoicesPage.Filter(InvoicesPage.FilterType.DateFrom, startDate);
            invoicesPage.Filter(InvoicesPage.FilterType.DateTo, endDate);

            // clear download directory 
            foreach (var file in new DirectoryInfo(downloadPath).GetFiles())
            {
                file.Delete();
            }
            // export
            invoicesPage.ClearDownloads();
            invoicesPage.ExportExcelFile();

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = invoicesPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern n'a été téléchargé.");
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadPath, fileName);
            var listInvoicesExcel = OpenXmlExcel.GetValuesInList("Invoice ID", "Invoices", filePath).ToList().Count();
            var TotalDetailsInvoice = serviceTotalNumber + TaxNameTotalNumber;
            Assert.AreEqual(TotalDetailsInvoice, listInvoicesExcel, "Le nombre d'invoices filtrées ne corespond pas au  nombre afficher sur ecran Invoice  ");

        }

        // -------------------utilitaires----------------------------
        private void CreateFlightForInvoice(HomePage homePage, string site, string customer, string service)
        {
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string siteTo = TestContext.Properties["SiteLP"].ToString();
            string flightNumber = new Random().Next().ToString();
            var state = "I";

            // On vérifie que la valeur de newVersionFlight est égale à 2
            //CheckNewVersionFlight(homePage);//old version

            // Create flight
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, site);

            var flightCreateModalpage = flightPage.FlightCreatePage();
            flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, site, siteTo);

            FlightDetailsPage editPage = flightPage.EditFirstFlight(flightNumber);
            editPage.AddGuestType();
            editPage.AddService(service);
            editPage.UnsetState("P");
            editPage.SetNewState(state);
            editPage.CloseViewDetails();

            WebDriver.Navigate().Refresh();

            // On vérifie que le state Invoice est sélectionné
            Assert.IsTrue(flightPage.GetFirstFlightStatus("I"), "Le state 'Invoiced' n'est pas sélectionné pour le flight créé.");
        }

        private void CheckNewVersionFlight(HomePage homePage)
        {
            var applicationSettings = homePage.GoToApplicationSettings();
            var applicationSettingsModalPage = applicationSettings.GetNewFlightGlobalPage();
            applicationSettingsModalPage.SetNewVersionFlightValue(2);
            applicationSettingsModalPage.Save();
        }

        public void CreateNewDispatch(HomePage homePage, string deliveryName, string customerName, string serviceName, string deliveryRoundName = null, bool isWithMenu = true, string menuName = null, string recipeName = null)
        {
            Random rnd = new Random();

            // Prepare delivery                       
            string deliverySite = TestContext.Properties["SiteLP"].ToString();
            string qty = "15";

            var sitePage = homePage.GoToParameters_Sites();
            sitePage.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, deliverySite);
            string siteID = sitePage.CollectNewSiteID();

            // Prepare recipe
            string recipeType = TestContext.Properties["RecipeType"].ToString();
            int nbPortions = new Random().Next(1, 30);
            string recipeVariant = TestContext.Properties["MenuVariantACE3"].ToString(); //ADULTOS 300G - ALMUERZO
            // Prepare menu
            if (menuName == null)
            {
                menuName = "MenuDispatch-" + rnd.Next().ToString();
                recipeName = "RecipeDispatch-" + rnd.Next().ToString();
            }
            string variant = TestContext.Properties["MenuVariantACE3"].ToString();

            if (isWithMenu)
            {
                //2. Creation de la recette
                if (menuName != null)
                {
                    var recipesPage = homePage.GoToMenus_Recipes();
                    var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                    var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());
                    recipeGeneralInfosPage.AddVariantWithSite(deliverySite, recipeVariant);
                    recipeGeneralInfosPage.SelectFirstVariantFromList();
                }

                //3. Creation du menu pour le service
                var menusPage = homePage.GoToMenus_Menus();
                var menusCreateModalPage = menusPage.MenuCreatePage();
                var menuDayViewPage = menusCreateModalPage.FillField_CreateNewMenu(menuName, DateUtils.Now.AddDays(-10), DateUtils.Now.AddDays(+30), deliverySite, variant, serviceName);

                menuDayViewPage.AddRecipe(recipeName);
            }

            //4. Création du delivery avec le service
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
            deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, customerName, deliverySite, true);
            var loadingPage = deliveryCreateModalPage.Create();

            loadingPage.AddService(serviceName);
            loadingPage.AddQuantity(qty);

            //5. Création d'un delivery round
            if (deliveryRoundName != null)
            {
                var deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
                var deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
                var deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRound(deliveryRoundName, siteID, DateUtils.Now, DateUtils.Now.AddDays(+31));

                var deliveryRoundDeliveriesPage = deliveryRoundGeneralInfoPage.ClickOnDeliveryTab();
                deliveryRoundDeliveriesPage.AddDelivery(deliveryName);
                deliveryRoundDeliveriesPage.BackToList();
            }
        }


        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_Invoices_access()
        {
            //arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            var invoicePage = homePage.GoToAccounting_InvoicesPage();
            bool isAccessOK = invoicePage.AccessPage();

            // Assert
            Assert.IsTrue(isAccessOK, "Page inaccessible");
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_Invoices_PDFSummary_Attached_Email()
        {
            var homePage = LogInAsAdmin();
            var invoicePage = homePage.GoToAccounting_InvoicesPage();
            invoicePage.ResetFilters();
            invoicePage.ShowInvoiceStep();
            invoicePage.FilterInvoiceStep(FilterInvoiceStepType.VALIDATED_ACCOUNTED, true);
            var invoicenumber = invoicePage.GetFirstinvoiceNumber().Replace("N°", "").Trim();
            var invoiceDetailsPage = invoicePage.SelectFirstInvoice();
            invoiceDetailsPage.ClickOnSendByEmail();
            string subject = invoiceDetailsPage.GetSubject();
            string extractedNumbers = invoiceDetailsPage.ExtractNumbersFromElement();
            Assert.IsTrue(extractedNumbers == invoicenumber, " le numéros d'invoice sur la piéce joint n'est pas correspondante  ");
            bool isPdfAttachementVisible = invoiceDetailsPage.IsPdfAttachmentVisible();
            Assert.IsTrue(isPdfAttachementVisible, " la pièce jointe PDF summary n'est pas  rattacher dans le mail avec l'Invioce");
            invoiceDetailsPage.Submit();
            invoiceDetailsPage.BackToList();
            invoicePage.ResetFilters();
            invoicePage.Filter(FilterType.ByNumber, invoicenumber);
            var resultat = invoicePage.IsInvoiceSentByMail();
            Assert.IsTrue(resultat, " le mail n'a pas été envoyé");
            MailPage mailPage = invoicePage.RedirectToOutlookMailbox();
            mailPage.FillFields_LogToOutlookMailbox_MoreThanOneMonth("");
            mailPage.ClickOnSpecifiedOutlookMail(subject);
            bool isPieceJointeOK = mailPage.IsOutlookPieceJointeOK("Invoice_" + invoicenumber + ".pdf");
            Assert.IsTrue(isPieceJointeOK, "La pièce jointe n'est pas présente dans le mail.");
            mailPage.DeleteCurrentOutlookMail();
            mailPage.Close();
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_Shifted_Column()
        {
            //arrange
            HomePage homePage = LogInAsAdmin();

            //act
            var invoicePage = homePage.GoToAccounting_InvoicesPage();
            var grossAmounts = invoicePage.GetGrossAmounts();
            var totalInclTaxes = invoicePage.GetTotalInclTaxes();
            var grossAmountsNbr = grossAmounts.Count;
            var totalInclTaxesNbr = totalInclTaxes.Count;
            // Assert
            Assert.AreEqual(grossAmountsNbr, totalInclTaxesNbr,
                "Le nombre d'entrées pour les 'Montants Bruts' doit correspondre au nombre d'entrées pour les 'Total TTC'");

            for (int i = 0; i < grossAmountsNbr - 1; i++)
            {
                int grossElementX = invoicePage.GetGrossAmountXPosition(i + 1);
                int totalInclTaxesElementX = invoicePage.GetTotalInclTaxesXPosition(i + 1);
                
                // Assert
                Assert.IsTrue(grossElementX <= totalInclTaxesElementX, $"Element {i} de 'Gross Amounts' (X: {grossElementX}) devrait être à gauche de ou aligné avec 'Total incl Taxes' (X: {totalInclTaxesElementX})");
            }
        }
        [Timeout(_timeout)]
        [TestMethod]

        public void AC_INVO_Align_Disks()
        {
            //arrange
            var homePage = LogInAsAdmin();
            var invoicePage = homePage.GoToAccounting_InvoicesPage();

            invoicePage.ResetFilters();
            invoicePage.ShowInvoiceStep();
            invoicePage.FilterInvoiceStep(FilterInvoiceStepType.ACCOUNTED, true);
            string invoicenumber = invoicePage.GetFirstinvoiceNumber();
            invoicePage.IsSameDistance();
            Assert.IsTrue(invoicePage.IsSameDistance(), "the disquette element is not aligned");

        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_Print_Invoices()
        {
            //Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            //arrange
            var homePage = LogInAsAdmin();

            //Act
            var invoicePage = homePage.GoToAccounting_InvoicesPage();
            invoicePage.ResetFilters();
            invoicePage.ShowInvoiceStep();
            invoicePage.FilterInvoiceStep(FilterInvoiceStepType.VALIDATED_ACCOUNTED, true);
            var invoiceDetailsPage = invoicePage.SelectFirstInvoice();
            //Information qui doit etre affiche dans le print
            var servicecategory = invoiceDetailsPage.GetServiceCategory();
            var quantity = invoiceDetailsPage.Get_Quantity();
            var unitPrice = invoiceDetailsPage.Get_Unit_Price() / 10.0;
            invoiceDetailsPage.WaitPageLoading();
            //Print
            invoiceDetailsPage.ClearDownloads();
            invoiceDetailsPage.ShowExtendedMenu();
            PrintReportPage printReport = invoiceDetailsPage.PrintInvoiceResults(true);
            Assert.IsTrue(printReport.IsReportGenerated(), "Fichier non généré");
            printReport.Close();
            printReport.Purge(downloadsPath, "Print Invoice_-_", "All_files_");
            string trouve = printReport.PrintAllZip(downloadsPath, "Print Invoice_-_", "All_files_", false);
            FileInfo file = new FileInfo(trouve);
            //Le fichier PDF est généré
            Assert.IsTrue(file.Exists, "Fichier PDF non généré");
            //Vérifier que les données correspondent
            PdfDocument document = PdfDocument.Open(file.FullName);
            List<string> mots = new List<string>();
            foreach (UglyToad.PdfPig.Content.Page p in document.GetPages())
            {
                foreach (var mot in p.GetWords())
                {
                    mots.Add(mot.Text);
                }
            }
            //Assert
            var total_price = (quantity * unitPrice * invoiceDetailsPage.GetTotal_Flight()).ToString("F2").Replace(".", ",");
            string AllmotsInPdf = String.Join(" ", mots).Replace(" ", "").Trim();
            bool Exist_Total = mots.Contains(total_price);
            int count = Regex.Matches(AllmotsInPdf, servicecategory.Replace(" ", "").Trim()).Count;
            Assert.AreEqual(count, 1, "Pas une seule ligne pour le même service il ya deux lignes ou plus et ils n'ont pas était regroupés");
            Assert.IsTrue(Exist_Total, "Pas une seule ligne pour le même service il ya deux lignes ou plus lorsqu'ils ont était regroupés");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_CreateErrorMessage()
        {
            // Prepare
            string customerType = "BC";
            string serviceName = "INVOICE AUTO FLIGHTT" + DateTime.Now.Day + DateTime.Now.Month + DateTime.Now.Hour + DateTime.Now.Minute + DateTime.Now.Second;
            string MsgError = "Invoice cannot be created because validation error(s) have occured :";
            string site = TestContext.Properties["SiteLP"].ToString();
            DateTime fromDate = DateUtils.Now.AddDays(-10);
            DateTime toDate = DateUtils.Now.AddDays(10);

            //flight
            string flightNumber = new Random().Next().ToString();
            string siteTo = TestContext.Properties["SiteToFlightBob"].ToString();
            string aircraft = TestContext.Properties["AircraftBis"].ToString();
            string customerBob = TestContext.Properties["CustomerSchedule"].ToString();
            var homePage = LogInAsAdmin();
            try
            {
                // Création du service
                var servicePage = homePage.GoToCustomers_ServicePage();
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
                servicePage.Filter(ServicePage.FilterType.ShowAll, true);

                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, null, null, null, customerType);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                priceModalPage.FillFields_CustomerPrice(site, customerBob, fromDate, toDate);
                pricePage.UnfoldAll();
                ServiceCreatePriceModalPage serviceEditPriceModal = pricePage.ServiceEitPriceModal(1);
                //Add method Scale
                serviceEditPriceModal.SetMethodScaleWithZeroNbPersonsWithoutSAVE("Scale", 1, "0");
                pricePage = priceModalPage.Save();
                pricePage.UnfoldAll();
                pricePage.ServiceEitPriceModal(1);
                var scaleListt = serviceEditPriceModal.GetAllScalePrice();
                serviceEditPriceModal.Close();
                //search service by Name
                homePage.GoToCustomers_ServicePage();
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
                Assert.IsTrue(servicePage.GetFirstServiceName().Contains(serviceName), "Le service n'a pas été créé.");
                //create new flight 
                var flightPage = homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, site);
                var flightCreatePageModal = flightPage.FlightCreatePage();
                flightCreatePageModal.FillField_CreatNewFlight(flightNumber, customerBob, aircraft, site, siteTo);
                var flightDetailsPage = flightPage.EditFirstFlight(flightNumber);
                flightDetailsPage.AddGuestType(customerType);
                flightDetailsPage.WaitPageLoading();
                flightDetailsPage.AddService(serviceName);

                flightPage = homePage.GoToFlights_FlightPage();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                flightPage.SetNewState("P");
                flightPage.SetNewState("V");
                flightPage.SetNewState("I");
                var invoiceAutoPage = homePage.GoToAccounting_InvoicesPage();
                invoiceAutoPage.ResetFilters();
                var autoInvoiceCreateModalPage = invoiceAutoPage.AutoInvoiceCreatePage();
                autoInvoiceCreateModalPage.FillField_CreateNewAutoInvoice(customerBob, site, CustomerPickMethod.AllFlightsInSelectedPeriod);
                var selectedFlightsNumber = autoInvoiceCreateModalPage.GetSelectedFlightsNumber();
                autoInvoiceCreateModalPage.Submit();
                bool IsMsgErreurAffiche = autoInvoiceCreateModalPage.TestMsgAfficher(MsgError);
                Assert.IsTrue(IsMsgErreurAffiche, "Il faut avoir un msg d'erreur.");
            }
            finally
            {
                //Delete service 
                var servicePage = homePage.GoToCustomers_ServicePage();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
                var serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();
            }

        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_AffichagederniereEtapeACO()
        {
            string site = TestContext.Properties["SiteACE"].ToString();
            string customer = TestContext.Properties["Bob_Customer"].ToString();
            string customerPickMethod = TestContext.Properties["AllFlightsInSelectedPeriod"].ToString();
            DateTime dateFrom = DateTime.ParseExact("01/09/2024", "dd/MM/yyyy", CultureInfo.InvariantCulture);

            var homePage = LogInAsAdmin();
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            invoicesPage.ResetFilters();

            var invoiceCreateAutoModalpage = invoicesPage.AutoInvoiceCreatePage();
            invoiceCreateAutoModalpage.FillField_CreateNewAutoInvoice(customer, site, customerPickMethod, dateFrom);
            var invoiceDetails = invoiceCreateAutoModalpage.Submit();

            // Vérifier la présence du titre de la pop-up
            var displayed = invoiceDetails.IsDisplayed();
            Assert.IsTrue(displayed, "Le titre 'New Invoice' n'est pas visible.");

            // Vérifier que les éléments ne se chevauchent pas
            Assert.IsTrue(invoiceDetails.AreElementsNotOverlapping(), "Les éléments se chevauchent.");

            // Vérifier que le bouton "OK" est bien positionné et visible
            var okButton = WebDriver.FindElement(By.XPath(OK_BTN));
            Assert.IsTrue(okButton.Displayed, "Le bouton 'OK' n'est pas visible.");
            Assert.IsTrue(okButton.Enabled, "Le bouton 'OK' n'est pas cliquable.");

            // Cliquer sur le bouton OK
            okButton.Click();
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_AffichageURLBacktoListe()
        {

            string site = TestContext.Properties["SiteBCN"].ToString();
            //string customerName = "595442398 - TestCustomer_COInvoiceAuto";
            DateTime dateFrom = DateTime.ParseExact("01/06/2024", "dd/MM/yyyy", CultureInfo.InvariantCulture);

            var homePage = LogInAsAdmin();

            var invoicePage = homePage.GoToAccounting_InvoicesPage();
            string customerName = invoicePage.IsDev() ? "595442398 - TestCustomer_COInvoiceAuto" : "290378940 - TestCustomer_COInvoiceAuto";
            invoicePage.ResetFilters();
            try
            {
                var autoInvoiceCreateModalPage = invoicePage.AutoInvoiceCreatePage();
                autoInvoiceCreateModalPage.FillField_CreateNewAutoInvoice(customerName, site, CustomerPickMethod.AllCustomerOrdersInSelectedPeriod, dateFrom);
                var invoiceDetails = autoInvoiceCreateModalPage.FillFieldSelectSomes(1);
                invoiceDetails.BackToList();
                var invoiceDetailPage = invoicePage.SelectFirstInvoice();

                // Verification: Check that the URL contains 'OrderInvoice'
                string currentUrl = WebDriver.Url;
                Assert.IsTrue(currentUrl.Contains("OrderInvoice"), "The URL does not contain 'OrderInvoice' as expected.");

                // Click on the "BackList" button
                invoiceDetailPage.BackToList();

                // Verification: Check that the URL contains '?isBack=True'
                string backUrl = WebDriver.Url;
                Assert.IsTrue(backUrl.Contains("Invoice?isBack=True"), "The URL does not contain '?isBack=True' after clicking the back button.");

            }
            finally
            {
                var invoicesPage = homePage.GoToAccounting_InvoicesPage();
                invoicePage.Filter(InvoicesPage.FilterType.Customer, customerName);
                invoicePage.DeleteFirstInvoice();
            }

        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_SendByEmailContactInvoiceManager()
        {

            string userEmail = TestContext.Properties["Admin_UserName"].ToString();
            string site = TestContext.Properties["SiteBCN"].ToString();
            DateTime dateFrom = DateTime.ParseExact("01/06/2024", "dd/MM/yyyy", CultureInfo.InvariantCulture);

            var homePage = LogInAsAdmin();
            var sitePage = homePage.GoToParameters_Sites();
            sitePage.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, site);
            sitePage.ClickOnFirstSite();
            sitePage.ClickToContacts();
            sitePage.SetInvoiceManager("test", userEmail);
            homePage.Navigate();
            var invoicePage = homePage.GoToAccounting_InvoicesPage();
            string customerName = invoicePage.IsDev() ? "EMIRATES AIRLINES" : "TRE3 TRAVEL RETAIL SL ANE";
            invoicePage.ResetFilters();

            invoicePage.ShowInvoiceStep();
            invoicePage.FilterInvoiceStep(InvoicesPage.FilterInvoiceStepType.VALIDATED_ACCOUNTED, true);
            invoicePage.Filter(InvoicesPage.FilterType.Site, site);
            invoicePage.Filter(InvoicesPage.FilterType.Customer, customerName);
            invoicePage.Filter(InvoicesPage.FilterType.DateFrom, DateUtils.Now.AddDays(-14));
            invoicePage.Filter(InvoicesPage.FilterType.DateTo, DateUtils.Now.AddDays(+15));
            var invoiceDetailPage = invoicePage.SelectFirstInvoice();
            var invoiceNumber = invoiceDetailPage.GetInvoiceNumber();
            var invoicesPage = invoiceDetailPage.BackToList();
            invoicePage.Filter(InvoicesPage.FilterType.ByNumber, invoiceNumber);
            var invoice = invoicePage.GetInvoiceNumber();

            invoicePage.SendByMail();
            MailPage mailPage = invoicePage.RedirectToOutlookMailbox();
            mailPage.FillFields_LogToOutlookMailbox_MoreThanOneMonth(userEmail);
            bool isMailSent = mailPage.CheckIfSpecifiedOutlookMailExist("Winrest - 1 Invoice" + " - " + site + " - " + invoice);

            Assert.IsTrue(isMailSent, "La réclamation n'a pas été envoyée par mail.");
            mailPage.DeleteCurrentOutlookMail();
            mailPage.Close();
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_Customs_report_Export()
        {
            string site = TestContext.Properties["SiteBCN"].ToString();
            var downloadPath = TestContext.Properties["DownloadsPath"].ToString();
            DateTime date = DateTime.Now.AddMonths(-1);
            string month = "September";
            //arrange
            HomePage homePage = LogInAsAdmin();
            //act
            var invoicesPage = homePage.GoToAccounting_InvoicesPage();
            invoicesPage.ResetFilters();
            invoicesPage.Filter(InvoicesPage.FilterType.DateFrom, DateTime.Now.AddMonths(-1));
            DeleteAllFileDownload();
            invoicesPage.ClearDownloads();
            invoicesPage.ExportCustomFile(downloadPath, site, month);
            invoicesPage.WaitLoading();
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();
            var correctDownloadedFile = invoicesPage.GetExportExcelFileInvoiceCustoms(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Le fichier exporté est introuvable.");
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadPath, fileName);
            int resultNumber = OpenXmlExcel.GetExportResultNumber("Customs", filePath);

            //Assert
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void AC_INVO_FilterShow_Showflightinvoices()
        {

            // Arrange
            HomePage homePage = LogInAsAdmin();

            InvoicesPage invoicesPage = homePage.GoToAccounting_InvoicesPage();
            invoicesPage.ResetFilters();

            // Application des filtres et affichage des factures.
            invoicesPage.Filter(InvoicesPage.FilterType.ByNumber, AutoInvoiceNumber);
            invoicesPage.ShowInvoiceShow();
            invoicesPage.FilterShow(InvoicesPage.FilterShowType.ShowFlightInvoice, true);

            // Récupération du premier numéro de facture affiché.
            string firstInvoicedNumber = invoicesPage.GetFirstInvoiceOnlyNumber();

            //Assert
            Assert.AreEqual(AutoInvoiceNumber, firstInvoicedNumber, "La facture nouvellement crée n'apparait pas.");

        }

    }
}
