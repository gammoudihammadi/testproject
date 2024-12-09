using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Wordprocessing;
using GemBox.Email;
using GemBox.Email.Imap;
using GemBox.Email.Smtp;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Accounting.Invoice;
using Newrest.Winrest.FunctionalTests.PageObjects.CustomerPortal;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Customer;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Delivery;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Reinvoice;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.CustomerOrder;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.Dispatch;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Database;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Policy;
using System.Text.RegularExpressions;
using System.Threading;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using static Newrest.Winrest.FunctionalTests.PageObjects.CustomerPortal.CustomerPortalQtyByWeekPage;

namespace Newrest.Winrest.FunctionalTests.CustomerPortal
{
    [TestClass]
    public class CustomerPortalTests : TestBase
    {
        private const int _timeout = 600000;
        private static string _deliceryNamOfCreateCustomerOrder = "";

        [Priority(1)]
		[TestMethod]
        [Timeout(_timeout)]
        public void CUPO_AddUserForCustomerPortal()
        {
            // Prepare
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();

            // boolean d'ouverture des nouveaux onglets
            bool mailboxTab = false;
            bool updatePasswordTab = false;

            var mailPage = new MailPage(WebDriver, TestContext);
            var updatePage = new CustomerPortalUpdatePage(WebDriver, TestContext);

            // boolean d'accès aux pages //à reconfigurer
            bool quantityByWeek = true;
            bool customerOrder = true;
            bool quantityByDay = true;

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var settingsPortalPage = homePage.GoToParameters_PortalPage();

            // si le user n'existe pas, on le créé
            if (!settingsPortalPage.IsExistingUser(userMail))
            {
                settingsPortalPage.AddNewUser(userMail);
                settingsPortalPage.SearchAndSelectUser(userMail);
                //settingsPortalPage.ConfigureUser(quantityByWeek, customerOrder, quantityByDay);

                try
                {
                    // Connexion à la mailbox pour voir si le mail a été envoyé
                    mailPage = settingsPortalPage.RedirectToOutlookMailbox();
                    mailboxTab = true;

                    mailPage.FillFields_LogToOutlookMailbox(userMail);

                    // Recherche du mail envoyé pour le resend password
                    mailPage.ClickOnSpecifiedOutlookMail("Bienvenido a nuestro portal cliente");

                    updatePage = mailPage.ClickOnLinkForPassword("aquí");
                    updatePasswordTab = true;

                    var customerPortalLoginPage = updatePage.UpdatePassword(userPassword);
                    customerPortalLoginPage.Close();
                    updatePasswordTab = false;

                    mailPage.DeleteCurrentOutlookMail();
                    mailPage.Close();
                    mailboxTab = false;

                }
                finally
                {
                    // On ferme les onglets ouverts
                    if (updatePasswordTab)
                    {
                        updatePage.Close();
                    }

                    if (mailboxTab)
                    {
                        mailPage.Close();
                    }
                }
            }
            else
            {
                settingsPortalPage.SearchAndSelectUser(userMail);
                //à reconfigurer
                //settingsPortalPage.ConfigureUser(quantityByWeek, customerOrder, quantityByDay);
            }

            settingsPortalPage.SetLanguage("English");
            settingsPortalPage.SaveAndConfirm();
            settingsPortalPage = homePage.GoToParameters_PortalPage();
            Assert.IsTrue(settingsPortalPage.IsExistingUser(userMail), "Le user n'est pas présent pour accéder au customer portal.");
        }

        [Priority(2)]
		[TestMethod]
        [Timeout(_timeout)]
        public void CUPO_CreateServiceForCustomerPortal()
        {
            // Prepare
            string customer = TestContext.Properties["CustomerDelivery"].ToString();
            string site = TestContext.Properties["Site"].ToString(); // MAD

            string serviceName = TestContext.Properties["ServiceCustomerPortal"].ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);

            DateTime dateFrom = DateUtils.Now.AddDays(-7);
            DateTime dateTo = DateUtils.Now.AddDays(+15);

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var servicesPage = homePage.GoToCustomers_ServicePage();
            servicesPage.ResetFilters();
            servicesPage.Filter(ServicePage.FilterType.Sites, site);
            servicesPage.Filter(ServicePage.FilterType.Search, serviceName);

            if (servicesPage.CheckTotalNumber() < 2)
            {
                string ext = servicesPage.CheckTotalNumber() == 0 ? "" : servicesPage.CheckTotalNumber().ToString();
                var serviceModalPage = servicesPage.ServiceCreatePage();
                serviceModalPage.FillFields_CreateServiceModalPage(serviceName + ext, serviceCode, serviceProduction);
                var serviceGeneralInformationsPage = serviceModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, dateFrom, dateTo);
                servicesPage = pricePage.BackToList();
            }
            else
            {
                var pricePage = servicesPage.ClickOnFirstService();
                pricePage.SearchPriceForCustomer(customer, site, dateFrom, dateTo);
                servicesPage = pricePage.BackToList();
                servicesPage.Filter(ServicePage.FilterType.Sites, site);
                servicesPage.Filter(ServicePage.FilterType.Search, serviceName + "1");
                pricePage = servicesPage.ClickOnFirstService();
                pricePage.SearchPriceForCustomer(customer, site, dateFrom, dateTo);
                servicesPage = pricePage.BackToList();
            }

            servicesPage.ResetFilters();
            servicesPage.Filter(ServicePage.FilterType.Sites, site);
            servicesPage.Filter(ServicePage.FilterType.Search, serviceName);
            Assert.IsTrue(servicesPage.GetFirstServiceName().Contains(serviceName), MessageErreur.FILTRE_ERRONE, "Search");

            servicesPage.Filter(ServicePage.FilterType.ShowServicesWithExpiredPrice, true);
            Assert.AreEqual(0, servicesPage.CheckTotalNumber(), "Le service apparaît dans la liste des services expirés.");
        }



        // Créé un delivery associé à l'utilisateur via les settings du portal
        [Priority(5)]
		[TestMethod]
        [Timeout(_timeout)]
        public void CUPO_CreateDeliveryForCustomerPortal()
        {
            string customer = TestContext.Properties["CustomerDelivery"].ToString();
            string site = TestContext.Properties["Site"].ToString(); // MAD
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();


            // Prepare for delivery
            string deliveryName = TestContext.Properties["DeliveryCustomerPortal"].ToString();
            string serviceName = TestContext.Properties["ServiceCustomerPortal"].ToString();
            string qty = "15";

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var loadingPage = new DeliveryLoadingPage(WebDriver, TestContext);

            // Create delivery
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.ResetFilter();
            deliveryPage.Filter(DeliveryPage.FilterType.Sites, site);
            deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName);

            if (deliveryPage.CheckTotalNumber() == 0)
            {
                var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, customer, site, true);
                var deliveryLoadingPage = deliveryCreateModalPage.Create();
                deliveryLoadingPage.AddService(serviceName);
                deliveryLoadingPage.AddQuantity(qty);
            }
            else
            {
                loadingPage = deliveryPage.ClickOnFirstDelivery();
                loadingPage.AddQuantity(qty);
            }

            deliveryPage = loadingPage.BackToList();
            deliveryPage.ResetFilter();
            deliveryPage.Filter(DeliveryPage.FilterType.Sites, site);
            deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName);

            Assert.AreEqual(deliveryName, deliveryPage.GetFirstDeliveryName(), "Le delivery " + deliveryName + " n'a pas été trouvé dans la liste des delivery.");


            // Link delivery to user
            var portalSettingsPage = homePage.GoToParameters_PortalPage();
            portalSettingsPage.SearchAndSelectUser(userMail);
            string isDeliveryLinked = portalSettingsPage.LinkDeliveryToUser(deliveryName);

            Assert.AreEqual("true", isDeliveryLinked, "Le delivery n'est pas associé à l'utilisateur du customer portal.");

            //ajouter un deuxième delivery A.F.A. CASAS

            string deliveryNameAFA = "A.F.A. CASAS";
            deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.ResetFilter();
            deliveryPage.Filter(DeliveryPage.FilterType.Sites, site);
            deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryNameAFA);

            if (deliveryPage.CheckTotalNumber() == 0)
            {
                var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryNameAFA, customer, site, true);
                var deliveryLoadingPage = deliveryCreateModalPage.Create();
                deliveryLoadingPage.AddService(serviceName);
                deliveryLoadingPage.AddQuantity(qty);
            }
            else
            {
                loadingPage = deliveryPage.ClickOnFirstDelivery();
                loadingPage.AddQuantity(qty);
            }


            // Link delivery to user
            portalSettingsPage = homePage.GoToParameters_PortalPage();
            portalSettingsPage.SearchAndSelectUser(userMail);
            isDeliveryLinked = portalSettingsPage.LinkDeliveryToUser(deliveryNameAFA);
            Assert.AreEqual("true", isDeliveryLinked, "Le delivery AFA n'est pas associé à l'utilisateur du customer portal.");
            isDeliveryLinked = portalSettingsPage.LinkDeliveryToUser(deliveryName);
            Assert.AreEqual("true", isDeliveryLinked, "Le delivery DeliveryForCustomerPortal n'est pas associé à l'utilisateur du customer portal.");



            // Aller vérifier sur customer Portal, se logguer > Quantity By Week > Verifier dans le filtre des delivery, si les delivery ajoutées sont bien présentes
            homePage.Navigate();
            var customerPortalLoginPage = homePage.GoToCustomerPortal();
            try
            {
                customerPortalLoginPage.LogIn(userMail, userPassword);

                bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
                bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();

                Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
                Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page Quantity By Week.");

                var customerPortalQtyByWeekPage = customerPortalLoginPage.ClickOnQtyByWeek();
                customerPortalQtyByWeekPage.clickDeliveryFilter();
                Assert.AreEqual(3, customerPortalQtyByWeekPage.FilterDeliveryCount(), "liste de checkbox avec 3 éléments");
                List<string> labels = new List<string>();
                labels.Add(customerPortalQtyByWeekPage.FilterDeliveryLabel(1));
                labels.Add(customerPortalQtyByWeekPage.FilterDeliveryLabel(2));
                labels.Add(customerPortalQtyByWeekPage.FilterDeliveryLabel(3));
                Assert.IsTrue(labels.Contains(deliveryNameAFA), "A.F.A. non trouvé dans la liste checkbox DELIVERY");

                if (isUserLogged)
                {
                    customerPortalLoginPage.LogOut();
                }
            }
            finally
            {
                customerPortalLoginPage.Close();
            }
        }



        [Priority(3)]
		[TestMethod]
        [Timeout(_timeout)]
        public void CUPO_CO_AddSitesForCustomerPortal()
        {
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            string siteACE = TestContext.Properties["SiteACE"].ToString();
            string siteBCN = TestContext.Properties["SiteBCN"].ToString();
            string Lanzarotecustomer = TestContext.Properties["CustomerDelivery"].ToString();

            //1.Sur Winrest, aller dans Parameters, puis Portal
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var portalPage = homePage.GoToParameters_PortalPage();

            //2.Chercher votre USER
            portalPage.SearchAndSelectUser(userMail);

            //3.Cliquer sur le sous onglet General Information
            //déjà ouvert

            //4.Vérifier si les sites sont cochés ou les cocher et save(ACE &BCN&MAD)
            Assert.AreEqual("ACE - ACE", portalPage.SearchAndSelectCustomerSite(true, siteACE, 1, true), "Selection raté pour ACE");
            Assert.AreEqual("BCN - BCN", portalPage.SearchAndSelectCustomerSite(false, siteBCN, 5, true), "Selection raté pour BCN");
            Assert.AreEqual("MAD - MAD", portalPage.SearchAndSelectCustomerSite(false, "MAD", 5, true), "Selection raté pour MAD");

            //5.Vérifier si les customers sont cochés ou les cocher et save(EMIRATES AIRLINES &TRE3 TRAVEL RETAIL SL ANE)
            Assert.AreEqual("EMIRATES AIRLINES", portalPage.SearchAndSelectCustomer(true, "EMIRATES AIRLINES", 469, true), "Selection raté pour EMIRATES AIRLINES");
            Assert.AreEqual("TRE3 TRAVEL RETAIL SL ANE", portalPage.SearchAndSelectCustomer(false, "TRE3 TRAVEL RETAIL SL ANE", 646, true), "Selection raté pour TRE3 TRAVEL RETAIL SL ANE");
            Assert.AreEqual(Lanzarotecustomer, portalPage.SearchAndSelectCustomer(false, Lanzarotecustomer, 646, true), "Selection raté pour " + Lanzarotecustomer);
            // Save
            portalPage.SaveAndConfirm();
        }



        [Priority(4)]
		[TestMethod]
        [Timeout(_timeout)]
        public void CUPO_CO_AddCustomersForCustomerPortal()
        {

            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();
            string Lanzarotecustomer = TestContext.Properties["CustomerDelivery"].ToString();

            //1.Sur Winrest, aller dans Parameters, puis Portal
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var portalPage = homePage.GoToParameters_PortalPage();

            //2.Chercher votre USER
            portalPage.SearchAndSelectUser(userMail);

            //3.Cliquer sur le sous onglet General Information (déjà ouvert)
            //déjà ouvert

            //4.Vérifier si les customers sont cochés ou les cocher et save(EMIRATES AIRLINES &TRE3 TRAVEL RETAIL SL ANE)
            Assert.AreEqual("EMIRATES AIRLINES", portalPage.SearchAndSelectCustomer(true, "EMIRATES AIRLINES", 469, true), "Selection raté pour EMIRATES AIRLINES");
            Assert.AreEqual("TRE3 TRAVEL RETAIL SL ANE", portalPage.SearchAndSelectCustomer(false, "TRE3 TRAVEL RETAIL SL ANE", 646, true), "Selection raté pour TRE3 TRAVEL RETAIL SL ANE");
            Assert.AreEqual(Lanzarotecustomer, portalPage.SearchAndSelectCustomer(false, Lanzarotecustomer, 646, true), "Selection raté pour " + Lanzarotecustomer);

            // Save
            portalPage.SaveAndConfirm();

            // go back to main page
            homePage.Navigate();

            //5.Aller sur Customer Portal, se logguer, aller sur Customer Orders
            // Go to CustomerPortal Page
            var customerPortalLoginPage = homePage.GoToCustomerPortal();

            try
            {
                customerPortalLoginPage.LogIn(userMail, userPassword);

                bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
                bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();

                Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
                Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page.");

                var customerPortalCustomerOrdersPage = customerPortalLoginPage.ClickOnCustomerOrders();

                //6.Vérifier dans le filtre 'CUSTOMERS' si on retrouve les customers sélectionnés sur Winrest(EMIRATES AIRLINES &TRE3 TRAVEL RETAIL SL ANE)
                customerPortalCustomerOrdersPage.ShowFilterCustomersContent();

                customerPortalCustomerOrdersPage.FilterCustomersUnCheckAll();
                Assert.AreEqual("SHANE - TRE3 TRAVEL RETAIL SL ANE", customerPortalCustomerOrdersPage.FilterCustomerSelect("SHANE - TRE3 TRAVEL RETAIL SL ANE"), "filtre CUSTOMERS sans TRE3 TRAVEL RETAIL SL ANE");
                Assert.AreEqual("UAE - EMIRATES AIRLINES", customerPortalCustomerOrdersPage.FilterCustomerSelect("UAE - EMIRATES AIRLINES"), "filtre CUSTOMERS sans UAE - EMIRATES AIRLINES");
                Assert.AreEqual("366 - A.F.A. LANZAROTE", customerPortalCustomerOrdersPage.FilterCustomerSelect("366 - A.F.A. LANZAROTE"), "filtre CUSTOMERS sans 366 - A.F.A. LANZAROTE");

                if (isUserLogged)
                {
                    customerPortalLoginPage.LogOut();
                }
            }
            finally
            {
                customerPortalLoginPage.Close();
            }
        }

        [Priority(6)]
		[TestMethod]
        [Timeout(_timeout)]
        public void CUPO_CO_CheckWinrestCustomerOrderCreation()
        {
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();

            //1.Sur Winrest, aller dans Production puis Customer orders et créer un nouvel customer order(cf PR_CO_CreateNewCustomerOrder)
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();
            // Act
            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
            customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customerName, aircraft);
            var deliveryFlightNumber = customerOrderCreateModalPage.GetFlightNumber();
            DateTime orderDateHour = DateUtils.Now.Date;
            var customerOrderDetailPage = customerOrderCreateModalPage.Submit();

            var generalInfo = customerOrderDetailPage.ClickOnGeneralInformationTab();
            var customerOrderNumber = generalInfo.GetOrderNumber();
            var price = generalInfo.GetPrice();
            customerOrderDetailPage.BackToList();

            var name = customerOrderPage.GetCustomerOrderNumber();

            //Assert
            Assert.IsTrue(name.Contains(customerOrderNumber), "Le customer order n'a pas été créé.");
            Assert.IsNotNull(customerName);
            Assert.IsNotNull(deliveryFlightNumber);
            Assert.IsNotNull(orderDateHour);
            Assert.IsNotNull(price);


            //2.Aller sur Customer Portal, se logguer, aller sur Customer Orders
            homePage.Navigate();
            var customerPortalLoginPage = homePage.GoToCustomerPortal();
            try
            {
                customerPortalLoginPage.LogIn(userMail, userPassword);

                bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
                bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();

                Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
                Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page.");

                var customerPortalCustomerOrdersPage = customerPortalLoginPage.ClickOnCustomerOrders();
                string customerOrderNumberId = Regex.Match(customerOrderNumber, @"\d+").Value;
                string customerPortalCustomerOrderId = Regex.Match(customerPortalCustomerOrdersPage.getTableFirstLine(1), @"\d+").Value;
                //3.Vérifier dans la table si on retrouve le customer order, créé sur Winrest(vérifier order no, customer name, delivery flight number, order date hour, pirce)
                Assert.AreEqual(customerOrderNumberId, customerPortalCustomerOrderId, "Première colonne ne contient pas " + customerOrderNumber);
                Assert.AreEqual(customerName, customerPortalCustomerOrdersPage.getTableFirstLine(2), "Seconde colonne ne contient pas " + customerName);
                Assert.AreEqual(deliveryFlightNumber, customerPortalCustomerOrdersPage.getTableFirstLine(3), "Troisième colonne ne contient pas " + deliveryFlightNumber);
                Assert.AreEqual(orderDateHour, DateTime.ParseExact(customerPortalCustomerOrdersPage.getTableFirstLine(4), "dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture), "Quatrième colonne ne contient pas " + orderDateHour);
                Assert.AreEqual(price, customerPortalCustomerOrdersPage.getTableFirstLine(5), "Cinquième colonne ne contient pas " + price);


                if (isUserLogged)
                {
                    customerPortalLoginPage.LogOut();
                }
            }
            finally
            {
                customerPortalLoginPage.Close();
            }
        }


        [Priority(7)]
		[TestMethod]
        [Timeout(_timeout)]
        public void CUPO_CO_CreateNewCustomerOrder()
        {
            //(test identique à PR_CO_CreateNewCustomerOrder mais sur Customer Portal)

            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();

            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string deliveryName = new Random().Next().ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();

            // Kept delivery name to filter for next tests.
            _deliceryNamOfCreateCustomerOrder = deliveryName;

            // Arrange
            CustomerPortalLoginPage customerPortalLoginPage = LogInCustomerPortal();
            customerPortalLoginPage.LogIn(userMail, userPassword);

            bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
            bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();

            Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
            Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page.");

            var customerPortalCustomerOrdersPage = customerPortalLoginPage.ClickOnCustomerOrders();

            //2.Cliquer sur "+"
            var customerPortalCustmerOrdersPageModal = customerPortalCustomerOrdersPage.ClickAdd();


            //2.Remplir le formulaire avec le site(BCN), le customer(EMIRATES AILRINES) et l'aircraft (B777)
            customerPortalCustmerOrdersPageModal.SelectSite(site);

            customerPortalCustmerOrdersPageModal.SelectCustomer(customer);

            customerPortalCustmerOrdersPageModal.SelectDeliveryName(deliveryName);

            customerPortalCustmerOrdersPageModal.SelectAircraft(aircraft);


            //3.Cliquer sur "create"
            CustomerPortalCustomerOrdersPageResult customerPortalCustomerOrdersPageResult = customerPortalCustmerOrdersPageModal.Create();

            //4.Récupérer le numéro du customer order crée, faire Back to List
            //result customer order
            var CustomerOrderText = customerPortalCustomerOrdersPageResult.ResultCustomerOrder();

            // back to list
            customerPortalCustomerOrdersPage.Navigate("/CustomerPortal/CustomerOrders");

            string customerOrderNumberId = Regex.Match(CustomerOrderText, @"\d+").Value;
            string customerPortalCustomerOrderId = Regex.Match(customerPortalCustomerOrdersPage.getTableFirstLine(1), @"\d+").Value;
            //5.Vérifier dans la table si on retrouve le customer order créé
            Assert.AreEqual(customerOrderNumberId, customerPortalCustomerOrderId);

            if (isUserLogged)
            {
                customerPortalLoginPage.LogOut();
            }

        }



        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_ConnectionToCustomerPortal()
        {
            // Prepare
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Go to CustomerPortal Page
            var customerPortalLoginPage = homePage.GoToCustomerPortal();

            try
            {
                customerPortalLoginPage.LogIn(userMail, userPassword);

                bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
                bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();

                if (isUserLogged)
                {
                    customerPortalLoginPage.LogOut();
                }

                Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
                Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page.");
            }
            finally
            {
                customerPortalLoginPage.Close();
            }
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_ResendMailForPassword()
        {
            // Prepare
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["MailBox_UserPassword"].ToString();
            var passwordPortal = TestContext.Properties["CustomerPortal_UserPassword"].ToString();

            // boolean d'ouverture des nouveaux onglets
            bool customerPortalTab = false;
            bool mailboxTab = false;
            bool updatePasswordTab = false;

            var customerPortalLoginPage = new CustomerPortalLoginPage(WebDriver, TestContext);
            var mailPage = new MailPage(WebDriver, TestContext);
            var customerPortalQtyByWeekPage = new CustomerPortalQtyByWeekPage(WebDriver, TestContext);

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            try
            {
                // Connexion au customerPortal via Winrest
                customerPortalLoginPage = homePage.GoToCustomerPortal();
                customerPortalTab = true;

                // Demande de renvoi du password par mail
                customerPortalLoginPage.ResendMailForPassword(userMail);

                // Connexion à la mailbox pour voir si le mail a été envoyé
                mailPage = customerPortalLoginPage.RedirectToOutlookMailbox();
                mailboxTab = true;
                mailPage.WaitPageLoading();
                mailPage.FillFields_LogToOutlookMailbox(userMail);

                // Recherche du mail envoyé pour le resend password
                mailPage.ClickOnSpecifiedOutlookMail("Bienvenido a nuestro portal cliente");

                var updatePage = mailPage.ClickOnLinkForPassword("aquí");
                updatePasswordTab = true;

                customerPortalQtyByWeekPage = updatePage.UpdatePassword(passwordPortal);

                bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
                bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();

                if (isUserLogged)
                {
                    customerPortalLoginPage.LogOut();
                }

                // Retour à l'onglet mail pour supprimer le mail reçu
                customerPortalQtyByWeekPage.Close();
                updatePasswordTab = false;

                mailPage.DeleteCurrentOutlookMail();
                mailPage.Close();
                mailboxTab = false;

                // On ferme l'onget CustomerPortal
                customerPortalLoginPage.Close();
                customerPortalTab = false;

                // Assert
                Assert.IsTrue(isUserLogged, "La connexion au customerPortal après la modification du password a échoué.");
                Assert.IsTrue(isPageLoaded, "Le chargement de la page après la modification du password a échoué.");

            }
            finally
            {
                // On ferme les onglets ouverts
                if (updatePasswordTab)
                {
                    customerPortalQtyByWeekPage.Close();
                }

                if (mailboxTab)
                {
                    mailPage.Close();
                }

                if (customerPortalTab)
                {
                    customerPortalLoginPage.Close();
                }
            }
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_QBW_AddAndValidateQuantities()
        {
            // Prepare
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();
            string site = TestContext.Properties["Site"].ToString();
            var delivery = "DeliveryForCustomerPortal";
            var demain = DateTime.Now.AddDays(1);
            var qty = "94";
            var quantityBeforeUpdate = string.Empty;


            var homePage = LogInAsAdmin();

            // Vérifier que le dispatch existe et invalider toutes les quantités

            #region set parametres to default of deliveryForCustomerPortal
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.Filter(DeliveryPage.FilterType.Search, delivery);
            var loadingTab = deliveryPage.ClickOnFirstDelivery();
            var serviceName = loadingTab.GetServiceName();
            var generalInfo = loadingTab.ClickOnGeneralInformation();
            generalInfo.SetToDefault();
            #endregion

            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.ResetFilter();
            dispatchPage.Filter(DispatchPage.FilterType.Site, site);
            dispatchPage.Filter(DispatchPage.FilterType.Search, serviceName);
            dispatchPage.Filter(DispatchPage.FilterType.Deliveries, delivery);

            //Vérifier et dévalider quantités
            dispatchPage.ClickQuantityToInvoice();
            dispatchPage.UnValidateAll();

            dispatchPage.ClickQuantityToProduce();
            dispatchPage.UnValidateAll();

            // On invalide les quantités et on valide les quantités de la semaine en cours
            dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.UnValidateAll();
            dispatchPage.ValidateAll();

            var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            var IsDispatchValidated = previsionnalQty.IsDispatchValidated();
            Assert.IsTrue(IsDispatchValidated, "Les previsionnal qty du dispatch sont déjà validées pour la semaine prochaine.");
            var qtyToProducePage = dispatchPage.ClickQuantityToProduce();
            var CanUpdateQty = qtyToProducePage.CanUpdateQty();
            Assert.IsTrue(CanUpdateQty, "Le dispatch n'est pas validé mais des quantités sont visibles.");

            // Go to CustomerPortal Page
            homePage.Navigate();
            var customerPortalLoginPage = homePage.GoToCustomerPortal();

            try
            {
                customerPortalLoginPage.LogIn(userMail, userPassword);
                bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
                Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
                var customerPortalQtyByWeekPage = customerPortalLoginPage.ClickOnQtyByWeek();
                customerPortalQtyByWeekPage.SearchFilterDelivery(delivery);
                customerPortalQtyByWeekPage.WaitPageLoading();
                var demainModifiable = customerPortalQtyByWeekPage.CheckIfModifiableByweek(demain, serviceName);
                if(demainModifiable)
                {
                    quantityBeforeUpdate = customerPortalQtyByWeekPage.GetQuantityEditble(serviceName, demain);
                    customerPortalQtyByWeekPage.SetQuantityByweek(serviceName, demain, qty); 
                    Assert.AreNotEqual(quantityBeforeUpdate, qty, "la modification n'est pas effectuer ");
                    customerPortalQtyByWeekPage.ValidateQuantities();

                }
                Assert.IsTrue(demainModifiable, "demain non modifiable lorsque delivery Default");
                customerPortalLoginPage.Close();
                dispatchPage = homePage.GoToProduction_DispatchPage();
                dispatchPage.Filter(DispatchPage.FilterType.Site, site);
                dispatchPage.Filter(DispatchPage.FilterType.Search, serviceName);
                dispatchPage.Filter(DispatchPage.FilterType.Deliveries, delivery);
                previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
                Assert.IsTrue(previsionnalQty.IsDispatchValidated(), "Les previsionnal qty du dispatch n'ont pas été validées pour la semaine renseignée via le portal.");
                qtyToProducePage = dispatchPage.ClickQuantityToProduce();
                var quantityToProduceForTomorrow = qtyToProducePage.GetQuantityToProduceForTomorrow();
                Assert.AreEqual(qty, quantityToProduceForTomorrow, "Pas de modification CP vers Dispatch");

            }
            finally
            {
                homePage.Navigate();
                homePage.GoToCustomerPortal();
                var customerPortalQtyByWeekPage = customerPortalLoginPage.ClickOnQtyByWeek();
                customerPortalQtyByWeekPage.SearchFilterDelivery(delivery);
                customerPortalQtyByWeekPage.WaitPageLoading();
                customerPortalQtyByWeekPage.SetQuantityByweek(serviceName, demain ,quantityBeforeUpdate);
                customerPortalQtyByWeekPage.ValidateQuantities();
            }
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_QBW_SendCommandByMail()
        {
            // Prepare
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userCustomerPortalPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();
            var userPassword = TestContext.Properties["MailBox_UserPassword"].ToString();
            var mailSubjectToEmpty = "Empty Mail To";

            bool customerPortalTab = false;
            bool mailboxTab = false;

            var mailPage = new MailPage(WebDriver, TestContext);


            HomePage homePage = LogInAsAdmin();

            // Go to CustomerPortal Page
            var customerPortalLoginPage = homePage.GoToCustomerPortal();
            customerPortalTab = true;

            try
            {
                customerPortalLoginPage.LogIn(userMail, userCustomerPortalPassword);

                bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
                Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");

                var customerPortalQtyByWeekPage = customerPortalLoginPage.ClickOnQtyByWeek();

                // ================ test mail To vide ================
                string pj = customerPortalQtyByWeekPage.SendQtyByMail("", mailSubjectToEmpty);
                customerPortalQtyByWeekPage.SendMail();
                WebDriver.Navigate().Refresh();
                // ===================================================

                string piecejointe = customerPortalQtyByWeekPage.SendQtyByMail(userMail);
                Assert.AreEqual(pj, piecejointe, "Piece Jointe modifiée");
                string mailSubject = customerPortalQtyByWeekPage.GetMailSubject().Replace("  ", " ");
                customerPortalQtyByWeekPage.SendMail();

                WebDriver.Navigate().Refresh();
                // Deconnexion du portal
                customerPortalLoginPage.LogOut();
                customerPortalLoginPage.Close();
                customerPortalTab = false;

                // Connexion à la mailbox pour voir si le mail a été reçu
                mailPage = customerPortalLoginPage.RedirectToOutlookMailbox();
                mailboxTab = true;

                mailPage.FillFields_LogToOutlookMailbox(userMail);

                WebDriver.Navigate().Refresh();
                // Recherche du mail envoyé
                mailPage.ClickOnSpecifiedOutlookMail(mailSubject);

                // Vérifier que la pièce jointe est présente et qu'elle n'est pas vide
                bool isPieceJointeOK = mailPage.IsOutlookPieceJointeOK(piecejointe);
                Assert.IsTrue(isPieceJointeOK, "La pièce jointe n'est pas présente dans le mail.");

                mailPage.DeleteCurrentOutlookMail();

                // ================ test mail To vide ================
                mailPage.ClickOnSpecifiedOutlookMail(mailSubjectToEmpty);

                isPieceJointeOK = mailPage.IsOutlookPieceJointeOK(piecejointe);
                Assert.IsTrue(isPieceJointeOK, "La pièce jointe n'est pas présente dans le mail.");

                mailPage.DeleteCurrentOutlookMail();
                // ===================================================

                mailPage.Close();
                mailboxTab = false;
            }

            finally
            {
                if (mailboxTab)
                {
                    mailPage.Close();
                }

                if (customerPortalTab)
                {
                    customerPortalLoginPage.Close();
                }
            }
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_QBW_Filter_SearchServices()
        {
            //1.Aller sur Customer Portal, se logguer, aller sur Quantity By Week
            // Prepare
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();
            string serviceName = TestContext.Properties["ServiceCustomerPortal"].ToString();

            CustomerPortalLoginPage customerPortalLoginPage = LogInCustomerPortal();
            customerPortalLoginPage.LogIn(userMail, userPassword);

            bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
            bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();

            Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
            Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page Quantity By Week.");

            var customerPortalQtyByWeekPage = customerPortalLoginPage.ClickOnQtyByWeek();

            //2.Dans le filtre 'SEARCH SERVICES', chercher le service 'ServiceForCustomerPortal'
            customerPortalQtyByWeekPage.Search(serviceName);

            //3.Vérifier que le service est affiché dans le tableau
            Assert.AreEqual(serviceName, customerPortalQtyByWeekPage.SearchFirstResult(), "ServiceForCustomerPortal non trouvé dans le tableau");

            //4.Dans le filtre 'SEARCH SERVICES', chercher le service 'ServiceForCustomerPortal2'
            customerPortalQtyByWeekPage.Search(serviceName + "2");

            //5.Vérifier que le service n'est pas affiché dans le tableau
            Assert.IsTrue(customerPortalQtyByWeekPage.IsSearchResultEmpty(), "Pas de message à la place du tableau (vide)");

            if (isUserLogged)
            {
                customerPortalLoginPage.LogOut();
            }
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_QBW_Filter_SearchDelivery()
        {
            //1. Aller sur Customer Portal, se logguer, aller sur Quantity By Week
            // Prepare
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();
            string serviceName = TestContext.Properties["ServiceCustomerPortal"].ToString();

            CustomerPortalLoginPage customerPortalLoginPage = LogInCustomerPortal();
            customerPortalLoginPage.LogIn(userMail, userPassword);

            bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
            bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();

            Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
            Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page Quantity By Week.");

            var customerPortalQtyByWeekPage = customerPortalLoginPage.ClickOnQtyByWeek();

            //2.Dans le filtre 'DELIVERY', cocher le delivery 'DELIVERYFORCUSTOMERPORTAL
            customerPortalQtyByWeekPage.SearchFilterDelivery("DeliveryForCustomerPortal");

            //3.Vérifier que le service 'ServiceForCustomerPortal' est affiché dans le tableau
            Assert.AreEqual(serviceName, customerPortalQtyByWeekPage.SearchFirstResult(), "ServiceForCustomerPortal non trouvé dans le tableau");

            //4.Dans le filtre 'DELIVERY', chercher le service 'A.F.A. CASAS'
            customerPortalQtyByWeekPage.SearchFilterDelivery("A.F.A. CASAS");

            //5.Vérifier que le message 'No previsional quantity' s'affiche
            Assert.IsTrue(customerPortalQtyByWeekPage.IsSearchResultEmpty(), "Pas de message à la place du tableau (vide)");

            if (isUserLogged)
            {
                customerPortalLoginPage.LogOut();
            }
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_QBW_Filter_Date()
        {
            //1.Aller sur Customer Portal, se logguer, aller sur Quantity By Week
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();

            CustomerPortalLoginPage customerPortalLoginPage = LogInCustomerPortal();
            customerPortalLoginPage.LogIn(userMail, userPassword);

            bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
            bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();

            Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
            Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page Quantity By Week.");

            var customerPortalQtyByWeekPage = customerPortalLoginPage.ClickOnQtyByWeek();

            //2.Cliquer sur la flèche pour se déplacer à la semaine suivante
            customerPortalQtyByWeekPage.ClickOnNextWeek();

            //3.Vérifier que la date de la semaine suivante s'affiche
            // 29/11/2021 To 05/12/2021
            DateTime date1 = customerPortalQtyByWeekPage.GetCalendarBegin();
            DateTime date2 = customerPortalQtyByWeekPage.GetCalendarEnd();
            Assert.IsNotNull(date1, "parsing date début raté");
            Assert.IsNotNull(date2, "parsing date fin raté");
            Assert.IsTrue(date1 > DateUtils.Now.Date, "début de la semaine suivante");
            Assert.IsTrue(date2 > DateUtils.Now.Date, "fin de la semaine suivante");

            //4.Cliquer sur la flèche pour se déplacer à la semaine précédante
            customerPortalQtyByWeekPage.ClickOnPreviousWeek();

            //5.Vérifier que la date de la semaine affichée au départ s'affiche
            // 29/11/2021 To 05/12/2021
            date1 = customerPortalQtyByWeekPage.GetCalendarBegin();
            date2 = customerPortalQtyByWeekPage.GetCalendarEnd();
            Assert.IsNotNull(date1, "parsing date début raté");
            Assert.IsNotNull(date2, "parsing date fin raté");
            Assert.IsTrue(date1 <= DateUtils.Now.Date, "début de cette semaine");
            Assert.IsTrue(date2 >= DateUtils.Now.Date, "fin de cette semaine");

            if (isUserLogged)
            {
                customerPortalLoginPage.LogOut();
            }
        }


        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_CO_Filter_Search()
        {
            //(test identique à PR_CO_Filter_Search mais à créer sur Customer Portal)
            //1.Aller sur Customer Portal, se logguer, aller sur Customer Orders
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();

            // Prepare
            string customer = TestContext.Properties["InvoiceCustomerAirportTax"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string deliveryName = new Random().Next().ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();

            // Arrange
            CustomerPortalLoginPage customerPortalLoginPage = LogInCustomerPortal();
            customerPortalLoginPage.LogIn(userMail, userPassword);

            bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
            bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();

            Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
            Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page.");

            var customerPortalCustomerOrdersPage = customerPortalLoginPage.ClickOnCustomerOrders();

            //2.Cliquer sur "+"
            var customerPortalCustmerOrdersPageModal = customerPortalCustomerOrdersPage.ClickAdd();

            //2.Remplir le formulaire avec le site(BCN), le customer(EMIRATES AILRINES) et l'aircraft (B777)
            customerPortalCustmerOrdersPageModal.SelectSite(site);
            customerPortalCustmerOrdersPageModal.SelectCustomer(customer);
            customerPortalCustmerOrdersPageModal.SelectDeliveryName(deliveryName);
            customerPortalCustmerOrdersPageModal.SelectAircraft(aircraft);

            //3.Cliquer sur "create"
            CustomerPortalCustomerOrdersPageResult customerPortalCustomerOrdersPageResult = customerPortalCustmerOrdersPageModal.Create();

            //4.Récupérer le numéro du customer order créé, faire Back to List
            var resultText = customerPortalCustomerOrdersPageResult.ResultCustomerOrder();
            Regex r = new Regex("Customer order ([0-9]+) - TRE3 TRAVEL RETAIL SL ANE");
            Match m = r.Match(resultText);
            string CustomerOrder = m.Groups[1].Value;

            // back to list
            customerPortalCustomerOrdersPage.Navigate("/CustomerPortal/CustomerOrders");

            //5.Dans le filtre 'SEARCH', rechercher le numéro
            customerPortalCustomerOrdersPage.Search(CustomerOrder);
            int customerOrderNumber = customerPortalCustomerOrdersPage.CheckTotalNumber();
            string customerOrderNo = customerPortalCustomerOrdersPage.getTableFirstLine(1);
            //6.Vérifier si on filtre bien sur le bon customer
            Assert.AreEqual(1, customerOrderNumber, "Le filtre SEARCH ne fonctionne pas.");
            Assert.AreEqual("Id " + CustomerOrder, customerOrderNo, "Le customer filtré n'est pas trouvé.");

            if (isUserLogged)
            {
                customerPortalLoginPage.LogOut();
            }
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_CO_Filter_Date()
        {
            //(test identique à PR_CO_Filter_Date mais à créer sur Customer Portal)
            //1.Aller sur Customer Portal, se logguer, aller sur Customer Orders
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();

            // Arrange
            CustomerPortalLoginPage customerPortalLoginPage = LogInCustomerPortal();
            customerPortalLoginPage.LogIn(userMail, userPassword);

            bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
            bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();

            Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
            Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page.");

            var customerPortalCustomerOrdersPage = customerPortalLoginPage.ClickOnCustomerOrders();

            //2.Dans le filtre 'FROM', date d'aujourd'hui - 3 jours
            DateTime dateFrom = DateUtils.Now.Date.AddDays(-3);
            //3.Dans le filtre 'TO', date d'aujourd'hui + 7 jours
            DateTime dateTo = DateUtils.Now.Date.AddDays(7);
            customerPortalCustomerOrdersPage.FilterDate(dateFrom, dateTo);

            //4.Vérifier la date des customer order filtrés
            Assert.IsTrue(customerPortalCustomerOrdersPage.CheckTotalNumber() > 0, "Pas de ligne dans le tableau");
            for (int i = 1; i < customerPortalCustomerOrdersPage.CheckTotalNumber() && i < 8; i++)
            {
                DateTime dateTable = DateTime.ParseExact(customerPortalCustomerOrdersPage.getTableFirstLine(4, i), "dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture);
                Assert.IsTrue(dateFrom <= dateTable);
                Assert.IsTrue(dateTo >= dateTable);
            }

            if (isUserLogged)
            {
                customerPortalLoginPage.LogOut();
            }
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_CO_Filter_Customer()
        {
            //(test identique à PR_CO_Filter_Customer mais à créer sur Customer Portal)
            //1.Aller sur Customer Portal, se logguer, aller sur Customer Orders
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();

            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();

            // Arrange
            CustomerPortalLoginPage customerPortalLoginPage = LogInCustomerPortal();
            customerPortalLoginPage.LogIn(userMail, userPassword);

            bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
            bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();

            Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
            Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page.");

            var customerPortalCustomerOrdersPage = customerPortalLoginPage.ClickOnCustomerOrders();
            //2.Dans le filtre 'CUSTOMER', faire un Uncheck All, rechercher le customer 'EMIRATES AILRINES' et le cocher
            customerPortalCustomerOrdersPage.ShowFilterCustomersContent();
            customerPortalCustomerOrdersPage.FilterCustomersUnCheckAll();

            Assert.AreEqual("UAE - EMIRATES AIRLINES", customerPortalCustomerOrdersPage.FilterCustomerSelect("UAE - EMIRATES AIRLINES"), "filtre CUSTOMERS sans EMIRATES AIRLINES");

            //6.Vérifier si on filtre bien sur le bon customer order
            Assert.IsTrue(customerPortalCustomerOrdersPage.CheckTotalNumber() > 0, "Pas de ligne dans le tableau");
            for (int i = 1; i < customerPortalCustomerOrdersPage.CheckTotalNumber() && i < 8; i++)
            {
                string customerTable = customerPortalCustomerOrdersPage.getTableFirstLine(2, i);
                Assert.AreEqual(customer, customerTable);
            }

            if (isUserLogged)
            {
                customerPortalLoginPage.LogOut();
            }
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_CO_Add_Item()
        {
            //(test identique à PR_CO_Add_Item mais à créer sur Customer Portal)
            //1.Aller sur Customer Portal, se logguer, aller sur Customer Orders
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();

            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string deliveryName = new Random().Next().ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            // Arrange
            //1.Aller sur Customer Portal, se logguer, aller sur Customer Orders
            CustomerPortalLoginPage customerPortalLoginPage = LogInCustomerPortal();
            customerPortalLoginPage.LogIn(userMail, userPassword);
            bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
            bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();
            Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
            Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page.");
            var customerPortalCustomerOrdersPage = customerPortalLoginPage.ClickOnCustomerOrders();
            //2.Créer un Customer order(CUPO_CO_CreateNewCustomerOrder)
            var customerPortalCustmerOrdersPageModal = customerPortalCustomerOrdersPage.ClickAdd();
            //2.Remplir le formulaire avec le site(BCN), le customer(EMIRATES AILRINES) 
            customerPortalCustmerOrdersPageModal.SelectSite(site);
            customerPortalCustmerOrdersPageModal.SelectCustomer(customer);
            customerPortalCustmerOrdersPageModal.SelectDeliveryName(deliveryName);
            customerPortalCustmerOrdersPageModal.SelectAircraft(aircraft);
            //3.Cliquer sur "create"
            var customerPortalCustomerOrdersPageResult = customerPortalCustmerOrdersPageModal.Create();
            //3.Cliquer sur '+'
            //4.Ajouter un item et une quantity(HOT MEAL YC UAE et 10)
            customerPortalCustomerOrdersPageResult.AddLine("HOT MEAL YC UAE", 10);
            //5.Vérifier l'ajout de l'item
            // attente du rechargement du tableau
            customerPortalCustomerOrdersPageResult.WaitPageLoading();
            customerPortalCustomerOrdersPageResult.Validate();
            var resultText = customerPortalCustomerOrdersPageResult.ResultCustomerOrder();
            Regex r = new Regex("##### Customer order ##### ([0-9]+) - EMIRATES AIRLINES");
            Match m = r.Match(resultText);
            string customerOrderNumberId = Regex.Match(resultText, @"\d+").Value;
            // back to list
            customerPortalCustomerOrdersPage.Navigate("/CustomerPortal/CustomerOrders");
            string customerPortalCustomerOrderId = Regex.Match(customerPortalCustomerOrdersPage.getTableFirstLine(1), @"\d+").Value;
            string customerOrderIconValidation = customerPortalCustomerOrdersPage.CheckIconsValidated();
            //Vérifier dans la table si on retrouve le customer order créé
            Assert.AreEqual(customerOrderNumberId, customerPortalCustomerOrderId, "Première colonne ne contient pas " + customerOrderNumberId);
            Assert.AreEqual("OK", customerOrderIconValidation, "petite icone manquante sur les 3");
            if (isUserLogged)
            {
                customerPortalLoginPage.LogOut();
            }
        }

        /**
         * Inspiré de CUPO_CO_CreateNewCustomerOrder
         */
        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_CO_Delete_Item()
        {
            //1.Aller sur Customer Portal, se logguer, aller sur Customer Orders
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();

            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string deliveryName = new Random().Next().ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();

            // Arrange
            CustomerPortalLoginPage customerPortalLoginPage = LogInCustomerPortal();
            customerPortalLoginPage.LogIn(userMail, userPassword);
            bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
            bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();
            Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
            Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page.");
            var customerPortalCustomerOrdersPage = customerPortalLoginPage.ClickOnCustomerOrders();

            //2.Cliquer sur "+"
            var customerPortalCustmerOrdersPageModal = customerPortalCustomerOrdersPage.ClickAdd();
            //2.Remplir le formulaire avec le site(BCN), le customer(EMIRATES AILRINES)
            customerPortalCustmerOrdersPageModal.SelectSite(site);
            customerPortalCustmerOrdersPageModal.SelectCustomer(customer);
            customerPortalCustmerOrdersPageModal.SelectDeliveryName(deliveryName);
            customerPortalCustmerOrdersPageModal.SelectAircraft(aircraft);
            //3.Cliquer sur "create"
            var customerPortalCustomerOrdersPageResult = customerPortalCustmerOrdersPageModal.Create();
            //4.Récupérer le numéro du customer order crée, faire Back to List
            //result customer order
            var resultText = customerPortalCustomerOrdersPageResult.ResultCustomerOrder();
            Regex r = new Regex("##### Customer order ##### ([0-9]+) - EMIRATES AIRLINES");
            Match m = r.Match(resultText);
            string customerOrderNumberId = Regex.Match(resultText, @"\d+").Value;
            Assert.IsNotNull(customerOrderNumberId, "Customer order number not found");
            //3.Cliquer sur '+'
            //4.Ajouter un item et une quantity(HOT MEAL YC UAE et 10)
            customerPortalCustomerOrdersPageResult.AddLine("HOT MEAL YC UAE", 10);
            //5.Vérifier l'ajout de l'item
            // attente du rechargement du tableau
            customerPortalCustomerOrdersPageResult.WaitPageLoading();
            Assert.AreEqual(1, customerPortalCustomerOrdersPageResult.CheckTableauSize(), "tableau vide");
            //6.Supprimer l'item
            customerPortalCustomerOrdersPageResult.DeleteFirstLine();
            //7.Vérifier que l'item a été supprimé
            bool IsItemDeleted =customerPortalCustomerOrdersPageResult.CheckTableauEmpty();
            Assert.IsTrue(IsItemDeleted, "tableau pas vide");

            if (isUserLogged)
            {
                customerPortalLoginPage.LogOut();
            }
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_CO_Add_BillinComment_Item()
        {
            //1.Aller sur Customer Portal, se logguer, aller sur Customer Orders
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();

            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string deliveryName = new Random().Next().ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            CustomerPortalLoginPage customerPortalLoginPage = LogInCustomerPortal();
            customerPortalLoginPage.LogIn(userMail, userPassword);
            bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
            bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();
            Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
            Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page.");
            var customerPortalCustomerOrdersPage = customerPortalLoginPage.ClickOnCustomerOrders();
            //2.Créer un Customer order(CUPO_CO_CreateNewCustomerOrder)
            var customerPortalCustmerOrdersPageModal = customerPortalCustomerOrdersPage.ClickAdd();
            //Remplir le formulaire avec le site(BCN), le customer(EMIRATES AILRINES) 
            customerPortalCustmerOrdersPageModal.SelectSite(site);
            customerPortalCustmerOrdersPageModal.SelectCustomer(customer);
            customerPortalCustmerOrdersPageModal.SelectDeliveryName(deliveryName);
            customerPortalCustmerOrdersPageModal.SelectAircraft(aircraft);
            //Cliquer sur "create"
            var customerPortalCustomerOrdersPageResult = customerPortalCustmerOrdersPageModal.Create();
            //4.Récupérer le numéro du customer order crée, faire Back to List
            //result customer order
            var resultText = customerPortalCustomerOrdersPageResult.ResultCustomerOrder();
            Regex r = new Regex("##### Customer order ##### ([0-9]+) - EMIRATES AIRLINES");
            Match m = r.Match(resultText);
            string customerOrderNumberId = Regex.Match(resultText, @"\d+").Value;
            Assert.IsNotNull(customerOrderNumberId, "Customer order number not found");
            //3.Cliquer sur '+'
            //4.Ajouter un item et une quantity(HOT MEAL YC UAE et 10)
            customerPortalCustomerOrdersPageResult.AddLine("HOT MEAL YC UAE", 10);
            //5.Ajouter un Billing comment(comment)
            customerPortalCustomerOrdersPageResult.FillBilling("test Billing");
            // back to list
            customerPortalCustomerOrdersPage.Navigate("/CustomerPortal/CustomerOrders");
            string customerPortalCustomerOrderId = Regex.Match(customerPortalCustomerOrdersPage.getTableFirstLine(1), @"\d+").Value;
            //Vérifier dans la table si on retrouve le customer order créé
            Assert.AreEqual(customerOrderNumberId, customerPortalCustomerOrderId, "Première colonne ne contient pas " + customerOrderNumberId);
            //6.Vérifier l'ajout du comment sur l'item(bien affiché et identique à ce qui a été entrée)
            customerPortalCustomerOrdersPageResult = customerPortalCustomerOrdersPage.GoToFirstCustomer();
            string comment_text = customerPortalCustomerOrdersPageResult.GetBillingComment();
            Assert.AreEqual("test Billing", comment_text, "mauvais billing comment");
            // ferme la popup, ainsi le bouton logOut est accessible
            customerPortalCustomerOrdersPageResult.CancelBilling();
            if (isUserLogged)
            {
                customerPortalLoginPage.LogOut();
            }
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_CO_Delete_Customer_Order()
        {
            //(test identique à PR_CO_Delete_Customer_Order mais à créer sur Customer Portal)
            //1.Aller sur Customer Portal, se logguer, aller sur Customer Orders
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();

            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string deliveryName = new Random().Next().ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();

            // Arrange
            CustomerPortalLoginPage customerPortalLoginPage = LogInCustomerPortal();
            customerPortalLoginPage.LogIn(userMail, userPassword);

            bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
            bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();

            Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
            Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page.");

            var customerPortalCustomerOrdersPage = customerPortalLoginPage.ClickOnCustomerOrders();

            //2.Créer un Customer order(CUPO_CO_CreateNewCustomerOrder)
            //Cliquer sur "+"
            var customerPortalCustmerOrdersPageModal = customerPortalCustomerOrdersPage.ClickAdd();
            //Remplir le formulaire avec le site(BCN), le customer(EMIRATES AILRINES) et l'aircraft (B777)
            customerPortalCustmerOrdersPageModal.SelectSite(site);
            customerPortalCustmerOrdersPageModal.SelectCustomer(customer);
            customerPortalCustmerOrdersPageModal.SelectDeliveryName(deliveryName);
            customerPortalCustmerOrdersPageModal.SelectAircraft(aircraft);
            //Cliquer sur "create"
            CustomerPortalCustomerOrdersPageResult customerPortalCustomerOrdersPageResult = customerPortalCustmerOrdersPageModal.Create();
            //Récupérer le numéro du customer order crée, faire Back to List
            //result customer order
            var resultText = customerPortalCustomerOrdersPageResult.ResultCustomerOrder();
            Regex r = new Regex("##### Customer order ##### ([0-9]+) - EMIRATES AIRLINES");
            Match m = r.Match(resultText);
            string customerOrderNumberId = Regex.Match(resultText, @"\d+").Value;

            //3.Faire 'BackToList'
            // back to list
            customerPortalCustomerOrdersPage.Navigate("/CustomerPortal/CustomerOrders");

            //4.Search le customer order crée par son order no
            customerPortalCustomerOrdersPage.Search(customerOrderNumberId);
            //Vérifier dans la table si on retrouve le customer order créé
            string customerPortalCustomerOrderId = Regex.Match(customerPortalCustomerOrdersPage.getTableFirstLine(1), @"\d+").Value;
            //Vérifier dans la table si on retrouve le customer order créé
            Assert.AreEqual(customerOrderNumberId, customerPortalCustomerOrderId, "Première colonne ne contient pas " + customerOrderNumberId);

            //5.Cliquer sur l'icone 'poubelle'
            //6.Cliquer sur 'DELETE'
            customerPortalCustomerOrdersPage.DeleteFirstItem();
            // relance la recherche car on a reset le moteur de recherche
            customerPortalCustomerOrdersPage.Search(customerOrderNumberId);

            //7.Vérifier qu'il est bien supprimé
            Assert.AreEqual(0, customerPortalCustomerOrdersPage.CheckTotalNumber());
            if (isUserLogged)
            {
                customerPortalLoginPage.LogOut();
            }
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_CO_Filter_OrderStatus_Validated()
        {
            //(test identique à PR_CO_Filter_Show_InvoiceStep_Validate mais à créer sur Customer Portal)
            //1.Aller sur Customer Portal, se logguer, aller sur Customer Orders
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();
            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string deliveryName = new Random().Next().ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            // Arrange
            CustomerPortalLoginPage customerPortalLoginPage = LogInCustomerPortal();
            customerPortalLoginPage.LogIn(userMail, userPassword);
            bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
            bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();
            Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
            Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page.");
            var customerPortalCustomerOrdersPage = customerPortalLoginPage.ClickOnCustomerOrders();
            //2.Créer un Customer order et le valider s'il n'y en a pas
            //Cliquer sur "+"
            var customerPortalCustmerOrdersPageModal = customerPortalCustomerOrdersPage.ClickAdd();
            //Remplir le formulaire avec le site(BCN), le customer(EMIRATES AILRINES) et l'aircraft (B777)
            customerPortalCustmerOrdersPageModal.SelectSite(site);
            customerPortalCustmerOrdersPageModal.SelectCustomer(customer);
            customerPortalCustmerOrdersPageModal.SelectDeliveryName(deliveryName);
            customerPortalCustmerOrdersPageModal.SelectAircraft(aircraft);
            //Cliquer sur "create"
            customerPortalCustmerOrdersPageModal.WaitPageLoading();
            CustomerPortalCustomerOrdersPageResult customerPortalCustomerOrdersPageResult = customerPortalCustmerOrdersPageModal.Create();

            //Cliquer sur '+'
            //Ajouter un item et une quantity(HOT MEAL YC UAE et 10)
            customerPortalCustomerOrdersPageResult.AddLine("HOT MEAL YC UAE", 10);
            //Vérifier l'ajout de l'item
            // attente du rechargement du tableau
            customerPortalCustomerOrdersPageResult.WaitPageLoading();
            customerPortalCustomerOrdersPageResult.Validate();

            //Récupérer le numéro du customer order crée, faire Back to List
            //result customer order
            var resultText = customerPortalCustomerOrdersPageResult.ResultCustomerOrder();
            Regex r = new Regex("Customer order ([0-9]+) - EMIRATES AIRLINES");
            Match m = r.Match(resultText);
            string CustomerOrder = m.Groups[1].Value;

            //3. Faire 'Back To List'
            customerPortalCustomerOrdersPage.Navigate("/CustomerPortal/CustomerOrders");
            //4. Filtrer 'ORDER STATUS' à 'IN VALIDATED'
            customerPortalCustomerOrdersPage.ResetFilter();
            customerPortalCustomerOrdersPage.ShowFilterCustomersOrder("validated");
            // tableau se charge
            //customerPortalCustomerOrdersPage.WaitPageLoading();
            int nbValidated = customerPortalCustomerOrdersPage.CheckTotalNumber_customerPortal();
            Assert.AreNotEqual(0, nbValidated, "pas de CUSTOMER ORDER validated");

            //5. Vérifier que le customer order crée/validé est bien validé ( icone validation vert à gauche)
            //Vérifier dans la table si on retrouve le customer order créé
            Assert.AreEqual("Id " + CustomerOrder, customerPortalCustomerOrdersPage.getTableFirstLine(1));
            Assert.AreEqual("OK", customerPortalCustomerOrdersPage.CheckIconsValidated(), "Manque l'icone validated");

            // purger le dossier Download
            string directory = TestContext.Properties["DownloadsPath"].ToString();
            foreach (string f in Directory.GetFiles(directory))
            {
                //export-orders 17-12-2021_2-33-56 PM.xlsx
                if (f.Contains("export-orders") && f.EndsWith(".xlsx"))
                {
                    File.Delete(f);
                }
            }
            customerPortalCustomerOrdersPage.PageSize("100");
            var validatedList = customerPortalCustomerOrdersPage.GetPriceLinesGreaterThanZero().Count;
            //6. Faire un export excel de ce customer
            // hack the link to be clickable
            customerPortalCustomerOrdersPage.ExportExcel();

            //8.Vérifier les informations : order ID, order Date, Site, Customer, Price, Detail name, quantity, unit
            string trouve = null;
            foreach (string f in Directory.GetFiles(directory))
            {
                //export-orders 17-12-2021_2-33-56 PM.xlsx
                if (f.Contains("export-orders") && f.EndsWith(".xlsx"))
                {
                    trouve = f;
                }
            }
            Assert.IsNotNull(trouve);
            int resultNumber = OpenXmlExcel.GetExportResultNumber("Sheet 1", trouve);
            //Assert
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.AreEqual(validatedList, resultNumber, "Le nombre de lignes dans l'export Excel ne correspond pas au nombre de lignes validées dans le filtre appliqué.");


            //7.Vérifier dans le fichier d'export excel que le customer order donné est à 'Validated' dans la colonne 'Invoice step'
            Assert.AreEqual(CustomerOrder, OpenXmlExcel.GetValuesInList("Order ID", "Sheet 1", trouve)[0], " excel colonne Order ID");
            Assert.AreEqual("Validated", OpenXmlExcel.GetValuesInList("Invoice step", "Sheet 1", trouve)[0], " excel colonne Invoice step");


            if (isUserLogged)
            {
                customerPortalLoginPage.LogOut();
            }
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_CO_ResetFilter()
        {
            //1.Aller sur Customer Portal, se logguer, aller sur Customer Orders
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();

            // Arrange
            CustomerPortalLoginPage customerPortalLoginPage = LogInCustomerPortal();
            customerPortalLoginPage.LogIn(userMail, userPassword);

            bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
            bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();

            Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
            Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page.");

            var customerPortalCustomerOrdersPage = customerPortalLoginPage.ClickOnCustomerOrders();

            //2.Récupérer le numéro de l'id du premier customer order
            string CustomerOrder = customerPortalCustomerOrdersPage.getTableFirstLine(1).Substring(3);

            //3.Dans le filtre 'SEARCH', rechercher le numéro
            customerPortalCustomerOrdersPage.Search(CustomerOrder);

            //4.Dans le filtre 'FROM', date d'aujourd'hui - 3 jours
            DateTime dateFrom = DateUtils.Now.Date.AddDays(-3);

            //5.Dans le filtre 'TO', date d'aujourd'hui + 7 jours
            DateTime dateTo = DateUtils.Now.Date.AddDays(7);
            customerPortalCustomerOrdersPage.FilterDate(dateFrom, dateTo);



            //6.Dans le filtre 'CUSTOMER', faire un Uncheck All, rechercher le customer 'EMIRATES AILRINES' et le cocher
            customerPortalCustomerOrdersPage.ShowFilterCustomersContent();
            customerPortalCustomerOrdersPage.FilterCustomersUnCheckAll();
            Assert.AreEqual("UAE - EMIRATES AIRLINES", customerPortalCustomerOrdersPage.FilterCustomerSelect("UAE - EMIRATES AIRLINES"), "filtre CUSTOMERS sans UAE - EMIRATES AIRLINES");


            //7.Dans le filtre 'ORDER STATUS', sélectionner 'VALIDATED'
            customerPortalCustomerOrdersPage.ShowFilterCustomersOrder("validated");

            //8.Cliquer sur 'Reset Filter'
            customerPortalCustomerOrdersPage.ResetFilter();

            //9.Les champs 'SEARCH' et 'To' sont vidés, la date 'FROM' est réinitialisée à la date d'aujourd'hui - 7 jours, tous les 'CUSTOMERS' sont sélectionnés, l''ORDER STATUS' est à 'ALL'
            Assert.AreEqual("", customerPortalCustomerOrdersPage.SearchInput(), "Filter Search pas vide");
            Assert.AreEqual("", customerPortalCustomerOrdersPage.DateToInput(), "Filter Date To pas vide");
            DateTime dateFrom7 = DateUtils.Now.Date.AddDays(-7);
            Assert.AreEqual(dateFrom7.ToString("dd/MM/yyyy"), customerPortalCustomerOrdersPage.DateFromInput(), "Filter Date From pas J-7");

            var filterCustomersText = customerPortalCustomerOrdersPage.FilterCustomerLabel();
            Regex r = new Regex("([0-9]+) of ([0-9]+) selected");
            Match m = r.Match(filterCustomersText);
            string select1 = m.Groups[1].Value;
            string select2 = m.Groups[2].Value;
            Assert.AreEqual(select1, select2, "Filter Customers pas tous selectionnés : " + filterCustomersText);

            Assert.IsTrue(customerPortalCustomerOrdersPage.ShowFilterCustomersOrderIsSelected("all"), "all non sélectionné");

            if (isUserLogged)
            {
                customerPortalLoginPage.LogOut();
            }
        }
        //[Timeout(_timeout)]
        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_CO_Validate()
        {
            //(test identique à PR_CO_Validate mais à créer sur Customer Portal)
            //1.Aller sur Customer Portal, se logguer, aller sur Customer Orders
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();

            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string deliveryName = new Random().Next().ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            // Arrange
            CustomerPortalLoginPage customerPortalLoginPage = LogInCustomerPortal();
            customerPortalLoginPage.LogIn(userMail, userPassword);
            bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
            bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();
            Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
            Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page.");
            var customerPortalCustomerOrdersPage = customerPortalLoginPage.ClickOnCustomerOrders();
            //2.Créer un Customer order(CUPO_CO_CreateNewCustomerOrder)
            var customerPortalCustmerOrdersPageModal = customerPortalCustomerOrdersPage.ClickAdd();
            //Remplir le formulaire avec le site(BCN), le customer(EMIRATES AILRINES)
            customerPortalCustmerOrdersPageModal.SelectSite(site);
            customerPortalCustmerOrdersPageModal.SelectCustomer(customer);
            customerPortalCustmerOrdersPageModal.SelectDeliveryName(deliveryName);
            customerPortalCustmerOrdersPageModal.SelectAircraft(aircraft);
            //Cliquer sur "create"
            CustomerPortalCustomerOrdersPageResult customerPortalCustomerOrdersPageResult = customerPortalCustmerOrdersPageModal.Create();
            //3.Cliquer sur '+'
            //4.Ajouter un item et une quantity(HOT MEAL YC UAE et 10)
            customerPortalCustomerOrdersPageResult.AddLine("HOT MEAL YC UAE", 10);
            //Vérifier l'ajout de l'item
            // attente du rechargement du tableau
            customerPortalCustomerOrdersPageResult.WaitPageLoading();
            //5. 'Validate'le customer order
            //6. Cliquer sur 'Confirm'
            customerPortalCustomerOrdersPageResult.Validate();
            //Récupérer le numéro du customer order crée, faire Back to List
            //result customer order
            var resultText = customerPortalCustomerOrdersPageResult.ResultCustomerOrder();
            Regex r = new Regex("Customer order ([0-9]+) - EMIRATES AIRLINES");
            Match m = r.Match(resultText);
            string CustomerOrder = m.Groups[1].Value;
            //7. Faire 'Back to List'
            customerPortalCustomerOrdersPage.Navigate("/CustomerPortal/CustomerOrders");
            //8. 'Search' ce customer order validé
            customerPortalCustomerOrdersPage.Search(CustomerOrder);
            string OrderNo = customerPortalCustomerOrdersPage.getTableFirstLine(1);
            Assert.AreEqual("Id " + CustomerOrder, OrderNo);
            string isValidationIconDisplayed = customerPortalCustomerOrdersPage.CheckIconsValidated();
            //9. Vérifier la validation (icone validation vert à gauche)
            Assert.AreEqual("OK", isValidationIconDisplayed, "Manque l'icone validated");
            if (isUserLogged)
            {
                customerPortalLoginPage.LogOut();
            }
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_CO_PrintCustomerOrder()
        {
            //(test identique à PR_CO_Print_CustomerOrder_NewVersion mais à créer sur Customer Portal)
            //1.Aller sur Customer Portal, se logguer, aller sur Customer Orders
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();
            
            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string deliveryName = new Random().Next().ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            // Arrange
            CustomerPortalLoginPage customerPortalLoginPage = LogInCustomerPortal();
            customerPortalLoginPage.LogIn(userMail, userPassword);
            customerPortalLoginPage.PurgeDownloads();

            bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
            bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();

            Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
            Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page.");

            var customerPortalCustomerOrdersPage = customerPortalLoginPage.ClickOnCustomerOrders();

            //2.Créer un Customer order(CUPO_CO_CreateNewCustomerOrder)
            var customerPortalCustmerOrdersPageModal = customerPortalCustomerOrdersPage.ClickAdd();
            //Remplir le formulaire avec le site(BCN), le customer(EMIRATES AILRINES) et l'aircraft (B777)
            customerPortalCustmerOrdersPageModal.SelectSite(site);
            customerPortalCustmerOrdersPageModal.SelectCustomer(customer);
            customerPortalCustmerOrdersPageModal.SelectDeliveryName(deliveryName);
            customerPortalCustmerOrdersPageModal.SelectAircraft(aircraft);
            //Cliquer sur "create"
            CustomerPortalCustomerOrdersPageResult customerPortalCustomerOrdersPageResult = customerPortalCustmerOrdersPageModal.Create();

            //3.Cliquer sur '+'
            //4.Ajouter un item et une quantity(HOT MEAL YC UAE et 10)
            customerPortalCustomerOrdersPageResult.AddLine("HOT MEAL YC UAE", 10);
            // récupérer le CustomerOrder
            var resultText = customerPortalCustomerOrdersPageResult.ResultCustomerOrder();
            Regex r = new Regex("Customer order ([0-9]+) - EMIRATES AIRLINES");
            Match m = r.Match(resultText);
            string CustomerOrder = m.Groups[1].Value;

            //5.Faire un 'Back to list'
            customerPortalCustomerOrdersPage.Navigate("/CustomerPortal/CustomerOrders");

            //6.Rechercher le customer order crée
            customerPortalCustomerOrdersPage.Search(CustomerOrder);

            string directory = TestContext.Properties["DownloadsPath"].ToString();
            string file = directory + Path.DirectorySeparatorChar + "Customer Orders.pdf";
            File.Delete(file);

            //7.Cliquer sur Print-(bouton imprimante)
            customerPortalCustomerOrdersPage.Print();
            //8.Vérifier que le rapport est généré
            long length = new FileInfo(file).Length;
            if (length > 1024)
            {
                Assert.IsTrue(length > 1024, "Fichier " + file + " généré vide");
            }
            else
            {
                Assert.Fail("Fichier " + file + " généré inexistant");
            }

            if (isUserLogged)
            {
                customerPortalLoginPage.LogOut();
            }
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_CO_Update_Item()
        {
            //(test identique à PR_CO_Update_Item mais à créer sur Customer Portal)
            //1.Aller sur Customer Portal, se logguer, aller sur Customer Orders
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();

            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string deliveryName = new Random().Next().ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            // Arrange
            CustomerPortalLoginPage customerPortalLoginPage = LogInCustomerPortal();
            customerPortalLoginPage.LogIn(userMail, userPassword);

            bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
            bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();
            Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
            Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page.");
            var customerPortalCustomerOrdersPage = customerPortalLoginPage.ClickOnCustomerOrders();
            //2.Créer un Customer order(CUPO_CO_CreateNewCustomerOrder)
            var customerPortalCustmerOrdersPageModal = customerPortalCustomerOrdersPage.ClickAdd();
            //Remplir le formulaire avec le site(BCN), le customer(EMIRATES AILRINES)
            customerPortalCustmerOrdersPageModal.SelectSite(site);
            customerPortalCustmerOrdersPageModal.SelectCustomer(customer);
            customerPortalCustmerOrdersPageModal.SelectDeliveryName(deliveryName);
            customerPortalCustmerOrdersPageModal.SelectAircraft(aircraft);
            //Cliquer sur "create"
            CustomerPortalCustomerOrdersPageResult customerPortalCustomerOrdersPageResult = customerPortalCustmerOrdersPageModal.Create();
            //3.Cliquer sur '+'
            //4.Ajouter un item et une quantity(HOT MEAL YC UAE et 10)
            customerPortalCustomerOrdersPageResult.AddLine("HOT MEAL YC UAE", 10);
            //5.Modifier l'item (COLD CUT AND CHEESE) et la quantity (5)
            customerPortalCustomerOrdersPageResult.UpdateLine("COLD CUT AND CHEESE", 5);
            // attendre que le tableau se rafraichie
            customerPortalCustomerOrdersPageResult.WaitPageLoading();
            //6.Cliquer sur l'onglet 'General Information' et revenir sur l'onglet 'Details'
            customerPortalCustomerOrdersPageResult.GoToGeneralInformation();
            customerPortalCustomerOrdersPageResult.GoToDetails();
            customerPortalCustomerOrdersPageResult.WaitPageLoading();
            string itemName = customerPortalCustomerOrdersPageResult.GetTableFirstLine(1);
            string customerOrderQty = customerPortalCustomerOrdersPageResult.GetTableFirstLine(2);
            //7.Vérifier les changements
            Assert.AreEqual("COLD CUT AND CHEESE", itemName, "Ne trouve pas COLD CUT AND CHEESE pour Name");
            Assert.AreEqual("5", customerOrderQty, "Ne trouve pas 5 pour Qty");

            if (isUserLogged)
            {
                customerPortalLoginPage.LogOut();
            }
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_CO_SendByMail()
        {
            //(test similaire à PR_CO_SendByMail mais à créer sur Customer Portal)
            //1.Aller sur Customer Portal, se logguer, aller sur Customer Orders
            var customPortalUserMail = TestContext.Properties["Admin_UserName"].ToString();
            var userCustomerPortalPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();
            var userPassword = TestContext.Properties["Admin_Password"].ToString();
            var userMail = TestContext.Properties["Admin_UserName"].ToString();

            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string deliveryName = new Random().Next().ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();


            // Arrange
            CustomerPortalLoginPage customerPortalLoginPage = LogInCustomerPortal();
            customerPortalLoginPage.LogIn(customPortalUserMail, userCustomerPortalPassword);

            bool isUserLogged = customerPortalLoginPage.IsUserLogged(customPortalUserMail);
            bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();

            Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
            Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page.");

            var customerPortalCustomerOrdersPage = customerPortalLoginPage.ClickOnCustomerOrders();

            //2.Créer un Customer order(CUPO_CO_CreateNewCustomerOrder)
            var customerPortalCustmerOrdersPageModal = customerPortalCustomerOrdersPage.ClickAdd();
            //Remplir le formulaire avec le site(BCN), le customer(EMIRATES AILRINES) et l'aircraft (B777)
            customerPortalCustmerOrdersPageModal.SelectSite(site);
            customerPortalCustmerOrdersPageModal.SelectCustomer(customer);
            customerPortalCustmerOrdersPageModal.SelectDeliveryName(deliveryName);
            customerPortalCustmerOrdersPageModal.SelectAircraft(aircraft);
            //Cliquer sur "create"
            CustomerPortalCustomerOrdersPageResult customerPortalCustomerOrdersPageResult = customerPortalCustmerOrdersPageModal.Create();

            //3.Ajouter un item et une quantity(HOT MEAL YC UAE et 10)
            customerPortalCustomerOrdersPageResult.AddLine("HOT MEAL YC UAE", 10);

            //4.Validate
            customerPortalCustomerOrdersPageResult.Validate();

            // récupérer le CustomerOrder
            var resultText = customerPortalCustomerOrdersPageResult.ResultCustomerOrder().Replace("#", "").Trim();
            Regex r = new Regex("([0-9]+) - EMIRATES AIRLINES");
            Match m = r.Match(resultText);
            string CustomerOrder = m.Groups[1].Value;

            //5.Survoler la coche et cliquer sur 'Send by email'
            //6.Remplir le formulaire To: testauto.newrest @gmail.com
            //7.Cliquer sur send
            customerPortalCustomerOrdersPageResult.SendMailAfterValidated(userMail);

            customerPortalCustomerOrdersPage.Navigate("/CustomerPortal/CustomerOrders");

            Assert.IsTrue(customerPortalCustomerOrdersPageResult.CheckEnveloppe(CustomerOrder), "pas d'enveloppe");

            //8.Vérifier sur gmail, la réception de l'email
            // Connexion à la mailbox pour voir si le mail a été reçu
            MailPage mailPage = customerPortalLoginPage.RedirectToOutlookMailbox();

            mailPage.FillFields_LogToOutlookMailbox(userMail, userPassword);

            WebDriver.Navigate().Refresh();
            // Recherche du mail envoyé

            Assert.IsTrue(mailPage.CheckIfSpecifiedOutlookMailExist("Customer order n° " + CustomerOrder + " validated"), "Mail Customer order n° " + CustomerOrder + " (quote) non trouvé");

            mailPage.DeleteCurrentOutlookMail();
            mailPage.Close();

            // purger le dossier Download
            DeleteAllFileDownload();

            if (isUserLogged)
            {
                customerPortalLoginPage.LogOut();
            }
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_CO_Export_CustomerOrder()
        {
            //(test similaire à PR_CO_Export_CustomerOrder_NewVersion mais à créer sur Customer Portal)
            //1.Aller sur Customer Portal, se logguer, aller sur Customer Orders
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userCustomerPortalPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();

            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string deliveryName = new Random().Next().ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();


            // Arrange
            CustomerPortalLoginPage customerPortalLoginPage = LogInCustomerPortal();
            customerPortalLoginPage.LogIn(userMail, userCustomerPortalPassword);

            bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
            bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();

            Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
            Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page.");

            var customerPortalCustomerOrdersPage = customerPortalLoginPage.ClickOnCustomerOrders();

            //2.Créer un Customer order(CUPO_CO_CreateNewCustomerOrder)
            var customerPortalCustmerOrdersPageModal = customerPortalCustomerOrdersPage.ClickAdd();
            //Remplir le formulaire avec le site(BCN), le customer(EMIRATES AILRINES) et l'aircraft (B777)
            customerPortalCustmerOrdersPageModal.SelectSite(site);
            customerPortalCustmerOrdersPageModal.SelectCustomer(customer);
            customerPortalCustmerOrdersPageModal.SelectDeliveryName(deliveryName);
            customerPortalCustmerOrdersPageModal.SelectAircraft(aircraft);
            //Cliquer sur "create"
            CustomerPortalCustomerOrdersPageResult customerPortalCustomerOrdersPageResult = customerPortalCustmerOrdersPageModal.Create();

            //3.Ajouter un item et une quantity(HOT MEAL YC UAE et 10)
            customerPortalCustomerOrdersPageResult.AddLine("HOT MEAL YC UAE", 10);

            //4.Validate
            customerPortalCustomerOrdersPageResult.Validate();
            var unitPrice = customerPortalCustomerOrdersPageResult.GetUnitPrice();
            var totalPrice = customerPortalCustomerOrdersPageResult.GetTotalPrice();
            // récupérer le CustomerOrder
            var resultText = customerPortalCustomerOrdersPageResult.ResultCustomerOrder();
            Regex r = new Regex("##### Customer order ##### ([0-9]+) - EMIRATES AIRLINES");
            Match m = r.Match(resultText);
            string customerOrderNumberId = Regex.Match(resultText, @"\d+").Value;
            Assert.IsNotNull(customerOrderNumberId, "Customer order number not found");

            //5.Faire 'Back to List'
            customerPortalCustomerOrdersPage.Navigate("/CustomerPortal/CustomerOrders");
            //6.Chercher le customer créé et validé
            customerPortalCustomerOrdersPage.Search(customerOrderNumberId);
            Assert.AreEqual("OK", customerPortalCustomerOrdersPage.CheckIconsValidated());

            // purger le dossier Download
            string directory = TestContext.Properties["DownloadsPath"].ToString();
            foreach (string f in Directory.GetFiles(directory))
            {
                //export-orders 17-12-2021_2-33-56 PM.xlsx
                if (f.Contains("export-orders") && f.EndsWith(".xlsx"))
                {
                    File.Delete(f);
                }
            }

            //7.Cliquer sur la flèche ddirigée vers le bas, puis Export
            // hack the link to be clickable
            customerPortalCustomerOrdersPage.ExportExcel();

            // barre de progression
            customerPortalCustomerOrdersPage.WaitPageLoading();

            //8.Vérifier les informations : order ID, order Date, Site, Customer, Price, Detail name, quantity, unit
            string trouve = null;
            foreach (string f in Directory.GetFiles(directory))
            {
                //export-orders 17-12-2021_2-33-56 PM.xlsx
                if (f.Contains("export-orders") && f.EndsWith(".xlsx"))
                {
                    trouve = f;
                }
            }
            Assert.IsNotNull(trouve);

            int resultNumber = OpenXmlExcel.GetExportResultNumber("Sheet 1", trouve);
            customerPortalCustomerOrdersPage.CLICKFIRSTLINE();
            string vat = customerPortalCustomerOrdersPage.GetVAT();
            string commercialname = customerPortalCustomerOrdersPage.GetCommercialName();
            string qty = customerPortalCustomerOrdersPage.GetQTY();
            string price = customerPortalCustomerOrdersPage.GETPRICE();
            string unit = customerPortalCustomerOrdersPage.GETUNIT();


            //Assert
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);

            Assert.AreEqual(customerOrderNumberId, OpenXmlExcel.GetValuesInList("Order ID", "Sheet 1", trouve)[0], " excel colonne Order ID");
            Assert.AreEqual(DateUtils.Now.Date, DateTime.FromOADate(Int32.Parse(OpenXmlExcel.GetValuesInList("Order Date", "Sheet 1", trouve)[0])).Date, " excel colonne Order Date");
            Assert.AreEqual(site, OpenXmlExcel.GetValuesInList("Site", "Sheet 1", trouve)[0], " excel colonne Site");
            Assert.AreEqual(customer, OpenXmlExcel.GetValuesInList("Customer", "Sheet 1", trouve)[0], " excel colonne Customer");
            Assert.AreEqual(vat, OpenXmlExcel.GetValuesInList("Price excl. VAT", "Sheet 1", trouve)[0], " excel colonne Price");
            Assert.AreEqual(commercialname, OpenXmlExcel.GetValuesInList("Detail Name", "Sheet 1", trouve)[0], " excel colonne Detail Name");
            Assert.AreEqual(qty, OpenXmlExcel.GetValuesInList("Quantity", "Sheet 1", trouve)[0], " excel colonne Quantity");
            Assert.AreEqual(price, OpenXmlExcel.GetValuesInList("Unit Price", "Sheet 1", trouve)[0], " excel colonne Unit");

            if (isUserLogged)
            {
                customerPortalLoginPage.LogOut();
            }
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_CO_CheckTotalPrice()
        {
            //1.Aller sur Customer Portal, se logguer, aller sur Customer Orders
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userCustomerPortalPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();

            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string deliveryName = new Random().Next().ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            int qty = 5;
            int qty2 = 3;

            // Arrange
            CustomerPortalLoginPage customerPortalLoginPage = LogInCustomerPortal();
            customerPortalLoginPage.LogIn(userMail, userCustomerPortalPassword);
            bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
            bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();
            Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
            Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page.");
            var customerPortalCustomerOrdersPage = customerPortalLoginPage.ClickOnCustomerOrders();
            //2.Créer un Customer order(CUPO_CO_CreateNewCustomerOrder)
            var customerPortalCustmerOrdersPageModal = customerPortalCustomerOrdersPage.ClickAdd();
            //Remplir le formulaire avec le site(BCN), le customer(EMIRATES AILRINES) et l'aircraft (B777)
            customerPortalCustmerOrdersPageModal.SelectSite(site);
            customerPortalCustmerOrdersPageModal.SelectCustomer(customer);
            customerPortalCustmerOrdersPageModal.SelectDeliveryName(deliveryName);
            customerPortalCustmerOrdersPageModal.SelectAircraft(aircraft);
            //Cliquer sur "create"
            CustomerPortalCustomerOrdersPageResult customerPortalCustomerOrdersPageResult = customerPortalCustmerOrdersPageModal.Create();
            //3.Cliquer sur '+'
            //4.Ajouter un item et une quantity(COLD CUT AND CHEESE et 5)
            customerPortalCustomerOrdersPageResult.AddLine("COLD CUT AND CHEESE", qty, 1);
            customerPortalCustomerOrdersPageResult.WaitLoading();
            customerPortalCustomerOrdersPageResult.AddLine("COLD CUT AND CHEESE", qty2, 2);
            customerPortalCustomerOrdersPageResult.WaitLoading();
            //5.Vérifier le total price
            double UnitPrice = Convert.ToDouble(customerPortalCustomerOrdersPageResult.GetTableFirstLine(4, 1));
            double TotalPrice = Convert.ToDouble(customerPortalCustomerOrdersPageResult.GetTableFirstLine(5, 1));
            double UnitPrice2 = Convert.ToDouble(customerPortalCustomerOrdersPageResult.GetTableFirstLine(4, 2));
            double TotalPrice2 = Convert.ToDouble(customerPortalCustomerOrdersPageResult.GetTableFirstLine(5, 2));
            Assert.AreEqual(qty * UnitPrice, TotalPrice, 0.00000001, "erreur CheckTotalPrice");
            Assert.AreEqual(qty2 * UnitPrice2, TotalPrice2, 0.00000001, "erreur CheckTotalPrice");
            string total = customerPortalCustomerOrdersPageResult.PriceExclVAT();
            total = total.Substring(2);
            Assert.AreEqual(qty * UnitPrice + qty2 * UnitPrice2, Convert.ToDouble(total), 0.00000001, "erreur CheckPriceExclVAT");
            if (isUserLogged)
            {
                customerPortalLoginPage.LogOut();
            }
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_CO_General_information_Modification()
        {
            //1.Aller sur Customer Portal, se logguer, aller sur Customer Orders
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userCustomerPortalPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();

            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string deliveryName = new Random().Next().ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            // Arrange
            CustomerPortalLoginPage customerPortalLoginPage = LogInCustomerPortal();
            customerPortalLoginPage.LogIn(userMail, userCustomerPortalPassword);
            bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
            bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();
            Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
            Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page.");
            var customerPortalCustomerOrdersPage = customerPortalLoginPage.ClickOnCustomerOrders();
            //2.Cliquer sur "+"
            var customerPortalCustmerOrdersPageModal = customerPortalCustomerOrdersPage.ClickAdd();
            //2.Remplir le formulaire avec le site(BCN), le customer(EMIRATES AILRINES) et l'aircraft (B777)
            customerPortalCustmerOrdersPageModal.SelectSite(site);
            customerPortalCustmerOrdersPageModal.SelectCustomer(customer);
            customerPortalCustmerOrdersPageModal.SelectDeliveryName(deliveryName);
            //3.Cliquer sur "create"
            CustomerPortalCustomerOrdersPageResult customerPortalCustomerOrdersPageResult = customerPortalCustmerOrdersPageModal.Create();
            //4.Cliquer sur l'onglet 'General Information'
            customerPortalCustomerOrdersPageResult.GoToGeneralInformation();
            customerPortalCustomerOrdersPageResult.SelectAirCraft(aircraft);
            //5.Modifier des informations 'Delivery date', 'Hour', 'Delivery', 'Delivery comment', 'Comment"
            string dd = DateUtils.Now.AddDays(5).ToString("dd/MM/yyyy");
            string h = "16:00";
            string d = "MyDelivery";
            string dc = "My Delivery Comment!";
            string c = "My comment.";
            customerPortalCustomerOrdersPageResult.GeneralInfoFill("creationDatePicker", dd, true);
            customerPortalCustomerOrdersPageResult.GeneralInfoFill("hour", h, false);
            customerPortalCustomerOrdersPageResult.GeneralInfoFill("DeliveryName", d, false);
            customerPortalCustomerOrdersPageResult.GeneralInfoFill("DeliveryComment", dc, false);
            // click pour sortir du calendrier
            customerPortalCustomerOrdersPageResult.GeneralInfoFill("OrderComment", c, true);
            //6.Aller sur l'onglet 'Details' et revenir sur l'onglet 'General Information'
            customerPortalCustomerOrdersPageResult.GoToDetails();
            customerPortalCustomerOrdersPageResult.GoToGeneralInformation();
            //7.Vérifier les données entrées
            Assert.AreEqual(dd, customerPortalCustomerOrdersPageResult.GetGeneralInformation("creationDatePicker", true), "mauvaise valeur de champs DeliveryDate");
            Assert.AreEqual(h, customerPortalCustomerOrdersPageResult.GetGeneralInformation("hour", true), "mauvaise valeur de champs Hour");
            Assert.AreEqual(d, customerPortalCustomerOrdersPageResult.GetGeneralInformation("DeliveryName", true), "mauvaise valeur de champs Delivery");
            Assert.AreEqual(dc, customerPortalCustomerOrdersPageResult.GetGeneralInformation("DeliveryComment", false).Trim(), "mauvaise valeur de champs Delivery Comment");
            Assert.AreEqual(c, customerPortalCustomerOrdersPageResult.GetGeneralInformation("OrderComment", false), "mauvaise valeur de champs Comment");

            if (isUserLogged)
            {
                customerPortalLoginPage.LogOut();
            }
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_CO_General_information_Modification_Delivery()
        {
            //1.Aller sur Customer Portal, se logguer, aller sur Customer Orders
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userCustomerPortalPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();

            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString(); //BCN
            string deliveryName = new Random().Next().ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();

            // Arrange
            CustomerPortalLoginPage customerPortalLoginPage = LogInCustomerPortal();
            customerPortalLoginPage.LogIn(userMail, userCustomerPortalPassword);

            bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
            bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();

            Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
            Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page.");

            var customerPortalCustomerOrdersPage = customerPortalLoginPage.ClickOnCustomerOrders();

            //2.Cliquer sur "+"
            var customerPortalCustmerOrdersPageModal = customerPortalCustomerOrdersPage.ClickAdd();

            //2.Remplir le formulaire avec le site(BCN), le customer(EMIRATES AILRINES) et l'aircraft (B777)
            customerPortalCustmerOrdersPageModal.SelectSite(site);
            customerPortalCustmerOrdersPageModal.SelectCustomer(customer);
            customerPortalCustmerOrdersPageModal.SelectDeliveryName(deliveryName);
            customerPortalCustmerOrdersPageModal.SelectAircraft(aircraft);
            //3.Cliquer sur "create"
            CustomerPortalCustomerOrdersPageResult customerPortalCustomerOrdersPageResult = customerPortalCustmerOrdersPageModal.Create();
            //4.Cliquer sur l'onglet 'General Information'
            customerPortalCustomerOrdersPageResult.GoToGeneralInformation();
            //5.Modifier des informations 'Delivery date', 'Hour', 'Delivery', 'Delivery comment', 'Comment"
            string deliverydate = DateUtils.Now.AddDays(5).ToString("dd/MM/yyyy");
            string hour = "16:00";
            string delivery = "MyDelivery";
            string deliverycomment = "My Delivery Comment!";
            string comment = "My comment.";
            customerPortalCustomerOrdersPageResult.GeneralInfoFill("creationDatePicker", deliverydate, true);
            customerPortalCustomerOrdersPageResult.GeneralInfoFill("hour", hour, false);
            customerPortalCustomerOrdersPageResult.GeneralInfoFill("DeliveryName", delivery, false);
            customerPortalCustomerOrdersPageResult.GeneralInfoFill("DeliveryComment", deliverycomment, false);
            // click pour sortir du calendrier
            customerPortalCustomerOrdersPageResult.GeneralInfoFill("OrderComment", comment, true);
            customerPortalCustomerOrdersPageResult.GoToDetails();
            customerPortalCustomerOrdersPageResult.AddLine("CHICKEN SALAD", 10);
            customerPortalCustomerOrdersPageResult.WaitLoading();
            customerPortalCustomerOrdersPageResult.Validate();
            //6.Aller sur l'onglet 'Details' et revenir sur l'onglet 'General Information'
            customerPortalCustomerOrdersPageResult.WaitLoading();
            customerPortalCustomerOrdersPageResult.GoToGeneralInformation();
            customerPortalCustomerOrdersPageResult.WaitLoading();
            //7.Vérifier les données entrées, le champ 'Delivery' n'est pas modifiable
            string xPathDeliveryDate = "//*[contains(text(),'Delivery date')]/following::div";
            Assert.AreEqual(deliverydate, customerPortalCustomerOrdersPageResult.WaitForElementIsVisible(By.XPath(xPathDeliveryDate)).Text, "mauvaise valeur de champs DeliveryDate");
            string xPathDelivery = "//*[normalize-space(text())='Delivery']/following::div/span";
            Assert.AreEqual(delivery, customerPortalCustomerOrdersPageResult.WaitForElementIsVisible(By.XPath(xPathDelivery)).Text, "mauvaise valeur de champs Delivery");
            string xPathDeliveryComment = "//*[normalize-space(text())='Delivery comment']/following::div/textarea";
            Assert.AreEqual(deliverycomment, customerPortalCustomerOrdersPageResult.WaitForElementIsVisible(By.XPath(xPathDeliveryComment)).Text, "mauvaise valeur de champs Delivery Comment");
            string xPathComment = "//*[normalize-space(text())='Comment']/following::div/textarea";
            Assert.AreEqual(comment, customerPortalCustomerOrdersPageResult.WaitForElementIsVisible(By.XPath(xPathComment)).Text, "mauvaise valeur de champs Comment");

            var Delivery = customerPortalCustomerOrdersPageResult.WaitForElementIsVisible(By.XPath(xPathDeliveryComment));

            Assert.AreEqual("true", Delivery.GetAttribute("disabled"), "DeliveryComment not readonly");

            var Comment = customerPortalCustomerOrdersPageResult.WaitForElementIsVisible(By.XPath(xPathComment));
            Assert.AreEqual("true", Comment.GetAttribute("disabled"), "DeliveryComment not readonly");

            if (isUserLogged)
            {
                customerPortalLoginPage.LogOut();
            }
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_CO_Filter_OrderStatus_Invoiced()
        {
            string greenColor = "rgba(0, 128, 0, 1)";
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userCustomerPortalPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString(); //BCN
            string deliveryName = new Random().Next().ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            bool stepEnd = false;

            var homePage = LogInAsAdmin();
            var customerPortalLoginPage = homePage.GoToCustomerPortal();
            try
            {
                customerPortalLoginPage.LogIn(userMail, userCustomerPortalPassword);
                bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
                bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();
                Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
                Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page.");

                CustomerPortalCustomerOrdersPage customerPortalCustomerOrdersPage = customerPortalLoginPage.ClickOnCustomerOrders();
                var customerPortalCustmerOrdersPageModal = customerPortalCustomerOrdersPage.ClickAdd();
                customerPortalCustmerOrdersPageModal.SelectSite(site);
                customerPortalCustmerOrdersPageModal.SelectCustomer(customer);
                customerPortalCustmerOrdersPageModal.SelectDeliveryName(deliveryName);
                customerPortalCustmerOrdersPageModal.SelectAircraft(aircraft);
                CustomerPortalCustomerOrdersPageResult customerPortalCustomerOrdersPageResult = customerPortalCustmerOrdersPageModal.Create();
                customerPortalCustomerOrdersPageResult.AddLine("HOT MEAL YC UAE", 10);
                customerPortalCustomerOrdersPageResult.WaitLoading();
                customerPortalCustomerOrdersPageResult.Validate();
                var resultText = customerPortalCustomerOrdersPageResult.ResultCustomerOrder();
                Regex r = new Regex("Customer order ([0-9]+) - EMIRATES AIRLINES");
                Match m = r.Match(resultText);
                string CustomerOrder = m.Groups[1].Value;
                customerPortalCustomerOrdersPage.Close();
                InvoicesPage invoicesPage = homePage.GoToAccounting_InvoicesPage();
                AutoInvoiceCreateModalPage autoInvoiceCreateModalPage = invoicesPage.AutoInvoiceCreatePage();
                autoInvoiceCreateModalPage.FillField_CreateNewAutoInvoice(customer, site, CustomerPickMethod.AllCustomerOrdersInSelectedPeriod);
                InvoiceDetailsPage invoiceDetailsPage = autoInvoiceCreateModalPage.FillFieldSelectAll();
                if (invoicesPage.CheckIfExistValidate())
                {
                    invoiceDetailsPage.Validate();
                }
                else
                {
                    invoiceDetailsPage = invoicesPage.ClickFirstLine();
                    invoiceDetailsPage.Validate();
                }
                homePage.Navigate();
                customerPortalLoginPage = homePage.GoToCustomerPortal();
                customerPortalCustomerOrdersPage = customerPortalLoginPage.ClickOnCustomerOrders();
                customerPortalCustomerOrdersPage.Search(CustomerOrder);
                customerPortalCustomerOrdersPage.ShowFilterCustomersOrder("invoiced");
                customerPortalCustomerOrdersPage.Search(CustomerOrder);
                bool isAdded = customerPortalCustomerOrdersPage.CheckTotalNumber() != 0;
                Assert.IsTrue(isAdded, $"there is no data with the current filter params search : {CustomerOrder} , status : invoiced");
                string isTheAddedOrder = customerPortalCustomerOrdersPage.getTableFirstLine(1);
                bool exist = isTheAddedOrder.Contains(CustomerOrder);
                Assert.IsTrue(exist, $"Le CustomerOrder ({CustomerOrder}) n'est pas ajouté.");
                bool isValidated = customerPortalCustomerOrdersPage.CheckValidationIcon();
                Assert.IsTrue(isValidated, $"CustomerOrder ({CustomerOrder}) n'est pas valide");
                bool isDollarExists = customerPortalCustomerOrdersPage.IsDollarIconExist();
                Assert.IsTrue(isDollarExists, "l'icon dollar n'existe pas");
                bool isValidate = customerPortalCustomerOrdersPage.CheckDollarIconColor(greenColor);
                Assert.IsTrue(isValidate, "l'icon dollar n'est pas verte");

                if (isUserLogged)
                {
                    customerPortalLoginPage.LogOut();
                }
                stepEnd = true;
            }
            finally
            {
                if (stepEnd)
                {
                    customerPortalLoginPage.Close();
                }
            }
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_CO_Filter_OrderStatus_InProgress()
        {
            //(test identique à PR_CO_Filter_Show_InvoiceStep_InProgress mais à créer sur Customer Portal)
            //1.Aller sur Customer Portal, se logguer, aller sur Customer Orders
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userCustomerPortalPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();

            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string deliveryName = new Random().Next().ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();


            // Arrange
            CustomerPortalLoginPage customerPortalLoginPage = LogInCustomerPortal();
            customerPortalLoginPage.LogIn(userMail, userCustomerPortalPassword);

            bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
            bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();

            Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
            Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page.");

            var customerPortalCustomerOrdersPage = customerPortalLoginPage.ClickOnCustomerOrders();
            //2.Créer un Customer order
            //Cliquer sur "+"
            var customerPortalCustmerOrdersPageModal = customerPortalCustomerOrdersPage.ClickAdd();
            //Remplir le formulaire avec le site(BCN), le customer(EMIRATES AILRINES) et l'aircraft (B777)
            customerPortalCustmerOrdersPageModal.SelectSite(site);
            customerPortalCustmerOrdersPageModal.SelectCustomer(customer);
            customerPortalCustmerOrdersPageModal.SelectDeliveryName(deliveryName);
            customerPortalCustmerOrdersPageModal.SelectAircraft(aircraft);
            //Cliquer sur "create"
            CustomerPortalCustomerOrdersPageResult customerPortalCustomerOrdersPageResult = customerPortalCustmerOrdersPageModal.Create();
            // ajouter une ligne pour l'export Excel
            customerPortalCustomerOrdersPageResult.AddLine("HOT MEAL YC UAE", 10);
            // récupérer le CustomerOrder
            var resultText = customerPortalCustomerOrdersPageResult.ResultCustomerOrder();
            Regex r = new Regex("Customer order ([0-9]+) - EMIRATES AIRLINES");
            Match m = r.Match(resultText);
            string CustomerOrder = m.Groups[1].Value;
            Thread.Sleep(4000);
            //3.Faire 'Back To List'
            // back to list
            customerPortalCustomerOrdersPage.Navigate("/CustomerPortal/CustomerOrders");

            //4.Filtrer 'ORDER STATUS' à 'IN PROGRESS'
            customerPortalCustomerOrdersPage.Search(CustomerOrder);
            customerPortalCustomerOrdersPage.ShowFilterCustomersOrder("inProgress");
            // rafraichisement du tableau
            Thread.Sleep(4000);
            //5.Vérifier que le customer order crée est bien affiché et n'a pas d'icone validation vert à gauche, ni d'icone $
            Assert.AreEqual("OK", customerPortalCustomerOrdersPage.CheckIconsInProgress());


            //6.Faire un export excel de ce customer
            // purger le dossier Download
            string directory = TestContext.Properties["DownloadsPath"].ToString();
            foreach (string f in Directory.GetFiles(directory))
            {
                //export-orders 17-12-2021_2-33-56 PM.xlsx
                if (f.Contains("export-orders") && f.EndsWith(".xlsx"))
                {
                    File.Delete(f);
                }
            }

            //Cliquer sur la flèche dirigée vers le bas, puis Export
            customerPortalCustomerOrdersPage.ExportExcel();

            // barre de progression
            Thread.Sleep(4000);

            //Vérifier les informations : order ID, order Date, Site, Customer, Price, Detail name, quantity, unit
            string trouve = null;
            foreach (string f in Directory.GetFiles(directory))
            {
                //export-orders 17-12-2021_2-33-56 PM.xlsx
                if (f.Contains("export-orders") && f.EndsWith(".xlsx"))
                {
                    trouve = f;
                }
            }
            Assert.IsNotNull(trouve);

            int resultNumber = OpenXmlExcel.GetExportResultNumber("Sheet 1", trouve);

            //7.Vérifier dans le fichier d'export excel que le customer order donné est à 'InProgress' dans la colonne 'Invoice step'
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.AreEqual(CustomerOrder, OpenXmlExcel.GetValuesInList("Order ID", "Sheet 1", trouve)[0], " excel colonne Order ID");
            Assert.AreEqual("InProgress", OpenXmlExcel.GetValuesInList("Invoice step", "Sheet 1", trouve)[0], " excel colonne Invoice step");


            if (isUserLogged)
            {
                customerPortalLoginPage.LogOut();
            }
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_CO_Filter_InvoiceStep_All()
        {
            //(test identique à PR_CO_Filter_Show_InvoiceStep_All mais à créer sur Customer Portal)
            //1.Aller sur Customer Portal, se logguer, aller sur Customer Orders
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userCustomerPortalPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();

            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string deliveryName = new Random().Next().ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();


            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var customerPortalLoginPage = homePage.GoToCustomerPortal();
            try
            {
                customerPortalLoginPage.LogIn(userMail, userCustomerPortalPassword);

                bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
                bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();

                Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
                Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page.");

                CustomerPortalCustomerOrdersPage customerPortalCustomerOrdersPage = customerPortalLoginPage.ClickOnCustomerOrders();

                //2.Vérifier si 1 customer order à chaque invoice step existe sinon en crée(minimum 1 In Progress, 1 Validated et 1 Invoiced)
                // Créer un Customer order "In Progress"
                //Cliquer sur "+"
                var customerPortalCustmerOrdersPageModal = customerPortalCustomerOrdersPage.ClickAdd();
                //Remplir le formulaire avec le site(BCN), le customer(EMIRATES AILRINES) et l'aircraft (B777)
                customerPortalCustmerOrdersPageModal.SelectSite(site);
                customerPortalCustmerOrdersPageModal.SelectCustomer(customer);
                customerPortalCustmerOrdersPageModal.SelectDeliveryName(deliveryName);
                customerPortalCustmerOrdersPageModal.SelectAircraft(aircraft);
                //Cliquer sur "create"
                CustomerPortalCustomerOrdersPageResult customerPortalCustomerOrdersPageResult = customerPortalCustmerOrdersPageModal.Create();
                // ajouter une ligne pour l'export Excel
                customerPortalCustomerOrdersPageResult.AddLine("HOT MEAL YC UAE", 10);
                // récupérer le CustomerOrder
                var resultText = customerPortalCustomerOrdersPageResult.ResultCustomerOrder();
                Regex r = new Regex("Customer order ([0-9]+) - EMIRATES AIRLINES");
                Match m = r.Match(resultText);
                string CustomerOrder1 = m.Groups[1].Value;
                // back to list
                customerPortalCustomerOrdersPage.Navigate("/CustomerPortal/CustomerOrders");

                // Créer un Customer order "Invoiced"
                //Cliquer sur "+"
                customerPortalCustmerOrdersPageModal = customerPortalCustomerOrdersPage.ClickAdd();
                //Remplir le formulaire avec le site(BCN), le customer(EMIRATES AILRINES) et l'aircraft (B777)
                customerPortalCustmerOrdersPageModal.SelectSite(site);
                customerPortalCustmerOrdersPageModal.SelectCustomer(customer);
                customerPortalCustmerOrdersPageModal.SelectDeliveryName(deliveryName);
                customerPortalCustmerOrdersPageModal.SelectAircraft(aircraft);
                //Cliquer sur "create"
                customerPortalCustomerOrdersPageResult = customerPortalCustmerOrdersPageModal.Create();
                // ajouter une ligne pour l'export Excel
                customerPortalCustomerOrdersPageResult.AddLine("HOT MEAL YC UAE", 10);
                customerPortalCustomerOrdersPageResult.Validate();
                // récupérer le CustomerOrder
                resultText = customerPortalCustomerOrdersPageResult.ResultCustomerOrder();
                r = new Regex("Customer order ([0-9]+) - EMIRATES AIRLINES");
                m = r.Match(resultText);
                string CustomerOrder2 = m.Groups[1].Value;
                // back to list
                customerPortalCustomerOrdersPage.Navigate("/CustomerPortal/CustomerOrders");
                customerPortalCustomerOrdersPage.Close();
                InvoicesPage invoicesPage = homePage.GoToAccounting_InvoicesPage();
                //Créer un Invoice Auto(BCN, All customer order, EMIRATES AILRINES, customer order crée) et le valider
                AutoInvoiceCreateModalPage autoInvoiceCreateModalPage = invoicesPage.AutoInvoiceCreatePage();
                autoInvoiceCreateModalPage.FillField_CreateNewAutoInvoice(customer, site, CustomerPickMethod.AllCustomerOrdersInSelectedPeriod); //CustomerOrder
                autoInvoiceCreateModalPage.FillFieldSelectAll();
                // on s'est perdu entre Invoice Details et Invoice index
                homePage.GoToWinrestHome();
                invoicesPage = homePage.GoToAccounting_InvoicesPage();
                invoicesPage.SelectFirstInvoice();
                invoicesPage.Validate();
                //Revenir sur Customer Portal, sur Customer orders
                homePage.GoToWinrestHome();
                customerPortalLoginPage = homePage.GoToCustomerPortal();
                customerPortalCustomerOrdersPage = customerPortalLoginPage.ClickOnCustomerOrders();

                // Créer un Customer order "Validated"
                //Cliquer sur "+"
                customerPortalCustmerOrdersPageModal = customerPortalCustomerOrdersPage.ClickAdd();
                //Remplir le formulaire avec le site(BCN), le customer(EMIRATES AILRINES) et l'aircraft (B777)
                customerPortalCustmerOrdersPageModal.SelectSite(site);
                customerPortalCustmerOrdersPageModal.SelectCustomer(customer);
                customerPortalCustmerOrdersPageModal.SelectDeliveryName(deliveryName);
                customerPortalCustmerOrdersPageModal.SelectAircraft(aircraft);
                //Cliquer sur "create"
                customerPortalCustomerOrdersPageResult = customerPortalCustmerOrdersPageModal.Create();
                // ajouter une ligne pour l'export Excel
                customerPortalCustomerOrdersPageResult.AddLine("HOT MEAL YC UAE", 10);
                customerPortalCustomerOrdersPageResult.Validate();
                // récupérer le CustomerOrder
                resultText = customerPortalCustomerOrdersPageResult.ResultCustomerOrder();
                r = new Regex("Customer order ([0-9]+) - EMIRATES AIRLINES");
                m = r.Match(resultText);
                string CustomerOrder3 = m.Groups[1].Value;
                // back to list
                customerPortalCustomerOrdersPage.Navigate("/CustomerPortal/CustomerOrders");





                Assert.AreEqual("Id " + CustomerOrder1, customerPortalCustomerOrdersPage.getTableFirstLine(1, 3));
                Assert.AreEqual("Id " + CustomerOrder2, customerPortalCustomerOrdersPage.getTableFirstLine(1, 2));
                Assert.AreEqual("Id " + CustomerOrder3, customerPortalCustomerOrdersPage.getTableFirstLine(1, 1));

                // Toggle "ORDER STATUS"
                customerPortalCustomerOrdersPage.ShowFilterCustomersOrder();

                //3.Filtrer 'ORDER STATUS' à 'ALL' et noter le nombre total de customer orders
                customerPortalCustomerOrdersPage.ShowFilterCustomersOrder("all", false);
                int nbAll = customerPortalCustomerOrdersPage.CheckTotalNumber();
                Assert.AreNotEqual(0, nbAll, "pas de CUSTOMER ORDER all");

                //4.Filtrer 'ORDER STATUS' à 'IN PROGRESS' et noter le nombre total de customer orders In progress
                customerPortalCustomerOrdersPage.ShowFilterCustomersOrder("inProgress", false);
                customerPortalCustomerOrdersPage.WaitPageLoading();
                int nbInProgress = customerPortalCustomerOrdersPage.CheckTotalNumber();
                Assert.AreNotEqual(0, nbInProgress, "pas de CUSTOMER ORDER inProgress");

                //5.Filtrer 'ORDER STATUS' à 'VALIDATED' et noter le nombre total de customer orders Validated
                customerPortalCustomerOrdersPage.ShowFilterCustomersOrder("validated", false);
                customerPortalCustomerOrdersPage.WaitPageLoading();
                int nbValidated = customerPortalCustomerOrdersPage.CheckTotalNumber();
                Assert.AreNotEqual(0, nbValidated, "pas de CUSTOMER ORDER validated");

                //6.Filtrer 'ORDER STATUS' à 'INVOICED' et noter le nombre total de customer orders Invoiced
                customerPortalCustomerOrdersPage.ShowFilterCustomersOrder("invoiced", false);
                customerPortalCustomerOrdersPage.WaitPageLoading();
                int nbInvoiced = customerPortalCustomerOrdersPage.CheckTotalNumber();
                Assert.AreNotEqual(0, nbInvoiced, "pas de CUSTOMER ORDER invoiced");

                //7.Vérifier le total des 3 = à ALL
                // nbValidated (v) comprend les non invoice ($ jaune)
                // nbInvoiced (v) comprend les invoice ($ vert)
                Assert.AreEqual(nbAll, nbInProgress + nbValidated + nbInvoiced, "mauvais nb valeurs nbAll(" + nbAll + ")=nbInProgress(" + nbInProgress + ")+nbValidated(" + nbValidated + ")+nbInvoiced(" + nbInvoiced + ")");

            }
            finally
            {
                bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);

                if (isUserLogged)
                {
                    customerPortalLoginPage.LogOut();
                }
            }
        }

        [DeploymentItem("Resources\\pizza.png")]
        [DeploymentItem("Resources\\pizza_petite.png")]
        [DeploymentItem("chromedriver.exe")]
        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_CO_PrintLogo()
        {
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string deliveryName = new Random().Next().ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            var CustomerOrder = string.Empty;



            //Arrange

            var homePage = LogInAsAdmin();
            // Etre dans un ""customer order"".
            // Mettre un logo sur le customer dans les settings
            CustomerPage customerPage = homePage.GoToCustomers_CustomerPage();
            customerPage.Filter(CustomerPage.FilterType.Search, customer);
            CustomerGeneralInformationPage generalInfo = customerPage.SelectFirstCustomer();

            // serialize + sha1sum (filesize)
            FileInfo fiUpload = new FileInfo(TestContext.TestDeploymentDir + "\\pizza.png");
            generalInfo.UploadLogo(fiUpload);

            homePage.Navigate();

            //Act
            var customerPortalLoginPage = homePage.GoToCustomerPortal();
            try
            {
                customerPortalLoginPage.LogIn(userMail, userPassword);
                bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
                bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();
                Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
                Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page.");
                var customerPortalCustomerOrdersPage = customerPortalLoginPage.ClickOnCustomerOrders();

                //Créer un Customer order(CUPO_CO_CreateNewCustomerOrder)
                var customerPortalCustmerOrdersPageModal = customerPortalCustomerOrdersPage.ClickAdd();
                //Remplir le formulaire avec le site(BCN), le customer(EMIRATES AILRINES) et l'aircraft (B777)
                customerPortalCustmerOrdersPageModal.SelectSite(site);
                customerPortalCustmerOrdersPageModal.SelectCustomer(customer);
                customerPortalCustmerOrdersPageModal.SelectDeliveryName(deliveryName);
                customerPortalCustmerOrdersPageModal.SelectAircraft(aircraft);
                //Cliquer sur "create"
                CustomerPortalCustomerOrdersPageResult customerPortalCustomerOrdersPageResult = customerPortalCustmerOrdersPageModal.Create();

                //Cliquer sur '+'
                //Ajouter un item et une quantity(HOT MEAL YC UAE et 10)
                customerPortalCustomerOrdersPageResult.AddLine("HOT MEAL YC UAE", 10);

                // récupérer le CustomerOrder
                var resultText = customerPortalCustomerOrdersPageResult.ResultCustomerOrder();
                Regex r = new Regex("Customer order ([0-9]+) - EMIRATES AIRLINES");
                Match m = r.Match(resultText);
                 CustomerOrder = m.Groups[1].Value;

                //Faire un 'Back to list'
                customerPortalCustomerOrdersPage = customerPortalCustomerOrdersPageResult.BackToList();

                //Rechercher le customer order crée
                customerPortalCustomerOrdersPage.Search(CustomerOrder);

                string directory = TestContext.Properties["DownloadsPath"].ToString();
                foreach (var files in new DirectoryInfo(downloadsPath).GetFiles())
                {
                    files.Delete();
                }
                string file = Path.Combine(directory, "ORDER_" + CustomerOrder + ".pdf");

                CustomerPortalCustomerOrdersPageResult customerOrder = customerPortalCustomerOrdersPage.GoToFirstCustomer();
                customerOrder.DownloadPDFFromEmail();

                //8.Vérifier que le rapport est généré
                customerOrder.WaitPageLoading();
                FileInfo fi = new FileInfo(file);
                customerOrder.WaitPageLoading();
                fi.Refresh();
                long length = fi.Length;
                if (length > 1024)
                {
                    Assert.IsTrue(length > 1024, "Fichier généré vide");
                }
                else
                {
                    Assert.Fail("Fichier généré " + file + " inexistant");
                }

                // 2.Verifier le logo (dans le Pdf du mail)
                PdfDocument document = PdfDocument.Open(fi.FullName);
                Page page = document.GetPage(1);
                List<IPdfImage> images = page.GetImages().ToList<IPdfImage>();
                Assert.AreEqual(1, images.Count, "Pas d'image dans le PDF");
                


                if (isUserLogged)
                {
                    customerPortalLoginPage.LogOut();
                }
            }
            finally
            {
                customerPortalLoginPage.Close();
                homePage.Navigate();
                CustomerOrderPage customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, CustomerOrder);
                customerOrderPage.DeleteCustomerOrder();


            }
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_QBW_Export()
        {
            // Prepare
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string siteMAD = TestContext.Properties["SiteMAD"].ToString();
            string siteACE = TestContext.Properties["SiteACE"].ToString();

            DeleteAllFileDownload();
            var homePage = LogInAsAdmin();

            CustomerPortalLoginPage customerPortalLoginPage = LogInCustomerPortal();
            customerPortalLoginPage.LogIn(userMail, userPassword);

            bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
            bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();

            Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
            Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page Quantity By Week.");

            var customerPortalQtyByWeekPage = customerPortalLoginPage.ClickOnQtyByWeek();
            //Avoir des données disponibles. (fonctionne sur dev5 site PMI)
            var lignes = WebDriver.FindElements(By.XPath("//*/tr[string(@data-service-id)]"));
            Assert.IsTrue(lignes.Count > 0, "pas de données disponibles");

            customerPortalQtyByWeekPage.Export(siteACE);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();
            var correctDownloadedFile = customerPortalQtyByWeekPage.GetQMExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

            // Récupération du nom du fichier et construction de l'URL du fichier Excel à ouvrir   
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //Assert
            int resultNumber = OpenXmlExcel.GetExportResultNumber("Quantity to produce", filePath);
            resultNumber += OpenXmlExcel.GetExportResultNumber("Total invoice by customer", filePath);

            customerPortalQtyByWeekPage.Export(siteMAD);//site ACE sans delivery 

            // On récupère les fichiers du répertoire de téléchargement
            taskDirectory = new DirectoryInfo(downloadsPath);
            taskFiles = taskDirectory.GetFiles();
            correctDownloadedFile = customerPortalQtyByWeekPage.GetQMExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

            // Récupération du nom du fichier et construction de l'URL du fichier Excel à ouvrir   
            fileName = correctDownloadedFile.Name;
            filePath = Path.Combine(downloadsPath, fileName);

            //Assert
            resultNumber = OpenXmlExcel.GetExportResultNumber("Quantity to produce", filePath);
            Assert.AreEqual(lignes.Count, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
            resultNumber += OpenXmlExcel.GetExportResultNumber("Total invoice by customer", filePath);
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
            
            List<string> customerNames = OpenXmlExcel.GetValuesInList("CustomerName", "Quantity to produce", filePath);
            List<string> deliverys = OpenXmlExcel.GetValuesInList("Delivery", "Quantity to produce", filePath);
            var counter = customerPortalQtyByWeekPage.GetCounter();
            Assert.AreEqual(counter, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);

            // Arrange
            homePage.Navigate();
            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteMAD);
            dispatchPage.FilterCustomersUnCheckAllCustomer();
            dispatchPage.FilterDeliverysUnCheckAllDelivery();

            dispatchPage.FilterCustomers(customerNames);
            dispatchPage.FilterDeliverys(deliverys);
            dispatchPage.WaitPageLoading();

            List<string> listOfDispatchPrevisionalQty = dispatchPage.GetTotalListDeliveryPrevisionalQty();
            List<string> listOfDispatchQtyToProduce = dispatchPage.GetTotalListDeliveryQtyToProduce();
            List<string> listOfDispatchQtyToInvoice = dispatchPage.GetTotalListDeliveryQtyToInvoice();
            dispatchPage.WaitPageLoading();
            bool verifDeliveryPrevisionalQuantity = dispatchPage.VerifyDispatchInDeliveries(deliverys, listOfDispatchPrevisionalQty);
            Assert.IsTrue(verifDeliveryPrevisionalQuantity, "Les éléments dans l'excel ne correspond pas à celui de listOfDispatch");

            var quantityToProducePage = dispatchPage.ClickQuantityToProduce();
            quantityToProducePage.WaitPageLoading();
            bool verifDeliveryQuantityToProduce = dispatchPage.VerifyDispatchInDeliveries(deliverys, listOfDispatchQtyToProduce);
            Assert.IsTrue(verifDeliveryQuantityToProduce, "Les éléments dans l'excel ne correspond pas à celui de listOfDispatch");

            var qtytoInvoice = dispatchPage.ClickQuantityToInvoice();
            qtytoInvoice.WaitPageLoading();
            bool verifDeliveryQuantityToInvoice = dispatchPage.VerifyDispatchInDeliveries(deliverys, listOfDispatchQtyToInvoice);
            Assert.IsTrue(verifDeliveryQuantityToInvoice, "Les éléments dans l'excel ne correspond pas à celui de listOfDispatch");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_QBW_DuplicateDay()
        {
            // Prepare
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();
            string qty = "27";
            string comment = "blah";

            //. Avoir un compte sur le customer portal
            CustomerPortalLoginPage customerPortalLoginPage = LogInCustomerPortal();
            customerPortalLoginPage.LogIn(userMail, userPassword);

            bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
            bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();

            Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
            Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page Quantity By Week.");

            //. Avoir des delivery associés au compte
            var customerPortalQtyByWeekPage = customerPortalLoginPage.ClickOnQtyByWeek();
            //Avoir des données disponibles. (fonctionne sur dev5 site PMI)
            var lignes = WebDriver.FindElements(By.XPath("//*/tr[string(@data-service-id)]"));
            Assert.IsTrue(lignes.Count > 0, "pas de données disponibles");

            //Modifier la quantité d'une journée
            customerPortalQtyByWeekPage.Search("ServiceForCustomerPortal");
            //ajouter une étape où on valide la quantité de demain puis on change la quantité et revalider, vérifier
            //sauf si c'est dimanche
            CultureInfo.CurrentCulture = new CultureInfo("en-US", false);
            var dateToUpdate = DateUtils.Now.AddDays(1);

            if (DateUtils.Now.AddDays(1).ToString("dddd") != "Sunday")
            {
                customerPortalQtyByWeekPage.SetQuantities(dateToUpdate.ToString("dddd"), qty);
                customerPortalQtyByWeekPage.Validate((int)dateToUpdate.DayOfWeek);

                //et ajouter un commentaire, valider
                customerPortalQtyByWeekPage.SetComment(dateToUpdate.ToString("dddd"), comment);
                customerPortalQtyByWeekPage.Validate((int)dateToUpdate.DayOfWeek);

                //Cliquer sur icone dupliquer(à droite du jour) et sélectionner le jour à dupliquer
                customerPortalQtyByWeekPage.DuplicateOneDay(dateToUpdate.ToString("ddd"), (int)dateToUpdate.DayOfWeek);

                //Vérifier que la quantité et le commentaire ont bien été dupliqué
                //sur le jour sélectionné
                customerPortalQtyByWeekPage.Search("ServiceForCustomerPortal");
                Assert.AreEqual(qty, customerPortalQtyByWeekPage.GetQuantity(DateUtils.Now.AddDays(2).ToString("dddd").ToLower()), "Qty non dupliqué");
                Assert.AreEqual(comment, customerPortalQtyByWeekPage.GetComment(DateUtils.Now.AddDays(2).ToString("dddd").ToLower()), "Commentaire non dupliqué");
            }
        }

        [Ignore]
        [Timeout(_timeout)]
		[TestMethod]
        public void imapYOPmail()
        {
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["MailBox_UserPassword"].ToString();

            bool doSmtp = false;
            bool doImap = true;
            // If using Professional version, put your serial key below.
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");

            Console.WriteLine(DateUtils.Now.ToString("dd/MM/yyyy hh:mm:ss"));
            if (doSmtp)
            {
                // Create new email message.
                MailMessage message = new MailMessage(
                    new MailAddress("sender@example.com", "Sender"),
                    new MailAddress("test@test.fr", "First receiver"));

                // Add subject and body.
                message.Subject = "Send Email in C# / VB.NET / ASP.NET";
                message.BodyText = "Hi 👋,\n" +
                    "This message was created and sent with GemBox.Email.\n" +
                    "Read more about it on https://www.gemboxsoftware.com/email";

                // Create new SMTP client and send an email message.
                using (SmtpClient smtp = new SmtpClient("127.0.0.1"))
                //using (SmtpClient smtp = new SmtpClient("smtp.gmail.com"))
                {
                    smtp.Connect();
                    smtp.Authenticate("test@test.fr", "test");
                    //smtp.Authenticate(userMail, userPassword);
                    smtp.SendMessage(message);
                }
            }
            if (doImap)
            {
                // Create new IMAP client.
                using (var imap = new ImapClient("127.0.0.1"))
                //using (var imap = new ImapClient("mail-imap.yopmail.com", 993,GemBox.Email.Security.ConnectionSecurity.Ssl))
                //using (var imap = new ImapClient("imap.gmail.com"))
                {
                    // By default the connect timeout is 5 sec.
                    imap.ConnectTimeout = TimeSpan.FromSeconds(4);

                    // Connect to IMAP server.
                    imap.Connect();
                    Console.WriteLine("Connected.");

                    // Authenticate using the credentials; username and password.
                    imap.Authenticate("test@test.fr", "test");
                    //imap.Authenticate("testauto.newrest@yopmail.com", "test");
                    //imap.Authenticate(userMail, userPassword);
                    Console.WriteLine("Authenticated.");

                    // Select INBOX folder.
                    imap.SelectInbox();

                    // Get information about all available messages in INBOX folder.
                    ReadOnlyCollection<string> messagesUid = imap.SearchMessageUids("UNSEEN");
                    IList<ImapMessageInfo> infos = imap.ListMessages();

                    // Display messages information.
                    foreach (string uid in messagesUid)
                    {
                        //Console.WriteLine($"{info.Number} - [{info.Uid}] - {info.Size} Byte(s)");
                        MailMessage m = imap.GetMessage(uid);
                        Console.WriteLine(m.Subject);
                        Console.WriteLine(m.BodyText);
                    }


                }

                Console.WriteLine("Disconnected.");
            }
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_CO_ChangePageSize()
        {
            // Prepare
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userCustomerPortalPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();

            HomePage homePage = LogInAsAdmin();

            //Arrange
            var customerPortalLoginPage = homePage.GoToCustomerPortal();

            try
            {
                customerPortalLoginPage.LogIn(userMail, userCustomerPortalPassword);

                bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
                bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();

                Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
                Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page.");
                // Avoir des CO
                var customerPortalCustomerOrdersPage = customerPortalLoginPage.ClickOnCustomerOrders();
                customerPortalCustomerOrdersPage.ResetFilter();
                //Changer de page size, vérifier le nombre d'éléments affichés
                //Rechanger de page size, vérifier le nombre d'éléments affichés
                int nbLignes = customerPortalCustomerOrdersPage.CheckTotalNumber();
                Assert.IsTrue(nbLignes > 50, "Pas assez de lignes CO");
                customerPortalCustomerOrdersPage.PageSize("50");
                Assert.AreEqual(50, customerPortalCustomerOrdersPage.GetLinesCount());
                customerPortalCustomerOrdersPage.PageSize("30");
                Assert.AreEqual(30, customerPortalCustomerOrdersPage.GetLinesCount());
                customerPortalCustomerOrdersPage.PageSize("16");
                Assert.AreEqual(16, customerPortalCustomerOrdersPage.GetLinesCount());
                customerPortalCustomerOrdersPage.PageSize("8");
                Assert.AreEqual(8, customerPortalCustomerOrdersPage.GetLinesCount());
                if (isUserLogged)
                {
                    customerPortalLoginPage.LogOut();
                }
            }
            finally
            {
                customerPortalLoginPage.Close();
            }

        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_QBD_Filter_Date()
        {
            //1.Aller sur Customer Portal, se logguer, aller sur Quantity By Week
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();

            CustomerPortalLoginPage customerPortalLoginPage = LogInCustomerPortal();
            customerPortalLoginPage.LogIn(userMail, userPassword);

            bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
            bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();

            Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
            Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page Quantity By Day.");

            var customerPortalQtyByDayPage = customerPortalLoginPage.ClickOnQtyByDay();


            //2.Cliquer sur la flèche pour se déplacer au jour suivant
            customerPortalQtyByDayPage.ClickOnNextDay();

            //3.Vérifier que la date de la semaine suivante s'affiche
            // 29/11/2021 To 05/12/2021
            DateTime nextdateCalendar = customerPortalQtyByDayPage.GetCalendarDay();

            Assert.IsNotNull(nextdateCalendar, "parsing date  raté");
            Assert.IsTrue(nextdateCalendar > DateUtils.Now.Date, "Le jour suivant n'apparaît pas.");

            //4.Cliquer sur la flèche pour se déplacer à la semaine précédante
            customerPortalQtyByDayPage.ClickOnPreviousDay();

            //5.Vérifier que la date de la semaine affichée au départ s'affiche
            // 29/11/2021 To 05/12/2021
            DateTime prevdateCalendar = customerPortalQtyByDayPage.GetCalendarDay();
            Assert.IsNotNull(prevdateCalendar, "parsing date raté");
            Assert.IsTrue(prevdateCalendar <= nextdateCalendar, "Le début de cette semaine n'est pas encore présent.");

            if (isUserLogged)
            {
                customerPortalLoginPage.LogOut();
            }

        }
        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_QBD_Filter_SearchDelivery()
        {
            //1.Aller sur Customer Portal, se logguer, aller sur Quantity By Week
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();
            var serviceName = "ServiceForCustomerPortal";
            CustomerPortalLoginPage customerPortalLoginPage = LogInCustomerPortal();
            customerPortalLoginPage.LogIn(userMail, userPassword);

            bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
            bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();

            Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
            Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page Quantity By Day.");

            var customerPortalQtyByDayPage = customerPortalLoginPage.ClickOnQtyByDay();


            //2.Dans le filtre 'DELIVERY', cocher le delivery 'DELIVERYFORCUSTOMERPORTAL
            customerPortalQtyByDayPage.SearchFilterDelivery("DeliveryForCustomerPortal");

            //3.Vérifier que le service 'ServiceForCustomerPortal' est affiché dans le tableau
            Assert.AreEqual(serviceName, customerPortalQtyByDayPage.SearchFirstResult(), "ServiceForCustomerPortal non trouvé dans le tableau");

            //4.Dans le filtre 'DELIVERY', chercher le service 'A.F.A. CASAS'
            customerPortalQtyByDayPage.SearchFilterDelivery("A.F.A. CASAS");

            //5.Vérifier que le message 'No previsional quantity' s'affiche
            Assert.IsTrue(customerPortalQtyByDayPage.IsSearchResultEmpty(), "Pas de message à la place du tableau (vide)");
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_QBD_Filter_SearchServices()
        {
            //1.Aller sur Customer Portal, se logguer, aller sur Quantity By Week
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();
            var serviceName = "ServiceForCustomerPortal";
            var secondServiceName = "ServiceForCustomerPortal2";
            CustomerPortalLoginPage customerPortalLoginPage = LogInCustomerPortal();
            customerPortalLoginPage.LogIn(userMail, userPassword);

            bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
            bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();

            Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
            Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page Quantity By Day.");

            var customerPortalQtyByDayPage = customerPortalLoginPage.ClickOnQtyByDay();
            customerPortalQtyByDayPage.Search(serviceName);

            var serviceExist = customerPortalQtyByDayPage.VerifierServiceExist(serviceName);
            Assert.IsTrue(serviceExist, serviceName + "n'existe pas");

            var secondServiceExist = customerPortalQtyByDayPage.VerifierServiceExist(secondServiceName);
            Assert.IsFalse(secondServiceExist, secondServiceName + "existe");

        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_QBD_SendCommandByMail()
        {
            // Prepare
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userCustomerPortalPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();
            var userPassword = TestContext.Properties["MailBox_UserPassword"].ToString();

            bool customerPortalTab = false;
            bool mailboxTab = false;

            var mailPage = new MailPage(WebDriver, TestContext);


            HomePage homePage = LogInAsAdmin();

            // Go to CustomerPortal Page
            var customerPortalLoginPage = homePage.GoToCustomerPortal();
            customerPortalTab = true;

            try
            {
                customerPortalLoginPage.LogIn(userMail, userCustomerPortalPassword);

                bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
                Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");

                CustomerPortalQtyByDayPage customerPortalQtyByDayPage = customerPortalLoginPage.ClickOnQtyByDay();

                string piecejointe = customerPortalQtyByDayPage.SendQtyByMail(userMail);
                string mailSubject = customerPortalQtyByDayPage.GetMailSubject().Replace("  ", " ");
                customerPortalQtyByDayPage.SendMail();

                WebDriver.Navigate().Refresh();
                // Deconnexion du portal
                customerPortalLoginPage.LogOut();
                customerPortalLoginPage.Close();
                customerPortalTab = false;

                // Connexion à la mailbox pour voir si le mail a été reçu
                mailPage = customerPortalLoginPage.RedirectToOutlookMailbox();
                mailboxTab = true;

                mailPage.FillFields_LogToOutlookMailbox(userMail);

                // Recherche du mail envoyé
                mailPage.ClickOnSpecifiedOutlookMail(mailSubject);

                // Vérifier que la pièce jointe est présente et qu'elle n'est pas vide
                bool isPieceJointeOK = mailPage.IsOutlookPieceJointeOK(piecejointe);
                Assert.IsTrue(isPieceJointeOK, "La pièce jointe n'est pas présente dans le mail.");

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

                if (customerPortalTab)
                {
                    customerPortalLoginPage.Close();
                }
            }
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_QBD_SendCommandByEmptyMail()
        {
            // Prepare
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userCustomerPortalPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();
            var userPassword = TestContext.Properties["MailBox_UserPassword"].ToString();
            var mailSubjectToEmpty = "Empty Mail To";

            bool customerPortalTab = false;
            bool mailboxTab = false;

            var mailPage = new MailPage(WebDriver, TestContext);


            HomePage homePage = LogInAsAdmin();

            // Go to CustomerPortal Page
            var customerPortalLoginPage = homePage.GoToCustomerPortal();
            customerPortalTab = true;

            try
            {
                customerPortalLoginPage.LogIn(userMail, userCustomerPortalPassword);

                bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
                Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");

                CustomerPortalQtyByDayPage customerPortalQtyByDayPage = customerPortalLoginPage.ClickOnQtyByDay();

                // ================ test mail To vide ================
                string pj = customerPortalQtyByDayPage.SendQtyByMail("", mailSubjectToEmpty);
                customerPortalQtyByDayPage.SendMail();
                // ===================================================

                WebDriver.Navigate().Refresh();
                // Deconnexion du portal
                customerPortalLoginPage.LogOut();
                customerPortalLoginPage.Close();
                customerPortalTab = false;

                // Connexion à la mailbox pour voir si le mail a été reçu
                mailPage = customerPortalLoginPage.RedirectToOutlookMailbox();
                mailboxTab = true;

                mailPage.FillFields_LogToOutlookMailbox(userMail);

                WebDriver.Navigate().Refresh();

                // ================ test mail To vide ================
                mailPage.ClickOnSpecifiedOutlookMail(mailSubjectToEmpty);

                bool isPieceJointeOK = mailPage.IsOutlookPieceJointeOK(pj);
                Assert.IsTrue(isPieceJointeOK, "La pièce jointe n'est pas présente dans le mail.");

                mailPage.DeleteCurrentOutlookMail();
                // ===================================================

                mailPage.Close();
                mailboxTab = false;
            }

            finally
            {
                if (mailboxTab)
                {
                    mailPage.Close();
                }

                if (customerPortalTab)
                {
                    customerPortalLoginPage.Close();
                }
            }
        }


        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_QBD_Export()
        {
            // Prepare
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string siteACE = TestContext.Properties["SiteACE"].ToString();
            string siteMAD = TestContext.Properties["SiteMAD"].ToString();

            DeleteAllFileDownload();
            var homePage = LogInAsAdmin();

            CustomerPortalLoginPage customerPortalLoginPage = LogInCustomerPortal();
            customerPortalLoginPage.LogIn(userMail, userPassword);

            bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
            bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();

            Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
            Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page Quantity By Week.");

            CustomerPortalQtyByDayPage customerPortalQtyByDayPage = customerPortalLoginPage.ClickOnQtyByDay();
            //Avoir des données disponibles. (fonctionne sur dev5 site PMI)
            var lignes = WebDriver.FindElements(By.XPath("//*/tr[string(@data-service-id)]"));
            Assert.IsTrue(lignes.Count > 0, "pas de données disponibles");

            customerPortalQtyByDayPage.Export(siteMAD);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();
            var correctDownloadedFile = customerPortalQtyByDayPage.GetQMExportExcelFile(taskFiles);
            customerPortalQtyByDayPage.WaitPageLoading();
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

            // Récupération du nom du fichier et construction de l'URL du fichier Excel à ouvrir   
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            //Assert
            int resultNumber = OpenXmlExcel.GetExportResultNumber("Quantity to produce", filePath);
            var counter = customerPortalQtyByDayPage.GetCounter();
            Assert.AreEqual(counter, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);

            resultNumber += OpenXmlExcel.GetExportResultNumber("Total invoice by customer", filePath);
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);

            customerPortalQtyByDayPage.Export(siteACE);//site ACE sans delivery 

            // On récupère les fichiers du répertoire de téléchargement
            taskDirectory = new DirectoryInfo(downloadsPath);
            taskFiles = taskDirectory.GetFiles();
            correctDownloadedFile = customerPortalQtyByDayPage.GetQMExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

            // Récupération du nom du fichier et construction de l'URL du fichier Excel à ouvrir   
            fileName = correctDownloadedFile.Name;
            filePath = Path.Combine(downloadsPath, fileName);

            //Assert
            resultNumber = OpenXmlExcel.GetExportResultNumber("Quantity to produce", filePath);
            resultNumber += OpenXmlExcel.GetExportResultNumber("Total invoice by customer", filePath);
            List<string> customerNames = OpenXmlExcel.GetValuesInList("CustomerName", "Quantity to produce", filePath);
            List<string> deliverys = OpenXmlExcel.GetValuesInList("Delivery", "Quantity to produce", filePath);
            Assert.AreEqual(2, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
            // Arrange
            homePage.Navigate();
            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Site, siteACE);
            dispatchPage.FilterCustomersUnCheckAll();
            dispatchPage.FilterDeliverysUnCheckAll();

            dispatchPage.FilterCustomers(customerNames);
            dispatchPage.FilterDeliverys(deliverys);
            dispatchPage.WaitPageLoading();

            List<string> listOfDispatchPrevisionalQty = dispatchPage.GetTotalListDeliveryPrevisionalQty();
            List<string> listOfDispatchQtyToProduce = dispatchPage.GetTotalListDeliveryQtyToProduce();
            List<string> listOfDispatchQtyToInvoice = dispatchPage.GetTotalListDeliveryQtyToInvoice();
            dispatchPage.WaitPageLoading();
            bool verifDeliveryPrevisionalQuantity = dispatchPage.VerifyDispatchInDeliveries(deliverys, listOfDispatchPrevisionalQty);
            Assert.IsTrue(verifDeliveryPrevisionalQuantity, "Les éléments dans l'excel ne correspond pas à celui de listOfDispatch");

            var quantityToProducePage = dispatchPage.ClickQuantityToProduce();
            quantityToProducePage.WaitPageLoading();
            bool verifDeliveryQuantityToProduce = dispatchPage.VerifyDispatchInDeliveries(deliverys, listOfDispatchQtyToProduce);
            Assert.IsTrue(verifDeliveryQuantityToProduce, "Les éléments dans l'excel ne correspond pas à celui de listOfDispatch");

            var qtytoInvoice = dispatchPage.ClickQuantityToInvoice();
            qtytoInvoice.WaitPageLoading();
            bool verifDeliveryQuantityToInvoice = dispatchPage.VerifyDispatchInDeliveries(deliverys, listOfDispatchQtyToInvoice);
            Assert.IsTrue(verifDeliveryQuantityToInvoice, "Les éléments dans l'excel ne correspond pas à celui de listOfDispatch");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_QBD_AddAndValidateQuantities()
        {
            // Prepare
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();

            string site = TestContext.Properties["Site"].ToString();
            string deliveryName = TestContext.Properties["DeliveryCustomerPortal"].ToString();
            string serviceName = "ServiceForCustomerPortal";
            var dateNow = DateTime.Now;
            var demain = dateNow.AddDays(1);
            var quantityToUpadte = "35";

            var homePage = LogInAsAdmin();

             //Vérifier le paramétrage du delivery
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName);
            var loadingTab = deliveryPage.ClickOnFirstDelivery();
            var generalInfo = loadingTab.ClickOnGeneralInformation();
            generalInfo.SetToDefault();

            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Site, site);
            dispatchPage.Filter(DispatchPage.FilterType.Search, serviceName);
            dispatchPage.Filter(DispatchPage.FilterType.Deliveries, deliveryName);

            dispatchPage.ClickQuantityToInvoice();
            dispatchPage.UnValidateAll();

            dispatchPage.ClickQuantityToProduce();
            dispatchPage.UnValidateAll();

            dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.UnValidateAll();
            dispatchPage.ValidateAll();

           var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
           var qty = previsionnalQty.GetQuantity();
           var IsDispatchValidated = previsionnalQty.IsDispatchValidated();
           Assert.IsTrue(IsDispatchValidated, "Les previsionnal qty du dispatch sont déjà validées.");

           var qtyToProducePage = dispatchPage.ClickQuantityToProduce();
           var CanUpdateQty = qtyToProducePage.CanUpdateQty();
           Assert.IsTrue(CanUpdateQty, "Le dispatch est validé.");

           homePage.Navigate();
           var customerPortalLoginPage = homePage.GoToCustomerPortal();
           customerPortalLoginPage.LogIn(userMail, userPassword);
           bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
           Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
           CustomerPortalQtyByDayPage customerPortalQtyByDayPage = customerPortalLoginPage.ClickOnQtyByDay();
           CultureInfo.CurrentCulture = new CultureInfo("en-US", false);
           try
            {
                customerPortalQtyByDayPage.GoToDate(demain);
                customerPortalQtyByDayPage.SearchFilterDelivery(deliveryName);
                customerPortalQtyByDayPage.WaitPageLoading();
                var quantityEdit = customerPortalQtyByDayPage.GetQuantitytoproduce(serviceName);
                Assert.AreEqual(qty, quantityEdit ,"La quantite au dispatsh ne convient pas au quantite dans customer portal");
                customerPortalQtyByDayPage.SetQuantitytoproduce(quantityToUpadte, serviceName);
                customerPortalQtyByDayPage.ValidateQuantities();
                WebDriver.Navigate().Refresh();
                customerPortalQtyByDayPage.SearchFilterDelivery(deliveryName);
                customerPortalQtyByDayPage.WaitPageLoading();
                var quantityAfter = customerPortalQtyByDayPage.GetQuantitytoproduce(serviceName);
                Assert.AreEqual(quantityToUpadte, quantityAfter , "La quantite n'a pas Changé");

                //les quantités saisies dans le dispatch doivent être les mêmes que celles saisies dans le portal	
                customerPortalQtyByDayPage.Close();
                customerPortalQtyByDayPage.WaitPageLoading();
                dispatchPage = homePage.GoToProduction_DispatchPage();
                dispatchPage.Filter(DispatchPage.FilterType.Site, site);
                dispatchPage.Filter(DispatchPage.FilterType.Search, serviceName);
                dispatchPage.Filter(DispatchPage.FilterType.Deliveries, deliveryName);

                //attraper la Qty pour le jour de la semaine dans la semaine
                if (DateUtils.Now.AddDays(1).DayOfWeek == DayOfWeek.Monday)
                {
                    dispatchPage.ClickNextWeek();
                }
                var quantityToProducePage = dispatchPage.ClickQuantityToProduce();
                var quantityToProduceForTomorrow = quantityToProducePage.GetQuantityToProduceForTomorrow();
                Assert.AreEqual(quantityToUpadte, quantityToProduceForTomorrow, "Pas de modification CP vers Dispatch");

            }
            finally
            {
                homePage.Navigate();
                customerPortalLoginPage = homePage.GoToCustomerPortal();
                customerPortalQtyByDayPage = customerPortalLoginPage.ClickOnQtyByDay();
                customerPortalQtyByDayPage.GoToDate(demain);
                customerPortalQtyByDayPage.SearchFilterDelivery(deliveryName);
                CultureInfo.CurrentCulture = new CultureInfo("en-US", false);
                customerPortalQtyByDayPage.SetQuantities(qty);
                customerPortalQtyByDayPage.ValidateQuantities();
                WebDriver.Navigate().Refresh();

            }

        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_QBD_DuplicateDay()
        {
            // Prepare
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();
            string serviceName = TestContext.Properties["ServiceCustomerPortal"].ToString(); 
            string qty = "27";
            string comment = "comment";
            var date = DateUtils.Now.AddDays(1);
            var Today = DateUtils.Now;


            CustomerPortalLoginPage customerPortalLoginPage = LogInCustomerPortal();
            customerPortalLoginPage.LogIn(userMail, userPassword);

            bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
            bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();

            Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
            Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page Quantity By Week.");
            //. Avoir des delivery associés au compte
            var customerPortalQtyByDayPage = customerPortalLoginPage.ClickOnQtyByDay();
            customerPortalQtyByDayPage.GoToDate(Today);
            var lignes = customerPortalQtyByDayPage.CheckServiceNumber();

            if (lignes.Count == 0)
            {
                customerPortalQtyByDayPage.GoToDate(date);
                lignes = customerPortalQtyByDayPage.CheckServiceNumber();
                Assert.IsTrue(lignes.Count > 0, "pas de données disponibles");

                customerPortalQtyByDayPage.Search(serviceName);
                CultureInfo.CurrentCulture = new CultureInfo("en-US", false);

                if (date.ToString("dddd") != "Sunday")
                {

                    customerPortalQtyByDayPage.GoToDate(date);
                    customerPortalQtyByDayPage.UpdateQunatitySetComment(qty, comment);
                    customerPortalQtyByDayPage.Validate();

                    //Cliquer sur icone dupliquer(à droite du jour) et sélectionner le jour à dupliquer
                    customerPortalQtyByDayPage.DuplicateOneDay(date.ToString("dddd"), 1);

                    customerPortalQtyByDayPage.Search(serviceName);
                    //Vérifier que la quantité et le commentaire ont bien été dupliqué
                    //sur le jour sélectionné
                    customerPortalQtyByDayPage.GoToDate(DateUtils.Now.AddDays(2));
                    customerPortalQtyByDayPage.Search(serviceName);
                    string Qty = customerPortalQtyByDayPage.GetQuantityValue();
                    string commentdisplyed = customerPortalQtyByDayPage.GetComment();
                    Assert.AreEqual(qty, Qty, "Qty non dupliquée");
                    Assert.AreEqual(comment, commentdisplyed, "Commentaire non dupliqué");


                }
            }
            else
            {
                customerPortalQtyByDayPage.Search(serviceName);
                //ajouter une étape où on valide la quantité de demain puis on change la quantité et revalider, vérifier
                //sauf si c'est dimanche
                CultureInfo.CurrentCulture = new CultureInfo("en-US", false);

                if (Today.ToString("dddd") != "Sunday")
                {

                    customerPortalQtyByDayPage.GoToDate(Today);
                    customerPortalQtyByDayPage.UpdateQunatitySetComment(qty, comment);
                    customerPortalQtyByDayPage.Validate();

                    //Cliquer sur icone dupliquer(à droite du jour) et sélectionner le jour à dupliquer
                    customerPortalQtyByDayPage.DuplicateOneDay(Today.ToString("dddd"), 1);

                    customerPortalQtyByDayPage.Search(serviceName);
                    //Vérifier que la quantité et le commentaire ont bien été dupliqué
                    //sur le jour sélectionné
                    customerPortalQtyByDayPage.GoToDate(DateUtils.Now.AddDays(1));

                    customerPortalQtyByDayPage.Search(serviceName);
                    string Qty = customerPortalQtyByDayPage.GetQuantityValue();
                    string commentdisplyed = customerPortalQtyByDayPage.GetComment();
                    Assert.AreEqual(qty, Qty, "Qty non dupliquée");
                    Assert.AreEqual(comment, commentdisplyed, "Commentaire non dupliqué");

                }
            }
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_QBD_DailyConfiguration()
        {
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();
            var delivery = "DeliveryForCustomerPortal";
            var demain = DateTime.Now.AddDays(1);
            var homePage = LogInAsAdmin();
            #region set parametres to default of deliveryForCustomerPortal
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.Filter(DeliveryPage.FilterType.Search, delivery);
            var loadingTab = deliveryPage.ClickOnFirstDelivery();
            var serviceName = loadingTab.GetServiceName();
            var generalInfo = loadingTab.ClickOnGeneralInformation();
            generalInfo.SetToDefault();
            #endregion

            try
            {

                generalInfo.SetAllowedInLimitedAccess("100");
                CustomerPortalLoginPage customerPortalLoginPage = LogInCustomerPortal();
                customerPortalLoginPage.LogIn(userMail, userPassword);
                CustomerPortalQtyByDayPage customerPortalQtyByDayPage = customerPortalLoginPage.ClickOnQtyByDay();
                customerPortalQtyByDayPage.GoToDate(demain);
                customerPortalQtyByDayPage.SearchFilterDelivery(delivery);
                var dayName = customerPortalQtyByDayPage.GetDayName();
                var demainModifiable = customerPortalQtyByDayPage.CheckIfModifiableByDay(dayName, serviceName);
                Assert.IsFalse(demainModifiable, "demain modifiable when AllowedInLimitedAccess changed to 100");

            }
            finally
            {
                homePage.Navigate();
                deliveryPage = homePage.GoToCustomers_DeliveryPage();
                deliveryPage.Filter(DeliveryPage.FilterType.Search, delivery);
                var deliveryLoadingPage = deliveryPage.ClickOnFirstDelivery();
                var deliveryGeneralinformation = deliveryLoadingPage.ClickOnGeneralInformation();
                deliveryGeneralinformation.SetToDefault();
            }
        }


        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_QBW_DailyConfiguration()
        {
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();
            var delivery = "DeliveryForCustomerPortal";
            var demain = DateTime.Now.AddDays(1);

            var homePage = LogInAsAdmin();

            #region set parametres to default of deliveryForCustomerPortal
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.Filter(DeliveryPage.FilterType.Search, delivery);
            var loadingTab = deliveryPage.ClickOnFirstDelivery();
            var serviceName = loadingTab.GetServiceName();
            var generalInfo = loadingTab.ClickOnGeneralInformation();
            generalInfo.SetToDefault();
            #endregion


            try
            {
                generalInfo.SetAllowedInLimitedAccess("100");

                CustomerPortalLoginPage customerPortalLoginPage = LogInCustomerPortal();
                customerPortalLoginPage.LogIn(userMail, userPassword);
                customerPortalLoginPage = LogInCustomerPortal();
                CustomerPortalQtyByWeekPage customerPortalQtyByWeekPage = customerPortalLoginPage.ClickOnQtyByWeek();
                customerPortalQtyByWeekPage.GoToDate(demain);
                customerPortalQtyByWeekPage.SearchFilterDelivery(delivery);
                customerPortalQtyByWeekPage.WaitPageLoading();
                var demainModifiable = customerPortalQtyByWeekPage.CheckIfModifiableByweek(demain, serviceName);
                Assert.IsFalse(demainModifiable, "demain modifiable when AllowedInLimitedAccess changed to 100");
            }
            finally
            {
                homePage.Navigate();
                // Set parametres to default of deliveryForCustomerPortal
                deliveryPage = homePage.GoToCustomers_DeliveryPage();
                deliveryPage.Filter(DeliveryPage.FilterType.Search, delivery);
                var deliveryLoadingPage = deliveryPage.ClickOnFirstDelivery();
                var deliveryGeneralinformation = deliveryLoadingPage.ClickOnGeneralInformation();
                deliveryGeneralinformation.SetToDefault();
            }
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_QBD_ModularConfiguration()
        {
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();
            var delivery = "DeliveryForCustomerPortal";
            var demain = DateTime.Now.AddDays(1);

            var homePage = LogInAsAdmin();

            #region set parametres to default of deliveryForCustomerPortal
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.Filter(DeliveryPage.FilterType.Search, delivery);
            var loadingTab = deliveryPage.ClickOnFirstDelivery();
            var serviceName = loadingTab.GetServiceName();
            var generalInfo = loadingTab.ClickOnGeneralInformation();
            generalInfo.SetToDefault();
            #endregion

            try
            {
                generalInfo.SetStopAccessToCustomerPortal("Modular");
                generalInfo.SetAllowedInLimitedAccess("100");
                var customerPortalLoginPage = LogInCustomerPortal();
                customerPortalLoginPage.LogIn(userMail, userPassword);
                var customerPortalQtyByDayPage = customerPortalLoginPage.ClickOnQtyByDay();
                customerPortalQtyByDayPage.GoToDate(demain);
                customerPortalQtyByDayPage.SearchFilterDelivery(delivery);
                var dayName = customerPortalQtyByDayPage.GetDayName();
                var demainModifiable = customerPortalQtyByDayPage.CheckIfModifiableByDay(dayName, serviceName);               
                Assert.IsFalse(demainModifiable, "demain modifiable Allowed in limited Access");

            }
            finally
            {
                homePage.Navigate();
                // Set parametres to default of deliveryForCustomerPortal
                deliveryPage = homePage.GoToCustomers_DeliveryPage();
                deliveryPage.Filter(DeliveryPage.FilterType.Search, delivery);
                var deliveryLoadingPage = deliveryPage.ClickOnFirstDelivery();
                var deliveryGeneralinformation = deliveryLoadingPage.ClickOnGeneralInformation();
                deliveryGeneralinformation.SetToDefault();
            }



        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_QBW_ModularConfiguration()
        {
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();
            var delivery = "DeliveryForCustomerPortal";
            string site = TestContext.Properties["Site"].ToString();
            string customer = TestContext.Properties["CustomerDelivery"].ToString();
            var aujoudhui = DateTime.Now;
            var demain = DateTime.Now.AddDays(1);


            var homePage = LogInAsAdmin();

            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.ResetFilter();
            deliveryPage.Filter(DeliveryPage.FilterType.Search, delivery);
            DeliveryLoadingPage loadingTab = deliveryPage.ClickOnFirstDelivery();
            var serviceName = loadingTab.GetServiceName();

            var generalInfo = loadingTab.ClickOnGeneralInformation();
            generalInfo.SetToDefault();

            try
            {
                //set mode modular
                generalInfo.SetStopAccessToCustomerPortal("Modular");

                //1. Delivery time: 23:00 => utc +1
                generalInfo.SetDeliveryTime("23", "00");
                //données de demain non modifiable et celles d'aujourd'hui modifiable
                var customerPortalLoginPage = LogInCustomerPortal();
                customerPortalLoginPage.LogIn(userMail, userPassword);
                var customerPortalQtyByWeekPage = customerPortalLoginPage.ClickOnQtyByWeek();
                customerPortalQtyByWeekPage.SearchFilterDelivery(delivery);
                customerPortalQtyByWeekPage.GoToDate(aujoudhui);
                var aujourdhuiModifiable = customerPortalQtyByWeekPage.CheckIfModifiableByweek(aujoudhui, serviceName);
                customerPortalQtyByWeekPage.GoToDate(demain);
                var demainModifiable = customerPortalQtyByWeekPage.CheckIfModifiableByweek(demain, serviceName);
                //données de demain non modifiable et celles d'aujourd'hui modifiable
                Assert.IsFalse(demainModifiable, "demain modifiable when deliveryTime changed to 23:00");
                Assert.IsTrue(aujourdhuiModifiable, "aujourd'hui non modifiable when deliveryTime changed to 23:00");
                homePage.Navigate();

                //2. % allowed in limited access: 100
                deliveryPage = homePage.GoToCustomers_DeliveryPage();
                deliveryPage.Filter(DeliveryPage.FilterType.Search, delivery);
                var deliveryLoadingPage = deliveryPage.ClickOnFirstDelivery();
                generalInfo = deliveryLoadingPage.ClickOnGeneralInformation();
                generalInfo.SetToDefault();
                generalInfo.SetStopAccessToCustomerPortal("Modular");
                generalInfo.SetAllowedInLimitedAccess("100");
                customerPortalLoginPage = LogInCustomerPortal();
                customerPortalQtyByWeekPage = customerPortalLoginPage.ClickOnQtyByWeek();
                customerPortalQtyByWeekPage.SearchFilterDelivery(delivery); 
                customerPortalQtyByWeekPage.GoToDate(demain);
                demainModifiable = customerPortalQtyByWeekPage.CheckIfModifiableByweek(demain, serviceName);
                //données de demain non modifiable
                Assert.IsFalse(demainModifiable, "demain modifiable Allowed in limited Access");
                homePage.Navigate();

                //3. Starting time to limited access (hours): 0
                deliveryPage = homePage.GoToCustomers_DeliveryPage();
                deliveryPage.Filter(DeliveryPage.FilterType.Search, delivery);
                deliveryLoadingPage = deliveryPage.ClickOnFirstDelivery();
                generalInfo = deliveryLoadingPage.ClickOnGeneralInformation();
                generalInfo.SetToDefault();
                generalInfo.SetStopAccessToCustomerPortal("Modular");
                //Modifier sur la journée de demain :
                var demainText = demain.ToString("dddd", CultureInfo.GetCultureInfo("en-US"));
                generalInfo.ClickDay(demainText);
                generalInfo.SetDayStartingTimeToLimitedAccess(demainText, "0");
                customerPortalLoginPage = LogInCustomerPortal();
                customerPortalQtyByWeekPage = customerPortalLoginPage.ClickOnQtyByWeek();
                customerPortalQtyByWeekPage.SearchFilterDelivery(delivery);
                customerPortalQtyByWeekPage.GoToDate(demain);
                demainModifiable = customerPortalQtyByWeekPage.CheckIfModifiableByweek(demain, serviceName);
                //données de demain non modifiable
                Assert.IsFalse(demainModifiable, "demain modifiable StartingTimeToLimitedAccess changed to 0");
                homePage.Navigate();
                //4. Ending time to limited access (hours): 24
                deliveryPage = homePage.GoToCustomers_DeliveryPage();
                deliveryPage.Filter(DeliveryPage.FilterType.Search, delivery);
                deliveryLoadingPage = deliveryPage.ClickOnFirstDelivery();
                generalInfo = deliveryLoadingPage.ClickOnGeneralInformation();
                generalInfo.SetToDefault();
                generalInfo.SetStopAccessToCustomerPortal("Modular");
                //Modifier sur la journée de demain :
                generalInfo.ClickDay(demainText);
                generalInfo.SetDayEndingTimeToLimitedAccess(demainText, "24");
               customerPortalLoginPage = LogInCustomerPortal();
                customerPortalQtyByWeekPage = customerPortalLoginPage.ClickOnQtyByWeek();
                customerPortalQtyByWeekPage.SearchFilterDelivery(delivery);
                customerPortalQtyByWeekPage.GoToDate(demain);
                demainModifiable = customerPortalQtyByWeekPage.CheckIfModifiableByweek(demain, serviceName);
                //données de demain non modifiable
                Assert.IsFalse(demainModifiable, "demain modifiable EndingTimeToLimitedAccess changed to 24");
            }
            finally
            {
                homePage.Navigate();
                // Set parametres to default of deliveryForCustomerPortal
                deliveryPage = homePage.GoToCustomers_DeliveryPage();
                deliveryPage.Filter(DeliveryPage.FilterType.Search, delivery);
                loadingTab = deliveryPage.ClickOnFirstDelivery();
                generalInfo = loadingTab.ClickOnGeneralInformation();

                var demainText = demain.ToString("dddd", CultureInfo.GetCultureInfo("en-US"));
                generalInfo.SetStopAccessToCustomerPortal("Modular");
                generalInfo.ClickDay(demainText);
                generalInfo.SetDayStartingTimeToLimitedAccess(demainText, "24");
                generalInfo.SetDayEndingTimeToLimitedAccess(demainText, "0");
                generalInfo.SetToDefault();


            }
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_CO_index_date()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);

            //1.Aller sur Customer Portal, se logguer, aller sur Customer Orders
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userCustomerPortalPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();
            // Prepare
            var culture = CultureInfo.InvariantCulture;
            string dateFormat = "dd/MM/yyyy HH:mm";
            // Arrange
            CustomerPortalLoginPage customerPortalLoginPage = LogInCustomerPortal();
            customerPortalLoginPage.LogIn(userMail, userCustomerPortalPassword);

            bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
            bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();

            Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
            Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page.");
            var customerPortalCustomerOrdersPage = customerPortalLoginPage.ClickOnCustomerOrders();
            customerPortalCustomerOrdersPage.ResetFilter();
            // Prendre le Delivery date + hour de premier CO depuis l'index.
            var deliveryDatePlusHour = customerPortalCustomerOrdersPage.GetFirstCODeliveryDatePlusHour();
            var deliveryDatePlusHourParsed = DateTime.TryParseExact(deliveryDatePlusHour, dateFormat, culture, DateTimeStyles.None, out var deliveryDate);
            Assert.IsTrue(deliveryDatePlusHourParsed, "La delivery date + hour n'a pas pu être analysé.");
            // Prendre le Delivery date + hour de premier CO depuis la page General Information de CO.
            var customerPortalCOPageResult = customerPortalCustomerOrdersPage.GoToFirstCustomer();
            customerPortalCOPageResult.GoToGeneralInformation();
            var generalInfoDate = customerPortalCOPageResult.GetDeliveryDate();
            var generalInfoTime = customerPortalCOPageResult.GetHOUR();
            var generalInfoDateTimeString = $"{generalInfoDate} {generalInfoTime}";
            var generalInfoDatePlusHourParsed = DateTime.TryParseExact(generalInfoDateTimeString, dateFormat, culture, DateTimeStyles.None, out var generalInfoDeliveryDate);
            Assert.IsTrue(generalInfoDatePlusHourParsed, "La delivery date + hour de la page General Information de CO n'a pas pu être analysé.");
            // Vérifier que la delivery date + hour dans sont les mêmes dans l'index et dans la CO
            Assert.AreEqual(deliveryDate, generalInfoDeliveryDate, "La delivery date + hour de l'index ne corresponde pas à la delivery date et l'hour de la page General Information de CO.");


        }
        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_CO_index_delivery()
        {
            //1.Aller sur Customer Portal, se logguer, aller sur Customer Orders
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userCustomerPortalPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();
            // Arrange
            CustomerPortalLoginPage customerPortalLoginPage = LogInCustomerPortal();
            customerPortalLoginPage.LogIn(userMail, userCustomerPortalPassword);
            bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
            bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();
            Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
            Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page.");
            // Rechercher l'ordre crée par le test prioritaire CUPO_CO_CreateNewCustomerOrder.
            var customerPortalCustomerOrdersPage = customerPortalLoginPage.ClickOnCustomerOrders();
            customerPortalCustomerOrdersPage.ResetFilter();
            customerPortalCustomerOrdersPage.Search(_deliceryNamOfCreateCustomerOrder);
            Assert.IsTrue(customerPortalCustomerOrdersPage.CheckTotalNumber() > 0, $"Aucun CO avec le delivery name {_deliceryNamOfCreateCustomerOrder} a été trouvé");

            // Prendre le Delivery name de premier CO depuis l'index.
            var deliveryName = customerPortalCustomerOrdersPage.GetFirstCODeliveryName();
            // Prendre le Delivery name de premier CO depuis la page General Information de CO.
            var customerPortalCOPageResult = customerPortalCustomerOrdersPage.GoToFirstCustomer();
            customerPortalCOPageResult.GoToGeneralInformation();
            var generalInfoDeliveryName = customerPortalCOPageResult.GetGeneralInformationDeliveryName("DeliveryName");
            // Vérifier que les deliveries s'affichent bien dans l'index.
            Assert.AreEqual(deliveryName, generalInfoDeliveryName, "Le delivery name de l'index ne corresponde pas au delivery name de la page General Information de CO.");


        }
        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_CO_index_Prix()
        {
            //1.Aller sur Customer Portal, se logguer, aller sur Customer Orders
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userCustomerPortalPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();

            // Arrange
            CustomerPortalLoginPage customerPortalLoginPage = LogInCustomerPortal();
            customerPortalLoginPage.LogIn(userMail, userCustomerPortalPassword);

            bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
            bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();
            Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
            Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page.");
            // Rechercher un ordre éxistant.
            var customerPortalCustomerOrdersPage = customerPortalLoginPage.ClickOnCustomerOrders();
            customerPortalCustomerOrdersPage.ResetFilter();
            Assert.IsTrue(customerPortalCustomerOrdersPage.CheckTotalNumber() > 0, $"Aucun CO a été trouvé");
            // Prendre le prix de premier CO depuis l'index.
            var price = customerPortalCustomerOrdersPage.GetFirstCOPrice();
            // Prendre le prix de premier CO depuis la page des détails de CO.
            var customerPortalCOPageResult = customerPortalCustomerOrdersPage.GoToFirstCustomer();
            var coDetailPrice = customerPortalCOPageResult.PriceExclVAT();
            // Vérifier que le prix affiché dans l'index correspond bien au prix quand on rentre dans la CO
            Assert.AreEqual(price, coDetailPrice, "Le prix de l'index ne corresponde pas au prix de la page des détails de CO.");

        }
        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_CO_index_customer_name()
        {
            //1.Aller sur Customer Portal, se logguer, aller sur Customer Orders
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userCustomerPortalPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();
            // Arrange
            CustomerPortalLoginPage customerPortalLoginPage = LogInCustomerPortal();
            customerPortalLoginPage.LogIn(userMail, userCustomerPortalPassword);
            bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
            bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();
            Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
            Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page.");
            // Rechercher un ordre éxistant.
            var customerPortalCustomerOrdersPage = customerPortalLoginPage.ClickOnCustomerOrders();
            customerPortalCustomerOrdersPage.ResetFilter();
            Assert.IsTrue(customerPortalCustomerOrdersPage.CheckTotalNumber() > 0, $"Aucun CO été trouvé");
            // Prendre le Delivery name de premier CO depuis l'index.
            var customerName = customerPortalCustomerOrdersPage.GetFirstCOCustomerName().Trim();
            // Prendre le Delivery name de premier CO depuis la page General Information de CO.
            var customerPortalCOPageResult = customerPortalCustomerOrdersPage.GoToFirstCustomer();
            customerPortalCOPageResult.GoToGeneralInformation();
            var generalInfoCustomerNamePlusCode = customerPortalCOPageResult.GetGICustomerName();
            var CustomerNamePlusCodeSplited = generalInfoCustomerNamePlusCode.Split(new char[] { '-' }, 2);
            Assert.AreEqual(2, CustomerNamePlusCodeSplited.Length, "Le format du Customer Name dans les general informations est incorrect.");
            var generalInfocustomerName = CustomerNamePlusCodeSplited.Length > 1 ? CustomerNamePlusCodeSplited[1].Trim() : string.Empty;
            // Vérifier que le customer name s'affiche bien dans l'index + qu'il correspond au nom dans la CO.
            Assert.AreEqual(customerName, generalInfocustomerName, "Le delivery name de l'index ne corresponde pas au delivery name de la page General Information de CO.");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_CO_commercial_name()
        {
            //(test identique à PR_CO_CreateNewCustomerOrder mais sur Customer Portal)

            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();

            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string deliveryName = new Random().Next().ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();

            // Kept delivery name to filter for next tests.
            _deliceryNamOfCreateCustomerOrder = deliveryName;

            // Arrange
            CustomerPortalLoginPage customerPortalLoginPage = LogInCustomerPortal();
            customerPortalLoginPage.LogIn(userMail, userPassword);

            bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
            bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();

            Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
            Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page.");

            var customerPortalCustomerOrdersPage = customerPortalLoginPage.ClickOnCustomerOrders();

            //2.Cliquer sur "+"
            var customerPortalCustmerOrdersPageModal = customerPortalCustomerOrdersPage.ClickAdd();


            //2.Remplir le formulaire avec le site(BCN), le customer(EMIRATES AILRINES) et l'aircraft (B777)
            customerPortalCustmerOrdersPageModal.SelectSite(site);

            customerPortalCustmerOrdersPageModal.SelectCustomer(customer);

            customerPortalCustmerOrdersPageModal.SelectDeliveryName(deliveryName);

            customerPortalCustmerOrdersPageModal.SelectAircraft(aircraft);


            //3.Cliquer sur "create"
            CustomerPortalCustomerOrdersPageResult customerPortalCustomerOrdersPageResult = customerPortalCustmerOrdersPageModal.Create();


            //3.Cliquer sur '+'
            //4.Ajouter un item et une quantity(HOT MEAL YC UAE et 10)
            customerPortalCustomerOrdersPageResult.AddLine("HOT MEAL YC UAE", 10);

            //5.Vérifier l'ajout de l'item
            // attente du rechargement du tableau
            customerPortalCustomerOrdersPageResult.WaitPageLoading();
            customerPortalCustomerOrdersPageResult.Validate();

            //4.Récupérer le numéro du customer order crée, faire Back to List
            //result customer order
            var resultText = customerPortalCustomerOrdersPageResult.ResultCustomerOrder();
            Regex r = new Regex("##### Customer order ##### ([0-9]+) - EMIRATES AIRLINES");
            Match m = r.Match(resultText);
            string customerOrderNumberId = Regex.Match(resultText, @"\d+").Value;

            // back to list
            customerPortalCustomerOrdersPage.Navigate("/CustomerPortal/CustomerOrders");

            string customerPortalCustomerOrderId = Regex.Match(customerPortalCustomerOrdersPage.getTableFirstLine(1), @"\d+").Value;
            //Vérifier dans la table si on retrouve le customer order créé
            Assert.AreEqual(customerOrderNumberId, customerPortalCustomerOrderId, "Première colonne ne contient pas " + customerOrderNumberId);

            //6.Vérifier l'ajout du comment sur l'item(bien affiché et identique à ce qui a été entrée)
            customerPortalCustomerOrdersPageResult = customerPortalCustomerOrdersPage.GoToFirstCustomer();
            customerPortalCustomerOrdersPageResult.GoToDetails();

            var commercialName = customerPortalCustomerOrdersPageResult.GetCommercialNameItem();
            Assert.AreEqual(customerPortalCustomerOrdersPageResult.GetCommercialNameItem(), "HOT MEAL YC UAE", "Les noms commerciaux n'apparaissent pas bien.");
            if (isUserLogged)
            {
                customerPortalLoginPage.LogOut();
            }

        }
        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_user_denglish()
        {   //(test identique à PR_CO_CreateNewCustomerOrder mais sur Customer Portal)

            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();
            // boolean d'ouverture des nouveaux onglets
            bool mailboxTab = false;
            bool updatePasswordTab = false;

            var mailPage = new MailPage(WebDriver, TestContext);
            var updatePage = new CustomerPortalUpdatePage(WebDriver, TestContext);
            // Kept delivery name to filter for next tests.


            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var settingsPortalPage = homePage.GoToParameters_PortalPage();

            // si le user n'existe pas, on le créé
            if (!settingsPortalPage.IsExistingUser(userMail))
            {
                settingsPortalPage.AddNewUser(userMail);
                settingsPortalPage.SearchAndSelectUser(userMail);

                try
                {
                    // Connexion à la mailbox pour voir si le mail a été envoyé
                    mailPage = settingsPortalPage.RedirectToOutlookMailbox();
                    mailboxTab = true;

                    mailPage.FillFields_LogToOutlookMailbox(userMail);

                    // Recherche du mail envoyé pour le resend password
                    mailPage.ClickOnSpecifiedOutlookMail("Bienvenido a nuestro portal cliente");

                    updatePage = mailPage.ClickOnLinkForPassword("aquí");
                    updatePasswordTab = true;

                    var customerPortalLoginPage = updatePage.UpdatePassword(userPassword);
                    customerPortalLoginPage.Close();
                    updatePasswordTab = false;

                    mailPage.DeleteCurrentOutlookMail();
                    mailPage.Close();
                    mailboxTab = false;

                }
                finally
                {
                    // On ferme les onglets ouverts
                    if (updatePasswordTab)
                    {
                        updatePage.Close();
                    }

                    if (mailboxTab)
                    {
                        mailPage.Close();
                    }
                }
            }
            else
            {
                settingsPortalPage.SearchAndSelectUser(userMail);
                //settingsPortalPage.ConfigureUser(quantityByWeek, customerOrder, quantityByDay);
            }

            settingsPortalPage = homePage.GoToParameters_PortalPage();
            Assert.IsTrue(settingsPortalPage.IsExistingUser(userMail), "Le user n'est pas présent pour accéder au customer portal.");

            try
            {
                settingsPortalPage.SearchAndSelectUser(userMail);
                settingsPortalPage.SetLanguage("English");
                settingsPortalPage.SaveAndConfirm();
                CustomerPortalLoginPage customerPortalLogin = LogInCustomerPortal(true);
                customerPortalLogin.LogIn(userMail, userPassword);

                bool isUserLogged = customerPortalLogin.IsUserLogged(userMail);
                bool isPageLoaded = customerPortalLogin.IsPageLoadedWithoutError();

                Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
                Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page.");
            }
            finally
            {
                //return to english language (to avoid traduction problems)

                homePage.Navigate();
                settingsPortalPage = homePage.GoToParameters_PortalPage();
                settingsPortalPage.SearchAndSelectUser(userMail);
                settingsPortalPage.SetLanguage("English");
                settingsPortalPage.SaveAndConfirm();

                //le save reset les delivery --> relink delivery to user
                string deliveryName = TestContext.Properties["DeliveryCustomerPortal"].ToString();
                string deliveryNameAFA = "A.F.A. CASAS";

                var portalSettingsPage = homePage.GoToParameters_PortalPage();
                portalSettingsPage.SearchAndSelectUser(userMail);
                portalSettingsPage.LinkDeliveryToUser(deliveryNameAFA);
                portalSettingsPage.LinkDeliveryToUser(deliveryName);
            }
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_user_spanish()
        {   //(test identique à PR_CO_CreateNewCustomerOrder mais sur Customer Portal)

            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();
            // boolean d'ouverture des nouveaux onglets
            bool mailboxTab = false;
            bool updatePasswordTab = false;

            var mailPage = new MailPage(WebDriver, TestContext);
            var updatePage = new CustomerPortalUpdatePage(WebDriver, TestContext);
            // Kept delivery name to filter for next tests.


            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var settingsPortalPage = homePage.GoToParameters_PortalPage();

            // si le user n'existe pas, on le créé
            if (!settingsPortalPage.IsExistingUser(userMail))
            {
                settingsPortalPage.AddNewUser(userMail);
                settingsPortalPage.SearchAndSelectUser(userMail);



                try
                {
                    // Connexion à la mailbox pour voir si le mail a été envoyé
                    mailPage = settingsPortalPage.RedirectToOutlookMailbox();
                    mailboxTab = true;

                    mailPage.FillFields_LogToOutlookMailbox(userMail);

                    // Recherche du mail envoyé pour le resend password
                    mailPage.ClickOnSpecifiedOutlookMail("Bienvenido a nuestro portal cliente");

                    updatePage = mailPage.ClickOnLinkForPassword("aquí");
                    updatePasswordTab = true;

                    var customerPortalLoginPage = updatePage.UpdatePassword(userPassword);
                    customerPortalLoginPage.Close();
                    updatePasswordTab = false;

                    mailPage.DeleteCurrentOutlookMail();
                    mailPage.Close();
                    mailboxTab = false;

                }
                finally
                {
                    // On ferme les onglets ouverts
                    if (updatePasswordTab)
                    {
                        updatePage.Close();
                    }

                    if (mailboxTab)
                    {
                        mailPage.Close();
                    }
                }
            }
            else
            {
                settingsPortalPage.SearchAndSelectUser(userMail);
                //settingsPortalPage.ConfigureUser(quantityByWeek, customerOrder, quantityByDay);
            }

            settingsPortalPage = homePage.GoToParameters_PortalPage();
            Assert.IsTrue(settingsPortalPage.IsExistingUser(userMail), "Le user n'est pas présent pour accéder au customer portal.");

            try
            {
                settingsPortalPage.SearchAndSelectUser(userMail);
                settingsPortalPage.SetLanguage("Spanish");
                settingsPortalPage.SaveAndConfirm();
                settingsPortalPage = homePage.GoToParameters_PortalPage();
                settingsPortalPage.SearchAndSelectUser(userMail);
                Assert.AreEqual(settingsPortalPage.GetLanguage(), "Spanish", "La langue n'est pas modifiée en spanish");
                CustomerPortalLoginPage customerPortalLogin = LogInCustomerPortal(true);
                customerPortalLogin.LogIn(userMail, userPassword);

                bool isUserLogged = customerPortalLogin.IsUserLogged(userMail);
                bool isPageLoaded = customerPortalLogin.IsPageLoadedWithoutError();

                Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
                Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page.");
            }
            finally
            {
                //return to english language (to avoid traduction problems)

                homePage.Navigate();
                settingsPortalPage = homePage.GoToParameters_PortalPage();
                settingsPortalPage.SearchAndSelectUser(userMail);
                settingsPortalPage.SetLanguage("English");
                settingsPortalPage.SaveAndConfirm();

                //le save reset les delivery --> relink delivery to user
                string deliveryName = TestContext.Properties["DeliveryCustomerPortal"].ToString();
                string deliveryNameAFA = "A.F.A. CASAS";

                var portalSettingsPage = homePage.GoToParameters_PortalPage();
                portalSettingsPage.SearchAndSelectUser(userMail);
                portalSettingsPage.LinkDeliveryToUser(deliveryNameAFA);
                portalSettingsPage.LinkDeliveryToUser(deliveryName);
            }
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_user_others()
        {   //(test identique à PR_CO_CreateNewCustomerOrder mais sur Customer Portal)

            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();

            // boolean d'ouverture des nouveaux onglets
            bool mailboxTab = false;
            bool updatePasswordTab = false;

            var mailPage = new MailPage(WebDriver, TestContext);
            var updatePage = new CustomerPortalUpdatePage(WebDriver, TestContext);
            // Kept delivery name to filter for next tests.


            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var settingsPortalPage = homePage.GoToParameters_PortalPage();

            // si le user n'existe pas, on le créé
            if (!settingsPortalPage.IsExistingUser(userMail))
            {
                settingsPortalPage.AddNewUser(userMail);
                settingsPortalPage.SearchAndSelectUser(userMail);

                try
                {
                    // Connexion à la mailbox pour voir si le mail a été envoyé
                    mailPage = settingsPortalPage.RedirectToOutlookMailbox();
                    mailboxTab = true;

                    mailPage.FillFields_LogToOutlookMailbox(userMail);

                    // Recherche du mail envoyé pour le resend password
                    mailPage.ClickOnSpecifiedOutlookMail("Bienvenido a nuestro portal cliente");

                    updatePage = mailPage.ClickOnLinkForPassword("aquí");
                    updatePasswordTab = true;

                    var customerPortalLoginPage = updatePage.UpdatePassword(userPassword);
                    customerPortalLoginPage.Close();
                    updatePasswordTab = false;

                    mailPage.DeleteCurrentOutlookMail();
                    mailPage.Close();
                    mailboxTab = false;

                }
                finally
                {
                    // On ferme les onglets ouverts
                    if (updatePasswordTab)
                    {
                        updatePage.Close();
                    }

                    if (mailboxTab)
                    {
                        mailPage.Close();
                    }
                }
            }
            else
            {
                settingsPortalPage.SearchAndSelectUser(userMail);
                //settingsPortalPage.ConfigureUser(quantityByWeek, customerOrder, quantityByDay);
            }

            settingsPortalPage = homePage.GoToParameters_PortalPage();
            Assert.IsTrue(settingsPortalPage.IsExistingUser(userMail), "Le user n'est pas présent pour accéder au customer portal.");

            try
            {
                settingsPortalPage.SearchAndSelectUser(userMail);
                settingsPortalPage.SetLanguage("Others");
                settingsPortalPage.SaveAndConfirm();
                settingsPortalPage = homePage.GoToParameters_PortalPage();
                settingsPortalPage.SearchAndSelectUser(userMail);
                Assert.AreEqual(settingsPortalPage.GetLanguage(), "Others", "La langue n'est pas modifiée en Others");
                CustomerPortalLoginPage customerPortalLogin = LogInCustomerPortal(true);
                customerPortalLogin.LogIn(userMail, userPassword);

                bool isUserLogged = customerPortalLogin.IsUserLogged(userMail);
                bool isPageLoaded = customerPortalLogin.IsPageLoadedWithoutError();

                Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
                Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page.");
            }
            finally
            {
                //return to english language (to avoid traduction problems)

                homePage.Navigate();
                settingsPortalPage = homePage.GoToParameters_PortalPage();
                settingsPortalPage.SearchAndSelectUser(userMail);
                settingsPortalPage.SetLanguage("English");
                settingsPortalPage.SaveAndConfirm();

                //le save reset les delivery --> relink delivery to user
                string deliveryName = TestContext.Properties["DeliveryCustomerPortal"].ToString();
                string deliveryNameAFA = "A.F.A. CASAS";

                var portalSettingsPage = homePage.GoToParameters_PortalPage();
                portalSettingsPage.SearchAndSelectUser(userMail);
                portalSettingsPage.LinkDeliveryToUser(deliveryNameAFA);
                portalSettingsPage.LinkDeliveryToUser(deliveryName);
            }
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_user_french()
        {   //(test identique à PR_CO_CreateNewCustomerOrder mais sur Customer Portal)

            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();

            // boolean d'ouverture des nouveaux onglets
            bool mailboxTab = false;
            bool updatePasswordTab = false;

            var mailPage = new MailPage(WebDriver, TestContext);
            var updatePage = new CustomerPortalUpdatePage(WebDriver, TestContext);
            // Kept delivery name to filter for next tests.


            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var settingsPortalPage = homePage.GoToParameters_PortalPage();

            // si le user n'existe pas, on le créé
            if (!settingsPortalPage.IsExistingUser(userMail))
            {
                settingsPortalPage.AddNewUser(userMail);
                settingsPortalPage.SearchAndSelectUser(userMail);

                try
                {
                    // Connexion à la mailbox pour voir si le mail a été envoyé
                    mailPage = settingsPortalPage.RedirectToOutlookMailbox();
                    mailboxTab = true;

                    mailPage.FillFields_LogToOutlookMailbox(userMail);

                    // Recherche du mail envoyé pour le resend password
                    mailPage.ClickOnSpecifiedOutlookMail("Bienvenido a nuestro portal cliente");

                    updatePage = mailPage.ClickOnLinkForPassword("aquí");
                    updatePasswordTab = true;

                    var customerPortalLoginPage = updatePage.UpdatePassword(userPassword);
                    customerPortalLoginPage.Close();
                    updatePasswordTab = false;

                    mailPage.DeleteCurrentOutlookMail();
                    mailPage.Close();
                    mailboxTab = false;

                }
                finally
                {
                    // On ferme les onglets ouverts
                    if (updatePasswordTab)
                    {
                        updatePage.Close();
                    }

                    if (mailboxTab)
                    {
                        mailPage.Close();
                    }
                }
            }
            else
            {
                settingsPortalPage.SearchAndSelectUser(userMail);
                //settingsPortalPage.ConfigureUser(quantityByWeek, customerOrder, quantityByDay);
            }

            settingsPortalPage = homePage.GoToParameters_PortalPage();
            Assert.IsTrue(settingsPortalPage.IsExistingUser(userMail), "Le user n'est pas présent pour accéder au customer portal.");

            try
            {
                settingsPortalPage.SearchAndSelectUser(userMail);
                settingsPortalPage.SetLanguage("Français");
                settingsPortalPage.SaveAndConfirm();
                settingsPortalPage = homePage.GoToParameters_PortalPage();
                settingsPortalPage.SearchAndSelectUser(userMail);
                Assert.AreEqual(settingsPortalPage.GetLanguage(), "Français", "La langue n'est pas modifiée en Français");
                CustomerPortalLoginPage customerPortalLogin = LogInCustomerPortal(true);
                customerPortalLogin.LogIn(userMail, userPassword);

                bool isUserLogged = customerPortalLogin.IsUserLogged(userMail);
                bool isPageLoaded = customerPortalLogin.IsPageLoadedWithoutError();

                Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
                Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page.");
                customerPortalLogin.Close();
            }
            finally
            {
                //return to english language (to avoid traduction problems)

                homePage.Navigate();
                settingsPortalPage = homePage.GoToParameters_PortalPage();
                settingsPortalPage.SearchAndSelectUser(userMail);
                settingsPortalPage.SetLanguage("English");
                settingsPortalPage.SaveAndConfirm();

                //le save reset les delivery --> relink delivery to user
                string deliveryName = TestContext.Properties["DeliveryCustomerPortal"].ToString();
                string deliveryNameAFA = "A.F.A. CASAS";

                var portalSettingsPage = homePage.GoToParameters_PortalPage();
                portalSettingsPage.SearchAndSelectUser(userMail);
                portalSettingsPage.LinkDeliveryToUser(deliveryNameAFA);
                portalSettingsPage.LinkDeliveryToUser(deliveryName);
            }

        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_CO_index_comment()
        {
            //1.Aller sur Customer Portal, se logguer, aller sur Customer Orders
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userCustomerPortalPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();

            // Arrange
            var testComment = "testComment";
            CustomerPortalLoginPage customerPortalLoginPage = LogInCustomerPortal();
            customerPortalLoginPage.LogIn(userMail, userCustomerPortalPassword);

            bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
            bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();

            Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
            Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page.");

            // Rechercher un ordre éxistant.
            var customerPortalCustomerOrdersPage = customerPortalLoginPage.ClickOnCustomerOrders();
            customerPortalCustomerOrdersPage.ResetFilter();
            customerPortalCustomerOrdersPage.ShowFilterCustomersOrder("inProgress");

            Assert.IsTrue(customerPortalCustomerOrdersPage.CheckTotalNumber() > 0, "Aucun CO été trouvé");

            // Prendre le Delivery name de premier CO depuis la page General Information de CO.
            customerPortalCustomerOrdersPage.WaitPageLoading();
            var customerPortalCOPageResult = customerPortalCustomerOrdersPage.GoToFirstCustomer();
            customerPortalCOPageResult.GoToGeneralInformation();

            // Sauvegarder le numéro d'ordre.
            var orderNumber = customerPortalCOPageResult.GetGICustomerOrderNo();

            // Ajouter un commentaire à la customer order
            customerPortalCOPageResult.GeneralInfoFill("OrderComment", testComment, true);

            // revenir sur l'index et vérifier que le picto commentaire apparait
            customerPortalCustomerOrdersPage = customerPortalCOPageResult.ClickOnBackToCustomerOrderIndex();

            customerPortalCustomerOrdersPage.ResetFilter();
            customerPortalCustomerOrdersPage.Search(orderNumber);

            Assert.AreEqual(1, customerPortalCustomerOrdersPage.CheckTotalNumber(), $"Aucun CO été trouvé avec le numéro {orderNumber}.");

            Assert.IsTrue(customerPortalCustomerOrdersPage.FirstLineCommentIconIsVisible(), "L'icône de commentaire ne s'affiche pas.");

            // Supprimer le commentaire
            customerPortalCustomerOrdersPage.WaitPageLoading();
            customerPortalCOPageResult = customerPortalCustomerOrdersPage.GoToFirstCustomer();
            customerPortalCOPageResult.GoToGeneralInformation();
            customerPortalCOPageResult.GeneralInfoFill("OrderComment", " ", true);

            // vérifier que le picto a disparu dans l'index
            customerPortalCustomerOrdersPage = customerPortalCOPageResult.ClickOnBackToCustomerOrderIndex();

            customerPortalCustomerOrdersPage.ResetFilter();
            customerPortalCustomerOrdersPage.Search(orderNumber);

            Assert.AreEqual(1, customerPortalCustomerOrdersPage.CheckTotalNumber(), $"Aucun CO été trouvé avec le numéro {orderNumber}.");

            Assert.IsTrue(!customerPortalCustomerOrdersPage.FirstLineCommentIconIsVisible(), "L'icône de commentaire ne se cache pas.");

        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_CO_index_IHM()
        {
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string deliveryName = new Random().Next().ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();


            var customerPortalLoginPage = LogInCustomerPortal();
            customerPortalLoginPage.LogIn(userMail, userPassword);

            bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
            bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();

            Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
            Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page.");

            var customerPortalCustomerOrdersPage = customerPortalLoginPage.ClickOnCustomerOrders();
            customerPortalCustomerOrdersPage.ResetFilter();

            //2.Cliquer sur "+"
            var customerPortalCustmerOrdersPageModal = customerPortalCustomerOrdersPage.ClickAdd();


            //2.Remplir le formulaire avec le site(BCN), le customer(EMIRATES AILRINES) et l'aircraft (B777)
            customerPortalCustmerOrdersPageModal.SelectSite(site);

            customerPortalCustmerOrdersPageModal.SelectCustomer(customer);

            customerPortalCustmerOrdersPageModal.SelectDeliveryName(deliveryName);

            customerPortalCustmerOrdersPageModal.SelectAircraft(aircraft);


            //3.Cliquer sur "create"
            CustomerPortalCustomerOrdersPageResult customerPortalCustomerOrdersPageResult = customerPortalCustmerOrdersPageModal.Create();

            //4.Récupérer le numéro du customer order crée, faire Back to List
            //result customer order
            var resultText = customerPortalCustomerOrdersPageResult.ResultCustomerOrder();
            Regex r = new Regex("Customer order ([0-9]+) - EMIRATES AIRLINES");
            Match m = r.Match(resultText);
            string CustomerOrder = m.Groups[1].Value;

            // back to list
            customerPortalCustomerOrdersPage.Navigate("/CustomerPortal/CustomerOrders");

            //5.Vérifier le header de la table
            bool testHeader = customerPortalCustomerOrdersPage.CheckHeader();
            bool testColumns = customerPortalCustomerOrdersPage.CheckColumns("Id " + CustomerOrder, customer, deliveryName);

            Assert.AreEqual(testHeader, true);
            Assert.AreEqual(testColumns, true);

        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_CO_index_expedition_date()
        {
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();
            var homePage = LogInAsAdmin();
            var customerPortalLoginPage = homePage.GoToCustomerPortal();
            var orderNumber = string.Empty;
            var ExpeditionDate = string.Empty;
            try
            {
                customerPortalLoginPage.LogIn(userMail, userPassword);

                bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
                bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();

                Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
                Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page.");

                var customerPortalCustomerOrdersPage = customerPortalLoginPage.ClickOnCustomerOrders();

                orderNumber = customerPortalCustomerOrdersPage.getTableFirstLine(1);
                ExpeditionDate = customerPortalCustomerOrdersPage.getTableFirstLine(9);
                if (isUserLogged)
                {
                    customerPortalLoginPage.LogOut();
                }
            }
            finally
            {
                customerPortalLoginPage.Close();
            }

            var customerOrdersPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrdersPage.ResetFilter();
            customerOrdersPage.Filter(CustomerOrderPage.FilterType.Search, orderNumber.Replace("Id ", ""));
            var customerOrderFirstExpDate = customerOrdersPage.GetFirstExpeditionDate();
            Assert.AreEqual(ExpeditionDate.Trim(), customerOrderFirstExpDate.Trim());
            var customerOrderItem = customerOrdersPage.SelectFirstCustomerOrder();
            var customerOrderGeneralInformationPage = customerOrderItem.ClickOnGeneralInformationTab();
            var y = customerOrderGeneralInformationPage.GetExpeditionDate();
            Assert.AreEqual(ExpeditionDate.Trim(), customerOrderFirstExpDate.Trim());

        }
        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_CO_index_message()
        {
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string deliveryName = new Random().Next().ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userCustomerPortalPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();
            CustomerPortalCustomerOrdersPage customerPortalCustomerOrdersPage = null;
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            // Go to CustomerPortal Page
            CustomerPortalLoginPage customerPortalLoginPage = homePage.GoToCustomerPortal();
            try
            {
                customerPortalLoginPage.LogIn(userMail, userCustomerPortalPassword);
                bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
                Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
                customerPortalCustomerOrdersPage = customerPortalLoginPage.ClickOnCustomerOrders();
                // Créer un Customer order    
                var customerPortalCustmerOrdersPageModal = customerPortalCustomerOrdersPage.ClickAdd();
                //Remplir le formulaire avec le site(BCN), le customer(EMIRATES AILRINES) et l'aircraft (B777)
                customerPortalCustmerOrdersPageModal.SelectSite(site);
                customerPortalCustmerOrdersPageModal.SelectCustomer(customer);
                customerPortalCustmerOrdersPageModal.SelectDeliveryName(deliveryName);
                customerPortalCustmerOrdersPageModal.SelectAircraft(aircraft);
                //Cliquer sur "create"
                CustomerPortalCustomerOrdersPageResult customerPortalCustomerOrdersPageResult = customerPortalCustmerOrdersPageModal.Create();
                // récupérer le CustomerOrder
                var resultText = customerPortalCustomerOrdersPageResult.ResultCustomerOrder();
                Regex r = new Regex("Customer order ([0-9]+) - EMIRATES AIRLINES");
                Match m = r.Match(resultText);
                string CustomerOrder = m.Groups[1].Value;
                customerPortalCustomerOrdersPageResult.GoToMessage();
                CreateNewMessageGroupModal createNewMessageGroupModal = customerPortalCustomerOrdersPageResult.CreateNewMessage();
                createNewMessageGroupModal.SetTitleMessage("Message Customer Portal");
                createNewMessageGroupModal.SetMessage("Message Send From Customer Portal");
                // Deconnexion du portal
                customerPortalLoginPage.LogOut();
                customerPortalLoginPage.Close();
                var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
                customerOrderPage.ResetFilter();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, CustomerOrder);
                Assert.IsTrue(customerOrderPage.VerifyGreenPictoMessageInCutomerOrderWhenNotView(), "Le picto Message en Vert n'apparait pas bien dans la colonne message de l'index apres envoie message.");
                var customerOrderItem = customerOrderPage.SelectFirstCustomerOrder();
                CustomerOrderIMessagesPage customerOrderIMessagesPage = customerOrderItem.GoToMessagesTab();
                customerOrderIMessagesPage.ViewMessageSubject();
                customerOrderIMessagesPage.ViewMessage();
                bool isMassageDisplayed = customerOrderIMessagesPage.VerifyViewMessage();
                Assert.IsTrue(isMassageDisplayed, "Le picto Message en Vert n'apparait pas bien dans la colonne message de l'index.");
                customerOrderIMessagesPage.CloseMessageModal();
                customerOrderIMessagesPage.BackToList();
                customerOrderPage.ResetFilter();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, CustomerOrder);
                bool isGreyNotifDisplayed = customerOrderPage.VerifyGrayPictoMessageInCutomerOrderWhenIsView();
                Assert.IsTrue(isGreyNotifDisplayed, "Le picto Message en Gris n'apparait pas bien dans la colonne message de l'index apres view message.");
                customerOrderPage.DeleteCustomerOrder();
                // Winrest To Customer Portal
                var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customer, aircraft);
                var customerOrderDetailPage = customerOrderCreateModalPage.Submit();
                var generalInfo = customerOrderDetailPage.ClickOnGeneralInformationTab();
                var customerOrderNumber = generalInfo.GetOrderNumber();
                homePage.Navigate();
                // Go to CustomerPortal Page
                customerPortalLoginPage = homePage.GoToCustomerPortal();
                customerPortalLoginPage.LogIn(userMail, userCustomerPortalPassword);
                customerPortalCustomerOrdersPage = customerPortalLoginPage.ClickOnCustomerOrders();
                customerPortalCustomerOrdersPage.Search(customerOrderNumber);
                bool isPictoMessage = customerPortalCustomerOrdersPage.VerifyPictoMessageInCutomerOrderPortal();
                Assert.IsFalse(isPictoMessage, "Le Picto ne doit pas apparaitre si le message est envoyé depuis Winrest");
            }
            finally
            {
                customerPortalCustomerOrdersPage.DeleteCustomerOrder();
            }
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_CO_messagerie_statut()
        {
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userCustomerPortalPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();
            CustomerOrderPage customerOrderPage = null;
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            try
            {
                customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
                var customerOrderCreateModalPage = customerOrderPage.CustomerOrderCreatePage();
                customerOrderCreateModalPage.FillField_CreatNewCustomerOrder(site, customer, aircraft);
                var customerOrderItem = customerOrderCreateModalPage.Submit();
                var generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
                var customerOrderNumber = generalInfo.GetOrderNumber();
                CustomerOrderIMessagesPage customerOrderIMessagesPage = customerOrderItem.GoToMessagesTab();
                customerOrderIMessagesPage.AddNewMessage("Message Customer Order To Portal", "Message Send From Customer Order To Portal");
                // Go to CustomerPortal Page
                homePage.Navigate();
                CustomerPortalLoginPage customerPortalLoginPage = homePage.GoToCustomerPortal();
                customerPortalLoginPage.LogIn(userMail, userCustomerPortalPassword);
                bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
                Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
                CustomerPortalCustomerOrdersPage customerPortalCustomerOrdersPage = customerPortalLoginPage.ClickOnCustomerOrders();
                customerPortalCustomerOrdersPage.Search(customerOrderNumber);
                CustomerPortalCustomerOrdersPageResult customerPortalCustomerOrdersPageResult = customerPortalCustomerOrdersPage.GoToFirstCustomer();
                CustomerOrderPortalMessagesPage customerOrderPortalMessagesPage = customerPortalCustomerOrdersPageResult.GoToMessagesTab();
                customerOrderPortalMessagesPage.ViewMessage();
                bool isMessageDisplayed = customerOrderPortalMessagesPage.VerifyViewMessage();
                Assert.IsTrue(isMessageDisplayed, "Le Message envoyer depuis Customer Order n'apparait pas dans Customer Order Portal.");
                customerOrderPortalMessagesPage.AnswerMessage("Message Send From Customer Portal");
                // Deconnexion du portal
                customerPortalLoginPage.LogOut();
                customerPortalLoginPage.Close();
                customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
                customerOrderPage.ResetFilter();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrderNumber);
                bool isGreenPictoDisplayed = customerOrderPage.VerifyGreenPictoMessageInCutomerOrderWhenNotView();
                Assert.IsTrue(isGreenPictoDisplayed, "Le picto Message en Vert n'apparait pas bien dans la colonne message de l'index apres envoie message.");
                customerOrderPage.SelectFirstCustomerOrder();
                customerOrderIMessagesPage = customerOrderItem.GoToMessagesTab();
                customerOrderIMessagesPage.ViewMessageBody();
                bool isMessageSent = customerOrderIMessagesPage.VerifyViewMessage();
                Assert.IsTrue(isMessageSent, "Le Message envoyer depuis Customer Order Portal n'apparait pas dans Customer Order.");
                customerOrderIMessagesPage.CloseMessageModal();
                customerOrderIMessagesPage.BackToList();
                customerOrderPage.ResetFilter();
                customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, customerOrderNumber);
                bool isGreyNotifDisplyaed = customerOrderPage.VerifyGrayPictoMessageInCutomerOrderWhenIsView();
                Assert.IsTrue(isGreyNotifDisplyaed, "Le picto Message en Gris n'apparait pas bien dans la colonne message de l'index apres view message.");

            }
            finally
            {
                customerOrderPage.DeleteCustomerOrder();
            }
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_CO_ModificationSuppressionImpossible_LastUpdate()
        {
            string userMail = TestContext.Properties["Admin_UserName"].ToString();
            string userCustomerPortalPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();
            CustomerPortalLoginPage customerPortalLoginPage = LogInCustomerPortal();
            customerPortalLoginPage.LogIn(userMail, userCustomerPortalPassword);
            bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
            bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();
            Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
            Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page.");
            CustomerPortalCustomerOrdersPage customerPortalCustomerOrdersPage = customerPortalLoginPage.ClickOnCustomerOrders();
            DateTime dateFrom = new DateTime(2024, 7, 31);
            DateTime dateTo = new DateTime(2024, 8, 8);
            customerPortalCustomerOrdersPage.FilterDate(dateFrom, dateTo);
            customerPortalCustomerOrdersPage.ShowFilterCustomersOrder("inProgress", true);
            CustomerPortalCustomerOrdersPageResult customerPortalCustomerOrdersPageResult = customerPortalCustomerOrdersPage.ClickOnFirstCutomerOrder();
            customerPortalCustomerOrdersPageResult.GoToGeneralInformation();
            DateTime lastUpdateDate = customerPortalCustomerOrdersPageResult.GetLastUpdateDate();
            bool isGreater = DateTime.Now.Date > lastUpdateDate;
            Assert.IsTrue(isGreater, $"Last update ({lastUpdateDate}) n'est pas < à la date du jour ({DateTime.Now})");
            customerPortalCustomerOrdersPageResult.GoToDetails();
            bool hasData = customerPortalCustomerOrdersPageResult.DetailPageHasData();
            Assert.IsTrue(hasData, "No items");
            bool isEditable = customerPortalCustomerOrdersPageResult.ItemIsEditable();
            Assert.IsFalse(isEditable, "Possibilité de faire des modifications sur les items");
            customerPortalCustomerOrdersPage = customerPortalCustomerOrdersPageResult.BackToList();
            customerPortalCustomerOrdersPage.FilterDate(dateFrom, dateTo);
            customerPortalCustomerOrdersPage.ShowFilterCustomersOrder("inProgress", true);
            bool hasDeleteIcon = customerPortalCustomerOrdersPage.CanBeDeleted();
            Assert.IsFalse(hasDeleteIcon, "l'un des COs a une icone de suppression");
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_CO_AffichageŒil_ApresModifDeliveryDate()
        {
            // Prepare
            string customerName = TestContext.Properties["InvoiceCustomer"].ToString();
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string deliveryName = new Random().Next().ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string comment = "This is a comment";
            DateTime date = DateUtils.Now.AddDays(10);
            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var customerPortalLoginPage = homePage.GoToCustomerPortal();
            customerPortalLoginPage.LogIn(userMail, userPassword);
            var customerPortalCustomerOrdersPage = customerPortalLoginPage.ClickOnCustomerOrders();

            var customerPortalCustomerOrdersPageModal = customerPortalCustomerOrdersPage.ClickAdd();

            customerPortalCustomerOrdersPageModal.SelectSite(site);
            customerPortalCustomerOrdersPageModal.SelectCustomer(customer);
            customerPortalCustomerOrdersPageModal.SelectDeliveryName(deliveryName);
            customerPortalCustomerOrdersPageModal.SelectAircraft(aircraft);

            CustomerPortalCustomerOrdersPageResult customerPortalCustomerOrdersPageResult = customerPortalCustomerOrdersPageModal.Create();
            var CustomerOrderText = customerPortalCustomerOrdersPageResult.ResultCustomerOrder();

            string orderNumber = CustomerOrderText.Replace("Customer order", "").Replace("- EMIRATES AIRLINES", "").Trim();
            customerPortalLoginPage.LogOut();
            customerPortalLoginPage.Close();

            var customerOrderPage = homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, orderNumber);
            Assert.IsTrue(customerOrderPage.IsEyeIconGrayForOrder(), "The eye icon should be green after verification.");
            var customerOrderItem = customerOrderPage.ClickFirstLine();
            customerOrderItem.DoVerify();
            homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, orderNumber);
            Assert.IsTrue(customerOrderPage.IsEyeIconGreenForOrder(), "The eye icon should be green after verification.");
            customerOrderPage.ClickFirstLine();
            var generalInfo = customerOrderItem.ClickOnGeneralInformationTab();
            generalInfo.UpdateGeneralInfo(deliveryName, date, comment);
            generalInfo.ClickOnDetailTab();
            homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, orderNumber);
            customerOrderPage.ClickFirstLine();
            var verifyVisible = customerOrderItem.IsVerifyButtonVisible();
            Assert.IsTrue(verifyVisible, "Le bouton 'Verify' n'a pas été trouvé sur la page.");
            homePage.GoToProduction_CustomerOrderPage();
            customerOrderPage.Filter(CustomerOrderPage.FilterType.Search, orderNumber);
            Assert.IsTrue(customerOrderPage.IsEyeIconGrayForOrder(), "The eye icon should be green after verification.");


        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_CO_AffichageCommercialName()
        {

            string customerOrderNumberId = "";

            //(test identique à PR_CO_CreateNewCustomerOrder mais sur Customer Portal)

            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();
            // Prepare
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString();
            string deliveryName = new Random().Next().ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            string ItemName = "DUVET BC";
            // Kept delivery name to filter for next tests.
            _deliceryNamOfCreateCustomerOrder = deliveryName;

            // Arrange
            CustomerPortalLoginPage customerPortalLoginPage = LogInCustomerPortal();
            customerPortalLoginPage.LogIn(userMail, userPassword);

            bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
            bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();

            Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
            Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page.");

            var customerPortalCustomerOrdersPage = customerPortalLoginPage.ClickOnCustomerOrders();
            try
            {

                //2.Cliquer sur "+"
                var customerPortalCustmerOrdersPageModal = customerPortalCustomerOrdersPage.ClickAdd();


                //2.Remplir le formulaire avec le site(BCN), le customer(EMIRATES AILRINES) et l'aircraft (B777)
                customerPortalCustmerOrdersPageModal.SelectSite(site);

                customerPortalCustmerOrdersPageModal.SelectCustomer(customer);

                customerPortalCustmerOrdersPageModal.SelectDeliveryName(deliveryName);

                customerPortalCustmerOrdersPageModal.SelectAircraft(aircraft);


                //3.Cliquer sur "create"
                CustomerPortalCustomerOrdersPageResult customerPortalCustomerOrdersPageResult = customerPortalCustmerOrdersPageModal.Create();

                //4.Récupérer le numéro du customer order crée, faire Back to List
                //result customer order
                var CustomerOrderText = customerPortalCustomerOrdersPageResult.ResultCustomerOrder();

                // back to list
                customerPortalCustomerOrdersPage.Navigate("/CustomerPortal/CustomerOrders");

                customerOrderNumberId = Regex.Match(CustomerOrderText, @"\d+").Value;
                string customerPortalCustomerOrderId = Regex.Match(customerPortalCustomerOrdersPage.getTableFirstLine(1), @"\d+").Value;
                //5.Vérifier dans la table si on retrouve le customer order créé
                Assert.AreEqual(customerOrderNumberId, customerPortalCustomerOrderId);
                customerPortalCustomerOrdersPage.Search(customerOrderNumberId);
                customerPortalCustomerOrdersPage.WaitLoading();
                customerPortalCustomerOrdersPage.ClickOnFirstCutomerOrder();
                customerPortalCustomerOrdersPage.WaitLoading();
                customerPortalCustomerOrdersPage.AddItem(ItemName);

                customerPortalCustomerOrdersPageResult.GoToGeneralInformation();
                customerPortalCustomerOrdersPageResult.GoToDetails();
                bool hasData = customerPortalCustomerOrdersPageResult.DetailPageHasData();
                Assert.IsTrue(hasData, "No items");

                bool isCommercialNameExist = customerPortalCustomerOrdersPageResult.CommercialNameExist(ItemName);
                Assert.IsTrue(hasData, "Il faut que le commercial name de l'item s'afficher");
                customerPortalCustomerOrdersPageResult.BackToList();
            }
            finally
            {
                if (!string.IsNullOrEmpty(customerOrderNumberId))
                {
                    customerPortalCustomerOrdersPage.WaitLoading();
                    customerPortalCustomerOrdersPage.Search(customerOrderNumberId);
                    customerPortalCustomerOrdersPage.WaitLoading();
                    customerPortalCustomerOrdersPage.DeleteFirstItem();
                }


            }

        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_QBW_DuplicateQuantity()
        {
            //(test identique à PR_CO_CreateNewCustomerOrder mais sur Customer Portal)
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();

            // Prepare
            string deliveryName = new Random().Next().ToString();
            string qty = "27";
            string customerName = "ServiceForCustomerPortal";
            // Kept delivery name to filter for next tests.
            _deliceryNamOfCreateCustomerOrder = deliveryName;

            // Arrange
            CustomerPortalLoginPage customerPortalLoginPage = LogInCustomerPortal();
            customerPortalLoginPage.LogIn(userMail, userPassword);

            bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
            bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();

            Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
            Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page.");

            var customerPortalQtyByWeekPage = customerPortalLoginPage.ClickOnQtyByWeek();

            //Avoir des données disponibles. (fonctionne sur dev5 site PMI)
            var counter = customerPortalQtyByWeekPage.GetCounter();
            Assert.IsTrue(counter > 0, "pas de données disponibles");

            //Modifier la quantité d'une journée
            customerPortalQtyByWeekPage.Search(customerName);

            //ajouter une étape où on valide la quantité de demain puis on change la quantité et revalider, vérifier

            CultureInfo.CurrentCulture = new CultureInfo("en-US", false);
            var value = customerPortalQtyByWeekPage.GetDateToUpdate();
            Days dayEnum = (Days)Enum.Parse(typeof(Days), value);
            var date = customerPortalQtyByWeekPage.GetCaptionWeeekContent();
            var dateToUpdate = customerPortalQtyByWeekPage.GetDateOfInput(dayEnum, date);

            if (DateUtils.Now.ToString("dddd") != "Sunday")
            {
                customerPortalQtyByWeekPage.SetQuantities(dateToUpdate.ToString("dddd"), qty);
                customerPortalQtyByWeekPage.Validate((int)dateToUpdate.DayOfWeek);

                //Cliquer sur icone dupliquer(à droite du jour) et sélectionner le jour à dupliquer
                customerPortalQtyByWeekPage.DuplicateOneDay(dateToUpdate.ToString("ddd"), (int)dateToUpdate.DayOfWeek);

                //Vérifier que la quantité et le commentaire ont bien été dupliqué
                customerPortalQtyByWeekPage.Search(customerName);
                string qtyDuplicate = customerPortalQtyByWeekPage.GetQuantity(dateToUpdate.AddDays(1).ToString("dddd").ToLower());

                //ASSERT
                Assert.AreEqual(qty, qtyDuplicate, "la duplication des quantités d'un jour à un autre doit se faire sans erreur.");
            }
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CUPO_QBD_ValidateButton()
        {
            // Prepare
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();
            string deliveryName = "DeliveryForCustomerPortal";
            string qty = "15";
            string serviceName = "ServiceForCustomerPortal";
            string site = TestContext.Properties["Site"].ToString(); // MAD

            // Arrange
            
            var homePage = LogInAsAdmin();

            var loadingPage = new DeliveryLoadingPage(WebDriver, TestContext);

            // Create delivery
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.ResetFilter();
            deliveryPage.Filter(DeliveryPage.FilterType.Sites, site);
            deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName);
            loadingPage = deliveryPage.ClickOnFirstDelivery();
            var deliveryGeneralInfo = loadingPage.ClickOnGeneralInformation();
            //deliveryGeneralInfo.SetToDefault();
            deliveryGeneralInfo.SetStartingTimeToLimitedAccess("24");
            deliveryPage = deliveryGeneralInfo.BackToList();

            homePage.Navigate();
            var customerPortalLoginPage = homePage.GoToCustomerPortal();
            try
            {
                customerPortalLoginPage.LogIn(userMail, userPassword);

                bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
                bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();

                Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
                Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page Quantity By Week.");

                var customerPortalQtyByDayPage = customerPortalLoginPage.ClickOnQtyByDay();

                customerPortalQtyByDayPage.Search(deliveryName);
                Assert.AreEqual(serviceName, customerPortalQtyByDayPage.SearchFirstResult(), "ServiceForCustomerPortal non trouvé dans le tableau");
                var demain = DateTime.Now.AddDays(1);
                customerPortalQtyByDayPage.GoToDate(demain);
                var dayName = customerPortalQtyByDayPage.GetDayName();
                var isModifiable = customerPortalQtyByDayPage.CheckIfModifiableByDay(dayName, serviceName);
                if(isModifiable) customerPortalQtyByDayPage.SetQuantities(qty);
                Assert.IsTrue(isModifiable, "n'est pas modifiable");

                // Démarrer le chronomètre
                var stopwatch = new Stopwatch();
                stopwatch.Start();
                customerPortalQtyByDayPage.Validate();
                stopwatch.Stop();

                //Assert - Vérification que l'opération ne prend pas trop de temps
                double maxAllowedTime = 5000;
                var checkDelay = stopwatch.ElapsedMilliseconds < maxAllowedTime;
                Assert.IsTrue(checkDelay, $"L'opération a pris trop de temps : {stopwatch.ElapsedMilliseconds}ms, alors que le maximum autorisé est {maxAllowedTime}ms.");

            }
            finally
            {
                customerPortalLoginPage.Close();
            }

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CUPO_CO_Export_Orders_StatusInvoiced()
        {
            string customer = TestContext.Properties["InvoiceCustomer"].ToString();
            string site = TestContext.Properties["InvoiceSite"].ToString(); //BCN
            string deliveryName = new Random().Next().ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            var userMail = TestContext.Properties["Admin_UserName"].ToString();
            var userCustomerPortalPassword = TestContext.Properties["CustomerPortal_UserPassword"].ToString();

            HomePage homePage = LogInAsAdmin();
            CustomerPortalLoginPage customerPortalLoginPage = homePage.GoToCustomerPortal();
            customerPortalLoginPage.LogIn(userMail, userCustomerPortalPassword);
            bool isUserLogged = customerPortalLoginPage.IsUserLogged(userMail);
            bool isPageLoaded = customerPortalLoginPage.IsPageLoadedWithoutError();
            Assert.IsTrue(isUserLogged, "La connexion au customerPortal depuis la page d'accueil Winrest a échoué.");
            Assert.IsTrue(isPageLoaded, "Erreur lors du chargement de la page.");

            CustomerPortalCustomerOrdersPage customerPortalCustomerOrdersPage = customerPortalLoginPage.ClickOnCustomerOrders();
            customerPortalCustomerOrdersPage.ShowFilterCustomersOrder("invoiced");
            int invoiced = customerPortalCustomerOrdersPage.CheckTotalNumber();
            if (invoiced == 0)
            {
                CustomerPortalCustomerOrdersPageModal customerPortalCustmerOrdersPageModal = customerPortalCustomerOrdersPage.ClickAdd();
                customerPortalCustmerOrdersPageModal.SelectSite(site);
                customerPortalCustmerOrdersPageModal.SelectCustomer(customer);
                customerPortalCustmerOrdersPageModal.SelectDeliveryName(deliveryName);
                customerPortalCustmerOrdersPageModal.SelectAircraft(aircraft);
                CustomerPortalCustomerOrdersPageResult customerPortalCustomerOrdersPageResult = customerPortalCustmerOrdersPageModal.Create();
                customerPortalCustomerOrdersPageResult.AddLine("HOT MEAL YC UAE", 10);
                customerPortalCustomerOrdersPageResult.WaitLoading();
                customerPortalCustomerOrdersPageResult.Validate();
                customerPortalCustomerOrdersPage.Close();
                InvoicesPage invoicesPage = homePage.GoToAccounting_InvoicesPage();
                AutoInvoiceCreateModalPage autoInvoiceCreateModalPage = invoicesPage.AutoInvoiceCreatePage();
                autoInvoiceCreateModalPage.FillField_CreateNewAutoInvoice(customer, site, CustomerPickMethod.AllCustomerOrdersInSelectedPeriod);
                autoInvoiceCreateModalPage.FillFieldSelectAll();
                invoicesPage.ValidateInvoice();
                homePage.Navigate();
                customerPortalLoginPage = homePage.GoToCustomerPortal();
                customerPortalCustomerOrdersPage = customerPortalLoginPage.ClickOnCustomerOrders();
                customerPortalCustomerOrdersPage.ShowFilterCustomersOrder("invoiced");
            }
            string firstCustomer = customerPortalCustomerOrdersPage.getTableFirstLine(1);
            string CustomerId = firstCustomer.Split(' ')[1];
            customerPortalCustomerOrdersPage.Search(CustomerId);
            customerPortalCustomerOrdersPage.ExportExcel();
            string directory = TestContext.Properties["DownloadsPath"].ToString();
            foreach (string f in Directory.GetFiles(directory))
            {
                if (f.Contains("export-orders") && f.EndsWith(".xlsx"))
                {
                    File.Delete(f);
                }
            }

            customerPortalCustomerOrdersPage.ExportExcel();
            Thread.Sleep(4000);
            string trouve = null;
            foreach (string f in Directory.GetFiles(directory))
            {
                if (f.Contains("export-orders") && f.EndsWith(".xlsx"))
                {
                    trouve = f;
                }
            }
            Assert.IsNotNull(trouve);
            int resultNumber = OpenXmlExcel.GetExportResultNumber("Sheet 1", trouve);
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
            string customerOrder = OpenXmlExcel.GetValuesInList("Order ID", "Sheet 1", trouve)[0];
            Assert.AreEqual(CustomerId, customerOrder, " excel colonne Order ID");
            string isInvoiced = OpenXmlExcel.GetValuesInList("Invoice step", "Sheet 1", trouve)[0];
            Assert.AreEqual("Invoiced", isInvoiced, " excel colonne Order ID");

        }
    }
}