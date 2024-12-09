using DocumentFormat.OpenXml.Office2010.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Accounting.Edi;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Customer;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Delivery;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.DeliveryRound;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.Dispatch;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Policy;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace Newrest.Winrest.FunctionalTests.Production
{
    [TestClass]
    public class DispatchTest : TestBase
    {
        private const int _timeout = 600000;
        readonly string customerName = "Customer for Dispatch";
        readonly string serviceName = "Service for Dispatch";
        readonly string customerCode = "CFD";
        readonly string serviceNameToday = "Service-" + DateUtils.Now.ToString("dd/MM/yyyy");
        readonly string deliveryNameToday = "Delivery-" + DateUtils.Now.ToString("dd/MM/yyyy");

        // Créer un nouveau client-31-
        [Priority(0)]
        [TestMethod]
        [Timeout(_timeout)]
        public void PR_DISP_Create_New_Customer()
        {
            // Prepare
            string customerType = TestContext.Properties["CustomerType3"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            ClearCache();

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
            var firstCustomerName = customersPage.GetFirstCustomerName();
            Assert.AreEqual(customerName, firstCustomerName, "Le customer n'a pas été créé.");
        }

        [TestMethod]
        [Priority(1)]
        [Timeout(_timeout)]
        public void PR_DISP_Create_Service()
        {
            // Prepare
            string siteACE = TestContext.Properties["Production_Site1"].ToString();

            string serviceCategorie = TestContext.Properties["Production_Service1"].ToString();
            string customerReference = customerCode + " - " + customerName;
            string serviceType = TestContext.Properties["Production_ServiceType"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var servicePage = homePage.GoToCustomers_ServicePage();

            // service 1
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);

            if (servicePage.CheckTotalNumber() == 0)
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, null, null, serviceCategorie, null, serviceType);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                serviceGeneralInformationsPage.SetProduced(true);

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(siteACE, customerReference, DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2), "3");
                servicePage = pricePage.BackToList();
            }
            else
            {
                var pricePage = servicePage.ClickOnFirstService();
                pricePage.ToggleFirstPrice();
                var serviceCreatePriceModalPage = pricePage.EditFirstPrice(siteACE, customerName);
                serviceCreatePriceModalPage.EditPriceDates(DateUtils.Now.AddDays(-10), DateUtils.Now.AddMonths(2));
                if (!serviceCreatePriceModalPage.VerifySuccessEditPriceDates())
                {
                    serviceCreatePriceModalPage.Close();
                }

                var serviceGeneralInformationsPage = pricePage.ClickOnGeneralInformationTab();
                serviceGeneralInformationsPage.SetProduced(true);
                servicePage = serviceGeneralInformationsPage.BackToList();
            }

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, serviceName);
            Assert.IsTrue(servicePage.GetFirstServiceName().Contains(serviceName), "Le service " + serviceName + " n'existe pas.");
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

            string recipeVariant;
            recipeVariant = TestContext.Properties["Production_RecipeVariant1ForACE"].ToString();

            // Prepare menu
            if (menuName == null)
            {
                menuName = "MenuDispatch-" + rnd.Next().ToString();
                recipeName = "RecipeDispatch-" + rnd.Next().ToString();
            }
            string variant = TestContext.Properties["MenuVariantACE1"].ToString();

            if (isWithMenu)
            {
                //2. Creation de la recette
                if (menuName != null)
                {
                    var recipesPage = homePage.GoToMenus_Recipes();
                    var recipesCreateModalPage = recipesPage.CreateNewRecipe();
                    var recipeGeneralInfosPage = recipesCreateModalPage.FillField_CreateNewRecipe(recipeName, recipeType, nbPortions.ToString());
                    recipeGeneralInfosPage.AddVariantWithSite(deliverySite, recipeVariant);
                    var recipeVariantPage = recipeGeneralInfosPage.SelectFirstVariantFromList();
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
                var deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRound(deliveryRoundName, deliverySite, DateUtils.Now, DateUtils.Now.AddDays(+31));

                var deliveryRoundDeliveriesPage = deliveryRoundGeneralInfoPage.ClickOnDeliveryTab();
                deliveryRoundDeliveriesPage.AddDelivery(deliveryName);
                deliveryRoundDeliveriesPage.BackToList();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_DISP_Filter_Search_Delivery()
        {
            Random rnd = new Random();
            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();
            string Site = TestContext.Properties["SiteACE"].ToString();
            string qty = "100";

            // Arrange
            var homePage = LogInAsAdmin();
            DeliveryPage deliveryPage = homePage.GoToCustomers_DeliveryPage();
            try
            {
                DeliveryCreateModalPage deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, customerName, Site, true);
                DeliveryLoadingPage loadingPage = deliveryCreateModalPage.Create();
                loadingPage.AddService(serviceName);
                loadingPage.AddQuantity(qty);
                homePage.Navigate();
                DispatchPage dispatchPage = homePage.GoToProduction_DispatchPage();
                dispatchPage.Filter(DispatchPage.FilterType.Site, Site);
                dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryName);
                var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
                var firstNameDelivery = previsionnalQty.GetFirstDeliveryName();
                Assert.AreEqual(deliveryName, firstNameDelivery, MessageErreur.FILTRE_ERRONE, "Search by deliveryName");
            }
            finally
            {
                deliveryPage = homePage.GoToCustomers_DeliveryPage();
                DeliveryMassiveDeletePage deliveryMassiveDeletePage = deliveryPage.MassiveDelete();
                deliveryMassiveDeletePage.SearchDeliveryName(deliveryName);
                deliveryMassiveDeletePage.ClickSearchButton();
                deliveryMassiveDeletePage.SelectAll();
                deliveryMassiveDeletePage.ClickDeleteButton();
                deliveryMassiveDeletePage.ClickConfirmDeleteButton();
            }

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_DISP_Filter_Search_Service()
        {
            // Prepare
            Random rnd = new Random();
            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();
            string Site = TestContext.Properties["SiteACE"].ToString();
            string qty = "100";


            var homePage = LogInAsAdmin();
            var dispatchPage = homePage.GoToProduction_DispatchPage();
            try
            {
                dispatchPage.Filter(DispatchPage.FilterType.Site, Site);
                dispatchPage.Filter(DispatchPage.FilterType.Search, serviceName);
                if (dispatchPage.CheckTotalNumber() == 0)
                {
                    DeliveryPage deliveryPage = homePage.GoToCustomers_DeliveryPage();
                    DeliveryCreateModalPage deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                    deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, customerName, Site, true);
                    DeliveryLoadingPage loadingPage = deliveryCreateModalPage.Create();
                    loadingPage.AddService(serviceName);
                    loadingPage.AddQuantity(qty);
                    homePage.Navigate();
                    dispatchPage = homePage.GoToProduction_DispatchPage();
                    dispatchPage.Filter(DispatchPage.FilterType.Site, Site);
                    dispatchPage.Filter(DispatchPage.FilterType.Search, serviceName);
                }
                var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
                Assert.AreEqual(serviceName, previsionnalQty.GetFirstServiceName(), MessageErreur.FILTRE_ERRONE, "Search by serviceName");
            }
            finally
            {
                DeliveryPage deliveryPage = homePage.GoToCustomers_DeliveryPage();
                deliveryPage = homePage.GoToCustomers_DeliveryPage();
                deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName);
                if (deliveryPage.CheckTotalNumber() > 0)
                {
                    DeliveryMassiveDeletePage deliveryMassiveDeletePage = deliveryPage.MassiveDelete();
                    deliveryMassiveDeletePage.SearchDeliveryName(deliveryName);
                    deliveryMassiveDeletePage.ClickSearchButton();
                    deliveryMassiveDeletePage.SelectAll();
                    deliveryMassiveDeletePage.ClickDeleteButton();
                    deliveryMassiveDeletePage.ClickConfirmDeleteButton();
                }
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_DISP_Filter_Site()
        {
            string deliverySite = TestContext.Properties["SiteACE"].ToString();

            var homePage = LogInAsAdmin();

            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Site, deliverySite);
            var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            ServicePricePage servicePricePage = previsionnalQty.ClickFirstService();

            Assert.IsTrue(servicePricePage.GetPriceName().Contains(deliverySite), MessageErreur.FILTRE_ERRONE, "Site");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_DISP_Filter_SortByDelivery()
        {

            //Random rnd = new Random();
            // Prepare delivery 
            //string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();           
            string Site = TestContext.Properties["SiteACE"].ToString();
            //string qty = "100";

            var homePage = LogInAsAdmin();

            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.ResetFilter();
            dispatchPage.Filter(DispatchPage.FilterType.Site, Site);
            dispatchPage.Filter(DispatchPage.FilterType.Customers, customerName);

            //if (dispatchPage.CheckTotalNumber() < 20)
            //{
            //    DeliveryPage deliveryPage = homePage.GoToCustomers_DeliveryPage();
            //    DeliveryCreateModalPage deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
            //    deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, customerName, Site, true);
            //    DeliveryLoadingPage loadingPage = deliveryCreateModalPage.Create();
            //    loadingPage.AddService(serviceName);
            //    loadingPage.AddQuantity(qty);
            //    homePage.Navigate();
            //    dispatchPage = homePage.GoToProduction_DispatchPage();
            //dispatchPage.ResetFilter();
            //dispatchPage.Filter(DispatchPage.FilterType.Site, Site);
            //dispatchPage.Filter(DispatchPage.FilterType.Customers, customerName);
            //}
            dispatchPage.Filter(DispatchPage.FilterType.SortBy, "DeliveryName");
            var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();

            if (!previsionnalQty.isPageSizeEqualsTo100())
            {
                previsionnalQty.PageSize("8");
                previsionnalQty.PageSize("100");
            }
            var DeliveriesNames = previsionnalQty.GetDeliveryName();
            Assert.IsTrue(DeliveriesNames, "les résultats ne s'accordent pas bien au filtre appliqué");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_DISP_Filter_SortByService()
        {
            //Random rnd = new Random();
            // Prepare delivery 
            //string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();           
            //string qty = "100";
            string Site = TestContext.Properties["SiteACE"].ToString();

            var homePage = LogInAsAdmin();
            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.ResetFilter();
            dispatchPage.Filter(DispatchPage.FilterType.Site, Site);
            dispatchPage.Filter(DispatchPage.FilterType.Customers, customerName);

            //if (dispatchPage.CheckTotalNumber() < 20)
            //{
            //    DeliveryPage deliveryPage = homePage.GoToCustomers_DeliveryPage();
            //    DeliveryCreateModalPage deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
            //    deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, customerName, Site, true);
            //    DeliveryLoadingPage loadingPage = deliveryCreateModalPage.Create();
            //    loadingPage.AddService(serviceName);
            //    loadingPage.AddQuantity(qty);
            //    homePage.Navigate();
            //    dispatchPage = homePage.GoToProduction_DispatchPage();
            //dispatchPage.ResetFilter();
            //dispatchPage.Filter(DispatchPage.FilterType.Site, Site);
            //dispatchPage.Filter(DispatchPage.FilterType.Customers, customerName);
            //}
            dispatchPage.Filter(DispatchPage.FilterType.SortBy, "ServiceName");
            var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();

            previsionnalQty.PageSize("8");
            var listService = previsionnalQty.GetAllService();
            Assert.IsTrue(listService.SequenceEqual(listService.OrderBy(service => service)) || listService.SequenceEqual(listService.OrderByDescending(service => service)), "Le service n'est pas en ordre alphabétique");

        }

        /*[Timeout(_timeout)]
        [TestMethod]
        public void PR_DISP_Filter_SortByDefault()
        {
            //Ce test a été supprimé après discussion avec les responsables, car il ne correspond pas à la description et ne présente pas de valeur ajoutée. 
             Random rnd = new Random();

             // Prepare delivery 
             string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();
             string deliverySite = TestContext.Properties["SiteLP"].ToString();
             bool isActivated = true;
             string qty = "1546";

             string deliveryCustomer = TestContext.Properties["CustomerDelivery"].ToString();
             string deliveryCustomer1 = "AMPA TOMARE CEIP EL QUINTERO";
             string deliveryCustomer2 = "AYTO.DE TEGUISE-GUARDERIA MUNICIPAL";

             List<string> customerList = new List<string>();
             customerList.Add(deliveryCustomer);
             customerList.Add(deliveryCustomer1);
             customerList.Add(deliveryCustomer2);


             // Prepare service           
             string serviceName = serviceNameToday + " - " + rnd.Next().ToString();
             string serviceCode = rnd.Next().ToString();
             string serviceProduction = GenerateName(4);

             // Prepare menu
             string menuName = "MenuDispatch-" + rnd.Next().ToString();
             string variant = TestContext.Properties["MenuVariantACE1"].ToString();

             // Arrange
             var homePage = LogInAsAdmin();

             // Accès aux dispatch
             var dispatchPage = homePage.GoToProduction_DispatchPage();
             dispatchPage.Filter(DispatchPage.FilterType.SortBy, "Default");
             dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryName);
             var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();

             // Récupération de la liste des customers associés aux dispatch
             var customerNames = previsionnalQty.GetCustomer();

             var customerPage = homePage.GoToCustomers_CustomerPage();
             List<string> finalCustomerNames = new List<string>();

             for (int i = 0; i < customerNames.Count; i++)
             {
                 customerPage.Filter(CustomerPage.FilterType.SortBy, "CUSTOMER");
                 customerPage.Filter(CustomerPage.FilterType.Search, customerNames[i]);
                 var customerName = customerPage.GetFirstCustomerName();
                 finalCustomerNames.Add(customerName);
             }
             var expectedList = finalCustomerNames.OrderBy(x => x).ToList();

             //Assert
             Assert.IsTrue(expectedList.SequenceEqual(finalCustomerNames));
            
        }*/

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_DISP_Filter_ShowForValidation()
        {

            Random rnd = new Random();
            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();
            string Site = TestContext.Properties["SiteACE"].ToString();
            string qty = "100";

            // Arrange
            var homePage = LogInAsAdmin();
            DeliveryPage deliveryPage = homePage.GoToCustomers_DeliveryPage();
            try
            {
                DeliveryCreateModalPage deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, customerName, Site, true);
                DeliveryLoadingPage loadingPage = deliveryCreateModalPage.Create();
                loadingPage.AddService(serviceName);
                loadingPage.AddQuantity(qty);
                homePage.Navigate();
                var dispatchPage = homePage.GoToProduction_DispatchPage();
                dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryName);
                dispatchPage.Filter(DispatchPage.FilterType.Site, Site);
                var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
                //Show all
                dispatchPage.Filter(DispatchPage.FilterType.ShowAll, true);
                Assert.AreEqual(1, previsionnalQty.CheckTotalNumber(), MessageErreur.FILTRE_ERRONE, "Show all");

                //Not Validated Only
                dispatchPage.Filter(DispatchPage.FilterType.NotValidateOnly, true);
                Assert.AreEqual(1, previsionnalQty.CheckTotalNumber(), MessageErreur.FILTRE_ERRONE, "Not Validated only");

                //Validated Only
                previsionnalQty.ValidateFirstDispatch();
                dispatchPage.Filter(DispatchPage.FilterType.ValidateOnly, true);
                Assert.AreEqual(1, previsionnalQty.CheckTotalNumber(), MessageErreur.FILTRE_ERRONE, "Validated only");

            }
            finally
            {
                deliveryPage = homePage.GoToCustomers_DeliveryPage();
                DeliveryMassiveDeletePage deliveryMassiveDeletePage = deliveryPage.MassiveDelete();
                deliveryMassiveDeletePage.SearchDeliveryName(deliveryName);
                deliveryMassiveDeletePage.ClickSearchButton();
                deliveryMassiveDeletePage.SelectAll();
                deliveryMassiveDeletePage.ClickDeleteButton();
                deliveryMassiveDeletePage.ClickConfirmDeleteButton();
            }

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_DISP_Filter_ShowCustomerTypes()
        {
            Random rnd = new Random();

            // Prepare delivery 
            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();
            string deliverySite = TestContext.Properties["SiteLP"].ToString();
            string deliveryCustomer = TestContext.Properties["CustomerDelivery"].ToString();

            // Prepare service           
            string serviceName = serviceNameToday + " - " + rnd.Next().ToString();
            string serviceName1 = serviceNameToday + " - " + rnd.Next().ToString();
            string serviceCode = rnd.Next().ToString();
            string serviceProduction = GenerateName(4);

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Create services
            var servicePage = homePage.GoToCustomers_ServicePage();
            var ServiceCreateModalPage = servicePage.ServiceCreatePage();
            ServiceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            var serviceGeneralInformationsPage = ServiceCreateModalPage.Create();

            var pricePage = serviceGeneralInformationsPage.GoToPricePage();
            var priceModalPage = pricePage.AddNewCustomerPrice();
            priceModalPage.FillFields_CustomerPrice(deliverySite, customerName, DateUtils.Now, DateUtils.Now.AddDays(10));
            var customer = priceModalPage.GetCustomerName();
            serviceGeneralInformationsPage.BackToList();

            servicePage.ServiceCreatePage();
            ServiceCreateModalPage.FillFields_CreateServiceModalPage(serviceName1, serviceCode, serviceProduction);
            ServiceCreateModalPage.Create();

            serviceGeneralInformationsPage.GoToPricePage();
            pricePage.AddNewCustomerPrice();
            priceModalPage.FillFields_CustomerPrice(deliverySite, deliveryCustomer, DateUtils.Now, DateUtils.Now.AddDays(10));
            serviceGeneralInformationsPage.BackToList();

            CreateNewDispatch(homePage, deliveryName, deliveryCustomer, serviceName1);
            CreateNewDispatch(homePage, deliveryName, customer, serviceName);

            //Act
            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryName);
            dispatchPage.Filter(DispatchPage.FilterType.Site, deliverySite);
            var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();

            dispatchPage.Filter(DispatchPage.FilterType.CustomersTypes, "Colectividades");
            Assert.AreEqual(1, previsionnalQty.CheckTotalNumber(), MessageErreur.FILTRE_ERRONE, "Customer types");

            dispatchPage.Filter(DispatchPage.FilterType.CustomersTypes, "Financiero");
            Assert.AreEqual(1, previsionnalQty.CheckTotalNumber(), MessageErreur.FILTRE_ERRONE, "Customer types");

            dispatchPage.Filter(DispatchPage.FilterType.CustomersTypes, "None");
            Assert.AreEqual(2, previsionnalQty.CheckTotalNumber(), MessageErreur.FILTRE_ERRONE, "Customer types");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_DISP_Filter_Customer()
        {
            Random rnd = new Random();
            string site = "ACE";
            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();

            var homePage = LogInAsAdmin();

            var customerPage = homePage.GoToCustomers_CustomerPage();
            customerPage.ResetFilters();
            customerPage.Filter(CustomerPage.FilterType.Search, customerName);
            string customerCode = customerPage.GetFirstCustomerIcao();

            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.ResetFilter();
            dispatchPage.Filter(DispatchPage.FilterType.Site, site);
            dispatchPage.Filter(DispatchPage.FilterType.Customers, customerName);
            var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            if (!dispatchPage.isPageSizeEqualsTo100())
            {
                dispatchPage.PageSize("8");
                dispatchPage.PageSize("100");
            }

            Assert.IsTrue(previsionnalQty.VerifyCustomer(customerCode), MessageErreur.FILTRE_ERRONE, "Customers");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_DISP_Filter_ServiceCategorie()
        {
            Random rnd = new Random();

            // Prepare delivery 
            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();
            string serviceCategorie = TestContext.Properties["Production_Service1"].ToString();
            string Site = TestContext.Properties["SiteACE"].ToString();
            string qty = "100";

            // Arrange
            var homePage = LogInAsAdmin();
            DeliveryPage deliveryPage = homePage.GoToCustomers_DeliveryPage();
            try
            {
                DeliveryCreateModalPage deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, customerName, Site, true);
                DeliveryLoadingPage loadingPage = deliveryCreateModalPage.Create();
                loadingPage.AddService(serviceName);
                loadingPage.AddQuantity(qty);
                homePage.Navigate();
                DispatchPage dispatchPage = homePage.GoToProduction_DispatchPage();
                dispatchPage.Filter(DispatchPage.FilterType.Site, Site);
                dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryName);
                dispatchPage.Filter(DispatchPage.FilterType.ServiceCategories, serviceCategorie);
                var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
                Assert.AreEqual(1, previsionnalQty.CheckTotalNumber(), MessageErreur.FILTRE_ERRONE, "Services categories");
                ServicePricePage servicePricePage = dispatchPage.ShowService();
                ServiceGeneralInformationPage serviceGeneralInformationPage = servicePricePage.ClickOnGeneralInformationTab();
                var Categorie = serviceGeneralInformationPage.GetCategory();
                Assert.AreEqual(Categorie, serviceCategorie, MessageErreur.FILTRE_ERRONE, "Services categories");

            }
            finally
            {
                deliveryPage = homePage.GoToCustomers_DeliveryPage();
                DeliveryMassiveDeletePage deliveryMassiveDeletePage = deliveryPage.MassiveDelete();
                deliveryMassiveDeletePage.SearchDeliveryName(deliveryName);
                deliveryMassiveDeletePage.ClickSearchButton();
                deliveryMassiveDeletePage.SelectAll();
                deliveryMassiveDeletePage.ClickDeleteButton();
                deliveryMassiveDeletePage.ClickConfirmDeleteButton();
            }

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_DISP_Filter_Deliveries()
        {
            Random rnd = new Random();
            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();
            string Site = TestContext.Properties["SiteACE"].ToString();
            string qty = "100";

            var homePage = LogInAsAdmin();
            DeliveryPage deliveryPage = homePage.GoToCustomers_DeliveryPage();
            try
            {
                DeliveryCreateModalPage deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, customerName, Site, true);
                DeliveryLoadingPage loadingPage = deliveryCreateModalPage.Create();
                loadingPage.AddService(serviceName);
                loadingPage.AddQuantity(qty);
                homePage.Navigate();
                var dispatchPage = homePage.GoToProduction_DispatchPage();
                dispatchPage.Filter(DispatchPage.FilterType.Site, Site);
                dispatchPage.Filter(DispatchPage.FilterType.Deliveries, deliveryName);
                PrevisionalQtyPage previsionalQtyPage = dispatchPage.ClickPrevisonalQuantity();
                var NameDelivery = previsionalQtyPage.GetFirstDeliveryName();
                Assert.AreEqual(1, dispatchPage.CheckTotalNumber(), MessageErreur.FILTRE_ERRONE, "Deliveries");
                Assert.AreEqual(NameDelivery, deliveryName, MessageErreur.FILTRE_ERRONE, "Deliveries");
            }
            finally
            {
                deliveryPage = homePage.GoToCustomers_DeliveryPage();
                DeliveryMassiveDeletePage deliveryMassiveDeletePage = deliveryPage.MassiveDelete();
                deliveryMassiveDeletePage.SearchDeliveryName(deliveryName);
                deliveryMassiveDeletePage.ClickSearchButton();
                deliveryMassiveDeletePage.SelectAll();
                deliveryMassiveDeletePage.ClickDeleteButton();
                deliveryMassiveDeletePage.ClickConfirmDeleteButton();
            }


        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_DISP_Filter_Delivery_Rounds()
        {
            Random rnd = new Random();
            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();
            string Site = TestContext.Properties["SiteACE"].ToString();
            string qty = "100";
            string deliveryRoundName = "DR-" + rnd.Next().ToString();
            DateTime startDate = DateTime.Now;
            DateTime endDate = DateTime.Now.AddDays(10);
            var homePage = LogInAsAdmin();

            DeliveryPage deliveryPage = homePage.GoToCustomers_DeliveryPage();
            try
            {
                DeliveryCreateModalPage deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, customerName, Site, true);
                DeliveryLoadingPage loadingPage = deliveryCreateModalPage.Create();
                loadingPage.AddService(serviceName);
                loadingPage.AddQuantity(qty);

                // Create delivery round
                homePage.Navigate();
                var deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
                var deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
                var deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRound(deliveryRoundName, Site, startDate, endDate);
                var deliveryRoundDeliveriesPage = deliveryRoundGeneralInfoPage.ClickOnDeliveryTab();
                deliveryRoundDeliveriesPage.AddDelivery(deliveryName);

                homePage.Navigate();
                DispatchPage dispatchPage = homePage.GoToProduction_DispatchPage();
                dispatchPage.Filter(DispatchPage.FilterType.Site, Site);
                dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryName);
                var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
                dispatchPage.Filter(DispatchPage.FilterType.DeliveryRounds, deliveryRoundName);
                var totalNumber = dispatchPage.CheckTotalNumber();
                //Assert
                Assert.AreEqual(1, totalNumber, MessageErreur.FILTRE_ERRONE, "Delivery rounds");
            }
            finally
            {
                deliveryPage = homePage.GoToCustomers_DeliveryPage();
                DeliveryMassiveDeletePage deliveryMassiveDeletePage = deliveryPage.MassiveDelete();
                deliveryMassiveDeletePage.SearchDeliveryName(deliveryName);
                deliveryMassiveDeletePage.ClickSearchButton();
                deliveryMassiveDeletePage.SelectAll();
                deliveryMassiveDeletePage.ClickDeleteButton();
                deliveryMassiveDeletePage.ClickConfirmDeleteButton();
                var deliveryRoundsPage = homePage.GoToCustomers_DeliveryRoundPage();
                deliveryRoundsPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRoundName);
                deliveryRoundsPage.DeleteFirstDeliveryRound();
            }



        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_DISP_ResetFilter()
        {
            // Prepare
            Random rnd = new Random();
            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            CreateNewDispatch(homePage, deliveryName, customerName, serviceName);

            //Act
            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.ResetFilter();
            var initNumber = dispatchPage.CheckTotalNumber();

            dispatchPage.Filter(DispatchPage.FilterType.Deliveries, deliveryName);
            var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();

            Assert.AreEqual(1, previsionnalQty.CheckTotalNumber(), MessageErreur.FILTRE_ERRONE, "Deliveries");

            dispatchPage.ResetFilter();

            Assert.AreNotEqual(previsionnalQty.CheckTotalNumber(), 1, "Le resetFilter ne fonctionne pas correctement.");
            Assert.AreEqual(initNumber, previsionnalQty.CheckTotalNumber(), MessageErreur.FILTRE_ERRONE, "Reset filter");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_DISP_Date()
        {
            // Prepare
            Random rnd = new Random();
            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();
            string autreQty = "20";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            CreateNewDispatch(homePage, deliveryName, customerName, serviceName);

            //Act
            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryName);
            dispatchPage.Filter(DispatchPage.FilterType.Site, "ACE");
            var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();

            previsionnalQty.SetQuantity(autreQty);
            dispatchPage.ClickNextWeek();
            Assert.AreNotEqual(previsionnalQty.GetQuantity(), autreQty, "Les quantités du dispatch sont les mêmes la semaine suivante alors qu'elles ont été modifiées.");

            dispatchPage.ClickPreviousWeek();
            Assert.AreEqual(previsionnalQty.GetQuantity(), autreQty, "les quantités du dispatch de la semaine ne sont pas les données modifiées.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_DISP_NotMenuLinked()
        {
            // Prepare
            Random random = new Random();
            string deliveryName = $"{deliveryNameToday}-{random.Next()}";
            string deliverySite = TestContext.Properties["SiteLP"].ToString();
            string serviceName = $"{serviceNameToday}-{random.Next()}";
            string serviceCode = random.Next().ToString();
            string serviceProduction = GenerateName(4);
            // Arrange
            HomePage homePage = LogInAsAdmin();
            // Create service
            ServicePage servicePage = homePage.GoToCustomers_ServicePage();
            ServiceCreateModalPage ServiceCreateModalPage = servicePage.ServiceCreatePage();
            ServiceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            ServiceGeneralInformationPage serviceGeneralInformationsPage = ServiceCreateModalPage.Create();
            ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
            ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
            priceModalPage.FillFields_CustomerPrice(deliverySite, customerName, DateUtils.Now, DateUtils.Now.AddDays(10));
            string customer = priceModalPage.GetCustomerName();
            serviceGeneralInformationsPage.BackToList();
            CreateNewDispatch(homePage, deliveryName, customer, serviceName, null, false);
            //Act
            DispatchPage dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryName);
            PrevisionalQtyPage previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            //Assert
            bool isNotLinked = previsionnalQty.IsNotLinked();
            Assert.IsTrue(isNotLinked, "l'icone d'alerte n'est pas visible quand un dispatch n'a pas de menu associé.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_DISP_PrevisonnalQty_Update()
        {
            // Prepare
            Random rnd = new Random();
            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();
            string newQty = "100";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            CreateNewDispatch(homePage, deliveryName, customerName, serviceName);

            //Act
            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryName);
            dispatchPage.ClickNextWeek();
            var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            previsionnalQty.UpdateQuantities(newQty);

            dispatchPage.ResetFilter();
            dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryName);
            dispatchPage.ClickNextWeek();
            bool updated = previsionnalQty.IsQuantitiesUpdated(newQty);

            //Assert
            Assert.IsTrue(updated, "Les quantités n'ont pas été mises à jour.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_DISP_PrevisonnalQty_Validate()
        {
            // Prepare
            Random rnd = new Random();
            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();
            string deliverySite = TestContext.Properties["SiteLP"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            CreateNewDispatch(homePage, deliveryName, customerName, serviceName);

            //Act
            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryName);
            dispatchPage.Filter(DispatchPage.FilterType.Site, deliverySite);

            var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            previsionnalQty.ValidateFirstDispatch();

            //Assert
            Assert.IsTrue(previsionnalQty.IsDispatchValidated(), "Le dispatch n'a pas été validé.");
            Assert.IsFalse(previsionnalQty.CanUpdateQty(), "Malgré la validation, les quantités du dispatch sont toujours modifiables.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_DISP_PrevisonnalQty_Day_Validate()
        {
            // Prepare
            string deliveryName = $"{deliveryNameToday}-{new Random().Next()}";
            // Arrange
            HomePage homePage = LogInAsAdmin();
            CreateNewDispatch(homePage, deliveryName, customerName, serviceName);
            //Act
            DispatchPage dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryName);
            dispatchPage.ValidateSunday();
            PrevisionalQtyPage previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            //Assert
            bool isValidated = previsionnalQty.IsSundayValidatedByColorDay();
            Assert.IsTrue(isValidated, "Le dispatch n'a pas été validé pour le jour choisi.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_DISP_PrevisonnalQty_ValidateAll()
        {
            // Prepare
            Random rnd = new Random();
            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();
            string Site = TestContext.Properties["SiteACE"].ToString();
            string qty = "100";

            // Arrange
            var homePage = LogInAsAdmin();
            DeliveryPage deliveryPage = homePage.GoToCustomers_DeliveryPage();
            try
            {
                DeliveryCreateModalPage deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, customerName, Site, true);
                DeliveryLoadingPage loadingPage = deliveryCreateModalPage.Create();
                loadingPage.AddService(serviceName);
                loadingPage.AddQuantity(qty);
                homePage.Navigate();
                var dispatchPage = homePage.GoToProduction_DispatchPage();
                dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryName);
                var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
                dispatchPage.ValidateAll();
                //Assert
                var IsValidatedByColorDay = previsionnalQty.IsValidatedByColorDay();
                Assert.IsTrue(IsValidatedByColorDay, "Le dispatch n'a pas été validé pour tous les jours de la semaine.");
            }
            finally
            {
                deliveryPage = homePage.GoToCustomers_DeliveryPage();
                DeliveryMassiveDeletePage deliveryMassiveDeletePage = deliveryPage.MassiveDelete();
                deliveryMassiveDeletePage.SearchDeliveryName(deliveryName);
                deliveryMassiveDeletePage.ClickSearchButton();
                deliveryMassiveDeletePage.SelectAll();
                deliveryMassiveDeletePage.ClickDeleteButton();
                deliveryMassiveDeletePage.ClickConfirmDeleteButton();
            }


        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_DISP_QuantityToProduce_ValidationAndMenu()
        {
            // Prepare
            Random rnd = new Random();
            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            CreateNewDispatch(homePage, deliveryName, customerName, serviceName);

            // Display dispatch
            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryName);
            var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            Assert.IsTrue(previsionnalQty.IsValidatedByColorDay(), "Le dispatch n'a pas été validé pour tous les jours de la semaine.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_DISP_QuantityToProduce_Update()
        {
            // Prepare
            Random rnd = new Random();
            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();
            string newQty = "100";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            CreateNewDispatch(homePage, deliveryName, customerName, serviceName);

            //Act
            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryName);
            var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.ClickNextWeek();
            previsionnalQty.ValidateFirstDispatch();

            var qtyToProduce = dispatchPage.ClickQuantityToProduce();
            qtyToProduce.UpdateQuantities(newQty);

            dispatchPage.ResetFilter();
            dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryName);
            qtyToProduce = dispatchPage.ClickQuantityToProduce();
            dispatchPage.ClickNextWeek();
            bool updated = qtyToProduce.IsQuantitiesUpdated(newQty);

            //Assert
            Assert.IsTrue(updated, "Les quantités to produce n'ont pas été mises à jour.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_DISP_QuantityToProduce_Validate()
        {
            // Prepare
            Random rnd = new Random();
            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();
            string Site = TestContext.Properties["SiteACE"].ToString();
            string qty = "100";

            var homePage = LogInAsAdmin();
            DeliveryPage deliveryPage = homePage.GoToCustomers_DeliveryPage();
            try
            {
                DeliveryCreateModalPage deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, customerName, Site, true);
                DeliveryLoadingPage loadingPage = deliveryCreateModalPage.Create();
                loadingPage.AddService(serviceName);
                loadingPage.AddQuantity(qty);
                homePage.Navigate();
                //Act
                var dispatchPage = homePage.GoToProduction_DispatchPage();
                dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryName);
                var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
                previsionnalQty.ValidateFirstDispatch();
                var qtyToProduce = dispatchPage.ClickQuantityToProduce();
                qtyToProduce.ValidateFirstDispatch();
                //Assert
                Assert.IsTrue(qtyToProduce.IsDispatchValidated(), "Le dispatch n'a pas été validé dans les qty to produce.");
                Assert.IsFalse(qtyToProduce.CanUpdateQty(), "Malgré la validation, les quantités du dispatch sont toujours modifiables dans le qty to produce.");
            }
            finally
            {
                deliveryPage = homePage.GoToCustomers_DeliveryPage();
                DeliveryMassiveDeletePage deliveryMassiveDeletePage = deliveryPage.MassiveDelete();
                deliveryMassiveDeletePage.SearchDeliveryName(deliveryName);
                deliveryMassiveDeletePage.ClickSearchButton();
                deliveryMassiveDeletePage.SelectAll();
                deliveryMassiveDeletePage.ClickDeleteButton();
                deliveryMassiveDeletePage.ClickConfirmDeleteButton();
            }


        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_DISP_QuantityToProduce_Day_Validate()
        {
            // Prepare
            string deliveryName = $"{deliveryNameToday}-{new Random().Next()}";
            // Arrange
            HomePage homePage = LogInAsAdmin();
            CreateNewDispatch(homePage, deliveryName, customerName, serviceName);
            //Act
            DispatchPage dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.ResetFilter();
            dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryName);
            PrevisionalQtyPage previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.ValidateSunday();
            bool isValidated = previsionnalQty.IsSundayValidatedByColorDay();
            Assert.IsTrue(isValidated, "Le dispatch n'a pas été validé pour le jour choisi pour les previsional qty.");
            QuantityToProducePage qtyToProduce = dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateSunday();
            isValidated = qtyToProduce.IsSundayValidatedByColorDay();
            Assert.IsTrue(isValidated, "Le dispatch n'a pas été validé pour le jour choisi pour les qty to produce.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_DISP_QuantityToProduce_ValidateAll()
        {
            // Prepare
            Random rnd = new Random();
            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();
            string Site = TestContext.Properties["SiteACE"].ToString();
            string qty = "100";

            // Arrange
            var homePage = LogInAsAdmin();
            DeliveryPage deliveryPage = homePage.GoToCustomers_DeliveryPage();
            try
            {
                DeliveryCreateModalPage deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, customerName, Site, true);
                DeliveryLoadingPage loadingPage = deliveryCreateModalPage.Create();
                loadingPage.AddService(serviceName);
                loadingPage.AddQuantity(qty);
                homePage.Navigate();
                var dispatchPage = homePage.GoToProduction_DispatchPage();
                dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryName);
                var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
                dispatchPage.ValidateAll();
                Assert.IsTrue(previsionnalQty.IsValidatedByColorDay(), "Le dispatch n'a pas été validé pour tous les jours de la semaine en previsional qty.");
                var qtyToProduce = dispatchPage.ClickQuantityToProduce();
                dispatchPage.ValidateAll();
                //Assert
                Assert.IsTrue(qtyToProduce.IsValidatedByColorDay(), "Le dispatch n'a pas été validé pour tous les jours de la semaine en qty to produce.");
            }
            finally
            {
                deliveryPage = homePage.GoToCustomers_DeliveryPage();
                DeliveryMassiveDeletePage deliveryMassiveDeletePage = deliveryPage.MassiveDelete();
                deliveryMassiveDeletePage.SearchDeliveryName(deliveryName);
                deliveryMassiveDeletePage.ClickSearchButton();
                deliveryMassiveDeletePage.SelectAll();
                deliveryMassiveDeletePage.ClickDeleteButton();
                deliveryMassiveDeletePage.ClickConfirmDeleteButton();
            }

            //Act

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_DISP_QuantityToProduce_ValidateAll_Error()
        {
            // Prepare
            Random rnd = new Random();
            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();
            string Site = TestContext.Properties["SiteACE"].ToString();
            string qty = "100";

            var homePage = LogInAsAdmin();
            DeliveryPage deliveryPage = homePage.GoToCustomers_DeliveryPage();
            try
            {
                DeliveryCreateModalPage deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, customerName, Site, true);
                DeliveryLoadingPage loadingPage = deliveryCreateModalPage.Create();
                loadingPage.AddService(serviceName);
                loadingPage.AddQuantity(qty);
                homePage.Navigate();
                var dispatchPage = homePage.GoToProduction_DispatchPage();
                dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryName);
                var qtyToProduce = dispatchPage.ClickQuantityToProduce();
                dispatchPage.ValidateAll();
                //Assert
                Assert.IsTrue(qtyToProduce.IsErrorValidation(), "Le dispatch a été validé en qty to produce malgré le fait que toutes les qté n'ont pas été validées en previsional qty.");
            }
            finally
            {
                deliveryPage = homePage.GoToCustomers_DeliveryPage();
                DeliveryMassiveDeletePage deliveryMassiveDeletePage = deliveryPage.MassiveDelete();
                deliveryMassiveDeletePage.SearchDeliveryName(deliveryName);
                deliveryMassiveDeletePage.ClickSearchButton();
                deliveryMassiveDeletePage.SelectAll();
                deliveryMassiveDeletePage.ClickDeleteButton();
                deliveryMassiveDeletePage.ClickConfirmDeleteButton();
            }


        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_DISP_QuantityToInvoice_ValidationAndMenu()
        {
            // Prepare
            Random rnd = new Random();
            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            CreateNewDispatch(homePage, deliveryName, customerName, serviceName);

            // Display dispatch
            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryName);
            dispatchPage.ValidateAll();

            var qtyToProduce = dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();

            //Assert
            Assert.IsTrue(qtyToProduce.IsValidatedByColorDay(), "Le dispatch n'a pas été validé pour tous les jours de la semaine en qty to produce.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_DISP_QuantityToInvoice_Update()
        {
            // Prepare
            Random rnd = new Random();
            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();
            string newQty = "100";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            CreateNewDispatch(homePage, deliveryName, customerName, serviceName);

            //Act
            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryName);
            var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            dispatchPage.ClickNextWeek();
            previsionnalQty.ValidateFirstDispatch();

            var qtyToProduce = dispatchPage.ClickQuantityToProduce();
            qtyToProduce.ValidateFirstDispatch();

            var qtytoInvoice = dispatchPage.ClickQuantityToInvoice();
            qtytoInvoice.UpdateQuantities(newQty);

            dispatchPage.ResetFilter();
            dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryName);
            qtytoInvoice = dispatchPage.ClickQuantityToInvoice();
            dispatchPage.ClickNextWeek();
            bool updated = qtytoInvoice.IsQuantitiesUpdated(newQty);

            //Assert
            Assert.IsTrue(updated, "Les quantités to produce n'ont pas été mises à jour.");
        }


        [Timeout(_timeout)]
        [TestMethod]
        public void PR_DISP_QuantityToInvoice_Validate()
        {
            // Prepare
            Random rnd = new Random();
            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            CreateNewDispatch(homePage, deliveryName, customerName, serviceName);

            //Act
            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryName);

            var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
            previsionnalQty.ValidateFirstDispatch();

            var qtyToProduce = dispatchPage.ClickQuantityToProduce();
            qtyToProduce.ValidateFirstDispatch();

            var qtytoInvoice = dispatchPage.ClickQuantityToInvoice();
            qtytoInvoice.ValidateTheFirst();

            //Assert
            Assert.IsTrue(qtytoInvoice.IsDispatchValidated(), "Le dispatch n'a pas été validé dans les qty to invoice.");
            Assert.IsFalse(qtytoInvoice.CanUpdateQty(), "Malgré la validation, les quantités du dispatch sont toujours modifiables dans le qty to invoice.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_DISP_QuantityToInvoice_Day_Validate()
        {
            // Prepare
            Random rnd = new Random();
            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();
            string Site = TestContext.Properties["SiteACE"].ToString();
            string qty = "100";

            // Arrange
            var homePage = LogInAsAdmin();
            DeliveryPage deliveryPage = homePage.GoToCustomers_DeliveryPage();
            try
            {
                DeliveryCreateModalPage deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, customerName, Site, true);
                DeliveryLoadingPage loadingPage = deliveryCreateModalPage.Create();
                loadingPage.AddService(serviceName);
                loadingPage.AddQuantity(qty);
                homePage.Navigate();
                var dispatchPage = homePage.GoToProduction_DispatchPage();
                dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryName);
                var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
                dispatchPage.ValidateSunday();
                var IsSundayValidatedByColorDay = previsionnalQty.IsSundayValidatedByColorDay();
                Assert.IsTrue(IsSundayValidatedByColorDay, "Le dispatch n'a pas été validé pour le jour choisi pour les previsional qty.");
                var qtyToProduce = dispatchPage.ClickQuantityToProduce();
                dispatchPage.ValidateSunday();
                IsSundayValidatedByColorDay = qtyToProduce.IsSundayValidatedByColorDay();
                Assert.IsTrue(IsSundayValidatedByColorDay, "Le dispatch n'a pas été validé pour le jour choisi pour les qty to produce.");
                var qtytoInvoice = dispatchPage.ClickQuantityToInvoice();
                dispatchPage.ValidateSunday();
                //Assert
                IsSundayValidatedByColorDay = qtytoInvoice.IsSundayValidatedByColorDay();
                Assert.IsTrue(IsSundayValidatedByColorDay, "Le dispatch n'a pas été validé pour le jour choisi pour les qty to invoice.");
            }
            finally
            {
                deliveryPage = homePage.GoToCustomers_DeliveryPage();
                DeliveryMassiveDeletePage deliveryMassiveDeletePage = deliveryPage.MassiveDelete();
                deliveryMassiveDeletePage.SearchDeliveryName(deliveryName);
                deliveryMassiveDeletePage.ClickSearchButton();
                deliveryMassiveDeletePage.SelectAll();
                deliveryMassiveDeletePage.ClickDeleteButton();
                deliveryMassiveDeletePage.ClickConfirmDeleteButton();
            }


        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_DISP_QuantityToInvoice_ValidateAll()
        {
            // Prepare
            Random rnd = new Random();
            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();
            string Site = TestContext.Properties["SiteACE"].ToString();
            string qty = "100";

            // Arrange
            var homePage = LogInAsAdmin();
            DeliveryPage deliveryPage = homePage.GoToCustomers_DeliveryPage();
            try
            {
                DeliveryCreateModalPage deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, customerName, Site, true);
                DeliveryLoadingPage loadingPage = deliveryCreateModalPage.Create();
                loadingPage.AddService(serviceName);
                loadingPage.AddQuantity(qty);
                homePage.Navigate();
                var dispatchPage = homePage.GoToProduction_DispatchPage();
                dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryName);
                var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
                dispatchPage.ValidateAll();
                Assert.IsTrue(previsionnalQty.IsValidatedByColorDay(), "Le dispatch n'a pas été validé pour tous les jours de la semaine en previsional qty.");
                var qtyToProduce = dispatchPage.ClickQuantityToProduce();
                dispatchPage.ValidateAll();
                Assert.IsTrue(qtyToProduce.IsValidatedByColorDay(), "Le dispatch n'a pas été validé pour tous les jours de la semaine en qty to produce.");
                var qtyToInvoice = dispatchPage.ClickQuantityToInvoice();
                dispatchPage.ValidateAll();
                //Assert
                Assert.IsTrue(qtyToInvoice.IsValidatedByColorDay(), "Le dispatch n'a pas été validé pour tous les jours de la semaine en qty to invoice.");
            }
            finally
            {
                deliveryPage = homePage.GoToCustomers_DeliveryPage();
                DeliveryMassiveDeletePage deliveryMassiveDeletePage = deliveryPage.MassiveDelete();
                deliveryMassiveDeletePage.SearchDeliveryName(deliveryName);
                deliveryMassiveDeletePage.ClickSearchButton();
                deliveryMassiveDeletePage.SelectAll();
                deliveryMassiveDeletePage.ClickDeleteButton();
                deliveryMassiveDeletePage.ClickConfirmDeleteButton();
            }


        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_DISP_QuantityToInvoice_ValidateAll_Error()
        {
            // Prepare
            Random rnd = new Random();
            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();
            string Site = TestContext.Properties["SiteACE"].ToString();
            string qty = "100";

            var homePage = LogInAsAdmin();
            DeliveryPage deliveryPage = homePage.GoToCustomers_DeliveryPage();
            try
            {
                DeliveryCreateModalPage deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, customerName, Site, true);
                DeliveryLoadingPage loadingPage = deliveryCreateModalPage.Create();
                loadingPage.AddService(serviceName);
                loadingPage.AddQuantity(qty);
                homePage.Navigate();
                var dispatchPage = homePage.GoToProduction_DispatchPage();
                dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryName);
                var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
                dispatchPage.ValidateAll();
                Assert.IsTrue(previsionnalQty.IsValidatedByColorDay(), "Le dispatch n'a pas été validé pour tous les jours de la semaine en previsional qty.");
                var qtyToProduce = dispatchPage.ClickQuantityToProduce();
                dispatchPage.ValidateSunday();
                Assert.IsTrue(qtyToProduce.IsSundayValidatedByColorDay(), "Le dispatch n'a pas été validé pour le jour choisi pour les qty to produce.");
                var qtyToInvoice = dispatchPage.ClickQuantityToInvoice();
                dispatchPage.ValidateAll();
                //Assert
                Assert.IsTrue(qtyToInvoice.IsErrorValidation(), "Le dispatch a été validé en qty to invoice malgré le fait que toutes les qté n'ont pas été validées avant.");
            }
            finally
            {
                deliveryPage = homePage.GoToCustomers_DeliveryPage();
                DeliveryMassiveDeletePage deliveryMassiveDeletePage = deliveryPage.MassiveDelete();
                deliveryMassiveDeletePage.SearchDeliveryName(deliveryName);
                deliveryMassiveDeletePage.ClickSearchButton();
                deliveryMassiveDeletePage.SelectAll();
                deliveryMassiveDeletePage.ClickDeleteButton();
                deliveryMassiveDeletePage.ClickConfirmDeleteButton();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_DISP_Generate_SupplierOrder()
        {
            // Prepare
            Random rnd = new Random();
            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();
            string deliveryPlace = TestContext.Properties["Place"].ToString();
            string deliveryRoundName = "DR-" + rnd.Next().ToString();
            // Arrange
            HomePage homePage = LogInAsAdmin();

            CreateNewDispatch(homePage, deliveryName, customerName, serviceName, deliveryRoundName);

            //Act
            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryName);
            dispatchPage.Filter(DispatchPage.FilterType.Deliveries, deliveryName);
            dispatchPage.Filter(DispatchPage.FilterType.DeliveryRounds, deliveryRoundName);
            dispatchPage.ValidateAll();


            var supplierOrderModal = dispatchPage.ClickGenerateSupplierOrder();
            var supplerId = supplierOrderModal.GenerateSupplyOrder(deliveryPlace);

            var supplyOrderItem = supplierOrderModal.Generate();
            var supplyOrderPage = supplyOrderItem.BackToList();
            supplyOrderPage.ResetFilter();
            var result = supplyOrderPage.GetFirstSONumber();
            supplyOrderPage.Close();

            //Assert
            Assert.AreEqual(result, supplerId, "Le supply order associé au dispatch n'a pas été créé.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_DISP_Previsionnal_Qty_UnvalidateAll()
        {
            // Prepare
            Random rnd = new Random();
            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();
            string Site = TestContext.Properties["SiteACE"].ToString();
            string qty = "100";

            // Arrange

            var homePage = LogInAsAdmin();
            DeliveryPage deliveryPage = homePage.GoToCustomers_DeliveryPage();
            try
            {
                DeliveryCreateModalPage deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, customerName, Site, true);
                DeliveryLoadingPage loadingPage = deliveryCreateModalPage.Create();
                loadingPage.AddService(serviceName);
                loadingPage.AddQuantity(qty);
                homePage.Navigate();
                var dispatchPage = homePage.GoToProduction_DispatchPage();
                dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryName);
                dispatchPage.Filter(DispatchPage.FilterType.Site, Site);
                var previsionnalQty = dispatchPage.ClickPrevisonalQuantity();
                dispatchPage.ValidateAll();
                Assert.IsTrue(previsionnalQty.IsDispatchValidated(), "Le dispatch n'a pas été validé.");
                Assert.IsFalse(previsionnalQty.CanUpdateQty(), "Malgré la validation, les quantités du dispatch sont toujours modifiables.");
                dispatchPage.UnValidateAll();
                Assert.IsFalse(previsionnalQty.IsDispatchValidated(), "Le dispatch est toujours validé.");
                Assert.IsTrue(previsionnalQty.CanUpdateQty(), "Malgré la dévalidation, les quantités du dispatch ne sont pas modifiables pour le prévisionnal qty.");
            }
            finally
            {
                deliveryPage = homePage.GoToCustomers_DeliveryPage();
                DeliveryMassiveDeletePage deliveryMassiveDeletePage = deliveryPage.MassiveDelete();
                deliveryMassiveDeletePage.SearchDeliveryName(deliveryName);
                deliveryMassiveDeletePage.ClickSearchButton();
                deliveryMassiveDeletePage.SelectAll();
                deliveryMassiveDeletePage.ClickDeleteButton();
                deliveryMassiveDeletePage.ClickConfirmDeleteButton();
            }


        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_DISP_Qty_to_Produce_UnvalidateAll()
        {
            // Prepare
            Random rnd = new Random();
            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();
            string Site = TestContext.Properties["SiteACE"].ToString();
            string qty = "100";

            // Arrange
            var homePage = LogInAsAdmin();
            DeliveryPage deliveryPage = homePage.GoToCustomers_DeliveryPage();
            try
            {
                DeliveryCreateModalPage deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, customerName, Site, true);
                DeliveryLoadingPage loadingPage = deliveryCreateModalPage.Create();
                loadingPage.AddService(serviceName);
                loadingPage.AddQuantity(qty);
                homePage.Navigate();
                var dispatchPage = homePage.GoToProduction_DispatchPage();
                dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryName);
                dispatchPage.ValidateAll();
                var qtyToProduce = dispatchPage.ClickQuantityToProduce();
                dispatchPage.ValidateAll();
                var IsDispatchValidated = qtyToProduce.IsDispatchValidated();
                var CanUpdateQty = qtyToProduce.CanUpdateQty();
                Assert.IsTrue(IsDispatchValidated, "Le dispatch n'a pas été validé.");
                Assert.IsFalse(CanUpdateQty, "Malgré la validation, les quantités du dispatch sont toujours modifiables.");
                dispatchPage.UnValidateAll();
                //Assert
                IsDispatchValidated = qtyToProduce.IsDispatchValidated();
                CanUpdateQty = qtyToProduce.CanUpdateQty();
                Assert.IsFalse(IsDispatchValidated, "Le dispatch est toujours validé.");
                Assert.IsTrue(CanUpdateQty, "Malgré la dévalidation, les quantités du dispatch ne sont pas modifiables pour le qty to produce.");
            }
            finally
            {
                deliveryPage = homePage.GoToCustomers_DeliveryPage();
                DeliveryMassiveDeletePage deliveryMassiveDeletePage = deliveryPage.MassiveDelete();
                deliveryMassiveDeletePage.SearchDeliveryName(deliveryName);
                deliveryMassiveDeletePage.ClickSearchButton();
                deliveryMassiveDeletePage.SelectAll();
                deliveryMassiveDeletePage.ClickDeleteButton();
                deliveryMassiveDeletePage.ClickConfirmDeleteButton();
            }


        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_DISP_Qty_to_Invoice_UnvalidateAll()
        {
            // Prepare
            Random rnd = new Random();
            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();
            string Site = TestContext.Properties["SiteACE"].ToString();
            string qty = "100";

            // Arrange
            var homePage = LogInAsAdmin();
            DeliveryPage deliveryPage = homePage.GoToCustomers_DeliveryPage();
            try
            {
                DeliveryCreateModalPage deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, customerName, Site, true);
                DeliveryLoadingPage loadingPage = deliveryCreateModalPage.Create();
                loadingPage.AddService(serviceName);
                loadingPage.AddQuantity(qty);
                homePage.Navigate();
                var dispatchPage = homePage.GoToProduction_DispatchPage();
                dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryName);
                dispatchPage.ValidateAll();
                dispatchPage.ClickQuantityToProduce();
                dispatchPage.ValidateAll();
                var qtyToInvoice = dispatchPage.ClickQuantityToInvoice();
                dispatchPage.ValidateAll();
                //Assert
                var IsDispatchValidated = qtyToInvoice.IsDispatchValidated();
                var CanUpdateQty = qtyToInvoice.CanUpdateQty();
                Assert.IsTrue(IsDispatchValidated, "Le dispatch n'a pas été validé.");
                Assert.IsFalse(CanUpdateQty, "Malgré la validation, les quantités du dispatch sont toujours modifiables.");
                dispatchPage.UnValidateAll();
                //Assert
                IsDispatchValidated = qtyToInvoice.IsDispatchValidated();
                CanUpdateQty = qtyToInvoice.CanUpdateQty();
                Assert.IsFalse(IsDispatchValidated, "Le dispatch est toujours validé.");
                Assert.IsTrue(CanUpdateQty, "Malgré la dévalidation, les quantités du dispatch ne sont pas modifiables pour le qty to invoice.");
            }
            finally
            {
                deliveryPage = homePage.GoToCustomers_DeliveryPage();
                DeliveryMassiveDeletePage deliveryMassiveDeletePage = deliveryPage.MassiveDelete();
                deliveryMassiveDeletePage.SearchDeliveryName(deliveryName);
                deliveryMassiveDeletePage.ClickSearchButton();
                deliveryMassiveDeletePage.SelectAll();
                deliveryMassiveDeletePage.ClickDeleteButton();
                deliveryMassiveDeletePage.ClickConfirmDeleteButton();
            }


        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_DISP_Export_NewVersion()
        {
            // Prepare
            Random rnd = new Random();
            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            bool newVersionPrint = true;
            string site = "ACE";
            // Arrange
            HomePage homePage = LogInAsAdmin();

            CreateNewDispatch(homePage, deliveryName, customerName, serviceName);

            //Act
            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.ResetFilter();

            dispatchPage.ClearDownloads();

            dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryName);
            dispatchPage.Filter(DispatchPage.FilterType.Site, site);
            dispatchPage.ValidateAll();

            dispatchPage.ExportExcel(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = dispatchPage.GetExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            // Récupération du nom du fichier et construction de l'URL du fichier Excel à ouvrir   
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber("Previsional Quantity", filePath);

            //Assert
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_DISP_PrintTurnover_NewVersion()
        {
            // Prepare
            Random rnd = new Random();
            string deliveryName = deliveryNameToday + "-" + rnd.Next().ToString();
            string deliveryRoundName = "DR" + "-" + rnd.Next(1000, 9000).ToString();

            bool newVersionPrint = true;
            string downloadPath = TestContext.Properties["DownloadsPath"].ToString();
            string docFileNamePdfBegin = "Print Turnover_-_";
            string docFileNameZipBegin = "All_files_";
            // Arrange
            HomePage homePage = LogInAsAdmin();
            CreateNewDispatch(homePage, deliveryName, customerName, serviceName, deliveryRoundName);

            //Act
            var dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.ResetFilter();

            dispatchPage.ClearDownloads();

            dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryName);
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToProduce();
            dispatchPage.ValidateAll();
            dispatchPage.ClickQuantityToInvoice();
            dispatchPage.ValidateAll();

            var reportPage = dispatchPage.PrintTurnover(newVersionPrint);
            var isReportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();
            //Assert
            Assert.IsTrue(isReportGenerated, "Le document PDF n'a pas pu être généré par l'application.");

            // Service (dans un nouveau onglet)
            ServicePricePage pricePage = dispatchPage.ShowService();
            ServiceGeneralInformationPage priceGeneralInfo = pricePage.ClickOnGeneralInformationTab();
            Assert.IsTrue(priceGeneralInfo.IsProduced(), "Service non IsProduced");
            priceGeneralInfo.Close();

            reportPage.Purge(downloadPath, docFileNamePdfBegin, docFileNameZipBegin);
            string trouve = reportPage.PrintAllZip(downloadPath, docFileNamePdfBegin, docFileNameZipBegin);
            FileInfo filePdf = new FileInfo(trouve);
            Assert.IsTrue(filePdf.Exists, "fichier PDF non généré");

            PdfDocument document = PdfDocument.Open(filePdf.FullName);
            List<string> mots = new List<string>();
            foreach (Page p in document.GetPages())
            {
                foreach (var mot in p.GetWords())
                {
                    mots.Add(mot.Text);
                }
            }
            Assert.IsTrue(mots.Count > 10, "Nombre de mots dans le PDF : " + mots.Count);
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_DISP_Go_To_Related_Menus()
        {
            // Prepare
            Random random = new Random();
            string deliveryName = $"{deliveryNameToday}-{random.Next()}";
            string deliverySite = TestContext.Properties["SiteLP"].ToString();
            string serviceName = $"{serviceNameToday}-{random.Next()}";
            string serviceCode = random.Next().ToString();
            string serviceProduction = GenerateName(4);
            string menuName = $"MenuDispatch-{random.Next()}";
            string recipeName = $"RecipeDispatch-{random.Next()}";
            string newMethod = "Std";
            string newCoef = "1";
            // Arrange
            HomePage homePage = LogInAsAdmin();
            // Create services
            ServicePage servicePage = homePage.GoToCustomers_ServicePage();
            ServiceCreateModalPage ServiceCreateModalPage = servicePage.ServiceCreatePage();
            ServiceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
            ServiceGeneralInformationPage serviceGeneralInformationsPage = ServiceCreateModalPage.Create();
            ServicePricePage pricePage = serviceGeneralInformationsPage.GoToPricePage();
            ServiceCreatePriceModalPage priceModalPage = pricePage.AddNewCustomerPrice();
            priceModalPage.FillFields_CustomerPrice(deliverySite, customerName, DateUtils.Now, DateUtils.Now.AddDays(10));
            serviceGeneralInformationsPage.BackToList();
            CreateNewDispatch(homePage, deliveryName, customerName, serviceName, null, true, menuName, recipeName);
            //Act
            DispatchPage dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.Search, deliveryName);
            //Check popup open
            bool isShown = dispatchPage.GoToRelatedMenusPopup();
            Assert.IsTrue(isShown, "La popup 'Related menus ne s'est pas ouverte.");
            //Check menu
            string relatedMenuName = dispatchPage.GetRelatedMenu();
            Assert.AreEqual(menuName, relatedMenuName, "Le related menu n'est pas le bon.");
            //Check date
            dispatchPage.ClickOnRelatedMenuOfToday();
            string relatedMenuDate = dispatchPage.GetRelatedMenuDate();
            bool isValidDay = relatedMenuDate.Contains(DateUtils.Now.DayOfWeek.ToString());
            Assert.IsTrue(isValidDay, "Le jour sélectionné n'est pas le bon.");
            bool isValidDate = relatedMenuDate.Contains(DateUtils.Now.ToString("dd/MM/yyyy"));
            Assert.IsTrue(isValidDate, "Le jour sélectionné n'est pas à la bonne date.");
            // Check recipe
            string relatedMenuRecipeName = dispatchPage.GetRelatedMenuRecipeName();
            Assert.AreEqual(recipeName, relatedMenuRecipeName, "La recette affichée associée au menu n'est pas la bonne");
            bool isValidRelatedMenuRecipeName = dispatchPage.CheckRelatedMenuRecipe().Contains(recipeName.ToUpper());
            Assert.IsTrue(isValidRelatedMenuRecipeName, "La page d'édition de la recette donnée n'est pas visible ou le nom de la recette ne correspond pas.");
            //Change method & coef & Check
            dispatchPage.SetRelatedMenuRecipeMethod(newMethod);
            PageObjects.Menus.Menus.MenusDayViewPage menuDayViewPage = dispatchPage.GoToRelatedMenuLink(); //verif nom menui et coef
            string finalMethod = menuDayViewPage.GetRecipeMethod();
            string finalCoef = menuDayViewPage.GetRecipeCoef();
            Assert.AreEqual(newCoef, finalCoef, "La valeur du coef de la recette du menu n'a pas été modifiée.");
            Assert.AreEqual(newMethod, finalMethod, "La valeur de la méthode de calcul de la recette du menu n'a pas été modifiée.");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_DISP_Import()
        {
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            HomePage homePage = LogInAsAdmin();
            homePage.ClearDownloads();
            DispatchPage dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.ExportExcel_Dispatch(true);
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();
            FileInfo correctDownloadedFile = dispatchPage.GetExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, $"The file {correctDownloadedFile} is null");
            string fileName = correctDownloadedFile.Name;
            string filePath = Path.Combine(downloadsPath, fileName);
            dispatchPage.Import(filePath);
            bool isImported = dispatchPage.VerifyImportFile();
            Assert.IsTrue(isImported, "The file is not imported");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_DISP_NavigationFlecheClavier()
        {
            HomePage homePage = LogInAsAdmin();
            DispatchPage dispatchPage = homePage.GoToProduction_DispatchPage();
            dispatchPage.Filter(DispatchPage.FilterType.NotValidateOnly, true);
            dispatchPage.ClickOnFirstInput();
            bool rightClicked = dispatchPage.ArrowKeyboardClick("right");
            Assert.IsTrue(rightClicked, $"Navigation {Keys.ArrowRight} does not work!");
            bool downClicked = dispatchPage.ArrowKeyboardClick("down");
            Assert.IsTrue(downClicked, $"Navigation {Keys.ArrowDown} does not work!");
            bool leftClicked = dispatchPage.ArrowKeyboardClick("left");
            Assert.IsTrue(leftClicked, $"Navigation {Keys.ArrowLeft} does not work!");
            bool upClicked = dispatchPage.ArrowKeyboardClick("up");
            Assert.IsTrue(upClicked, $"Navigation {Keys.ArrowUp} does not work!");
        }
    }
}
