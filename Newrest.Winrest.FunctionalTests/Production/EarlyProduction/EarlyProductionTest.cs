using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.Flight;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.LoadingPlans;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.Schedule;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Datasheet;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.EarlyProduction;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Item;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using System;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.Production
{
    [TestClass]
    public class EarlyProductionTest : TestBase
    {
        private const int _timeout = 600000;

        [TestMethod]
        [Priority(0)]
        [Timeout(_timeout)]
        public void PR_EP_flight_Prepare()
        {
            string tentativeNo = "5";
            // Prepare
            string DataSheetName = "Menu Test " + tentativeNo;
            string Site = "ACE"; // GRO déjà utilisé
            string Customer = "$$ - CAT Genérico";
            string CustomerShort = "CAT Genérico";
            string GuestType = "BOB"; // or YC
            string CommercialName = "Test Salade";
            // --IsSanitization
            // select IsThawing, IsSanitization from Items where Name like '%MANDARINAS COLECTIVIDADES%';
            // --IsThawing
            // select IsThawing, IsSanitization from Items where Name like '%CANAPES SURTIDOS PRESTIGE%'
            // --IsThawing bis
            // select IsThawing, IsSanitization from Items where Name like '%HAMBURGUESA TERNERA 110 GR CONGELADA  ACE%'
            string Ingredient1 = "MANDARINAS COLECTIVIDADES";
            string Ingredient2 = "CANAPES SURTIDOS PRESTIGE";
            string Ingredient3 = "HAMBURGUESA TERNERA 110 GR CONGELADA  ACE";
            string Ingredient4 = "Test Item 0" + tentativeNo;
            // par défaut : Cocina Caliente
            string CookingMode = "Vacío";
            string ServiceName = "Service Test " + tentativeNo;
            string ServiceNameWithoutDatasheet = "Service Test Without Datasheet " + tentativeNo;
            string LoadingPlanName = "SYDNEY" + tentativeNo;
            string Route = "ACE-ACE";// "ACE-ACE"; //Site = "MAD", après clique sur CREATE, "MAD-ACE" disparait
            string Aircraft = "B717";
            string aircraftRegistration = Aircraft;
            string FlightNo = "XU71" + tentativeNo;
            // pas de choix
            string AirportFrom = "ACE";
            string AirportTo = Site;
            string ItemGroup = "AIR CANADA";
            string workshop = "Cocina Caliente";
            string taxType = TestContext.Properties["Item_TaxType"].ToString();
            string prodUnit = TestContext.Properties["Item_ProdUnit"].ToString();



            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // item hidden
            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.Filter(ItemPage.FilterType.Search, Ingredient4);
            if (itemPage.CheckTotalNumber() == 0)
            {
                ItemCreateModalPage itemCreatePage = itemPage.ItemCreatePage();
                itemCreatePage.WaitForElementExists(By.Id("IsSanitization")).Click();
                itemCreatePage.WaitForElementExists(By.Id("IsThawing")).Click();
                itemCreatePage.WaitForElementExists(By.Id("IsHACCPRecordExcluded")).Click();
                ItemGeneralInformationPage generalInfo = itemCreatePage.FillField_CreateNewItem(Ingredient4, ItemGroup, workshop, taxType, prodUnit);
                ItemCreateNewPackagingModalPage pack = generalInfo.NewPackaging();
                pack.FillField_CreateNewPackaging(Site, "KG", "1", "UD", "2", "CANARD S.A.");

            }

            // Menu>Datasheet
            DatasheetPage dataSheetPage = homePage.GoToMenus_Datasheet();
            dataSheetPage.ResetFilter();
            dataSheetPage.Filter(DatasheetPage.FilterType.DatasheetName, DataSheetName);
            dataSheetPage.Filter(DatasheetPage.FilterType.ShowAll, null);
            // pas de csutomer, on uncheck all
            dataSheetPage.Filter(DatasheetPage.FilterType.Customers, null);
            dataSheetPage.Filter(DatasheetPage.FilterType.Categories, null);
            int totalNumberDatasheet = dataSheetPage.CheckTotalNumber();
            if (totalNumberDatasheet == 0)
            {
                DatasheetCreateModalPage datasheetModal = dataSheetPage.CreateNewDatasheet();
                DatasheetDetailsPage datasheetDetails = datasheetModal.FillField_CreateNewDatasheet(DataSheetName, GuestType, Site);
                DatasheetCreateNewRecipePage newRecipe = datasheetDetails.CreateNewRecipe();
                datasheetDetails = newRecipe.FillFields_AddNewRecipeToDatasheet(CommercialName, Ingredient1);
                newRecipe = datasheetDetails.CreateNewRecipe();
                datasheetDetails = newRecipe.FillFields_AddNewRecipeToDatasheet(CommercialName, Ingredient2);
                newRecipe = datasheetDetails.CreateNewRecipe();
                datasheetDetails = newRecipe.FillFields_AddNewRecipeToDatasheet(CommercialName, Ingredient3, CookingMode);
                newRecipe = datasheetDetails.CreateNewRecipe();
                datasheetDetails = newRecipe.FillFields_AddNewRecipeToDatasheet(CommercialName, Ingredient4);
            }

            homePage.Navigate();

            // Customer>Service
            ServicePage service = homePage.GoToCustomers_ServicePage();
            service.ResetFilters();
            service.Filter(ServicePage.FilterType.Search, ServiceName);
            int totalNumberService = service.CheckTotalNumber();
            if (totalNumberService == 0)
            {
                ServiceCreateModalPage serviceModal = service.ServiceCreatePage();
                serviceModal.FillFields_CreateServiceModalPage(ServiceName, null, null, null, GuestType);
                ServiceGeneralInformationPage generalInfo = serviceModal.Create();
                ServicePricePage pricePage = generalInfo.GoToPricePage();
                ServiceCreatePriceModalPage createPrice = pricePage.AddNewCustomerPrice();
                // fait le save à la fin
                //DateTo à 3 mois car loading plan +1 mois -2 jour
                createPrice.FillFields_CustomerPrice(Site, Customer, DateUtils.Now.Date.AddDays(-7), DateUtils.Now.Date.AddMonths(3), "10", DataSheetName);
            }
            else
            {
                // maj Date From/To
                ServicePricePage pricePage = service.ClickOnFirstService();
                pricePage.UnfoldAll();
                ServiceCreatePriceModalPage editPrice = pricePage.EditFirstPrice(Site, "");
                // fait le save à la fin
                //DateTo à 3 mois car loading plan +1 mois -2 jour
                editPrice.FillFields_EditCustomerPrice(Site, Customer, DateUtils.Now.Date.AddDays(-7), DateUtils.Now.Date.AddMonths(3));
            }

            service = homePage.GoToCustomers_ServicePage();
            service.ResetFilters();
            service.Filter(ServicePage.FilterType.Search, ServiceNameWithoutDatasheet);
            totalNumberService = service.CheckTotalNumber();
            if (totalNumberService == 0)
            {
                ServiceCreateModalPage serviceModal = service.ServiceCreatePage();
                serviceModal.FillFields_CreateServiceModalPage(ServiceNameWithoutDatasheet, null, null, null, GuestType);
                ServiceGeneralInformationPage generalInfo = serviceModal.Create();
                ServicePricePage pricePage = generalInfo.GoToPricePage();
                ServiceCreatePriceModalPage createPrice = pricePage.AddNewCustomerPrice();
                // fait le save à la fin
                //DateTo à 3 mois car loading plan +1 mois -2 jour
                createPrice.FillFields_CustomerPrice(Site, Customer, DateUtils.Now.Date.AddDays(-7), DateUtils.Now.Date.AddMonths(3), "10", null);
            }
            else
            {
                // maj Date From/To
                ServicePricePage pricePage = service.ClickOnFirstService();
                pricePage.UnfoldAll();
                ServiceCreatePriceModalPage editPrice = pricePage.EditFirstPrice(Site, "");
                // fait le save à la fin
                //DateTo à 3 mois car loading plan +1 mois -2 jour
                editPrice.FillFields_EditCustomerPrice(Site, Customer, DateUtils.Now.Date.AddDays(-7), DateUtils.Now.Date.AddMonths(3));
            }

            homePage.Navigate();

            // Flight>Loading plans
            LoadingPlansPage plans = homePage.GoToFlights_LoadingPlansPage();
            plans.ResetFilter();
            plans.Filter(LoadingPlansPage.FilterType.Site, Site);
            plans.Filter(LoadingPlansPage.FilterType.Search, LoadingPlanName);
            plans.Filter(LoadingPlansPage.FilterType.StartDate, DateUtils.Now.AddDays(-7));
            int totalNumberLoadingPlans = plans.CheckTotalNumber();
            if (totalNumberLoadingPlans == 0)
            {
                LoadingPlansCreateModalPage plansCreateModal = plans.LoadingPlansCreatePage();
                // DateFrom -2 jour, DateTo 1 mois
                plansCreateModal.FillField_CreateNewLoadingPlan(LoadingPlanName, CustomerShort, Route, aircraftRegistration, Site);
                //Name already exist for the same site and customer!
                plansCreateModal.WaiForLoad();
                plansCreateModal.FillFieldLoadingPlanInformations(DateUtils.Now.Date.AddMonths(1), DateUtils.Now.Date.AddDays(-2));
                LoadingPlansDetailsPage details = plansCreateModal.Create();
                details.ClickAddGuestBtn();
                details.SelectGuest(GuestType);
                details.ClickCreateGuestBtn();
                if (GuestType == "BOB")
                {
                    details.ClickGuestBOBBtn();
                }
                else
                {
                    // YC
                    details.ClickGuestBtn();
                }
                details.AddServiceBtn();
                details.AddNewService(ServiceName);
                details.AddServiceBtn();
                details.AddNewService(ServiceNameWithoutDatasheet);
            }
            else
            {
                LoadingPlansGeneralInformationsPage plansEditModal = plans.ClickOnFirstLoadingPlan();
                plansEditModal.EditLoadingPlanInformations(DateUtils.Now.Date.AddMonths(1), DateUtils.Now.Date.AddDays(-2));
            }

            homePage.Navigate();

            // Flight>Flight
            FlightPage flight = homePage.GoToFlights_FlightPage();
            flight.ResetFilter();
            //flight.Filter(FlightPage.FilterType.HideCancelledFlight, true);
            flight.Filter(FlightPage.FilterType.SearchFlight, FlightNo);
            flight.Filter(FlightPage.FilterType.Sites, AirportFrom);
            int totalNumberFlight = flight.CheckTotalNumber();
            if (totalNumberFlight == 0)
            {
                FlightCreateModalPage flightModal = flight.FlightCreatePage();
                // fait le save à la fin
                flightModal.FillField_CreatNewFlight(FlightNo, Customer, Aircraft, AirportFrom, AirportTo);
                FlightDetailsPage flightDetails = flight.EditFirstFlight(FlightNo);

                flightDetails.AddGuestType(GuestType);
                // FIXME ComboBox trop petit sur mon écran hack
                var javaScriptExecutor = WebDriver as IJavaScriptExecutor;
                javaScriptExecutor.ExecuteScript("document.evaluate(\"//*/th[text()='Service']\", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.setAttribute(\"style\", \"width: 60px;\");");
                flight.WaiForLoad();

                flightDetails.AddService(ServiceName);
                flightDetails.AddService(ServiceNameWithoutDatasheet);

                flight = flightDetails.CloseViewDetails();
                // des colonnes du tableau disparaissent rapidement
                flight.WaiForLoad();
                Thread.Sleep(1000);
                flight.SetNewState("V");
                Thread.Sleep(1000);
                flight.WaiForLoad();
            }

            homePage.Navigate();

            // Flight>Schedules
            // D-1 => D
            SchedulePage schedule = homePage.GoToFlights_FlightSelectionPage();
            schedule.Filter(SchedulePage.FilterType.Site, "GRO");
            schedule.Filter(SchedulePage.FilterType.Site, Site);
            if (!schedule.isPageSizeEqualsTo100())
            {
                schedule.PageSize("100");
            }

            schedule.UnfoldAll();
            schedule.SetFlightProduced(FlightNo, true);
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_EP_Filter_Site()
        {
            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            EarlyProductionPage earlyProductionPage = homePage.GoToProduction_EarlyProduction();
             earlyProductionPage.ResetFilter();
            //1. Appliquer Filter site
            earlyProductionPage.Filter(EarlyProductionPage.FilterType.Site, "MAD");
            earlyProductionPage.Filter(EarlyProductionPage.FilterType.Site, "ACE");
            //2.Appliquer Filter Date
            earlyProductionPage.Filter(EarlyProductionPage.FilterType.StartDate, DateUtils.Now.AddDays(-11));
            earlyProductionPage.Filter(EarlyProductionPage.FilterType.EndDate, DateUtils.Now);
            //3.Cliquer sur Show
            earlyProductionPage.Show();
            //Données sur onglet Favorite (onglet Early Production plutôt ?)
            Assert.IsTrue(earlyProductionPage.HasEarlyProduction(), "No production");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_EP_Filter_ShowHiddenArticles()
        {
            // Arrange
            HomePage homePage = LogInAsAdmin();
            EarlyProductionPage earlyProductionPage = homePage.GoToProduction_EarlyProduction();

            earlyProductionPage.Filter(EarlyProductionPage.FilterType.ShowHiddenArticles, true);
            earlyProductionPage.Filter(EarlyProductionPage.FilterType.ShowHiddenArticles, false);

            // aller sur le second onglet
            earlyProductionPage.Filter(EarlyProductionPage.FilterType.StartDate, DateUtils.Now.AddDays(-11));
            earlyProductionPage.Filter(EarlyProductionPage.FilterType.EndDate, DateUtils.Now);
            earlyProductionPage.Show();

            //Cocher "Show hidden articles"
            // FIXME : dans le second onglet le checkbox est read-only (mais on peut cocher le label)
            earlyProductionPage.Filter(EarlyProductionPage.FilterType.ShowHiddenArticles, true, false);
            //Les données sont filtré et correspondent
            Assert.IsTrue(earlyProductionPage.HasEarlyProduction(), "No production");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_EP_FilterByFromDate()
        {
            //Prepare
            string siteACE = TestContext.Properties["SiteACE"].ToString();
            DateTime fromDate = DateUtils.Now.AddDays(-1);

            // Arrange
            var homePage = LogInAsAdmin();

            // Cliquer sur l'onglet Early Favorite
            EarlyProductionPage earlyProductionPage = homePage.GoToProduction_EarlyProduction();
            earlyProductionPage.ResetFilter();

            //1.Appliquer Filter site
            earlyProductionPage.Filter(EarlyProductionPage.FilterType.Site, siteACE);

            //2.Appliquer Filter Date J-1
            earlyProductionPage.Filter(EarlyProductionPage.FilterType.StartDate, fromDate);

            //3.Cliquer sur Show"
            earlyProductionPage.Show();

            //Assert
            bool hasEarlyProduction = earlyProductionPage.HasEarlyProduction();
            Assert.IsTrue(hasEarlyProduction, "No production");
            earlyProductionPage.UnfoldAll();

            // Vérifier que tous les dates sont supérieures à la date de From 
            bool verifFromDate = earlyProductionPage.VerifFromDate(fromDate);
            Assert.IsTrue(verifFromDate, "Le filtrage de start Day est incorrecte");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_EP_Filter_Customer()
        {
            // Arrange
            HomePage homePage = LogInAsAdmin();
            EarlyProductionPage earlyProductionPage = homePage.GoToProduction_EarlyProduction();
            //Appliquer Filter Customer
            earlyProductionPage.Filter(EarlyProductionPage.FilterType.Customer, "CAT Genérico");
            earlyProductionPage.Filter(EarlyProductionPage.FilterType.StartDate, DateUtils.Now.AddDays(-11));
            earlyProductionPage.Filter(EarlyProductionPage.FilterType.EndDate, DateUtils.Now);
            earlyProductionPage.Show();
            Assert.IsTrue(earlyProductionPage.HasEarlyProduction(), "No production");
            earlyProductionPage.Filter(EarlyProductionPage.FilterType.Customer, "CANARYFLY");
            Assert.IsFalse(earlyProductionPage.HasEarlyProduction(), "Production présent");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_EP_FoldAll_UnfoldAll()
        {
            // Arrange
            HomePage homePage = LogInAsAdmin();
            EarlyProductionPage earlyProductionPage = homePage.GoToProduction_EarlyProduction();
            earlyProductionPage.Filter(EarlyProductionPage.FilterType.StartDate, DateUtils.Now.AddDays(-11));
            earlyProductionPage.Filter(EarlyProductionPage.FilterType.EndDate, DateUtils.Now);
            earlyProductionPage.Show();

            //Survoler … Unfold all
            earlyProductionPage.UnfoldAll();
            Assert.IsTrue(earlyProductionPage.IsUnfold(), "Pas déplié");
            earlyProductionPage.FoldAll();
            Assert.IsFalse(earlyProductionPage.IsUnfold(), "Déplié");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_EP_Generate_Raw_Materials()
        {
            DateTime date = DateTime.Now.AddDays(-7);
            LogInAsAdmin();
            HomePage homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            EarlyProductionPage earlyProductionPage = homePage.GoToProduction_EarlyProduction();
            earlyProductionPage.Filter(EarlyProductionPage.FilterType.StartDate,date);
            earlyProductionPage.Filter(EarlyProductionPage.FilterType.ShowHiddenArticles, true);
            earlyProductionPage.Show();
            earlyProductionPage.GenerateRawMaterials();
            earlyProductionPage.UnfoldAll();
            bool hasRowMaterials = earlyProductionPage.HasRawMaterials();
            Assert.IsTrue(hasRowMaterials, "Pas de Raw Materials");
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void PR_EP_Add_Selection_As_Favorite()
        {
            // Prepare
            string favoriteName = "MonFav - " + new Random().Next();
            // Arrange
            HomePage homePage = LogInAsAdmin();
            EarlyProductionPage earlyProductionPage = homePage.GoToProduction_EarlyProduction();

            earlyProductionPage.Filter(EarlyProductionPage.FilterType.Customer, "CAT Genérico");
            earlyProductionPage.Filter(EarlyProductionPage.FilterType.StartDate, DateUtils.Now.AddDays(-11));
            earlyProductionPage.Filter(EarlyProductionPage.FilterType.EndDate, DateUtils.Now);
            earlyProductionPage.Show();

            earlyProductionPage.UnfoldAll();
            earlyProductionPage.SelectFirstRecipe();
            //"1. Survoler +
            //2.Cliquer Add selection as favorite"	
            EarlyProductionFavoriteModal favModal = earlyProductionPage.AddSelectionAsFavorite();
            favModal.Fill(favoriteName);
            favModal.Submit();
            earlyProductionPage.CLickOnFavoritesTab();
            Assert.IsTrue(earlyProductionPage.HasFavorite(favoriteName), "Pas dans favorite");
            earlyProductionPage.DeleteFavorite(favoriteName);
        }
    }
}
