using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Accounting.Reporting;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.Flight;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.LoadingPlans;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.Schedule;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Datasheet;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.ProductionManagement;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Item;
using Newrest.Winrest.FunctionalTests.PageObjects.TabletApp;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace Newrest.Winrest.FunctionalTests.TabletApp
{
    [TestClass]
    public class HACCPProdInflightTest : TestBase
    {
        private const int Timeout = 600000;
        /// <summary>
        /// Mise en place du paramétrage Prod. inflight 
        /// </summary>
        [TestMethod]
        [Priority(0)]
        [Timeout(Timeout)]
        public void TB_ProdInflight_SetConfigWinrest()
        {
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Check all documents in application settings
            var applicationSettings = homePage.GoToApplicationSettings();

            //Add documents & columns
            var customizableColumns = applicationSettings.GoToCustomizableColumns();
            customizableColumns.AddAllCustomizableColumns();
            //Attention ! Seul 3 documents sont paramétrés HACCP - Sanithization, HACCP - Thawing, HACCP - TSU et HACCP - Tray Setup
            //Vérifier si documents bien cochés dans App settings au niveau de la ligne DocumentsDisplay
        }

        [TestMethod]
        [Priority(1)]
        [Timeout(Timeout)]
        public void TB_HACCP_ProdInflight_Prepare()
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
            string Ingredient3 = "HAMBURGUESA TERNERA 110 GR CONGELADA ACE";
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
            var homePage = LogInAsAdmin();

            // item hidden
            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.Filter(ItemPage.FilterType.Search, Ingredient1);
            if (itemPage.CheckTotalNumber() == 0)
            {
                ItemCreateModalPage itemCreatePage = itemPage.ItemCreatePage();
                itemCreatePage.WaitForElementExists(By.Id("IsSanitization")).Click();
                itemCreatePage.WaitForElementExists(By.Id("IsThawing")).Click();
                itemCreatePage.WaitForElementExists(By.Id("IsHACCPRecordExcluded")).Click();
                ItemGeneralInformationPage generalInfo = itemCreatePage.FillField_CreateNewItem(Ingredient1, ItemGroup, workshop, taxType, prodUnit);
                ItemCreateNewPackagingModalPage pack = generalInfo.NewPackaging();
                pack.FillField_CreateNewPackaging(Site, "KG", "1", "UD", "2", "CANARD S.A.");
            }
            else
            {
                ItemGeneralInformationPage generalInfo = itemPage.ClickFirstItem();
                generalInfo.Check_Is(ItemGeneralInformationPage.isValueChecked.IsSanitization, false);

            }

            itemPage.BackToList();
            itemPage.Filter(ItemPage.FilterType.Search, Ingredient2);

            if (itemPage.CheckTotalNumber() == 0)
            {
                ItemCreateModalPage itemCreatePage = itemPage.ItemCreatePage();
                itemCreatePage.WaitForElementExists(By.Id("IsSanitization")).Click();
                itemCreatePage.WaitForElementExists(By.Id("IsThawing")).Click();
                itemCreatePage.WaitForElementExists(By.Id("IsHACCPRecordExcluded")).Click();
                ItemGeneralInformationPage generalInfo = itemCreatePage.FillField_CreateNewItem(Ingredient2, ItemGroup, workshop, taxType, prodUnit);
                ItemCreateNewPackagingModalPage pack = generalInfo.NewPackaging();
                pack.FillField_CreateNewPackaging(Site, "KG", "1", "UD", "2", "CANARD S.A.");
            }
            else
            {
                ItemGeneralInformationPage generalInfo = itemPage.ClickFirstItem();
                generalInfo.Check_Is(ItemGeneralInformationPage.isValueChecked.IsSanitization, false);

            }

            itemPage.BackToList();
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
            else
            {
                ItemGeneralInformationPage generalInfo = itemPage.ClickFirstItem();
                generalInfo.Check_Is(ItemGeneralInformationPage.isValueChecked.IsSanitization, true);

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
                editPrice.WaitPageLoading();
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
                plansCreateModal.WaitPageLoading();
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
                flight.WaitPageLoading();

                flightDetails.AddService(ServiceName);
                flightDetails.AddService(ServiceNameWithoutDatasheet);

                flight = flightDetails.CloseViewDetails();
                // des colonnes du tableau disparaissent rapidement
                flight.WaitPageLoading();
                var validation = flight.WaitForElementExists(By.XPath("//*[@id=\"flightTable\"]/tbody/tr[2]/td[20]/div/ul/li[2]"));
                // ça marche mais pas dynamiquement, faut recharger la page
                string js = validation.GetAttribute("onclick");
                javaScriptExecutor = WebDriver as IJavaScriptExecutor;
                javaScriptExecutor.ExecuteScript(js);
                flight.WaitPageLoading();
            }

            homePage.Navigate();

            // Flight>Schedules 
            // D-1 => D
            SchedulePage schedule = homePage.GoToFlights_FlightSelectionPage();
            schedule.Filter(SchedulePage.FilterType.Site, Site);
            schedule.WaitPageLoading();
            if (!schedule.isPageSizeEqualsTo100())
            {
                schedule.WaitPageLoading();
                schedule.PageSize("100");
            }
            schedule.UnfoldAll();
            schedule.SetFlightProduced(FlightNo, true);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_ProdInflight_SaveDocumentHACCP()
        {
            // Prepare
            string DocSite = "ACE";
            string DocTitle = "HACCP - Sanitization";
            string DocReport = "HACCPSanitization";
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            //HACCP_Sanitization_9_20220301_015007.pdf
            string DocFileNamePdfBegin = "HACCP_Sanitization_9_";

            // Arrange
            var homePage = LogInAsAdmin();


            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Purge(downloadsPath, DocFileNamePdfBegin);

            //1.sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);

            //2. aller sur ""prod inflight""
            tabletAppPage.GotoTabletApp_ProdInFlight();
            tabletAppPage.WaitPageLoading();

            //3. Sélectionne une date FROM TO
            tabletAppPage.Filter(TabletAppPage.FilterType.From, DateUtils.Now.Date);
            tabletAppPage.Filter(TabletAppPage.FilterType.To, DateUtils.Now.Date.AddDays(4));

            //4. Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.Select("mat-select-2", DocTitle);

            //5.Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");

            //6. choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            tabletAppPage.WaitHACCPHorizontalProgressBar();

            //7.remplir une ligne de données
            tabletAppPage.addLineSanitizationProdInFlight("HAMBURGUESA TERNERA 110 GR CONGELADA ACE", "2.2", "80.0", "10", true, false, "MyComments", "MyPreparedBy");

            //8.Cliquer sur le bouton ""save and validate "" , une pop up s'ouvre de validation s'ouvre, puis cliquer sur ""validate"""
            tabletAppPage.ClickButton("Save and validate");
            tabletAppPage.ClickButton("Validate");
            tabletAppPage.ClickButton("Ok");

            tabletAppPage.Close();
            tabletAppPage.WaitPageLoading();
            // verify Accouting>Reporting
            homePage.Navigate();
            homePage.ClearDownloads();

            ReportingPage reportingPage = homePage.GoToAccounting_Reporting();
            string DocTypeOfReport = reportingPage.DevPathDocTypeOfRecord();

            //1.Sélectionner type of report: HACCP
            //2. sélectionner le site sélectionné et sélectionner la production date
            reportingPage.FillReportingPage(DocTypeOfReport, DocSite);
            //3. recupérer le rapport HACCP
            var offset = reportingPage.TableHasDocument(DocReport);
            Assert.IsNotNull(offset, "Document non trouvé");
            reportingPage.PrintDownload(offset);

            //4.Les données pré remplies dans le document sont les mêmes dans le rapport
            //5. Vérifier dans le print...
            FileInfo fi = reportingPage.FindProdInFlightPDF(downloadsPath, DocFileNamePdfBegin);
            Assert.IsNotNull(fi, DocFileNamePdfBegin + " pas trouvé");
            Assert.IsTrue(fi.Exists, "pdf non généré");

            reportingPage.FixPDFMagic(fi.FullName);
            PdfDocument document = PdfDocument.Open(fi.FullName);
            IEnumerable<Word> allWords = Enumerable.Empty<Word>();
            for (int i = 1; i <= document.NumberOfPages; i++)
            {
                Page currentPage = document.GetPage(i);
                IEnumerable<Word> words = currentPage.GetWords();
                allWords = allWords.Concat(words);
            }
            // le nom du rapport
            var wordArray = allWords.ToArray();
            var title = wordArray[0].Text + wordArray[1].Text;
            Assert.AreEqual(1, (title.Contains("HACCPSANITIZATION") ? 1 : 0) + (title.Contains("HACCPSanitization") ? 1 : 0), "nom du rapport non trouvé dans " + fi.FullName);
            // la date de production,
            Assert.AreNotEqual(0, allWords.Count(w => w.Text == DateUtils.Now.Date.ToString("dd/MM/yyyy")), "date de production non trouvé dans " + fi.FullName);
            // le site,
            Assert.IsTrue(allWords.Any(w => w.Text == DocSite), "site non trouvé dans " + fi.FullName);
            // les dates from et to,
            Assert.AreNotEqual(0, allWords.Count(w => w.Text == DateUtils.Now.Date.ToString("dd/MM/yyyy")), "date from non trouvé dans " + fi.FullName);
            Assert.AreEqual(1, allWords.Count(w => w.Text == DateUtils.Now.Date.AddDays(4).ToString("dd/MM/yyyy")), "date to non trouvé dans " + fi.FullName);
            // l'intitulé des colonnes
            Assert.AreEqual(3, allWords.Count(w => w.Text == "Group" || w.Text == "Item"), "colonne 'Group - Item' non trouvé dans " + fi.FullName);
            Assert.AreEqual(4, allWords.Count(w => w.Text == "Total" || w.Text == "Forecast" || w.Text == "Gross" || w.Text == "weight"), "colonne 'Total Forecast Gross weight' non trouvé dans " + fi.FullName);
            Assert.AreEqual(2, allWords.Count(w => w.Text == "Quantity"), "colonnes 'Quantity' non trouvé dans " + fi.FullName);
            Assert.AreEqual(3, allWords.Count(w => w.Text == "Desinfection" || w.Text == "Quantity"), "colonne 'Desinfection Quantity' non trouvé dans " + fi.FullName);
            Assert.AreEqual(2, allWords.Count(w => w.Text == "Disinfection" || w.Text == "Time"), "colonne 'Disinfection Time (min)' non trouvé dans " + fi.FullName);
            Assert.AreEqual(1, allWords.Count(w => w.Text == "Rinsing"), "colonne 'Rinsing' non trouvé dans " + fi.FullName);
            Assert.AreEqual(3, allWords.Count(w => w.Text == "Final" || w.Text == "visual" || w.Text == "inspection"), "colonne 'Final visual inspection' non trouvé dans " + fi.FullName);
            Assert.AreEqual(3, allWords.Count(w => w.Text == "Comments/" || w.Text == "corrective" || w.Text == "action"), "colonne 'Comments/ corrective action' non trouvé dans " + fi.FullName);
            Assert.AreEqual(2, allWords.Count(w => w.Text == "Prepared" || w.Text == "by"), "colonne 'Prepared by' non trouvé dans " + fi.FullName);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_ProdInflight_OpenDocumentHACCP()
        {
            // Prepare
            string DocSite = "ACE";
            string DocTitle = "HACCP - Sanitization";
            // Arrange
            var homePage = LogInAsAdmin();


            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            //1.sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);

            //2. aller sur ""prod inflight""
            tabletAppPage.GotoTabletApp_ProdInFlight();
            Thread.Sleep(2000);

            //sélectionner un document haccp dans le menu déroulant
            tabletAppPage.Select("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-2", DocTitle);

            //Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");

            //choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            tabletAppPage.WaitHACCPHorizontalProgressBar();

            //Le document s'ouvre bien
            tabletAppPage.ClickButton("Save");
            tabletAppPage.ClickButton("Cancel");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_ProdInflight_SaveComment()
        {
            // Prepare
            string DocSite = "ACE";
            string DocCommentary = "Hello World ProdInflight !";
            string DocTitle = "HACCP - Sanitization";

            // Arrange
            var homePage = LogInAsAdmin();


            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            //1.sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);

            //2. aller sur ""prod inflight""
            tabletAppPage.GotoTabletApp_ProdInFlight();
            Thread.Sleep(2000);

            //3. Sélectionne une date FROM TO
            // FIXME peut-être du code déjà existant ces datepicker
            tabletAppPage.Filter(TabletAppPage.FilterType.From, DateUtils.Now.Date.AddDays(-5));
            tabletAppPage.Filter(TabletAppPage.FilterType.To, DateUtils.Now.Date.AddDays(5));

            //1.saisir un commentaire dans la case ""comment""
            var option = tabletAppPage.WaitForElementIsVisible(By.Id("mat-input-2"));
            option.Click();
            Thread.Sleep(2000);
            option.SendKeys(DocCommentary);
            Thread.Sleep(2000);

            //5. Choisir le doc haccp dans le menu déroulant ,
            tabletAppPage.Select("mat-select-2", DocTitle);
            //puis cliquer sur next step pour l'ouvrir
            tabletAppPage.ClickBaseButton("NEXT STEP");
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            tabletAppPage.WaitHACCPHorizontalProgressBar();

            //6.enregistrer le document définitivement en cliquant sur save and validate
            tabletAppPage.ClickButton("Save and validate");
            tabletAppPage.ClickButton("Validate");
            tabletAppPage.ClickButton("Ok");

            //7.vérifier sur la partie winrest dans accounting => reporting
            //que le commentaire est présent dans le champs ""commentary ""
            //(après avoir sélectionné le type de rapport ""HACCP""
            //et avoir choisi la même date de production et le même site "
            homePage.Navigate();
            ReportingPage reportingPage = homePage.GoToAccounting_Reporting();
            string DocTypeOfReport = reportingPage.DevPathDocTypeOfRecord();
            reportingPage.FillReportingPage(DocTypeOfReport, DocSite);
            string xPathTableauNbColonne = "//*[@id=\"table-records\"]/tbody/tr[1]/td";
            int tableauNbColonne = WebDriver.FindElements(By.XPath(xPathTableauNbColonne)).Count;
            string xPathTableau = "//*[@id=\"table-records\"]/tbody/tr[*]/td[4]/div/p";
            if (tableauNbColonne == 6)
            {
                // sur Dev
                xPathTableau = "//*[@id=\"table-records\"]/tbody/tr[*]/td[5]/div/p";
            }
            if (tableauNbColonne == 7)
            {
                // sur Dev
                xPathTableau = "//*[@id=\"table-records\"]/tbody/tr[*]/td[6]/div/p";
            }
            var tableau = WebDriver.FindElements(By.XPath(xPathTableau));
            Assert.IsTrue(tableau.Count > 0, "commentaire non présent dans le Accounting>Reporting");
            Assert.AreEqual(DocCommentary, tableau[tableau.Count - 1].Text, "commentaire non présent dans le Accounting>Reporting (cas 2)");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_ProdInflight_SaveDocumentTemporary()
        {
            // Prepare
            string DocSite = TestContext.Properties["SiteACE"].ToString();
            string DocTitle = "HACCP - Sanitization";
            string DocFileName = "TempTestPIF_wrfr.user01";
            // Arrange
            var homePage = LogInAsAdmin();


            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            //sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);

            //aller sur ""prod inflight""
            tabletAppPage.GotoTabletApp_ProdInFlight();
            tabletAppPage.WaitPageLoading();

            //1.Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.Select("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-2", DocTitle);

            //2. Cliquer sur next pour ouvrir le document

            tabletAppPage.clickNextStep(DocTitle);
            //3.choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            tabletAppPage.WaitHACCPHorizontalProgressBar();

            //4. remplir une ligne de données
            tabletAppPage.addLineSanitizationProdInFlight("HAMBURGUESA TERNERA 110 GR CONGELADA ACE", "2.2", "80.0", "10", true, false, "MyComments", "MyPreparedBy");
            //5. Cliquer sur le bouton ""save "" => winrest propose un nom du temporary par défaut qui peut être modifié
            tabletAppPage.ClickButton("Save");
            tabletAppPage.FileFileNameProdInFlight1(DocFileName);
            //6.cliquer sur ""validate""
            tabletAppPage.ClickButton("Validate");

            // écraser le fichier
            tabletAppPage.ClickButton("Yes");

            //7.une pop de validation apparait et puis cliquer sur OK
            tabletAppPage.ClickButton("Ok");
            tabletAppPage.Close();
            tabletAppPage.WaitPageLoading();
            //8.Revenir vers la page d'acceuil de prod inflight , sélectionner le doc haccp
            homePage.Navigate();
            tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            // si je suis déjà passé par /tablet, ça affiche ACE en haut à gauche, et le reste de la page est blanche
            tabletAppPage.GotoTabletApp_ProdInFlight();
            tabletAppPage.WaitPageLoading();

            //Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.Select("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-2", DocTitle);
            //Dans la barre des ""temporary"" on devrait retrouver le doc temporary sauvegardé"" avec la bonne nomination"
            tabletAppPage.SelectAction("mat-select-6", DocFileName, true);

            //2.vérifier qu'on récupère bien les données enregistrées sont OK
            //Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");
            //choisir une date de production et cliquer sur print
            //tabletAppPage.ClickButton("Print");
            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            tabletAppPage.WaitHACCPHorizontalProgressBar();

            //Assert
            var qty = tabletAppPage.WaitForElementIsVisible(By.Id("mat-input-1"));
            Assert.AreEqual("2.2", qty.GetAttribute("value"), "mauvais qty");
            var desinfection = tabletAppPage.WaitForElementIsVisible(By.Id("mat-input-3"));
            Assert.AreEqual("80", desinfection.GetAttribute("value"), "mauvais Desinfection");
            var desinfectionTime = tabletAppPage.WaitForElementIsVisible(By.Id("mat-input-5"));
            Assert.AreEqual("10", desinfectionTime.GetAttribute("value"), "mauvais Desinfection");
            IWebElement Rinsing;
            Rinsing = tabletAppPage.WaitForElementExists(By.Id("mat-mdc-checkbox-1-input"));
            string RinsingBool = Rinsing.GetAttribute("class");
            Assert.IsFalse(RinsingBool.Contains("mdc-checkbox--selected"), "mauvais Rising");
            IWebElement Final;
            Final = tabletAppPage.WaitForElementExists(By.Id("mat-mdc-checkbox-2-input"));
            string FinalBool = Final.GetAttribute("class");
            Assert.IsTrue(FinalBool.Contains("mdc-checkbox--selected"), "mauvais Final");
            var PreparedBy = tabletAppPage.WaitForElementIsVisible(By.Id("mat-input-10"));
            Assert.AreEqual("MyPreparedBy", PreparedBy.GetAttribute("value"), "mauvais PreparedBy");

        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_ProdInflight_FilterServices()
        {
            // Prepare
            string DocSite = "ACE";
            string DocTitle = "HACCP - Sanitization";
            string DocIngredient = "ItemForNormalProductOnly";
            DateTime dateFrom = new DateTime(2024, 9, 15);
            //DateTime dateTo = new DateTime(2024, 9, 17);

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            //1.sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);

            //2. aller sur ""prod inflight""
            tabletAppPage.GotoTabletApp_ProdInFlight();
            tabletAppPage.WaitPageLoading();
            //3. Sélectionne une date FROM TO
            tabletAppPage.Filter(TabletAppPage.FilterType.From, dateFrom);
            //tabletAppPage.Filter(TabletAppPage.FilterType.To, dateTo);

            string DocService = tabletAppPage.IsDev() ? "Service Test 5" : "ServiceForNormalProductOnly";
            int nbLigneSelected = tabletAppPage.Filter(TabletAppPage.FilterType.Services, DocService);
            Assert.IsTrue(nbLigneSelected > 0, "Tableau des services vide");

            tabletAppPage.Select("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-2", DocTitle);

            //2. Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");
            //3.choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            tabletAppPage.WaitHACCPHorizontalProgressBar();

            string xPathResult = "//*[contains(text(),'" + DocIngredient + "')]/parent::div";
            Assert.IsNotNull(tabletAppPage.WaitForElementIsVisible(By.XPath(xPathResult)));

        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_ProdInflight_ChoiceDocumentHACCP()
        {
            // Prepare
            string DocSite = "ACE";
            string DocTitle = "HACCP - Sanitization";
            // Arrange
            var homePage = LogInAsAdmin();


            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            //sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);

            //aller sur ""prod inflight""
            tabletAppPage.GotoTabletApp_ProdInFlight();
            Thread.Sleep(2000);

            // sélectionner un document haccp dans le menu déroulant
            tabletAppPage.Select("mat-select-2", DocTitle);
            // le document remonte bien dans la liste des documents à sélectionner lorsqu'il est activé
            tabletAppPage.SelectAction("mat-select-2", DocTitle);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_ProdInflight_SaveDocumentTemporaryCheckDate()
        {
            // Prepare
            string DocSite = "ACE";
            string DocTitle = "HACCP - Thawing";
            string DocTmpName = "TempTestPIFCheckDate_wrfr.user01";
            DateTime dateFrom = new DateTime(2024, 9, 11);
            DateTime dateTo = new DateTime(2024, 9, 17);
            string DocFileName = "HACCP_Thawing_" + DocSite + "_" + dateFrom.ToString("ddMMyyyy") + "_" + dateTo.ToString("ddMMyyyy") + ".pdf";
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            // Arrange
            var homePage = LogInAsAdmin();

            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Purge(downloadsPath, "HACCP_Thawing_" + DocSite + "_");

            //1. sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);

            //2. aller sur ""prod inflight""
            tabletAppPage.GotoTabletApp_ProdInFlight();
            Thread.Sleep(2000);

            //3. Sélectionne une date FROM TO
            tabletAppPage.Filter(TabletAppPage.FilterType.From, dateFrom);
            tabletAppPage.Filter(TabletAppPage.FilterType.To, dateTo);

            //4. Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.WaitPageLoading();
            tabletAppPage.Select("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-2", DocTitle);

            //5. Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");
            //6. choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            tabletAppPage.WaitHACCPHorizontalProgressBar();

            // 7. renseigner une date dans la colonne où le format date est paramétré
            string ExpDate = DateUtils.Now.Date.AddYears(2).ToString("dd/MM/yyyy");
            tabletAppPage.addLineThawingProdInFlight("CANAPES SURTIDOS PRESTIGE", "2", ExpDate, "15", DateUtils.Now.Date.ToString("dd/MM/yyyy"), DateUtils.Now.Date.AddYears(1).ToString("dd/MM/yyyy"), "MyComments", "MyPreparedBy");

            // 8. cliquer sur save pour enregistrer un document temporary
            tabletAppPage.ClickButton("Save");
            tabletAppPage.FileFileNameProdInFlight0(DocTmpName);
            tabletAppPage.ClickButton("Validate");
            try
            {
                // écraser le fichier
                tabletAppPage.ClickButton("Yes");
            }
            catch
            {
                // nouveau fichier
            }
            //7.une pop de validation apparait et puis cliquer sur OK
            tabletAppPage.ClickButton("Ok");

            tabletAppPage.Close();
            Thread.Sleep(2000);

            homePage.Navigate();
            tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            //9. récupérer ensuite le document temporary dans la page d'acceuil de prod inflight dans la section "" temporary""
            tabletAppPage.GotoTabletApp_ProdInFlight();
            Thread.Sleep(2000);
            //Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.WaitPageLoading();
            tabletAppPage.Select("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-2", DocTitle);
            //Dans la barre des ""temporary"" on devrait retrouver le doc temporary sauvegardé"" avec la bonne nomination"
            tabletAppPage.SelectAction("mat-select-6", DocTmpName, true);
            //Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");
            //choisir une date de production et cliquer sur print
            //tabletAppPage.ClickButton("Print");
            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            tabletAppPage.WaitHACCPHorizontalProgressBar();

            //10. vérifier que la date s'est bien enregistrée sur le document temporary
            var qty = tabletAppPage.WaitForElementIsVisible(By.Id("mat-input-3"));
            Assert.AreEqual(ExpDate, qty.GetAttribute("value"), "mauvais ExpDate");

            //11.puis cliquer sur print pour générer le rapport haccp en pdf
            tabletAppPage.ClickButton("Print");
            FileInfo fi = new FileInfo(Path.Combine(downloadsPath, DocFileName));
            int counter = 0;
            while (!fi.Exists && counter < 10)
            {
                Thread.Sleep(2000);
                fi.Refresh();
                counter++;
            }

            //12.vérifier que la date est correcte sur le rapport généré "
            Assert.IsTrue(fi.Exists, DocFileName + " non généré");

            tabletAppPage.Close();
            Thread.Sleep(2000);

            PdfDocument document = PdfDocument.Open(fi.FullName);

            List<string> mots = new List<string>();
            foreach (Page page in document.GetPages())
            {
                IEnumerable<Word> words = page.GetWords();
                foreach (Word word in words)
                {
                    mots.Add(word.Text);
                }
            }
            // le nom du rapport
            Assert.AreNotEqual(0, mots.Count(w => w == "CANAPES" || w == "SURTIDOS" || w == "PRESTIGE"), "CANAPES SURTIDOS PRESTIGE non trouvé dans " + fi.FullName);
            Assert.AreNotEqual(0, mots.Count(w => w == "2"), "2 non trouvé dans " + fi.FullName);
            Assert.AreNotEqual(0, mots.Count(w => w == ExpDate), ExpDate + " non trouvé dans " + fi.FullName);
            Assert.AreNotEqual(0, mots.Count(w => w == "15"), "15 non trouvé dans " + fi.FullName);
            Assert.AreNotEqual(0, mots.Count(w => w == DateUtils.Now.Date.ToString("dd/MM/yyyy")), DateUtils.Now.Date.ToString("dd/MM/yyyy") + "non trouvé dans " + fi.FullName);
            Assert.AreNotEqual(0, mots.Count(w => w == DateUtils.Now.Date.AddYears(1).ToString("dd/MM/yyyy")), DateUtils.Now.Date.AddYears(1).ToString("dd/MM/yyyy") + "non trouvé dans " + fi.FullName);
            Assert.AreNotEqual(0, mots.Count(w => w == "MyComments"), "MyComments non trouvé dans " + fi.FullName);
            Assert.AreNotEqual(0, mots.Count(w => w == "MyPreparedBy"), "MyPreparedBy non trouvé dans " + fi.FullName);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_ProdInflight_SaveDocumentCheckDate()
        {
            // Prepare
            string DocSite = "ACE";
            string DocTitle = "HACCP - Thawing";
            string DocReport = "HACCPThawing";
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            //HACCP_Sanitization_9_20220301_015007.pdf
            string DocFileNamePdfBegin = "HACCP_Thawing_9_";

            var homePage = LogInAsAdmin();

            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Purge(downloadsPath, DocFileNamePdfBegin);

            //1. sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);

            //2. aller sur ""prod inflight""
            tabletAppPage.GotoTabletApp_ProdInFlight();
            Thread.Sleep(2000);

            //3. Sélectionne une date FROM TO
            tabletAppPage.Filter(TabletAppPage.FilterType.From, DateUtils.Now.Date.AddDays(-6));
            tabletAppPage.Filter(TabletAppPage.FilterType.To, DateUtils.Now.Date.AddDays(6));

            //4. Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.Select("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-2", DocTitle);

            //5. Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");

            //6. choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            tabletAppPage.WaitHACCPHorizontalProgressBar();

            // 7. renseigner une date dans la colonne où le format date est paramétré
            string ExpDate = DateUtils.Now.Date.AddYears(2).ToString("dd/MM/yyyy");
            tabletAppPage.addLineThawingProdInFlight("CANAPES SURTIDOS PRESTIGE", "2", ExpDate, "15", DateUtils.Now.Date.ToString("dd/MM/yyyy"), DateUtils.Now.Date.AddYears(1).ToString("dd/MM/yyyy"), "MyComments", "MyPreparedBy");

            //8. puis cliquer sur ""save and validate"" pour enregistrer le document définitivement
            tabletAppPage.ClickButton("Save and validate");
            tabletAppPage.ClickButton("Validate");
            tabletAppPage.ClickButton("Ok");
            tabletAppPage.Close();
            Thread.Sleep(2000);
            //9.aller sur winrest dans la partie accounting=> reporting...
            homePage.Navigate();
            homePage.ClearDownloads();

            ReportingPage reportingPage = homePage.GoToAccounting_Reporting();
            string DocTypeOfReport = reportingPage.DevPathDocTypeOfRecord();

            //...pour récupérer le rapport pdf généré(sélectionner le type de rapport HACCP, puis sélectionner le même site et la même date)"
            reportingPage.FillReportingPage(DocTypeOfReport, DocSite);
            //recupérer le rapport HACCP
            var offset = reportingPage.TableHasDocument(DocReport);
            Assert.IsNotNull(offset, "Document non trouvé");
            reportingPage.PrintDownload(offset);

            FileInfo fi = reportingPage.FindProdInFlightPDF(downloadsPath, DocFileNamePdfBegin);
            Assert.IsNotNull(fi, "pdf non trouvé");
            Assert.IsTrue(fi.Exists, "pdf non généré");

            reportingPage.FixPDFMagic(fi.FullName);
            PdfDocument document = PdfDocument.Open(fi.FullName);
            Page page1 = document.GetPage(1);
            IEnumerable<Word> words = page1.GetWords();
            Assert.AreEqual(3, words.Count(w => w.Text == "HACCP" || w.Text == "Thawing"), "nom du rapport non trouvé dans " + fi.FullName);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_ProdInflight_FilterCustomers()
        {
            // Prepare
            string DocSite = "ACE";
            string DocTitle = "HACCP - Tray Setup";
            string DocCustomer = "CAT Genérico";
            DateTime dateFrom = new DateTime(2024, 9, 11);

            // Arrange
            var homePage = LogInAsAdmin();

            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            //1. sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);

            //2. aller sur ""prod inflight""
            tabletAppPage.GotoTabletApp_ProdInFlight();
            Thread.Sleep(2000);

            tabletAppPage.Filter(TabletAppPage.FilterType.From, dateFrom);
            int nbLigneSelected = tabletAppPage.Filter(TabletAppPage.FilterType.Customers, DocCustomer);
            Assert.IsTrue(nbLigneSelected > 0, "Tableau des customers vide");

            tabletAppPage.Select("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-2", DocTitle);

            //2. Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");
            //3.choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            tabletAppPage.WaitHACCPHorizontalProgressBar();
            string xPathResult = "//*[contains(text(),'" + DocCustomer + "')]";
            IWebElement customerElement = tabletAppPage.WaitForElementExists(By.XPath(xPathResult));

            // Vérification que l'élément est trouvé
            Assert.IsNotNull(customerElement, "Les données qui remontent dans les documents HACCP ne correspondent pas au groupe d'éléments sélectionné");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_ProdInflight_FilterServiceCategories()
        {
            // Prepare
            string DocSite = "ACE";
            string DocTitle = "HACCP - Tray Setup";
            string DocServiceCategories = "BEBIDAS";
            //DateTime dateFrom = new DateTime(2024, 9, 11);
            //DateTime dateTo = new DateTime(2024, 9, 17);
            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            ServicePage servicePage = homePage.GoToCustomers_ServicePage();
            string DocService = servicePage.IsDev() ? "Service Test 5" : "Service for flight";

            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Sites, DocSite);
            servicePage.Filter(ServicePage.FilterType.Search, DocService);
            ServicePricePage servicePricePage = servicePage.ClickOnFirstService();
            servicePricePage.ToggleLastPrice();
            DateTime dateFrom = DateTime.Parse(servicePricePage.GetPriceDateFrom().Replace("From:", "").Trim());
            DateTime dateTo = DateTime.Parse(servicePricePage.GetPriceDateTo().Replace("To:", "").Trim());

            homePage.Navigate();
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            //1. sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);

            //2. aller sur ""prod inflight""
            tabletAppPage.GotoTabletApp_ProdInFlight();
            tabletAppPage.WaitPageLoading();


            //3. Sélectionne une date FROM TO
            tabletAppPage.Filter(TabletAppPage.FilterType.From, dateFrom);
            tabletAppPage.Filter(TabletAppPage.FilterType.To, dateTo);

            int nbLigneSelected = tabletAppPage.Filter(TabletAppPage.FilterType.ServiceCategories, DocServiceCategories);
            Assert.IsTrue(nbLigneSelected > 0, "Tableau des services vide");

            tabletAppPage.Select("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-2", DocTitle);

            //2. Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");
            //3.choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            // le ServiceCategorie est dans le Service info.
            string xPathResult = "//*[contains(text(),'" + DocService + "')]";
            Assert.IsNotNull(tabletAppPage.WaitForElementExists(By.XPath(xPathResult)));
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_ProdInflight_FilterGuestType()
        {
            // Prepare
            string DocSite = TestContext.Properties["SiteACE"].ToString();
            string DocTitle = "HACCP - Tray Setup";
            string DocServiceGuestType = TestContext.Properties["GuestNameBob"].ToString();
            DateTime dateFrom = new DateTime(2024, 9, 15);

            // Arrange
            var homePage = LogInAsAdmin();

            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            //1. sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);

            //2. aller sur ""prod inflight""
            tabletAppPage.GotoTabletApp_ProdInFlight();
            tabletAppPage.WaitPageLoading();
             tabletAppPage.Filter(TabletAppPage.FilterType.From, dateFrom);
            tabletAppPage.WaitPageLoading();

            int nbLigneSelected = tabletAppPage.Filter(TabletAppPage.FilterType.GuestType, DocServiceGuestType);
            Assert.IsTrue(nbLigneSelected > 0, "Tableau des services vide");

            tabletAppPage.Select("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-2", DocTitle);

            //2. Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");
            //3.choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            // le ServiceCategorie est dans le Service info.
            var guestType = tabletAppPage.GetGuestType();
            
            Assert.AreEqual(guestType, DocServiceGuestType);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_ProdInflight_FilterRecipeType()
        {
            // Prepare
            string DocSite = "ACE";
            string DocTitle = "HACCP - Bakery"; //"Recipe report detailled v2";
            // dans le datasheet Menu Test 4,, on a les Recipe Test Salade      
            string DocRecipeType = "ENSALADA CREW";
            DateTime dateFrom = new DateTime(2024, 9, 11);
            DateTime dateTo = new DateTime(2024, 9, 17);
            // Arrange
            var homePage = LogInAsAdmin();


            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            //1. sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);

            //2. aller sur ""prod inflight""
            tabletAppPage.GotoTabletApp_ProdInFlight();
            tabletAppPage.WaitPageLoading();

            //3. Sélectionne une date FROM TO
            tabletAppPage.Filter(TabletAppPage.FilterType.From, dateFrom);
            tabletAppPage.Filter(TabletAppPage.FilterType.To, dateTo);

            int nbLigneSelected = tabletAppPage.Filter(TabletAppPage.FilterType.RecipeType, DocRecipeType);
            Assert.IsTrue(nbLigneSelected > 0, "Tableau des services vide");

            tabletAppPage.WaitPageLoading();
            tabletAppPage.Select("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-2", DocTitle);

            //2. Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");
            //3.choisir une date de production et cliquer sur print
            // cocher Datasheet
            //tabletAppPage.WaitForElementIsVisible(By.XPath("//*[@formcontrolname='showDatasheets']/mat-radio-button[2]")).Click();
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            tabletAppPage.WaitHACCPHorizontalProgressBar();
            tabletAppPage.MenuTestModal();
            tabletAppPage.ToggleTablet();
            string ingredientName = tabletAppPage.GetNameIngredient();
            tabletAppPage.Close();

            homePage.Navigate();
            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();

            itemPage.ResetFilter();
            itemPage.Filter(ItemPage.FilterType.Search, ingredientName);
            itemPage.Filter(ItemPage.FilterType.ShowAll, true);
            var itemGeneralInformationsPage = itemPage.ClickOnFirstItem();
            var itemUseCasePage = itemGeneralInformationsPage.ClickOnUseCasePage();
            itemUseCasePage.ResetFilter();

            // on utilise le filtre search sur le recipe name du premier élément de la liste
            itemUseCasePage.Filter(ItemUseCasePage.FilterType.UncheckAllRecipeType, DocRecipeType);
            var numberOfItem = itemUseCasePage.CheckTotalNumber();

            Assert.AreNotEqual(0, numberOfItem, "Les données qui remontent dans les documents haccp ne correspondent pas au Filtrages sélectionné.");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_ProdInflight_FilterItemGroups()
        {
            string DocSite = TestContext.Properties["SiteACE"].ToString();
            string DocTitle = "HACCP - Sanitization";
            string DocIngredientItemGroup = "A REFERENCIA";

            // Arrange
            var homePage = LogInAsAdmin();


            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            //1. sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);

            //2. aller sur ""prod inflight""
            tabletAppPage.GotoTabletApp_ProdInFlight();
            Thread.Sleep(2000);

            int nbLigneSelected = tabletAppPage.Filter(TabletAppPage.FilterType.ItemGroups, DocIngredientItemGroup);
            tabletAppPage.Filter(TabletAppPage.FilterType.From, DateTime.Now.AddMonths(-2));
            Assert.IsTrue(nbLigneSelected > 0, "Tableau des services vide");


            tabletAppPage.Select("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-2", DocTitle);

            //2. Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");
            //3.choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            // le RecipeType est dans le Datasheet info.
            string xPathResult = "//*[contains(text(),'" + DocIngredientItemGroup + "')]";
            Assert.IsNotNull(tabletAppPage.WaitForElementIsVisible(By.XPath(xPathResult)), "Les données neremontent pas dans les documents haccp correspondent au item group sélectionné");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_ProdInflight_FilterWorkshop()
        {
            string DocSite = TestContext.Properties["SiteACE"].ToString();
            string DocTitle = "HACCP - Sanitization";
            string DocIngredient = "ItemForNormalProductOnly";
            string DocIngredientWorkshop = "Cocina Caliente";
            DateTime dateFrom = new DateTime(2024, 9, 11);

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            //1. sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);

            //2. aller sur ""prod inflight""
            tabletAppPage.GotoTabletApp_ProdInFlight();
            tabletAppPage.WaitPageLoading();

            tabletAppPage.Filter(TabletAppPage.FilterType.From, dateFrom);

            int nbLigneSelected = tabletAppPage.Filter(TabletAppPage.FilterType.Workshop, DocIngredientWorkshop);
            Assert.IsTrue(nbLigneSelected > 0, "Tableau des services vide");

            tabletAppPage.Select("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-2", DocTitle);

            //2. Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");
            //3.choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            tabletAppPage.WaitHACCPHorizontalProgressBar();

            //Assert
            // le Workshop est dans Item info.
            string xPathResult = "//*[contains(text(),'" + DocIngredient + "')]";
            var elementIsVisible = tabletAppPage.WaitForElementIsVisible(By.XPath(xPathResult));
            Assert.IsNotNull(elementIsVisible, "Les données ne remontent pas dans les documents haccp");
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_ProdInflight_SaveDocumentHACCPFromTemporary()
        {
            // Prepare
            string DocSite = "ACE";
            string DocTitle = "HACCP - Sanitization";
            string DocFileName = "TempTest_wrfr.user01";
            string DocReport = "HACCPSanitization";
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            //HACCP_Sanitization_9_20220301_015007.pdf
            string DocFileNamePdfBegin = "HACCP_Sanitization_9_";


            // Arrange
            var homePage = LogInAsAdmin();

            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Purge(downloadsPath, DocFileNamePdfBegin);

            //1. sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);

            //aller sur ""prod inflight""
            tabletAppPage.GotoTabletApp_ProdInFlight();
            Thread.Sleep(2000);

            // FROM TO mémorisé dans le Temporary (confirmé par Nicolas)
            //3.Sélectionne une date FROM TO
            tabletAppPage.Filter(TabletAppPage.FilterType.From, DateUtils.Now.Date);
            tabletAppPage.Filter(TabletAppPage.FilterType.To, DateUtils.Now.Date.AddDays(3));

            //1.Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.Select("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-2", DocTitle);

            //2. Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");
            //3.choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            Thread.Sleep(2000);

            //4. remplir une ligne de données
            tabletAppPage.addLineSanitizationProdInFlight("HAMBURGUESA TERNERA 110 GR CONGELADA ACE", "2.2", "80.0", "10", true, false, "MyComments", "MyPreparedBy");
            //5. Cliquer sur le bouton ""save "" => winrest propose un nom du temporary par défaut qui peut être modifié
            tabletAppPage.ClickButton("Save");
            tabletAppPage.FileFileNameProdInFlight1(DocFileName);
            //6.cliquer sur ""validate""
            tabletAppPage.ClickButton("Validate");
            try
            {
                // écraser le fichier
                tabletAppPage.ClickButton("Yes");
            }
            catch
            {
                // c'était un nouveau fichier
            }
            //une pop de validation apparait et puis cliquer sur OK
            tabletAppPage.ClickButton("Ok");

            tabletAppPage.Close();
            Thread.Sleep(2000);
            //Revenir vers la page d'acceuil de prod inflight , sélectionner le doc haccp
            homePage.Navigate();
            tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            // si je suis déjà passé par /tablet, ça affiche ACE en haut à gauche, et le reste de la page est blanche
            //2.aller sur ""prod inflight""
            tabletAppPage.GotoTabletApp_ProdInFlight();
            Thread.Sleep(2000);

            //4. Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.Select("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-2", DocTitle);
            Thread.Sleep(2000);
            //5. ouvrir un document temporaire dans le champs ""temporary""
            tabletAppPage.SelectAction("mat-select-6", DocFileName, true);

            //2.vérifier qu'on récupère bien les données enregistrées sont OK
            //Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");
            //choisir une date de production et cliquer sur print
            //tabletAppPage.ClickButton("Print");
            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();

            //6. Cliquer sur le bouton ""save and validate "" , une pop up s'ouvre de validation s'ouvre, puis cliquer sur ""validate"""
            tabletAppPage.ClickButton("Save and validate");
            tabletAppPage.ClickButton("Validate");
            //tabletAppPage.WaitHACCPHorizontalProgressBar();
            tabletAppPage.ClickButton("Ok");

            tabletAppPage.Close();
            Thread.Sleep(2000);
            //"Dans winrest dans la partie accouting => reporting :
            homePage.Navigate();
            homePage.ClearDownloads();

            ReportingPage reportingPage = homePage.GoToAccounting_Reporting();
            string DocTypeOfReport = reportingPage.DevPathDocTypeOfRecord();

            //1. Sélectionner type of report : HACCP
            //2. sélectionner le site sélectionné et sélectionner la production date
            reportingPage.FillReportingPage(DocTypeOfReport, DocSite);
            //3. recupérer le rapport HACCP
            var offset = reportingPage.TableHasDocument(DocReport);
            Assert.IsNotNull(offset, "Document non trouvé");
            reportingPage.PrintDownload(offset);

            //4.Les données pré remplies dans le document sont les mêmes dans le rapport
            //5. Vérifier dans le print...
            FileInfo fi = reportingPage.FindProdInFlightPDF(downloadsPath, DocFileNamePdfBegin);
            Assert.IsNotNull(fi, "pdf non trouvé");
            Assert.IsTrue(fi.Exists, "pdf non généré");

            reportingPage.FixPDFMagic(fi.FullName);
            PdfDocument document = PdfDocument.Open(fi.FullName);
            Page page1 = document.GetPage(1);
            IEnumerable<Word> words = page1.GetWords();
            // le nom du rapport
            var wordArray = words.ToArray();
            var title = wordArray[0].Text + wordArray[1].Text;
            Assert.AreEqual(1, (title.Contains("HACCPSANITIZATION") ? 1 : 0) + (title.Contains("HACCPSanitization") ? 1 : 0), "nom du rapport non trouvé dans " + fi.FullName);
            // la date de production, + FROM
            Assert.AreEqual(2, words.Count(w => w.Text == DateUtils.Now.Date.ToString("dd/MM/yyyy")), "date de production non trouvé dans " + fi.FullName);
            // le site,
            Assert.AreEqual(3, words.Count(w => w.Text == DocSite), "site non trouvé dans " + fi.FullName);
            // les dates from et to, (dont date de production égale à FROM)
            Assert.AreEqual(2, words.Count(w => w.Text == DateUtils.Now.Date.ToString("dd/MM/yyyy")), "date from non trouvé dans " + fi.FullName);
            Assert.AreEqual(1, words.Count(w => w.Text == DateUtils.Now.Date.AddDays(3).ToString("dd/MM/yyyy")), "date to non trouvé dans " + fi.FullName);
            // l'intitulé des colonnes
            Assert.AreEqual(3, words.Count(w => w.Text == "Group" || w.Text == "Item"), "colonne 'Group - Item' non trouvé dans " + fi.FullName);
            Assert.AreEqual(4, words.Count(w => w.Text == "Total" || w.Text == "Forecast" || w.Text == "Gross" || w.Text == "weight"), "colonne 'Total Forecast Gross weight' non trouvé dans " + fi.FullName);
            Assert.AreEqual(2, words.Count(w => w.Text == "Quantity"), "colonnes 'Quantity' non trouvé dans " + fi.FullName);
            Assert.AreEqual(3, words.Count(w => w.Text == "Desinfection" || w.Text == "Quantity"), "colonne 'Desinfection Quantity' non trouvé dans " + fi.FullName);
            Assert.AreEqual(2, words.Count(w => w.Text == "Disinfection" || w.Text == "Time"), "colonne 'Disinfection Time (min)' non trouvé dans " + fi.FullName);
            Assert.AreEqual(1, words.Count(w => w.Text == "Rinsing"), "colonne 'Rinsing' non trouvé dans " + fi.FullName);
            Assert.AreEqual(3, words.Count(w => w.Text == "Final" || w.Text == "visual" || w.Text == "inspection"), "colonne 'Final visual inspection' non trouvé dans " + fi.FullName);
            Assert.AreEqual(3, words.Count(w => w.Text == "Comments/" || w.Text == "corrective" || w.Text == "action"), "colonne 'Comments/ corrective action' non trouvé dans " + fi.FullName);
            Assert.AreEqual(2, words.Count(w => w.Text == "Prepared" || w.Text == "by"), "colonne 'Prepared by' non trouvé dans " + fi.FullName);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_ProdInflight_FilterShowNormalProductionOnly()
        {
            // Prepare
            string DocSite = "ACE";
            string DocTitle = "HACCP - Thawing";

            // Arrange
            var homePage = LogInAsAdmin();


            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            //1. sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);

            //2. aller sur ""prod inflight""
            tabletAppPage.GotoTabletApp_ProdInFlight();
            Thread.Sleep(2000);

            tabletAppPage.Filter(TabletAppPage.FilterType.From, DateTime.Now);
            tabletAppPage.Filter(TabletAppPage.FilterType.ShowNormalProductionOnly, null);

            tabletAppPage.Select("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-2", DocTitle);

            //2. Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");
            //3.choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            var listNameIngredient = tabletAppPage.GetAllNameIngredient();

            homePage.Navigate();
            var prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
            prodManagement.ResetFilter();

            prodManagement.Filter(FilterAndFavoritesPage.FilterType.DateFrom, DateTime.Now);
            var qtyAjustementPage = prodManagement.DoneToQtyAjustement();

            var resultPage = qtyAjustementPage.GoToResultPage();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.ShowNormalProd, true);
            resultPage.UnfoldAll();

            var items = resultPage.GetItemNames();

            foreach (var ingredient in listNameIngredient)
            {
                if (ingredient == "NO DATA")
                    Assert.IsTrue(items.Contains("No production"), "La liste contient l'item" + ingredient + "hors que dans tablet app No Data");
                else
                    Assert.IsTrue(items.Contains(ingredient), "La liste initiale ne contient pas l'item " + ingredient);
            }
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_ProdInflight_FilterShowVacuumProductionOnly()
        {
            // Prepare
            string DocSite = "ACE";
            string DocTitle = "HACCP - Thawing";
            DateTime dateFrom = new DateTime(2024, 9, 11);
            DateTime dateTo = new DateTime(2024, 9, 17);

            // Arrange
            var homePage = LogInAsAdmin();


            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            //1. sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);

            //2. aller sur ""prod inflight""
            tabletAppPage.GotoTabletApp_ProdInFlight();
            Thread.Sleep(2000);

            //3. Sélectionne une date FROM TO
            tabletAppPage.Filter(TabletAppPage.FilterType.From, dateFrom);
            tabletAppPage.Filter(TabletAppPage.FilterType.To, dateTo);

            tabletAppPage.Filter(TabletAppPage.FilterType.ShowVacuumProductionOnly, true);

            tabletAppPage.WaitPageLoading();
            tabletAppPage.Select("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-2", DocTitle);

            //2. Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");
            //3.choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();

            var listNameIngredient = tabletAppPage.GetAllNameIngredient();
            homePage.Navigate();
            ItemPage itemPage = homePage.GoToPurchasing_ItemPage();
            itemPage.ResetFilter();
            int nbPage = 2;
            foreach (var item in listNameIngredient)
            {
                itemPage.Filter(ItemPage.FilterType.Search, item);
                itemPage.Filter(ItemPage.FilterType.ShowAll, true);
                var itemGeneralInformationsPage = itemPage.ClickOnFirstItem();
                var itemUseCasePage = itemGeneralInformationsPage.ClickOnUseCasePage();
                var datasheetDetails = itemUseCasePage.EditFirstUseCase();

                itemUseCasePage.Go_To_New_Navigate(nbPage);
                nbPage++;

                var recipePage = datasheetDetails.EditRecipe();
                itemUseCasePage.ClickOnGeneralInformation();
                itemUseCasePage.WaitPageLoading();
                string cookingMode = itemUseCasePage.GetCookingModeSelected();
                itemUseCasePage.CloseEditUseCaseDatasheetForRecipe();

                homePage.Navigate();
                var applicationSettings = homePage.GoToApplicationSettings();
                var parametersProduction = applicationSettings.GoToParameters_ProductionPage();

                parametersProduction.GoToTab_CookingMode();
                var checkCookingModeIsVacuum = parametersProduction.isVacuumPackedCooking(cookingMode);

                Assert.IsTrue(checkCookingModeIsVacuum);

                homePage.Navigate();
                homePage.GoToPurchasing_ItemPage();
                itemPage.ResetFilter();
                itemPage.WaitPageLoading();
            }
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_ProdInflight_FilterShowNormalAndVacuumProduction()
        {
            // Prepare
            string DocSite = "ACE";
            string DocTitle = "HACCP - Thawing";
            DateTime dateFrom = new DateTime(2024, 9, 11);
            DateTime dateTo = new DateTime(2024, 9, 17);

            // Arrange
            var homePage = LogInAsAdmin();


            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            //1. sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);

            //2. aller sur ""prod inflight""
            tabletAppPage.GotoTabletApp_ProdInFlight();
            Thread.Sleep(2000);

            //3. Sélectionne une date FROM TO
            tabletAppPage.Filter(TabletAppPage.FilterType.From, dateFrom);
            tabletAppPage.Filter(TabletAppPage.FilterType.To, dateTo);

            tabletAppPage.Filter(TabletAppPage.FilterType.ShowNormalAndVacuumProduction, null);

            tabletAppPage.WaitPageLoading();
            tabletAppPage.Select("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-2", DocTitle);

            //4. Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");
            //5.choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            var listNameIngredient = tabletAppPage.GetAllNameIngredient();

            homePage.Navigate();
            var prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
            prodManagement.ResetFilter();

            prodManagement.Filter(FilterAndFavoritesPage.FilterType.DateFrom, dateFrom);
            prodManagement.Filter(FilterAndFavoritesPage.FilterType.DateTo, dateTo);
            var qtyAjustementPage = prodManagement.DoneToQtyAjustement();

            var resultPage = qtyAjustementPage.GoToResultPage();
            qtyAjustementPage.Filter(QuantityAdjustmentsPage.FilterType.ShowNormalAndVacuumProd, true);
            resultPage.UnfoldAll();

            var items = resultPage.GetItemNames();

            foreach (var ingredient in listNameIngredient)
            {
                Assert.IsTrue(items.Contains(ingredient), "La liste initiale ne contient pas l'item " + ingredient);
            }

        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_ProdInflight_FilterValidatedFlightsOnly()
        {
            // Prepare
            string DocSite = "ACE";
            string DocTitle = "HACCP - Tray Setup";
            //string DocService = "Service Test 5";
            string DocFlight = "XU715";
            DateTime dateFrom = new DateTime(2024, 9, 11);
            DateTime dateTo = new DateTime(2024, 9, 17);

            // Arrange
            var homePage = LogInAsAdmin();


            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            //1. sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);

            //2. aller sur ""prod inflight""
            tabletAppPage.GotoTabletApp_ProdInFlight();
            Thread.Sleep(2000);

            //3. Sélectionne une date FROM TO
            tabletAppPage.Filter(TabletAppPage.FilterType.From, dateFrom);
            tabletAppPage.Filter(TabletAppPage.FilterType.To, dateTo);

            tabletAppPage.Filter(TabletAppPage.FilterType.ValidatedFlightsOnly, true);
            string DocService = tabletAppPage.IsDev() ? "Service Test 5" : "Service for flight";
            tabletAppPage.Select("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-2", DocTitle);

            //2. Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");
            //3.choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            string xPathResult = "//*[contains(text(),'" + DocFlight + "')]";
            Assert.IsNotNull(tabletAppPage.WaitForElementExists(By.XPath(xPathResult)));
            string xPathResult2 = "//*[contains(text(),'" + DocService + "')]";
            Assert.IsNotNull(tabletAppPage.WaitForElementExists(By.XPath(xPathResult2)));
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_ProdInflight_Print()
        {
            // Prepare
            string DocSite = "ACE";
            string DocTitle = "HACCP - Thawing";
            string DocFileName = "HACCP_Thawing_" + DocSite + "_" + DateUtils.Now.Date.ToString("ddMMyyyy") + "_" + DateUtils.Now.Date.AddMonths(1).ToString("ddMMyyyy") + ".pdf";
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();




            // Arrange
            var homePage = LogInAsAdmin();


            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            tabletAppPage.Purge(downloadsPath, "HACCP_Thawing_" + DocSite + "_");

            //1. sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);

            //2.aller sur ""prod inflight""
            tabletAppPage.GotoTabletApp_ProdInFlight();
            Thread.Sleep(2000);

            //3.Sélectionne une date FROM TO
            tabletAppPage.Filter(TabletAppPage.FilterType.From, DateUtils.Now.Date);
            tabletAppPage.Filter(TabletAppPage.FilterType.To, DateUtils.Now.Date.AddMonths(1));


            //4.Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.Select("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-2", DocTitle);

            //5.Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");
            //6.choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            //tabletAppPage.WaitHACCPHorizontalProgressBar();

            //7.remplir une ligne de données
            string ExpDate = DateUtils.Now.Date.AddYears(2).ToString("dd/MM/yyyy");
            tabletAppPage.addLineThawingProdInFlight("CANAPES SURTIDOS PRESTIGE", "2", ExpDate, "15", DateUtils.Now.Date.AddMonths(2).ToString("dd/MM/yyyy"), DateUtils.Now.Date.AddYears(1).ToString("dd/MM/yyyy"), "MyComments", "MyPreparedBy");

            //8.Cliquer sur le bouton ""print """
            tabletAppPage.ClickButton("Print");

            FileInfo fi = new FileInfo(Path.Combine(downloadsPath, DocFileName));
            int counter = 0;
            while (!fi.Exists && counter < 20)
            {
                Thread.Sleep(2000);
                fi.Refresh();
                counter++;
            }
            Assert.IsTrue(fi.Exists, DocFileName + " non généré");

            //1. le rapport haccp est généré en PDF avec toutes les données pré remplies
            //2.Les données pré remplies dans le document sont les mêmes dans le rapport
            //3.Vérifier dans le print
            PdfDocument document = PdfDocument.Open(fi.FullName);
            IEnumerable<Word> allWords = Enumerable.Empty<Word>();
            for (int i = 1; i <= document.NumberOfPages; i++)
            {
                Page currentPage = document.GetPage(i);
                IEnumerable<Word> words = currentPage.GetWords();
                allWords = allWords.Concat(words);
            }


            // la date de production + FROM
            Assert.AreNotEqual(0, allWords.Count(w => w.Text == DateUtils.Now.Date.ToString("dd/MM/yyyy")), DateUtils.Now.Date.ToString("dd/MM/yyyy") + " non trouvé dans " + fi.FullName);
            // le site,
            Assert.AreNotEqual(0, allWords.Count(w => w.Text == DocSite), DocSite + " non trouvé dans " + fi.FullName);
            // les dates from et to ,
            Assert.AreNotEqual(0, allWords.Count(w => w.Text == DateUtils.Now.Date.ToString("dd/MM/yyyy")), DateUtils.Now.Date.ToString("dd/MM/yyyy") + "non trouvé dans " + fi.FullName);
            Assert.AreNotEqual(0, allWords.Count(w => w.Text == DateUtils.Now.Date.AddMonths(1).ToString("dd/MM/yyyy")), DateUtils.Now.Date.AddMonths(1).ToString("dd/MM/yyyy") + "non trouvé dans " + fi.FullName);
            // l'intitulé des colonnes
            Assert.AreNotEqual(0, allWords.Count(w => w.Text == "Group" || w.Text == "Item"), "'Group - Item' non trouvé dans " + fi.FullName);
            //Assert.AreEqual(4, words.Count(w => w.Text == "Total" || w.Text == "Forecast" || w.Text == "Gross" || w.Text == "weight"), "'Total Forecast Gross weight' non trouvé dans " + fi.FullName);
            Assert.AreNotEqual(0, allWords.Count(w => w.Text == "Real" || w.Text == "Quantity"), "'Real Quantity' non trouvé dans " + fi.FullName);
            Assert.AreNotEqual(0, allWords.Count(w => w.Text == "Exp." || w.Text == "Date"), "'Exp. Date' non trouvé dans " + fi.FullName);
            Assert.AreNotEqual(0, allWords.Count(w => w.Text == "Batch" || w.Text == "number" || w.Text == "if" || w.Text == "available"), "'Batch number if available' non trouvé dans " + fi.FullName);
            Assert.AreNotEqual(0, allWords.Count(w => w.Text == "Thawing" || w.Text == "starting" || w.Text == "date"), "'Thawing starting date' non trouvé dans " + fi.FullName);
            Assert.AreNotEqual(0, allWords.Count(w => w.Text == "Maximum" || w.Text == "End" || w.Text == "using"), "'Maximum End using' non trouvé dans " + fi.FullName);
            Assert.AreNotEqual(0, allWords.Count(w => w.Text == "Comments/" || w.Text == "corrective" || w.Text == "action"), "'Comments/ corrective action' non trouvé dans " + fi.FullName);
            Assert.AreNotEqual(0, allWords.Count(w => w.Text == "Prepared" || w.Text == "By"), "'Prepared By' non trouvé dans " + fi.FullName);

            // Les données
            Assert.AreNotEqual(0, allWords.Count(w => w.Text == "CANAPES" || w.Text == "SURTIDOS" || w.Text == "PRESTIGE"), "CANAPES SURTIDOS PRESTIGE non trouvé dans " + fi.FullName);

            if (tabletAppPage.IsDev())
            {
                Assert.AreNotEqual(0, allWords.Count(w => w.Text == "2"), "2 non trouvé dans " + fi.FullName);
                Assert.AreNotEqual(0, allWords.Count(w => w.Text == DateUtils.Now.Date.AddYears(2).ToString("dd/MM/yyyy")), DateUtils.Now.Date.AddYears(2).ToString("dd/MM/yyyy") + "non trouvé dans " + fi.FullName);
                Assert.AreNotEqual(0, allWords.Count(w => w.Text == "15"), "15 non trouvé dans " + fi.FullName);
                Assert.AreNotEqual(0, allWords.Count(w => w.Text == DateUtils.Now.Date.AddMonths(2).ToString("dd/MM/yyyy")), DateUtils.Now.Date.ToString("dd/MM/yyyy") + "non trouvé dans " + fi.FullName);
                Assert.AreNotEqual(0, allWords.Count(w => w.Text == DateUtils.Now.Date.AddYears(1).ToString("dd/MM/yyyy")), DateUtils.Now.Date.AddYears(1).ToString("dd/MM/yyyy") + "non trouvé dans " + fi.FullName);
                Assert.AreNotEqual(0, allWords.Count(w => w.Text == "MyComments"), "MyComments non trouvé dans " + fi.FullName);
                Assert.AreNotEqual(0, allWords.Count(w => w.Text == "MyPreparedBy"), "MyPreparedBy non trouvé dans " + fi.FullName);
            }
            else
            {
                Assert.AreNotEqual(0, allWords.Count(w => w.Text == "1"), "1 non trouvé dans " + fi.FullName);
                Assert.AreNotEqual(0, allWords.Count(w => w.Text == "2"), "2 non trouvé dans " + fi.FullName);
                Assert.AreNotEqual(0, allWords.Count(w => w.Text == DateUtils.Now.Date.AddYears(2).ToString("dd/MM/yyyy")), DateUtils.Now.Date.AddYears(2).ToString("dd/MM/yyyy") + "non trouvé dans " + fi.FullName);
                Assert.AreNotEqual(0, allWords.Count(w => w.Text == "15"), "15 non trouvé dans " + fi.FullName);
                Assert.AreNotEqual(0, allWords.Count(w => w.Text == DateUtils.Now.Date.AddMonths(2).ToString("dd/MM/yyyy")), DateUtils.Now.Date.ToString("dd/MM/yyyy") + "non trouvé dans " + fi.FullName);
                Assert.AreNotEqual(0, allWords.Count(w => w.Text == DateUtils.Now.Date.AddYears(1).ToString("dd/MM/yyyy")), DateUtils.Now.Date.AddYears(1).ToString("dd/MM/yyyy") + "non trouvé dans " + fi.FullName);
                Assert.AreNotEqual(0, allWords.Count(w => w.Text == "MyComments"), "MyComments non trouvé dans " + fi.FullName);
                Assert.AreNotEqual(0, allWords.Count(w => w.Text == "MyPreparedBy"), "MyPreparedBy non trouvé dans " + fi.FullName);
            }
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_ProdInflight_PrintTemporaryDocument()
        {
            // Prepare
            string DocSite = "ACE";
            string DocTitle = "HACCP - Thawing";
            string DocFileName = "HACCP_Thawing_" + DocSite + "_" + DateUtils.Now.Date.ToString("ddMMyyyy") + "_" + DateUtils.Now.Date.ToString("ddMMyyyy") + ".pdf";
            string DocTmpName = "Temp2.user";
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            IEnumerable<Word> words2 = new List<Word>();


            // Arrange
            var homePage = LogInAsAdmin();


            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            tabletAppPage.Purge(downloadsPath, "HACCP_Thawing_" + DocSite + "_");

            //sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);

            //aller sur ""prod inflight""
            tabletAppPage.GotoTabletApp_ProdInFlight();
            Thread.Sleep(2000);

            //Sélectionne une date FROM TO
            tabletAppPage.Filter(TabletAppPage.FilterType.From, DateUtils.Now.Date);

            //Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.Select("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-2", DocTitle);

            //5.Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");
            //6.choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            //tabletAppPage.WaitHACCPHorizontalProgressBar();
            var listNameIngredient = tabletAppPage.GetAllNameIngredient();

            homePage.Navigate();
            var prodManagement = homePage.GoToProduction_ProductionManagemenentPage();
            prodManagement.ResetFilter();

            prodManagement.Filter(FilterAndFavoritesPage.FilterType.DateFrom, DateUtils.Now.Date);
            QuantityAdjustmentsPage qtyAjustementPage = prodManagement.DoneToQtyAjustement();

            ResultPage resultPage = qtyAjustementPage.GoToResultPage();
            resultPage.UnfoldAll();

            var items = resultPage.GetItemNames();

            foreach (var ingredient in listNameIngredient)
            {
                if (ingredient == "NO DATA")
                    Assert.IsTrue(items.Contains("No production"), "La liste contient l'item" + ingredient + "hors que dans tablet app No Data");
                else
                    Assert.IsTrue(items.Contains(ingredient), "La liste initiale ne contient pas l'item " + ingredient);
            }

            homePage.Navigate();
            tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Purge(downloadsPath, "HACCP_Thawing_" + DocSite + "_");
            tabletAppPage.Select("mat-select-0", DocSite);
            tabletAppPage.GotoTabletApp_ProdInFlight();
            Thread.Sleep(2000);
            tabletAppPage.Filter(TabletAppPage.FilterType.From, DateUtils.Now.Date);
            tabletAppPage.Select("mat-select-2", DocTitle);

            tabletAppPage.SelectAction("mat-select-2", DocTitle);
            tabletAppPage.ClickBaseButton("NEXT STEP");
            tabletAppPage.ClickButton("Print");
            tabletAppPage.GotoTabletApp_DocumentTab();

            //7.remplir une ligne de données
            string ExpDate = DateUtils.Now.Date.AddYears(2).ToString("dd/MM/yyyy");
            int offset = 0;

            for (int i = 0; i < listNameIngredient.Count; i++)
            {
                string ingredientName = listNameIngredient[i];
                tabletAppPage.addLineThawingProdInFlight(ingredientName, "3", ExpDate, "16", DateUtils.Now.Date.ToString("dd/MM/yyyy"), DateUtils.Now.Date.AddYears(1).ToString("dd/MM/yyyy"), "MyComments", "MyPreparedBy", offset);
                offset += 14;
            }

            tabletAppPage.ClickButton("Save");

            tabletAppPage.FileFileNameProdInFlight0(DocTmpName);
            tabletAppPage.ClickButton("Validate");
            try
            {
                tabletAppPage.ClickButton("Yes");
            }
            catch
            {
                // nouveau fichier
            }
            tabletAppPage.ClickButton("Ok");

            tabletAppPage.Close();
            Thread.Sleep(2000);
            //8.Revenir vers la page d'acceuil de prod inflight , sélectionner le doc haccp
            homePage.Navigate();
            tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            // si je suis déjà passé par /tablet, ça affiche ACE en haut à gauche, et le reste de la page est blanche
            tabletAppPage.GotoTabletApp_ProdInFlight();
            Thread.Sleep(2000);

            tabletAppPage.Select("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-6", DocTmpName, true);

            //5.Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            //tabletAppPage.WaitHACCPHorizontalProgressBar();

            //8.Cliquer sur le bouton ""print """
            tabletAppPage.ClickButton("Print");
            //tabletAppPage.WaitHACCPHorizontalProgressBar();
            FileInfo fi = new FileInfo(Path.Combine(downloadsPath, DocFileName));
            int counter = 0;
            while (!fi.Exists && counter < 10)
            {
                Thread.Sleep(2000);
                fi.Refresh();
                counter++;
            }
            Assert.IsTrue(fi.Exists, DocFileName + " non généré");
            //1. le rapport haccp est généré en PDF avec toutes les données pré remplies
            //2.Les données pré remplies dans le document sont les mêmes dans le rapport
            //3.Vérifier dans le print le nom du rapport, la date de production, le site, les dates from et to , l'intitulé des colonnes
            PdfDocument document = PdfDocument.Open(fi.FullName);
            IEnumerable<Word> allWords = Enumerable.Empty<Word>();
            for (int i = 1; i <= document.NumberOfPages; i++)
            {
                Page currentPage = document.GetPage(i);
                IEnumerable<Word> words = currentPage.GetWords();
                allWords = allWords.Concat(words);
            }

            // le site,
            Assert.AreNotEqual(0, allWords.Count(w => w.Text == DocSite), DocSite + " non trouvé dans " + fi.FullName);
            // les dates from et to
            Assert.AreNotEqual(0, allWords.Count(w => w.Text == DateUtils.Now.Date.ToString("dd/MM/yyyy")), DateUtils.Now.Date.ToString("dd/MM/yyyy") + " non trouvé dans " + fi.FullName);
            Assert.AreNotEqual(0, allWords.Count(w => w.Text == DateUtils.Now.Date.ToString("dd/MM/yyyy")), DateUtils.Now.Date.AddYears(1).ToString("dd/MM/yyyy") + " non trouvé dans " + fi.FullName);
            // l'intitulé des colonnes
            Assert.AreNotEqual(0, allWords.Count(w => w.Text == "Group" || w.Text == "Item"), "'Group - Item' non trouvé dans " + fi.FullName);
            //Assert.AreEqual(4, words.Count(w => w.Text == "Total" || w.Text == "Forecast" || w.Text == "Gross" || w.Text == "weight"), "'Total Forecast Gross weight' non trouvé dans " + fi.FullName);
            Assert.AreNotEqual(0, allWords.Count(w => w.Text == "Real" || w.Text == "Quantity"), "'Real Quantity' non trouvé dans " + fi.FullName);
            Assert.AreNotEqual(0, allWords.Count(w => w.Text == "Exp." || w.Text == "Date"), "'Exp. Date' non trouvé dans " + fi.FullName);
            Assert.AreNotEqual(0, allWords.Count(w => w.Text == "Batch" || w.Text == "number" || w.Text == "if" || w.Text == "available"), "'Batch number if available' non trouvé dans " + fi.FullName);
            Assert.AreNotEqual(0, allWords.Count(w => w.Text == "Thawing" || w.Text == "starting" || w.Text == "date"), "'Thawing starting date' non trouvé dans " + fi.FullName);
            Assert.AreNotEqual(0, allWords.Count(w => w.Text == "Maximum" || w.Text == "End" || w.Text == "using"), "'Maximum End using' non trouvé dans " + fi.FullName);
            Assert.AreNotEqual(0, allWords.Count(w => w.Text == "Comments/" || w.Text == "corrective" || w.Text == "action"), "'Comments/ corrective action' non trouvé dans " + fi.FullName);
            Assert.AreNotEqual(0, allWords.Count(w => w.Text == "Prepared" || w.Text == "By"), "'Prepared By' non trouvé dans " + fi.FullName);
            // Les données
            Assert.AreNotEqual(0, allWords.Count(w => w.Text == "HAMBURGUESA" || w.Text == "TERNERA" || w.Text == "110" || w.Text == "GR" || w.Text == "CONGELADA" || w.Text == "ACE"), "HAMBURGUESA TERNERA 110 GR CONGELADA ACE non trouvé dans " + fi.FullName);
            Assert.AreNotEqual(0, allWords.Count(w => w.Text == "3"), "3 non trouvé dans " + fi.FullName);
            Assert.AreNotEqual(0, allWords.Count(w => w.Text == ExpDate), ExpDate + " non trouvé dans " + fi.FullName);
            Assert.AreNotEqual(0, allWords.Count(w => w.Text == "16"), "16 non trouvé dans " + fi.FullName);
            Assert.AreNotEqual(0, allWords.Count(w => w.Text == DateUtils.Now.Date.ToString("dd/MM/yyyy")), DateUtils.Now.Date.ToString("dd/MM/yyyy") + "non trouvé dans " + fi.FullName);
            Assert.AreNotEqual(0, allWords.Count(w => w.Text == DateUtils.Now.Date.AddYears(1).ToString("dd/MM/yyyy")), DateUtils.Now.Date.AddYears(1).ToString("dd/MM/yyyy") + "non trouvé dans " + fi.FullName);
            Assert.AreNotEqual(0, allWords.Count(w => w.Text == "MyComments"), "MyComments non trouvé dans " + fi.FullName);
            Assert.AreNotEqual(0, allWords.Count(w => w.Text == "MyPreparedBy"), "MyPreparedBy non trouvé dans " + fi.FullName);

        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_ProdInflight_FilterFromTo()
        {
            // Prepare
            string DocSite = "ACE";
            string DocService = "Service Test 5";

            // Arrange
            var homePage = LogInAsAdmin();


            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            //1. sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);

            //2. aller sur ""prod inflight""
            tabletAppPage.GotoTabletApp_ProdInFlight();
            Thread.Sleep(2000);

            //Sélectionner une date From et une date TO
            //Les données qui remontent dans les documents haccp correspondent aux services prévus dans cet intervalle de temps
            //check Service demain
            tabletAppPage.Filter(TabletAppPage.FilterType.From, DateUtils.Now.Date.AddDays(10));
            tabletAppPage.Filter(TabletAppPage.FilterType.To, DateUtils.Now.Date.AddDays(20));
            Assert.AreEqual(0, tabletAppPage.Filter(TabletAppPage.FilterType.Services, DocService));

            homePage.Navigate();
            tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            tabletAppPage.GotoTabletApp_ProdInFlight();
            Thread.Sleep(2000);
            //check Service aujourd'hui
            tabletAppPage.Filter(TabletAppPage.FilterType.From, DateUtils.Now.Date.AddDays(-10));
            tabletAppPage.Filter(TabletAppPage.FilterType.To, DateUtils.Now.Date.AddDays(20));
            Assert.AreEqual(1, tabletAppPage.Filter(TabletAppPage.FilterType.Services, DocService));
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_ProdInflight_FilterFavorite()
        {
            // Prepare
            string DocSite = "ACE";
            string DocFavorite = "TestFavoriteConfig";
            string DocTitle = "HACCP - Sanitization";
            // Arrange
            var homePage = LogInAsAdmin();

            //Avoir enregistré un filtre favori dans winrest dans la partie production -> production management par rapport au :
            //site, workshop, customer, guest type, service categories, services, recipe type, item group "
            FilterAndFavoritesPage favoritePage = homePage.GoToProduction_ProductionManagemenentPage();
            favoritePage.DeleteFavorite(DocFavorite);

            //site
            favoritePage.Filter(FilterAndFavoritesPage.FilterType.Site, "MAD");
            favoritePage.Filter(FilterAndFavoritesPage.FilterType.Site, "ACE");
            //workshop
            string WorkShop = "Ensamblaje";
            favoritePage.Filter(FilterAndFavoritesPage.FilterType.Workshops, WorkShop);
            //customer
            string CustomerShort = "CAT Genérico";
            favoritePage.Filter(FilterAndFavoritesPage.FilterType.Customers, CustomerShort);
            //guest type
            string GuestType = "BOB";
            favoritePage.Filter(FilterAndFavoritesPage.FilterType.GuestType, GuestType);
            //service categories
            string Category = "BEBIDAS";
            favoritePage.Filter(FilterAndFavoritesPage.FilterType.ServicesCategorie, Category);
            //services
            string Service = favoritePage.IsDev() ? "Service Test 5" : "Service for flight";
            favoritePage.Filter(FilterAndFavoritesPage.FilterType.Service, Service);
            //recipe type
            string RecipeType = "ENSALADA CREW";
            favoritePage.Filter(FilterAndFavoritesPage.FilterType.RecipeType, RecipeType);
            //item group
            string ItemGroup = "FRUTA Y VERDURA FRESCA";
            favoritePage.Filter(FilterAndFavoritesPage.FilterType.ItemGroups, ItemGroup);

            //1. avoir enregsitré une combinaison de filtre en tant que favori dans winrest dans la partie production-> production management et donner un nom à ce filtre
            favoritePage.MakeFavorite(DocFavorite);
            homePage.Navigate();

            //2. dans la tablette, choisir le site et choisir dans le champs ""favorite"" la combinaison de filtre sauvegardée
            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            //sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);
            //aller sur ""prod inflight""
            tabletAppPage.GotoTabletApp_ProdInFlight();
            Thread.Sleep(2000);
            tabletAppPage.Select("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-2", DocTitle);

            //3.et vérifier que la combinaison sauvegardée est correcte par rapport au workshop, customer, guest type, service categories, services, recipe type, item group
            //3.vérifier que les données qui remontent dans les docs haccp correspondent bien à cette combinaison de filtre"
            //1.En cliquant sur le filtre favori, on devrait retrouver la même combinaison sauvergardée dans production -> production management
            tabletAppPage.SelectAction("mat-select-4", DocFavorite, true);

            //3. Quand on sélectionne ce filtre, on doit retrouver la même combinaison sauvegardée
            //workshop
            Assert.AreEqual(1, tabletAppPage.FilterCounter("Workshops"));
            //customer
            Assert.AreEqual(1, tabletAppPage.FilterCounter("Customers"));
            //guest type
            Assert.AreEqual(1, tabletAppPage.FilterCounter("GuestTypes"));
            //service categories
            Assert.AreEqual(1, tabletAppPage.FilterCounter("ServiceCategories"));
            //services
            Assert.AreEqual(1, tabletAppPage.FilterCounter("Services"));
            //recipe type
            Assert.AreEqual(1, tabletAppPage.FilterCounter("RecipeTypes"));
            //item group
            Assert.AreEqual(1, tabletAppPage.FilterCounter("ItemGroups"));

            //2.Les données qui remontent dans les docs haccp doivent correspondre à cette combinaison de filtre en terme de:
            //workshop, customer, guest type, service categories, services, recipe type, item group "
            tabletAppPage.ClickBaseButton("NEXT STEP");
            //6. choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            tabletAppPage.WaitHACCPHorizontalProgressBar();
            tabletAppPage.Close();
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_ProdInflight_CheckBoxCNC()
        {
            // Prepare
            string DocSite = "ACE";
            string DocTitle = "HACCP - Tray Setup";
            string DocTmpName = "Tmp.user";
            string DocFileName = "HACCP_Tray_Setup_" + DocSite + "_" + DateUtils.Now.Date.ToString("ddMMyyyy") + "_" + DateUtils.Now.Date.AddDays(10).ToString("ddMMyyyy") + ".pdf";
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            DateTime dateFrom = new DateTime(2024, 9, 11);
            DateTime dateTo = new DateTime(2024, 9, 17);

            // Arrange
            var homePage = LogInAsAdmin();


            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            tabletAppPage.Purge(downloadsPath, "HACCP_Tray_Setup_" + DocSite + "_");

            //1. sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);
            //2. aller sur ""prod inflight""
            tabletAppPage.GotoTabletApp_ProdInFlight();
            Thread.Sleep(2000);
            //3. Sélectionne une date FROM TO
            tabletAppPage.Filter(TabletAppPage.FilterType.From, DateUtils.Now.Date);
            tabletAppPage.Filter(TabletAppPage.FilterType.To, DateUtils.Now.Date.AddDays(10));
            //4. Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.Select("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-2", DocTitle);

            //5. Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");
            //6. choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");
            tabletAppPage.WaitHACCPHorizontalProgressBar();

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            //7. Dans la colonne où la checkbox est paramétré, cocher une case
            tabletAppPage.addLineTraySetup(true, true);

            //8. cliquer sur save pour enregistrer un document temporary
            tabletAppPage.ClickButton("Save");
            tabletAppPage.FileFileNameProdInFlight0(DocTmpName);
            //cliquer sur ""validate""
            tabletAppPage.ClickButton("Validate");
            try
            {
                // écraser le fichier
                tabletAppPage.ClickButton("Yes");
            }
            catch
            {
                // nouveau fichier
            }
            //7.une pop de validation apparait et puis cliquer sur OK
            tabletAppPage.ClickButton("Ok");

            tabletAppPage.Close();
            //9. récupérer ensuite le document temporary dans la page d'acceuil de prod inflight dans la section "" temporary""
            homePage.Navigate();
            tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            // si je suis déjà passé par /tablet, ça affiche ACE en haut à gauche, et le reste de la page est blanche
            tabletAppPage.GotoTabletApp_ProdInFlight();
            tabletAppPage.Select("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-6", DocTmpName, true);

            //5. Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");
            //6. choisir une date de production et cliquer sur print
            //tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();


            //10.vérifier que la case coché dans l'enregistrement reste bien cochée sur le document temporary
            tabletAppPage.EditLineTraySetup();
            Assert.IsTrue(tabletAppPage.CheckValueBox(1), "CNC non checked");
            //11. puis cliquer sur print pour générer le rapport haccp en pdf
            tabletAppPage.ClickButton("Print");
            //tabletAppPage.WaitHACCPHorizontalProgressBar();
            //12. vérifier que la case cochée renvoie bien l'information ""C"" et que les autres cases non cochés renvoient l'information ""NA""
            //1. la case cochée dans l'enregistrement reste bien coché sur le document temporary
            //2.Dans le print, la case cochée doit renvoyer l'information ""C"" et les autres cases non cochées renvoient l'information ""NA""
            FileInfo fi = new FileInfo(downloadsPath + "\\" + DocFileName);
            int counter = 0;
            while (!fi.Exists && counter < 20)
            {
                Thread.Sleep(1000);
                fi.Refresh();
                counter++;
            }
            Assert.IsTrue(fi.Exists, DocFileName + " non généré cas 1");

            PdfDocument document = PdfDocument.Open(fi.FullName);
            Page page1 = document.GetPage(1);
            IEnumerable<Word> words = page1.GetWords();
            Assert.AreEqual(2, words.Count(w => w.Text == "C"), "C non trouvé dans " + fi.FullName);
            Assert.AreEqual(0, words.Count(w => w.Text == "NC"), "NC trouvé dans " + fi.FullName);
            Assert.AreEqual(0, words.Count(w => w.Text == "NA"), "NA trouvé dans " + fi.FullName);


            tabletAppPage.Close();

            homePage.Navigate();
            tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            tabletAppPage.GotoTabletApp_ProdInFlight();

            tabletAppPage.Select("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-2", DocTitle);

            tabletAppPage.ClickBaseButton("NEXT STEP");
            //6. choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            //7. Dans la colonne où la checkbox est paramétré, cocher une case
            tabletAppPage.addLineTraySetup(false, true);
            Assert.IsFalse(tabletAppPage.CheckValueBox(1), "CNC checked");
            tabletAppPage.ClickButton("Print");
            //tabletAppPage.WaitHACCPHorizontalProgressBar();

            fi = new FileInfo(downloadsPath + "\\" + DocFileName.Replace(".", " (1)."));
            counter = 0;
            while (!fi.Exists && counter < 20)
            {
                Thread.Sleep(1000);
                fi.Refresh();
                counter++;
            }
            Assert.IsTrue(fi.Exists, DocFileName + " non généré cas 2");

            document = PdfDocument.Open(fi.FullName);
            page1 = document.GetPage(1);
            words = page1.GetWords();
            Assert.AreEqual(1, words.Count(w => w.Text == "C"), "C non trouvé dans " + fi.FullName);
            Assert.AreEqual(1, words.Count(w => w.Text == "NC"), "NC trouvé dans " + fi.FullName);
            Assert.AreEqual(0, words.Count(w => w.Text == "NA"), "NA trouvé dans " + fi.FullName);
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_ProdInfight_CheckBoxCNA()
        {
            // Prepare
            string DocSite = "ACE";
            string DocTitle = "HACCP - Tray Setup";
            string DocTmpName = "Tmp.user";
            string DocFileName = "HACCP_Tray_Setup_" + DocSite + "_" + DateUtils.Now.Date.AddDays(-6).ToString("ddMMyyyy") + "_" + DateUtils.Now.Date.ToString("ddMMyyyy") + ".pdf";
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            tabletAppPage.Purge(downloadsPath, "HACCP_Tray_Setup_" + DocSite + "_");

            //1. sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);
            //2. aller sur ""prod inflight""
            tabletAppPage.GotoTabletApp_ProdInFlight();
            Thread.Sleep(2000);
            //3. Sélectionne une date FROM TO
            tabletAppPage.Filter(TabletAppPage.FilterType.From, DateUtils.Now.Date.AddDays(-6));
            tabletAppPage.Filter(TabletAppPage.FilterType.To, DateUtils.Now.Date);
            //4. Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.Select("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-2", DocTitle);

            //5. Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");
            //6. choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");
            tabletAppPage.WaitHACCPHorizontalProgressBar();

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            //7. Dans la colonne où la checkbox est paramétré, cocher une case
            tabletAppPage.addLineTraySetup(true, true);

            //8. cliquer sur save pour enregistrer un document temporary
            tabletAppPage.ClickButton("Save");
            tabletAppPage.FileFileNameProdInFlight0(DocTmpName);
            //cliquer sur ""validate""
            tabletAppPage.ClickButton("Validate");
            try
            {
                // écraser le fichier
                tabletAppPage.ClickButton("Yes");
            }
            catch
            {
                // nouveau fichier
            }
            //7.une pop de validation apparait et puis cliquer sur OK
            tabletAppPage.ClickButton("Ok");

            tabletAppPage.Close();

            //9. récupérer ensuite le document temporary dans la page d'acceuil de prod inflight dans la section "" temporary""
            homePage.Navigate();
            // si je suis déjà passé par /tablet, ça affiche ACE en haut à gauche, et le reste de la page est blanche
            tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            tabletAppPage.GotoTabletApp_ProdInFlight();
            tabletAppPage.Select("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-6", DocTmpName, true);

            //5. Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");
            //6. choisir une date de production et cliquer sur print
            //tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();


            //10.vérifier que la case coché dans l'enregistrement reste bien cochée sur le document temporary
            tabletAppPage.EditLineTraySetup();
            Assert.IsTrue(tabletAppPage.CheckValueBox(2), "CNA not checked");
            //11. puis cliquer sur print pour générer le rapport haccp en pdf
            tabletAppPage.ClickButton("Print");
            tabletAppPage.WaitHACCPHorizontalProgressBar();
            //12. vérifier que la case cochée renvoie bien l'information ""C"" et que les autres cases non cochés renvoient l'information ""NA""
            //1. la case cochée dans l'enregistrement reste bien coché sur le document temporary
            //2.Dans le print, la case cochée doit renvoyer l'information ""C"" et les autres cases non cochées renvoient l'information ""NA""
            FileInfo fi = new FileInfo(downloadsPath + "\\" + DocFileName);
            int counter = 0;
            while (!fi.Exists && counter < 20)
            {
                Thread.Sleep(1000);
                fi.Refresh();
                counter++;
            }
            Assert.IsTrue(fi.Exists, DocFileName + " non généré cas 1");

            PdfDocument document = PdfDocument.Open(fi.FullName);
            List<string> mots = new List<string>();
            foreach (Page page in document.GetPages())
            {
                IEnumerable<Word> words = page.GetWords();
                foreach (Word word in words)
                {
                    mots.Add(word.Text);
                }
            }
            Assert.AreEqual(2, mots.Count(w => w == "C"), "C non trouvé dans " + fi.FullName);
            if (document.NumberOfPages == 1)
            {
                Assert.AreEqual(0, mots.Count(w => w == "NC"), "NC trouvé dans " + fi.FullName);
                Assert.AreEqual(0, mots.Count(w => w == "NA"), "NA trouvé dans " + fi.FullName);
            }
            else
            {
                Assert.AreNotEqual(0, mots.Count(w => w == "NC"), "NC trouvé dans " + fi.FullName);
                Assert.AreNotEqual(0, mots.Count(w => w == "NA"), "NA trouvé dans " + fi.FullName);
            }

            tabletAppPage.Close();

            homePage.Navigate();
            tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            tabletAppPage.GotoTabletApp_ProdInFlight();

            tabletAppPage.Select("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-2", DocTitle);

            tabletAppPage.ClickBaseButton("NEXT STEP");
            //6. choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            //7. Dans la colonne où la checkbox est paramétré, cocher une case
            tabletAppPage.addLineTraySetup(true, false);
            Assert.IsFalse(tabletAppPage.CheckValueBox(2), "CNA checked");
            tabletAppPage.ClickButton("Print");
            tabletAppPage.WaitHACCPHorizontalProgressBar();

            fi = new FileInfo(downloadsPath + "\\" + DocFileName.Replace(".", " (1)."));
            counter = 0;
            while (!fi.Exists && counter < 20)
            {
                Thread.Sleep(1000);
                fi.Refresh();
                counter++;
            }
            Assert.IsTrue(fi.Exists, DocFileName + " non généré cas 2");

            document = PdfDocument.Open(fi.FullName);
            foreach (Page page in document.GetPages())
            {
                IEnumerable<Word> words = page.GetWords();
                foreach (Word word in words)
                {
                    mots.Add(word.Text);
                }
            }
            Assert.AreEqual(3, mots.Count(w => w == "C"), "C non trouvé dans " + fi.FullName);
            if (document.NumberOfPages == 1)
            {
                Assert.AreEqual(0, mots.Count(w => w == "NC"), "NC trouvé dans " + fi.FullName);
                Assert.AreEqual(1, mots.Count(w => w == "NA"), "NA trouvé dans " + fi.FullName);
            }
            else
            {
                Assert.AreNotEqual(0, mots.Count(w => w == "NC"), "NC trouvé dans " + fi.FullName);
                Assert.AreNotEqual(0, mots.Count(w => w == "NA"), "NA trouvé dans " + fi.FullName);
            }
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_ProdInflight_SaveDocumentCheckMultiDate()
        {
            // Prepare
            string DocSite = "ACE";
            string DocTitle = "HACCP - Tray Setup";
            string DocReport = "HACCPTraySetup";
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNamePdfBegin = "HACCP_Tray_Setup_9_";

            // Arrange
            var homePage = LogInAsAdmin();


            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Purge(downloadsPath, DocFileNamePdfBegin);

            //1. sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);
            //2. aller sur ""prod inflight""
            tabletAppPage.GotoTabletApp_ProdInFlight();
            Thread.Sleep(2000);
            //3. Sélectionne une date FROM TO
            tabletAppPage.Filter(TabletAppPage.FilterType.From, DateUtils.Now.Date);
            tabletAppPage.Filter(TabletAppPage.FilterType.To, DateUtils.Now.Date.AddMonths(1));
            //4. Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.Select("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-2", DocTitle);

            //5. Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");
            //6. choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();

            //7. renseigner plusieurs dates dans la colonne où le format multi date est paramétré
            tabletAppPage.addLineTraySetup(true, false);
            Assert.IsTrue(tabletAppPage.CheckMultiDate(), "Probleme MultiDate");

            //8.puis cliquer sur ""save and validate"" pour enregistrer le document définitivement
            tabletAppPage.ClickButton("Save and validate");
            tabletAppPage.ClickButton("Validate");
            tabletAppPage.ClickButton("Ok");

            tabletAppPage.Close();
            Thread.Sleep(2000);
            //9. aller sur winrest dans la partie accounting=> reporting pour récupérer le rapport pdf généré ( sélectionner le type de rapport HACCP , puis sélectionner le même site et la même date )"
            // verify Accouting>Reporting
            homePage.Navigate();
            homePage.ClearDownloads();

            ReportingPage reportingPage = homePage.GoToAccounting_Reporting();
            string DocTypeOfReport = reportingPage.DevPathDocTypeOfRecord();

            //Sélectionner type of report: HACCP
            //sélectionner le site sélectionné et sélectionner la production date
            reportingPage.FillReportingPage(DocTypeOfReport, DocSite);
            //recupérer le rapport HACCP
            var offset = reportingPage.TableHasDocument(DocReport);
            Assert.IsNotNull(offset, "Document non trouvé");
            reportingPage.PrintDownload(offset);
            //Les date doivent être conformes dans le rapport généré dans le bon format également défini dans les global settings (settings=> global settings=> DateFormatDisplay)
            FileInfo fi = reportingPage.FindProdInFlightPDF(downloadsPath, DocFileNamePdfBegin);
            Assert.IsNotNull(fi);

            reportingPage.FixPDFMagic(fi.FullName);
            PdfDocument document = PdfDocument.Open(fi.FullName);

            List<string> mots = new List<string>();
            foreach (Page page in document.GetPages())
            {
                IEnumerable<Word> words = page.GetWords();
                foreach (Word word in words)
                {
                    mots.Add(word.Text);
                }
            }
            string yearAndMonth = DateUtils.Now.ToString("/MM/yyyy");
            for (int i = 10; i < 17; i++)
            {
                Assert.IsTrue(mots.Any(w => w == (i + yearAndMonth)), (i + yearAndMonth) + " MultiDate non trouvé dans " + fi.FullName);
            }
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_ProdInflight_SaveDocumentTemporaryCheckMultiDate()
        {
            // Prepare
            string DocSite = "ACE";
            string DocTitle = "HACCP - Tray Setup";
            string DocTmpName = "Tmp.user";
            string DocFileName = "HACCP_Tray_Setup_" + DocSite + "_" + DateUtils.Now.Date.ToString("ddMMyyyy") + "_" + DateUtils.Now.Date.AddMonths(1).ToString("ddMMyyyy") + ".pdf";
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();


            // Arrange
            LogInAsAdmin();
            var homePage = LogInAsAdmin();

            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            tabletAppPage.Purge(downloadsPath, "HACCP_Tray_Setup_" + DocSite + "_");

            //1. sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);
            //2. aller sur ""prod inflight""
            tabletAppPage.GotoTabletApp_ProdInFlight();
            Thread.Sleep(2000);
            //3. Sélectionne une date FROM TO
            tabletAppPage.Filter(TabletAppPage.FilterType.From, DateUtils.Now.Date);
            tabletAppPage.Filter(TabletAppPage.FilterType.To, DateUtils.Now.Date.AddMonths(1));
            //4. Sélectioner un doc haccp dans le menu déroulant
            tabletAppPage.Select("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-2", DocTitle);

            //5. Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");
            //6. choisir une date de production et cliquer sur print
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();

            //7. renseigner plusieurs dates dans la colonne où le format multi date est paramétré
            tabletAppPage.addLineTraySetup(true, false);
            Assert.IsTrue(tabletAppPage.CheckMultiDate(), "Probleme MultiDate");

            //8. cliquer sur save pour enregistrer un document temporary
            tabletAppPage.ClickButton("Save");
            tabletAppPage.FileFileNameProdInFlight0(DocTmpName);
            //cliquer sur ""validate""
            tabletAppPage.ClickButton("Validate");
            try
            {
                // écraser le fichier
                tabletAppPage.ClickButton("Yes");
            }
            catch
            {
                // nouveau fichier
            }
            //7.une pop de validation apparait et puis cliquer sur OK
            tabletAppPage.ClickButton("Ok");

            tabletAppPage.Close();
            Thread.Sleep(2000);
            //9. récupérer ensuite le document temporary dans la page d'acceuil de prod inflight dans la section "" temporary""
            homePage.Navigate();
            tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            // si je suis déjà passé par /tablet, ça affiche ACE en haut à gauche, et le reste de la page est blanche
            tabletAppPage.GotoTabletApp_ProdInFlight();
            Thread.Sleep(2000);
            tabletAppPage.Select("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-2", DocTitle);
            tabletAppPage.SelectAction("mat-select-6", DocTmpName, true);

            //5. Cliquer sur next pour ouvrir le document
            tabletAppPage.ClickBaseButton("NEXT STEP");
            //6. choisir une date de production et cliquer sur print

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();


            //10.vérifier que la case coché dans l'enregistrement reste bien cochée sur le document temporary
            tabletAppPage.EditLineTraySetup();
            Assert.IsTrue(tabletAppPage.CheckMultiDate(), "Probleme MultiDate");
            //11. puis cliquer sur print pour générer le rapport haccp en pdf
            tabletAppPage.ClickButton("Print");

            //12. vérifier que la date est correcte sur le rapport généré "
            //1. On récupère bien les dates enregistrés dans le document temporary et au bon format défini dans les global settings (settings=> global settings=> DateFormatDisplay)
            //2.Les dates doivent être conformes dans le print généré dans le bon format également
            //tabletAppPage.WaitHACCPHorizontalProgressBar();
            FileInfo fi = new FileInfo(downloadsPath + "\\" + DocFileName);

            int counter = 0;
            do
            {
                fi.Refresh();
                Thread.Sleep(1000);
                counter++;
            } while (!fi.Exists && counter < 20);

            Assert.IsTrue(fi.Exists, DocFileName + " non généré cas 1");

            PdfDocument document = PdfDocument.Open(fi.FullName);
            Page page1 = document.GetPage(1);
            IEnumerable<Word> words = page1.GetWords();
            string yearAndMonth = DateUtils.Now.ToString("/MM/yyyy");
            if (DateUtils.Now.Day != 01)
            {
                Assert.IsFalse(words.Any(w => w.Text == (01 + yearAndMonth)), (01 + yearAndMonth) + " MultiDate trouvé dans " + fi.FullName);
            }
            for (int i = 10; i < 01; i++)
            {
                Assert.IsTrue(words.Any(w => w.Text == (i + yearAndMonth)), (i + yearAndMonth) + " MultiDate non trouvé dans " + fi.FullName);
            }

        }


        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_ProdInflight_FilterShowHiddenArticles()
        {
            // Prepare
            string DocSite = "ACE";
            string DocTitle = "HACCP - Thawing";
            string ItemGroup = "AIR CANADA";

            // Arrange
            var homePage = LogInAsAdmin();


            // Prepare Data : Avoir des articles "hidden" dans les recettes contenues dans les datasheets liés aux services prévus dans l'intervalle de temps "FROM-TO" sélectionné
            FilterAndFavoritesPage filterFavPage = homePage.GoToProduction_ProductionManagemenentPage();
            filterFavPage.ResetFilter();
            filterFavPage.Filter(FilterAndFavoritesPage.FilterType.DateFrom, DateTime.Now.AddDays(-2));
            QuantityAdjustmentsPage ajdPage = filterFavPage.DoneToQtyAjustement();
            ResultPage resultPage = ajdPage.GoToResultPage();
            resultPage.ToggleItem();
            Thread.Sleep(2000);

            if (ItemGroup == resultPage.GetFirstItemGroup())
            {
                resultPage.HideFirstArticleResult();
            }
            else
            {
                ItemGroup = resultPage.GetFirstItemGroup();
            }
            homePage.Navigate();
            Thread.Sleep(2000);

            TabletAppPage tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);

            tabletAppPage.GotoTabletApp_ProdInFlight();
            Thread.Sleep(2000);
            tabletAppPage.Filter(TabletAppPage.FilterType.From, DateUtils.Now.Date.AddDays(-1));
            tabletAppPage.Filter(TabletAppPage.FilterType.To, DateUtils.Now.Date.AddDays(1));
            tabletAppPage.Select("mat-select-2", DocTitle);

            //Cliquer sur le filtre show hidden articles
            tabletAppPage.Filter(TabletAppPage.FilterType.ShowHiddenArticles, true);

            tabletAppPage.ClickBaseButton("NEXT STEP");
            tabletAppPage.ClickButton("Print");

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();
            Thread.Sleep(2000);
            // busy indicator
            tabletAppPage.WaitHACCPHorizontalProgressBar();

            // Afficher également les articles "hidden" dans les documents haccp
            string countAirCanada = "//*[text()='" + ItemGroup + "']/parent::div";
            Assert.AreEqual(1, WebDriver.FindElements(By.XPath(countAirCanada)).Count, "article hidden anormalement non visible");
            tabletAppPage.Close();
            homePage.Navigate();
            tabletAppPage = homePage.GotoTabletApp();
            tabletAppPage.Select("mat-select-0", DocSite);
            tabletAppPage.GotoTabletApp_ProdInFlight();

            // unshow (en cache ?!?)
            tabletAppPage.Filter(TabletAppPage.FilterType.ShowHiddenArticles, true);

            tabletAppPage.ClickBaseButton("NEXT STEP");
            tabletAppPage.ClickButton("Print");
            tabletAppPage.WaitForLoad();

            // je suis dans un nouvel onglet !!!!
            tabletAppPage.GotoTabletApp_DocumentTab();

            Assert.AreEqual(0, WebDriver.FindElements(By.XPath(countAirCanada)).Count, "article hidden anormalement visible");
            tabletAppPage.Close();
            homePage.Navigate();
            homePage.GoToProduction_ProductionManagemenentPage();

            filterFavPage.ResetFilter();
            filterFavPage.Filter(FilterAndFavoritesPage.FilterType.DateFrom, DateTime.Now.AddDays(-2));
            ajdPage = filterFavPage.DoneToQtyAjustement();

            resultPage = ajdPage.GoToResultPage();
            resultPage.Filter(ResultPage.FilterType.ShowHiddenArticlesResults, true);
            resultPage.ToggleItem();
            resultPage.HideFirstArticleResult();
            WebDriver.Navigate().Refresh();
            resultPage = ajdPage.GoToResultPage();
        }

        [TestMethod]
        [Timeout(Timeout)]
        public void TB_HACCP_ProdInflight_FilterStartEndTime()
        {
            // Prepare
            string DocSite = "ACE";

            // Arrange
            var homePage = LogInAsAdmin();


            TabletAppPage tabletAppPage = homePage.GotoTabletApp();

            //sélectionner un site sur la tablette
            tabletAppPage.Select("mat-select-0", DocSite);
            //aller sur ""prod inflight""
            tabletAppPage.GotoTabletApp_ProdInFlight();

            // Mettre FROM TO sur la même date
            tabletAppPage.Filter(TabletAppPage.FilterType.From, DateUtils.Now.Date);
            tabletAppPage.Filter(TabletAppPage.FilterType.To, DateUtils.Now.Date);
            tabletAppPage.Filter(TabletAppPage.FilterType.StartTime, "0500AM");
            tabletAppPage.Filter(TabletAppPage.FilterType.EndTime, "0500PM");
            Assert.IsTrue(tabletAppPage.CanNextStep(), "Bouton Next Step non valide");
        }
       
    }
}