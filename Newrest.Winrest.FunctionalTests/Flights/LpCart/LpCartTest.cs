using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.Flight;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.LoadingPlans;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.LpCart;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using UglyToad.PdfPig;
using static Newrest.Winrest.FunctionalTests.PageObjects.Flights.LpCart.LpCartPage;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Customer;
using DocumentFormat.OpenXml;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2010.Excel;
using System.Security.Policy;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using System.Web.UI.WebControls.WebParts;
using System.Windows.Media.Media3D;
using static Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service.ServiceMassiveDeleteModalPage;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Web.Configuration;
using System.Web.Routing;
using DocumentFormat.OpenXml.Bibliography;
using System.Runtime.InteropServices.ComTypes;
using Microsoft.VisualBasic;

namespace Newrest.Winrest.FunctionalTests.Flights
{
    [TestClass]
    public class LpCartTest : TestBase
    {
        private string lpCartName = "LpCartName"+  new Random().Next(1,100000) ;

        // Impossible de évaluer l’expression, car un frame natif se trouve en haut de la pile des appels.
        private const int _timeout = 60 * 10 * 1000;
        // test initialize
        [TestInitialize]
        public override void TestInitialize()
        {
            base.TestInitialize();
            var testMethod = TestContext.TestName;
            switch (testMethod)
            {
                case nameof(FL_LP_CART_FusionTrolley):
                    TestInitialiazeCreateLpCart();
                    break;

                case nameof(FL_LP_CART_UndoTrolley):
                    TestInitialiazeCreateLpCart();
                    break;

                default:
                    break;
            }
        }

        private void TestInitialiazeCreateLpCart()
        {
            string site = TestContext.Properties["SiteACE"].ToString();
            string customersname = "$$ - CAT Genérico";
            string aircraft = "AB310";
            DateTime from = DateUtils.Now.AddDays(-1);
            DateTime to = DateUtils.Now.AddDays(10);
            string comment = "Bob comment";
            HomePage homePage = LogInAsAdmin();
            // Act
            var lpCartPage = homePage.GoToFlights_LpCartPage();
            var lpCartModalCreate = lpCartPage.LpCartCreatePage();
            var LpCartDetailPage = lpCartModalCreate.FillField_CreateNewLpCart(lpCartName, lpCartName, site, customersname, aircraft, from, to, comment);
            //Acces Onglet Cart                
            LpCartDetailPage.BackToList();
        }
        
        [Priority(0)]
        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_DeleteLoadingPlanForTrolley()
        {
            var adminName = TestContext.Properties["Admin_UserName"].ToString();
            string userName = adminName.Substring(0, adminName.IndexOf("@"));

            string loadingPlanName = "test loading plan tr";
            string loadingPlanNameBis = "Bis loading plan tr bis";

            //Loading plan
            string loadingPlanName3 = TestContext.Properties["LoadingPlanBob"].ToString();
            string loadingPlanName4 = TestContext.Properties["LoadingPlanBobBis"].ToString();
            string site = TestContext.Properties["SiteLP"].ToString();

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            ClearCache();

            //Prepare
            var parametersUserPage = homePage.GoToParameters_User();
            parametersUserPage.SearchAndSelectUser(userName);
            parametersUserPage.ClickOnAffectedSite();
            parametersUserPage.SelectUnselectAllSites(true);

            var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
            loadingPlanPage.ResetFilter();
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, "MAD");
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName);
            if (loadingPlanPage.CheckTotalNumber() != 0)
            {
                var generalInformationsPage = loadingPlanPage.ClickOnFirstLoadingPlan();
                generalInformationsPage.EditLoadingPlanInformations(DateUtils.Now.AddMonths(1), DateUtils.Now.AddDays(-14));
                generalInformationsPage.BackToList();
            }


            loadingPlanPage.ResetFilter();
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, "MAD");
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanNameBis);
            if (loadingPlanPage.CheckTotalNumber() != 0)
            {
                var generalInformationsPage = loadingPlanPage.ClickOnFirstLoadingPlan();
                generalInformationsPage.EditLoadingPlanInformations(DateUtils.Now.AddMonths(1), DateUtils.Now.AddDays(-14));
                generalInformationsPage.BackToList();

            }

            ////Desactive all loading Plan et LpCart
            loadingPlanPage.ResetFilter();
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName4);
            if (loadingPlanPage.CheckTotalNumber() != 0)
            {
                var generalInformationsPage = loadingPlanPage.ClickOnFirstLoadingPlan();
                generalInformationsPage.DeleteLoadingPlan();
            }

            loadingPlanPage.ResetFilter();
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
            loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName3);
            if (loadingPlanPage.CheckTotalNumber() != 0)
            {
                var generalInformationsPage = loadingPlanPage.ClickOnFirstLoadingPlan();
                generalInformationsPage.DeleteLoadingPlan();
            }
        }

        [Priority(1)]
        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_DesactiveAll_Lp_CardForConfig()
        {

            //prepare
            string name = TestContext.Properties["LpCartName2"].ToString();
            string lpCartValue = "None";

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // On récupère le format de date utilisé dans les date picker
            string dateFormat = homePage.GetDateFormatPickerValue();

            // Act
            var lpCartPage = homePage.GoToFlights_LpCartPage();

            //all lpcart with flight
            lpCartPage.ResetFilter();
            lpCartPage.Filter(LpCartPage.FilterType.Search, name);

            if (lpCartPage.CheckTotalNumber() != 0)
            {

                //Recupereation de Flights
                var lpCartDetail = lpCartPage.LpCartCartDetailPage();
                var lpCartFlightdetail = lpCartDetail.LpCartFlightDetailPage();

                lpCartFlightdetail.Filter(LpCartFlightDetailPage.FilterType.StartDate, DateUtils.Now.AddDays(-4));
                var flights = lpCartFlightdetail.GetFlightName();
                var flights_dates = lpCartFlightdetail.GetFlightDates();
                Assert.AreEqual(flights.Count, flights_dates.Count, "Problème d'alignement tableau - cas 1");
                // Récupération du type de séparateur (, ou . selon les pays)
                CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

                Desactive_LpCarts(flights, flights_dates, ci, lpCartValue);

            }

            // Act
            homePage.GoToFlights_LpCartPage();

            //all lpcart with flight
            lpCartPage.ResetFilter();
            lpCartPage.Filter(LpCartPage.FilterType.Search, "trolleyFlightBob");

            if (lpCartPage.CheckTotalNumber() != 0)
            {
                var lpCartDetail = lpCartPage.LpCartCartDetailPage();
                var lpCartFlightdetail = lpCartDetail.LpCartFlightDetailPage();

                lpCartFlightdetail.Filter(LpCartFlightDetailPage.FilterType.StartDate, DateUtils.Now.AddDays(-4));
                var flights = lpCartFlightdetail.GetFlightName();
                var flights_dates = lpCartFlightdetail.GetFlightDates();
                Assert.AreEqual(flights.Count, flights_dates.Count, "Problème d'alignement tableau - cas 2");
                // Récupération du type de séparateur (, ou . selon les pays)
                CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

                Desactive_LpCarts(flights, flights_dates, ci, lpCartValue);
            }
        }


        public void Desactive_LpCarts(List<string> flights, List<string> flights_dates, CultureInfo ci, string lpCartValue)
        {
            var homePage = new HomePage(WebDriver, TestContext);
            var flightPage = homePage.GoToFlights_FlightPage();

            for (int i = 0; i < flights.Count; i++)
            {
                var flight = flights[i];
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.DispalyFlight, "Display both");
                flightPage.Filter(FlightPage.FilterType.Sites, "MAD");
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flight);

                var date = flights_dates[i];
                var dateParse = DateTime.Parse(date, ci);
                flightPage.SetDateState(dateParse);

                if (flightPage.IsArrival(i))
                {
                    continue;
                }

                if (flightPage.IsFlightExist())
                {
                    if (flightPage.IsValidated())
                    {
                        flightPage.UnSetNewState("V");
                    }

                    var flightModalEdit = flightPage.EditFirstFlight(flight);
                    flightModalEdit.SelectLpCart(lpCartValue);
                    flightModalEdit.CloseViewDetails();
                }
                else
                {
                    flightPage.Filter(FlightPage.FilterType.Sites, "MAD");

                    if (flightPage.IsArrival(1))
                    {
                        continue;
                    }

                    if (flightPage.IsFlightExist())
                    {
                        if (flightPage.IsValidated())
                        {
                            flightPage.UnSetNewState("V");
                        }

                        var flightModalEdit = flightPage.EditFirstFlight(flight);
                        flightModalEdit.SelectLpCart(lpCartValue);
                        flightModalEdit.CloseViewDetails();
                    }

                }
            }
        }
        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_Delete_Lp_Cart()
        {
            // Arrange
            Random rnd = new Random();
            // Prepare premier LpCart
            string code = rnd.Next(0, 500).ToString();
            string name = "LP_Cart delete test";
            string site = TestContext.Properties["SiteLpCart"].ToString();
            string customer = TestContext.Properties["CustomerLpCart"].ToString();
            string aircraft = TestContext.Properties["AircraftLpCart"].ToString();
            DateTime from = DateUtils.Now.AddDays(5);
            DateTime to = DateUtils.Now.AddDays(10);
            string comment = TestContext.Properties["CommentLpCart"].ToString();

            // Act
            HomePage homePage = LogInAsAdmin();
            var lpCartPage = homePage.GoToFlights_LpCartPage();

            //Create LP cart
            var lpCartModalCreate = lpCartPage.LpCartCreatePage();
            var LpCartDetailPage = lpCartModalCreate.FillField_CreateNewLpCart(code, name, site, customer, aircraft, from, to, comment);
            LpCartDetailPage.BackToList();
            //Delete LpCart
            lpCartPage.ResetFilter();
            lpCartPage.Filter(LpCartPage.FilterType.Search, name);
            lpCartPage.DeleteLpCart();
            var isLpCartDeleted = lpCartPage.CheckTotalNumber() == 0;
            Assert.IsTrue(isLpCartDeleted, "Le lpCart n'est pas supprimé");
        }

        [Priority(2)]
        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_Create_Lp_CardForConfig()
        {
            Random rnd = new Random();
            Random rnd1 = new Random();
            Random rnd2 = new Random();
            // Prepare premier LpCart
            string code = rnd.Next(0, 500).ToString();
            string name = TestContext.Properties["LpCartName"].ToString();
            string site = TestContext.Properties["SiteLpCart"].ToString();
            string customer = TestContext.Properties["CustomerLpCart"].ToString();
            string aircraft = TestContext.Properties["AircraftLpCart"].ToString();
            DateTime from = DateUtils.Now.AddDays(5);
            DateTime to = DateUtils.Now.AddDays(10);
            string comment = TestContext.Properties["CommentLpCart"].ToString();


            // Prepare second LpCart
            string code1 = rnd1.Next(0, 5000).ToString();
            string name1 = TestContext.Properties["LpCartName1"].ToString();
            string site1 = TestContext.Properties["SiteLpCart1"].ToString();
            string customer1 = TestContext.Properties["CustomerLpCart1"].ToString();
            string aircraft1 = TestContext.Properties["AircraftLpCart"].ToString();
            DateTime from1 = DateUtils.Now.AddDays(0);
            DateTime to1 = DateUtils.Now.AddDays(4);
            string comment1 = TestContext.Properties["CommentLpCart"].ToString();

            // Prepare troisieme LpCart with flight
            string code2 = rnd2.Next(0, 50000).ToString();
            string name2 = TestContext.Properties["LpCartName2"].ToString();
            string site2 = TestContext.Properties["SiteLpCart2"].ToString();
            string customer2 = TestContext.Properties["CustomerLpCart2"].ToString();
            string aircraft2 = TestContext.Properties["AircraftLpCart"].ToString();
            DateTime from2 = DateUtils.Now.AddDays(-5);
            DateTime to2 = DateUtils.Now.AddDays(10);
            string comment2 = TestContext.Properties["CommentLpCart"].ToString();

            // Prepare quatrieme LpCart avec trolley
            string code3 = rnd2.Next(0, 500000).ToString();
            string name3 = TestContext.Properties["LpCartName4"].ToString();
            string site3 = TestContext.Properties["SiteLpCart2"].ToString();
            string customer3 = TestContext.Properties["CustomerLpCart2"].ToString();
            string aircraft3 = TestContext.Properties["AircraftLpCart"].ToString();
            DateTime from3 = DateUtils.Now.AddDays(1);
            DateTime to3 = DateUtils.Now.AddDays(10);
            string comment3 = TestContext.Properties["CommentLpCart"].ToString();
            string trolleyName = TestContext.Properties["TrolleyName"].ToString();
            string trolleyName1 = TestContext.Properties["TrolleyName1"].ToString();
            string trolleyScheme = TestContext.Properties["TrolleySchemeName"].ToString();
            string trolleyScheme1 = TestContext.Properties["TrolleySchemeName1"].ToString();


            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var lpCartPage = homePage.GoToFlights_LpCartPage();

            // Create
            //--------------------------------------------LpCart 1 avec trolleyS --------------------------------------------
            lpCartPage.Filter(LpCartPage.FilterType.Search, name3);
            if (lpCartPage.CheckTotalNumber() == 0)
            {
                var lpCartModalCreate = lpCartPage.LpCartCreatePage();
                var lpCartCartDetailPage = lpCartModalCreate.FillField_CreateNewLpCart(code3, name3, site3, customer3, aircraft3, from3, to3, comment3);

                //Acces Onglet Cart                
                lpCartCartDetailPage.ClickAddtrolley();
                lpCartCartDetailPage.AddTrolley(trolleyName);

                //Create LpCart Scheme
                var lpCartSchemeModal = lpCartPage.LpCartCreateLpCartSchemeModal();
                lpCartSchemeModal.CreateLpCartscheme(trolleyScheme, "2", "2");

                lpCartCartDetailPage.ClickAddtrolley();
                lpCartCartDetailPage.AddTrolley(trolleyName1);

                //Create LpCart Scheme
                lpCartPage.LpCartCreateLpCartSchemeModal();
                lpCartSchemeModal.CreateLpCartscheme(trolleyScheme1, "2", "2");

                lpCartCartDetailPage.BackToList();
            }
            else
            {
                var lpCartCartDetailPage = lpCartPage.LpCartCartDetailPage();
                if (lpCartCartDetailPage.CheckTotalNumber() > 2)
                {
                    lpCartCartDetailPage.DeleteAllLpCartScheme();

                    //Acces Onglet Cart                
                    lpCartCartDetailPage.ClickAddtrolley();
                    lpCartCartDetailPage.AddTrolley(trolleyName);

                    //Create LpCart Scheme
                    var lpCartSchemeModal = lpCartPage.LpCartCreateLpCartSchemeModal();
                    lpCartSchemeModal.CreateLpCartscheme(trolleyScheme, "2", "2");

                    lpCartCartDetailPage.ClickAddtrolley();
                    lpCartCartDetailPage.AddTrolley(trolleyName1);

                    //Create LpCart Scheme
                    lpCartPage.LpCartCreateLpCartSchemeModal();
                    lpCartSchemeModal.CreateLpCartscheme(trolleyScheme1, "2", "2");
                }

                lpCartCartDetailPage.BackToList();
            }

            //
            //-----------------------LpCart 2 -----------------------
            lpCartPage.Filter(LpCartPage.FilterType.Search, name);
            if (lpCartPage.CheckTotalNumber() == 0)
            {
                var lpCartModalCreate = lpCartPage.LpCartCreatePage();
                var LpCartDetailPage = lpCartModalCreate.FillField_CreateNewLpCart(code, name, site, customer, aircraft, from, to, comment);

                LpCartDetailPage.BackToList();

            }

            //-----------------------LpCart 3 -----------------------
            lpCartPage.Filter(LpCartPage.FilterType.Search, name1);
            if (lpCartPage.CheckTotalNumber() == 0)
            {
                var lpCartModalCreate = lpCartPage.LpCartCreatePage();
                var LpCartDetailPage = lpCartModalCreate.FillField_CreateNewLpCart(code1, name1, site1, customer1, aircraft1, from1, to1, comment1);

                LpCartDetailPage.BackToList();
            }

            //-----------------------LpCart Calomato -----------------------
            lpCartPage.Filter(LpCartPage.FilterType.Search, name2);
            if (lpCartPage.CheckTotalNumber() == 0)
            {
                var lpCartModalCreate = lpCartPage.LpCartCreatePage();
                var LpCartDetailPage = lpCartModalCreate.FillField_CreateNewLpCart(code2, name2, site2, customer2, aircraft2, from2, to2, comment2);

                //Acces Onglet Cart                
                LpCartDetailPage.ClickAddtrolley();
                LpCartDetailPage.AddTrolley(trolleyName);
                var lpCartSchemeModal = lpCartPage.LpCartCreateLpCartSchemeModal();
                lpCartSchemeModal.CreateLpCartscheme(trolleyScheme, "2", "2");
                LpCartDetailPage.BackToList();
            }
            else
            {
                var lpCartCartDetailPage = lpCartPage.LpCartCartDetailPage();
                if (lpCartCartDetailPage.CheckTotalNumber() > 2)
                {
                    lpCartCartDetailPage.DeleteAllLpCartScheme();

                    //Acces Onglet Cart                
                    lpCartCartDetailPage.ClickAddtrolley();
                    lpCartCartDetailPage.AddTrolley(trolleyName);

                    //Create LpCart Scheme
                    var lpCartSchemeModal = lpCartPage.LpCartCreateLpCartSchemeModal();
                    lpCartSchemeModal.CreateLpCartscheme(trolleyScheme, "2", "2");
                }

                lpCartCartDetailPage.BackToList();
            }
        }

        //Création d'un flight lié à un lpCart 
        [Priority(3)]
        [Timeout(_timeout)]
        [TestMethod]
        public void LP_CART_Create_New_FlightsForLpCart()
        {
            string aircraft = TestContext.Properties["AircraftLpCart"].ToString();
            string customer = TestContext.Properties["CustomerLpCart2"].ToString();
            string BCN = TestContext.Properties["SiteToFlight"].ToString();
            string ACE = TestContext.Properties["Site"].ToString();
            string lpcart = TestContext.Properties["LpCartName2"].ToString();

            var flightNumber = "Albata - " + DateUtils.Now.AddDays(2).ToString("dd-MM-yyyy");
            var flightNumber1 = "Bonnata - " + DateUtils.Now.ToString("dd-MM-yyyy");
            var flightNumber1Bis = "Connata - " + DateUtils.Now.AddDays(1).ToString("dd-MM-yyyy");
            var flightNumber2 = "Caramata - " + DateUtils.Now.AddDays(1).ToString("dd-MM-yyyy");


            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Act
            var flightPage = homePage.GoToFlights_FlightPage();


            // Create
            //figth site Ace lastDate middle EDT
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, "MAD");
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            // LEN est un airport, pas un site en fait, donc on prend le site destination : BCN
            flightPage.Filter(FlightPage.FilterType.DispalyFlight, "Display flight arrival");
            flightPage.SetDateState(DateUtils.Now.AddDays(2));

            if (flightPage.IsFlightExist())
            {
                Assert.IsTrue(flightPage.IsArrival(1), "avion " + flightNumber + " n'est pas un vol d'arrivée");
            }
            else
            {
                //figth site LEN lastDate middle EDT
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber.ToString(), customer, aircraft, "LEN", "MAD", null, "13", "15", lpcart, DateUtils.Now.AddDays(2));
            }

            //Assert
            Assert.AreEqual(flightNumber, flightPage.GetFirstFlightNumber(), "Le vol n'a pas été créé.");

            // Create
            flightPage.MakeSureModalIsClosed();
            //flight site MAD firstDate last EDT
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, "MAD");
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber1);
            flightPage.SetDateState(DateUtils.Now);

            if (flightPage.IsFlightExist())
            {
                Assert.IsFalse(flightPage.IsArrival(1), "avion " + flightNumber1 + " est un vol d'arrivée");
                if (flightPage.IsValidated())
                {
                    flightPage.UnSetNewState("V");
                }

                var flightModalEdit = flightPage.EditFirstFlight(flightNumber1);
                flightModalEdit.SelectLpCart(lpcart);
                flightModalEdit.CloseViewDetails();
            }
            else
            {
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber1.ToString(), customer, aircraft, "MAD", ACE, null, "13", "18", lpcart);
            }

            //Assert
            Assert.AreEqual(flightNumber1, flightPage.GetFirstFlightNumber(), "Le vol n'a pas été créé.");

            // Create
            //flight site MAD firstDate last EDT
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, "MAD");
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber1Bis);
            flightPage.SetDateState(DateUtils.Now.AddDays(1));

            if (flightPage.IsFlightExist())
            {
                Assert.IsFalse(flightPage.IsArrival(1), "avion " + flightNumber1Bis + " est un vol d'arrivée");
                if (flightPage.IsValidated())
                {
                    flightPage.UnSetNewState("V");
                }
                var flightModalEdit = flightPage.EditFirstFlight(flightNumber1Bis);
                flightModalEdit.SelectLpCart(lpcart);
                flightModalEdit.CloseViewDetails();
                flightPage.UnSetNewState("V");
            }
            else
            {
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber1Bis.ToString(), customer, aircraft, "MAD", ACE, null, "13", "18", lpcart, DateUtils.Now.AddDays(1));
            }

            //Assert
            Assert.AreEqual(flightNumber1Bis, flightPage.GetFirstFlightNumber(), "Le vol n'a pas été créé.");

            // Create
            //flight site SLM middleDate first EDT
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, "MAD");
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber2);
            flightPage.Filter(FlightPage.FilterType.DispalyFlight, "Display flight arrival");
            flightPage.SetDateState(DateUtils.Now.AddDays(1));

            if (flightPage.IsFlightExist())
            {
                Assert.IsTrue(flightPage.IsArrival(1), "avion " + flightNumber2 + " n'est pas un vol d'arrivée");
            }
            else
            {
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber2.ToString(), customer, aircraft, "SLM", "MAD", null, "13", "13", lpcart, DateUtils.Now.AddDays(1));
            }

            //Assert
            flightPage.Filter(FlightPage.FilterType.DispalyFlight, "Display flight arrival");
            Assert.AreEqual(flightNumber2, flightPage.GetFirstFlightNumber(), "Le vol n'a pas été créé.");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_Create_Lp_Card()
        {
            // Prepare
            string code = new Random().Next().ToString();
            var name = "Albata - " + DateUtils.Now.AddDays(2).ToString("dd-MM-yyyy");
            string site = TestContext.Properties["SiteLpCart"].ToString();
            string customer = TestContext.Properties["CustomerLpCart"].ToString();
            string aircraft = TestContext.Properties["AircraftLpCart"].ToString();
            DateTime from = DateUtils.Now.AddDays(5);
            DateTime to = DateUtils.Now.AddDays(10);
            string comment = TestContext.Properties["CommentLpCart"].ToString();
            // Arrange
            HomePage homePage = LogInAsAdmin();
            // Act
            var lpCartPage = homePage.GoToFlights_LpCartPage();
            try {
                var lpCartModalCreate = lpCartPage.LpCartCreatePage();
                var LpCartDetailPage = lpCartModalCreate.FillField_CreateNewLpCart(code, name, site, customer, aircraft, from, to, comment);
                LpCartDetailPage.BackToList();
                lpCartPage.ResetFilter();
                lpCartPage.Filter(LpCartPage.FilterType.Search, name);
                var list = lpCartPage.GetLPCartsFiltred();
                Assert.IsTrue(list.Contains(name), $"Le nom {name} n'a pas été trouvé dans la liste.");
            }
            finally
            {
                lpCartPage.ResetFilter();
                lpCartPage.Filter(LpCartPage.FilterType.Search, name);
                lpCartPage.DeleteLpCart();
            }
        }



        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_FilterName()
        {
            // Prepare
            string name = TestContext.Properties["LpCartName"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var lpCartPage = homePage.GoToFlights_LpCartPage();

            //Filter by Search
            lpCartPage.ResetFilter();
            lpCartPage.Filter(LpCartPage.FilterType.Search, name);

            // Assert
            Assert.AreEqual(name, lpCartPage.GetFirstLpCartName(), "Aucun lp cart trouvé.");

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_FilterSortByDate()
        {
            //Prepare
            //lpCart & cart
            var code = new Random()
                .Next(0, 5000)
                .ToString();
            var customerFilter = TestContext.Properties["CustomerLpFilter"].ToString();
            var site = TestContext.Properties["SiteACE"].ToString();
            var aircraft = TestContext.Properties["AircraftLpCart"].ToString();
            var lpCartName1 = $"LP Cart1 {DateTime.Now.ToString("dd/MM/yyyy")}";
            var lpCartName2 = $"LP Cart2 {DateTime.Now.ToString("dd/MM/yyyy")}";
            var fromLpCart = DateTime.Today.AddMonths(-3);
            var toLpCart = DateTime.Today.AddMonths(+2);
            var comment = "comment";

            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var lpCartPage = homePage.GoToFlights_LpCartPage();
            var lpCartNumber = lpCartPage.CheckTotalNumber();
            try
            {
                if (lpCartNumber < 2)
                {
                    var lpCartModalCreate = lpCartPage.LpCartCreatePage();
                    //Create lp cart 1
                    LpCartCartDetailPage lpCartCartDetailPage = lpCartModalCreate
                        .FillField_CreateNewLpCart(code, lpCartName1, site, customerFilter, aircraft, fromLpCart, toLpCart, comment);
                    lpCartCartDetailPage.BackToList();

                    lpCartModalCreate = lpCartPage.LpCartCreatePage();
                    //Create lp cart 2
                    lpCartCartDetailPage = lpCartModalCreate
                        .FillField_CreateNewLpCart(code, lpCartName2, site, customerFilter, aircraft, fromLpCart, toLpCart, comment);
                    lpCartCartDetailPage.BackToList();
                }

                //Filter by Name
                lpCartPage.Filter(LpCartPage.FilterType.SortBy, "StartDate");
                var isSortedByDate = lpCartPage.IsSortedByDate();

                // Assert
                Assert.IsTrue(isSortedByDate, MessageErreur.FILTRE_ERRONE, "Sort by date");
            }
            finally
            {
                if(lpCartNumber < 2)
                {
                    lpCartPage = homePage.GoToFlights_LpCartPage();
                    lpCartPage.ResetFilter();
                    lpCartPage.Filter(FilterType.Search, lpCartName1);
                    lpCartPage.DeleteLpCart();

                    lpCartPage.ResetFilter();
                    lpCartPage.Filter(FilterType.Search, lpCartName2);
                    lpCartPage.DeleteLpCart();
                }
            }
        }
        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_FilterSortByName()
        {
            //lpCart & cart
            var code = new Random()
                .Next(0, 5000)
                .ToString();
            var customerFilter = TestContext.Properties["CustomerLpFilter"].ToString();
            var site = TestContext.Properties["SiteACE"].ToString();
            var aircraft = TestContext.Properties["AircraftLpCart"].ToString();
            var lpCartName1 = $"LP Cart1 {DateTime.Now.ToString("dd/MM/yyyy")}";
            var lpCartName2 = $"LP Cart2 {DateTime.Now.ToString("dd/MM/yyyy")}";
            var fromLpCart = DateTime.Today.AddMonths(-3);
            var toLpCart = DateTime.Today.AddMonths(+2);
            var comment = "comment";

            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var lpCartPage = homePage.GoToFlights_LpCartPage();
            var lpCartNumber = lpCartPage.CheckTotalNumber();
            try
            {
                if (lpCartNumber < 2)
                {
                    var lpCartModalCreate = lpCartPage.LpCartCreatePage();
                    //Create lp cart 1
                    LpCartCartDetailPage lpCartCartDetailPage = lpCartModalCreate
                        .FillField_CreateNewLpCart(code, lpCartName1, site, customerFilter, aircraft, fromLpCart, toLpCart, comment);
                    lpCartCartDetailPage.BackToList();

                    lpCartModalCreate = lpCartPage.LpCartCreatePage();
                    //Create lp cart 2
                    lpCartCartDetailPage = lpCartModalCreate
                        .FillField_CreateNewLpCart(code, lpCartName2, site, customerFilter, aircraft, fromLpCart, toLpCart, comment);
                    lpCartCartDetailPage.BackToList();
                }

                //Filter by Name
                lpCartPage.Filter(LpCartPage.FilterType.SortBy, "StartDate");
                lpCartPage.Filter(LpCartPage.FilterType.SortBy, "Number");
                var isSortedByName = lpCartPage.IsSortedByName();

                // Assert
                Assert.IsTrue(isSortedByName, MessageErreur.FILTRE_ERRONE, "Sort by name");
            }
            finally
            {
                if (lpCartNumber < 2)
                {
                    lpCartPage = homePage.GoToFlights_LpCartPage();
                    lpCartPage.ResetFilter();
                    lpCartPage.Filter(FilterType.Search, lpCartName1);
                    lpCartPage.DeleteLpCart();

                    lpCartPage.ResetFilter();
                    lpCartPage.Filter(FilterType.Search, lpCartName2);
                    lpCartPage.DeleteLpCart();
                }
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_Settings_Filter()
        {
            string siteToSelect = "ACE";
            string site1 = TestContext.Properties["SiteLpCart"].ToString();
            var adminName = TestContext.Properties["Admin_UserName"].ToString();
            string userName = adminName.Substring(0, adminName.IndexOf("@"));
            HomePage homePage = LogInAsAdmin();
            PageObjects.Parameters.User.ParametersUser parametersUserPage = homePage.GoToParameters_User();
            parametersUserPage.SearchAndSelectUser(userName);
            parametersUserPage.ClickOnAffectedSite();
            parametersUserPage.SelectUnselectAllSites(false);
            try
            {
                parametersUserPage.SelectOneSite(siteToSelect);
                LpCartPage lpCartPage = homePage.GoToFlights_LpCartPage();
                lpCartPage.ResetFilter();
                bool isVerified = lpCartPage.VerifyFilterSite(site1);
                Assert.IsTrue(isVerified, MessageErreur.FILTRE_ERRONE, "Sites");
            }
            finally
            {
                parametersUserPage = homePage.GoToParameters_User();
                parametersUserPage.SearchAndSelectUser(userName);
                parametersUserPage.ClickOnAffectedSite();
                parametersUserPage.SelectUnselectAllSites(true);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_Filter_Site()
        {
            string site = "ACE - ACE";
            string site1 = TestContext.Properties["SiteLpCart"].ToString();
            HomePage homePage = LogInAsAdmin();
            LpCartPage lpCartPage = homePage.GoToFlights_LpCartPage();
            lpCartPage.ResetFilter();
            lpCartPage.Filter(LpCartPage.FilterType.Site, site);
            bool isVerified = lpCartPage.VerifySite(site1);
            Assert.IsTrue(lpCartPage.VerifySite(site1), MessageErreur.FILTRE_ERRONE, "Sites");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_Filter_Customer()
        {
            string customer = TestContext.Properties["CustomerLpCart"].ToString();
            HomePage homePage = LogInAsAdmin();
            LpCartPage lpCartPage = homePage.GoToFlights_LpCartPage();
            lpCartPage.ResetFilter();
            lpCartPage.Filter(LpCartPage.FilterType.Customers, customer);
            bool isVerified = lpCartPage.VerifyCustomer(customer.Substring(0, 3));
            Assert.IsTrue(isVerified, MessageErreur.FILTRE_ERRONE, "Customers");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_Filter_CreateNewTrolley()
        {
            // Prepare
            string trolleyName = TestContext.Properties["TrolleyName"].ToString();
            string trolleyScheme = TestContext.Properties["TrolleySchemeName"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var lpCartPage = homePage.GoToFlights_LpCartPage();
            lpCartPage.ResetFilter();
            lpCartPage.Filter(LpCartPage.FilterType.Search, "Dercamte");

            //Acces Onglet Cart 
            var detailCartPage = lpCartPage.LpCartCartDetailPage();
            if (detailCartPage.CheckTotalNumber() > 0)
            {
                detailCartPage.DeleteAllLpCartScheme();
            }
            detailCartPage.ClickAddtrolley();
            detailCartPage.AddTrolley(trolleyName);

            //Create LpCart Scheme
            var lpCartSchemeModal = lpCartPage.LpCartCreateLpCartSchemeModal();
            lpCartPage.okWarning();
            lpCartSchemeModal.CreateLpCartscheme(trolleyScheme, "2", "2");

            //Assert
            Assert.AreEqual(1, detailCartPage.CheckTotalNumber(), String.Format(MessageErreur.OBJET_NON_CREE, "Le trolley"));
            Assert.IsTrue(detailCartPage.IstrolleySchemaExist(), String.Format(MessageErreur.OBJET_NON_CREE, "Le schema du trolley"));

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_DeleteTrolley()
        {

            // Prepare
            string name = TestContext.Properties["LpCartName"].ToString();
            string trolleyName = TestContext.Properties["TrolleyName"].ToString();
            string trolleyScheme = TestContext.Properties["TrolleySchemeName"].ToString();


            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var lpCartPage = homePage.GoToFlights_LpCartPage();
            lpCartPage.ResetFilter();

            lpCartPage.Filter(LpCartPage.FilterType.Search, name);
            var detailCartPage = lpCartPage.LpCartCartDetailPage();
            if (detailCartPage.CheckTotalNumber() > 0)
            {
                detailCartPage.DeleteAllLpCartScheme();
            }
            detailCartPage.ClickAddtrolley();
            detailCartPage.AddTrolley(trolleyName);

            //Create LpCart Scheme
            var lpCartSchemeModal = lpCartPage.LpCartCreateLpCartSchemeModal();
            lpCartSchemeModal.CreateLpCartscheme(trolleyScheme, "2", "2");

            Assert.AreEqual(1, detailCartPage.CheckTotalNumber(), String.Format(MessageErreur.OBJET_NON_CREE, "Le trolley"));

            detailCartPage.DeleteAllLpCartScheme();

            Assert.AreEqual(0, detailCartPage.CheckTotalNumber(), "Le trolley n'est pas supprimé");

        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_Filter_SearchOnCart()
        {

            // Prepare
            string name = TestContext.Properties["LpCartName4"].ToString();
            string trolleyName = TestContext.Properties["TrolleyName"].ToString();


            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var lpCartPage = homePage.GoToFlights_LpCartPage();
            lpCartPage.ResetFilter();

            lpCartPage.Filter(LpCartPage.FilterType.Search, name);
            var detailCartPage = lpCartPage.LpCartCartDetailPage();

            detailCartPage.Filter(LpCartCartDetailPage.FilterType.Search, trolleyName);

            //Assert
            Assert.IsTrue(detailCartPage.IsTrolleyName(trolleyName), String.Format(MessageErreur.FILTRE_ERRONE, "Search"));

        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_Filter_SearchFlightOnFlightTab()
        {

            string site = TestContext.Properties["SiteACE"].ToString();
            var aircraft = "AB310";
            var route = "ACE-AAA";
            Random rnd = new Random();
            string serviceName = "serviceNameToday" + "-" + new Random().Next().ToString();
            string serviceCode = new Random().Next().ToString();
            string serviceProduction = GenerateName(4);
            DateTime fromDate = DateUtils.Now.AddMonths(-1);
            DateTime toDate = DateUtils.Now.AddMonths(3);
            string loadingPlanName = "LgPlan " + rnd.Next().ToString();
            string guestName = "FC";
            string name = TestContext.Properties["LpCartNamebob"].ToString();
            string code ="lpCart - "+new Random().Next().ToString();
            DateTime from = DateUtils.Now.AddDays(-1);
            DateTime to = DateUtils.Now.AddDays(10);
            string comment = "Bob comment";
            string flightNumber = new Random().Next().ToString();
            string siteTo = TestContext.Properties["SiteAalTrolley"].ToString();
            LpCartPage lpCartPage = null;
            string customer = "$$ - CAT Genérico";
            string siteFrom = TestContext.Properties["SiteACE"].ToString(); ;
            string flightNumber2 = flightNumber +25;
            // Arrange
            HomePage homePage = LogInAsAdmin();
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            try
            {
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, serviceCode, serviceProduction);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();

                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                priceModalPage.FillFields_CustomerPrice(site, customer, fromDate, toDate);

                var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customer, route, aircraft, site);
                var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();

                loadingPlanDetailsPage.ClickAddGuestBtn();
                loadingPlanDetailsPage.SelectGuest(guestName);
                loadingPlanDetailsPage.ClickCreateGuestBtn();
                loadingPlanDetailsPage.ClickFirstGuest();


                loadingPlanDetailsPage.AddServiceBtn();
                loadingPlanDetailsPage.AddNewService(serviceName);
                servicePage = homePage.GoToCustomers_ServicePage();
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
                pricePage = servicePage.ClickOnFirstService();
                var serviceLoadingPage = pricePage.GoToLoadingPlanTab();
                var loadingPlanName1 = serviceLoadingPage.GetLoadingPlanName();
                lpCartPage = homePage.GoToFlights_LpCartPage();
                var lpCartCreateModalPage = lpCartPage.LpCartCreatePage();
                var LpCartDetailPage = lpCartCreateModalPage.FillField_CreateNewLpCartWithRoutes(code, code, site, customer, aircraft, fromDate,toDate, comment, route);
                FlightPage flightPage = homePage.GoToFlights_FlightPage();
                var flightCreatePageModal = flightPage.FlightCreatePage();
                flightCreatePageModal.WaitPageLoading();
                flightCreatePageModal.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo, null, "00", "23",code, null);
                flightCreatePageModal = flightPage.FlightCreatePage();
                flightCreatePageModal.WaitPageLoading();
                flightCreatePageModal.FillField_CreatNewFlight(flightNumber2, customer, aircraft, siteFrom, siteTo, null, "00", "23", code, null);
                lpCartPage = homePage.GoToFlights_LpCartPage();
                lpCartPage.Filter(LpCartPage.FilterType.Search, code);
                string FlightCount = lpCartPage.GetFlightCountNumber();
                Assert.AreEqual("2", FlightCount, "Le nombre de vols n'est pas correct");
                LpCartDetailPage = lpCartPage.ClickFirstLpCart();
                var lpCartFlightdetail = LpCartDetailPage.LpCartFlightDetailPage();
                lpCartFlightdetail.Filter(LpCartFlightDetailPage.FilterType.Search, flightNumber);
                var flightNames = lpCartFlightdetail.GetFlightNames();
                Assert.IsTrue(flightNames.Contains(flightNumber), $"Le numéro de vol '{flightNumber}'n'a pas été trouvé dans la liste.");

            }
            finally 
            {
                //Delete Flight
                var flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteFlightPage = flightPage.ClickMassiveDelete();
                massiveDeleteFlightPage.SetFlightName(flightNumber);
                massiveDeleteFlightPage.ClickSearchButton();
                massiveDeleteFlightPage.SelectFirstFlight();
                massiveDeleteFlightPage.SelectLastFlight();
                massiveDeleteFlightPage.Delete();
                //Delete LoadingPlan
                var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
                loadingPlanPage.MassiveDeleteLoadingPlan(loadingPlanName, null, null, fromDate.AddMonths(2));
                loadingPlanPage.ResetFilter();
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Site, site);
                loadingPlanPage.Filter(LoadingPlansPage.FilterType.Search, loadingPlanName);
                int totalNumber = loadingPlanPage.CheckTotalNumber();
                Assert.AreEqual(0, totalNumber, "La massive delete ne fonctionne pas.");
                //Delete LpCart
                lpCartPage = homePage.GoToFlights_LpCartPage();
                lpCartPage.ResetFilter();
                lpCartPage.Filter(LpCartPage.FilterType.Search, code);
                lpCartPage.DeleteLpCart();
                var isLpCartDeleted = lpCartPage.CheckTotalNumber() == 0;
                Assert.IsTrue(isLpCartDeleted, "Le lpCart n'est pas supprimé");
                servicePage = homePage.GoToCustomers_ServicePage();
                //Delete service2
                var serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.From, DateUtils.Now.AddMonths(-1).ToString("dd/MM/yyyy"));
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.To, DateUtils.Now.AddMonths(3).ToString("dd/MM/yyyy"));
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
                int total = servicePage.CheckTotalNumber();
                var isService2Deleted = total == 0;
                Assert.IsTrue(isService2Deleted, "La suppression d'un service ne fonctionne pas.");
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_FlightTab_SortByStartDate()
        {

            // prepare 
            string serviceName = "ServicName" + new Random().Next();
            string serviceCategory = TestContext.Properties["CategoryServiceDefault"].ToString();
            string site = TestContext.Properties["SiteLP"].ToString(); //ACE
            string customer = TestContext.Properties["CustomerLPFlight"].ToString();
            string customerLpCart = "$$ - CAT Genérico";
            DateTime dateFrom = DateTime.Now;
            DateTime dateTo = DateTime.Now.AddMonths(+12);

            string loadingPlanName = "lp" + new Random().Next();
            string type = "BuyOnBoard";
            string route = TestContext.Properties["RouteLP"].ToString();
            string guestName = TestContext.Properties["GuestNameBob"].ToString();
            string deleteFrom = DateUtils.Now.AddMonths(-1).ToString("dd/MM/yyyy");
            string deleteTo = DateUtils.Now.AddMonths(3).ToString("dd/MM/yyyy");
            string aircraft = TestContext.Properties["Registration"].ToString();
            DateTime lpDateTo = DateTime.Now;
            DateTime lpDateFrom = DateTime.Now.AddMonths(12);

            var code = "code"+ new Random().Next(0, 5000).ToString();
            var lpCartName = $"LP Cart {DateTime.Now}";
            var fromLpCart = DateTime.Now ;
            var toLpCart = DateTime.Today.AddMonths(+12);
            var comment = "comment";
            string flightNumber= "flighNumber" + new Random().Next();
            string flightNumber2 = "flighNumber2" + new Random().Next();
            string flightNumber3 = "flighNumber3" + new Random().Next();
            string siteTo = TestContext.Properties["SiteToFlightBob"].ToString();
            DateTime date1 = DateTime.Now;
            DateTime date2 = DateTime.Now.AddMonths(+1);
            DateTime date3 = DateTime.Now.AddMonths(+2);
            string etaHours = "00";
            string etdHours = "23";
            string filterBystartDate = "Date";

            //arrange
            var homePage = LogInAsAdmin();
            //act
            var servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();

            bool isFlightExist1 = false;
            bool isFlightExist2 = false;
            bool isFlightExist3 = false;
            bool isLpCartExist = false;
            bool isLoadingPlanExist = false;
            bool isServiceExist = false;
            try
            {
                // create a service
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, null, null, serviceCategory);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                isServiceExist = true;
                // add a price
                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, dateFrom, dateTo);
                servicePage = pricePage.BackToList();
                // create a loading plan
                var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
                loadingPlanPage.ResetFilter();
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customer, route, aircraft, site, type);
                loadingPlanCreateModalpage.FillFieldLoadingPlanInformations(lpDateFrom, lpDateTo);
                var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();
                loadingPlanDetailsPage.ClickAddGuestBtn();
                loadingPlanDetailsPage.SelectGuest(guestName);
                loadingPlanDetailsPage.ClickCreateGuestBtn();
                loadingPlanDetailsPage.ClickGuestBtnBOB(guestName);
                loadingPlanDetailsPage.AddServiceBtn();
                loadingPlanDetailsPage.AddNewService(serviceName);
                isLoadingPlanExist = true;
                // create a lpcart
                LpCartPage lpCartPage = homePage.GoToFlights_LpCartPage();
                lpCartPage.ResetFilter();
                var lpCartModalCreate = lpCartPage.LpCartCreatePage();
                LpCartCartDetailPage lpCartCartDetailPage = lpCartModalCreate
                    .FillField_CreateNewLpCartWithRoutes(code, lpCartName, site, customerLpCart, aircraft, fromLpCart, toLpCart, comment, route);
                isLpCartExist = true;
                // create 3 flights with different start date
                var flightPage = homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, site, siteTo, loadingPlanName, etaHours, etdHours, lpCartName,date1);
                isFlightExist1 = true;
                flightCreateModalpage = flightPage.FlightCreatePage();
                 flightCreateModalpage.FillField_CreatNewFlight(flightNumber2, customer, aircraft, site, siteTo, loadingPlanName, etaHours, etdHours, lpCartName, date2);
                isFlightExist2 = true;
                flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber3, customer, aircraft, site, siteTo, loadingPlanName, etaHours, etdHours, lpCartName, date3);
                isFlightExist3 = true;
                lpCartPage = flightPage.GoToFlights_LpCartPage();
                lpCartPage.ResetFilter();
                lpCartPage.Filter(LpCartPage.FilterType.Search, lpCartName);
                lpCartPage.WaitPageLoading();
                var detailPage = lpCartPage.ClickFirstLpCart();
                var lpCartFlightDetailPage = detailPage.ClickOnFlightTab();
                lpCartFlightDetailPage.Filter(LpCartFlightDetailPage.FilterType.SortBy, filterBystartDate);
                var isSortedByStartDate = lpCartFlightDetailPage.IsSortedByDate();
                Assert.IsTrue(isSortedByStartDate,"le filtre sur start date ne s'applique pas");


            }
            finally
            {
                //Delete Flight
                var flightPage = homePage.GoToFlights_FlightPage();
                FlightMassiveDeleteModalPage massiveDeleteFlightPage;
                if (isFlightExist1)
                {
                    massiveDeleteFlightPage = flightPage.ClickMassiveDelete();
                    massiveDeleteFlightPage.SetFlightName(flightNumber);
                    massiveDeleteFlightPage.ClickSearchButton();
                    massiveDeleteFlightPage.SelectFirstFlight();
                    massiveDeleteFlightPage.Delete();
                }
                if (isFlightExist2)
                {
                    massiveDeleteFlightPage = flightPage.ClickMassiveDelete();
                    massiveDeleteFlightPage.SetFlightName(flightNumber2);
                    massiveDeleteFlightPage.ClickSearchButton();
                    massiveDeleteFlightPage.SelectFirstFlight();
                    massiveDeleteFlightPage.Delete();
                }
                if (isFlightExist3)
                {
                    massiveDeleteFlightPage = flightPage.ClickMassiveDelete();
                    massiveDeleteFlightPage.SetFlightName(flightNumber3);
                    massiveDeleteFlightPage.ClickSearchButton();
                    massiveDeleteFlightPage.SelectFirstFlight();
                    massiveDeleteFlightPage.Delete();
                }

                //Delete LoadingPlan
                if(isLoadingPlanExist)
                {
                    var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
                    loadingPlanPage.MassiveDeleteLoadingPlan(loadingPlanName, null, null, lpDateTo.AddMonths(2));
                }

                //Delete LpCart
                if (isLpCartExist)
                {
                    var lpCartPage = homePage.GoToFlights_LpCartPage();
                    lpCartPage.ResetFilter();
                    lpCartPage.Filter(LpCartPage.FilterType.Search, lpCartName);
                    lpCartPage.DeleteLpCart();
                }

                //Delete service
                if (isServiceExist)
                {
                    servicePage = homePage.GoToCustomers_ServicePage();
                    var serviceMassiveDelete = servicePage.ClickMassiveDelete();
                    serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
                    serviceMassiveDelete.ClickSearchButton();
                    serviceMassiveDelete.DeleteFirstService();
                }
            }


        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_FlightTab_SortByFrom()
        {
            // prepare 
            string serviceName = "ServicName" + new Random().Next();
            string serviceCategory = TestContext.Properties["CategoryServiceDefault"].ToString();
            string site = TestContext.Properties["SiteLP"].ToString();
            string customer = TestContext.Properties["CustomerLPFlight"].ToString();
            string customerLpCart = "$$ - CAT Genérico";
            DateTime dateFrom = DateTime.Now;
            DateTime dateTo = DateTime.Now.AddMonths(+12);

            string loadingPlanName = "lp" + new Random().Next();
            string type = "BuyOnBoard";
            string route = TestContext.Properties["RouteLP"].ToString();
            string guestName = TestContext.Properties["GuestNameBob"].ToString();
            string deleteFrom = DateUtils.Now.AddMonths(-1).ToString("dd/MM/yyyy");
            string deleteTo = DateUtils.Now.AddMonths(3).ToString("dd/MM/yyyy");
            string aircraft = TestContext.Properties["Registration"].ToString();
            DateTime lpDateTo = DateTime.Now;
            DateTime lpDateFrom = DateTime.Now.AddMonths(12);

            var code = "code" + new Random().Next(0, 5000).ToString();
            var lpCartName = $"LP Cart {DateTime.Now}";
            var fromLpCart = DateTime.Now;
            var toLpCart = DateTime.Today.AddMonths(+12);
            var comment = "comment";
            string flightNumber = "flighNumber" + new Random().Next();
            string flightNumber2 = "flighNumber2" + new Random().Next();
            string flightNumber3 = "flighNumber3" + new Random().Next();
            string siteTo = TestContext.Properties["SiteToFlightBob"].ToString();
            DateTime date1 = DateTime.Now;
            string etaHours = "05";
            string etdHours = "23";
            string etaHours2 = "01";
            string etaHours3 = "02";
            string filterByFrom = "From";

            //arrange
            var homePage = LogInAsAdmin();
            //act
            var servicePage = homePage.GoToCustomers_ServicePage();
            try
            {
                // create a service
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, null, null, serviceCategory);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                // add a price
                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, dateFrom, dateTo);
                servicePage = pricePage.BackToList();
                // create a loading plan
                var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customer, route, aircraft, site, type);
                loadingPlanCreateModalpage.FillFieldLoadingPlanInformations(lpDateFrom, lpDateTo);
                var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();
                loadingPlanDetailsPage.ClickAddGuestBtn();
                loadingPlanDetailsPage.SelectGuest(guestName);
                loadingPlanDetailsPage.ClickCreateGuestBtn();
                loadingPlanDetailsPage.ClickGuestBtnBOB(guestName);
                loadingPlanDetailsPage.AddServiceBtn();
                loadingPlanDetailsPage.AddNewService(serviceName);
                // create a lpcart
                LpCartPage lpCartPage = homePage.GoToFlights_LpCartPage();
                var lpCartModalCreate = lpCartPage.LpCartCreatePage();
                LpCartCartDetailPage lpCartCartDetailPage = lpCartModalCreate
                    .FillField_CreateNewLpCartWithRoutes(code, lpCartName, site, customerLpCart, aircraft, fromLpCart, toLpCart, comment, route);
                // create 3 flights with different start date
                var flightPage = homePage.GoToFlights_FlightPage();
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, site, siteTo, loadingPlanName, etaHours, etdHours, lpCartName, date1);
                flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber2, customer, aircraft, site, siteTo, loadingPlanName, etaHours2, etdHours, lpCartName, date1);
                flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber3, customer, aircraft, site, siteTo, loadingPlanName, etaHours3, etdHours, lpCartName, date1);
                lpCartPage = flightPage.GoToFlights_LpCartPage();
                lpCartPage.ResetFilter();
                lpCartPage.Filter(LpCartPage.FilterType.Search, lpCartName);
                lpCartPage.WaitPageLoading();
                var detailPage = lpCartPage.ClickFirstLpCart();
                var lpCartFlightDetailPage = detailPage.ClickOnFlightTab();
                lpCartFlightDetailPage.Filter(LpCartFlightDetailPage.FilterType.SortBy, filterByFrom);
                lpCartFlightDetailPage.WaitPageLoading();
                var isSortedByFrom = lpCartFlightDetailPage.IsSortedByFrom();
                Assert.IsTrue(isSortedByFrom, "le filtre sur from date ne s'applique pas");


            }
            finally
            {
                //Delete Flight
                var flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteFlightPage = flightPage.ClickMassiveDelete();
                massiveDeleteFlightPage.SetFlightName(flightNumber);
                massiveDeleteFlightPage.ClickSearchButton();
                massiveDeleteFlightPage.SelectFirstFlight();
                massiveDeleteFlightPage.Delete();
                massiveDeleteFlightPage = flightPage.ClickMassiveDelete();
                massiveDeleteFlightPage.SetFlightName(flightNumber2);
                massiveDeleteFlightPage.ClickSearchButton();
                massiveDeleteFlightPage.SelectFirstFlight();
                massiveDeleteFlightPage.Delete();
                massiveDeleteFlightPage = flightPage.ClickMassiveDelete();
                massiveDeleteFlightPage.SetFlightName(flightNumber3);
                massiveDeleteFlightPage.ClickSearchButton();
                massiveDeleteFlightPage.SelectFirstFlight();
                massiveDeleteFlightPage.Delete();

                //Delete LoadingPlan
                var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
                loadingPlanPage.MassiveDeleteLoadingPlan(loadingPlanName, null, null, lpDateTo.AddMonths(2));

                //Delete LpCart
                var lpCartPage = homePage.GoToFlights_LpCartPage();
                lpCartPage.ResetFilter();
                lpCartPage.Filter(LpCartPage.FilterType.Search, lpCartName);
                lpCartPage.DeleteLpCart();

                //Delete service
                servicePage = homePage.GoToCustomers_ServicePage();
                var serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_FlightTab_SortByETD()
        {
            string serviceName = "ServicName" + new Random().Next();
            string serviceCategory = TestContext.Properties["CategoryServiceDefault"].ToString();
            string site = TestContext.Properties["SiteLP"].ToString();
            string customer = TestContext.Properties["CustomerLPFlight"].ToString();
            string customerLpCart = "$$ - CAT Genérico";
            DateTime dateFrom = DateTime.Now;
            DateTime dateTo = DateTime.Now.AddMonths(+12);

            string loadingPlanName = "lp" + new Random().Next();
            string type = "BuyOnBoard";
            string route = TestContext.Properties["RouteLP"].ToString();
            string guestName = TestContext.Properties["GuestNameBob"].ToString();
            string deleteFrom = DateUtils.Now.AddMonths(-1).ToString("dd/MM/yyyy");
            string deleteTo = DateUtils.Now.AddMonths(3).ToString("dd/MM/yyyy");
            string aircraft = TestContext.Properties["Registration"].ToString();
            DateTime lpDateTo = DateTime.Now;
            DateTime lpDateFrom = DateTime.Now.AddMonths(12);

            var code = "code" + new Random().Next(0, 5000).ToString();
            var lpCartName = $"LP Cart {DateTime.Now}";
            var fromLpCart = DateTime.Now;
            var toLpCart = DateTime.Today.AddMonths(+12);
            var comment = "comment";
            string flightNumber = "flighNumber" + new Random().Next();
            string flightNumber2 = "flighNumber2" + new Random().Next();
            string flightNumber3 = "flighNumber3" + new Random().Next();
            string siteTo = TestContext.Properties["SiteToFlightBob"].ToString();
            DateTime date1 = DateTime.Now;
            string etaHours = "05";
            string etdHours = "20";
            string etdHours2 = "22";
            string etdHours3 = "21";
            string filterByEtd = "ETD";

            //arrange
            var homePage = LogInAsAdmin();
            //act
            var servicePage = homePage.GoToCustomers_ServicePage();
            try
            {
                // create a service
                var serviceCreateModalPage = servicePage.ServiceCreatePage();
                serviceCreateModalPage.FillFields_CreateServiceModalPage(serviceName, null, null, serviceCategory);
                var serviceGeneralInformationsPage = serviceCreateModalPage.Create();
                // add a price
                var pricePage = serviceGeneralInformationsPage.GoToPricePage();
                var priceModalPage = pricePage.AddNewCustomerPrice();
                pricePage = priceModalPage.FillFields_CustomerPrice(site, customer, dateFrom, dateTo);
                servicePage = pricePage.BackToList();
                // create a loading plan
                var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
                var loadingPlanCreateModalpage = loadingPlanPage.LoadingPlansCreatePage();
                loadingPlanCreateModalpage.FillField_CreateNewLoadingPlan(loadingPlanName, customer, route, aircraft, site, type);
                loadingPlanCreateModalpage.FillFieldLoadingPlanInformations(lpDateFrom, lpDateTo);
                var loadingPlanDetailsPage = loadingPlanCreateModalpage.Create();
                loadingPlanDetailsPage.ClickAddGuestBtn();
                loadingPlanDetailsPage.SelectGuest(guestName);
                loadingPlanDetailsPage.ClickCreateGuestBtn();
                loadingPlanDetailsPage.ClickGuestBtnBOB(guestName);
                loadingPlanDetailsPage.AddServiceBtn();
                loadingPlanDetailsPage.AddNewService(serviceName);
                // create a lpcart
                LpCartPage lpCartPage = homePage.GoToFlights_LpCartPage();
                var lpCartModalCreate = lpCartPage.LpCartCreatePage();
                LpCartCartDetailPage lpCartCartDetailPage = lpCartModalCreate
                    .FillField_CreateNewLpCartWithRoutes(code, lpCartName, site, customerLpCart, aircraft, fromLpCart, toLpCart, comment, route);
                // create 3 flights with different start date
                var flightPage = homePage.GoToFlights_FlightPage();
                var flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber, customer, aircraft, site, siteTo, loadingPlanName, etaHours, etdHours, lpCartName, date1);
                flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber2, customer, aircraft, site, siteTo, loadingPlanName, etaHours, etdHours2, lpCartName, date1);
                flightCreateModalpage = flightPage.FlightCreatePage();
                flightCreateModalpage.FillField_CreatNewFlight(flightNumber3, customer, aircraft, site, siteTo, loadingPlanName, etaHours, etdHours3, lpCartName, date1);
                lpCartPage = flightPage.GoToFlights_LpCartPage();
                lpCartPage.ResetFilter();
                lpCartPage.Filter(LpCartPage.FilterType.Search, lpCartName);
                lpCartPage.WaitPageLoading();
                var detailPage = lpCartPage.ClickFirstLpCart();
                var lpCartFlightDetailPage = detailPage.ClickOnFlightTab();
                lpCartFlightDetailPage.Filter(LpCartFlightDetailPage.FilterType.SortBy, filterByEtd);
                lpCartFlightDetailPage.WaitPageLoading();
                var isSortedByEtd = lpCartFlightDetailPage.IsSortedByETD();
                Assert.IsTrue(isSortedByEtd, "le filtre sur etd date ne s'applique pas");
            }
            finally
            {
                //Delete Flight
                var flightPage = homePage.GoToFlights_FlightPage();
                var massiveDeleteFlightPage = flightPage.ClickMassiveDelete();
                massiveDeleteFlightPage.SetFlightName(flightNumber);
                massiveDeleteFlightPage.ClickSearchButton();
                massiveDeleteFlightPage.SelectFirstFlight();
                massiveDeleteFlightPage.Delete();
                massiveDeleteFlightPage = flightPage.ClickMassiveDelete();
                massiveDeleteFlightPage.SetFlightName(flightNumber2);
                massiveDeleteFlightPage.ClickSearchButton();
                massiveDeleteFlightPage.SelectFirstFlight();
                massiveDeleteFlightPage.Delete();
                massiveDeleteFlightPage = flightPage.ClickMassiveDelete();
                massiveDeleteFlightPage.SetFlightName(flightNumber3);
                massiveDeleteFlightPage.ClickSearchButton();
                massiveDeleteFlightPage.SelectFirstFlight();
                massiveDeleteFlightPage.Delete();

                //Delete LoadingPlan
                var loadingPlanPage = homePage.GoToFlights_LoadingPlansPage();
                loadingPlanPage.MassiveDeleteLoadingPlan(loadingPlanName, null, null, lpDateTo.AddMonths(2));

                //Delete LpCart
                var lpCartPage = homePage.GoToFlights_LpCartPage();
                lpCartPage.ResetFilter();
                lpCartPage.Filter(LpCartPage.FilterType.Search, lpCartName);
                lpCartPage.DeleteLpCart();

                //Delete service
                servicePage = homePage.GoToCustomers_ServicePage();
                var serviceMassiveDelete = servicePage.ClickMassiveDelete();
                serviceMassiveDelete.Filter(ServiceMassiveDeleteFilterType.ServiceName, serviceName);
                serviceMassiveDelete.ClickSearchButton();
                serviceMassiveDelete.DeleteFirstService();
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_EditScheme()
        {
            //// Prepare
            string name = TestContext.Properties["LpCartName4"].ToString();
            string trolleyName = TestContext.Properties["TrolleyName"].ToString();
            string trolleyScheme = TestContext.Properties["TrolleySchemeName"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var lpCartPage = homePage.GoToFlights_LpCartPage();
            lpCartPage.ResetFilter();

            lpCartPage.Filter(LpCartPage.FilterType.Search, name);
            var detailCartPage = lpCartPage.LpCartCartDetailPage();
            detailCartPage.Filter(LpCartCartDetailPage.FilterType.Search, trolleyName);
            if (lpCartPage.CheckTotalNumber() == 0)
            {
                detailCartPage.ClickAddtrolley();
                lpCartPage.okWarning();
                detailCartPage.AddTrolley(trolleyName);
                var lpCartCreateSchemeModal = lpCartPage.LpCartCreateLpCartSchemeModal();
                lpCartCreateSchemeModal.CreateLpCartscheme(trolleyScheme, "2", "2");
            }
            var lpCartSchemeModal = lpCartPage.LpCartEditLpCartSchemeModal();
            lpCartSchemeModal.EditLpCartscheme("Update", "1", "1", "1", "1", "");

            detailCartPage.Filter(LpCartCartDetailPage.FilterType.Search, trolleyName);
            lpCartPage.LpCartEditLpCartSchemeModal();
            var result = lpCartSchemeModal.GetLpCartschemeValues();

            //Assert
            Assert.IsTrue(result.Contains("Update"), "Le Scheme n'a pas été modifié correctement");
            Assert.IsTrue(result.Contains("1"), "Le Scheme n'a pas été modifié correctement");
            Assert.IsTrue(!result.Contains(trolleyScheme), "Le Scheme n'a pas été modifié correctement");

        }

        // Test de duplication de LpCart
        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_DuplicateLpCart()
        {
            Random rnd = new Random();

            // Préparation
            string lpCartName = "lpCartTest" + rnd.Next().ToString();
            string nameDuplicate = lpCartName + " - Dup";

            // Données pour créer le LP cart
            string code3 = rnd.Next(0, 500000).ToString();
            string code = rnd.Next(0, 500000).ToString();
            string site3 = TestContext.Properties["SiteLpCart2"].ToString();
            string customer3 = TestContext.Properties["CustomerLpCart2"].ToString();
            string aircraft3 = TestContext.Properties["AircraftLpCart"].ToString();
            DateTime from3 = DateUtils.Now.AddDays(1);
            DateTime to3 = DateUtils.Now.AddDays(10);
            string comment3 = TestContext.Properties["CommentLpCart"].ToString();
            string trolleyName = TestContext.Properties["TrolleyName"].ToString();
            string trolleyScheme = TestContext.Properties["TrolleySchemeName"].ToString();
            string column = "2";
            string rows = "2";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act - Accéder à la page de LpCart et réinitialiser le filtre
            var lpCartPage = homePage.GoToFlights_LpCartPage();
            lpCartPage.ResetFilter();

            // Création du LP cart
            var lpCartModalCreate = lpCartPage.LpCartCreatePage();
            var lpCartCartDetailPage = lpCartModalCreate.FillField_CreateNewLpCart(code3, lpCartName, site3, customer3, aircraft3, from3, to3, comment3);

            // Ajouter un trolley
            lpCartCartDetailPage.ClickAddtrolley();
            lpCartCartDetailPage.AddTrolley(trolleyName);

            // Création du schéma de LpCart
            var lpCartSchemeModal = lpCartPage.LpCartCreateLpCartSchemeModal();
            lpCartSchemeModal.CreateLpCartscheme(trolleyScheme, rows, column);

            var namesTrolley = lpCartCartDetailPage.GetAllTrolleyNames();

            // Retourner à la liste et appliquer le filtre pour rechercher le LpCart créé
            lpCartCartDetailPage.BackToList();
            lpCartPage.Filter(LpCartPage.FilterType.Search, lpCartName);

            // Dupliquer le LpCart
            var duplicateLpCartModal = lpCartPage.DuplicateLpCart();
            duplicateLpCartModal.SetValuesForDuplication(lpCartName, code, nameDuplicate, site3, customer3, aircraft3);

            // Vérifier que le nombre total de LpCart a doublé
            var totalDuplicate = lpCartPage.CheckTotalNumber();
            Assert.IsTrue(totalDuplicate == 2, "Le LpCart n'a pas été dupliqué correctement.");

            // Filtrer par le nom du duplicata et sélectionner le premier LpCart
            lpCartPage.Filter(LpCartPage.FilterType.Search, nameDuplicate);
            lpCartPage.ClickFirstLpCart();

            // Vérifier que les noms de trolley sont identiques entre l'original et le duplicata
            var namesTrolleyDup = lpCartCartDetailPage.GetAllTrolleyNames();
            CollectionAssert.AreEqual(namesTrolley, namesTrolleyDup, "Les noms de trolleys ne sont pas identiques entre l'original et le duplicata.");

            //vérifier les valeurs de lignes et colonnes
            var schemaLineVal = lpCartPage.GetSchemaLineCount();
            var schemaColumnLine = lpCartPage.GetSchemeColumnCount();

            Assert.AreEqual(schemaLineVal, rows, "La valeur de 'rows' du schéma n'est pas correcte dans le duplicata.");
            Assert.AreEqual(schemaColumnLine, column, "La valeur de 'columns' du schéma n'est pas correcte dans le duplicata.");
        }

        //duplicate Trolley 
        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_DuplicateTrolley()
        {
            // Préparation
            Random rnd = new Random();

            string lpCartName = "lpCartTest" + rnd.Next().ToString();
            string nameDuplicate = lpCartName + " - Dup";

            // Données pour créer le LP cart
            string code3 = rnd.Next(0, 500000).ToString();
            string code = rnd.Next(0, 500000).ToString();
            string site3 = TestContext.Properties["SiteLpCart2"].ToString();
            string customer3 = TestContext.Properties["CustomerLpCart2"].ToString();
            string aircraft3 = TestContext.Properties["AircraftLpCart"].ToString();
            DateTime from3 = DateUtils.Now.AddDays(1);
            DateTime to3 = DateUtils.Now.AddDays(10);
            string comment3 = TestContext.Properties["CommentLpCart"].ToString();
            string trolleyName = TestContext.Properties["TrolleyName"].ToString();
            string trolleyScheme = TestContext.Properties["TrolleySchemeName"].ToString();
            string column = "2";
            string rows = "2";

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act 
            var lpCartPage = homePage.GoToFlights_LpCartPage();
            lpCartPage.ResetFilter();

            // Création du LP cart
            var lpCartModalCreate = lpCartPage.LpCartCreatePage();
            var lpCartCartDetailPage = lpCartModalCreate.FillField_CreateNewLpCart(code3, lpCartName, site3, customer3, aircraft3, from3, to3, comment3);

            // Ajouter un trolley
            lpCartCartDetailPage.ClickAddtrolley();
            lpCartCartDetailPage.AddTrolley(trolleyName);

            // Création du schéma de LpCart
            var lpCartSchemeModal = lpCartPage.LpCartCreateLpCartSchemeModal();
            lpCartSchemeModal.CreateLpCartscheme(trolleyScheme, rows, column);

            var namesTrolley = lpCartCartDetailPage.GetAllTrolleyNames();

            // Retourner à la liste et appliquer le filtre pour rechercher le LpCart créé
            lpCartCartDetailPage.BackToList();
            lpCartPage.Filter(LpCartPage.FilterType.Search, lpCartName);

            // Dupliquer le LpCart
            var duplicateLpCartModal = lpCartPage.DuplicateLpCart();
            duplicateLpCartModal.SetValuesForDuplication(lpCartName, code, nameDuplicate, site3, customer3, aircraft3);

            // Vérifier que le nombre total de LpCart a doublé
            var totalDuplicate = lpCartPage.CheckTotalNumber();
            Assert.IsTrue(totalDuplicate == 2, "Le LpCart n'a pas été dupliqué correctement.");

            // Filtrer par le nom du duplicata et sélectionner le premier LpCart
            lpCartPage.Filter(LpCartPage.FilterType.Search, nameDuplicate);
            lpCartPage.ClickFirstLpCart();

            // Vérifier que les noms de trolley sont identiques entre l'original et le duplicata
            var namesTrolleyDup = lpCartCartDetailPage.GetAllTrolleyNames();
            CollectionAssert.AreEqual(namesTrolley, namesTrolleyDup, "Les noms de trolleys ne sont pas identiques entre l'original et le duplicata.");

            // Vérifier les valeurs de lignes et colonnes
            var schemaLineVal = lpCartPage.GetSchemaLineCount();
            var schemaColumnLine = lpCartPage.GetSchemeColumnCount();
            var beforeNumber = lpCartPage.CheckTotalNumber();

            // dupliquer le Trolley
            var duplicateTrolleyModal = lpCartCartDetailPage.DuplicateTrolley();
            duplicateTrolleyModal.SetValuesForDuplication(code);

            // Assert
            Assert.IsTrue(lpCartPage.CheckTotalNumber() == beforeNumber + 1, "Le CartTrolley n'a pas été dupliqué");
            Assert.AreEqual(schemaLineVal, rows, "La valeur de 'rows' du schéma n'est pas correcte dans le duplicata.");
            Assert.AreEqual(schemaColumnLine, column, "La valeur de 'columns' du schéma n'est pas correcte dans le duplicata.");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_ChangeOrderTrolley()
        {

            //// Prepare
            string name = TestContext.Properties["LpCartName4"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var lpCartPage = homePage.GoToFlights_LpCartPage();
            lpCartPage.ResetFilter();

            lpCartPage.Filter(LpCartPage.FilterType.Search, name);
            var detailCartPage = lpCartPage.LpCartCartDetailPage();

            if (lpCartPage.CheckTotalNumber() == 1)
            {
                detailCartPage.ClickAddtrolley();
                lpCartPage.okWarning();
                detailCartPage.AddTrolley(name);
            }
            var firstList = detailCartPage.GetAllTrolleyNames();

            detailCartPage.ClickFirstArrowDown();

            var secondList = detailCartPage.GetAllTrolleyNames();

            //Assert
            Assert.IsFalse(firstList.SequenceEqual(secondList), "L'ordre des trolley ne s'est pas effectué");
            Assert.IsTrue(firstList.OrderBy(x => x).SequenceEqual(secondList.OrderBy(x => x)), "L'ordre des trolley ne s'est pas effectué");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_FusionTrolley()
        {
            // Prepare
            string trolley = "trolley"+ new Random().Next(1,50000).ToString();
            string trolleyEdtited = "TestTrolley";
            int nombreDrawer = 4;
            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var lpCartPage = homePage.GoToFlights_LpCartPage();
            lpCartPage.ResetFilter();
            // be on the lpcart created
            lpCartPage.Filter(LpCartPage.FilterType.Search, lpCartName);
            var detailCartPage = lpCartPage.LpCartCartDetailPage();
            //add a trolley
            detailCartPage.ClickAddtrolley();
            detailCartPage.AddTrolley(trolley);
            //Create LpCart Scheme with 4 drawers (2 ligne 2 column)
            LpCartCreateLpCartSchemeModal lpCartSchemeModal = lpCartPage.LpCartCreateLpCartSchemeModal();
            lpCartSchemeModal.CreateLpCartscheme(trolley, "2", "2");
            // check if th drawers are created
            LpCartEditLpCartSchemeModal lpCartEditSchemeModal = lpCartPage.LpCartEditLpCartSchemeModal();
            int drawerGenerated = lpCartEditSchemeModal.GetDrawerCount();
            Assert.AreEqual(nombreDrawer, drawerGenerated, "les drawers générés ne sont pas 4");
            // fusionner drawers
            lpCartEditSchemeModal.ActivateFusion();
            lpCartEditSchemeModal.SelectFusionElement();
            lpCartEditSchemeModal.ClickFusion();
            lpCartEditSchemeModal = lpCartPage.LpCartEditLpCartSchemeModal();
            drawerGenerated = lpCartEditSchemeModal.GetDrawerCount();
            nombreDrawer = 1;
            Assert.AreEqual(nombreDrawer, drawerGenerated, "la fusion des drawers ne s'applique pas");
            // Renseigner les données 
            lpCartEditSchemeModal.EditDrawInputs(trolleyEdtited);
            lpCartEditSchemeModal.ClickConfirm();
            lpCartEditSchemeModal = lpCartPage.LpCartEditLpCartSchemeModal();
            bool isDrawerEdited = lpCartEditSchemeModal.IsDrawersEdited(trolleyEdtited);
            Assert.IsTrue(isDrawerEdited,"drawers n'a pas modifié");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_UndoTrolley()
        {
            // Prepare
            string trolley = "trolley" + new Random().Next(1, 50000).ToString();
            string trolleyEdtited = "TestTrolley";
            int nombreDrawer = 4;
            string column = "2";
            string row = "2";
            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var lpCartPage = homePage.GoToFlights_LpCartPage();
            lpCartPage.ResetFilter();
            // be on the lpcart created
            lpCartPage.Filter(LpCartPage.FilterType.Search, lpCartName);
            var detailCartPage = lpCartPage.LpCartCartDetailPage();
            //add a trolley
            detailCartPage.ClickAddtrolley();
            detailCartPage.AddTrolley(trolley);
            //Create LpCart Scheme with 4 drawers (2 ligne 2 column)
            LpCartCreateLpCartSchemeModal lpCartSchemeModal = lpCartPage.LpCartCreateLpCartSchemeModal();
            lpCartSchemeModal.CreateLpCartschemeWithoutDrawes(trolley);
            // check if th drawers are created
            LpCartEditLpCartSchemeModal lpCartEditSchemeModal = lpCartPage.LpCartEditLpCartSchemeModal();
            lpCartEditSchemeModal.EditLpCartschemeWithoutConfirm(trolley, trolley, trolley, row, column, trolley);
            int drawerGenerated = lpCartEditSchemeModal.GetDrawerCount();
            Assert.AreEqual(nombreDrawer, drawerGenerated, "les drawers générés ne sont pas 4");
            // fusionner drawers
            lpCartEditSchemeModal.ActivateFusion();
            lpCartEditSchemeModal.SelectFusionElement();
            lpCartEditSchemeModal.ClickFusionWithoutConfirm();
            lpCartEditSchemeModal.ClickDesactivateFusion();
            drawerGenerated = lpCartEditSchemeModal.GetDrawerCount();
            nombreDrawer = 1;
            Assert.AreEqual(nombreDrawer, drawerGenerated, "la fusion des drawers ne s'applique pas");
            // Renseigner les données 
            lpCartEditSchemeModal.EditDrawInputs(trolleyEdtited);
            // defusionner les drawers
            lpCartEditSchemeModal.ActivateFusion();
            lpCartEditSchemeModal.SelectFusionElement();
            lpCartEditSchemeModal.ClickUndoFusionWithoutConfirm();
            lpCartEditSchemeModal.ClickDesactivateFusion();
            lpCartEditSchemeModal.Insert_Postion_PP_Line(row, column);
            lpCartEditSchemeModal = lpCartPage.LpCartEditLpCartSchemeModal();
            drawerGenerated = lpCartEditSchemeModal.GetDrawerCount();
            nombreDrawer = 4;
            Assert.AreEqual(nombreDrawer, drawerGenerated, "undofusion des drawers ne s'applique pas");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_Export()
        {
            //// Prepare
            string lpCartYacamil = TestContext.Properties["LpCartName4"].ToString();
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            bool newVersion = true;

            //Data to create LPcart yacamil
            Random rnd2 = new Random();
            string code3 = rnd2.Next(0, 500000).ToString();
            string site3 = TestContext.Properties["SiteLpCart2"].ToString();
            string customer3 = TestContext.Properties["CustomerLpCart2"].ToString();
            string aircraft3 = TestContext.Properties["AircraftLpCart"].ToString();
            DateTime from3 = DateUtils.Now.AddDays(1);
            DateTime to3 = DateUtils.Now.AddDays(10);
            string comment3 = TestContext.Properties["CommentLpCart"].ToString();
            string trolleyName = TestContext.Properties["TrolleyName"].ToString();
            string trolleyName1 = TestContext.Properties["TrolleyName1"].ToString();
            string trolleyScheme = TestContext.Properties["TrolleySchemeName"].ToString();
            string trolleyScheme1 = TestContext.Properties["TrolleySchemeName1"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var lpCartPage = homePage.GoToFlights_LpCartPage();
            lpCartPage.ResetFilter();



            lpCartPage.Filter(LpCartPage.FilterType.Search, lpCartYacamil);
            if (lpCartPage.CheckTotalNumber() == 0)
            {
                // Create Yacamil


                var lpCartModalCreate = lpCartPage.LpCartCreatePage();
                var lpCartCartDetailPage = lpCartModalCreate.FillField_CreateNewLpCart(code3, lpCartYacamil, site3, customer3, aircraft3, from3, to3, comment3);

                //Acces Onglet Cart                
                lpCartCartDetailPage.ClickAddtrolley();
                lpCartCartDetailPage.AddTrolley(trolleyName);

                //Create LpCart Scheme
                var lpCartSchemeModal = lpCartPage.LpCartCreateLpCartSchemeModal();
                lpCartSchemeModal.CreateLpCartscheme(trolleyScheme, "2", "2");

                lpCartCartDetailPage.ClickAddtrolley();
                lpCartCartDetailPage.AddTrolley(trolleyName1);

                //Create LpCart Scheme
                lpCartPage.LpCartCreateLpCartSchemeModal();
                lpCartSchemeModal.CreateLpCartscheme(trolleyScheme1, "2", "2");

                lpCartCartDetailPage.BackToList();
            }
            DeleteAllFileDownload();

            lpCartPage.ClearDownloads();

            lpCartPage.ExportLpcart(ExportType.Export, newVersion);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = lpCartPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            int resultNumber = OpenXmlExcel.GetExportResultNumber("LPCarts", filePath);
            var listResult = OpenXmlExcel.GetValuesInList("Name", "LPCarts", filePath);

            //Assert
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.IsFalse(!listResult.Contains(lpCartYacamil), MessageErreur.EXCEL_DONNEES_KO);

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_Import()
        {
            //ATTENTION, fichier excel à revoir à chaque changement d'année
            // Prepare
            string name = TestContext.Properties["LpCartName5"].ToString();

            // Répertoire du fichier à importer
            var path = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase).Substring(6);
            var excelPath = path + "\\PageObjects\\Flights\\LpCart\\LPCarts_test.xlsx";
            Assert.IsTrue(new FileInfo(excelPath).Exists, "Fichier " + excelPath + "non trouvé");

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var lpCartPage = homePage.GoToFlights_LpCartPage();
            lpCartPage.ResetFilter();

            try
            {
                lpCartPage.Import(excelPath);

                lpCartPage.Filter(LpCartPage.FilterType.Search, name);

                //Assert
                Assert.AreEqual(1, lpCartPage.CheckTotalNumber(), "L'import du fichier ne s'est pas executé");
                Assert.AreEqual(lpCartPage.GetFirstLpCartName(), name, "L'import du fichier ne s'est pas éxecuté");
            }
            finally
            {
                lpCartPage.DeleteLpCart();
            }

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_FilterKeyword()
        {
            // Prepare
            string name = TestContext.Properties["LpCartName4"].ToString();
            string trolleyName = TestContext.Properties["TrolleyName1"].ToString();


            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var lpCartPage = homePage.GoToFlights_LpCartPage();
            lpCartPage.ResetFilter();

            lpCartPage.Filter(LpCartPage.FilterType.Search, name);

            var detailCartPage = lpCartPage.LpCartCartDetailPage();
            var duplicateTrolleyModal = detailCartPage.DuplicateTrolley();

            duplicateTrolleyModal.SetKeywordFilter(trolleyName);

            //Assert
            Assert.IsTrue(duplicateTrolleyModal.IsGoodValueFilterKeyword(trolleyName), String.Format(MessageErreur.FILTRE_ERRONE, "Keyword"));

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_ResetFilter()
        {
            string name = TestContext.Properties["LpCartName4"].ToString();
            HomePage homePage = LogInAsAdmin();
            LpCartPage lpCartPage = homePage.GoToFlights_LpCartPage();
            lpCartPage.ResetFilter();
            int initNumber = lpCartPage.CheckTotalNumber();
            lpCartPage.Filter(LpCartPage.FilterType.Search, name);
            Assert.IsTrue(lpCartPage.CheckTotalNumber() == 1, "Le filtre Search ne marche.");
            lpCartPage.ResetFilter();
            int totalNumber = lpCartPage.CheckTotalNumber();
            Assert.AreNotEqual(1, lpCartPage.CheckTotalNumber(), "Le resetFilter ne fonctionne pas correctement.");
            totalNumber = lpCartPage.CheckTotalNumber();
            Assert.AreEqual(initNumber, totalNumber, MessageErreur.FILTRE_ERRONE, "Reset filter");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_Carts_SchemeColors()
        {
            #region data
            //data
            string lpCartName = "Calomato";

            //colors names
            string color1 = "Black";
            string color2 = "Red";
            string color3 = "Orange";
            string color4 = "darkgray";

            //colors rgb 

            int color1R = 0;
            int color1G = 0;
            int color1B = 0;

            int color2R = 255;
            int color2G = 0;
            int color2B = 0;

            
            int color3R = 255;
            int color3G = 165;
            int color3B = 0;

            int color4R = 169;
            int color4G = 169;
            int color4B = 169;
            #endregion
            #region process
            //login 
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            homePage.ClearDownloads();

            if (homePage.IsDev())
            {
                // Orange sombre, plus professionnel.
                color3G = 69;
            }

            var lpCartPage = homePage.GoToFlights_LpCartPage();
            lpCartPage.Filter(FilterType.Search, lpCartName);
            Assert.IsTrue(lpCartPage.VerifyLpCartExist(lpCartName), "LpCart n'existe pas");
            var lpCartDetailsPage = lpCartPage.ClickFirstLpCart();
            var editCartSchemePage = lpCartDetailsPage.EditCartScheme();
            if (!editCartSchemePage.ColorsTableIsVisible())
            {
                editCartSchemePage.GenerateRowsColumns(2, 2);
            }

            editCartSchemePage.ClickChooseLabelColors();
            editCartSchemePage.ChangeColors(color1, color2, color3, color4);
            var colorsByOrderObjects = editCartSchemePage.GetColors();
            editCartSchemePage = lpCartDetailsPage.EditCartScheme();
            editCartSchemePage.ClickChooseLabelColors();
            var confirmColorsbyOrdersObjects = editCartSchemePage.GetColors();
            //verif colors are updated
            Assert.AreEqual(colorsByOrderObjects.ToString(), confirmColorsbyOrdersObjects.ToString(), "Colors not updated");

            //get flight details attached to lp
            lpCartDetailsPage.FlightsTab();
            var firstFlightNumber = lpCartDetailsPage.GetFirstFlightNumber();
            var firstFlightDate = lpCartDetailsPage.GetFirstFlightDate();
            var firstFlightSiteFrom = lpCartDetailsPage.GetFirstFlightSiteFrom();
            var firstFlightSiteTo = lpCartDetailsPage.GetFirstFlightSiteTo();
            var flightPage = homePage.GoToFlights_FlightPage();
            //filter
            flightPage.Filter(FlightPage.FilterType.SearchFlight, firstFlightNumber);
            flightPage.SetDateState(DateTime.ParseExact(firstFlightDate, "dd/MM/yyyy", null));
            flightPage.Filter(FlightPage.FilterType.Sites, firstFlightSiteFrom);
            if (!flightPage.IsFlightExist())
            {
                flightPage.Filter(FlightPage.FilterType.Sites, firstFlightSiteTo);
                flightPage.Filter(FlightPage.FilterType.DispalyFlight, "Display flight arrival");
            }

            //print pdf
            var printPage = flightPage.PrintReport(FlightPage.PrintType.FlightsLabels, true);

            //take screenshot
            var imgName = printPage.TakeScreenshot();
            var color1Exist = printPage.VerifyRGBExistInImg(imgName, color1R, color1G, color1B);
            var color2Exist = printPage.VerifyRGBExistInImg(imgName, color2R, color2G, color2B);
            var color3Exist = printPage.VerifyRGBExistInImg(imgName, color3R, color3G, color3B);
            var color4Exist = printPage.VerifyRGBExistInImg(imgName, color4R, color4G, color4B);
            #endregion
            #region assertion
            Assert.IsTrue(color1Exist, color1 + "n'existe pas dans le pdf");
            Assert.IsTrue(color2Exist, color2 + "n'existe pas dans le pdf");
            Assert.IsTrue(color3Exist, color3 + "n'existe pas dans le pdf");
            Assert.IsTrue(color4Exist, color4 + "n'existe pas dans le pdf");
            #endregion
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_Carts_SchemeDetails()
        {
            #region data
            string lpCartName = "Calomato";
            string position = "NEWPOSITION";
            string pandp = "NEWP&P";
            string positionDetail = "NEWPOSITIONDETAIL";
            string DocFileNamePdfBegin = "Flights Print Label";
            string DocFileNameZipBegin = "All_files_";
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            #endregion
            #region process
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var lpCartPage = homePage.GoToFlights_LpCartPage();
            lpCartPage.ResetFilter();
            lpCartPage.Filter(FilterType.Search, lpCartName);
            bool lpCartExist = lpCartPage.VerifyLpCartExist(lpCartName);
            var lpCartDetailsPage = lpCartPage.ClickFirstLpCart();
            var editlpCart = lpCartDetailsPage.EditCartScheme();
            editlpCart.EditFirstPosition(position, pandp, positionDetail);
            editlpCart = lpCartDetailsPage.EditCartScheme();
            bool verifDataChanged = editlpCart.VerifyPosition(position, pandp, positionDetail);

            lpCartDetailsPage.FlightsTab();
            var firstFlightNumber = lpCartDetailsPage.GetFirstFlightNumber();
            var firstFlightDate = lpCartDetailsPage.GetFirstFlightDate();
            var firstFlightSite = lpCartDetailsPage.GetFirstFlightSiteFrom();
            var firstFlightSiteTo = lpCartDetailsPage.GetFirstFlightSiteTo();
            var flightPage = homePage.GoToFlights_FlightPage();

            flightPage.Filter(FlightPage.FilterType.SearchFlight, firstFlightNumber);
            flightPage.SetDateState(DateTime.ParseExact(firstFlightDate, "dd/MM/yyyy", null));
            flightPage.Filter(FlightPage.FilterType.Sites, firstFlightSiteTo);
            DeleteAllFileDownload();
            flightPage.ClearDownloads();

            //generate pdf file
            var printPage = flightPage.PrintReport(FlightPage.PrintType.FlightsLabels, true);
            var isReportGenerated = printPage.IsReportGenerated();
            printPage.Close();
            //download pdf file

            printPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            string trouve = printPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            FileInfo fi = new FileInfo(trouve);
            fi.Refresh();
            var fileIndexExist = fi.Exists;
            var allWordsExist = printPage.VerifyNewPositions(fi, pandp, positionDetail);
            printPage.ClosePrintButton();
            flightPage.Filter(FlightPage.FilterType.SearchFlight, firstFlightNumber);
            flightPage.Filter(FlightPage.FilterType.Sites, firstFlightSiteTo);
            var flightDetailsPage = flightPage.EditFirstFlight(firstFlightNumber);
            printPage = flightDetailsPage.PrintFlightLabels();
            //generate second pdf file
            var issecondReportGenerated = printPage.IsReportGenerated();
            printPage.Close();
            //download second pdf file
            printPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            printPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            FileInfo secondfi = new FileInfo(trouve);
            secondfi.Refresh();
            var fileflightExist = secondfi.Exists;
            var allWordssecondExist = printPage.VerifyNewPositions(fi, pandp, positionDetail);
            #endregion
            #region assertion
            Assert.IsTrue(lpCartExist, "LpCart n'existe pas");
            Assert.IsTrue(verifDataChanged, "données ne sont pas changés");
            Assert.IsTrue(isReportGenerated, "pdf file non généré");
            Assert.IsTrue(fileIndexExist, trouve + " non généré");
            Assert.IsTrue(allWordsExist, "position , P&P et position detail ne sont pas présents dans le pdf");
            Assert.IsTrue(issecondReportGenerated, "deuxième pdf file non généré");
            Assert.IsTrue(fileflightExist, trouve + " non généré");
            Assert.IsTrue(allWordssecondExist, "position , P&P et position detail ne sont pas présents dans le deuxième pdf");
            #endregion
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_Filter_SortByOnCartsTab()
        {
            var filterByLegOrder = "LegAndOrder";
            var filterByCode = "Code";
            var filterByName = "Name";
            // Prepare
            //lpCart
            Random rnd = new Random();
            string name = TestContext.Properties["LpCartNamebob"].ToString();
            string code = TestContext.Properties["LpCartCodebob"].ToString() + rnd.Next().ToString();
            string nameBis = TestContext.Properties["LpCartNameBis"].ToString();
            string codeBis = TestContext.Properties["LpCartCodeBis"].ToString();
            string site = TestContext.Properties["SiteLpCart"].ToString();
            string customer = TestContext.Properties["Bob_CustomerFilter"].ToString();
            string aircraft = "AB310";
            //trolley
            string trolleyName = "troll";

            DateTime from = DateUtils.Now.AddDays(5);
            DateTime to = DateUtils.Now.AddDays(10);
            string comment = "Bob comment";

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            var lpCartPage = homePage.GoToFlights_LpCartPage();
            var lpCartCreateModalPage = lpCartPage.LpCartCreatePage();
            var LpCartDetailPage = lpCartCreateModalPage.FillField_CreateNewLpCart(code, name, site, customer, aircraft, from, to, comment);
            for (int i = 0; i < 3; i++)
            {
                LpCartDetailPage.WaitPageLoading();
                LpCartDetailPage.ClickAddtrolley();
                LpCartDetailPage.WaitPageLoading();
                LpCartDetailPage.AddTrolley(trolleyName + i);
                LpCartDetailPage.WaitPageLoading();
            }
            LpCartDetailPage.Filter(LpCartCartDetailPage.FilterType.SortBy, filterByLegOrder);
            var verifyFilterByLegOrder = LpCartDetailPage.VerifyFilterByLegOrder();
            LpCartDetailPage.Filter(LpCartCartDetailPage.FilterType.SortBy, filterByCode);
            var verifyFilterByCode = LpCartDetailPage.VerifyFilterBycode();
            LpCartDetailPage.Filter(LpCartCartDetailPage.FilterType.SortBy, filterByName);
            var verifyFilterByName = LpCartDetailPage.VerifyFilterByName();
            Assert.IsTrue(verifyFilterByLegOrder, "erreur de filtrage by leg order");
            Assert.IsTrue(verifyFilterByCode, "erreur de filtrage by code");
            Assert.IsTrue(verifyFilterByName, "erreur de filtrage by name");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_Export_Import()
        {
            Random rnd = new Random();
            //// Prepare
            string lpCartName = "lpCartTest" + rnd.Next().ToString();
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            bool newVersion = true;

            //Data to create LPcart yacamil

            string code3 = rnd.Next(0, 500000).ToString();
            string site3 = TestContext.Properties["SiteLpCart2"].ToString();
            string customer3 = TestContext.Properties["CustomerLpCart2"].ToString();
            string aircraft3 = TestContext.Properties["AircraftLpCart"].ToString();
            DateTime from3 = DateUtils.Now.AddDays(1);
            DateTime to3 = DateUtils.Now.AddDays(10);
            string comment3 = TestContext.Properties["CommentLpCart"].ToString();
            string trolleyName = TestContext.Properties["TrolleyName"].ToString();
            string trolleyName1 = TestContext.Properties["TrolleyName1"].ToString();
            string trolleyScheme = TestContext.Properties["TrolleySchemeName"].ToString();
            string trolleyScheme1 = TestContext.Properties["TrolleySchemeName1"].ToString();
            //Excel 
            string columnName = "Trolley name";
            string sheetName = "LPCarts";
            string value = "trolley123";
            // Arrange
            HomePage homePage = LogInAsAdmin();
            // Act
            var lpCartPage = homePage.GoToFlights_LpCartPage();
            lpCartPage.ResetFilter();

            DeleteAllFileDownload();

            lpCartPage.ClearDownloads();

            lpCartPage.Filter(LpCartPage.FilterType.Search, lpCartName);
            if (lpCartPage.CheckTotalNumber() == 0)
            {
                // Create
                var lpCartModalCreate = lpCartPage.LpCartCreatePage();
                var lpCartCartDetailPage = lpCartModalCreate.FillField_CreateNewLpCart(code3, lpCartName, site3, customer3, aircraft3, from3, to3, comment3);

                //Acces Onglet Cart                
                lpCartCartDetailPage.ClickAddtrolley();
                lpCartCartDetailPage.AddTrolley(trolleyName);

                //Create LpCart Scheme
                var lpCartSchemeModal = lpCartPage.LpCartCreateLpCartSchemeModal();
                lpCartSchemeModal.CreateLpCartscheme(trolleyScheme, "2", "2");

                lpCartCartDetailPage.ClickAddtrolley();
                lpCartCartDetailPage.AddTrolley(trolleyName1);

                //Create LpCart Scheme
                lpCartPage.LpCartCreateLpCartSchemeModal();
                lpCartSchemeModal.CreateLpCartscheme(trolleyScheme1, "2", "2");

                lpCartCartDetailPage.BackToList();
            }

            lpCartPage.ExportLpcart(ExportType.Export, newVersion);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = lpCartPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            int resultNumber = OpenXmlExcel.GetExportResultNumber("LPCarts", filePath);
            var listResult = OpenXmlExcel.GetValuesInList("Name", "LPCarts", filePath);

            //Assert
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.IsFalse(!listResult.Contains(lpCartName), MessageErreur.EXCEL_DONNEES_KO);

            OpenXmlExcel.WriteDataInColumn(columnName, sheetName, filePath, value, CellValues.String);


            lpCartPage.Import(filePath);

            lpCartPage.Filter(LpCartPage.FilterType.Search, lpCartName);

            //Assert
            Assert.AreEqual(1, lpCartPage.CheckTotalNumber(), "L'import du fichier ne s'est pas executé");
            Assert.AreEqual(lpCartPage.GetFirstLpCartName(), lpCartName, "L'import du fichier ne s'est pas éxecuté");

            var lpCartDetailPage = lpCartPage.ClickFirstLpCart();
            var verifImport = lpCartPage.verifierImport(value);
            Assert.IsTrue(verifImport, "mise a jour des Lpcarts erroné");

            lpCartDetailPage.BackToList();

            lpCartPage.DeleteLpCart();
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_Index_Filter_ByLPCartValidityDate()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            // Act
            var lpCartPage = homePage.GoToFlights_LpCartPage();
            lpCartPage.ResetFilter();
            lpCartPage.Filter(FilterType.ValidityDate, true);
            var filterWorks = lpCartPage.VerifyFilterByValidityDate();
            Assert.IsTrue(filterWorks, "erreur de filtrage by valdity date");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_DuplicateLpCartSeveralAircrafts()
        {
            Random rnd = new Random();
            //// Prepare
            string name = TestContext.Properties["LpCartName1"].ToString();
            string nameDuplicate = TestContext.Properties["LpCartName1"].ToString() + " - Duplicate";
            string code = rnd.Next(0, 500000).ToString();
            string site = TestContext.Properties["SiteLpCart2"].ToString();
            string customer = TestContext.Properties["CustomerLpCart2"].ToString();
            string aircraft = TestContext.Properties["AircraftLpCart"].ToString();

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var lpCartPage = homePage.GoToFlights_LpCartPage();
            lpCartPage.ResetFilter();

            var lpCartDuplicateLpCartModalPage = lpCartPage.DuplicateLpCart();
            lpCartDuplicateLpCartModalPage.SetValuesForDuplication(name, code, nameDuplicate, site, customer, aircraft);
            var aircraftsByLpCart = lpCartPage.GetAircraftsByLpCartName(name);
            lpCartPage = homePage.GoToFlights_LpCartPage();
            lpCartPage.Filter(FilterType.Search, nameDuplicate);
            var lpCartDetailPage = lpCartPage.ClickFirstLpCart();
            var lpCartGeneralInformationsPage = lpCartDetailPage.LpCartGeneralInformationPage();
            var aircraftsByLpCartDuplicate = lpCartPage.GetAircraftsByLpCartName(nameDuplicate);
            Assert.IsTrue(aircraftsByLpCart.ToString().Equals(aircraftsByLpCartDuplicate.ToString()), "duplication avec des aircrafts differents");
            lpCartPage.RemoveDuplicate(nameDuplicate);

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_CreateLPCart_WithRoutes()
        {
            var site = "ACE";
            var customer = "TVS - SMARTWINGS, A.S. (TVS)";
            var rnd = new Random();
            var code = "Test " + DateTime.Now.ToString("yyyy-MM-dd") + " " + rnd.Next();
            var name = code;
            var route = "ACE-AGP";
            var newRouteInput = "ACE-BCN";
            var aircraft = "AB310";
            DateTime from = DateUtils.Now.AddDays(5);
            DateTime to = DateUtils.Now.AddDays(10);
            string comment = "Label comment";
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            LpCartPage lpCartPage = homePage.GoToFlights_LpCartPage();
            lpCartPage.DeleteCodeLpCartTestIfExist(code, lpCartPage);


            var lpCartCreateModalPage = lpCartPage.LpCartCreatePage();
            lpCartCreateModalPage.FillField_CreateNewLpCartWithRoutes(code, name, site, customer, aircraft, from, to, comment, route);



            Assert.IsTrue(lpCartPage.VerifRouteSelect(route), "La route ajoutée n'est pas vérifiée.");
            lpCartPage.AddNewRouteLpCartDetail(newRouteInput);

            Assert.IsTrue(lpCartPage.VerifyAfterChangedTab(newRouteInput), "failed");

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_EditLPCart_Name()
        {
            Random rnd = new Random();
            string code = rnd.Next(0, 500).ToString(); ;
            var name = $"LPC {DateTime.Now}";
            var newname = $"{name} - Test";
            string site = TestContext.Properties["SiteLpCart"].ToString();
            string customer = TestContext.Properties["CustomerLpCart"].ToString();
            string aircraft = TestContext.Properties["AircraftLpCart"].ToString();
            DateTime from = DateUtils.Now.AddDays(5);
            DateTime to = DateUtils.Now.AddDays(10);
            string comment = "Label comment";

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            LpCartPage lpCartPage = homePage.GoToFlights_LpCartPage();
            var lpCartModalCreate = lpCartPage.LpCartCreatePage();
            var LpCartDetailPage = lpCartModalCreate.FillField_CreateNewLpCart(code, name, site, customer, aircraft, from, to, comment);
            var LpCartGeneralePage = LpCartDetailPage.LpCartGeneralInformationPage();
            LpCartGeneralePage.UpdateLpCartName(newname);
            lpCartPage = LpCartDetailPage.BackToList();
            lpCartPage.Filter(FilterType.Search, code);

            Assert.IsTrue(lpCartPage.CheckTotalNumber() != 0, "La création de lp carte n'est pas réalisée.");

            var lpCartName = lpCartPage.GetFirstLpCartName();
            Assert.AreNotEqual(lpCartName, newname, $"La mise à jour du nom de lp carte dû {name} a {newname} n'est pas réalisée.");
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_Filter_SearchOnCart_byName()
        {
            // Prepare
            string name = TestContext.Properties["LpCartName4"].ToString();
            string trolleyName = TestContext.Properties["TrolleyName"].ToString();
            string trolleyName1 = TestContext.Properties["TrolleyName1"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var lpCartPage = homePage.GoToFlights_LpCartPage();

            //Filter by Search
            lpCartPage.ResetFilter();
            lpCartPage.Filter(LpCartPage.FilterType.Search, name);

            LpCartCartDetailPage lpCartCartDetailPage = lpCartPage.ClickFirstLpCart();
            var cartsNumber = lpCartCartDetailPage.CheckTotalNumber();
            lpCartCartDetailPage.Filter(LpCartCartDetailPage.FilterType.Search, trolleyName);
            Assert.AreNotEqual(cartsNumber, lpCartCartDetailPage.CheckTotalNumber(), MessageErreur.FILTRE_ERRONE);
            lpCartCartDetailPage.Filter(LpCartCartDetailPage.FilterType.Search, trolleyName1);
            Assert.AreNotEqual(cartsNumber, lpCartCartDetailPage.CheckTotalNumber(), MessageErreur.FILTRE_ERRONE);
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_Filter_SortByOnCartsTab_LegOrder()
        {
            // Prepare
            string name = TestContext.Properties["LpCartName4"].ToString();
            string trolleyName = TestContext.Properties["TrolleyName"].ToString();
            string trolleyName1 = TestContext.Properties["TrolleyName1"].ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var lpCartPage = homePage.GoToFlights_LpCartPage();

            //Filter by Search
            lpCartPage.ResetFilter();
            lpCartPage.Filter(LpCartPage.FilterType.Search, name);

            LpCartCartDetailPage lpCartCartDetailPage = lpCartPage.ClickFirstLpCart();
            lpCartCartDetailPage.Filter(LpCartCartDetailPage.FilterType.SortBy, "LegAndOrder");
            Assert.IsTrue(lpCartCartDetailPage.VerifyFilterByLegOrder(), MessageErreur.FILTRE_ERRONE);
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_EditLPCart_Comment()
        {
            //// Prepare
            var customerFilter = TestContext.Properties["CustomerLpFilter"].ToString();
            var lpCartName = string.Empty;
            var site = TestContext.Properties["SiteACE"].ToString();
            var siteFilter = TestContext.Properties["SiteACEFilter"].ToString();
            var aircraft = TestContext.Properties["AircraftLpCart"].ToString();
            var comment = "This_is_a_comment_for_lp_cart";
            var rnd = new Random();
            var code = rnd.Next(0, 5000).ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            string DocFileNamePdfBegin = "Flights Print Label";
            string DocFileNameZipBegin = "All_files_";

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            homePage.ClearDownloads();
            homePage.PurgeDownloads();

            //Create lp cart if not existe and check if comment is updated after creation or update
            var lpCartPage = homePage.GoToFlights_LpCartPage();
            FilterLpCart();
            if (lpCartPage.CheckTotalNumber() == 0)
            {
                lpCartName = $"LPC {DateTime.Now}";
                DateTime from = DateUtils.Now;
                DateTime to = DateUtils.Now.AddDays(10);

                var lpCartModalCreate = lpCartPage.LpCartCreatePage();
                lpCartModalCreate.FillField_CreateNewLpCart(code, lpCartName, site, customerFilter, aircraft, from, to, comment)
                    .BackToList();
            }

            var lpCartCartDetailPage = lpCartPage.ClickFirstLpCart();
            var lpCartGeneralInformation = lpCartCartDetailPage.LpCartGeneralInformationPage();
            lpCartGeneralInformation.UpdateToDateTo();
            lpCartName = lpCartGeneralInformation.GetLpCartName();
            lpCartGeneralInformation.PageUp();
            lpCartGeneralInformation.AddAircraf(aircraft);
            lpCartGeneralInformation.UpdateLpCartComment(comment);
            lpCartGeneralInformation.ClickOnCarts().LpCartGeneralInformationPage();
            var updatedComment = lpCartGeneralInformation.GetLpCartComment();
            Assert.AreEqual(comment, updatedComment, $"La mise à jour du comentaire de lp carte n'est pas réalisée.");

            //update flight with created aircraft and print flight labels
            homePage.Navigate();
            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.ClearDownloads();
            flightPage.ShowCustomersMenu();
            flightPage.Filter(FlightPage.FilterType.Customer, customerFilter);
            flightPage.Filter(FlightPage.FilterType.Sites, site);
            flightPage.Filter(FlightPage.FilterType.Status, "Preval");

            //If flight not existe 
            Assert.IsTrue(flightPage.CheckTotalNumber() != 0, $"Impossible de trouver un vol pour le client {customerFilter} et le site {site}");
            var flightEditPage = flightPage.EditFirstFlight(string.Empty);
            flightEditPage.ChangeAircaftForLoadingPlan(aircraft);
            flightEditPage.SelectLpCart(lpCartName);

            //Execute print labes and check if comment is updated
            var printReportPage = flightEditPage.PrintFlightLabels();
            printReportPage.Close();
            Thread.Sleep(2000);
            printReportPage.Purge(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin);
            string filePath = printReportPage.PrintAllZip(downloadsPath, DocFileNamePdfBegin, DocFileNameZipBegin, false);
            FileInfo fi = new FileInfo(filePath);
            fi.Refresh();
            Assert.IsTrue(fi.Exists, filePath + " non généré");
            PdfDocument document = PdfDocument.Open(fi.FullName);

            var pages = document.GetPages();
            var hasCommentUpdated = pages
                .SelectMany(p => p.GetWords())
                .Any(w => w.Text == comment);

            Assert.IsTrue(hasCommentUpdated, $"{comment} : non présent dans le Pdf");

            void FilterLpCart()
            {
                lpCartPage.ResetFilter();
                lpCartPage.Filter(FilterType.Customers, customerFilter);
                lpCartPage.Filter(FilterType.Site, siteFilter);
            }
        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_EditLPCart_From()
        {
            // Prepare
            //lpCart & cart
            var code = new Random()
                .Next(0, 5000)
                .ToString();
            var from = DateTime.Today.AddDays(-5);
            var customerFilter = TestContext.Properties["CustomerLpFilter"].ToString();
            var site = TestContext.Properties["SiteACE"].ToString();
            var siteFilter = TestContext.Properties["SiteACEFilter"].ToString();
            var aircraft = TestContext.Properties["AircraftLpCart"].ToString();
            var lpCartName = $"LP Cart {DateTime.Now.ToString("dd/MM/yyyy")}";
            var fromLpCart = DateTime.Today.AddMonths(-3);
            var toLpCart = DateTime.Today.AddMonths(2);

            //Flight
            string flightNumber = new Random().Next().ToString();
            string siteTo = TestContext.Properties["SiteToFlightBob"].ToString();
            string etaHours = "00";
            string etdHours = "23";
            string loadingPlanName = null;
            DateTime date = DateTime.Today.AddDays(+1);
            string lpCart = null;

            //Arrange
            var homePage = LogInAsAdmin();

            bool isLpCartCreated = false;
            bool isFlightCreated = false;
            //Act
            //Create lp cart 
            #region LP cart - Update the existing LP cart with the dates (from - to ) with less than today.
            var lpCartPage = homePage.GoToFlights_LpCartPage();
            FilterLpCart();

            try
            {
                var lpCartModalCreate = lpCartPage.LpCartCreatePage();
                //Create lp cart and back to list
                lpCartModalCreate
                    .FillField_CreateNewLpCart(code, lpCartName, site, customerFilter, aircraft, fromLpCart, toLpCart, string.Empty)
                    .BackToList();
                isLpCartCreated = true;

                lpCartPage.Filter(FilterType.Search, lpCartName);
                var lpCartCartDetailPage = lpCartPage.ClickFirstLpCart();
                var lpCartGeneralInformation = lpCartCartDetailPage.LpCartGeneralInformationPage();

                //Update date from to befor today 
                lpCartGeneralInformation
                    .UpdateFromDate(from);
                lpCartCartDetailPage.BackToList();
                lpCartPage.Filter(FilterType.Search, lpCartName);
                lpCartCartDetailPage = lpCartPage.ClickFirstLpCart();
                lpCartGeneralInformation = lpCartCartDetailPage.LpCartGeneralInformationPage();

                var dateFromUpdated =DateTime.ParseExact(lpCartGeneralInformation.GetLpCartDateFrom(), "dd/MM/yyyy", CultureInfo.InvariantCulture);

                //Assert
                Assert.AreEqual(dateFromUpdated, from, " la dateFrom modifiée n'est pas bien enregistrée dans la général information. ");
                #endregion

                #region Flight

                //update flight with aircraft
                homePage.Navigate();
                var flightPage = homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.ShowCustomersMenu();
                flightPage.Filter(FlightPage.FilterType.Customer, customerFilter);
                flightPage.Filter(FlightPage.FilterType.Sites, site);

                var flightCreatePageModal = flightPage.FlightCreatePage();

                flightCreatePageModal.FillField_CreatNewFlight(flightNumber, customerFilter, aircraft, site, siteTo, loadingPlanName, etaHours, etdHours, lpCart, date);
                isFlightCreated = true;

                var flightEditPage = flightPage.EditFirstFlight(flightNumber);

                flightEditPage.ChangeAircaftForLoadingPlan(aircraft);

                var lPCartIsExist = flightEditPage.LPCartIsExist(lpCartName);

                Assert.IsTrue(lPCartIsExist, $"LP cart existe dans le flight, car le date from inferieur a la date du flight.");

                #endregion
            }

            //Suppression de LPCart et Flight crées
            finally
            {
                if (isLpCartCreated)
                {
                    lpCartPage = homePage.GoToFlights_LpCartPage();
                    lpCartPage.ResetFilter();
                    lpCartPage.Filter(FilterType.Search, lpCartName);
                    lpCartPage.DeleteLpCart();
                    lpCartPage.ResetFilter();
                }

                if (isFlightCreated)
                {
                    var flightPage = homePage.GoToFlights_FlightPage();
                    flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                    flightPage.MassiveDeleteMenus(flightNumber, site, null, false);
                    flightPage.ResetFilter();
                }
            }

            void FilterLpCart()
            {
                lpCartPage.Filter(FilterType.Customers, customerFilter);
                lpCartPage.Filter(FilterType.Site, siteFilter);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_Create_Lp_Cart()
        {

            var lpCartName = $"LPC {DateTime.Now}";
            var code = new Random()
                .Next(0, 500)
                .ToString() + "_" + DateTime.Now;
            var from = DateTime.Today.AddDays(-5);
            var to = DateTime.Today.AddDays(-1);

            var customerFilter = "366 - A.F.A. LANZAROTE";
            var site = TestContext.Properties["SiteACE"].ToString();
            var aircraft = TestContext.Properties["AircraftLpCart"].ToString();
            string comment = TestContext.Properties["CommentLpCart"].ToString();

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var lpCartPage = homePage.GoToFlights_LpCartPage();
            var lpCartModalCreate = lpCartPage
                .LpCartCreatePage();

            lpCartModalCreate.FillField_CreateNewLpCart(code, lpCartName, site, customerFilter, aircraft, from, to, comment);

            var result = lpCartPage.VerifCreateLPCartInfo(code, lpCartName, site, customerFilter, aircraft, from, to, comment);

            Assert.IsTrue(result, "Les informations du lp cart ne correspondent pas");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_Filter_SortByOnCartsTab_Name()
        {
            string lpCartName = "LPCART FOR FLIGHT";

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            LpCartPage lpCartPage = homePage.GoToFlights_LpCartPage();
            lpCartPage.Filter(FilterType.Search, lpCartName);
            LpCartCartDetailPage lpCartPageDetail = lpCartPage.LpCartCartDetailPage();

            Assert.IsTrue(lpCartPageDetail.VerifySortByName(), "Le sort by name ne fonctionne pas");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_CartsTab_SortByName()
        {
            // Prepare
            //lpCart & cart
            var code = new Random()
                .Next(0, 5000)
                .ToString();
            var customerFilter = TestContext.Properties["CustomerLpFilter"].ToString();
            var site = TestContext.Properties["SiteACE"].ToString();
            var aircraft = TestContext.Properties["AircraftLpCart"].ToString();
            var lpCartName = $"LP Cart {DateTime.Now}";
            var fromLpCart = DateTime.Today.AddMonths(-3);
            var toLpCart = DateTime.Today.AddMonths(+2);
            var comment = "comment";

            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            LpCartPage lpCartPage = homePage.GoToFlights_LpCartPage();
            try
            {
                var lpCartModalCreate = lpCartPage.LpCartCreatePage();
                //Create lp cart 
                LpCartCartDetailPage lpCartCartDetailPage = lpCartModalCreate
                    .FillField_CreateNewLpCart(code, lpCartName, site, customerFilter, aircraft, fromLpCart, toLpCart, comment);

                lpCartCartDetailPage.ClickAddtrolley();
                lpCartCartDetailPage.AddTrolleyCustom("1","code1", "name 1", 0);
                lpCartCartDetailPage.ClickAddtrolley();
                lpCartCartDetailPage.AddTrolleyCustom("2","code2", "name 2", 1);
                lpCartCartDetailPage.ClickAddtrolley();
                lpCartCartDetailPage.AddTrolleyCustom("3","code3", "name 3", 2);

                //Assert
                var checkSortByLegAndOrder = lpCartCartDetailPage.VerifySortByName();
                Assert.IsTrue(checkSortByLegAndOrder, "Le SortByName pour les Carts ne fonctionne pas correctement.");
            }
            finally
            {
                lpCartPage = homePage.GoToFlights_LpCartPage();
                lpCartPage.ResetFilter();
                lpCartPage.Filter(FilterType.Search, lpCartName);
                LpCartCartDetailPage lpCartCartDetailPage = lpCartPage.LpCartCartDetailPage();

                //Delete Carts
                lpCartCartDetailPage.DeleteAllLpCartScheme();
                lpCartCartDetailPage.BackToList();

                //Delete LpCart
                lpCartPage.DeleteLpCart();
                lpCartPage.ResetFilter();
            }

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_Filter_SortByOnCartsTab_Code()
        {
            string lpCartName = "LPCART FOR FLIGHT";

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            LpCartPage lpCartPage = homePage.GoToFlights_LpCartPage();
            lpCartPage.Filter(FilterType.Search, lpCartName);
            LpCartCartDetailPage lpCartPageDetail = lpCartPage.LpCartCartDetailPage();

            Assert.IsTrue(lpCartPageDetail.VerifySortByCode(), "Le sort by code ne fonctionne pas");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_Create_Cart_Plus()
        {
            // Prepare premier LpCart
            Random rnd = new Random();
            string code = rnd.Next(0, 500).ToString();
            string name = "LP cart delete test";
            string site = TestContext.Properties["SiteLpCart"].ToString();
            string customer = TestContext.Properties["CustomerLpCart"].ToString();
            string aircraft = TestContext.Properties["AircraftLpCart"].ToString();
            DateTime from = DateUtils.Now.AddDays(5);
            DateTime to = DateUtils.Now.AddDays(10);
            string comment = TestContext.Properties["CommentLpCart"].ToString();
            HomePage homePage = LogInAsAdmin();
            LpCartPage lpCartPage = homePage.GoToFlights_LpCartPage();
            LpCartCreateModalPage lpCartCreateModalPage = lpCartPage.LpCartCreatePage();
            LpCartCartDetailPage lpCartCartDetailPage = lpCartCreateModalPage.FillField_CreateNewLpCart(code, name, site, customer, aircraft, from, to, comment);
            lpCartCartDetailPage.Click_PlusBTN();
            bool isDefaultLpCartDetailsDisplayed = lpCartCartDetailPage.VerifyDefaultFields();
            Assert.IsTrue(isDefaultLpCartDetailsDisplayed, "Les champs ne sont pas vides et le leg par défaut n'est pas fixé à 1.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_Create_Cart_Info_Cart_Edit_Confirm()
        {
            // Login 
            HomePage homePage = LogInAsAdmin();

            // Prepare premier LpCart
            Random rnd = new Random();
            string code = rnd.Next(0, 500).ToString();
            string name = "LP cart - Cart edit test";
            string site = TestContext.Properties["SiteLpCart"].ToString();
            string customer = TestContext.Properties["CustomerLpCart"].ToString();
            string aircraft = TestContext.Properties["AircraftLpCart"].ToString();
            DateTime from = DateUtils.Now.AddDays(5);
            DateTime to = DateUtils.Now.AddDays(10);
            string comment = TestContext.Properties["CommentLpCart"].ToString();

            LpCartPage lpCartPage = homePage.GoToFlights_LpCartPage();
            LpCartCreateModalPage lpCartCreateModalPage = lpCartPage.LpCartCreatePage();
            LpCartCartDetailPage lpCartCartDetailPage = lpCartCreateModalPage.FillField_CreateNewLpCart(code, name, site, customer, aircraft, from, to, comment);
            Assert.IsTrue(lpCartCreateModalPage.IsFormPopupClosed(), "La popup n'est pas fermée correctement.");
            lpCartCartDetailPage.Click_PlusBTN();
            string trolleyName = TestContext.Properties["TrolleyName"].ToString();
            string trolleyScheme = TestContext.Properties["TrolleySchemeName"].ToString();
            lpCartCartDetailPage.AddTrolley(trolleyName);
            //Create LpCart Scheme
            var lpCartSchemeModal = lpCartPage.LpCartCreateLpCartSchemeModal();
            lpCartSchemeModal.CreateLpCartscheme(trolleyScheme, "2", "4", "Prepzone Text", "Short", "Comment Text", "5", "2", "EqpNameLabel Text", "cartColor Text", "mainPos Text");
            // Check values of Schema line and columns
            var schemaLineVal = lpCartPage.GetSchemaLineCount();
            var schemaColumnLine = lpCartPage.GetSchemeColumnCount();
            Assert.AreEqual(schemaLineVal, "2");
            Assert.AreEqual(schemaColumnLine, "4");
            var lpCartEditLpCartSchemeModal = lpCartPage.LpCartEditLpCartSchemeModal();
            var title = lpCartPage.GetTitleEditLpCart();
            Assert.AreEqual(title, trolleyName + "Scheme");
            var prepzpne = lpCartPage.GetPrepzoneEditLpCart();
            Assert.AreEqual(prepzpne, "Prepzone Text");
            var editComment = lpCartPage.GetCommentEditLpCart();
            Assert.AreEqual(editComment, "Comment Text");
            var epsNameLabel = lpCartPage.GetEqpNameLabelEditLpCart();
            Assert.AreEqual(epsNameLabel, "EqpNameLabel Text");
            var salsNumber = lpCartPage.GetSalsNumberEditLpCart();
            Assert.AreEqual(salsNumber, "5");
            var labelPageNumber = lpCartPage.GetLabelPageNbrEditLpCart();
            Assert.AreEqual(labelPageNumber, "2");
            var cartColor = lpCartPage.GetCartColorEditLpCart();
            Assert.AreEqual(cartColor, "cartColor Text");
            var mainPos = lpCartPage.GetMainPositionEditLpCart();
            Assert.AreEqual(mainPos, "mainPos Text");
            var rowsNumber = lpCartPage.GetRowsNumberEditLpCart();
            Assert.AreEqual(rowsNumber, "2");
            var colsCumber = lpCartPage.GetColumnsNumberEditLpCart();
            Assert.AreEqual(colsCumber, "4");



            //lpCartCartDetailPage.BackToList();
            //Assert.IsTrue(lpCartCartDetailPage.VerifyDefaultFields(), "Les Champs sont vide et Leg par defaut set à 1.");

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_Create_Cart_Info_Cart_Scheme_PostionPP()
        {
            Random rnd = new Random();
            string code = rnd.Next(0, 1000).ToString();
            string name = "LP cart" + code;
            string site = TestContext.Properties["SiteLpCart"].ToString();
            string customer = TestContext.Properties["CustomerLpCart"].ToString();
            string aircraft = TestContext.Properties["AircraftLpCart"].ToString();
            DateTime from = DateUtils.Now.AddDays(5);
            DateTime to = DateUtils.Now.AddDays(10);
            string comment = TestContext.Properties["CommentLpCart"].ToString();
            string trolleyName = TestContext.Properties["TrolleyName1"].ToString();

            HomePage homePage = LogInAsAdmin();

            LpCartPage lpCartPage = homePage.GoToFlights_LpCartPage();
            LpCartCreateModalPage lpCartCreateModalPage = lpCartPage.LpCartCreatePage();
            LpCartCartDetailPage lpCartCartDetailPage = lpCartCreateModalPage.FillField_CreateNewLpCart(code, name, site, customer, aircraft, from, to, comment);
            lpCartCartDetailPage.ClickAddtrolley();
            lpCartCartDetailPage.AddTrolley(trolleyName);
            //Create LpCart Scheme
            var lpCartSchemeModal = lpCartPage.LpCartCreateLpCartSchemeModal();
            lpCartSchemeModal.CreateLpCartscheme(string.Empty, "4", "2", "Prepzone Text", "Short", "Comment Text", "5", "2", "EqpNameLabel Text", "cartColor Text", "mainPos Text");
            lpCartPage = lpCartCartDetailPage.BackToList();
            lpCartPage.Filter(LpCartPage.FilterType.Search, name);
            var lpCartDetail = lpCartPage.LpCartCartDetailPage();
            var editlpCart = lpCartDetail.EditCartScheme();
            var liste_position_pp_line = editlpCart.Insert_Postion_PP_Line("4", "2");
            var editlpCartAfterInsert = lpCartDetail.EditCartScheme();
            var liste_position_pp_lineAfterEdit = editlpCartAfterInsert.List_Of_Postion_PP_Line_After_Insert();
            bool areEqual = liste_position_pp_line.SequenceEqual(liste_position_pp_lineAfterEdit);
            Assert.IsTrue(areEqual, "Le Cart ne contient pas encore les informations sur la Position et le PP Line");

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_Create_Cart_Info_Cart_Generate()
        {
            string trolleyName1 = TestContext.Properties["TrolleyName1"].ToString();
            var lpCartName = $"LP Cart {DateTime.Now}";
            var code = new Random()
                .Next(0, 500)
                .ToString() + "-" + DateTime.Now.ToString("yyyy-MM-dd");
            var from = DateTime.Today.AddDays(-5);
            var to = DateTime.Today.AddDays(-1);

            var customerFilter = "$$ - CAT Genérico";
            var site = TestContext.Properties["SiteACE"].ToString();
            var aircraft = TestContext.Properties["AircraftLpCart"].ToString();
            string comment = TestContext.Properties["CommentLpCart"].ToString();
            string shortComment = "comm";
            string titleTrolley = TestContext.Properties["TrolleySchemeName"].ToString();
            HomePage homePage = LogInAsAdmin();

            LpCartPage lpCartPage = homePage.GoToFlights_LpCartPage();
            var lpCartModalCreate = lpCartPage
               .LpCartCreatePage();
            LpCartCartDetailPage lpCartCartDetailPage = lpCartModalCreate.FillField_CreateNewLpCart(code, lpCartName, site, customerFilter, aircraft, from, to, comment);
            lpCartCartDetailPage.ClickAddtrolley();
            lpCartCartDetailPage.AddTrolley(trolleyName1, shortComment);
            lpCartCartDetailPage.BackToList();
            lpCartPage.ResetFilter();
            lpCartPage.Filter(LpCartPage.FilterType.Search, lpCartName);
            if (lpCartPage.CheckTotalNumber() != 0)
            {
                var lpCartDetail = lpCartPage.LpCartCartDetailPage();
                LpCartEditLpCartSchemeModal modalCartScheme = lpCartDetail.EditCartScheme();

                modalCartScheme.EditLpCartscheme(titleTrolley, "1", "2", "4", "2", "red");

                var rows = lpCartDetail.GetRowsValueFromEditModal();
                var cols = lpCartDetail.GetColumnsValueFromEditModal();
                Assert.IsTrue(modalCartScheme.VerifyCountLineColumn(rows, cols), "le Line count et le Column count ne sont pas enregistrés correctement");

            }

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_Create_Cart_Info_Cart_Scheme_EditText_Save()
        {
            //prepare
            Random rnd = new Random();
            string code = rnd.Next(0, 500).ToString();
            string name = "LP cart delete test";
            string site = TestContext.Properties["SiteLpCart"].ToString();
            string customer = TestContext.Properties["CustomerLpCart"].ToString();
            string aircraft = TestContext.Properties["AircraftLpCart"].ToString();
            DateTime from = DateUtils.Now.AddDays(5);
            DateTime to = DateUtils.Now.AddDays(10);
            string comment = TestContext.Properties["CommentLpCart"].ToString();
            string trolleyName = TestContext.Properties["TrolleyName1"].ToString();
            string position = "NEWPOSITION";
            string pandp = "NEWP&P";
            string positionDetail = "NEWPOSITIONDETAIL";
            HomePage homePage = LogInAsAdmin();
            LpCartPage lpCartPage = homePage.GoToFlights_LpCartPage();
            LpCartCreateModalPage lpCartCreateModalPage = lpCartPage.LpCartCreatePage();
            LpCartCartDetailPage lpCartCartDetailPage = lpCartCreateModalPage.FillField_CreateNewLpCart(code, name, site, customer, aircraft, from, to, comment);
            lpCartCartDetailPage.ClickAddtrolley();
            lpCartCartDetailPage.AddTrolley(trolleyName);
            lpCartPage = lpCartCartDetailPage.BackToList();
            lpCartPage.Filter(LpCartPage.FilterType.Search, name);
            var lpCartDetail = lpCartPage.LpCartCartDetailPage();
            var editlpCart = lpCartDetail.EditCartScheme();
            editlpCart.EditLpCartscheme("trolley", "1", "2", "2", "2", "red");
            editlpCart = lpCartDetail.EditCartScheme();
            var oldPositionValue = "50";
            var weightValue = "10";
            editlpCart.EditWeightPosition(position, pandp, positionDetail, weightValue, oldPositionValue);
            editlpCart.ClickOnSavePositionDetails();
            Assert.IsTrue(editlpCart.IsEditPositionPopupClosed(), "La popup ne se ferme pas correctement.");
            editlpCart.ClickOnAddLpCartSheme();
            editlpCart = lpCartDetail.EditCartScheme();
            bool isAddAll = editlpCart.VerifyWeightPosition(position, pandp, positionDetail, weightValue, oldPositionValue);
            Assert.IsTrue(isAddAll, "L'ajout du weight et position sur le cart ne fonctionne pas.");

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_Create_Cart_Info_Cart_Scheme_EditText_Popup()
        {
            string trolleyName1 = TestContext.Properties["TrolleyName1"].ToString();
            var lpCartName = $"LP Cart {DateTime.Now}";
            var code = new Random().Next(0, 500) .ToString();
            var from = DateTime.Today.AddDays(-5);
            var to = DateTime.Today.AddDays(-1);

            var customerFilter = "$$ - CAT Genérico";
            var site = TestContext.Properties["SiteACE"].ToString();
            var aircraft = TestContext.Properties["AircraftLpCart"].ToString();
            string comment = TestContext.Properties["CommentLpCart"].ToString();
            string shortComment = "comm";
            string titleTrolley = TestContext.Properties["TrolleySchemeName"].ToString();

            HomePage homePage = LogInAsAdmin();
            LpCartPage lpCartPage = homePage.GoToFlights_LpCartPage();
            var lpCartModalCreate = lpCartPage.LpCartCreatePage();
            LpCartCartDetailPage lpCartCartDetailPage = lpCartModalCreate.FillField_CreateNewLpCart(code, lpCartName, site, customerFilter, aircraft, from, to, comment);
            lpCartCartDetailPage.ClickAddtrolley();
            lpCartCartDetailPage.AddTrolley(trolleyName1, shortComment);
            lpCartCartDetailPage.BackToList();
            lpCartPage.ResetFilter();
            lpCartPage.Filter(LpCartPage.FilterType.Search, lpCartName);
            if (lpCartPage.CheckTotalNumber() != 0)
            {
                var lpCartDetail = lpCartPage.LpCartCartDetailPage();
                LpCartEditLpCartSchemeModal modalCartScheme = lpCartDetail.EditCartScheme();
                Assert.IsTrue(lpCartDetail.IsEditPopupDisplayed(), "Une nouvelle popup ne s'ouvre pas correctement");
                modalCartScheme.EditLpCartscheme(titleTrolley, "1", "2", "2", "2", "red");
                lpCartDetail.EditCartScheme();
                Assert.IsTrue(modalCartScheme.FieldsPSDetailsIsEmpty(), "Les Champs: Poids , Previous position et le text Doivent être vides");
                Assert.IsTrue(modalCartScheme.IsEditPositionDetailsPopupDisplayed(), "Une nouvelle popup ne s'ouvre pas correctement");

            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_Filter_SearchOnCart_byCode()
        {
            // Prepare
            string name = TestContext.Properties["LpCartName4"].ToString();
            string trolleyName = TestContext.Properties["TrolleyName"].ToString();
            string trolleyCode = TestContext.Properties["CodeTrolley"].ToString();
            string trolleyName1 = TestContext.Properties["TrolleyName1"].ToString();
            string trolleyCode1 = TestContext.Properties["CodeTrolley1"].ToString();
            var lpCartName = $"LP Cart {DateTime.Now}";
            var code = DateTime.Now.ToString("yyyy-MM-dd") + "-" + new Random().Next(0, 500).ToString();
            var from = DateTime.Today.AddDays(-5);
            var to = DateTime.Today.AddDays(-1);
            var customerFilter = "$$ - CAT Genérico";
            var site = TestContext.Properties["SiteACE"].ToString();
            var aircraft = "AB310";
            string comment = TestContext.Properties["CommentLpCart"].ToString();
            string titleTrolley = TestContext.Properties["TrolleySchemeName"].ToString();
            // Arrange
            HomePage homePage = LogInAsAdmin();
            LpCartPage lpCartPage = homePage.GoToFlights_LpCartPage();
            try
            {
                var lpCartModalCreate = lpCartPage.LpCartCreatePage();
                LpCartCartDetailPage lpCartCartDetailPage = lpCartModalCreate.FillField_CreateNewLpCart(code, lpCartName, site, customerFilter, aircraft, from, to, comment);
                lpCartCartDetailPage.ClickAddtrolley();
                lpCartCartDetailPage.FillNewTrolleyLine(trolleyName, trolleyCode, trolleyName, trolleyName, trolleyName, trolleyName);
                lpCartCartDetailPage.ClickAddtrolley();
                lpCartCartDetailPage.FillNewTrolleyLine(trolleyName1, trolleyCode1, trolleyName1, trolleyName1, trolleyName1, trolleyName1);
                lpCartCartDetailPage.Filter(LpCartCartDetailPage.FilterType.Search, trolleyCode);
                bool isSearchByCodeOk = lpCartCartDetailPage.CheckSearchByCode(trolleyCode);
                Assert.IsTrue(isSearchByCodeOk, "La recherche par code ne fonctionne pas");
                lpCartPage = lpCartCartDetailPage.BackToList();
            }
            finally
            {
                lpCartPage.Filter(LpCartPage.FilterType.Search, code);
                lpCartPage.DeleteLpCart();
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_Create_Cart_Info_Cart_Default()
        {
            // Login 
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            string trolleyName1 = TestContext.Properties["TrolleyName1"].ToString();
            var lpCartName = $"LP Cart {DateTime.Now}";
            var code = new Random().Next(0, 500).ToString();
            var from = DateTime.Today.AddDays(-5);
            var to = DateTime.Today.AddDays(-1);
            var customerFilter = "$$ - CAT Genérico";
            var site = TestContext.Properties["SiteACE"].ToString();
            var aircraft = TestContext.Properties["AircraftLpCart"].ToString();
            string comment = TestContext.Properties["CommentLpCart"].ToString();
            string shortComment = "comm";
            LpCartPage lpCartPage = homePage.GoToFlights_LpCartPage();
            var lpCartModalCreate = lpCartPage.LpCartCreatePage();
            LpCartCartDetailPage lpCartCartDetailPage = lpCartModalCreate.FillField_CreateNewLpCart(code, lpCartName, site, customerFilter, aircraft, from, to, comment);
            lpCartCartDetailPage.ClickAddtrolley();
            lpCartCartDetailPage.AddTrolley(trolleyName1, shortComment);
            lpCartCartDetailPage.BackToList();
            lpCartPage.ResetFilter();
            lpCartPage.Filter(LpCartPage.FilterType.Search, lpCartName);
            if (lpCartPage.CheckTotalNumber() != 0)
            {
                var lpCartDetail = lpCartPage.LpCartCartDetailPage();
                lpCartDetail.EditCartScheme();
                Assert.IsTrue(lpCartDetail.IsEditPopupDisplayed(), "Une nouvelle popup ne s'ouvre pas correctement");
                Assert.IsTrue(lpCartDetail.VerifyLpCartName(trolleyName1), " Le nom du LP Cart n s'affiche pas ou il néest pas grisé.");
                Assert.IsTrue(lpCartDetail.VerifyShortComment(shortComment), "Le Short Comment n'est pas identique");
                Assert.IsTrue(lpCartDetail.VerifyLabelPage(), "le Label Page Number n'est pas égale à 1.");
                Assert.IsTrue(lpCartDetail.VerifyPrintPositionsChecked(), "La mention Print Positions Doivent être Coché par défaut .");
                Assert.IsTrue(lpCartDetail.VerifyRowsAndCols(), "Les Rows et Columns doivent être automatiquement à 0");
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_Create_Cart_Info_Index()
        {
            // Prepare premier LpCart
            Random rnd = new Random();
            string code = "LP cart delete test" + rnd.Next(0, 500).ToString();
            string name = "LP cart delete test" + rnd.Next(0, 500).ToString(); ;
            string site = TestContext.Properties["SiteLpCart"].ToString();
            string customer = TestContext.Properties["CustomerLpCart"].ToString();
            string aircraft = TestContext.Properties["AircraftLpCart"].ToString();
            DateTime to = DateUtils.Now.AddDays(10);
            DateTime from = DateUtils.Now.AddDays(5);
            string comment = TestContext.Properties["CommentLpCart"].ToString();
            // ADD New Row Lpart Details
            string trolleyName = TestContext.Properties["TrolleyName"].ToString();
            string trolleycode = TestContext.Properties["CodeTrolley"].ToString(); ;
            string trolleyloc = "localisation";
            string galleycode = "code";
            string galleyloc = "localisation";
            string equipement = "equipement";
            string shortcomment = "short";

            HomePage homePage = LogInAsAdmin();

            LpCartPage lpCartPage = homePage.GoToFlights_LpCartPage();
            try
            {
                LpCartCreateModalPage lpCartCreateModalPage = lpCartPage.LpCartCreatePage();
                LpCartCartDetailPage lpCartCartDetailPage = lpCartCreateModalPage.FillField_CreateNewLpCart(code, name, site, customer, aircraft, from, to, comment);
                lpCartCartDetailPage.Click_PlusBTN();
                lpCartCartDetailPage.CreateNewRowLpCart(1, trolleyName, trolleycode, trolleyloc, galleycode, galleyloc, equipement, shortcomment);
                var elementplus = lpCartCartDetailPage.GetElementPlus();
                lpCartCartDetailPage.LpCartGeneralInformationPage();
                lpCartCartDetailPage.GoToCartsTab();
                bool cerifChangeBtnPlusToEditAndVerifIconDelete = lpCartCartDetailPage.VerifyChangePlusToEdit(elementplus);
                Assert.IsTrue(cerifChangeBtnPlusToEditAndVerifIconDelete, "Le picto (+) ne s'est pas transformé en picto stylo && Le picto delete n'est pas tout à gauche de la colonne.");
                lpCartCartDetailPage.BackToList();
            }
            finally
            {
                lpCartPage.ResetFilter();
                lpCartPage.DeleteCodeLpCartTestIfExist(name, lpCartPage);
            }
            

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_Create_Cart_Info_Cart_Scheme_EditText_Properties_Sq()
        {
            Random rnd = new Random();
            string customer = TestContext.Properties["CustomerLPFlight"].ToString();
            string code = rnd.Next(0, 9999).ToString();
            string name = "LPCART Testing SQ";
            string site = TestContext.Properties["SiteLpCart2"].ToString();
            string aircraft = TestContext.Properties["Aircraft"].ToString();
            DateTime from = DateUtils.Now.AddYears(-2);
            DateTime to = DateUtils.Now.AddYears(2);
            string comment = TestContext.Properties["CommentLpCart"].ToString();


            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            try
            {
                CustomerPage customerPage = homePage.GoToCustomers_CustomerPage();
                customerPage.Filter(CustomerPage.FilterType.Search, customer);
                CustomerGeneralInformationPage customerGeneralInfo = customerPage.SelectFirstCustomer();
                customerGeneralInfo.SetPrintLabelFormat("Square");

                var lpCartPage = homePage.GoToFlights_LpCartPage();
                var lpCartModalCreate = lpCartPage.LpCartCreatePage();
                var LpCartDetailPage = lpCartModalCreate.FillField_CreateNewLpCartWithRoutes(code, name, site, "$$ - " + customer, aircraft, from, to, comment, "MAD - BCN");

                LpCartDetailPage.Click_PlusBTN();

                string trolleyName = TestContext.Properties["TrolleyName"].ToString();
                string trolleyScheme = TestContext.Properties["TrolleySchemeName"].ToString();

                LpCartDetailPage.AddTrolley(trolleyName);

                //Create LpCart Scheme
                var lpCartSchemeModal = lpCartPage.LpCartCreateLpCartSchemeModal();
                lpCartSchemeModal.CreateLpCartscheme(trolleyScheme, "2", "4", "Prepzone Text", "Short", "Comment Text", "5", "2", "EqpNameLabel Text", "cartColor Text", "mainPos Text");

                var flightPage = homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.Sites, site);
                flightPage.Filter(FlightPage.FilterType.SearchFlight, aircraft);
                flightPage.ShowCustomersMenu();
                flightPage.Filter(FlightPage.FilterType.Customer, customer);
                string flightNumber = flightPage.GetFirstFlightNumber();
                var flightDetail = flightPage.EditFirstFlight(flightNumber);
                Assert.IsTrue(flightDetail.LPCartIsExist("LPCART Testing SQ"));
            }
            finally
            {
                CustomerPage customerPage = homePage.GoToCustomers_CustomerPage();
                customerPage.Filter(CustomerPage.FilterType.Search, customer);
                CustomerGeneralInformationPage customerGeneralInfo = customerPage.SelectFirstCustomer();
                customerGeneralInfo.SetPrintLabelFormat("Portrait");
            }

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_EditLPCart_Aircrafts()
        {
            //lpCart & cart
            string trolleyName1 = TestContext.Properties["TrolleyName1"].ToString();
            var lpCartName = $"LP Cart {DateTime.Now}";
            var code = new Random()
                .Next(0, 500)
                .ToString();
            var from = DateTime.Today.AddDays(-700);
            var to = DateTime.Today.AddDays(600);
            var customer1 = "$$ - CAT Genérico";
            var customer = TestContext.Properties["CustomerLPFlight"].ToString();
            var site = TestContext.Properties["SiteMAD"].ToString();
            var aircraft1 = TestContext.Properties["AircraftBis"].ToString();
            var route1 = TestContext.Properties["Route"].ToString();
            string comment = TestContext.Properties["CommentLpCart"].ToString();
            string shortComment = "comm";
            string titleTrolley = TestContext.Properties["TrolleySchemeName"].ToString();
            List<string> routes = new List<string>();
            List<string> aircrafts = new List<string>();
            //flight
            string flightNumber = new Random().Next().ToString();
            string siteFrom = TestContext.Properties["Site"].ToString();
            string siteTo = TestContext.Properties["SiteLP"].ToString();
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            LpCartPage lpCartPage = homePage.GoToFlights_LpCartPage();
            var lpCartModalCreate = lpCartPage
               .LpCartCreatePage();
            LpCartCartDetailPage lpCartCartDetailPage = lpCartModalCreate.FillField_CreateNewLpCartWithRoutes(code, lpCartName, site, customer1, aircraft1, from, to, comment, route1);
            aircrafts.Add(aircraft1);
            routes.Add(route1);
            lpCartCartDetailPage.ClickAddtrolley();
            lpCartCartDetailPage.AddTrolley(trolleyName1, shortComment);
            lpCartCartDetailPage.BackToList();
            lpCartPage.ResetFilter();
            lpCartPage.Filter(LpCartPage.FilterType.Search, lpCartName);
            if (lpCartPage.CheckTotalNumber() != 0)
            {
                var lpCartDetail = lpCartPage.LpCartCartDetailPage();
                LpCartEditLpCartSchemeModal modalCartScheme = lpCartDetail.EditCartScheme();
                modalCartScheme.EditLpCartscheme(titleTrolley, "1", "2", "1", "1", "red");
                LpCartGeneralInformationPage lpCartGeneralInformationTab = lpCartCartDetailPage.LpCartGeneralInformationPage();
                Assert.IsTrue(lpCartPage.VerifCreateLPCartInfoWithRoutes(code, lpCartName, site, customer1, aircrafts, from, to, comment, routes), "un problème est survenu lors de l'enregistrement dans la général information");
            }
            lpCartCartDetailPage.BackToList();
            string airCraftsToString = String.Join(" | ", aircrafts);
            var air = lpCartPage.GetAircraftTypeLpCart();
            Assert.AreEqual(airCraftsToString, air, "Le(s) aircraft(s) ne sont pas affiché(s) sur l'index des LP Carts colonne 'Aircraft Type'");
            FlightPage flightPage = homePage.GoToFlights_FlightPage();
            flightPage.Filter(FlightPage.FilterType.Sites, site);
            var flightCreatePageModal = flightPage.FlightCreatePage();
            flightCreatePageModal.FillField_CreatNewFlight(flightNumber, customer, aircraft1, siteFrom, siteTo);
            FlightDetailsPage flightDetailsPage = flightPage.EditFirstFlight(flightNumber);
            flightDetailsPage.ChangeAircaftForLoadingPlan(aircraft1);
            Assert.IsTrue(flightDetailsPage.LPCartIsExist(lpCartName), "le champ 'LP Cart' ne propose pas l'option de LP Cart éditée comme prévu");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_Create_Cart_Info_Cart_Scheme_EditText_Properties_Portrait()
        {
            Random rnd = new Random();
            var flightN = rnd.Next(250, 99999).ToString();
            string customer = TestContext.Properties["CustomerLpFilter"].ToString();
            string code = rnd.Next(0, 9999).ToString();
            string name = "LPCART Testing Portrait";
            string site = TestContext.Properties["SiteLpCart2"].ToString();
            string aircraft = TestContext.Properties["AircraftBis"].ToString();
            string siteFrom = TestContext.Properties["Site"].ToString();
            DateTime from = DateUtils.Now.AddYears(-2);
            DateTime to = DateUtils.Now.AddYears(2);
            string comment = TestContext.Properties["CommentLpCart"].ToString();


            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var lpCartPage = homePage.GoToFlights_LpCartPage();
            var lpCartModalCreate = lpCartPage.LpCartCreatePage();
            var LpCartDetailPage = lpCartModalCreate.FillField_CreateNewLpCartWithRoutes(code, name, site, customer, aircraft, from, to, comment, "MAD - BCN");

            LpCartDetailPage.Click_PlusBTN();

            string trolleyName = TestContext.Properties["TrolleyName"].ToString();
            string trolleyScheme = TestContext.Properties["TrolleySchemeName"].ToString();

            LpCartDetailPage.AddTrolley(trolleyName);

            //Create LpCart Scheme
            var lpCartSchemeModal = lpCartPage.LpCartCreateLpCartSchemeModal();
            lpCartSchemeModal.CreateLpCartscheme(trolleyScheme, "2", "4", "Prepzone Text", "Short Comment Text", "Comment Text", "5", "2", "EqpNameLabel Text", "cartColor Text", "mainPos Text");

            var flightPage = homePage.GoToFlights_FlightPage();
            flightPage.ResetFilter();
            flightPage.Filter(FlightPage.FilterType.Sites, site);
            FlightCreateModalPage flightCreateModalPage = flightPage.FlightCreatePage();
            flightCreateModalPage.FillField_CreatNewFlight(flightN, customer, aircraft, siteFrom, site);
            flightPage.Filter(FlightPage.FilterType.SearchFlight, aircraft);
            flightPage.ShowCustomersMenu();
            flightPage.Filter(FlightPage.FilterType.Customer, "TVS - SMARTWINGS, A.S.");
            string flightNumber = flightPage.GetFirstFlightNumber();
            var flightDetail = flightPage.EditFirstFlight(flightNumber);
            Assert.IsTrue(flightDetail.LPCartIsExist("LPCART Testing Portrait"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_Routes_Names_Change()
        {
            var site = "ACE";
            var customer = "TVS - SMARTWINGS, A.S. (TVS)";
            var codeRnd = new Random();
            var lpCartRnd = new Random();

            var code = $"TestTest-{codeRnd.Next()}";
            var lpCartName = $"LP Cart {DateTime.Now:yyyyMMdd}-{lpCartRnd.Next()}";
            // champs limité en nombre de 30 caractères
            var newLpCartName = $"NewLPCart {DateTime.Now:yyyyMMdd}-{lpCartRnd.Next()}";
            var route = "ACE-AGP";
            var aircraft = "AB310";
            DateTime from = DateUtils.Now.AddDays(5);
            DateTime to = DateUtils.Now.AddDays(10);
            string comment = "Comment Test";

            var homePage = LogInAsAdmin();
            var lpCartPage = homePage.GoToFlights_LpCartPage();
            lpCartPage.ResetFilter();

            // Create LP Cart with route
            var lpCartCreateModalPage = lpCartPage.LpCartCreatePage();
            var lpCartCartDetailPage = lpCartCreateModalPage.FillField_CreateNewLpCartWithRoutes(code, lpCartName, site, customer, aircraft, from, to, comment, route);
            lpCartCartDetailPage.LpCartGeneralInformationPage();

            bool IsRoutesAvailable = lpCartPage.VerifLPCartHasRoutes(new List<string> { route });
            var routeName = lpCartPage.GetRouteNames();

            Assert.IsTrue(IsRoutesAvailable, "LP Cart n'a pas des routes");

            lpCartCartDetailPage.BackToList();

            lpCartPage.ResetFilter();
            lpCartPage.Filter(LpCartPage.FilterType.Search, lpCartName);

            lpCartPage.ClickFirstLpCart();
            lpCartCartDetailPage.LpCartGeneralInformationPage();
            lpCartCartDetailPage.ChangeLpCartName(newLpCartName);

            lpCartCartDetailPage.BackToList();

            lpCartPage.ResetFilter();
            lpCartPage.Filter(LpCartPage.FilterType.Search, newLpCartName);
            var totalLpCart = lpCartCartDetailPage.CheckTotalNumber();

            if (totalLpCart == 0)
            {
                Assert.Fail("Aucun LP Cart trouvé avec le nom modifié.");
            }
            else if (totalLpCart == 1)
            {
                lpCartPage.ClickFirstLpCart();
                lpCartCartDetailPage.LpCartGeneralInformationPage();
            }
            Assert.AreNotEqual(newLpCartName, routeName, "Le nom de la route prend le nom du LP Cart.");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_ImportUpdateAircraft()
        {
            // Répertoire du fichier à importer
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            bool newVersion = true;

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            // Act
            var lpCartPage = homePage.GoToFlights_LpCartPage();
            lpCartPage.ResetFilter();

            var FirstNameLpCart = lpCartPage.GetFirstLpCartName();
            var FirstAircraftLpCart = lpCartPage.GetAircraftTypeLpCart();

            DeleteAllFileDownload();
            lpCartPage.ClearDownloads();

            lpCartPage.Filter(LpCartPage.FilterType.Search, FirstNameLpCart);
            lpCartPage.ExportLpcart(ExportType.Export, newVersion);

            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            var correctDownloadedFile = lpCartPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber("LPCarts", filePath);
            bool result7377 = OpenXmlExcel.ReadAllDataInColumn("Aircraft(s) type", "LPCarts", filePath, "7377");
            bool resultB777 = OpenXmlExcel.ReadAllDataInColumn("Aircraft(s) type", "LPCarts", filePath, "B777");
            long fileLengthBefore = new FileInfo(filePath).Length;
            if (result7377)
            {
                // nouvelle valeur
                OpenXmlExcel.WriteDataInColumn("Aircraft(s) type", "LPCarts", filePath, "B777", CellValues.SharedString);
            }
            else
            {
                // nouvelle valeur
                OpenXmlExcel.WriteDataInColumn("Aircraft(s) type", "LPCarts", filePath, "7377", CellValues.SharedString);
            }

            //Thread.Sleep(5000);
            Assert.IsTrue(new FileInfo(filePath).Exists);
            long fileLengthAfter = new FileInfo(filePath).Length;

            Assert.AreNotEqual(fileLengthBefore, fileLengthAfter);

            //Assert
            Assert.AreNotEqual(0, resultNumber, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.IsTrue(result7377 || resultB777, MessageErreur.EXCEL_DONNEES_KO);

            lpCartPage.Import(filePath);
            lpCartPage.ResetFilter();
            lpCartPage.Filter(LpCartPage.FilterType.Search, FirstNameLpCart);

            var updatedAircraftLpCart = lpCartPage.GetAircraftTypeLpCart();
            Assert.AreNotEqual(FirstAircraftLpCart, updatedAircraftLpCart, "Le type d'avion n'a pas été mis à jour.");

        }


        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_EditLPCart_AddAircraft()
        {
            //lpCart & cart
            var lpCartName1 = $"LP Cart1 {DateTime.Now.ToString("dd/MM/yyyy")}";
            var lpCartName2 = $"LP Cart2 {DateTime.Now.ToString("dd/MM/yyyy")}";
            var code = new Random()
                .Next(0, 5000)
                .ToString();
            var code1 = new Random()
                .Next(5001, 10000)
                .ToString();
            var from = DateTime.Today.AddDays(-700);
            var to = DateTime.Today.AddDays(600);
            var customer1 = "$$ - CAT Genérico";
            var customer = TestContext.Properties["CustomerLPFlight"].ToString();
            var site = TestContext.Properties["SiteMAD"].ToString();
            var site2 = TestContext.Properties["SiteACE"].ToString();
            var aircraft1 = TestContext.Properties["AircraftBis"].ToString();
            var aircraft2 = TestContext.Properties["Aircraft"].ToString();
            var route1 = TestContext.Properties["Route"].ToString();
            var route2 = TestContext.Properties["RouteLP"].ToString();
            string comment = TestContext.Properties["CommentLpCart"].ToString();

            //flight
            string flightNumberLpC1 = new Random().Next().ToString();
            string flightNumberLpC2 = new Random().Next().ToString();
            string siteFrom1 = TestContext.Properties["Site"].ToString(); //MAD
            string siteTo1 = TestContext.Properties["SiteLP"].ToString(); //ACE
            string siteFrom2 = TestContext.Properties["SiteLP"].ToString(); //ACE
            string siteTo2 = TestContext.Properties["SiteToFlightBob"].ToString(); //AGP

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            LpCartPage lpCartPage = homePage.GoToFlights_LpCartPage();
            try
            {
                // create LPCard 1
                var lpCartModalCreate1 = lpCartPage.LpCartCreatePage();
                var lpCartCartDetailPage1 = lpCartModalCreate1.FillField_CreateNewLpCartWithRoutes(
                    code, lpCartName1, site, customer1, aircraft1, from, to, comment, route1);
             
                //Add aircraft to LPCard 1
                lpCartCartDetailPage1.BackToList();
                lpCartPage.Filter(LpCartPage.FilterType.Search, lpCartName1);
                LpCartCartDetailPage lpCartCartDetailPage = lpCartPage.LpCartCartDetailPage();
                LpCartGeneralInformationPage lpCartGeneralInformationPage = lpCartCartDetailPage.LpCartGeneralInformationPage();
                lpCartGeneralInformationPage.AddAircraf(aircraft2);

                // Vérifier LPCard 1
                lpCartCartDetailPage1.BackToList();
                bool aircraftVerified = VerifyAircraft(lpCartPage, lpCartName1);
                Assert.IsTrue(aircraftVerified, "Le(s) aircraft(s) ne sont pas affiché(s) sur l'index des LP Carts colonne 'Aircraft Type'");
                lpCartCartDetailPage1.BackToList();

                // create LPCard 2
                var lpCartModalCreate2 = lpCartPage.LpCartCreatePage();
                var lpCartCartDetailPage2 = lpCartModalCreate2.FillField_CreateNewLpCartWithRoutes(
                    code1, lpCartName2, site2, customer1, aircraft2, from, to, comment, route2);

                //Add aircraft to LPCard 2
                lpCartCartDetailPage2.BackToList();
                lpCartPage.Filter(LpCartPage.FilterType.Search, lpCartName2);
                lpCartCartDetailPage2 = lpCartPage.LpCartCartDetailPage();
                lpCartGeneralInformationPage = lpCartCartDetailPage.LpCartGeneralInformationPage();
                lpCartGeneralInformationPage.AddAircraf(aircraft1);

                // Vérifier LPCard 2
                lpCartCartDetailPage2.BackToList();
                aircraftVerified = VerifyAircraft(lpCartPage, lpCartName2);
                Assert.IsTrue(aircraftVerified, "Le(s) aircraft(s) ne sont pas affiché(s) sur l'index des LP Carts colonne 'Aircraft Type'");
                lpCartCartDetailPage2.BackToList();

                //create flight 1
                var flightPage = homePage.GoToFlights_FlightPage();
                CreateFlightAircraft( flightPage, flightNumberLpC1, customer, aircraft2, siteFrom1, siteTo1);
                VerifyFlightAircraft( flightPage, flightNumberLpC1, aircraft2, lpCartName1);

                //create flight 2
                flightPage = homePage.GoToFlights_FlightPage();
                CreateFlightAircraft(flightPage, flightNumberLpC2,  customer,  aircraft1,  siteFrom2,  siteTo2);
                VerifyFlightAircraft(flightPage, flightNumberLpC2, aircraft1, lpCartName2);
            }
            finally 
            {
                var lpCartPage0 = homePage.GoToFlights_LpCartPage();
                lpCartPage.ResetFilter();
                lpCartPage.Filter(FilterType.Search, lpCartName1);
                lpCartPage.DeleteLpCart();

                lpCartPage.ResetFilter();
                lpCartPage.Filter(FilterType.Search, lpCartName2);
                lpCartPage.DeleteLpCart();

                var flightPage = homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumberLpC1);
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom1);
                flightPage.MassiveDeleteMenus(flightNumberLpC1, siteFrom1, customer, false);

                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumberLpC2);
                flightPage.Filter(FlightPage.FilterType.Sites, siteFrom2);
                flightPage.MassiveDeleteMenus(flightNumberLpC2, siteFrom2, customer, false);
            }
        }

        public bool VerifyAircraft(LpCartPage lpCartPage, string lpCartName)
        {   
            lpCartPage.ResetFilter();
            lpCartPage.Filter(LpCartPage.FilterType.Search, lpCartName);
            var displayedAircrafts = lpCartPage.GetAircraftTypeLpCart();
            LpCartCartDetailPage lpCartCartDetailPage = lpCartPage.LpCartCartDetailPage();
            LpCartGeneralInformationPage lpCartGeneralInformationPage = lpCartCartDetailPage.LpCartGeneralInformationPage();
            List<string> aircrafts = lpCartGeneralInformationPage.GetSelectedAircrafts();
            string airCraftsToString = string.Join(" | ", aircrafts);

            return airCraftsToString == displayedAircrafts;
        }

        private void CreateFlightAircraft(FlightPage flightPage, string flightNumber, string customer, string aircraft, string siteFrom, string siteTo)
        {
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
            var flightCreatePageModal = flightPage.FlightCreatePage();
            flightCreatePageModal.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);
        }

        private void VerifyFlightAircraft(FlightPage flightPage, string flightNumber, string aircraft, string lpCartName)
        {
            var flightDetailsPage = flightPage.EditFirstFlight(flightNumber);
            flightDetailsPage.ChangeAircaftForLoadingPlan(aircraft);

            //Assert
            bool lPCartIsExist = flightDetailsPage.LPCartIsExist(lpCartName);
            Assert.IsTrue(lPCartIsExist, $"le champ 'LP Cart' : {lpCartName} ne propose pas l'option de LP Cart éditée comme prévu");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_EditLPCart_AddRoute()
        {
            //lpCart & cart
            var lpCartName1 = $"LP Cart1 {DateTime.Now.ToString("dd/MM/yyyy")}";
            var lpCartName2 = $"LP Cart2 {DateTime.Now.ToString("dd/MM/yyyy")}";
            var code = new Random()
                .Next(0, 5000)
                .ToString();
            var code1 = new Random()
                .Next(5001, 10000)
                .ToString();
            var from = DateTime.Today.AddDays(-700);
            var to = DateTime.Today.AddDays(600);
            var customer1 = "$$ - CAT Genérico";
            var customer = TestContext.Properties["CustomerLPFlight"].ToString();
            var site = TestContext.Properties["SiteMAD"].ToString();
            var site2 = TestContext.Properties["SiteACE"].ToString();
            var aircraft1 = TestContext.Properties["AircraftBis"].ToString();
            var aircraft2 = TestContext.Properties["Aircraft"].ToString();
            var route1 = TestContext.Properties["Route"].ToString();
            var route2 = TestContext.Properties["RouteLP"].ToString();
            string comment = TestContext.Properties["CommentLpCart"].ToString();

            //flight
            string flightNumberLpC1 = new Random().Next().ToString();
            string flightNumberLpC2 = new Random().Next().ToString();
            string siteFrom1 = TestContext.Properties["Site"].ToString();
            string siteTo1 = TestContext.Properties["SiteLP"].ToString();
            string siteFrom2 = TestContext.Properties["SiteLP"].ToString();
            string siteTo2 = TestContext.Properties["SiteToFlightBob"].ToString();

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            LpCartPage lpCartPage = homePage.GoToFlights_LpCartPage();
            try
            {
                // create LPCard 1
                var lpCartModalCreate1 = lpCartPage.LpCartCreatePage();
                var lpCartCartDetailPage1 = lpCartModalCreate1.FillField_CreateNewLpCart(
                    code, lpCartName1, site, customer1, aircraft1, from, to, comment);

                //Add aircraft to LPCard 1
                lpCartCartDetailPage1.BackToList();
                lpCartPage.Filter(LpCartPage.FilterType.Search, lpCartName1);
                LpCartCartDetailPage lpCartCartDetailPage = lpCartPage.LpCartCartDetailPage();
                LpCartGeneralInformationPage lpCartGeneralInformationPage = lpCartCartDetailPage.LpCartGeneralInformationPage();
                lpCartGeneralInformationPage.AddRoute(route1);

                // Vérifier LPCard 1
                lpCartCartDetailPage1.BackToList();
                bool aircraftVerified = VerifyRoute(lpCartPage, lpCartName1, route1);
                Assert.IsTrue(aircraftVerified, "Le(s) route(s) ne sont pas affiché(s) sur l'index des LP Carts colonne 'Aircraft Type'");
                lpCartCartDetailPage1.BackToList();

                // create LPCard 2
                var lpCartModalCreate2 = lpCartPage.LpCartCreatePage();
                var lpCartCartDetailPage2 = lpCartModalCreate2.FillField_CreateNewLpCart(
                    code1, lpCartName2, site2, customer1, aircraft2, from, to, comment);

                //Add aircraft to LPCard 2
                lpCartCartDetailPage2.BackToList();
                lpCartPage.Filter(LpCartPage.FilterType.Search, lpCartName2);
                lpCartCartDetailPage2 = lpCartPage.LpCartCartDetailPage();
                lpCartGeneralInformationPage = lpCartCartDetailPage.LpCartGeneralInformationPage();
                lpCartGeneralInformationPage.AddRoute(route2); 

                // Vérifier LPCard 2
                lpCartCartDetailPage2.BackToList();
                aircraftVerified = VerifyRoute(lpCartPage, lpCartName2,route2);
                Assert.IsTrue(aircraftVerified, "Le(s) route(s) ne sont pas affiché(s) sur l'index des LP Carts colonne 'Aircraft Type'");
                lpCartCartDetailPage2.BackToList();

                //create flight 1
                var flightPage = homePage.GoToFlights_FlightPage();
                CreateFlightRoute(flightPage, flightNumberLpC1, customer, aircraft1, siteFrom1, siteTo1);
                VerifyFlightRoute(flightPage, flightNumberLpC1, aircraft1, lpCartName1, siteFrom1, siteTo1);

                //create flight 2
                flightPage = homePage.GoToFlights_FlightPage();
                CreateFlightRoute(flightPage, flightNumberLpC2, customer, aircraft2, siteFrom2, siteTo2);
                VerifyFlightRoute(flightPage, flightNumberLpC2, aircraft2, lpCartName2, siteFrom2, siteTo2);   

            }
            finally
            {
                var lpCartPage0 = homePage.GoToFlights_LpCartPage();
                lpCartPage.ResetFilter();
                lpCartPage.Filter(FilterType.Search, lpCartName1);
                lpCartPage.DeleteLpCart();

                lpCartPage.ResetFilter();
                lpCartPage.Filter(FilterType.Search, lpCartName2);
                lpCartPage.DeleteLpCart();

                var flightPage = homePage.GoToFlights_FlightPage();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumberLpC1);
                flightPage.MassiveDeleteMenus(flightNumberLpC1, siteFrom1, customer, false);

                flightPage.ResetFilter();
                flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumberLpC2);
                flightPage.MassiveDeleteMenus(flightNumberLpC2, siteFrom2, customer, false);
            }
        }

        public bool VerifyRoute(LpCartPage lpCartPage, string lpCartName, string route)
        {
            lpCartPage.ResetFilter();
            lpCartPage.Filter(LpCartPage.FilterType.Search, lpCartName);
            LpCartCartDetailPage lpCartCartDetailPage = lpCartPage.LpCartCartDetailPage();
            LpCartGeneralInformationPage lpCartGeneralInformationPage = lpCartCartDetailPage.LpCartGeneralInformationPage();
            List<string> routes = lpCartGeneralInformationPage.GetSelectedRoutes();
            return routes.Contains(route);
        }

        private void CreateFlightRoute(FlightPage flightPage, string flightNumber, string customer, string aircraft, string siteFrom, string siteTo)
        {
            flightPage.Filter(FlightPage.FilterType.Sites, siteFrom);
            var flightCreatePageModal = flightPage.FlightCreatePage();
            flightCreatePageModal.FillField_CreatNewFlight(flightNumber, customer, aircraft, siteFrom, siteTo);
        }

        private void VerifyFlightRoute(FlightPage flightPage, string flightNumber, string aircraft, string lpCartName, string siteFrom, string siteTo)
        {
            flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
            var from = flightPage.GetFromValue();
            var to = flightPage.GetToValue();
            Assert.AreEqual(from, siteFrom);
            Assert.AreEqual(to, siteTo);
            var flightDetailsPage = flightPage.EditFirstFlight(flightNumber);
            flightDetailsPage.ChangeAircaftForLoadingPlan(aircraft);

            //Assert
            bool lPCartIsExist = flightDetailsPage.LPCartIsExist(lpCartName);
            Assert.IsTrue(lPCartIsExist, $"le champ 'LP Cart' : {lpCartName} ne propose pas l'option de LP Cart éditée comme prévu");
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_EditLPCart_To()
        {
            // Prepare
            //lpCart & cart
            var code = new Random()
                .Next(0, 5000)
                .ToString();
            var to = DateTime.Today.AddDays(-5);
            var customerFilter = TestContext.Properties["CustomerLpFilter"].ToString();
            var site = TestContext.Properties["SiteACE"].ToString();
            var siteFilter = TestContext.Properties["SiteACEFilter"].ToString();
            var aircraft = TestContext.Properties["AircraftLpCart"].ToString();
            var lpCartName = $"LP Cart {DateTime.Now}";
            var fromLpCart = DateTime.Today.AddMonths(-3);
            var toLpCart = DateTime.Today.AddMonths(+2);

            //Flight
            string flightNumber = new Random().Next().ToString();
            string siteTo = TestContext.Properties["SiteToFlightBob"].ToString();
            string etaHours = "00";
            string etdHours = "23";
            string loadingPlanName = null;
            DateTime date = DateTime.Today.AddDays(+1);
            string lpCart = null;

            //Arrange
            var homePage = LogInAsAdmin();

            //Act
            //Create lp cart 
            #region LP cart - Update the existing LP cart with the dates (from - to ) with less than today.
            var lpCartPage = homePage.GoToFlights_LpCartPage();
            FilterLpCart();

            bool isLpCartCreated = false;
            bool isFlightCreated = false;
            try
            {
                var lpCartModalCreate = lpCartPage.LpCartCreatePage();
                //Create lp cart and back to list
                var lpCartCartDetailPage = lpCartModalCreate
                    .FillField_CreateNewLpCart(code, lpCartName, site, customerFilter, aircraft, fromLpCart, toLpCart, "comment");
                isLpCartCreated = true;

                lpCartCartDetailPage.BackToList();

                lpCartPage.Filter(FilterType.Search, lpCartName);
                lpCartCartDetailPage = lpCartPage.ClickFirstLpCart();
                var lpCartGeneralInformation = lpCartCartDetailPage.LpCartGeneralInformationPage();

                //Update date from to befor today 
                lpCartGeneralInformation
                    .UpdateToDate(to);

                lpCartCartDetailPage.BackToList();
                lpCartPage.Filter(FilterType.Search, lpCartName);

                lpCartCartDetailPage = lpCartPage.ClickFirstLpCart();
                lpCartGeneralInformation = lpCartCartDetailPage.LpCartGeneralInformationPage();

                var dateToUpdated = DateTime.ParseExact(lpCartGeneralInformation.GetLpCartDateTo(), "dd/MM/yyyy", CultureInfo.InvariantCulture);

                //Assert
                Assert.AreEqual(dateToUpdated, to, " la dateFrom modifiée n'est pas bien enregistrée dans la général information. ");
                #endregion

                #region Flight

                //update flight with aircraft
                homePage.Navigate();
                var flightPage = homePage.GoToFlights_FlightPage();
                flightPage.ResetFilter();
                flightPage.ShowCustomersMenu();
                flightPage.Filter(FlightPage.FilterType.Customer, customerFilter);
                flightPage.Filter(FlightPage.FilterType.Sites, site);

                var flightCreatePageModal = flightPage.FlightCreatePage();

                flightCreatePageModal.FillField_CreatNewFlight(flightNumber, customerFilter, aircraft, site, siteTo, loadingPlanName, etaHours, etdHours, lpCart, date);
                isFlightCreated = true;

                var flightEditPage = flightPage.EditFirstFlight(flightNumber);

                flightEditPage.ChangeAircaftForLoadingPlan(aircraft);

                var lPCartIsExist = flightEditPage.LPCartIsExist(lpCartName);

                Assert.IsTrue(!lPCartIsExist, $"LP cart existe dans le flight, car le date from inferieur a la date du flight.");

                #endregion
            }

            //Suppression de LPCart et Flight crées
            finally
            {
                if (isLpCartCreated)
                {
                    lpCartPage = homePage.GoToFlights_LpCartPage();
                    lpCartPage.ResetFilter();
                    lpCartPage.Filter(FilterType.Search, lpCartName);
                    lpCartPage.DeleteLpCart();
                    lpCartPage.ResetFilter();
                }

                if (isFlightCreated)
                {
                    var flightPage = homePage.GoToFlights_FlightPage();
                    flightPage.Filter(FlightPage.FilterType.SearchFlight, flightNumber);
                    flightPage.MassiveDeleteMenus(flightNumber, site, null, false);
                    flightPage.ResetFilter();
                }
            }

            void FilterLpCart()
            {
                lpCartPage.Filter(FilterType.Customers, customerFilter);
                lpCartPage.Filter(FilterType.Site, siteFilter);
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_CartsTab_SortByCode()
        {
            // Prepare
            //lpCart & cart
            var code = new Random()
                .Next(0, 5000)
                .ToString();
            var customerFilter = TestContext.Properties["CustomerLpFilter"].ToString();
            var site = TestContext.Properties["SiteACE"].ToString();
            var aircraft = TestContext.Properties["AircraftLpCart"].ToString();
            var lpCartName = $"LP Cart {DateTime.Now}";
            var fromLpCart = DateTime.Today.AddMonths(-3);
            var toLpCart = DateTime.Today.AddMonths(+2);
            var comment = "comment";
            //Arrange
            HomePage homePage = LogInAsAdmin();
            //Act
            LpCartPage lpCartPage = homePage.GoToFlights_LpCartPage();
            try
            {
                var lpCartModalCreate = lpCartPage.LpCartCreatePage();
                //Create lp cart 
                LpCartCartDetailPage lpCartCartDetailPage = lpCartModalCreate
                    .FillField_CreateNewLpCart(code, lpCartName, site, customerFilter, aircraft, fromLpCart, toLpCart, comment);
                lpCartCartDetailPage.ClickAddtrolley();
                lpCartCartDetailPage.AddTrolleyCustom("1", "code1", "name 1", 0);
                lpCartCartDetailPage.ClickAddtrolley();
                lpCartCartDetailPage.AddTrolleyCustom("2","code2", "name 2", 1);
                lpCartCartDetailPage.ClickAddtrolley();
                lpCartCartDetailPage.AddTrolleyCustom("3","code3", "name 3", 2);

                //Assert
                var checkSortByLegAndOrder = lpCartCartDetailPage.VerifySortByCode();
                Assert.IsTrue(checkSortByLegAndOrder, "Le SortByCode pour les Carts ne fonctionne pas correctement.");
            }
            finally
            {
                lpCartPage = homePage.GoToFlights_LpCartPage();
                lpCartPage.ResetFilter();
                lpCartPage.Filter(FilterType.Search, lpCartName);
                LpCartCartDetailPage lpCartCartDetailPage = lpCartPage.LpCartCartDetailPage();
                //Delete Carts
                lpCartCartDetailPage.DeleteAllLpCartScheme();
                lpCartCartDetailPage.BackToList();
                //Delete LpCart
                lpCartPage.DeleteLpCart();
                lpCartPage.ResetFilter();
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_CartsTab_SortByLegAndOrder()
        {
            // Prepare
            //lpCart & cart
            var code = new Random()
                .Next(0, 5000)
                .ToString();
            var customerFilter = TestContext.Properties["CustomerLpFilter"].ToString();
            var site = TestContext.Properties["SiteACE"].ToString();
            var aircraft = TestContext.Properties["AircraftLpCart"].ToString();
            var lpCartName = $"LP Cart {DateTime.Now}";
            var fromLpCart = DateTime.Today.AddMonths(-3);
            var toLpCart = DateTime.Today.AddMonths(+2);
            var comment = "comment";
            string code1 = "code1";
            string name1 = "name1";
            string leg1 = "5";
            int row1 = 0;
            string code2 = "code2";
            string name2 = "name2";
            string leg2 = "4";
            int row2 = 1;
            string code3 = "code3";
            string name3 = "name3";
            string leg3 = "1";
            int row3 = 2;
            string LegAndOrder = "LegAndOrder";
            //Arrange
            HomePage homePage = LogInAsAdmin();

            //Act
            LpCartPage lpCartPage = homePage.GoToFlights_LpCartPage();
            try
            {
                var lpCartModalCreate = lpCartPage.LpCartCreatePage();
                //Create lp cart 
                LpCartCartDetailPage lpCartCartDetailPage = lpCartModalCreate
                    .FillField_CreateNewLpCart(code, lpCartName, site, customerFilter, aircraft, fromLpCart, toLpCart, comment);

                lpCartCartDetailPage.ClickAddtrolley();
                lpCartCartDetailPage.AddTrolleyCustom(leg1, code1, name1, row1);
                lpCartCartDetailPage.ClickAddtrolley();
                lpCartCartDetailPage.AddTrolleyCustom(leg2, code2, name2, row2);
                lpCartCartDetailPage.ClickAddtrolley();
                lpCartCartDetailPage.AddTrolleyCustom(leg3, code3, name3, row3);
                WebDriver.Navigate().Refresh();
                lpCartCartDetailPage.Filter(LpCartCartDetailPage.FilterType.SortBy, LegAndOrder);
                lpCartCartDetailPage.WaitPageLoading();
                //Assert
                var checkSortByLegAndOrder = lpCartCartDetailPage.VerifyFilterByLegOrder();
                Assert.IsTrue(checkSortByLegAndOrder, "Le SortByleg pour les Carts ne fonctionne pas correctement.");
            }
            finally
            {
                lpCartPage = homePage.GoToFlights_LpCartPage();
                lpCartPage.ResetFilter();
                lpCartPage.Filter(FilterType.Search, lpCartName);
                lpCartPage.DeleteLpCart();
                lpCartPage.ResetFilter();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void FL_LP_CART_FlightTab_FilterStartDate()
        {
            // Prepare
            Random rnd = new Random();
            string code = rnd.Next(500, 6000).ToString();
            string name = "LP cart " + code;
            string customer = TestContext.Properties["CustomerLpCart"].ToString();
            string aircraft = TestContext.Properties["AircraftLpCart"].ToString();
            string siteTo = TestContext.Properties["SiteToFlightBob"].ToString();
            string comment = TestContext.Properties["CommentLpCart"].ToString();
            string siteFrom = TestContext.Properties["SiteLpCart"].ToString();
            string etaHours = "00";
            string etdHours = "23";
            DateTime from = DateUtils.Now;
            DateTime flight1Date = DateUtils.Now.AddDays(3);
            DateTime flight2Date = DateUtils.Now.AddDays(7);
            DateTime to = DateUtils.Now.AddDays(10);
            string flightNumber1 = new Random().Next().ToString();
            string flightNumber2 = flightNumber1 + 25;

            // Act
            HomePage homePage = LogInAsAdmin();
            string dateFormat = homePage.GetDateFormatPickerValue();
            LpCartPage lpCartPage = homePage.GoToFlights_LpCartPage();
            try
            {
                LpCartCreateModalPage lpCartModalCreate = lpCartPage.LpCartCreatePage();
                lpCartModalCreate.FillField_CreateNewLpCart(code, name, siteFrom, customer, aircraft, from, to, comment);
                FlightPage flightPage = homePage.GoToFlights_FlightPage();
                FlightCreateModalPage createFlightpage = flightPage.FlightCreatePage();
                createFlightpage.FillField_CreatNewFlight(flightNumber1, customer, aircraft, siteFrom, siteTo, null, etaHours, etdHours, name, flight1Date);
                createFlightpage = flightPage.FlightCreatePage();
                createFlightpage.FillField_CreatNewFlight(flightNumber2, customer, aircraft, siteFrom, siteTo, null, etaHours, etdHours, name, flight2Date);
                homePage.GoToFlights_LpCartPage();
                lpCartPage.Filter(LpCartPage.FilterType.Search, name);
                LpCartCartDetailPage lpCartDetail = lpCartPage.LpCartCartDetailPage();
                LpCartFlightDetailPage lpCartFlightdetail = lpCartDetail.LpCartFlightDetailPage();
                int totFlightNumber = lpCartFlightdetail.CheckTotalNumber();
                //Filter by Search          
                lpCartFlightdetail.Filter(LpCartFlightDetailPage.FilterType.StartDate, flight1Date);
                lpCartFlightdetail.CloseDatePicker();
                bool isdaterespected = lpCartFlightdetail.IsDateRespected(flight1Date.Date, to.Date, dateFormat);
                Assert.IsTrue(isdaterespected, string.Format(MessageErreur.FILTRE_ERRONE, "Date"));

                lpCartFlightdetail.Filter(LpCartFlightDetailPage.FilterType.StartDate, flight2Date);
                int flightNumber = lpCartFlightdetail.CheckTotalNumber();
                Assert.AreNotEqual(totFlightNumber, flightNumber, string.Format(MessageErreur.FILTRE_ERRONE, "Date"));

                isdaterespected = lpCartFlightdetail.IsDateRespected(flight2Date.Date, to.Date, dateFormat);
                Assert.IsTrue(isdaterespected, string.Format(MessageErreur.FILTRE_ERRONE, "Date"));
            }
            finally
            {
                //Delete Flights
                FlightPage flightPage = homePage.GoToFlights_FlightPage();
                FlightMassiveDeleteModalPage massiveDeleteFlightPage = flightPage.ClickMassiveDelete();
                massiveDeleteFlightPage.SetFlightName(flightNumber1);
                massiveDeleteFlightPage.ClickSearchButton();
                massiveDeleteFlightPage.ClickSelectAllButton();
                massiveDeleteFlightPage.Delete();

                //Delete LpCart
                homePage.GoToFlights_LpCartPage();
                lpCartPage.ResetFilter();
                lpCartPage.Filter(LpCartPage.FilterType.Search, name);
                lpCartPage.DeleteLpCart();
            }
        }

        [Timeout(_timeout)]
        [TestMethod]
        public void FL_LP_CART_FlightTab_FilterEndDate()
        {
            // Prepare
            Random rnd = new Random();
            string code = rnd.Next(500, 6000).ToString();
            string name = "LP cart " + code;
            string customer = TestContext.Properties["CustomerLpCart"].ToString();
            string aircraft = TestContext.Properties["AircraftLpCart"].ToString();
            string siteTo = TestContext.Properties["SiteToFlightBob"].ToString();
            string comment = TestContext.Properties["CommentLpCart"].ToString();
            string siteFrom = TestContext.Properties["SiteLpCart"].ToString();
            string etaHours = "00";
            string etdHours = "23";
            DateTime from = DateUtils.Now;
            DateTime flight1Date = DateUtils.Now.AddDays(3);
            DateTime flight2Date = DateUtils.Now.AddDays(7);
            DateTime to = DateUtils.Now.AddDays(10);
            string flightNumber1 = new Random().Next().ToString();
            string flightNumber2 = flightNumber1 + 25;

            // Act
            HomePage homePage = LogInAsAdmin();
            string dateFormat = homePage.GetDateFormatPickerValue();
            LpCartPage lpCartPage = homePage.GoToFlights_LpCartPage();
            try
            {
                LpCartCreateModalPage lpCartModalCreate = lpCartPage.LpCartCreatePage();
                lpCartModalCreate.FillField_CreateNewLpCart(code, name, siteFrom, customer, aircraft, from, to, comment);
                FlightPage flightPage = homePage.GoToFlights_FlightPage();
                FlightCreateModalPage createFlightpage = flightPage.FlightCreatePage();
                createFlightpage.FillField_CreatNewFlight(flightNumber1, customer, aircraft, siteFrom, siteTo, null, etaHours, etdHours, name, flight1Date);
                createFlightpage = flightPage.FlightCreatePage();
                createFlightpage.FillField_CreatNewFlight(flightNumber2, customer, aircraft, siteFrom, siteTo, null, etaHours, etdHours, name, flight2Date);
                homePage.GoToFlights_LpCartPage();
                lpCartPage.Filter(LpCartPage.FilterType.Search, name);
                LpCartCartDetailPage lpCartDetail = lpCartPage.LpCartCartDetailPage();
                LpCartFlightDetailPage lpCartFlightdetail = lpCartDetail.LpCartFlightDetailPage();
                int totFlightNumber = lpCartFlightdetail.CheckTotalNumber();
                //Filter by Search          
                lpCartFlightdetail.Filter(LpCartFlightDetailPage.FilterType.EndDate, flight2Date);
                bool isdaterespected = lpCartFlightdetail.IsDateRespected(from.Date, flight2Date.Date, dateFormat);
                Assert.IsTrue(isdaterespected, string.Format(MessageErreur.FILTRE_ERRONE, "Date"));

                lpCartFlightdetail.Filter(LpCartFlightDetailPage.FilterType.EndDate, flight1Date);
                int flightNumber = lpCartFlightdetail.CheckTotalNumber();
                Assert.AreNotEqual(totFlightNumber, flightNumber, string.Format(MessageErreur.FILTRE_ERRONE, "Date"));

                isdaterespected = lpCartFlightdetail.IsDateRespected(from.Date, flight1Date.Date, dateFormat);
                Assert.IsTrue(isdaterespected, string.Format(MessageErreur.FILTRE_ERRONE, "Date"));
            }
            finally
            {
                //Delete Flights
                FlightPage flightPage = homePage.GoToFlights_FlightPage();
                FlightMassiveDeleteModalPage massiveDeleteFlightPage = flightPage.ClickMassiveDelete();
                massiveDeleteFlightPage.SetFlightName(flightNumber1);
                massiveDeleteFlightPage.ClickSearchButton();
                massiveDeleteFlightPage.ClickSelectAllButton();
                massiveDeleteFlightPage.Delete();

                //Delete LpCart
                homePage.GoToFlights_LpCartPage();
                lpCartPage.ResetFilter();
                lpCartPage.Filter(LpCartPage.FilterType.Search, name);
                lpCartPage.DeleteLpCart();
            }
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_Index_FilterByFlightStartDate()
        {
            // Prepare
            DateTime from1 = DateUtils.Now.AddDays(10);
            DateTime from2 = DateUtils.Now.AddDays(20);

            // Act
            HomePage homePage = LogInAsAdmin();
            string dateFormat = homePage.GetDateFormatPickerValue();
            LpCartPage lpCartPage = homePage.GoToFlights_LpCartPage();
            lpCartPage.ResetFilter();
            lpCartPage.Filter(FilterType.ByFlightdate, true);
            int totLPCartNumber = lpCartPage.CheckTotalNumber();
            lpCartPage.Filter(FilterType.StartDate, from1);
            lpCartPage.CloseDatePicker();
            int lPCartNumber = lpCartPage.CheckTotalNumber();
            //Assert filtre 1
            Assert.AreNotEqual(totLPCartNumber, lPCartNumber, string.Format(MessageErreur.FILTRE_ERRONE, "Date"));

            lpCartPage.Filter(FilterType.StartDate, from2);
            lPCartNumber = lpCartPage.CheckTotalNumber();
            //Assert filtre 2
            Assert.AreNotEqual(totLPCartNumber, lPCartNumber, string.Format(MessageErreur.FILTRE_ERRONE, "Date"));

        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_Index_FilterByFlightEndDate()
        {
            // Prepare
            DateTime to1test = DateUtils.Now.AddDays(3);
            DateTime to2test = DateUtils.Now.AddDays(1);

            // Act
            HomePage homePage = LogInAsAdmin();
            LpCartPage lpCartPage = homePage.GoToFlights_LpCartPage();
            lpCartPage.ResetFilter();
            lpCartPage.Filter(FilterType.ByFlightdate, true);
            int totLPCartNumber = lpCartPage.CheckTotalNumber();
            lpCartPage.Filter(FilterType.EndDate, to1test);
            int lPCartNumber = lpCartPage.CheckTotalNumber();
            //Assert filter 1
            Assert.AreNotEqual(totLPCartNumber, lPCartNumber, string.Format(MessageErreur.FILTRE_ERRONE, "Date"));

            lpCartPage.Filter(FilterType.EndDate, to2test);
            lPCartNumber = lpCartPage.CheckTotalNumber();
            //Assert filtre 2
            Assert.AreNotEqual(totLPCartNumber, lPCartNumber, string.Format(MessageErreur.FILTRE_ERRONE, "Date"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_Index_FilterByLPCartValidityStartDate()
        {
            // Prepare
            DateTime from1 = DateUtils.Now.AddDays(10);
            DateTime from2 = DateUtils.Now.AddDays(20);

            // Act
            HomePage homePage = LogInAsAdmin();
            string dateFormat = homePage.GetDateFormatPickerValue();
            LpCartPage lpCartPage = homePage.GoToFlights_LpCartPage();
            lpCartPage.ResetFilter();
            lpCartPage.Filter(FilterType.ValidityDate, true);
            int totLPCartNumber = lpCartPage.CheckTotalNumber();
            lpCartPage.Filter(FilterType.StartDate, from1);
            lpCartPage.CloseDatePicker();
            int lPCartNumber = lpCartPage.CheckTotalNumber();
            //Assert filtre 1
            Assert.AreNotEqual(totLPCartNumber, lPCartNumber, string.Format(MessageErreur.FILTRE_ERRONE, "Date"));

            lpCartPage.Filter(FilterType.StartDate, from2);
            lPCartNumber = lpCartPage.CheckTotalNumber();
            //Assert filtre 2
            Assert.AreNotEqual(totLPCartNumber, lPCartNumber, string.Format(MessageErreur.FILTRE_ERRONE, "Date"));
        }

        [TestMethod]
        [Timeout(_timeout)]
        public void FL_LP_CART_Index_FilterByLPCartValidityEndDate()
        {
            // Prepare
            DateTime to1test = DateUtils.Now.AddDays(3);
            DateTime to2test = DateUtils.Now.AddDays(1);

            // Act
            HomePage homePage = LogInAsAdmin();
            LpCartPage lpCartPage = homePage.GoToFlights_LpCartPage();
            lpCartPage.ResetFilter();
            lpCartPage.Filter(FilterType.ValidityDate, true);
            int totLPCartNumber = lpCartPage.CheckTotalNumber();
            lpCartPage.Filter(FilterType.EndDate, to1test);
            int lPCartNumber = lpCartPage.CheckTotalNumber();
            //Assert filter 1
            Assert.AreNotEqual(totLPCartNumber, lPCartNumber, string.Format(MessageErreur.FILTRE_ERRONE, "Date"));

            lpCartPage.Filter(FilterType.EndDate, to2test);
            lPCartNumber = lpCartPage.CheckTotalNumber();
            //Assert filtre 2
            Assert.AreNotEqual(totLPCartNumber, lPCartNumber, string.Format(MessageErreur.FILTRE_ERRONE, "Date"));
        }
    }
}