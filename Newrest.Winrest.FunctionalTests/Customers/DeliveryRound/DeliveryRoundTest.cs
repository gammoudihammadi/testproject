using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Delivery;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.DeliveryRound;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Policy;

namespace Newrest.Winrest.FunctionalTests.Customers
{
    [TestClass]
    public class DeliveryRoundTest : TestBase
    {
        private const int _timeout = 600000; 
        private const string DELIVERY_ROUND_EXCEL_SHEET_NAME = "Deliveries";

        readonly string deliveryRoundNameToday = "deliveryRound-" + DateUtils.Now.ToString("dd/MM/yyyy");

        string deliveryName = "deliveryForDeliveryRound-" + DateUtils.Now.ToString("dd/MM/yyyy");
        private static Random random = new Random();
        string deliveryNameForDR = "";
        string deliveryNameForDR1 = "";
        string deliveryNameForDR2 = "";
        //Dates For From
        private static DateTime fromDate = DateTime.Today;
        private static DateTime dateBeforeFrom = DateTime.Today.AddDays(-10);
        private static DateTime dateAfterFrom = DateTime.Today.AddDays(+5);
        //Dates For To
        private static DateTime dateBeforeTo = DateTime.Today.AddDays(+21);
        private static DateTime dateAfterTo = DateTime.Today.AddDays(+41);
        private static DateTime toDate = DateTime.Today.AddDays(+31);

        private static DateTime dateTimeUnique = DateTime.Today;
        private static DateTime deliveryRoundNameDateTo = DateTime.Today;

        //Dates Names For From
        string deliveryRoundNameBeforeDateFrom = "BeforeDateFrom-" + dateTimeUnique.ToString("dd/MM/yyyy") + "-" +random.Next(100, 999);
        string deliveryRoundNameEqualDateFrom = "EqualDateFrom-" + dateTimeUnique.ToString("dd/MM/yyyy") + "-" + random.Next(100, 999);
        string deliveryRoundNameAfterDateFrom = "AfterDateFrom-" + dateTimeUnique.ToString("dd/MM/yyyy") + random.Next(100, 999);

        //Dates Names For To
        string deliveryRoundNameEqualDateTo = "EqualDateTo-" + dateTimeUnique.ToString("dd/MM/yyyy") + "-" + random.Next(100, 999);
        string deliveryRoundNameAfterDateTo = "AfterToDate-" + dateTimeUnique.ToString("dd/MM/yyyy") + "-" + random.Next(100, 999);
        string deliveryRoundNameBeforeDateTo = "BeforeDateTo-" + dateTimeUnique.ToString("dd/MM/yyyy") + "-" + random.Next(100, 999);
        string deliveryRoundNameDateFrom = "DateFrom-" + dateTimeUnique.ToString("dd/MM/yyyy") + "-" + random.Next(100, 999);

        // ______________________________________ DELIVERY_ROUND ___________________________________________________
        [TestInitialize]
        public override void TestInitialize()
        {
            base.TestInitialize();

            var testMethod = TestContext.TestName;
            switch (testMethod)
            {
                case nameof(CU_DERO_Filter_DateFrom):
                    CreateDeliveryForDeliveryRound_TestInitialize();
                    CU_DERO_CreateDero_FilterDateFrom_TestInitialize();
                    break;
                case nameof(CU_DERO_Create_New_Delivery_Round):
                    CreateDeliveryForDeliveryRound_TestInitialize();
                    break; 
                case nameof(CU_DERO_Create_Delivery):
                    CreateDeliveryForDeliveryRound_TestInitialize();
                    break; 
                case nameof(CU_DERO_Check_Delivery):
                    CreateDeliveryForDeliveryRound_TestInitialize();
                    break; 
                case nameof(CU_DERO_Change_General_Information):
                    CreateDeliveryForDeliveryRound_TestInitialize();
                    break; 
                case nameof(CU_DERO_Delete_Delivery):
                    CreateDeliveryForDeliveryRound_TestInitialize();
                    break; 
                case nameof(CU_DERO_Change_Delivery_Round_Name_Already_Existing):
                    CreateDeliveryForDeliveryRound_TestInitialize();
                    break; 
                case nameof(CU_DERO_Filter_Customers):
                    CreateDeliveryForDeliveryRound_TestInitialize();
                    break;
                     case nameof(CU_DERO_Export_Delivery_Round_NewVersion):
                    CreateDeliveryForDeliveryRound_TestInitialize();
                    break;
                case nameof(CU_DERO_Filter_ShowInactive):
                    CreateDeliveryForDeliveryRound_TestInitialize();
                    CU_DERO_CreateDero_Delivery_TestInitialize();
                    break;
                case nameof(CU_DERO_Export_Filter_Site_Delivery_Round_NewVersion):
                    CreateDeliveryForDeliveryRound_TestInitialize();
                    CU_DERO_CreateDero_Delivery_TestInitialize();
                    break;
                case nameof(CU_DERO_Duplicate_Delivery_Round):
                    CreateDeliveryForDeliveryRound_TestInitialize();
                 
                    break;
                case nameof(CU_DERO_Export_Filter_Search_Delivery_Round_NewVersion):
                    CreateDeliveryForDeliveryRound_TestInitialize();

                    break;
                case nameof(CU_DERO_Filter_SortByName):
                    CreateDeliveryForDeliveryRound_TestInitialize();
                    break;

                case nameof(CU_DERO_Filter_SortBySite):
                    CreateDeliveryForDeliveryRound_TestInitialize();
                    break;
                case nameof(CU_DERO_NewDR):
                    CreateDeliveriesForDeliveryRound_TestInitialize();
                    break;

                case nameof(CU_DERO_Filter_DateTo):
                    CreateDeliveryForDeliveryRound_TestInitialize();
                    CU_DERO_CreateDero_FilterDateTo_TestInitialize();
                    break;

                default:
                    break;
            }
        }
       
        [TestCleanup]
        public override void TestCleanup()
        {

            var testMethod = TestContext.TestName;
            switch (testMethod)
            {
                case nameof(CU_DERO_Filter_DateFrom):
                    CU_DERO_CreateDero_FilterDateFromTo_TestCleanUp();
                    DeleteDeliveryForDeliveryRound_TestCleanUp();
                    break;
                case nameof(CU_DERO_Export_Delivery_Round_NewVersion):
                    DeleteDeliveryForDeliveryRound_TestCleanUp();
                    break;
                case nameof(CU_DERO_Filter_ShowInactive):
                    CU_DERO_DeleteeDero_Delivery_TestCleanUp();
                    DeleteDeliveryForDeliveryRound_TestCleanUp();
                    break;
                case nameof(CU_DERO_Export_Filter_Site_Delivery_Round_NewVersion):
                    DeleteDeliveryForDeliveryRound_TestCleanUp();
                    break;
                case nameof(CU_DERO_Duplicate_Delivery_Round):
                    DeleteDeliveryForDeliveryRound_TestCleanUp();
                    break;
                case nameof(CU_DERO_Export_Filter_Search_Delivery_Round_NewVersion):
                    DeleteDeliveryForDeliveryRound_TestCleanUp();
                    break;
                case nameof(CU_DERO_Filter_SortByName):
                    DeleteDeliveryForDeliveryRound_TestCleanUp();
                    break;
                case nameof(CU_DERO_Filter_SortBySite):
                    DeleteDeliveryForDeliveryRound_TestCleanUp();
                    break;
                case nameof(CU_DERO_NewDR):
                    DeliveriesFrDR_TestCleanUp();
                    break;
                case nameof(CU_DERO_Filter_DateTo):
                    CU_DERO_CreateDero_FilterDateFromTo_TestCleanUp();
                    DeleteDeliveryForDeliveryRound_TestCleanUp();
                    break;
                default:
                    break;
            }
            base.TestCleanup();
        }

        private void CU_DERO_CreateDero_FilterDateFrom_TestInitialize()
        {
            string site = TestContext.Properties["Site"].ToString();

            TestContext.Properties[string.Format("DeliveryId1")] = InsertDeliveryRound(deliveryRoundNameBeforeDateFrom, dateBeforeFrom, dateBeforeFrom, site);
            TestContext.Properties[string.Format("DeliveryId2")] = InsertDeliveryRound(deliveryRoundNameEqualDateFrom, fromDate, fromDate, site);
            TestContext.Properties[string.Format("DeliveryId3")] = InsertDeliveryRound(deliveryRoundNameAfterDateFrom, dateAfterFrom, dateAfterFrom, site);
        }

        private void CU_DERO_CreateDero_FilterDateTo_TestInitialize()
        {
            string site = TestContext.Properties["Site"].ToString();

            TestContext.Properties[string.Format("DeliveryId1")] = InsertDeliveryRound(deliveryRoundNameBeforeDateTo, dateBeforeTo, dateBeforeTo, site);
            TestContext.Properties[string.Format("DeliveryId2")] = InsertDeliveryRound(deliveryRoundNameEqualDateTo, toDate, toDate, site);
            TestContext.Properties[string.Format("DeliveryId3")] = InsertDeliveryRound(deliveryRoundNameAfterDateTo, dateAfterTo, dateAfterTo, site);
        }
        private void CU_DERO_CreateDero_Delivery_TestInitialize()
        {
            string site = TestContext.Properties["Site"].ToString();

            TestContext.Properties[string.Format("DeliveryId")] = InsertDeliveryRound(deliveryRoundNameDateFrom, deliveryRoundNameDateTo, deliveryRoundNameDateTo, site);
      
        }
        private void CU_DERO_DeleteeDero_Delivery_TestCleanUp()
        {
            var deliveryId = TestContext.Properties[string.Format("DeliveryId")];
            Console.WriteLine(deliveryId);
            DeleteDeliveryRound(int.Parse(deliveryId.ToString()));
      
        }
        private void CU_DERO_CreateDero_FilterDateFromTo_TestCleanUp()
        {
            List<int> deliveryRoundIds = new List<int>();

            for (int i = 1; i <= 6; i++)
            {
                if (TestContext.Properties.Contains($"DeliveryId{i}"))
                {
                    deliveryRoundIds.Add((int)TestContext.Properties[$"DeliveryId{i}"]);
                }
            }

            DeleteMultipleDeliveryRounds(deliveryRoundIds);
        }

        // Effectuer des recherches via les filtres - From/To
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_DERO_Filter_DateFrom()
        {
            // Arrange
            var homePage = LogInAsAdmin();
            // Act
            var deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            deliveryRoundPage.ResetFilter();
            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.DateFrom, fromDate);
            //Asserts
            VerifyDeliveryRound(deliveryRoundPage, deliveryRoundNameBeforeDateFrom, false, "deliveryRoundNameBeforeDateFrom affiché dans la liste des DRs");
            VerifyDeliveryRound(deliveryRoundPage, deliveryRoundNameEqualDateFrom, true, "deliveryRoundNameEqualDateFrom n'est pas affiché dans la liste des DRs");
            VerifyDeliveryRound(deliveryRoundPage, deliveryRoundNameAfterDateFrom, true, "deliveryRoundNameAfterDateTo n'est pas affiché dans la liste des DRs");

        }

        [Timeout(_timeout)]
        [TestMethod]
        public void CU_DERO_Filter_DateTo()
        {
            // Arrange
            var homePage = LogInAsAdmin();
            // Act
            var deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            deliveryRoundPage.ResetFilter();
            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.DateTo, toDate);
            //Assert
            VerifyDeliveryRound(deliveryRoundPage, deliveryRoundNameBeforeDateTo, true, "deliveryRoundNameAfterDateTo n'est pas affiché dans la liste des DRs");
            VerifyDeliveryRound(deliveryRoundPage, deliveryRoundNameEqualDateTo, true, "deliveryRoundNameEqualDateTo n'est pas affiché dans la liste des DRs");
            VerifyDeliveryRound(deliveryRoundPage, deliveryRoundNameAfterDateTo, false, "deliveryRoundNameAfterDateTo est affiché dans la liste des DRs");

        }
        private void CreateDeliveryForDeliveryRound_TestInitialize()
        {
            // Prepare
            string deliveryCustomer = TestContext.Properties["CustomerDeliveryRound"].ToString();
            string deliverySite = TestContext.Properties["Site"].ToString();
            deliveryNameForDR = "deliveryForDeliveryRound - " + DateUtils.Now.ToString("dd/MM/yyyy") + "-" + random.Next(10,999);

            TestContext.Properties[string.Format("DeliveryFrDRId")] =  InsertDeliveryForDeliveryRound(deliveryNameForDR, fromDate, deliverySite, deliveryCustomer);
        }
        
        private void CreateDeliveriesForDeliveryRound_TestInitialize()
        {
            // Prepare
            string deliveryCustomer = TestContext.Properties["CustomerDeliveryRound"].ToString();
            string deliverySite = TestContext.Properties["Site"].ToString();
            deliveryNameForDR1 = "deliveryForDeliveryRound1 - " + DateUtils.Now.ToString("dd / MM / yyyy") + "-" + random.Next(10,999);
            deliveryNameForDR2 = "deliveryForDeliveryRound2 - " + DateUtils.Now.ToString("dd / MM / yyyy") + "-" + random.Next(10,999);

            TestContext.Properties[string.Format("DeliveryFrDRId1")] =  InsertDeliveryForDeliveryRound(deliveryNameForDR1, fromDate, deliverySite, deliveryCustomer);
            Assert.IsNotNull(DeliveryExist((int)TestContext.Properties[string.Format("DeliveryFrDRId1")]), "Le delivery n'est pas crée.");

            TestContext.Properties[string.Format("DeliveryFrDRId2")] =  InsertDeliveryForDeliveryRound(deliveryNameForDR2, fromDate, deliverySite, deliveryCustomer);
            Assert.IsNotNull(DeliveryExist((int)TestContext.Properties[string.Format("DeliveryFrDRId2")]), "Le delivery n'est pas crée.");
        }

        private void DeleteDeliveryForDeliveryRound_TestCleanUp()
        {
            if (TestContext.Properties.Contains($"DeliveryFrDRId"))
            {
                DeleteDelivery((int)TestContext.Properties["DeliveryFrDRId"]);
            }
        }
        private void DeleteDeliveriesForDeliveryRound_TestCleanUp()
        {
            if (TestContext.Properties.Contains($"DeliveryFrDRId"))
            {
                DeleteDelivery((int)TestContext.Properties["DeliveryFrDRId"]);
            }
        }

        //[Priority(0)]
        //[Timeout(_timeout)]
		//[TestMethod]
        //public void CU_DERO_CreateDeliveryForDeliveryRound()
        //{
        //    // Prepare
        //    string deliveryCustomer = "A.F.B. LANZAROTE";
        //    string deliverySite = TestContext.Properties["Site"].ToString();

        //    // Arrange
        //    LogInAsAdmin();
        //    var homePage = new HomePage(WebDriver, TestContext);
        //    homePage.Navigate();

        //    // Act
        //    var deliveryPage = homePage.GoToCustomers_DeliveryPage();
        //    deliveryPage.ResetFilter();
        //    deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName);

        //    if (deliveryPage.CheckTotalNumber() == 0)
        //    {
        //        var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
        //        deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, deliveryCustomer, deliverySite, true);
        //        var deliveryLoadingPage = deliveryCreateModalPage.Create();
        //        deliveryPage = deliveryLoadingPage.BackToList();
        //        deliveryPage.ResetFilter();
        //        deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName);
        //    }

        //    // Assert
        //    Assert.AreEqual(deliveryPage.GetFirstDeliveryName(), deliveryName, "Le delivery n'a pas été créé.");
        //}

        // Créer une nouvelle tournée de livraison
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_DERO_Create_New_Delivery_Round()
        {
            // Prepare
            Random rnd = new Random();
            string site = TestContext.Properties["Site"].ToString();

            string deliveryRoundName = deliveryRoundNameToday + "-" + rnd.Next(1000, 9000).ToString();


            // Arrange
            HomePage homePage = LogInAsAdmin();

            var sitePage = homePage.GoToParameters_Sites();
            sitePage.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, site);
            string siteID = sitePage.CollectNewSiteID();

            // Act
            var deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            deliveryRoundPage.ResetFilter();

            var deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
            var deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRound(deliveryRoundName, site, DateUtils.Now, DateUtils.Now.AddDays(+31));

            var deliveryRoundDeliveriesPage = deliveryRoundGeneralInfoPage.ClickOnDeliveryTab();
            deliveryRoundDeliveriesPage.AddDelivery(deliveryNameForDR);
            deliveryRoundPage = deliveryRoundDeliveriesPage.BackToList();

            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRoundName);

            //Assert
            var firstDeliveryRound= deliveryRoundPage.GetFirstDeliveryRound();
            Assert.AreEqual(deliveryRoundName, firstDeliveryRound, "Le delivery round n'a pas été créé.");
        }

        // Dupliquer tournée de livraison
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_DERO_Duplicate_Delivery_Round()
        {
            // Prepare
            Random rnd = new Random();
            string site = TestContext.Properties["Site"].ToString();
            string deliveryRoundName = deliveryRoundNameToday + "-" + rnd.Next(1000, 9000).ToString();
            string deliveryRoundName2 = deliveryRoundNameToday + "-" + rnd.Next(1000, 9000).ToString();
            // Arrange
            HomePage homePage = LogInAsAdmin();
            var sitePage = homePage.GoToParameters_Sites();
            sitePage.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, site);
            string siteID = sitePage.CollectNewSiteID();
            // Act
            var deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            deliveryRoundPage.ResetFilter();
            var deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
            var deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRound(deliveryRoundName, site, DateUtils.Now, DateUtils.Now.AddDays(+31));

            var deliveryRoundDeliveriesPage = deliveryRoundGeneralInfoPage.ClickOnDeliveryTab();
            deliveryRoundDeliveriesPage.AddDelivery(deliveryNameForDR);
            deliveryRoundPage = deliveryRoundDeliveriesPage.BackToList();
            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRoundName);
            deliveryRoundDeliveriesPage = deliveryRoundPage.SelectFirstDeliveryRound();

            deliveryRoundDeliveriesPage.Duplicate(deliveryRoundName2, DateUtils.Now, DateUtils.Now.AddDays(+31));
            deliveryRoundPage = deliveryRoundDeliveriesPage.BackToList();
            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRoundName2);
            //Assert
            Assert.AreEqual(deliveryRoundName2, deliveryRoundPage.GetFirstDeliveryRound(), "Le delivery round n'a pas été dupliqué.");
        }

        // créer une nouvelle tournée de livraison sans remplir les informations obligatoires
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_DERO_Create_New_Delivery_Round_Without_Fill_Fields()
        {
            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            var deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
            deliveryRoundCreateModalpage.Save();

            // Assert 
            var errorMessage =deliveryRoundCreateModalpage.ErrorMessageNameRequired();
            Assert.IsTrue(errorMessage, "Aucun message d'erreur n'est apparu alors que les champs du delivery round n'ont pas été renseignés.");
        }

        // créer une nouvelle tournée de livraison déjà existante
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_DERO_Create_New_Delivery_Already_Existing()
        {
            // Prepare
            Random rnd = new Random();
            string site = TestContext.Properties["Site"].ToString();
            string deliveryRoundName = deliveryRoundNameToday + "-" + rnd.Next(1000, 9000).ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            var sitePage = homePage.GoToParameters_Sites();
            sitePage.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, site);
            string siteID = sitePage.CollectNewSiteID();

            // Act
            var deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            var deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
            var deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRound(deliveryRoundName, site, DateUtils.Now, DateUtils.Now.AddDays(+31));
            deliveryRoundPage = deliveryRoundGeneralInfoPage.BackToList();
            deliveryRoundPage.DeliveryRoundCreatePage();
            deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRound(deliveryRoundName, site, DateUtils.Now, DateUtils.Now.AddDays(+31));

            // Assert
            var errorMessage=deliveryRoundCreateModalpage.ErrorMessageNameAlreadyExists();
            Assert.IsTrue(errorMessage, "Aucun message d'erreur n'est apparu alors que le nom du delivery round est déjà attribué.");
        }


        // Effectuer des recherches via les filtres - Search
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_DERO_Filter_Search()
        {
            // Prepare
            Random rnd = new Random();
            string site = TestContext.Properties["Site"].ToString();

            string deliveryRoundName = deliveryRoundNameToday + "-" + rnd.Next(1000, 9000).ToString();

            string deliveryCustomer = TestContext.Properties["CustomerDeliveryRound"].ToString();
            string deliverySite = TestContext.Properties["Site"].ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            var sitePage = homePage.GoToParameters_Sites();
            sitePage.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, site);
            string siteID = sitePage.CollectNewSiteID();
            //Check
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.ResetFilter();
            deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName);

            if (deliveryPage.CheckTotalNumber() == 0)
            {
                var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryName, deliveryCustomer, deliverySite, true);
                var deliveryLoadingPage = deliveryCreateModalPage.Create();
                deliveryPage = deliveryLoadingPage.BackToList();
                deliveryPage.ResetFilter();
                deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryName);
            }

            // Assert
            Assert.AreEqual(deliveryPage.GetFirstDeliveryName(), deliveryName, "Le delivery n'a pas été créé.");
            // Act
            var deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            deliveryRoundPage.ResetFilter();

            var deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
            var deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRound(deliveryRoundName, site, DateUtils.Now, DateUtils.Now.AddDays(+31));

            var deliveryRoundDeliveriesPage = deliveryRoundGeneralInfoPage.ClickOnDeliveryTab();
            deliveryRoundDeliveriesPage.AddDelivery(deliveryName);
            deliveryRoundPage = deliveryRoundDeliveriesPage.BackToList();

            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRoundName);

            Assert.AreEqual(deliveryRoundName, deliveryRoundPage.GetFirstDeliveryRound(), string.Format(MessageErreur.FILTRE_ERRONE, "Search"));
        }

        // Effectuer des recherches via le filtre - SortBy Name
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_DERO_Filter_SortByName()
        {
            //Prepare
            Random rnd = new Random();
            string site = TestContext.Properties["Site"].ToString();
            string deliveryRoundName = deliveryRoundNameToday + "-" + rnd.Next(1000, 9000).ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            deliveryRoundPage.ResetFilter();
            var deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
            var deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRound(deliveryRoundName, site, DateUtils.Now, DateUtils.Now.AddDays(+31));

            var deliveryRoundDeliveriesPage = deliveryRoundGeneralInfoPage.ClickOnDeliveryTab();
            deliveryRoundDeliveriesPage.AddDelivery(deliveryNameForDR);
            deliveryRoundPage = deliveryRoundDeliveriesPage.BackToList();
            try
            {
                //Filter
                deliveryRoundPage.ResetFilter();
                deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.SortBy, "Name");
                bool isSortedByName = deliveryRoundPage.IsSortedByName();

                // Assert
                Assert.IsTrue(isSortedByName, String.Format(MessageErreur.FILTRE_ERRONE, "'Sort by name'"));
            }
            finally
            {
                // delete 
                deliveryRoundPage.ResetFilter();
                deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRoundName);
                deliveryRoundPage.DeleteDelivery();

            }
        }

        // Effectuer des recherches via le filtre - SortBy Site
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_DERO_Filter_SortBySite()
        {

            // Arrange
            var homePage = LogInAsAdmin();

            // Act
            var deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            deliveryRoundPage.ResetFilter();
            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.SortBy, "Site");
            bool isSortedBySite = deliveryRoundPage.IsSortedBySite();

            // Assert
            Assert.IsTrue(isSortedBySite, String.Format(MessageErreur.FILTRE_ERRONE, "'Sort by site'"));
        }

        // Effectuer des recherches via les filtres - Show all delivery round
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_DERO_Filter_ShowAll()
        {
            // Prepare
            Random rnd = new Random();
            string site = TestContext.Properties["Site"].ToString();

            string deliveryRoundName = deliveryRoundNameToday + "-" + rnd.Next(1000, 9000).ToString();


            // Arrange
            HomePage homePage = LogInAsAdmin();

            var sitePage = homePage.GoToParameters_Sites();
            sitePage.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, site);
            string siteID = sitePage.CollectNewSiteID();

            // Act
            var deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            deliveryRoundPage.ResetFilter();
            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.ShowAll, true);

            if (deliveryRoundPage.CheckTotalNumber() < 50)
            {
                var deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
                var deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRound(deliveryRoundName, siteID, DateUtils.Now, DateUtils.Now.AddDays(+31));

                var deliveryRoundDeliveriesPage = deliveryRoundGeneralInfoPage.ClickOnDeliveryTab();
                deliveryRoundDeliveriesPage.AddDelivery(deliveryName);
                deliveryRoundPage = deliveryRoundDeliveriesPage.BackToList();
                deliveryRoundPage.ResetFilter();
            }

            if (!deliveryRoundPage.isPageSizeEqualsTo100())
            {
                deliveryRoundPage.PageSize("8");
                deliveryRoundPage.PageSize("100");
            }

            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.ShowOnlyActive, true);
            var nbrActive = deliveryRoundPage.CheckTotalNumber();

            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.ShowOnlyInactive, true);
            var nbrInactive = deliveryRoundPage.CheckTotalNumber();

            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.ShowAll, true);
            var realNbr = deliveryRoundPage.CheckTotalNumber();

            //Assert
            Assert.AreEqual(nbrActive + nbrInactive, realNbr, String.Format(MessageErreur.FILTRE_ERRONE, "Show all"));
        }

        // Effectuer des recherches via les filtres - Show inactive delivery round
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_DERO_Filter_ShowInactive()
        {
            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            deliveryRoundPage.ResetFilter();
            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.ShowOnlyInactive, true);                           
            if (!deliveryRoundPage.isPageSizeEqualsTo100())
            {
                deliveryRoundPage.PageSize("8");
                deliveryRoundPage.PageSize("100");
            }

            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.ShowAll, true);
            var numberShowAll = deliveryRoundPage.CheckTotalNumber(); 
            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.ShowOnlyActive, true);
            var numberShowActive = deliveryRoundPage.CheckTotalNumber();

            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.ShowOnlyInactive, true);
            var numberInactive= deliveryRoundPage.CheckTotalNumber();

            //Assert
            Assert.IsFalse(deliveryRoundPage.CheckStatus(false), String.Format(MessageErreur.FILTRE_ERRONE, "'Show only inactive'"));
            Assert.AreEqual(numberShowAll- numberShowActive, numberInactive, "le filtre ne fonctionne pas correctement");
        }

        // Effectuer des recherches via les filtres - Show all delivery round
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_DERO_Filter_ShowActive()
        {
            // Prepare
            Random rnd = new Random();
            string site = TestContext.Properties["Site"].ToString();

            string deliveryRoundName = deliveryRoundNameToday + "-" + rnd.Next(1000, 9000).ToString();


            // Arrange
            HomePage homePage = LogInAsAdmin();
            

            var sitePage = homePage.GoToParameters_Sites();
            sitePage.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, site);
            string siteID = sitePage.CollectNewSiteID();

            // Act
            var deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            deliveryRoundPage.ResetFilter();
            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.ShowOnlyActive, true);

            if (deliveryRoundPage.CheckTotalNumber() < 20)
            {
                var deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
                var deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRound(deliveryRoundName, siteID, DateUtils.Now, DateUtils.Now.AddDays(+31));

                var deliveryRoundDeliveriesPage = deliveryRoundGeneralInfoPage.ClickOnDeliveryTab();
                deliveryRoundDeliveriesPage.AddDelivery(deliveryName);
                deliveryRoundPage = deliveryRoundDeliveriesPage.BackToList();

                deliveryRoundPage.ResetFilter();
                deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.ShowOnlyActive, true);
            }

            if (!deliveryRoundPage.isPageSizeEqualsTo100())
            {
                deliveryRoundPage.PageSize("8");
                deliveryRoundPage.PageSize("100");
            }

            //Assert
            Assert.IsTrue(deliveryRoundPage.CheckStatus(true), String.Format(MessageErreur.FILTRE_ERRONE, "'Show only active'"));
        }


        [Timeout(_timeout)]
		[TestMethod]
        public void CU_DERO_Filter_Site()
        {
            // Prepare
            Random rnd = new Random();
            string site = TestContext.Properties["Site"].ToString();
            string deliveryRoundName = deliveryRoundNameToday + "-" + rnd.Next(1000, 9000).ToString();

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var sitePage = homePage.GoToParameters_Sites();
            sitePage.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, site);
            string siteID = sitePage.CollectNewSiteID();

            // Act
            var deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            deliveryRoundPage.ResetFilter();
            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Site, site);

            if (deliveryRoundPage.CheckTotalNumber() < 20)
            {
                var deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
                var deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRound(deliveryRoundName, siteID, DateUtils.Now, DateUtils.Now.AddDays(+31));

                var deliveryRoundDeliveriesPage = deliveryRoundGeneralInfoPage.ClickOnDeliveryTab();
                deliveryRoundDeliveriesPage.AddDelivery(deliveryNameForDR);
                deliveryRoundPage = deliveryRoundDeliveriesPage.BackToList();

                deliveryRoundPage.ResetFilter();
                deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Site, site);
            }

            if (!deliveryRoundPage.isPageSizeEqualsTo100())
            {
                deliveryRoundPage.PageSize("8");
                deliveryRoundPage.PageSize("100");
            }
           var verifySite= deliveryRoundPage.VerifySite(site);
            Assert.IsTrue(verifySite, MessageErreur.FILTRE_ERRONE, "Sites");
        }

        // Effectuer des recherches via les filtres - Customers
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_DERO_Filter_Customers()
        {
            // Prepare
            Random rnd = new Random();
            string site = TestContext.Properties["Site"].ToString();
            string deliveryCustomer = TestContext.Properties["CustomerDeliveryRound"].ToString();
            string deliveryRoundName = deliveryRoundNameToday + "-" + rnd.Next(1000, 9000).ToString();


            // Arrange
            HomePage homePage = LogInAsAdmin();

            //var sitePage = homePage.GoToParameters_Sites();
            //sitePage.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, site);
            //string siteID = sitePage.CollectNewSiteID();

            // Act
            var deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            deliveryRoundPage.ResetFilter();
            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Customers, deliveryCustomer);
            if (deliveryRoundPage.CheckTotalNumber() < 20)
            {
                var deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
                var deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRound(deliveryRoundName, site, DateUtils.Now, DateUtils.Now.AddDays(+31));

                var deliveryRoundDeliveriesPage = deliveryRoundGeneralInfoPage.ClickOnDeliveryTab();
                deliveryRoundDeliveriesPage.AddDelivery(deliveryNameForDR);
                deliveryRoundPage = deliveryRoundDeliveriesPage.BackToList();

                deliveryRoundPage.ResetFilter();
                deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Customers, deliveryCustomer);
            }
            if (!deliveryRoundPage.isPageSizeEqualsTo100())
            {
                //deliveryRoundPage.PageSize("8");
                deliveryRoundPage.PageSize("100");
            }

            deliveryRoundPage.UnfoldAll();
            //Assert 
            bool isGoodCustomer = deliveryRoundPage.VerifyCustomer(deliveryCustomer);
            Assert.IsTrue(isGoodCustomer, MessageErreur.FILTRE_ERRONE, "Customers");
        }

        // Modifier les informations générales
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_DERO_Change_General_Information()
        {
            // Prepare
            Random rnd = new Random();
            string site = TestContext.Properties["Site"].ToString();
            string siteUpdated = TestContext.Properties["SiteLP"].ToString();
            string deliveryCustomer = TestContext.Properties["CustomerDeliveryRound"].ToString();
            string deliverySite = TestContext.Properties["Site"].ToString();
            string deliveryRoundName = deliveryRoundNameToday + "-" + rnd.Next(1000, 9000).ToString();
            string deliveryRoundNameUpdated = deliveryRoundNameToday + "-" + rnd.Next(1000, 9000).ToString();
            DateTime initStartDate = DateUtils.Now;
            DateTime initEndDate = DateUtils.Now.AddDays(+31);
            DateTime finalStartDate = DateUtils.Now.AddDays(+2);
            DateTime finalEndDate = DateUtils.Now.AddDays(+20);

            // Arrange
            HomePage homePage = LogInAsAdmin();
            // Act
            var dateFormat = homePage.GetDateFormatPickerValue();
            var sitePage = homePage.GoToParameters_Sites();
            sitePage.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, site);
            string siteID = sitePage.CollectNewSiteID();
            WebDriver.Navigate().Refresh();
            sitePage.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, siteUpdated);
            string siteUpdatedID = sitePage.CollectNewSiteID();
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.ResetFilter();
            deliveryPage.Filter(DeliveryPage.FilterType.Search, deliveryNameForDR);
            if (deliveryPage.CheckTotalNumber() == 0)
            {
                var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryNameForDR, deliveryCustomer, deliverySite, true);
                var deliveryLoadingPage = deliveryCreateModalPage.Create();
                deliveryPage = deliveryLoadingPage.BackToList();
              
            }
            var deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            deliveryRoundPage.ResetFilter();
            var deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
            deliveryRoundPage.WaitPageLoading();
            var deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRound(deliveryRoundName, site, initStartDate, initEndDate);
            var deliveryRoundDeliveriesPage = deliveryRoundGeneralInfoPage.ClickOnDeliveryTab();
            deliveryRoundDeliveriesPage.AddDelivery(deliveryNameForDR);
            var deliveryRoundGeneralInformationPage = deliveryRoundDeliveriesPage.ClickOnGeneralInfoTab();
            deliveryRoundDeliveriesPage.WaitPageLoading();
            deliveryRoundGeneralInformationPage.UpdateGeneralInformation(deliveryRoundNameUpdated, siteUpdatedID, finalStartDate, finalEndDate);
            deliveryRoundPage = deliveryRoundGeneralInformationPage.BackToList();
            deliveryRoundPage.ResetFilter();
            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRoundName);
            Assert.AreEqual(0, deliveryRoundPage.CheckTotalNumber(), "Le delivery existe toujours alors qu'il a été renommé.");
            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRoundNameUpdated);
            deliveryRoundDeliveriesPage = deliveryRoundPage.SelectFirstDeliveryRound();
            deliveryRoundGeneralInformationPage = deliveryRoundDeliveriesPage.ClickOnGeneralInfoTab();
            var newName = deliveryRoundGeneralInformationPage.GetDeliveryRoundName();
            var newSite = deliveryRoundGeneralInformationPage.GetDeliveryRoundSite();
            var newStartDate = deliveryRoundGeneralInformationPage.GetDeliveryRoundStartDate(dateFormat);
            var newEndDate = deliveryRoundGeneralInformationPage.GetDeliveryRoundEndDate(dateFormat);

            // Assert
            Assert.AreEqual(deliveryRoundNameUpdated, newName, "Le nom du delivery round n'a pas été mis à jour.");
            Assert.AreEqual(siteUpdatedID, newSite, "Le site du delivery round n'a pas été mis à jour.");
            Assert.AreEqual(finalStartDate.Date, newStartDate.Date, "La start du delivery round n'a pas été mise à jour.");
            Assert.AreEqual(finalEndDate.Date, newEndDate.Date, "La end du delivery round n'a pas été mise à jour.");
        }


        // Modifier le nom du delivery round pour mettre un déjà existant
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_DERO_Change_Delivery_Round_Name_Already_Existing()
        {
            // Prepare
            Random rnd = new Random();
            string site = TestContext.Properties["Site"].ToString();
            string deliveryRoundName = deliveryRoundNameToday + "-" + rnd.Next(1000, 9000).ToString();
            string deliveryRoundName2 = deliveryRoundNameToday + "-" + rnd.Next(1000, 9000).ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Création de 2 delivery round
            var deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            deliveryRoundPage.ResetFilter();

            var deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
            var deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRound(deliveryRoundName, site, DateUtils.Now, DateUtils.Now.AddDays(+31));
            var deliveryRoundDeliveriesPage = deliveryRoundGeneralInfoPage.ClickOnDeliveryTab();
            deliveryRoundDeliveriesPage.AddDelivery(deliveryNameForDR);
            deliveryRoundPage = deliveryRoundDeliveriesPage.BackToList();

            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRoundName);
            Assert.AreEqual(1, deliveryRoundPage.CheckTotalNumber(), "Le deliveryRound n'a pas été créé.");

            deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
            deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRound(deliveryRoundName2, site, DateUtils.Now, DateUtils.Now.AddDays(+31));
            deliveryRoundDeliveriesPage = deliveryRoundGeneralInfoPage.ClickOnDeliveryTab();
            deliveryRoundDeliveriesPage.AddDelivery(deliveryNameForDR);

            deliveryRoundGeneralInfoPage = deliveryRoundDeliveriesPage.ClickOnGeneralInfoTab();
            deliveryRoundGeneralInfoPage.UpdateDeliveryRoundName(deliveryRoundName);
            deliveryRoundGeneralInfoPage.WaitPageLoading();
            //Assert 
            var errorMessage = deliveryRoundGeneralInfoPage.GetErrorMessageAlreadyExisting();
            Assert.IsTrue(errorMessage, "Aucun message d'erreur n'est apparu alors qu'on tente de renommer un delivery round avec un nom existant.");
        }

        // Ajouter un delivery
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_DERO_Create_Delivery()
        {
            // Prepare
            Random rnd = new Random();
            string site = TestContext.Properties["Site"].ToString();
            string deliveryRoundName = deliveryRoundNameToday + "-" + rnd.Next(1000, 9000).ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            deliveryRoundPage.ResetFilter();
            var deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
            var deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRound(deliveryRoundName, site, DateUtils.Now, DateUtils.Now.AddDays(+31));
            var deliveryRoundDeliveriesPage = deliveryRoundGeneralInfoPage.ClickOnDeliveryTab();
            deliveryRoundDeliveriesPage.AddDelivery(deliveryNameForDR);

            // Assert
            var isDeliveryVisible=deliveryRoundDeliveriesPage.IsDeliveryVisible();
            Assert.IsTrue(isDeliveryVisible, "Le delivery n'a pas été ajouté au delivery round.");
        }

        // Consulter un delivery
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_DERO_Check_Delivery()
        {
            // Prepare
            Random rnd = new Random();
            string site = TestContext.Properties["Site"].ToString();
            string deliveryRoundName = deliveryRoundNameToday + "-" + rnd.Next(1000, 9000).ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            deliveryRoundPage.ResetFilter();

            var deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
            var deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRound(deliveryRoundName, site, DateUtils.Now, DateUtils.Now.AddDays(+31));
            var deliveryRoundDeliveriesPage = deliveryRoundGeneralInfoPage.ClickOnDeliveryTab();
            deliveryRoundDeliveriesPage.AddDelivery(deliveryNameForDR);
            var deliveryLoadingPage = deliveryRoundDeliveriesPage.EditDelivery();
            string deliveryNameReturn = deliveryLoadingPage.GetDeliveryName();
            var totalTab = deliveryRoundDeliveriesPage.GetTotalTabs();

            // Assert
            Assert.IsTrue(totalTab,"le nouveau onglet n'est pas ouvert" ); 
            Assert.AreEqual(deliveryNameForDR.ToLower().Trim(), deliveryNameReturn.ToLower().Trim(), "La consultation du delivery a échoué.");
        }

        // Supprimer un delivery
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_DERO_Delete_Delivery()
        {
            // Prepare
            Random rnd = new Random();
            string site = TestContext.Properties["Site"].ToString();

            string deliveryRoundName = deliveryRoundNameToday + "-" + rnd.Next(1000, 9000).ToString();


            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            deliveryRoundPage.ResetFilter();
            var deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
            var deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRound(deliveryRoundName, site, DateUtils.Now, DateUtils.Now.AddDays(+31));
            var deliveryRoundDeliveriesPage = deliveryRoundGeneralInfoPage.ClickOnDeliveryTab();
            deliveryRoundDeliveriesPage.AddDelivery(deliveryNameForDR);
            //Assert
            Assert.IsTrue(deliveryRoundDeliveriesPage.IsDeliveryVisible(), "Le delivery n'a pas été ajouté au delivery round.");
            deliveryRoundDeliveriesPage.DeleteDelivery();
            var isDeliveryVisible= deliveryRoundDeliveriesPage.IsDeliveryVisible();
            Assert.IsFalse(isDeliveryVisible, "Le delivery n'a pas été supprimé du delivery round.");
        }

        // exporter un delivery 
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_DERO_Export_Delivery_Round_NewVersion()
        {
            // Prepare
            Random rnd = new Random();
            string site = TestContext.Properties["Site"].ToString();

            string deliveryRoundName = deliveryRoundNameToday + "-" + rnd.Next(1000, 9000).ToString();


            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            bool newVersionPrint = true;

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            deliveryRoundPage.ResetFilter();

            deliveryRoundPage.ClearDownloads();

            var deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
            var deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRound(deliveryRoundName, site, DateUtils.Now, DateUtils.Now.AddDays(+31));

            var deliveryRoundDeliveriesPage = deliveryRoundGeneralInfoPage.ClickOnDeliveryTab();
            deliveryRoundDeliveriesPage.AddDelivery(deliveryNameForDR);
            deliveryRoundPage = deliveryRoundDeliveriesPage.BackToList();

            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.DateFrom, DateUtils.Now.AddDays(-1));
            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.DateTo, DateUtils.Now.AddMonths(+2));

            deliveryRoundPage.Export(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = deliveryRoundPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber(DELIVERY_ROUND_EXCEL_SHEET_NAME, filePath);

            //Assert
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);
            DeleteAllFileDownload();
        }

        // Exporter les données des "delivery rounds" en utilisant le filtre avec search 
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_DERO_Export_Filter_Search_Delivery_Round_NewVersion()
        {
            // Prepare
            Random rnd = new Random();
            string site = TestContext.Properties["Site"].ToString();
            string deliveryRoundName = deliveryRoundNameToday + "-" + rnd.Next(1000, 9000).ToString();
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            bool newVersionPrint = true;
            HomePage homePage = LogInAsAdmin();
            var sitePage = homePage.GoToParameters_Sites();
            sitePage.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, site);
            string siteID = sitePage.CollectNewSiteID();
            var deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            deliveryRoundPage.ResetFilter();
            deliveryRoundPage.ClearDownloads();
            var deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
            var deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRound(deliveryRoundName, site, DateUtils.Now, DateUtils.Now.AddDays(+31));

            var deliveryRoundDeliveriesPage = deliveryRoundGeneralInfoPage.ClickOnDeliveryTab();
            deliveryRoundDeliveriesPage.AddDelivery(deliveryNameForDR);
            deliveryRoundPage = deliveryRoundDeliveriesPage.BackToList();
            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRoundName);
            deliveryRoundPage.Export(newVersionPrint);
            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();
            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = deliveryRoundPage.GetExportExcelFile(taskFiles, deliveryRoundName);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");
            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);
            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber(DELIVERY_ROUND_EXCEL_SHEET_NAME, filePath);
            bool result = OpenXmlExcel.ReadAllDataInColumn("Code/Name", DELIVERY_ROUND_EXCEL_SHEET_NAME, filePath, deliveryRoundName);
            //Assert
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.IsTrue(result, MessageErreur.EXCEL_DONNEES_KO);
        }

        // Exporter les données des "delivery rounds" en utilisant le filtre avec Site
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_DERO_Export_Filter_Site_Delivery_Round_NewVersion()
        {
            // Prepare
            Random rnd = new Random();
            string site = TestContext.Properties["Site"].ToString();

            string deliveryRoundName = deliveryRoundNameToday + "-" + rnd.Next(1000, 9000).ToString();


            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();
            bool newVersionPrint = true;

            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            var sitePage = homePage.GoToParameters_Sites();
            sitePage.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, site);
            string siteID = sitePage.CollectNewSiteID();

            // Act
            var deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            deliveryRoundPage.ResetFilter();

            deliveryRoundPage.ClearDownloads();
            deliveryRoundPage.WaitPageLoading();

            var deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
            var deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRound(deliveryRoundName, site, DateUtils.Now, DateUtils.Now.AddDays(+31));

            var deliveryRoundDeliveriesPage = deliveryRoundGeneralInfoPage.ClickOnDeliveryTab();
            deliveryRoundDeliveriesPage.AddDelivery(deliveryNameForDR);
            deliveryRoundPage = deliveryRoundDeliveriesPage.BackToList();

            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Site, site);
            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.DateFrom, DateUtils.Now.AddDays(-1));
            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.DateTo, DateUtils.Now.AddMonths(+2));

            deliveryRoundPage.Export(newVersionPrint);

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = deliveryRoundPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);

            // Exploitation du fichier Excel
            int resultNumber = OpenXmlExcel.GetExportResultNumber(DELIVERY_ROUND_EXCEL_SHEET_NAME, filePath);
            var listResults = OpenXmlExcel.GetValuesInList("Code/Name", DELIVERY_ROUND_EXCEL_SHEET_NAME, filePath);
            listResults.RemoveAll(x => x == null);

            var listSites = OpenXmlExcel.GetValuesOnListDeliveryRound("Site", DELIVERY_ROUND_EXCEL_SHEET_NAME, filePath, site);

            //Assert
            Assert.AreNotEqual(resultNumber, 0, MessageErreur.EXCEL_PAS_DE_DONNEES);
            Assert.AreEqual(listResults.Count, listSites.Count, MessageErreur.EXCEL_DONNEES_KO);
        }

        // ---------------------------- Methodes --------------------------
        
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_DERO_Access()
        {

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);

            var deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            bool isAccessOK = deliveryRoundPage.AccessPage();

            Assert.IsTrue(isAccessOK, "Page inaccessible");
             
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_DERO_NewDR()
        {
            // Prepare
            Random rnd = new Random();
            string site = TestContext.Properties["Site"].ToString();
            string deliveryRoundName = deliveryRoundNameToday + "-" + rnd.Next(1000, 9000).ToString();

            // Arrange
            HomePage homePage = LogInAsAdmin();

            // Act
            var deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
            deliveryRoundPage.ResetFilter();
            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRoundName);
            int totalCountBefore = deliveryRoundPage.CheckTotalNumber();
            try
            {
                DeliveryRoundCreateModalpage deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
                DeliveryRoundGeneralInformationPage deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRound(deliveryRoundName, site, DateUtils.Now, DateUtils.Now.AddDays(+31));

                deliveryRoundGeneralInfoPage.BackToList();
                deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRoundName);

                //Assert
                int totalCountAfter = deliveryRoundPage.CheckTotalNumber();
                Assert.IsTrue(totalCountAfter == totalCountBefore + 1, "Le delivery round n'a pas été créé.");
            }
            finally
            {
                deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRoundName);
                deliveryRoundPage.DeleteFirstDeliveryRound();
            }

        }

        ////////////////////////////////////////////// Methods ///////////////////////////////////////////
        private int InsertDeliveryRound(string deliveryRoundName, DateTime dateFrom, DateTime dateTo, string site)
        {
            string query = @"
            DECLARE @siteId INT;
            SELECT TOP 1 @siteId = Id FROM sites WHERE Name LIKE @site;
            INSERT INTO deliveryRounds (Name, IsActive, SiteId, Monday, Tuesday, Wednesday, Thursday, Friday, Saturday, Sunday, StartDate, EndDate) 
            VALUES (@deliveryRoundName, 1, @siteId, 0, 0, 0, 0, 0, 0, 0, @dateFrom, @dateTo);
            SELECT SCOPE_IDENTITY();";
 
            return ExecuteAndGetInt(
                query,
                new KeyValuePair<string, object>("deliveryRoundName", deliveryRoundName),
                new KeyValuePair<string, object>("dateFrom", dateFrom),
                new KeyValuePair<string, object>("dateTo", dateTo),
                new KeyValuePair<string, object>("site", site)
            );
        }

        private int InsertDeliveryForDeliveryRound(string deliveryNameForDR, DateTime date, string deliverySite, string deliveryCustomer)
        {
            Random random = new Random();
            int cst = random.Next(1, 1000);
            string query = @"
                DECLARE @siteId INT;
                SELECT TOP 1 @siteId = Id FROM sites WHERE Name LIKE @deliverySite;
                DECLARE @customerId INT;
                SELECT TOP 1 @customerId = Id FROM customers WHERE Name LIKE @deliveryCustomer;
                INSERT INTO FlightDeliveries 
                    (Name, IsActive, SiteId, CustomerId, DeliveryTime, HoursBeforeAccess, AllowedModificationsPercent, HoursBeforeModifications, Method, CustomerPortalBlockAccessType) 
                VALUES 
                    (@deliveryNameForDR, 1, @siteId, @customerId, '00:00:00.0000000', 1, 1, 1, 1, 1);

                SELECT SCOPE_IDENTITY();";

            return ExecuteAndGetInt(
                query,
                new KeyValuePair<string, object>("deliveryNameForDR", deliveryNameForDR),
                new KeyValuePair<string, object>("deliveryCustomer", deliveryCustomer),
                new KeyValuePair<string, object>("deliverySite", deliverySite));
        }
        private void DeleteDeliveryRound(int deliveryRoundId)
        {
            string query = @"
            DELETE FROM DeliveryRoundToFlightDeliveries 
                    WHERE DeliveryRoundId = @deliveryRoundId;

            DELETE FROM deliveryRounds 
                    WHERE Id = @deliveryRoundId;";

            ExecuteNonQuery(query, new KeyValuePair<string, object>("deliveryRoundId", deliveryRoundId));
        }

        private void DeleteMultipleDeliveryRounds(List<int> deliveryRoundIds)
        {
            foreach (var deliveryRoundId in deliveryRoundIds)
            {
                DeleteDeliveryRound(deliveryRoundId);
            }
        }

        private void DeleteDelivery(int deliveryId)
        {
            string query = @"
            DELETE FROM FlightDeliveries 
                    WHERE Id = @deliveryId;";

            TestContext.Properties["deliveryId"] = ExecuteAndGetInt(
                query,
                new KeyValuePair<string, object>("deliveryId", deliveryId));
        }
        private void VerifyDeliveryRound(DeliveryRoundPage deliveryRoundPage, string deliveryRoundName, bool shouldExist, string errorMessage)
        {
            deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRoundName);
            if (shouldExist)
            {
                Assert.IsTrue(deliveryRoundPage.VerifyDRNamesExist(deliveryRoundName), errorMessage);
            }
            else
            {
                Assert.IsFalse(deliveryRoundPage.VerifyDRNamesExist(deliveryRoundName), errorMessage);
            }
        }

        public void DeliveriesFrDR_TestCleanUp()
        {
            List<int> Deliveries = new List<int>();

            for (int i = 1; i <= 2; i++)
            {
                if (TestContext.Properties.Contains($"DeliveryFrDRId{i}"))
                {
                    Deliveries.Add((int)TestContext.Properties[$"DeliveryFrDRId{i}"]);
                }
            }
            DeleteDeliveries(Deliveries);
        }
        private void DeleteDeliveries(List<int> deliveries)
        {
            foreach (var deliveryId in deliveries)
            {
                DeleteDelivery(deliveryId);
            }
        }

        private int? DeliveryExist(int id)
        {
            string query = @"
             select Id from FlightDeliveries where Id = @id

              SELECT SCOPE_IDENTITY();
            "
            ;
            int? result = ExecuteAndGetInt(
            query,
            new KeyValuePair<string, object>("id", id)
            );
            return result == 0 ? (int?)null : result;
        }
        [Timeout(_timeout)]
        [TestMethod]
        public void CU_DERO_GeneralInformation()
        {
            // Prepare
            Random rnd = new Random();
            string site = TestContext.Properties["Site"].ToString();
            string siteUpdated = TestContext.Properties["SiteLP"].ToString();
            string deliveryCustomer = "A.F.B. LANZAROTE";
            string deliverySite = TestContext.Properties["Site"].ToString();
            string deliveryRoundName = deliveryRoundNameToday + "-" + rnd.Next(1000, 9000).ToString();
            string deliveryRoundNameUpdated = deliveryRoundNameToday + "-" + rnd.Next(1000, 9000).ToString();
            deliveryNameForDR = "deliveryForDeliveryRound - " + DateUtils.Now.ToString("dd / MM / yyyy") + "-" + random.Next(10, 999);
            DateTime initStartDate = DateUtils.Now;
            DateTime initEndDate = DateUtils.Now.AddDays(+31);
            DateTime finalStartDate = DateUtils.Now.AddDays(+2);
            DateTime finalEndDate = DateUtils.Now.AddDays(+20);

            // Arrange
            HomePage homePage = LogInAsAdmin();
            // Act
            var dateFormat = homePage.GetDateFormatPickerValue();
            var sitePage = homePage.GoToParameters_Sites();
            sitePage.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, site);
            string siteID = sitePage.CollectNewSiteID();
            WebDriver.Navigate().Refresh();
            sitePage.Filter(PageObjects.Parameters.Sites.ParametersSites.FilterType.SearchSite, siteUpdated);
            string siteUpdatedID = sitePage.CollectNewSiteID();
            var deliveryPage = homePage.GoToCustomers_DeliveryPage();
            deliveryPage.ResetFilter();
            try
            {
                var deliveryCreateModalPage = deliveryPage.DeliveryCreatePage();
                deliveryCreateModalPage.FillFields_CreateDeliveryModalPage(deliveryNameForDR, deliveryCustomer, deliverySite, true);
                var deliveryLoadingPage = deliveryCreateModalPage.Create();
                deliveryPage = deliveryLoadingPage.BackToList();

                var deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
                deliveryRoundPage.ResetFilter();
                var deliveryRoundCreateModalpage = deliveryRoundPage.DeliveryRoundCreatePage();
                deliveryRoundPage.WaitPageLoading();
                var deliveryRoundGeneralInfoPage = deliveryRoundCreateModalpage.FillField_CreateNewDeliveryRound(deliveryRoundName, site, initStartDate, initEndDate);
                var deliveryRoundDeliveriesPage = deliveryRoundGeneralInfoPage.ClickOnDeliveryTab();
                var deliveryRoundGeneralInformationPage = deliveryRoundDeliveriesPage.ClickOnGeneralInfoTab();
                deliveryRoundDeliveriesPage.WaitPageLoading();
                deliveryRoundGeneralInformationPage.UpdateGeneralInformation(deliveryRoundNameUpdated, siteUpdatedID, finalStartDate, finalEndDate);
                deliveryRoundPage = deliveryRoundGeneralInformationPage.BackToList();
                deliveryRoundPage.ResetFilter();
                deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRoundName);
                Assert.AreEqual(0, deliveryRoundPage.CheckTotalNumber(), "Le delivery existe toujours alors qu'il a été renommé.");
                deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRoundNameUpdated);
                deliveryRoundDeliveriesPage = deliveryRoundPage.SelectFirstDeliveryRound();
                deliveryRoundGeneralInformationPage = deliveryRoundDeliveriesPage.ClickOnGeneralInfoTab();
                var newName = deliveryRoundGeneralInformationPage.GetDeliveryRoundName();
                var newSite = deliveryRoundGeneralInformationPage.GetDeliveryRoundSite();
                var newStartDate = deliveryRoundGeneralInformationPage.GetDeliveryRoundStartDate(dateFormat);
                var newEndDate = deliveryRoundGeneralInformationPage.GetDeliveryRoundEndDate(dateFormat);

                // Assert
                Assert.AreEqual(deliveryRoundNameUpdated, newName, "Le nom du delivery round n'a pas été mis à jour.");
                Assert.AreEqual(siteUpdatedID, newSite, "Le site du delivery round n'a pas été mis à jour.");
                Assert.AreEqual(finalStartDate.Date, newStartDate.Date, "La start du delivery round n'a pas été mise à jour.");
                Assert.AreEqual(finalEndDate.Date, newEndDate.Date, "La end du delivery round n'a pas été mise à jour.");
            }
            finally
            {
                var deliveryRoundPage = homePage.GoToCustomers_DeliveryRoundPage();
                deliveryRoundPage.Filter(DeliveryRoundPage.FilterType.Search, deliveryRoundNameUpdated);
                deliveryRoundPage.DeleteDeliveryRound();
            }
        }
    }
}