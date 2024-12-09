using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.FoodPackets;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using System.Diagnostics;
using System.IO;
using System.Net.Sockets;
using System.Security.Policy;
using System.Web.UI;

namespace Newrest.Winrest.FunctionalTests.Customers
{
    [TestClass]
    public class FoodPacketsTest : TestBase
    {
        private const int _timeout = 600000;


        [Timeout(_timeout)]
		[TestMethod]
        public void CU_FOPA_Filter_Search()
        {
            //prepare
            var searchInputPacket = "ATLANTIC";
            var searchInputNotExist = "TEST";
            var customer = "UNITED AIRLINES, INC. SUC EN ESPAÑA";
            var serviceName = "ServiceForTest";
            var site = TestContext.Properties["Site"].ToString();
            var packetOne = "ATLANTIC BEVERAGES UPB";

            //arrange
            HomePage homePage = LogInAsAdmin();

            try
            {
                var servicePage = homePage.GoToCustomers_ServicePage();
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
                var servicepricePage = servicePage.SelectFirstRowInCustmerService();
                var foodpacketpage = servicepricePage.GoToFoodPacketsTab();
                var number = foodpacketpage.GetNumberOfPacketFoodsInService();
                if (number > 0)
                {
                    foodpacketpage.DeleteRows(number);
                }
                ServiceCreateFoodPacketModal createfoodpackets = foodpacketpage.CreateFoodPackets();
                createfoodpackets.FillServiceFoodPacket(site, packetOne, customer);

                //apply filter on search
                FoodPacketPage foodPacketPages = homePage.GoToCustomers_FoodPacketsPage();
                foodPacketPages.ResetFilters();
                foodPacketPages.Filter(FoodPacketPage.FilterType.Search, searchInputPacket);
                foodPacketPages.Filter(FoodPacketPage.FilterType.ChildrenFoodPackets, true);
                foodPacketPages.Filter(FoodPacketPage.FilterType.Site, site);
                foodPacketPages.Filter(FoodPacketPage.FilterType.Customers, customer);
                var filteredResults = foodPacketPages.GetNumberOfFoodPacketHeader();

                //Assert
                Assert.IsTrue(int.Parse(filteredResults) > 0, "search filter ne fonctionne pas");
                //Assert
                foodPacketPages.Filter(FoodPacketPage.FilterType.Search, searchInputNotExist);
                filteredResults = foodPacketPages.GetNumberOfFoodPacketHeader();
                Assert.IsTrue(int.Parse(filteredResults) == 0, "search filter ne fonctionne pas");

            }
            finally
            {
                var servicePage = homePage.GoToCustomers_ServicePage();
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
                var servicepricePage = servicePage.SelectFirstRowInCustmerService();
                var foodpacketpage = servicepricePage.GoToFoodPacketsTab();
                var number = foodpacketpage.GetNumberOfPacketFoodsInService();
                if (number > 0)
                {
                    foodpacketpage.DeleteRows(number);
                }
            }
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_FOPA_Filter_Customer()
        {
            //prepare
            var customer = "UNITED AIRLINES, INC. SUC EN ESPAÑA";
            var serviceName = "ServiceForTest";
            var site = TestContext.Properties["Site"].ToString();
            var packetOne = "ATLANTIC BEVERAGES UPB";

            //arrange
            HomePage homePage = LogInAsAdmin();
          
            try
            {
                var servicePage = homePage.GoToCustomers_ServicePage();
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
                var servicepricePage = servicePage.SelectFirstRowInCustmerService();
                var foodpacketpage = servicepricePage.GoToFoodPacketsTab();
                var number = foodpacketpage.GetNumberOfPacketFoodsInService();
                if (number > 0)
                {
                    foodpacketpage.DeleteRows(number);
                }
                ServiceCreateFoodPacketModal createfoodpackets = foodpacketpage.CreateFoodPackets();
                createfoodpackets.FillServiceFoodPacket(site, packetOne, customer);
                //apply filter on customer
                FoodPacketPage foodPacketPages = homePage.GoToCustomers_FoodPacketsPage();
                foodPacketPages.ResetFilters();
                foodPacketPages.Filter(FoodPacketPage.FilterType.ChildrenFoodPackets, true);
                foodPacketPages.Filter(FoodPacketPage.FilterType.Customers, customer);
                foodPacketPages.Filter(FoodPacketPage.FilterType.Site, site);
                var filteredResults = foodPacketPages.GetNumberOfFoodPacketHeader();
                //Assert
                Assert.IsTrue(int.Parse(filteredResults) > 0, "Customer filter ne fonctionne pas");
            }
            finally
            {
                var servicePage = homePage.GoToCustomers_ServicePage();
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
                var servicepricePage = servicePage.SelectFirstRowInCustmerService();
                var foodpacketpage = servicepricePage.GoToFoodPacketsTab();
                var number = foodpacketpage.GetNumberOfPacketFoodsInService();
                if (number > 0)
                {
                    foodpacketpage.DeleteRows(number);
                }
            }
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_FOPA_Filter_Site()
        {
            //prepare
            var customer = "UNITED AIRLINES, INC. SUC EN ESPAÑA";
            var serviceName = "ServiceForTest";
            var site = "ACE";
            var packetOne = "ATLANTIC BEVERAGES UPB";

            //arrange
            HomePage homePage =  LogInAsAdmin();

            try
            {
                var servicePage = homePage.GoToCustomers_ServicePage();
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
                var servicepricePage = servicePage.SelectFirstRowInCustmerService();
                var foodpacketpage = servicepricePage.GoToFoodPacketsTab();
                var number = foodpacketpage.GetNumberOfPacketFoodsInService();
                if (number > 0)
                {
                    foodpacketpage.DeleteRows(number);
                }
                ServiceCreateFoodPacketModal createfoodpackets = foodpacketpage.CreateFoodPackets();
                createfoodpackets.FillServiceFoodPacket(site, packetOne, customer);

                //apply filter on site
                FoodPacketPage foodPacketPages = homePage.GoToCustomers_FoodPacketsPage();
                foodPacketPages.ResetFilters();
                foodPacketPages.Filter(FoodPacketPage.FilterType.ChildrenFoodPackets, true);
                foodPacketPages.Filter(FoodPacketPage.FilterType.Site, site);
                foodPacketPages.Filter(FoodPacketPage.FilterType.Customers, customer);
                var filteredResults = foodPacketPages.GetNumberOfFoodPacketHeader();
                //Assert
                Assert.IsTrue(int.Parse(filteredResults) > 0, "Site filter ne fonctionne pas");
            }
            finally
            {
                var servicePage = homePage.GoToCustomers_ServicePage();
                servicePage.ResetFilters();
                servicePage.Filter(ServicePage.FilterType.Search, serviceName);
                var servicepricePage = servicePage.SelectFirstRowInCustmerService();
                var foodpacketpage = servicepricePage.GoToFoodPacketsTab();
                var number = foodpacketpage.GetNumberOfPacketFoodsInService();
                if (number > 0)
                {
                    foodpacketpage.DeleteRows(number);
                }
            }
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_FOPA_Filter_ShowOnlyServicesWithPrices()
        {
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            FoodPacketPage foodPacketPage = homePage.GoToCustomers_FoodPacketsPage();
            foodPacketPage.ResetFilters();
            var numberDefaultListFoodPacket = foodPacketPage.GetNumberOfFoodPacketHeader();
            foodPacketPage.Filter(FoodPacketPage.FilterType.ShowOnlyServicesWithPrices, true);
            var numberFilterListFoodPacket = foodPacketPage.GetNumberOfFoodPacketHeader();
            Assert.AreNotEqual(numberFilterListFoodPacket, numberDefaultListFoodPacket, "Show Only Services With Prices filter ne fonctionne pas.");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_FOPA_ResetFilter()
        {
            HomePage homePage =  LogInAsAdmin();
            FoodPacketPage foodPacketPage = homePage.GoToCustomers_FoodPacketsPage();
            foodPacketPage.ResetFilters();
            var numberDefaultListSiteFilter = foodPacketPage.GetNumberSelectedSiteFilter();
            var numberDefaultListCustomerFilter = foodPacketPage.GetNumberSelectedCustomerFilter();

            foodPacketPage.Filter(FoodPacketPage.FilterType.Search, "ATLANTIC BEVERAGES");
            foodPacketPage.Filter(FoodPacketPage.FilterType.Site, "ACE");
            foodPacketPage.Filter(FoodPacketPage.FilterType.Customers, "1000 - CP PINTOR JOAN MIRO");
            foodPacketPage.Filter(FoodPacketPage.FilterType.ShowOnlyServicesWithPrices, true);
            foodPacketPage.ResetFilters();

            var searchValue = foodPacketPage.GetFilterValue(FoodPacketPage.FilterType.Search);
            Assert.AreEqual("", searchValue, "ResetFilter Search ''");
            var numberSelectedSiteFilter = foodPacketPage.GetNumberSelectedSiteFilter();
            Assert.AreEqual(numberSelectedSiteFilter, numberDefaultListSiteFilter, "ResetFilter Site ''");
            var numberSelectedCustomerFilter = foodPacketPage.GetNumberSelectedCustomerFilter();
            Assert.AreEqual(numberSelectedCustomerFilter, numberDefaultListCustomerFilter, "ResetFilter Customer ''");
            var showOnlyServicesWithPrices = foodPacketPage.GetFilterValue(FoodPacketPage.FilterType.ShowOnlyServicesWithPrices);
            Assert.AreEqual(false, showOnlyServicesWithPrices, "ResetFilter ShowOnlyServicesWithPrices");
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_FOPA_Export()
        {
            string downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            HomePage homePage =  LogInAsAdmin();
            FoodPacketPage foodPacketPage = homePage.GoToCustomers_FoodPacketsPage();
            foodPacketPage.ResetFilters();

            var listFoodPacketName = foodPacketPage.GetListFoodPacketName();
            //clear download directory 
            foreach (var file in new DirectoryInfo(downloadsPath).GetFiles())
            {
                file.Delete();
            }
            //export
            foodPacketPage.Export();

            // On récupère les fichiers du répertoire de téléchargement
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();


            // On recherche le fichier exporté avec le nom et la date d'écriture du fichier.
            var correctDownloadedFile = foodPacketPage.GetExportExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile);

            var fileName = correctDownloadedFile.Name;
            var filePath = Path.Combine(downloadsPath, fileName);
            var namesExcel = OpenXmlExcel.GetValuesInList("Food packet name", "export-foodpackets", filePath);
            //verifier les données 
            var verify = foodPacketPage.VerifyExcel(listFoodPacketName, namesExcel);
            Assert.IsTrue(verify, "les donées du grid et du fichier Excel ne sont pas identiques");
        }

        [Timeout(_timeout)]
		[TestMethod]
        [Ignore]
        public void CU_FOPA_FoldAndUnfold()
        {
            // Prepare
            string service = "FoodPacketService";
            string site = "ACE";
            string customer = "AIR CPU SL";
            string foodPacket = "ATLANTIC BEVERAGES UPB";
            // Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            // Act
            // faut-il créer un service (avec price) relié à ce Food Packet avant ?
            // en faire un prérequis ?
            ServicePage servicePage = homePage.GoToCustomers_ServicePage();
            servicePage.ResetFilters();
            servicePage.Filter(ServicePage.FilterType.Search, service);
            if (servicePage.CheckTotalNumber() == 0)
            {
                //creation
                ServiceCreateModalPage modal = servicePage.ServiceCreatePage();
                modal.FillFields_CreateServiceModalPage(service, service, service);
                ServiceGeneralInformationPage generalInfo = modal.Create();
                ServicePricePage prices = generalInfo.GoToPricePage();
                ServiceCreatePriceModalPage modalPrice = prices.AddNewCustomerPrice();
                prices = modalPrice.FillFields_CustomerPrice(site, customer, DateUtils.Now,DateUtils.Now.AddDays(20),"20");
                ServiceFoodPacketsPage packets = prices.GoToFoodPacketsTab();
                ServiceCreateFoodPacketModal packetsModal = packets.CreateFoodPackets();
                packetsModal.FillServiceFoodPacket(site, customer, foodPacket);
            }
            else
            {
                //maj date price
                ServicePricePage pricePage = servicePage.ClickOnFirstService();
                pricePage.UnfoldAll();
                ServiceCreatePriceModalPage modalFirstPrice = pricePage.EditFirstPrice(site, customer);
                pricePage = modalFirstPrice.FillFields_EditCustomerPrice(site, customer, DateUtils.Now, DateUtils.Now.AddDays(20));
                //FIXME prices.ClickOnFoodPacksTab()
                //FIXME à retirer lorsque ce sera stabilisé le ticket dev
                ServiceFoodPacketsPage packets = pricePage.GoToFoodPacketsTab();
                ServiceCreateFoodPacketModal packetsModal = packets.CreateFoodPackets();
                packetsModal.FillServiceFoodPacket(site, customer, foodPacket);
            }
            FoodPacketPage foodPacketPage = homePage.GoToCustomers_FoodPacketsPage();
            foodPacketPage.ResetFilters();
            foodPacketPage.Filter(FoodPacketPage.FilterType.Search, foodPacket);
            //Cliquer sur ˃
            foodPacketPage.FoldOrUnfold();
            //Une liste déroulante se déplie avec le nom du service
            Assert.IsFalse(foodPacketPage.IsServiceError(), "Message d'erreur aucun service pour ce Food Packet");
            foodPacketPage.FoldOrUnfold();
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void CU_FOPA_Access()
        {
            HomePage homePage =  LogInAsAdmin();
            FoodPacketPage foodPacketPage = homePage.GoToCustomers_FoodPacketsPage();
            Assert.IsNotNull(foodPacketPage, "Erreur : La page Food Packet n'a pas pu être chargée.");
            Assert.IsTrue(foodPacketPage.IsPageLoaded(), "Erreur : Timeout lors du chargement de la page Customer > Food Packet.");
        }
        }
}
