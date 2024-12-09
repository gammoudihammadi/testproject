using DocumentFormat.OpenXml.Office2010.ExcelAc;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Reconciliation;
using Newrest.Winrest.FunctionalTests.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using static Newrest.Winrest.FunctionalTests.PageObjects.Customers.Reconciliation.ReconciliationPage;

namespace Newrest.Winrest.FunctionalTests.Customers
{
    [TestClass]
    public class ReconciliationTest : TestBase
    {
        private const int _timeout = 600000;
        /// <summary>
        /// DATA FOR RECON FILTER TESTS
        /// </summary>
        private List<string> FlightKeys = new List<string>() { "999999", "123456" };
        private List<string> FlightNumbers = new List<string>() { "999999", "IB0123" };
        private List<string> CustomerIATA = new List<string>() { "AFT", "AFN" };
        private List<NotificationErrorType> ErrorTypes = new List<NotificationErrorType>() { NotificationErrorType.Barset };
        private List<ReconciliationPage.NotificationType> NotificationStates = new List<ReconciliationPage.NotificationType>() { ReconciliationPage.NotificationType.InProgress, ReconciliationPage.NotificationType.Done , ReconciliationPage.NotificationType.Failed };
        private List<ReconciliationPage.NotificationDataType> NotificationDataTypes = new List<ReconciliationPage.NotificationDataType>() { ReconciliationPage.NotificationDataType.AirfiPurchase, ReconciliationPage.NotificationDataType.AirfiStockConfirmation , ReconciliationPage.NotificationDataType.AirfiStockMutation };
        private readonly DateTime FlightDateFrom = new DateTime(2024, 05, 13);
        private readonly DateTime FlightDateTo = new DateTime(2024, 05, 17);
        private readonly DateTime ReceptionDateFrom = new DateTime(2024, 05, 14);
        private readonly DateTime ReceptionDateTo = new DateTime(2024, 05, 17);
        private readonly DateTime ReceptionDateForError = new DateTime(2024, 05, 25);


        #region Notifications

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_RE_FilterCustomer()
        {
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Process
            var reconPage = homePage.GoToCustomers_ReconPage();
            reconPage.GoToNotificationsTab();
            reconPage.ResetFilters();
            reconPage.Filter(ReconciliationPage.FilterType.Customers, CustomerIATA[0]);
            reconPage.Filter(ReconciliationPage.FilterType.ReceptionDateFrom, ReceptionDateFrom);
            reconPage.Filter(ReconciliationPage.FilterType.ReceptionDateTo, ReceptionDateTo);

            int totalNbr = reconPage.GetNumberOfNotifications();
            Assert.IsTrue(totalNbr == 1);
        }


        [Timeout(_timeout)]
		[TestMethod]
        public void CU_RE_FilterSearchById()
        {
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Process
            var reconPage = homePage.GoToCustomers_ReconPage();
            reconPage.GoToNotificationsTab();
            reconPage.ResetFilters();
            reconPage.Filter(ReconciliationPage.FilterType.SearchbyNotificationId, "2");
            reconPage.Filter(ReconciliationPage.FilterType.ReceptionDateFrom, ReceptionDateFrom);
            reconPage.Filter(ReconciliationPage.FilterType.ReceptionDateTo, ReceptionDateTo);

            int totalNbr = reconPage.GetNumberOfNotifications();
            Assert.IsTrue(totalNbr == 1);
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_RE_FilterReceptionDateFromAndTo()
        {
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Process
            var reconPage = homePage.GoToCustomers_ReconPage();
            reconPage.GoToNotificationsTab();
            reconPage.ResetFilters();
            reconPage.Filter(ReconciliationPage.FilterType.ReceptionDateFrom, ReceptionDateFrom);
            reconPage.Filter(ReconciliationPage.FilterType.ReceptionDateTo, ReceptionDateTo);

            int totalNbr = reconPage.GetNumberOfNotifications();
            Assert.IsTrue(totalNbr == 1);
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_RE_FilterSortBy()
        {
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Process
            var reconPage = homePage.GoToCustomers_ReconPage();
            reconPage.GoToNotificationsTab();
            reconPage.ResetFilters();
            reconPage.Filter(ReconciliationPage.FilterType.ReceptionDateFrom, ReceptionDateFrom);
            reconPage.Filter(ReconciliationPage.FilterType.ReceptionDateTo, new DateTime(2024, 05, 25));

            foreach (ReconciliationPage.SortByType value in Enum.GetValues(typeof(ReconciliationPage.SortByType)))
            {
                string errorMsg = $"An error occured while sorting by {value.ToString()}";
                //As it will be sorted by Id firstly, we will have it in ascending order in this test
                if (value == ReconciliationPage.SortByType.Id)
                {
                    reconPage.SortBy(value);
                    var customerList = reconPage.GetCustemerNotificationsList();
                    Assert.IsTrue(customerList.OrderBy(x => x).SequenceEqual(customerList) || customerList.OrderByDescending(x => x).SequenceEqual(customerList), "La liste des notifications client n'est pas ordonnée");

                }
            }
        } 

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_RE_FilterNotificationState()
        {
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Process
            var reconPage = homePage.GoToCustomers_ReconPage();
            reconPage.GoToNotificationsTab();
            reconPage.ResetFilters();
            reconPage.Filter(ReconciliationPage.FilterType.ReceptionDateFrom, ReceptionDateFrom);
            reconPage.Filter(ReconciliationPage.FilterType.ReceptionDateTo, ReceptionDateTo);
            reconPage.Filter(ReconciliationPage.FilterType.NotificationStates, NotificationStates[2]); 

            int totalNbr = reconPage.GetNumberOfNotifications();
            Assert.IsTrue(totalNbr >= 1, "Could not filter properly by state 'Failed'");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_RE_FilterSearchByFlightInfos()
        {
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            string errorMsg = "An error occured while filtering by {0}";

            //Process
            var reconPage = homePage.GoToCustomers_ReconPage();
            reconPage.GoToNotificationsTab();
            reconPage.ResetFilters();
            reconPage.Filter(ReconciliationPage.FilterType.ReceptionDateFrom, ReceptionDateFrom);
            reconPage.Filter(ReconciliationPage.FilterType.ReceptionDateTo, ReceptionDateTo);

            //Filter by Barset
            string method = "Barset";
            reconPage.Filter(ReconciliationPage.FilterType.SearchByBarset, FlightKeys[0]);
            int totalNbr = reconPage.GetNumberOfNotifications();
            Assert.IsTrue(totalNbr == 1, string.Format(errorMsg, method));

            //Filter by flight number
            method = "Flight Number";
            reconPage.Filter(ReconciliationPage.FilterType.SearchByBarset, "");
            reconPage.Filter(ReconciliationPage.FilterType.SearchByFlightNumber, FlightNumbers[0]);
            totalNbr = reconPage.GetNumberOfNotifications();
            Assert.IsTrue(totalNbr == 1, string.Format(errorMsg, method));

            //Filter by flight date from
            method = "Flight Date From";
            reconPage.Filter(ReconciliationPage.FilterType.SearchByFlightNumber, "");
            reconPage.Filter(ReconciliationPage.FilterType.SearchByFlightDateFrom, FlightDateFrom);
            reconPage.Filter(ReconciliationPage.FilterType.SearchByFlightDateTo, FlightDateFrom);
            totalNbr = reconPage.GetNumberOfNotifications();
            Assert.IsTrue(totalNbr == 0, string.Format(errorMsg, method));

            //Filter by flight date to
            method = "Flight Date To";
            reconPage.Filter(ReconciliationPage.FilterType.SearchByFlightDateFrom, FlightDateTo);
            reconPage.Filter(ReconciliationPage.FilterType.SearchByFlightDateTo, FlightDateTo);
            totalNbr = reconPage.GetNumberOfNotifications();
            Assert.IsTrue(totalNbr == 1, string.Format(errorMsg, method));
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_RE_FilterNotificationType()
        {
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Process
            var reconPage = homePage.GoToCustomers_ReconPage();
            reconPage.GoToNotificationsTab();
            reconPage.ResetFilters();
            reconPage.Filter(ReconciliationPage.FilterType.ReceptionDateFrom, ReceptionDateFrom);
            reconPage.Filter(ReconciliationPage.FilterType.ReceptionDateTo, ReceptionDateTo);

            reconPage.Filter(ReconciliationPage.FilterType.NotificationType, NotificationDataTypes[1].ToString());
            int totalNbr = reconPage.GetNumberOfNotifications();
            Assert.AreEqual(1, totalNbr);
            reconPage.Filter(ReconciliationPage.FilterType.ReceptionDateTo, new DateTime(2024, 05, 25));
            reconPage.Filter(ReconciliationPage.FilterType.NotificationType, NotificationDataTypes[2].ToString());
            totalNbr = reconPage.GetNumberOfNotifications();
            //Not removing previous filter, so we expect 2 results
            Assert.AreEqual(2, totalNbr);
        }


        [Timeout(_timeout)]
		[TestMethod]
        public void CU_RE_FilterErrorType()
        {
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Process
            var reconPage = homePage.GoToCustomers_ReconPage();
            reconPage.GoToNotificationsTab();
            reconPage.ResetFilters();
            reconPage.Filter(ReconciliationPage.FilterType.ReceptionDateFrom, ReceptionDateForError);
            reconPage.Filter(ReconciliationPage.FilterType.ReceptionDateTo, ReceptionDateForError.AddDays(1));

            reconPage.Filter(ReconciliationPage.FilterType.ErrorType, ErrorTypes[0].ToString());
            int totalNbr = reconPage.GetNumberOfNotifications();
            Assert.IsTrue(totalNbr == 1);
        }


        [Timeout(_timeout)]
		[TestMethod]
        public void CU_RE_FilterSearchByDisplayTests()
        {
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Process
            var reconPage = homePage.GoToCustomers_ReconPage();
            reconPage.GoToNotificationsTab();
            reconPage.ResetFilters();
            reconPage.Filter(ReconciliationPage.FilterType.ReceptionDateFrom, ReceptionDateForError);
            reconPage.Filter(ReconciliationPage.FilterType.ReceptionDateTo, ReceptionDateForError.AddDays(1));

            reconPage.Filter(ReconciliationPage.FilterType.DisplayTestCases, true);
            int totalNbr = reconPage.GetNumberOfNotifications();
            Assert.IsTrue(totalNbr == 1);
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_RE_ExportNotifications()
        {
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Process
            var reconPage = homePage.GoToCustomers_ReconPage();
            reconPage.ResetFilters();
            reconPage.ExportExcel(true);
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            var correctDownloadedFile = reconPage.GetExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_RE_ExportOrders()
        {
            //Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            //Arrange
            HomePage homePage= LogInAsAdmin();
            homePage.ClearDownloads();

            //ACT
            var reconPage = homePage.GoToCustomers_ReconPage();
            reconPage.ResetFilters();
            reconPage.Go_To_Orders_ActionLink();
            reconPage.ExportExcel(true);
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();
            //Assert
            var correctDownloadedFile = reconPage.GetSalesSummaryExcelFile(taskFiles);
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_RE_ExportSalesSummary()
        {
            //Prepare
            var downloadsPath = TestContext.Properties["DownloadsPath"].ToString();

            //Arrange
            HomePage homePage= LogInAsAdmin();

            //Process
            var reconPage = homePage.GoToCustomers_ReconPage();
            reconPage.ResetFilters();
            reconPage.Go_To_Orders_ActionLink();
            reconPage.ExportSalesSummaryExcel(true);
            DirectoryInfo taskDirectory = new DirectoryInfo(downloadsPath);
            FileInfo[] taskFiles = taskDirectory.GetFiles();

            var correctDownloadedFile = reconPage.GetSalesSummaryExcelFile(taskFiles);
            reconPage.WaitPageLoading();
            Assert.IsNotNull(correctDownloadedFile, "Aucun fichier correspondant au pattern recherché n'a été trouvé.");

        }

        [Timeout(_timeout)]
		[TestMethod]
        public void CU_RE_FoldAndUnfold()
        {

            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();

            //Process
            var reconPage = homePage.GoToCustomers_ReconPage();
            reconPage.ResetFilters();
            reconPage.Filter(FilterType.ReceptionDateFrom, ReceptionDateFrom);
            reconPage.Fold();
            var IsFoldAl = reconPage.IsFoldAll();
            //assert 
            Assert.IsTrue(IsFoldAl, "le fold n'est fonctionne pas.");

            reconPage.Fold();

            var IsUnfoldAll = reconPage.IsUnfoldAll(); 
            Assert.IsTrue(reconPage.IsUnfoldAll(), "le unfold n'est fonctionne pas.");

        }
        #endregion


    }
}
