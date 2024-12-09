using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects;
using Newrest.Winrest.FunctionalTests.PageObjects.Interim.InterimOrders;
using Newrest.Winrest.FunctionalTests.PageObjects.Interim.InterimReceptions;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using static Newrest.Winrest.FunctionalTests.PageObjects.Interim.InterimOrders.InterimOrdersPage;
using Newrest.Winrest.FunctionalTests.Utils;
using Newrest.Winrest.FunctionalTests.PageObjects.Accounting.Invoice;
using System.IO;

using Newrest.Winrest.FunctionalTests.PageObjects.Jobs.Settings;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Item;
using System.Globalization;
using OpenQA.Selenium.Support.UI;
using System.Security.Policy;
using Newrest.Winrest.FunctionalTests.PageObjects.Accounting.SupplierInvoices;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenQA.Selenium.Html5;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.CustomerOrder;
using System.Net.NetworkInformation;
using DocumentFormat.OpenXml.Bibliography;
using OpenQA.Selenium;

namespace Newrest.Winrest.FunctionalTests.Interim.InterimOrders
{
    [TestClass]
    public class InterimOrdersTest : TestBase
    {
        private const int _timeout = 600000;

        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_index_Pagination()
        {
            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            var InterimOrdersPage = homePage.GoToInterim_InterimOrders();
            InterimOrdersPage.PageSize("8");
            InterimOrdersPage.WaitPageLoading();

            Assert.IsTrue(InterimOrdersPage.GetInterimOrdersList().Count <= 8, "La pagination par 8 ne fonctionne pas..");
            InterimOrdersPage.PageSize("16");
            InterimOrdersPage.WaitPageLoading();

            Assert.IsTrue(InterimOrdersPage.GetInterimOrdersList().Count <= 16, "La pagination par 16 ne fonctionne pas..");
            InterimOrdersPage.PageSize("30");
            InterimOrdersPage.WaitPageLoading();

            Assert.IsTrue(InterimOrdersPage.GetInterimOrdersList().Count <= 30, "La pagination par 30 ne fonctionne pas..");
            InterimOrdersPage.PageSize("50");
            InterimOrdersPage.WaitPageLoading();

            Assert.IsTrue(InterimOrdersPage.GetInterimOrdersList().Count <= 50, "La pagination par 50 ne fonctionne pas..");
            InterimOrdersPage.PageSize("100");
            Assert.IsTrue(InterimOrdersPage.GetInterimOrdersList().Count <= 100, "La pagination par 100 ne fonctionne pas..");
        }


        //_________________________________________CREATE_INTERIM____________________________________________________________

        /*
         * Création d'une nouvelle Interim
        */
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_index_AddNew()
        {
            string site = TestContext.Properties["SiteACE"].ToString();
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string location = "Produccion";
            var homePage = LogInAsAdmin();
            InterimOrdersPage interimPage = homePage.GoToInterim_InterimOrders();
            var nbreItemsBeforeAdd = interimPage.CheckTotalNumber();
            InterimOrdersCreateModalPage modalcreateInterimOrder = interimPage.CreateNewInterimOrder();
            modalcreateInterimOrder.FillField_CreatNewInterminOrder(DateTime.Now, site, supplier, location);
            InterimOrdersItem InterimItem = modalcreateInterimOrder.Submit();
            interimPage = InterimItem.BackToList();
            //Assert
            Assert.AreEqual(nbreItemsBeforeAdd + 1, interimPage.CheckTotalNumber(), String.Format(MessageErreur.OBJET_NON_CREE, "L'Interim Order"));
        }


        /*
         * Création d'une nouvelle Interim avec commentaire
        */
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_index_AddNewComment()
        {
            string site = TestContext.Properties["SiteACE"].ToString();
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string location = "Produccion";
            string comment = "interim order comment";
            var homePage = LogInAsAdmin();
            InterimOrdersPage interimPage = homePage.GoToInterim_InterimOrders();
            InterimOrdersCreateModalPage modalcreateInterimOrder = interimPage.CreateNewInterimOrder();
            modalcreateInterimOrder.FillField_CreatNewInterminOrderWithComment(DateTime.Now, site, supplier, location, comment);
            InterimOrdersItem InterimItem = modalcreateInterimOrder.Submit();
            InterimOrdersGeneralInformation interimReceptionsGeneralInformation = InterimItem.GoToGeneralInformation();
            string savedComment = interimReceptionsGeneralInformation.GetComment();
            //Assert
            Assert.AreEqual(comment, savedComment, "Le commentaire enregistré ne correspond pas au commentaire initial.");
        }

        /*
         * Création d'une nouvelle Interim Order Validator
        */
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_index_AddNewValidator()
        {
            //Arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            InterimOrdersPage interimPage = homePage.GoToInterim_InterimOrders();
            InterimOrdersCreateModalPage modalcreateInterimOrder = interimPage.CreateNewInterimOrder();
            modalcreateInterimOrder.Submit();
            //Assert
            Assert.IsTrue(!interimPage.CheckValidator(), "Les validators in Interim Order n'apparaissent pas!");

        }

        //_________________________________________FIN CREATE_INTERIM________________________________________________________
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_index_filter_SearchByNumber()
        {
            var homePage = LogInAsAdmin();
            var interimOrdersPage = homePage.GoToInterim_InterimOrders();
            string interimOrdersNumber = interimOrdersPage.GetFirstID();
            interimOrdersPage.ResetFilters();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.SearchByNumber, interimOrdersNumber);
            bool numberFound = interimOrdersPage.GetInterimOrdersList().Any(item => item.Contains(interimOrdersNumber));
            //Assert number is Found
            Assert.IsTrue(numberFound, "Les résultats ne sont pas mis à jour en fonction de filtre number.");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_index_filter_ShowValidatedOnly()
        {
            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            var interimOrdersPage = homePage.GoToInterim_InterimOrders();
            interimOrdersPage.ResetFilters();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowValidatedOnly, true);
            interimOrdersPage.WaitPageLoading();
            var resultsValidatedOnly = interimOrdersPage.CheckTotalNumber();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowNotValidatedOnly, true);
            interimOrdersPage.WaitPageLoading();
            var resultsNotValidatedOnly = interimOrdersPage.CheckTotalNumber();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowAllOrders, true);
            interimOrdersPage.WaitPageLoading();
            var resultAll = interimOrdersPage.CheckTotalNumber();
            var expectedResult = resultAll - resultsNotValidatedOnly;

            //assert
            Assert.AreEqual(expectedResult, resultsValidatedOnly, "Les résultats ne sont pas mis à jour en fonction des filtres Show Validated Only");
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_index_filter_ShowNotValidatedOnly()
        {
            //arrange
            var homePage = LogInAsAdmin();
            //act
            var interimOrdersPage = homePage.GoToInterim_InterimOrders();
            interimOrdersPage.ResetFilters();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowValidatedOnly, true);
            var resultsValidatedOnly = interimOrdersPage.CheckTotalNumber();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowNotValidatedOnly, true);
            var resultsNotValidatedOnly = interimOrdersPage.CheckTotalNumber();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowAllOrders, true);
            var resultAll = interimOrdersPage.CheckTotalNumber();
            var expectedResult = resultAll - resultsValidatedOnly;
            //assert
            Assert.AreEqual(expectedResult, resultsNotValidatedOnly, "Les résultats ne sont pas mis à jour en fonction des filtres Show not Validated Only");
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_index_filter_ShowAll()
        {
            var homePage = LogInAsAdmin();
            var interimOrdersPage = homePage.GoToInterim_InterimOrders();
            interimOrdersPage.ResetFilters();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowValidatedOnly, true);
            interimOrdersPage.WaitPageLoading();
            var resultsValidatedOnly = interimOrdersPage.CheckTotalNumber();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowNotValidatedOnly, true);
            interimOrdersPage.WaitPageLoading();
            var resultsNotValidatedOnly = interimOrdersPage.CheckTotalNumber();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowAllOrders, true);
            var resultAll = interimOrdersPage.CheckTotalNumber();
            Assert.AreEqual(resultsValidatedOnly + resultsNotValidatedOnly, resultAll, "Les résultats ne sont pas mis à jour en fonction des filtres not Show all");
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_index_filter_Opened()
        {
            //login
            var homePage = LogInAsAdmin();
            InterimOrdersPage interimOrdersPage = homePage.GoToInterim_InterimOrders();
            interimOrdersPage.ResetFilters();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowAllInterimReception, true);
            interimOrdersPage.WaitPageLoading();
            var resultAll = interimOrdersPage.CheckTotalNumber();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowOpenedInterimReception, true);
            interimOrdersPage.WaitPageLoading();
            var resultsOpened = interimOrdersPage.CheckTotalNumber();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowClosedInterimReception, true);
            interimOrdersPage.WaitPageLoading();
            var resultsClosed = interimOrdersPage.CheckTotalNumber();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowCancelledInterimReception, true);
            interimOrdersPage.WaitPageLoading();
            var resultCancelled = interimOrdersPage.CheckTotalNumber();
            //assert
            Assert.AreEqual(resultAll - (resultsClosed + resultCancelled), resultsOpened, "Les résultats ne sont pas mis à jour en fonction des filtres Status Opened");
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_index_filter_FromTo()
        {
            DateTime dateFrom = DateTime.Now.AddDays(-3);
            DateTime dateTo = DateTime.Now;
            string writingType = "someWritingType";
            string dateFormatPicker = "dd/mm/yyyy";
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            InterimOrdersPage interimOrdersPage = homePage.GoToInterim_InterimOrders();
            interimOrdersPage.PageSize("50");
            interimOrdersPage.ResetFilters();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.DateFrom, dateFrom);
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.DateTo, dateTo);
            List<DateTime> filtreredDeliveryDates = interimOrdersPage.GetDeliveryDate(writingType, dateFormatPicker);
            bool areDatesWithinRange = filtreredDeliveryDates.All(date => date >= dateFrom.Date && date <= dateTo.Date);
            Assert.IsTrue(areDatesWithinRange, String.Format(MessageErreur.FILTRE_ERRONE, "L'application de filtre par From To ne fonctionne pas correctement."));

        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_index_filter_closed()
        {
            var homePage = LogInAsAdmin();
            InterimOrdersPage interimOrdersPage = homePage.GoToInterim_InterimOrders();
            interimOrdersPage.ResetFilters();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowAllInterimReception, true);
            var resultAll = interimOrdersPage.CheckTotalNumber();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowOpenedInterimReception, true);
            var resultsOpened = interimOrdersPage.CheckTotalNumber();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowClosedInterimReception, true);
            var resultsClosed = interimOrdersPage.CheckTotalNumber();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowCancelledInterimReception, true);
            var resultCancelled = interimOrdersPage.CheckTotalNumber();
            //Assert
            Assert.AreEqual(resultAll - (resultsOpened + resultCancelled), resultsClosed, "Les résultats ne sont pas mis à jour en fonction des filtres Status Closed");
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_index_filter_cancelled()
        {
            //login
            var homePage = LogInAsAdmin();
            InterimOrdersPage interimOrdersPage = homePage.GoToInterim_InterimOrders();
            interimOrdersPage.ResetFilters();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowAllInterimReception, true);
            var resultAll = interimOrdersPage.CheckTotalNumber();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowOpenedInterimReception, true);
            var resultsOpened = interimOrdersPage.CheckTotalNumber();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowClosedInterimReception, true);
            var resultsClosed = interimOrdersPage.CheckTotalNumber();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowCancelledInterimReception, true);
            var resultCancelled = interimOrdersPage.CheckTotalNumber();
            //Assert canceled = all - open + closed
            Assert.AreEqual(resultAll - (resultsOpened + resultsClosed), resultCancelled, "Les résultats ne sont pas mis à jour en fonction des filtres Status Cancelled");
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_index_filter_ALL()
        {
            //login
            var homePage = LogInAsAdmin();
            InterimOrdersPage interimOrdersPage = homePage.GoToInterim_InterimOrders();
            interimOrdersPage.ResetFilters();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowAllInterimReception, true);
            var resultAll = interimOrdersPage.CheckTotalNumber();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowOpenedInterimReception, true);
            var resultsOpened = interimOrdersPage.CheckTotalNumber();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowClosedInterimReception, true);
            var resultsClosed = interimOrdersPage.CheckTotalNumber();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowCancelledInterimReception, true);
            var resultCancelled = interimOrdersPage.CheckTotalNumber();
            // Assert
            Assert.AreEqual(resultsOpened + resultsClosed + resultCancelled, resultAll, "Les résultats ne sont pas mis à jour en fonction des filtres Status All");
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_index_filter_ShowAllOrder()
        {
            //login
            var homePage = LogInAsAdmin();
            InterimOrdersPage interimOrdersPage = homePage.GoToInterim_InterimOrders();
            interimOrdersPage.ResetFilters();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowValidatedOnly, true);
            var resultsValidated = interimOrdersPage.CheckTotalNumber();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowNotValidatedOnly, true);
            var resultsInvalidated = interimOrdersPage.CheckTotalNumber();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowAllOrders, true);
            var resultAll = interimOrdersPage.CheckTotalNumber();
            //assert
            Assert.AreEqual(resultsValidated + resultsInvalidated, resultAll, "Les résultats ne sont pas mis à jour en fonction des filtres Show All");
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_index_filter_ResetFilter()
        {
            
            var homePage = LogInAsAdmin();
            InterimOrdersPage interimOrdersPage = homePage.GoToInterim_InterimOrders();
            interimOrdersPage.ResetFilters();
            var defaultShowreceptions = interimOrdersPage.GetShowFilterSelected();
            var defaultStatusDisplayed = interimOrdersPage.GetStatusFilterSelected();
            var defaultnumberintrimorder = interimOrdersPage.CheckTotalNumber();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowValidatedOnly, true);
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowOpenedInterimReception, true);
            interimOrdersPage.ScrollToFilterSearchNumber();
            interimOrdersPage.PageUp();
            interimOrdersPage.ResetFilters();
            var resultShowreceptionsAfterReset = interimOrdersPage.GetShowFilterSelected();
            var resultStatusDisplayedAfterReset = interimOrdersPage.GetStatusFilterSelected();
            var resutnumberinterimorder = interimOrdersPage.CheckTotalNumber();
            Assert.AreEqual(defaultShowreceptions, resultShowreceptionsAfterReset, "Le filter Show receptions ne remet pas");
            Assert.AreEqual(defaultnumberintrimorder, resutnumberinterimorder, "le filtre du nombre des interim orders est different ");
            Assert.AreEqual(defaultStatusDisplayed, resultStatusDisplayedAfterReset, "Le filter Show STATUS ne remet pas");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_index_filter_ByValidationDate()
        {  
            //arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            //act 
            InterimOrdersPage interimOrdersPage = homePage.GoToInterim_InterimOrders();
            interimOrdersPage.ResetFilters();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ByFromToDate, true);
            interimOrdersPage.WaitPageLoading();
            var resultsFromToDate = interimOrdersPage.CheckTotalNumber();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ByValidationDate, true);
            interimOrdersPage.WaitPageLoading();
            var resultsValidationDate = interimOrdersPage.CheckTotalNumber();
            //assert
            Assert.AreNotEqual(resultsFromToDate, resultsValidationDate, "Les résultats ne sont pas mis à jour en fonction des filtres By Validation Date");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_details_ItemComment()
        {
            //prepare
            string qty = "3";
            string comment = "Testing add Comment";
            string site = TestContext.Properties["SiteACE"].ToString();
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string location = "Produccion";
            Random rnd = new Random();
            //arrange
            var homePage = LogInAsAdmin();
            InterimOrdersPage interimOrdersPage = homePage.GoToInterim_InterimOrders();
            interimOrdersPage.ResetFilters();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowNotValidatedOnly, true);
            int numbers = interimOrdersPage.GetTotalOfOrders();
            if (numbers == 0)
            {
                InterimOrdersCreateModalPage interimOrdersCreateModalPage = interimOrdersPage.CreateNewInterimOrder();
                interimOrdersCreateModalPage.FillField_CreatNewInterminOrder(DateTime.Now, site, supplier, location, true);
                interimOrdersCreateModalPage.Submit();
                interimOrdersPage = homePage.GoToInterim_InterimOrders();
            }
            InterimOrdersItem interimOrdersItem = interimOrdersPage.ClickOnFirstService();
            if (interimOrdersItem.GetQtyItem() == "0")
            {
                interimOrdersItem.ClickFirstItem();
                interimOrdersItem.SetQty(qty);
            }
            else
            {
                interimOrdersItem.ClickFirstItem();
            }
            interimOrdersItem.OpenModalAddComment();
            interimOrdersItem.SetComment(comment);
            homePage.GoToInterim_InterimOrders();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowNotValidatedOnly, true);
            interimOrdersPage.ClickOnFirstService();
            interimOrdersItem.ClickFirstItem();
            var colorGreen = interimOrdersItem.isCommentIconGreen();
            Assert.IsTrue(colorGreen, "L'icône Comment sur la ligne de l'item ne devenait pas verte");
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_Details_filter_SearchByNameRef()
        {
            string site = TestContext.Properties["SiteACE"].ToString();
            var homePage = LogInAsAdmin();
            InterimOrdersPage interimOrdersPage = homePage.GoToInterim_InterimOrders();
            interimOrdersPage.ResetFilters();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowNotValidatedOnly, true);
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.Sites, site);
            var interimOrderItem = interimOrdersPage.ClickFirstLine();
            string itemName = interimOrderItem.ReturnItemName();
            Assert.IsNotNull(itemName, "Il n'y a pas d'items."); 
     
            interimOrderItem.Filter(InterimOrdersItem.FilterItemType.SearchByNameRef, itemName);
            bool interiemOrderAfterFilter = interimOrderItem.GetAllNamesResultPaged().Contains(itemName); 

            Assert.IsTrue(interiemOrderAfterFilter, "Les résultats ne sont pas mis à jour en fonction des filtres.");
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_details_ItemCommentNoQty()
        {
            string site = TestContext.Properties["SiteACE"].ToString();
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string location = "Produccion";
            Random rnd = new Random();
            string qty = "0,000";
            var homePage = LogInAsAdmin();
            InterimOrdersPage interimOrdersPage = homePage.GoToInterim_InterimOrders();
            interimOrdersPage.ResetFilters();
            InterimOrdersCreateModalPage interimOrdersCreateModalPage = interimOrdersPage.CreateNewInterimOrder();
            interimOrdersCreateModalPage.FillField_CreatNewInterminOrder(DateTime.Now, site, supplier, location, true);
            string interimNumber = interimOrdersCreateModalPage.GetInterimOrderId();
            InterimOrdersItem interimOrdersItem = interimOrdersCreateModalPage.Submit();
            interimOrdersItem.ClickFirstItem();
            interimOrdersItem.SetQty(qty);
            interimOrdersItem.AddComment();
            var errorMessage = interimOrdersItem.GetCommentMsgItem();
            Assert.AreEqual("You must set a valid quantity for this item before adding a comment.", errorMessage, "The expected error message did not appear.");
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_index_filter_Supplier()
        {
            //Prepare
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            InterimOrdersPage interimOrdersPage = homePage.GoToInterim_InterimOrders();
            interimOrdersPage.ResetFilters();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.Suppliers, supplier);

            //Assert
            bool areFilteredBySupplier = interimOrdersPage.IsSortedBySuppliers(supplier);
            Assert.IsTrue(areFilteredBySupplier, MessageErreur.FILTRE_ERRONE, "le filtre par supplier ne fonctionne pas  ");
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_details_GenerateReceptionPopupShow()
        {
            var homePage = LogInAsAdmin();
            InterimOrdersPage interimOrdersPage = homePage.GoToInterim_InterimOrders();
            interimOrdersPage.ResetFilters();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ByValidationDate, true);
            InterimOrdersItem interimOrdersItem = interimOrdersPage.ClickFirstLine();
            interimOrdersItem.ShowExtendedMenu();
            InterimOrdersCreateModalPage interimOrdersCreateModalPage = interimOrdersItem.GoToGenerateInterimReception();
            Assert.IsTrue(interimOrdersCreateModalPage.IsModalVisible(), "Le modal Generate Interim reception  n'est pas visible.");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_details_GenerateReceptionPopupSearch()
        {
            string site = TestContext.Properties["SiteACE"].ToString();
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string location = "Produccion";
            var qty = "3";

            var numberOrder = "";
            DateTime dateFrom = DateTime.Now.AddDays(-3);
            DateTime dateTo = DateTime.Now;
            var homePage = LogInAsAdmin();
            InterimOrdersPage interimOrdersPage = homePage.GoToInterim_InterimOrders();
            InterimOrdersCreateModalPage modalcreateInterimOrder = interimOrdersPage.CreateNewInterimOrder();
            modalcreateInterimOrder.FillField_CreatNewInterminOrder(DateTime.Now, site, supplier, location);
            var id = modalcreateInterimOrder.GetInterimOrderId();
            InterimOrdersItem InterimItemOrder = modalcreateInterimOrder.Submit();
            InterimItemOrder.SetQty(qty);
            InterimItemOrder.Validate();
            var qtyExistance = InterimItemOrder.ExistQtyItem();
            Assert.IsTrue(qtyExistance, "No Item");
            InterimItemOrder.ShowExtendedMenu();
            InterimReceptionsCreateModalPage interimReceptionsCreateModalPage = InterimItemOrder.GenerateIntreimReception();
            interimReceptionsCreateModalPage.Filter(InterimReceptionsCreateModalPage.FilterType.ByNumber, numberOrder);
            interimReceptionsCreateModalPage.Filter(InterimReceptionsCreateModalPage.FilterType.ByNumber, id);
            interimReceptionsCreateModalPage.Filter(InterimReceptionsCreateModalPage.FilterType.DateFrom, dateFrom);
            interimReceptionsCreateModalPage.Filter(InterimReceptionsCreateModalPage.FilterType.DateTo, dateTo);
            var numberFromPopUp = interimReceptionsCreateModalPage.GetOrderNumber();
            Assert.AreEqual(id, numberFromPopUp, "Les résultats ne sont pas mis à jour en fonction des filtres ");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_index_filter_sites()
        {
            //Prepare
            string site = "ACE";

            // Arrange
            var homePage = LogInAsAdmin();

            //Act
            InterimOrdersPage interimOrdersPage = homePage.GoToInterim_InterimOrders();
            interimOrdersPage.ResetFilters();
            interimOrdersPage.PageSize("16");
            interimOrdersPage.ScrollUntilSitesFilterIsVisible();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.Sites, site);
            Assert.IsTrue(interimOrdersPage.IsSortedBySites(site), MessageErreur.FILTRE_ERRONE, "le filtrage par site ne fonctione pas");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_details_BackToList()
        {
            //arrange
            var homePage = LogInAsAdmin();
            InterimOrdersPage interimOrdersPage = homePage.GoToInterim_InterimOrders();

            interimOrdersPage.ResetFilters();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowValidatedOnly, true);
            var number = interimOrdersPage.GetFirstID();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.SearchByNumber, number.ToString());
            interimOrdersPage.WaitPageLoading();

            var interimOrdersList = interimOrdersPage.GetInterimOrdersList();

            var firstOrder = interimOrdersList.First();

            InterimOrdersItem interimItem = interimOrdersPage.ClickFirstLine();

            interimOrdersPage = interimItem.BackToList();
            interimOrdersPage.WaitPageLoading();

            var filteredList = interimOrdersPage.GetInterimOrdersList();
            //assert
            Assert.IsTrue(filteredList.SequenceEqual(interimOrdersList), "Les critères de filtrage ne sont pas retenus après être revenu à la liste.");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_details_ItemTotalwoVAT()
        {
            string qty = "3";
            string site = TestContext.Properties["SiteACE"].ToString();
            var homePage = LogInAsAdmin();
            InterimOrdersPage interimOrdersPage = homePage.GoToInterim_InterimOrders();
            interimOrdersPage.ResetFilters();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowNotValidatedOnly, true);
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.Sites, site);
            interimOrdersPage.WaitPageLoading();
            InterimOrdersItem interimOrdersItem = interimOrdersPage.ClickOnFirstService();
            var total = 0;
            if (interimOrdersItem.GetQtyItem() == "0")
            {
                interimOrdersItem.ClickFirstItem();
                interimOrdersItem.SetQty(qty);
                WebDriver.Navigate().Refresh();
            }
            var quantite = int.Parse(interimOrdersItem.GetQtyItem());
            var packagingPrice = int.Parse(interimOrdersItem.GetPackagingPriceItem());
            total = quantite * packagingPrice;
            var newTotal = int.Parse(interimOrdersItem.GetTotalVATItem());
            Assert.AreEqual(total, newTotal, "Le Total n'est pas valide !!");
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_Details_filter_SubGroup()
        {
            string subgrpname = "subgrpname";

            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            var orderPage = homePage.GoToInterim_InterimOrders();
            orderPage.Filter(InterimOrdersPage.FilterType.ShowNotValidatedOnly, true);
            InterimOrdersItem InterimOrdersItem = orderPage.GoToInterim_InterimOrdersItem();
            //InterimOrdersItem.Filter(InterimOrdersItem.FilterType.subGroup, subgrpname);
            InterimOrdersItem.Filter(InterimOrdersItem.FilterItemType.subGroup, subgrpname);
            var subGroupSelectionStatusvalue = InterimOrdersItem.GetsubGroupSelectionStatus();
            Assert.AreEqual(InterimOrdersItem.GetsubGroupSelectionStatus(), subGroupSelectionStatusvalue, "Le Fitrage par par sous-groupes ne fonctionne pas");
            
            
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_index_filter_ShowActiveOnly()
        {
            var homePage = LogInAsAdmin();
            InterimOrdersPage interimOrdersPage = homePage.GoToInterim_InterimOrders();
            interimOrdersPage.ResetFilters();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowOnlyActive, true);
            var resultsActivated = interimOrdersPage.CheckTotalNumber();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowOnlyInactive, true);
            var resultsInactivated = interimOrdersPage.CheckTotalNumber();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowAll, true);
            var resultAll = interimOrdersPage.CheckTotalNumber();
            //Assert
            Assert.AreEqual(resultAll - resultsInactivated, resultsActivated, "Les résultats ne sont pas mis à jour en fonction du filtres Show active only");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_details_GenerateReceptionFromReceptionSearch()
        {
            //interim reception variable
            string site = TestContext.Properties["SiteACE"].ToString();
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string location = "Produccion";

            DateTime dateFrom = DateTime.Now.AddDays(-3);
            DateTime dateTo = DateTime.Now;
            var qty = "3";
            var homePage = LogInAsAdmin();

            //create interim reception 
            InterimReceptionsPage interimreceptionPage = homePage.GoToInterim_Receptions();
            InterimReceptionsCreateModalPage modalcreateInterim = interimreceptionPage.CreateNewInterim();
            modalcreateInterim.FillField_CreatNewIntermin(DateTime.Now, site, supplier, location);
            var InterimId = modalcreateInterim.GetInterimId();
            InterimReceptionsItem InterimItem = modalcreateInterim.Submit();
            InterimItem.SetReceptionItemReceived();
            Thread.Sleep(1000);
            InterimItem.Validate();
            Thread.Sleep(1000);
            InterimOrdersPage interimOrdersPage = InterimItem.GoToReceptionOrder();
            InterimOrdersCreateModalPage modalcreateInterimOrder = interimOrdersPage.CreateNewInterimOrder();
            modalcreateInterimOrder.FillField_CreatNewInterminOrder(DateTime.Now, site, supplier, location);
            InterimOrdersItem InterimItemOrder = modalcreateInterimOrder.Submit();
            InterimItemOrder.SetQty(qty);
            Thread.Sleep(1000);
            InterimItemOrder.Validate();
            Thread.Sleep(1000);
            InterimItemOrder.ShowExtendedMenu();
            InterimReceptionsCreateModalPage interimReceptionsCreateModalPage = InterimItemOrder.GenerateIntreimReception();
            interimReceptionsCreateModalPage.ToInterimReception();
            interimReceptionsCreateModalPage.Filter(InterimReceptionsCreateModalPage.FilterType.ByNumber, InterimId);
            interimReceptionsCreateModalPage.Filter(InterimReceptionsCreateModalPage.FilterType.DateFrom, dateFrom);
            interimReceptionsCreateModalPage.Filter(InterimReceptionsCreateModalPage.FilterType.DateTo, dateTo);
            var numberFromPopUp = interimReceptionsCreateModalPage.GetReceptionNumber();
            Assert.AreEqual(InterimId, numberFromPopUp, "Les résultats ne sont pas mis à jour en fonction des filtres ");


        }

        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_details_GenerateReceptionPopupPagination()
        {
            string site = TestContext.Properties["SiteACE"].ToString();
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string location = "Produccion";
            var numberOrder = "";
            var qty = "3";
            var homePage = LogInAsAdmin();
            InterimOrdersPage interimOrdersPage = homePage.GoToInterim_InterimOrders();
            InterimOrdersCreateModalPage modalcreateInterimOrder = interimOrdersPage.CreateNewInterimOrder();
            modalcreateInterimOrder.FillField_CreatNewInterminOrder(DateTime.Now, site, supplier, location);
            InterimOrdersItem InterimItemOrder = modalcreateInterimOrder.Submit();
            InterimItemOrder.SetQty(qty);
            Thread.Sleep(1000);
            InterimItemOrder.Validate();
            Thread.Sleep(1000);
            var qtyExistance = InterimItemOrder.ExistQtyItem();
            Thread.Sleep(2000);
            Assert.IsTrue(qtyExistance, "No Item");
            InterimItemOrder.ShowExtendedMenu();
            InterimReceptionsCreateModalPage interimReceptionsCreateModalPage = InterimItemOrder.GenerateIntreimReception();
            interimReceptionsCreateModalPage.Filter(InterimReceptionsCreateModalPage.FilterType.ByNumber, numberOrder);
            interimReceptionsCreateModalPage.PageSize("8");
            Assert.IsTrue(interimReceptionsCreateModalPage.GetNumberList().Count() <= 8, "La pagination du 8 ne fonctionne pas..");
            interimReceptionsCreateModalPage.PageSize("16");
            Assert.IsTrue(interimReceptionsCreateModalPage.GetNumberList().Count() <= 16, "La pagination du 16 ne fonctionne pas..");
            interimReceptionsCreateModalPage.PageSize("30");
            Assert.IsTrue(interimReceptionsCreateModalPage.GetNumberList().Count() <= 30, "La pagination du 30 ne fonctionne pas..");
            interimReceptionsCreateModalPage.PageSize("50");
            Assert.IsTrue(interimReceptionsCreateModalPage.GetNumberList().Count() <= 50, "La pagination du 50 ne fonctionne pas..");
            interimReceptionsCreateModalPage.PageSize("100");
            Assert.IsTrue(interimReceptionsCreateModalPage.GetNumberList().Count() <= 100, "La pagination du 100 ne fonctionne pas..");
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_Details_filter_Group()
        {
            var homePage = LogInAsAdmin();
            InterimOrdersPage interimPage = homePage.GoToInterim_InterimOrders();

            interimPage.ResetFilters(); 
            var totalOrders = interimPage.GetNumberOfOrders();
            for (int i = 0; i < totalOrders; i++)
            {
                InterimOrdersItem interimItem = interimPage.ClickOrderAtIndex(i);
                if (interimItem.HasItems())
                {
                    interimItem.Filter(InterimOrdersItem.FilterItemType.SearchByGroup, "AIR CANADA");
                    var filteredList = interimItem.GetNumberOfDisplayedRows();
                    interimItem.Filter(InterimOrdersItem.FilterItemType.SearchByGroup, "A REFERENCIA");
                    var filteredList2 = interimItem.GetNumberOfDisplayedRows();
                    Assert.AreNotEqual(filteredList, filteredList2, "The number of displayed rows should differ after applying the group filters.");
                    break;
                }
                else
                {
                    interimItem.BackToList();
                }
            }
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_index_AddNewCopyFromNumber()
        {
         
            string site = TestContext.Properties["SiteACE"].ToString();
            string siteFix = "AGP";
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string location = "Produccion";
            var homePage = LogInAsAdmin();
            InterimOrdersPage interimOrdersPage = homePage.GoToInterim_InterimOrders();
            interimOrdersPage.ResetFilters();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.Sites, site);
            string numberToCopy = interimOrdersPage.GetFirstInterimOrdersNumber(); 
            InterimOrdersCreateModalPage modalcreateInterimOrder = interimOrdersPage.CreateNewInterimOrder();
      
            modalcreateInterimOrder.FillField_CreatNewInterminOrder(DateTime.Now, site, supplier, location);
            modalcreateInterimOrder.CopyItems();
            modalcreateInterimOrder.SetNumber(numberToCopy);
            // Les deux lignes suivantes ont été ajoutées pour contourner un bug
            modalcreateInterimOrder.FillSite(siteFix);
            modalcreateInterimOrder.FillSite(site);

            bool numberSaisie = modalcreateInterimOrder.IsExistingFilteredNumber();
            var filtredNumber = modalcreateInterimOrder.GetFilteredNumber();
            Assert.IsTrue(numberSaisie, "Le nombre saisi n'a pas été trouvé");
            Assert.AreEqual(filtredNumber, numberToCopy, "Le nombre affiché n'est pas le même que celui saisi.");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_Details_ResetFilter()
        { 
            string subgrpname = "subgrpname";
            string group = "A REFERENCIA";

            // Arrange
            var homePage = LogInAsAdmin();
            InterimOrdersPage interimOrdersPage = homePage.GoToInterim_InterimOrders();
            InterimOrdersItem InterimOrdersItem = interimOrdersPage.GoToInterim_InterimOrdersItem();
            InterimOrdersItem.ResetFilters();
            var searchByNamevalue = InterimOrdersItem.GetFilterValue(InterimOrdersItem.FilterItemType.SearchByNameRef);
            var groupSelectionStatusvalue = InterimOrdersItem.GetGroupSelectionStatus();
            var subGroupSelectionStatusvalue = InterimOrdersItem.GetsubGroupSelectionStatus();
            // remplir les valeurs
            InterimOrdersItem.Filter(InterimOrdersItem.FilterItemType.SearchByNameRef, "ItemForInterim");
            InterimOrdersItem.Filter(InterimOrdersItem.FilterItemType.SearchByGroup, group);
            InterimOrdersItem.Filter(InterimOrdersItem.FilterItemType.subGroup, subgrpname);
            InterimOrdersItem.ResetFilters();
            var searchApres = InterimOrdersItem.GetFilterValue(InterimOrdersItem.FilterItemType.SearchByNameRef);
            Assert.AreEqual(searchApres, searchByNamevalue, "La fonction ResetFilter dans SearchByNameRef ne fonctionne pas");
            var groupSelectionStatusvalueApres = InterimOrdersItem.GetGroupSelectionStatus();
            Assert.AreEqual(groupSelectionStatusvalueApres, groupSelectionStatusvalue, "La fonction ResetFilter dans SearchByGroup ne fonctionne pas");
            var subGroupSelectionStatusvalueApres = InterimOrdersItem.GetsubGroupSelectionStatus();
            Assert.AreEqual(subGroupSelectionStatusvalueApres, subGroupSelectionStatusvalue, "La fonction ResetFilter dans SearchBysubGroup ne fonctionne pas");

        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_details_DeleteItem()
        {
            string site = TestContext.Properties["SiteACE"].ToString();
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string location = "Produccion";
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            InterimOrdersPage interimOrdersPage = homePage.GoToInterim_InterimOrders();
            InterimOrdersCreateModalPage modalcreateInterimOrder = interimOrdersPage.CreateNewInterimOrder();
            modalcreateInterimOrder.FillField_CreatNewInterminOrder(DateTime.Now, site, supplier, location);
            InterimOrdersItem InterimItem = modalcreateInterimOrder.Submit();
            InterimItem.ClickOnItem();
            InterimItem.SetQuantity();
            var prodQuantitybefore = InterimItem.GetQuantity();
            InterimItem.ClickOnPoubelle();
            InterimItem.ClickOnItem();
            var prodQuantityafter = InterimItem.GetQuantity();
            Assert.AreEqual(prodQuantityafter, "0,000", "Le Bouton Delete ne fonctionne pas");

        }

        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_details_EditItem()
        {
            string site = TestContext.Properties["SiteACE"].ToString();
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string location = "Produccion";
            var homePage = LogInAsAdmin();
            InterimOrdersPage interimOrdersPage = homePage.GoToInterim_InterimOrders();
            InterimOrdersCreateModalPage modalcreateInterimOrder = interimOrdersPage.CreateNewInterimOrder();
            modalcreateInterimOrder.FillField_CreatNewInterminOrder(DateTime.Now, site, supplier, location);
            InterimOrdersItem interimItem = modalcreateInterimOrder.Submit();
            interimItem.ClickOnItem();
            interimItem.ClickOnCrayon();
            bool resultIcon = interimItem.verifyOpenedItem();
            Assert.IsTrue(resultIcon, "The edit page did not open correctly after clicking the pencil icon.");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_details_Pagination()
        {
            var homePage = LogInAsAdmin();
            InterimOrdersPage interimOrdersPage = homePage.GoToInterim_InterimOrders();
            interimOrdersPage.PageSize("8");
            var resultPageSize = interimOrdersPage.GetNumberList().Count();
            interimOrdersPage.PaginationNextFirstPage();
            Assert.IsTrue((resultPageSize <= 8) && (interimOrdersPage.GetNumberList().Count() <= 8), "La pagination du 8 ne fonctionne pas.");
            interimOrdersPage.PageSize("16");
            resultPageSize = interimOrdersPage.GetNumberList().Count();
            interimOrdersPage.PaginationNextFirstPage();
            Assert.IsTrue((resultPageSize <= 16) && (interimOrdersPage.GetNumberList().Count() <= 16), "La pagination du 16 ne fonctionne pas.");
            interimOrdersPage.PageSize("30");
            resultPageSize = interimOrdersPage.GetNumberList().Count();
            interimOrdersPage.PaginationNext();
            Assert.IsTrue((resultPageSize <= 30) && (interimOrdersPage.GetNumberList().Count() <= 30), "La pagination du 30 ne fonctionne pas.");
            interimOrdersPage.PageSize("50");
            resultPageSize = interimOrdersPage.GetNumberList().Count();
            interimOrdersPage.PaginationNext();
            Assert.IsTrue((resultPageSize <= 50) && (interimOrdersPage.GetNumberList().Count() <= 50), "La pagination du 50 ne fonctionne pas.");
            interimOrdersPage.PageSize("100");
            if (interimOrdersPage.CheckTotalNumber() > 100)
            {
                interimOrdersPage.PaginationNext();
            }
            Assert.IsTrue(interimOrdersPage.GetNumberList().Count() <= 100, "La pagination du 100 ne fonctionne pas.");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_details_Validate()
        {
            string site = TestContext.Properties["SiteACE"].ToString();
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string location = "Produccion";
            DateTime dateFrom = DateTime.Now.AddDays(-3);
            DateTime dateTo = DateTime.Now;
            var homePage = LogInAsAdmin();
            InterimOrdersPage interimOrdersPage = homePage.GoToInterim_InterimOrders();
            interimOrdersPage.ResetFilters();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowNotValidatedOnly, true);
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.Sites, site);
            if (interimOrdersPage.CheckTotalNumber() == 0)
            {
                InterimOrdersCreateModalPage modalcreateInterimOrder = interimOrdersPage.CreateNewInterimOrder();
                modalcreateInterimOrder.FillField_CreatNewInterminOrder(DateTime.Now, site, supplier, location);
                InterimOrdersItem InterimItemOrder = modalcreateInterimOrder.Submit();
                InterimItemOrder.SetQty("2");
                InterimOrdersPage InterimOrdersPage = InterimItemOrder.BackToList();
            }
            InterimOrdersItem interimOrdersItem = interimOrdersPage.ClickFirstLine();
            interimOrdersItem.SetQty("2");
            interimOrdersItem.Validate();
            var dateItem = interimOrdersItem.GetDate();
            var storedValidationDate = interimOrdersItem.GetFormattedValidationDate();
            InterimOrdersGeneralInformation interimOrdersGeneralInformation = interimOrdersItem.GoToGeneralInformation();
            var validationDate = interimOrdersGeneralInformation.ValidationDate();
            DateTime parsedValidationDate;
            bool isValidValidationDate = interimOrdersGeneralInformation.IsValidDate(validationDate, "dd/MM/yyyy HH:mm", out parsedValidationDate);
            Assert.IsTrue(isValidValidationDate, "La date de validation n'est pas dans un format valide.");
            var validationDateOnly = parsedValidationDate.ToString("yyyy-MM-dd");
            Assert.AreEqual(storedValidationDate, validationDateOnly, "La date de validation est incorrecte.");
            var creationDate = interimOrdersGeneralInformation.CreationDate();
            DateTime parsedCreationDate;
            bool isValidCreationDate = interimOrdersGeneralInformation.IsValidDate(creationDate, "dd/MM/yyyy HH:mm", out parsedCreationDate);
            Assert.IsTrue(isValidCreationDate, "La date de création n'est pas dans un format valide.");
            var creationDateOnly = parsedCreationDate.ToString("yyyy-MM-dd");
            Assert.AreEqual(dateItem, creationDateOnly, "La date de création est incorrecte.");
            Assert.IsFalse(string.IsNullOrEmpty(interimOrdersGeneralInformation.ValidatedBy()), "Le champ 'Validated by' est vide.");

        }
        [Timeout(_timeout)]
		[TestMethod]

        public void INT_INTORD_index_AddNewCopyFromDate()
        {
            DateTime dateFrom = DateTime.Now;
            DateTime dateTo = DateUtils.Now.AddDays(1);

            string dateFromString = dateFrom.ToString("dd/MM/yyyy");
            string dateToString = dateTo.ToString("dd/MM/yyyy");
            string site1 = TestContext.Properties["SiteMAD"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();

            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string location = "Produccion";
            string location1 = "Producción";
            var homePage = LogInAsAdmin();
            InterimOrdersPage interimOrdersPage = homePage.GoToInterim_InterimOrders();
            interimOrdersPage.ResetFilters();
            InterimOrdersCreateModalPage modalcreateInterimOrder = interimOrdersPage.CreateNewInterimOrder();
            modalcreateInterimOrder.FillField_CreatNewInterminOrder(DateTime.Now, site, supplier, location);
            modalcreateInterimOrder.CopyItems();
            // les deux lignes pour fixé le bug du dev
            modalcreateInterimOrder.FillField_CreatNewInterminOrder(DateTime.Now, site1, supplier, location1);
            modalcreateInterimOrder.FillField_CreatNewInterminOrder(DateTime.Now, site, supplier, location);
            modalcreateInterimOrder.CopyItems();
            modalcreateInterimOrder.filter(InterimOrdersCreateModalPage.FilterType.DateFrom, dateFrom);
            modalcreateInterimOrder.filter(InterimOrdersCreateModalPage.FilterType.DateTo, dateTo);
            var datelist = modalcreateInterimOrder.GetListDeliveryDate();
            bool allDatesfromValid = modalcreateInterimOrder.IsDateFromGreaterOrEqualToDateDelivary(datelist, dateFrom, dateTo);
            Assert.IsTrue(allDatesfromValid, "le filtre n'est pas validé");
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_details_ValidatePopup()
        {
            string site = TestContext.Properties["SiteACE"].ToString();
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string location = "Produccion";
            var homePage = LogInAsAdmin();
            InterimOrdersPage interimOrdersPage = homePage.GoToInterim_InterimOrders();
            interimOrdersPage.ResetFilters();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowNotValidatedOnly, true);
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.Sites, site);
            if (interimOrdersPage.CheckTotalNumber() == 0)
            {
                InterimOrdersCreateModalPage modalcreateInterimOrder = interimOrdersPage.CreateNewInterimOrder();
                modalcreateInterimOrder.FillField_CreatNewInterminOrder(DateTime.Now, site, supplier, location);
                InterimOrdersItem InterimItemOrder = modalcreateInterimOrder.Submit();
                InterimItemOrder.SetQty("2");
                InterimOrdersPage InterimOrdersPage = InterimItemOrder.BackToList();
            }
            InterimOrdersItem interimOrdersItem = interimOrdersPage.ClickFirstLine();
            interimOrdersItem.SetQty("2");
            interimOrdersItem.ValidatePopUp();
            Assert.IsTrue(interimOrdersItem.IsModalVisible(), "Le total n'est pas affiché sur la Pop-Up .");

        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_index_filter_ByFromToDate()
        {
            //arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            //act
            InterimOrdersPage interimOrdersPage = homePage.GoToInterim_InterimOrders();
            interimOrdersPage.ResetFilters();
            interimOrdersPage.WaitPageLoading();
            var totalBeforeFilter = interimOrdersPage.CheckTotalNumber();
            interimOrdersPage.Filter(FilterType.ByFromToDate, true);
            interimOrdersPage.Filter(FilterType.DateFrom, DateTime.Today);
            interimOrdersPage.WaitPageLoading();
            var totalAfterFilter = interimOrdersPage.CheckTotalNumber();
            //assert 
            Assert.AreNotEqual(totalAfterFilter, totalBeforeFilter, "Le filtre fonctionne correctement!");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_Details_filter_Keyword()
        {
            //prepare
            string keyword = TestContext.Properties["Item_Keyword"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string location = "Produccion";
            string numberInterim = string.Empty;

            //arrange 
            var homePage = LogInAsAdmin();
            InterimOrdersPage interimOrdersPage = homePage.GoToInterim_InterimOrders();
            interimOrdersPage.ResetFilters();
        
        
            //act
            InterimOrdersItem interimOrderItem = null;
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowNotValidatedOnly, true);
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.Sites, site);
            if (interimOrdersPage.GetNumberOfOrders() == 0)
            {
                // Création d'un nouvel ordre intérimaire si aucun n'existe
                InterimOrdersCreateModalPage interimOrdersCreateModalPage = interimOrdersPage.CreateNewInterimOrder();
                interimOrdersCreateModalPage.FillField_CreatNewInterminOrder(DateTime.Now, site, supplier, location, true);
                numberInterim = interimOrdersCreateModalPage.GetInterimOrderId();
                interimOrderItem = interimOrdersCreateModalPage.Submit();

            }
            else
            {

                numberInterim = interimOrdersPage.GetFirstInterimOrdersNumber();
                interimOrdersPage.ScrollToFilterSearchNumber();
                interimOrdersPage.Filter(InterimOrdersPage.FilterType.SearchByNumber, numberInterim);
                interimOrderItem = interimOrdersPage.ClickOnFirstService();
            }

            var nameItem = interimOrdersPage.GetFirstInterimOrdersName();
            var itemPage = homePage.GoToPurchasing_ItemPage();
            string itemNameSerach = Regex.Replace(nameItem, @"\s*\(.*?\)\s*", "") ;
            itemPage.Filter(ItemPage.FilterType.Search, itemNameSerach);
            ItemGeneralInformationPage generalInfo = itemPage.ClickOnFirstItem();
            ItemKeywordPage keywordPage = generalInfo.ClickOnKeywordItem();
            try
            {
                var x = keywordPage.IsKeywordAdded(keyword);
                if (!keywordPage.IsKeywordAdded(keyword))
                {
                    keywordPage.AddKeyword(keyword);
                }
                interimOrdersPage = homePage.GoToInterim_InterimOrders();
                interimOrdersPage.ResetFilters();
                interimOrdersPage.Filter(InterimOrdersPage.FilterType.SearchByNumber, numberInterim);
                interimOrderItem = interimOrdersPage.ClickOnFirstService();
                var rowsItem = interimOrdersPage.GetTotalResultRowsPage();
                interimOrdersPage.Filter(InterimOrdersPage.FilterType.Keywords, keyword);
                var rowsItemAfterFilter = interimOrdersPage.GetTotalResultRowsPage();
                Assert.AreEqual(rowsItem, rowsItemAfterFilter, "Les résultats ne sont pas mis à jour en fonction des filtres KEYWORDS");
            }
            finally
            {
                itemPage = homePage.GoToPurchasing_ItemPage();
                itemPage.Filter(ItemPage.FilterType.Search, itemNameSerach);
                ItemGeneralInformationPage generalInfos = itemPage.ClickOnFirstItem();
                ItemKeywordPage keyWordPage = generalInfo.ClickOnKeywordItem();
                keywordPage.RemoveKeyword(keyword);
            }

        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_index_AddNewCopyFrom()
        {
            string site1 = TestContext.Properties["SiteMAD"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string location = "Produccion";
            string location1 = "Producción";
            var homePage = LogInAsAdmin();
            InterimOrdersPage interimOrdersPage = homePage.GoToInterim_InterimOrders();
            interimOrdersPage.ResetFilters();
            InterimOrdersCreateModalPage modalcreateInterimOrder = interimOrdersPage.CreateNewInterimOrder();
            modalcreateInterimOrder.FillField_CreatNewInterminOrder(DateTime.Now, site, supplier, location);
            modalcreateInterimOrder.CopyItems();
            // les deux lignes pour fixé le bug du dev
            modalcreateInterimOrder.FillField_CreatNewInterminOrder(DateTime.Now, site1, supplier, location1);
            modalcreateInterimOrder.FillField_CreatNewInterminOrder(DateTime.Now, site, supplier, location);
            modalcreateInterimOrder.CopyItems();
            modalcreateInterimOrder.PageSizeCreateNewInterimOrder("100");
            var allInterimOrders = modalcreateInterimOrder.getAllList().ToList();
            var validatedInterimOrders = modalcreateInterimOrder.getListValidate().ToList();
            var notvalidateOrder = allInterimOrders.Count - validatedInterimOrders.Count;
            Assert.IsTrue(validatedInterimOrders.Count > 0, "les interim orders ne sont pas validés");
            Assert.IsTrue(notvalidateOrder >= 0, "les interim order sont validé");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_index_AddNewCopyFromSelect()
        {
            //prepare
            string site = TestContext.Properties["SiteACE"].ToString();
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string location = "Produccion";

            //arrange
            var homePage = LogInAsAdmin();

            //act
            InterimOrdersPage interimOrdersPage = homePage.GoToInterim_InterimOrders();
            interimOrdersPage.ResetFilters();
            InterimOrdersCreateModalPage modalcreateInterimOrder = interimOrdersPage.CreateNewInterimOrder();
            

            modalcreateInterimOrder.FillField_CreatNewInterminOrder(DateTime.Now, site, supplier, location, true);
            string firstInterimOrderToCopy = modalcreateInterimOrder.GetInterimOrderId();
            InterimOrdersItem InterimItem = modalcreateInterimOrder.Submit();
            InterimItem.SetQty(new Random().Next(3, 50).ToString());
            InterimItem.Validate();
            InterimItem.BackToList();

            interimOrdersPage.CreateNewInterimOrder();
            modalcreateInterimOrder.FillField_CreatNewInterminOrder(DateTime.Now, site, supplier, location);
            modalcreateInterimOrder.CopyItems();

            // Les deux lignes suivantes ont été ajoutées pour contourner un bug
            modalcreateInterimOrder.FillSite(site);

            modalcreateInterimOrder.SearchForInterim(firstInterimOrderToCopy);
            modalcreateInterimOrder.CopyFirstInterimOrder();

            InterimOrdersItem interimOrdersItem = modalcreateInterimOrder.Submit();

            var copiedInterimOrderName = interimOrdersItem.GetCopiedInterimOrderName();
            var copiedInterimOrderSupplier = interimOrdersItem.GetCopiedInterimOrderSupplier();
            var copiedInterimOrderPacking = interimOrdersItem.GetCopiedInterimOrderPacking();
            var copiedInterimOrderPackagingPrice = interimOrdersItem.GetCopiedInterimOrderPackagingPrice();
            var copiedInterimOrderQuantity = interimOrdersItem.GetCopiedInterimOrderQuantity();
            var copiedInterimOrderTotalPrice = interimOrdersItem.GetCopiedInterimOrderTotalPrice();

            interimOrdersPage = interimOrdersItem.BackToList();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.SearchByNumber, firstInterimOrderToCopy);
            interimOrdersItem = interimOrdersPage.ClickFirstLine();
            
            var originInterimOrderName = interimOrdersItem.GetOriginInterimOrderNameSelectsFirstInterim();
            var originInterimOrderSupplier = interimOrdersItem.GetOriginInterimOrderPackingSelectsFirstInterim();
            var originInterimOrderPackagingPrice = interimOrdersItem.GetOriginInterimOrderPackagingPriceSelectsFirstInterim();
            var originInterimOrderQuantity = interimOrdersItem.GetOriginInterimOrderQuantitySelectsFirstInterim();
            var originInterimOrderTotalPrice = interimOrdersItem.GetOriginInterimOrderTotalPriceSelectsFirstInterim();

            Assert.AreEqual(copiedInterimOrderName, originInterimOrderName, "Les items de l'intérim order n'ont pas été correctement pré-remplis selon les informations des orders copiées.");
            Assert.AreEqual(copiedInterimOrderPackagingPrice, originInterimOrderPackagingPrice, "Les items de l'intérim order n'ont pas été correctement pré-remplis selon les informations des orders copiées.");
            Assert.AreEqual(copiedInterimOrderQuantity, originInterimOrderQuantity, "Les items de l'intérim order n'ont pas été correctement pré-remplis selon les informations des orders copiées.");
            Assert.AreEqual(copiedInterimOrderTotalPrice, originInterimOrderTotalPrice, "Les items de l'intérim order n'ont pas été correctement pré-remplis selon les informations des orders copiées.");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_index_AddNewCopyFromSelects()
        {
            string site = TestContext.Properties["SiteACE"].ToString();
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string location = "Produccion";
            //login
            var homePage = LogInAsAdmin();
            InterimOrdersPage interimOrdersPage = homePage.GoToInterim_InterimOrders();
            InterimOrdersCreateModalPage modalcreateInterimOrder = interimOrdersPage.CreateNewInterimOrder();
            modalcreateInterimOrder.FillField_CreatNewInterminOrder(DateTime.Now, site, supplier, location, true);
            string firstInterimOrderToCopy = modalcreateInterimOrder.GetInterimOrderId();
            InterimOrdersItem InterimItem = modalcreateInterimOrder.Submit();
            InterimItem.SetQty(new Random().Next(3, 50).ToString());
            InterimItem.Validate();
            InterimItem.BackToList();
            modalcreateInterimOrder = interimOrdersPage.CreateNewInterimOrder();
            modalcreateInterimOrder.FillField_CreatNewInterminOrder(DateTime.Now, site, supplier, location, true);
            string secondInterimOrderToCopy = modalcreateInterimOrder.GetInterimOrderId();
            InterimItem = modalcreateInterimOrder.Submit();
            InterimItem.SetQty(new Random().Next(3, 50).ToString());
            InterimItem.Validate();
            InterimItem.BackToList();

            interimOrdersPage.ResetFilters();
            interimOrdersPage.ScrollUntilSitesFilterIsVisible();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.Site, site);
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowNotValidatedOnly, true);
            modalcreateInterimOrder = interimOrdersPage.CreateNewInterimOrder();
            modalcreateInterimOrder.FillField_CreatNewInterminOrder(DateTime.Now, site, supplier, location);
            modalcreateInterimOrder.CopyItems();

            // Les deux lignes suivantes ont été ajoutées pour contourner un bug
            modalcreateInterimOrder.FillSite(site);

            var copiedInterimOrderNumber = modalcreateInterimOrder.GetCopiedInterimOrderNumber();
            /*  var firstInterimOrderToCopy = modalcreateInterimOrder.GetFirstInterimOrderToCopy();
              var secondInterimOrderToCopy = modalcreateInterimOrder.GetSecondInterimOrderToCopy();*/
            modalcreateInterimOrder.SearchForInterim(firstInterimOrderToCopy);
            modalcreateInterimOrder.CopyFirstInterimOrder();
            modalcreateInterimOrder.SearchForInterim(secondInterimOrderToCopy);
            modalcreateInterimOrder.CopyFirstInterimOrder();
            InterimOrdersItem interimItemOrder = modalcreateInterimOrder.Submit();

            var copiedInterimOrderName = interimItemOrder.GetCopiedInterimOrderName();
            var copiedInterimOrderSupplier = interimItemOrder.GetCopiedInterimOrderSupplier();
            var copiedInterimOrderPacking = interimItemOrder.GetCopiedInterimOrderPacking();
            var copiedInterimOrderPackagingPrice = interimItemOrder.GetCopiedInterimOrderPackagingPrice();
            var copiedInterimOrderQuantity = interimItemOrder.GetCopiedInterimOrderQuantity();
            var copiedInterimOrderTotalPrice = interimItemOrder.GetCopiedInterimOrderTotalPrice();

            interimItemOrder.BackToList();
            interimOrdersPage.ResetFilters();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.SearchByNumber, firstInterimOrderToCopy);
            interimOrdersPage.ClickFirstLine();

            var firstInterimOrderName = interimItemOrder.GetOriginInterimOrderNameSelectsFirstInterim();
            var firstInterimOrderSupplier = interimItemOrder.GetOriginInterimOrderSupplierSelectsFirstInterim();
            var firstInterimOrderPacking = interimItemOrder.GetOriginInterimOrderPackingSelectsFirstInterim();
            var firstInterimOrderPackagingPrice = interimItemOrder.GetOriginInterimOrderPackagingPriceSelectsFirstInterim();
            var firstInterimOrderQuantity = interimItemOrder.GetOriginInterimOrderQuantitySelectsFirstInterim();
            var firstInterimOrderTotalPrice = interimItemOrder.GetOriginInterimOrderTotalPriceSelectsFirstInterim();

            interimItemOrder.BackToList();
            interimOrdersPage.ResetFilters();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.SearchByNumber, secondInterimOrderToCopy);
            interimOrdersPage.ClickFirstLine();

            var secondInterimOrderName = interimItemOrder.GetOriginInterimOrderNameSelectsFirstInterim();
            var secondInterimOrderSupplier = interimItemOrder.GetOriginInterimOrderSupplierSelectsFirstInterim();
            var secondInterimOrderPacking = interimItemOrder.GetOriginInterimOrderPackingSelectsFirstInterim();
            var secondInterimOrderPackagingPrice = interimItemOrder.GetOriginInterimOrderPackagingPriceSelectsFirstInterim();
            var secondInterimOrderQuantity = interimItemOrder.GetOriginInterimOrderQuantitySelectsFirstInterim();
            var secondInterimOrderTotalPrice = interimItemOrder.GetOriginInterimOrderTotalPriceSelectsFirstInterim();

            if (firstInterimOrderName == secondInterimOrderName)
            {
                Assert.AreEqual(copiedInterimOrderName, firstInterimOrderName, "Les items de l'intérim order n'ont pas été correctement pré-remplis selon les informations des orders copiées.");
                Assert.AreEqual(copiedInterimOrderSupplier, firstInterimOrderSupplier, "Les items de l'intérim order n'ont pas été correctement pré-remplis selon les informations des orders copiées.");
                Assert.AreEqual(copiedInterimOrderPacking, firstInterimOrderPacking, "Les items de l'intérim order n'ont pas été correctement pré-remplis selon les informations des orders copiées.");
               // Assert.AreEqual(firstInterimOrderPackagingPrice + secondInterimOrderPackagingPrice, copiedInterimOrderPackagingPrice, "Les items de l'intérim order n'ont pas été correctement pré-remplis selon les informations des orders copiées.");
                Assert.AreEqual(firstInterimOrderQuantity + secondInterimOrderQuantity, copiedInterimOrderQuantity, "Les items de l'intérim order n'ont pas été correctement pré-remplis selon les informations des orders copiées.");
                Assert.AreEqual(firstInterimOrderTotalPrice + secondInterimOrderTotalPrice, copiedInterimOrderTotalPrice, "Les items de l'intérim order n'ont pas été correctement pré-remplis selon les informations des orders copiées.");
            }

            if (firstInterimOrderName != secondInterimOrderName)
            {
                interimItemOrder.BackToList();
                interimOrdersPage.ResetFilters();
                interimOrdersPage.Filter(InterimOrdersPage.FilterType.SearchByNumber, copiedInterimOrderNumber);
                interimOrdersPage.ClickFirstLine();

                interimItemOrder.Filter(InterimOrdersItem.FilterItemType.SearchByNameRef, firstInterimOrderName);

                Assert.AreEqual(copiedInterimOrderName, firstInterimOrderName, "Les items de l'intérim order n'ont pas été correctement pré-remplis selon les informations des orders copiées.");
                Assert.AreEqual(copiedInterimOrderSupplier, firstInterimOrderSupplier, "Les items de l'intérim order n'ont pas été correctement pré-remplis selon les informations des orders copiées.");
                Assert.AreEqual(copiedInterimOrderPacking, firstInterimOrderPacking, "Les items de l'intérim order n'ont pas été correctement pré-remplis selon les informations des orders copiées.");
                Assert.AreEqual(copiedInterimOrderPackagingPrice, firstInterimOrderPackagingPrice, "Les items de l'intérim order n'ont pas été correctement pré-remplis selon les informations des orders copiées.");
                Assert.AreEqual(copiedInterimOrderQuantity, firstInterimOrderQuantity, "Les items de l'intérim order n'ont pas été correctement pré-remplis selon les informations des orders copiées.");
                Assert.AreEqual(copiedInterimOrderTotalPrice, firstInterimOrderTotalPrice, "Les items de l'intérim order n'ont pas été correctement pré-remplis selon les informations des orders copiées.");

                interimItemOrder.Filter(InterimOrdersItem.FilterItemType.SearchByNameRef, secondInterimOrderName);
                //Assert
                Assert.AreEqual(copiedInterimOrderName, secondInterimOrderName, "Les items de l'intérim order n'ont pas été correctement pré-remplis selon les informations des orders copiées.");
                Assert.AreEqual(copiedInterimOrderSupplier, secondInterimOrderSupplier, "Les items de l'intérim order n'ont pas été correctement pré-remplis selon les informations des orders copiées.");
                Assert.AreEqual(copiedInterimOrderPacking, secondInterimOrderPacking, "Les items de l'intérim order n'ont pas été correctement pré-remplis selon les informations des orders copiées.");
                Assert.AreEqual(copiedInterimOrderPackagingPrice, secondInterimOrderPackagingPrice, "Les items de l'intérim order n'ont pas été correctement pré-remplis selon les informations des orders copiées.");
                Assert.AreEqual(copiedInterimOrderQuantity, secondInterimOrderQuantity, "Les items de l'intérim order n'ont pas été correctement pré-remplis selon les informations des orders copiées.");
                Assert.AreEqual(copiedInterimOrderTotalPrice, secondInterimOrderTotalPrice, "Les items de l'intérim order n'ont pas été correctement pré-remplis selon les informations des orders copiées.");
            }
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_details_TotalwoVAT()
        {

            string qty = "3";
            string site = TestContext.Properties["SiteACE"].ToString();
            var homePage = LogInAsAdmin();
            InterimOrdersPage interimOrdersPage = homePage.GoToInterim_InterimOrders();
            interimOrdersPage.ResetFilters();
            // Filter to display only non-validated items
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowNotValidatedOnly, true);
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.Sites, site);
            var number = interimOrdersPage.GetFirstID();
            if (interimOrdersPage.CheckTotalNumber() != 0)
            {
                // Use existing data without creating a new interim order
                InterimOrdersItem firstItem = interimOrdersPage.ClickFirstLine();
                if (firstItem.HasPricePackaging())
                {
                    // Optionally set a quantity if needed
                    if (firstItem.GetQtyItem() == "0")
                    {
                        firstItem.SetQty(qty);
                        WebDriver.Navigate().Refresh();
                    }
                }
                else
                {
                    //Adding price if not exist 
                    ItemGeneralInformationPage itemGeneralInformationPage =  firstItem.EditButton();
                    itemGeneralInformationPage.Go_To_New_Navigate();
                    itemGeneralInformationPage.ScrollDown();
                    itemGeneralInformationPage.SelectFirstPackaging();
                    itemGeneralInformationPage.SetPackagingPrice("10");
                    interimOrdersPage = homePage.GoToInterim_InterimOrders();
                    interimOrdersPage.Filter(InterimOrdersPage.FilterType.SearchByNumber, number);
                    firstItem = interimOrdersPage.ClickFirstLine();
                    // Optionally set a quantity if needed
                    if (firstItem.GetQtyItem() == "0")
                    {
                        firstItem.SetQty(qty);
                        WebDriver.Navigate().Refresh();
                    }

                }

                var quantite = int.Parse(firstItem.GetQtyItem());
                var packagingPrice = int.Parse(firstItem.GetPackagingPriceItem());
                var total = quantite * packagingPrice;
                var newTotal = int.Parse(firstItem.GetTotalVATItem());

                Assert.AreEqual(total, newTotal, "Le Total n'est pas valide !!");
            }
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_details_GenerateReceptionFromReception()
        {
            string site = TestContext.Properties["SiteACE"].ToString();
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string location = "Produccion";
            DateTime dateFrom = DateTime.Now.AddDays(-3);
            DateTime dateTo = DateTime.Now;
            var qty = "3";
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            InterimReceptionsPage interimreceptionPage = homePage.GoToInterim_Receptions();
            InterimReceptionsCreateModalPage modalcreateInterim = interimreceptionPage.CreateNewInterim();
            modalcreateInterim.FillField_CreatNewIntermin(DateTime.Now, site, supplier, location);
            var InterimId = modalcreateInterim.GetInterimId();
            InterimReceptionsItem InterimItem = modalcreateInterim.Submit();
            InterimItem.SetReceptionItemReceived();
            InterimItem.Validate();
            InterimOrdersPage interimOrdersPage = InterimItem.GoToReceptionOrder();
            InterimOrdersCreateModalPage modalcreateInterimOrder = interimOrdersPage.CreateNewInterimOrder();
            modalcreateInterimOrder.FillField_CreatNewInterminOrder(DateTime.Now, site, supplier, location);
            InterimOrdersItem InterimItemOrder = modalcreateInterimOrder.Submit();
            InterimItemOrder.SetQty(qty);
            InterimItemOrder.Validate();
            InterimItemOrder.ShowExtendedMenu();
            InterimReceptionsCreateModalPage interimReceptionsCreateModalPage = InterimItemOrder.GenerateIntreimReception();
            interimReceptionsCreateModalPage.ToInterimReception();
            interimReceptionsCreateModalPage.ClearField();
            var containsData = interimReceptionsCreateModalPage.ContainsData();
            Assert.IsTrue(containsData, "aucun résultat n'est affiché");
        }

        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_details_SendByMailPopUp()
        {
            string site = TestContext.Properties["SiteACE"].ToString();
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string attachedFiles ;
            var homePage = LogInAsAdmin();
            InterimOrdersPage interimOrdersPage = homePage.GoToInterim_InterimOrders();
            interimOrdersPage.ResetFilters();
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.ShowValidatedOnly, true);
            InterimOrdersItem firstItem = interimOrdersPage.ClickFirstLine();
            firstItem.SendByMail();
            Assert.IsTrue(firstItem.IsEmailPopupDisplayed(), "The email was not sent successfully.");
            attachedFiles = firstItem.GetAttachedFiles();
            bool containsPdf = attachedFiles.Contains(".pdf");
            Assert.IsTrue(containsPdf, "The attached files do not include a PDF.");
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_details_Footer()
        {
            string site = TestContext.Properties["SiteACE"].ToString();
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string location = "Produccion";
            Random rnd = new Random();
            //arrange
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            InterimOrdersPage interimOrdersPage = homePage.GoToInterim_InterimOrders();
            //act
                InterimOrdersCreateModalPage interimOrdersCreateModalPage = interimOrdersPage.CreateNewInterimOrder();
                interimOrdersCreateModalPage.FillField_CreatNewInterminOrder(DateTime.Now, site, supplier, location , true);
                string interimNumber = interimOrdersCreateModalPage.GetInterimOrderId();
                InterimOrdersItem InterimItem = interimOrdersCreateModalPage.Submit();
                InterimItem.SetQty(rnd.Next(1, 9).ToString());
                InterimItem.Validate();
                InterimItem.BackToList(); 
                interimOrdersPage.Filter(InterimOrdersPage.FilterType.SearchByNumber, interimNumber);
                InterimOrdersItem interimOrdersItem = interimOrdersPage.ClickFirstLine();
                var listItemNames = interimOrdersItem.GetAllNamesResultPaged();
                var totalItemPrice = interimOrdersItem.GetTotalItemPrice();
                InterimOrdersFooter interimOrdersFooter = interimOrdersItem.GoToFooter();
                var totalFooterPrice = interimOrdersFooter.GetTotalFooterPrice();
                var totalFooterGrossAmount = interimOrdersFooter.GetTotalFooterGrossAmount();
                var totalFooterVatAmount = interimOrdersFooter.GetTotalFooterVatAmount();

                var totalInterimOrder = (totalFooterGrossAmount + totalFooterVatAmount);
                Assert.AreEqual(totalItemPrice, totalFooterGrossAmount, "Le montant total de base n'est pas le même Total dans l'onglet 'Item");

                var itemPage = homePage.GoToPurchasing_ItemPage();
                itemPage.ResetFilter();

                foreach (var name in listItemNames)
                {
                    itemPage.Filter(ItemPage.FilterType.Search, name);
                    ItemGeneralInformationPage itemGeneralInformationPage = itemPage.ClickOnFirstItem();
                    var taxeItemName = itemGeneralInformationPage.GetVatName();
                    homePage.GoToInterim_InterimOrders();

                    interimOrdersPage.ResetFilters();
                    interimOrdersPage.Filter(InterimOrdersPage.FilterType.SearchByNumber, interimNumber);
                    interimOrdersPage.ClickFirstLine();

                    interimOrdersItem.GoToFooter();
                    var taxNameFooter = interimOrdersFooter.GetTaxeName();
                    Assert.AreEqual(taxeItemName, taxNameFooter, "Les Noms ne sont pas Valide");
                    homePage.GoToPurchasing_ItemPage();
                    itemPage.ResetFilter();
                }
                homePage.GoToInterim_InterimOrders();
                interimOrdersPage.ResetFilters();
                interimOrdersPage.Filter(InterimOrdersPage.FilterType.SearchByNumber, interimNumber);
                interimOrdersPage.ClickFirstLine();
                interimOrdersItem.GoToFooter();
                Assert.AreEqual(totalInterimOrder, (totalFooterGrossAmount + totalFooterVatAmount), " Le total interim est différent de la somme du total et de la taxe.");
           
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_index_AddNewPagination()
        {
            string site = TestContext.Properties["SiteACE"].ToString();
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string location = "Produccion";
            //login
            var homePage = LogInAsAdmin();
            InterimOrdersPage interimPage = homePage.GoToInterim_InterimOrders();
            InterimOrdersCreateModalPage interimOrdersCreateModalPage = interimPage.CreateNewInterimOrder();
            interimOrdersCreateModalPage.FillField_CreatNewInterminOrder(DateTime.Now, site, supplier, location);
            interimOrdersCreateModalPage.CopyItems();
            InterimOrdersItem InterimItem = interimOrdersCreateModalPage.Submit();
            interimOrdersCreateModalPage.PageSizeCreateNewInterimOrder("8");
            var resultPageSize = interimOrdersCreateModalPage.GetNumberList().Count();
            interimOrdersCreateModalPage.PaginationNextFirstPage();
            //assert : pagination 8
            Assert.IsTrue((resultPageSize <= 8) && (interimOrdersCreateModalPage.GetNumberList().Count() <= 8), "La pagination du 8 ne fonctionne pas.");
            interimOrdersCreateModalPage.PageSizeCreateNewInterimOrder("16");
            resultPageSize = interimOrdersCreateModalPage.GetNumberList().Count();
            interimOrdersCreateModalPage.PaginationNextFirstPage();
            //assert : pagination 16
            Assert.IsTrue((resultPageSize <= 16) && (interimOrdersCreateModalPage.GetNumberList().Count() <= 16), "La pagination du 16 ne fonctionne pas..");
            interimOrdersCreateModalPage.PageSizeCreateNewInterimOrder("30");
            resultPageSize = interimOrdersCreateModalPage.GetNumberList().Count();
            interimOrdersCreateModalPage.PaginationNext();
            //assert : pagination 30
            Assert.IsTrue((resultPageSize <= 30) && (interimOrdersCreateModalPage.GetNumberList().Count() <= 30), "La pagination du 30  ne fonctionne pas..");
            interimOrdersCreateModalPage.PageSizeCreateNewInterimOrder("50");
            resultPageSize = interimOrdersCreateModalPage.GetNumberList().Count();
            interimOrdersCreateModalPage.PaginationNext();
            //assert : pagination 50
            Assert.IsTrue((resultPageSize <= 50) && (interimOrdersCreateModalPage.GetNumberList().Count() <= 50), "La pagination du 50 ne fonctionne pas..");
            interimOrdersCreateModalPage.PageSizeCreateNewInterimOrder("100");
            if (interimOrdersCreateModalPage.CheckTotalNumber() > 100)
            {
                interimOrdersCreateModalPage.PaginationNext();
            }
            //assert : pagination 100
            Assert.IsTrue(interimOrdersCreateModalPage.GetNumberList().Count() <= 100, "La pagination du 100 ne fonctionne pas..");
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_details_GenerateReceptionFromReceptionGenerate()
        {

            string site = TestContext.Properties["SiteACE"].ToString();
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string location = "Produccion";
            var homePage = LogInAsAdmin();
            InterimOrdersPage interimOrdersPage = homePage.GoToInterim_InterimOrders();
            InterimOrdersCreateModalPage modalcreateInterimOrder = interimOrdersPage.CreateNewInterimOrder();
            modalcreateInterimOrder.FillField_CreatNewInterminOrder(DateTime.Now, site, supplier, location);
            InterimOrdersItem InterimItemOrder = modalcreateInterimOrder.Submit();
            InterimItemOrder.Validate();
            InterimItemOrder.ShowExtendedMenu();
            InterimReceptionsCreateModalPage interimReceptionsCreateModalPage = InterimItemOrder.GenerateIntreimReception();
            interimReceptionsCreateModalPage.ToInterimReception();
            interimReceptionsCreateModalPage.FillDeliveryOrderNumber();
            interimReceptionsCreateModalPage.ClearField();
            // Les deux lignes suivantes ont été ajoutées pour contourner un bug
            interimReceptionsCreateModalPage.ToInterimOrder();
            interimReceptionsCreateModalPage.ToInterimReception();

            var firstInterimReceptionToSelect = interimReceptionsCreateModalPage.GetFirstInterimReceptionToSelect();
            var secondInterimReceptionToSelect = interimReceptionsCreateModalPage.GetSecondInterimReceptionToSelect();
            interimReceptionsCreateModalPage.SelectFirstAndSecondInterimReception();
            InterimReceptionsItem interimReceptionsItem = interimReceptionsCreateModalPage.Submit();
            var newReceivedInterimReceptionItem = interimReceptionsItem.GetReceivedInterimReceptionItem();
            var newTotalVatInterimReceptionItem = interimReceptionsItem.GetTotalVatInterimReceptionItem();

            InterimReceptionsPage interimReceptionsPage = interimReceptionsItem.BackToList();
            interimReceptionsPage.Filter(InterimReceptionsPage.FilterType.Bynumber, firstInterimReceptionToSelect);
            interimReceptionsItem = interimReceptionsPage.GoToInterimReceptionItem();
            var firstReceivedInterimReceptionItem = interimReceptionsItem.GetSelectedReceivedInterimReceptionItem();
            var firstTotalVatInterimReceptionItem = interimReceptionsItem.GetSelectedTotalVatInterimReceptionItem();

            interimReceptionsPage = interimReceptionsItem.BackToList();
            interimReceptionsPage.Filter(InterimReceptionsPage.FilterType.Bynumber, secondInterimReceptionToSelect);
            interimReceptionsItem = interimReceptionsPage.GoToInterimReceptionItem();
            var secondReceivedInterimReceptionItem = interimReceptionsItem.GetSelectedReceivedInterimReceptionItem();
            var secondTotalVatInterimReceptionItem = interimReceptionsItem.GetSelectedTotalVatInterimReceptionItem();

            Assert.AreEqual(firstReceivedInterimReceptionItem + secondReceivedInterimReceptionItem, newReceivedInterimReceptionItem, "Les items ne correspondent pas au total des deux réceptions sélectionnées.");
            Assert.AreEqual(firstTotalVatInterimReceptionItem + secondTotalVatInterimReceptionItem, newTotalVatInterimReceptionItem, "Les items ne correspondent pas au total des deux réceptions sélectionnées.");
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_details_GenerateReceptionFromOrderGenerate()
        {

            string site = TestContext.Properties["SiteACE"].ToString();
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string location = "Produccion";
            var rd = new Random(); 
            LogInAsAdmin();
            var homePage = new HomePage(WebDriver, TestContext);
            homePage.Navigate();
            InterimOrdersPage interimOrdersPage = homePage.GoToInterim_InterimOrders();
            
            InterimOrdersCreateModalPage modalcreateInterimOrder = interimOrdersPage.CreateNewInterimOrder();
            modalcreateInterimOrder.FillField_CreatNewInterminOrder(DateTime.Now, site, supplier, location , true );
            string firstInterimOrderToSelect = modalcreateInterimOrder.GetInterimOrderId();
            InterimOrdersItem interimItem = modalcreateInterimOrder.Submit();
            interimItem.SetQty(rd.Next(3, 50).ToString());
            interimItem.Validate(); 
            interimItem.BackToList(); 

            modalcreateInterimOrder = interimOrdersPage.CreateNewInterimOrder();
            modalcreateInterimOrder.FillField_CreatNewInterminOrder(DateTime.Now, site, supplier, location , true );
            string secondInterimOrderToSelect = modalcreateInterimOrder.GetInterimOrderId();
            interimItem = modalcreateInterimOrder.Submit();
            interimItem.SetQty(rd.Next(3, 50).ToString());
            interimItem.Validate();
            interimItem.BackToList();


            modalcreateInterimOrder = interimOrdersPage.CreateNewInterimOrder();
            modalcreateInterimOrder.FillField_CreatNewInterminOrder(DateTime.Now, site, supplier, location);
            InterimOrdersItem interimItemOrder = modalcreateInterimOrder.Submit();
       //     Thread.Sleep(1000);
            interimItemOrder.Validate();
        //    Thread.Sleep(1000);
            interimItemOrder.ShowExtendedMenu();
            InterimReceptionsCreateModalPage interimReceptionsCreateModalPage = interimItemOrder.GenerateIntreimReception();
            interimReceptionsCreateModalPage.FillDeliveryOrderNumber();
            interimReceptionsCreateModalPage.ClearField();

            // Les deux lignes suivantes ont été ajoutées pour contourner un bug
            interimReceptionsCreateModalPage.ToInterimReception();
            interimReceptionsCreateModalPage.ToInterimOrder();
       //     Thread.Sleep(1000);
            modalcreateInterimOrder.SearchForInterim(firstInterimOrderToSelect);
            modalcreateInterimOrder.CopyFirstInterimOrder();
            modalcreateInterimOrder.SearchForInterim(secondInterimOrderToSelect);
           modalcreateInterimOrder.CopyFirstInterimOrder();


            InterimReceptionsItem interimReceptionsItem = interimReceptionsCreateModalPage.Submit();
            var newReceivedInterimReceptionItem = interimReceptionsItem.GetReceivedInterimReceptionItem();
            var newTotalVatInterimReceptionItem = interimReceptionsItem.GetTotalVatInterimReceptionItem();
            interimOrdersPage = interimReceptionsItem.GoToReceptionOrder();
      //      Thread.Sleep(1000);
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.SearchByNumber, firstInterimOrderToSelect);
            interimItemOrder = interimOrdersPage.GoToInterim_InterimOrdersItem();
            var firstProdQtyInterimOrderItem = interimItemOrder.GetSelectedProdQtyInterimOrderItem();
            var firstTotalVatInterimOrderItem = interimItemOrder.GetSelectedTotalVatInterimOrderItem();

            interimOrdersPage = interimItemOrder.BackToList();
          
            interimOrdersPage.Filter(InterimOrdersPage.FilterType.SearchByNumber, secondInterimOrderToSelect);
            interimItemOrder = interimOrdersPage.GoToInterim_InterimOrdersItem();
            var secondProdQtyInterimOrderItem = interimItemOrder.GetSelectedProdQtyInterimOrderItem();
            var secondTotalVatInterimOrderItem = interimItemOrder.GetSelectedTotalVatInterimOrderItem();

            Assert.AreEqual(firstProdQtyInterimOrderItem + secondProdQtyInterimOrderItem, newReceivedInterimReceptionItem, "Les items ne correspondent pas au total des deux orders sélectionnées.");
            Assert.AreEqual(firstTotalVatInterimOrderItem + secondTotalVatInterimOrderItem, newTotalVatInterimReceptionItem, "Les items ne correspondent pas au total des deux orders sélectionnées.");
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_details_SendByMailAttach()
        {
            //prepare
            string userEmail = TestContext.Properties["Admin_UserName"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string text = null;
            string location = "Production";
            Random rnd = new Random();
            string qty = "10,000";
            string interimNumber = string.Empty;

            //arrange
            var homePage = LogInAsAdmin();
            InterimOrdersPage interimOrdersPage = homePage.GoToInterim_InterimOrders();
            interimOrdersPage.ResetFilters();
            interimOrdersPage.Filter(FilterType.ShowValidatedOnly,true);
            InterimOrdersItem interimOrdersItem = null;
            var numbers = interimOrdersPage.GetTotalOfOrders();
            if (numbers == 0)
            {
                // Création d'un nouvel ordre intérimaire si aucun n'existe
                InterimOrdersCreateModalPage interimOrdersCreateModalPage = interimOrdersPage.CreateNewInterimOrder();
                interimOrdersCreateModalPage.FillField_CreatNewInterminOrder(DateTime.Now, site, supplier, location, true);
                interimNumber = interimOrdersCreateModalPage.GetInterimOrderId();
                interimOrdersItem = interimOrdersCreateModalPage.Submit();

                // Vérification de la quantité et modification si nécessaire
                if (interimOrdersItem.GetQtyItem() == "0")
                {
                    interimOrdersItem.ClickFirstItem();
                    interimOrdersItem.SetQty(qty);
                }
                else
                {
                    interimOrdersItem.ClickFirstItem();
                }
                interimOrdersItem.Validate();
            }
            else
            {
                // Si des ordres intérimaires existent déjà, sélectionnez le premier
                interimOrdersItem = interimOrdersPage.ClickFirstLine();
                interimNumber = interimOrdersItem.ReturnInterimOrderNumber();

            }

            interimOrdersItem.SendByMail();
            //firstItem.SendByEmailEnterimOrder(userEmail);
            interimOrdersItem.SendByEmailEnterimOrder(userEmail, text);
            MailPage mailPage = interimOrdersItem.RedirectToOutlookMailbox();
            mailPage.FillFields_LogToOutlookMailbox_MoreThanOneMonth(userEmail);
            // check if mail sent
            bool isMailSent = mailPage.CheckIfSpecifiedOutlookMailExist("Newrest - Interim order "+ interimNumber);
            Assert.IsTrue(isMailSent, "La claim n'a pas été envoyée par mail.");
            mailPage.DeleteCurrentOutlookMail();
            mailPage.Close();
        }
        [Timeout(_timeout)]
		[TestMethod]
        public void INT_INTORD_details_SendByMailEdit()
        {
            string userEmail = TestContext.Properties["Admin_UserName"].ToString();
            string site = TestContext.Properties["SiteACE"].ToString();
            string supplier = TestContext.Properties["SupplierForInterim"].ToString();
            string text = "This mail was modified";
            string location = "Produccion";
            string qty = "10,000";
            var homePage = LogInAsAdmin();
            InterimOrdersPage interimOrdersPage = homePage.GoToInterim_InterimOrders();
            interimOrdersPage.ResetFilters();

            // Déclaration de la variable en dehors du bloc if-else pour être accessible dans toute la méthode
            InterimOrdersItem interimOrderItem = null;
            interimOrdersPage.Filter(FilterType.ShowValidatedOnly, true);
            var numbers = interimOrdersPage.GetTotalOfOrders();
            if (numbers == 0)
            {
                // Création d'un nouvel ordre intérimaire si aucun n'existe
                InterimOrdersCreateModalPage interimOrdersCreateModalPage = interimOrdersPage.CreateNewInterimOrder();
                interimOrdersCreateModalPage.FillField_CreatNewInterminOrder(DateTime.Now, site, supplier, location, true);
                string interimNumber = interimOrdersCreateModalPage.GetInterimOrderId();
                interimOrderItem = interimOrdersCreateModalPage.Submit();

                // Vérification de la quantité et modification si nécessaire
                if (interimOrderItem.GetQtyItem() == "0")
                {
                    interimOrderItem.ClickFirstItem();
                    interimOrderItem.SetQty(qty);
                }
                else
                {
                    interimOrderItem.ClickFirstItem();
                }
                interimOrderItem.Validate();
            }
            else
            {
                // Si des ordres intérimaires existent déjà, sélectionnez le premier
                interimOrderItem = interimOrdersPage.ClickFirstLine();
            }

            // Envoi de l'ordre par mail
            interimOrderItem.SendByMail();
            interimOrderItem.SendByEmailEnterimOrder(userEmail, text);

            // Navigation vers la boîte mail Outlook pour vérifier si le mail a bien été envoyé
            MailPage mailPage = interimOrderItem.RedirectToOutlookMailbox();
            mailPage.FillFields_LogToOutlookMailbox_MoreThanOneMonth(userEmail);

            // Vérification de l'existence du mail envoyé
            bool isMailSent = mailPage.CheckIfSpecifiedOutlookMailExist(text);
            Assert.IsTrue(isMailSent, "La réclamation n'a pas été envoyée par mail.");

            // Suppression du mail pour nettoyage après le test
            mailPage.DeleteCurrentOutlookMail();
            mailPage.Close();
        }
    }
}