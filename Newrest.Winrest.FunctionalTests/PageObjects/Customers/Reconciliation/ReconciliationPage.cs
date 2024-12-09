using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.Extensions;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Text.RegularExpressions;
using System.Web.UI.WebControls;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.Reconciliation
{
    public class ReconciliationPage : PageBase
    {
        public ReconciliationPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //---------------------------------------CONSTANTES---------------------------------------------------------------
        private const string NOTIFICATIONS_ACTION_LINK = "/html/body/div[2]/div/div[2]/div[1]/nav/ul/li[1]/a";
        private const string ORDERS_ACTION_LINK = "/html/body/div[2]/div/div[2]/div[1]/nav/ul/li[2]/a";
        private const string COLLECTOR_MOUVEMENTS_ACTION_LINK = "/html/body/div[2]/div/div[2]/div[1]/nav/ul/li[3]/a";
        private const string BANK_MOUVEMENTS_ACTION_LINK = "/html/body/div[2]/div/div[2]/div[1]/nav/ul/li[4]/a";
        private const string PAYMENTS_CONTROL_ACTION_LINK = "/html/body/div[2]/div/div[2]/div[1]/nav/ul/li[5]/a";
        private const string ROUTES_ACTION_LINK = "/html/body/div[2]/div/div[2]/div[1]/nav/ul/li[6]/a";
        private const string FLIGHTS_ACTION_LINK = "/html/body/div[2]/div/div[2]/div[1]/nav/ul/li[7]/a";
        private const string BARSETS_ACTION_LINK = "/html/body/div[2]/div/div[2]/div[1]/nav/ul/li[8]/a";
        private const string NOTIFICATION_TYPE_ELEMENT = "ui-multiselect-1-NotificationsIndexSelectedNotificationTypes-option-{0}";
        private const string NOTIFICATION_ERRORTYPE_ELEMENT = "/html/body/div[12]/ul/li[*]/label/input[@value=\"{0}\"]";
        private const string NOTIFICATION_TYPE_DIV = "NotificationsIndexSelectedNotificationTypes_ms";
        private const string NOTIFICATION_ERRORTYPE_DIV = "NotificationsIndexSelectedErrorTypes_ms";

        private const string NOTIFICATIONS_RESULTS_LIST = "/html/body/div[2]/div/div[2]/div[2]/div/div/div/table/tbody/tr";
        public const string EXPORT_BUTTON = "//*[@id=\"nav-tab\"]/li[10]/div/div/div[1]/a";
        public const string EXPORT_BUTTON_VALIDATE = "//*[@id=\"btn-submit-export-params\"]";
        public const string EXPORT_SALES_SUMMARY_BUTTON = "//*[@id=\"nav-tab\"]/li[10]/div/div/div[2]/a";
        public const string UNFOLD = "unfoldBtn";
        public const string FOLD_ARROW_BUTTON = "/html/body/div[2]/div/div[2]/div[2]/div/div/div/table/tbody/tr[*]/td[1]/div/div/div"; 
        public const string CUSTOMER_NOTIFICATION_ID = "//*[contains(@id,\"id\")and starts-with(@id,\"airfi-notification\") and not(contains(@id, 'identifier'))]";



        //---------------------------------------FILTERS CONSTANTES-------------------------------------------------------
        private const string RESET_FILTER = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[1]/div/a";
        private const string CUSTOMERS_FILTER = "NotificationsIndexSelectedCustomers_ms";
        private const string UNSELECT_ALL_CUSTOMERS = "/html/body/div[10]/div/ul/li[2]/a";
        private const string SELECT_ALL_CUSTOMERS = "/html/body/div[10]/div/ul/li[1]/a";
        private const string CUSTOMERS_SEARCH = "/html/body/div[10]/div/div/label/input";
        private const string RECEPTION_DATE_FROM = "start-date-picker";
        private const string RECEPTION_DATE_TO = "end-date-picker";
        private const string BY_ID_SPAN = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[2]/a";
        private const string BY_ID_SEARCH = "search-notification-id-filter";
        private const string NOTIFICATION_STATE_HEADER = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[9]/a/label";
        private const string NOTIFICATION_DATATYPE_STATE_HEADER = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[7]/div/a/label";
        private const string NOTIFICATION_ERRORTYPE_STATE_HEADER = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[8]/div/a/label";
        private const string NOTIFICATION_DISPLAY_TEST_HEADER = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[10]/a/label";
        private const string SEARCH_BY_FLIGHTKEY = "search-by-flight-key-filter";
        private const string SEARCH_BY_FLIGHTNUMBER = "search-by-flight-number-filter";
        private const string SEARCH_BY_FLIGHTDATEFROM = "flight-date-picker-from";
        private const string SEARCH_BY_FLIGHTDATETO = "flight-date-picker-to";

        //--------------------------------------SORT BY-------------------------------------------------------------------
        //*[@id=\"row-result-{0}\"]/div[2]/table/tbody/tr";
        private const string SORT_BY_ROW_ID = "//table[@id='notifications-table']/tbody/tr[2*{0}-1][not(contains(@id,'span-to-show-'))]";
        private const string SORT_BY_ID = "header-span-id";
        private const string SORT_BY_FLIGHTNUMBER = "header-span-flightnumber";
        private const string SORT_BY_FLIGHTDATE = "header-span-flightdate";
        private const string SORT_BY_BARSET = "header-span-barset";
        private const string SORT_BY_RECEIVED = "header-span-received";
        private const string SORT_BY_IDENTIFIER = "header-span-identifier";
        private const string SORT_BY_ERRORTYPE = "header-span-error-type";
        private const string SORT_BY_DATATYPE = "header-span-data-type";
        private const string SORT_BY_TEST = "filter-checkbox-notification-state-test";


        // ____________________________________ Variables _______________________________________________
        [FindsBy(How = How.XPath, Using = NOTIFICATIONS_ACTION_LINK)]
        private IWebElement _notifications;

        [FindsBy(How = How.XPath, Using = ORDERS_ACTION_LINK)]
        private IWebElement _orders;

        [FindsBy(How = How.XPath, Using = COLLECTOR_MOUVEMENTS_ACTION_LINK)]
        private IWebElement _collectorMvt;

        [FindsBy(How = How.XPath, Using = BANK_MOUVEMENTS_ACTION_LINK)]
        private IWebElement _bankMvt;

        [FindsBy(How = How.XPath, Using = PAYMENTS_CONTROL_ACTION_LINK)]
        private IWebElement _paymentsControl;

        [FindsBy(How = How.XPath, Using = ROUTES_ACTION_LINK)]
        private IWebElement _routes;

        [FindsBy(How = How.XPath, Using = FLIGHTS_ACTION_LINK)]
        private IWebElement _flights;

        [FindsBy(How = How.XPath, Using = BARSETS_ACTION_LINK)]
        private IWebElement _barsets;

        [FindsBy(How = How.XPath, Using = EXPORT_BUTTON)]
        private IWebElement _export;

        [FindsBy(How = How.XPath, Using = EXPORT_BUTTON_VALIDATE)]
        private IWebElement _validateExport;

        [FindsBy(How = How.XPath, Using = EXPORT_SALES_SUMMARY_BUTTON)]
        private IWebElement _salesSummaryExport;

        [FindsBy(How = How.XPath, Using = UNFOLD)]
        private IWebElement _unfold;

        [FindsBy(How = How.XPath, Using = FOLD_ARROW_BUTTON)]
        private IWebElement _foldArrowButton;

        //_______________________________________________________ Filters _____________________________________________________________

        [FindsBy(How = How.XPath, Using = RESET_FILTER)]
        private IWebElement _resetFilterDev;

        [FindsBy(How = How.Id, Using = CUSTOMERS_FILTER)]
        private IWebElement _customersFilter;

        [FindsBy(How = How.XPath, Using = UNSELECT_ALL_CUSTOMERS)]
        private IWebElement _unselectAllCustomers;

        [FindsBy(How = How.XPath, Using = CUSTOMERS_SEARCH)]
        private IWebElement _searchCustomers;

        [FindsBy(How = How.Id, Using = RECEPTION_DATE_FROM)]
        private IWebElement _receptionDateFrom;

        [FindsBy(How = How.XPath, Using = RECEPTION_DATE_TO)]
        private IWebElement _receptionDateTo;


        [FindsBy(How = How.XPath, Using = BY_ID_SPAN)]
        private IWebElement _spanById;

        [FindsBy(How = How.Id, Using = BY_ID_SEARCH)]
        private IWebElement _searchById;

        [FindsBy(How = How.Id, Using = SEARCH_BY_FLIGHTKEY)]
        private IWebElement _searchByFlightKey;

        [FindsBy(How = How.Id, Using = SEARCH_BY_FLIGHTNUMBER)]
        private IWebElement _searchByFlightNumber;

        [FindsBy(How = How.Id, Using = SEARCH_BY_FLIGHTDATEFROM)]
        private IWebElement _searchByFlightDateFrom;

        [FindsBy(How = How.Id, Using = SEARCH_BY_FLIGHTDATETO)]
        private IWebElement _searchByFlightDateTo;

        //--------------------------------------SORT BY-------------------------------------------------------------------

        [FindsBy(How = How.Id, Using = SORT_BY_ID)]
        private IWebElement _divSortById;

        [FindsBy(How = How.Id, Using = SORT_BY_FLIGHTNUMBER)]
        private IWebElement _divSortByFlightNbr;

        [FindsBy(How = How.Id, Using = SORT_BY_FLIGHTDATE)]
        private IWebElement _divSortByFlightDate;

        [FindsBy(How = How.Id, Using = SORT_BY_BARSET)]
        private IWebElement _divSortByBarset;

        [FindsBy(How = How.Id, Using = SORT_BY_RECEIVED)]
        private IWebElement _divSortByReceived;

        [FindsBy(How = How.Id, Using = SORT_BY_IDENTIFIER)]
        private IWebElement _divSortByIdentifier;

        [FindsBy(How = How.Id, Using = SORT_BY_ERRORTYPE)]
        private IWebElement _divSortByErrorType;

        [FindsBy(How = How.Id, Using = SORT_BY_DATATYPE)]
        private IWebElement _divSortByDataType;

        #region Enums

        private static bool IsFirstComboboxCall = true;

        public enum FilterType
        {
            SearchbyNotificationId,
            SearchByBarset,
            SearchByFlightNumber,
            SearchByFlightDateFrom,
            SearchByFlightDateTo,

            ReceptionDateFrom,
            ReceptionDateTo,
            Customers,
            NotificationType,
            ErrorType,
            NotificationStates,
            DisplayTestCases
        }

        public enum SortByType
        {
            Id,
            FlightNumber,
            FlightDate,
            Barset,
            Received,
            Identifier,
            ErrorType,
            DataType
        }

        public enum NotificationType
        {
            None = 0,
            Scheduled = 1,
            Queued = 2,
            InProgress = 3,
            Failed = 4,
            Done = 5,
            Unscheduled = 6,
            Processed = 7,
            RouteToCompute = 8,
            FailedToComputeRoute = 9,
            ManuallyScheduled = 10
        }

        public enum NotificationDataType
        {
            Unknown = 0,
            AirfiPurchase = 1,
            AirfiStockMutation = 2,
            AirfiForm = 3,
            AirfiFlightLegSubscription = 4,
            AirfiStockConfirmation = 5,
            JusteatPurchase = 6,
            AirfiEndOfCycle = 7,
            AirfiEndOfFlight = 8,
        }

        public enum NotificationErrorType
        {
            Barset = 400,
            Crew = 600,
        }

        #endregion




        public void Go_To_Orders_ActionLink()
        {
            _orders = WaitForElementIsVisible(By.XPath(ORDERS_ACTION_LINK));
            _orders.Click();
            WaitPageLoading();
        }

        public void GoToNotificationsTab()
        {
            _notifications = WaitForElementIsVisible(By.XPath(NOTIFICATIONS_ACTION_LINK));
            _notifications.Click();
            WaitPageLoading();
        }

        public void ResetFilters()
        {
            _resetFilterDev = WaitForElementExists(By.XPath(RESET_FILTER));
            _resetFilterDev.Click();
        }

        public void Filter(FilterType filterType, object value)
        {
            switch (filterType)
            {
                case FilterType.SearchbyNotificationId:
                    _spanById = WaitForElementToBeClickable(By.XPath(BY_ID_SPAN));
                    _spanById.Click();
                    _searchById = WaitForElementIsVisible(By.Id(BY_ID_SEARCH));
                    _searchById.SetValue(ControlType.TextBox, value);
                    break;
                case FilterType.SearchByBarset:
                    _searchByFlightKey = WaitForElementExists(By.Id(SEARCH_BY_FLIGHTKEY));
                    _searchByFlightKey.SetValue(ControlType.TextBox, value);
                    break;
                case FilterType.SearchByFlightNumber:
                    _searchByFlightNumber = WaitForElementExists(By.Id(SEARCH_BY_FLIGHTNUMBER));
                    _searchByFlightNumber.SetValue(ControlType.TextBox, value);
                    break;
                case FilterType.SearchByFlightDateFrom:
                    _searchByFlightDateFrom = WaitForElementExists(By.Id(SEARCH_BY_FLIGHTDATEFROM));
                    _searchByFlightDateFrom.SetValue(ControlType.DateTime, value, "yyyy-MM-dd");
                    _searchByFlightDateFrom.SendKeys(Keys.Tab);
                    break;
                case FilterType.SearchByFlightDateTo:
                    _searchByFlightDateTo = WaitForElementExists(By.Id(SEARCH_BY_FLIGHTDATETO));
                    _searchByFlightDateTo.SetValue(ControlType.DateTime, value, "yyyy-MM-dd");
                    _searchByFlightDateTo.SendKeys(Keys.Tab);
                    break;
                case FilterType.ReceptionDateFrom:
                    _receptionDateFrom = WaitForElementExists(By.Id(RECEPTION_DATE_FROM));
                    _receptionDateFrom.SetValue(ControlType.DateTime, value, "yyyy-MM-dd");
                    _receptionDateFrom.SendKeys(Keys.Tab);
                    break;
                case FilterType.ReceptionDateTo:
                    _receptionDateTo = WaitForElementExists(By.Id(RECEPTION_DATE_TO));
                    _receptionDateTo.SetValue(ControlType.DateTime, value, "yyyy-MM-dd");
                    _receptionDateTo.SendKeys(Keys.Tab);
                    break;
                case FilterType.Customers:
                    _customersFilter = WaitForElementExists(By.Id(CUSTOMERS_FILTER));

                    //If the element is not displayed we might need to scroll to position
                    if (_customersFilter.Displayed == false)
                    {
                        ScrollToElement(_customersFilter);
                    }
                    _customersFilter.Click();

                    _unselectAllCustomers = WaitForElementIsVisible(By.XPath(UNSELECT_ALL_CUSTOMERS));
                    _unselectAllCustomers.Click();

                    _searchCustomers = WaitForElementIsVisible(By.XPath(CUSTOMERS_SEARCH));
                    _searchCustomers.SetValue(ControlType.TextBox, value);
                    WaitForLoad();

                    var customerToCheck = WaitForElementIsVisible(By.XPath("//span[contains(text(), '" + value + "')]"));
                    customerToCheck.SetValue(ControlType.CheckBox, true);

                    _customersFilter.Click();
                    break;
                case FilterType.NotificationType:
                    //Scroll to filter
                    ScrollToElement(WaitForElementExists(By.XPath(NOTIFICATION_DATATYPE_STATE_HEADER)));
                    if (IsFirstComboboxCall)
                    {
                        ComboBoxSelectById(new ComboBoxOptions(NOTIFICATION_TYPE_DIV, (string)value));
                        IsFirstComboboxCall = false;
                    }
                    else
                    {
                        ComboBoxSelectById(new ComboBoxOptions(NOTIFICATION_TYPE_DIV, (string)value) {ClickUncheckAllAtStart = false });
                    }
                    break;
                case FilterType.ErrorType:
                    //Scroll to filter
                    ScrollToElement(WaitForElementExists(By.XPath(NOTIFICATION_ERRORTYPE_STATE_HEADER)));
                    if (IsFirstComboboxCall)
                    {
                        ComboBoxSelectById(new ComboBoxOptions(NOTIFICATION_ERRORTYPE_DIV, (string)value));
                        IsFirstComboboxCall = false;
                    }
                    else
                    {
                        ComboBoxSelectById(new ComboBoxOptions(NOTIFICATION_ERRORTYPE_DIV, (string)value) { ClickUncheckAllAtStart = false });
                    }
                    break;
                case FilterType.NotificationStates:
                    //Scroll to filter
                    ScrollToElement(WaitForElementExists(By.XPath(NOTIFICATION_STATE_HEADER)));
                    //In Notifications tab, we only display New (None), Failed, In progress, Scheduled, Done
                    string elementId = GetNotificationStateCheckBoxId((ReconciliationPage.NotificationType)value);
                    var checkBoxFilter = WaitForElementExists(By.Id(elementId));
                    checkBoxFilter.Click();
                    break;
                case FilterType.DisplayTestCases:
                    ScrollToElement(WaitForElementExists(By.XPath(NOTIFICATION_DISPLAY_TEST_HEADER)));
                    var displayTestinput = WaitForElementExists(By.Id(SORT_BY_TEST));
                    bool check = (bool)value;
                    string checkedstr = displayTestinput.GetAttribute("checked");
                    if (displayTestinput.GetCssValue("checked") == "true")
                    {
                        if (!check)
                            displayTestinput.Click();
                    }
                    else
                    {
                        if (check)
                            displayTestinput.Click();
                    }
                    break;
                default:
                    throw new NotImplementedException();
            }
            WaitPageLoading();
        }

        /// <summary>
        /// Notification types are grouped
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        /// <exception cref="NotImplementedException"></exception>
        private string GetNotificationStateCheckBoxId(ReconciliationPage.NotificationType type)
        {
            switch (type)
            {
                case NotificationType.None:
                case NotificationType.Unscheduled:
                    return "filter-checkbox-notification-state-new";
                case NotificationType.Scheduled:
                case NotificationType.ManuallyScheduled:
                    return "filter-checkbox-notification-state-scheduled";
                case NotificationType.Queued:
                case NotificationType.InProgress:
                case NotificationType.Processed:
                case NotificationType.RouteToCompute:
                    return "filter-checkbox-notification-state-in-progress";
                case NotificationType.Failed:
                case NotificationType.FailedToComputeRoute:
                    return "filter-checkbox-notification-state-failed";
                case NotificationType.Done:
                    return "filter-checkbox-notification-state-done";
                default:
                    throw new NotImplementedException();
            }
        }

        public int GetNumberOfNotifications()
        {
            WaitPageLoading();
            var results = _webDriver.FindElements(By.XPath(NOTIFICATIONS_RESULTS_LIST));
            //As we have 1 TR hidden for each row, we need to divide by 2 the result to ge the real result
            return results != null ? results.Count/2 : 0;
        }

        public void SortBy(SortByType sortBy)
        {
            switch (sortBy)
            {
                case SortByType.Id:
                    _divSortById = WaitForElementExists(By.Id(SORT_BY_ID));
                    TryToClickOnElement(_divSortById);
                    break;
                case SortByType.FlightNumber:
                    _divSortByFlightNbr = WaitForElementExists(By.Id(SORT_BY_FLIGHTNUMBER));
                    TryToClickOnElement(_divSortByFlightNbr);
                    break;
                case SortByType.FlightDate:
                    _divSortByFlightDate = WaitForElementExists(By.Id(SORT_BY_FLIGHTDATE));
                    TryToClickOnElement(_divSortByFlightDate);
                    break;
                case SortByType.Barset:
                    _divSortByBarset = _webDriver.FindElement(By.Id(SORT_BY_BARSET));
                    TryToClickOnElement(_divSortByBarset);
                    break;
                case SortByType.Received:
                    _divSortByReceived = _webDriver.FindElement(By.Id(SORT_BY_RECEIVED));
                    TryToClickOnElement(_divSortByReceived);
                    break;
                case SortByType.Identifier:
                    _divSortByIdentifier = _webDriver.FindElement(By.Id(SORT_BY_IDENTIFIER));
                    TryToClickOnElement(_divSortByIdentifier);
                    break;
                case SortByType.ErrorType:
                    _divSortByErrorType = _webDriver.FindElement(By.Id(SORT_BY_ERRORTYPE));
                    TryToClickOnElement(_divSortByErrorType);
                    break;
                case SortByType.DataType:
                    _divSortByDataType = _webDriver.FindElement(By.Id(SORT_BY_DATATYPE));
                    TryToClickOnElement(_divSortByDataType);
                    break;
                default:
                    throw new NotImplementedException();
            }
            WaitPageLoading();
        }


        public List<string> GetCustemerNotificationsList()
        {
            var customerNotificationsListId = new List<string>();

            var customerNotificationsList = _webDriver.FindElements(By.XPath(CUSTOMER_NOTIFICATION_ID));

            foreach (var jobNotification in customerNotificationsList)
            {
                customerNotificationsListId.Add(jobNotification.Text);
            }

            return customerNotificationsListId;
        }

        /// <summary>
        /// Get sortBy information by line in a table (starting at line 1)
        /// </summary>
        /// <param name="sortBy"></param>
        /// <param name="line"></param>
        /// <returns></returns>

        public void ExportExcel(bool versionPrint)
        {

            ShowExtendedMenu();
            _export = WaitForElementIsVisible(By.XPath(EXPORT_BUTTON));
            _export.Click();
            WaitForLoad();

            _validateExport = WaitForElementIsVisible(By.XPath(EXPORT_BUTTON_VALIDATE));
            _validateExport.Click();
            WaitForLoad();

            if (versionPrint)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));
            }

            WaitForDownload();
            Close();
        }


        public FileInfo GetExcelFile(FileInfo[] taskFiles)
        {
            List<FileInfo> correctDownloadFiles = new List<FileInfo>();

            foreach (var file in taskFiles)
            {
                if (IsExcelFileCorrect(file.Name))
                {
                    correctDownloadFiles.Add(file);
                }
            }

            if (correctDownloadFiles.Count == 0)
            {
                return null;
            }

            var time = correctDownloadFiles[0].LastWriteTimeUtc;
            var correctFile = correctDownloadFiles[0];

            correctDownloadFiles.ForEach(file =>
            {
                if (time < file.LastWriteTimeUtc)
                {
                    time = file.LastWriteTimeUtc;
                    correctFile = file;
                }
            });

            return correctFile;
        }

        public bool IsExcelFileCorrect(string filePath)
        {
            string moisName = "(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)";
            string annee = "\\d{4}";
            string jour = "(?:0?[1-9]|[12]\\d|3[01])";
            string moisNbr = "\\d{1,2}";

            string date = $"{annee}_{moisNbr}_{jour}";

            Regex reg = new Regex($"^AirfiPuchaseRaws{jour}_{moisName}_{annee}_from_{annee}_{moisNbr}_{jour}_to_{annee}_{moisNbr}_{jour}.*\\.xlsx$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = reg.Match(filePath);

            return m.Success;
        }

        public FileInfo GetOrdersExcelFile(FileInfo[] taskFiles)
        {
            List<FileInfo> correctDownloadFiles = new List<FileInfo>();

            foreach (var file in taskFiles)
            {
                if (IsOrdersExcelFileCorrect(file.Name))
                {
                    correctDownloadFiles.Add(file);
                }
            }

            if (correctDownloadFiles.Count == 0)
            {
                return null;
            }

            var time = correctDownloadFiles[0].LastWriteTimeUtc;
            var correctFile = correctDownloadFiles[0];

            correctDownloadFiles.ForEach(file =>
            {
                if (time < file.LastWriteTimeUtc)
                {
                    time = file.LastWriteTimeUtc;
                    correctFile = file;
                }
            });

            return correctFile;
        }
        public bool IsOrdersExcelFileCorrect(string filePath)
        {
            string moisName = "(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)";
            string annee = "\\d{4}";
            string jour = "(?:0?[1-9]|[12]\\d|3[01])";
            string moisNbr = "\\d{1,2}";

            string date = $"{annee}_{moisNbr}_{jour}";

            Regex reg = new Regex($"^ReconOrder{jour}_{moisName}_{annee}_from_{annee}_{moisNbr}_{jour}_to_{annee}_{moisNbr}_{jour}.*\\.xlsx$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = reg.Match(filePath);

            return m.Success;
        }

        public void ExportSalesSummaryExcel(bool versionPrint)
        {

            ShowExtendedMenu();
            _salesSummaryExport = WaitForElementIsVisible(By.XPath(EXPORT_SALES_SUMMARY_BUTTON));
            _salesSummaryExport.Click();
            WaitForLoad();

            _validateExport = WaitForElementIsVisible(By.XPath(EXPORT_BUTTON_VALIDATE));
            _validateExport.Click();
            WaitForLoad();

            if (versionPrint)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));
            }

            WaitForDownload();
            Close();
        }
     
        public FileInfo GetSalesSummaryExcelFile(FileInfo[] taskFiles)
        {
            if (taskFiles == null || taskFiles.Length == 0)
            {
                return null;
            }

            FileInfo latestFile = null;

            foreach (var file in taskFiles)
            {
                if (IsSalesSummaryExcelFileCorrect(file.Name))
                {
                    if (latestFile == null || file.LastWriteTimeUtc > latestFile.LastWriteTimeUtc)
                    {
                        latestFile = file;
                    }
                }
            }

            return latestFile;
        }
        
        public bool IsSalesSummaryExcelFileCorrect(string fileName)
        {
            Console.WriteLine("Vérification du fichier : " + fileName);

            // Vérifier si le fichier est un fichier Excel
            if (!fileName.EndsWith(".xls") && !fileName.EndsWith(".xlsx"))
            {
                Console.WriteLine("Le fichier n'est pas un fichier Excel : " + fileName);
                return false;
            }

            // Vérifier si le nom du fichier contient le mot "SalesSummary"
            if (!fileName.Contains("SalesSummary"))
            {
                Console.WriteLine("Le fichier ne contient pas 'SalesSummary' : " + fileName);
                return false;
            }

            // Le fichier correspond au pattern recherché
            Console.WriteLine("Fichier valide : " + fileName);
            return true;
        }
        public void Fold()
        {
            WaitForLoad();

            ShowExtendedMenu();
            _unfold = WaitForElementIsVisible(By.Id(UNFOLD));
            _unfold.Click();
            WaitForLoad();

        }

        public Boolean IsFoldAll()
        {
            bool valueBool = false;

            var contents = _webDriver.FindElements(By.XPath(FOLD_ARROW_BUTTON));

            // Temps nécessaire pour que l'élément change de classe
            WaitPageLoading();
            for(var i = 0; i < contents.Count; i += 2)
            {
                if (!contents[i].GetAttribute("class").Contains("collapsed"))
                    valueBool = true;
            }
           
            return valueBool;
        }

        public Boolean IsUnfoldAll()
        {
            bool valueBool = false;

            var contents = _webDriver.FindElements(By.XPath(FOLD_ARROW_BUTTON));

            // Temps nécessaire pour que l'élément change de classe
            WaitPageLoading();
            for (var i = 0; i < contents.Count; i += 2)
            {
                if (contents[i].GetAttribute("class").Contains("collapsed"))
                    valueBool = true;
            }
            return valueBool;
        }
    }
}
