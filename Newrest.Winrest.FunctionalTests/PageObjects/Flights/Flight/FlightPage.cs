using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Drawing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.LpCart;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.Trolleys;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.Events;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Flights.Flight

{
    public class FlightPage : PageBase
    {

        public FlightPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //__________________________________ Constantes _____________________________________

        // General
        private const string PLUS_BTN = "//*[@id=\"tabContentItemContainer\"]/div[1]/div/div/div[2]/div[2]/button";
        private const string IMPORT_CUSTOMER_FILE = "//*[@id=\"tabContentItemContainer\"]/div[1]/div/div/div[2]/div[2]/div/a[text()='Import customer file']";
        private const string DIV_CUSTOMER_TO_IMPORT = "//*[@id=\"import-form\"]/div[1]/div[1]/div/div/div[1]";
        private const string CUSTOMER_TO_IMPORT_NAME = "dropdown-customer-selectized";
        private const string FIRST_CUSTOMER_TO_IMPORT = "//*[@id=\"import-form\"]/div[1]/div[1]/div/div/div[2]";
        private const string UPLOAD_FILE = "input-files";
        private const string IMPORT = "//*[@id=\"import-form\"]/div[2]/button[2]";
        private const string NEW_FLIGHT_BUTTON = "//*[@id=\"flights-flight-create-new-flight\"]";
        private const string NEW_FLIGHT_MULTIPLE = "btnAddMultipleFlight";
        private const string VALIDATE_BTN_DEV = "//*[@id=\"modal-1\"]/div[3]/button[3]";
        private const string VALIDATEALL_BTN = "/html/body/div[2]/div/div[2]/div/div[1]/div/div/div[2]/div[1]/div/div[4]/div[2]/a";
        private const string VALIDATE_ALL = "//*[@id=\"flights-validation-model\"]";
        private const string VALIDATEBT = "//*[@id=\"filter-for-flights\"]/div/div[2]/ul/li[5]/div/a";      
        private const string VALIDE = "btn-validate";
        private const string INVOICE_ALL = "//*[@id=\"flights-validation-model\"]";

        

        private const string EXTENDED_BTN = "//*[@id=\"tabContentItemContainer\"]/div[1]/div/div/div[2]/div[1]/button";
        private const string EXTENDED_BTN_FOLD_UNFOLD = "//button[text()='...']";
        private const string PRINT_EXTENDED_BTN = "//a[contains(text(), 'PRINT')]";
        private const string EXPORT_EXTENDED_BTN = "//a[contains(text(), 'EXPORT')]";
        private const string IMPORT_EXTENDED_BTN = "//a[contains(text(), 'IMPORT')]";
        private const string OTHERS_EXTENDED_BTN = "//a[contains(text(), 'OTHERS')]";
        private const string PRINT_DELIVERYNOTE_BTN = "//a[text()='Print delivery note']";
        private const string PRINT_TRACKSHEET_BTN = "//a[text()='Print track sheet']";
        private const string PRINT_TURNOVER_BTN = "//a[text()='Print turnover']";
        private const string PRINT_TROLLEYSLABELSB_BTN = "//a[text()='Print trolleys labels']";
        private const string PRINT_SWAPLPCARTREPORT_BTN = "//a[text()='SwapLPCart Report']";
        private const string PRINT_LPCARTSEALS_BTN = "//a[text()='Print LP Cart Seals']";
        private const string PRINT_SPECIALMEALS_BTN = "//a[text()='Print Special Meals']";
        private const string PRINT_CHECKLIST_BTN = "//a[text()='Print check list']";
        private const string CONFIRM_PRINT = "printButton";
        private const string EXPORTGUESTANDSERVICES = "btn-export-guestAndService-file-excel";
        private const string CREATEPRINTLABEL = "create-print-label";
        private const string CREATEPRINTLABELLPCARTSEALS = "printButton";
        private const string EXPORT = "btn-export-excel";
        private const string EXPORTLOG = "btn-exportlogs-excel";
        private const string EXPORTLOADINGBOB = "btn-export-loadingBOB";
        private const string EXPORTOP = "//a[contains(text(),'OP245')]";
        private const string EXPORTDOCKNUMBERFILE = "//a[text()='Export Dock Numbers File']";
        private const string IMPORTDOCKNUMBERFILE = "//a[text()='Import Dock Numbers File']";
        private const string EXPORTBARSETSTOCK = "//a[contains(text(),'Barset stock export')]";


        private const string PRINT_DELIVERYBOB_BTN = "//a[text()='Print Delivery BOB']";
        private const string VALIDATE_PRINTBOB = "btn-print";
        private const string MASSIVE_DELETE = "//*[@id=\"btn-massiveDelete\"]";
        private const string SEARCH_MENUS_BTN = "SearchFlightsBtn";
        private const string MULTIPLE_DELETE = "//*[@id=\"collapseOthers\"]/a[1]";
        private const string CONTENT = "//*[@id=\"list-item-with-action\"]/div/table/tbody/tr[1]/td[2]/a/span";
        private const string DETAILS = "//*[@id=\"list-item-with-action\"]/div/table/tbody/tr[2]/td";
        
        private const string ONE_FLIGHT_PER_PAGE = "//input[@id=\"isPageBreakByFlight\" and @type=\"checkbox\"]";
        private const string QUANTITY_NULL = "//input[@id=\"IsQuantityNull\" and @type=\"checkbox\"]";
        private const string OTHERS_GROUP_BTN = "Others-Group-btn";
        private const string ARROW_SPLIT = "//*[@id=\"rowFlightBarset\"]/a";
        // Views
        private const string CALENDAR_VIEW = "//*[@id=\"tabContentItemContainer\"]/div[1]/div/div/ul/li[2]/a";
        private const string CALENDAR_VIEW_ACTIVATED = "//*[@id=\"tabContentItemContainer\"]/div[1]/ul/li[2]";
        private const string DATE_LABEL_ID = "start-date-picker-detail";
        private const string CUSTOMERBOB_BTN = "li-bobcustomers-view";
        private const string DATE_FROM = "flight-to-delete-dateFrom";
        private const string DATE_TO = "flight-to-delete-dateTo";
        // Tableau
        private const string FIRST_FLIGHT = "//*[@id=\"flightTable\"]/tbody/tr[{0}]";
        private const string FIRST_FLIGHT_NUMBER = "//*[@id=\"flightTable\"]/tbody/tr[2]/td[2]/span";
        private const string FIRST_FLIGHT_AIRCRAFT = "item_AircraftName";
        private const string FIRST_FLIGHT_KEY = "//*[@id=\"flightTable\"]/tbody/tr[2]/td[2]/span[2]/a";
        private const string FLIGHT_NUMBER_INPUT = "item_Name";
        private const string EDIT_FLIGHT_BTN = "//*[@id=\"flightTable\"]/tbody/tr[*]/td[2]/span/../../td[*]/a/span[@class='glyphicon glyphicon-edit']";
        private const string EDIT_FIRST_FLIGHT_BTN = "//*/table[@id='flightTable']/tbody/tr[2]/td[22]/a";
        private const string P_STATE = "//*[@id=\"flightTable\"]/tbody/tr[*]/td[*]/div/ul/li[text()='P']";
        private const string V_STATE = "//*[@id=\"flightTable\"]/tbody/tr[*]/td[*]/div/ul/li[text()='V']";
        private const string I_STATE = "//*[@id=\"flightTable\"]/tbody/tr[*]/td[*]/div/ul/li[text()='I']";
        private const string C_STATE = "//*[@id=\"flightTable\"]/tbody/tr[*]/td[*]/div/div[2]";
        private const string FLIGHT_LEGS_NUMBER = "//*[@id=\"flightTable\"]/tbody/tr[{0}]/td[18]/div";
        // Filtres
        private const string RESET_FILTER = "ResetFilter";

        private const string SEARCH_INPUT = "SearchPattern";
        private const string SITES = "cbSites";
        private const string SORTBY = "cbSortBy";
        private const string CUSTOMERBOB_FILTER = "cbCustomers";
        private const string CUSTOMERS_MENU = "customers-filter";
        private const string CUSTOMER_FILTER = "SelectedCustomerIds_ms";
        private const string SEARCH_CUSTOMER_FILTER = "/html/body/div[14]/div/div/label/input";
        private const string UNCHECK_ALL_CUSTOMERS_FILTER = "/html/body/div[14]/div/ul/li[2]/a/span[2]";
        private const string STATUS_MENU = "status-filter";
        private const string EQUIPMENTTYPE_MENU = "/html/body/div[2]/div/div[1]/div[2]/div/form/div[17]/a";
        private const string EQUIPMENTTYPE_MENU_ID = "flights-flight-filter-equipment-type";
        private const string ACETYPE_0 = "acetype-0";
        private const string ACETYPE_1 = "acetype-1";
        private const string ACETYPE_2 = "acetype-2";
        private const string ATLAS = "ATLAS";
        private const string KSSU = "KSSU";
        private const string UNKNOWN = "UNKNOWN";
        private const string STATUS_FILTER = "SelectedStatus_ms";
        private const string TABLET_FLIGHT_TYPE_FILTER = "SelectedTabletFlightTypeIds_ms";
        private const string DAYTIME_ALL = "//*[@value=\"All\"]";
        private const string FLIGHTS_FILTER_ETD_FROM = "ETDFrom";
        private const string FLIGHTS_FILTER_ETD_TO = "ETDTo";
        private const string HIDE_CANCELLED_FLIGHT = "HideCancelledFlights";
        private const string NO_ACTIVE_PRICE = "NoActivePrice";
        private const string SHOW_NOT_COUNTED = "ShowOnlyNotTrolleyCounted";
        private const string FLIGHT_NUMBER = "//*[@id=\"flightTable\"]/tbody/tr[*]/td[2]/span";
        private const string CUSTOMER = "//*[@id=\"flightTable\"]/tbody/tr[*]/td[3]";
        private const string AIRCRAFT = "//*[@id=\"flightTable\"]/tbody/tr[*]/td[10]/span";
        private const string DOCKNUMBER = "//*[@id=\"flightTable\"]/tbody/tr[*]/td[12]/span[text()!='']";
        private const string DRIVER = "//*[@id=\"flightTable\"]/tbody/tr[*]/td[13]/span";
        private const string GATE = "//*[@id=\"flightTable\"]/tbody/tr[*]/td[9]/span";
        private const string FIRST_GATE = "//tr[contains(@class,\"line-selected\")]//input[@id=\"item_Gate\"]";
        private const string REGISTRATION = "//*[@id=\"flightTable\"]/tbody/tr[*]/td[11]/span";
        private const string SITE = "//*[@id=\"flightTable\"]/tbody/tr[*]/td[4]/span";
        private const string ETA = "//*[@id=\"flightTable\"]/tbody/tr[*]/td[6]/span";
        private const string ETD = "//*[@id=\"flightTable\"]/tbody/tr[*]/td[8]/span";
        private const string PRINT_FLIGHTS_LABELS = "//a[text()='Print Flights Labels']";
        private const string PRINT_CONFIRM = "create-print-label-Portrait";
        private const string DISPLAY_FLIGHT = "cbDisplayArrivalDeparture";
        private const string FILTER_SITE = "//*[@id=\"cbSites\"]/option[@selected=\"selected\"]";
        private const string SITES_FROM = "//*[@id=\"flightTable\"]/tbody/tr[*]/td[5]/span";
        private const string SITES_TO = "//*[@id=\"flightTable\"]/tbody/tr[*]/td[6]/span";
        private const string FLIGHTS_ASSIGNMENTS = "//a[text()='Print Flights assignment']";
        private const string FLIGHTS_ASSIGNMENTS_CONFIRM_PRINT = "create-print-flight-assignment";
        private const string FLIGHTS_TEMPS_CHECK_LIST = "//a[text()='Print flights TempChecks list']";
        private const string CONFIRM_FLIGHTS_TEMPS_CHECK_LIST = "create-print-flight-temp-check";
        private const string DELIVERY_NOTES_REPORTS = "//*[@id=\"SelectedTargetReports\"]/option";
        private const string FROM_DELIVERY_NOTES_REPORTS = "//*[@id=\"SelectedSourceReports\"]/option";
        private const string REMOVE_DELIVERY_NOTES_REPORT = "//*[@id=\"deliveryReportForm\"]/div[1]/div[2]/div/a[2]";
        private const string ADD_DELIVERY_NOTES_REPORT = "//*[@id=\"deliveryReportForm\"]/div[1]/div[2]/div/a[1]";
        private const string SAVE_BTN = "//*[@id=\"deliveryReportForm\"]/div[2]/div/button";
        private const string SHOW_STATUS = "//*[@id=\"status-filter\"]";
        //private const string SHOW_NOT_COUNTED = "//*[@id=\"tableListMenu\"]/tbody/tr[{0}]/td[1]/img[@alt='Inactive']";
        private const string UNFOLD_ALL = "//a[text()='Unfold All']";// "//*[@id=\"tabContentItemContainer\"]/div[1]/div/div/div/a[4]";
        private const string FOLD_ALL = "//a[text()='Fold All']";// "//*[@id=\"tabContentItemContainer\"]/div[1]/div/div/div/a[4]";
        private const string INCLUDE_QUANTITY = "IncludeDetails";
        private const string GROUPED_BY_CATEGORY = "isGroupedByCategory";

        private const string FLIGHTS_LIST = "//*[@id=\"rowFlightBarset\"]/span";
        private const string FLIGHTS_LIST_INVOICE = "//*[@id=\"flightTable\"]/tbody/tr[*]/td[20]/div/ul/li[3]";

        private const string PRIVAL_CLIC = "//*[@id=\"flights-flight-steps-1\"]/li[1]";
        private const string VALIDATE_CLIC = "//*[@id=\"flights-flight-steps-1\"]/li[2]";
        private const string INVOICE_CLIC = "//*[@id=\"flights-flight-steps-1\"]/li[3]";

        private const string INDEXAIRCRAFTVALUE = "//*[@id=\"flightTable\"]/tbody/tr[2]/td[11]/span/span[2]";
        private const string INDEXARRNUMBERVALUE = "//*[@id=\"flightTable\"]/tbody/tr[2]/td[9]";
        private const string INDEXDOCKNUMBERVALUE = "//*[@id=\"flightTable\"]/tbody/tr[2]/td[13]";
        private const string INDEX_ARR_NUMBER_INPUT = "item_ArrFlightNo";

        private const string FLIGHT_ALERTS = "FlightAlertFiltersSelectedFlightAlerts_ms";
        private const string FLIGHT_ALERTS_DROPDOWN = "//*[starts-with(@id,\"flight-alerts-dropdown\")]";
        private const string FlIGHT_ALERTS_DROPDOWN_ITEM = "//*[starts-with(@id,'flight-alerts-dropdown')]/../div[contains(@class,'dropdown-menu')]/p";
        private const string INDEXFLIGHTSNUMBER = "//*[@id=\"tabContentItemContainer\"]/div[1]/div/div/h1/span[2]";

        private const string DOWNLOADS_FIRST_ROW = "//*[starts-with(@id, 'printData')]";
        private const string DOWNLOAD_BUTTON = "header-print-button";
        private const string DOWNLOAD_PANEL = "//*[starts-with(@id,'popover')]";
        private const string EDIT_BUTTON = "//*[starts-with(@id,\"btnedit\")]";
        private const string FIRST_LINE = "//*[@id=\"flightTable\"]/tbody/tr[2]";
        private const string LIST_FLIGHT = "/html/body/div[2]/div/div[2]/div/div[2]/div[1]/div[1]/table/tbody/tr[*]/td[2]";

        private const string PREVAL = "//*[@id=\"flights-flight-steps-1\"]/li[1]";
        private const string VALIDATE = "//*[@id=\"flights-flight-steps-1\"]/li[2]";
        private const string INVOICE = "//*[@id=\"flights-flight-steps-1\"]/li[3]";

        private const string EDIT_FLIGHT = "//*/span[contains(@class,'edit')]/..";

        private const string FIRSTFLIGHT = "/html/body/div[2]/div/div[2]/div/div[2]/div[1]/div[1]/table/tbody/tr[2]/td[22]/a";
        private const string LP_CART_TAB = "//*[@id=\"LPCartTabNav\"]/a";
        private const string SELECT_ALL_MASSIVE_DELETE = "selectAll";
        private const string FIRST_FLIGHT_NOTVALID = "(//*[@id=\"flightTable\"]/tbody/tr[not(contains(@class,'readonly inactive'))])[2]";
        private const string FIRST_DRIVER_FLIGHT_NOTVALID = "(//*[@id=\"flightTable\"]/tbody/tr[not(contains(@class,'readonly inactive'))])[2]/td[14]/input";
        private const string CLICK_ON_TOP = "bobtopbarsets-view-mode-toggle-btn";
        private const string INDEX_DOCK_NUMBER_INPUT = "item_DockNumber";
        private const string DRIVER_VALUE = "//*[@id=\"flightTable\"]/tbody/tr[2]/td[14]/span";
        private const string ETA_VALUE = "//*[@id=\"flightTable\"]/tbody/tr[2]/td[7]/span";
        private const string ETA_INPUT = "ETA-line-flight";
        private const string REGISTRATION_INPUT = "item_RegistrationTypeName";
        private const string AIRCRAFT_INPUT = "/html/body/div[2]/div/div[2]/div/div[2]/div[1]/div[1]/table/tbody/tr[2]/td[11]/span/div/div/div/strong";
        private const string DRIVER_ITEM = "/html/body/div[2]/div/div[2]/div/div[2]/div[1]/div[1]/table/tbody/tr[2]/td[14]";

        private const string PRINTBUTTON = "header-print-button";
        
        //__________________________________ Variables _____________________________________
        // General
        [FindsBy(How = How.Id, Using = PRINTBUTTON)]
        private IWebElement _printButton;
        [FindsBy(How = How.Id, Using = VALIDE)]
        private IWebElement _confirm;
        [FindsBy(How = How.XPath, Using = EDIT_BUTTON)]
        private IWebElement _editButton;
        [FindsBy(How = How.Id, Using = DOWNLOAD_BUTTON)]
        private IWebElement _downloadButton;
        [FindsBy(How = How.XPath, Using = DOWNLOAD_PANEL)]
        private IWebElement _exportPanel;

        [FindsBy(How = How.XPath, Using = FIRSTFLIGHT)]
        private IWebElement _firstflight;

        [FindsBy(How = How.XPath, Using = EXPORTBARSETSTOCK)]
        private IWebElement _exportBarsetStock;

        [FindsBy(How = How.Id, Using = INCLUDE_QUANTITY)]
        private IWebElement _includequantity;

        [FindsBy(How = How.XPath, Using = PLUS_BTN)]
        private IWebElement _plusBtn;

        [FindsBy(How = How.XPath, Using = IMPORT_CUSTOMER_FILE)]
        private IWebElement _importCustomerFile;

        [FindsBy(How = How.XPath, Using = DIV_CUSTOMER_TO_IMPORT)]
        private IWebElement _divCustomerToImport;

        [FindsBy(How = How.Id, Using = CUSTOMER_TO_IMPORT_NAME)]
        private IWebElement _customerToImportName;

        [FindsBy(How = How.XPath, Using = FIRST_CUSTOMER_TO_IMPORT)]
        private IWebElement _firstCustomerToImport;

        [FindsBy(How = How.Id, Using = UPLOAD_FILE)]
        private IWebElement _uploadFile;

        [FindsBy(How = How.XPath, Using = IMPORT)]
        private IWebElement _importBtn;

        [FindsBy(How = How.XPath, Using = NEW_FLIGHT_BUTTON)]
        private IWebElement _createNewFlight;

        [FindsBy(How = How.Id, Using = NEW_FLIGHT_MULTIPLE)]
        private IWebElement _creatMultiFlight;

        [FindsBy(How = How.XPath, Using = EXTENDED_BTN)]
        private IWebElement _extendedButton;

        [FindsBy(How = How.XPath, Using = EXTENDED_BTN_FOLD_UNFOLD)]
        private IWebElement _extendedButtonFold;

        [FindsBy(How = How.XPath, Using = PRINT_EXTENDED_BTN)]
        private IWebElement _printExtendedButton;

        [FindsBy(How = How.XPath, Using = EXPORT_EXTENDED_BTN)]
        private IWebElement _exportExtendedButton;

        [FindsBy(How = How.XPath, Using = IMPORT_EXTENDED_BTN)]
        private IWebElement _importExtendedButton;

        [FindsBy(How = How.XPath, Using = OTHERS_EXTENDED_BTN)]
        private IWebElement _othersExtendedButton;

        [FindsBy(How = How.XPath, Using = PRINT_DELIVERYNOTE_BTN)]
        private IWebElement _printDeliveryNoteBtn;

        [FindsBy(How = How.XPath, Using = PRINT_TRACKSHEET_BTN)]
        private IWebElement _printTrackSheetBtn;

        [FindsBy(How = How.XPath, Using = PRINT_TURNOVER_BTN)]
        private IWebElement _printTurnoverBtn;

        [FindsBy(How = How.XPath, Using = PRINT_TROLLEYSLABELSB_BTN)]
        private IWebElement _printTrolleysLablesBtn;

        [FindsBy(How = How.XPath, Using = PRINT_SWAPLPCARTREPORT_BTN)]
        private IWebElement _printSwapLpCartReportBtn;

        [FindsBy(How = How.XPath, Using = PRINT_LPCARTSEALS_BTN)]
        private IWebElement _printLpCartSealsBtn;

        [FindsBy(How = How.XPath, Using = PRINT_SPECIALMEALS_BTN)]
        private IWebElement _printSpecialMealsBtn;

        [FindsBy(How = How.XPath, Using = PRINT_CHECKLIST_BTN)]
        private IWebElement _printChecklistBtn;

        [FindsBy(How = How.Id, Using = CONFIRM_PRINT)]
        private IWebElement _confirmPrint;

        [FindsBy(How = How.Id, Using = EXPORTGUESTANDSERVICES)]
        private IWebElement _exportGuestAndServices;

        [FindsBy(How = How.Id, Using = EXPORT)]
        private IWebElement _export;

        [FindsBy(How = How.Id, Using = EXPORTLOG)]
        private IWebElement _exportLog;

        [FindsBy(How = How.Id, Using = EXPORTLOADINGBOB)]
        private IWebElement _exportLoadingBob;

        [FindsBy(How = How.XPath, Using = EXPORTOP)]
        private IWebElement _exportOP;

        [FindsBy(How = How.XPath, Using = EXPORTDOCKNUMBERFILE)]
        private IWebElement _exportDOCKNUMBERFILE;

        [FindsBy(How = How.XPath, Using = IMPORTDOCKNUMBERFILE)]
        private IWebElement _importDockNumberFile;

        [FindsBy(How = How.XPath, Using = PRINT_DELIVERYBOB_BTN)]
        private IWebElement _printDeliveryBOBBtn;

        [FindsBy(How = How.Id, Using = VALIDATE_PRINTBOB)]
        private IWebElement _confirmPrintBOB;

        [FindsBy(How = How.XPath, Using = UNFOLD_ALL)]
        private IWebElement _unFoldAll;

        [FindsBy(How = How.XPath, Using = FOLD_ALL)]
        private IWebElement _foldAll;

        // Views
        [FindsBy(How = How.XPath, Using = CALENDAR_VIEW)]
        private IWebElement _calendarView;

        [FindsBy(How = How.Id, Using = DATE_LABEL_ID)]
        private IWebElement _dateLabel;

        [FindsBy(How = How.Id, Using = CUSTOMERBOB_BTN)]
        private IWebElement _customerBobBtn;


        [FindsBy(How = How.XPath, Using = QUANTITY_NULL)]
        private IWebElement _quantityNull;


        //[FindsBy(How = How.XPath, Using = ONE_FLIGHT_PER_PAGE)]
        private IWebElement _oneFlightPerPage;

        // Tableau
        [FindsBy(How = How.XPath, Using = FIRST_FLIGHT)]
        private IWebElement _firstFlight;

        [FindsBy(How = How.XPath, Using = FIRST_FLIGHT_NUMBER)]
        private IWebElement _flightNumber;

        [FindsBy(How = How.Id, Using = FIRST_FLIGHT_AIRCRAFT)]
        private IWebElement _flightAircraft;

        [FindsBy(How = How.XPath, Using = FIRST_FLIGHT_KEY)]
        private IWebElement _flightKey;

        [FindsBy(How = How.Id, Using = FLIGHT_NUMBER_INPUT)]
        private IWebElement _flightNumberInput;

        [FindsBy(How = How.XPath, Using = EDIT_FLIGHT_BTN)]
        private IWebElement _editFlightBtn;

        [FindsBy(How = How.XPath, Using = P_STATE)]
        private IWebElement _pStateBtn;

        [FindsBy(How = How.XPath, Using = V_STATE)]
        private IWebElement _vStateBtn;

        [FindsBy(How = How.XPath, Using = I_STATE)]
        private IWebElement _iStateBtn;

        [FindsBy(How = How.XPath, Using = C_STATE)]
        private IWebElement _cStateBtn;

        [FindsBy(How = How.XPath, Using = FLIGHT_LEGS_NUMBER)]
        private IWebElement _flightLegNb;

        [FindsBy(How = How.XPath, Using = MASSIVE_DELETE)]
        private IWebElement _massiveDelete;

        [FindsBy(How = How.XPath, Using = MULTIPLE_DELETE)]
        private IWebElement _multipleDelete;

        [FindsBy(How = How.Id, Using = DATE_FROM)]
        private IWebElement _Datefrom;

        [FindsBy(How = How.Id, Using = DATE_TO)]
        private IWebElement _Dateto;


        [FindsBy(How = How.Id, Using = GROUPED_BY_CATEGORY)]
        private IWebElement _groupedByCategory;

        [FindsBy(How = How.XPath, Using = VALIDATEALL_BTN)]
        private IWebElement _validateAllBtn;
        //__________________________________Filters_______________________________________

        [FindsBy(How = How.Id, Using = RESET_FILTER)]
        private IWebElement _resetFilter;

        [FindsBy(How = How.Id, Using = SEARCH_INPUT)]
        private IWebElement _searchInput;

        [FindsBy(How = How.Id, Using = SITES)]
        private IWebElement _sites;

        [FindsBy(How = How.Id, Using = SORTBY)]
        private IWebElement _sortBy;

        [FindsBy(How = How.Id, Using = CUSTOMERS_MENU)]
        private IWebElement _customersMenu;

        [FindsBy(How = How.Id, Using = CUSTOMER_FILTER)]
        private IWebElement _customerFilter;

        [FindsBy(How = How.XPath, Using = SEARCH_CUSTOMER_FILTER)]
        private IWebElement _searchCustomer;

        [FindsBy(How = How.XPath, Using = UNCHECK_ALL_CUSTOMERS_FILTER)]
        private IWebElement _uncheckAllCustomers;

        [FindsBy(How = How.Id, Using = STATUS_MENU)]
        private IWebElement _statusMenu;

        [FindsBy(How = How.XPath, Using = DAYTIME_ALL)]
        private IWebElement _dayTimeAll;

        [FindsBy(How = How.Id, Using = FLIGHTS_FILTER_ETD_FROM)]
        private IWebElement _ETDFrom;

        [FindsBy(How = How.Id, Using = FLIGHTS_FILTER_ETD_TO)]
        private IWebElement _ETDTo;

        [FindsBy(How = How.Id, Using = HIDE_CANCELLED_FLIGHT)]
        private IWebElement _hideCancelledFlight;

        [FindsBy(How = How.Id, Using = NO_ACTIVE_PRICE)]
        private IWebElement _noActivePrice;

        [FindsBy(How = How.Id, Using = SHOW_NOT_COUNTED)]
        private IWebElement _showNotCounted;

        [FindsBy(How = How.Id, Using = CUSTOMERBOB_FILTER)]
        private IWebElement _customerBob;

        [FindsBy(How = How.XPath, Using = PRINT_FLIGHTS_LABELS)]
        private IWebElement _printFlightsLabels;
        [FindsBy(How = How.Id, Using = INDEXAIRCRAFTVALUE)]
        private IWebElement _indexaircraftvalue;

        [FindsBy(How = How.Id, Using = INDEXDOCKNUMBERVALUE)]
        private IWebElement _indexdocknumbervalue;

        [FindsBy(How = How.Id, Using = FLIGHT_ALERTS)]
        private IWebElement _flightAlerts;
        [FindsBy(How = How.Id, Using = FIRST_LINE)]
        private IWebElement _firstLine;
        [FindsBy(How = How.ClassName, Using = "counter")]
        private IWebElement _totalNumber;

        [FindsBy(How = How.XPath, Using = PREVAL)]
        private IWebElement _preval;
        [FindsBy(How = How.XPath, Using = VALIDATE)]
        private IWebElement _validate;
        [FindsBy(How = How.XPath, Using = INVOICE)]
        private IWebElement _invoice;

        [FindsBy(How = How.Id, Using = CLICK_ON_TOP)]
        private IWebElement _clickOnTop;

        [FindsBy(How = How.XPath, Using = VALIDATE_ALL)]
        private IWebElement _valid;

        public enum FilterType
        {
            SearchFlight,
            Sites,
            SortBy,
            Customer,
            Status,
            dayTimeAll,
            ETDFrom,
            ETDTo,
            HideCancelledFlight,
            noActivePrice,
            showNotCounted,
            CustomerBob,
            Driver,
            DispalyFlight,
            DockNumber,
            TabletFlight,
            EQUIPMENTTYPE_Unknown,
            EQUIPMENTTYPE_KSSU,
            EQUIPMENTTYPE_Atlas,
            FlightAlerts,
        }

        public enum VerifyFilter
        {
            DockNumber,
            Driver,
            Gate
        }

        public void ShowStatusMenu()
        {
            _statusMenu = WaitForElementIsVisible(By.Id(STATUS_MENU));
            _statusMenu.Click();
            WaitPageLoading();
        }

        public void ShowCustomersMenu()
        {
            _customersMenu = WaitForElementIsVisible(By.Id(CUSTOMERS_MENU));
            _customersMenu.Click();
            WaitPageLoading();
        }

        public void ShowEquipmentTypeMenu()
        {
            ScrollUntilElementIsVisible(By.Id(EQUIPMENTTYPE_MENU_ID));
            _customersMenu = WaitForElementIsVisible(By.Id(EQUIPMENTTYPE_MENU_ID));
            _customersMenu.Click();
            WaitPageLoading();
        }

        public void Filter(FilterType filterType, object value)
        {
            WaitLoading();
            switch (filterType)
            {
                case FilterType.SearchFlight:
                    _searchInput = WaitForElementIsVisibleNew(By.Id(SEARCH_INPUT));
                    _searchInput.SetValue(ControlType.TextBox, value);
                    // pourquoi ?
                    WaitPageLoading();
                    WaitLoading();
                    break;
                case FilterType.Sites:
                    _sites = WaitForElementIsVisibleNew(By.Id(SITES));
                    _sites.SetValue(ControlType.DropDownList, value);
                    WaitForLoad();
                    break;
                case FilterType.SortBy:
                    if (IsDev())
                    {
                        _sortBy = WaitForElementIsVisibleNew(By.Id(SORTBY));
                        _sortBy.SetValue(ControlType.DropDownList, value);
                        // pourquoi ?
                        WaitPageLoading();
                        WaitForLoad();
                    }
                    else
                    {
                        _sortBy = WaitForElementIsVisibleNew(By.Id(SORTBY));
                        _sortBy.Click();
                        //var element = WaitForElementIsVisible(By.XPath("//option[contains(@value,'" + value + "')]"));
                        IWebElement element = WaitForElementIsVisible(By.XPath($"//*[@id=\"cbSortBy\"]/option[contains(text(), '{value}')]"));
                        _sortBy.SetValue(ControlType.DropDownList, element.Text);
                        _sortBy.Click();
                        // pourquoi ?
                        WaitPageLoading();
                        WaitForLoad();
                    }
                    break;
                case FilterType.Status:
                    string COLLAPSE_STATUS = "//*/label[text()='Status']/../../a[contains(@class,'collapsed')]";
                    if (isElementExists(By.XPath(COLLAPSE_STATUS)))
                    {
                        IWebElement collapse = WaitForElementExists(By.XPath(COLLAPSE_STATUS));
                        collapse.Click();
                        Thread.Sleep(2000);
                        WaitForLoad();
                    }
                    ComboBoxSelectById(new ComboBoxOptions(STATUS_FILTER, (string)value));
                    // pourquoi ?
                    WaitPageLoading();
                    break;
                case FilterType.Customer:
                    string COLLAPSE_CUSTOMER = "//*/label[text()='Customers']/../../a[contains(@class,'collapsed')]";
                    if (isElementVisible(By.XPath(COLLAPSE_CUSTOMER)))
                    {
                        IWebElement collapse = WaitForElementIsVisibleNew(By.XPath(COLLAPSE_CUSTOMER));
                        collapse.Click();
                        Thread.Sleep(2000);
                        WaitForLoad();
                    }
                    ComboBoxSelectById(new ComboBoxOptions(CUSTOMER_FILTER, (string)value));
                    break;
                case FilterType.dayTimeAll:
                    _dayTimeAll = WaitForElementExists(By.XPath(DAYTIME_ALL));
                    _dayTimeAll.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.ETDFrom:
                    _ETDFrom = WaitForElementIsVisibleNew(By.Id(FLIGHTS_FILTER_ETD_FROM), nameof(FLIGHTS_FILTER_ETD_FROM));
                    _ETDFrom.Clear();
                    _ETDFrom.Click();
                    _ETDFrom.SendKeys((string)value);
                    // pourquoi ?
                    WaitPageLoading();
                    WaitForLoad();
                    break;
                case FilterType.ETDTo:
                    _ETDTo = WaitForElementIsVisibleNew(By.Id(FLIGHTS_FILTER_ETD_TO), nameof(FLIGHTS_FILTER_ETD_TO));
                    _ETDTo.Clear();
                    _ETDTo.Click();
                    _ETDTo.SendKeys((string)value);
                    // pourquoi ?
                    WaitPageLoading();
                    WaitForLoad();
                    break;
                case FilterType.HideCancelledFlight:
                    _hideCancelledFlight = WaitForElementExists(By.Id(HIDE_CANCELLED_FLIGHT));
                    _hideCancelledFlight.SetValue(ControlType.CheckBox, value);
                    // pourquoi ?
                    WaitPageLoading();
                    WaitForLoad();
                    break;
                case FilterType.noActivePrice:
                    _noActivePrice = WaitForElementExists(By.Id(NO_ACTIVE_PRICE));
                    _noActivePrice.SetValue(ControlType.CheckBox, value);
                    break;
                case FilterType.showNotCounted:
                    _showNotCounted = WaitForElementExists(By.Id(SHOW_NOT_COUNTED));
                    _showNotCounted.SetValue(ControlType.CheckBox, value);
                    break;
                case FilterType.CustomerBob:
                    _customerBob = WaitForElementIsVisibleNew(By.Id(CUSTOMERBOB_FILTER));
                    _customerBob.SetValue(ControlType.DropDownList, value);
                    WaitForLoad();
                    break;
                case FilterType.DispalyFlight:
                    IWebElement input = WaitForElementIsVisibleNew(By.Id(DISPLAY_FLIGHT));
                    input.SetValue(ControlType.DropDownList, value);
                    // pourquoi ?
                    WaitPageLoading();
                    WaitForLoad();
                    break;
                case FilterType.DockNumber:
                    ComboBoxSelectById(new ComboBoxOptions("SelectedDockNumbers_ms", (string)value, false));
                    WaitPageLoading();
                    break;
                case FilterType.FlightAlerts:
                    ComboBoxSelectById(new ComboBoxOptions(FLIGHT_ALERTS, (string)value, false));
                    break;
                case FilterType.Driver:
                    Thread.Sleep(2000);
                    ComboBoxSelectById(new ComboBoxOptions("SelectedDrivers_ms", (string)value, false));
                    WaitPageLoading();
                    break;
                case FilterType.TabletFlight:
                    ScrollUntilElementIsVisible(By.XPath("//*[@id=\"SelectedTabletFlightTypeIds_ms\"]"));
                    ComboBoxSelectById(new ComboBoxOptions(TABLET_FLIGHT_TYPE_FILTER, (string)value, false));
                    break;
                case FilterType.EQUIPMENTTYPE_Unknown:

                    IWebElement collapsed = WaitForElementIsVisibleNew(By.XPath("//*/label[text()='Equipement type']/.."));
                    if (collapsed.GetAttribute("class").Contains("collapsed"))
                        collapsed.Click();
                    if ((bool)value == true)
                    {
                        IWebElement equipementId = WaitForElementExists(By.XPath("//*[@id=\"collapseEquipTypesFilter\"]/div[1]/label[text()='Unknown']/../div/input"));
                        equipementId.SetValue(ControlType.CheckBox, true);
                        WaitPageLoading();
                        WaitForLoad();
                    }
                    else
                    {
                        IWebElement equipementId = WaitForElementExists(By.XPath("//*[@id=\"collapseEquipTypesFilter\"]/div[1]/label[text()='Unknown']/../div/input"));
                        equipementId.SetValue(ControlType.CheckBox, false);
                        WaitPageLoading();
                        WaitForLoad();
                    }

                    break;
                case FilterType.EQUIPMENTTYPE_KSSU:

                    collapsed = WaitForElementIsVisible(By.XPath("//*/label[text()='Equipement type']/.."));
                    if (collapsed.GetAttribute("class").Contains("collapsed"))
                        collapsed.Click();
                    if ((bool)value == true)
                    {
                        IWebElement equipementId = WaitForElementExists(By.XPath("//*[@id=\"collapseEquipTypesFilter\"]/div[2]/label[text()='KSSU']/../div/input"));
                        equipementId.SetValue(ControlType.CheckBox, true);
                        WaitPageLoading();
                        WaitForLoad();
                    }
                    else
                    {
                        IWebElement equipementId = WaitForElementExists(By.XPath("//*[@id=\"collapseEquipTypesFilter\"]/div[2]/label[text()='KSSU']/../div/input"));
                        equipementId.SetValue(ControlType.CheckBox, false);
                        WaitPageLoading();
                        WaitForLoad();
                    }

                    break;
                case FilterType.EQUIPMENTTYPE_Atlas:

                    collapsed = WaitForElementIsVisible(By.XPath("//*/label[text()='Equipement type']/.."));
                    if (collapsed.GetAttribute("class").Contains("collapsed"))
                        collapsed.Click();
                    if ((bool)value == true)
                    {
                        IWebElement equipementId = WaitForElementExists(By.XPath("//*[@id=\"collapseEquipTypesFilter\"]/div[3]/label[text()='Atlas']/../div/input"));
                        equipementId.SetValue(ControlType.CheckBox, true);
                        WaitPageLoading();
                        WaitForLoad();
                    }
                    else
                    {
                        IWebElement equipementId = WaitForElementExists(By.XPath("//*[@id=\"collapseEquipTypesFilter\"]/div[3]/label[text()='Atlas']/../div/input"));
                        equipementId.SetValue(ControlType.CheckBox, false);
                        WaitPageLoading();
                        WaitForLoad();
                    }
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);
            }

            WaitPageLoading();
            WaitLoading();
        }

        public int GetTotalNumberLegs()
        {
            int totalNbLegs = 0;
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            for (int i = 2; i < CheckTotalNumber() + 2; i++)
            {
                WaitForLoad();
                _flightLegNb = WaitForElementIsVisible(By.XPath(String.Format("//*[@id=\"flightTable\"]/tbody/tr[{0}]/td[19]/div", i)));

                javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", _flightLegNb);

                if (!string.IsNullOrEmpty(_flightLegNb.Text))
                {
                    totalNbLegs += int.Parse(_flightLegNb.Text);
                }
            }
            return totalNbLegs;
        }
        public int GetTotalSplittedFlight()
        {
            int splittedFlightNb = 0;
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;

            //scroll up
            var scrollUpAll = _webDriver.FindElements(By.XPath(string.Format(FIRST_FLIGHT, 1)));
            var scrollUp = scrollUpAll.Last();
            javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", scrollUp);

            if (isElementVisible(By.XPath("//span[@class='glyphicon glyphicon-share-alt']")))
            {
                var splittedFlights = _webDriver.FindElements(By.XPath("//span[@class='glyphicon glyphicon-share-alt']"));
                splittedFlightNb = splittedFlights.Count;
            }
            else
            {
                var splittedFlights = _webDriver.FindElements(By.XPath("//span[@class='fas fa-share']"));
                splittedFlightNb = splittedFlights.Count;

            }
            return splittedFlightNb;
        }
        public void ResetFilter()
        {
            _resetFilter = WaitForElementIsVisibleNew(By.Id(RESET_FILTER));
            _resetFilter.Click();
            WaitLoading();
            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                SetDateState(DateUtils.Now);
            }
        }

        public void AddRegistrationOnFirstFlight(string registrationToSet)
        {
            //click on the first line 
            var line = WaitForElementIsVisible(By.XPath(FIRST_LINE));
            line.Click();
            //enter value in input
            var registration = _webDriver.FindElements(By.Id(REGISTRATION_INPUT)).FirstOrDefault();
            registration.SetValue(PageBase.ControlType.TextBox, registrationToSet);
            registration.SendKeys(Keys.Tab);
            WaitPageLoading();
        }

        public bool VerifyCustomer(string customer)
        {
            // taille de l'écran : 12 lignes du tableau visible
            int tot = CheckTotalNumber() > 12 ? 12 : CheckTotalNumber();


            if (tot == 0)
                return false;

            int compteur = 0;
            IEnumerable<IWebElement> customers;
            customers = _webDriver.FindElements(By.XPath("//*[@id=\"flightTable\"]/tbody/tr[*]/td[4]"));

            foreach (var elm in customers)
            {
                if (compteur >= tot)
                    break;

                if (!elm.GetAttribute("innerText").Contains(customer))
                {
                    return false;
                }

                compteur++;
            }
            return true;
        }

        public bool VerifyEquipementType(string equipementType)
        {
            // taille de l'écran : 12 lignes du tableau visible
            int tot = CheckTotalNumber() > 12 ? 12 : CheckTotalNumber();

            if (tot == 0)
                return false;

            IEnumerable<IWebElement> equipmentType;
            equipmentType = _webDriver.FindElements(By.XPath("//*[@id=\"flightTable\"]/tbody/tr[*]/td[11]"));

            if (equipementType == UNKNOWN)
            {
                foreach (var elem in equipmentType)
                {
                    if (!elem.Text.StartsWith("K") && !elem.Text.StartsWith("A"))
                    {
                        return false;
                    }
                }
            }

            if (equipementType == KSSU)
            {
                foreach (var elem in equipmentType)
                {
                    if (!elem.Text.StartsWith("A\r") && elem.Text.Contains("\r"))
                    {
                        return false;
                    }
                }
            }

            if (equipementType == ATLAS)
            {
                foreach (var elem in equipmentType)
                {
                    if (!elem.Text.StartsWith("K\r") && elem.Text.Contains("\r"))
                    {
                        return false;
                    }
                }
            }

            return true;
        }

        public bool VerifySite(List<string> airports)
        {
            int tot = CheckTotalNumber() > 12 ? 12 : CheckTotalNumber();

            if (tot == 0)
                return false;

            int compteur = 0;
            IEnumerable<IWebElement> sites;

            sites = _webDriver.FindElements(By.XPath("//*[@id=\"flightTable\"]/tbody/tr[*]/td[5]/span"));


            if (sites.Count() == 0)
                return false;

            foreach (var elm in sites)
            {
                if (compteur >= tot)
                    break;


                if (!airports.Contains(elm.GetAttribute("innerText")))
                {
                    return false;
                }

                compteur++;
            }
            return true;
        }
        public List<string> GetDockNumbers()
        {
            List<string> liste = new List<string>();
            var dockNumbersElements = _webDriver.FindElements(By.XPath("//*[@id=\"flightTable\"]/tbody/tr[*]/td[13]/span[text()!='']"));

            foreach (var elm in dockNumbersElements)
            {
                liste.Add(elm.GetAttribute("innerText").Trim());
            }
            return liste;
        }
        public IEnumerable<IWebElement> GetDrivers()
        {
            IEnumerable<IWebElement> driversElements;
            Thread.Sleep(2000);
            driversElements = _webDriver.FindElements(By.XPath("//*[@id=\"flightTable\"]/tbody/tr[*]/td[14]/span"));
            WaitForLoad();
            return driversElements;
        }
        public IEnumerable<IWebElement> GetGates()
        {
            Thread.Sleep(2000);
            ReadOnlyCollection<IWebElement> gatesElements;
            gatesElements = _webDriver.FindElements(By.XPath("//*[@id='flightTable']/tbody/tr[*]/td[10]/span"));
            WaitForLoad();
            return gatesElements;
        }
        public IEnumerable<IWebElement> GetRegistrations()
        {
            Thread.Sleep(2000);
            var registrationsElements = _webDriver.FindElements(By.XPath(REGISTRATION));
            return registrationsElements;
        }

        public bool VerifyAllDockNumbersExist(string var1, string var2, string var3)
        {
            var docknumber1exist = false;
            var docknumber2exist = false;
            var docknumber3exist = false;
            var docknumbers = GetDockNumbers();
            foreach (var docknum in docknumbers)
            {
                if (docknum.Equals(var1, StringComparison.InvariantCulture))
                {
                    docknumber1exist = true;
                }
                if (docknum.Equals(var2, StringComparison.InvariantCulture))
                {
                    docknumber2exist = true;
                }
                if (docknum.Equals(var3, StringComparison.InvariantCulture))
                {
                    docknumber3exist = true;
                }
            }
            if (docknumber1exist && docknumber2exist && docknumber3exist)
            {
                return true;
            }
            return false;
        }
        public bool VerifyAllDriversExist(string var1)
        {
            var driver1exist = false;
            var drivers = GetDrivers();
            foreach (var driver in drivers)
            {
                if (driver.Text.Equals(var1))
                {
                    driver1exist = true;
                }
            }
            if (driver1exist)
            {
                return true;
            }
            return false;
        }
        public bool VerifyAllGatesExist(string var1)
        {
            var gate1exist = false;
            var gates = GetGates();
            foreach (var docknum in gates)
            {
                if (docknum.Text.Equals(var1))
                {
                    gate1exist = true;
                }
            }
            if (gate1exist)
            {
                return true;
            }
            return false;
        }
        public bool VerifyRegistrationsNotExist(string var1, string var2, string var3)
        {
            var reg1exist = false;
            var reg2existt = false;
            var reg3exist = false;
            var registrations = GetRegistrations();
            foreach (var reg in registrations)
            {
                if (reg.Text.Equals(var1))
                {
                    reg1exist = true;
                }
                if (reg.Text.Equals(var2))
                {
                    reg2existt = true;
                }
                if (reg.Text.Equals(var3))
                {
                    reg3exist = true;
                }
            }
            if (!reg1exist && !reg2existt && !reg3exist)
            {
                return true;
            }
            return false;
        }

        public bool VerifyAircraft(string aircraft)
        {
            int tot = CheckTotalNumber() > 12 ? 12 : CheckTotalNumber();
            if (tot == 0)
                return false;
            int compteur = 0;
            var aircraftelements = _webDriver.FindElements(By.XPath(AIRCRAFT));
            if (aircraftelements.Count == 0)
                return false;
            foreach (var elm in aircraftelements)
            {
                if (compteur >= tot)
                    break;
                if (!aircraft.Contains(elm.Text))
                {
                    return false;
                }
                compteur++;
            }
            return true;
        }
        public bool VerifyCustomerICAO(string customerICAO)
        {
            int tot = CheckTotalNumber() > 12 ? 12 : CheckTotalNumber();
            if (tot == 0)
                return false;
            int compteur = 0;
            var cutomerICAOElements = _webDriver.FindElements(By.XPath(CUSTOMER));
            if (cutomerICAOElements.Count == 0)
                return false;
            foreach (var elm in cutomerICAOElements)
            {
                if (compteur >= tot)
                    break;
                if (!customerICAO.Contains(elm.Text))
                {
                    return false;
                }
                compteur++;
            }
            return true;
        }
        public bool IsSortedByFrom()
        {
            int tot = CheckTotalNumber() > 12 ? 12 : CheckTotalNumber();

            if (tot == 0)
                return false;
            ReadOnlyCollection<IWebElement> sites;
            sites = _webDriver.FindElements(By.XPath("//*[@id='flightTable']/tbody/tr[*]/td[5]/span"));

            var ancienSite = sites[0].GetAttribute("innerText");
            int compteur = 0;

            foreach (var elm in sites)
            {
                // Problème exploitation résultats à cause du type d'affichage des données de la catégorie (on ne prend que les 12 premiers)
                if (compteur >= tot)
                    break;

                if (String.Compare(ancienSite, elm.GetAttribute("innerText")) > 0)
                    return false;


                ancienSite = elm.GetAttribute("innerText");
                compteur++;
            }

            return true;
        }

        public bool IsSortedByEtd()
        {
            int tot = CheckTotalNumber() > 12 ? 12 : CheckTotalNumber();

            if (tot == 0)
                return false;

            ReadOnlyCollection<IWebElement> listeEtd;
            listeEtd = _webDriver.FindElements(By.XPath("//*[@id=\"flightTable\"]/tbody/tr[*]/td[8]/span"));

            var ancienEtd = listeEtd[0].GetAttribute("innerText");
            int compteur = 0;

            foreach (var elm in listeEtd)
            {
                // Problème exploitation résultats à cause du type d'affichage des données de la catégorie (on ne prend que les 12 premiers)
                if (compteur >= tot)
                    break;

                if (String.Compare(ancienEtd, elm.GetAttribute("innerText")) > 0)
                    return false;


                ancienEtd = elm.GetAttribute("innerText");
                compteur++;
            }

            return true;
        }

        public bool IsSortedByEta()
        {
            int tot = CheckTotalNumber() > 12 ? 12 : CheckTotalNumber();

            if (tot == 0)
                return false;
            ReadOnlyCollection<IWebElement> listeEta;
            listeEta = _webDriver.FindElements(By.XPath("//*[@id=\"flightTable\"]/tbody/tr[*]/td[7]/span"));

            var ancienEta = listeEta[0].GetAttribute("innerText");
            int compteur = 0;

            foreach (var elm in listeEta)
            {
                // Problème exploitation résultats à cause du type d'affichage des données de la catégorie (on ne prend que les 12 premiers)
                if (compteur >= tot)
                    break;

                if (String.Compare(ancienEta, elm.GetAttribute("innerText")) > 0)
                    return false;


                ancienEta = elm.GetAttribute("innerText");
                compteur++;
            }

            return true;
        }

        public bool IsSortedByFlightNo()
        {
            int tot = CheckTotalNumber() > 12 ? 12 : CheckTotalNumber();

            if (tot == 0)
                return false;

            var listeFlightNumber = _webDriver.FindElements(By.XPath(FLIGHT_NUMBER));

            var ancienFlightNo = listeFlightNumber[0].GetAttribute("innerText");
            int compteur = 0;

            foreach (var elm in listeFlightNumber)
            {
                // Problème exploitation résultats à cause du type d'affichage des données de la catégorie (on ne prend que les 12 premiers)
                if (compteur >= tot)
                    break;

                if (String.Compare(ancienFlightNo, elm.GetAttribute("innerText")) > 0)
                    return false;


                ancienFlightNo = elm.GetAttribute("innerText");
                compteur++;
            }

            return true;
        }

        public bool IsSortedByAirline()
        {
            int tot = CheckTotalNumber() > 12 ? 12 : CheckTotalNumber();

            if (tot == 0)
                return false;

            ReadOnlyCollection<IWebElement> listeAirlines;
            listeAirlines = _webDriver.FindElements(By.XPath("//*[@id='flightTable']/tbody/tr[*]/td[4]"));

            var ancienAirline = listeAirlines[0].GetAttribute("innerText");
            int compteur = 0;

            foreach (var elm in listeAirlines)
            {
                // Problème exploitation résultats à cause du type d'affichage des données de la catégorie (on ne prend que les 12 premiers)
                if (compteur >= tot)
                    break;

                if (String.Compare(ancienAirline, elm.GetAttribute("innerText")) > 0)
                    return false;


                ancienAirline = elm.GetAttribute("innerText");
                compteur++;
            }

            return true;
        }

        public bool IsSortedByAircraft()
        {
            int tot = CheckTotalNumber() > 12 ? 12 : CheckTotalNumber();

            if (tot == 0)
                return false;
            ReadOnlyCollection<IWebElement> listeAircrafts;
            listeAircrafts = _webDriver.FindElements(By.XPath("//*[@id='flightTable']/tbody/tr[*]/td[11]/span"));

            var ancienAircraft = listeAircrafts[0].Text;
            int compteur = 0;

            foreach (var elm in listeAircrafts)
            {
                if (elm.Text == "") continue;

                // Problème exploitation résultats à cause du type d'affichage des données de la catégorie (on ne prend que les 12 premiers)
                if (compteur >= tot)
                    break;

                if (ancienAircraft != "")
                {
                    if (String.Compare(ancienAircraft, elm.Text) > 0)
                        return false;
                }

                ancienAircraft = elm.Text;
                compteur++;
            }

            return true;
        }


        public bool IsSortedByDriver()
        {
            int tot = CheckTotalNumber() > 12 ? 12 : CheckTotalNumber();

            if (tot == 0)
                return false;
            ReadOnlyCollection<IWebElement> listeDrivers;
            listeDrivers = _webDriver.FindElements(By.XPath("//*[@id='flightTable']/tbody/tr[*]/td[14]/span"));

            var ancienDriver = listeDrivers[0].Text;
            int compteur = 0;

            foreach (var elm in listeDrivers)
            {
                // Problème exploitation résultats à cause du type d'affichage des données de la catégorie (on ne prend que les 12 premiers)
                if (compteur >= tot)
                    break;

                if (String.Compare(ancienDriver, elm.Text) > 0)
                    return false;


                ancienDriver = elm.Text;
                compteur++;
            }
            return true;
        }
        public bool IsSortedByGate()
        {
            int tot = CheckTotalNumber() > 12 ? 12 : CheckTotalNumber();

            if (tot == 0)
                return false;
            ReadOnlyCollection<IWebElement> listeGates;

            listeGates = _webDriver.FindElements(By.XPath("//*[@id=\"flightTable\"]/tbody/tr[*]/td[10]/span"));

            var ancienGate = listeGates[0].Text;
            int compteur = 0;

            foreach (var elm in listeGates)
            {
                // Problème exploitation résultats à cause du type d'affichage des données de la catégorie (on ne prend que les 12 premiers)
                if (compteur >= tot)
                    break;

                if (String.Compare(ancienGate, elm.Text) > 0)
                    return false;


                ancienGate = elm.Text;
                compteur++;
            }
            return true;
        }
        public bool IsSortedByResgistration()
        {
            int tot = CheckTotalNumber() > 12 ? 12 : CheckTotalNumber();

            if (tot == 0)
                return false;

            ReadOnlyCollection<IWebElement> listeRegistrations;
            listeRegistrations = _webDriver.FindElements(By.XPath("//*[@id='flightTable']/tbody/tr[*]/td[12]/span"));

            var ancienRegistration = listeRegistrations[0].Text;
            int compteur = 0;

            foreach (var elm in listeRegistrations)
            {
                // Problème exploitation résultats à cause du type d'affichage des données de la catégorie (on ne prend que les 12 premiers)
                if (compteur >= tot)
                    break;

                if (String.Compare(ancienRegistration, elm.Text) > 0)
                    return false;


                ancienRegistration = elm.Text;
                compteur++;
            }
            return true;
        }
        //___________________________________ Méthodes ____________________________________________

        // General
        public override void ShowPlusMenu()
        {
            Actions actions = new Actions(_webDriver);

            _plusBtn = WaitForElementIsVisibleNew(By.XPath(PLUS_BTN), nameof(PLUS_BTN));
            actions.MoveToElement(_plusBtn).Perform();
            WaitPageLoading();
        }

        public void ImportCustomerFile(string customer, string fileName)
        {
            ShowPlusMenu();

            _importCustomerFile = WaitForElementIsVisible(By.XPath(IMPORT_CUSTOMER_FILE), nameof(IMPORT_CUSTOMER_FILE));
            _importCustomerFile.Click();
            WaitForLoad();

            _divCustomerToImport = WaitForElementIsVisible(By.XPath(DIV_CUSTOMER_TO_IMPORT), nameof(DIV_CUSTOMER_TO_IMPORT));
            _divCustomerToImport.Click();
            _divCustomerToImport.Click();

            _customerToImportName = WaitForElementIsVisible(By.Id(CUSTOMER_TO_IMPORT_NAME));
            _customerToImportName.SetValue(ControlType.TextBox, customer);

            _firstCustomerToImport = WaitForElementToBeClickable(By.XPath(FIRST_CUSTOMER_TO_IMPORT));
            _firstCustomerToImport.Click();
            WaitForLoad();

            _uploadFile = WaitForElementIsVisible(By.Id(UPLOAD_FILE));
            _uploadFile.SendKeys(fileName);
            WaitPageLoading();
            Thread.Sleep(3000);

            _importBtn = WaitForElementIsVisible(By.XPath(IMPORT), nameof(IMPORT));
            _importBtn.Click();
            WaitForLoad();
        }

        public FlightCreateModalPage FlightCreatePage()
        {
            ShowPlusMenu();
            _createNewFlight = WaitForElementIsVisibleNew(By.XPath(NEW_FLIGHT_BUTTON), nameof(NEW_FLIGHT_BUTTON));
            _createNewFlight.Click();
            WaitPageLoading();

            return new FlightCreateModalPage(_webDriver, _testContext);
        }

        public FlightMultiCreateModalPage CreateMultiFlights()
        {
            ShowPlusMenu();

            _creatMultiFlight = WaitForElementIsVisible(By.Id(NEW_FLIGHT_MULTIPLE), nameof(NEW_FLIGHT_MULTIPLE));
            _creatMultiFlight.Click();
            WaitForLoad();

            return new FlightMultiCreateModalPage(_webDriver, _testContext);
        }

        public override void ShowExtendedMenu()
        {
            ClosePrintButton();//popup still visible sometimes
            var actions = new Actions(_webDriver);
            _extendedButton = WaitForElementIsVisible(By.XPath(EXTENDED_BTN), nameof(EXTENDED_BTN));
            actions.MoveToElement(_extendedButton).MoveByOffset(-3, 0).Perform();
            WaitForLoad();
        }

        public void ShowExtendedMenuFold()
        {
            var actions = new Actions(_webDriver);
            _extendedButton = WaitForElementIsVisible(By.XPath(EXTENDED_BTN_FOLD_UNFOLD), nameof(EXTENDED_BTN_FOLD_UNFOLD));
            actions.MoveToElement(_extendedButtonFold).MoveByOffset(-3, 0).Perform();
            WaitForLoad();
        }

        public enum PrintType
        {
            DeliveryNote,
            TrackSheet,
            Turnover,
            Checklist,
            DeliveryBOB,
            FlightsLabels,
            FlightAssignments,
            FlightsTempChecksList,
            Trolleyslabels,
            SwapLpCartReport,
            LpCartSeals,
            SpecialMeals
        }

        public enum ExportType
        {
            ExportGuestAndServices,
            Export,
            ExportLogs,
            ExportOP,
            ExportLoadingBob,
            ExportDockNumberFile,
            BarsetStockExport
        }
        public enum ImportType
        {
            ImportGuestAndServices,
            ImportDockNumbersFile
        }
        public enum OtherFlightType
        {
            ValidateAll
        }

        public PrintReportPage PrintReport(PrintType reportType, bool versionPrint, bool DeliveryNote = false, bool isOneFlightPerPage = false, bool isQuantityNull = false, bool groupedBy = false, bool safetyNote = false)
        {
            WaitForLoad();

            ShowExtendedMenu();

            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            var actions = new Actions(_webDriver);
            WaitForLoad();
            if (!PrintMenuIsClicked())
            {
                _printExtendedButton = WaitForElementIsVisible(By.Id("Print-Group-btn"));
                _printExtendedButton.Click();
                WaitForLoad();
            }

            switch (reportType)
            {
                case PrintType.DeliveryNote:
                    _printDeliveryNoteBtn = WaitForElementIsVisible(By.XPath(PRINT_DELIVERYNOTE_BTN), nameof(PRINT_DELIVERYNOTE_BTN));
                    _printDeliveryNoteBtn.Click();
                    WaitForLoad();

                    //if (DeliveryNote)
                    //{
                    //WebDriverWait waith = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(20));
                    //}
                    // animation fold
                    Thread.Sleep(2000);
                    _includequantity = WaitForElementExists(By.Id(INCLUDE_QUANTITY));
                    actions.MoveToElement(_includequantity).Perform();
                    _includequantity.SetValue(ControlType.CheckBox, DeliveryNote);
                    WaitForLoad();

                    // seulement le parameter
                    // Settings>Flights>Delivery report> checkbox "Print safety notes and safety comments"
                    // if (safetyNote)
                    // {
                    // var safe = WaitForElementIsVisible(By.Id("scYES"));
                    // safe.SetValue(ControlType.RadioButton, true);
                    // WaitForLoad();
                    // }
                    if (groupedBy)
                    {
                        _groupedByCategory = WaitForElementExists(By.Id(GROUPED_BY_CATEGORY));
                        actions.MoveToElement(_groupedByCategory).Perform();
                        _groupedByCategory.Click();
                        WaitForLoad();
                    }

                    _confirmPrint = WaitForElementIsVisible(By.Id(CONFIRM_PRINT));
                    _confirmPrint.Click();
                    WaitForLoad();
                    break;

                case PrintType.TrackSheet:
                    _printTrackSheetBtn = WaitForElementIsVisible(By.XPath(PRINT_TRACKSHEET_BTN), nameof(PRINT_TRACKSHEET_BTN));
                    _printTrackSheetBtn.Click();
                    WaitPageLoading();
                    WaitForLoad();

                    if (isOneFlightPerPage)
                    {
                        if (IsDev())
                        {
                            _oneFlightPerPage = WaitForElementExists(By.XPath(ONE_FLIGHT_PER_PAGE));
                        }
                        else
                        {
                            _oneFlightPerPage = WaitForElementExists(By.Id("isPageBreakByFlight"));
                        }
                        _oneFlightPerPage.SetValue(ControlType.CheckBox, true);
                        WaitForLoad();
                    }
                    if (isQuantityNull)
                    {
                        if (IsDev())
                        {
                            _quantityNull = WaitForElementExists(By.XPath(QUANTITY_NULL));
                        }
                        else
                        {
                            _quantityNull = WaitForElementExists(By.Id("IsQuantityNull"));

                        }
                        _quantityNull.SetValue(ControlType.CheckBox, true);
                        WaitForLoad();
                    }

                    _confirmPrint = WaitForElementIsVisible(By.Id(CONFIRM_PRINT));
                    _confirmPrint.Click();
                    WaitForLoad();
                    break;

                case PrintType.Turnover:
                    _printTurnoverBtn = WaitForElementIsVisible(By.XPath(PRINT_TURNOVER_BTN), nameof(PRINT_TURNOVER_BTN));
                    _printTurnoverBtn.Click();
                    WaitForLoad();

                    _confirmPrint = WaitForElementIsVisible(By.Id(CONFIRM_PRINT));
                    _confirmPrint.Click();
                    WaitForLoad();
                    break;

                case PrintType.Checklist:
                    _printChecklistBtn = WaitForElementIsVisible(By.XPath(PRINT_CHECKLIST_BTN), nameof(PRINT_CHECKLIST_BTN));
                    _printChecklistBtn.Click();
                    WaitForLoad();

                    _confirmPrint = WaitForElementIsVisible(By.Id(CONFIRM_PRINT));
                    _confirmPrint.Click();
                    WaitForLoad();
                    break;

                case PrintType.DeliveryBOB:
                    _printDeliveryBOBBtn = WaitForElementIsVisible(By.XPath(PRINT_DELIVERYBOB_BTN), nameof(PRINT_DELIVERYBOB_BTN));
                    javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", _printDeliveryBOBBtn);
                    _printDeliveryBOBBtn.Click();
                    WaitForLoad();

                    if (versionPrint)
                    {
                        _confirmPrintBOB = WaitForElementIsVisible(By.Id(VALIDATE_PRINTBOB));
                        _confirmPrintBOB.Click();
                        WaitForLoad();
                    }
                    break;
                case PrintType.FlightsLabels:

                    _printFlightsLabels = WaitForElementIsVisible(By.XPath(PRINT_FLIGHTS_LABELS));
                    _printFlightsLabels.Click();

                    var printPortrait = WaitForElementIsVisible(By.Id(PRINT_CONFIRM));
                    printPortrait.Click();
                    WaitForLoad();
                    break;
                case PrintType.FlightAssignments:
                    var flightAssignments = WaitForElementIsVisible(By.XPath(FLIGHTS_ASSIGNMENTS));
                    flightAssignments.Click();
                    var confirmPrint = WaitForElementIsVisible(By.Id(FLIGHTS_ASSIGNMENTS_CONFIRM_PRINT));
                    confirmPrint.Click();
                    WaitForLoad();
                    break;
                case PrintType.FlightsTempChecksList:
                    var flightsTempChecks = WaitForElementIsVisible(By.XPath(FLIGHTS_TEMPS_CHECK_LIST));
                    flightsTempChecks.Click();
                    WaitForLoad();
                    var printBtn = WaitForElementIsVisible(By.Id(CONFIRM_FLIGHTS_TEMPS_CHECK_LIST));
                    printBtn.Click();
                    WaitPageLoading();
                    // animation fenêtre qui va vers le haut
                    Thread.Sleep(2000);
                    WaitForLoad();
                    break;
                case PrintType.Trolleyslabels:
                    _printTrolleysLablesBtn = WaitForElementIsVisible(By.XPath(PRINT_TROLLEYSLABELSB_BTN), nameof(PRINT_TROLLEYSLABELSB_BTN));
                    _printTrolleysLablesBtn.Click();
                    WaitForLoad();
                    _printTrolleysLablesBtn = WaitForElementIsVisible(By.Id(CREATEPRINTLABEL));
                    _printTrolleysLablesBtn.Click();
                    WaitForLoad();
                    break;
                case PrintType.SwapLpCartReport:
                    _printSwapLpCartReportBtn = WaitForElementIsVisible(By.XPath(PRINT_SWAPLPCARTREPORT_BTN), nameof(PRINT_SWAPLPCARTREPORT_BTN));
                    _printSwapLpCartReportBtn.Click();
                    WaitForLoad();
                    break;
                case PrintType.LpCartSeals:
                    _printLpCartSealsBtn = WaitForElementIsVisible(By.XPath(PRINT_LPCARTSEALS_BTN), nameof(PRINT_LPCARTSEALS_BTN));
                    _printLpCartSealsBtn.Click();
                    WaitForLoad();
                    _printLpCartSealsBtn = WaitForElementIsVisible(By.Id(CREATEPRINTLABELLPCARTSEALS));
                    _printLpCartSealsBtn.Click();
                    WaitForLoad();
                    break;
                case PrintType.SpecialMeals:
                    _printSpecialMealsBtn = WaitForElementIsVisible(By.XPath(PRINT_SPECIALMEALS_BTN), nameof(PRINT_SPECIALMEALS_BTN));
                    _printSpecialMealsBtn.Click();
                    WaitForLoad();
                    _printSpecialMealsBtn = WaitForElementIsVisible(By.Id(CREATEPRINTLABELLPCARTSEALS));
                    _printSpecialMealsBtn.Click();
                    WaitForLoad();
                    break;
                default:
                    break;
            }

            if (versionPrint)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-pdf']"));
                ClosePrintButton();
            }

            //Results are opened in a new tab, switch the driver to the newly created one
            WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);
            return new PrintReportPage(_webDriver, _testContext);
        }
        public bool ExportMenuIsClicked()
        {
            if (isElementVisible(By.Id("btn-export-guestAndService-file-excel")))
                return true;
            return false;
        }
        public bool PrintMenuIsClicked()
        {
            //to be changed by id
            if (isElementVisible(By.Id("btn-print-pdf")))
                return true;
            return false;
        }
        public void ExportFlights(ExportType exportType, bool versionPrint, string status = "", bool export = false, bool downloadsfile = true)
        {
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;

            ShowExtendedMenu();

            if (!ExportMenuIsClicked())
            {
                _exportExtendedButton = WaitForElementIsVisible(By.Id("Export-Group-btn"));
                _exportExtendedButton.Click();
            }
            WaitForLoad();

            switch (exportType)
            {
                case ExportType.ExportGuestAndServices:
                    _exportGuestAndServices = WaitForElementIsVisible(By.Id(EXPORTGUESTANDSERVICES));
                    _exportGuestAndServices.Click();
                    WaitForLoad();
                    break;

                case ExportType.Export:
                    _export = WaitForElementIsVisible(By.Id(EXPORT));
                    _export.Click();
                    WaitForLoad();
                    break;

                case ExportType.ExportLogs:
                    _exportLog = WaitForElementIsVisible(By.Id(EXPORTLOG));
                    _exportLog.Click();
                    WaitForLoad();
                    break;

                case ExportType.ExportOP:
                    _exportOP = WaitForElementIsVisible(By.XPath(EXPORTOP), nameof(EXPORTOP));
                    _exportOP.Click();
                    WaitForLoad();

                    _confirmPrint = WaitForElementIsVisible(By.Id(CONFIRM_PRINT));
                    _confirmPrint.Click();
                    WaitForLoad();
                    break;

                case ExportType.ExportLoadingBob:
                    _exportLoadingBob = WaitForElementIsVisible(By.Id(EXPORTLOADINGBOB));
                    javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", _exportLoadingBob);
                    _exportLoadingBob.Click();
                    WaitForLoad();
                    break;

                case ExportType.ExportDockNumberFile:
                    _exportDOCKNUMBERFILE = WaitForElementIsVisible(By.XPath(EXPORTDOCKNUMBERFILE), nameof(EXPORTDOCKNUMBERFILE));
                    _exportDOCKNUMBERFILE.Click();
                    WaitForLoad();
                    break;

                case ExportType.BarsetStockExport:
                    _exportBarsetStock = WaitForElementIsVisible(By.XPath(EXPORTBARSETSTOCK), nameof(EXPORTBARSETSTOCK));
                    _exportBarsetStock.Click();
                    WaitForLoad();
                    if (!String.IsNullOrEmpty(status))
                    {
                        ComboBoxSelectById(new ComboBoxOptions("SelectedFlightStatus_ms", (string)status, false));
                        WaitForLoad();
                    }
                    break;

                default:
                    break;
            }

            if (versionPrint)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));
                ClickPrintButton();
            }

            if (export)
            {
                WaitForLoad();
                var exportbarsetstock = WaitForElementIsVisible(By.XPath("//*[@id=\"btn-export-barsetstock\"]"));
                exportbarsetstock.Click();
                WaitForLoad();

                try
                {
                    IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));
                    ClickPrintButton();
                }
                catch (Exception ex)
                {
                    return;
                }
            }

            if (downloadsfile)
            {
                WaitForDownload();
                //Close();
            }

        }

        public string GetMessagebarsetstock()
        {
            WaitForLoad();
            if (isElementVisible(By.XPath("//*[@id=\"no-selection-error\"]/span")))
            {
                var isViewChanged = WaitForElementExists(By.XPath("//*[@id=\"no-selection-error\"]/span"));
                return isViewChanged.Text;
            }
            else
            {
                return "";
            }
        }
        public void ImportFlights(ImportType importType, string filePath)
        {
            ShowExtendedMenu();
            _importExtendedButton = WaitForElementIsVisible(By.Id("Import-Group-btn"));
            _importExtendedButton.Click();
            WaitForLoad();
            switch (importType)
            {
                case ImportType.ImportDockNumbersFile:
                    _importDockNumberFile = WaitForElementIsVisible(By.XPath(IMPORTDOCKNUMBERFILE), nameof(IMPORTDOCKNUMBERFILE));
                    _importDockNumberFile.Click();
                    WaitForLoad();
                    WaitPageLoading();
                    break;
            }
        }
        public void OtherFlights(OtherFlightType otherFlightType)
        {
            ShowExtendedMenu();
            _othersExtendedButton = WaitForElementIsVisible(By.Id(OTHERS_GROUP_BTN));
            _othersExtendedButton.Click();
            WaitForLoad();
            switch (otherFlightType)
            {
                case OtherFlightType.ValidateAll:
                    _validateAllBtn = WaitForElementIsVisible(By.XPath(VALIDATEALL_BTN), nameof(VALIDATEALL_BTN));
                    _validateAllBtn.Click();
                    WaitForLoad();
                    WaitPageLoading();
                    _confirm = WaitForElementIsVisible(By.Id(VALIDE));
                    _confirm.Click();
                    
                    break;
            }
        }


        public FileInfo GetExportExcelFile(FileInfo[] taskFiles)
        {
            List<FileInfo> correctDownloadFiles = new List<FileInfo>();

            StringBuilder sb = new StringBuilder();

            foreach (var file in taskFiles)
            {
                sb.Append(file.Name + " ");
                //  Test REGEX
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
            Regex r = new Regex("[export?flights\\s\\d.-]", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = r.Match(filePath);

            return m.Success;
        }

        // Views
        public void GetCalendarView()
        {
            _calendarView = WaitForElementIsVisible(By.XPath(CALENDAR_VIEW), nameof(CALENDAR_VIEW));
            _calendarView.Click();
            WaitForLoad();
        }

        public void ClickCustomerBob()
        {
            _customerBobBtn = WaitForElementIsVisible(By.Id(CUSTOMERBOB_BTN), nameof(CUSTOMERBOB_BTN));
            _customerBobBtn.Click();
            WaitPageLoading();
        }

        public bool IsViewChanged()
        {
            WaitForLoad();
            if (isElementVisible(By.XPath(CALENDAR_VIEW_ACTIVATED)))
            {
                var isViewChanged = WaitForElementExists(By.XPath(CALENDAR_VIEW_ACTIVATED));

                if (isViewChanged.GetAttribute("class") != "active")
                {
                    return false;
                }
                return true;
            }
            else
            {
                return false;
            }
        }

        public string GetDateState()
        {
            _dateLabel = WaitForElementIsVisible(By.Id(DATE_LABEL_ID));
            return _dateLabel.GetAttribute("value");
        }

        public void SetDateState(DateTime date)
        {
            _dateLabel = WaitForElementIsVisible(By.Id(DATE_LABEL_ID));
            ////*[@id="start-date-picker-detail"]
            _dateLabel.SetValue(ControlType.DateTime, date);
            _dateLabel.SendKeys(Keys.Enter);
            WaitPageLoading();
            WaitForLoad();
        }

        public void SetBobTop()
        {
            _customerBobBtn = WaitForElementIsVisible(By.Id(CUSTOMERBOB_BTN), nameof(CUSTOMERBOB_BTN));
            _customerBobBtn.Click();
            WaitPageLoading();
            WaitForLoad();
        }


        // Tableau
        public void CliCkOnFirstFlight(int i)
        {
            _firstFlight = WaitForElementIsVisible(By.XPath(string.Format(FIRST_FLIGHT, i)), nameof(FIRST_FLIGHT));
            _firstFlight.Click();
            WaitForLoad();
        }

        public void ConsultFirstFlight()
        {

            var _firstFlight = WaitForElementIsVisible(By.XPath(FIRSTFLIGHT));
            _firstFlight.Click();
            WaitForLoad();
        }

        public string GetFirstFlightNumber()
        {
            var count = CheckTotalNumber();
            Assert.IsTrue(count != 0, "no flights in list");
            var firstFlightNumber = WaitForElementExists(By.XPath("//*[@id=\"rowFlightBarset\"]//span[@class=\"readonly-input\"]"));
            return firstFlightNumber.GetAttribute("innerText").Trim();
        }


        public string GetFirstFlightAircraft()
        {
            WaitForLoad();

            _flightAircraft = WaitForElementIsVisible(By.Id(FIRST_FLIGHT_AIRCRAFT));
            return _flightAircraft.Text;
        }

        public string GetFirstFlightKey()
        {

            _flightKey = WaitForElementIsVisible(By.XPath("//*[@id=\"flightTable\"]/tbody/tr[2]/td[2]/span[1]/a"), nameof(FIRST_FLIGHT_KEY));

            return _flightKey.Text;
        }



        public TrolleyAdjustingPage ClickFirstFlightKey()
        {
            _flightKey = WaitForElementIsVisible(By.XPath("//*[@id=\"flightTable\"]/tbody/tr[2]/td[2]/span/a"), nameof(FIRST_FLIGHT_KEY));

            _flightKey.Click();
            WaitForLoad();

            //Results are opened in a new tab, switch the driver to the newly created one
            WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(30));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            Thread.Sleep(1000); //waiting to update
            WaitForLoad();

            return new TrolleyAdjustingPage(_webDriver, _testContext);
        }

        public void ModifyFlightName(string newFlightNumber)
        {
            CliCkOnFirstFlight(2);

            _flightNumberInput = WaitForElementIsVisible(By.XPath("//*/input[@name='item.Name']"));
            _flightNumberInput.SetValue(ControlType.TextBox, newFlightNumber);

            // Temps d'enregistrement de la modif
            Thread.Sleep(2000);
            WaitPageLoading();
            WaitForLoad();
        }

        public bool IsArrival(int i)
        {
            bool IsArrival;
            //le premier tr c'est les titres de colonne
            IsArrival = isElementVisible(By.XPath("//*[@id='flightTable']/tbody/tr[" + (i + 1) + "]/td[1]/span/i[contains(@class,'plane-arrival')]"));
            WaitForLoad();
            return IsArrival;
        }

        public bool IsValidated()
        {

            var btnValid = WaitForElementIsVisible(By.XPath(V_STATE), nameof(V_STATE));
            string className = btnValid.GetAttribute("class");

            if (className.Contains("state-valid") && className.Contains("active"))
            {
                return true;
            }

            return false;
        }

        public bool IsPreval()
        {

            var btnPreval = WaitForElementIsVisible(By.XPath(P_STATE), nameof(P_STATE));
            string className = btnPreval.GetAttribute("class");

            if (className.Contains("state-preval") && className.Contains("active "))
            {
                return true;
            }

            return false;
        }

        public void SetNewState(string status)
        {
            CliCkOnFirstFlight(1);

            switch (status)
            {
                case "P":
                    _pStateBtn = WaitForElementIsVisible(By.XPath(P_STATE), nameof(P_STATE));
                    _pStateBtn.Click();
                    WaitForLoad();
                    break;
                case "V":
                    _vStateBtn = WaitForElementIsVisible(By.XPath(V_STATE), nameof(V_STATE));
                    _vStateBtn.Click();
                    WaitForLoad();
                    break;
                case "I":
                    _iStateBtn = WaitForElementIsVisible(By.XPath(I_STATE), nameof(I_STATE));
                    _iStateBtn.Click();
                    WaitForLoad();
                    break;
                case "C":
                    _cStateBtn = WaitForElementIsVisible(By.XPath(C_STATE), nameof(C_STATE));
                    _cStateBtn.Click();
                    WaitForLoad();
                    break;
                case "":
                    _pStateBtn = WaitForElementIsVisible(By.XPath(P_STATE), nameof(P_STATE));
                    _pStateBtn.Click();
                    WaitForLoad();
                    break;
                default:
                    break;
            }

            Thread.Sleep(5000);
            WaitForLoad();
        }

        public void UnSetNewState(string status)
        {
            CliCkOnFirstFlight(1);

            switch (status)
            {
                case "P":
                    _pStateBtn = WaitForElementIsVisible(By.XPath(P_STATE), nameof(P_STATE));
                    _pStateBtn.Click();
                    break;
                case "V":
                    _vStateBtn = WaitForElementIsVisible(By.XPath(V_STATE), nameof(V_STATE));
                    _vStateBtn.Click();
                    break;
                case "I":
                    _iStateBtn = WaitForElementIsVisible(By.XPath(I_STATE), nameof(I_STATE));
                    _iStateBtn.Click();
                    break;
                case "C":
                    _cStateBtn = WaitForElementIsVisible(By.XPath(C_STATE), nameof(C_STATE));
                    _cStateBtn.Click();
                    if (_cStateBtn.GetAttribute("class").Contains("active"))
                    {
                        _cStateBtn.Click();
                        IJavaScriptExecutor js = (IJavaScriptExecutor)_webDriver;
                        string newClassValue = "cancel-flight";
                        string script = $"arguments[0].setAttribute('class', '{newClassValue}');";
                        js.ExecuteScript(script, _cStateBtn);
                    }
                    break;
                case "":
                    _pStateBtn = WaitForElementIsVisible(By.XPath(P_STATE), nameof(P_STATE));
                    _pStateBtn.Click();
                    Thread.Sleep(1000);
                    break;
                default:
                    break;
            }

            WaitPageLoading();
            WaiForLoad();
        }

        public bool GetFirstFlightStatus(string status)
        {
            switch (status)
            {
                case "P":
                    _pStateBtn = WaitForElementIsVisible(By.XPath(P_STATE), nameof(P_STATE));
                    if (_pStateBtn.GetAttribute("class").Contains("state-preval") && _pStateBtn.GetAttribute("class").Contains("active"))
                    {
                        return true;
                    }
                    break;
                case "V":
                    _vStateBtn = WaitForElementIsVisible(By.XPath(V_STATE), nameof(V_STATE));
                    if (_vStateBtn.GetAttribute("class").Contains("state-valid") && _vStateBtn.GetAttribute("class").Contains("active"))
                    {
                        return true;
                    }
                    break;
                case "I":
                    _iStateBtn = WaitForElementIsVisible(By.XPath(I_STATE), nameof(I_STATE));
                    if (_iStateBtn.GetAttribute("class").Contains("state-invoice") && _iStateBtn.GetAttribute("class").Contains("active"))
                    {
                        return true;
                    }
                    break;
                default:
                    return false;
            }
            WaitForLoad();
            return false;
        }

        public FlightDetailsPage EditFirstFlight(string flightNumber)
        {
            // EDITABLE
            // parfois non éditable (pas de crayon edit) - plane-arrimval
            //WaitPageLoading();
            _editFlightBtn = WaitForElementIsVisibleNew(By.XPath(EDIT_FLIGHT));
            new Actions(_webDriver).MoveToElement(_editFlightBtn).Click().Perform();
            WaitPageLoading();
            WaitForLoad();
            return new FlightDetailsPage(_webDriver, _testContext);
        }

        public FlightDetailsPage EditFlightByName(string flightName)
        {
            var flight = WaitForElementIsVisible(By.XPath(string.Format("//*[@id=\"item_Name\" and contains(@value, '{0}')]/../../td[22]/a", flightName)));

            flight.Click();
            WaitPageLoading();
            WaitForLoad();

            return new FlightDetailsPage(_webDriver, _testContext);
        }

        public bool IsFlightDisplayed()
        {
            WaitPageLoading();
            if (isElementVisible(By.XPath(FIRST_FLIGHT)))
            {
                return true;
            }
            else
            {
                return false;
            }
        }


        public bool IsFlightExist()
        {
            if (CheckTotalNumber() >= 1) return true;

            return false;
        }

        public bool IsMultiFlightAdded(string flightNumber)
        {
            Filter(FlightPage.FilterType.SearchFlight, flightNumber);

            SetDateState(DateUtils.Now);
            bool day1 = CheckTotalNumber() == 1;

            SetDateState(DateUtils.Now.AddDays(1));
            bool day2 = CheckTotalNumber() == 1;

            SetDateState(DateUtils.Now.AddDays(2));
            bool day3 = CheckTotalNumber() == 1;

            SetDateState(DateUtils.Now.AddDays(3));
            bool day4 = CheckTotalNumber() == 1;

            SetDateState(DateUtils.Now.AddDays(4));
            bool day5 = CheckTotalNumber() == 1;

            SetDateState(DateUtils.Now.AddDays(5));
            bool day6 = CheckTotalNumber() == 1;

            SetDateState(DateUtils.Now.AddDays(6));
            bool day7 = CheckTotalNumber() == 1;

            return day1 && day2 && day3 && day4 && day5 && day6 && day7;
        }

        public void CreateXmlFileForImport(string path, string initFileName, string finalFileName)
        {
            string date = DateUtils.Now.ToString("D", CultureInfo.CreateSpecificCulture("en-US"));

            var dateFinale = date.Substring(date.IndexOf(",") + 1).Trim().Replace(",", "").Split(' ');

            XmlDocument docxml = new XmlDocument();
            docxml.Load(path + initFileName);

            var elemList = docxml.GetElementsByTagName("DateInFrequency");
            XmlElement elem = (XmlElement)elemList[0];

            // On modifie les attributs de la date
            elem.GetAttributeNode("Month").InnerText = dateFinale[0];
            elem.GetAttributeNode("Year").InnerText = dateFinale[2];
            elem.GetAttributeNode("Day").InnerText = dateFinale[1];

            //On sauvegarde la modification dans le fichier xml  
            docxml.Save(path + finalFileName);
        }

        public bool IsImportFilePresent(string path, string fileName)
        {
            DirectoryInfo taskDirectory = new DirectoryInfo(path);
            FileInfo[] taskFiles = taskDirectory.GetFiles(fileName);

            if (taskFiles.Length == 0)
                return false;

            return true;
        }

        public void DeleteXmlFile(string path, string fileName)
        {
            DirectoryInfo taskDirectory = new DirectoryInfo(path);
            FileInfo[] taskFiles = taskDirectory.GetFiles(fileName);

            foreach (var file in taskFiles)
            {
                file.Delete();
            }
        }

        public void WaiForLoad()
        {
            base.WaitForLoad();
        }

        public int CheckPdfNumberOfPage(FileInfo filePdf)
        {
            PdfDocument document = PdfDocument.Open(filePdf.FullName);
            List<string> mots = new List<string>();
            foreach (Page p in document.GetPages())
            {
                foreach (var mot in p.GetWords())
                {
                    mots.Add(mot.Text);
                }
            }
            return document.NumberOfPages;
        }
        public void CheckPrint(FileInfo filePdf)
        {
            int count = Math.Min(CheckTotalNumber(), 100);
            PdfDocument document = PdfDocument.Open(filePdf.FullName);
            List<string> mots = new List<string>();
            foreach (Page p in document.GetPages())
            {
                foreach (var mot in p.GetWords())
                {
                    mots.Add(mot.Text);
                }
            }

            var numberStrings = _webDriver.FindElements(By.Id("item_Name"));
            int counter = 0;
            foreach (var numberString in numberStrings)
            {
                if (numberString.GetAttribute("value") == "") continue;
                counter++;
                Assert.IsTrue(mots.Contains(numberString.GetAttribute("value")), "number string " + numberString.GetAttribute("value") + " non présent dans le PDF");
            }
            Assert.AreEqual(count, counter, "Différence entre les SupplierInvoices number string du tableau et du PDF");
        }
        public string GetSiteFilter()
        {
            return WaitForElementIsVisible(By.XPath(FILTER_SITE)).Text;
        }

        public bool SetSiteFilter(string DocSite)
        {
            var docComboBox = WaitForElementIsVisible(By.XPath(FILTER_SITE));

            int countdown = 30;
            while (docComboBox.Text != DocSite)
            {
                countdown--;
                if (countdown == 0)
                {
                    break;
                }
                docComboBox.SendKeys(Keys.ArrowDown);
            }
            WaitForLoad();
            return docComboBox.Text == DocSite;
        }

        public bool VerifySiteFrom(string site)
        {
            var sitesFrom = _webDriver.FindElements(By.XPath(SITES_FROM));
            foreach (var s in sitesFrom)
            {
                if (s.Text.Trim() != site)
                    return false;
            }
            return true;
        }
        public bool VerifySiteTo(string site)
        {
            var sitesTo = _webDriver.FindElements(By.XPath(SITES_TO));
            foreach (var s in sitesTo)
            {
                if (s.Text.Trim() != site)
                    return false;
            }
            return true;
        }
        public bool VerifierFlightNumberInPdf(FileInfo filePdf, string flightNo)
        {
            PdfDocument document = PdfDocument.Open(filePdf.FullName);
            Page page1 = document.GetPage(1);
            foreach (var word in page1.GetWords())
            {
                if (word.Text == flightNo)
                {
                    return true;
                }
            }
            return false;
        }

        public void RefreshInBoundInfo(string flight)
        {
            WaitForLoad();

            ShowExtendedMenu();
            _othersExtendedButton = WaitForElementIsVisible(By.XPath("//*[@id = 'Others-Group-btn']"));

            _othersExtendedButton.Click();
            WaitForLoad();

            var refreshInboundInfos = WaitForElementIsVisible(By.XPath("//*/a[text()='Refresh Inbound Infos']"));
            refreshInboundInfos.Click();
            var flightInput = WaitForElementIsVisible(By.Id("flight-number-pattern"));
            flightInput.SetValue(ControlType.TextBox, flight);
            var updateButton = WaitForElementIsVisible(By.Id("refresh-flights-data"));
            updateButton.Click();
            WaitPageLoading();
            var message = WaitForElementIsVisible(By.Id("updates-container"));
            Assert.AreEqual("No updates.", message.Text);
        }

        public void SetDockNumber(string dockNumber)
        {
            var docNumber = WaitForElementIsVisible(By.Id("item_DockNumber"));
            docNumber.SetValue(ControlType.TextBox, dockNumber);
            docNumber.SendKeys(Keys.Enter);
            WaitPageLoading();
            WaitForLoad();
        }
        public void RemoveAllFromSelectedDeliveryNoteReport()
        {
            WaitPageLoading();
            var listTo = _webDriver.FindElements(By.XPath(DELIVERY_NOTES_REPORTS));
            foreach (var delivery in listTo)
            {
                delivery.Click();
                var remove = WaitForElementIsVisible(By.XPath(REMOVE_DELIVERY_NOTES_REPORT));
                remove.Click();
            }
            //save
            var btnSave = WaitForElementIsVisible(By.XPath(SAVE_BTN));
            btnSave.Click();
        }

        public void SelectDeliveryNoteReport(string deliveryNoteReport)
        {
            WaitPageLoading();
            var listFrom = _webDriver.FindElements(By.XPath(FROM_DELIVERY_NOTES_REPORTS));
            foreach (var report in listFrom)
            {
                if (report.GetAttribute("value").Equals(deliveryNoteReport))
                {
                    report.Click();
                    var add = WaitForElementIsVisible(By.XPath(ADD_DELIVERY_NOTES_REPORT));
                    add.Click();
                }
            }
            //save
            var btnSave = WaitForElementIsVisible(By.XPath(SAVE_BTN));
            btnSave.Click();
        }
        public void ShowStatusFilter()
        {
            var btnStatus = WaitForElementIsVisible(By.XPath(SHOW_STATUS));
            btnStatus.Click();
        }

        public bool HasGuest(string flightNo, List<string> mots)
        {
            Filter(FilterType.SearchFlight, flightNo);
            FlightDetailsPage details = this.EditFirstFlight(flightNo);
            int count = details.CountServices();
            // flight dans NonPDF ACE1468162870
            //int countWithoutPrice = details.CountServiceWithoutPrice();
            // flight pas dans PDF 1780499352
            /*if (count > 0)
            {
                List<string> guests = details.GetGuestTypes();
                foreach (var guest in guests)
                {
                    details.SelectGuestType(guest);

                    List<string> services = details.GetServiceNames();

                    foreach (string s in services)
                    {
                        foreach (var mot in s.Split(' '))
                        {
                            mots.Add(mot);
                        }
                    }

                    foreach (var mot in guest.Split(' '))
                    {
                        mots.Add(mot);
                    }
                }
            }*/
            details.CloseViewDetails();
            //return (count - countWithoutPrice) > 0;
            return count > 0;
        }


        public bool HasService(string flightNo, List<string> mots = null)
        {
            Filter(FilterType.SearchFlight, flightNo);
            FlightDetailsPage details = this.EditFirstFlight(flightNo);
            List<string>  services = details.GetServiceNames();
            int count = details.CountServices();
            Assert.AreEqual(count, services.Count);

            foreach (var service in services)
            {
                if (mots == null)
                {
                    details.CloseViewDetails();
                    return true;
                }
                foreach (var mot in mots)
                {
                    if (service == mot)
                    {
                        details.CloseViewDetails();
                        return true;
                    }
                }
            }

            details.CloseViewDetails();
            return false;
        }

        public void SetFirstGate(string gate)
        {
            // select first line editable
            var gateLine = WaitForElementIsVisible(By.XPath("//*[@id='flightTable']/tbody/tr[not(contains(@class,'first-line') or contains(@class,'inactive'))]"));
            gateLine.Click();
            WaitForLoad();
            var gateInput = WaitForElementIsVisible(By.XPath(FIRST_GATE));
            gateInput.SetValue(ControlType.TextBox, gate);
            WaitForLoad();
        }
        public bool VerifyIsEditable(IWebElement line)
        {
            if (line.GetAttribute("class").Contains("selected line-selected"))
            {
                return true;
            }
            return false;
        }

        public void SetFirstDriver(string drive)
        {
            // select first line
            var driverLine = WaitForElementIsVisible(By.XPath("//*[@id='flightTable']/tbody/tr[2]"));
            driverLine.Click();
            WaitLoading();
            var driverInput = WaitForElementIsVisible(By.XPath("//*[@id='item_Teams']"));
            driverInput.SetValue(ControlType.TextBox, drive);
            WaitPageLoading();
        }

        public void MassiveDeleteMenus(string flightNumber, string siteFrom, string customer, bool selectAll = false)
        {
            ShowExtendedMenu();
            _othersExtendedButton = WaitForElementIsVisible(By.XPath("//*[@id = 'Others-Group-btn']"));
            _othersExtendedButton.Click();
            WaitForLoad();
            _massiveDelete = WaitForElementIsVisible(By.XPath(MASSIVE_DELETE));
            _massiveDelete.Click();
            WaitForLoad();
            var searchPattern = WaitForElementIsVisible(By.XPath("//*[@id=\"formMassiveDeleteFlight\"]/div/div[1]/div/input[@id=\"SearchPattern\"]"));
            searchPattern.SendKeys(flightNumber);
            ComboBoxSelectById(new ComboBoxOptions("SelectedSiteIds_ms", siteFrom, false));
            ComboBoxSelectById(new ComboBoxOptions("SelectedCustomersForDeletion_ms", customer, false));
            var searchMenusbtn = WaitForElementIsVisible(By.Id(SEARCH_MENUS_BTN));
            searchMenusbtn.Click();
            WaitForLoad();
            if (isElementVisible(By.XPath("//*[@id=\"tableFlights\"]/tbody/tr")))
            {
                if (selectAll)
                {
                    var selectAllResult = WaitForElementIsVisible(By.Id(SELECT_ALL_MASSIVE_DELETE));
                    selectAllResult.Click();
                }
                else
                {
                    var checkInput = WaitForElementIsVisible(By.XPath("//*[@id=\"tableFlights\"]/tbody/tr/td[1]"));
                    checkInput.Click();
                }

                WaitForLoad();
            }
            var deleteBtn = WaitForElementIsVisible(By.Id("deleteFlightsBtn"));
            deleteBtn.Click();
            WaitForLoad();
            var deleteConfirmBtn = WaitForElementIsVisible(By.Id("dataConfirmOK"));
            deleteConfirmBtn.Click();
            WaitLoading();
            WaitForLoad();
            var okButton = WaitForElementIsVisible(By.XPath("//*/button[text()='Ok']"));
            okButton.Click();
            WaitForLoad();
        }
        public void MultipleDelete(string customer, string siteFrom, object DateFrom, DateTime DateTo)
        {
            ShowExtendedMenu();
            _othersExtendedButton = WaitForElementIsVisible(By.XPath("//*[@id = 'Others-Group-btn']"));
            _othersExtendedButton.Click();
            WaitForLoad();
            _multipleDelete = WaitForElementIsVisible(By.XPath(MULTIPLE_DELETE));
            _multipleDelete.Click();
            WaitForLoad();

            ComboBoxSelectById(new ComboBoxOptions("SelectedCustomers_ms", customer, false));
            var site = WaitForElementIsVisible(By.Id("selected-site"));
            site.SetValue(ControlType.DropDownList, siteFrom);

            _Datefrom = WaitForElementIsVisible(By.Id(DATE_FROM));
            _Datefrom.SetValue(ControlType.DateTime, DateFrom);
            _Datefrom.SendKeys(Keys.Enter);

            _Dateto = WaitForElementIsVisible(By.Id(DATE_TO));
            _Dateto.SetValue(ControlType.DateTime, DateTo);
            _Dateto.SendKeys(Keys.Enter);

            WaitForLoad();

            var deleteBtn = WaitForElementIsVisible(By.XPath("//*[@id=\"modal-1\"]/div[3]/button[2]"));
            deleteBtn.Click();
            WaitForLoad();
        }

        public bool CheckStatus(bool active)
        {
            bool isActive = false;
            int tot = CheckTotalNumber() > 100 ? 100 : CheckTotalNumber();

            if (tot == 0)
                return false;

            for (int i = 0; i < tot; i++)
            {
                try
                {
                    _webDriver.FindElement(By.XPath(String.Format(SHOW_NOT_COUNTED, i + 1)));

                    if (active)
                        return false;
                }
                catch
                {
                    isActive = true;
                    if (!active)
                        return true;
                }
            }
            return isActive;
        }

        public void UnfoldAll()
        {
            ShowExtendedMenuFold();
            _unFoldAll = WaitForElementIsVisible(By.XPath(UNFOLD_ALL));
            _unFoldAll.Click();
            WaitForLoad();
        }

        public void FoldAll()
        {
            ShowExtendedMenuFold();
            _foldAll = WaitForElementIsVisible(By.XPath(FOLD_ALL));
            _foldAll.Click();
            WaitForLoad();
        }


        public Boolean IsUnfoldAll()
        {
            bool valueBool = false;

            var content = WaitForElementIsVisible(By.XPath(CONTENT));

            // Temps nécessaire pour que l'élément change de classe
            WaitPageLoading();
            if (content.GetAttribute("class") == "fas fa-angle-down")
                valueBool = true;

            return valueBool;
        }
        public Boolean IsDetail()
        {
            bool valueBool = false;

            var content = WaitForElementIsVisible(By.XPath(DETAILS));

            WaitPageLoading();
            if (content.GetAttribute("class") == "sub-table-col")
                valueBool = true;

            return valueBool;
        }

        public Boolean IsFoldAll()
        {
            bool valueBool = false;

            var content = WaitForElementExists(By.XPath(CONTENT));

            // Temps nécessaire pour que l'élément change de classe
            WaitPageLoading();

            if (content.GetAttribute("class") == "fas fa-chevron-right")
                valueBool = true;

            return valueBool;
        }

        public FlightMassiveDeleteModalPage ClickMassiveDelete()
        {
            ShowExtendedMenu();
            WaitLoading();
            _othersExtendedButton = WaitForElementIsVisible(By.Id(OTHERS_GROUP_BTN));
            _othersExtendedButton.Click();
            WaitForLoad();
            WaitForElementIsVisible(By.XPath(MASSIVE_DELETE)).Click();
            WaitForLoad();
            return new FlightMassiveDeleteModalPage(_webDriver, _testContext);
        }

        public void ScrollUntilElementIsVisible(By by)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)_webDriver;
            var element = WaitForElementExists(by);
            if (element != null)
            {
                js.ExecuteScript("arguments[0].scrollIntoView(true);", element);
            }
        }

        public void Go_To_New_Navigate()
        {
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.Close();
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[0]);

        }

        public int GetFirstFlightLeg()
        {
            int totalNbLegs = 0;
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;

            WaitForLoad();

            _flightLegNb = WaitForElementIsVisible(By.XPath("//*[@id=\"flightTable\"]/tbody/tr[2]/td[19]/div"));

            javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", _flightLegNb);

            if (!string.IsNullOrEmpty(_flightLegNb.Text))
            {
                totalNbLegs = int.Parse(_flightLegNb.Text);
            }

            return totalNbLegs;
        }

        public bool IsFlightNumberExist()
        {
            WaitPageLoading();
            if (isElementVisible(By.XPath(FIRST_FLIGHT_NUMBER)))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public string GetETD()
        {
            var _etd = WaitForElementIsVisible(By.XPath(ETD));
            return _etd.Text;
        }

        public string GetSearchValue()
        {
            var _searchInput = WaitForElementIsVisible(By.Id(SEARCH_INPUT));
            return _searchInput.Text;
        }
        public List<string> GetFlightList()
        {
            var FlightListId = new List<string>();

            var FlightsList = _webDriver.FindElements(By.XPath(FLIGHTS_LIST));

            foreach (var Flight in FlightsList)
            {
                FlightListId.Add(Flight.Text);
            }

            return FlightListId;
        }

        public int GetFirstIndexInvoiceAllowed()
        {
            var FlightListId = new List<string>();

            var FlightsInvoiceList = _webDriver.FindElements(By.XPath(FLIGHTS_LIST_INVOICE));
            var index = 0;
            foreach (var FlightInvoice in FlightsInvoiceList)
            {
                index++;
                if (FlightInvoice.GetAttribute("style") != "cursor: not-allowed;")
                {

                    return index;
                }
            }

            return -1;
        }

        public string GetState(int index)
        {
            var test = WaitForElementExists(By.XPath($"//*[@id=\"flightTable\"]/tbody/tr[{index + 1}]/td[20]/div/ul"));
            Actions act = new Actions(_webDriver);
            act.MoveToElement(test).Perform();
            // P
            // V
            // I
            if (WaitForElementIsVisible(By.XPath($"//*[@id=\"flightTable\"]/tbody/tr[{index + 1}]/td[20]/div/ul/li[3]")).GetAttribute("class").Contains("active"))
            {
                return "I";
            }
            if (WaitForElementIsVisible(By.XPath($"//*[@id=\"flightTable\"]/tbody/tr[{index + 1}]/td[20]/div/ul/li[2]")).GetAttribute("class").Contains("active"))
            {
                return "V";
            }
            if (WaitForElementIsVisible(By.XPath($"//*[@id=\"flightTable\"]/tbody/tr[{index + 1}]/td[20]/div/ul/li[1]")).GetAttribute("class").Contains("active"))
            {
                return "P";
            }
            return "";
        }

        public void ChangeStateInvoice(int index, int state)
        {
            var test = WaitForElementExists(By.XPath($"//*[@id=\"flightTable\"]/tbody/tr[{index + 1}]/td[20]/div/ul/li[{state}]"));
            Actions act = new Actions(_webDriver);
            act.MoveToElement(test).Click().Perform();
            WaitPageLoading();
            WaitForLoad();
        }
        public string GetIndexAirCraftValue()
        {
            _indexaircraftvalue = WaitForElementIsVisible(By.XPath(INDEXAIRCRAFTVALUE));
            return _indexaircraftvalue.Text;
        }
        public string GetIndexArrNumberValue()
        {
            var _indexarrnumbertvalue = WaitForElementIsVisible(By.XPath(INDEXARRNUMBERVALUE));
            return _indexarrnumbertvalue.Text;


        }
        public void SetArrNumberValue(string arrNumber)
        {
            var arrnumber = WaitForElementIsVisible(By.XPath(INDEXARRNUMBERVALUE));
            arrnumber.Click();
            WaitForLoad();
            var arrnumberInput = WaitForElementIsVisible(By.Id(INDEX_ARR_NUMBER_INPUT));
            arrnumberInput.SetValue(ControlType.TextBox, arrNumber);
            WaitForLoad();
        }
        public string GetIndexDockNumberValue()
        {
            _indexdocknumbervalue = WaitForElementIsVisible(By.XPath(INDEXDOCKNUMBERVALUE));
            return _indexdocknumbervalue.Text;
        }

        public bool IsResultFlightAlertsFiltreVisibile(string flightAlert)
        {

            var flightsAlerts = _webDriver.FindElements(By.XPath("/html/body/div[2]/div/div[2]/div/div[2]/div[1]/div[1]/table/tbody/tr[*]/td[1]/div/button"));
            var flightCount = flightsAlerts.Count();

            Actions a = new Actions(_webDriver);

            for (var i = 0; i < flightCount; i++)
            {
                var flightcol = WaitForElementIsVisible(By.XPath($"/html/body/div[2]/div/div[2]/div/div[2]/div[1]/div[1]/table/tbody/tr[{i + 2}]/td[1]/div/button[starts-with(@id,'flight-alerts-dropdown')]"));
                flightcol.Click();


                var flight = WaitForElementIsVisible(By.XPath($"//*[@id=\"flightTable\"]/tbody/tr[{i + 2}]/td[1]/div/div/p"));

                if (flight.Text.Trim() != flightAlert)
                    return false;
                a.MoveToElement(flightcol, 100, 150).Perform();
            }

            return true;
        }
        public void SetETD(string etd)
        {
            var firstFlight = WaitForElementIsVisible(By.XPath("//*[@id=\"flightTable\"]/tbody/tr[2]"));
            firstFlight.Click();
            var _etd = WaitForElementIsVisible(By.Id("item_ETD"));
            _etd.SendKeys(Keys.ArrowRight);
            _etd.SendKeys(Keys.ArrowRight);
            _etd.SendKeys(Keys.Backspace);
            _etd.SendKeys(Keys.Backspace);
            WaitForLoad();
            _etd.SendKeys(etd);
            _etd.SendKeys(Keys.ArrowRight);
            _etd.SendKeys(Keys.ArrowRight);
            WaitPageLoading();
            //_webDriver.Navigate().Refresh();
        }


        public void SetETA(string eta)
        {
            var firstFlight = WaitForElementIsVisible(By.XPath("//*[@id=\"flightTable\"]/tbody/tr[2]"));
            firstFlight.Click();
            var _eta = WaitForElementIsVisible(By.Id(ETA_INPUT));
            _eta.SendKeys(Keys.ArrowRight);
            _eta.SendKeys(Keys.ArrowRight);
            _eta.SendKeys(Keys.Backspace);
            _eta.SendKeys(Keys.Backspace);
            WaitForLoad();
            _eta.SendKeys(eta);
            _eta.SendKeys(Keys.ArrowRight);
            _eta.SendKeys(Keys.ArrowRight);
            WaitPageLoading();

        }

        public void SetAC(string ac)
        {
            //click First Flight
            WaitPageLoading();
            var firstFlight = WaitForElementIsVisible(By.XPath("//*[@id=\"flightTable\"]/tbody/tr[2]"));
            firstFlight.Click();

            //Set the element A/C
            var _ac = WaitForElementIsVisible(By.Id("item_AircraftName"));
            _ac.Click();
            _ac.SetValue(ControlType.TextBox, ac);
            WaitPageLoading();
            //_ac.Click();

            //choose from dropdown
            var _acDropdown = WaitForElementToBeClickable(By.XPath("//*[@id=\"flightTable\"]/tbody/tr[2]/td[11]/span/div/div/div[1]"));
            new Actions(_webDriver).MoveToElement(_acDropdown).Perform();
            _acDropdown.Click();

            WaitPageLoading();
            _webDriver.Navigate().Refresh();

        }

        public void SetDriver(string driver)
        {
            var firstFlight = WaitForElementIsVisible(By.XPath(FIRST_FLIGHT_NOTVALID));
            firstFlight.Click();
            Assert.IsTrue(VerifyIsEditable(firstFlight), "line was not editable");
            // pour info @Id = item_Teams => duplicated
            var _driver = WaitForElementIsVisible(By.XPath(FIRST_DRIVER_FLIGHT_NOTVALID));
            _driver.SetValue(ControlType.TextBox, driver);
            _driver.Click();
            WaitPageLoading();
            _webDriver.Navigate().Refresh();
        }

        public void SettDockNumber(string dock)
        {
            var firstFlight = WaitForElementIsVisible(By.XPath("//*[@id=\"flightTable\"]/tbody/tr[2]"));
            firstFlight.Click();
            var docknumberInputIsEditable = isElementVisible(By.Id("item_DockNumber"));
            Assert.IsTrue(docknumberInputIsEditable, "input is not editable");
            var docknumber = WaitForElementIsVisible(By.Id("item_DockNumber"));
            docknumber.Click();
            docknumber.SetValue(ControlType.TextBox, dock);
            WaitPageLoading();
            WaitForLoad();

            _webDriver.Navigate().Refresh();
        }


        public void SetGATE(string gate)
        {
            var _gate = WaitForElementIsVisible(By.XPath(GATE));
            _gate.Click();
            var gateValue = WaitForElementIsVisible(By.Id("item_Gate"));
            gateValue.Click();
            gateValue.SetValue(ControlType.TextBox, gate);
            WaitPageLoading();
        }
        public string GetETDValue()
        {
            var _etd = WaitForElementIsVisible(By.XPath("//*[@id=\"flightTable\"]/tbody/tr[2]/td[8]/span"));
            return _etd.Text;
        }
        public string GetIndexRegist(int index)
        {
            var element = WaitForElementExists(By.XPath($"//*[@id=\"flightTable\"]/tbody/tr[{index}]/td[12]/span"));
            var textIndexRegist = element.GetAttribute("innerText");
            // parfois "REG_999" donc pas un int
            return textIndexRegist;
        }
        public void ChangeIndexRegist(int index, string newIndexRegist)
        {
            ScrollUntilElementIsInView(By.XPath($"//*[@id=\"flightTable\"]/tbody/tr[{index}]"));
            WaitForElementExists(By.XPath($"//*[@id=\"flightTable\"]/tbody/tr[{index}]")).Click();
            WaitForLoad();
            var indexRegist = WaitForElementExists(By.XPath($"//*[@id=\"flightTable\"]/tbody/tr[{index}]/td[12]/input[contains(@class,'registrationTypeName')]"));
            indexRegist.SetValue(ControlType.TextBox, newIndexRegist);
            WaitPageLoading();
            WaitForLoad();
        }

        public string GetETAValue()
        {
            var _eta = WaitForElementIsVisible(By.XPath(ETA_VALUE));
            WaitPageLoading();
            return _eta.Text.Trim();
        }

        public string GetACValue()
        {
            var _ac = WaitForElementIsVisible(By.XPath("//*[@id=\"flightTable\"]/tbody/tr[2]/td[11]/span/span[2]"));
            return _ac.Text;
        }

        public string GetDriverValue()
        {
            var _driver = WaitForElementIsVisible(By.XPath(DRIVER_VALUE));
            return _driver.Text.Trim();
        }

        public void SetIndexDockNumberValue(string dockNumber)
        {
            var docknumber = WaitForElementIsVisible(By.XPath(INDEXDOCKNUMBERVALUE));
            docknumber.Click();
            WaitForLoad();
            var docknumberInput = WaitForElementIsVisible(By.Id(INDEX_DOCK_NUMBER_INPUT));
            docknumberInput.SetValue(ControlType.TextBox, dockNumber);
            WaitForLoad();
        }
        public HedomadairePage GoToCalendarView()
        {
            _calendarView = WaitForElementIsVisible(By.XPath(CALENDAR_VIEW), nameof(CALENDAR_VIEW));
            _calendarView.Click();

            WaitPageLoading();
            WaitForLoad();
            return new HedomadairePage(_webDriver, _testContext);
        }

        public void Refrech()
        {
            _webDriver.Navigate().Refresh();
        }
        public string GetFromValue()
        {
            var _from = WaitForElementIsVisible(By.XPath("//*[@id=\"flightTable\"]/tbody/tr[2]/td[5]/span"));
            return _from.Text;
        }

        public string GetToValue()
        {
            var _to = WaitForElementIsVisible(By.XPath("//*[@id=\"flightTable\"]/tbody/tr[2]/td[6]/span"));
            WaitPageLoading();
            return _to.Text;
        }

        public void SetFrom(string from)
        {
            var firstFlight = WaitForElementIsVisible(By.XPath("//*[@id=\"flightTable\"]/tbody/tr[2]"));
            firstFlight.Click();

            var _from = WaitForElementIsVisible(By.Id("item_AirportFromName"));
            _from.Click();
            WaitForLoad();
            _from.SetValue(ControlType.TextBox, "");

            _from.SendKeys(Keys.ArrowRight);
            _from.SendKeys(Keys.ArrowRight);
            _from.SendKeys(Keys.Backspace);
            _from.SendKeys(Keys.Backspace);
            _from.SendKeys(from);
            _from.SendKeys(Keys.ArrowRight);
            _from.SendKeys(Keys.ArrowRight);

            _from.Click();
            WaitForLoad();
            WaitPageLoading();
            _webDriver.Navigate().Refresh();
        }
        public void SetTo(string to)
        {
            var firstFlight = WaitForElementIsVisible(By.XPath("//*[@id=\"flightTable\"]/tbody/tr[2]"));
            firstFlight.Click();
            var _to = WaitForElementIsVisible(By.Id("item_AirportToName"));
            _to.Click();
            _to.SetValue(ControlType.TextBox, "");
            _to.SendKeys(Keys.ArrowRight);
            _to.SendKeys(Keys.ArrowRight);
            _to.SendKeys(Keys.Backspace);
            _to.SendKeys(Keys.Backspace);
            _to.SendKeys(to);
            _to.SendKeys(Keys.ArrowRight);
            _to.SendKeys(Keys.ArrowRight);
            _to.Click();
            TimeSpan.FromSeconds(30);
            WaitPageLoading();

            WaitForLoad();

        }
        public string GetIndexFlightsNumber()
        {
            _indexdocknumbervalue = WaitForElementIsVisible(By.XPath(INDEXFLIGHTSNUMBER));
            return _indexdocknumbervalue.Text.Trim();
        }

        public void ScrollUntilDockNumberFilterIsVisible()
        {
            ScrollUntilElementIsInView(By.Id("SelectedDockNumbers_ms"));
        }
        public bool CheckDownloadsExists()
        {
            if (!isElementExists(By.XPath(DOWNLOAD_PANEL)))
            {
                _downloadButton = WaitForElementIsVisible(By.Id(DOWNLOAD_BUTTON));
                _downloadButton.Click();
            }
            WaitForLoad();
            return isElementExists(By.XPath(DOWNLOADS_FIRST_ROW));
        }
        public void OpenDownloadPanel()
        {
            _downloadButton = WaitForElementIsVisible(By.Id(DOWNLOAD_BUTTON));
            _downloadButton.Click();
            WaiForLoad();
        }
        public bool CheckPanelPosition()
        {
            _exportPanel = WaitForElementIsVisible(By.XPath(DOWNLOAD_PANEL));

            string rightValue = _exportPanel.GetCssValue("right");
            string leftValue = _exportPanel.GetCssValue("left");
            string topValue = _exportPanel.GetCssValue("top");

            return rightValue == "0px" && leftValue == "0px" && topValue == "0px";
        }

        public string GetGateValue()
        {
            var _gate = WaitForElementIsVisible(By.XPath("//*[@id=\"flightTable\"]/tbody/tr[2]/td[10]/span"));
            return _gate.Text;
        }
        public FlightDetailsPage GoToFlightDetailsPage()
        {
            _firstLine = WaitForElementIsVisible(By.XPath("//*[@id=\"flightTable\"]/tbody/tr[2]/td[22]/a"));
            _firstLine.Click();
            WaitForLoad();
            return new FlightDetailsPage(_webDriver, _testContext);

        }
        public void ClickPrevalidate()
        {
            _preval = WaitForElementIsVisible(By.XPath(PREVAL));
            _preval.Click();
            WaitForLoad();



        }
        public int GetCountResult()
        {
            WaitForLoad();
            _totalNumber = WaitForElementExists(By.ClassName("counter"));
            int nombre = Int32.Parse(_totalNumber.Text);
            return nombre;
        }
        public bool CheckFightNumber(string flightnumber)
        {
            var list_flight = _webDriver.FindElements(By.XPath(LIST_FLIGHT));

            foreach (var element in list_flight)
            {
                if (element.Text == flightnumber)
                {
                    return true;  // Le vol a été trouvé
                }
            }

            return false;  // Le vol n'a pas été trouvé
        }

        public FlightEditModalPage GoToFlightEditModalPage()
        {
            _editButton = WaitForElementIsVisible(By.XPath(EDIT_BUTTON));
            _editButton.Click();
            return new FlightEditModalPage(_webDriver, _testContext);
        }
        public bool CheckChangedStatus(string status)
        {
            string elementXPath;
            if (status == "P")
            {
                elementXPath = PRIVAL_CLIC;
            }
            else if (status == "V")
            {
                elementXPath = VALIDATE_CLIC;
            }
            else
            {
                elementXPath = INVOICE_CLIC;
            }
            var element = WaitForElementIsVisible(By.XPath(elementXPath));
            string classAttribute = element.GetAttribute("class");
            return classAttribute != null && classAttribute.Contains("active");
        }

        public void MakeSureModalIsClosed()
        {
            var modalIsVisible = isElementVisible(By.Id("modal-1"));
            if (modalIsVisible)
            {
                var exitBtn = WaitForElementIsVisible(By.XPath("//*[@id=\"modal-1\"]/div[@class=\"modal-header\"]//button[@data-dismiss]"));
                exitBtn.Click();
            }
        }
        public int CheckTotalNumberOfLines()
        {
            WaitForLoad();

            var tableElement = _webDriver.FindElements(By.XPath("/html/body/div[5]/div/div/div[1]/table/tbody/tr/td[2]/div/div/table/tbody/tr/td[2]/div/div/div[2]/div/table/tbody/tr[*][contains(@class, 'flight-detail-lin')]"));

            int rows = tableElement.Count();


            return rows;
        }
        public LpCartPage GO_To_LPCART_Tab()
        {
            var lpCartTab = WaitForElementIsVisible(By.XPath(LP_CART_TAB));
            lpCartTab.Click();
            WaitPageLoading();
            return new LpCartPage(_webDriver, _testContext);
        }
        public FlightEditModalPage EditFlightRow(string flightNumber)
        {
            _editFlightBtn = WaitForElementIsVisible(By.XPath(EDIT_FIRST_FLIGHT_BTN));

            new Actions(_webDriver).MoveToElement(_editFlightBtn).Click().Perform();
            WaitPageLoading();
            WaitForLoad();

            return new FlightEditModalPage(_webDriver, _testContext);
        }



        public void Export(ExportType exportType)
        {
            ShowExtendedMenu();
            if (!ExportMenuIsClicked())
            {
                _exportExtendedButton = WaitForElementIsVisible(By.Id("Export-Group-btn"));
                _exportExtendedButton.Click();
            }
            WaitForLoad();
            switch (exportType)
            {
                case (ExportType.Export):
                    var _exportResults = WaitForElementIsVisible(By.Id(EXPORT));
                    _exportResults.Click();
                    WaitForLoad();

                    IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));

                    WaitForDownload();
                    WaitForLoad();
                    break;
            }

        }
        public bool IsFlightNameExist()
        {
            if (CheckTotalNumber() == 1)
            {
                return true;
            }
            return false;
        }

        public BarsetsPage ClickOnTop()
        {
            _clickOnTop = WaitForElementIsVisible(By.Id(CLICK_ON_TOP));
            _clickOnTop.Click();
            WaitForLoad();
            return new BarsetsPage(_webDriver, _testContext);
        }

        public bool IsArrowSplitExist()
        {
            return isElementVisible(By.XPath(ARROW_SPLIT));
        }

        public bool DetectClosingPeriod(string columnToCheck)
        {
            var columnElements = _webDriver.FindElements(By.CssSelector("ul.header-center li.menu-item"));

            List<string> columnNames = new List<string>();

            foreach (var column in columnElements)
            {
                var link = column.FindElement(By.TagName("a"));
                if (link.Text != "")
                {
                    columnNames.Add(link.Text);
                }
            }

            bool columnExists = columnNames.Contains(columnToCheck);
            return columnExists;



        }
        public void EditAircraftAC(string Aircraft)
        {
            var s = WaitForElementIsVisible(By.Id(FIRST_FLIGHT_AIRCRAFT));
            s.SetValue(ControlType.TextBox, Aircraft);

            var selectFirstItemC = WaitForElementToBeClickable(By.XPath(AIRCRAFT_INPUT));
            selectFirstItemC.Click();
            WaitForLoad();
            WaitPageLoading();
        }
        public string GetIndexDriverValue()
        {
            var _driver = WaitForElementIsVisible(By.XPath(DRIVER_ITEM));
            return _driver.Text;
        }

        public bool AfficheFlightValid(List<string> mots, string flightNumber, string flightNumber1)
        {
            return (mots.Any(a=>a.Contains(flightNumber)) && (mots.Any(a => a.Contains(flightNumber1))));

        }
        private void WaitForModalToClose()
        {
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(10));
            wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.InvisibilityOfElementLocated(By.ClassName("modal-backdrop")));
        }
        public PrintReportPage ClickPrinterIcon()
        {
            WaitForLoad();
            WaitForModalToClose(); // Wait for the modal to close
            _printButton = _webDriver.FindElement(By.Id(PRINTBUTTON));
            _printButton.Click();
            return new PrintReportPage(_webDriver, _testContext);
        }

        public bool VerifyETDInRange(string minValue, string maxValue)
        {
            DateTime minTime = DateTime.ParseExact(minValue, "hhmmtt", System.Globalization.CultureInfo.InvariantCulture);
            DateTime maxTime = DateTime.ParseExact(maxValue, "hhmmtt", System.Globalization.CultureInfo.InvariantCulture);

            // Récupérer tous les éléments de la colonne sauf le premier
            var elements = _webDriver.FindElements(By.XPath("//*[@id='flightTable']/tbody/tr[position() > 1]/td[8]/span"));

            foreach (var element in elements)
            {
                    DateTime elementTime = DateTime.ParseExact(element.Text, "HH:mm", System.Globalization.CultureInfo.InvariantCulture);
                    if (elementTime < minTime || elementTime > maxTime)
                    {
                        return false; // Retourne false si une valeur n'est pas dans l'intervalle
                    }
            }
            return true;
        }

        public void ShowStatusCancelled()
        {
            var status = _webDriver.FindElement(By.XPath(C_STATE));
            status.Click();
            WaitPageLoading();
        }
   

    }
}
