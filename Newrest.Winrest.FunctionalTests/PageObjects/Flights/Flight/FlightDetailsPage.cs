using DocumentFormat.OpenXml.Bibliography;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Flights.Flight
{
    public class FlightDetailsPage : PageBase
    {
        public FlightDetailsPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // _________________________________________ Constantes ________________________________________

        // General
        private const string CLOSE_BTN = "viewDetails-close";
        private const string FLIGHT_NUMBER = "//*[@id=\"flightDetails\"]/div/h4";
        private const string CHANGE_FLIGHT_NUMBER = "//*[@id=\"flights-list\"]/tbody/tr[*]/td[2]/div/a[normalize-space(text())!=\"{0}\"]";
        private const string FIRST_FLIGHT_NUMBER = "/html/body/div[5]/div/div/div[2]/div/div/table/tbody/tr[{0}]/td[2]/div/a";

        // Informations vols
        private const string INFO_FLIGHT_EXTENDED_BTN = "dropdownBtn";

        private const string IMPORT_GUEST_AND_SERVICES_DEV = "importGuestAndServices";

        private const string CHOOSE_FILE = "fileSent";
        private const string VALIDATE_IMPORT = "//*[@id=\"ImportFileForm\"]/div[2]/button[2]";
        private const string OK_IMPORT = "//*[@id=\"modal-1\"]/div/div/div[3]/button";
        private const string DUPLICATE_FLIGHT = "duplicateFlights";
        private const string DUPLICATE_FLIGHT_NUMBER = "NewFlightNumber";
        private const string VALIDATE_DUPLICATE = "last";
        private const string REPLACE_GUEST = "replaceGuest";
        private const string INITIAL_GUEST = "GuestToReplace";
        private const string NEW_GUEST = "GuestToSelect";
        private const string CONFIRM_REPLACE = "confirmReplace";
        private const string UPDATE_GUEST_QTY = "updateGuestQuantities";
        private const string INPUT_GUEST_QTY = "item_NbGuestTypes";
        private const string CONFIRM_UPDATE_QTY = "btn-update-submit";

        private const string FLIGHT_HISTORY = "flightHistory";

        private const string HISTORY_ROW1 = "//*[@id=\"modal-1\"]/div/div/div[2]/table/tbody/tr[2]";
        private const string HISTORY_ROW2 = "//*[@id=\"modal-1\"]/div/div/div[2]/table/tbody/tr[3]";
        private const string HISTORY_ROW1_DEV = "//*[@id=\"modal-1\"]/div[2]/table/tbody/tr[2]";
        private const string HISTORY_ROW2_DEV = "//*[@id=\"modal-1\"]/div[2]/table/tbody/tr[3]";

        private const string SPLIT_FLIGHT_ROUTE = "splitFlightRoute";

        private const string VALIDATE_SPLIT = "last";

        private const string EDIT_FLIGHT = "greenEditBtn";

        private const string MODAL_FLIGHT_DATE = "FlightDate";
        private const string MODAL_FLIGHT_NUMBER = "Name";
        private const string SAVE_FLIGHT_ONLY = "//*[@id=\"form-createdit-flight\"]/div[2]/button[3]";
        private const string P_STATE = "prevalPopUp";
        private const string P_STATE_SELECTED = "//*[@id=\"statuss\"]/li[1][@class='active state-preval-PopUp']";
        private const string P_STATE_UNSELECTED = "//*[@id=\"statuss\"]/li[1][@class=' state-preval-PopUp']";
        private const string V_STATE = "validatePopup";
        private const string V_STATE_SELECTED = "//*[@id=\"statuss\"]/li[2][@class='active state-valid-PopUp']";
        private const string V_STATE_UNSELECTED = "//*[@id=\"statuss\"]/li[2][@class=' state-valid-PopUp']";
        private const string I_STATE = "invoicePopup";
        private const string I_STATE_SELECTED = "//*[@id=\"statuss\"]/li[3][@class='active state-invoice-PopUp']";
        private const string I_STATE_UNSELECTED = "//*[@id=\"statuss\"]/li[3][@class=' state-valid-PopUp']";

        private const string COMMENT_TEXTAREA_1 = "InternalFlightRemarks";
        private const string COMMENT_TEXTAREA_2 = "DeliveryNoteComment";

        // Guests
        private const string ADD_GUEST_BTN = "//a[text()='Add guest type']";
        private const string GUEST_NAME = "GuestTypeId";
        private const string CREATE_GUEST_BUTTON = "//*[@id=\"form-create-guest-type\"]/div[4]/button[2]";
        private const string GUEST_TYPE = "//*[@id=\"leg-list\"]/tbody/tr/td[1 and contains(@class,'guest-type-name')]";
        private const string GUEST_TYPE_SELECT = "//*[@id='leg-list']/tbody/tr/td[1 and contains(@class,'guest-type-name') and text()='{0}']";
        private const string GUESTTYPE_QTY = "NbGuestTypes";
        private const string REMOVE_GUEST = "//span[contains(@class, 'remove')]";

        private const string LEG_NAME = "//*[@id=\"leg-list\"]/tbody/tr[1]/td/div[contains(@class,'header-detail-box-leg')]";
        private const string DELIVERY_NUM = "//*[@id=\"flightDetails\"]/div/h4";

        //Lpcart
        private const string LPCART = "dropdown-lpcart-selectized";

        //Aircraft
        private const string AIRCRAFT_NUMBER = "dropdown-aircraft-details-selectized";

        // Service
        private const string ADD_SERVICE_MENU = "//*/span[@class='btn-icon btn-icon-add-white']";
        private const string ADD_SERVICE = "//a[text()='add service']";
        private const string ADD_GENERIC_SERVICE = "//*[@id=\"leg-details-with-services\"]/div/div/div/div/a[2]";
        private const string CREATE_GENERIC_SERVICE = "//*[@id=\"leg-details-with-services\"]/div/div/div/div/a[3]";
        private const string GENERIC_SERVICE_NAME = "first";
        private const string SAVE_GENERIC_SERVICE = "last";

        private const string GENERIC_SERVICE = "//*[@id=\"services-details\"]/div/table/tbody/tr[3]/td[2]/span";
        private const string SERVICE_CONTAINER = "//*[@id=\"ServiceIdContainer_0\"]/div/div[2]/div/div[1]";
        private const string SERVICE_SPAN = "//*[@id=\"ServiceIdContainer_0\"]/span";
        private const string SECOND_SERVICE_LINE_INPUT = "//*[@id=\"ServiceIdContainer_0\"]/div/div[1]/div";

        private const string SERVICE_LINE = "//*[@id=\"ServiceIdContainer_0\"]/div/div[1]";
        private const string SERVICE_LINE_INPUT = "//*[@id=\"ServiceIdContainer_0\"]/div/div[1]/input";
        private const string SERVICE_LINE_FROM_BOB = "//tr[contains(@id, 'service')]";
        private const string FIRST_SERVICE = "//*[@id=\"ServiceIdContainer_0\"]/div/div[2]/div/div[1]";
        private const string ICON_SAVE = "//*/span[@class='save-state btn-icon btn-icon-status-save']";

        private const string SERVICE_EXTENDED_MENU = "[class = 'btn mini-btn open-service-menu']";
        private const string SERVICE_EXTENDED_MENU_ID = "flights-flight-details-service-extended-btn";
        private const string DELETE_SERVICE = "//*[starts-with(@id,\"service-menu-\")]/li[*]/a[@class='d-flex justify-content-center align-items-center btn-delete-serviceFV']";
        private const string DELETE_SERVICE_ID = "flights-flight-details-service-delete-btn";
        private const string CONFIRM_DELETE = "dataConfirmOK";

        private const string SERVICE_QUANTITY = "//*[@id=\"services-details\"]/div/table/tbody/tr[2]/td[4]";
        private const string SERVICENAME_VALIDE = "//*[starts-with(@id,\"service-\")]/td[1]/span";
        private const string SERVICENAME_INPUT = "/html/body/div[5]/div/div/div[1]/table/tbody/tr/td[2]/div/div/table/tbody/tr/td[2]/div/div/div[2]/div/table/tbody/tr[2]/td[1]/span";
        private const string SERVICENAME_READONLY_ALL = "//*/span[@class='readonly-text']";
        private const string SERVICENAME_READONLY = "//*/span[@class='readonly-text' and text()='{0}']";
        private const string QUANTITY_UPDATE_DISKETTE_ICON = "//span[@class='save-state btn-icon btn-icon-status-save']";
        private const string EXTEND_FLIGHT_MENU = "dropdownBtn";
        private const string PRINT_FLIGHT_LABELS_BUTTON = "//*[@id=\"form-createdit-flight-detail\"]/div[1]/div[1]/div/div/div/a[text()='Print Flights Labels']";
        private const string CONFIRM_PRINT_FLIGHT_LABELS = "//*/button[contains(text(),'Print (')]";
        private const string CLOSE_DETAILS_FLIGHTS_BTN = "viewDetails-close";
        private const string SERVICE_TAB = "//*[@id=\"services-details\"]/div/table/tbody/tr[*]/td[6]/a/span";
        private const string PREVALIDATE_ALL = "/html/body/div[5]/div/div/div[3]/div/div/div/form/div/div[2]/ul/li[5]/div/ul/li[1]/a";
        private const string VALIDATE_ALL = "//*[starts-with(@id,\"validation-menu\")]/li[2]/a";
        private const string INVOICE_ALL = "/html/body/div[5]/div/div/div[3]/div/div/div/form/div/div[2]/ul/li[5]/div/ul/li[3]/a";
        private const string EXTENDED_BTN = "//*[@id=\"filter-for-flights\"]/div/div[2]/ul/li[5]/div/a";
        private const string LIST_FLIGHT_DETAIL = "//*[@id=\"flights-list\"]/tbody/tr[*]/td[3]/div/a";
        private const string SEARCH_INPUT = "SearchPattern";
        private const string STATUS_FILTER = "SelectedStatusPopUpFilter_ms";
        private const string CHECK_ALL_STATUS = "/html/body/div[20]/div/ul/li[1]/a";
        private const string UNCHECK_ALL_STATUS = "/html/body/div[20]/div/ul/li[2]/a";
        private const string SEARCH_STATUS = "/html/body/div[20]/div/div/label/input";
        private const string IS_LPCART_APPLY_TO_FUTURE_Flights = "ApplyLPCartToFutureFlights";
        private const string APPLY_TO_FUTURE_FLIGHTS_MODAL_INFO = "//*[@id=\"dataAlertModal\"]/div";
        private const string CONFIRM_APPLY_TO_FUTURE_FLIGHTS = "//*[@id=\"dataAlertModal\"]/div/div/div[3]/button[contains(text(),'OK')]";
        private const string MESSAGE_APPLY_LPCART_MODAL = "//*[@id=\"dataAlertModal\"]/div/div/div[2]/p[contains(.,\"LPCart has been selected on\") and contains(.,\" future flight\")]";
        private const string VALIDE = "btn-validate";

        private const string SERVICE_HEADER = "//*[@id=\"services-details\"]/div/table/tbody/tr[1]/th[1]";
        private const string SERVICE_LINE_PATH = "(//*/div[contains(@class,'selectize-input')]/div)[last()]\"";
        private const string SERVICES_EXIST = "//tr[contains(@id,\"service-\")]";
        private const string SERVICES_ALREADY_EXIST = "(//*[contains(@id,\"ServiceIdContainer\")]/span)[last()]";
        private const string GUEST_TYPE_SELECTED = "//*[@id=\"leg-list\"]/tbody/tr[*]/td[contains(@class,\"leg-guest-type-name\")and contains(text(),{0})]";
        private const string SEVICESNAMES_ONPATCH = "//*[starts-with(@id,\"ServiceIdContainer\")]/div/div[1]/div";
        private const string SEVICESNAMES_READONLY_ONPATCH = "//*[starts-with(@id,\"service-\")]/td[1]/span";
        private const string SEARCH_FLIGHT_NUMBER = "/html/body/div[4]/div/div/div[2]/div/div/form/div[1]/div[1]/div[1]/input";
        private const string SERVICE_NAMES_LIST = "//*[starts-with(@id,\"service-\")]/td[1]/span";
        private const string CLOSE_MENU = "leg-details-with-services";
        private const string FLIGHT_DETAIL = "flightDetails";
        private const string SWAP = "//*[@id=\"form-createdit-flight-detail\"]/div[1]/div[1]/div/div/div/a[text()='SwapLPCart Report']";
        private const string FLIGHT_LEG = "FlightLegId";
        private const string PICO = "//*[starts-with(@id, 'service-')]/td[6]/a/span";

        private const string FLIGHT_TYPE = "dropdown-flight-type";
        private const string ARR_FLIGHT_NUMBER = "ArrFlightNo";
        private const string REGISTRATION = "tb-registration-type-details";
        private const string TABLET_FLIGHT_TYPE = "dropdown-tablet-flight-type";
        private const string INPUTFLIGHTNUMBEFILTER = "/html/body/div[4]/div/div/div[2]/div/div/form/div[1]/div[1]/div[1]/input";
        private const string INPUTSITEFILTER = "/html/body/div[4]/div/div/div[2]/div/div/form/div[1]/div[1]/div[2]/select";
        private const string DATA_SHEET = "datasheet-span";
        // ________________________________________ Variables ________________________________________________

        // General
        [FindsBy(How = How.XPath, Using = LEG_NAME)]
        private IWebElement _legname;

        [FindsBy(How = How.XPath, Using = DELIVERY_NUM)]
        private IWebElement _delivery_num;


        [FindsBy(How = How.XPath, Using = SERVICE_TAB)]
        private IWebElement _servicetab;

        [FindsBy(How = How.Id, Using = CLOSE_BTN)]
        private IWebElement _closeBtn;

        [FindsBy(How = How.Id, Using = LPCART)]
        private IWebElement _lpCart;

        [FindsBy(How = How.XPath, Using = FLIGHT_NUMBER)]
        private IWebElement _flightNumber;

        // Informations vols
        [FindsBy(How = How.Id, Using = INFO_FLIGHT_EXTENDED_BTN)]
        private IWebElement _infoFlightExtendedBtn;

        [FindsBy(How = How.Id, Using = IMPORT_GUEST_AND_SERVICES_DEV)]
        private IWebElement _importGuestAndServicesDev;

        [FindsBy(How = How.Id, Using = CHOOSE_FILE)]
        private IWebElement _chooseFile;

        [FindsBy(How = How.XPath, Using = VALIDATE_IMPORT)]
        private IWebElement _validateImport;

        [FindsBy(How = How.XPath, Using = OK_IMPORT)]
        private IWebElement _okImport;

        [FindsBy(How = How.Id, Using = DUPLICATE_FLIGHT)]
        private IWebElement _duplicateFlight;

        [FindsBy(How = How.Id, Using = DUPLICATE_FLIGHT_NUMBER)]
        private IWebElement _duplicateFlightNumber;

        [FindsBy(How = How.Id, Using = VALIDATE_DUPLICATE)]
        private IWebElement _validateDuplicate;

        [FindsBy(How = How.Id, Using = REPLACE_GUEST)]
        private IWebElement _replaceGuest;

        [FindsBy(How = How.Id, Using = INITIAL_GUEST)]
        private IWebElement _initialGuest;

        [FindsBy(How = How.Id, Using = NEW_GUEST)]
        private IWebElement _newGuest;

        [FindsBy(How = How.Id, Using = CONFIRM_REPLACE)]
        private IWebElement _confirmReplace;

        [FindsBy(How = How.Id, Using = UPDATE_GUEST_QTY)]
        private IWebElement _updateGuestQty;

        [FindsBy(How = How.Id, Using = INPUT_GUEST_QTY)]
        private IWebElement _inputGuestQty;

        [FindsBy(How = How.Id, Using = CONFIRM_UPDATE_QTY)]
        private IWebElement _confirmUpdateQty;

        [FindsBy(How = How.Id, Using = FLIGHT_HISTORY)]
        private IWebElement _historyFlight;

        [FindsBy(How = How.Id, Using = SPLIT_FLIGHT_ROUTE)]
        private IWebElement _splitFlightRoute;

        [FindsBy(How = How.Id, Using = VALIDATE_SPLIT)]
        private IWebElement _validateSplit;

        [FindsBy(How = How.Id, Using = EDIT_FLIGHT)]
        private IWebElement _editFlight;

        [FindsBy(How = How.Id, Using = MODAL_FLIGHT_DATE)]
        private IWebElement _modalFlightDate;

        [FindsBy(How = How.XPath, Using = MODAL_FLIGHT_NUMBER)]
        private IWebElement _modalFlightNumber;

        [FindsBy(How = How.XPath, Using = SAVE_FLIGHT_ONLY)]
        private IWebElement _saveFlightOnly;

        [FindsBy(How = How.Id, Using = P_STATE)]
        private IWebElement _pStateBtn;

        [FindsBy(How = How.Id, Using = V_STATE)]
        private IWebElement _vStateBtn;

        [FindsBy(How = How.Id, Using = I_STATE)]
        private IWebElement _iStateBtn;

        // Guests
        [FindsBy(How = How.XPath, Using = ADD_GUEST_BTN)]
        private IWebElement _addGuestBtn;

        [FindsBy(How = How.Id, Using = GUEST_NAME)]
        private IWebElement _guestName;

        [FindsBy(How = How.XPath, Using = CREATE_GUEST_BUTTON)]
        private IWebElement _createGuest;

        [FindsBy(How = How.XPath, Using = GUEST_TYPE)]
        private IWebElement _guestType;

        [FindsBy(How = How.Id, Using = GUESTTYPE_QTY)]
        private IWebElement _guestTypeQty;

        [FindsBy(How = How.XPath, Using = REMOVE_GUEST)]
        private IWebElement _removeGuest;

        //Aircraft
        [FindsBy(How = How.Id, Using = AIRCRAFT_NUMBER)]
        private IWebElement _aircraftNumber;


        // Service
        [FindsBy(How = How.XPath, Using = ADD_SERVICE_MENU)]
        private IWebElement _addServiceMenu;

        [FindsBy(How = How.XPath, Using = ADD_SERVICE)]
        private IWebElement _addService;

        [FindsBy(How = How.XPath, Using = ADD_GENERIC_SERVICE)]
        private IWebElement _addGenericService;

        [FindsBy(How = How.XPath, Using = CREATE_GENERIC_SERVICE)]
        private IWebElement _createGenericService;

        [FindsBy(How = How.Id, Using = GENERIC_SERVICE_NAME)]
        private IWebElement _genericServiceName;

        [FindsBy(How = How.Id, Using = SAVE_GENERIC_SERVICE)]
        private IWebElement _saveGenericServiceBtn;

        [FindsBy(How = How.XPath, Using = SERVICE_LINE)]
        private IWebElement _serviceLine;

        [FindsBy(How = How.XPath, Using = SERVICE_LINE_INPUT)]
        private IWebElement _serviceLineInput;

        [FindsBy(How = How.XPath, Using = FIRST_SERVICE)]
        private IWebElement _firstService;

        [FindsBy(How = How.Id, Using = SERVICE_EXTENDED_MENU_ID)]
        private IWebElement _serviceExtendedMenu;

        [FindsBy(How = How.CssSelector, Using = DELETE_SERVICE)]
        private IWebElement _deleteService;

        [FindsBy(How = How.Id, Using = CONFIRM_DELETE)]
        private IWebElement _confirmDelete;

        [FindsBy(How = How.XPath, Using = SERVICE_QUANTITY)]
        private IWebElement _serviceQty;

        [FindsBy(How = How.XPath, Using = SERVICENAME_INPUT)]
        private IWebElement _serviceName;
        [FindsBy(How = How.XPath, Using = VALIDATE_ALL)]
        private IWebElement _validateAll;
        [FindsBy(How = How.XPath, Using = PREVALIDATE_ALL)]
        private IWebElement _prevalidateAll;
        [FindsBy(How = How.XPath, Using = INVOICE_ALL)]
        private IWebElement _invoiceAll;
        [FindsBy(How = How.XPath, Using = EXTENDED_BTN)]
        private IWebElement _extendedButton;
        [FindsBy(How = How.XPath, Using = SEARCH_FLIGHT_NUMBER)]
        private IWebElement _searchInput;
        [FindsBy(How = How.Id, Using = CHECK_ALL_STATUS)]
        private IWebElement _checkAllStatus;
        [FindsBy(How = How.Id, Using = UNCHECK_ALL_STATUS)]
        private IWebElement _uncheckAllStatus;
        [FindsBy(How = How.Id, Using = SEARCH_STATUS)]
        private IWebElement _serchStatus;

        [FindsBy(How = How.Id, Using = STATUS_FILTER)]
        private IWebElement _StatusFilter;

        [FindsBy(How = How.XPath, Using = SERVICE_SPAN)]
        private IWebElement _serviceSpan;

        [FindsBy(How = How.XPath, Using = SERVICE_CONTAINER)]
        private IWebElement _serviceContainer;

        [FindsBy(How = How.XPath, Using = SECOND_SERVICE_LINE_INPUT)]
        private IWebElement _secondServiceLineInput;

        [FindsBy(How = How.Id, Using = FLIGHT_LEG)]
        private IWebElement _flightLeg;

        [FindsBy(How = How.Id, Using = FLIGHT_TYPE)]
        private IWebElement _flightType;

        [FindsBy(How = How.Id, Using = ARR_FLIGHT_NUMBER)]
        private IWebElement _arrFlightNumber;

        // ___________________________________________ Methodes ____________________________________________________

        // General
        public FlightPage CloseViewDetails()
        {
            WaitPageLoading();
            _closeBtn = WaitForElementIsVisible(By.Id(CLOSE_BTN));
            // légère animation (plus clair) lorsqu'on passe la sourie devant
            new Actions(_webDriver).MoveToElement(_closeBtn).Click().Perform();
            WaitForLoad();

            return new FlightPage(_webDriver, _testContext);
        }

        public string GetFlightNumber()
        {
            _flightNumber = WaitForElementIsVisible(By.XPath(FLIGHT_NUMBER));
            return _flightNumber.Text.Split(' ')[4];
        }

        public string ChangeFlight(string flightToExclude)
        {
            string selectedFlightName = "";

            // On recherche le nombre de vols différents de celui créé
            var flightList = _webDriver.FindElements(By.XPath(String.Format(CHANGE_FLIGHT_NUMBER, flightToExclude)));

            if (flightList.Count > 0)
            {
                selectedFlightName = flightList[0].Text;
                flightList[0].Click();
                WaitForLoad();
            }

            return selectedFlightName;
        }

        // Informations vols
        public void ShowInfoFlightExtendedMenu()
        {
            Actions actions = new Actions(_webDriver);

            _infoFlightExtendedBtn = WaitForElementIsVisible(By.Id(INFO_FLIGHT_EXTENDED_BTN));
            actions.MoveToElement(_infoFlightExtendedBtn).Perform();
            WaitForLoad();
        }

        public FlightPage ImportService(string filePath)
        {
            ShowInfoFlightExtendedMenu();

            _importGuestAndServicesDev = WaitForElementIsVisible(By.Id(IMPORT_GUEST_AND_SERVICES_DEV));
            _importGuestAndServicesDev.Click();
            WaitForLoad();

            _chooseFile = WaitForElementIsVisible(By.Id(CHOOSE_FILE));
            _chooseFile.SendKeys(filePath);

            _validateImport = WaitForElementIsVisible(By.XPath(VALIDATE_IMPORT));
            _validateImport.Click();
            WaitForLoad();
            _okImport = WaitForElementIsVisible(By.XPath("//*[@id=\"modal-1\"]/div[3]/button"));
            _okImport.Click();
            WaitForLoad();

            return new FlightPage(_webDriver, _testContext);
        }

        public FlightPage DuplicateFlight(string flightUpdate)
        {
            if (isElementVisible(By.XPath("//*[@id=\"modal-1\"]/div/div/div[1]/button")))
            {
                var close = WaitForElementIsVisible(By.XPath("//*[@id=\"modal-1\"]/div/div/div[1]/button"));
                close.Click();
                WaitForLoad();
            }

            ShowInfoFlightExtendedMenu();

            _duplicateFlight = WaitForElementIsVisible(By.Id(DUPLICATE_FLIGHT));
            _duplicateFlight.Click();
            WaitPageLoading();

            _duplicateFlightNumber = WaitForElementIsVisible(By.Id(DUPLICATE_FLIGHT_NUMBER));
            _duplicateFlightNumber.SetValue(ControlType.TextBox, flightUpdate);
            WaitForLoad();

            _validateDuplicate = WaitForElementIsVisible(By.Id(VALIDATE_DUPLICATE));
            _validateDuplicate.Click();
            WaitForLoad();

            return new FlightPage(_webDriver, _testContext);
        }

        public FlightPage ReplaceGuest(string initGuest, string newGuest)
        {
            if (isElementVisible(By.XPath("//*[@id=\"modal-1\"]/div/div/div/div/form/div[1]/button")))
            {
                var close = WaitForElementIsVisible(By.XPath("//*[@id=\"modal-1\"]/div/div/div/div/form/div[1]/button"));
                close.Click();
                WaitForLoad();
            }

            ShowInfoFlightExtendedMenu();

            _replaceGuest = WaitForElementIsVisible(By.Id(REPLACE_GUEST));
            _replaceGuest.Click();
            WaitForLoad();

            _initialGuest = WaitForElementIsVisible(By.Id(INITIAL_GUEST));
            _initialGuest.SetValue(ControlType.DropDownList, initGuest);
            WaitForLoad();

            _newGuest = WaitForElementIsVisible(By.Id(NEW_GUEST));
            _newGuest.SetValue(ControlType.DropDownList, newGuest);
            WaitForLoad();

            _confirmReplace = WaitForElementIsVisible(By.Id(CONFIRM_REPLACE));
            _confirmReplace.Click();
            WaitForLoad();

            return new FlightPage(_webDriver, _testContext);
        }

        public FlightPage UpdateGuestTypeQty(string qty)
        {
            if (isElementVisible(By.XPath("//*[@id=\"modal-1\"]/div/div/div[1]/button")))
            {
                var close = WaitForElementIsVisible(By.XPath("//*[@id=\"modal-1\"]/div/div/div[1]/button"));
                close.Click();
                WaitForLoad();
            }

            ShowInfoFlightExtendedMenu();

            _updateGuestQty = WaitForElementIsVisible(By.Id(UPDATE_GUEST_QTY));
            _updateGuestQty.Click();
            WaitForLoad();

            _inputGuestQty = WaitForElementIsVisible(By.Id(INPUT_GUEST_QTY));
            _inputGuestQty.SetValue(ControlType.TextBox, qty);
            WaitForLoad();

            _confirmUpdateQty = WaitForElementIsVisible(By.Id(CONFIRM_UPDATE_QTY));
            _confirmUpdateQty.Click();
            WaitForLoad();

            return new FlightPage(_webDriver, _testContext);
        }

        public bool GetFlightHistory()
        {
            if (isElementVisible(By.XPath("//*[@id=\"modal-1\"]/div/div/div[1]/button")))
            {
                var close = WaitForElementIsVisible(By.XPath("//*[@id=\"modal-1\"]/div/div/div[1]/button"));
                close.Click();
                WaitForLoad();
            }
            ShowInfoFlightExtendedMenu();

            _historyFlight = WaitForElementIsVisible(By.Id(FLIGHT_HISTORY));
            _historyFlight.Click();

            WaitForLoad();
            WaitForElementIsVisible(By.XPath(HISTORY_ROW1_DEV));
            WaitForElementIsVisible(By.XPath(HISTORY_ROW2_DEV));


            return true;
        }

        public FlightPage SplitFlight()
        {
            WaitForLoad();

            ShowInfoFlightExtendedMenu();

            _splitFlightRoute = WaitForElementIsVisible(By.Id(SPLIT_FLIGHT_ROUTE));
            _splitFlightRoute.Click();
            WaitForLoad();

            _validateSplit = WaitForElementIsVisible(By.Id(VALIDATE_SPLIT));
            _validateSplit.Click();
            WaitForLoad();

            return new FlightPage(_webDriver, _testContext);
        }

        public void EditFirstFlightDetails(string flightNumberBis)
        {
            _editFlight = WaitForElementIsVisible(By.Id(EDIT_FLIGHT));
            _editFlight.Click();

            WaitForLoad();

            _modalFlightNumber = WaitForElementIsVisible(By.Id(MODAL_FLIGHT_NUMBER));
            _modalFlightNumber.SetValue(ControlType.TextBox, flightNumberBis);

            _saveFlightOnly = WaitForElementIsVisible(By.XPath(SAVE_FLIGHT_ONLY));
            _saveFlightOnly.Click();
            WaitForLoad();
        }

        public string GetDateConfigured()
        {
            _editFlight = WaitForElementIsVisible(By.Id(EDIT_FLIGHT));
            _editFlight.Click();

            WaitForLoad();

            _modalFlightDate = WaitForElementIsVisible(By.Id(MODAL_FLIGHT_DATE));
            return _modalFlightDate.GetAttribute("value").ToString();
        }

        public void SetNewState(string status)
        {
            switch (status)
            {
                case "P":
                    _pStateBtn = WaitForElementIsVisible(By.Id(P_STATE));
                    _pStateBtn.Click();
                    WaitForLoad();
                    WaitForElementIsVisible(By.XPath(P_STATE_SELECTED));
                    break;
                case "V":
                    _vStateBtn = WaitForElementIsVisible(By.Id(V_STATE));
                    _vStateBtn.Click();
                    WaitForLoad();
                    WaitForElementIsVisible(By.XPath(V_STATE_SELECTED));
                    break;
                case "I":
                    _iStateBtn = WaitForElementIsVisible(By.Id(I_STATE));
                    _iStateBtn.Click();
                    WaitForLoad();
                    WaitForElementIsVisible(By.XPath(I_STATE_SELECTED));
                    break;
                default:
                    break;
            }
        }

        public void UnsetState(string status)
        {
            switch (status)
            {
                case "P":
                    _pStateBtn = WaitForElementIsVisible(By.Id(P_STATE));
                    _pStateBtn.Click();
                    WaitForLoad();
                    WaitForElementIsVisible(By.XPath(P_STATE_UNSELECTED));
                    break;
                case "V":
                    _vStateBtn = WaitForElementIsVisible(By.Id(V_STATE));
                    _vStateBtn.Click();
                    WaitForLoad();
                    WaitForElementIsVisible(By.XPath(V_STATE_UNSELECTED));
                    break;
                case "I":
                    _iStateBtn = WaitForElementIsVisible(By.Id(I_STATE));
                    _iStateBtn.Click();
                    WaitForLoad();
                    WaitForElementIsVisible(By.XPath(I_STATE_UNSELECTED));
                    break;
                default:
                    break;
            }
        }

        public bool GetFlightStatus(string status)
        {
            if (!FlightIsCancelled())
            {
                switch (status)
                {
                    case "P":
                        _pStateBtn = WaitForElementIsVisible(By.Id(P_STATE));
                        if (_pStateBtn.GetAttribute("class") == "active state-preval-PopUp")
                        {
                            return true;
                        }
                        break;
                    case "V":
                        _vStateBtn = WaitForElementIsVisible(By.Id(V_STATE));
                        if (_vStateBtn.GetAttribute("class") == "active state-valid-PopUp")
                        {
                            return true;
                        }
                        break;
                    case "I":
                        _iStateBtn = WaitForElementIsVisible(By.Id(I_STATE));
                        if (_iStateBtn.GetAttribute("class") == "active state-invoice-PopUp")
                        {
                            return true;
                        }
                        break;
                    default:
                        return false;
                }
            }

            return false;
        }
        public bool FlightIsCancelled()
        {
            var cancelledBtn = WaitForElementIsVisible(By.XPath("//*/div[contains(@class,'cancel-flight')]"));
            return cancelledBtn.GetAttribute("class").Contains("active");

        }
        // Guests
        public void AddGuestType(string guestType = null)
        {
            _addGuestBtn = WaitForElementIsVisibleNew(By.XPath(ADD_GUEST_BTN));
            _addGuestBtn.Click();
            WaitForLoad();

            if (guestType != null)
            {
                try
                {
                    _guestName = WaitForElementIsVisibleNew(By.Id(GUEST_NAME));
                    _guestName.SetValue(ControlType.DropDownList, guestType);
                    WaitForLoad();

                }
                catch
                {
                    _guestName = WaitForElementIsVisibleNew(By.Id(GUEST_NAME));
                    WaitForLoad();

                }
            }

            _createGuest = WaitForElementIsVisibleNew(By.XPath(CREATE_GUEST_BUTTON), nameof(CREATE_GUEST_BUTTON));
            _createGuest.Click();
            WaitForLoad();
        }
        public void AddGuestTypeMeal(string route, string guestType = null)
        {
            _addGuestBtn = WaitForElementIsVisible(By.XPath(ADD_GUEST_BTN));
            _addGuestBtn.Click();
            WaitForLoad();

            if (!string.IsNullOrEmpty(guestType))
            {
                _flightLeg = WaitForElementIsVisible(By.Id(FLIGHT_LEG));
                _flightLeg.SetValue(ControlType.DropDownList, route);
                _guestName = WaitForElementIsVisible(By.Id(GUEST_NAME));
                _guestName.SetValue(ControlType.DropDownList, guestType);
                WaitForLoad();
            }
            _createGuest = WaitForElementIsVisible(By.XPath(CREATE_GUEST_BUTTON), nameof(CREATE_GUEST_BUTTON));
            _createGuest.Click();
            WaitForLoad();
            WaitPageLoading();
        }


        public void SelectLpCart(string value)
        {
            _lpCart = WaitForElementIsVisible(By.XPath("//*/select[@id='dropdown-lpcart']/.."));
            _lpCart.Click();
            WaitForLoad();

            _lpCart = WaitForElementIsVisible(By.Id("dropdown-lpcart-selectized"));
            _lpCart.SendKeys(value);
            _lpCart.SendKeys(Keys.Enter);
            WaitPageLoading();
            WaitForLoad();
            // <h3 id="dataAlertLabel">Information (1)</h3>
            if (isElementVisible(By.Id("dataAlertLabel")))
            {
                var button = WaitForElementIsVisible(By.Id("dataAlertCancel"));
                button.Click();
                WaitForLoad();
            }
        }

        public bool LPCartIsExist(string value)
        {
            _lpCart = WaitForElementIsVisible(By.XPath("//*/select[@id='dropdown-lpcart']/.."));
            _lpCart.Click();
            WaitForLoad();
            return isElementExists(By.XPath($"//*[@id=\"dropdown-lpcart\"]/../div[1]/div[2]/div[1]/div[text()=\"{value}\"]"));

        }

        public void ChangeAircaftForLoadingPlan(string aircraft, bool ignoreAlert = true)
        {

            var _aircraft = WaitForElementExists(By.XPath("//*[@id=\"form-createdit-flight-detail\"]/div[1]/div[3]/div/div[17]/div"));
            _aircraft.Click();
            WaitForLoad();
            _aircraftNumber = WaitForElementExists(By.Id(AIRCRAFT_NUMBER));
            WaitForLoad();
            _aircraftNumber.SendKeys(Keys.End);
            _aircraftNumber.SendKeys(Keys.Control + "a");
            _aircraftNumber.SendKeys(Keys.Backspace);
            _aircraftNumber.SendKeys(aircraft.ToString());
            WaitForLoad();
            _aircraftNumber.SendKeys(Keys.Enter);
            WaitPageLoading();
            WaitForLoad();
            if (ignoreAlert && isElementVisible(By.XPath("//*[@id=\"dataAlertModal\"]/div/div")))
            {
                var okbutton = WaitForElementIsVisible(By.XPath("//*[@id=\"dataAlertModal\"]/div/div/div[3]/button[@id=\"dataAlertCancel\"]"));
                okbutton.Click();
                WaitForLoad();
            }
        }

        public void CloseDataAlertCancel()
        {
            if (isElementVisible(By.Id("dataAlertCancel")))
            {
                var okButton = WaitForElementExists(By.Id("dataAlertCancel"));
                okButton.Click();
                WaitForLoad();
            }
        }
        public string ChangeAircaftForLoadingPlanAndReturnPoputTitle(string aircraft)
        {
            _aircraftNumber = WaitForElementExists(By.XPath("//*[@id=\"form-createdit-flight-detail\"]/div[1]/div[3]/div/div[17]/div"));
            _aircraftNumber.Click();
            var airctaftDrp = WaitForElementIsVisible(By.Id(AIRCRAFT_NUMBER));

            airctaftDrp.SendKeys(Keys.End);
            airctaftDrp.SendKeys(Keys.Control + "a");
            airctaftDrp.SendKeys(Keys.Backspace);
            airctaftDrp.SendKeys(aircraft.ToString());
            WaitForLoad();
            airctaftDrp.SendKeys(Keys.Enter);
            WaitForLoad();
            WaitPageLoading();
            WaitForLoad();
            if (isElementVisible(By.Id("dataAlertCancel")))
            {
                var popupTitle = WaitForElementIsVisible(By.Id("dataAlertLabel"));
                var poTitle = popupTitle.Text;
                return poTitle;
            }
            return string.Empty;
        }

        public bool IsLoadingPlanLoaded(string loadingPlanName)
        {
            if (isElementVisible(By.XPath(String.Format("//*[@id=\"InternalFlightRemarks\" and contains(text(), '{0}')]", loadingPlanName))))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public bool IsLoadingPlanLoadedWithTrolley()
        {
            var result = WaitForElementIsVisible(By.XPath("//*[@id=\"dataAlertModal\"]/div/div/div[2]/p"));
            if (result.Text.Contains("LPCart"))
            {
                return true;
            }

            return false;
        }


        public bool IsGuestTypeVisible()
        {
            WaitLoading();
            if (isElementVisible(By.XPath(GUEST_TYPE)))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public string GetGuestType()
        {
            if (isElementVisible(By.XPath(GUEST_TYPE)))
            {
                _guestType = WaitForElementIsVisible(By.XPath(GUEST_TYPE));
                return _guestType.Text.Trim();
            }
            else
            {
                return "";
            }
        }

        public string GetLegName()
        {
            if (isElementVisible(By.XPath(LEG_NAME)))
            {
                _legname = WaitForElementIsVisible(By.XPath(LEG_NAME));
                string[] parts = _legname.Text.Split(new string[] { "\r\n" }, StringSplitOptions.None);

                var result = parts[0].Split(' ')[3] + parts[0].Split(' ')[4] + parts[0].Split(' ')[5];
                return result;
            }
            else
            {
                return "";
            }
        }

        public string GetdeliveryNum()
        {
            if (isElementVisible(By.XPath(DELIVERY_NUM)))
            {
                _delivery_num = WaitForElementIsVisible(By.XPath(DELIVERY_NUM));
                return _delivery_num.Text.Split(' ')[6];
            }
            else
            {
                return "";
            }
        }

        public string GetGuestTypeQty()
        {
            _guestTypeQty = WaitForElementIsVisible(By.Id(GUESTTYPE_QTY));
            return _guestTypeQty.GetAttribute("value");
        }

        public void DeleteGuestType()

        {
            if (GetFlightStatus("V"))
            {
                UnsetState("V");
            }

            _removeGuest = WaitForElementIsVisible(By.XPath(REMOVE_GUEST));
            _removeGuest.Click();
            WaitLoading();

            _confirmDelete = WaitForElementIsVisible(By.Id(CONFIRM_DELETE));
            _confirmDelete.Click();
            WaitPageLoading();
        }

        // Service
        public void ShowAddServiceMenu()

        {

            WaitPageLoading();
            Actions action = new Actions(_webDriver);

            _addServiceMenu = WaitForElementIsVisible(By.XPath("//*[@id=\"leg-details-with-services\"]/div/div/div/button/span"));
            action.MoveToElement(_addServiceMenu).Perform();
            WaitForLoad();


        }

        public void AddService(string serviceName)
        {
            ShowAddServiceMenu();
            _addService = WaitForElementIsVisibleNew(By.XPath(ADD_SERVICE));
            _addService.Click();
            Thread.Sleep(2000);
            WaitForLoad();
            if (serviceName != null)
            {
                IWebElement serviceInputHeader = WaitForElementIsVisibleNew(By.XPath(SERVICE_HEADER));
                IJavaScriptExecutor javaScriptExecutor = _webDriver as IJavaScriptExecutor;
                javaScriptExecutor.ExecuteScript("arguments[0].style.width = '100px';", serviceInputHeader);
                WaitForLoad();

                _serviceLine = WaitForElementIsVisibleNew(By.XPath("(//*/div[contains(@class,'selectize-input')]/div)[last()]"));
                _serviceLine.Click();
                Thread.Sleep(2000);
                WaitForLoad();

                _serviceLineInput = WaitForElementExists(By.XPath(string.Format("(//*[contains(@id,\"ServiceIdContainer_\")]/div/div[2]/div/div[contains(text(),'{0}')])[last()]", serviceName)));
                WaitPageLoading();
                _serviceLineInput.Click();
                Thread.Sleep(2000);
                WaitPageLoading();

                //validateur inefficace  if (isElementVisible(By.XPath(SERVICES_ALREADY_EXIST))) return;
                serviceInputHeader = WaitForElementIsVisibleNew(By.XPath(SERVICE_HEADER));
                javaScriptExecutor = _webDriver as IJavaScriptExecutor;
                javaScriptExecutor.ExecuteScript("arguments[0].style.width = null;", serviceInputHeader);
                WaitForLoad();

                // disquette
                WaitForElementIsVisibleNew(By.XPath(ICON_SAVE));
                // fin de disquette
                Thread.Sleep(2000);
                WaitLoading();
            }

        }


        public int CountServices()
        {
            int countInput = _webDriver.FindElements(By.XPath(SERVICENAME_INPUT)).Count;
            int countReadOnly = _webDriver.FindElements(By.XPath(SERVICENAME_READONLY_ALL)).Count;
            int servicesPatchNumber = _webDriver.FindElements(By.XPath(SEVICESNAMES_ONPATCH)).Count;
            int servicesNumberReadOnlyOnPatch = _webDriver.FindElements(By.XPath(SEVICESNAMES_READONLY_ONPATCH)).Count;
            return IsDev() ? countInput + countReadOnly : servicesPatchNumber + servicesNumberReadOnlyOnPatch;
        }

        public int CountServiceWithoutPrice()
        {
            var countWithoutPrice = _webDriver.FindElements(By.XPath("//*/span[contains(@title,'No Price')]")).Count;

            return countWithoutPrice;
        }


        public int CountServicesFromBob()
        {
            return _webDriver.FindElements(By.XPath(SERVICE_LINE_FROM_BOB)).Count;
        }

        public void DeleteService()
        {
            WaitForLoad();
            var services = _webDriver.FindElements(By.XPath(SERVICES_EXIST));
            Assert.IsTrue(services.Any(), "no services available");
            ZoomOut(0.8m);
            try
            {
                _serviceExtendedMenu = WaitForElementIsVisible(By.Id(SERVICE_EXTENDED_MENU_ID));
                _serviceExtendedMenu.Click();
            }
            catch (ElementClickInterceptedException ex)
            {
                ClickWithJs(_serviceExtendedMenu);
            }
            WaitPageLoading();

            _deleteService = WaitForElementIsVisible(By.Id(DELETE_SERVICE_ID));
            _deleteService.Click();
            WaitPageLoading();
            _confirmDelete = WaitForElementIsVisible(By.Id(CONFIRM_DELETE));
            _confirmDelete.Click();
            WaitPageLoading();
        }
        public void ClickWithJs(IWebElement element)
        {
            ((IJavaScriptExecutor)_webDriver).ExecuteScript("arguments[0].click();", element);
        }
        public string GetServiceQty()
        {
            _serviceQty = WaitForElementIsVisible(By.Id("FinalQuantity"));
            return _serviceQty.GetAttribute("value");
        }



        public bool GetServiceFinalQty1(string value)
        {
            var elements = _webDriver.FindElements(By.ClassName("final-quantity"));
            foreach (var elem in elements)
            {
                if (elem.GetAttribute("value") != value)
                {
                    return false;
                }
            }

            return true;
        }

        public void SetFinalQty(string value)
        {
            var elements = _webDriver.FindElements(By.ClassName("final-quantity"));

            foreach (var elem in elements)
            {
                elem.SetValue(ControlType.TextBox, value);
                elem.SendKeys(Keys.Enter);
                WaitForLoad();
                WaitForElementIsVisible(By.XPath(QUANTITY_UPDATE_DISKETTE_ICON), nameof(QUANTITY_UPDATE_DISKETTE_ICON));

            }
            
            WaitPageLoading();
            WaitForLoad();
        }

        public ServicePricePage SetInvoicedService(string value)
        {
            _servicetab = WaitForElementIsVisible(By.XPath(SERVICE_TAB));
            _servicetab.Click();
            WaitPageLoading();
            WaitForLoad();

            return new ServicePricePage(_webDriver, _testContext);
        }

        public string GetServiceName(bool serviceValide = false)
        {
            // selon screenshoot serveur
            WaitPageLoading();
            WaitForLoad();

            var serviceInputHeader = WaitForElementIsVisible(By.XPath("//*[@id=\"services-details\"]/div/table/tbody/tr[1]/th[1]"));
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            javaScriptExecutor.ExecuteScript("arguments[0].style.width = '100px';", serviceInputHeader);
            string name;
            if (!serviceValide)
            {
                // évite de cliquer partout
                Thread.Sleep(5000);
                WaitForLoad();

                if (isElementExists(By.XPath(SERVICENAME_READONLY_ALL)))
                {
                    _serviceName = WaitForElementExists(By.XPath(SERVICENAME_READONLY_ALL));
                    name = _serviceName.Text.Trim();
                }
                else if (isElementExists(By.XPath(SERVICENAME_INPUT)))
                {
                    _serviceName = WaitForElementExists(By.XPath(SERVICENAME_INPUT));
                    name = _serviceName.Text.Trim();
                }
                else
                {
                    name = "";
                }
            }
            else
            {
                _serviceName = WaitForElementExists(By.XPath(SERVICENAME_VALIDE));
                name = _serviceName.Text;
            }

            // selon screenshoot serveur (rétablissement pour voir le bouton Add Service)
            serviceInputHeader = WaitForElementIsVisible(By.XPath("//*[@id=\"services-details\"]/div/table/tbody/tr[1]/th[1]"));
            javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            javaScriptExecutor.ExecuteScript("arguments[0].style.width = null;", serviceInputHeader);

            return name;
        }

        public bool isMoreThenOneSrvice()
        {
            if (_webDriver.FindElements(By.XPath(SERVICENAME_READONLY_ALL)).Count > 1)
            {
                return true;
            }
            return false;
        }
        public string GetServicesName(string serviceNameSansDatasheet)
        {
            _serviceName = WaitForElementExists(By.XPath(string.Format(SERVICENAME_READONLY, serviceNameSansDatasheet)));
            return _serviceName.Text;
        }
        public List<string> GetServiceNames()
        {
            List<string> serviceNames = new List<string>();

            var liste = _webDriver.FindElements(By.XPath(SERVICENAME_INPUT));
            foreach (var ele in liste)
            {
                serviceNames.Add(ele.Text);
            }
            liste = _webDriver.FindElements(By.XPath(SERVICENAME_READONLY));
            foreach (var ele in liste)
            {
                serviceNames.Add(ele.Text);
            }
            return serviceNames;
        }


        public void AddGenericService()
        {
            ShowAddServiceMenu();
            _addGenericService = WaitForElementIsVisible(By.XPath(ADD_GENERIC_SERVICE));
            _addGenericService.Click();
            WaitForLoad();

            var fermerMenu = WaitForElementIsVisible(By.Id(CLOSE_MENU));
            new Actions(_webDriver).MoveToElement(fermerMenu).Perform();
            WaitForLoad();
        }

        public void CreateGenericService(string serviceName)
        {
            ShowAddServiceMenu();

            _createGenericService = WaitForElementIsVisible(By.XPath(CREATE_GENERIC_SERVICE));
            _createGenericService.Click();
            WaitForLoad();

            _genericServiceName = WaitForElementIsVisible(By.Id(GENERIC_SERVICE_NAME));
            _genericServiceName.SetValue(ControlType.TextBox, serviceName);
            WaitForLoad();

            _saveGenericServiceBtn = WaitForElementIsVisible(By.Id(SAVE_GENERIC_SERVICE));
            _saveGenericServiceBtn.Click();
            WaitPageLoading();
            WaitForLoad();
        }

        public bool IsGenericServiceAdded()
        {
            if (isElementVisible(By.XPath(GENERIC_SERVICE)))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public FlightEditModalPage EditFlightInModal()
        {
            _editFlight = WaitForElementIsVisible(By.Id(EDIT_FLIGHT));
            _editFlight.Click();
            WaitPageLoading();
            WaitForLoad();

            return new FlightEditModalPage(_webDriver, _testContext);
        }

        public void CloseModal()
        {
            var close = WaitForElementIsVisible(By.Id(CLOSE_BTN));
            close.Click();
            WaitLoading();
        }
        public void SetTabletFlightType(string flightType)
        {
            var input = WaitForElementIsVisible(By.Id("dropdown-tablet-flight-type"));
            input.SetValue(ControlType.DropDownList, flightType);
            WaitPageLoading();
            WaitLoading();
        }

        public PrintReportPage PrintFlightLabels()
        {
            var extendMenu = WaitForElementIsVisible(By.Id(EXTEND_FLIGHT_MENU));
            Actions actions = new Actions(_webDriver);
            actions.MoveToElement(extendMenu).Perform();
            var printFlightlabelsBtn = WaitForElementIsVisible(By.XPath(PRINT_FLIGHT_LABELS_BUTTON));
            printFlightlabelsBtn.Click();
            WaitPageLoading();
            var confirmPrintFlightBtn = WaitForElementIsVisible(By.XPath(CONFIRM_PRINT_FLIGHT_LABELS));
            confirmPrintFlightBtn.Click();
            WaitPageLoading();
            WaitForLoad();
            WaitPageLoading();
            // animation
            WaitPageLoading();
            var closeDetailsBtn = WaitForElementIsVisible(By.Id(CLOSE_DETAILS_FLIGHTS_BTN));
            closeDetailsBtn.Click();
            WaitForLoad();

            IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-pdf']"));
            ClickPrintButton();

            //Results are opened in a new tab, switch the driver to the newly created one
            WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new PrintReportPage(_webDriver, _testContext);
        }

        public void SetRegistration(string reg)
        {
            var regInput = WaitForElementIsVisible(By.Id("tb-registration-type-details"));
            regInput.SetValue(ControlType.TextBox, reg);
            WaitForLoad();
            
            try {
                var confirmDefaultLpCart = WaitForElementIsVisible(By.Id("dataAlertCancel"));
                confirmDefaultLpCart.Click();
                WaitForLoad();
            } catch {
                // la boite s'affiche quand on maj aircraft avec une liaison
                // la boite s'affiche quand on maj registration avec une liaison (test FL_FLIG_Index_PrintFlightLabels)
            }
        }

        public List<string> GetGuestTypes()
        {
            List<string> guestTypes = new List<string>();
            if (isElementVisible(By.XPath(GUEST_TYPE)))
            {
                var _guestTypes = _webDriver.FindElements(By.XPath(GUEST_TYPE));
                foreach (var _guestType in _guestTypes)
                {
                    guestTypes.Add(_guestType.Text.Trim());
                }
            }
            return guestTypes;

        }

        public void SelectGuestType(string guest)
        {
            var _guestTypes = WaitForElementIsVisible(By.XPath(string.Format(GUEST_TYPE_SELECT, guest)));
            _guestTypes.Click();
            WaitPageLoading();
        }

        public void GetComments(out string comment1Flight, out string comment2Flight)
        {
            var flesh = WaitForElementIsVisible(By.Id("foldUnfoldRemarksSquare"));
            flesh.Click();
            var comment1textarea = WaitForElementIsVisible(By.Id(COMMENT_TEXTAREA_1));
            comment1Flight = comment1textarea.Text;
            var flesh1 = WaitForElementIsVisible(By.Id("foldUnfoldDNCommentSquare"));
            flesh1.Click();
            var comment2textarea = WaitForElementIsVisible(By.Id(COMMENT_TEXTAREA_2));
            comment2Flight = comment2textarea.Text;
        }

        public void UnfoldInternal_Flg_Remarks(string comment)
        {
            //var listsitesActif = WaitForElementExists(By.Id("foldUnfoldRemarksSquare"));
            //listsitesActif.Click();
            var RemarksInput = WaitForElementExists(By.Id("InternalFlightRemarks"));
            RemarksInput.SetValue(ControlType.TextBox, comment);
        }

        public void SetFirstDriver(string driver)
        {
            throw new NotImplementedException();
        }

        public string GetFirstFlightNumber()
        {
            WaitForLoad();
            var _flightDetails = _webDriver.FindElements(By.XPath("//*[@id=\"flightsListFilterDate\"]//div[contains(@class,\"active\")]/a"));
            var fNumber = _flightDetails[1].Text;
            CloseModal();
            WaitForLoad();
            return fNumber;
        }
        public override void ShowExtendedMenu()
        {
            _extendedButton = WaitForElementExists(By.XPath(EXTENDED_BTN));
            var actions = new Actions(_webDriver);
            actions.MoveToElement(_extendedButton).Perform();
            WaitForLoad();
        }

        public bool CheckFlightDetailStatus(string expectedStatus)
        {
            WaitPageLoading(); // Attendre le chargement initial de la page
            var listFlightDetails = _webDriver.FindElements(By.XPath(LIST_FLIGHT_DETAIL));

            foreach (var element in listFlightDetails)
            {
                WaitForLoad(); // Assurez-vous que l'élément est complètement chargé
                if (element.Text != expectedStatus)
                {
                    return false; // Retourner false dès qu'un élément ne correspond pas
                }
            }
            return true; // Retourner true si tous les éléments correspondent
        }
        public bool CheckPrivalidate(string privalidate)
        {
            return CheckFlightDetailStatus(privalidate);
        }

        public bool Checkvalidate(string validate)
        {
            return CheckFlightDetailStatus(validate);
        }

        public bool CheckInvoice(string invoice)
        {
            return CheckFlightDetailStatus(invoice);
        }
        public FlightDetailsPage ClickOnValidateAll()
        {
            ShowExtendedMenu();
            ClickElement(By.XPath(VALIDATE_ALL));
            ClickElement(By.Id(VALIDE));
            LoadingPage();
            return new FlightDetailsPage(_webDriver, _testContext);
        }

        public FlightDetailsPage ClickOnInvoieAll()
        {
            ShowExtendedMenu();
            ClickElement(By.XPath(INVOICE_ALL));
            ClickElement(By.Id(VALIDE));
            LoadingPage();
            return new FlightDetailsPage(_webDriver, _testContext);
        }

        public FlightDetailsPage ClickOnPrevalidateAll()
        {
            ShowExtendedMenu();
            ClickElement(By.XPath(PREVALIDATE_ALL));
            ClickElement(By.Id(VALIDE));
            LoadingPage();
            return new FlightDetailsPage(_webDriver, _testContext);
        }


        private void ClickElement(By by)
        {
            var element = WaitForElementExists(by);
            WaitForLoad(); // Attendre que l'élément soit complètement chargé
            element.Click();
        }

        public bool CheckCancelled(string Cancelled)
        {
            Thread.Sleep(1500);
            var list_flight_detail = _webDriver.FindElements(By.XPath(LIST_FLIGHT_DETAIL));
            if (list_flight_detail.Count == 0)
            {
                return false;
            }
            WaitPageLoading(); // Attendre le chargement de la page
            WaitPageLoading(); // Attendre à nouveau si nécessaire


            foreach (var element in list_flight_detail)
            {
                if (element.Text != Cancelled)
                {
                    return false; // Retourner false dès qu'un élément ne correspond pas
                }
            }

            return true; // Retourner true si tous les éléments correspondent
        }


        public enum FilterType
        {
            SearchFlight,
            Status
        }



        public void Filter(FilterType filterType, object value)
        {
            ClickOnFilter();
            switch (filterType)
            {
                case FilterType.SearchFlight:
                    _searchInput = WaitForElementIsVisible(By.XPath(SEARCH_FLIGHT_NUMBER));
                    _searchInput.SetValue(ControlType.TextBox, value);
                    WaitPageLoading();
                    break;
                case FilterType.Status:
                    ComboBoxSelectById(new ComboBoxOptions(STATUS_FILTER, (string)value, false));
                    break;


                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);

            }
            Submit();
        }
        public void ClickOnFilter()
        {
            var _clickOnFilter = WaitForElementIsVisible(By.Id("filter-model"));
            _clickOnFilter.Click();
            WaitPageLoading();
        }
        public void Submit()
        {
            var _submitFilter = WaitForElementIsVisible(By.Id("flights"));
            _submitFilter.Click();
            WaitPageLoading();
        }
        public void AddGenericServiceWithName(string serviceName)
        {
            ShowAddServiceMenu();
            _addGenericService = WaitForElementIsVisible(By.XPath("//*[@id=\"leg-details-with-services\"]/div/div/div/div/a[2]"));
            _addGenericService.Click();
            WaitForLoad();
            _serviceLineInput.SetValue(ControlType.TextBox, serviceName);
            _serviceContainer.Click();
            WaitForLoad();

        }
        public void AddSecondGenericServiceWithName(string serviceName)
        {
            ShowAddServiceMenu();
            _addGenericService = WaitForElementIsVisible(By.XPath("//*[@id=\"leg-details-with-services\"]/div/div/div/div/a[2]"));
            _addGenericService.Click();
            WaitForLoad();
            _serviceLineInput = WaitForElementIsVisible(By.XPath("/html/body/div[5]/div/div/div[1]/table/tbody/tr/td[2]/div/div/table/tbody/tr/td[2]/div/div/div[2]/div/table/tbody/tr[4]/td[1]/div/div/div[1]/input"));
            _serviceLineInput.SetValue(ControlType.TextBox, serviceName);
            _serviceLineInput.SendKeys(Keys.Enter);
            WaitForLoad();

        }
        public bool CheckServiceExists()
        {
            WaitForLoad();
            _serviceSpan = WaitForElementIsVisible(By.XPath(SERVICES_ALREADY_EXIST));
            string message = _serviceSpan.Text;
            if (message == "service already exists")
            {
                return true;
            }
            return false;
        }

        public void SetFinalQtyServiceFlight(string value)
        {
            WaitForLoad();
            var elements = _webDriver.FindElements(By.ClassName("final-quantity"));

            foreach (var elem in elements)
            {
                elem.SetValue(ControlType.TextBox, value);
                elem.SendKeys(Keys.Enter);
                WaitForLoad();

            }
            WaitForLoad();
            WaitPageLoading();
        }

        // Guests
        public void AddGuestTypeFlight(string guestType = null)
        {
            string typeGuest = "YC";
            string selectedGuestType = guestType ?? typeGuest;
            if (GetGuestTypes().Contains(typeGuest))
            {
                var guestYC = WaitForElementIsVisible(By.XPath(string.Format(GUEST_TYPE_SELECTED, typeGuest)));
                guestYC.Click();
                WaitForLoad();
                DeleteGuestType();
            }
            _addGuestBtn = WaitForElementIsVisible(By.XPath(ADD_GUEST_BTN));
            _addGuestBtn.Click();
            WaitForLoad();

            if (string.IsNullOrEmpty(guestType))
            {
                _guestName = WaitForElementIsVisible(By.Id(GUEST_NAME));
                _guestName.SetValue(ControlType.DropDownList, typeGuest);
                WaitForLoad();
            }

            _createGuest = WaitForElementIsVisible(By.XPath(CREATE_GUEST_BUTTON), nameof(CREATE_GUEST_BUTTON));
            _createGuest.Click();
            WaitForLoad();

            var guestSelected = WaitForElementIsVisible(By.XPath(string.Format(GUEST_TYPE_SELECTED, selectedGuestType)));
            guestSelected.Click();
            WaitForLoad();
        }

        public void ApplyLPCartToSameFutureFlightsWithoutConfirm(bool isChecked)
        {
            var lpCartToFutureFlights = WaitForElementIsVisible(By.Id(IS_LPCART_APPLY_TO_FUTURE_Flights));
            lpCartToFutureFlights.SetValue(ControlType.CheckBox, isChecked);
            WaitForLoad();
            WaitPageLoading();
        }
        public bool ModalLPCartToSameFutureFlightsInfo()
        {
            Thread.Sleep(1000);
            return isElementVisible(By.XPath(APPLY_TO_FUTURE_FLIGHTS_MODAL_INFO)) && isElementVisible(By.XPath(MESSAGE_APPLY_LPCART_MODAL));
        }
        public void ConfirmApplyLPCartToSameFutureFlights()
        {
            var confirm = WaitForElementIsVisible(By.XPath(CONFIRM_APPLY_TO_FUTURE_FLIGHTS));
            confirm.Click();
            WaitPageLoading();
        }
        public bool IsFlightDetailsPopupOpen()
        {
            WaitPageLoading();
            return _webDriver.FindElement(By.Id(FLIGHT_DETAIL)).Displayed;
        }
        public bool ServiceExist()
        {
            WaitLoading();
            var services = _webDriver.FindElements(By.XPath(SERVICES_EXIST));
            return services.Count > 0;
        }
        public void validateFlight()
        {
            WaitPageLoading();
            var preval = WaitForElementIsVisible(By.Id("prevalPopUp"));
            preval.Click();
            WaitForLoad();
            WaitPageLoading();
            var val = WaitForElementIsVisible(By.Id("validatePopup"));
            val.Click();
            WaitForLoad();
            WaitPageLoading();


        }

        public void ClickFirstFlightInList(string flightNumber)
        {
            WaitPageLoading();
            var count = 1;
            var _flightNumber = WaitForElementIsVisible(By.XPath(string.Format(FIRST_FLIGHT_NUMBER, count)));

            if (_flightNumber.Text.Equals(flightNumber))
            {
                count++;
                _flightNumber = WaitForElementIsVisible(By.XPath(string.Format(FIRST_FLIGHT_NUMBER, count)));
            }

            _flightNumber.Click();
            WaitPageLoading();

        }

        public List<string> GetListService()
        {
            List<string> serviceNames = new List<string>();

            var liste = _webDriver.FindElements(By.XPath(SERVICE_NAMES_LIST));
            foreach (var ele in liste)
            {
                serviceNames.Add(ele.Text.Trim());
            }
            return serviceNames;
        }
        public PrintReportPage SwapLPCartReport()
        {
            var extendMenu = WaitForElementIsVisible(By.Id(EXTEND_FLIGHT_MENU));
            Actions actions = new Actions(_webDriver);
            actions.MoveToElement(extendMenu).Perform();
            var swapLPCartReportBtn = WaitForElementIsVisible(By.XPath(SWAP));
            swapLPCartReportBtn.Click();
            WaitPageLoading();
            WaitForLoad();
            WaitPageLoading();
            // animation
            WaitPageLoading();
            var closeDetailsBtn = WaitForElementIsVisible(By.Id(CLOSE_DETAILS_FLIGHTS_BTN));
            closeDetailsBtn.Click();
            WaitForLoad();

            IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-pdf']"));
            ClickPrintButton();

            //Results are opened in a new tab, switch the driver to the newly created one
            WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new PrintReportPage(_webDriver, _testContext);
        }

        public void EditFlightType(string flightType)
        {
            var type = WaitForElementIsVisible(By.Id(FLIGHT_TYPE));
            type.Click();
            WaitForLoad();
            _flightType.SetValue(ControlType.TextBox, flightType);
            type.Click();
        }

        public string GetFlightType()
        {
            return WaitForElementIsVisible(By.Id(FLIGHT_TYPE)).GetAttribute("value");
        }

        public void EditGuestTypeQT(string quantity)
        {
            var type = WaitForElementIsVisible(By.Id(GUESTTYPE_QTY));
            type.SetValue(ControlType.TextBox, quantity);
            WaitForLoad();
        }

        public string GetFlightAircraft()
        {
            var _aircraft = WaitForElementExists(By.XPath("//*[@id=\"form-createdit-flight-detail\"]/div[1]/div[3]/div/div[17]/div"));
            _aircraft.Click();
            WaitForLoad();
            _aircraftNumber = WaitForElementExists(By.Id(AIRCRAFT_NUMBER));
            WaitForLoad();
            return WaitForElementExists(By.Id(AIRCRAFT_NUMBER)).GetAttribute("value");
        }

        public string GetArrFlightNumber()
        {
            return WaitForElementIsVisible(By.Id(ARR_FLIGHT_NUMBER)).GetAttribute("value");
        }

        public void AddArrFlightNumber(string number)
        {
            var flightnumber = WaitForElementIsVisible(By.Id(ARR_FLIGHT_NUMBER));
            flightnumber.SetValue(ControlType.TextBox, number);
            WaitForLoad();
        }

        public string GetRegistration()
        {
            return WaitForElementIsVisible(By.Id(REGISTRATION)).GetAttribute("value");
        }
        public string GetTabletFlightType()
        {
            SelectElement selectElement = new SelectElement(_webDriver.FindElement(By.Id(TABLET_FLIGHT_TYPE)));
            return selectElement.SelectedOption.Text;
        }

        public string GetLpCart()
        {
            _lpCart = WaitForElementIsVisible(By.XPath("//*/select[@id='dropdown-lpcart']/.."));
            _lpCart.Click();
            WaitForLoad();
            return WaitForElementExists(By.Id(LPCART)).GetAttribute("value");
        }
        public void OkBtn()
        {

            var ok = WaitForElementIsVisible(By.Id("dataAlertCancel"));
            ok.Click();
        }

        public void FillFilterParamaters(string flightNo, string site, string status = null)
        {
            ClickOnFilter();

            IWebElement siteElment = WaitForElementIsVisible(By.XPath(INPUTSITEFILTER));
            siteElment.SetValue(ControlType.DropDownList, site);
            WaitLoading();

            IWebElement fligntNoElment = WaitForElementIsVisible(By.XPath(INPUTFLIGHTNUMBEFILTER));
            fligntNoElment.Clear();
            fligntNoElment.SetValue(ControlType.TextBox, flightNo);
            WaitLoading();

            if (status != null)
            {
                ComboBoxSelectById(new ComboBoxOptions(STATUS_FILTER, (string)status, false));
            }

            Submit();
        }

        public bool VerifyDatasheetExist()
        {
            var datasheet = WaitForElementIsVisible(By.Id(DATA_SHEET));
            if (!datasheet.Displayed)
            {
                return false;
            }
            string actualColor = datasheet.GetCssValue("color");
            return actualColor == "rgba(105, 163, 5, 1)";
        }

        public bool VerifyPicoExist()
        {
            var datasheet = WaitForElementIsVisible(By.Id(DATA_SHEET));
            if (!datasheet.Displayed)
            {
                return false;
            }
            string actualColor = datasheet.GetCssValue("color");
            return actualColor == "rgba(255, 165, 0, 1)";
        }

        public bool PicoExist()
        {
            WaitForLoad();
            // affichage lent de la page
            Thread.Sleep(5000);
            var pico = _webDriver.FindElement(By.XPath(PICO));
            return pico.Displayed;

        }

        public int GetNumberOfFlight()
        {
            int result;
            var list_flight_detail = _webDriver.FindElements(By.XPath(LIST_FLIGHT_DETAIL));
            result = list_flight_detail.Count;
            return result;
        }

        public void ClickOnPrevalidate()
        {
            var status = WaitForElementIsVisible(By.Id(P_STATE));
            status.Click();
        }
        public void ClickCancelledFlight()
        {
            var cancelled = WaitForElementIsVisible(By.Id("cancelFlight"));
            cancelled.Click();
        }
    }
}
