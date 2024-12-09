using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Wordprocessing;
using iText.StyledXmlParser.Jsoup.Nodes;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.Extensions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Security.Policy;
using System.Threading;
using System.Web.Caching;

namespace Newrest.Winrest.FunctionalTests.PageObjects.TabletApp
{
    public class FlightTabletAppPage : PageBase
    {
        string[] monthNames = CultureInfo.CurrentCulture.DateTimeFormat.MonthNames;
        string[] monthNamesFR = CultureInfo.GetCultureInfo("fr-FR").DateTimeFormat.MonthNames;
        //______________________________________CONSTANTES_____________________________________________________
        private const string COMMENT_BUTTON_1 = "//*/td[8]/div/mat-icon";
        private const string COMMENT_BUTTON_2 = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight/div/div[2]/table/div/virtual-scroller/div[2]/tbody/tr[1]/td[12]/div/mat-icon";
        private const string COMMENT_BUTTONS = "//*/mat-icon[text()='speaker_notes']";
        private const string COMMENT_INPUT_1 = "//*/div[text()='InternalFlightRemarks: ']//input";
        private const string COMMENT_INPUT_2 = "//*/div[text()='DeliveryNoteComment: ']//input";
        private const string COMMENT_SAVE_BUTTON = "//*/span[text()='Save']/parent::button";
        private const string FILTER_BUTTON = "filterButton";
        private const string OK_FILTER_BUTTON = "//button[@class= 'okFilterButton shadow']";
        private const string FIRST_FLIGHT_NUMBER = "//*/td[contains(@class,'FlightName')]";
        private const string LINES_FLIGHTS = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight/div/div[2]/table/div/virtual-scroller/div[2]/tbody/tr[*]/td[4]";
        private const string FLIGHT_COLUMN = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight/div/div[2]/table/div/virtual-scroller/div[2]/tbody/tr/td[4]";

        private const string SEARCH_INPUT_FILTER = "//div/label[text()='Search']/../input";
        private const string DATE_ICONE = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight/div/div[1]/div[2]/div[2]";
        private const string YEARS = "//*[@id=\"mat-datepicker-0\"]/div/mat-multi-year-view/table/tbody/tr[*]/td[*]/button/span[1]";
        private const string YEARS_BTN = "/html/body/div[2]/div[2]/div/mat-datepicker-content/div[2]/mat-calendar/mat-calendar-header/div/div/button[1]";
        private const string DAYS = "//*[@id=\"mat-datepicker-0\"]/div/mat-month-view/table/tbody/tr[*]/td[*]/button/span[1]";
        private const string MONTHS = "//*[@id=\"mat-datepicker-0\"]/div/mat-year-view/table/tbody/tr[*]/td[*]/button/span[1]";

        private const string FILTER_ETD_FROM = "//div/label[text()='ETD From:']/../input";
        private const string FILTER_ETD_TO = "//div/label[text()='ETD To:']/../input";
        private const string ETD_COLUMNS = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight/div/div[2]/table/div/virtual-scroller/div[2]/tbody/tr[*]/td[7]";
        private const string SHOW_DOCK_CHECKED = "//flight-filters/div/div[*]/mat-slide-toggle/div/label[text()=' Show dock ']";
        private const string SHOW_DRIVER_CHECKED = "//flight-filters/div/div[*]/mat-slide-toggle/div/label[text()=' Show driver ']";
        private const string SHOW_ETA_CHECKED = "//*/label[text()=' Show ETA ']";
        private const string SHOW_GUEST_CHECKED = "//*/label[text()=' Show Guest ']";
        private const string SHOW_NOTIFICATION = "//*[@id=\"mat-mdc-slide-toggle-30-label\"]";
        private const string SHOW_REGISTRATION_TYPE = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight/div/div[1]/flight-filters/div/div[6]/mat-slide-toggle/div/button";
        private const string SHOW_ROBOT_TSU = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight/div/div[1]/flight-filters/div/div[11]/mat-slide-toggle/div/button";

        private const string COUNT_COLUMNS = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight/div/div[2]/table/div/thead/tr/th";
        private const string SHOW_LEG_CHECKED = "//*/label[text()=' Show Leg ']";
        private const string SHOW_SPEC_MEAL_CHECKED = "//*/label[text()=' Show spec. meal ']";
        private const string SHOW_COLUMN = "//tr/th[text()='{0}']";
        private const string SHOW_WORK_SHOP_CHECKED = "//flight-filters/div/div[*]/mat-slide-toggle/div/label[text()=' Show Workshop ']";
        private const string SHOW_WITH_ALERT_ONLY_FILTER = "//*/label[text()=' Show flights - with alerts only ']/../button";
        private const string SHOW_VALIDATED_ONLY_FILTER = "//*/label[text()='Show flights - validated only']/../button";
        private const string SHOW_MAJOR_ONLY_FILTER = "//*/label[text()='Show flights - major only']/../button";
        private const string SHOW = "//*/label[text()=\"{0}\"]/../button";
        private const string FILTER_COMBOBOX = "//*/label[text()='{0}']/../ng-multiselect-dropdown";
        private const string FILTER_COMBOBOX_SELECT_ALL = "//*/label[text()='{0}']/../ng-multiselect-dropdown//input[@aria-label='multiselect-select-all']";
        private const string FILTER_COMBOBOX_SELECT_ALL2 = "//*/label[text()='{0}']/..//div[text()='{1}']/../input";
        private const string FILTER_COMBOBOX_SELECT = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight/div/div[1]/flight-filters/div/div[11]/div[*]/label[text()='{0}']/../ng-multiselect-dropdown/div/div[2]/ul[2]/li[*]/input[@aria-label='{1}']/../div";
        private const string SHOW_WITH_NOTIFICATION_ONLY = "//*/label[text()='Show with notification only']/../button";
        private const string SAVE_VIEW_BUTTON = "//button[contains(@class, 'saveViewButton')]";
        private const string VIEWS = "/html/body/div[2]/div[2]/div/div/div[*]/mat-option/span";
        private const string VIEW_SELECTED = "//*[@id=\"mat-select-value-3\"]/span/mat-select-trigger";
        private const string DELETE_VIEW_ICONE = "/html/body/div[2]/div[2]/div/div/div/mat-icon";
        private const string DELETE_VIEW_ICONE_ALL = "/html/body/div[2]/div[2]/div/div/div[*]/mat-icon";
        private const string EXPAND_ALL_FLIGHTS_FILTER = "//*/label[text()=' Expand all flights ']/../button\"";
        private const string FIRST_FLIGHT_EDIT_BTN = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight/div/div[2]/table/div/virtual-scroller/div[2]/tbody/tr[1]/td[12]/div";
        private const string FIRST_FLIGHT_QTY = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-detail/div[5]/div/flight-detail-leg/div/div[2]/flight-detail-service-list/div/table/tr[2]/td[3]/div[1]";
        private const string FIRST_FLIGHT_QTY_INPUT = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-detail/div[5]/div/flight-detail-leg/div/div[2]/flight-detail-service-list/div/table/tr[2]/td[3]/input";
        private const string FIRST_GUEST_INPUT = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-detail/div[5]/div/flight-detail-leg/div/div[1]/div/div[2]/table/tr[1]/td[2]/input";
        private const string FLIGHT_STATUS = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight/div/div[2]/table/div/virtual-scroller/div[2]/tbody/tr[1]/td[13]/cp-component-4button/div/div";
        private const string CHOOSE_VIEW_LIST = "//mat-option[@role='option']";
        private const string FINAL_STATE_FIRST_LINE = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight/div/div[2]/table/div/virtual-scroller/div[2]/tbody/tr[1]/td[13]/cp-component-4button/div";
        private const string FINAL_STATE_SECOND_LINE = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight/div/div[2]/table/div/virtual-scroller/div[2]/tbody/tr[2]/td[13]/cp-component-4button/div";
        private const string FIRSTLINEDETAILS = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight/div/div[2]/table/div/virtual-scroller/div[2]/tbody/tr[1]/td[23]";
        private const string HISTORY_BUTTON_FIRST_LINE = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight/div/div[2]/table/div/virtual-scroller/div[2]/tbody/tr[1]/td[10]/cp-component-3button/div/button";
        private const string HISTORY_BUTTON_SECOND_LINE = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight/div/div[2]/table/div/virtual-scroller/div[2]/tbody/tr[2]/td[10]/cp-component-3button/div/button";

        private const string CLICKNOTIFICATION = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight/div/div[2]/table/div/virtual-scroller/div[2]/tbody/tr[1]/td[12]";
        private const string ADDNOTIFICATIONBUTTON = "/html/body/div[2]/div[2]/div/mat-dialog-container/div/div/flight/div/div[3]/div/button[2]";
        private const string NOTIFICATIONTEXT = "/html/body/div[2]/div[4]/div/mat-dialog-container/div/div/notifcreation/div/div[2]/div[2]/ckeditor/div[2]/div[2]/div";
        private const string ADDNOTIFICATIONAFTERTEXT = "/html/body/div[2]/div[4]/div/mat-dialog-container/div/div/notifcreation/div/div[3]/div/button[2]";
        private const string CONFIRMNOTIFICATIONFORADD = "/html/body/div[2]/div[2]/div/mat-dialog-container/div/div/flight/div/div[2]/div/div[1]/table/tr/th[1]/button";
        private const string SAVEBUTTONFORNOTIFICATION = "/html/body/div[2]/div[2]/div/mat-dialog-container/div/div/flight/div/div[3]/div/button[3]";
        private const string CLICKONCHAMPS = "/html/body/div[2]/div[4]/div/mat-dialog-container/div/div/notifcreation/div/div[2]/div[2]/ckeditor/div[2]/div[2]/div";
        private const string VERIFNOTIFADDED = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight/div/div[2]/table/div/virtual-scroller/div[2]/tbody/tr[1]/td[12]/div/mat-icon";
        private const string FINAL_STATE_BUTTON = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight/div/div[2]/table/div/virtual-scroller/div[2]/tbody/tr[{0}]/td[13]/cp-component-4button/div";
        private const string FLIGHT_NAME = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight/div/div[2]/table/div/virtual-scroller/div[2]/tbody/tr[{0}]/td[4]";
        private const string DELETENOTIF = "/html/body/div[2]/div[2]/div/mat-dialog-container/div/div/flight/div/div[2]/div/div[1]/table/tr/th[5]";
        private const string LIST_FLIGHT_NB_SERVICES = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight/div/div[2]/table/div/virtual-scroller/div[2]/tbody/tr[*]/td[9]";
        private const string FLIGHT_HAVE_SERVICE = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight/div/div[2]/table/div/virtual-scroller/div[2]/tbody/tr[{0}]/td[12]/div";
        private const string NUMBER_OF_FLIGHTS = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight/div/div[1]/div[2]/div[1]/div/span[1]";
        private const string CUSTOMER_NAME = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight/div/div[2]/table/div/virtual-scroller/div[2]/tbody/tr[1]/td[2]";

        private const string FLIGHTS = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight/div/div[2]/table/div/virtual-scroller/div[2]/tbody/tr[*]/td[4]";
        private const string VIEW = "//div[@role='listbox']";
        private const string DEFAULT_VIEW = "/html/body/div[2]/div[2]/div//mat-option/span[text()=\"Default view\"]";
        private const string CLICK_VIEW = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight/div/div[1]/div[1]/mat-form-field/div[1]/div[2]/div";
        private const string FLIGHT_DETAILS = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight/div/div[2]/table/div/virtual-scroller/div[2]/tbody/tr[{0}]/td[contains(@class,\"ShowDetails\")]/div[contains(@class,\"btn-icon-edit-flight\")]";
        private const string LIST_OF_NBR_SERVICES = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight/div/div[2]/table/div/virtual-scroller/div[2]/tbody/tr[*]/td[contains(@class,'nbServProd')]";
        private const string VIEW_INPUT = "//*/mat-dialog-container[contains(@id,'mat-mdc-dialog-')]//input";
        private const string VIEW_SAVE = "//*/mat-dialog-container[contains(@id,'mat-mdc-dialog-')]//button[2]";
        private const string CHOOSE_VIEW_COMBO = "//*/mat-select[contains(@id,'mat-select')]";
        private const string FLIGHTS_STATE = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight/div/div[2]/table/div/virtual-scroller/div[2]/tbody/tr[*]/td[10]/ul/li[2]";
        private const string DATASHEET_NAME = "//*[@id=\"flight-detail-datasheet-dialog\"]/div/div/app-flight-detail-datasheet/datasheet-shared-view/div/div/div/div[1]/div/div[2]/div[1]/span";
        private const string FIRST_SERVICE_NAME = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-detail/div[5]/div/flight-detail-leg/div/div[2]/flight-detail-service-list/div/table/tr[2]/td[1]";
        private const string FIRST_ITEM_NAME = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-detail/div[5]/div/flight-detail-leg/div/div[2]/flight-detail-service-list/div/table/tr[2]/td[1]";
        // actions
        private const string CHANGE_STATUS = "//*/div[text()=' {0} ']";

        private const string ENSAMBL = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-workshop/div/div[2]/table/thead/tr/th[9]";
        private const string CORTE = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-workshop/div/div[2]/table/thead/tr/th[10]";
        private const string HISTORY = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-workshop/div/div[2]/table/thead/tr/th[11]";
        private const string CLICK_FLIGHT = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/div/div[3]";
        private const string TIME_BLOCK = "/html/body/app-root/mat-sidenav-container/mat-sidenav/div/div/ul/li[4]/a";
        private const string CLICK_RECIPE = "//*[@id=\"flight-detail-datasheet-dialog\"]/div/div/app-flight-detail-datasheet/datasheet-shared-view/div/div/div/div[3]/mat-list/details/summary/div/div[1]/span";
        private const string CLICK_DATASHEET = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-detail/div[5]/div/flight-detail-leg/div/div[2]/flight-detail-service-list/div/table/tr[2]/td[6]/em";
        
        
        public FlightTabletAppPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // ____________________________________ Variables _______________________________________________
        [FindsBy(How = How.Id, Using = FLIGHT_STATUS)]
        private IWebElement _flightStatus;
        [FindsBy(How = How.XPath, Using = CLICKNOTIFICATION)]
        private IWebElement _clicknotification;
        [FindsBy(How = How.XPath, Using = DELETENOTIF)]
        private IWebElement _deletenotif;
        [FindsBy(How = How.XPath, Using = CLICKONCHAMPS)]
        private IWebElement _clickonchamps;
        [FindsBy(How = How.XPath, Using = SAVEBUTTONFORNOTIFICATION)]
        private IWebElement _savenotification;
        [FindsBy(How = How.XPath, Using = CONFIRMNOTIFICATIONFORADD)]
        private IWebElement _confirmnotificationforadd;
        [FindsBy(How = How.XPath, Using = NOTIFICATIONTEXT)]
        private IWebElement _notificationtext;
        [FindsBy(How = How.XPath, Using = ADDNOTIFICATIONAFTERTEXT)]
        private IWebElement _notificationaftertext;
        [FindsBy(How = How.XPath, Using = ADDNOTIFICATIONBUTTON)]
        private IWebElement _addnotificationbutton;
        [FindsBy(How = How.Id, Using = FIRSTLINEDETAILS)]
        private IWebElement _commentbutton2;
        [FindsBy(How = How.Id, Using = FILTER_BUTTON)]
        private IWebElement _filterBtn;

        [FindsBy(How = How.XPath, Using = SHOW_DOCK_CHECKED)]
        private IWebElement _showDockChecked;

        [FindsBy(How = How.XPath, Using = OK_FILTER_BUTTON)]
        private IWebElement _okFilterButton;

        [FindsBy(How = How.XPath, Using = SHOW_DRIVER_CHECKED)]
        private IWebElement _showDriverChecked;

        [FindsBy(How = How.XPath, Using = SHOW_ETA_CHECKED)]
        private IWebElement _showETAChecked;
        [FindsBy(How = How.XPath, Using = SHOW_NOTIFICATION)]
        private IWebElement _shownotification;


        [FindsBy(How = How.XPath, Using = SHOW_REGISTRATION_TYPE)]
        private IWebElement _showRegistration;

        [FindsBy(How = How.XPath, Using = SHOW_ROBOT_TSU)]
        private IWebElement _showRobotTsu;



        [FindsBy(How = How.XPath, Using = SHOW_GUEST_CHECKED)]
        private IWebElement _showGuest;

        [FindsBy(How = How.XPath, Using = SHOW_LEG_CHECKED)]
        private IWebElement _showLeg;

        [FindsBy(How = How.XPath, Using = SHOW_SPEC_MEAL_CHECKED)]
        private IWebElement _specMeal;

        [FindsBy(How = How.XPath, Using = SHOW_WORK_SHOP_CHECKED)]
        private IWebElement _showWorkShop;

        [FindsBy(How = How.XPath, Using = SEARCH_INPUT_FILTER)]
        private IWebElement _searchfilter;

        [FindsBy(How = How.XPath, Using = FLIGHT_COLUMN)]
        private IWebElement _flightColumn;

        [FindsBy(How = How.XPath, Using = DATE_ICONE)]
        private IWebElement _dateIcone;

        [FindsBy(How = How.XPath, Using = FILTER_ETD_FROM)]
        private IWebElement _etdFrom;

        [FindsBy(How = How.XPath, Using = FILTER_ETD_TO)]
        private IWebElement _etdTo;

        [FindsBy(How = How.XPath, Using = SHOW_WITH_ALERT_ONLY_FILTER)]
        private IWebElement _showWithAlertOnly;

        [FindsBy(How = How.XPath, Using = SHOW_VALIDATED_ONLY_FILTER)]
        private IWebElement _showValidatedOnly;

        [FindsBy(How = How.XPath, Using = SHOW_WITH_NOTIFICATION_ONLY)]
        private IWebElement _showWithNotificationOnly;

        [FindsBy(How = How.XPath, Using = SAVE_VIEW_BUTTON)]
        private IWebElement _saveViewButton;

        [FindsBy(How = How.XPath, Using = SHOW_MAJOR_ONLY_FILTER)]
        private IWebElement _showMajorOnly;
        [FindsBy(How = How.XPath, Using = EXPAND_ALL_FLIGHTS_FILTER)]
        private IWebElement _expandallFlightsFilter;

        [FindsBy(How = How.XPath, Using = FIRST_FLIGHT_EDIT_BTN)]
        private IWebElement _firstFlightEditBtn;

        [FindsBy(How = How.XPath, Using = FIRST_FLIGHT_QTY)]
        private IWebElement _firstFlightQty;

        [FindsBy(How = How.XPath, Using = FIRST_FLIGHT_QTY_INPUT)]
        private IWebElement _firstFlightQtyInput;

        [FindsBy(How = How.XPath, Using = FIRST_GUEST_INPUT)]
        private IWebElement _firstGuestInput;

        [FindsBy(How = How.XPath, Using = FLIGHTS)]
        private IWebElement _flights;
        [FindsBy(How = How.XPath, Using = VIEW)]
        private IWebElement _view;

        [FindsBy(How = How.XPath, Using = DEFAULT_VIEW)]
        private IWebElement _defaultView;

        [FindsBy(How = How.XPath, Using = CLICK_VIEW)]
        private IWebElement _clickView;

        [FindsBy(How = How.XPath, Using = ENSAMBL)]
        private IWebElement _ensambl;

        [FindsBy(How = How.XPath, Using = CORTE)]
        private IWebElement _corte;

        [FindsBy(How = How.XPath, Using = HISTORY)]
        private IWebElement _history;

        [FindsBy(How = How.XPath, Using = CLICK_FLIGHT)]
        private IWebElement _clickFlight;

        [FindsBy(How = How.XPath, Using = TIME_BLOCK)]
        private IWebElement _timeBlock;
        [FindsBy(How = How.XPath, Using = CLICK_RECIPE)]
        private IWebElement _clickRecipe;
        [FindsBy(How = How.XPath, Using = CLICK_DATASHEET)]
        private IWebElement _clickDatasheet;

        public enum FilterType
        {
            AutoScroll,
            ShowETA,
            ShowGuest,
            ShowWorkshop,
            ShowLeg,
            ShowDriver,
            ShowDock,
            ShowNotification,
            ShowSpecMeal,
            ShowFlights_ByflightDate,
            ShowFlights_ByTabletFlightType,
            ShowFlights_MajorOnly,
            ShowWithNotificationOnly,
            ShowFlights_WithAlertsOnly,
            ShowFlights_ValidatedOnly,
            ShowFlights_NotWithDoneStatus,
            ShowFlights_BarsetsToo,
            ExpandAllFlights,
            Customer,
            Workshop,
            Search,
            EDTFrom,
            EDTTo,
            ShowWithAlertOnly,
            ShowValidatedOnly,
            Driver,
            TabletFlightType,
            ShowRegistration,
            ShowRobotTSU
        }

        public void Filter(FilterType filterType, object value, object value2 = null)
        {
            switch (filterType)
            {
                case FilterType.ShowDock:
                    _showDockChecked = WaitForElementIsVisible(By.XPath(string.Format(SHOW, " Show dock ")));
                    if (_showDockChecked.GetAttribute("aria-checked") != value.ToString().ToLower())
                    {
                        _showDockChecked.Click();
                    }
                    break;

                 case FilterType.ShowRegistration:
                    _showRegistration = WaitForElementIsVisible(By.XPath(string.Format(SHOW_REGISTRATION_TYPE, " Show Registration ")));
                    if (_showRegistration.GetAttribute("aria-checked") != value.ToString().ToLower())
                    {
                        _showRegistration.Click();
                    }
                    break;


                case FilterType.ShowRobotTSU:
                    _showRobotTsu = WaitForElementIsVisible(By.XPath(string.Format(SHOW_ROBOT_TSU, " Show RobotTSU ")));
                    if (_showRobotTsu.GetAttribute("aria-checked") != value.ToString().ToLower())
                    {
                        _showRobotTsu.Click();
                    }
                    break;


                case FilterType.ShowDriver:
                    _showDriverChecked = WaitForElementIsVisible(By.XPath(string.Format(SHOW, " Show driver ")));
                    if (_showDriverChecked.GetAttribute("aria-checked") != value.ToString().ToLower())
                    {
                        _showDriverChecked.Click();
                    }
                    break;

                case FilterType.ShowETA:
                    _showETAChecked = WaitForElementIsVisible(By.XPath(string.Format(SHOW, " Show ETA ")));
                    if (_showETAChecked.GetAttribute("aria-checked") != value.ToString().ToLower())
                    {
                        _showETAChecked.Click();
                    }
                    break;
                case FilterType.ShowNotification:
                    _shownotification = WaitForElementIsVisible(By.XPath(string.Format(SHOW, " Show notification ")));
                    if (_shownotification.GetAttribute("aria-checked") != value.ToString().ToLower())
                    {
                        _shownotification.Click();
                    }
                    break;

                case FilterType.ShowGuest:
                    _showGuest = WaitForElementIsVisible(By.XPath(string.Format(SHOW, " Show Guest ")));
                    if (_showGuest.GetAttribute("aria-checked") != value.ToString().ToLower())
                    {
                        _showGuest.Click();
                    }
                    break;

                case FilterType.ShowLeg:
                    _showLeg = WaitForElementIsVisible(By.XPath(string.Format(SHOW, " Show Leg ")));
                    if (_showLeg.GetAttribute("aria-checked") != value.ToString().ToLower())
                    {
                        _showLeg.Click();
                    }
                    break;

                case FilterType.ShowSpecMeal:
                    _specMeal = WaitForElementIsVisible(By.XPath(string.Format(SHOW, " Show spec. meal ")));
                    if (_specMeal.GetAttribute("aria-checked") != value.ToString().ToLower())
                    {
                        _specMeal.Click();
                    }
                    break;

                case FilterType.ShowWorkshop:
                    _showWorkShop = WaitForElementIsVisible(By.XPath(string.Format(SHOW, " Show Workshop ")));
                    if (_showWorkShop.GetAttribute("aria-checked") != value.ToString().ToLower())
                    {
                        WaitLoading();
                        new Actions(_webDriver).MoveToElement(_showWorkShop).Click().Perform();
                    }
                    break;

                case FilterType.ShowFlights_ByTabletFlightType:
                    var _showByFlightType = WaitForElementIsVisible(By.XPath(string.Format(SHOW, " Show flights - by tablet flight type ")));
                    if (_showByFlightType.GetAttribute("aria-checked") != value.ToString().ToLower())
                    {
                        _showByFlightType.Click();
                    }
                    break;

                case FilterType.Customer:
                    if (!isElementVisible(By.XPath("/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight/div/div[1]/flight-filters/div/div[11]/div[16]/ng-multiselect-dropdown/div/div[2]")))
                    {
                        var CustomerCombobox = WaitForElementIsVisible(By.XPath(string.Format(FILTER_COMBOBOX, "Customers")));
                        CustomerCombobox.Click();
                    }
                    if ((string)value == "UnSelect All")
                    {
                        var customerCheckBox = WaitForElementIsVisible(By.XPath(string.Format(FILTER_COMBOBOX_SELECT_ALL, "Customers")));
                        if (customerCheckBox.Selected)
                        {   //UnSelect All
                            new Actions(_webDriver).MoveToElement(customerCheckBox).Click().Perform();
                        }
                    }
                    else if ((string)value == "Select All")
                    {
                        var customerCheckBox = WaitForElementIsVisible(By.XPath(string.Format(FILTER_COMBOBOX_SELECT_ALL, "Customers")));
                        if (!customerCheckBox.Selected)
                        {   //Select All
                            new Actions(_webDriver).MoveToElement(customerCheckBox).Click().Perform();
                        }
                    }
                    else
                    {
                        var customerCheckBox = WaitForElementIsVisible(By.XPath(string.Format(FILTER_COMBOBOX_SELECT, "Customers", (string)value)));
                        if (!customerCheckBox.Selected)
                        {
                            new Actions(_webDriver).MoveToElement(customerCheckBox).Click().Perform();
                        }
                    }
                    var customerCombobox = WaitForElementIsVisible(By.XPath(string.Format(FILTER_COMBOBOX, "Customers")));
                    customerCombobox.Click();
                    //customerCombobox.Click();
                    break;
                case FilterType.Driver:
                    var driverCombobox = WaitForElementIsVisible(By.XPath(string.Format(FILTER_COMBOBOX, "Drivers")));
                    driverCombobox.Click();
                    if ((string)value == "UnSelect All")
                    {
                        var driverCheckBox = WaitForElementIsVisible(By.XPath(string.Format(FILTER_COMBOBOX_SELECT_ALL, "Drivers")));
                        if (driverCheckBox.Selected)
                        {   //UnSelect All
                            new Actions(_webDriver).MoveToElement(driverCheckBox).Click().Perform();
                        }
                    }
                    else if ((string)value == "Select All")
                    {
                        var driverCheckBox = WaitForElementIsVisible(By.XPath(string.Format(FILTER_COMBOBOX_SELECT_ALL, "Drivers")));
                        if (!driverCheckBox.Selected)
                        {   //Select All
                            new Actions(_webDriver).MoveToElement(driverCheckBox).Click().Perform();
                        }
                    }
                    else
                    {
                        Actions action = new Actions(_webDriver);
                        var driverCheckBox = WaitForElementExists(By.XPath(string.Format(FILTER_COMBOBOX_SELECT, "Drivers", (string)value)));
                        action.MoveToElement(driverCheckBox).Perform();
                        if (!driverCheckBox.Selected)
                        {
                            new Actions(_webDriver).MoveToElement(driverCheckBox).Click().Perform();
                        }
                    }
                    driverCombobox = WaitForElementIsVisible(By.XPath(string.Format(FILTER_COMBOBOX, "Drivers")));
                    driverCombobox.Click();
                    //driverCombobox.Click();
                    break;

                case FilterType.TabletFlightType:
                    var tabletFlightTypeCombobox = WaitForElementIsVisible(By.XPath(string.Format(FILTER_COMBOBOX, "Tablet Flight Types")));
                    tabletFlightTypeCombobox.Click();
                    if ((string)value == "UnSelect All")
                    {
                        var tabletFlightTypeCheckBox = WaitForElementIsVisible(By.XPath(string.Format(FILTER_COMBOBOX_SELECT_ALL, "Tablet Flight Types")));
                        if (tabletFlightTypeCheckBox.Selected)
                        {   //UnSelect All
                            new Actions(_webDriver).MoveToElement(tabletFlightTypeCheckBox).Click().Perform();
                        }
                    }
                    else if ((string)value == "Select All")
                    {
                        var tabletFlightTypeCheckBox = WaitForElementIsVisible(By.XPath(string.Format(FILTER_COMBOBOX_SELECT_ALL, "Tablet Flight Types")));
                        if (!tabletFlightTypeCheckBox.Selected)
                        {   //Select All
                            new Actions(_webDriver).MoveToElement(tabletFlightTypeCheckBox).Click().Perform();
                        }
                    }
                    else
                    {
                        var tabletFlightTypeCheckBox = WaitForElementIsVisible(By.XPath(string.Format(FILTER_COMBOBOX_SELECT, "Tablet Flight Types", (string)value)));
                        if (!tabletFlightTypeCheckBox.Selected)
                        {
                            new Actions(_webDriver).MoveToElement(tabletFlightTypeCheckBox).Click().Perform();
                        }
                    }
                    tabletFlightTypeCombobox = WaitForElementIsVisible(By.XPath(string.Format(FILTER_COMBOBOX, "Tablet Flight Types")));
                    tabletFlightTypeCombobox.Click();
                    // tabletFlightTypeCombobox.Click();
                    break;
                case FilterType.Workshop:
                    FlightTabletComboBox("Workshops", (string)value);
                    break;

                case FilterType.Search:
                    _searchfilter = WaitForElementIsVisibleNew(By.XPath(SEARCH_INPUT_FILTER));
                    _searchfilter.SetValue(ControlType.TextBox, value);
                    WaitForLoad();
                    break;

                case FilterType.EDTFrom:
                    if (isElementVisible(By.XPath(FILTER_ETD_FROM)))
                    {
                        _etdFrom = WaitForElementIsVisible(By.XPath(FILTER_ETD_FROM));
                        _etdFrom.Click();
                        _etdFrom.SendKeys(Keys.ArrowLeft);
                        _etdFrom.SendKeys(Keys.ArrowLeft);
                        _etdFrom.SendKeys((string)value);
                        _etdFrom.SendKeys((string)value2);
                        _etdFrom.SendKeys(Keys.ArrowRight);

                        WaitForLoad();
                    }
                    break;

                case FilterType.EDTTo:
                    if (isElementVisible(By.XPath(FILTER_ETD_TO)))
                    {
                        _etdTo = WaitForElementIsVisible(By.XPath(FILTER_ETD_TO));
                        _etdTo.Click();
                        _etdTo.SendKeys(Keys.ArrowLeft);
                        _etdTo.SendKeys(Keys.ArrowLeft);
                        _etdTo.SendKeys((string)value);
                        _etdTo.SendKeys((string)value2);
                        WaitForLoad();
                    }
                    break;

                case FilterType.ShowFlights_NotWithDoneStatus:
                    var _showNotWithDoneStatus = WaitForElementIsVisible(By.XPath(string.Format(SHOW, "Show flights - not with 'Done' status")));
                    if (_showNotWithDoneStatus.GetAttribute("aria-checked") != value.ToString().ToLower())
                    {
                        _showNotWithDoneStatus.Click();
                    }
                    WaitForLoad();

                    break;
                case FilterType.ShowFlights_WithAlertsOnly:
                    _showWithAlertOnly = WaitForElementIsVisible(By.XPath(string.Format(SHOW, " Show flights - with alerts only ")));
                    if (_showWithAlertOnly.GetAttribute("aria-checked") != value.ToString().ToLower())
                    {
                        _showWithAlertOnly.Click();
                    }
                    WaitForLoad();
                    break;

                case FilterType.ShowFlights_ValidatedOnly:
                    _showValidatedOnly = WaitForElementIsVisible(By.XPath(string.Format(SHOW, "Show flights - validated only")));
                    if (_showValidatedOnly.GetAttribute("aria-checked") != value.ToString().ToLower())
                    {
                        _showValidatedOnly.Click();
                    }
                    WaitForLoad();
                    break;

                case FilterType.ShowWithNotificationOnly:
                    _showWithNotificationOnly = WaitForElementIsVisible(By.XPath(string.Format(SHOW, "Show with notification only")));
                    if (_showWithNotificationOnly.GetAttribute("aria-checked") != value.ToString().ToLower())
                    {
                        _showWithNotificationOnly.Click();
                    }
                    WaitForLoad();
                    break;
                case FilterType.ShowFlights_ByflightDate:
                    var _showWithFlightDate = WaitForElementIsVisible(By.XPath(string.Format(SHOW, " Show flights - by flight date ")));
                    if (_showWithFlightDate.GetAttribute("aria-checked") != value.ToString().ToLower())
                    {
                        _showWithFlightDate.Click();
                    }
                    WaitForLoad();
                    break;
                case FilterType.ShowFlights_BarsetsToo:
                    var _showBarsetToo = WaitForElementIsVisible(By.XPath(string.Format(SHOW, "Show flights - barsets too")));
                    if (_showBarsetToo.GetAttribute("aria-checked") != value.ToString().ToLower())
                    {
                        _showBarsetToo.Click();
                    }
                    WaitForLoad();
                    break;
                case FilterType.ShowFlights_MajorOnly:
                    _showMajorOnly = WaitForElementIsVisible(By.XPath(string.Format(SHOW, "Show flights - major only")));
                    if (_showMajorOnly.GetAttribute("aria-checked") != value.ToString().ToLower())
                    {
                        _showMajorOnly.Click();
                    }
                    WaitForLoad();
                    break;
                case FilterType.ExpandAllFlights:
                    _expandallFlightsFilter = WaitForElementIsVisible(By.XPath(string.Format(SHOW, " Expand all flights ")));
                    if (_expandallFlightsFilter.GetAttribute("aria-checked") != value.ToString().ToLower())
                    {
                        _expandallFlightsFilter.Click();
                    }
                    WaitForLoad();
                    break;
                default:
                    break;
            }
            WaitPageLoading();
        }

        private void FlightTabletComboBox(string name, string value)
        {
            var workshopCombobox = WaitForElementIsVisible(By.XPath(string.Format(FILTER_COMBOBOX, name)));
            workshopCombobox.Click();
            if ((string)value == "UnSelect All")
            {
                if (isElementVisible(By.XPath(string.Format(FILTER_COMBOBOX_SELECT_ALL2, name, value))))
                {
                    var workshopCheckBox = WaitForElementIsVisible(By.XPath(string.Format(FILTER_COMBOBOX_SELECT_ALL2, name, value)));
                    if (workshopCheckBox.Selected)
                    {   //UnSelect All
                        new Actions(_webDriver).MoveToElement(workshopCheckBox).Click().Perform();
                    }
                }
                else
                {
                    workshopCombobox.Click();
                }
            }
            else if ((string)value == "Select All")
            {
                if (isElementVisible(By.XPath(string.Format(FILTER_COMBOBOX_SELECT_ALL2, name, value))))
                {
                    var workshopCheckBox = WaitForElementIsVisible(By.XPath(string.Format(FILTER_COMBOBOX_SELECT_ALL2, name, value)));
                    if (!workshopCheckBox.Selected)
                    {   //Select All
                        new Actions(_webDriver).MoveToElement(workshopCheckBox).Click().Perform();
                    }
                }
                else
                {
                    workshopCombobox.Click();
                }
            }
            else
            {
                var workshopCheckBox = WaitForElementIsVisible(By.XPath(string.Format(FILTER_COMBOBOX_SELECT, name, (string)value)));
                if (!workshopCheckBox.Selected)
                {
                    new Actions(_webDriver).MoveToElement(workshopCheckBox).Click().Perform();
                }
            }
            workshopCombobox = WaitForElementIsVisible(By.XPath(string.Format(FILTER_COMBOBOX, name)));
            // pourquoi ?
            workshopCombobox.Click();
        }

        public void CliqueSurOKFilterBtn()
        {
            try
            {
                _okFilterButton = WaitForElementIsVisibleNew(By.XPath(OK_FILTER_BUTTON));
                _okFilterButton.Click();
            }
            catch (ElementClickInterceptedException ex)
            {

                IJavaScriptExecutor js = (IJavaScriptExecutor)_webDriver;
                js.ExecuteScript("arguments[0].click();", _okFilterButton);
            }

            WaitPageLoading();
            WaitForLoad();
        }
        public void ClickFirstNotification()
        {

            _clicknotification = WaitForElementIsVisible(By.XPath(CLICKNOTIFICATION));
            _clicknotification.Click();


            WaitPageLoading();
            WaitForLoad();
        }
        public void Savenotificationforadd()
        {

            _savenotification = WaitForElementIsVisible(By.XPath(SAVEBUTTONFORNOTIFICATION));
            _savenotification.Click();


            WaitPageLoading();
            WaitForLoad();
        }
        public void AcceptJavaScriptConfirmation()
        {
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.AlertIsPresent());

            IAlert alert = _webDriver.SwitchTo().Alert();
            alert.Accept(); // Accepte la confirmation
        }

        public void ConfirmNotificationForAdd()
        {
            _confirmnotificationforadd = WaitForElementIsVisible(By.XPath(CONFIRMNOTIFICATIONFORADD));
            _confirmnotificationforadd.Click();

        }

        public void AddNotificationAfterText()
        {

            _notificationaftertext = WaitForElementIsVisible(By.XPath(ADDNOTIFICATIONAFTERTEXT));
            _notificationaftertext.Click();


            WaitPageLoading();
            WaitForLoad();
        }
        public void DeleteNotif()
        {

            _deletenotif = WaitForElementIsVisible(By.XPath(DELETENOTIF));
            _deletenotif.Click();


        }
        public void AddNotificationButton()
        {

            _addnotificationbutton = WaitForElementIsVisible(By.XPath(ADDNOTIFICATIONBUTTON));
            _addnotificationbutton.Click();


            WaitPageLoading();
            WaitForLoad();
        }
        public void AddNotificationText(string notificationtext)
        {
            _clickonchamps = WaitForElementIsVisible(By.XPath(CLICKONCHAMPS));
            _clickonchamps.Click();

            var notificatinotext = WaitForElementIsVisible(By.XPath(NOTIFICATIONTEXT));

            notificatinotext.SendKeys(notificationtext);


            WaitPageLoading();
            WaitForLoad();
        }

        public string GetFlightName(int row)
        {
            var first = WaitForElementIsVisible(By.XPath(string.Format(FLIGHT_NAME, row)));
            return first.Text.Trim();
        }
        public void PutStateToDone()
        {

            var button_start = WaitForElementIsVisible(By.XPath(FINAL_STATE_FIRST_LINE));
            if (button_start.Text.ToString() == "WARN.")
            {
                var button = WaitForElementIsVisible(By.XPath(HISTORY_BUTTON_FIRST_LINE));
                button.Click();
                WaitPageLoading();
            }

            if (button_start.Text.ToString() == "START")
            {
                button_start.Click();
                button_start = WaitForElementIsVisible(By.XPath(FINAL_STATE_FIRST_LINE));
                button_start.Click();
                WaitPageLoading();
            }
            else if (button_start.Text.ToString() == "STARTED")
            {
                button_start.Click();
                WaitPageLoading();
            }
            else
            {
                return;
            }
            WaitForLoad();
        }
        public void PutStateToStarted()
        {

            var button_start = WaitForElementIsVisible(By.XPath(FINAL_STATE_SECOND_LINE));
            if (button_start.Text.ToString() == "WARN.")
            {
                var button = WaitForElementIsVisible(By.XPath(HISTORY_BUTTON_SECOND_LINE));
                button.Click();
                WaitPageLoading();
            }
            if (button_start.Text.ToString() == "DONE")
            {
                button_start.Click();
                button_start = WaitForElementIsVisible(By.XPath(FINAL_STATE_SECOND_LINE));
                button_start.Click();
                WaitPageLoading();
            }
            else if (button_start.Text.ToString() == "START")
            {
                button_start.Click();
                WaitPageLoading();
            }
            else
            {
                return;
            }
            WaitForLoad();
        }
        public bool CheckFlightState(string fligthName, int row, string state)
        {
            var isStateCorrect = WaitForElementIsVisible(By.XPath(string.Format(FINAL_STATE_BUTTON, row))).Text.Trim().Equals(state);
            var fligthNameCorrect = WaitForElementIsVisible(By.XPath(string.Format(FLIGHT_NAME, row))).Text.Trim().Equals(fligthName);

            return isStateCorrect && fligthNameCorrect;
        }

        public void CliqueSurFilterIcone()
        {
            WaitForLoad();
            _filterBtn = WaitForElementIsVisibleNew(By.Id(FILTER_BUTTON));
            IJavaScriptExecutor executor = (IJavaScriptExecutor)_webDriver;
            executor.ExecuteScript("arguments[0].click();", _filterBtn);
            //new Actions(_webDriver).MoveToElement(_filterBtn).Click().Perform();
            WaitPageLoading();
        }
        public void SetFirstComment(string comment1, string comment2)
        {
            IWebElement buttonComment;
            if (isElementVisible(By.XPath(COMMENT_BUTTON_1)))
            {
                try
                {
                    buttonComment = WaitForElementIsVisible(By.XPath(COMMENT_BUTTON_1));
                    buttonComment.Click();
                    WaitForLoad();
                }
                catch (ElementClickInterceptedException ex)
                {
                    buttonComment = WaitForElementIsVisible(By.XPath(COMMENT_BUTTON_1));
                    IJavaScriptExecutor js = (IJavaScriptExecutor)_webDriver;
                    js.ExecuteScript("arguments[0].click();", buttonComment);
                }
            }
            else
            {
                buttonComment = WaitForElementIsVisible(By.XPath(COMMENT_BUTTON_2));
                buttonComment.Click();
                WaitForLoad();
            }

            var comment1input = WaitForElementIsVisible(By.XPath(COMMENT_INPUT_1));
            comment1input.Clear();
            comment1input.SetValue(ControlType.TextBox, comment1);
            var comment2input = WaitForElementIsVisible(By.XPath(COMMENT_INPUT_2));
            comment2input.Clear();
            comment2input.SetValue(ControlType.TextBox, comment2);
            WaitForLoad();
            IWebElement addButton;

            addButton = WaitForElementIsVisible(By.XPath(COMMENT_SAVE_BUTTON));

            addButton.Click();
            WaitPageLoading();
            WaitForLoad();
        }
        public void SetCommentForFlight(string flightNumber, string comment1, string comment2)
        {
            WaitPageLoading();
            string baseXPath = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight/div/div[2]/table/div/virtual-scroller/div[2]/tbody/tr";
            int rowCount = _webDriver.FindElements(By.XPath(baseXPath)).Count;

            for (int i = 1; i <= rowCount; i++)
            {
                string flightXPath = $"{baseXPath}[{i}]/td[4]";
                IWebElement flightElement = _webDriver.FindElement(By.XPath(flightXPath));

                if (flightElement.Text.Equals(flightNumber, StringComparison.OrdinalIgnoreCase))
                {
                    
                        string commentButtonXPath = $"{baseXPath}[{i}]/td[8]/div/mat-icon";
                        IWebElement buttonComment = WaitForElementIsVisible(By.XPath(commentButtonXPath));
                        buttonComment.Click();
                        WaitForLoad();
                    
                  
                    WaitPageLoading() ;
                    var comment1Input = WaitForElementIsVisible(By.XPath(COMMENT_INPUT_1));
                    comment1Input.Clear();
                    comment1Input.SendKeys(comment1);

                    WaitPageLoading();
                    var comment2Input = WaitForElementIsVisible(By.XPath(COMMENT_INPUT_2));
                    comment2Input.Clear();
                    comment2Input.SendKeys(comment2);

                    WaitForLoad();

                    IWebElement addButton = WaitForElementIsVisible(By.XPath(COMMENT_SAVE_BUTTON));
                    addButton.Click();
                    WaitPageLoading();
                    WaitForLoad();

                    break; 
                }
            }
        }


        public void GetFirstComment(out string comment1, out string comment2)
        {
            WaitPageLoading();
            IWebElement buttonComment;
            if (isElementVisible(By.XPath(COMMENT_BUTTON_1)))
            {
                buttonComment = WaitForElementIsVisible(By.XPath(COMMENT_BUTTON_1));
                buttonComment.Click();
                WaitForLoad();
            }
            else
            {
                buttonComment = WaitForElementIsVisible(By.XPath(COMMENT_BUTTON_2));
                buttonComment.Click();
                WaitForLoad();
            }

            var comment1input = WaitForElementIsVisible(By.XPath(COMMENT_INPUT_1));
            comment1 = comment1input.GetAttribute("value");
            var comment2input = WaitForElementIsVisible(By.XPath(COMMENT_INPUT_2));
            comment2 = comment2input.GetAttribute("value");
            IWebElement cancelButton;

            cancelButton = WaitForElementIsVisible(By.XPath(COMMENT_SAVE_BUTTON));

            cancelButton.Click();
            WaitForLoad();

        }

        protected new void WaitForLoad()
        {
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(61));

            Func<IWebDriver, bool> readyCondition = webDriver =>
                (bool)javaScriptExecutor.ExecuteScript("return (document.readyState == 'complete')");

            wait.Until(readyCondition);
        }

        public string GetFirstFlightNumber()
        {
            return WaitForElementIsVisible(By.XPath(FIRST_FLIGHT_NUMBER)).Text;
        }
        public DateTime GetFirstDateFlight()
        {
            var dateString = WaitForElementIsVisible(By.XPath("//*/td[contains(@class,'FlightDate')]")).Text;
            string day = dateString.Substring(0, 2);
            string month = ConvertMonthNameToNumber(dateString.Substring(3));
            int year = DateTime.Now.Year;
            var stringDate = day + "/" + month + "/" + year;
            DateTime date = DateTime.ParseExact(stringDate, "dd/MM/yyyy", null);

            return date;
        }
        public string ConvertMonthNameToNumber(string monthName)
        {
            // "Mai", "Juil."
            var month = monthName.Replace(".", "");
            for (int i = 0; i < monthNames.Length; i++)
            {
                if (monthNames[i].Contains(month) ||
                   monthNamesFR[i].Contains(month))
                {
                    var ind = i + 1;
                    return ind.ToString("00"); // Ajouter 1 car les mois sont indexés à partir de 1
                }
            }
            throw new ArgumentException("Nom du mois non valide", nameof(month));
        }

        public int GetNumberOfColumns()
        {
            var countColumns = _webDriver.FindElements(By.XPath(COUNT_COLUMNS));
            var number = countColumns.Count;
            return number;
        }
        public bool VerifyColonneAjoutee(string name)
        {
            WaitForLoad();
            return isElementVisible(By.XPath(string.Format(SHOW_COLUMN, name)));
        }

        public bool VerifNotif()
        {
            var customers = _webDriver.FindElements(By.XPath(VERIFNOTIFADDED));

            if (customers.Count == 0)
                return false;

            foreach (var customer in customers)
            {
                string backgroundColor = customer.GetCssValue("color");

                if (backgroundColor == "rgba(0, 128, 0, 1)")
                {
                    return true;
                }
            }

            return false;
        }


        public void DatePickerToday(DateTime date)
        {
            string xPathDatePicker1 = "//*/mat-datepicker-toggle[@data-mat-calendar='mat-datepicker-0']/button";
            var datePicker1 = WaitForElementIsVisible(By.XPath(xPathDatePicker1));
            datePicker1.Click();
            DateTime reference = DateUtils.Now.Date;
            while (reference.Year > date.Year)
            {
                reference = reference.AddMonths(-1);
                string xPathDatePicker1Month = "//*[@id=\"mat-datepicker-0\"]/mat-calendar-header/div/div/button[2]";
                var datePicker1Month = WaitForElementIsVisible(By.XPath(xPathDatePicker1Month));
                datePicker1Month.Click();
                WaitForLoad();
            }
            while (reference.Month > date.Month)
            {
                reference = reference.AddMonths(-1);
                string xPathDatePicker1Month = "//*[@id=\"mat-datepicker-0\"]/mat-calendar-header/div/div/button[2]";
                var datePicker1Month = WaitForElementIsVisible(By.XPath(xPathDatePicker1Month));
                datePicker1Month.Click();
                WaitForLoad();
            }
            while (reference.Year < date.Year)
            {
                reference = reference.AddMonths(1);
                string xPathDatePicker2Month = "//*[@id=\"mat-datepicker-0\"]/mat-calendar-header/div/div/button[3]";
                var datePicker1Month = WaitForElementIsVisible(By.XPath(xPathDatePicker2Month));
                datePicker1Month.Click();
                WaitForLoad();
            }
            while (reference.Month < date.Month)
            {
                reference = reference.AddMonths(1);
                string xPathDatePicker1Month = "//*[@id=\"mat-datepicker-0\"]/mat-calendar-header/div/div/button[3]";
                var datePicker1Month = WaitForElementIsVisible(By.XPath(xPathDatePicker1Month));
                datePicker1Month.Click();
                WaitForLoad();
            }
            int jour = date.Day;
            // on peut faire mieux, tirer le jour du premier coup
            string xPathDatePicker1Day;
            xPathDatePicker1Day = "//*[@id='mat-datepicker-0']//span[text()=' " + jour + " ']/parent::button";
            var datePicker1Day = WaitForElementIsVisible(By.XPath(xPathDatePicker1Day));
            datePicker1Day.Click();
            Thread.Sleep(2000);
            WaitForLoad();
            if (isElementVisible(By.Id("mat-datepicker-0")))
            {
                //parfois ne se repli pas
                datePicker1Day = WaitForElementIsVisible(By.XPath(xPathDatePicker1Day));
                datePicker1Day.Click();
            }
            WaitForLoad();
        }

        public bool IsAllCustomer(string customer)
        {
            var customers = _webDriver.FindElements(By.XPath("//*/td[contains(@class,'CustomerCode')]"));
            if (customers.Count == 0) return false;
            foreach (var c in customers)
            {
                if (c.Text != customer) return false;
            }

            return true;
        }
        public bool IsAllDriver(string driver)
        {
            var drivers = _webDriver.FindElements(By.XPath("//*/td[contains(@class,'driverColumn')]/p"));
            if (drivers.Count == 0) return false;
            foreach (var c in drivers)
            {
                if (c.Text != driver) return false;
            }

            return true;
        }

        public bool IsAllTabletFlightType(string tabletFlightType)
        {
            var allStars = _webDriver.FindElements(By.XPath("//*/star-image-svg"));
            if (allStars.Count == 0) return false;
            foreach (var c in allStars)
            {
                if (c.GetAttribute("title") != tabletFlightType) return false;
            }

            return true;
        }


        public int GetNumberFlightLines()
        {
            WaitForLoad();
            if (isElementVisible(By.XPath(LINES_FLIGHTS)))
            {
                var listLinesFlights = _webDriver.FindElements(By.XPath(LINES_FLIGHTS));
                var numberLinesFlights = listLinesFlights.Count;
                return numberLinesFlights;
            }
            return 0;
        }
        public bool VerifyDataExist()
        {
            var numberLinesFlights = GetNumberFlightLines();
            if (numberLinesFlights == 0)
            {
                return false;
            }
            return true;
        }

        public bool VerifySearchFilterList(string flight)
        {
            var flightColumns = _webDriver.FindElements(By.XPath(LINES_FLIGHTS));

            foreach (var column in flightColumns)
            {
                if (column.Text == flight)
                {
                    return true; 
                }
            }

            return false;
        }
        private void SetDay(int day)
        {
            ReadOnlyCollection<IWebElement> days;
            days = _webDriver.FindElements(By.XPath(DAYS));

            foreach (var element in days)
            {
                if (element.Text.Contains(day.ToString()))
                {
                    element.Click();
                    break;
                }
            }
        }
        private void SetMonth(string month)
        {
            var currentMonths = _webDriver.FindElements(By.XPath(MONTHS));
            foreach (var currentMonth in currentMonths)
            {
                if ((currentMonth.Text).ToUpper() == month)
                {
                    currentMonth.Click();
                    break;
                }
            }
        }
        private void SetYear(int year)
        {
            ReadOnlyCollection<IWebElement> years;
            years = _webDriver.FindElements(By.XPath(YEARS));
            foreach (var element in years)
            {
                if (int.Parse(element.Text) == year)
                {
                    element.Click();
                    break;
                }
            }
        }
        public void SetDate(DateTime date)
        {
            WaitPageLoading();
            _dateIcone = WaitForElementIsVisible(By.XPath(DATE_ICONE));
            _dateIcone.Click();
            WaitForLoad();

            CultureInfo ci = new CultureInfo("en-US");
            var jour = date.Day;
            var month = date.ToString("MMM", ci).ToUpper();
            var year = date.Year;
            var yearBtn = WaitForElementIsVisible(By.XPath(YEARS_BTN));
            yearBtn.Click();
            SetYear(year);
            SetMonth(month);
            SetDay(jour);
            WaitPageLoading();
            WaitForLoad();
        }
        public bool VerifyFilterETDFrom(string EDTFromH, string EDTFromM)
        {
            var etdColumns = _webDriver.FindElements(By.XPath(ETD_COLUMNS));
            foreach (var column in etdColumns)
            {
                var etdH = column.Text.Substring(0, 2);
                var columnEtdLength = column.Text.Length;
                var etdM = column.Text.Substring(columnEtdLength - 2, 2);
                if (int.Parse(etdH) < int.Parse(EDTFromH))
                {
                    return false;
                }
                if (int.Parse(etdM) < int.Parse(EDTFromM))
                {
                    return false;
                }
            }
            return true;
        }
        public bool VerifyFilterETDTo(string EDTToH, string EDTToM)
        {
            var etdColumns = _webDriver.FindElements(By.XPath(ETD_COLUMNS));
            foreach (var column in etdColumns)
            {
                var etdH = column.Text.Substring(0, 2);
                var columnEtdLength = column.Text.Length;
                var etdM = column.Text.Substring(columnEtdLength - 2, 2);
                if (int.Parse(etdH) > int.Parse(EDTToH))
                {
                    return false;
                }
                if (int.Parse(etdM) > int.Parse(EDTToM))
                {
                    return false;
                }
            }
            return true;
        }
        public void SaveView(string nameView)
        {
            WaitPageLoading();
            _saveViewButton = WaitForElementExists(By.XPath(SAVE_VIEW_BUTTON));

            try
            {
                _webDriver.ExecuteJavaScript("arguments[0].click();", _saveViewButton);

            }
            catch
            {
                new Actions(_webDriver).MoveToElement(_saveViewButton).Click().Perform();
            }

            WaitPageLoading();
            var viewInput = WaitForElementIsVisible(By.XPath(VIEW_INPUT));
            viewInput.SetValue(ControlType.TextBox, nameView);
            WaitPageLoading();
            var viewSave = WaitForElementIsVisible(By.XPath(VIEW_SAVE));
            viewSave.Click();
            WaitPageLoading();
        }
        public void ClickSurChooseAView()
        {
            if (!isElementVisible(By.XPath(VIEW)))
            {
                var chooseViewCombo = WaitForElementIsVisible(By.XPath(CHOOSE_VIEW_COMBO));
                //chooseViewCombo.Click();
                IJavaScriptExecutor js = (IJavaScriptExecutor)_webDriver;
                js.ExecuteScript("arguments[0].click();", chooseViewCombo);
            }
            WaitPageLoading();
        }
        public void UnclickSurChooseAView()
        {
            if (isElementVisible(By.XPath(VIEW)))
            {
                var chooseViewCombo = WaitForElementIsVisible(By.XPath(CHOOSE_VIEW_COMBO));
                IJavaScriptExecutor js = (IJavaScriptExecutor)_webDriver;
                js.ExecuteScript("arguments[0].click();", chooseViewCombo);
            }
            WaitForLoad();
        }
        public bool isViewExist(string viewName)
        {
            WaitForLoad();
            var chooseViewList = _webDriver.FindElements(By.XPath(VIEWS));

            if (chooseViewList.Count > 0)
            {
                foreach (var chooseView in chooseViewList)
                {
                    if (chooseView.Text == viewName)
                    {
                        return true;
                    }
                    continue;
                }
            }
            return false;
        }
        public bool VerifyViewSaveData(string viewName, string workshop, int nbreOfColumns)
        {
            WaitForLoad();
            var chooseViewList = _webDriver.FindElements(By.XPath(VIEWS));

            if (chooseViewList.Count > 0)
            {
                foreach (var chooseView in chooseViewList)
                {
                    if (chooseView.Text == viewName)
                    {
                        chooseView.Click();
                        break;
                    }
                    continue;
                }
            }
            WaitPageLoading();
            var viewSelected = WaitForElementIsVisible(By.XPath(VIEW_SELECTED));
            int columns = GetNumberOfColumns();
            WaitForLoad();
            if (viewSelected.Text.Contains(workshop))
            {
                if (columns == nbreOfColumns)
                {
                    return false;
                }
                return true;
            }
            else
            {
                if (columns != nbreOfColumns)
                {
                    return false;
                }
                return true;
            }
        }
        public void DeleteView(string nameView)
        {
            var chooseViewList = _webDriver.FindElements(By.XPath(VIEWS));

            if (chooseViewList.Count > 0)
            {
                foreach (var chooseView in chooseViewList)
                {
                    if (chooseView.Text == nameView)
                    {
                        chooseView.Click();
                        WaitForLoad();
                        break;
                    }
                    continue;
                }
                ClickSurChooseAView();
                WaitForLoad();
                if (isElementVisible(By.XPath(DELETE_VIEW_ICONE)))
                {
                    var deleteViewIcone = WaitForElementIsVisible(By.XPath(DELETE_VIEW_ICONE));
                    deleteViewIcone.Click();
                }
                WaitPageLoading();
            }
        }
        public bool DeleteAllViews()
        {
            var chooseViewList = _webDriver.FindElements(By.XPath(VIEWS));

            if (chooseViewList.Count > 0)
            {
                WaitPageLoading();
                var deleteViewIconeAll = _webDriver.FindElements(By.XPath(DELETE_VIEW_ICONE_ALL));
                //foreach (var deleteViewIcone in deleteViewIconeAll)
                for (var i = deleteViewIconeAll.Count - 1; i >= 0; i--)
                {
                    try
                    {
                        deleteViewIconeAll[i].Click();
                    }
                    catch (StaleElementReferenceException ex)
                    {
                        deleteViewIconeAll = _webDriver.FindElements(By.XPath(DELETE_VIEW_ICONE_ALL));
                        deleteViewIconeAll[i].Click();
                    }

                }
                WaitPageLoading();

                chooseViewList = _webDriver.FindElements(By.XPath(VIEWS));

                return chooseViewList.Count == 0;
            }
            return false;
        }
        public DetailFlightTabletAppPage EditFlightHaveServicesExceptXUFlights()
        {
            var listFlightsNbServices = _webDriver.FindElements(By.XPath(LIST_OF_NBR_SERVICES));
            int i = 0;
            for (i = 0; i < listFlightsNbServices.Count; i++)
            {
                if (int.Parse(listFlightsNbServices[i].Text) > 0)
                {
                    var flightName = WaitForElementExists(By.XPath(string.Format(FLIGHT_NAME, i + 1)));
                    if (!flightName.Text.Contains("Flight"))
                    {
                        continue;
                    }
                    var flightHaveServices = WaitForElementExists(By.XPath(string.Format(FLIGHT_DETAILS, i + 1)));
                    flightHaveServices.Click();
                    break;
                }
            }
            WaitPageLoading();
            WaitForLoad();
            return new DetailFlightTabletAppPage(_webDriver, _testContext, i);


        }
        public int checkNumberOfViews()
        {
            WaitPageLoading();
            var chooseViewList = _webDriver.FindElements(By.XPath(CHOOSE_VIEW_LIST));
            return chooseViewList.Count;
        }

        public List<string> GetFlightTypes()
        {
            List<string> liste = new List<string>();
            var stars = _webDriver.FindElements(By.XPath("//*/star-image-svg"));
            foreach (var star in stars)
            {
                liste.Add(star.GetAttribute("title"));
            }
            return liste;
        }

        public void WaitTabletFlightProgressBar()
        {
            // attente de la progress bar
            int compteur = 1;
            bool vueSablier = false;
            while (compteur <= 1000)
            {
                try
                {
                    _webDriver.FindElement(By.XPath("//*/div[@role='alert']"));
                    vueSablier = true;
                    break;
                }
                catch
                {
                    compteur++;
                }
            }

            // attente de la fin de la progress bar
            compteur = 1;

            while (compteur <= 600 && vueSablier)
            {
                try
                {
                    var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(1));
                    wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//*/div[@role='alert']")));
                    compteur++;
                    // Ophélie : ajout d'un sleep pour augmenter le temps d'attente (équivalent à 1 minute max au total)
                    Thread.Sleep(100);
                }
                catch
                {
                    vueSablier = false;
                }
            }

            if (vueSablier)
            {
                throw new Exception("Délai d'attente dépassé pour le chargement de la page.");
            }
            WaitForLoad();
        }

        public string GetInternalFlightRemarks(int iNoFLight)
        {
            var list = _webDriver.FindElements(By.XPath(COMMENT_BUTTONS));
            list[iNoFLight].Click();
            WaitForLoad();
            var text = WaitForElementIsVisible(By.XPath(COMMENT_INPUT_1));
            return text.GetAttribute("value");
        }

        public string GetDeliveryNoteComment(int iNoFLight)
        {
            var list = _webDriver.FindElements(By.XPath(COMMENT_BUTTONS));
            list[iNoFLight].Click();
            WaitForLoad();
            var text = WaitForElementIsVisible(By.XPath(COMMENT_INPUT_2));
            return text.GetAttribute("value");
        }

        public void ChangeStatus(string oldStatus)
        {
            if (HasStatus(oldStatus))
            {
                var changeStatus = WaitForElementIsVisible(By.XPath(string.Format(CHANGE_STATUS, oldStatus)));
                changeStatus.Click();
                WaitForLoad();
            }
        }

        public bool HasStatus(string status)
        {
            return isElementVisible(By.XPath(string.Format(CHANGE_STATUS, status)));
        }

        public bool IsSortedByDate(string date)
        {
            var dates = _webDriver.FindElements(By.XPath("//*/td[contains(@class,'FlightDate')]"));
            if (dates.Count == 0) return false;
            var frenchCulture = new CultureInfo("fr-FR");
            var englishCulture = new CultureInfo("en-US");
            foreach (var c in dates)
            {
                if (DateTime.TryParse(c.Text, frenchCulture, DateTimeStyles.None, out DateTime parsedDate))
                {
                    var formattedDate = parsedDate.ToString("dd-MMM", frenchCulture);

                    if (formattedDate != date)
                    { return false; }
                }
            }

            return true;
        }
        
        public void EditFirstFlight()
        {
            var _firstFlightEditBtnflightNames = _webDriver.FindElements(By.XPath(FIRST_FLIGHT_EDIT_BTN));
            _firstFlightEditBtn.Click();
            WaitForLoad();


        }
        public void EditFlightQTY(string qty)
        {
            var _firstFlightQty = WaitForElementIsVisible(By.XPath(FIRST_FLIGHT_QTY));
            _firstFlightQty.Click();

            var _firstFlightQtyInput = WaitForElementIsVisible(By.XPath(FIRST_FLIGHT_QTY_INPUT));
            _firstFlightQtyInput.SetValue(ControlType.TextBox, qty);


        }
        public void EditGuestQTY(string qty)
        {
            var _firstGuestInput = WaitForElementIsVisible(By.XPath(FIRST_GUEST_INPUT));
            _firstGuestInput.SetValue(ControlType.TextBox, qty);
            WaitForLoad();

        }
        public string GetFlightQTY()
        {

            var _firstFlightQty = WaitForElementIsVisible(By.XPath(FIRST_FLIGHT_QTY));
            _firstFlightQty.Click();
            var _firstFlightQtyInput = WaitForElementIsVisible(By.XPath(FIRST_FLIGHT_QTY_INPUT));
            string Qty = _firstFlightQtyInput.GetAttribute("value");
            return Qty;


        }
        public string GetGuestQTY()
        {
            var _firstGuestInput = WaitForElementIsVisible(By.XPath(FIRST_GUEST_INPUT));
            string Qty = _firstGuestInput.GetAttribute("value");
            return Qty;

        }

        public bool CheckAlertExist()
        {
            return isElementExists(By.XPath("/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight/div/div[2]/table/div/thead/tr/th[7]"));
        }

        public bool IsValidFlightWithServices()
        {
            var listFlightsNbServices = _webDriver.FindElements(By.XPath(LIST_FLIGHT_NB_SERVICES));
            int i = 0;
            for (i = 0; i < listFlightsNbServices.Count; i++)
            {
                if (int.Parse(listFlightsNbServices[i].Text) > 0)
                {
                    var flightName = WaitForElementIsVisible(By.XPath(string.Format(FLIGHT_NAME, i + 1)));
                    if (!flightName.Text.Contains("Flight"))
                    {
                        continue;
                    }
                    return true;

                }
            }
            return false;
        }
        public IEnumerable<string> GetFlights()
        {
            var flightsNames = _webDriver.FindElements(By.XPath(FLIGHTS));
            return flightsNames.Select(e => e.Text);
        }
        public void ClickSurChooseAViewDefault()
        {
            WaitPageLoading();
            if (isElementVisible(By.XPath(VIEW)))
            {
                var chooseViewCombo = WaitForElementIsVisible(By.XPath("/html/body/div[2]/div[2]/div/div/mat-option/span"));
                chooseViewCombo.Click();

            }
            else
            {
                WaitPageLoading();
                var input = WaitForElementIsVisible(By.XPath(CLICK_VIEW));
                input.Click();
                WaitPageLoading();
                var chooseViewCombo = WaitForElementIsVisible(By.XPath(DEFAULT_VIEW));
                chooseViewCombo.Click();
            }
            WaitPageLoading();
        }
        public void ClickSurChooseAView(string name)
        {
            var input = WaitForElementIsVisible(By.XPath(CLICK_VIEW));
            input.Click();

            var view = WaitForElementExists(By.XPath(VIEW));
            if (view.Displayed)
            {
                var chooseViewCombo = WaitForElementIsVisible(By.XPath($"/html/body/div[2]/div[2]/div//mat-option/span[contains(text(),'{name}')]"));
                chooseViewCombo.Click();
            }
            else
            {
                input = WaitForElementIsVisible(By.XPath(CLICK_VIEW));
                input.Click();
                var chooseViewCombo = WaitForElementIsVisible(By.XPath($"/html/body/div[2]/div[2]/div//mat-option/span[contains(text(),'{name}')]"));
                chooseViewCombo.Click();
            }
            WaitForLoad();
            WaitPageLoading();
        }

        public void clickonTimeBlock()
        {

            _clickFlight = WaitForElementIsVisible(By.XPath(CLICK_FLIGHT));
            _clickFlight.Click();
            WaitForLoad(); 

            _timeBlock = WaitForElementIsVisible(By.XPath(TIME_BLOCK));
            _timeBlock.Click();

            WaitPageLoading();
        }

        public bool VerifiedOrdreWorkshop(string ensambl, string corte, string history)
        {
            _ensambl = WaitForElementIsVisible(By.XPath(ENSAMBL));
            _corte = WaitForElementIsVisible(By.XPath(CORTE));
            _history = WaitForElementIsVisible(By.XPath(HISTORY));
            WaitForLoad();

            if ((_ensambl.Text == ensambl) && (_corte.Text == corte) && (_history.Text == history))
            {
                return true;
            }
            return false;
        }
        public string GetCustomer()
        {
            var customerName = WaitForElementIsVisible(By.XPath(CUSTOMER_NAME));
            return customerName.Text.Trim();
        }

        public int GetNumberOfFligths()
        {
            var nbFlights = WaitForElementIsVisible(By.XPath(NUMBER_OF_FLIGHTS));
            return int.Parse(nbFlights.Text.Trim());
        }

        public bool AreAllColumnTitlesAligned()
        {
            var headerElements = _webDriver.FindElements(By.XPath("/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight/div/div[2]/table/div/thead/tr/th"));

            if (headerElements.Count == 0)
            {
                return true;
            }
            int firstYPosition = headerElements[0].Location.Y;
            for (int i = 1; i < headerElements.Count; i++)
            {
                int currentYPosition = headerElements[i].Location.Y;

                if (currentYPosition != firstYPosition)
                {
                    return false;
                }
            }
            return true;
        }
        public void ScrollToEndOfHorizontalScroll()
        {
            var element = _webDriver.FindElement(By.XPath("/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight/div/div[2]"));
            ((IJavaScriptExecutor)_webDriver).ExecuteScript("arguments[0].scrollLeft = arguments[0].scrollWidth;", element);
        }

        public bool IsHorizontalScrollBarVisible()
        {
            var element = _webDriver.FindElement(By.XPath("/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight/div/div[2]"));
            var scrollWidth = element.GetAttribute("scrollWidth");
            var clientWidth = element.GetAttribute("clientWidth");

            return Convert.ToInt32(scrollWidth) > Convert.ToInt32(clientWidth);
        }
        public bool VerifyAddedColonnes()
        {
            var thElemeents =_webDriver.FindElements(By.ClassName("workshopColumn"));
            return thElemeents.Count > 0;

        }

        public List<string> GetAllFlightWithAlert(string flight)
        {
            WaitLoading();
            var alert = _webDriver.FindElements(By.XPath(string.Format("/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight/div/div[2]/table/div/virtual-scroller/div[2]/tbody/tr[td[3][contains(text(), '{0}')]]/th", flight)));
            WaitPageLoading();
            return alert.Select(e => e.Text).ToList();
        }

        public bool VerifyAllFlightWithAlert(string flight , string alert)
        {
            WaitLoading();
            var allFligth = GetAllFlightWithAlert(flight);
            WaitPageLoading();
            if (!allFligth.Contains(alert))
            {
                return false;
            }
            return true;
        }
        public bool IsFlightValidatedOnly()
        {
            var flightStates = _webDriver.FindElements(By.XPath(FLIGHTS_STATE));
            foreach (var state in flightStates)
            {
                if (!state.GetAttribute("class").Contains("active"))
                {
                    return false;
                }
            }
            return true;
        }
        public string GetFirstServiceName()
        {
            var service = WaitForElementIsVisibleNew(By.XPath(FIRST_SERVICE_NAME));
            return service.Text;
        } 
        public string GetFirstDatasheetName()
        {
            return WaitForElementIsVisibleNew(By.XPath(DATASHEET_NAME)).Text;
        } 
        public bool CheckFirstItem()
        {
            var items = _webDriver.FindElements(By.XPath(FIRST_ITEM_NAME));
            return items.Count()>=1;
        } 
        
        public void ClickFirstDatasheet()
        {
            var datasheet = WaitForElementIsVisibleNew(By.XPath(CLICK_DATASHEET));
            datasheet.Click();
        }
        public void ClickFirstRecipe()
        {
            _clickRecipe = WaitForElementIsVisibleNew(By.XPath(CLICK_RECIPE));
            _clickRecipe.Click();
        }
        public bool IsInvoiced()
        {
            var flight = WaitForElementIsVisibleNew(By.XPath("/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight/div/div[2]/table/div/virtual-scroller/div[2]/tbody/tr[1]/td[10]/ul/li[3]"));

            return flight.Text=="I";
        }

    }
}
