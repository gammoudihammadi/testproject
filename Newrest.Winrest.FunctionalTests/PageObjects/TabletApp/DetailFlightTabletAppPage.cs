using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using iText.StyledXmlParser.Jsoup.Internal;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Linq;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.TabletApp
{
    public class DetailFlightTabletAppPage : PageBase
    {
        int noFlight;
        public DetailFlightTabletAppPage(IWebDriver webDriver, TestContext testContext, int i) : base(webDriver, testContext)
        {
            noFlight = i;
        }

        //______________________________________CONSTANTES_____________________________________________________
        private const string FILTER_BUTTON = "//*[@id=\"mat-menu-panel-0\"]/div/button[3]";
        private const string ADD_GUEST_BUTTON = "//*[@id=\"mat-menu-panel-0\"]/div/button[2]";
        private const string ADD_SERVICE_BUTTON = "//*[@id=\"mat-menu-panel-0\"]/div/button[1]";
        private const string BACK_BUTTON = "//*[@id='mat-menu-panel-0']/div/a";
        private const string EXTENDED_BUTTON = "menuButtonDetail";
        private const string NOTIFICATION_FILTER_1 = "//*[@id=\"mat-mdc-slide-toggle-4-button\"]";
        private const string NOTIFICATION_FILTER_2 = "//*[@id=\"mat-mdc-slide-toggle-10-button\"]";
        private const string CLOCHE_ICONE = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-detail/div[4]/div/flight-detail-leg/div/div[2]/flight-detail-service-list/div/table/tr[2]/td[4]/div/mat-icon";
        private const string NOTIFICATION_FILTER_SELECTED = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-detail/flight-detail-filter/div/div/div[4]/mat-slide-toggle/div/button[@aria-checked = \"true\"]";
        private const string OK_FILTER_BUTTON = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-detail/flight-detail-filter/div/div/div[7]/button";
        private const string SHOW = "//*/label[text()=' {0} ']/../button";
        private const string ADD_NOTIFICATION_BUTTON_1 = "//*[@id=\"mat-mdc-dialog-0\"]/div/div/flight/div/div[3]/div/button[2]";
        private const string LIST_NOTIFICATIONS = "//*/mat-dialog-container/div/div/flight/div/div[2]/div[*]/div[2]/ckeditor/div[2]/div[2]/div/p";
        private const string ADD_NOTIFICATION_BUTTON_2 = "//*[@id=\"mat-mdc-dialog-1\"]/div/div/notifcreation/div/div[3]/div/button[2]";
        private const string CANCEL_NOTIFICATION_BUTTON = "/html/body/div[1]/div[2]/div/mat-dialog-container/div/div/flight/div/div[3]/div/button[1]";
        private const string SAVE_MODIF_NOTIFICATION = "/html/body/div[1]/div[2]/div/mat-dialog-container/div/div/flight/div/div[2]/div/div[2]/div/button";
        private const string ADD_NOTIFICATION_INPUT = "//*/ckeditor/div[2]/div[2]/div/p";
        private const string DELETE_NOTIFICATION_ICONE_ALL = "/html/body/div[1]/div[2]/div/mat-dialog-container/div/div/flight/div/div[2]/div[*]/div[1]/table/tr/th[5]/mat-icon[text() = \"delete\"]";
        private const string DELETE_NOTIFICATION_ICONE_1 = "//*/flight/div/div[2]/div[1]/div[1]/table/tr/th[5]/mat-icon";
        private const string DELETE_NOTIFICATION_ICONE_2 = "//*/flight/div/div[2]/div/div[1]/table/tr/th[5]/mat-icon";
        private const string ACTIVATE_NOTIFICATION_BUTTON = "//*/flight/div/div[2]/div/div[1]/table/tr/th[2]/button";
        private const string EDIT_BUTTON_DELIVERY_NOTE = "//*/mat-icon[text()='edit']";
        private const string EDIT_BUTTON_FLIGHT_REMARKS = "//*/mat-icon[text()=' edit ']";
        private const string EDIT_PLATEAU = "//*/flight-detail-service-list/div/table/tr[2]/td[6]/em";
        private const string COMMENT_INPUT_1 = "//*/div[text()='InternalFlightRemarks: ']//input";
        private const string COMMENT_INPUT_2 = "//*/div[text()='DeliveryNoteComment: ']//input";
        private const string COMMENT_SAVE = "//*/span[text()='Save']";
        private const string EDIT_BUTTON_NOTIFICATION = "//*/mat-icon[text() = \"edit\"]";
        private const string VALID_BUTTON_NOTIFICATION = "//*/mat-icon[text() = \"check_circle_outline\"]";
        private const string VALID_NOTIFICATION = "//*/mat-icon[@class = \"mat-icon notranslate shadow radius normal-icon material-icons mat-ligature-font Valid mat-icon-no-color\"]";
        private const string COLUMNS_TH_SERVICES = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-detail/div[5]/div/flight-detail-leg/div/div[2]/flight-detail-service-list/div/table/tr/th";
        private const string LIST_FLIGHT_BTN = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-detail/div[3]/div[1]/div/span";
        private const string LIST_FLIGHT_ITEM = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-detail/mat-sidenav/div/div[2]/div[*]/flight-list-item/div/div[1]";
        private const string FLIGHT_NAME = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-detail/div[3]/div[2]";
        private const string SERVICE_BTN = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-detail/div[4]/div/flight-detail-leg/div/div[2]/flight-detail-service-list/div/table/tr[2]/td[8]/cp-component-3button/div/button";
        private const string LIST_START_BTN = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-detail/div[5]/div/flight-detail-leg/div/div[1]/div/div[2]/table/tr[*]/td[4]/cp-component-3button/div/button";
        private const string LIST_START_SERVICES = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-detail/div[5]/div/flight-detail-leg/div/div[2]/flight-detail-service-list/div/table/tr[*]/td[8]/cp-component-3button/div/button";
        private const string START_SERVICES = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-detail/div[5]/div/flight-detail-leg/div/div[2]/flight-detail-service-list/div/table/tr[2]/td[7]/cp-component-3button/div/button";
        private const string IS_STARTED = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-detail/div[5]/div/flight-detail-leg/div/div[1]/div/div[2]/table/tr/td[4]/cp-component-3button/div/button[text() = \" Started \"]";
        private const string DONE_SERVICE = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-detail/div[5]/div/flight-detail-leg/div/div[2]/flight-detail-service-list/div/table/tr[2]/td[7]/cp-component-3button/div/button";
        private const string LIST_DONE_SERVICES = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-detail/div[5]/div/flight-detail-leg/div/div[2]/flight-detail-service-list/div/table/tr[*]/td[7]/cp-component-3button/div/button";
        private const string RECIPE_BUTTON = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-detail/div[5]/div/flight-detail-leg/div/div[2]/flight-detail-service-list/div/table/tr[2]/td[6]/em";
        private const string FOLD_UNFOLD_BUTTON = "//*/datasheet-shared-view/div/div/div/div[3]/div/button";
        private const string RECIPE_DETAIL = "//*/datasheet-shared-view/div/div/div/div[3]/mat-list/details";
        private const string ADD_GUEST_INPUT = "mat-dialog-container flight-detail input";
        private const string ADD_GUEST_ITEM = "div[role='listbox'] mat-option";
        private const string ADD_GUEST_POPUP_BUTTON = "mat-dialog-container flight-detail mat-dialog-actions button:last-child";
        private const string GUEST_LIST_ELEMENTS = ".leg-list > table > tr > td:first-child";
        private const string ADD_SERVICE_INPUT = "mat-dialog-container flight-detail input";
        private const string ADD_SERVICE_ITEM = "div[role='listbox'] mat-option";
        private const string ADD_SERVICE_POPUP_BUTTON = "mat-dialog-container flight-detail mat-dialog-actions button:last-child";
        private const string SERVICE_LIST_ELEMENTS = "flight-detail-service-list table.main-table tr.flightDetailLine td:first-child";
        private const string BUTTON_STATE_PATCH = "//*/flight-detail-service-list/div/table/tr[*]/td[8]/cp-component-3button/div/button";
        private const string BUTTON_STATE ="//*/flight-detail-service-list/div/table/tr[*]/td[8]/cp-component-3button/div/button";
        private const string VALIDATE_ADD_GUEST_ITEM = "//*[@id=\"mat-mdc-dialog-0\"]/div/div/flight-detail/div/div[3]/mat-dialog-actions/button[2]/span[2]";
        private const string CANCEL_GUEST_ITEM = "//*/span[@class='mdc-button__label'][text()='Cancel']";
        private const string SELECT_SERVICE_TO_ADD = "//*/label[text()='Search SERVICE']/../input";
        private const string SELECT_ADD_SERVICE = "//*[@id=\"mat-autocomplete-1\"]";
        private const string DELETEICOUN = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-detail/div[5]/div/flight-detail-leg/div/div[2]/flight-detail-service-list/div/table/tr[2]/td[7]";
        private const string DONE_BTN = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-detail/div[5]/div/flight-detail-leg/div/div[2]/flight-detail-service-list/div/table/tr[2]/td[8]/cp-component-3button/div/button";
        private const string EDIT_BUTTON_FLIGHT_REMARKS2 = "//*/mat-icon[text()=' add_circle_outline ']";
        private const string FLIGHT_LIST_ITEM = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-detail/mat-sidenav/div/div[2]/div[*]/flight-list-item/div/div[1]/span";
        private const string FLIGHT_TO_CHANGE = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-detail/mat-sidenav/div/div[2]/div[{0}]";
        private const string EACH_GUEST = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-detail/div[5]/div/flight-detail-leg/div/div[1]/div/div[2]/table/tr[{0}]/td[1]";
        private const string FLIGHT_DATE = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-detail/div[3]/div[3]/table/tr[2]/td[1]";
        private const string SAVE_NOTIFICATION = "/html/body/div[1]/div[2]/div/mat-dialog-container/div/div/flight/div/div[3]/div/button[3]";
        private const string START = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-detail/div[4]/div/flight-detail-leg/div/div[2]/flight-detail-service-list/div/table/tr[2]/td[8]/cp-component-3button/div/button";
        private const string ICON_NOTIFICATION = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-detail/div[5]/div/flight-detail-leg/div/div[2]/flight-detail-service-list/div/table/tr[2]/td[4]/div/mat-icon";
        private const string CLOSE_NOTIFICATION = "//*[@id=\"mat-mdc-dialog-2\"]/div/div/flight/div/div[3]/div/button[1]";
        private const string NOTIF_ICONE = "//td[@class='Notification']//mat-icon[contains(@class, 'mat-icon') and text()='notifications']";

        
        // ____________________________________ Variables _______________________________________________
        [FindsBy(How = How.XPath, Using = DONE_BTN)]
        private IWebElement _doneBtn;

        [FindsBy(How = How.XPath, Using = FILTER_BUTTON)]
        private IWebElement _filterBtn;

        [FindsBy(How = How.XPath, Using = DELETEICOUN)]
        private IWebElement _deletebtn;

        [FindsBy(How = How.XPath, Using = ADD_GUEST_BUTTON)]
        private IWebElement _addGuestBtn;

        [FindsBy(How = How.XPath, Using = ADD_SERVICE_BUTTON)]
        private IWebElement _addServiceBtn;

        [FindsBy(How = How.Id, Using = EXTENDED_BUTTON)]
        private IWebElement _extendedButton;

        [FindsBy(How = How.XPath, Using = NOTIFICATION_FILTER_1)]
        private IWebElement _notificationFilter;

        [FindsBy(How = How.XPath, Using = CLOCHE_ICONE)]
        private IWebElement _clocheIcone;

        [FindsBy(How = How.XPath, Using = OK_FILTER_BUTTON)]
        private IWebElement _OKFilterBtn;

        [FindsBy(How = How.XPath, Using = ADD_NOTIFICATION_BUTTON_1)]
        private IWebElement _addBtn1;

        [FindsBy(How = How.XPath, Using = ADD_NOTIFICATION_BUTTON_2)]
        private IWebElement _addBtn2;

        [FindsBy(How = How.XPath, Using = ADD_NOTIFICATION_INPUT)]
        private IWebElement _addInput;

        [FindsBy(How = How.XPath, Using = EDIT_BUTTON_NOTIFICATION)]
        private IWebElement _editNotificationBtn;

        [FindsBy(How = How.XPath, Using = LIST_FLIGHT_BTN)]
        private IWebElement _listFlightBtn;

        [FindsBy(How = How.XPath, Using = FLIGHT_NAME)]
        private IWebElement _flightName;

        [FindsBy(How = How.XPath, Using = NOTIF_ICONE)]
        private IWebElement _clocheIconeNotif;

        public enum FilterType
        {
            ShowOnlyIsCheckList,
            ShowProductionName,
            ShowByProdName,
            Notification,
            ShowDeliveryNoteComment,
            ShowInternalFlightRemarks,
        }

        public void Filter(FilterType filterType, object value)
        {
            switch (filterType)
            {
                case FilterType.Notification:
                    if (value.ToString() == true.ToString())
                    {
                        if (IsDev())
                        {
                            _notificationFilter = WaitForElementIsVisible(By.XPath(NOTIFICATION_FILTER_1));
                            _notificationFilter.SetValue(ControlType.CheckBox, value);
                        }
                        else
                        {
                            _notificationFilter = WaitForElementIsVisible(By.XPath("/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-detail/flight-detail-filter/div/div/div[4]/mat-slide-toggle/div/button/div[2]/div/div[3]"));
                            _notificationFilter.SetValue(ControlType.CheckBox, value);
                        }
                        
                    }
                    else
                    {
                        _notificationFilter = WaitForElementIsVisible(By.XPath(NOTIFICATION_FILTER_2));
                        _notificationFilter.SetValue(ControlType.CheckBox, true);
                    }
                    break;

                case FilterType.ShowDeliveryNoteComment:
                    var _showNoteComment = WaitForElementIsVisible(By.XPath(string.Format(SHOW, "Show Delivery note comment")));
                    if (_showNoteComment.GetAttribute("aria-checked") != value.ToString().ToLower())
                    {
                        _showNoteComment.Click();
                    }
                    break;

                case FilterType.ShowInternalFlightRemarks:
                    var _showRemarks = WaitForElementIsVisible(By.XPath(string.Format(SHOW, "Show Internal flight remarks")));
                    if (_showRemarks.GetAttribute("aria-checked") != value.ToString().ToLower())
                    {
                        _showRemarks.Click();
                    }
                    break;

                default:
                    break;
            }
            WaitPageLoading();
        }

        public void ShowExtendedMenu()
        {
            _extendedButton = WaitForElementIsVisibleNew(By.Id(EXTENDED_BUTTON));
            _extendedButton.Click();
        }

        public FlightTabletAppPage ClickBackButton()
        {
            var _backBtn = WaitForElementIsVisible(By.XPath(BACK_BUTTON));
            _backBtn.Click();
            WaitPageLoading();
            WaitForLoad();
            return new FlightTabletAppPage(_webDriver, _testContext);
        }

        public void ClickFilterButton()
        {
            _filterBtn = WaitForElementIsVisibleNew(By.XPath(FILTER_BUTTON));
            _filterBtn.Click();
        }
        public void ClickDeleteBtn()
        {
            _deletebtn = WaitForElementIsVisible(By.XPath(DELETEICOUN));
            _deletebtn.Click();
        }
        public void ClickAddGuestButton()
        {
            _addGuestBtn = WaitForElementIsVisible(By.XPath(ADD_GUEST_BUTTON));
            _addGuestBtn.Click();
        }
        public void ClickAddServiceButton()
        {
            _addGuestBtn = WaitForElementIsVisible(By.XPath(ADD_SERVICE_BUTTON));
            _addGuestBtn.Click();
        }
        public void ClickFilterOKBtn()
        {
            _OKFilterBtn = WaitForElementIsVisibleNew(By.XPath(OK_FILTER_BUTTON));
            _OKFilterBtn.Click();
        }
        public bool isNotificationFilterSelected()
        {
            if(isElementVisible(By.XPath(NOTIFICATION_FILTER_SELECTED)))
            {
                return true;
            }
            return false;
        }
        public void ClickSurClocheIcone()
        {
            WaitPageLoading();
            _clocheIcone = WaitForElementIsVisible(By.XPath(CLOCHE_ICONE));
            _clocheIcone.Click();
        }

        public void AddNotification(string notification)
        {
            WaitPageLoading();
            _addBtn1 = WaitForElementIsVisibleNew(By.XPath(ADD_NOTIFICATION_BUTTON_1));
            _addBtn1.Click();
            WaitPageLoading();
            _addInput = WaitForElementIsVisibleNew(By.XPath(ADD_NOTIFICATION_INPUT));
            for (var s=0;s<notification.Length; s++)
            {
                _addInput.SendKeys(notification[s].ToString());
            }
            WaitPageLoading();
            _addBtn2 = WaitForElementIsVisibleNew(By.XPath(ADD_NOTIFICATION_BUTTON_2));
            _addBtn2.Click();
            WaitPageLoading();
        }
        public bool VerifyNotificationAdded(string notification)
        {
            WaitPageLoading();
            ClickSurClocheIcone();
            WaitLoading();
            var listNotification = _webDriver.FindElements(By.XPath(LIST_NOTIFICATIONS));
            foreach(var notifi in listNotification)
            {
                WaitLoading();
                if (notifi.Text == notification)
                {
                    return true;
                }
            }
            return false;
        }
        public bool VerifyNotificationActive()
        {
            _clocheIcone = WaitForElementIsVisible(By.XPath(CLOCHE_ICONE));
            return _clocheIcone.GetAttribute("class").Contains("Error");
        }
        protected new void WaitForLoad()
        {
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(61));

            Func<IWebDriver, bool> readyCondition = webDriver =>
                (bool)javaScriptExecutor.ExecuteScript("return (document.readyState == 'complete')");

            wait.Until(readyCondition);
        }
        public void DeleteNotifications()
        {
            _webDriver.Navigate().Refresh();
            ClickSurClocheIcone();

            var deleteNotificationIconeAll = _webDriver.FindElements(By.XPath(DELETE_NOTIFICATION_ICONE_ALL));
            for (var i = 0; i < deleteNotificationIconeAll.Count; i++)
            {
                if(isElementVisible(By.XPath(DELETE_NOTIFICATION_ICONE_1)))
                {
                    var deleteIcone1 = WaitForElementIsVisible(By.XPath(DELETE_NOTIFICATION_ICONE_1));
                    deleteIcone1.Click();

                }
                else
                {
                    var deleteIcone2 = WaitForElementIsVisible(By.XPath(DELETE_NOTIFICATION_ICONE_2));
                    deleteIcone2.Click();
                }
                _webDriver.SwitchTo().Alert().Accept();
                _webDriver.Navigate().Refresh();
                ClickSurClocheIcone();
                WaitForLoad();
            }
        }
        public void SetInternalFlightRemarks(string text)
        {
            if(isElementVisible(By.XPath(EDIT_BUTTON_FLIGHT_REMARKS)))
            {
                var editButton = WaitForElementIsVisible(By.XPath(EDIT_BUTTON_FLIGHT_REMARKS));
                editButton.Click();
                WaitLoading();
            }
            else
            {
                var editButton = WaitForElementIsVisible(By.XPath(EDIT_BUTTON_FLIGHT_REMARKS2));
                editButton.Click();
                WaitLoading();
            }
            
            var input = WaitForElementIsVisible(By.XPath(COMMENT_INPUT_1));
            input.SetValue(ControlType.TextBox, text);
            WaitPageLoading();
            var saveButton = WaitForElementIsVisible(By.XPath(COMMENT_SAVE));
            saveButton.Click();
            WaitPageLoading();
        }

        public void SetDeliveryNoteComment(string text)
        {
            if(isElementVisible(By.XPath(EDIT_BUTTON_DELIVERY_NOTE)))
            {
                var editButton = WaitForElementIsVisible(By.XPath(EDIT_BUTTON_DELIVERY_NOTE));
                editButton.Click();
                WaitForLoad();
            }
            else
            {
                var addButton = WaitForElementIsVisible(By.XPath("//*/mat-icon[text()='add_circle_outline']"));
                addButton.Click();
                WaitForLoad();
            }

            var input = WaitForElementIsVisible(By.XPath(COMMENT_INPUT_2));
            input.SetValue(ControlType.TextBox, text);
            WaitForLoad();
            var saveButton = WaitForElementIsVisible(By.XPath(COMMENT_SAVE));
            saveButton.Click();
            WaitForLoad();
        }

        public int GetNoFlight()
        {
            return noFlight;
        }

        public PlateauFlightTabletAppPage EditPlateau()
        {
            var plateau = WaitForElementIsVisible(By.XPath(EDIT_PLATEAU));
            plateau.Click();
            WaitForLoad();
            return new PlateauFlightTabletAppPage(_webDriver, _testContext);
        }
        public void EditNotification(string notificationModif)
        {
            WaitPageLoading();
            _editNotificationBtn = WaitForElementIsVisibleNew(By.XPath(EDIT_BUTTON_NOTIFICATION));
            _editNotificationBtn.Click();

            _addInput = WaitForElementIsVisibleNew(By.XPath(ADD_NOTIFICATION_INPUT));
            _addInput.Clear();
            for (var s = 0; s < notificationModif.Length; s++)
            {
                _addInput.SendKeys(notificationModif[s].ToString());
            }

            var saveButton = WaitForElementIsVisibleNew(By.XPath(COMMENT_SAVE));
            saveButton.Click();
            WaitPageLoading();
        }
        public void ActivateNotification()
        {
            var activateBtn = WaitForElementIsVisible(By.XPath(ACTIVATE_NOTIFICATION_BUTTON));
            activateBtn.Click();

            var saveBtn = WaitForElementIsVisible(By.XPath(SAVE_NOTIFICATION));
            saveBtn.Click();
            WaitForLoad();
        }
        public bool VerifyNotificationEdited(string notification)
        {
            _addInput = WaitForElementIsVisibleNew(By.XPath(ADD_NOTIFICATION_INPUT));

            if (_addInput.Text != notification)
            {
                return true;
            }
            return false;
        }
        public void ValideNotification()
        {
            WaitPageLoading();
            _editNotificationBtn = WaitForElementIsVisibleNew(By.XPath(EDIT_BUTTON_NOTIFICATION));
            _editNotificationBtn.Click();

            var validBtn = WaitForElementIsVisibleNew(By.XPath(VALID_BUTTON_NOTIFICATION));
            validBtn.Click();
            WaitPageLoading();

            var saveButton = WaitForElementIsVisibleNew(By.XPath(COMMENT_SAVE));
            saveButton.Click();
            WaitPageLoading();

        }
        public bool VerifyNotificationValidee()
        {
            if(isElementVisible(By.XPath(VALID_NOTIFICATION)))
            {
                return true;
            }
            return false;
        }
        public bool NotifacationIsVisible()
        {
            try
            {
                var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(10));
                var element = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(
                    By.XPath("/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-detail/div[4]/div/flight-detail-leg/div/div[2]/flight-detail-service-list/div/table/tr[1]/th[4]")));

                return element.Text == "Notifications";
            }
            catch (WebDriverTimeoutException)
            {
                return false;
            }
        }

        public int GetNumberOfColumnsServiceTable()
        {
            WaitPageLoading();
            var nbreColumnsOfServices = _webDriver.FindElements(By.XPath(COLUMNS_TH_SERVICES));
            return nbreColumnsOfServices.Count;
        }
        public bool VerifyNotificationFiltre(int nbreBeforeFilter)
        {
            var nbreColumnsAfterFilter = GetNumberOfColumnsServiceTable();
            if(nbreColumnsAfterFilter == nbreBeforeFilter + 1)
            {
                return true;
            }
            return false;
        }
        public void ClickListFlight()
        {
            _listFlightBtn = WaitForElementIsVisible(By.XPath(LIST_FLIGHT_BTN));
            _listFlightBtn.Click();
            WaitForLoad();
        }
        public string getFlightNameDetailPage()
        {
            _flightName = WaitForElementIsVisible(By.XPath(FLIGHT_NAME));
            var splitted = _flightName.Text.Split(' ');
            Console.WriteLine(splitted[1]);
            return splitted[1];
        }
        public void ChangeFlight(string flightname)
        {
            WaitPageLoading();
            var listFlightItem = _webDriver.FindElements(By.XPath(FLIGHT_LIST_ITEM));
            var listFlightName = listFlightItem.Select(x => x.Text).ToList();
            for (int i = 0; i < listFlightName.Count; i++)
            {
                if (listFlightName[i] != flightname && listFlightName[0] != "")
                {
                    var flight = WaitForElementExists(By.XPath(string.Format(FLIGHT_TO_CHANGE, i+1)));
                    //IJavaScriptExecutor executor = (IJavaScriptExecutor)_webDriver;
                    //executor.ExecuteScript("arguments[0].click();", flight);
                    flight.Click();
                    WaitLoading();
                    WaitPageLoading();
                    break;
                }
            }
        }
        public void StartAllServices()
        {
            var startsButtons = _webDriver.FindElements(By.XPath(BUTTON_STATE));

            for (int i = 0; i < startsButtons.Count; i++)
            {
                var button = startsButtons[i];

                if (button.Displayed && button.Enabled) 
                {
                    button.Click();
                    WaitForLoad();
                }
            }
        }

        public void StartService()
        {
            var startBtn = WaitForElementIsVisibleNew(By.XPath(SERVICE_BTN));
            startBtn.Click();
            WaitForLoad();
        } 
        public void ClickOnStart()
        {
            var startBtn = WaitForElementIsVisible(By.XPath("/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-detail/div[4]/div/flight-detail-leg/div/div[1]/div/div[2]/table/tr/td[4]/cp-component-3button/div/button"));
            startBtn.Click();
            WaitForLoad();
        }  
        public void ClickOnStarted()
        {
            var startBtn = WaitForElementIsVisibleNew(By.XPath("/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-detail/div[4]/div/flight-detail-leg/div/div[1]/div/div[2]/table/tr/td[4]/cp-component-3button/div/button"));
            startBtn.Click();
            WaitForLoad();
        }  
        public void ClickOnDone()
        {
            var startBtn = WaitForElementIsVisible(By.XPath("/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-detail/div[4]/div/flight-detail-leg/div/div[1]/div/div[2]/table/tr/td[4]/cp-component-3button/div/button"));
            startBtn.Click();
            WaitForLoad();
        }

        public void ChangeServiceStatusTostart()
        {
            var startBtn = WaitForElementIsVisible(By.XPath(DONE_BTN));
            if (startBtn.Text == "Started")
            {
                startBtn.Click();
                Thread.Sleep(2000); 
                startBtn = WaitForElementIsVisible(By.XPath(DONE_BTN));
                startBtn.Click();
                Thread.Sleep(2000);
            }
            else if (startBtn.Text == "Done") {
                startBtn.Click();
                Thread.Sleep(2000);
            }

        }
        public bool isAlreadyStarted()
        {
            if (!isElementVisible(By.XPath(IS_STARTED)))
            {
                return false;
            }
            return true;
        }
        public bool isStartedAll()
        {
            IReadOnlyCollection<IWebElement> listStartServices = new List<IWebElement>();

            listStartServices = _webDriver.FindElements(By.XPath(BUTTON_STATE));

            bool result = false;
            foreach (var startService in listStartServices)
            {
                if (startService.Text != " Started " || startService.Text != "Started")
                {
                    result = false;
                }
                result = true;
            }
            return result;
        }
        public bool isStarted()
        {
            IWebElement startService;
            // secouse
            Thread.Sleep(2000);
            startService = WaitForElementIsVisible(By.XPath(BUTTON_STATE));

            if (startService.Text.Trim() == "Started")
            {
                return true;
            }
            return false;
        }
        public void DoneServices()
        {
            IWebElement startService;
            startService = _webDriver.FindElement(By.XPath(BUTTON_STATE));

            startService.Click();
            WaitForLoad();
        }
        public void RemoveState()
        {
            IWebElement startService;
            startService = _webDriver.FindElement(By.XPath("/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-detail/div[5]/div/flight-detail-leg/div/div[2]/flight-detail-service-list/div/table/tr[3]/td[8]/cp-component-3button/div/button"));

            startService.Click();
            WaitForLoad();
        }
        public bool isDoneAll()
        {
            ReadOnlyCollection<IWebElement> listDoneServices;
            listDoneServices = _webDriver.FindElements(By.XPath(BUTTON_STATE));

            bool result = false;
            foreach (var doneService in listDoneServices)
            {
                if (doneService.Text == "Done")
                {
                    result = true;
                }
                else
                {
                    result = false;
                }

            }
            return result;
        }
        public bool isDone()
        {
            try
            {
                IWebElement doneService = _webDriver.FindElement(By.XPath(BUTTON_STATE));

                return doneService.Text.Trim() == "Done";
            }
            catch (NoSuchElementException)
            {
                // Retourne false si l'élément n'est pas trouvé
                return false;
            }
        }
        public void ClickRecipeButton()
        {
            var recipeBtn = WaitForElementIsVisible(By.XPath(RECIPE_BUTTON));
            recipeBtn.Click();
        }
        public void ClickFoldUnfoldButton()
        {
            var foldUnfoldBtn = WaitForElementIsVisible(By.XPath(FOLD_UNFOLD_BUTTON));
            foldUnfoldBtn.Click();
        }
        public bool isRecipeDetailOpen()
        {
            var recipeDetail = WaitForElementIsVisible(By.XPath(RECIPE_DETAIL));
            return recipeDetail.GetAttribute("open") != null;
        }

        public void SelectGuestToAdd(string guest)
        {
            var addGuestInput = WaitForElementIsVisible(By.CssSelector(ADD_GUEST_INPUT));
            addGuestInput.SendKeys(guest);
            addGuestInput.SendKeys(Keys.Enter);
            WaitForLoad();
        }
        public void ValidateAddGuest()
        {
            var addGuestItem = WaitForElementIsVisible(By.XPath(VALIDATE_ADD_GUEST_ITEM));
            //addGuestItem.Click();
            IJavaScriptExecutor executor = (IJavaScriptExecutor)_webDriver;
            executor.ExecuteScript("arguments[0].click();", addGuestItem);
            WaitPageLoading();
            if (isElementVisible(By.XPath(VALIDATE_ADD_GUEST_ITEM)))
            {
                // guest déjà sélectionné (second lancement du test le même jour)
                var cancelGuestItem = WaitForElementExists(By.XPath(CANCEL_GUEST_ITEM));
                cancelGuestItem.Click();
            }
        }
        public bool VerifyGuestIsPresent(string guest)
        {
            var listElements = _webDriver.FindElements(By.CssSelector(GUEST_LIST_ELEMENTS));
            int i = listElements.Count + 1;
            List<string> list = new List<string>();
            for (var j = 1; j < i; j++)
            {
                var element = WaitForElementExists(By.XPath(string.Format(EACH_GUEST, j)));
                list.Add(element.Text);
            }
            if (!list.Contains(guest))
            {
                return false;
            }
            return true;
        }
        public void SelectGuest(string guest)
        {
            var listElements = _webDriver.FindElements(By.CssSelector(GUEST_LIST_ELEMENTS));
            int i = listElements.Count + 1;
            for(var j = 1; j<i; j++)
            {
                var element = WaitForElementExists(By.XPath(string.Format(EACH_GUEST, j)));
                if (element.Text == guest)
                {
                    element.Click();
                }
            }
        }
        public void SelectServiceToAdd(string service)
        {
            var addServiceInput = WaitForElementExists(By.XPath(SELECT_SERVICE_TO_ADD));
            addServiceInput.SendKeys(service);
            WaitForLoad();
            var addServiceItem = WaitForElementExists(By.XPath(string.Format(SELECT_ADD_SERVICE, service)));
            addServiceItem.Click();
            WaitForLoad();
        }
        public void ValidateAddService()
        {
            var addServicePopupButton = WaitForElementExists(By.CssSelector(ADD_SERVICE_POPUP_BUTTON));
            addServicePopupButton.Click();
            WaitForLoad();
        }
        public bool VerifyServiceIsPresent(string service)
        {
            var listElements = _webDriver.FindElements(By.CssSelector(SERVICE_LIST_ELEMENTS));
            foreach (var element in listElements)
            {
                if (element.Text == service)
                {
                    return true;
                }
            }
            return false;
        }
        public string GetFlightName()
        {
            _flightName = WaitForElementIsVisible(By.XPath(FLIGHT_NAME));
            var flightName = _flightName.Text.Replace("N° ","").Trim();
            return flightName;
        }
        public string GetFlightDate()
        {
            var flightDate = WaitForElementIsVisible(By.XPath(FLIGHT_DATE));
            var result = flightDate.Text.Trim();
            return result;
        }
        public void CloseNotification()
        {
            WaitPageLoading();
            var closeButton = WaitForElementIsVisible(By.XPath(CLOSE_NOTIFICATION));
            closeButton.Click();
        }
        public void ClickSurIconeNotification()
        {
            _clocheIcone = WaitForElementIsVisibleNew(By.XPath(ICON_NOTIFICATION));
            _clocheIcone.Click();
        }
        public bool VerifyNotificationisAdded(string notification)
        {
            WaitPageLoading();
            ClickSurIconeNotification();
            WaitLoading();
            var listNotification = _webDriver.FindElements(By.XPath(LIST_NOTIFICATIONS));
            foreach (var notifi in listNotification)
            {
                WaitLoading();
                if (notifi.Text == notification)
                {
                    return true;
                }
            }
            return false;
        }
        public void DeleteNotificationsInServiceFlight()
        {
            _webDriver.Navigate().Refresh();
            ClickSurIconeNotification();

            var deleteNotificationIconeAll = _webDriver.FindElements(By.XPath(DELETE_NOTIFICATION_ICONE_ALL));
            for (var i = 0; i < deleteNotificationIconeAll.Count; i++)
            {
                if (isElementVisible(By.XPath(DELETE_NOTIFICATION_ICONE_1)))
                {
                    var deleteIcone1 = WaitForElementIsVisibleNew(By.XPath(DELETE_NOTIFICATION_ICONE_1));
                    deleteIcone1.Click();

                }
                else
                {
                    var deleteIcone2 = WaitForElementIsVisibleNew(By.XPath(DELETE_NOTIFICATION_ICONE_2));
                    deleteIcone2.Click();
                }
                _webDriver.SwitchTo().Alert().Accept();
                _webDriver.Navigate().Refresh();
                ClickSurIconeNotification();
                WaitForLoad();
            }
        }
        public string GetButtonState()
        {
            IWebElement button = _webDriver.FindElement(By.XPath("/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-detail/div[4]/div/flight-detail-leg/div/div[1]/div/div[2]/table/tr/td[4]/cp-component-3button/div/button"));

            string buttonText = button.Text.Trim();

            // Retourner l'état selon le texte du bouton
            if (buttonText.Equals("Start", StringComparison.OrdinalIgnoreCase))
            {
                return "Start";
            }
            else if (buttonText.Equals("Started", StringComparison.OrdinalIgnoreCase))
            {
                return "Started";
            }
            else if (buttonText.Equals("Done", StringComparison.OrdinalIgnoreCase))
            {
                return "Done";
            }

            throw new InvalidOperationException($"Impossible de déterminer l'état du bouton. Texte trouvé : '{buttonText}'");
        }

        public void ClickSurIconeNotificationCloche()
        {
            WaitPageLoading();
            WaitForLoad();
            try
            {
                _clocheIconeNotif = WaitForElementIsVisible(By.XPath(NOTIF_ICONE));
                _clocheIconeNotif.Click();
            }
            catch (Exception ex)
            {
                throw new Exception($"Impossible de localiser l'icône de notification avec le XPath : {CLOCHE_ICONE}. Détails : {ex.Message}");
            }

        }
    }
}
