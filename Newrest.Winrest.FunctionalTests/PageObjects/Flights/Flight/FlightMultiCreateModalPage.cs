using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using iText.StyledXmlParser.Jsoup.Nodes;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Security.Policy;
using System.Threading;
using System.Windows;
using System.Windows.Forms;
using Keys = OpenQA.Selenium.Keys;


namespace Newrest.Winrest.FunctionalTests.PageObjects.Flights.Flight
{
    public class FlightMultiCreateModalPage : PageBase
    {
        public FlightMultiCreateModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //______________________________________ Constantes ____________________________________________________

        private const string ADD_NEW_RULE = "//*[@id=\"addMultipleFlightForm\"]/div[1]/div[2]/div[4]/div/div/div/a[2]";
        private const string SITE = "ddl-site";
        private const string CUSTOMER_NAME = "//*[@id=\"addMultipleFlightForm\"]/div[1]/div[2]/div[2]/div/div/div[1]/div";
        private const string LIST_LP = "//*[@id='addMultipleFlightForm']/div[2]/div/div/div[5]/div[1]/table/tbody/tr";

        private const string FLIGHT_NUMBER = "//*[@id=\"Rules_0_FlightRuleLegs_0_FlightNo\"]";
        private const string AIRPORT_FROM = "//*[@id=\"Rules_0_FlightRuleLegs_0_AirportFromId\"]";
        private const string AIRPORT_TO = "//*[@id=\"Rules_0_FlightRuleLegDataVM_0_AirportToId\"]";
        private const string START_DATE = "//*[@id=\"Rules_0_StartDate\"]";
        private const string END_DATE = "//*[@id=\"Rules_0_EndDate\"]";
        private const string MONDAY = "//*[@id=\"addMultipleFlightForm\"]/div[2]/div/div/div[3]/div[3]/div/div[1]";
        private const string TUESDAY = "//*[@id=\"addMultipleFlightForm\"]/div[2]/div/div/div[3]/div[3]/div/div[2]";
        private const string WEDNESDAY = "//*[@id=\"addMultipleFlightForm\"]/div[2]/div/div/div[3]/div[3]/div/div[3]";
        private const string THURSDAY = "//*[@id=\"addMultipleFlightForm\"]/div[2]/div/div/div[3]/div[3]/div/div[4]";
        private const string FRIDAY = "//*[@id=\"addMultipleFlightForm\"]/div[2]/div/div/div[3]/div[3]/div/div[5]";
        private const string SATURDAY = "//*[@id=\"addMultipleFlightForm\"]/div[2]/div/div/div[3]/div[3]/div/div[6]";
        private const string SUNDAY = "//*[@id=\"addMultipleFlightForm\"]/div[2]/div/div/div[3]/div[3]/div/div[7]";
        private const string AIRCRAFT = "//*[@id=\"Rules_0_AircraftId-selectized\"]";
        private const string AIRCRAFT_VALUE = "//*[@id=\"addMultipleFlightForm\"]/div[2]/div/div[1]/div[4]/div[3]/div/div[1]/div";

        private const string FLIGHT_NUMBER_DUPLIQUE = "//*[@id=\"Rules_1_FlightRuleLegs_0_FlightNo\"]";
        private const string AIRPORT_FROM_DUPLIQUE = "//*[@id=\"Rules_1_FlightRuleLegs_0_AirportFromId\"]";
        private const string AIRPORT_TO_DUPLIQUE = "//*[@id=\"Rules_1_FlightRuleLegDataVM_0_AirportToId\"]";
        private const string START_DATE_DUPLIQUE = "//*[@id=\"Rules_1_StartDate\"]";
        private const string END_DATE_DUPLIQUE = "//*[@id=\"Rules_1_EndDate\"]";
        private const string MONDAY_DUPLIQUE = "//*[@id=\"addMultipleFlightForm\"]/div[2]/div/div[2]/div[3]/div[3]/div/div[1]";
        private const string TUESDAY_DUPLIQUE = "//*[@id=\"addMultipleFlightForm\"]/div[2]/div/div[2]/div[3]/div[3]/div/div[2]";
        private const string WEDNESDAY_DUPLIQUE = "//*[@id=\"addMultipleFlightForm\"]/div[2]/div/div[2]/div[3]/div[3]/div/div[3]";
        private const string THURSDAY_DUPLIQUE = "//*[@id=\"addMultipleFlightForm\"]/div[2]/div/div[2]/div[3]/div[3]/div/div[4]";
        private const string FRIDAY_DUPLIQUE = "//*[@id=\"addMultipleFlightForm\"]/div[2]/div/div[2]/div[3]/div[3]/div/div[5]";
        private const string SATURDAY_DUPLIQUE = "//*[@id=\"addMultipleFlightForm\"]/div[2]/div/div[2]/div[3]/div[3]/div/div[6]";
        private const string SUNDAY_DUPLIQUE = "//*[@id=\"addMultipleFlightForm\"]/div[2]/div/div[2]/div[3]/div[3]/div/div[7]";
        private const string AIRCRAFT_DUPLIQUE = "//*[@id=\"Rules_1_AircraftId-selectized\"]";
        private const string AIRCRAFT_VALUE_DUPLIQUE = "//*[@id=\"addMultipleFlightForm\"]/div[2]/div/div[2]/div[4]/div[3]/div/div[1]/div";

        private const string DUPLICATE_BTN = "//*[@id=\"addMultipleFlightForm\"]/div[2]/div/div/div[1]/a[1]/span";
        private const string TRASH_ICON = "//*[@id=\"addMultipleFlightForm\"]/div[2]/div/div[1]/div[1]/a[2]/span";

        private const string VALIDATE_BTN = "//*[@id=\"modal-1\"]/div/div/div[3]/button[3]";
        private const string VALIDATE_BTN_DEV = "//*[@id=\"modal-1\"]/div[3]/button[3]";
        private const string ERROR_MESSAGE = "//*[@id=\"modal-1\"]/div/div/div[3]/div/p[1]/span";
        private const string ERROR_MESSAGE_DEV = "//*[@id=\"modal-1\"]/div[3]/div/p[1]/span";
        private const string GO_TOP_BTN = "//*[@id=\"scroll-to-top\"]";
        private const string GO_BOTTOM_BTN = "//*[@id=\"modal-1\"]/div[3]/button[2]";
        private const string ETA_HOUR = "//*[@id=\"addMultipleFlightForm\"]/div[2]/div/div[{0}]/div[4]/div[1]/input[2]";
        private const string ETA_MINUTE = "//*[@id=\"addMultipleFlightForm\"]/div[2]/div/div[{0}]/div[4]/div[1]/input[3]";
        private const string ETD_HOUR = "//*[@id=\"addMultipleFlightForm\"]/div[2]/div/div[{0}]/div[4]/div[2]/input[2]";
        private const string ETD_MINUTE = "//*[@id=\"addMultipleFlightForm\"]/div[2]/div/div[{0}]/div[4]/div[2]/input[3]";
        private const string LOADING_PLAN = "//*[@id='addMultipleFlightForm']/div[2]/div/div/div[5]/div[1]/table/tbody/tr";
        private const string LEG = "//*[@id=\"addMultipleFlightForm\"]/div[2]/div/div/div[5]/div[2]/table/tbody/tr[1]/td[1]";
        private const string GUEST = "//*[@id=\"addMultipleFlightForm\"]/div[2]/div/div/div[5]/div[2]/table/tbody/tr[1]/td[2]";
        private const string QTY = "Rules_0_LegsGuestsQuantities_0_Quantity";
        private const string ADD_NEW_BTN = "/html/body/div[3]/div/div/div[3]/button[2]";
        private const string MESSAGE_ERROR = "/html/body/div[3]/div/div/div[2]/div[1]/div/form/div[2]/div/div/div[2]/span";
        private const string SITE_FROM = "//*[@id=\"Rules_0_FlightRuleLegs_0_AirportFromId\"]/option[@selected]";
        private const string SITE_TO = "//*[@id=\"Rules_0_FlightRuleLegDataVM_0_AirportToId\"]/option[@selected]";
        private const string CREATE_BN = "/html/body/div[3]/div/div/div[3]/button[3]"; 
        //_________________________VARIABLES____________________________________________________

        [FindsBy(How = How.Id, Using = SITE)]
        private IWebElement _site;

        [FindsBy(How = How.XPath, Using = ADD_NEW_RULE)]
        private IWebElement _addNewRule;

        [FindsBy(How = How.XPath, Using = FLIGHT_NUMBER)]
        private IWebElement _flightNumber;

        [FindsBy(How = How.XPath, Using = FLIGHT_NUMBER_DUPLIQUE)]
        private IWebElement _flightNumberDuplique;

        [FindsBy(How = How.XPath, Using = MONDAY)]
        private IWebElement _mondayBtn;

        [FindsBy(How = How.XPath, Using = MONDAY_DUPLIQUE)]
        private IWebElement _mondayBtnDuplique;

        [FindsBy(How = How.XPath, Using = TUESDAY)]
        private IWebElement _tuesdayBtn;

        [FindsBy(How = How.XPath, Using = TUESDAY_DUPLIQUE)]
        private IWebElement _tuesdayBtnDuplique;

        [FindsBy(How = How.XPath, Using = WEDNESDAY)]
        private IWebElement _wednesdayBtn;

        [FindsBy(How = How.XPath, Using = WEDNESDAY_DUPLIQUE)]
        private IWebElement _wednesdayBtnDuplique;

        [FindsBy(How = How.XPath, Using = THURSDAY)]
        private IWebElement _thursdayBtn;

        [FindsBy(How = How.XPath, Using = THURSDAY_DUPLIQUE)]
        private IWebElement _thursdayBtnDuplique;

        [FindsBy(How = How.XPath, Using = FRIDAY)]
        private IWebElement _fridayBtn;

        [FindsBy(How = How.XPath, Using = FRIDAY_DUPLIQUE)]
        private IWebElement _fridayBtnDuplique;

        [FindsBy(How = How.XPath, Using = SATURDAY)]
        private IWebElement _saturdayBtn;

        [FindsBy(How = How.XPath, Using = SATURDAY_DUPLIQUE)]
        private IWebElement _saturdayBtnDuplique;

        [FindsBy(How = How.XPath, Using = SUNDAY)]
        private IWebElement _sundayBtn;

        [FindsBy(How = How.XPath, Using = SUNDAY_DUPLIQUE)]
        private IWebElement _sundayBtnDuplique;

        [FindsBy(How = How.XPath, Using = AIRCRAFT)]
        private IWebElement _aircraftBtn;

        [FindsBy(How = How.XPath, Using = AIRCRAFT_DUPLIQUE)]
        private IWebElement _aircraftBtnDuplique;

        [FindsBy(How = How.XPath, Using = AIRPORT_FROM)]
        private IWebElement _airportFrom;

        [FindsBy(How = How.XPath, Using = AIRPORT_FROM_DUPLIQUE)]
        private IWebElement _airportFromDuplique;

        [FindsBy(How = How.XPath, Using = AIRPORT_TO)]
        private IWebElement _airportTo;

        [FindsBy(How = How.XPath, Using = AIRPORT_TO_DUPLIQUE)]
        private IWebElement _airportToDuplique;

        [FindsBy(How = How.XPath, Using = START_DATE)]
        private IWebElement _startDate;

        [FindsBy(How = How.XPath, Using = START_DATE_DUPLIQUE)]
        private IWebElement _startDateDuplique;

        [FindsBy(How = How.XPath, Using = END_DATE)]
        private IWebElement _endDate;

        [FindsBy(How = How.XPath, Using = END_DATE_DUPLIQUE)]
        private IWebElement _endDateDuplique;

        [FindsBy(How = How.XPath, Using = DUPLICATE_BTN)]
        private IWebElement _duplicateBtn;

        [FindsBy(How = How.XPath, Using = TRASH_ICON)]
        private IWebElement _trashIcon;

        [FindsBy(How = How.XPath, Using = VALIDATE_BTN)]
        private IWebElement _validateBtn;

        [FindsBy(How = How.XPath, Using = ERROR_MESSAGE)]
        private IWebElement _errorMessage;

        [FindsBy(How = How.XPath, Using = CUSTOMER_NAME)]
        private IWebElement _customer;


        [FindsBy(How = How.XPath, Using = GO_TOP_BTN)]
        private IWebElement _goTopBtn;

        [FindsBy(How = How.XPath, Using = GO_BOTTOM_BTN)]
        private IWebElement _goBottomBtn;

        [FindsBy(How = How.XPath, Using = LEG)]
        private IWebElement _leg;

        [FindsBy(How = How.XPath, Using = GUEST)]
        private IWebElement _guest;

        [FindsBy(How = How.Id, Using = QTY)]
        private IWebElement _qty;

        [FindsBy(How = How.XPath, Using = ADD_NEW_BTN)]
        private IWebElement _addNewBtn;

        [FindsBy(How = How.XPath, Using = MESSAGE_ERROR)]
        private IWebElement _messageErrror;

        [FindsBy(How = How.XPath, Using = SITE_FROM)]
        private IWebElement _siteFrom;

        [FindsBy(How = How.XPath, Using = SITE_TO)]
        private IWebElement _siteto;

        // ____________________________________________ Methodes ____________________________________________

        public void FillField_CreatMultiFlight(string site, string flightNumber, string airportTo, string aircraft, DateTime? startDate = null, DateTime? endDate = null, string customer = null)
        {
            DateTime actualStartDate = startDate ?? DateUtils.Now;
            DateTime actualEndDate = endDate ?? DateUtils.Now.AddDays(7);
            _site = WaitForElementIsVisible(By.Id(SITE));
            _site.SetValue(ControlType.DropDownList, site);
            WaitForLoad();
            if (!string.IsNullOrEmpty(customer))
            {
                _customer = WaitForElementIsVisible(By.XPath(CUSTOMER_NAME));
                if (_customer.TagName == "div")
                {
                    var actions = new Actions(_webDriver);
                    var inputField = WaitForElementIsVisible(By.XPath(CUSTOMER_NAME));

                    actions.MoveToElement(inputField).Click().Perform();
                    actions.Click(inputField).Perform();
                    actions.SendKeys(Keys.Delete).Perform();
                    actions.SendKeys(customer).Perform();
                    WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(10));
                    wait.Until(drv => drv.FindElement(By.XPath("//div[@class='item' and @data-value]")).Displayed);
                    var firstOption = WaitForElementIsVisible(By.XPath("//*[@id='addMultipleFlightForm']/div[1]/div[2]/div[2]/div/div/div[2]/div/div[1]"));
                    firstOption.Click();

                }
                else
                {
                    _customer.Clear();
                    _customer.SetValue(ControlType.TextBox, customer);
                }
            }
            _addNewRule = WaitForElementToBeClickable(By.XPath(ADD_NEW_RULE));
            _addNewRule.Click();
            WaitForLoad();

            _flightNumber = WaitForElementIsVisible(By.XPath(FLIGHT_NUMBER));
            _flightNumber.SetValue(ControlType.TextBox, flightNumber);

            _airportFrom = WaitForElementToBeClickable(By.XPath(AIRPORT_FROM));
            _airportFrom.SetValue(ControlType.DropDownList, site);

            _airportTo = WaitForElementToBeClickable(By.XPath(AIRPORT_TO));
            _airportTo.SetValue(ControlType.DropDownList, airportTo);

            _startDate = WaitForElementIsVisible(By.XPath(START_DATE));
            _startDate.SetValue(ControlType.DateTime, actualStartDate);

            _endDate = WaitForElementIsVisible(By.XPath(END_DATE));
            _endDate.SetValue(ControlType.DateTime, actualEndDate);

            CheckDays();

            _aircraftBtn = WaitForElementToBeClickable(By.XPath(AIRCRAFT));
            _aircraftBtn.SetValue(ControlType.TextBox, aircraft);
            _aircraftBtn.SendKeys(Keys.Enter);

            //validation of modal
            WaitForLoad();
        }

        private void CheckDays()
        {
            _mondayBtn = WaitForElementIsVisible(By.XPath(MONDAY));
            _mondayBtn.Click();

            _tuesdayBtn = WaitForElementIsVisible(By.XPath(TUESDAY));
            _tuesdayBtn.Click();

            _wednesdayBtn = WaitForElementIsVisible(By.XPath(WEDNESDAY));
            _wednesdayBtn.Click();

            _thursdayBtn = WaitForElementIsVisible(By.XPath(THURSDAY));
            _thursdayBtn.Click();

            _fridayBtn = WaitForElementIsVisible(By.XPath(FRIDAY));
            _fridayBtn.Click();

            _saturdayBtn = WaitForElementIsVisible(By.XPath(SATURDAY));
            _saturdayBtn.Click();

            _sundayBtn = WaitForElementIsVisible(By.XPath(SUNDAY));
            _sundayBtn.Click();
        }

        public FlightPage Validate()
        {
            _validateBtn = WaitForElementToBeClickable(By.XPath(VALIDATE_BTN_DEV));
            _validateBtn.Click();
            WaitForLoad();
            return new FlightPage(_webDriver, _testContext);
        }

        public bool IsErrorMessage()
        {
            if (isElementVisible(By.XPath(ERROR_MESSAGE)))
            {
                _errorMessage = _webDriver.FindElement(By.XPath(ERROR_MESSAGE));

                if (_errorMessage.Text == "You must define at least one rule")
                {
                    return true;
                }
            }
            else if (isElementVisible(By.XPath(ERROR_MESSAGE_DEV)))
            {
                _errorMessage = _webDriver.FindElement(By.XPath(ERROR_MESSAGE_DEV));

                if (_errorMessage.Text == "You must define at least one rule")
                {
                    return true;
                }
            }
            else
            {
                return false;
            }

            return false;
        }

        public void DuplicateFlight()
        {
            _duplicateBtn = WaitForElementToBeClickable(By.XPath(DUPLICATE_BTN));
            _duplicateBtn.Click();
            WaitForLoad();
        }

        public void RemoveFirstFlight()
        {
            _trashIcon = WaitForElementToBeClickable(By.XPath(TRASH_ICON));
            _trashIcon.Click();
            WaitForLoad();
        }

        public bool IsFirstFlightPresent()
        {
            if (isElementVisible(By.XPath(FLIGHT_NUMBER)))
            {
                _webDriver.FindElement(By.XPath(FLIGHT_NUMBER));
                return true;
            }
            else
            {
                return false;
            }
        }

        public bool CompareDuplicatedFlightNumber()
        {
            _flightNumber = WaitForElementExists(By.XPath(FLIGHT_NUMBER));
            string initValue = _flightNumber.GetAttribute("value");

            _flightNumberDuplique = WaitForElementExists(By.XPath(FLIGHT_NUMBER_DUPLIQUE));
            string newValue = _flightNumberDuplique.GetAttribute("value");

            return initValue.Equals(newValue);
        }

        public bool CompareDuplicatedAirport()
        {
            _airportFrom = WaitForElementExists(By.XPath(AIRPORT_FROM));
            string initAirportFrom = _airportFrom.GetAttribute("value");

            _airportTo = WaitForElementExists(By.XPath(AIRPORT_TO));
            string initAirportTo = _airportTo.GetAttribute("value");

            _airportFromDuplique = WaitForElementExists(By.XPath(AIRPORT_FROM_DUPLIQUE));
            string newAirportFrom = _airportFromDuplique.GetAttribute("value");

            _airportToDuplique = WaitForElementExists(By.XPath(AIRPORT_TO_DUPLIQUE));
            string newAirportTo = _airportToDuplique.GetAttribute("value");

            return initAirportFrom.Equals(newAirportFrom) && initAirportTo.Equals(newAirportTo);
        }

        public bool CompareDate()
        {
            _startDate = WaitForElementExists(By.XPath(START_DATE));
            string initStartDate = _startDate.GetAttribute("value");

            _endDate = WaitForElementExists(By.XPath(END_DATE));
            string initEndDate = _endDate.GetAttribute("value");

            _startDateDuplique = WaitForElementExists(By.XPath(START_DATE_DUPLIQUE));
            string newStartDate = _startDateDuplique.GetAttribute("value");

            _endDateDuplique = WaitForElementExists(By.XPath(END_DATE_DUPLIQUE));
            string newEndDate = _endDateDuplique.GetAttribute("value");

            return initStartDate.Equals(newStartDate) && initEndDate.Equals(newEndDate);
        }

        public bool CompareAircraft()
        {
            var aircraftValue = WaitForElementExists(By.XPath(AIRCRAFT_VALUE));
            string initAircraft = aircraftValue.Text;

            var aircraftValueDuplique = WaitForElementExists(By.XPath(AIRCRAFT_VALUE_DUPLIQUE));
            string newAircraft = aircraftValueDuplique.Text;

            return initAircraft.Equals(newAircraft);
        }

        public bool CompareDays()
        {
            _mondayBtn = WaitForElementExists(By.XPath(MONDAY));
            string initMonday = _mondayBtn.GetAttribute("class");

            _tuesdayBtn = WaitForElementExists(By.XPath(TUESDAY));
            string initTuesday = _tuesdayBtn.GetAttribute("class");

            _wednesdayBtn = WaitForElementExists(By.XPath(WEDNESDAY));
            string initWednesday = _wednesdayBtn.GetAttribute("class");

            _thursdayBtn = WaitForElementExists(By.XPath(THURSDAY));
            string initThursday = _thursdayBtn.GetAttribute("class");

            _fridayBtn = WaitForElementExists(By.XPath(FRIDAY));
            string initFriday = _fridayBtn.GetAttribute("class");

            _saturdayBtn = WaitForElementExists(By.XPath(SATURDAY));
            string initSaturday = _saturdayBtn.GetAttribute("class");

            _sundayBtn = WaitForElementExists(By.XPath(SUNDAY));
            string initSunday = _sundayBtn.GetAttribute("class");


            // Dupliqués
            _mondayBtnDuplique = WaitForElementExists(By.XPath(MONDAY_DUPLIQUE));
            string newMonday = _mondayBtnDuplique.GetAttribute("class");

            _tuesdayBtnDuplique = WaitForElementExists(By.XPath(TUESDAY_DUPLIQUE));
            string newTuesday = _tuesdayBtnDuplique.GetAttribute("class");

            _wednesdayBtnDuplique = WaitForElementExists(By.XPath(WEDNESDAY_DUPLIQUE));
            string newWednesday = _wednesdayBtnDuplique.GetAttribute("class");

            _thursdayBtnDuplique = WaitForElementExists(By.XPath(THURSDAY_DUPLIQUE));
            string newThursday = _thursdayBtnDuplique.GetAttribute("class");

            _fridayBtnDuplique = WaitForElementExists(By.XPath(FRIDAY_DUPLIQUE));
            string newFriday = _fridayBtnDuplique.GetAttribute("class");

            _saturdayBtnDuplique = WaitForElementExists(By.XPath(SATURDAY_DUPLIQUE));
            string newSaturday = _saturdayBtnDuplique.GetAttribute("class");

            _sundayBtnDuplique = WaitForElementExists(By.XPath(SUNDAY_DUPLIQUE));
            string newSunday = _sundayBtnDuplique.GetAttribute("class");

            return initMonday.Equals(newMonday) && initTuesday.Equals(newTuesday)
                && initWednesday.Equals(newWednesday) && initThursday.Equals(newThursday)
                && initFriday.Equals(newFriday) && initSaturday.Equals(newSaturday)
                && initSunday.Equals(newSunday);
        }
        public void FillField_AddMultipleFlights(string site, string CustomerLpFilter)
        {
            _site = WaitForElementIsVisible(By.Id(SITE));

            _site.SetValue(ControlType.DropDownList, site);

            Thread.Sleep(2000);
            _customer = WaitForElementIsVisible(By.XPath(CUSTOMER_NAME));
            if (_customer.TagName == "div")
            {
                var actions = new Actions(_webDriver);
                var inputField = WaitForElementIsVisible(By.XPath(CUSTOMER_NAME));

                actions.MoveToElement(inputField).Click().Perform();
                actions.Click(inputField).Perform();
                actions.SendKeys(Keys.Delete).Perform();
                actions.SendKeys(CustomerLpFilter).Perform();
                WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(10));
                wait.Until(drv => drv.FindElement(By.XPath("//div[@class='item' and @data-value]")).Displayed);
                var firstOption = WaitForElementIsVisible(By.XPath("//*[@id='addMultipleFlightForm']/div[1]/div[2]/div[2]/div/div/div[2]/div/div[1]"));
                firstOption.Click();

            }
            else
            {
                _customer.Clear();
                _customer.SetValue(ControlType.TextBox, CustomerLpFilter);
            }


            _addNewRule = WaitForElementToBeClickable(By.XPath(ADD_NEW_RULE));
            _addNewRule.Click();

            WaitForLoad();
        }

        public void AircraftMultipleFlights(string aircraft)
        {
            _aircraftBtn = WaitForElementToBeClickable(By.XPath(AIRCRAFT));
            _aircraftBtn.SetValue(ControlType.TextBox, aircraft);
            _aircraftBtn.SendKeys(Keys.Enter);

            //validation of modal
            WaitPageLoading();
        }

        public bool VerifyGoTopBtn()
        {
            _addNewRule = WaitForElementToBeClickable(By.XPath(ADD_NEW_RULE));
            _addNewRule.Click();
            _addNewRule.Click();
            WaitForLoad();
            if (isElementVisible(By.XPath(ADD_NEW_RULE)))
            {
                var goTopBtn = WaitForElementIsVisible(By.XPath(GO_TOP_BTN));
                goTopBtn.Click();  
            }
            return isElementVisible(By.XPath(GO_BOTTOM_BTN));
        }

        public bool VerifyGoToBottomBtn()
        {
            _addNewRule = WaitForElementToBeClickable(By.XPath(ADD_NEW_RULE));
            _addNewRule.Click();
            _addNewRule.Click();
            WaitForLoad();
            if (isElementVisible(By.XPath(ADD_NEW_RULE)))
            {
                var goTopBtn = WaitForElementIsVisible(By.XPath(GO_BOTTOM_BTN));
                goTopBtn.Click();
            }
            return isElementVisible(By.XPath(CREATE_BN));
        }
        public void SetETA(string time, int ruleNumber)
        {
            var hour = time.Split(':')[0];
            var minute = time.Split(':')[1];
            var hourEta = WaitForElementIsVisible(By.XPath(string.Format(ETA_HOUR, ruleNumber)));
            hourEta.ClearElement();
            hourEta.SetValue(ControlType.TextBox, hour);
            hourEta.SendKeys(Keys.Tab);
            var minuteEta = WaitForElementIsVisible(By.XPath(string.Format(ETA_MINUTE, ruleNumber)));
            minuteEta.SetValue(ControlType.TextBox, minute);
            WaitForLoad();
        }
        public void SetETD(string time, int ruleNumber)
        {
            var hour = time.Split(':')[0];
            var minute = time.Split(':')[1];
            var hourEtd = WaitForElementIsVisible(By.XPath(string.Format(ETD_HOUR, ruleNumber)));
            hourEtd.ClearElement();
            hourEtd.SetValue(ControlType.TextBox, hour);
            hourEtd.SendKeys(Keys.Tab);
            var minuteEtd = WaitForElementIsVisible(By.XPath(string.Format(ETD_MINUTE, ruleNumber)));
            minuteEtd.SetValue(ControlType.TextBox, minute);
            minuteEtd.SendKeys(Keys.Enter);
        }

        public void pickMultipleLodingPlan()
        {
            var loadingPlans = _webDriver.FindElements(By.XPath(LOADING_PLAN));
            int i = 0;

            if (loadingPlans.Count >= 2)
            {
                loadingPlans[i].Click();
                WaitForLoad();
                i++;

                loadingPlans[i].Click();
                WaitForLoad();
                i++;

                WaitPageLoading();

            }
            var legExist = GetFirstLegDetails();

            while (legExist == "No leg associated" && i < loadingPlans.Count)
            {
                loadingPlans[i].Click();
                WaitForLoad();
                i++;
                WaitPageLoading();
                legExist = GetFirstLegDetails();
            }

        }
        public string GetFirstLegDetails()
        {
            _leg = WaitForElementIsVisible(By.XPath(LEG));
            return _leg.Text;
        }
        public string GetFirstGuestDetails()
        {
            _guest = WaitForElementIsVisible(By.XPath(GUEST));
            return _guest.Text;
        }
        public string GetQtyDetails()
        {
            _qty = _webDriver.FindElement(By.Id(QTY));
            return _qty.GetAttribute("value");
        }
        public void CheckLoadingPlan(string loadingPlanName)
        {
            _webDriver.Manage().Window.Maximize();


            var table = _webDriver.FindElement(By.XPath("//*[@id=\"addMultipleFlightForm\"]/div[2]/div/div/div[5]/div[1]/table/tbody"));


            int trIndex = -1;

            IList<IWebElement> rows = table.FindElements(By.TagName("tr"));

            for (int j = 0; j < rows.Count; j++)
            {
                IList<IWebElement> columns = rows[j].FindElements(By.TagName("td"));

                for (int i = 0; i < columns.Count; i++)
                {
                    if (columns[i].Text == loadingPlanName)
                    {
                        trIndex = j;
                        string xpath = $"/html/body/div[3]/div/div/div[2]/div[1]/div/form/div[2]/div/div/div[5]/div[1]/table/tbody/tr[{trIndex + 1}]/td[1]/input";
                        var checkboxLP = WaitForElementExists(By.XPath(xpath));
                        ((IJavaScriptExecutor)_webDriver).ExecuteScript("arguments[0].scrollIntoView(true);", checkboxLP);
                        try
                        {
                            var checkBox = WaitForElementIsVisible(By.XPath(xpath));
                            checkBox.SetValue(ControlType.CheckBox, true);
                        }
                        catch (ElementNotInteractableException)
                        {
                            ((IJavaScriptExecutor)_webDriver).ExecuteScript("arguments[0].click();", checkboxLP);
                        }
                        WaitPageLoading();
                        WaitForLoad();
                        break;

                    }
                    if (trIndex > 0)
                    {
                        break;

                    }
                }
            }





        }
        public FlightPage AddNew()
        {
            _addNewBtn = WaitForElementToBeClickable(By.XPath(ADD_NEW_BTN));
            _addNewBtn.Click();
            WaitForLoad();
            return new FlightPage(_webDriver, _testContext);
        }
        public bool AfficheMsgError(string messageError)
        {
            _messageErrror = WaitForElementIsVisible(By.XPath(MESSAGE_ERROR));
            var msgError = _messageErrror.Text;
            WaitForLoad();
            if (msgError.Contains(messageError))
            {

                return true;


            }

            else return false;
        }

        public bool UpdateDateFlightVerifInformation(string siteFrom, string flightNumber, string siteTo, DateTime startDate)
        {
            //siteFrom
            _siteFrom = WaitForElementIsVisible(By.XPath(SITE_FROM));
            // siteTo
            _siteto = WaitForElementToBeClickable(By.XPath(SITE_TO));
            //flightNumber
            _flightNumber = WaitForElementIsVisible(By.XPath(FLIGHT_NUMBER));


            string siteFrm = _siteFrom.GetAttribute("innerHTML").Trim();
            string number = _flightNumber.GetAttribute("value").Trim();
            string siteto = _siteto.GetAttribute("innerHTML").Trim();

            if (siteFrm == siteFrom && number == flightNumber && siteto == siteTo)
            {
                _startDate = WaitForElementIsVisible(By.XPath(START_DATE));
                _startDate.SetValue(ControlType.DateTime, startDate);

                CheckDays();
                WaitLoading();

                return true;
            }
            else return false;


        }

        public bool AddNewFlightMemory()
        {
            bool memoryDelai = false;
            for (int i = 0; i < 5; i++)
            {
                long memoryBefore = GC.GetTotalMemory(false);

                AddNew();

                long memoryAfter = GC.GetTotalMemory(false);

                memoryDelai = memoryAfter > memoryBefore;

            }
            return memoryDelai;
        }
        public bool IsLoadingPlanListLoaded()
        {
            try
            {           
                var loadingPlanRows = _webDriver.FindElements(By.XPath(LIST_LP));

                return loadingPlanRows.Count > 0;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
            catch (WebDriverTimeoutException)
            {
                return false;
            }
        }

        public void Create()
        {
            var create = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div/div[3]/button[3]"));
            create.Click();
        }


        public bool isVisibleYoumustdefineAtLeastOneRule()
        {
            var message= WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div/div[3]/div/p[1]/span"));
            if (message.Text == "You must define at least one rule")
            {
                return true;
            }
            return false;
        }
        public void ClickOnNewRule()
        {
            var newRule = WaitForElementIsVisible(By.XPath("//*[@id=\"addMultipleFlightForm\"]/div[1]/div[2]/div[4]/div/div/div/a[2]"));
            newRule.Click();
        }

        public string GetDate()
        {
            var date = WaitForElementIsVisible(By.Id("Rules_0_StartDate"));
            return date.GetAttribute("value");
        }
    }
}
