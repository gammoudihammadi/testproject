using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;


namespace Newrest.Winrest.FunctionalTests.PageObjects.Flights.Flight
{
    public class BarsetMultiCreateModalPage : PageBase
    {
        public BarsetMultiCreateModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //______________________________________ Constantes ____________________________________________________

        private const string ADD_NEW_RULE = "//*[@id=\"addMultipleFlightForm\"]/div[1]/div[2]/div[4]/div/div/div/a[2]";
        private const string SITE = "ddl-site";
        private const string CUSTOMER_NAME = "//*[@id=\"addMultipleFlightForm\"]/div[1]/div[2]/div[2]/div/div/div[1]/div"; 
        private const string FLIGHT_NUMBER = "//*[@id=\"Rules_0_FlightRuleLegs_0_FlightNo\"]";
        private const string AIRPORT_FROM = "//*[@id=\"Rules_0_FlightRuleLegs_0_AirportFromId\"]";
        private const string AIRPORT_TO = "//*[@id=\"Rules_0_FlightRuleLegDataVM_0_AirportToId\"]";
        private const string START_DATE = "//*[@id=\"Rules_0_StartDate\"]";
        private const string END_DATE = "//*[@id=\"Rules_0_EndDate\"]";
        private const string AIRCRAFT = "//*[@id=\"Rules_0_AircraftId-selectized\"]";
        //Days
        private const string MONDAY = "//*[@id=\"addMultipleFlightForm\"]/div[2]/div/div/div[3]/div[3]/div/div[1]";
        private const string TUESDAY = "//*[@id=\"addMultipleFlightForm\"]/div[2]/div/div/div[3]/div[3]/div/div[2]";
        private const string WEDNESDAY = "//*[@id=\"addMultipleFlightForm\"]/div[2]/div/div/div[3]/div[3]/div/div[3]";
        private const string THURSDAY = "//*[@id=\"addMultipleFlightForm\"]/div[2]/div/div/div[3]/div[3]/div/div[4]";
        private const string FRIDAY = "//*[@id=\"addMultipleFlightForm\"]/div[2]/div/div/div[3]/div[3]/div/div[5]";
        private const string SATURDAY = "//*[@id=\"addMultipleFlightForm\"]/div[2]/div/div/div[3]/div[3]/div/div[6]";
        private const string SUNDAY = "//*[@id=\"addMultipleFlightForm\"]/div[2]/div/div/div[3]/div[3]/div/div[7]";
        private const string MONDAY_CHECKED = "//*[@id=\"addMultipleFlightForm\"]/div[2]/div/div/div[3]/div[3]/div/div[1][contains(@class,  'btn-toggle btn-toggle-on')]";
        private const string TUESDAYY_CHECKED = "//*[@id=\"addMultipleFlightForm\"]/div[2]/div/div/div[3]/div[3]/div/div[2][contains(@class,  'btn-toggle btn-toggle-on')]";
        private const string WEDNESDAYY_CHECKED = "//*[@id=\"addMultipleFlightForm\"]/div[2]/div/div/div[3]/div[3]/div/div[3][contains(@class,  'btn-toggle btn-toggle-on')]";
        private const string THURSDAYY_CHECKED = "//*[@id=\"addMultipleFlightForm\"]/div[2]/div/div/div[3]/div[3]/div/div[4][contains(@class,  'btn-toggle btn-toggle-on')]";
        private const string FRIDAYY_CHECKED = "//*[@id=\"addMultipleFlightForm\"]/div[2]/div/div/div[3]/div[3]/div/div[5][contains(@class,  'btn-toggle btn-toggle-on')]";
        private const string SATURDAYY_CHECKED = "//*[@id=\"addMultipleFlightForm\"]/div[2]/div/div/div[3]/div[3]/div/div[6][contains(@class,  'btn-toggle btn-toggle-on')]";
        private const string SUNDAYY_CHECKED = "//*[@id=\"addMultipleFlightForm\"]/div[2]/div/div/div[3]/div[3]/div/div[7][contains(@class,  'btn-toggle btn-toggle-on')]";
        private const string CREATE_BTN = "//*[@id=\"modal-1\"]/div[3]/button[3]";
        private const string LP_CART = "Rules_0_LoadingPlanCartId";


        //_________________________VARIABLES____________________________________________________

        [FindsBy(How = How.Id, Using = SITE)]
        private IWebElement _site;

        [FindsBy(How = How.XPath, Using = ADD_NEW_RULE)]
        private IWebElement _addNewRule;

        [FindsBy(How = How.XPath, Using = FLIGHT_NUMBER)]
        private IWebElement _flightNumber;

        [FindsBy(How = How.XPath, Using = CUSTOMER_NAME)]
        private IWebElement _customer;

        [FindsBy(How = How.XPath, Using = AIRCRAFT)]
        private IWebElement _aircraftBtn;

        [FindsBy(How = How.XPath, Using = AIRPORT_FROM)]
        private IWebElement _airportFrom;

        [FindsBy(How = How.XPath, Using = AIRPORT_TO)]
        private IWebElement _airportTo;

        [FindsBy(How = How.XPath, Using = START_DATE)]
        private IWebElement _startDate;

        [FindsBy(How = How.XPath, Using = END_DATE)]
        private IWebElement _endDate;

        [FindsBy(How = How.XPath, Using = MONDAY)]
        private IWebElement _mondayBtn;

        [FindsBy(How = How.XPath, Using = TUESDAY)]
        private IWebElement _tuesdayBtn;

        [FindsBy(How = How.XPath, Using = WEDNESDAY)]
        private IWebElement _wednesdayBtn;

        [FindsBy(How = How.XPath, Using = THURSDAY)]
        private IWebElement _thursdayBtn;

        [FindsBy(How = How.XPath, Using = FRIDAY)]
        private IWebElement _fridayBtn;

        [FindsBy(How = How.XPath, Using = SATURDAY)]
        private IWebElement _saturdayBtn;

        [FindsBy(How = How.XPath, Using = SUNDAY)]
        private IWebElement _sundayBtn;

        [FindsBy(How = How.XPath, Using = CREATE_BTN)]
        private IWebElement _createBtn;

        [FindsBy(How = How.Id, Using = LP_CART)]
        private IWebElement _lpCart;

        public void FillField_CreateMultiBarsets(string site, string flightNumber, string aircraft, DateTime startDate , DateTime endDate, string customer ,string LpCartName)
        {
            _site = WaitForElementIsVisible(By.Id(SITE));
            _site.SetValue(ControlType.DropDownList, site);
            WaitForLoad();

            if (!string.IsNullOrEmpty(customer))
            {
                _customer = WaitForElementToBeClickable(By.XPath("//*[@id=\"addMultipleFlightForm\"]/div[1]/div[2]/div[2]/div/div"));
                _customer.Click();
                WaitLoading();
                var input = WaitForElementToBeClickable(By.XPath("//*[@id=\"ddl-customer-selectized\"]"));
                input.SetValue(ControlType.TextBox, customer);
                //var clickcustomer = WaitForElementIsVisible(By.XPath($"//*[@id='ddl-customer-selectized' and @value='{customer}']"));
                var clickcustomer = WaitForElementIsVisible(By.XPath("//*[@id=\"addMultipleFlightForm\"]/div[1]/div[2]/div[2]/div/div/div[2]/div/div[1]"));
                clickcustomer.Click();

            }
            _addNewRule = WaitForElementToBeClickable(By.XPath(ADD_NEW_RULE));
            _addNewRule.Click();
            WaitForLoad();

            _flightNumber = WaitForElementIsVisible(By.XPath(FLIGHT_NUMBER));
            _flightNumber.SetValue(ControlType.TextBox, flightNumber);

            _startDate = WaitForElementIsVisible(By.XPath(START_DATE));
            _startDate.SetValue(ControlType.DateTime, startDate);

            _endDate = WaitForElementIsVisible(By.XPath(END_DATE));
            _endDate.SetValue(ControlType.DateTime, endDate);
            _endDate.SendKeys(Keys.Enter);


            CheckDays();

            _aircraftBtn = WaitForElementToBeClickable(By.XPath(AIRCRAFT));
            _aircraftBtn.SetValue(ControlType.TextBox, aircraft);
            _aircraftBtn.SendKeys(Keys.Enter);
            WaitPageLoading();
            _aircraftBtn = WaitForElementToBeClickable(By.Id(LP_CART));
            _aircraftBtn.SetValue(ControlType.DropDownList, LpCartName);

            //validation of modal
            WaitForLoad();
        }

        private void CheckDays()
        {
            if (!isElementVisible(By.XPath(MONDAY_CHECKED)))
            {
                _mondayBtn = WaitForElementIsVisible(By.XPath(MONDAY));
                _mondayBtn.Click();
            }
            if (!isElementVisible(By.XPath(TUESDAYY_CHECKED)))
            {
                _tuesdayBtn = WaitForElementIsVisible(By.XPath(TUESDAY));
                _tuesdayBtn.Click();
            }
            if (!isElementVisible(By.XPath(WEDNESDAYY_CHECKED)))
            {
                _wednesdayBtn = WaitForElementIsVisible(By.XPath(WEDNESDAY));
                _wednesdayBtn.Click();
            }
            if (!isElementVisible(By.XPath(THURSDAYY_CHECKED)))
            {
                _thursdayBtn = WaitForElementIsVisible(By.XPath(THURSDAY));
                _thursdayBtn.Click();
            }
            if (!isElementVisible(By.XPath(FRIDAYY_CHECKED)))
            {
                _fridayBtn = WaitForElementIsVisible(By.XPath(FRIDAY));
                _fridayBtn.Click();
            }
            if (!isElementVisible(By.XPath(SATURDAYY_CHECKED)))
            {
                _saturdayBtn = WaitForElementIsVisible(By.XPath(SATURDAY));
                _saturdayBtn.Click();
            }
            if (!isElementVisible(By.XPath(SUNDAYY_CHECKED)))
            {
                _sundayBtn = WaitForElementIsVisible(By.XPath(SUNDAY));
                _sundayBtn.Click();
            }
        }

        public FlightPage CreateButton()
        {
            _createBtn = WaitForElementToBeClickable(By.XPath(CREATE_BTN));
            _createBtn.Click();
            WaitForLoad();
            return new FlightPage(_webDriver, _testContext);
        }


    }
}
