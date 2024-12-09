using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;


namespace Newrest.Winrest.FunctionalTests.PageObjects.Flights.Flight
{
    public class FlightCreateModalPage : PageBase
    {
        public FlightCreateModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // __________________________________ Constantes ____________________________________________________

        private const string GO_BOTTOM = "scroll-to-bottom";

        private const string DATE = "FlightDate";
        private const string FLIGHT_NUMBER = "Name";
        private const string DIV_CUSTOMER = "//*[@id=\"customer-form-group\"]/div/div[1]";
        private const string CUSTOMER_NAME = "dropdown-customer-selectized";
        private const string AIRPORT_FROM = "LegParts_0";
        private const string AIRPORT_FROM_SELECTED = "//*[@id=\"LegParts_0\"]/option[@selected]";
        private const string AIRPORT_TO = "LegParts_1";
        private const string AIRPORT_TO_SELECTED = "//*[@id=\"LegParts_1\"]/option[@selected]";
        private const string ADD_LEG = "//*[@id=\"createFormVipOrder\"]/div[4]/div[2]/div[2]/a[2]";
        private const string AIRPORT_LEG2 = "LegParts_2";
        private const string ETA_HOURS = "//*[@id=\"createFormVipOrder\"]/div[5]/div[1]/div/input[2]";
        private const string ETD_HOURS = "//*[@id=\"createFormVipOrder\"]/div[5]/div[2]/div/input[2]";
        private const string AIRCRAFT_NUMBER = "dropdown-aircraft-selectized";
        private const string CHECK_BOX_LP = "//*[@id=\"createFormVipOrder\"]/div[8]/div/table/tbody/tr[*]/td[2]/input[2][@value=\"{0}\"]/../../td[1]/div/input";
        private const string LPCART = "dropdown-lpcart";
        private const string FLIGHTTYPE = "/html/body/div[3]/div/div/div[2]/div/form/div[1]/div[6]/div[3]/div/select";
        private const string FLIGHTREMARKS = "//*[@id=\"InternalFlightRemarks\"]";

        private const string VALIDATE = "//*[@id=\"form-createdit-flight\"]/div[2]/button[2]";
        private const string VALIDATE_LABEL = "//*[@id=\"createFormVipOrder\"]/div[3]/div[1]/div/span";

        private const string FLIGHT_DATE = "FlightDate";

        // __________________________________ Variables ______________________________________________________

        [FindsBy(How = How.Id, Using = FLIGHT_NUMBER)]
        private IWebElement _flightNo;

        [FindsBy(How = How.XPath, Using = DIV_CUSTOMER)]
        private IWebElement _divCustomer;

        [FindsBy(How = How.XPath, Using = DATE)]
        private IWebElement _date;

        [FindsBy(How = How.Id, Using = CUSTOMER_NAME)]
        private IWebElement _customerName;

        [FindsBy(How = How.Id, Using = FLIGHTREMARKS)]
        private IWebElement _flightRemarks;

        [FindsBy(How = How.Id, Using = LPCART)]
        private IWebElement _lpCartName;

        [FindsBy(How = How.Id, Using = AIRPORT_FROM)]
        private IWebElement _airportFrom;

        [FindsBy(How = How.Id, Using = AIRPORT_FROM_SELECTED)]
        private IWebElement _airportFromSelected;

        [FindsBy(How = How.Id, Using = AIRPORT_TO)]
        private IWebElement _airportTo;

        [FindsBy(How = How.Id, Using = AIRPORT_TO_SELECTED)]
        private IWebElement _airportToSelected;

        [FindsBy(How = How.XPath, Using = ADD_LEG)]
        private IWebElement _addLeg;

        [FindsBy(How = How.Id, Using = AIRPORT_LEG2)]
        private IWebElement _airportLeg2;

        [FindsBy(How = How.Id, Using = AIRCRAFT_NUMBER)]
        private IWebElement _aircraftNumber;

        [FindsBy(How = How.XPath, Using = ETA_HOURS)]
        private IWebElement _etaHours;

        [FindsBy(How = How.XPath, Using = ETD_HOURS)]
        private IWebElement _etdHours;

        [FindsBy(How = How.XPath, Using = VALIDATE)]
        private IWebElement _validateBtn;

        [FindsBy(How = How.XPath, Using = VALIDATE_LABEL)]
        private IWebElement _validateLabel;

        // ___________________________________ Méthodes _____________________________________________

        public void FillField_CreatNewFlight(string flightNo, string customer, string aircraft, string airportFrom, string airportTo,
            string loadingPlanName = null, string etaHours = "00", string etdHours = "23", string lpCart = null, Nullable<DateTime> date = null)
        {
            // hidden : <button id="scroll-to-bottom" type="button" class="btn-icon-arrowBot"></button>
            //WaitForElementIsVisible(By.Id(GO_BOTTOM));
            //WaitForElementExists(By.Id(GO_BOTTOM));
            // animation
            //Thread.Sleep(2000);

            if (date != null)
            {
                DateTime dateNotNull = (DateTime)date;
                var dateInput = WaitForElementIsVisibleNew(By.Id(FLIGHT_DATE));
                dateInput.SetValue(ControlType.TextBox, dateNotNull.ToString("dd/MM/yyyy"));
                dateInput.SendKeys(Keys.Tab);
            }
            else
            {
                var dateInput = WaitForElementIsVisibleNew(By.Id(FLIGHT_DATE));
                dateInput.SetValue(ControlType.TextBox, DateUtils.Now.ToString("dd/MM/yyyy"));
                dateInput.SendKeys(Keys.Tab);
                WaitLoading();
            }

            if (flightNo != null)
            {
                _flightNo = WaitForElementIsVisibleNew(By.Id(FLIGHT_NUMBER));
                _flightNo.SetValue(ControlType.TextBox, flightNo);
            }

            _airportFrom = WaitForElementIsVisibleNew(By.Id(AIRPORT_FROM));
            _airportFromSelected = WaitForElementIsVisibleNew(By.XPath(AIRPORT_FROM_SELECTED));
            if (!_airportFromSelected.Text.Contains(airportFrom))
            {
                _airportFrom.SetValue(ControlType.DropDownList, airportFrom);
            }

            _airportTo = WaitForElementIsVisibleNew(By.Id(AIRPORT_TO));
            _airportToSelected = WaitForElementIsVisibleNew(By.XPath(AIRPORT_TO_SELECTED));
            if (!_airportToSelected.Text.Contains(airportTo))
            {
                _airportTo.SetValue(ControlType.DropDownList, airportTo);
            }

            if (customer != null)
            {
                _divCustomer = WaitForElementIsVisibleNew(By.XPath(DIV_CUSTOMER));
                _divCustomer.Click();
                _divCustomer.Click();

                _customerName = WaitForElementIsVisibleNew(By.Id(CUSTOMER_NAME));
                _customerName.SetValue(ControlType.TextBox, customer);
                _customerName.SendKeys(Keys.Enter);
                WaitLoading();
            }

            _aircraftNumber = WaitForElementIsVisibleNew(By.Id(AIRCRAFT_NUMBER));
            _aircraftNumber.SetValue(ControlType.TextBox, aircraft);
            _aircraftNumber.SendKeys(Keys.Tab);

            _etaHours = WaitForElementIsVisibleNew(By.XPath(ETA_HOURS));
            _etaHours.SendKeys(Keys.ArrowRight);
            _etaHours.SendKeys(Keys.ArrowRight);
            _etaHours.SendKeys(Keys.Backspace);
            _etaHours.SendKeys(Keys.Backspace);
            _etaHours.SendKeys(etaHours);

            _etdHours = WaitForElementIsVisibleNew(By.XPath(ETD_HOURS));
            _etdHours.SendKeys(Keys.ArrowRight);
            _etdHours.SendKeys(Keys.ArrowRight);
            _etdHours.SendKeys(Keys.Backspace);
            _etdHours.SendKeys(Keys.Backspace);
            _etdHours.SendKeys(etdHours);

            if (!string.IsNullOrEmpty(loadingPlanName))
            {
                var checkboxLP = WaitForElementExists(By.XPath(String.Format(CHECK_BOX_LP, loadingPlanName)));

                var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
                javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", checkboxLP);

                checkboxLP.SetValue(ControlType.CheckBox, true);
            }

            if (flightNo != null && flightNo.Contains("second flight"))
            {
                _date = WaitForElementIsVisibleNew(By.Id(DATE));
                _date.SetValue(ControlType.DateTime, DateUtils.Now.AddDays(10));
                _date.SendKeys(Keys.Tab);
                WaitForLoad();
            }

            if (flightNo != null && flightNo.Contains("For linked"))
            {
                _date = WaitForElementIsVisibleNew(By.Id(DATE));
                _date.SetValue(ControlType.DateTime, DateUtils.Now.AddDays(1));
                _date.SendKeys(Keys.Tab);
                WaitForLoad();
            }

            if (flightNo != null && flightNo.Contains("to Split"))
            {
                _addLeg = WaitForElementIsVisibleNew(By.XPath(ADD_LEG));
                _addLeg.Click();
                WaitForLoad();

                _airportLeg2 = WaitForElementIsVisibleNew(By.Id(AIRPORT_LEG2));
                _airportLeg2.SetValue(ControlType.DropDownList, airportFrom);
                _airportLeg2.SendKeys(Keys.Enter);
                WaitForLoad();
            }

            if (lpCart != null)
            {
                _lpCartName = WaitForElementIsVisibleNew(By.Id(LPCART));
                _lpCartName.SetValue(ControlType.DropDownList, lpCart);
                _lpCartName.SendKeys(Keys.Enter);
                WaitForLoad();
            }
            _validateBtn = WaitForElementToBeClickable(By.XPath(VALIDATE));
            _validateBtn.Click();
            WaitPageLoading();
            WaitLoading();
        }




        public void FillField_CreatNewFlightWithSelectionLP(string flightNo, string customer, string aircraft, string airportFrom, string airportTo,
          string loadingPlanName = null, string etaHours = "00", string etdHours = "23", string lpCart = null, DateTime? date = null)
        {
            // hidden : <button id="scroll-to-bottom" type="button" class="btn-icon-arrowBot"></button>
            //WaitForElementIsVisible(By.Id(GO_BOTTOM));
            //WaitForElementExists(By.Id(GO_BOTTOM));
            // animation
            Thread.Sleep(2000);

            if (date != null)
            {
                DateTime dateNotNull = (DateTime)date;
                var dateInput = WaitForElementIsVisible(By.Id("FlightDate"));
                dateInput.SetValue(ControlType.TextBox, dateNotNull.ToString("dd/MM/yyyy"));
                dateInput.SendKeys(Keys.Enter);
                WaitForLoad();
            }
            else
            {
                var dateInput = WaitForElementIsVisible(By.Id("FlightDate"));
                dateInput.SetValue(ControlType.TextBox, DateUtils.Now.ToString("dd/MM/yyyy"));
                dateInput.SendKeys(Keys.Enter);
                WaitForLoad();
            }

            if (flightNo != null)
            {
                _flightNo = WaitForElementIsVisible(By.Id(FLIGHT_NUMBER));
                _flightNo.SetValue(ControlType.TextBox, flightNo);
                WaitForLoad();
            }

            _airportFrom = WaitForElementIsVisible(By.Id(AIRPORT_FROM));
            _airportFromSelected = WaitForElementIsVisible(By.XPath(AIRPORT_FROM_SELECTED));
            if (!_airportFromSelected.Text.Contains(airportFrom))
            {
                _airportFrom.SetValue(ControlType.DropDownList, airportFrom);
            }

            _airportTo = WaitForElementIsVisible(By.Id(AIRPORT_TO));
            _airportToSelected = WaitForElementIsVisible(By.XPath(AIRPORT_TO_SELECTED));
            if (!_airportToSelected.Text.Contains(airportTo))
            {
                _airportTo.SetValue(ControlType.DropDownList, airportTo);
            }

            if (customer != null)
            {
                _divCustomer = WaitForElementIsVisible(By.XPath(DIV_CUSTOMER));
                _divCustomer.Click();
                _divCustomer.Click();

                _customerName = WaitForElementIsVisible(By.Id(CUSTOMER_NAME));
                _customerName.SetValue(ControlType.TextBox, customer);
                _customerName.SendKeys(Keys.Enter);
                WaitForLoad();
            }

            _aircraftNumber = WaitForElementIsVisible(By.Id(AIRCRAFT_NUMBER));
            _aircraftNumber.SetValue(ControlType.TextBox, aircraft);
            _aircraftNumber.SendKeys(Keys.Tab);

            _etaHours = WaitForElementIsVisible(By.XPath(ETA_HOURS));
            _etaHours.SendKeys(Keys.ArrowRight);
            _etaHours.SendKeys(Keys.ArrowRight);
            _etaHours.SendKeys(Keys.Backspace);
            _etaHours.SendKeys(Keys.Backspace);
            _etaHours.SendKeys(etaHours);

            _etdHours = WaitForElementIsVisible(By.XPath(ETD_HOURS));
            _etdHours.SendKeys(Keys.ArrowRight);
            _etdHours.SendKeys(Keys.ArrowRight);
            _etdHours.SendKeys(Keys.Backspace);
            _etdHours.SendKeys(Keys.Backspace);
            _etdHours.SendKeys(etdHours);

            if (!string.IsNullOrEmpty(loadingPlanName))
            {
                var checkboxLP = WaitForElementExists(By.XPath(String.Format(CHECK_BOX_LP, loadingPlanName)));

                var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
                javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", checkboxLP);
                WaitForLoad();

                checkboxLP.SetValue(ControlType.CheckBox, true);
            }

            if (flightNo != null && flightNo.Contains("second flight"))
            {
                _date = WaitForElementIsVisible(By.Id(DATE));
                _date.SetValue(ControlType.DateTime, DateUtils.Now.AddDays(10));
                _date.SendKeys(Keys.Tab);
                WaitForLoad();
            }

            if (flightNo != null && flightNo.Contains("For linked"))
            {
                _date = WaitForElementIsVisible(By.Id(DATE));
                _date.SetValue(ControlType.DateTime, DateUtils.Now.AddDays(1));
                _date.SendKeys(Keys.Tab);
                WaitForLoad();
            }

            if (flightNo != null && flightNo.Contains("to Split"))
            {
                _addLeg = WaitForElementIsVisible(By.XPath(ADD_LEG));
                _addLeg.Click();
                WaitForLoad();

                _airportLeg2 = WaitForElementIsVisible(By.Id(AIRPORT_LEG2));
                _airportLeg2.SetValue(ControlType.DropDownList, airportFrom);
                _airportLeg2.SendKeys(Keys.Enter);
                WaitForLoad();
            }

            if (lpCart != null)
            {
                _lpCartName = WaitForElementIsVisible(By.Id(LPCART));
                _lpCartName.SetValue(ControlType.DropDownList, lpCart);
                _lpCartName.SendKeys(Keys.Enter);
                WaitForLoad();
            }
            _validateBtn = WaitForElementToBeClickable(By.XPath(VALIDATE));
            _validateBtn.Click();
            WaitPageLoading();
            WaitForLoad();
        }



        public void FillField_CreatNewFlightWithComment(string flightNo, string customer, string aircraft, string airportFrom, string airportTo,
        string loadingPlanName = null, string etaHours = "00", string etdHours = "23", string lpCart = null, DateTime? date = null, string flightRemarks = "XxxxxX")
        {
            Thread.Sleep(2000);

            if (date != null)
            {
                DateTime dateNotNull = (DateTime)date;
                var dateInput = WaitForElementIsVisible(By.Id("FlightDate"));
                dateInput.SetValue(ControlType.TextBox, dateNotNull.ToString("dd/MM/yyyy"));
                dateInput.SendKeys(Keys.Enter);
                WaitForLoad();
            }
            else
            {
                var dateInput = WaitForElementIsVisible(By.Id("FlightDate"));
                dateInput.SetValue(ControlType.TextBox, DateUtils.Now.ToString("dd/MM/yyyy"));
                dateInput.SendKeys(Keys.Enter);
                WaitForLoad();
            }

            if (flightNo != null)
            {
                _flightNo = WaitForElementIsVisible(By.Id(FLIGHT_NUMBER));
                _flightNo.SetValue(ControlType.TextBox, flightNo);
                WaitForLoad();
            }

            _airportFrom = WaitForElementIsVisible(By.Id(AIRPORT_FROM));
            _airportFromSelected = WaitForElementIsVisible(By.XPath(AIRPORT_FROM_SELECTED));
            if (!_airportFromSelected.Text.Contains(airportFrom))
            {
                _airportFrom.SetValue(ControlType.DropDownList, airportFrom);
            }

            _airportTo = WaitForElementIsVisible(By.Id(AIRPORT_TO));
            _airportToSelected = WaitForElementIsVisible(By.XPath(AIRPORT_TO_SELECTED));
            if (!_airportToSelected.Text.Contains(airportTo))
            {
                _airportTo.SetValue(ControlType.DropDownList, airportTo);
            }

            if (customer != null)
            {
                _divCustomer = WaitForElementIsVisible(By.XPath(DIV_CUSTOMER));
                _divCustomer.Click();
                _divCustomer.Click();

                _customerName = WaitForElementIsVisible(By.Id(CUSTOMER_NAME));
                _customerName.SetValue(ControlType.TextBox, customer);
                _customerName.SendKeys(Keys.Enter);
                WaitForLoad();
            }

            _aircraftNumber = WaitForElementIsVisible(By.Id(AIRCRAFT_NUMBER));
            _aircraftNumber.SetValue(ControlType.TextBox, aircraft);
            _aircraftNumber.SendKeys(Keys.Tab);

            _etaHours = WaitForElementIsVisible(By.XPath(ETA_HOURS));
            _etaHours.SendKeys(Keys.ArrowRight);
            _etaHours.SendKeys(Keys.ArrowRight);
            _etaHours.SendKeys(Keys.Backspace);
            _etaHours.SendKeys(Keys.Backspace);
            _etaHours.SendKeys(etaHours);

            _etdHours = WaitForElementIsVisible(By.XPath(ETD_HOURS));
            _etdHours.SendKeys(Keys.ArrowRight);
            _etdHours.SendKeys(Keys.ArrowRight);
            _etdHours.SendKeys(Keys.Backspace);
            _etdHours.SendKeys(Keys.Backspace);
            _etdHours.SendKeys(etdHours);

            if (!string.IsNullOrEmpty(loadingPlanName))
            {
                var checkboxLP = WaitForElementExists(By.XPath(String.Format(CHECK_BOX_LP, loadingPlanName)));

                var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
                javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", checkboxLP);
                WaitForLoad();

                checkboxLP.SetValue(ControlType.CheckBox, true);
            }

            if (flightNo != null && flightNo.Contains("second flight"))
            {
                _date = WaitForElementIsVisible(By.Id(DATE));
                _date.SetValue(ControlType.DateTime, DateUtils.Now.AddDays(10));
                _date.SendKeys(Keys.Tab);
                WaitForLoad();
            }

            if (flightNo != null && flightNo.Contains("For linked"))
            {
                _date = WaitForElementIsVisible(By.Id(DATE));
                _date.SetValue(ControlType.DateTime, DateUtils.Now.AddDays(1));
                _date.SendKeys(Keys.Tab);
                WaitForLoad();
            }

            if (flightNo != null && flightNo.Contains("to Split"))
            {
                _addLeg = WaitForElementIsVisible(By.XPath(ADD_LEG));
                _addLeg.Click();
                WaitForLoad();

                _airportLeg2 = WaitForElementIsVisible(By.Id(AIRPORT_LEG2));
                _airportLeg2.SetValue(ControlType.DropDownList, airportFrom);
                _airportLeg2.SendKeys(Keys.Enter);
                WaitForLoad();
            }

            if (lpCart != null)
            {
                _lpCartName = WaitForElementIsVisible(By.Id(LPCART));
                _lpCartName.SetValue(ControlType.DropDownList, lpCart);
                _lpCartName.SendKeys(Keys.Enter);
                WaitForLoad();
            }

            if (flightRemarks != null)
            {
                _flightRemarks = WaitForElementIsVisible(By.XPath(FLIGHTREMARKS));
                _flightRemarks.SetValue(ControlType.TextBox, flightRemarks);
                WaitForLoad();
            }


            _validateBtn = WaitForElementToBeClickable(By.XPath(VALIDATE));
            _validateBtn.Click();
            WaitPageLoading();
            WaitForLoad();
        }


        public void FillField_CreatNewFlightWithFlightType(string flightNo, string customer, string aircraft, string airportFrom, string airportTo, string flightType = null
            , string loadingPlanName = null, string etaHours = "00", string etdHours = "23", string lpCart = null, DateTime? date = null)
        {
            //hidden  : <button id="scroll-to-bottom" type="button" class="btn-icon-arrowBot"></button>
            Thread.Sleep(2000);

            if (date != null)
            {
                DateTime dateNotNull = (DateTime)date;
                var dateInput = WaitForElementIsVisible(By.Id("FlightDate"));
                dateInput.SetValue(ControlType.TextBox, dateNotNull.ToString("dd/MM/yyyy"));
                dateInput.SendKeys(Keys.Tab);
                WaitForLoad();
            }
            else
            {
                var dateInput = WaitForElementIsVisible(By.Id("FlightDate"));
                dateInput.SetValue(ControlType.TextBox, DateUtils.Now.ToString("dd/MM/yyyy"));
                dateInput.SendKeys(Keys.Tab);
                WaitForLoad();
            }

            if (flightNo != null)
            {
                _flightNo = WaitForElementIsVisible(By.Id(FLIGHT_NUMBER));
                _flightNo.SetValue(ControlType.TextBox, flightNo);
                WaitForLoad();
            }

            if (customer != null)
            {
                _divCustomer = WaitForElementIsVisible(By.XPath(DIV_CUSTOMER));
                _divCustomer.Click();
                _divCustomer.Click();

                _customerName = WaitForElementIsVisible(By.Id(CUSTOMER_NAME));
                _customerName.SetValue(ControlType.TextBox, customer);
                _customerName.SendKeys(Keys.Enter);
                WaitForLoad();
            }

            _airportFrom = WaitForElementIsVisible(By.Id(AIRPORT_FROM));
            _airportFromSelected = WaitForElementIsVisible(By.XPath(AIRPORT_FROM_SELECTED));
            if (!_airportFromSelected.Text.Contains(airportFrom))
            {
                _airportFrom.SetValue(ControlType.DropDownList, airportFrom);
            }

            _airportTo = WaitForElementIsVisible(By.Id(AIRPORT_TO));
            _airportToSelected = WaitForElementIsVisible(By.XPath(AIRPORT_TO_SELECTED));
            if (!_airportToSelected.Text.Contains(airportTo))
            {
                _airportTo.SetValue(ControlType.DropDownList, airportTo);
            }

            if (flightNo != null && flightNo.Contains("to Split"))
            {
                _addLeg = WaitForElementIsVisible(By.XPath(ADD_LEG));
                _addLeg.Click();
                WaitForLoad();

                _airportLeg2 = WaitForElementIsVisible(By.Id(AIRPORT_LEG2));
                _airportLeg2.SetValue(ControlType.DropDownList, airportFrom);
                _airportLeg2.SendKeys(Keys.Enter);
            }

            _aircraftNumber = WaitForElementIsVisible(By.Id(AIRCRAFT_NUMBER));
            _aircraftNumber.SetValue(ControlType.TextBox, aircraft);
            _aircraftNumber.SendKeys(Keys.Enter);
            WaitForLoad();

            _etaHours = WaitForElementIsVisible(By.XPath(ETA_HOURS));
            _etaHours.SendKeys(Keys.ArrowRight);
            _etaHours.SendKeys(Keys.ArrowRight);
            _etaHours.SendKeys(Keys.Backspace);
            _etaHours.SendKeys(Keys.Backspace);
            _etaHours.SendKeys(etaHours);
            WaitForLoad();

            _etdHours = WaitForElementIsVisible(By.XPath(ETD_HOURS));
            _etdHours.SendKeys(Keys.ArrowRight);
            _etdHours.SendKeys(Keys.ArrowRight);
            _etdHours.SendKeys(Keys.Backspace);
            _etdHours.SendKeys(Keys.Backspace);
            _etdHours.SendKeys(etdHours);
            WaitForLoad();

            if (!string.IsNullOrEmpty(loadingPlanName))
            {
                var checkboxLP = WaitForElementExists(By.XPath(String.Format(CHECK_BOX_LP, loadingPlanName)));

                var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
                javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", checkboxLP);
                Thread.Sleep(1000);

                checkboxLP.SetValue(ControlType.CheckBox, true);
                WaitForLoad();
            }

            if (flightNo != null && flightNo.Contains("second flight"))
            {
                _date = WaitForElementIsVisible(By.Id(DATE));
                _date.SetValue(ControlType.DateTime, DateUtils.Now.AddDays(10));
                _date.SendKeys(Keys.Tab);
                WaitForLoad();
            }

            if (flightNo != null && flightNo.Contains("For linked"))
            {
                _date = WaitForElementIsVisible(By.Id(DATE));
                _date.SetValue(ControlType.DateTime, DateUtils.Now.AddDays(1));
                _date.SendKeys(Keys.Tab);
                WaitForLoad();
            }

            if (lpCart != null)
            {
                _lpCartName = WaitForElementIsVisible(By.Id(LPCART));
                _lpCartName.SetValue(ControlType.DropDownList, lpCart);
                _lpCartName.Click();
                WaitForLoad();
            }
            if (flightType != null)
            {
                var flightTypeInput = WaitForElementIsVisible(By.XPath(FLIGHTTYPE));
                if (!flightTypeInput.Text.Contains(flightType))
                {
                    flightTypeInput.SetValue(ControlType.DropDownList, flightType);
                }
            }
            _validateBtn = WaitForElementToBeClickable(By.XPath(VALIDATE));
            _validateBtn.Click();
            WaitPageLoading();
        }
        public bool FlightIsValid()
        {
            _validateLabel = WaitForElementIsVisible(By.XPath(VALIDATE_LABEL));

            if (_validateLabel.Text == "The Flight No field is required.")
            {
                return false;
            }

            return true;
        }
        public void CheckLoadingPlan(string loadingPlanName)
        {
            _webDriver.Manage().Window.Maximize();


            var table = _webDriver.FindElements(By.XPath("//*[@id=\"createFormVipOrder\"]/div[8]/div[1]/table/tbody/tr[*]/td[2]/input[contains(@name,'.Name')]"));
            var lp = table.FirstOrDefault(element => element.GetAttribute("value") == loadingPlanName);
            if (lp != null)
            {
                // Get the parent <tr> of the matched input
                var parentRow = lp.FindElement(By.XPath("./ancestor::tr"));

                // Now, find the first <td> in the same row
                var firstTd = parentRow.FindElement(By.XPath("./td[1]/div/input"));

                // Click on the first <td>
                firstTd.Click();
                WaitLoading();
            }
            WaitPageLoading();
        }
        public void SaveFlightAndReloadLP()
        {
            _validateBtn = WaitForElementIsVisible(By.XPath(VALIDATE));
            _validateBtn.Click();
            WaitLoading();
            WaitPageLoading();
        }
        public void FillField_CreatNewFlightWithStartDate(string flightNo, string customer, string aircraft, string airportFrom, string airportTo,
         string loadingPlanName = null, string etaHours = "00", string etdHours = "23", string lpCart = null, Nullable<DateTime> date = null)
        {
            if (date != null)
            {
                DateTime dateNotNull = (DateTime)date;
                var dateInput = WaitForElementIsVisibleNew(By.Id(FLIGHT_DATE));
                dateInput.SetValue(ControlType.TextBox, dateNotNull.ToString("dd/MM/yyyy"));
                dateInput.SendKeys(Keys.Tab);
            }
            else
            {
                var dateInput = WaitForElementIsVisibleNew(By.Id(FLIGHT_DATE));
                dateInput.SetValue(ControlType.TextBox, DateUtils.Now.ToString("dd/MM/yyyy"));
                dateInput.SendKeys(Keys.Tab);
                WaitLoading();
            }

            if (flightNo != null)
            {
                _flightNo = WaitForElementIsVisibleNew(By.Id(FLIGHT_NUMBER));
                _flightNo.SetValue(ControlType.TextBox, flightNo);
            }

            _airportFrom = WaitForElementIsVisibleNew(By.Id(AIRPORT_FROM));
            _airportFromSelected = WaitForElementIsVisibleNew(By.XPath(AIRPORT_FROM_SELECTED));
            if (!_airportFromSelected.Text.Contains(airportFrom))
            {
                _airportFrom.SetValue(ControlType.DropDownList, airportFrom);
            }

            _airportTo = WaitForElementIsVisibleNew(By.Id(AIRPORT_TO));
            _airportToSelected = WaitForElementIsVisibleNew(By.XPath(AIRPORT_TO_SELECTED));
            if (!_airportToSelected.Text.Contains(airportTo))
            {
                _airportTo.SetValue(ControlType.DropDownList, airportTo);
            }

            if (customer != null)
            {
                _divCustomer = WaitForElementIsVisibleNew(By.XPath(DIV_CUSTOMER));
                _divCustomer.Click();
                _divCustomer.Click();

                _customerName = WaitForElementIsVisibleNew(By.Id(CUSTOMER_NAME));
                _customerName.SetValue(ControlType.TextBox, customer);
                _customerName.SendKeys(Keys.Enter);
                WaitLoading();
            }

            _aircraftNumber = WaitForElementIsVisibleNew(By.Id(AIRCRAFT_NUMBER));
            _aircraftNumber.SetValue(ControlType.TextBox, aircraft);
            _aircraftNumber.SendKeys(Keys.Tab);

            _etaHours = WaitForElementIsVisibleNew(By.XPath(ETA_HOURS));
            _etaHours.SendKeys(Keys.ArrowRight);
            _etaHours.SendKeys(Keys.ArrowRight);
            _etaHours.SendKeys(Keys.Backspace);
            _etaHours.SendKeys(Keys.Backspace);
            _etaHours.SendKeys(etaHours);

            _etdHours = WaitForElementIsVisibleNew(By.XPath(ETD_HOURS));
            _etdHours.SendKeys(Keys.ArrowRight);
            _etdHours.SendKeys(Keys.ArrowRight);
            _etdHours.SendKeys(Keys.Backspace);
            _etdHours.SendKeys(Keys.Backspace);
            _etdHours.SendKeys(etdHours);

            if (!string.IsNullOrEmpty(loadingPlanName))
            {
                var checkboxLP = WaitForElementExists(By.XPath(String.Format(CHECK_BOX_LP, loadingPlanName)));

                var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
                javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", checkboxLP);

                checkboxLP.SetValue(ControlType.CheckBox, true);
            }

            if (flightNo != null && flightNo.Contains("second flight"))
            {
                _date = WaitForElementIsVisibleNew(By.Id(DATE));
                _date.SetValue(ControlType.DateTime, DateUtils.Now.AddDays(10));
                _date.SendKeys(Keys.Tab);
                WaitForLoad();
            }

            if (flightNo != null && flightNo.Contains("For linked"))
            {
                _date = WaitForElementIsVisibleNew(By.Id(DATE));
                _date.SetValue(ControlType.DateTime, DateUtils.Now.AddDays(1));
                _date.SendKeys(Keys.Tab);
                WaitForLoad();
            }

            if (flightNo != null && flightNo.Contains("to Split"))
            {
                _addLeg = WaitForElementIsVisibleNew(By.XPath(ADD_LEG));
                _addLeg.Click();
                WaitForLoad();

                _airportLeg2 = WaitForElementIsVisibleNew(By.Id(AIRPORT_LEG2));
                _airportLeg2.SetValue(ControlType.DropDownList, airportFrom);
                _airportLeg2.SendKeys(Keys.Enter);
                WaitForLoad();
            }

            if (lpCart != null)
            {
                _lpCartName = WaitForElementIsVisibleNew(By.Id(LPCART));
                _lpCartName.SetValue(ControlType.DropDownList, lpCart);
                _lpCartName.SendKeys(Keys.Enter);
                WaitForLoad();
            }

        }
        public bool IsExistLoadingPlan(string loadingPlanName)
        {
            _webDriver.Manage().Window.Maximize();

            var table = _webDriver.FindElements(By.XPath("//*[@id=\"createFormVipOrder\"]/div[8]/div[1]/table/tbody/tr[*]/td[2]/input[contains(@name,'.Name')]"));
            var lp = table.FirstOrDefault(element => element.GetAttribute("value") == loadingPlanName);
            if (lp != null)
            {
                return true;
            }
            else return false;

        }
    }
}