using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
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
    public class FlightEditModalPage : PageBase
    {
        public FlightEditModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // _________________________________________ Constantes ________________________________________

        private const string MONDAY = "/html/body/div[3]/div/div/div/div/form/div[2]/div[6]/div/label[1]";
        private const string TUESDAY = "/html/body/div[3]/div/div/div/div/form/div[2]/div[6]/div/label[2]";
        private const string WEDNESDAY = "/html/body/div[3]/div/div/div/div/form/div[2]/div[6]/div/label[3]";
        private const string THURSDAY = "/html/body/div[3]/div/div/div/div/form/div[2]/div[6]/div/label[4]";
        private const string FRIDAY = "/html/body/div[3]/div/div/div/div/form/div[2]/div[6]/div/label[5]";
        private const string SATURDAY = "/html/body/div[3]/div/div/div/div/form/div[2]/div[6]/div/label[6]";
        private const string SUNDAY = "/html/body/div[3]/div/div/div/div/form/div[2]/div[6]/div/label[7]";

        //Aircraft
        private const string AIRCRAFT_NUMBER = "dropdown-aircraft-selectized";
        private const string AIRCRAFT_INPUT = "//*[@id=\"createFormVipOrder\"]/div[6]/div[1]/div/div";

        //loading plan
        private const string CHECK_LOADING_PLAN = "//*[@id=\"createFormVipOrder\"]/div[8]/div[1]/table/tbody/tr[*]/td[1]/div/input";
        private const string SAVEANDRELOAD = "//*[@id=\"form-createdit-flight\"]/div[2]/button[2]";
        private const string SAVEONLY = "//*[@id=\"form-createdit-flight\"]/div[2]/button[3]";
        private const string UNCHECKREMOVE = "EnableLoadingPlansRemoveThenAdd";
        private const string MODAL = "//*[@id=\"dataAlertModal\"]/div/div/div[2]/p";
        private const string MODAL_ID = "dataAlertModal";
        private const string FLIGHTREMARKS = "//*[@id=\"InternalFlightRemarks\"]";
        private const string CLOSEBUTTON = "viewDetails-close";
        private const string EXTENDED_SERVICE_BTN = "//*[@id=\"leg-details-with-services\"]/div/div/div/button/span";
        private const string ADD_SERVICE_BTN = "//*[@id=\"leg-details-with-services\"]/div/div/div/div/a[1]";
        private const string ADD_GENERIC_SERVICE_BTN = "//*[@id=\"leg-details-with-services\"]/div/div/div/div/a[2]";
        private const string EXTENDED_FLIGHT_BTN = "dropdownBtn";
        private const string DUPLICATE_FLIGHT_BTN = "duplicateFlights";
        private const string DUPLICATE_BTN = "/html/body/div[3]/div/div/div/div/form/div[3]/button[2]";
        private const string GENERIC_SERVICE = "//*[starts-with(@id,\"ServiceIdContainer_\")]/div/div[1]/div";
        private const string GENERIC_SERVICE_INPUT = "/html/body/div[5]/div/div/div[1]/table/tbody/tr/td[2]/div/div/table/tbody/tr/td[2]/div/div/div[2]/div/table/tbody/tr[3]/td[1]/div/div/div[1]/input"; // "//*[contains(@id,\"-selectized\")]";
        private const string EDIT_DETAILS_FLIGHT = "greenEditBtn";



        //Variables
        [FindsBy(How = How.XPath, Using = EXTENDED_SERVICE_BTN)]
        private IWebElement _extendedServiceButton;
        [FindsBy(How = How.XPath, Using = ADD_SERVICE_BTN)]
        private IWebElement _addservice;
        [FindsBy(How = How.XPath, Using = ADD_GENERIC_SERVICE_BTN)]
        private IWebElement _addGenericService;
        [FindsBy(How = How.Id, Using = EXTENDED_FLIGHT_BTN)]
        private IWebElement _extendedFlightButton;
        [FindsBy(How = How.Id, Using = DUPLICATE_FLIGHT_BTN)]
        private IWebElement _duplicateFlight;
        [FindsBy(How = How.XPath, Using = DUPLICATE_BTN)]
        private IWebElement _duplicateFlightButton;
        [FindsBy(How = How.XPath, Using = GENERIC_SERVICE)]
        private IWebElement _genericService;
        [FindsBy(How = How.XPath, Using = GENERIC_SERVICE_INPUT)]
        private IWebElement _genericServiceInput;
        [FindsBy(How = How.XPath, Using = EDIT_DETAILS_FLIGHT)]
        private IWebElement _editDetailsFlight;


        public void ChangeAircaftForLoadingPlan(string aircraft)
        {
            var _aircraftNumber = WaitForElementExists(By.XPath(AIRCRAFT_INPUT));
            _aircraftNumber.Click();
            var airctaftDrp = WaitForElementIsVisible(By.Id(AIRCRAFT_NUMBER));

            airctaftDrp.SendKeys(Keys.End);
            airctaftDrp.SendKeys(Keys.Control + "a");
            airctaftDrp.SendKeys(Keys.Backspace);
            airctaftDrp.SendKeys(aircraft.ToString());
            airctaftDrp.SendKeys(Keys.Enter);
            WaitPageLoading();
            WaitForLoad();
        }

        public void SelectAllLoadingPlan()
        {

            var elements = _webDriver.FindElements(By.XPath(CHECK_LOADING_PLAN));
            foreach (var elm in elements)
            {
                elm.SetValue(PageBase.ControlType.CheckBox, true);
                WaitForLoad();
            }
            var save = WaitForElementExists(By.XPath(SAVEANDRELOAD));
            save.Click();
            WaitPageLoading();
            WaitForLoad();
        }

        public void SetFlightRemarks(string remarks)
        {
            var elements = WaitForElementExists(By.XPath(FLIGHTREMARKS));
            elements.SetValue(ControlType.TextBox, remarks);

            var save = WaitForElementExists(By.XPath(SAVEANDRELOAD));
            save.Click();
            WaitForLoad();
            WaitPageLoading();
        }

        public string getFlightRemarks()
        {
            var elements = WaitForElementExists(By.XPath(FLIGHTREMARKS));
            var remarks = elements.Text;

            return remarks;
        }

        public void UnCheckRemoveAll()
        {

            var result = WaitForElementExists(By.Id(UNCHECKREMOVE));
            result.Click();
            WaitForLoad();

            var save = WaitForElementExists(By.XPath(SAVEANDRELOAD));
            save.Click();
            WaitForLoad();
        }

        public void SaveForFlightOnly()
        {
            var save = WaitForElementExists(By.XPath(SAVEONLY));
            save.Click();
            WaitPageLoading();
            WaitForLoad();
        }

        public bool IsLpCartsDetectedInSelectedLoadingPlan()
        {
            WaitForLoad();
            var modalVisible = isElementVisible(By.Id(MODAL_ID));
            Assert.IsTrue(modalVisible, "modal did not appear");
            var result = WaitForElementIsVisible(By.XPath(MODAL));
            if (result.Text.Contains("Several possible LPCarts have been detected in selected LoadingPlans"))
            {
                return true;
            }
            if (result.Text.Contains("1 trolley labels has been created from LoadingPlan"))
            {
                return true;
            }
            if (result.Text.Contains("has been automatically selected by default"))
            {
                return true;
            }
            Console.WriteLine("erreur [" + result.Text + "] à la place de [Several possible LPCarts have been detected in selected LoadingPlans]");

            var okButton = WaitForElementExists(By.Id("dataAlertCancel"));
            okButton.Click();
            WaitForLoad();

            return false;
        }
        public void ClickOnCloseButton()
        {
            WaitForLoad();
            var result = WaitForElementExists(By.Id(CLOSEBUTTON));
            result.Click();
            WaitForLoad();

        }
        public void ClickOnPreval()
        {
            var prevalButton = WaitForElementIsVisible(By.Id("prevalPopUp"));
            prevalButton.Click();
            WaitForLoad();
            Thread.Sleep(1500);
            WaitForLoad();
        }
        public void ClickOnValidate()
        {
            var validateButton = WaitForElementIsVisible(By.Id("validatePopup"));
            validateButton.Click();
            WaitForLoad(); Thread.Sleep(1500);
            WaitForLoad();
        }
        public void ClickOnInvoice()
        {
            var invoiceButton = WaitForElementIsVisible(By.Id("invoicePopup"));
            invoiceButton.Click();
            WaitForLoad(); Thread.Sleep(1500);
            WaitForLoad();
        }
        public void ClickOnAddGuestType()
        {
            var addGuestTypeResult = WaitForElementExists(By.XPath("/html/body/div[5]/div/div/div[1]/table/tbody/tr/td[2]/div/div/table/tbody/tr/td[1]/div/div/div[2]/div/a[2]"));
            addGuestTypeResult.Click();
            WaitForLoad();
            var createResult = WaitForElementExists(By.XPath("//*[@id=\"form-create-guest-type\"]/div[4]/button[2]"));
            createResult.Click();
            WaitForLoad();
            WaitPageLoading();

        }
        public void ShowServiceExtendedMenu()
        {
            _extendedServiceButton = WaitForElementExists(By.XPath(EXTENDED_SERVICE_BTN));
            var actions = new Actions(_webDriver);
            actions.MoveToElement(_extendedServiceButton).Perform();
            WaitForLoad();
        }
        public void ShowFlightExtendedMenu()
        {
            _extendedFlightButton = WaitForElementExists(By.Id(EXTENDED_FLIGHT_BTN));
            var actions = new Actions(_webDriver);
            actions.MoveToElement(_extendedFlightButton).Perform();
            WaitForLoad();
        }
        public void AddNewService()
        {
            ShowServiceExtendedMenu();
            _addservice = WaitForElementIsVisible(By.XPath(ADD_SERVICE_BTN));
            _addservice.Click();
            WaitPageLoading();
        }
        public void ClickAddGenericService()
        {
            ShowServiceExtendedMenu();
            _addGenericService = WaitForElementIsVisible(By.XPath(ADD_GENERIC_SERVICE_BTN));
            _addGenericService.Click();
            WaitForLoad();
            WaitPageLoading();
        }
        public void AddGenericService(string genericService)
        {
            _genericService = WaitForElementIsVisible(By.XPath(GENERIC_SERVICE));
            _genericService.Click();
            //_genericServiceInput = WaitForElementIsVisible(By.XPath(GENERIC_SERVICE_INPUT));
            //_genericServiceInput.SetValue(ControlType.DropDownList, genericService);
            WaitForElementIsVisible(By.XPath("//*[starts-with(@id,\"ServiceIdContainer_\")]/div/div[2]/div/div[2]")).Click();
            WaitForLoad();

        }
        public void DuplicateFlight(string flightNumber, DateTime flightDateFrom, DateTime? flightDateTo)
        {
            ShowFlightExtendedMenu();
            WaitForLoad();
            _duplicateFlight = WaitForElementIsVisible(By.Id(DUPLICATE_FLIGHT_BTN));
            _duplicateFlight.Click();
            WaitPageLoading();
            WaitForLoad();

            SetFlightAttributes(flightNumber, true, flightDateFrom, flightDateTo);
            WaitForLoad();
            _duplicateFlightButton = WaitForElementIsVisible(By.XPath(DUPLICATE_BTN));
            _duplicateFlightButton.Click();
        }
        public void SetFlightAttributes(string flightNumber, bool overPeriod, DateTime flightDateFrom, DateTime? flightDateTo)
        {
            WaitLoading();
            WaitForLoad();
            var _newFlightNumber = WaitForElementExists(By.Id("NewFlightNumber"));
            _newFlightNumber.SetValue(ControlType.TextBox, flightNumber);
            var _overPeriod = WaitForElementExists(By.Id("over-period"));
            _overPeriod.SetValue(ControlType.CheckBox, overPeriod);
            var _datapickerNewValueFrom = WaitForElementIsVisible(By.Id("datapicker-new-flightdate"));
            _datapickerNewValueFrom.SetValue(ControlType.DateTime, flightDateFrom);
            _datapickerNewValueFrom.SendKeys(Keys.Tab);
            if (overPeriod)
            {
                var _datapickerNewValueTo = WaitForElementIsVisible(By.Id("datapicker-new-flight-dateTo"));
                _datapickerNewValueTo.SetValue(ControlType.DateTime, flightDateTo);
                _datapickerNewValueTo.SendKeys(Keys.Tab);
                CheckDays();
            }
        }
        private void CheckDays()
        {
            var _mondayBtn = WaitForElementIsVisible(By.XPath(MONDAY));
            _mondayBtn.Click();

            var _tuesdayBtn = WaitForElementIsVisible(By.XPath(TUESDAY));
            _tuesdayBtn.Click();

            var _wednesdayBtn = WaitForElementIsVisible(By.XPath(WEDNESDAY));
            _wednesdayBtn.Click();

            var _thursdayBtn = WaitForElementIsVisible(By.XPath(THURSDAY));
            _thursdayBtn.Click();

            var _fridayBtn = WaitForElementIsVisible(By.XPath(FRIDAY));
            _fridayBtn.Click();

            var _saturdayBtn = WaitForElementIsVisible(By.XPath(SATURDAY));
            _saturdayBtn.Click();

            var _sundayBtn = WaitForElementIsVisible(By.XPath(SUNDAY));
            _sundayBtn.Click();
        }
        public bool GetServicesNames(string newServiceName)
        {
            var servicesNames = new List<string>();
            var rows = _webDriver.FindElements(By.XPath("/html/body/div[5]/div/div/div[1]/table/tbody/tr/td[2]/div/div/table/tbody/tr/td[2]/div/div/div[2]/div/table/tbody/tr[*]"));


            for (int i = 1; i < rows.Count; i++)
            {
                Console.WriteLine(rows[i].Text);
                string[] lines = rows[i].Text.Split(new[] { "\r\n" }, StringSplitOptions.None);
                servicesNames.Add(lines[0]);

            }
            var serviceExist = servicesNames.Any(c => c.Contains(newServiceName));
            return serviceExist;
        }
        public FlightCreateModalPage OpenEditDetailsflightModal()
        {
            _editDetailsFlight = WaitForElementIsVisible(By.Id(EDIT_DETAILS_FLIGHT));
            _editDetailsFlight.Click();
            WaitLoading();
            return new FlightCreateModalPage(_webDriver, _testContext);
        }

    }
}
