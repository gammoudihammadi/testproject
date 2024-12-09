using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.PageObjects.TabletApp;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Web.UI.WebControls.WebParts;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Flights
{
    public class ParametresFlights : PageBase
    {
        private const string DELIVERY_REPORT = "//*[@id=\"paramFlightTab\"]/li[9]/a";
        private const string PRINT_SAFE_NOTES = "//*[@id=\"form-delivery-report-settings\"]/div[2]/div/div[2]/div/div/span[2]";
        private const string FLIGHT_ALERTS = "//*[@id=\"paramFlightTab\"]/li[11]/a";
        private const string ALERT_TYPES_TAB = "//*[@id=\"myTab\"]/li[*]/a[contains(text(),'{0}')]";
        private const string New_Btn = "//*[starts-with(@id,'{0}')]/table/thead/tr/th[6]/a";
        private const string CUSTOMER = "select-input-customer";
        private const string SITE = "select-input-site";
        private const string ACTIVATED = "checkbox-input-activated";
        private const string WEBAPP_ENABLED = "checkbox-input-winrest-webapp-enabled";
        private const string TABLET_ENABLED = "checkbox-input-winrest-tablet-enabled";
        private const string SAVE = "btn-submit-subscription";
        private const string DELETE = "//*[starts-with(@id,'{0}')]/table/tbody/tr[starts-with(@id,\"subscription\")]/td[contains(text(),'{1}')]/../td[contains(text(),'{2}')]/../td[6]/a/span[contains(@class,\"fa-trash\")]";
        private const string AIRCRAFT = "//*[@id=\"paramFlightTab\"]/li[2]/a";
        private const string HANDLER = "//*[@id=\"paramFlightTab\"]/li[3]/a";

        public ParametresFlights(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }
        public void GoToDeliveryReport()
        {
            var deliveryReport = WaitForElementIsVisible(By.XPath(DELIVERY_REPORT));
            deliveryReport.Click();
        }

        public void ClickPrintSafetyNote()
        {
            var switchYesNo = WaitForElementIsVisible(By.XPath("//*[@id='form-delivery-report-settings']/div[2]/div/div[2]/div"));
            if (!switchYesNo.GetAttribute("class").Contains("bootstrap-switch-on")) {
                var printSafety = WaitForElementExists(By.XPath(PRINT_SAFE_NOTES));
                printSafety.Click();
                Thread.Sleep(2000);
                WaitForLoad();
            }
        }
        public void UnClickPrintSafetyNote()
        {
            var switchYesNo = WaitForElementIsVisible(By.XPath("//*[@id='form-delivery-report-settings']/div[2]/div/div[2]/div"));
            if (switchYesNo.GetAttribute("class").Contains("bootstrap-switch-on"))
            {
                var printSafety = WaitForElementExists(By.XPath(PRINT_SAFE_NOTES));
                printSafety.Click();
                Thread.Sleep(2000);
                WaitForLoad();
            }
        }
        public void GoToFlightAlerts()
        {
            var flightAlerts = WaitForElementIsVisible(By.XPath(FLIGHT_ALERTS));
            flightAlerts.Click();
            WaitForLoad();
        }
        public void SelectAlertTypes(string alertType)
        {
            var tabSelected= WaitForElementIsVisible(By.XPath(string.Format(ALERT_TYPES_TAB,alertType)));
            tabSelected.Click();
            WaitForLoad();

        }
        public void AddCustomerSubscriptionToAlertTypes(string modalToOpen, string customer, string site, bool isActivated = true, bool isWebAppEnabled= true, bool isTabletEnabled=true)
        {        
            var modalToAddNew = WaitForElementIsVisible(By.XPath(string.Format(New_Btn, modalToOpen)));
            modalToAddNew.Click();
            WaitForLoad();

            var _customer = WaitForElementIsVisible(By.Id(CUSTOMER));
            _customer.SetValue(ControlType.DropDownList, customer);
            _customer.SendKeys(Keys.Enter);
            WaitForLoad();

            var _site = WaitForElementIsVisible(By.Id(SITE));
            _site.SetValue(ControlType.DropDownList, site);
            _site.SendKeys(Keys.Enter);
            WaitForLoad();

            var _isActivated = WaitForElementIsVisible(By.Id(ACTIVATED));
            _isActivated.SetValue(ControlType.CheckBox, isActivated);
            WaitForLoad();

            var _isWebAppEnabled = WaitForElementIsVisible(By.Id(WEBAPP_ENABLED));
            _isWebAppEnabled.SetValue(ControlType.CheckBox, isWebAppEnabled);
            WaitForLoad();

            var _isTabletEnabled = WaitForElementIsVisible(By.Id(TABLET_ENABLED));
            _isTabletEnabled.SetValue(ControlType.CheckBox, isTabletEnabled);
            WaitForLoad();

            var save = WaitForElementIsVisible(By.Id(SAVE));
            save.Click();
            WaitForLoad();
        }
        public void DeleteCustomerSubscription(string alertTypesId, string customer, string site)
        {
            var deleteRow = WaitForElementExists(By.XPath(string.Format(DELETE, alertTypesId, customer, site)));
            deleteRow.Click();
            WaitForLoad();
        }

        public ParametersFlightAircraft GoToAircraft()
        {
            var TimeBlock = WaitForElementIsVisible(By.XPath(AIRCRAFT));
            TimeBlock.Click();
            WaitForLoad();

            return new ParametersFlightAircraft(_webDriver, _testContext);
        }

        public ParametersFlightHandler GoToFlightsHandler()
        {
            var _handler = WaitForElementIsVisible(By.XPath(HANDLER));
            _handler.Click();
            WaitForLoad();

            return new ParametersFlightHandler(_webDriver, _testContext);
        }
        public List<string> GetListCustomersOfSubscriptions()
        {
            var customersNameList = new List<string>();
            var customersNameElements = _webDriver.FindElements(By.XPath("/html/body/div[2]/div/div/div/div/div/div[2]/div/div[5]/table/tbody/tr[*]/td[1]"));
            foreach (var customerNameElement in customersNameElements)
            {
                customersNameList.Add(customerNameElement.Text);
            }
            return customersNameList;
        }
        public bool IsCustomerInList(string customerName)
        {
            // Récupérer la liste des clients
            var customersList = GetListCustomersOfSubscriptions();
            if (customersList.Contains(customerName))
            {
                return true;
            }
            return false;
        }

    

        public void GoToFlightsRegistrationTypes()
        {
            var RegistrationTypes = WaitForElementIsVisible(By.XPath("//*[@id=\"paramFlightTab\"]/li[4]/a"));
            RegistrationTypes.Click();
            WaitForLoad();

        }


        public void ClickOnEdit()
        {

            var edit = WaitForElementIsVisible(By.XPath("/html/body/div[2]/div/div/div/div[3]/form/div/table/tbody/tr[2]/td[5]/a[1]/span"));
            edit.Click();
        }
        public void EditRegistrationTypes(string customer, string aircraft)
        {

            WaitPageLoading();
            if (customer != null)
            {
                var _divCustomer = WaitForElementIsVisible(By.XPath("//*[@id=\"RegistrationTypeModal\"]/div/div/div/div/form/div[2]/div[2]/div"));
                _divCustomer.Click();
                _divCustomer.Click();

                var _customerName = WaitForElementIsVisible(By.Id("CustomerId"));
                _customerName.SetValue(ControlType.TextBox, customer);
                WaitPageLoading();

                _customerName.SendKeys(Keys.Enter);
                WaitPageLoading();
            }     
            WaitPageLoading();
            if (aircraft != null)
            {
                var _divaircraft = WaitForElementIsVisible(By.XPath("//*[@id=\"RegistrationTypeModal\"]/div/div/div/div/form/div[2]/div[3]/div"));
                _divaircraft.Click();
                _divaircraft.Click();

                var _aircraftrName = WaitForElementIsVisible(By.Id("AircraftId"));
                _aircraftrName.SetValue(ControlType.TextBox, aircraft);
                _aircraftrName.SendKeys(Keys.Enter);
                WaitPageLoading();
            }
         


        }

        public void EditRegistrationTypeswithCustomer(string customer)
        {

            WaitPageLoading();
                if (customer != null)
            {
                var _divCustomer = WaitForElementIsVisible(By.XPath("//*[@id=\"RegistrationTypeModal\"]/div/div/div/div/form/div[2]/div[2]/div"));
                WaitPageLoading();

                _divCustomer.Click();
                _divCustomer.Click();
                WaitPageLoading();

                var _customerName = WaitForElementIsVisible(By.Id("CustomerId"));
                _customerName.SetValue(ControlType.TextBox, customer);
                _customerName.SendKeys(Keys.Enter);
                WaitPageLoading();


            }
        }


        public string GetCustomer()
        {

            var element = WaitForElementIsVisible(By.Id("CustomerId"));
           return element.Text;
        }   
        public void  save()
        {

            var element = WaitForElementIsVisible(By.Id("last"));
            element.Click();
        }

        public string GetCustomerFromTable()
        {

            var element = WaitForElementIsVisible(By.XPath("/html/body/div[2]/div/div/div/div[3]/form/div/table/tbody/tr[2]/td[4]"));
            return element.Text;
        }
        public string GetAirCraftFromTable()
        {

            var element = WaitForElementIsVisible(By.XPath("/html/body/div[2]/div/div/div/div[3]/form/div/table/tbody/tr[2]/td[2]"));
            return element.Text;
        }

    }

}
