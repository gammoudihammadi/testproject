using DocumentFormat.OpenXml.Bibliography;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Linq;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.Delivery
{
    public class DeliveryGeneralInformationPage : PageBase
    {
        public DeliveryGeneralInformationPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //_______________________________________________________Constantes____________________________________________________________

        // General
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";

        // Informations
        private const string NAME = "first";
        private const string CUSTOMER = "CustomerId";
        private const string SITE = "SiteId";
        private const string DELIVERY_LABEL_COMMENT = "LabelComment";
        private const string DELIVERY_NOTE_COMMENT = "FlightDeliveryNoteComment";
        private const string ACTIVE = "check-box-isactive";
        private const string ACTIVE2 = "/html/body/div[3]/div/div/div[2]/div/div/div/form/div/div[2]/div[1]/div/div[8]/div/div/input";
        private const string DELETE = "//*[@id=\"div-delete-delivery\"]/a";
        private const string CONFIRM_DELETE = "dataConfirmOK";

        private const string ERROR_MSG = "//*[@id=\"flightdelivery-filter-form\"]/div/div/div[1]/div/div[1]/div/span";
        private const string ERROR = "//*[@id=\"dataAlertModal\"]/div/div/div[2]/p";




        private const string HOURS = "/html/body/div[3]/div/div/div[2]/div/div/div/form/div/div[2]/div[1]/div/div[1]/div/input[2]";
        private const string MINUTES = "/html/body/div[3]/div/div/div[2]/div/div/div/form/div/div[2]/div[1]/div/div[1]/div/input[3]";
        private const string STOP_ACCESS_TO_CUSTOMER_PORTAL = "//*[@id=\"blockMode\"]";
        private const string STARTING_TIME_TO_LIMITED_ACCESS_HOURS = "//*[@id=\"HoursBeforeAccess\"]";
        private const string ENDING_TIME_TO_LIMITED_ACCESS_HOURS = "//*[@id=\"HoursBeforeModifications\"]";
        private const string ALLOWED_PRECENTAGE_IN_LIMIT_ACCESS = "//*[@id=\"AllowedModificationsPercent\"]";
        private const string DAY_TO_CLICK = "/html/body/div[3]/div/div/div[2]/div/div/div/form/div/div[2]/div[1]/div/div[6]/ul/li[*]/a[text()='{0}']";
        private const string DAY_STARTING_TIME_LIMITED_ACCESS = "/html/body/div[3]/div/div/div[2]/div/div/div/form/div/div[2]/div[1]/div/div[6]/div/div[2]/div[1]/div/input";
        private const string DAY_ENDING_TIME_LIMITED_ACCESS = "/html/body/div[3]/div/div/div[2]/div/div/div/form/div/div[2]/div[1]/div/div[6]/div/div[2]/div[2]/div/input";
        //_______________________________________________________Variables____________________________________________________________

        // General
        [FindsBy(How = How.XPath, Using = HOURS)]
        private IWebElement _hours;

        [FindsBy(How = How.XPath, Using = MINUTES)]
        private IWebElement _minutes;

        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;

        // Informations
        [FindsBy(How = How.XPath, Using = NAME)]
        private IWebElement _deliveryName;

        [FindsBy(How = How.Id, Using = CUSTOMER)]
        private IWebElement _customer;

        [FindsBy(How = How.Id, Using = SITE)]
        private IWebElement _site;

        [FindsBy(How = How.Id, Using = DELIVERY_LABEL_COMMENT)]
        private IWebElement _labelComment;

        [FindsBy(How = How.Id, Using = DELIVERY_NOTE_COMMENT)]
        private IWebElement _deliveryNoteComment;

        [FindsBy(How = How.Id, Using = ACTIVE)]
        private IWebElement _activate;

        [FindsBy(How = How.XPath, Using = ACTIVE2)]
        private IWebElement _activate2;

        [FindsBy(How = How.XPath, Using = DELETE)]
        private IWebElement _delete;

        [FindsBy(How = How.Id, Using = CONFIRM_DELETE)]
        private IWebElement _confirmDelete;

        [FindsBy(How = How.XPath, Using = ERROR)]
        private IWebElement _errorWarn;

        [FindsBy(How = How.XPath, Using = ERROR_MSG)]
        private IWebElement _error;


        //_______________________________________________________Pages______________________________________________________________________

        // General
        public DeliveryPage BackToList()
        {
            _backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            _backToList.Click();
            WaitForLoad();

            return new DeliveryPage(_webDriver, _testContext);
        }

        // Informations
        public void Update_General_Information(string value, string customer = null, string site = null)
        {
            // Définition du nom
            _deliveryName = WaitForElementIsVisible(By.Id(NAME));
            _deliveryName.SetValue(ControlType.TextBox, value);

            if (customer != null)
            {
                _customer = WaitForElementIsVisible(By.Id(CUSTOMER));
                _customer.SetValue(ControlType.DropDownList, customer);
            }

            if (site != null)
            {
                _site = WaitForElementIsVisible(By.Id(SITE));
                _site.SetValue(ControlType.DropDownList, site);
            }

            WaitPageLoading();
        }

        public string GetLabelComment()
        {
            _labelComment = WaitForElementIsVisible(By.Id(DELIVERY_LABEL_COMMENT));
            return _labelComment.Text;
        }

        public void SetLabelComment(string deliveryLabelComment)
        {
            _labelComment = WaitForElementIsVisible(By.Id(DELIVERY_LABEL_COMMENT));
            _labelComment.SetValue(ControlType.TextBox, deliveryLabelComment);
            WaitForLoad();
            WaitPageLoading(); //time to save value
        }

        public string GetDeliveryNotesComment()
        {
            _deliveryNoteComment = WaitForElementIsVisible(By.Id(DELIVERY_NOTE_COMMENT));
            return _deliveryNoteComment.Text;
        }

        public void SetDeliveryNotesComment(string deliveryLabelComment)
        {
            _deliveryNoteComment = WaitForElementIsVisible(By.Id(DELIVERY_NOTE_COMMENT));
            _deliveryNoteComment.SetValue(ControlType.TextBox, deliveryLabelComment);
            WaitForLoad();
            WaitPageLoading(); //time to save value
        }
        public void SetMethod(string method)
        {
            if (method != null)
            {
                _customer = WaitForElementIsVisible(By.Id("Method"));
                _customer.SetValue(ControlType.DropDownList, method);
            }

            WaitPageLoading();
        }

        public void SetActive(bool active)
        {
            if (isElementVisible(By.Id(ACTIVE)))
            {
                _activate = WaitForElementExists(By.Id(ACTIVE));
                _activate.SetValue(ControlType.CheckBox, active);
            }
            else
            {
                _activate2 = WaitForElementExists(By.XPath(ACTIVE2));
                _activate2.SetValue(ControlType.CheckBox, active);
            }
            WaitPageLoading();
        }

        public bool IsUpdated(string name, string customer, string site)
        {
            bool updatedName;
            bool updatedCustomer;
            bool updatedSite;

            _deliveryName = WaitForElementIsVisible(By.Id(NAME));
            updatedName = name.Equals(_deliveryName.GetAttribute("value"));

            var newCust = new SelectElement(_webDriver.FindElement(By.Id(CUSTOMER)));
            updatedCustomer = customer.Equals(newCust.AllSelectedOptions.FirstOrDefault()?.Text);

            var newSite = new SelectElement(_webDriver.FindElement(By.Id(SITE)));
            updatedSite = site.Equals(newSite.AllSelectedOptions.FirstOrDefault()?.Text);

            return updatedName && updatedCustomer && updatedSite;
        }

        public void DeleteDelivery()
        {
            Actions action = new Actions(_webDriver);
            WaitForLoad();
            WaitPageLoading();
            _delete = WaitForElementExists(By.XPath(DELETE));
            action.MoveToElement(_delete).Perform();
            _delete.Click();
            WaitForLoad();

            _confirmDelete = WaitForElementIsVisible(By.Id(CONFIRM_DELETE));
            _confirmDelete.Click();
            WaitPageLoading();
        }

        public bool GetDeleteErrorMessage(string errorText)
        {
            if (isElementVisible(By.XPath(ERROR)))
            {
                _errorWarn = WaitForElementExists(By.XPath(ERROR));
                return _errorWarn.Text.Equals(errorText);
            }
            else
            {
                return false;
            }
        }


        public bool IsError(string error)
        {
            if (isElementVisible(By.XPath(ERROR_MSG)))
            {
                _error = WaitForElementIsVisible(By.XPath(ERROR_MSG));
                return _error.Text.Equals(error);
            }
            else
            {
                return false;
            }
        }

        public void SetDeliveryTime(string hours , string minutes)
        {
            WaitForLoad();
            var hoursIn = WaitForElementIsVisible(By.XPath(HOURS));
            hoursIn.SendKeys(Keys.Control + "a");
            hoursIn.SendKeys(Keys.Backspace); 
            hoursIn.SendKeys(Keys.Backspace); 
            hoursIn.SetValue(ControlType.TextBox, hours);
            WaitPageLoading();
            WaitPageLoading();

            var minutesIn = WaitForElementIsVisible(By.XPath(MINUTES));
            minutesIn.Clear();
            minutesIn.SetValue(ControlType.TextBox, minutes);
            WaitPageLoading();
            WaitPageLoading();
        }
        public void SetStopAccessToCustomerPortal(string value)
        {
            var input = WaitForElementIsVisible(By.XPath(STOP_ACCESS_TO_CUSTOMER_PORTAL));
            input.SetValue(ControlType.DropDownList, value);
            WaitPageLoading();
            WaitForLoad();
        }
        public void SetStartingTimeToLimitedAccess(string value)
        {
            WaitForLoad();
            var input = WaitForElementIsVisible(By.XPath(STARTING_TIME_TO_LIMITED_ACCESS_HOURS));
            input.Clear();
            input.SendKeys(value);
            WaitPageLoading();
            WaitForLoad();
        }
        public void SetEndingTimeToLimitedAccess(string value)
        {
            var input = WaitForElementIsVisible(By.XPath(ENDING_TIME_TO_LIMITED_ACCESS_HOURS));
            input.Clear();
            input.SendKeys(value);
            WaitPageLoading();
            WaitForLoad();
        }
        public void SetAllowedInLimitedAccess(string value)
        {
            var input = WaitForElementIsVisible(By.XPath(ALLOWED_PRECENTAGE_IN_LIMIT_ACCESS));

            input.Clear();
            input.SendKeys(value);
            WaitPageLoading();
            WaitForLoad();
        }
        public void SetToDefault()
        {
            WaitPageLoading();
            SetDeliveryTime("00", "00");
            SetStopAccessToCustomerPortal("Daily");
            SetStartingTimeToLimitedAccess("24");
            SetEndingTimeToLimitedAccess("0");
            SetAllowedInLimitedAccess("0");
            WaitForLoad();
        }

        public void ClickDay(string value)
        {
            WaitPageLoading();
            WaitForLoad();
            ScrollUntilElementIsInView(By.XPath($@"//*[@id='tab{value}']"));
            WaitLoading();
            var day = WaitForElementExists(By.XPath($@"//*[@id='tab{value}']"));
            day.Click();
        }
        public void SetDayStartingTimeToLimitedAccess(string demainT,string value)
        {
            var pathString = "[id*="+ demainT +"][id*=StartModifications]";
            var startIn = WaitForElementIsVisible(By.CssSelector(pathString));
            startIn.Clear();
            startIn.SendKeys(value);
            WaitPageLoading();
            WaitForLoad();
        }
        public void SetDayEndingTimeToLimitedAccess(string demainT,string value)
        {
            var endIn = WaitForElementIsVisible(By.CssSelector("[id*='" + demainT + "'][id*=EndModifications]"));
            endIn.Clear();
            endIn.SendKeys(value);
            WaitPageLoading();
            WaitForLoad();
        }
    }
}
