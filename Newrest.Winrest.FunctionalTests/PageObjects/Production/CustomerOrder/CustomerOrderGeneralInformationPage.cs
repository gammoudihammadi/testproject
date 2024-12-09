using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Vml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Production.CustomerOrder
{
    public class CustomerOrderGeneralInformationPage : PageBase
    {
        public CustomerOrderGeneralInformationPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //___________________________________ Constantes ____________________________________________

        // General
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";

        // Onglets
        private const string DETAILS_TAB = "hrefTabContentItems";

        // Informations
        private const string SITE = "//*[@id=\"dropdown-site\"]//child::option[@selected]";
        private const string ORDER_NUMBER = "//*[@id=\"createFormOrder\"]/div[1]/div[1]/div[2]/div";
        private const string DELIVERY_DATE = "OrderDate";
        private const string DELIVERY_INPUT = "DeliveryName";
        private const string LOCAL_FOREIGN = "//div[contains(@class,'bootstrap-switch')]";
        private const string LOCAL_FOREIGN_VALUE = "IsOrderLocalName";
        private const string COMMENT = "OrderComment";
        private const string STATUS = "ddlStatus";
        private const string PRICE = "//*[@id=\"createFormOrder\"]/div[3]/div/div[1]/div[2]";
        private const string CUSTOMER_NAME = "customer-input-id";
        private const string MESSAGES_TAB = "//*[@id=\"hrefTabContentMessages\"]";
        private const string PRINT = "//*[@id=\"div-body\"]/div/div[1]/div/div[2]/div/a[1]";
        private const string EXPEDITION_DATE = "ExpeditionDate";
        private const string FLIGHT_NUMBER = "FlightNumber";
        private const string FLIGHT_DATE = "flightDate";
        private const string HOURS = "//*[@id=\"createFormOrder\"]/div[1]/div[1]/div[6]/div/input[2]";
        private const string MINUTES = "//*[@id=\"createFormOrder\"]/div[1]/div[1]/div[6]/div/input[3]";
        private const string AIRCRAFT = "Aircraftt";
        private const string AIRPORTDEST = "AirportDest";
        private const string ORDER_TYPE = "dropdown-order-type";
        private const string ETD = "etd";
        private const string ORIGIN = "AirportOrig";
        private const string REGISTRATION_TYPE = "CustomFlightRegistrationType";
        private const string LABEL_1 = "Label1";
        private const string LABEL_2 = "Label2";
        private const string LABEL_3 = "Label3";
        private const string LABEL_4 = "Label4";
        private const string LABEL_5 = "Label5";
        private const string HANDLER = "dropdown-handler-type";

        //___________________________________ Variables ______________________________________________

        // General
        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;

        // Onglets
        [FindsBy(How = How.Id, Using = DETAILS_TAB)]
        private IWebElement _detailsTab;

        // Informations
        [FindsBy(How = How.XPath, Using = SITE)]
        private IWebElement _site;

        [FindsBy(How = How.XPath, Using = ORDER_NUMBER)]
        private IWebElement _orderNumber;

        [FindsBy(How = How.Id, Using = DELIVERY_DATE)]
        private IWebElement _deliveryDate;

        [FindsBy(How = How.Id, Using = DELIVERY_INPUT)]
        private IWebElement _deliveryInput;

        [FindsBy(How = How.XPath, Using = LOCAL_FOREIGN)]
        private IWebElement _localForeign;

        [FindsBy(How = How.Id, Using = LOCAL_FOREIGN_VALUE)]
        private IWebElement _localForeignValue;

        [FindsBy(How = How.Id, Using = COMMENT)]
        private IWebElement _commentInput;

        [FindsBy(How = How.Id, Using = STATUS)]
        private IWebElement _status;

        [FindsBy(How = How.XPath, Using = PRICE)]
        private IWebElement _price;

        [FindsBy(How = How.XPath, Using = PRINT)]
        private IWebElement _print;

        [FindsBy(How = How.Id, Using = CUSTOMER_NAME)]
        private IWebElement _customerName;

        [FindsBy(How = How.Id, Using = EXPEDITION_DATE)]
        private IWebElement _expeditionDate;

        [FindsBy(How = How.Id, Using = FLIGHT_NUMBER)]
        private IWebElement _flightNumber;

        [FindsBy(How = How.Id, Using = FLIGHT_DATE)]
        private IWebElement _flightDate;

        [FindsBy(How = How.Id, Using = HOURS)]
        private IWebElement _hours;

        [FindsBy(How = How.Id, Using = MINUTES)]
        private IWebElement _minutes;

        [FindsBy(How = How.Id, Using = AIRCRAFT)]
        private IWebElement _aircraft;

        [FindsBy(How = How.Id, Using = AIRPORTDEST)]
        private IWebElement _airportDest;

        [FindsBy(How = How.Id, Using = ORDER_TYPE)]
        private IWebElement _orderType;

        [FindsBy(How = How.Id, Using = ETD)]
        private IWebElement _etd;

        [FindsBy(How = How.Id, Using = ORIGIN)]
        private IWebElement _origin;

        [FindsBy(How = How.Id, Using = REGISTRATION_TYPE)]
        private IWebElement _registrationType;

        [FindsBy(How = How.Id, Using = LABEL_1)]
        private IWebElement _label1;

        [FindsBy(How = How.Id, Using = LABEL_2)]
        private IWebElement _label2;

        [FindsBy(How = How.Id, Using = LABEL_3)]
        private IWebElement _label3;

        [FindsBy(How = How.Id, Using = LABEL_4)]
        private IWebElement _label4;

        [FindsBy(How = How.Id, Using = LABEL_5)]
        private IWebElement _label5;

        [FindsBy(How = How.Id, Using = HANDLER)]
        private IWebElement _handler;

        // _________________________________ Méthodes ________________________________________________

        // General
        public CustomerOrderPage BackToList()
        {
            WaitForLoad();
            _backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            _backToList.Click();
            WaitForLoad();

            return new CustomerOrderPage(_webDriver, _testContext);
        }

        public CustomerOrderIMessagesPage GoToMessagesTab()
        {
            var _messageTab = WaitForElementIsVisible(By.XPath(MESSAGES_TAB));
            _messageTab.Click();
            WaitForLoad();

            return new CustomerOrderIMessagesPage(_webDriver, _testContext);
        }

        // Onglets
        public CustomerOrderItem ClickOnDetailTab()
        {
            Thread.Sleep(1000);
            _detailsTab = WaitForElementToBeClickable(By.Id(DETAILS_TAB));
            //_detailsTab = WaitForElementIsVisible(By.Id(DETAILS_TAB));
            _detailsTab.Click();
            WaitForLoad();

            return new CustomerOrderItem(_webDriver, _testContext);
        }

        // Informations
        public void UpdateGeneralInfo(string delivery, DateTime date, string comment, bool validated = false)
        {
            if (!validated)
            {
                _deliveryInput = WaitForElementIsVisible(By.Id(DELIVERY_INPUT));
                _deliveryInput.SetValue(ControlType.TextBox, delivery);
            }

            _commentInput = WaitForElementIsVisible(By.Id(COMMENT));
            _commentInput.SetValue(ControlType.TextBox, comment);

            _deliveryDate = WaitForElementIsVisible(By.Id(DELIVERY_DATE));
            _deliveryDate.SetValue(ControlType.DateTime, date);
            _deliveryDate.SendKeys(Keys.PageUp);
            _deliveryDate.SendKeys(Keys.Tab);

            //Temps obligatoire pour mettre en base
            Thread.Sleep(2000);
        }

        public CustomerOrderItem UpdateLocalForeign()
        {
            _localForeign = WaitForElementIsVisible(By.XPath(LOCAL_FOREIGN));
            _localForeign.Click();
            WaitForLoad();

            return new CustomerOrderItem(_webDriver, _testContext);
        }

        public bool CanEditDelivery()
        {
            try
            {
                _deliveryInput = WaitForElementIsVisible(By.Id(DELIVERY_INPUT));
                return _deliveryInput.GetAttribute("readonly").Equals(null);
            }
            catch
            {
                return false;
            }
        }



        public bool IsSelectSiteEditable()
        {

            var selectElement = WaitForElementIsVisible(By.Id("dropdown-site"));

            bool isDisabled = selectElement.GetAttribute("disabled") != null;

            return !isDisabled;
        }

        public string GetComment()
        {
            _commentInput = WaitForElementIsVisible(By.Id(COMMENT));
            return _commentInput.Text;
        }
        public bool IsCommentExist(string comment)
        {
            if (isElementVisible(By.Id(COMMENT)))
            {
             _commentInput = WaitForElementIsVisible(By.Id(COMMENT));
                if (_commentInput.Text == comment)
                {
                    return true;
                }

                else return false;

            }
            else return false;
        }
        public void SetComment(string comment)
        {
            WaitForLoad();
            _commentInput = WaitForElementIsVisible(By.Id(COMMENT));
            _commentInput.SetValue(ControlType.TextBox, comment);
            WaitForLoad();
            //Temps obligatoire pour mettre en base
            Thread.Sleep(2000);

        }
        public string GetOrderDate(string dateFormat)
        {
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            return DateTime.Parse(_deliveryDate.GetAttribute("value"), ci).ToShortDateString();
        }

        public string GetDelivery()
        {
            _deliveryInput = WaitForElementIsVisible(By.Id(DELIVERY_INPUT));
            return _deliveryInput.GetAttribute("value");
        }

        public string GetLocalForeignValue()
        {
            _localForeignValue = WaitForElementExists(By.Id(LOCAL_FOREIGN_VALUE));
            return _localForeignValue.GetAttribute("value");
        }

        public void ChangeStatus(string status)
        {
            _status = WaitForElementIsVisible(By.Id(STATUS));
            _status.SetValue(ControlType.DropDownList, status);

            //Temps obligatoire pour mettre en base
            Thread.Sleep(1000);
        }

        public string GetOrderNumber()
        {
            _orderNumber = WaitForElementIsVisible(By.XPath(ORDER_NUMBER));
            return _orderNumber.Text;
        }

        public string GetSite()
        {
            _site = WaitForElementIsVisible(By.XPath(SITE));
            return _site.Text;
        }

        public string GetPrice()
        {
            _price = WaitForElementIsVisible(By.XPath(PRICE));
            return _price.Text;
        }
        public string GetFlightNumber()
        {
            var _flightNumber = WaitForElementIsVisible(By.Id(FLIGHT_NUMBER));
            return _flightNumber.GetAttribute("value");
        }
        public string GetCustomerName()
        {
            _customerName = WaitForElementIsVisible(By.Id(CUSTOMER_NAME));
            return _customerName.GetAttribute("value");
        }
        public PrintReportPage PrintCustomerOrder(bool printValue)
        {

            ShowExtendedMenu();

            _print = WaitForElementIsVisible(By.XPath(PRINT));
            _print.Click();
            WaitForLoad();

            var confirm = WaitForElementIsVisible(By.XPath("//*[@id=\"btn-print\"]"));
            confirm.Click();

            if (printValue)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-pdf']"));
                ClickPrintButton();
            }

            //Results are opened in a new tab, switch the driver to the newly created one
            WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new PrintReportPage(_webDriver, _testContext);
        }

        public string GetExpeditionDate()
        {
            var expeditionDate = WaitForElementIsVisible(By.XPath("//*[@id=\"ExpeditionDate\"]"));
            return expeditionDate.GetAttribute("value");
        }

        public void UpdateExpeditionDate(DateTime date)
        {
            _expeditionDate = WaitForElementIsVisible(By.Id(EXPEDITION_DATE));
            _expeditionDate.SetValue(ControlType.DateTime, date);
            _expeditionDate.SendKeys(Keys.PageUp);
            _expeditionDate.SendKeys(Keys.Tab);

            //Temps obligatoire pour mettre en base
            Thread.Sleep(2000);
        }

        public void UpdateFlightNumber(string updateNumber)
        {
            _flightNumber = WaitForElementIsVisible(By.Id(FLIGHT_NUMBER));
            _flightNumber.SetValue(ControlType.TextBox, updateNumber);

            //Temps obligatoire pour mettre en base
            Thread.Sleep(2000);
        }

        public void UpdateFlightDate(DateTime date)
        {
            _flightDate = WaitForElementIsVisible(By.Id(FLIGHT_DATE));
            _flightDate.SetValue(ControlType.DateTime, date);
            _flightDate.SendKeys(Keys.PageUp);
            _flightDate.SendKeys(Keys.Tab);

            //Temps obligatoire pour mettre en base
            Thread.Sleep(2000);
        }

        public string GetFlightDate()
        {
            var expeditionDate = WaitForElementIsVisible(By.Id(FLIGHT_DATE));
            return expeditionDate.GetAttribute("value");
        }

        public void UpdateHour(string hour)
        {
            string[] timeParts = hour.Split(':');
            string hours = timeParts[0];
            string minutes = timeParts[1];

            _hours = WaitForElementIsVisible(By.XPath(HOURS));
            _hours.SendKeys(Keys.ArrowRight);
            _hours.SendKeys(Keys.ArrowRight);
            _hours.SendKeys(Keys.Backspace);
            _hours.SendKeys(Keys.Backspace);
            _hours.SendKeys(hours);
            WaitForLoad();

            _minutes = WaitForElementIsVisible(By.XPath(MINUTES));
            _minutes.SendKeys(Keys.ArrowRight);
            _minutes.SendKeys(Keys.ArrowRight);
            _minutes.SendKeys(Keys.Backspace);
            _minutes.SendKeys(Keys.Backspace);
            _minutes.SendKeys(minutes);
            WaitForLoad();

            //Temps obligatoire pour mettre en base
            Thread.Sleep(2000);
        }

        public string GetHour()
        {
            _hours = WaitForElementIsVisible(By.XPath(HOURS));
            string hoursValue = _hours.GetAttribute("value");
            _minutes = WaitForElementIsVisible(By.XPath(MINUTES));
            string minutesValue = _minutes.GetAttribute("value");
            string time = $"{hoursValue}:{minutesValue}";
            return time;
        }

        public string GetStatus()
        {
            var _status = WaitForElementIsVisible(By.Id(STATUS));
            string statusValue = _status.GetAttribute("value");
            return statusValue;
        }

        public void UpdateAircraft(string aircraft)
        {
            _aircraft = WaitForElementExists(By.Id(AIRCRAFT));
            _aircraft.Click();
            _aircraft.SetValue(ControlType.DropDownList, aircraft);
            //Temps obligatoire pour mettre en base
            Thread.Sleep(2000);
        }

        public bool GetAirCraft(string aircraft)
        {
            var aircrafts = WaitForElementExists(By.Id(AIRCRAFT));
            var liste = new SelectElement(aircrafts).AllSelectedOptions;
            foreach (var element in liste)
            {
                if (element.GetAttribute("innerHTML") == aircraft)
                {
                    return true;
                }
            }
            return false;

        }

        public void UpdateAirportDest(string airportDest)
        {
            _airportDest = WaitForElementExists(By.Id(AIRPORTDEST));
            _airportDest.Click();
            _airportDest.SetValue(ControlType.DropDownList, airportDest);
            //Temps obligatoire pour mettre en base
            Thread.Sleep(2000);
        }

        public bool GetAirportDest(string airportDest)
        {
            _airportDest = WaitForElementExists(By.Id(AIRPORTDEST));
            var liste = new SelectElement(_airportDest).AllSelectedOptions;
            foreach (var element in liste)
            {
                if (element.GetAttribute("innerHTML") == airportDest)
                {
                    return true;
                }
            }
            return false;

        }

        public void UpdateEtd(string etd)
        {
            _etd = WaitForElementIsVisible(By.Id(ETD));
            _etd.Click();
            _etd.SetValue(ControlType.TextBox,etd);

            //Temps obligatoire pour mettre en base
            Thread.Sleep(2000);
        }

        public string GetETDValue()
        {
            _etd = WaitForElementIsVisible(By.Id(ETD));
            return _etd.GetAttribute("value");
        }

        public void UpdateOrderType(string orderType) 
        {
            _orderType = WaitForElementExists(By.Id(ORDER_TYPE));
            _orderType.SetValue(ControlType.DropDownList, orderType);
            //Temps obligatoire pour mettre en base
            Thread.Sleep(2000);
        }

        public bool GetOrderType(string orderType)
        {
            _orderType = WaitForElementExists(By.Id(ORDER_TYPE));
            var liste = new SelectElement(_orderType).AllSelectedOptions;
            foreach (var element in liste)
            {
                if (element.GetAttribute("innerHTML") == orderType)
                {
                    return true;
                }
            }
            return false;

        }

        public void UpdateOrigin(string origin)
        {
            _origin = WaitForElementExists(By.Id(ORIGIN));
            _origin.SetValue(ControlType.DropDownList, origin);
            //Temps obligatoire pour mettre en base
            Thread.Sleep(2000);
        }

        public bool GetOrigin(string origin)
        {
            var _origin = WaitForElementExists(By.Id(ORIGIN));
            var liste = new SelectElement(_origin).AllSelectedOptions;
            foreach (var element in liste)
            {
                if (element.GetAttribute("innerHTML") == origin)
                {
                    return true;
                }
            }
            return false;
        }

        public void UpdateRegistrationType(string registrationType)
        {
            _registrationType = WaitForElementIsVisible(By.Id(REGISTRATION_TYPE));
            _registrationType.Click();
            _registrationType.SetValue(ControlType.TextBox, registrationType);

            //Temps obligatoire pour mettre en base
            Thread.Sleep(2000);
        }

        public string GetRegistrationType()
        {
            _registrationType = WaitForElementIsVisible(By.Id(REGISTRATION_TYPE));
            return _registrationType.GetAttribute("value");
        }

        public string GetLabel(string label)
        {
            switch (label)
            {
                case "label1":
                    _label1 = WaitForElementIsVisible(By.Id(LABEL_1));
                    return _label1.GetAttribute("value");

                case "label2":
                    _label2 = WaitForElementIsVisible(By.Id(LABEL_2));
                    return _label2.GetAttribute("value");

                case "label3":
                    _label3 = WaitForElementIsVisible(By.Id(LABEL_3));
                    return _label3.GetAttribute("value");

                case "label4":
                    _label4 = WaitForElementIsVisible(By.Id(LABEL_4));
                    return _label4.GetAttribute("value");

                case "label5":
                    _label5 = WaitForElementIsVisible(By.Id(LABEL_5));
                    return _label5.GetAttribute("value");

                default:
                    throw new ArgumentException("Invalid label name");
            }
        }

        public void SetLabel(string label, string labelText)
        {
            IWebElement element = null;

            switch (label)
            {
                case "label1":
                    element = WaitForElementIsVisible(By.Id(LABEL_1));
                    element.SendKeys(labelText);
                    //Temps obligatoire pour mettre en base
                    Thread.Sleep(2000);
                    break;

                case "label2":
                    element = WaitForElementIsVisible(By.Id(LABEL_2));
                    element.SendKeys(labelText);
                    //Temps obligatoire pour mettre en base
                    Thread.Sleep(2000);
                    break;

                case "label3":
                    element = WaitForElementIsVisible(By.Id(LABEL_3));
                    element.SendKeys(labelText);
                    //Temps obligatoire pour mettre en base
                    Thread.Sleep(2000);
                    break;

                case "label4":
                    element = WaitForElementIsVisible(By.Id(LABEL_4));
                    element.SendKeys(labelText);
                    //Temps obligatoire pour mettre en base
                    Thread.Sleep(2000);
                    break;

                case "label5":
                    element = WaitForElementIsVisible(By.Id(LABEL_5));
                    element.SendKeys(labelText);
                    //Temps obligatoire pour mettre en base
                    Thread.Sleep(2000);
                    break;

                default:
                    throw new ArgumentException("Invalid label name");
            }
        }

        public bool GetHandler(string handler)
        {
            _handler = WaitForElementExists(By.Id(HANDLER));
            var liste = new SelectElement(_handler).AllSelectedOptions;
            foreach (var element in liste)
            {
                if (element.GetAttribute("innerHTML") == handler)
                {
                    return true;
                }
            }
            return false;
        }

        public void UpdateHandler(string handler)
        {
            _handler = WaitForElementExists(By.Id(HANDLER));
            _handler.SetValue(ControlType.DropDownList, handler);

            //Temps obligatoire pour mettre en base
            Thread.Sleep(2000);
        }

    }
}
