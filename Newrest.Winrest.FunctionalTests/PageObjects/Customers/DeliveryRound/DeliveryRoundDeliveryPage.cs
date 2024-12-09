using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Delivery;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.DeliveryRound
{
    public class DeliveryRoundDeliveryPage : PageBase
    {

        public DeliveryRoundDeliveryPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //__________________________________Constantes_______________________________________

        // General
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";

        // Onglets
        private const string GENERAL_INFORMATION = "hrefTabContentInformation";

        // Add delivery
        private const string ADD_DELIVERY = "FlightDeliveryName";
        private const string FIRST_DELIVERY_RESULT = "//*[@id=\"flightdelivery-list-result\"]/table/tbody/tr/td[2]";

        // Duplicate
        private const string BTN_DUPLICATE = "//*[@id=\"div-body\"]/div/div[1]/a";
        private const string NAME_DUPLICATE = "first";
        private const string STARTDATE_DUPLICATE = "DestinationStartDate";
        private const string ENDDATE_DUPLICATE = "DestinationEndDate";
        private const string CONFIRM_DUPLICATE = "last";

        // Tableau
        private const string FIRST_DELIVERY = "//*[@id=\"datasheet-list\"]/div/ul/li/div/div[1]/div[1]";

        private const string EDIT_DELIVERY = "//*[@id=\"datasheet-list\"]/div/ul/li/div/div[2]/div[4]/a[1]";
        private const string DELETE_BTN = "//*[@id=\"datasheet-list\"]/div/ul/li/div/div[2]/div[4]/a[2]/span";
        private const string CONFIRM_DELETE_BTN = "dataConfirmOK";


        //__________________________________Variables_______________________________________

        // General
        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;

        // Onglets
        [FindsBy(How = How.Id, Using = GENERAL_INFORMATION)]
        private IWebElement _generalInformationTab;

        // Add delivery
        [FindsBy(How = How.Id, Using = ADD_DELIVERY)]
        private IWebElement _addDelivery;

        [FindsBy(How = How.XPath, Using = FIRST_DELIVERY_RESULT)]
        private IWebElement _firstDeliveryResult;

        // Duplicate
        [FindsBy(How = How.XPath, Using = BTN_DUPLICATE)]
        private IWebElement _btnDuplicate;

        [FindsBy(How = How.XPath, Using = NAME_DUPLICATE)]
        private IWebElement _nameDuplicate;

        [FindsBy(How = How.Id, Using = STARTDATE_DUPLICATE)]
        private IWebElement _dateFromDuplicate;

        [FindsBy(How = How.Id, Using = ENDDATE_DUPLICATE)]
        private IWebElement _dateToDuplicate;

        [FindsBy(How = How.Id, Using = CONFIRM_DUPLICATE)]
        private IWebElement _confirmDuplicate;

        // Tableau

        [FindsBy(How = How.XPath, Using = EDIT_DELIVERY)]
        private IWebElement _editBtn;

        [FindsBy(How = How.XPath, Using = DELETE_BTN)]
        private IWebElement _deleteBtn;

        [FindsBy(How = How.Id, Using = CONFIRM_DELETE_BTN)]
        private IWebElement _confirmDeleteBtn;

        // _____________________________________________ Méthodes _________________________________________________

        //General
        public DeliveryRoundPage BackToList()
        {
            _backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            _backToList.Click();
            WaitForLoad();

            return new DeliveryRoundPage(_webDriver, _testContext);
        }

        // Onglets
        public DeliveryRoundGeneralInformationPage ClickOnGeneralInfoTab()
        {
            _generalInformationTab = WaitForElementIsVisible(By.Id(GENERAL_INFORMATION));
            _generalInformationTab.Click();
            WaitForLoad();

            return new DeliveryRoundGeneralInformationPage(_webDriver, _testContext);
        }

        // Add delivery
        public void AddDelivery(string deliveryName)
        {
            WaitForLoad();
            _addDelivery = WaitForElementIsVisible(By.Id(ADD_DELIVERY));
            _addDelivery.SetValue(ControlType.TextBox, deliveryName);
            WaitPageLoading();
            WaitForLoad();

            _firstDeliveryResult = WaitForElementIsVisible(By.XPath(FIRST_DELIVERY_RESULT));
            _firstDeliveryResult.Click();
            WaitForLoad();
        }

        // Duplicate
        public void Duplicate(string deliveryRound, DateTime start, DateTime end)
        {
            _btnDuplicate = WaitForElementIsVisible(By.XPath(BTN_DUPLICATE));
            _btnDuplicate.Click();
            WaitForLoad();

            _nameDuplicate = WaitForElementIsVisible(By.Id(NAME_DUPLICATE));
            _nameDuplicate.SetValue(ControlType.TextBox, deliveryRound);
            _nameDuplicate.SendKeys(Keys.Tab);

            _dateFromDuplicate = WaitForElementIsVisible(By.Id(STARTDATE_DUPLICATE));
            _dateFromDuplicate.SetValue(ControlType.DateTime, start);
            _dateFromDuplicate.Click();
            _dateFromDuplicate.SendKeys(Keys.Tab);

            _dateToDuplicate = WaitForElementIsVisible(By.Id(ENDDATE_DUPLICATE));
            _dateToDuplicate.SetValue(ControlType.DateTime, end);
            _dateToDuplicate.Click();
            _dateToDuplicate.SendKeys(Keys.Tab);

            _confirmDuplicate = WaitForElementIsVisible(By.Id(CONFIRM_DUPLICATE));
            _confirmDuplicate.Click();
            WaitForLoad();
        }

        // Tableau
        public Boolean IsDeliveryVisible()
        {
            if (isElementExists(By.XPath(FIRST_DELIVERY)))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public DeliveryLoadingPage EditDelivery()
        {
            _editBtn = WaitForElementIsVisible(By.XPath(EDIT_DELIVERY));
            _editBtn.Click();
            WaitForLoad();

            //Results are opened in a new tab, switch the driver to the newly created one
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            WaitForLoad();

            return new DeliveryLoadingPage(_webDriver, _testContext);
        }

        public void DeleteDelivery()
        {
            _deleteBtn = WaitForElementIsVisible(By.XPath(DELETE_BTN));
            _deleteBtn.Click();
            WaitForLoad();

            _confirmDeleteBtn = WaitForElementIsVisible(By.Id(CONFIRM_DELETE_BTN));
            _confirmDeleteBtn.Click();
            WaitForLoad();
        }
        public bool GetTotalTabs()
        {
            IList<string> allWindows = _webDriver.WindowHandles;
            return allWindows.Count > 1;
        }

    }
}