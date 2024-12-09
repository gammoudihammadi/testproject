using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Linq;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Portal
{
    public class ParametersPortal : PageBase
    {

        public ParametersPortal(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {

        }

        // ___________________________________ Constantes ______________________________________________

        private const string ADD_NEW_USER = "//*[@id=\"add-new-user-link\"]/a";
        private const string EMAIL = "first";
        private const string LANGUAGE = "User_Language_Id";
        private const string CREATE = "last";

        private const string SEARCH_USER = "tbSearchPattern";
        private const string SELECT_USER = "//h4[text()='{0}']";

        private const string SEARCH_CUSTOMER_COMBOBOX = "//*[@id=\"SelectedCustomers_ms\"]";
        private const string SEARCH_CUSTOMER_TEXT = "/html/body/div[10]/div/div/label/input";
        private const string SEARCH_CUSTOMER_CHECKBOX = "//*[@id=\"ui-multiselect-0-SelectedCustomers-option-{0}\"]";
        private const string SEARCH_CUSTOMER_NAME = "//*[@id=\"ui-multiselect-0-SelectedCustomers-option-{0}\"]/following-sibling::span";

        private const string SEARCH_SITE_COMBOBOX = "//*[@id=\"SelectedSites_ms\"]";
        private const string SEARCH_SITE_TEXT = "/html/body/div[11]/div/div/label/input";
        private const string SEARCH_SITE_CHECKBOX = "//*[@id=\"ui-multiselect-1-SelectedSites-option-{0}\"]";
        private const string SEARCH_SITE_NAME = "//*[@id=\"ui-multiselect-1-SelectedSites-option-{0}\"]/following-sibling::span";

        private const string QUANTITY_BY_WEEK = "User_QuantityByWeek";
        private const string CUSTOMER_ORDERS = "User_CustomerOrders";
        private const string QUANTITY_BY_DAY = "User_QuantityByDay";
        private const string SAVE_BTN = "last";
        private const string CONFIRM_BTN = "dataConfirmOK";

        private const string AFFECTED_DELIVERY_TAB = "EditUserDelivery";
        private const string DELIVERY_LINE = "//*[@id=\"delivery_affectation\"]/tbody/tr[*]/td[2]/div/div[2]/label[text()='{0}']";

        // ___________________________________ Variables _______________________________________________

        [FindsBy(How = How.XPath, Using = ADD_NEW_USER)]
        private IWebElement _addNewUser;

        [FindsBy(How = How.Id, Using = EMAIL)]
        private IWebElement _email;

        [FindsBy(How = How.Id, Using = LANGUAGE)]
        private IWebElement _language;

        [FindsBy(How = How.Id, Using = CREATE)]
        private IWebElement _create;

        [FindsBy(How = How.Id, Using = QUANTITY_BY_WEEK)]
        private IWebElement _quantityByWeek;

        [FindsBy(How = How.Id, Using = QUANTITY_BY_DAY)]
        private IWebElement _quantityByDay;

        [FindsBy(How = How.Id, Using = CUSTOMER_ORDERS)]
        private IWebElement _customerOrders;

        [FindsBy(How = How.Id, Using = SAVE_BTN)]
        private IWebElement _saveBtn;

        [FindsBy(How = How.Id, Using = CONFIRM_BTN)]
        private IWebElement _confirmBtn;

        [FindsBy(How = How.Id, Using = SEARCH_USER)]
        private IWebElement _searchUser;

        [FindsBy(How = How.Id, Using = AFFECTED_DELIVERY_TAB)]
        private IWebElement _affectedDeliveryTab;

        [FindsBy(How = How.XPath, Using = DELIVERY_LINE)]
        private IWebElement _deliveryLine;

        // __________________________________ Méthodes __________________________________________________

        public void AddNewUser(string userName)
        {
            _addNewUser = WaitForElementIsVisible(By.XPath(ADD_NEW_USER));
            _addNewUser.Click();
            WaitForLoad();

            _email = WaitForElementIsVisible(By.Id(EMAIL));
            _email.SetValue(ControlType.TextBox, userName);

            _language = WaitForElementIsVisible(By.Id(LANGUAGE));
            _language.SetValue(ControlType.DropDownList, "English");

            _create = WaitForElementToBeClickable(By.Id(CREATE));
            _create.Click();
            WaitForLoad();
        }

        public void SearchAndSelectUser(string userName)
        {
            //Recherche du user
            _searchUser = WaitForElementIsVisible(By.Id(SEARCH_USER));
            _searchUser.SetValue(ControlType.TextBox, userName);
            WaitPageLoading();

            var selectedUser = WaitForElementIsVisible(By.XPath(String.Format(SELECT_USER, userName)));
            selectedUser.Click();
            WaitPageLoading();
        }

        public bool IsExistingUser(string userName)
        {
            //Recherche du user
            _searchUser = WaitForElementIsVisible(By.Id(SEARCH_USER));
            _searchUser.SetValue(ControlType.TextBox, userName);
            WaitPageLoading();

            try
            {
                _webDriver.FindElement(By.XPath(String.Format(SELECT_USER, userName)));
                return true;

            }
            catch
            {
                return false;
            }

        }

        public void ConfigureUser(bool quantityByWeek, bool customerOrder, bool quantityByDay)
        {
            Actions action = new Actions(_webDriver);

            _quantityByWeek = WaitForElementExists(By.Id(QUANTITY_BY_WEEK));
            action.MoveToElement(_quantityByWeek).Perform();
            _quantityByWeek.SetValue(ControlType.CheckBox, quantityByWeek);

            _customerOrders = WaitForElementExists(By.Id(CUSTOMER_ORDERS));
            action.MoveToElement(_customerOrders).Perform();
            _customerOrders.SetValue(ControlType.CheckBox, customerOrder);

            _quantityByDay = WaitForElementExists(By.Id(QUANTITY_BY_DAY));
            action.MoveToElement(_quantityByDay).Perform();
            _quantityByDay.SetValue(ControlType.CheckBox, quantityByDay);

            _saveBtn = WaitForElementIsVisible(By.Id(SAVE_BTN));
            _saveBtn.Click();
            WaitForLoad();
        }

        public string LinkDeliveryToUser(string deliveryName)
        {
            Actions action = new Actions(_webDriver);

            _affectedDeliveryTab = WaitForElementIsVisible(By.Id(AFFECTED_DELIVERY_TAB));
            _affectedDeliveryTab.Click();
            WaitForLoad();

            _deliveryLine = WaitForElementExists(By.XPath(String.Format(DELIVERY_LINE, deliveryName)));
            action.MoveToElement(_deliveryLine).Perform();

            var deliveryBox = _webDriver.FindElement(By.Id(_deliveryLine.GetAttribute("for")));
            if (deliveryBox.GetAttribute("checked") != "true")
                deliveryBox.Click();

            // Temps d'enregistrement de la donnée
            WaitPageLoading();

            return deliveryBox.GetAttribute("checked");
        }

        public string SearchAndSelectCustomer(bool firstTime, string customer, int number, bool selected)
        {
            ComboBoxOptions cbOpt = new ComboBoxOptions("SelectedCustomers_ms", customer, false) { ClickCheckAllAtStart = false, ClickUncheckAllAtStart = firstTime };
            ComboBoxSelectById(cbOpt);

            return customer;
        }

        public string SearchAndSelectCustomerSite(bool firstTime, string site, int number, bool selected)
        {
            string siteSearch = site + " - " + site;

            ComboBoxOptions cbOpt = new ComboBoxOptions("SelectedSites_ms", siteSearch, false) { ClickCheckAllAtStart = false, ClickUncheckAllAtStart = firstTime };
            ComboBoxSelectById(cbOpt);

            return siteSearch;
        }

        public void SetLanguage(string language)
        {
            _language = WaitForElementIsVisible(By.Id(LANGUAGE));
            _language.SetValue(ControlType.DropDownList, language);
            WaitForLoad();
            WaitPageLoading();

        }
        public string GetLanguage()
        {
            _language = WaitForElementIsVisible(By.Id(LANGUAGE));
             string selectedLanguage = _language.FindElement(By.CssSelector("option[selected='selected']")).Text;
            WaitPageLoading();
           return selectedLanguage;
        }
        public void SaveAndConfirm()
        {
            _saveBtn = WaitForElementIsVisible(By.Id(SAVE_BTN));
            _saveBtn.Click();
            WaitForLoad();
            if (isElementVisible(By.Id(CONFIRM_BTN)))
            {
                _confirmBtn = WaitForElementIsVisible(By.Id(CONFIRM_BTN));
                _confirmBtn.Click();
                WaitPageLoading();
            }
        }
    }
}
