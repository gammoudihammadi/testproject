using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using static Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service.ServiceMassiveDeleteModalPage;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.Customer
{
    public class CustomerContactsPage : PageBase
    {

        public CustomerContactsPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //______________________________________ Constantes ____________________________________________

        // General
        private const string NEW_CONTACT = "//*[@id=\"tabContentDetails\"]/div/div[1]/div/a[2]";
        private const string CONTACT_NAME = "Contact_Name";
        private const string CONTACT_MAIL = "Contact_Mail";
        private const string CONTACT_PHONE = "Contact_Phone";
        private const string CONTACT_SITE = "SelectedSites_ms";
        private const string SEARCH_SITE = "/html/body/div[12]/div/div/label/input";
        private const string CANCEL = "//button[text()='Cancel']";
        private const string CREATE = "//button[text()='Create']";
        private const string SAVE = "//button[text()='Save']";
        private const string ERROR_MESSAGE = "//*[@id=\"modal-1\"]//form/div[1]/span";
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";

        // Tableau
        private const string FIRST_CONTACT =      "//*[@id=\"list-item-with-action\"]/div/div[1]/div/div[2]";
        private const string ADRESS_CUSTOMER = "//*[@id=\"list-item-with-action\"]/div/div[2]/div/dl/dt/dd";
        private const string FIRST_CONTACT_MAIL = "//div/dl/dd[2]";
        private const string EDIT_CONTACT = "//*[@id=\"list-item-with-action\"]/div/div[1]/div/div[3]/a[1]";
        private const string DELETE_CONTACT = "//*[@id=\"list-item-with-action\"]/div/div[1]/div/div[3]/a[2]";
        private const string CONFIRM_DELETE = "/html/body/div[11]/div/div/div[3]/a[1]";     

        //_____________________________________ Variables________________________________________________

        // General
        [FindsBy(How = How.XPath, Using = NEW_CONTACT)]
        private IWebElement _newContact;

        [FindsBy(How = How.Id, Using = CONTACT_NAME)]
        private IWebElement _contactName;

        [FindsBy(How = How.Id, Using = CONTACT_MAIL)]
        private IWebElement _contactMail;

        [FindsBy(How = How.Id, Using = CONTACT_PHONE)]
        private IWebElement _contactPhone;

        [FindsBy(How = How.Id, Using = CONTACT_SITE)]
        private IWebElement _contactSite;

        [FindsBy(How = How.XPath, Using = SEARCH_SITE)]
        private IWebElement _searchSite;

        [FindsBy(How = How.XPath, Using = CANCEL)]
        private IWebElement _cancelBtn;

        [FindsBy(How = How.XPath, Using = CREATE)]
        private IWebElement _createBtn;

        [FindsBy(How = How.XPath, Using = SAVE)]
        private IWebElement _saveBtn;

        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;

        // Tableau

        [FindsBy(How = How.XPath, Using = ADRESS_CUSTOMER)]
        private IWebElement _adresscustomer;

        [FindsBy(How = How.XPath, Using = FIRST_CONTACT)]
        private IWebElement _firstContact;

        [FindsBy(How = How.XPath, Using = FIRST_CONTACT_MAIL)]
        private IWebElement _firstContactMail;

        [FindsBy(How = How.XPath, Using = EDIT_CONTACT)]
        private IWebElement _editContact;

        [FindsBy(How = How.XPath, Using = DELETE_CONTACT)]
        private IWebElement _deleteContact;

        [FindsBy(How = How.XPath, Using = CONFIRM_DELETE)]
        private IWebElement _confirmDelete;
        

        //_____________________________________ Méthodes __________________________________________________

        // General
        public void CreateContact(string contactName, string mail, string phone, string site)
        {
            _newContact = WaitForElementIsVisible(By.XPath(NEW_CONTACT));
            _newContact.Click();
            WaitForLoad();

            _contactName = WaitForElementIsVisible(By.Id(CONTACT_NAME));
            _contactName.SetValue(ControlType.TextBox, contactName);

            _contactMail = WaitForElementIsVisible(By.Id(CONTACT_MAIL));
            _contactMail.SetValue(ControlType.TextBox, mail);

            _contactPhone = WaitForElementIsVisible(By.Id(CONTACT_PHONE));
            _contactPhone.SetValue(ControlType.TextBox, phone);

             ComboBoxSelectById(new ComboBoxOptions(CONTACT_SITE, site, false));
              WaitPageLoading();
            _createBtn = WaitForElementIsVisible(By.XPath(CREATE));
            _createBtn.Click();
            WaitForLoad();
        }

        public void EditContact(string mail)
        {
            WaitPageLoading();
            _firstContact = WaitForElementIsVisible(By.XPath(FIRST_CONTACT));
            _firstContact.Click();

            _editContact = WaitForElementIsVisible(By.XPath(EDIT_CONTACT));
            _editContact.Click();
            WaitForLoad();

            // Renseigner le mail
            _contactMail = WaitForElementIsVisible(By.Id(CONTACT_MAIL));
            _contactMail.SetValue(ControlType.TextBox, mail);

            _saveBtn = WaitForElementIsVisible(By.XPath(SAVE));
            _saveBtn.Click();
            WaitForLoad();
        }

        public void Cancel()
        {
            _cancelBtn = WaitForElementIsVisible(By.XPath(CANCEL));
            _cancelBtn.Click();
            WaitForLoad();
        }

        public bool IsErrorMessageDisplayed()
        {
            if(isElementVisible(By.XPath(ERROR_MESSAGE)))
            {
                _webDriver.FindElement(By.XPath(ERROR_MESSAGE));
                return true;
            }
            else
            {
                return false;
            }
        }

        // Tableau
        public void ClickFirstContact()
        {
            _firstContact = WaitForElementIsVisible(By.XPath(FIRST_CONTACT));
            _firstContact.Click();
            WaitForLoad();
        }

        public string GetFirstContactMail()
        {
            WaitPageLoading();
            _firstContactMail = WaitForElementIsVisible(By.XPath(FIRST_CONTACT_MAIL));
            return _firstContactMail.Text;
        }
        public string GetContactName()
        {
            if(isElementVisible(By.XPath(FIRST_CONTACT)))
            {
                _firstContact = WaitForElementIsVisible(By.XPath(FIRST_CONTACT));
                return _firstContact.Text;
            }
            else
            {
                return "";
            }
        }

        public string GetContactAdress()
        {
            if (isElementVisible(By.XPath(ADRESS_CUSTOMER)))
            {
                _adresscustomer = WaitForElementIsVisible(By.XPath(ADRESS_CUSTOMER));
                return _adresscustomer.Text;
            }
            else
            {
                return "";
            }
        }


        public bool IsContactDisplayed()
        {
            if(isElementVisible(By.XPath(FIRST_CONTACT)))
            {
                _firstContact = _webDriver.FindElement((By.XPath(FIRST_CONTACT)));
                return _firstContact.Displayed;
            }
            else
            {
                return false;
            }
        }

        public void DeleteContact()
        {
            _deleteContact = WaitForElementIsVisible(By.XPath(DELETE_CONTACT));
            _deleteContact.Click();
            WaitForLoad();

            _confirmDelete = WaitForElementIsVisible(By.XPath(CONFIRM_DELETE));
            _confirmDelete.Click();
            WaitForLoad();
        }
        public CustomerPage BackToList()
        {
            _backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            _backToList.Click();
            WaitForLoad();

            return new CustomerPage(_webDriver, _testContext);
        }

    }
}
