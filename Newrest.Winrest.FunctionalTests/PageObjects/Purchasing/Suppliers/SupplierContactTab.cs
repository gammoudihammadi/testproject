using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Suppliers
{
    public class SupplierContactTab : PageBase
    {
        public SupplierContactTab(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {

        }

        private const string PHONE = "//*/div[@aria-expanded='true' and contains(@id,'content_')]/div/div/div/div/dl/dd[1]";
        private const string FAX = "//*/div[@aria-expanded='true' and contains(@id,'content_')]/div/div/div/div/dl/dd[2]";
        private const string MAIL = "//*/div[@aria-expanded='true' and contains(@id,'content_')]/div/div/div/div/dl/dd[3]";
        private const string ADDRESS = "//*/div[@aria-expanded='true' and contains(@id,'content_')]/div/div/div/div/dl/dd[4]";
        private const string FOR_PO = "//*/div[@aria-expanded='true' and contains(@id,'content_')]/div/div/div/div/dl/dd[4]/b[1]";
        private const string FOR_CLAIM = "//*/div[@aria-expanded='true' and contains(@id,'content_')]/div/div/div/div/dl/dd[4]/b[2]";
        private const string DELETE_CONTACT_BUTTONS = "//div[contains(@class, 'row-content') and contains(@title, 'PMI') and @data-toggle='collapse']\r\n";
        private const string CONFIRM_BUTTON = "//*[@id=\"dataConfirmOK\"]";
        private const string NEW_CONTACT_BUTTON = "//*[@id=\"tabContentDetails\"]/div/div[1]/div/a[2]";
        private const string CONTACT_NAME_FIELD = "//*[@id=\"Contact_Name\"]";
        private const string CONTACT_EMAIL_FIELD = "//*[@id=\"Contact_Mail\"]";
        private const string CONTACT_PHONE_FIELD = "//*[@id=\"Contact_Phone\"]";
        private const string CONTACT_ADDRESS_FIELD = "//*[@id=\"Contact_Address1\"]";
        private const string CONTACT_ZIP_FIELD = "//*[@id=\"Contact_ZipCode\"]";
        private const string CONTACT_CITY_FIELD = "//*[@id=\"Contact_City\"]";
        private const string CONTACT_CREATE_BUTTON = "//*[@id=\"modal-1\"]/div[2]/div/form/div[2]/button[2]";
        private const string CONTACT_LIST_ITEMS = "//*[@id=\"list-item-with-action\"]";
        private const string CONTACT_TAB = "hrefTabContentContacts";
        private const string CONTACT_ELEMENT = "//*[@id=\"list-item-with-action\"]";


        [FindsBy(How = How.XPath, Using = PHONE)]
        private IWebElement _phone;

        [FindsBy(How = How.XPath, Using = FAX)]
        private IWebElement _fax;

        [FindsBy(How = How.XPath, Using = MAIL)]
        private IWebElement _mail;

        [FindsBy(How = How.XPath, Using = ADDRESS)]
        private IWebElement _address;

        [FindsBy(How = How.XPath, Using = FOR_PO)]
        private IWebElement _forPO;

        [FindsBy(How = How.XPath, Using = FOR_CLAIM)]
        private IWebElement _forClaim;



        public void Deplier(string site)
        {
            var contactDeplier = WaitForElementIsVisible(By.XPath("//*/span[contains(text(), '- " + site + "')]"));
            contactDeplier.Click();
        }

        public string GetPhone()
        {
                _phone = WaitForElementIsVisible(By.XPath("//*/div[@class='panel-collapse collapse show']/div/dl/dd[1]"));
            
            return _phone.Text.Trim();
        }

        public string GetFax()
        {
                _fax = WaitForElementIsVisible(By.XPath("//*/div[@class='panel-collapse collapse show']/div/dl/dd[2]"));
            
            return _fax.Text.Trim();
        }

        public string GetMail()
        {
                _mail = WaitForElementIsVisible(By.XPath("//*/div[@class='panel-collapse collapse show']/div/dl/dd[3]"));
            
            return _mail.Text.Trim();
        }

        public string[] GetAdress()
        {
                _address = WaitForElementIsVisible(By.XPath("//*/div[@class='panel-collapse collapse show']/div/dl/dd[4]"));
           
            var resultat2 = ((IJavaScriptExecutor)_webDriver).ExecuteScript("return arguments[0].innerHTML;", _address);
            var resultat3 = ((string)resultat2).Replace("\r\n", "");
            return resultat3.Split(new String[] { "<br>" }, StringSplitOptions.None);
        }

        public string GetForPO()
        {
                _forPO = WaitForElementIsVisible(By.XPath("//*/div[@class='panel-collapse collapse show']/div/dl/dd[4]/b[1]"));
            
            return _forPO.Text.Trim();
        }

        public string GetForClaim()
        {
                _forClaim = WaitForElementIsVisible(By.XPath("//*/div[@class='panel-collapse collapse show']/div/dl/dd[4]/b[2]"));
           
            return _forClaim.Text.Trim();
        }

        public SupplierAccountingsTab ClickOnAccountingsTab()
        {
            var tab = WaitForElementExists(By.Id("hrefTabContentAccounting"));
            tab.Click();
            WaitForLoad();
            return new SupplierAccountingsTab(_webDriver, _testContext);
        }

        public void FillNewContactForm(string name, string email, string phone, string address, string zip, string city)
        {
            var nameField = WaitForElementIsVisible(By.XPath(CONTACT_NAME_FIELD));
            nameField.SendKeys(name);

            var emailField = WaitForElementIsVisible(By.XPath(CONTACT_EMAIL_FIELD));
            emailField.SendKeys(email);

            var phoneField = WaitForElementIsVisible(By.XPath(CONTACT_PHONE_FIELD));
            phoneField.SendKeys(phone);

            var addressField = WaitForElementIsVisible(By.XPath(CONTACT_ADDRESS_FIELD));
            addressField.SendKeys(address);

            var zipField = WaitForElementIsVisible(By.XPath(CONTACT_ZIP_FIELD));
            zipField.SendKeys(zip);

            var cityField = WaitForElementIsVisible(By.XPath(CONTACT_CITY_FIELD));
            cityField.SendKeys(city);

        }

        public void SubmitNewContactForm()
        {
            var createButton = WaitForElementIsVisible(By.XPath(CONTACT_CREATE_BUTTON));
            createButton.Click();
            WaitForLoad();
        }
        private IReadOnlyCollection<IWebElement> WaitForElementsAreVisible(By by)
        {
            WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(100));
            return wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.VisibilityOfAllElementsLocatedBy(by));
        }

        public List<string> GetContactList()
        {
            // Locate all elements matching the XPath
            var contactElements = WaitForElementsAreVisible(By.XPath(CONTACT_LIST_ITEMS));

            // Print the number of elements found
            Console.WriteLine($"Found {contactElements.Count} contact elements.");

            // Print the text of each element before filtering
            foreach (var contactElement in contactElements)
            {
                Console.WriteLine($"Contact element text: '{contactElement.Text}'");
            }

            // Filter out elements with empty or whitespace text
            var validContacts = contactElements
                .Where(ce => !string.IsNullOrWhiteSpace(ce.Text))
                .ToList();

            // Return the list of valid contact texts
            return validContacts.Select(ce => ce.Text).ToList();
        }

        public SupplierContactTab GoToContact()
        {
            var contactsTabElement = _webDriver.FindElement(By.Id(CONTACT_TAB));
            contactsTabElement.Click();

            return new SupplierContactTab(_webDriver, _testContext);
        }

    }
}
