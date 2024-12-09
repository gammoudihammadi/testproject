using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Accounting.SupplierInvoices
{
    public class SupplierInvoicesGeneralInformation : PageBase
    {

        public SupplierInvoicesGeneralInformation(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //________________________________ Constantes ______________________________________

        // Général
        private const string BACK_TO_LIST = "BackToList";

        // Onglets
        private const string ITEMS_TAB = "hrefTabContentItems";

        // Tableau
        private const string INVOICE_NUMBER = "tb-new-supplier-invoice-number";
        private const string COMMENT = "commentTextarea";
        private const string SITE = "//*[@id='form-create-supplier-invoice']/div/div[1]/div[2]/div/div/input[3]";
        private const string SUPPLIER = "//*[@id='form-create-supplier-invoice']/div/div[2]/div/div/div/input[2]";
        private const string PLACE = "//*[@id='form-create-supplier-invoice']/div/div[3]/div/div/div/input";
        private const string DATE = "datapicker-new-supplier-invoice-date";
        private const string ACTIVE = "//*[@id='SupplierInvoice_IsActive']/parent::div";
        private const string RELATED_CLAIM_ID = "//*[@id=\"form-create-supplier-invoice\"]/div/div[9]/div/div/div/ul/li/a";

        //_________________________________ Variables ______________________________________

        // Général 
        [FindsBy(How = How.Id, Using = BACK_TO_LIST)]
        private IWebElement _backToList;

        // Onglets
        [FindsBy(How = How.Id, Using = ITEMS_TAB)]
        private IWebElement _itemsTab;

        // Tableau
        [FindsBy(How = How.Id, Using = INVOICE_NUMBER)]
        private IWebElement _invoiceNumber;

        [FindsBy(How = How.Id, Using = SITE)]
        private IWebElement _site;

        [FindsBy(How = How.Id, Using = SUPPLIER)]
        private IWebElement _supplier;

        [FindsBy(How = How.Id, Using = PLACE)]
        private IWebElement _place;

        [FindsBy(How = How.Id, Using = DATE)]
        private IWebElement _date;

        [FindsBy(How = How.Id, Using = ACTIVE)]
        private IWebElement _active;

        [FindsBy(How = How.Id, Using = COMMENT)]
        private IWebElement _comments;

        [FindsBy(How = How.Id, Using = RELATED_CLAIM_ID)]
        private IWebElement _relatedClaim;

        //________________________________ Méthodes ________________________________________

        // Général
        public SupplierInvoicesPage BackToList()
        {
            _backToList = WaitForElementIsVisible(By.Id(BACK_TO_LIST));
            _backToList.Click();
            WaitForLoad();

            return new SupplierInvoicesPage(_webDriver, _testContext);
        }

        // Onglets
        public SupplierInvoicesItem ClickOnItems()
        {
            _itemsTab = WaitForElementIsVisible(By.Id(ITEMS_TAB));
            _itemsTab.Click();
            WaitForLoad();

            return new SupplierInvoicesItem(_webDriver, _testContext);
        }

        // Tableau
        public void SetSupplierInvoiceNumber(string invoiceNumber)
        {
            _invoiceNumber = WaitForElementIsVisible(By.Id(INVOICE_NUMBER));
            _invoiceNumber.SetValue(ControlType.TextBox, invoiceNumber);
            WaitPageLoading(); 
        }

        public string GetSupplierInvoiceNumber()
        {
            _invoiceNumber = WaitForElementExists(By.Id(INVOICE_NUMBER));
            return _invoiceNumber.GetAttribute("value");
        }
        public string GetSite()
        {
            _site = WaitForElementIsVisible(By.XPath(SITE));
            return _site.GetAttribute("value");
        }

        public string GetSupplier()
        {
            _supplier = WaitForElementIsVisible(By.XPath(SUPPLIER));
            return _supplier.GetAttribute("value");
        }
        public string GetPlace()
        {
            _place = WaitForElementIsVisible(By.XPath(PLACE));
            return _place.GetAttribute("value");
        }

        public void SetActive(bool value)
        {
            _active = WaitForElementIsVisible(By.XPath(ACTIVE));
            _active.Click();
            // Ophélie : Temps d'enregistrement de la nouvelle valeur dans la page (ne pas enlever)
            WaitPageLoading();
            WaitForLoad();
        }

        public void SetComments(string newcoms)
        {
            _comments = WaitForElementIsVisible(By.Id(COMMENT));
            WaitPageLoading(); 
            _comments.SetValue(ControlType.TextBox, newcoms);
            WaitPageLoading();
        }

        public string GetComments()
        {
            _comments = WaitForElementIsVisible(By.Id(COMMENT));
            return _comments.GetAttribute("value");
        }

        public string GetRelatedClaimNumber()
        {
            
            _relatedClaim = WaitForElementIsVisible(By.XPath(RELATED_CLAIM_ID));
            return _relatedClaim.Text;
        }

        public string GetDate()
        {
            _date = WaitForElementIsVisible(By.Id(DATE));
            return _date.GetAttribute("value");
        }
    }
}
