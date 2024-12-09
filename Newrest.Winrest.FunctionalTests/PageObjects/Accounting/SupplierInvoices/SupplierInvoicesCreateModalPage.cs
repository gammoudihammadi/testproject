using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;


namespace Newrest.Winrest.FunctionalTests.PageObjects.Accounting.SupplierInvoices
{
    public class SupplierInvoicesCreateModalPage : PageBase
    {
        public SupplierInvoicesCreateModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // ___________________________________ Constantes __________________________________________

        private const string INVOICE_NUMBER = "tb-new-supplier-invoice-number";
        private const string SITE = "drop-down-sites";
        private const string SUPPLIER = "//*[@id=\"form-create-supplier-invoice\"]//input[@class='supplier-input tt-input']";
        private const string SITE_PLACE = "drop-down-siteplaces";
        private const string INVOICE_DATE = "datapicker-new-supplier-invoice-date";
        private const string CREDIT_NOTE = "check-box-credit-note";
        private const string ACTIVATED = "SupplierInvoice_IsActive";
        private const string CREATE_FROM = "check-box-create-from";

        private const string RECEIPT_NOTE_LINK = "btn-tab-receipt-note";
        private const string RECEIPT_NOTE_NUMBER = "form-copy-from-tbSearchName";
        private const string RECEIPT_NOTE_SELECTED = "//*[@id=\"table-copy-from-rn\"]/tbody/tr[*]/td[2][normalize-space(text())='{0}']/../td[7]/div";

        private const string PURCHASE_ORDER_LINK = "btn-tab-purchase-order";
        private const string PURCHASE_ORDER_NUMBER = "form-copy-from-tbSearchName";
        private const string PURCHASE_ORDER_SELECTED = "//*[@id=\"table-copy-from-po\"]/tbody/tr[*]/td[2][normalize-space(text())='{0}']/../td[6]/div";

        private const string CLAIM_NUMBER = "form-copy-from-tbSearchName";
        private const string CLAIM_SELECTED = "//*[@id=\"table-copy-from-rn\"]/tbody/tr[*]/td[2][normalize-space(text())='{0}']/../td[7]/div/input[1]";
        private const string SUBMIT = "btn-submit-form-create-supplier-invoice";

        private const string ERROR_MESSAGE_NUMBER = "//*[@id=\"form-create-supplier-invoice\"]/div/div[1]/div[1]/div/div/span";
        private const string ERROR_MESSAGE_SITE = "//*[@id=\"form-create-supplier-invoice\"]/div/div[1]/div[2]/div/div/span";
        private const string ERROR_MESSAGE_SUPPLIER = "//*[@id=\"form-create-supplier-invoice\"]/div/div[2]/div/div/div/span";

        // __________________________________ Variables ______________________________________________

        [FindsBy(How = How.Id, Using = INVOICE_NUMBER)]
        private IWebElement _number;

        [FindsBy(How = How.Id, Using = SITE)]
        private IWebElement _site;

        [FindsBy(How = How.XPath, Using = SUPPLIER)]
        private IWebElement _supplier;

        [FindsBy(How = How.Id, Using = SITE_PLACE)]
        private IWebElement _sitePlace;

        [FindsBy(How = How.Id, Using = INVOICE_DATE)]
        private IWebElement _date;

        [FindsBy(How = How.Id, Using = ACTIVATED)]
        private IWebElement _activated;

        [FindsBy(How = How.Id, Using = CREATE_FROM)]
        private IWebElement _createFrom;

        [FindsBy(How = How.Id, Using = RECEIPT_NOTE_LINK)]
        private IWebElement _receiptNoteLink;

        [FindsBy(How = How.Id, Using = RECEIPT_NOTE_NUMBER)]
        private IWebElement _receiptNoteNumber;

        [FindsBy(How = How.XPath, Using = RECEIPT_NOTE_SELECTED)]
        private IWebElement _receiptNoteSelected;

        [FindsBy(How = How.Id, Using = PURCHASE_ORDER_LINK)]
        private IWebElement _purchaseOrderLink;

        [FindsBy(How = How.Id, Using = PURCHASE_ORDER_NUMBER)]
        private IWebElement _purchaseOrderNumber;

        [FindsBy(How = How.XPath, Using = PURCHASE_ORDER_SELECTED)]
        private IWebElement _purchaseOrderSelected;

        [FindsBy(How = How.Id, Using = CREDIT_NOTE)]
        private IWebElement _creditNote;

        [FindsBy(How = How.Id, Using = CLAIM_NUMBER)]
        private IWebElement _claimNumber;

        [FindsBy(How = How.XPath, Using = CLAIM_SELECTED)]
        private IWebElement _claimSelected;

        [FindsBy(How = How.Id, Using = SUBMIT)]
        private IWebElement _submit;

        // __________________________________________ Méthodes ______________________________________________

        public void FillField_CreateNewSupplierInvoices(string supplierNumber, DateTime date, string site, string supplier, bool isActivate = true, bool createFrom = false, string sitePlace = null)
        {
            WaitForLoad();

            _number = WaitForElementIsVisible(By.Id(INVOICE_NUMBER));
            _number.SetValue(ControlType.TextBox, supplierNumber);

            _site = WaitForElementIsVisible(By.Id(SITE));
            _site.SetValue(ControlType.DropDownList, site);
            WaitForLoad();

            //input with search bar
            _supplier = WaitForElementIsVisible(By.XPath(SUPPLIER));
            _supplier.SetValue(ControlType.TextBox, supplier);
            var searchSupplierValue = WaitForElementIsVisible(By.XPath("//div[text()='" + supplier + "']"));
            searchSupplierValue.Click();

            if (sitePlace != null)
            {
                _sitePlace = WaitForElementIsVisible(By.Id(SITE_PLACE));
                _sitePlace.SetValue(ControlType.DropDownList, sitePlace);
                WaitForLoad();
            }

            _date = WaitForElementIsVisible(By.Id(INVOICE_DATE));
            _date.SetValue(ControlType.DateTime, date);
            _date.SendKeys(Keys.Tab);
            WaitForLoad();

            if (!isActivate)
            {
                _activated = _webDriver.FindElement(By.Id(ACTIVATED));
                _activated.Click();
            }

            _createFrom = WaitForElementExists(By.Id(CREATE_FROM));
            _createFrom.SetValue(ControlType.CheckBox, createFrom);
        }

        public void FillField_CreateNewSupplierInvoicesCreditNote(string supplierNumber, DateTime date, string site, string supplier, bool isActivate = true, bool createFrom = false, string sitePlace = null,bool creditNote= false)
        {
            WaitForLoad();

            _number = WaitForElementIsVisible(By.Id(INVOICE_NUMBER));
            _number.SetValue(ControlType.TextBox, supplierNumber);

            _site = WaitForElementIsVisible(By.Id(SITE));
            _site.SetValue(ControlType.DropDownList, site);
            WaitForLoad();

            //input with search bar
            _supplier = WaitForElementIsVisible(By.XPath(SUPPLIER));
            _supplier.SetValue(ControlType.TextBox, supplier);
            var searchSupplierValue = WaitForElementIsVisible(By.XPath("//div[text()='" + supplier + "']"));
            searchSupplierValue.Click();

            if (sitePlace != null)
            {
                _sitePlace = WaitForElementIsVisible(By.Id(SITE_PLACE));
                _sitePlace.SetValue(ControlType.DropDownList, sitePlace);
                WaitForLoad();
            }

            _date = WaitForElementIsVisible(By.Id(INVOICE_DATE));
            _date.SetValue(ControlType.DateTime, date);
            _date.SendKeys(Keys.Tab);
            WaitForLoad();

            if (!isActivate)
            {
                _activated = _webDriver.FindElement(By.Id(ACTIVATED));
                _activated.Click();
            }

            _createFrom = WaitForElementExists(By.Id(CREATE_FROM));
            _createFrom.SetValue(ControlType.CheckBox, createFrom);

            if (!creditNote)
            {
                _creditNote = WaitForElementExists(By.Id(CREDIT_NOTE));
                _creditNote.Click();
            }
        }


        public void CreateSIFromRN(string receiptNoteID)
        {
            _receiptNoteLink = WaitForElementIsVisible(By.Id(RECEIPT_NOTE_LINK));
            _receiptNoteLink.Click();

            _receiptNoteNumber = WaitForElementIsVisible(By.Id(RECEIPT_NOTE_NUMBER));
            _receiptNoteNumber.SetValue(ControlType.TextBox, receiptNoteID);
            WaitPageLoading();
            WaitForLoad();

            _receiptNoteSelected = WaitForElementIsVisible(By.XPath(string.Format(RECEIPT_NOTE_SELECTED, receiptNoteID)));
            _receiptNoteSelected.Click();
        }

        public void CreateSIFromPO(string purchaseOrderID)
        {
            _purchaseOrderLink = WaitForElementIsVisible(By.Id(PURCHASE_ORDER_LINK));
            _purchaseOrderLink.Click();

            _purchaseOrderNumber = WaitForElementIsVisible(By.Id(PURCHASE_ORDER_NUMBER));
            _purchaseOrderNumber.SetValue(ControlType.TextBox, purchaseOrderID);
            WaitPageLoading();
            WaitForLoad();

            _purchaseOrderSelected = WaitForElementIsVisible(By.XPath(string.Format(PURCHASE_ORDER_SELECTED, purchaseOrderID)));
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", _purchaseOrderSelected);
            _purchaseOrderSelected.Click();
        }

        public void CreateSIFromClaim(string claimID)
        {
            _creditNote = WaitForElementExists(By.Id(CREDIT_NOTE));
            _creditNote.Click();
            WaitForLoad();

            _claimNumber = WaitForElementIsVisible(By.Id(CLAIM_NUMBER));
            _claimNumber.SetValue(ControlType.TextBox, claimID);
            WaitPageLoading();
            WaitForLoad();

            _claimSelected = WaitForElementExists(By.XPath(string.Format(CLAIM_SELECTED, claimID)));
            _claimSelected.SetValue(ControlType.CheckBox, true);
            WaitPageLoading();
            WaitForLoad();
        }

        public void CreateSIFromIR(string site, string supplierName, string interimReceptionNumber, string supplierNumber)
        {
            _number = WaitForElementIsVisible(By.Id(INVOICE_NUMBER));
            _number.SetValue(ControlType.TextBox, new Random().Next().ToString());
            WaitForLoad();

            // séquence valide : site, onglet interim receptions,supplier,interimReceptionNumber

            //site
            _site = WaitForElementIsVisible(By.Id(SITE));
            _site.SetValue(ControlType.DropDownList, site);
            WaitForLoad();

            // onglet interim receptions
            var _interimReceptionLink = WaitForElementIsVisible(By.Id("btn-tab-interim-reception"));
            _interimReceptionLink.Click();
            WaitForLoad();

            //supplier
            _supplier = WaitForElementIsVisible(By.XPath(SUPPLIER));
            _supplier.SetValue(ControlType.TextBox, supplierName);
            var searchSupplierValue = WaitForElementIsVisible(By.XPath("//div[text()='" + supplierName + "']"));
            searchSupplierValue.Click();
            WaitForLoad();

            //interimReceptionNumber
            var _interimReceptionNumber = WaitForElementIsVisible(By.Id("form-copy-from-tbSearchName"));
            _interimReceptionNumber.SetValue(ControlType.TextBox, interimReceptionNumber);
            WaitPageLoading();
            WaitForLoad();
            
            var _interimReceptionSelected = WaitForElementExists(By.Id("item_IsSelected"));
            _interimReceptionSelected.SetValue(ControlType.CheckBox, true);
            WaitForLoad();
        }

        public SupplierInvoicesItem Submit()
        {
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;

            _submit = WaitForElementExists(By.Id(SUBMIT));
            javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", _submit);
            WaitForLoad();

            _submit.Click();
            WaitPageLoading();
            WaitForLoad();

            return new SupplierInvoicesItem(_webDriver, _testContext);
        }

        public bool ErrorMessageInvoiceNumberRequired()
        {
            var errorMessageNumber = _webDriver.FindElement(By.XPath(ERROR_MESSAGE_NUMBER));
            return errorMessageNumber.Displayed;
        }

        public bool ErrorMessageSiteRequired()
        {
            var errorMessageSite = _webDriver.FindElement(By.XPath(ERROR_MESSAGE_SITE));
            return errorMessageSite.Displayed;
        }

        public bool ErrorMessageSupplierRequired()
        {
            var errorMessageSupplier = _webDriver.FindElement(By.XPath(ERROR_MESSAGE_SUPPLIER));
            return errorMessageSupplier.Displayed;
        }
    }
}
