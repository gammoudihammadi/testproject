using System;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;


namespace Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.OutputForm
{
    public class OutputFormCreateModalPage : PageBase
    {
        public OutputFormCreateModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // _______________________________________ Constantes _________________________________________________

        private const string NUMBER = "tb-new-outputform-number";
        private const string SITE = "drop-down-sites";
        private const string PLACE_FROM = "drop-down-places-from";
        private const string PLACE_TO = "drop-down-places-to";
        private const string DATE = "datapicker-output-form-date";
        private const string ACTIVATED = "OutputForm_IsActive";
        private const string CREATE_FROM = "check-box-create-from";
        private const string TAB_SUPPLIER_ORDER = "btn-tab-supply-order";
        private const string TAB_SUPPLIER_ORDER_CHECKBOX = "item_IsSelected";
        private const string TAB_SUPPLIER_ORDER_NUMBER = "//*[@id='div-create-from-supply-orders']/table/tbody/tr[2]/td[1]";
        private const string TAB_RECEIPT_NOTE = "btn-tab-receipt-note";
        private const string TAB_RECEIPT_NOTE_CHECKBOX = "item_IsSelected";
        private const string TAB_RECEIPT_NOTE_NUMBER = "//*[@id='div-create-from-receipt-notes']/table/tbody/tr[2]/td[1]";
        private const string TAB_PURCHASE_ORDER = "btn-tab-purchase-order";
        private const string TAB_PURCHASE_ORDER_CHECKBOX = "item_IsSelected";
        private const string TAB_PURCHASE_ORDER_NUMBER = "//*[@id='div-create-from-purchase-orders']/table/tbody/tr[2]/td[1]";
        private const string TAB_OUTPUT_FORM = "btn-tab-output-form";
        private const string TAB_OUTPUT_FORM_CHECKBOX = "item_IsSelected";
        private const string TAB_OUTPUT_FORM_NUMBER = "//*[@id='table-copy-from-of']/tbody/tr[2]/td[1]";


        private const string SUBMIT = "btn-submit-create-output-form";

        // _______________________________________ Variables __________________________________________________

        [FindsBy(How = How.Id, Using = NUMBER)]
        private IWebElement _outputFormNumber;

        [FindsBy(How = How.Id, Using = DATE)]
        private IWebElement _date;

        [FindsBy(How = How.Id, Using = SITE)]
        private IWebElement _site;

        [FindsBy(How = How.Id, Using = PLACE_FROM)]
        private IWebElement _placeFrom;

        [FindsBy(How = How.Id, Using = PLACE_TO)]
        private IWebElement _placeTo;

        [FindsBy(How = How.Id, Using = ACTIVATED)]
        private IWebElement _activated;

        [FindsBy(How = How.Id, Using = CREATE_FROM)]
        private IWebElement _createFrom;

        [FindsBy(How = How.Id, Using = TAB_SUPPLIER_ORDER)]
        private IWebElement _tabSupplierOrder;

        [FindsBy(How = How.Id, Using = TAB_SUPPLIER_ORDER_CHECKBOX)]
        private IWebElement _tabSupplierOrderCheckbox;

        [FindsBy(How = How.Id, Using = TAB_SUPPLIER_ORDER_NUMBER)]
        private IWebElement _tabSupplierOrderNumber;

        [FindsBy(How = How.Id, Using = TAB_RECEIPT_NOTE)]
        private IWebElement _tabReceiptNote;

        [FindsBy(How = How.Id, Using = TAB_RECEIPT_NOTE_CHECKBOX)]
        private IWebElement _tabReceiptNoteCheckbox;

        [FindsBy(How = How.Id, Using = TAB_RECEIPT_NOTE_NUMBER)]
        private IWebElement _tabReceiptNoteNumber;

        [FindsBy(How = How.Id, Using = TAB_PURCHASE_ORDER)]
        private IWebElement _tabPurchaseOrder;

        [FindsBy(How = How.Id, Using = TAB_PURCHASE_ORDER_CHECKBOX)]
        private IWebElement _tabPurchaseOrderCheckbox;

        [FindsBy(How = How.Id, Using = TAB_PURCHASE_ORDER_NUMBER)]
        private IWebElement _tabPurchaseOrderNumber;

        [FindsBy(How = How.Id, Using = TAB_OUTPUT_FORM)]
        private IWebElement _tabOutputForm;

        [FindsBy(How = How.Id, Using = TAB_OUTPUT_FORM_CHECKBOX)]
        private IWebElement _tabOutputFormCheckbox;

        [FindsBy(How = How.Id, Using = TAB_OUTPUT_FORM_NUMBER)]
        private IWebElement _tabOutputFormNumber;
        

                [FindsBy(How = How.Id, Using = SUBMIT)]
        private IWebElement _submit;

        // ______________________________________ Méthodes ___________________________________________________

        public void FillField_CreatNewOutputForm(DateTime date, string site, string placeFrom, string placeTo, bool isActive = true, bool isCreateFrom = false)
        {
            _site = WaitForElementIsVisibleNew(By.CssSelector("#drop-down-sites"));
            _site.SetValue(ControlType.DropDownList, site);
            WaitForLoad();

            _placeFrom = WaitForElementIsVisibleNew(By.CssSelector("#drop-down-places-from"));
            _placeFrom.SetValue(ControlType.DropDownList, placeFrom);

            _placeTo = WaitForElementIsVisibleNew(By.CssSelector("#drop-down-places-to"));
            _placeTo.SetValue(ControlType.DropDownList, placeTo);

            _activated = WaitForElementExists(By.CssSelector("#OutputForm_IsActive"));
            _activated.SetValue(ControlType.CheckBox, isActive);

            _createFrom = WaitForElementExists(By.CssSelector("#check-box-create-from"));
            _createFrom.SetValue(ControlType.CheckBox, isCreateFrom);

            _date = WaitForElementIsVisible(By.CssSelector("#datapicker-output-form-date"));
            _date.SetValue(ControlType.DateTime, date);
            var unselectDate = WaitForElementIsVisibleNew(By.CssSelector("#tb-new-outputform-number"));
            unselectDate.Click();

        }

        public string GetOutputFormNumber()
        {
            _outputFormNumber = WaitForElementIsVisibleNew(By.CssSelector("#tb-new-outputform-number"));
            return _outputFormNumber.GetAttribute("value");
        }

        public OutputFormItem Submit()
        {
            _submit = WaitForElementIsVisibleNew(By.CssSelector("#btn-submit-create-output-form"));
            _submit.Click();
            WaitForLoad();

            return new OutputFormItem(_webDriver, _testContext);
        }

        public string SelectSupplierOrder()
        {
            _tabSupplierOrder = WaitForElementIsVisibleNew(By.Id(TAB_SUPPLIER_ORDER));
            _tabSupplierOrder.Click();
            WaitPageLoading();
            Thread.Sleep(2000);

            _tabSupplierOrderCheckbox = WaitForElementExists(By.Id(TAB_SUPPLIER_ORDER_CHECKBOX));
            _tabSupplierOrderCheckbox.SetValue(PageBase.ControlType.CheckBox, true);

            // renvoir le Supplier Number
            return WaitForElementIsVisibleNew(By.XPath(TAB_SUPPLIER_ORDER_NUMBER)).Text.Trim();
        }

        public string SelectReceiptNote(string rnNumber)
        {
            _tabReceiptNote = WaitForElementIsVisible(By.Id(TAB_RECEIPT_NOTE));
            _tabReceiptNote.Click();
            WaitPageLoading();
            Thread.Sleep(2000);

            var inputRN = WaitForElementIsVisible(By.Id("form-copy-from-tbSearchName"));
            inputRN.SetValue(ControlType.TextBox, rnNumber);


            _tabReceiptNoteCheckbox = WaitForElementExists(By.Id(TAB_RECEIPT_NOTE_CHECKBOX));
            _tabReceiptNoteCheckbox.SetValue(PageBase.ControlType.CheckBox, true);

            // renvoir le Receipt Note Number
            return WaitForElementIsVisible(By.XPath(TAB_RECEIPT_NOTE_NUMBER)).Text.Trim();
        }

        public string SelectPurchaseOrder()
        {
            _tabPurchaseOrder = WaitForElementIsVisible(By.Id(TAB_PURCHASE_ORDER));
            _tabPurchaseOrder.Click();
            WaitPageLoading();
            Thread.Sleep(2000);

            _tabPurchaseOrderCheckbox = WaitForElementExists(By.Id(TAB_PURCHASE_ORDER_CHECKBOX));
            _tabPurchaseOrderCheckbox.SetValue(PageBase.ControlType.CheckBox, true);

            // renvoir le Purchase Order Number
            return WaitForElementIsVisible(By.XPath(TAB_PURCHASE_ORDER_NUMBER)).Text.Trim();
        }

        public string SelectOutputForm()
        {
            _tabOutputForm = WaitForElementIsVisible(By.Id(TAB_OUTPUT_FORM));
            _tabOutputForm.Click();
            WaitPageLoading();
            Thread.Sleep(2000);

            _tabOutputFormCheckbox = WaitForElementExists(By.Id(TAB_OUTPUT_FORM_CHECKBOX));
            _tabOutputFormCheckbox.SetValue(PageBase.ControlType.CheckBox, true);

            // renvoir le Output From Number (from)
            return WaitForElementIsVisible(By.XPath(TAB_OUTPUT_FORM_NUMBER)).Text.Trim();
        }
    }
}
