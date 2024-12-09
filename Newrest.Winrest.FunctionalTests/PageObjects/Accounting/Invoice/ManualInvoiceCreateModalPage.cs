using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Accounting.Invoice
{
    public class ManualInvoiceCreateModalPage : PageBase
    {
        public ManualInvoiceCreateModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // _____________________________________ Constantes __________________________________________________

        private const string SITE = "dropdown-site";
        private const string INVOICE_DATE = "InvoiceDate";
        private const string CUSTOMER = "//*[@id=\"createFormVipInvoice\"]/div[4]/div/span[1]/input[2]";
        private const string CUSTOMER_SELECTED = "//*[@id=\"createFormVipInvoice\"]/div[4]/div/span[1]/div/div/div[contains(text(), '{0}')]";
        private const string APPLY_VAT = "//*[@id=\"createFormVipInvoice\"]/div[5]/div/div[1]";
        private const string SUBMIT = "btn-submit-form-create-vip-order";
        private const string CUSTOMER_INPUT = "//*[@id=\"createFormVipInvoice\"]/div[4]/div/span[1]/input[2]";


        // _____________________________________ Variables ___________________________________________________

        [FindsBy(How = How.Id, Using = SITE)]
        private IWebElement _site;

        [FindsBy(How = How.Id, Using = INVOICE_DATE)]
        private IWebElement _invoiceDate;

        [FindsBy(How = How.XPath, Using = CUSTOMER)]
        private IWebElement _customer;

        [FindsBy(How = How.XPath, Using = CUSTOMER_SELECTED)]
        private IWebElement _customerSelected;

        [FindsBy(How = How.XPath, Using = APPLY_VAT)]
        private IWebElement _applyVAT;

        [FindsBy(How = How.Id, Using = SUBMIT)]
        private IWebElement _submit;

        // ______________________________________ Methodes ____________________________________________________

        public InvoiceDetailsPage FillField_CreatNewManualInvoice(string site, DateTime date, string customer, bool isVATForced)
        {
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            
            _site = WaitForElementIsVisible(By.Id(SITE));
            _site.SetValue(ControlType.DropDownList, site);

            _invoiceDate = WaitForElementIsVisible(By.Id(INVOICE_DATE));
            _invoiceDate.SetValue(ControlType.DateTime, date);
            _invoiceDate.SendKeys(Keys.Tab);

            _customer = WaitForElementIsVisible(By.XPath(CUSTOMER));
            _customer.SetValue(ControlType.TextBox, customer);
            WaitForLoad();

            // Selection du premier élément de la liste (customer utilisé pour les flights)
            _customerSelected = WaitForElementIsVisible(By.XPath(string.Format(CUSTOMER_SELECTED, customer)));
            WaitPageLoading();
            _customerSelected.Click();
            WaitPageLoading();

            _applyVAT = WaitForElementIsVisible(By.XPath(APPLY_VAT));
            _applyVAT.SetValue(ControlType.CheckBox, isVATForced);
            WaitForLoad();

            _submit = WaitForElementExists(By.Id(SUBMIT));
            javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", _submit);
            _submit = WaitForElementExists(By.Id(SUBMIT));
            _submit.Click();
            WaitPageLoading();

            return new InvoiceDetailsPage(_webDriver, _testContext);
        }

        public bool CanCreate()
        {
            _submit = WaitForElementExists(By.Id(SUBMIT));
            return _submit.GetAttribute("disabled") == null;
        }

        public InvoiceDetailsPage Create()
        {
            var customerInput = WaitForElementIsVisible(By.XPath(CUSTOMER_INPUT));
            customerInput.Submit();
            return new InvoiceDetailsPage(_webDriver, _testContext);
        }
    }
}
