using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Globalization;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.ReceiptNotes
{
    public class ReceiptNotesFooterPage : PageBase
    {
        public ReceiptNotesFooterPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // _________________________________________ Constantes _______________________________________________

        // General
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";

        // Onglets
        private const string ITEMS_TAB = "hrefTabContentItems";
        private const string ACCOUNTING_TAB = "hrefTabContentExportSageWriting";

        // Tableau
        private const string TOTAL_RECEIPT_NOTE = "//*[@id=\"tabContentDetails\"]/div/table[1]/tbody/tr[5]/td[4]";

        // _________________________________________ Variables ________________________________________________

        // General
        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;

        // Onglets
        [FindsBy(How = How.Id, Using = ITEMS_TAB)]
        private IWebElement _itemsTab;

        [FindsBy(How = How.Id, Using = ACCOUNTING_TAB)]
        private IWebElement _accountingTab;

        // Tableau
        [FindsBy(How = How.Id, Using = TOTAL_RECEIPT_NOTE)]
        private IWebElement _totalReceiptNote;

        // _________________________________________ Méthodes _________________________________________________

        // General
        public ReceiptNotesPage BackToList()
        {
            _backToList = WaitForElementToBeClickable(By.XPath(BACK_TO_LIST));
            _backToList.Click();
            WaitForLoad();

            return new ReceiptNotesPage(_webDriver, _testContext);
        }

        // Onglets
        public ReceiptNotesItem ClickOnItemsTab()
        {
            _itemsTab = WaitForElementIsVisible(By.Id(ITEMS_TAB));
            _itemsTab.Click();
            WaitForLoad();

            return new ReceiptNotesItem(_webDriver, _testContext);
        }

        public ReceiptNotesAccountingPage ClickOnAccounting()
        {
            _accountingTab = WaitForElementIsVisible(By.Id(ACCOUNTING_TAB));
            _accountingTab.Click();
            WaitForLoad();

            return new ReceiptNotesAccountingPage(_webDriver, _testContext);
        }

        // Tableau
        public double GetReceiptNoteTotal(string currency, string decimalSeparatorValue)
        {
            CultureInfo ci = decimalSeparatorValue.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            _totalReceiptNote = WaitForElementIsVisible(By.XPath(TOTAL_RECEIPT_NOTE));
            string montant = _totalReceiptNote.Text.Replace(currency, "").Trim();

            return double.Parse(montant, ci);
        }

        public string GetTaxName()
        {
            var _taxName = WaitForElementIsVisible(By.XPath("//*/table[contains(@class,'footer-table')]/tbody/tr[2]/td[1]"));
            return _taxName.Text;
        }

        public double GetTaxBaseAmount(string decimalSeparator)
        {
            var _taxBaseAmount = WaitForElementIsVisible(By.XPath("//*/table[contains(@class,'footer-table')]/tbody/tr[2]/td[2]"));
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");
            return double.Parse(_taxBaseAmount.Text.Replace("€", "").Replace(" ", ""), ci);
        }

        public double GetVATRate(string decimalSeparator)
        {
            var _VATRate = WaitForElementIsVisible(By.XPath("//*/table[contains(@class,'footer-table')]/tbody/tr[2]/td[3]"));
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");
            return double.Parse(_VATRate.Text.Replace("%", "").Replace(" ", ""), ci);
        }

        public double GetVATAmount(string decimalSeparator)
        {
            var _VATAmount = WaitForElementIsVisible(By.XPath("//*/table[contains(@class,'footer-table')]/tbody/tr[2]/td[4]"));
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");
            return double.Parse(_VATAmount.Text.Replace("€", "").Replace(" ", ""), ci);
        }

        public double GetTotalReceiptNote(string decimalSeparator)
        {
            var _totalReceiptNote = WaitForElementIsVisible(By.XPath("//*/table[contains(@class,'footer-table')]/tbody/tr[5]/td[4]"));
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");
            return double.Parse(_totalReceiptNote.Text.Replace("€", "").Replace(" ", ""), ci);
        }
        public void Go_To_New_Navigate()
        {
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);
        }
    }
}
