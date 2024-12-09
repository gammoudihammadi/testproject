using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Globalization;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.OutputForm
{
    public class OutputFormFooterPage : PageBase
    {
        public OutputFormFooterPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // _________________________________________ Constantes _______________________________________________

        // Onglets
        private const string ITEMS_TAB = "hrefTabContentItems";
        private const string ACCOUNTING_TAB = "hrefTabContentExportSageWriting";

        // Tableau
        private const string TOTAL_OUTPUT_FORM = "//*[@id=\"tabContentDetails\"]/div/table[1]/tbody/tr[5]/td[4]";//*[@id="tabContentDetails"]/div/table[1]/tbody/tr[5]/td[4]

        // _________________________________________ Variables ________________________________________________

        // Onglets
        [FindsBy(How = How.Id, Using = ITEMS_TAB)]
        private IWebElement _itemsTab;

        [FindsBy(How = How.Id, Using = ACCOUNTING_TAB)]
        private IWebElement _accountingTab;

        // Tableau
        [FindsBy(How = How.XPath, Using = TOTAL_OUTPUT_FORM)]
        private IWebElement _totalOutputForm;

        // _________________________________________ Méthodes _________________________________________________

        // Onglets
        public OutputFormItem ClickOnItemsTab()
        {
            _itemsTab = WaitForElementToBeClickable(By.Id(ITEMS_TAB));
            _itemsTab.Click();
            WaitForLoad();

            return new OutputFormItem(_webDriver, _testContext);
        }

        public OutputFormAccountingPage ClickOnAccountingTab()
        {
            _accountingTab = WaitForElementIsVisible(By.Id(ACCOUNTING_TAB));
            _accountingTab.Click();
            WaitForLoad();

            return new OutputFormAccountingPage(_webDriver, _testContext);
        }

        // Tableau
        public double GetOutputFormTotalHT(string currency, string decimalSeparator)
        {
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            if(isElementVisible(By.XPath(TOTAL_OUTPUT_FORM)))
            {
                _totalOutputForm = WaitForElementIsVisible(By.XPath(TOTAL_OUTPUT_FORM));
            }
            else
            {
                //PATCH 24Aug
                _totalOutputForm = WaitForElementIsVisible(By.XPath("//*[@id=\"tabContentDetails\"]/div/table[1]/tbody/tr[3]/td[2]"));
            }

            return Math.Round(double.Parse(_totalOutputForm.Text.Replace(currency, "").Trim(), ci), 2);
        }
    }
}
