using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.Inventory
{
    public class InventoryFooterPage : PageBase
    {

        public InventoryFooterPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // _________________________________________ Constantes _______________________________________________

        // Onglets
        private const string ACCOUNTING_TAB = "hrefTabContentExportSageWriting";

        // Tableau
        private const string TOTAL_INVENTORY = "//*[@id=\"tabContentDetails\"]/div/table[1]/tbody/tr[*]/td[1][text()='TOTAL GROSS AMOUNT']/parent::*/td[2]";

        // _________________________________________ Variables ________________________________________________

        // Onglets
        [FindsBy(How = How.Id, Using = ACCOUNTING_TAB)]
        private IWebElement _accountingTab;

        // Tableau
        [FindsBy(How = How.XPath, Using = TOTAL_INVENTORY)]
        private IWebElement _totalInventory;

        // _________________________________________ Méthodes _________________________________________________

        // Onglets
        public InventoryAccountingPage ClickOnAccounting()
        {
            _accountingTab = WaitForElementIsVisible(By.Id(ACCOUNTING_TAB));
            _accountingTab.Click();
            WaitForLoad();

            return new InventoryAccountingPage(_webDriver, _testContext);
        }

        // Tableau
        public double GetInventoryTotalHT(string currency, string decimalSeparator)
        {
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            _totalInventory = WaitForElementIsVisible(By.XPath(TOTAL_INVENTORY));
            string montant = _totalInventory.Text.Replace(currency, "").Trim();

            return double.Parse(montant, ci);
        }
    }
}
