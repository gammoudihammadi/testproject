using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Globalization;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Accounting.Invoice
{
    public class InvoiceFooterPage : PageBase
    {
        public InvoiceFooterPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // _____________________________________________Constantes______________________________________________

        // Général
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";

        // Onglets
        private const string ITEMS_TAB = "hrefTabContentItems";
        private const string ACCOUNTING_TAB = "hrefTabContentExportSageWriting";
        private const string GENERAL_INFORMATION_TAB = "hrefTabContentInformations";

        // Tableau
        private const string TOTAL_GROSS_AMOUNT = "//*[@id=\"tabContentDetails\"]/div/table[1]/tbody/tr[*]/td[contains(text(), 'TOTAL GROSS AMOUNT')]/../td[2]";
        private const string TOTAL_TTC = "//*[@id=\"tabContentDetails\"]/div/table[1]/tbody/tr[*]/td[1][text()='TOTAL INVOICE']/../td[4]";
        private const string LOCAL_CURRENCY = "//*[@id=\"tabContentDetails\"]/div/table[2]/tbody/tr[*]/td[1][contains(text(),'Local currency')]/../td[2]";
        private const string CUSTOMER_CURRENCY = "//*[@id=\"tabContentDetails\"]/div/table[2]/tbody/tr[*]/td[1][contains(text(),'Customer currency')]/../td[2]";
        private const string AIRPORT_TAX_CONTENT = "//*[@id=\"tabContentDetails\"]/div/table[1]/tbody/tr[*]/td[1][text()='Airport Tax']/../td[4]";
        private const string TOTAL_INCL_TAXES = "/html/body/div[3]/div/div[3]/div[1]/div/div/div/div[1]/div[2]";
        private const string EXCHANGE_RATE = "/html/body/div[3]/div/div[3]/div[1]/div/div/div/div[2]/div[2]";
        private const string LOCAL_TOTAL_INCL_TAXES = "/html/body/div[3]/div/div[3]/div[1]/div/div/div/div[3]/div[2]";
        private const string TAX_NAME = "/html/body/div[3]/div/div[3]/div[1]/div/div/table[1]/tbody/tr[2]/td[1]";

        // _____________________________________________Variables_______________________________________________

        // Général
        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;

        // Onglets
        [FindsBy(How = How.Id, Using = ITEMS_TAB)]
        private IWebElement _itemsTab;

        [FindsBy(How = How.Id, Using = ACCOUNTING_TAB)]
        private IWebElement _accountingTab;

        [FindsBy(How = How.Id, Using = GENERAL_INFORMATION_TAB)]
        private IWebElement _generalInfoTab;

        // Tableau
        [FindsBy(How = How.XPath, Using = TOTAL_GROSS_AMOUNT)]
        private IWebElement _totalGrossAMount;

        [FindsBy(How = How.XPath, Using = TOTAL_TTC)]
        private IWebElement _totalTTC;

        [FindsBy(How = How.XPath, Using = LOCAL_CURRENCY)]
        private IWebElement _localCurrency;

        [FindsBy(How = How.XPath, Using = CUSTOMER_CURRENCY)]
        private IWebElement _customerCurrency;

        [FindsBy(How = How.XPath, Using = AIRPORT_TAX_CONTENT)]
        private IWebElement _airportTax;

        // _____________________________________________Pages___________________________________________________

        // Général
        public InvoicesPage BackToList()
        {
            _backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            _backToList.Click();
            WaitForLoad();

            return new InvoicesPage(_webDriver, _testContext);
        }

        // Onglets
        public InvoiceDetailsPage ClickOnItems()
        {
            _itemsTab = WaitForElementIsVisible(By.Id(ITEMS_TAB));
            _itemsTab.Click();
            WaitForLoad();

            return new InvoiceDetailsPage(_webDriver, _testContext);
        }

        public InvoiceAccountingPage ClickOnAccounting()
        {
            _accountingTab = WaitForElementIsVisible(By.Id(ACCOUNTING_TAB));
            _accountingTab.Click();
            WaitForLoad();

            return new InvoiceAccountingPage(_webDriver, _testContext);
        }

        public InvoiceGeneralInformations ClickOnGeneralInformation()
        {
            _generalInfoTab = WaitForElementIsVisible(By.Id(GENERAL_INFORMATION_TAB));
            _generalInfoTab.Click();
            WaitForLoad();

            // Tableau

            return new InvoiceGeneralInformations(_webDriver, _testContext);
        }
        public string GetTotalGrossAmount(string currency)
        {
            _totalGrossAMount = WaitForElementIsVisible(By.XPath(TOTAL_GROSS_AMOUNT));
            return _totalGrossAMount.Text.Replace(currency, "").Trim();
        }

        public double GetTotalTTC(string currency, string decimalSeparatorValue)
        {
            // Récupération du type de séparateur (, ou . selon les pays)
            CultureInfo ci = decimalSeparatorValue.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            _totalTTC = WaitForElementIsVisible(By.XPath(TOTAL_TTC));
            return Convert.ToDouble(_totalTTC.Text.Replace(currency, ""), ci);
        }

        public string GetTotalTTC(string currency)
        {
            _totalTTC = WaitForElementIsVisible(By.XPath(TOTAL_TTC));
            return _totalTTC.Text.Replace(currency, "").Trim();
        }

        public string GetLocalCurrency()
        {
            _localCurrency = WaitForElementIsVisible(By.XPath(LOCAL_CURRENCY));
            return _localCurrency.Text;
        }

        public string GetCustomerCurrency()
        {
            _customerCurrency = WaitForElementIsVisible(By.XPath(CUSTOMER_CURRENCY));
            return _customerCurrency.Text;
        }

        public bool IsAirportTaxPresent(string currency)
        {
            try
            {
                _airportTax = _webDriver.FindElement(By.XPath(AIRPORT_TAX_CONTENT));

                if (_airportTax.Text.Replace(currency, "").Trim() == "0")
                {
                    return false;
                }
            }
            catch
            {
                return false;
            }

            return true;
        }

        public string[] GetColumnVATrates()
        {
            string rates = "//*[@id='tabContentDetails']/div/table[1]/tbody/*/td[3]";
            var lignes = _webDriver.FindElements(By.XPath(rates));
            Assert.AreEqual(5, lignes.Count, "pas de ligne TVA");
            string[] result = new string[3];
            for (var x = 0; x < 3; x++)
            {
                result[x] = lignes[x].Text.Replace(" ", "");
            }
            return result;
        }

        public bool CheckVAT(int nbVAT)
        {
            var _nbVAT = _webDriver.FindElements(By.XPath("//*/tr[@class='first-line']/parent::tbody/tr[*]"));
            return _nbVAT.Count - 4 == nbVAT;
        }

        public int TaxNameTotalNumber()
        {
            return _webDriver.FindElements(By.XPath("//*[@id=\"tabContentDetails\"]/div/table[1]/tbody/tr[not(td[1][contains(text(), 'Total') or contains(text(), 'TOTAL INVOICE') or contains(text(), 'TOTAL GROSS AMOUNT')])]/td[1]")).Count;
        }

        public bool VerifyTotalInvoice()
        {
            var totalInclTaxe = WaitForElementIsVisible(By.XPath(TOTAL_INCL_TAXES));          
            var ExchangeRate = WaitForElementIsVisible(By.XPath(EXCHANGE_RATE));
            var localtotalInclTaxe = WaitForElementIsVisible(By.XPath(LOCAL_TOTAL_INCL_TAXES));
            // get value 
            string numbertotalInclTaxe = (totalInclTaxe.Text.Replace("$", "").Trim().Replace(",", "."));
            string numberLocaltotal = (localtotalInclTaxe.Text.Replace("€", "").Trim().Replace(",", "."));
            //convert string to double 
            double.TryParse(numbertotalInclTaxe, NumberStyles.Any, CultureInfo.InvariantCulture, out double _numbertotalInclTaxe);
            double.TryParse(ExchangeRate.Text, NumberStyles.Any, new CultureInfo("fr-FR"), out double _ExchangeRate); 
            double.TryParse(numberLocaltotal, NumberStyles.Any, CultureInfo.InvariantCulture, out double _numberLocaltotal);

            return _numberLocaltotal.Equals(_numbertotalInclTaxe * _ExchangeRate);
        }
        public string getTaxName()
        {
            var taxName = WaitForElementIsVisible(By.XPath(TAX_NAME));
            return taxName.Text;
        }
    }
}
