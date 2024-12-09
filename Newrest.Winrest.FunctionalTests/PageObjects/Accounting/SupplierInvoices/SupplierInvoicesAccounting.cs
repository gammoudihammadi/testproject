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

namespace Newrest.Winrest.FunctionalTests.PageObjects.Accounting.SupplierInvoices
{
    public class SupplierInvoicesAccounting : PageBase
    {
        public SupplierInvoicesAccounting(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // ________________________________________ Constantes ____________________________________________________

        // Général
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";

        // Onglets
        private const string ITEMS_TAB = "hrefTabContentItems";

        // Tableau
        private const string ERROR_MESSAGE = "//*[@id=\"tabContentDetails\"]/div/p[2]";
        private const string AMOUNT_CREDITIN_LOCAL = "//*[@id=\"tabContentDetails\"]/table/tbody/tr[1]/td[10]";

        private const string AMOUNT_DEBIT = "//*[@id=\"tabContentDetails\"]/table/tbody/tr[*]/td[14][contains(text(),'{0}')]/../td[9]";
        private const string AMOUNT_CREDIT = "//*[@id=\"tabContentDetails\"]/table/tbody/tr[*]/td[14][contains(text(),'{0}')]/../td[10]";
        private const string SPECIFIC_INVOICE_DATE = "//*[@id=\"tabContentDetails\"]/table/tbody/tr[*]/td[14][contains(text(),'{0}')]/../td[21]";

        
        // ________________________________________ Variables _____________________________________________________

        // Général
        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;

        // Onglets
        [FindsBy(How = How.Id, Using = ITEMS_TAB)]
        private IWebElement _itemsTab;

        // Tableau
        [FindsBy(How = How.XPath, Using = ERROR_MESSAGE)]
        private IWebElement _errorMessage;

        // ________________________________________ Méthodes_______________________________________________________

        // Général
        public SupplierInvoicesPage BackToList()
        {
            _backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
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
        public string GetErrorMessage()
        {
            if (isElementVisible(By.XPath(ERROR_MESSAGE)))
            {
                _errorMessage = _webDriver.FindElement(By.XPath(ERROR_MESSAGE));
                return _errorMessage.Text.Trim();
            }
            else
            {
                return "";
            }
        }

        public double GetInvoiceGrossAmount(string writingType, string decimalSeparator)
        {
            // Récupération du type de séparateur (, ou . selon les pays)
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            var amountDebitList = _webDriver.FindElements(By.XPath(string.Format(AMOUNT_DEBIT, writingType)));

            if (amountDebitList.Count == 0)
                return 0;

            double amountDebit = 0;

            foreach (var elm in amountDebitList)
            {
                amountDebit += double.Parse(elm.Text.Trim(), ci);
            }

            return amountDebit;
        }

        public double GetInvoiceDetailAmount(string writingType, string decimalSeparator)
        {
            // Récupération du type de séparateur (, ou . selon les pays)
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            var amountCreditList = _webDriver.FindElements(By.XPath(string.Format(AMOUNT_CREDIT, writingType)));

            if (amountCreditList.Count == 0)
                return 0;

            double amountCredit = 0;

            foreach (var elm in amountCreditList)
            {
                amountCredit += double.Parse(elm.Text.Trim(), ci);
            }

            return amountCredit;
        }

        public List<DateTime> GetInvoiceIntegrationDate(string writingType, string dateFormatPicker)
        {
            HashSet<DateTime> dates = new HashSet<DateTime>();

            // Récupération du type de séparateur (, ou . selon les pays)
            CultureInfo ci = dateFormatPicker.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            var integrationDates = _webDriver.FindElements(By.XPath(string.Format(SPECIFIC_INVOICE_DATE, writingType)));

            foreach (var elm in integrationDates)
            {
                dates.Add(DateTime.Parse(elm.Text.Trim(), ci).Date);
            }

            return new List<DateTime>(dates);
        }
        public double GetTotalAmountCreditInLocal(string decimalSeparator)
        {
            WaitPageLoading();
            // Logic to extract the sum of the AmountCreditInLocal column
            var amountCreditInLocalElements = _webDriver.FindElements(By.XPath(AMOUNT_CREDITIN_LOCAL));
            double totalAmountCreditInLocal = 0;

            foreach (var element in amountCreditInLocalElements)
            {
                string amountText = element.Text.Replace(decimalSeparator, ".");
                totalAmountCreditInLocal += Convert.ToDouble(amountText);
            }

            return totalAmountCreditInLocal;
        }


    }
}
