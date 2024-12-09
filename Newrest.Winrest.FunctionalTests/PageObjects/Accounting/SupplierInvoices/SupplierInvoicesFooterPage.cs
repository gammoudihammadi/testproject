using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.VariantTypes;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Purchasing;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Globalization;
using System.Linq;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Accounting.SupplierInvoices
{
    public class SupplierInvoicesFooterPage : PageBase
    {

        public SupplierInvoicesFooterPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // ________________________________________ Constantes ____________________________________________________

        // Onglets
        private const string ITEMS_TAB = "hrefTabContentItems";
        private const string ACCOUNTING_TAB = "hrefTabContentExportSageWriting";

        // Tableau
        private const string TOTAL_SUPPLIER_INVOICE = "//span[text()='TOTAL SUPPLIER INVOICE']/../..//span[contains(text(), '{0}')]";
        private const string TOTAL_PAYMENT = "//*[@id=\"tabContentDetails\"]/div/div[2]/div[2]";
        
        private const string NEW_TAX = "btn-add-tax";
        private const string NEW_TAX_TYPE = "drop-down-taxes";
        private const string NEW_TAX_AMOUNT = "FormTaxAmount";
        private const string NEW_TAX_ADD_BTN = "btn-save";

        private const string TAX_BASE_AMOUNT_LINE = "editable-row";
        private const string TAX_BASE_AMOUNT_AMOUNT = "//*[@id=\"tabContentDetails\"]/div/div/div[2]/div/div/div/form/div[2]/div[1]/span[text()='{0}']";
        private const string TAX_BASE_AMOUNT_DELETE = "//*[@id=\"tabContentDetails\"]/div/div/div[2]/div/div/div/form/div[2]/div[1]/span[text()='{0}']/../../div[3]/div/a/button";
        private const string TAX_AMOUNT = "taxLineVM_TotalBaseAmountDouble";

        private const string GROUP_LINE = "//*[@id=\"tabContentDetails\"]/div/table[2]/tbody/tr";
        private const string GROUP_TAX_AMOUNT = "//*[@id=\"tabContentDetails\"]/div/table[2]/tbody/tr[*]/td[text()='{0}']/parent::*/td[3]";
        private const string GROUP_VAT_AMOUNT = "//*[@id=\"tabContentDetails\"]/div/table[2]/tbody/tr[*]/td[text()='{0}']/parent::*/td[5]";
        private const string GROUP_NAME = "//*[@id=\"tabContentDetails\"]/div/table[2]/tbody/tr[*]/td[text()='{0}']/parent::*/td[1]";
        private const string VAT_AMOUNT_WITH_DEVIATION = "//*[@id=\"tabContentDetails\"]/div/table[2]/tbody/tr[3]/td[5]";
        private const string TOTAL_VAT_AMOUNT = "//*[@id=\"tabContentDetails\"]/div/table[1]/tbody/tr/td[4]";
        private const string DEVIATION_ROW = "//*[@id=\"tabContentDetails\"]/div/table[2]/tbody/tr[3]";

        // ________________________________________ Variables _____________________________________________________

        // Onglets
        [FindsBy(How = How.Id, Using = ITEMS_TAB)]
        private IWebElement _itemsTab;

        [FindsBy(How = How.Id, Using = ACCOUNTING_TAB)]
        private IWebElement _accountingTab;

        // Tableau
        [FindsBy(How = How.XPath, Using = TOTAL_SUPPLIER_INVOICE)]
        private IWebElement _totalSupplierInvoice;

        [FindsBy(How = How.Id, Using = NEW_TAX)]
        private IWebElement _newTax;

        [FindsBy(How = How.Id, Using = NEW_TAX_TYPE)]
        private IWebElement _newTaxType;

        [FindsBy(How = How.Id, Using = NEW_TAX_AMOUNT)]
        private IWebElement _newTaxAmount;

        [FindsBy(How = How.Id, Using = NEW_TAX_ADD_BTN)]
        private IWebElement _newTaxAddBtn;

        [FindsBy(How = How.Id, Using = TAX_AMOUNT)]
        private IWebElement _taxAmount;

        [FindsBy(How = How.XPath, Using = GROUP_NAME)]
        private IWebElement _groupName;
        [FindsBy(How = How.XPath, Using = TOTAL_VAT_AMOUNT)]
        private IWebElement _totalVatAmount;
        [FindsBy(How = How.XPath, Using = VAT_AMOUNT_WITH_DEVIATION)]
        private IWebElement _vatAmountWithDeviation;
        [FindsBy(How = How.XPath, Using = DEVIATION_ROW)]
        private IWebElement _deviation_Row;


        // ________________________________________ Méthodes ______________________________________________________

        // Onglets
        public SupplierInvoicesItem ClickOnItems()
        {
            _itemsTab = WaitForElementIsVisible(By.Id(ITEMS_TAB));
            _itemsTab.Click();
            WaitForLoad();

            return new SupplierInvoicesItem(_webDriver, _testContext);
        }

        public SupplierInvoicesAccounting ClickOnAccounting()
        {
            _accountingTab = WaitForElementIsVisible(By.Id(ACCOUNTING_TAB));
            _accountingTab.Click();
            WaitForLoad();

            return new SupplierInvoicesAccounting(_webDriver, _testContext);
        }

        // Tableau
        public double GetTotalSupplierInvoice(string currency, string decimalSeparator)
        {
            // Récupération du type de séparateur (, ou . selon les pays)
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            _totalSupplierInvoice = _webDriver.FindElement(By.XPath(String.Format(TOTAL_SUPPLIER_INVOICE, currency)));
            string value1 = _totalSupplierInvoice.Text.Replace(currency, "").Trim();
            return double.Parse(value1, ci);
        }
        public double GetTotalPayment(string currency, string decimalSeparator)
        {
            WaitPageLoading();
            // Récupération du type de séparateur (, ou . selon les pays)
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            _totalSupplierInvoice = _webDriver.FindElement(By.XPath(String.Format(TOTAL_PAYMENT, currency)));
            string value1 = _totalSupplierInvoice.Text.Replace(currency, "").Trim();
            return double.Parse(value1, ci);
        }
        public void AddNewTaxToSupplierInvoice(string taxType, string amount)
        {
            _newTax = WaitForElementIsVisible(By.Id(NEW_TAX));
            _newTax.Click();
            WaitForLoad();

            _newTaxType = WaitForElementIsVisible(By.Id(NEW_TAX_TYPE));
            _newTaxType.SetValue(ControlType.DropDownList, taxType);
            _newTaxType.SendKeys(Keys.Tab);

            _newTaxAmount = WaitForElementIsVisible(By.Id(NEW_TAX_AMOUNT));
            _newTaxAmount.SetValue(ControlType.TextBox, amount);

            _newTaxAddBtn = WaitForElementIsVisible(By.Id(NEW_TAX_ADD_BTN));
            _newTaxAddBtn.Click();
            WaitForLoad();
        }
        public bool IsAddNewTaxToSupplierInvoice(string taxType, string amount)
        {
            if (isElementVisible(By.XPath(NEW_TAX)))
            {
           
                _newTax = _webDriver.FindElement(By.Id(NEW_TAX));
                _newTax.Click();
                WaitForLoad();

                _newTaxType = WaitForElementIsVisible(By.Id(NEW_TAX_TYPE));
                _newTaxType.SetValue(ControlType.DropDownList, taxType);
                _newTaxType.SendKeys(Keys.Tab);

                _newTaxAmount = WaitForElementIsVisible(By.Id(NEW_TAX_AMOUNT));
                _newTaxAmount.SetValue(ControlType.TextBox, amount);

                _newTaxAddBtn = WaitForElementIsVisible(By.Id(NEW_TAX_ADD_BTN));
                _newTaxAddBtn.Click();
                WaitForLoad();
                return true;
            }
            else
            {
                return false;
            }
        }

        public int GetNumberOfTaxBaseAmount()
        {
            return _webDriver.FindElements(By.ClassName(TAX_BASE_AMOUNT_LINE)).Count();
        }

        public bool IsAmountPresent(string amount)
        {
            if(isElementVisible(By.XPath(String.Format(TAX_BASE_AMOUNT_AMOUNT, amount))))
            {
                _webDriver.FindElement(By.XPath(String.Format(TAX_BASE_AMOUNT_AMOUNT, amount)));
                return true;
            }
            else if (isElementVisible(By.XPath(String.Format("//*/input[@id='taxLineVM_TotalBaseAmountDouble'][@value='{0}']", amount.Substring(2)))))
            {
                // <input/>
                return true;
            }
            else
            {
                return false;
            }
        }

        public bool DeleteTaxInTaxAmountTable(string amount)
        {
            if(isElementVisible(By.XPath("//*[@id='btn_delete_manual_tax_1']/button"))) {
                var element = _webDriver.FindElement(By.XPath("//*[@id='btn_delete_manual_tax_1']/button"));
                element.Click();
                WaitForLoad();
                WaitPageLoading();

                return true;
            }
            if(isElementVisible(By.XPath(String.Format(TAX_BASE_AMOUNT_DELETE, amount))))
            {
                var element = _webDriver.FindElement(By.XPath(String.Format(TAX_BASE_AMOUNT_DELETE, amount)));
                element.Click();
                WaitForLoad();
                WaitPageLoading();

                return true;
            }
            else
            {
                return false;   
            }
        }

        public int GetNumberOfGroup()
        {
            return _webDriver.FindElements(By.XPath(GROUP_LINE)).Count();
        }

        public string GetGroupNameForTaxName(string taxType)
        {
            _groupName = _webDriver.FindElement(By.XPath(String.Format(GROUP_NAME, taxType)));
            return _groupName.Text;
        }

        public double GetTaxBaseAmountForGroup(string group, string currency, string decimalSeparator)
        {
            // Récupération du type de séparateur (, ou . selon les pays)
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            var element = _webDriver.FindElement(By.XPath(String.Format(GROUP_TAX_AMOUNT, group)));
            string taxAmount = element.Text.Replace(currency, "").Trim();

            return double.Parse(taxAmount, ci);
        }

        public double GetVatAmountForGroup(string group, string currency, string decimalSeparator)
        {
            // Récupération du type de séparateur (, ou . selon les pays)
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            var element = _webDriver.FindElement(By.XPath(String.Format(GROUP_VAT_AMOUNT, group)));
            string vatAmount = element.Text.Replace(currency, "").Trim();

            return double.Parse(vatAmount, ci);
        }

        public void ModifyAmount(string initAmount, string newAmount)
        {
            var amountToModify = WaitForElementIsVisible(By.Id("taxLineVM_TotalBaseAmountDouble"));
            amountToModify.SetValue(ControlType.TextBox, newAmount);
            amountToModify.Submit();
            WaitForLoad();
        }

        public bool IsModifyAmount(string initAmount, string newAmount)
        {
            if (isElementVisible(By.XPath(String.Format(TAX_BASE_AMOUNT_AMOUNT, initAmount))))
            {
                var amountToModify = _webDriver.FindElement(By.XPath(String.Format(TAX_BASE_AMOUNT_AMOUNT, initAmount)));
                amountToModify.Click();
                WaitForLoad();

                if (!isElementVisible(By.Id(TAX_AMOUNT)))
                {
                    return false;
                }

                _taxAmount = _webDriver.FindElement(By.Id(TAX_AMOUNT));
                _taxAmount.SetValue(ControlType.TextBox, newAmount);
                _taxAmount.Submit();
                WaitForLoad();

                return true;
            }
            else
            {
                return false;
            }
        }
        public void ModifyTAXBaseAmount(string initAmount, string newAmount)
        {
            var amountToModify = WaitForElementIsVisible(By.Id("taxLineVM_TotalBaseAmountDouble"));
            amountToModify.SetValue(ControlType.TextBox, newAmount);
            WaitForLoad();
            WaitPageLoading();
        }
        public double GetTotalVATAmount(string currency, string decimalSeparator)
        {
            WaitPageLoading();
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");
            _totalVatAmount = _webDriver.FindElement(By.XPath(String.Format(TOTAL_VAT_AMOUNT, currency)));
            string value1 = _totalVatAmount.Text.Replace(currency, "").Trim();
            return double.Parse(value1, ci);
        }
        public double GetVATAmountDeviation(string currency, string decimalSeparator)
        {
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");
            WaitForLoad();
            _vatAmountWithDeviation = _webDriver.FindElement(By.XPath(String.Format(VAT_AMOUNT_WITH_DEVIATION, currency)));
            string value1 = _vatAmountWithDeviation.Text.Replace(currency, "").Trim();
            return double.Parse(value1, ci);
        }
        public bool IsDeviationRowDisplayed()
        {
            var row = _webDriver.FindElement(By.XPath("//*[@id='tabContentDetails']/div/table[2]/tbody/tr[3]"));
            return row.Text.Contains("DEVIATION");
        }
        public string GetVatRate()
        {
            WaitLoading();
            var vatRateElement = _webDriver.FindElement(By.XPath("//*[@id='tabContentDetails']/div/div[1]/div[2]/div/div/div/form/div[1]/div[2]"));
            return vatRateElement.Text;
        }

    }
}
