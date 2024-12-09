using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml.Linq;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Purchasing
{
    public class ParametersPurchasing : PageBase
    {
        public ParametersPurchasing(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {

        }

        // ____________________________________________ Constantes ____________________________________________

        // VAT
        private const string VAT_TAB = "//*[@id=\"paramPurchasingTab\"]/li[2]";
        private const string SUPPLIER_TAB = "//*/a[text()='Supplier type']";

        private const string ADD_VAT = "//*[@id=\"tabContentParameters\"]/div[1]/a";
        private const string VAT_NAME = "first";
        private const string VAT_TYPE = "TaxType_KindOfTax";
        private const string VAT_VALUE = "//*[@id=\"TaxTypeModal\"]/div/div/div/div/form/div[2]/div[*]/div/input[3]";
        private const string PURCHASE_CODE = "//*[@id=\"TaxTypeModal\"]/div/div/div/div/form/div[2]/div[*]/label[contains(text(), 'Purchase code')]/../div/input";
        private const string PURCHASE_ACCOUNT = "//*[@id=\"TaxTypeModal\"]/div/div/div/div/form/div[2]/div[*]/label[contains(text(), 'Purchase account')]/../div/input";
        private const string SALES_CODE = "//*[@id=\"TaxTypeModal\"]/div/div/div/div/form/div[2]/div[*]/label[contains(text(), 'Sales code')]/../div/input";
        private const string SALES_ACCOUNT = "//*[@id=\"TaxTypeModal\"]/div/div/div/div/form/div[2]/div[*]/label[contains(text(), 'Sales account')]/../div/input";
        private const string DUE_TO_INVOICE_CODE = "//*[@id=\"TaxTypeModal\"]/div/div/div/div/form/div[2]/div[*]/label[contains(text(), 'Due to invoice code')]/../div/input";
        private const string DUE_TO_INVOICE_ACCOUNT = "//*[@id=\"TaxTypeModal\"]/div/div/div/div/form/div[2]/div[*]/label[contains(text(), 'Due to invoice account')]/../div/input";
        private const string SAVE = "last";

        private const string SPAIN_ID = "//*[@id=\"TaxTypeModal\"]/div/div/div/div/form/div[2]/h4[contains(text(), 'Spain')]/following-sibling::div/label";

        private const string VAT_LINES = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]";

        private const string VAT_BIS = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[(contains(td[1], '{0}') and contains(td[2], '{1}'))]";
        private const string VAT_EDIT = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[(contains(td[1], '{0}') and contains(td[2], '{1}'))]/td[4]/a[2]";
        private const string CURRENCY_DOLAR = "/html/body/div[2]/div/div/div/table/tbody/tr[7]";
        private const string EDIT_CURRENCY = "/html/body/div[2]/div/div/div/table/tbody/tr[7]/td[10]/a[1]";
        private const string CURRENCY_COEFF = "CoefficientToOfficialCurrency";

        // ______________________________________________ Variables ____________________________________________

        // VAT
        [FindsBy(How = How.XPath, Using = VAT_TAB)]
        private IWebElement _vatTab;

        [FindsBy(How = How.XPath, Using = ADD_VAT)]
        private IWebElement _addVat;

        [FindsBy(How = How.Id, Using = VAT_NAME)]
        private IWebElement _vatName;

        [FindsBy(How = How.Id, Using = VAT_TYPE)]
        private IWebElement _vatType;

        [FindsBy(How = How.Id, Using = SAVE)]
        private IWebElement _save;

        //SUPPLIER
        [FindsBy(How = How.XPath, Using = SUPPLIER_TAB)]
        private IWebElement _supplierTab;

        // _______________________________________________ Méthodes _____________________________________________

        // VAT 
        public void GoToTab_VAT()
        {
            if (IsDev())
            {
                _vatTab = WaitForElementIsVisible(By.XPath("//*[@id='paramPurchasingTab']/li/a[text()='VAT']"));
            }
            else
            {
                _vatTab = WaitForElementIsVisible(By.XPath(VAT_TAB));
            }
            _vatTab.Click();
            WaitForLoad();
        }

        // SUPPLIER TAB
        public void GoToTab_Supplier()
        {
            _supplierTab = WaitForElementIsVisible(By.XPath(SUPPLIER_TAB));
            _supplierTab.Click();
            WaitForLoad();
        }

        public bool IsTaxPresent(string taxName, string taxType)
        {
            try
            {
                _webDriver.FindElement(By.XPath(string.Format(VAT_EDIT, taxName, taxType)));
                return true;
            }
            catch
            {
                return false;
            }
        }

        public void SearchVAT(string taxName, string taxType)
        {
            var editVAT = WaitForElementIsVisible(By.XPath(String.Format(VAT_EDIT, taxName, taxType)));
            editVAT.Click();
            WaitForLoad();
        }

        public void CreateNewVAT(string taxName, string taxType, string taxValue)
        {
            Actions action = new Actions(_webDriver);

            _addVat = WaitForElementIsVisible(By.XPath(ADD_VAT));
            _addVat.Click();
            WaitForLoad();

            _vatName = WaitForElementIsVisible(By.Id(VAT_NAME));
            _vatName.SetValue(ControlType.TextBox, taxName);

            _vatType = WaitForElementIsVisible(By.Id(VAT_TYPE));
            _vatType.SetValue(ControlType.DropDownList, taxType);

            var vatValues = _webDriver.FindElements(By.XPath(VAT_VALUE));

            foreach (var elm in vatValues)
            {
                action.MoveToElement(elm).Perform();
                elm.SetValue(ControlType.TextBox, taxValue);
            }

            _save = WaitForElementExists(By.Id(SAVE));
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", _save);
            WaitPageLoading();

            _save.Click();
            WaitForLoad();
        }

        public void EditVATAccountForSpain(string purchaseCode = null, string purchaseAccount = null, string salesCode = null, string salesAccount = null, string RNCode = null, string RNAccount = null)
        {
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;

            var spainId = _webDriver.FindElement(By.XPath(SPAIN_ID)).GetAttribute("for");
            var spainIdNumber = Regex.Match(spainId, @"\d+").Value;


            if (purchaseCode != null)
            {
                var purchaseCodeValue = _webDriver.FindElement(By.Id("CountryValues_" + spainIdNumber + "__PurchaseCode"));
                purchaseCodeValue.SetValue(ControlType.TextBox, purchaseCode);
            }

            if (purchaseAccount != null)
            {
                var purchaseAccountValue = _webDriver.FindElement(By.Id("CountryValues_" + spainIdNumber + "__AccountingCode"));
                purchaseAccountValue.SetValue(ControlType.TextBox, purchaseAccount);
            }

            if (salesCode != null)
            {
                var salesCodeValue = _webDriver.FindElement(By.Id("CountryValues_" + spainIdNumber + "__SalesCode"));
                salesCodeValue.SetValue(ControlType.TextBox, salesCode);
            }

            if (salesAccount != null)
            {
                var salesAccountValue = _webDriver.FindElement(By.Id("CountryValues_" + spainIdNumber + "__SalesAccountingCode"));
                salesAccountValue.SetValue(ControlType.TextBox, salesAccount);
            }

            if (RNCode != null)
            {
                var RNCodeValue = _webDriver.FindElement(By.Id("CountryValues_" + spainIdNumber + "__DueToInvoiceCode"));
                RNCodeValue.SetValue(ControlType.TextBox, RNCode);
            }

            if (RNAccount != null)
            {
                var RNAccountValue = _webDriver.FindElement(By.Id("CountryValues_" + spainIdNumber + "__DueToInvoiceAccountingCode"));
                RNAccountValue.SetValue(ControlType.TextBox, RNAccount);
            }
            _save = WaitForElementExists(By.Id(SAVE));

            javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", _save);

            _save.Click();

            //Jusqu'à 1min30 pour sauver (Cécile), on contourne les lenteurs en attendant que la modal disparaisse
            WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(130));
            wait.Until(ExpectedConditions.InvisibilityOfElementLocated(By.Id(SAVE)));
            wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(3));

            WaitForLoad();
        }
        public string GetTaxValueByCountry(string country)
        {
            var el = WaitForElementExists(By.XPath(string.Format("//*[@id=\"TaxTypeModal\"]/div/div/div/div/form/div[2]/h4[contains(text(),'{0}')]//following-sibling::div/div//*[@id=\"CountryValues_2__Value\"]", country)));
            return el.GetAttribute("value");
        }
        public bool CheckForExistingType(string supplierType)
        {
            var existingTypes = _webDriver.FindElements(By.XPath("//*/tr[*]/td[1]"));
            bool IsExistingType = false;
            foreach (var type in existingTypes)
            {
                if (type.Text.Contains(supplierType))
                {
                    IsExistingType = true;
                    break;
                }
            }
            return IsExistingType;
        }

        public void CreateNewType(string supplierType)
        {
            var newType = WaitForElementIsVisible(By.XPath("//*/div[contains(@class,'title-bar')]/a[text()=' New']"));
            newType.Click();

            // modal fill
            var newTypeSupplierType = WaitForElementIsVisible(By.Id("first"));
            newTypeSupplierType.SendKeys(supplierType);
            var newTypeIsInterim = WaitForElementIsVisible(By.XPath("//*/label[text()='IsInterim']"));
            newTypeIsInterim.SetValue(PageBase.ControlType.CheckBox, true);

            var newTypeSave = WaitForElementIsVisible(By.Id("last"));
            newTypeSave.Click();
        }

        public void EditCurrency(string currency)
        {
            var currencyDolar = WaitForElementIsVisible(By.XPath(CURRENCY_DOLAR));
            currencyDolar.Click();
 
            if ((currencyDolar != null) && (currencyDolar.Text.Contains(currency)))
            {
                var EditBtn = WaitForElementIsVisible(By.XPath(EDIT_CURRENCY));
                EditBtn.Click();

                var ValueCurrency = WaitForElementIsVisible(By.Id(CURRENCY_COEFF));
                ValueCurrency.SetValue(ControlType.TextBox, "1.5");
                WaitForLoad();

                var save = WaitForElementIsVisible(By.Id(SAVE));
                save.Click();

            }
            WaitForLoad();
        }
    }

}
