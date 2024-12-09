using DocumentFormat.OpenXml.Drawing.Charts;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Globalization;
using System.Threading;
using System.Windows.Forms;
using Keys = OpenQA.Selenium.Keys;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Purchasing
{
    public class SupplyOrderGeneralInformation : PageBase
    {

        // __________________________________________ Constantes ______________________________________________

        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";

        private const string DELIVERY_DATE = "datapicker-new-supply-order-delivery";
        private const string DATE_FROM = "datapicker-new-supply-order-start";
        private const string DATE_TO = "datapicker-new-supply-order-end";
        private const string PURCHASE_ORDER = "//span[contains(text(), 'purchase')]/a";

        private const string ITEMS_TAB = "hrefTabContentItems";
        private const string BASED_ON_MAIN_SUPPLIER_CHECKBOX = "SupplyOrder_IsBasedOnMainSupplier";
        private const string ROUND_PREFILLED_QUANTITIES_CHECKBOX = "checkBoxRoundPackagingQtys";

        // __________________________________________ Variables _______________________________________________

        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;

        [FindsBy(How = How.Id, Using = DELIVERY_DATE)]
        private IWebElement _deliveryDate;

        [FindsBy(How = How.Id, Using = DATE_FROM)]
        private IWebElement _dateFrom;

        [FindsBy(How = How.Id, Using = DATE_TO)]
        private IWebElement _dateTo;

        [FindsBy(How = How.XPath, Using = PURCHASE_ORDER)]
        private IWebElement _purchaseOrder;

        [FindsBy(How = How.Id, Using = ITEMS_TAB)]
        private IWebElement _itemsTab;

        // __________________________________________ Méthodes _________________________________________________

        public SupplyOrderGeneralInformation(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }
        public SupplyOrderPage BackToList()
        {
            _backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            _backToList.Click();
            WaitForLoad();

            return new SupplyOrderPage(_webDriver, _testContext);
        }

        public string GetDeliveryDateValue()
        {
            var dateFormat = _deliveryDate.GetAttribute("data-date-format");
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? ci = new CultureInfo("fr-FR") : new CultureInfo("en-US");

            _deliveryDate = WaitForElementIsVisible(By.Id(DELIVERY_DATE));
            String deliveryDate = _deliveryDate.GetAttribute("value");
            
            return DateTime.Parse(deliveryDate, ci).ToShortDateString();
        }
   
        public void DeliveryDateUpdate(DateTime date)
        {
            _deliveryDate = WaitForElementIsVisible(By.Id(DELIVERY_DATE));
            _deliveryDate.SetValue(ControlType.DateTime, date);
            WaitForLoad();
        }

        public string GetFromDateValue()
        {
            var dateFormat = _dateFrom.GetAttribute("data-date-format");
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? ci = new CultureInfo("fr-FR") : new CultureInfo("en-US");

            _dateFrom = WaitForElementIsVisible(By.Id(DATE_FROM));
            String fromDate = _dateFrom.GetAttribute("value");
            
            return DateTime.Parse(fromDate, ci).ToShortDateString();
        }

        public void FromDateUpdate(DateTime date)
        {
            _dateFrom = WaitForElementIsVisible(By.Id(DATE_FROM));
            _dateFrom.SetValue(ControlType.DateTime, date);
            WaitForLoad();
        }

        public string GetToDateValue()
        {
            var dateFormat = _dateTo.GetAttribute("data-date-format");
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? ci = new CultureInfo("fr-FR") : new CultureInfo("en-US");

            _dateTo = WaitForElementIsVisible(By.Id(DATE_TO));
            String toDate = _dateTo.GetAttribute("value");

            return DateTime.Parse(toDate, ci).ToShortDateString();
        }

        public void ToDateUpdate(DateTime date)
        {
            _dateTo = WaitForElementIsVisible(By.Id(DATE_TO));
            _dateTo.SetValue(ControlType.DateTime, date);
            _dateTo.SendKeys(Keys.Enter);
            //_dateTo.SendKeys(Keys.Tab);
            // Prise en compte de la modification
            //Thread.Sleep(2000);
            WaitForLoad();
        }

        public string GetPurchaseOrderValue()
        {
            if (isElementVisible(By.XPath(PURCHASE_ORDER)))
            {
                _purchaseOrder = WaitForElementExists(By.XPath(PURCHASE_ORDER));
                return _purchaseOrder.Text;
            }
            else
            {
                return "";
            }
        }

        public SupplyOrderItem ClickOnItemsTab()
        {
            _itemsTab = WaitForElementIsVisible(By.Id(ITEMS_TAB));
            _itemsTab.Click();
            WaitForLoad();

            return new SupplyOrderItem(_webDriver, _testContext);
        }
        public void SetBasedOnMainSupplier(bool value)
        {
                    var basedOnMainSupplierCheckbox = WaitForElementExists(By.Id(BASED_ON_MAIN_SUPPLIER_CHECKBOX));
                    basedOnMainSupplierCheckbox.SetValue(ControlType.CheckBox, value);
            WaitForLoad();
        }
        public void SetRoundPrefilledQuantities(bool value)
        {
            var setRoundPrefilledQuantitiesCheckbox = WaitForElementExists(By.Id(ROUND_PREFILLED_QUANTITIES_CHECKBOX));
            setRoundPrefilledQuantitiesCheckbox.SetValue(ControlType.CheckBox, value);
            WaitForLoad();
        }
        public bool GetBasedOnMainSupplierValue()
        {
            var basedOnMainSupplierCheckbox = _webDriver.FindElement(By.Id(BASED_ON_MAIN_SUPPLIER_CHECKBOX));
            return basedOnMainSupplierCheckbox.Selected;
        }
        public bool GetRoundPrefiledQuantitiesValue()
        {
            var RoundPrefilledQuantitiesCheckbox = _webDriver.FindElement(By.Id(ROUND_PREFILLED_QUANTITIES_CHECKBOX));
            return RoundPrefilledQuantitiesCheckbox.Selected;
        }
    }


}
