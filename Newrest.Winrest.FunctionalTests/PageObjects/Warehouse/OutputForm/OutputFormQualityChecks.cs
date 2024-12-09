using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.OutputForm
{
    public class OutputFormQualityChecks : PageBase
    {
        public OutputFormQualityChecks(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // _____________________________________________ Constantes _________________________________________________

        // Onglets
        private const string ITEMS_TAB = "hrefTabContentItems";

        // Tableau
        private const string DELIVERY_ACCEPTED = "accepted";
        private const string VERIFIED_BY = "VerifiedBy";
        private const string FROZEN_TEMPERATURE = "//*[@id=\"form-update-quality-check\"]/div/div[1]/div[2]/div/div/input";

        // _____________________________________________ Variables __________________________________________________

        // Onglets
        [FindsBy(How = How.Id, Using = ITEMS_TAB)]
        private IWebElement _itemsTab;

        // Tableau
        [FindsBy(How = How.Id, Using = DELIVERY_ACCEPTED)]
        private IWebElement _deliveryAccepted;

        [FindsBy(How = How.Id, Using = VERIFIED_BY)]
        private IWebElement _verifiedBy;

        [FindsBy(How = How.XPath, Using = FROZEN_TEMPERATURE)]
        private IWebElement _frozenTemperature;

        // _____________________________________________ Méthodes ___________________________________________________

        // Onglets
        public OutputFormItem ClickOnItemsTab()
        {
            WaitForLoad();
            _itemsTab = WaitForElementIsVisible(By.Id(ITEMS_TAB));
            _itemsTab.Click();
            WaitForLoad();

            return new OutputFormItem(_webDriver, _testContext);
        }

        // Tableau
        public string GetVerifiedBy()
        {
            _verifiedBy = WaitForElementExists(By.Id(VERIFIED_BY));
            return _verifiedBy.GetAttribute("value");
        }

        public void SetFrozenTemperature(string temp)
        {
            _frozenTemperature = WaitForElementExists(By.XPath(FROZEN_TEMPERATURE));
            _frozenTemperature.SetValue(ControlType.TextBox, temp);
            _frozenTemperature.Submit();
            Thread.Sleep(2000);//long to save
            WaitForLoad();

            var refriTemperature = WaitForElementExists(By.XPath("//*[@id=\"form-update-quality-check\"]/div/div[1]/div[1]/div/div/input"));
            refriTemperature.SetValue(ControlType.TextBox, temp);
            refriTemperature.Submit();
            WaitForLoad();
            Thread.Sleep(2000);//long to save
        }

        public string GetFrozenTemperature(string decimalSeparator)
        {
            _frozenTemperature = WaitForElementExists(By.XPath(FROZEN_TEMPERATURE));
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");
            return Convert.ToString(_frozenTemperature.GetAttribute("value"), CultureInfo.InvariantCulture);
        }

        public void DeliveryAccepted()
        {
            _deliveryAccepted = WaitForElementIsVisible(By.Id("Delivery-Accepted"));
            _deliveryAccepted.Click();
            Thread.Sleep(2000);
            WaitForLoad();
        }

        public bool IsDeliveryAccepted()
        {
            _deliveryAccepted = WaitForElementIsVisible(By.Id("Delivery-Accepted"));

            if (_deliveryAccepted.GetAttribute("checked").Equals("true"))
                return true;
            else
                return false;
        }

    }
}
