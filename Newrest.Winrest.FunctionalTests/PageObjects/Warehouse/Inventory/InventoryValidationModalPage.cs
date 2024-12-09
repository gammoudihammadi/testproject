using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Globalization;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.Inventory
{
    public class InventoryValidationModalPage : PageBase
    {
        public InventoryValidationModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // _______________________________ Constantes _________;____________________________________
        private const string VALIDATE_INVENTORY_POPUP = "//h4[text()='Validate the inventory']";
        private const string PARTIAL_INVENTORY = "SetQtysToTheo";
        private const string TOTAL_INVENTORY = "SetQtysToZero";

        private const string CONFIRM_VALIDATE = "btn-submit-validate-inventory";
        private const string UPDATE = "btn-submit-update-items-inventory";

        // _______________________________ Variables ______________________________________________

        [FindsBy(How = How.Id, Using = UPDATE)]
        private IWebElement _updateItems;

        [FindsBy(How = How.Id, Using = CONFIRM_VALIDATE)]
        private IWebElement _confirmValidate;

        [FindsBy(How = How.Id, Using = PARTIAL_INVENTORY)]
        private IWebElement _copyTheo;

        [FindsBy(How = How.Id, Using = TOTAL_INVENTORY)]
        private IWebElement _setQtystoZero;

        // ________________________________ Méthodes ______________________________________________

        

        public InventoryItem ValidatePartialInventory()
        {
            // On coche la case permettant de copier les quantités théoriques dans les quantités physiques
            //(Pour cela, il faut que l'item n'est pas de theroical qty)
            //VALIDATE_INVENTORY_POPUP
            if (isElementExists(By.Id(PARTIAL_INVENTORY)))
            {
                _copyTheo = WaitForElementExists(By.Id(PARTIAL_INVENTORY));
                _copyTheo.SetValue(ControlType.CheckBox, true);
                Assert.IsTrue(CanUpdateItems(), "Le bouton de validation de l'inventaire est indisponible.");
            
            // Validation
            _updateItems = WaitForElementIsVisible(By.Id(UPDATE));
            _updateItems.Click();
            WaitForLoad();
            }

            _confirmValidate = WaitForElementIsVisible(By.Id(CONFIRM_VALIDATE));
            _confirmValidate.Click();
            WaitForLoad();

            return new InventoryItem(_webDriver, _testContext);
        }

        public InventoryItem ValidateTotalInventory()
        {
            // On coche la case permettant de mettre à 0 les qtty
            //(Pour cela, il faut que l'item n'est pas de theroical qty)
            if (isElementExists(By.Id(TOTAL_INVENTORY)))
            {
                _setQtystoZero = WaitForElementExists(By.Id(TOTAL_INVENTORY));
                _setQtystoZero.SetValue(ControlType.CheckBox, true);
                WaitForLoad();

                Assert.IsTrue(CanUpdateItems(), "Le bouton de validation de l'inventaire est indisponible.");

                // Validation
                _updateItems = WaitForElementIsVisible(By.Id(UPDATE));
                _updateItems.Click();
                WaitForLoad();
            }

            _confirmValidate = WaitForElementIsVisible(By.Id(CONFIRM_VALIDATE));
            _confirmValidate.Click();
            WaitForLoad();

            return new InventoryItem(_webDriver, _testContext);
        }

        public bool CanUpdateItems()
        {
            if(isElementVisible(By.Id(UPDATE)))
            {
                _updateItems = _webDriver.FindElement(By.Id(UPDATE));
                return _updateItems.GetAttribute("disabled") == null;
            }
            else
            {
                return false;
            }
        }

        public Dictionary<string, double> GetTop10(string currency, string decimalSeparatorValue)
        {
            CultureInfo ci = decimalSeparatorValue.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");
            Dictionary<string, double> top10 = new Dictionary<string, double>();
            var elements = _webDriver.FindElements(By.XPath("//*/table[@id='tableItem']/tbody/tr/td[1]"));
            foreach(var element in elements)
            {
                var value = _webDriver.FindElement(By.XPath("//*/table[@id='tableItem']/tbody/tr/td[position()=1 and contains(text(),\"" + element.Text+"\")]/../td[2]"));
                top10.Add(element.Text, double.Parse(value.Text.Replace(currency, ""), ci));
            }
            return top10;
        }
    }
}
