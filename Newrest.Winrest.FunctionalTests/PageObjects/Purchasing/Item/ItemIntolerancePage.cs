using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Purchasing
{
    public class ItemIntolerancePage : PageBase
    {
        public ItemIntolerancePage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // _____________________________________ Constantes _________________________________________________

        private const string ALLERGEN_LINE = "//*[@id=\"itemAllergenForm\"]/table/tbody/tr[*]/td[2]/label[text()='{0}']";
        private const string ALLERGEN_ALL_LINE = "//*[@id=\"itemAllergenForm\"]/table/tbody/tr[*]/td[2]/label";

        private const string SAVE_ERROR_ICON = "//*[@id=\"itemAllergenForm\"]/div/span[@class='glyphicon glyphicon-floppy-remove']";
        private const string SAVE_ICON = "//*[@id=\"itemAllergenForm\"]/div/span[@class='glyphicon glyphicon-floppy-saved']";
        private const string LIST_ALLERGENS_SELECTED = "//*[@id=\"allergenCheckBox\" and @checked]";

        private const string ALLERGEN_LINE_NAME = "//*[@id=\"tabContentDetails\"]/div[2]/table/tbody/tr[*]/td[2]/label[text()='{0}']";

        private const string KEYWORDS = "KeywordValuesIds_ms";
        // _____________________________________ Variables __________________________________________________

        [FindsBy(How = How.XPath, Using = ALLERGEN_LINE)]
        private IWebElement _allergenLine;

        [FindsBy(How = How.XPath, Using = ALLERGEN_LINE_NAME)]
        private IWebElement _allergenLineName;

        // _____________________________________ Méthodes ___________________________________________________

        public void AddAllergen(string allergen)
        {
            Actions action = new Actions(_webDriver);

            _allergenLine = WaitForElementExists(By.XPath(String.Format(ALLERGEN_LINE, allergen)));
            action.MoveToElement(_allergenLine).Perform();

            var allergenBox = _webDriver.FindElement(By.Id(_allergenLine.GetAttribute("for")));
            if(allergenBox.GetAttribute("checked") != "true")
                allergenBox.Click();

            // Temps d'enregistrement de la donnée
            WaitPageLoading();
        }

        public void AddFirstAllergen(string allergen)
        {
            Actions action = new Actions(_webDriver);
            _allergenLineName = WaitForElementExists(By.XPath(String.Format(ALLERGEN_LINE_NAME, allergen)));
            action.MoveToElement(_allergenLineName).Perform();

            var allergenBox = _webDriver.FindElement(By.Id("allergenCheckBox"));
                allergenBox.Click();
        }
        public void UncheckAllergen(string allergen)
        {
            Actions action = new Actions(_webDriver);

            _allergenLine = WaitForElementExists(By.XPath(String.Format(ALLERGEN_LINE, allergen)));
            action.MoveToElement(_allergenLine).Perform();

            var allergenBox = _webDriver.FindElement(By.Id(_allergenLine.GetAttribute("for")));
            if (allergenBox.GetAttribute("checked") == "true")
                allergenBox.Click();

            // Temps d'enregistrement de la donnée
            WaitPageLoading();
        }

        public void UncheckAllAllergen()
        {

            var checkallallergenh = _webDriver.FindElements(By.XPath(ALLERGEN_ALL_LINE));

      
            foreach (var element in checkallallergenh)
            {
                var allergenBox = _webDriver.FindElement(By.Id(element.GetAttribute("for")));
                if (allergenBox.GetAttribute("checked") == "true")
                    allergenBox.Click();
            }
            // Temps d'enregistrement de la donnée
            WaitPageLoading();
        }

        public bool IsAllergenChecked(string allergen) 
        {
            Actions action = new Actions(_webDriver);

            _allergenLine = WaitForElementExists(By.XPath(String.Format(ALLERGEN_LINE, allergen)));
            action.MoveToElement(_allergenLine).Perform();

            var allergenBox = _webDriver.FindElement(By.Id(_allergenLine.GetAttribute("for")));

            if (allergenBox.GetAttribute("checked") != "true")
                return false;
            else
                return true;
        }

        public string GetAllergenError(string allergen)
        {
            string errorMsg = "";
            if (isElementVisible(By.XPath(String.Format("//label[text()='{0}']/../..//input[contains(@id, 'IsSelected')]", allergen))))
            {
                var checkbox = WaitForElementExists(By.Id("Allergens_0__IsSelected"));
                errorMsg = checkbox.GetAttribute("title");
            }
            
            return errorMsg;
        }

        public bool SaveAllergenError(string allergen, bool newTab)
        {
            WaitPageLoading();
            bool isError = false;

            if (isElementVisible(By.XPath(String.Format(ALLERGEN_LINE, allergen))))
            {
                _allergenLine = WaitForElementExists(By.XPath(String.Format(ALLERGEN_LINE, allergen)));

                var checkbox = WaitForElementExists(By.Id("Allergens_0__IsSelected"));
                var errorMsg = checkbox.GetAttribute("title");

                var allergenBox = _webDriver.FindElement(By.Id(_allergenLine.GetAttribute("for")));
                if (allergenBox.GetAttribute("checked") != "true")
                    isError = true;
            }
            else
            {
                isError = false;
            }

            if (newTab)
            {
                Close();
            }

            return isError;
        }

        public bool SaveAllergen()
        {
            if(isElementVisible(By.XPath(SAVE_ICON)))
            {
                WaitForElementIsVisible(By.XPath(SAVE_ICON));
                return true;
            }
            else
            {
                return false;
            }
        }
        public int CalculateAllergens()
        {
            int count = 0; 
            var allergensSelected = _webDriver.FindElements(By.XPath("//*[@id=\"itemAllergenForm\"]/table/tbody/tr[*]/td[2]/label"));
            foreach (var element in allergensSelected)
            {
                var allergenBox = _webDriver.FindElement(By.Id(element.GetAttribute("for")));
                if (allergenBox.GetAttribute("checked") == "true")
                    count = count + 1;
            }
            return count;
        }
        public List<string> GetListAllergensSelected()
        {
            var allergensSelected = _webDriver.FindElements(By.XPath("//*[@id=\"itemAllergenForm\"]/table/tbody/tr[*]/td[2]/label"));
            var list = new List<string>();
            foreach (var element in allergensSelected)
            {
                var allergenBox = _webDriver.FindElement(By.Id(element.GetAttribute("for")));
                if (allergenBox.GetAttribute("checked") == "true")
                    list.Add(element.Text);
            }
            return list;
        }        
    }
}