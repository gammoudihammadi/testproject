using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Menus.Datasheet
{
    public class DatasheetIntolerancePage : PageBase
    {
        public DatasheetIntolerancePage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // ___________________________________ Constantes _____________________________________________

        private const string DATASHEET_INTOLERANCE_PAGE_ALLERGEN_LINE = "//*[@id=\"datasheet-detail-container\"]/table/tbody/tr[*]/td[2]/label[text()='{0}']";
        private const string LIST_ALLERGENS_SELECTED = "//*[starts-with(@id,\"Allergens\") and @checked]";


        // _____________________________________ Variables __________________________________________________

        [FindsBy(How = How.XPath, Using = DATASHEET_INTOLERANCE_PAGE_ALLERGEN_LINE)]
        private IWebElement _allergenLine;

        // _____________________________________ Méthodes ___________________________________________________

        public bool IsDatasheetAllergenChecked(string allergen)
        {
            Actions action = new Actions(_webDriver);

            _allergenLine = WaitForElementExists(By.XPath(String.Format(DATASHEET_INTOLERANCE_PAGE_ALLERGEN_LINE, allergen)));
            action.MoveToElement(_allergenLine).Perform();

            var allergenBox = _allergenLine.GetAttribute("for");

            if (allergenBox.Contains("IsSelected"))
                return true;
            else
                return false;
        }
        public int CalculateAllergens()
        {
            WaitForLoad();
            var allergensSelected = _webDriver.FindElements(By.XPath(LIST_ALLERGENS_SELECTED));
            return allergensSelected.Count;
        }
    }
}
