using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Item
{
    public class ItemSearchCiqualModalPage : PageBase
    {
        public ItemSearchCiqualModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // _____________________________________________ Constantes ________________________________________________
        private const string SEARCH_FIELD = "/html/body/div[4]/div/div/div[2]/div/div[1]/div/form/span/input[2]";
        private const string FIRST_SUGGESTED_CIQUAL = "//*[@id=\"ciqual-search-form\"]/span/div/div/div[1]";
        private const string FIRST_CIQUAL_DISPLAYED = "//a[contains(text(), 'Select')]";

        //private const string FIRST_CIQUAL_DISPLAYED_ENERGY_KJ = "//*[@id=\"EnergyKJ\"]";
        private const string FIRST_CIQUAL_DISPLAYED_ENERGY_KJ = "/html/body/div[4]/div/div/div[3]/div/form/div[1]/div/div/div[1]/div[2]/div[1]/div/input";
        private const string APPLY_FIRST_CIQUAL_DISPLAYED = "//*[@id=\"btn-apply\"]";



        // _____________________________________________ Variables _______________________________________________

        [FindsBy(How = How.XPath, Using = SEARCH_FIELD)]
        private IWebElement _searchField;

        [FindsBy(How = How.XPath, Using = FIRST_SUGGESTED_CIQUAL)]
        private IWebElement _firstSuggestedCiqual;

        [FindsBy(How = How.XPath, Using = FIRST_CIQUAL_DISPLAYED)]
        private IWebElement _firstCiqualDisplayed;

        [FindsBy(How = How.XPath, Using = FIRST_CIQUAL_DISPLAYED_ENERGY_KJ)]
        private IWebElement _firstCiqualDisplayedEnergyKJ;

        [FindsBy(How = How.XPath, Using = APPLY_FIRST_CIQUAL_DISPLAYED)]
        private IWebElement _applyfirstCiqualDisplayed;

        public void SearchCiqualItem(string value)
        {
            _searchField = WaitForElementIsVisible(By.XPath(SEARCH_FIELD));
            _searchField.SetValue(ControlType.TextBox, value);
            WaitForLoad();
        }

        public void ClickOnFirstSuggestedCiqualDisplayed()
        {
            _firstSuggestedCiqual = WaitForElementIsVisible(By.XPath(FIRST_SUGGESTED_CIQUAL));
            _firstSuggestedCiqual.Click();
            WaitForLoad();

        }

        public void SelectFirstCiqualDisplayed()
        {
                _firstCiqualDisplayed = WaitForElementIsVisible(By.XPath("//*/a[@class='btn btn-select-ciqual']"));
                _firstCiqualDisplayed.Click();
            WaitForLoad();

        }

        public string GetFirstCiqualDisplayedEnergyKJ()
        {
            _firstCiqualDisplayedEnergyKJ = WaitForElementIsVisible(By.XPath(FIRST_CIQUAL_DISPLAYED_ENERGY_KJ));
            return _firstCiqualDisplayedEnergyKJ.GetAttribute("value");
        }

        public void ApplyCiqualOnItem()
        {
            _applyfirstCiqualDisplayed = WaitForElementIsVisible(By.XPath(APPLY_FIRST_CIQUAL_DISPLAYED));
            _applyfirstCiqualDisplayed.Click();
            WaitForLoad();
        }
    }
}
