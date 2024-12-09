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
    public class ItemDieteticPage : PageBase
    {
        public ItemDieteticPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // _____________________________________________ Constantes ________________________________________________

        private const string ENERGYKJ_FIELD = "/html/body/div[3]/div/div[2]/div[2]/div/div/div/form/div[2]/div[1]/div[1]/div[1]/div/input";
        
        private const string SEARCH_CIQUAL = "//*[@id=\"nutrition-info-form\"]/div[1]/a";

        private const string GENERAL_INFO_TAB = "hrefTabContentItem";

        // _____________________________________________ Variables _______________________________________________

        [FindsBy(How = How.Id, Using = ENERGYKJ_FIELD)]
        private IWebElement _energyKJField;

        [FindsBy(How = How.XPath, Using = SEARCH_CIQUAL)]
        private IWebElement _searchCIQUAL;

        [FindsBy(How = How.Id, Using = GENERAL_INFO_TAB)]
        private IWebElement _generalInfoTab;

        public void SetEnergyKJ_numericalValue(string value)
        {
            _energyKJField = WaitForElementIsVisible(By.XPath(ENERGYKJ_FIELD));//*[@id="EnergyKJ"]
            _energyKJField.SetValue(ControlType.TextBox, value);
            Thread.Sleep(1000);
        }

        public string GetEnergyKJ_numericalValue()
        {
            _energyKJField = WaitForElementIsVisible(By.XPath(ENERGYKJ_FIELD));
            return _energyKJField.GetAttribute("value");
        }

        public ItemSearchCiqualModalPage ClickOnCiqualButton()
        {
            _searchCIQUAL = WaitForElementIsVisible(By.XPath(SEARCH_CIQUAL));
            _searchCIQUAL.Click();
            WaitForLoad();
            return new ItemSearchCiqualModalPage(_webDriver, _testContext);
        }

        public ItemGeneralInformationPage ClickOnGeneralInfo()
        {
            _generalInfoTab = WaitForElementIsVisible(By.Id(GENERAL_INFO_TAB));
            _generalInfoTab.Click();
            WaitPageLoading();

            return new ItemGeneralInformationPage(_webDriver, _testContext);
        }

        public string GetMacroNutriments(string id)
        {
            var alcohol = WaitForElementIsVisible(By.Id(id));
            return alcohol.GetAttribute("value");
        }

        public void SetMacroNutriments(string id, string value)
        {
            var alcohol = WaitForElementIsVisible(By.Id(id));
            alcohol.Clear();
            alcohol.SendKeys(value);
            WaitLoading();
        }
    }
}
