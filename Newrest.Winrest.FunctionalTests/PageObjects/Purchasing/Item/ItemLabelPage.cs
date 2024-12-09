using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Item
{
    public class ItemLabelPage : PageBase
    {
        public ItemLabelPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        private const string COMPOSITION = "Composition";
        private const string CARACTERISTICS = "dropdown-caracteristics";

        [FindsBy(How = How.Id, Using = COMPOSITION)]
        private IWebElement _compositionText;

        [FindsBy(How = How.Id, Using = CARACTERISTICS)]
        private IWebElement _caracteristics;

        public void EditLabel(string text)
        {
            _compositionText = WaitForElementIsVisible(By.Id(COMPOSITION));
            _compositionText.SetValue(ControlType.TextBox, text);

            _caracteristics = WaitForElementIsVisible(By.Id(CARACTERISTICS));
            _caracteristics.SetValue(ControlType.TextBox, text);
            //Temps de prise compte des valeurs
            Thread.Sleep(1500);
        }

        public void deleteText() 
        {
            _compositionText = WaitForElementIsVisible(By.Id(COMPOSITION));
            _caracteristics = WaitForElementIsVisible(By.XPath(CARACTERISTICS));
            _compositionText.SetValue(ControlType.TextBox, "");
            _caracteristics.Click();
        }

        public bool IsEditLabel(string value)
        {
            _compositionText = WaitForElementIsVisible(By.Id(COMPOSITION));
            _caracteristics = WaitForElementIsVisible(By.Id(CARACTERISTICS));

            if (_compositionText.Text == value && _caracteristics.GetAttribute("value") == value)
            {
                return true;
            }

            return false;
        }

        public void Validate()
        {
            var validateButton  = WaitForElementIsVisible(By.XPath("//button[contains(text(),'Validate')]"));
            validateButton.Click();
            WaitForLoad();
        }
    }
}
