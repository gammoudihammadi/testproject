using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Menus.Recipes
{
    public class RecipeDuplicateVariantPage : PageBase
    {
        public RecipeDuplicateVariantPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // ____________________________________________ Constantes ______________________________________________

        private const string VARIANT_SITE = "SearchVM_Site";
        private const string VARIANT_NAME = "SearchVM_Variant";
        private const string SELECTED_VARIANT = "//*[@id=\"recipe-siteVariant-container\"]/div/div/form/table/tbody/tr[2]";
        private const string VALIDATE = "//*/a[text()=' Validate']";
        private const string SAVED_ICON = "//img[@title = 'Duplication succeed']";
        private const string ERROR_ICON = "//img[@class = 'duplication-error-icon']";

        private const string CLOSE = "/html/body/div[7]/div/div/div[1]/button";

        // ____________________________________________ Variables _______________________________________________

        [FindsBy(How = How.Id, Using = VARIANT_SITE)]
        private IWebElement _variantSite;

        [FindsBy(How = How.Id, Using = VARIANT_NAME)]
        private IWebElement _variantName;

        [FindsBy(How = How.XPath, Using = SELECTED_VARIANT)]
        private IWebElement _selectedVariant;

        [FindsBy(How = How.XPath, Using = VALIDATE)]
        private IWebElement _validate;

        [FindsBy(How = How.XPath, Using = CLOSE)]
        private IWebElement _close;

        // ____________________________________________ Méthodes ________________________________________________

        public void FillFields_DuplicateVariant(string site, string variant)
        {
            _variantSite = WaitForElementIsVisible(By.Id(VARIANT_SITE));
            _variantSite.SetValue(ControlType.TextBox, site);

            _variantName = WaitForElementIsVisible(By.Id(VARIANT_NAME));
            _variantName.SetValue(ControlType.TextBox, variant);

            WaitForLoad();
            if(isElementVisible(By.XPath(SELECTED_VARIANT)))
            {
                _selectedVariant = WaitForElementIsVisible(By.XPath(SELECTED_VARIANT));
                _selectedVariant.Click();
            }            

            _validate = WaitForElementIsVisible(By.XPath(VALIDATE));
            _validate.Click();

            var selectedVariant = WaitForElementIsVisible(By.Id("recipe-select-all-text"));
            selectedVariant.Click();
            //selectedVariant.Click();
            //selectedVariant.Click();
            WaitForLoad();
            _validate.Click();
            WaitForLoad();
        }

        public bool IsDuplicateSaved()
        {
            try
            {
                WaitForElementIsVisible(By.XPath(SAVED_ICON));//*[@id="recipe-siteVariant-container"]/div/div/div/div/div/div/form/table/tbody/tr[2]/td[5]/img
                return true;
            }
            catch
            {
                return false;
            }
        }

        public bool IsDuplicateInError()
        {
            try
            {
                WaitForElementIsVisible(By.XPath(ERROR_ICON));
                return true;
            }
            catch
            {
                return false;
            }
        }

        public RecipesVariantPage CloseDuplicate()
        {
            _close = WaitForElementIsVisible(By.XPath(CLOSE));
            _close.Click();

            // Temps de fermeture de la page
            WaitPageLoading();

            return new RecipesVariantPage(_webDriver, _testContext);
        }
    }
}
