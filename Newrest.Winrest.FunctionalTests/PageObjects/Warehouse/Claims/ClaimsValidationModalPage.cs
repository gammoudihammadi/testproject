using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.Claims
{
    public class ClaimsValidationModalPage : PageBase
    {
        private const string UPDATE_ITEMS_XPATH = "//*[@id=\"btn-submit-update-items-inventory\"]";
        private const string CANCEL = "//*[@id=\"btn-cancel-validate-inventory\"]";
        private const string VALIDATE = " //*[@id=\"btn-submit-validate-inventory\"]";

        //Update Button
        [FindsBy(How = How.XPath, Using = UPDATE_ITEMS_XPATH)]
        private IWebElement _updateItems;

        [FindsBy(How = How.XPath, Using = VALIDATE)]
        private IWebElement _validate;

        [FindsBy(How = How.XPath, Using = "//*[@id=\"SetQtysToTheo\"]")]
        private IWebElement _copyTheo;

        [FindsBy(How = How.XPath, Using = "//*[@id=\"SetQtysToZero\"]")]
        private IWebElement _setQtystoZero;

        [FindsBy(How = How.XPath, Using = "//*[@id=\"modal-1\"]/div/div/div/div/div/form/div[2]/div[1]")]
        private IWebElement _valiadionError;

        //Update Button
        [FindsBy(How = How.XPath, Using = UPDATE_ITEMS_XPATH)]
        private IWebElement _cancel;

        public ClaimsItem Validate()
        {
            try
            {
                //Click on "update items "
                _setQtystoZero.SetValue(ControlType.CheckBox, true);
                _updateItems.Click();

                WaitForLoad();
            }
            catch
            {
                // don't update 
            }
            _validate.Click();
            WaitForLoad();
            return new ClaimsItem(_webDriver, _testContext);
        }
        public bool CanUpdateItems()
        {
            bool canUpdate = true;
            try
            {
                var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(2));
                _updateItems = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(UPDATE_ITEMS_XPATH)));
                canUpdate = _updateItems.GetAttribute("disabled") == null;
            }
            catch (NoSuchElementException e)
            {
                canUpdate = false;
            }

            return canUpdate;
        }
        public ClaimsValidationModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }
    }
}
