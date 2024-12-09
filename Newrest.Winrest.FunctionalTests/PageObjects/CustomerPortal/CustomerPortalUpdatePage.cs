using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.CustomerPortal
{
    public class CustomerPortalUpdatePage : PageBase
    {

        public CustomerPortalUpdatePage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // _________________________________________ Constantes _______________________________________________

        private const string PASSWORD = "//*[@id=\"Password\"]";
        private const string CONFIRM_PASSWORD = "//*[@id=\"ConfirmPassword\"]";
        private const string UPDATE = "/html/body/div[3]/div/div/section/form/div[4]/div/button";

        // _________________________________________ Variables _________________________________________________

        [FindsBy(How = How.XPath, Using = PASSWORD)]
        private IWebElement _password;

        [FindsBy(How = How.XPath, Using = CONFIRM_PASSWORD)]
        private IWebElement _confirmPassword;

        [FindsBy(How = How.XPath, Using = UPDATE)]
        private IWebElement _update;

        // _________________________________________ Méthodes __________________________________________________

        public CustomerPortalQtyByWeekPage UpdatePassword(string userPassword)
        {
                Thread.Sleep(2000);
                var _password = WaitForElementIsVisible(By.XPath(PASSWORD));
                _password.SetValue(ControlType.TextBox,userPassword);
                WaitForLoad();

                // Attendre et trouver l'élément de confirmation du mot de passe
                _confirmPassword = WaitForElementIsVisible(By.XPath(CONFIRM_PASSWORD));
                _confirmPassword.SetValue(ControlType.TextBox, userPassword);
                WaitForLoad();

                // Attendre et trouver le bouton de mise à jour
                _update = WaitForElementIsVisible(By.XPath(UPDATE));
                _update.Click();
                WaitForLoad();

            return new CustomerPortalQtyByWeekPage(_webDriver, _testContext);
        }

    }

}
