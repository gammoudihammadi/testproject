using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;

namespace S
{
    public class LoginPage : PageBase
    {
        private readonly IWebDriver _webDriver;



        [FindsBy(How = How.XPath, Using = "//*/input[@type='email']")]
        private IWebElement _userNameSignInField;

        [FindsBy(How = How.XPath, Using = "//*/input[@data-report-event='Signin_Submit']")]
        private IWebElement _loginSignInButton;

        [FindsBy(How = How.Id, Using = "userNameInput")]
        private IWebElement _userNameField;

        [FindsBy(How = How.Id, Using = "passwordInput")]
        private IWebElement _passwordField;

        [FindsBy(How = How.Id, Using = "submitButton")]
        private IWebElement _loginButton;

        [FindsBy(How = How.XPath, Using = "//*/input[@type='submit']")]
        private IWebElement _confirmButton;

        public LoginPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
            _webDriver = webDriver;
            // parasite
            //PageFactory.InitElements(_webDriver, this);
        }

        public bool Login(string userName, string password)
        {
            try
            {
                WaitForLoad();
                _userNameSignInField = WaitForElementIsVisible(By.XPath("//*/input[@type='email']"));
                _userNameSignInField.SendKeys(userName);
                WaitForLoad();
                _loginSignInButton = WaitForElementIsVisible(By.XPath("//*/input[@data-report-event='Signin_Submit']"));
                _loginSignInButton.Click();
                // animation
                WaitForLoad();
            } catch
            {
                // Autologin
            }
            _userNameField = WaitForElementIsVisible(By.Id("userNameInput"));
            _userNameField.Clear();
            _userNameField.SendKeys(userName);
            WaitForLoad();
            _passwordField = WaitForElementIsVisible(By.Id("passwordInput"));
            _passwordField.SendKeys(password);
            WaitForLoad();
            _loginButton = WaitForElementIsVisible(By.Id("submitButton"));
            _loginButton.Click();
            WaitForLoad();


            Thread.Sleep(2000);
            if (isElementVisible(By.XPath("//*/input[@type='button' and @id='idBtn_Back']")))
            {
                try
                {
                    // en théorie la page KMSI ne s'affiche pas si je clique non, en pratique heu
                    //_confirmButton = WaitForElementIsVisible(By.XPath("//*/input[@type='submit']"));
                    _confirmButton = WaitForElementIsVisible(By.XPath("//*/input[@type='button' and @id='idBtn_Back']"));
                    _confirmButton.Click();
                }
                catch
                {
                    // System Ready : https://winrest-testauto5dev-app.azurewebsites.net/
                }
                // _webDriver.Title le fait déjà
                //WaitForLoad();
            }
            else
            {
                // System Ready : https://winrest-testauto5dev-app.azurewebsites.net/
            }

            //OpenQA.Selenium.WebDriverException: The HTTP request to the remote WebDriver server for URL http://localhost:62595/session/85025c17f4781ba1e1538918cb54f04c/title timed out after 60 seconds. ---> System.Net.WebException: La demande a été abandonnée : Le délai d'attente de l'opération a expiré..
            bool ready = _webDriver.Title != "Connexion";

            WaitForLoad();
            return ready;
        }

        protected void WaitForLoad()
        {
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(61));

            Func<IWebDriver, bool> readyCondition = webDriver =>
                (bool)javaScriptExecutor.ExecuteScript("return (document.readyState == 'complete')");

            wait.Until(readyCondition);
        }
    }
}
