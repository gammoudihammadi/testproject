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
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;

namespace Newrest.Winrest.FunctionalTests.PageObjects.CustomerPortal
{
    public class CustomerPortalLoginPage : PageBase
    {

        public CustomerPortalLoginPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // _________________________________________ Constantes _______________________________________________

        private const string USER_NAME = "UserName";
        private const string PASSWORD = "Password";
        private const string LOG_IN_BTN = "/html/body/div[3]/div/div/section/form/div[4]/div/div/button";
        private const string LOG_IN_BTN_DEV = "/html/body/div[3]/div/div/div/section/form/div[4]/div/div/button";

        private const string RESEND_MAIL_PWD_BTN = "btnResendMail";

        private const string USER_PROFILE = "profile";
        private const string LOG_OUT = "//*[@id=\"logoutForm\"]/div/div/a";

        private const string QTY_BY_WEEK_BTN = "/html/body/div[3]/div/div/form/div/button";
        private const string QTY_BY_WEEK_BTN_DEV = "/html/body/div[3]/div/div[1]/div/form/div/button";
        private const string QTY_BY_DAY_BTN = "/html/body/div[3]/div/div/form[3]/div/button";
        private const string QTY_BY_DAY_BTN_DEV = "/html/body/div[3]/div/div[3]/div/form/div/button";
        //FIXME c'est pas plutot CUSTOMER_ORDER_BTN ?
        private const string CUSTOMER_ORDERS_BTN = "/html/body/div[3]/div/div/form[2]/div/button";
        private const string CUSTOMER_ORDERS_BTN_DEV = "/html/body/div[3]/div/div[2]/div/form/div/button";

        // _________________________________________ Variables _________________________________________________

        [FindsBy(How = How.Id, Using = USER_NAME)]
        private IWebElement _userName;

        [FindsBy(How = How.Id, Using = PASSWORD)]
        private IWebElement _password;

        [FindsBy(How = How.XPath, Using = LOG_IN_BTN)]
        private IWebElement _logIn;

        [FindsBy(How = How.Id, Using = RESEND_MAIL_PWD_BTN)]
        private IWebElement _resendMail;


        [FindsBy(How = How.Id, Using = USER_PROFILE)]
        private IWebElement _userProfile;

        [FindsBy(How = How.XPath, Using = LOG_OUT)]
        private IWebElement _logOut;

        [FindsBy(How = How.XPath, Using = QTY_BY_WEEK_BTN)]
        private IWebElement _qtyByWeekBtn;

        [FindsBy(How = How.XPath, Using = CUSTOMER_ORDERS_BTN)]
        private IWebElement _customerOrdersBtn;

        // _________________________________________ Méthodes __________________________________________________

        public void LogIn(string userName, string password)
        {
                _webDriver.FindElement(By.Id(USER_NAME));

                _userName = WaitForElementIsVisible(By.Id(USER_NAME));
                _userName.SetValue(ControlType.TextBox, userName);

                WaitForLoad();
                _password = WaitForElementIsVisible(By.Id(PASSWORD));
                _password.SetValue(ControlType.TextBox, password);

                //WaitForLoad();
                //var _rememberMe = WaitForElementIsVisible(By.Id("RememberMe"));
                //_rememberMe.SetValue(ControlType.CheckBox, true);

                WaitForLoad();
             _logIn = WaitForElementToBeClickable(By.XPath(LOG_IN_BTN_DEV));
            _logIn.Click();
            WaitForLoad();
        }

        // ____________________________________________ Connexion _____________________________________________________

        public bool IsUserLogged(string userName)
        {
            bool isLogged = true;

            if(isElementVisible(By.Id(USER_PROFILE)))
            {
                _userProfile = WaitForElementIsVisible(By.Id(USER_PROFILE));
                if (!_userProfile.Text.Equals(userName))
                    isLogged = false;
            }
            else
            {
                isLogged = false;
            }

            return isLogged;
        }

        public CustomerPortalLoginPage LogOut()
        {
            _logOut = WaitForElementIsVisible(By.XPath(LOG_OUT));
            _logOut.Click();
            WaitPageLoading();

            return new CustomerPortalLoginPage(_webDriver, _testContext);
        }

        public bool IsPageLoadedWithoutError()
        {
            if (isElementVisible(By.XPath(QTY_BY_WEEK_BTN_DEV)))
            {
                _qtyByWeekBtn = _webDriver.FindElement(By.XPath(QTY_BY_WEEK_BTN_DEV));
                return true;
            }
            else
            {
                return false;
            }
        }

        public CustomerPortalQtyByWeekPage ClickOnQtyByWeek()
        {
            _qtyByWeekBtn = WaitForElementIsVisible(By.XPath(QTY_BY_WEEK_BTN_DEV));
            _qtyByWeekBtn.Click();
            WaitForLoad();
            return new CustomerPortalQtyByWeekPage(_webDriver, _testContext);
        }

        public CustomerPortalCustomerOrdersPage ClickOnCustomerOrders()
        {
            _customerOrdersBtn = WaitForElementExists(By.XPath(CUSTOMER_ORDERS_BTN_DEV));
            _customerOrdersBtn.Click();
            WaitForLoad();
            return new CustomerPortalCustomerOrdersPage(_webDriver, _testContext);
        }
        public CustomerPortalQtyByDayPage ClickOnQtyByDay()
        {
            IWebElement qtyByDayBtn;
            qtyByDayBtn = WaitForElementIsVisible(By.XPath(QTY_BY_DAY_BTN_DEV));
            qtyByDayBtn.Click();
            WaitForLoad();
            return new CustomerPortalQtyByDayPage(_webDriver, _testContext);
        }
        public void ResendMailForPassword(string userName)
        {
            _userName = WaitForElementIsVisible(By.Id(USER_NAME));
            _userName.SetValue(ControlType.TextBox, userName);

            _resendMail = WaitForElementToBeClickable(By.Id(RESEND_MAIL_PWD_BTN));
            _resendMail.Click();
            WaitForLoad();
        }

    }
}
