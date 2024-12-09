using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.CustomerPortal
{
    public class CustomerOrderPortalMessagesPage : PageBase
    {
        public CustomerOrderPortalMessagesPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
            
        }

        // _________________________________________ Constantes _______________________________________________

        private const string CLOSE_SEND_MESSAGE = "//*[@id=\"bootstrap-modal\"]/div/div/div[1]/div/div[2]/button";

        private const string MESSAGE_TEXT = "//*[@id=\"txtMsg\"]";
        private const string CLOSE_ANSWER_SEND_MESSAGE = "//*[@id=\"bootstrap-modal\"]/div/div/div[1]/div/div[2]/button";
        private const string SEND_MESSAGE = "//*[@id=\"btnSend\"]";
        private const string VIEW_MESSAGE = "/html/body/div[5]/div/div/div[1]/div/div[1]/h4";
        private const string VIEW_MESSAGE_BODY = "/html/body/div[4]/div/div[2]/div/div/div/div[2]/table/tbody/tr/td[6]/a";

        // _________________________________________ Variables _________________________________________________

        [FindsBy(How = How.XPath, Using = CLOSE_SEND_MESSAGE)]
        private IWebElement _close;

        [FindsBy(How = How.XPath, Using = MESSAGE_TEXT)]
        private IWebElement _messagetext;

        [FindsBy(How = How.XPath, Using = SEND_MESSAGE)]
        private IWebElement _sendmessage;

        [FindsBy(How = How.XPath, Using = CLOSE_ANSWER_SEND_MESSAGE)]
        private IWebElement _closeanswermessage;
        [FindsBy(How = How.XPath, Using = VIEW_MESSAGE_BODY)]
        private IWebElement _viewMessage;

        public void ViewMessage()
        {
            WaitForLoad();
            _viewMessage = WaitForElementIsVisible(By.XPath(VIEW_MESSAGE_BODY));
            _viewMessage.Click();
            WaitForLoad();            
        }

        public bool VerifyViewMessage()
        {
            WaitForLoad();
            if (!isElementVisible(By.XPath(VIEW_MESSAGE)))
            {
                return false;
            }
            return true;
        }

        public void CloseViewMessage()
        {
            _close = WaitForElementToBeClickable(By.XPath(CLOSE_SEND_MESSAGE));
            _close.Click();
            WaitForLoad();
        }

        public void AnswerMessage(string message)
        {
            _messagetext = WaitForElementIsVisible(By.XPath(MESSAGE_TEXT));
            _messagetext.SetValue(ControlType.TextBox, message);

            _sendmessage = WaitForElementToBeClickable(By.XPath(SEND_MESSAGE));
            _sendmessage.Click();

            _closeanswermessage = WaitForElementToBeClickable(By.XPath(CLOSE_ANSWER_SEND_MESSAGE));
            _closeanswermessage.Click();
            WaitForLoad();
        }

    }
}
