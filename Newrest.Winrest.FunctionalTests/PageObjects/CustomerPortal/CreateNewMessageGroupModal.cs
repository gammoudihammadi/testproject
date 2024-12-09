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
    public class CreateNewMessageGroupModal : PageBase
    {
        public CreateNewMessageGroupModal(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
            
        }

        // _________________________________________ Constantes _______________________________________________

        private const string TITLE_MESSAGE= "//*[@id=\"messageTitleInput\"]"; 
        private const string CREATE_TITLE_MESSAGE = "//*[@id=\"createMessageBtn\"]";
        private const string SEND_MESSAGE = "//*[@id=\"btnSend\"]";
        private const string MESSAGE_TEXT = "//*[@id=\"txtMsg\"]";
        private const string CLOSE_SEND_MESSAGE = "//*[@id=\"messagesModal\"]/div[1]/div/div[2]/button";

        // _________________________________________ Variables _________________________________________________

        [FindsBy(How = How.XPath, Using = TITLE_MESSAGE)]
        private IWebElement _titlemessage;

        [FindsBy(How = How.XPath, Using = CREATE_TITLE_MESSAGE)]
        private IWebElement _createtitlemessage;

        [FindsBy(How = How.XPath, Using = MESSAGE_TEXT)]
        private IWebElement _messagetext;

        [FindsBy(How = How.XPath, Using = SEND_MESSAGE)]
        private IWebElement _sendmessage;

         [FindsBy(How = How.XPath, Using = CLOSE_SEND_MESSAGE)]
        private IWebElement _close;

        // _________________________________________ Méthodes __________________________________________________

        public void SetTitleMessage(string title)
        {
            _titlemessage = WaitForElementIsVisible(By.XPath(TITLE_MESSAGE));
            _titlemessage.SetValue(ControlType.TextBox, title);

            _createtitlemessage = WaitForElementToBeClickable(By.XPath(CREATE_TITLE_MESSAGE));
            _createtitlemessage.Click();
            WaitForLoad();
        }
        public void SetMessage(string message)
        {
            _messagetext = WaitForElementIsVisible(By.XPath(MESSAGE_TEXT));
            _messagetext.SetValue(ControlType.TextBox, message);

            _sendmessage = WaitForElementToBeClickable(By.XPath(SEND_MESSAGE));
            _sendmessage.Click();

            _close = WaitForElementToBeClickable(By.XPath(CLOSE_SEND_MESSAGE));
            _close.Click();
            WaitForLoad();
        }
    }
}
