using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;
using System.Linq;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Production.CustomerOrder
{
    public class CustomerOrderIMessagesPage : PageBase
    {
        public CustomerOrderIMessagesPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        private const string NEW_MESSAGE_BTN = "//*[@id=\"tabContentDetails\"]/div/div[1]/div/a/span";
        private const string TITLE_INPUT = "//*[@id=\"messageTitleInput\"]";
        private const string BODY_INPUT = "//*[@id=\"txtMsg\"]";
        private const string CREATE_BTN = "//*[@id=\"createMessageBtn\"]";
        private const string SEND_BTN = "//*[@id=\"btnSend\"]";
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";
        private const string DELETE_MESSAGE_BTN = "//button[@class='btn btn-clean deleteBtn']/span";
        private const string VIEW_MESSAGE_BTN = "//*[@id=\"row-38\"]/td[6]/a";
        private const string DATE_LINE = "/html/body/div[3]/div/div[3]/div[2]/div/div/div[2]/table/tbody/tr/td[1]";
        private const string TIME_LINE = "/html/body/div[3]/div/div[3]/div[2]/div/div/div[2]/table/tbody/tr/td[2]";
        private const string CREATED_BY = "/html/body/div[3]/div/div[3]/div[2]/div/div/div[2]/table/tbody/tr/td[3]";
        private const string SUBJECT = "/html/body/div[3]/div/div[3]/div[2]/div/div/div[2]/table/tbody/tr/td[4]";
        private const string SEEN_EYE_ICON = "//td/a/span[contains(@class, 'fas fa-eye')]";
        private const string CLOSE_VIEW_MESSAGE = "//*[@id=\"modal-1\"]/div[1]/div/div[2]/button";
        private const string VIEW_MESSAGE_BODY = "/html/body/div[3]/div/div[3]/div[2]/div/div/div[2]/table/tbody/tr/td[6]/a";
        private const string MESSAGE_ADDED = "/html/body/div[3]/div/div[3]/div[2]/div/div/div[2]/table/tbody/tr/td[4]";
        private const string DELETE_BTN = "/html/body/div[3]/div/div[3]/div[2]/div/div/div[2]/table/tbody/tr[2]/td[7]/button";
        private const string MESSAGE_DISPLAYED = "/html/body/div[3]/div/div[3]/div[2]/div/div/div[2]/table/tbody/tr/td[4]";
        private const string DELETE_MSG_BTN = "deleteBtn";


        //__________________________________ Variables ______________________________________

        [FindsBy(How = How.XPath, Using = NEW_MESSAGE_BTN)]
        private IWebElement _newMessageBtn;

        [FindsBy(How = How.XPath, Using = TITLE_INPUT)]
        private IWebElement _titleInput;

        [FindsBy(How = How.XPath, Using = BODY_INPUT)]
        private IWebElement _bodyInput;

        [FindsBy(How = How.XPath, Using = CREATE_BTN)]
        private IWebElement _createBtn;

        [FindsBy(How = How.XPath, Using = SEND_BTN)]
        private IWebElement _sendBtn;

        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;
        [FindsBy(How = How.XPath, Using = VIEW_MESSAGE_BODY)]
        private IWebElement _viewMessage;

        [FindsBy(How = How.XPath, Using = MESSAGE_ADDED)]
        private IWebElement _subject;

        [FindsBy(How = How.XPath, Using = DELETE_BTN)]
        private IWebElement _deleteBtn;

        public string AddNewMessage(string title, string body)
        {
            _newMessageBtn = WaitForElementIsVisible(By.XPath(NEW_MESSAGE_BTN));
            _newMessageBtn.Click();
            WaitForLoad();

            _titleInput = WaitForElementIsVisible(By.XPath(TITLE_INPUT));
            _titleInput.SetValue(ControlType.TextBox, title);

            _createBtn = WaitForElementIsVisible(By.XPath(CREATE_BTN));
            _createBtn.Click();
            WaitForLoad();

            // l'heure dépend de la zone (ES espagnol comme nous) et de l'horloge de la machine
            DateTime heure = DateTime.ParseExact(DateTime.Now.AddHours(-1).ToString("HH:mm"), "HH:mm", CultureInfo.InvariantCulture);
            string time = heure.ToString("h:mm tt");
            Console.WriteLine("time -2 : " + time);

            _bodyInput = WaitForElementIsVisible(By.XPath(BODY_INPUT));
            _bodyInput.SetValue(ControlType.TextBox, body);

            _sendBtn = WaitForElementIsVisible(By.XPath(SEND_BTN));
            _sendBtn.Click();
            WaitForLoad();

            var _closeBtn = WaitForElementIsVisible(By.XPath("//*[@id=\"messagesModal\"]/div[1]/div/div[2]/button[@class='close']"));
            _closeBtn.Click();
            WaitForLoad();

            return time;
        }

        public CustomerOrderPage BackToList()
        {
            _backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            _backToList.Click();
            WaitForLoad();

            return new CustomerOrderPage(_webDriver, _testContext);
        }

        public bool VerifyMessageAdded(string title)
        {
            if(!isElementVisible(By.XPath(MESSAGE_ADDED)))
            {
               return false;
            }
            else
            {
                var subject = WaitForElementIsVisible(By.XPath(MESSAGE_ADDED));
                if(subject.Text != title)
                {
                    return false;
                }
            }
            return true;
        }

        public bool VerifyMessageData(string title, string body, string date, string time)
        {
            try
            {
                if (!isElementVisible(By.XPath(DATE_LINE)))
                {
                    return false;
                }

                var dateLine = WaitForElementIsVisible(By.XPath(DATE_LINE));
                if (dateLine.Text != date)
                {
                    return false;
                }

                var timeLine = WaitForElementIsVisible(By.XPath(TIME_LINE));
                string timeToCompare = timeLine.Text.Length >= time.Length ? timeLine.Text.Substring(0, time.Length) : timeLine.Text;

                if (timeToCompare != time)
                {
                    return false;
                }

                var createdBy = WaitForElementIsVisible(By.XPath(CREATED_BY));
                if (!createdBy.Text.Contains("Winrest Test Auto"))
                {
                    return false;
                }

                var subject = WaitForElementIsVisible(By.XPath(SUBJECT));
                if (subject.Text != title)
                {
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error verifying message data: {ex.Message}");
                return false;
            }
        }
        public void DeleteMessage()
        {
            var _deleteBtn = WaitForElementIsVisible(By.XPath(DELETE_MESSAGE_BTN));
            _deleteBtn.Click();
            WaitForLoad();
        }
        public void ViewMessage()
        {
            var _viewMessage = WaitForElementIsVisible(By.XPath("//td/a/span[contains(@class, 'fas fa-eye')]"));
            _viewMessage.Click();
            WaitForLoad();
        }
        public void DeleteMessageDetail()
        {
            var _deleteBtn = WaitForElementIsVisible(By.Id(DELETE_MSG_BTN));
            _deleteBtn.Click();
            WaitForLoad();

        }
        public void CloseMessageModal()
        {
            var _closeBtn = WaitForElementIsVisible(By.XPath("//div[1]/div/div[2]/button[@class='close']"));
            _closeBtn.Click();
            WaitForLoad();
        }
        public bool VerifyMessageDetailDeleted()
        {
            if (isElementVisible(By.XPath("//*[@class=\"message bot-message\"]")))
            {
                return false;
            }
            return true;
        }
        public bool VerifyViewMessage()
        {
            if (!isElementVisible(By.XPath("//*[@id=\"messengerContainer\"]")))
            {
                return false;
            }
            return true;
        }

        public bool CheckIfMessageClosedSuccesfully()
        {
            return !isElementVisible(By.XPath("//div[1]/div/div[2]/button[@class='close']"));
        }
        public void ViewMessageSubject()
        {
            WaitPageLoading();
            var viewMessageElements = _webDriver.FindElements(By.XPath(SEEN_EYE_ICON));
            WaitPageLoading();
            for (int i = 0; i < viewMessageElements.Count; i++)
            {
                WaitPageLoading();
                // Re-fetch the list of elements to avoid stale references
                viewMessageElements = _webDriver.FindElements(By.XPath(SEEN_EYE_ICON));
                viewMessageElements[i].Click();
                WaitPageLoading();
                var tabMessage = WaitForElementIsVisible(By.XPath(CLOSE_VIEW_MESSAGE));
                tabMessage.Click();
            }
        }
        public void ViewMessageBody()
        {
            _viewMessage = WaitForElementIsVisible(By.XPath(VIEW_MESSAGE_BODY));
            _viewMessage.Click();
            WaitForLoad();
        }
        public bool VerifyMessageDisplayed(string title)
        {
            if (!isElementVisible(By.XPath(MESSAGE_DISPLAYED)))
            {
                return false;
            }
            else
            {
                var subject = WaitForElementIsVisible(By.XPath(MESSAGE_DISPLAYED));
                if (subject.Text != title)
                {
                    return false;
                }
            }
            return true;
        }




    }
}
