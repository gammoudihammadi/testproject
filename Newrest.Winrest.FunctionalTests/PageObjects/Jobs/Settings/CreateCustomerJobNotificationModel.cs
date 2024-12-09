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

namespace Newrest.Winrest.FunctionalTests.PageObjects.Jobs.Settings
{
    public class CreateCustomerJobNotificationModel : PageBase
    {
        private const string MESSAGE_VALIDATOR_EMAIL = "//*[@id=\"modal-1\"]/div/div/div/div/form/div[2]/div[1]/div/span";
        private const string MESSAGE_VALIDATOR_CUSTOMER = "//*[@id=\"modal-1\"]/div/div/div/div/form/div[2]/div[2]/div/span";
        private const string BNT_CREATE = "//*[@id=\"last\"]";
        private const string CLOSE_X = "//*[@id=\"modal-1\"]/div/div/form/div[3]/button[1]";
        private const string EMAIL = "Email";
        private const string CUSTOMER = "drop-down-customers";
        private const string FIRST_CUSTOMER = "//*[@id=\"drop-down-customers\"]/option[2]";
        public CreateCustomerJobNotificationModel(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
            
        }

        [FindsBy(How = How.Id, Using = EMAIL)]
        private IWebElement _email;

        [FindsBy(How = How.Id, Using = FIRST_CUSTOMER)]
        private IWebElement _firstcutomer;

        [FindsBy(How = How.XPath, Using = BNT_CREATE)]
        private IWebElement _btncreate;

        [FindsBy(How = How.XPath, Using = MESSAGE_VALIDATOR_CUSTOMER)]
        private IWebElement _messagevalidatorcustomer;

        [FindsBy(How = How.XPath, Using = MESSAGE_VALIDATOR_EMAIL)]
        private IWebElement _messagevalidatoremail;


        [FindsBy(How = How.Id, Using = CUSTOMER)]
        private IWebElement _cutomer;

        [FindsBy(How = How.XPath, Using = BNT_CREATE)]
        private IWebElement _btn_create;

        public bool GetEmailRequiredMessageNewSageJob()
        {
            if(isElementVisible(By.XPath(MESSAGE_VALIDATOR_EMAIL)))
            {
                return true;
            }
            return false;
        }


        public void CreateCustomerJobNotificationButton()
        {
            _btncreate = WaitForElementIsVisible(By.XPath(BNT_CREATE));
            _btncreate.Click();
            WaitForLoad();
        }

        public bool GetEmailRequiredMessage()
        {
            return isElementVisible(By.XPath(MESSAGE_VALIDATOR_EMAIL)) ? true : false;
        }
        public bool GetCustomerRequiredMessage()
        {
            return isElementVisible(By.XPath(MESSAGE_VALIDATOR_CUSTOMER)) ? true : false;
  
        }
        public JobNotifications FillFieldAddNewCustomerJobNotification(string email, string customer=null)
        {
            _email = WaitForElementIsVisible(By.Id(EMAIL));
            _email.SetValue(ControlType.TextBox,email);
            if (customer != null)
            {
                _cutomer = WaitForElementExists(By.Id(CUSTOMER));
                _cutomer.SetValue(ControlType.DropDownList, customer);
            }
            else
            {
                _firstcutomer = WaitForElementExists(By.XPath(FIRST_CUSTOMER));
                _firstcutomer.Click();
            }
            _btncreate = WaitForElementIsVisible(By.XPath(BNT_CREATE));
            _btncreate.Click();
            WaitForLoad();
            return new JobNotifications(_webDriver, _testContext);
        }
    }
}
