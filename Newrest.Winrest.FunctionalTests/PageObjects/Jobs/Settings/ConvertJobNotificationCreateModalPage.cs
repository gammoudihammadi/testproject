using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Jobs.Settings
{
    public class ConvertJobNotificationCreateModalPage : PageBase
    {
        private const string EMAIL = "Email";
        private const string FIRST_CUSTOMER = "//*[@id=\"drop-down-customers\"]/option[2]";
        private const string BTN_ADD_CONVERTERS = "last";
        private const string CONVERTERS_COMBO = "SelectedConverterIds_ms";
        private const string EMAIL_REQUIRED = "//*[@id=\"item-edit-form\"]/div[2]/div[1]/div/span";
        private const string CUSTOMER_REQUIRED = "//*[@id=\"SelectedCustomer-Group\"]/div/span";
        private const string CUSTOMER = "drop-down-customers";
        private const string BNT_CREATE = "last";
        private const string SEARCH = "/html/body/div[13]/div/div/label/input";
        private const string UNSELECT_ALL = "/html/body/div[13]/div/ul/li[2]/a/span[2]";
        private readonly string EMAIL_VALIDATOR = "//*[@id=\"item-edit-form\"]/div[2]/div[1]/div/span";
        private readonly string CUSTOMER_VALIDATOR = "//*[@id=\"SelectedCustomer-Group\"]/div/span";

        [FindsBy(How = How.Id, Using = EMAIL)]
        private IWebElement _email;

        [FindsBy(How = How.Id, Using = FIRST_CUSTOMER)]
        private IWebElement _firstcutomer;

        [FindsBy(How = How.Id, Using = BTN_ADD_CONVERTERS)]
        private IWebElement _btnaddconverters;

        [FindsBy(How = How.XPath, Using = EMAIL_REQUIRED)]
        private IWebElement _emailrequired;

        [FindsBy(How = How.XPath, Using = CUSTOMER_REQUIRED)]
        private IWebElement _customerrequired;

        [FindsBy(How = How.Id, Using = CUSTOMER)]
        private IWebElement _cutomer;

        [FindsBy(How = How.Id, Using = CONVERTERS_COMBO)]
        private IWebElement _converter;

        [FindsBy(How = How.Id, Using = BNT_CREATE)]
        private IWebElement _btncreate;
        public ConvertJobNotificationCreateModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }
        public JobNotifications FillField_CreateNewCONVERTERSJOBNOTIFICATION( string Email , string customer)
        {
            _email = WaitForElementIsVisible(By.Id(EMAIL));
            _email.SetValue(ControlType.TextBox, Email);
            _cutomer = WaitForElementExists(By.Id(CUSTOMER));
            _cutomer.SetValue(ControlType.DropDownList, customer);
            WaitForLoad();
            ComboBoxSelectById(new ComboBoxOptions(CONVERTERS_COMBO, "AC_HR_txt (.csv)", false));
            WaitForLoad();

            return new JobNotifications(_webDriver, _testContext);
        }
    
        public void BTN_add_convertersjobnotification()
        {
            _btnaddconverters = WaitForElementExists(By.Id(BTN_ADD_CONVERTERS));
            _btnaddconverters.Click();
            WaitForLoad();
        }
        public JobNotifications FillField_AddNewCONVERTERSJOBNOTIFICATION(string Email, string customer=null, string converter=null)
        {
            _email = WaitForElementIsVisible(By.Id(EMAIL));
            _email.SetValue(ControlType.TextBox, Email);
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

            if(converter != null)
            {
                ComboBoxSelectById(new ComboBoxOptions(CONVERTERS_COMBO, converter, false));
            }
            _btncreate = WaitForElementIsVisible(By.Id(BNT_CREATE));
            _btncreate.Click();
            WaitForLoad();
             return new JobNotifications(_webDriver, _testContext);

            }

        public bool IsValidatorsExists()
        {
            if (isElementVisible(By.XPath(CUSTOMER_VALIDATOR)) && isElementVisible(By.XPath(EMAIL_VALIDATOR))) return true;
            return false;
        }
        }
    }

