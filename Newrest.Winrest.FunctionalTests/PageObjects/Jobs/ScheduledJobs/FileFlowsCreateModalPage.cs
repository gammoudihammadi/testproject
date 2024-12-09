using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Jobs.ScheduledJobs
{
    public class FileFlowsCreateModalPage : PageBase
    {
        private const string CLOSE_X = "//*[@id=\"item-edit-form\"]/div[1]/button/span";
        private const string CUSTOMER = "drop-down-customers";
        private const string CONVERTER = "drop-down-converters";
        private const string BNT_CREATE = "last";
        private const string DIRECTION = "Direction";
        private const string TYPE = "FileFlowTypeId";
        private const string PROVIDER = "FileFlowProviderId";
        private const string CRON_EXPRESSION = "CronExpression";
        private const string IS_ACTIVE = "/html/body/div[3]/div/div/div/div/form/div[2]/div[8]/div/div";
        private const string NAME = "Name";
        private const string VALIDATOR_NAME = "//*[@id=\"item-edit-form\"]/div[2]/div[1]/div/span[contains(@class,'text-danger')]";
        private const string VALIDATOR_CUSTOMER = "//*[@id=\"item-edit-form\"]/div[2]/div[2]/div/span[contains(@class,'text-danger')]";
        private const string VALIDATOR_CONVERTER = "//*[@id=\"item-edit-form\"]/div[2]/div[3]/div/span[contains(@class,'text-danger')]";
        private const string VALIDATOR_CRON = "//*[@id='CronExpression']/..//span[contains(@class, 'text-danger col-md-12 field-validation-error')]";

        private const string SHOW_EDITOR_BUTTON = "btnShowEditor";
        private const string SECOND_TAB = "//*[@id=\"crongenerator\"]/li[1]/a";
        private const string MINUTE_TAB = "//*[@id=\"crongenerator\"]/li[2]/a";
        private const string HOUR_TAB = "//*[@id=\"crongenerator\"]/li[3]/a";
        private const string DAY_TAB = "//*[@id=\"crongenerator\"]/li[4]/a";
        private const string MONTH_TAB = "//*[@id=\"crongenerator\"]/li[5]/a";
        private const string YEAR_TAB = "//*[@id=\"crongenerator\"]/li[6]/a";
        private const string SPECIFIC_HOUR_RADIO_BUTTON = "cronHourSpecific";
        private const string CRON_HOUR = "cronHour";
        private const string CRON_MINUTES = "cronMinute";
        private const string SPECIFIC_MINUTES_BUTTON = "cronMinuteSpecific";
        public FileFlowsCreateModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {

        }

        [FindsBy(How = How.Id, Using = NAME)]
        private IWebElement _name;

        [FindsBy(How = How.Id, Using = CUSTOMER)]
        private IWebElement _customer;

        [FindsBy(How = How.Id, Using = CONVERTER)]
        private IWebElement _converter;

        [FindsBy(How = How.XPath, Using = BNT_CREATE)]
        private IWebElement _btn_create;

        [FindsBy(How = How.XPath, Using = DIRECTION)]
        private IWebElement _direction;

        [FindsBy(How = How.XPath, Using = TYPE)]
        private IWebElement _type;

        [FindsBy(How = How.XPath, Using = PROVIDER)]
        private IWebElement _provider;

        [FindsBy(How = How.XPath, Using = CRON_EXPRESSION)]
        private IWebElement _cronExpression;

        [FindsBy(How = How.Id, Using = IS_ACTIVE)]
        private IWebElement _is_Active;

        [FindsBy(How = How.Id, Using = VALIDATOR_NAME)]
        private IWebElement _validatorName;

        [FindsBy(How = How.Id, Using = VALIDATOR_CUSTOMER)]
        private IWebElement _validatorCustomer;

        [FindsBy(How = How.Id, Using = VALIDATOR_CONVERTER)]
        private IWebElement _validatorConverter;

        [FindsBy(How = How.Id, Using = SHOW_EDITOR_BUTTON)]
        private IWebElement _showEditorButton;
        [FindsBy(How = How.XPath, Using = SECOND_TAB)]
        private IWebElement _secondTab;
        [FindsBy(How = How.XPath, Using = MINUTE_TAB)]
        private IWebElement _minuteTab;
        [FindsBy(How = How.XPath, Using = HOUR_TAB)]
        private IWebElement _hoursTab;
        [FindsBy(How = How.XPath, Using = DAY_TAB)]
        private IWebElement _dayTab;
        [FindsBy(How = How.XPath, Using = MONTH_TAB)]
        private IWebElement _monthTab;
        [FindsBy(How = How.XPath, Using = YEAR_TAB)]
        private IWebElement _yearTab;
        [FindsBy(How = How.Id, Using = SPECIFIC_HOUR_RADIO_BUTTON)]
        private IWebElement _specificHourRadioButton;
        [FindsBy(How = How.Id, Using = CRON_HOUR)]
        private IWebElement _cronHour;
        [FindsBy(How = How.Id, Using = CRON_MINUTES)]
        private IWebElement _cronMinutes;
        [FindsBy(How = How.Id, Using = SPECIFIC_MINUTES_BUTTON)]
        private IWebElement _specificMinutesButton;
        public FileFlowsCreateModalPage FillFiled_CreateFileFlows(string name, string customer, object value, string type = null, string converter = null, string cronExpression = null)
        {
            var _name = WaitForElementIsVisible(By.Id(NAME));
            _name.SetValue(ControlType.TextBox, name);
            WaitForLoad();

            var _customer = WaitForElementIsVisible(By.Id(CUSTOMER));
            _customer.SetValue(ControlType.DropDownList, customer);
            WaitForLoad();

            if (type != null)
            {
                _type = WaitForElementIsVisible(By.Id(TYPE));
                _type.SetValue(PageBase.ControlType.DropDownList, type);
                WaitForLoad();
            }

            if (converter != null)
            {
                var _converter = WaitForElementIsVisible(By.Id("drop-down-converters"));
                _converter.SetValue(ControlType.DropDownList, converter+" (.csv)");
                WaitForLoad();
            }

            if (cronExpression != null)
            {
                var cron = WaitForElementIsVisible(By.Id("CronExpression"));
                cron.SetValue(ControlType.TextBox, cronExpression);
                WaitForLoad();
            }
            _is_Active = WaitForElementIsVisible(By.XPath(IS_ACTIVE));
            _is_Active.SetValue(ControlType.CheckBox, value);
            WaitForLoad();

            _btn_create = WaitForElementToBeClickable(By.Id(BNT_CREATE));
            _btn_create.Click();
            WaitForLoad();
            return new FileFlowsCreateModalPage(_webDriver, _testContext);

        }
        public bool VerifyCreateValidators()
        {
            var _validatorName = WaitForElementIsVisible(By.XPath(VALIDATOR_NAME));
            var _validatorCustomer = WaitForElementIsVisible(By.XPath(VALIDATOR_CUSTOMER));
            var _validatorConverter = WaitForElementIsVisible(By.XPath(VALIDATOR_CONVERTER));
            var _validatorCron = WaitForElementIsVisible(By.XPath(VALIDATOR_CRON));

            if (_validatorName.Displayed && !string.IsNullOrEmpty(_validatorName.Text) &&
                _validatorCustomer.Displayed && !string.IsNullOrEmpty(_validatorCustomer.Text) &&
                _validatorConverter.Displayed && !string.IsNullOrEmpty(_validatorConverter.Text)
                && _validatorCron.Displayed && !string.IsNullOrEmpty(_validatorCron.Text))
            {
                return true;
            }
            return false;
        }

        public void ClickOnEditorButton()
        {
            _showEditorButton = WaitForElementIsVisible(By.Id(SHOW_EDITOR_BUTTON));
            _showEditorButton.Click();
            WaitForLoad();
        }
        public void EditHourCronExpression()
        {
            _specificHourRadioButton = WaitForElementIsVisible(By.Id(SPECIFIC_HOUR_RADIO_BUTTON));
            _specificHourRadioButton.Click();
            WaitForLoad();
            _cronHour = WaitForElementIsVisible(By.Id($"{CRON_HOUR}3"));
            _cronHour.Click();
            WaitForLoad();
            _cronHour = WaitForElementIsVisible(By.Id($"{CRON_HOUR}0"));
            _cronHour.Click();
            WaitForLoad();
            _btn_create = WaitForElementIsVisible(By.Id(BNT_CREATE));
            _btn_create.Click();
            WaitForLoad() ;
        }

        public void EditMinutesCronExpression()
        {
            _specificMinutesButton = WaitForElementIsVisible(By.Id(SPECIFIC_MINUTES_BUTTON));
            _specificMinutesButton.Click();
            WaitForLoad();
            _cronMinutes = WaitForElementIsVisible(By.Id($"{CRON_MINUTES}2"));
            _cronMinutes.Click();
            WaitForLoad();
            _cronMinutes = WaitForElementIsVisible(By.Id($"{CRON_MINUTES}0"));
            _cronMinutes.Click();
            WaitForLoad();
        }
        public string GetCronExpression()
        {
            _cronExpression = WaitForElementIsVisible(By.Id(CRON_EXPRESSION));
            return _cronExpression.Text;
        }
        public void SelectTab(Tabs tab)
        {
            switch (tab)
            {
                case Tabs.SECONDS:
                    _secondTab = WaitForElementIsVisible(By.XPath(SECOND_TAB));
                    _secondTab.Click();
                    break;
                case Tabs.MINUTES:
                    _minuteTab = WaitForElementIsVisible(By.XPath(MINUTE_TAB));
                    _minuteTab.Click();
                    break;
                case Tabs.HOURS:
                    _hoursTab = WaitForElementIsVisible(By.XPath(HOUR_TAB));
                    _hoursTab.Click();
                    break;
                case Tabs.DAYS:
                    _dayTab = WaitForElementIsVisible(By.XPath(DAY_TAB));
                    _dayTab.Click();
                    break;
                case Tabs.MONTHS:
                    _monthTab = WaitForElementIsVisible(By.XPath(MONTH_TAB));
                    _monthTab.Click();
                    break;
                case Tabs.YEAR:
                    _yearTab = WaitForElementIsVisible(By.XPath(YEAR_TAB));
                    _yearTab.Click();
                    break;
                default:
                    _secondTab = WaitForElementIsVisible(By.XPath(SECOND_TAB));
                    _secondTab.Click();
                    break;
            }
        }
        public enum Tabs
        {
            SECONDS,
            MINUTES,
            HOURS,
            DAYS,
            MONTHS,
            YEAR
        }
    }
}
