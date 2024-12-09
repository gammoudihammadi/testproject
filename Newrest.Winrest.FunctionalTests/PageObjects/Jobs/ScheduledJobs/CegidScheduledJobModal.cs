using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Jobs.ScheduledJobs
{
    public class CegidScheduledJobModal : PageBase
    {
        public CegidScheduledJobModal(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {

        }
        //___________________________________ Constantes _____________________________________________

        private const string NAME = "Name";
        private const string CEGID_JOB_TYPE = "SageJobType";
        private const string CRON_EXPRESSION = "CronExpression";
        private const string FISCAL_ENTITIES = "SelectedFiscalEntities_ms";
        private const string IS_ACTIVE = "IsActive";
        private const string CREATE = "last";
        private const string VALIDATION_NAME = "//*[@id='Name']/..//span[contains(@class, 'text-danger col-md-12 field-validation-error')]";
        private const string VALIDATION_FISCALENTITIES = "//*[@id='SelectedFiscalEntities_ms']/../../../../..//span[contains(@class, 'field-validation-error')]";
        private const string VALIDATION_CRON = "//*[@id='CronExpression']/..//span[contains(@class, 'text-danger col-md-12 field-validation-error')]";
        private const string EDITOR = "btnShowEditor";
        private const string SECOND_TAB = "//*[@id=\"crongenerator\"]/li[1]/a";
        private const string MINUTE_TAB = "//*[@id=\"crongenerator\"]/li[2]/a";
        private const string HOUR_TAB = "//*[@id=\"crongenerator\"]/li[3]/a";
        private const string DAY_TAB = "//*[@id=\"crongenerator\"]/li[4]/a";
        private const string MONTH_TAB = "//*[@id=\"crongenerator\"]/li[5]/a";
        private const string YEAR_TAB = "//*[@id=\"crongenerator\"]/li[6]/a";
        private const string SPECIFIC_HOUR_BUTTON = "cronHourSpecific";
        private const string CRON_HOUR = "cronHour";
        private const string BNT_CREATE = "last";
        private const string CRON_MINUTES = "cronMinute";
        private const string SPECIFIC_MINUTES_BUTTON = "cronMinuteSpecific";

        //_____________________________________ Variables _____________________________________________
        // General
        [FindsBy(How = How.Id, Using = NAME)]
        private IWebElement _name;

        [FindsBy(How = How.Id, Using = CEGID_JOB_TYPE)]
        private IWebElement _cegidJobType;

        [FindsBy(How = How.Id, Using = CRON_EXPRESSION)]
        private IWebElement _cronExpression;

        [FindsBy(How = How.Id, Using = FISCAL_ENTITIES)]
        private IWebElement _fiscalEntities;

        [FindsBy(How = How.Id, Using = IS_ACTIVE)]
        private IWebElement _isActive;

        [FindsBy(How = How.Id, Using = CREATE)]
        private IWebElement _save;

        [FindsBy(How = How.Id, Using = EDITOR)]
        private IWebElement _editor;

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
        [FindsBy(How = How.Id, Using = SPECIFIC_HOUR_BUTTON)]
        private IWebElement _specificHourButton;
        [FindsBy(How = How.Id, Using = CRON_HOUR)]
        private IWebElement _cronHour;
        [FindsBy(How = How.XPath, Using = BNT_CREATE)]
        private IWebElement _btn_create;
        [FindsBy(How = How.Id, Using = CRON_MINUTES)]
        private IWebElement _cronMinutes;
        [FindsBy(How = How.Id, Using = SPECIFIC_MINUTES_BUTTON)]
        private IWebElement _specificMinutesButton;

        public void FillFieldsModalCegidScheduledJob(string name, string fiscalEntites = null, bool isActive = true, string cronExpression = null, string cegidJobType = "WRSAGEINVENTORY")
        {
            _name = WaitForElementIsVisible(By.Id(NAME));
            _name.SetValue(PageBase.ControlType.TextBox, name);
            WaitForLoad();

            _cegidJobType = WaitForElementIsVisible(By.Id(CEGID_JOB_TYPE));
            _cegidJobType.SetValue(PageBase.ControlType.DropDownList, cegidJobType);
            WaitForLoad();
            if (cronExpression != null)
            {
                _cronExpression = WaitForElementIsVisible(By.Id(CRON_EXPRESSION));
                _cronExpression.SetValue(PageBase.ControlType.TextBox, cronExpression);
                WaitForLoad();
            }
            if (fiscalEntites != null)
            {
                ComboBoxSelectById(new ComboBoxOptions(FISCAL_ENTITIES, (string)fiscalEntites, false));

            }

            _isActive = WaitForElementExists(By.Id(IS_ACTIVE));
            _isActive.SetValue(PageBase.ControlType.CheckBox, isActive);
            WaitForLoad();

            _save = WaitForElementIsVisible(By.Id(CREATE));
            _save.Click();
            WaitForLoad();
        }
        public bool VerifyCreateValidators()
        {
            if (isElementVisible(By.XPath(VALIDATION_NAME))
                && isElementVisible(By.XPath(VALIDATION_CRON))
                && isElementVisible(By.XPath(VALIDATION_FISCALENTITIES)))
            {
                return true;
            } 
            return false;
        }
        public void EditorCronExpression()
        {
            _editor = WaitForElementIsVisible(By.Id(EDITOR));
            _editor.Click();

        }
        public bool CheckTabsExistance()
        {
       
            return isElementVisible(By.XPath(SECOND_TAB)) && isElementVisible(By.XPath(MINUTE_TAB))
                && isElementVisible(By.XPath(HOUR_TAB)) && isElementVisible(By.XPath(DAY_TAB))
                && isElementVisible(By.XPath(MONTH_TAB)) && isElementVisible(By.XPath(YEAR_TAB));
        }
        public void ClickOnEditorButton()
        {
            _editor = WaitForElementIsVisible(By.Id(EDITOR));
            _editor.Click();
            WaitForLoad();
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
        public void EditHourCronExpression()
        {
            _specificHourButton = WaitForElementIsVisible(By.Id(SPECIFIC_HOUR_BUTTON));
            _specificHourButton.Click();
            WaitForLoad();
            _cronHour = WaitForElementIsVisible(By.Id($"{CRON_HOUR}5"));
            _cronHour.Click();
            WaitForLoad();
            _cronHour = WaitForElementIsVisible(By.Id($"{CRON_HOUR}0"));
            _cronHour.Click();
            WaitForLoad();
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
            _btn_create = WaitForElementIsVisible(By.Id(BNT_CREATE));
            _btn_create.Click();
            WaitForLoad();
        }

    }
}
