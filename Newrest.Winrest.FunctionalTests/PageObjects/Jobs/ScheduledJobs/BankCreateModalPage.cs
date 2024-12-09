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

namespace Newrest.Winrest.FunctionalTests.PageObjects.Jobs.ScheduledJobs
{
    public  class BankCreateModalPage : PageBase
    {
        public BankCreateModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {

        }

        private const string NAME = "Name";
        private const string CEGID_JOB_TYPE = "SageJobType";
        private const string CRON_EXPRESSION = "CronExpression";
        private const string FISCAL_ENTITIES = "SelectedFiscalEntities_ms";
        private const string IS_ACTIVE = "//*[@id=\"IsActive\"]";
        private const string CREATE = "last";
        private const string VALIDATION_NAME = "//*[@id=\"modal-1\"]/div/div/div/div/form/div[2]/div[1]/div/span[@class]";
        private const string EDITOR = "btnShowEditor";
        private const string SECOND_TAB = "//*[@id=\"crongenerator\"]/li[1]/a";
        private const string MINUTE_TAB = "//*[@id=\"crongenerator\"]/li[2]/a";
        private const string HOUR_TAB = "//*[@id=\"crongenerator\"]/li[3]/a";
        private const string DAY_TAB = "//*[@id=\"crongenerator\"]/li[4]/a";
        private const string MONTH_TAB = "//*[@id=\"crongenerator\"]/li[5]/a";
        private const string YEAR_TAB = "//*[@id=\"crongenerator\"]/li[6]/a";
        private const string FIRST_JOB_ROW = "//*[@id=\"tableListMenu\"]/tbody/tr[1]";
        private const string SHOW_EDITOR_BUTTON = "btnShowEditor";
        private const string SPECIFIC_HOUR_BUTTON = "cronHourSpecific";
        private const string CRON_HOUR = "cronHour";
        private const string BNT_CREATE = "last";
        private const string CRON_MINUTES = "cronMinute";
        private const string SPECIFIC_MINUTES_BUTTON = "cronMinuteSpecific";




        //_____________________________________ Variables _____________________________________________
        // General
        
        [FindsBy(How = How.Id, Using = CREATE)]
        private IWebElement _save;

        [FindsBy(How = How.Id, Using = EDITOR)]
        private IWebElement _editor;
        [FindsBy(How = How.XPath, Using = FIRST_JOB_ROW)]
        private IWebElement _firstJobRow;
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
       

        public void ClickSave()
        {
            _save = WaitForElementIsVisible(By.Id(CREATE));
            _save.Click();
            WaitForLoad();
        }
        public bool VerifyCreateValidators()
        {
            if (isElementVisible(By.XPath(VALIDATION_NAME)))
            {
                return true;
            }
            else { return false; }
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
            _showEditorButton = WaitForElementIsVisible(By.Id(SHOW_EDITOR_BUTTON));
            _showEditorButton.Click();
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
