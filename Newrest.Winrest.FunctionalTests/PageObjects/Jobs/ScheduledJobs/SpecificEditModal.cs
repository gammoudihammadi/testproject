using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Jobs.ScheduledJobs
{
    public class SpecificEditModal : PageBase
    {
        private const string SHOW_EDITOR_BUTTON = "btnShowEditor";
        private const string SECOND_TAB = "//*[@id=\"crongenerator\"]/li[1]/a";
        private const string MINUTE_TAB = "//*[@id=\"crongenerator\"]/li[2]/a";
        private const string HOUR_TAB = "//*[@id=\"crongenerator\"]/li[3]/a";
        private const string DAY_TAB = "//*[@id=\"crongenerator\"]/li[4]/a";
        private const string MONTH_TAB = "//*[@id=\"crongenerator\"]/li[5]/a";
        private const string YEAR_TAB = "//*[@id=\"crongenerator\"]/li[6]/a";
        private const string SPECIFIC_HOUR_RADIO_BUTTON = "cronHourSpecific";
        private const string CRON_HOUR = "cronHour";
        private const string BNT_CREATE = "last";
        private const string CRON_MINUTES = "cronMinute";
        private const string SPECIFIC_MINUTES_BUTTON = "cronMinuteSpecific";


        //_____________________________________ Variables _____________________________________________
        // General
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

        [FindsBy(How = How.XPath, Using = BNT_CREATE)]
        private IWebElement _btn_create;

        [FindsBy(How = How.Id, Using = CRON_MINUTES)]
        private IWebElement _cronMinutes;

        [FindsBy(How = How.Id, Using = SPECIFIC_MINUTES_BUTTON)]
        private IWebElement _specificMinutesButton;


        public SpecificEditModal(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
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
        }

        public bool CheckTabsExistance()
        {
            return isElementVisible(By.XPath(SECOND_TAB)) && isElementVisible(By.XPath(MINUTE_TAB))
                && isElementVisible(By.XPath(HOUR_TAB)) && isElementVisible(By.XPath(DAY_TAB))
                && isElementVisible(By.XPath(MONTH_TAB)) && isElementVisible(By.XPath(YEAR_TAB));
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
