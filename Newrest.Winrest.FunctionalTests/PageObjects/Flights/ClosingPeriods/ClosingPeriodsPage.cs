using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using System;
using System.Globalization;
using System.Linq;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Flights.ClosingPeriods
{
    public class ClosingPeriodsPage : PageBase
    {
        public ClosingPeriodsPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // __________________________________________ Constantes _________________________________________
        public const string EXTEND_MENU = "/html/body/div[2]/div/div[2]/div/div[1]/div/div/div/button";
        public const string CREATE_PERIOD_BUTTON = "/html/body/div[2]/div/div[2]/div/div[1]/div/div/div/div/a";
        public const string START_DATE_CREATE = "/html/body/div[3]/div/div/div/div/form/div[2]/div/div/div[1]/div/div/input";
        public const string END_DATE_CREATE = "/html/body/div[3]/div/div/div/div/form/div[2]/div/div/div[2]/div/div/input";
        public const string SITE_CREATE = "/html/body/div[3]/div/div/div/div/form/div[2]/div/div/div[3]/div/select";
        public const string SAVE_CREATE_BUTTON = "/html/body/div[3]/div/div/div/div/form/div[2]/div/div/div[4]/button[2]";
        public const string START_DATE_CLICK = "/html/body/div[3]/div/div/div/div/form/div[2]/div/div/div[2]/div";
        public const string LIST_CLOSING_PERIODS = "/html/body/div[2]/div/div[2]/div/div[2]/div";
        public const string END_DATE_FILTER = "/html/body/div[2]/div/div[1]/div[2]/div/form/div[3]/div/div/input";
        public const string START_DATE_FILTER = "/html/body/div[2]/div/div[1]/div[2]/div/form/div[2]/div/div/input";
        public const string FILTER_SITES = "/html/body/div[2]/div/div[1]/div[2]/div/form/div[4]/select";
        public const string START_DATES_GRID = "/html/body/div[2]/div/div[2]/div/div[2]/div/table/tbody/tr[*]/td[2]";
        public const string END_DATES_GRID = "/html/body/div[2]/div/div[2]/div/div[2]/div/table/tbody/tr[*]/td[3]";
        // __________________________________________ Méthodes ___________________________________________

        public enum FilterType
        {
            Startdate,
            Enddate,
            Site
        }

        public void Filter(FilterType filterType, object value)

        {
            switch (filterType)
            {
                case FilterType.Startdate:

                     var startDateFilter = WaitForElementIsVisible(By.XPath(START_DATE_FILTER));
                    startDateFilter.SetValue(ControlType.DateTime, value);
                    startDateFilter.SendKeys(Keys.Tab);
                    break;

                case FilterType.Enddate:
                    var endDateFilter = WaitForElementIsVisible(By.XPath(END_DATE_FILTER));
                    endDateFilter.SetValue(ControlType.DateTime, value);
                    endDateFilter.SendKeys(Keys.Tab);
                    break;

                case FilterType.Site:
                    ComboBoxSelectById(new ComboBoxOptions(FILTER_SITES, (string)value));
                    break;

                default:
                    break;

            }
            WaitPageLoading();
            WaitForLoad();
        }

        public void SetStartDate(DateTime date)
        {
            var startDate = WaitForElementIsVisible(By.XPath(START_DATE_CREATE));
            startDate.SetValue(ControlType.DateTime, date);
        }

        public void SetEndDate(DateTime date)
        {
            var endDate = WaitForElementIsVisible(By.XPath(END_DATE_CREATE));
            endDate.SetValue(ControlType.DateTime, date);
        }

        public void SetSite(string siteInput)
        {
            var site = WaitForElementIsVisible(By.XPath(SITE_CREATE));
            site.SendKeys(siteInput);
        }

        public void CreateNewClosingPeriod(DateTime startDate, DateTime endDate, string site)
        {
            var extendMenu = WaitForElementIsVisible(By.XPath(EXTEND_MENU));
            Actions actions = new Actions(_webDriver);
            actions.MoveToElement(extendMenu).Perform();
            var createPeriodBtn = WaitForElementIsVisible(By.XPath(CREATE_PERIOD_BUTTON));
            WaitForLoad();
            createPeriodBtn.Click();
            SetStartDate(startDate);
            var startDateClick = WaitForElementIsVisible(By.XPath(START_DATE_CLICK));
            startDateClick.Click();
            SetEndDate(endDate);
            SetSite(site);
            var saveBtn = WaitForElementIsVisible(By.XPath(SAVE_CREATE_BUTTON));
            saveBtn.Click();

            WaitForLoad();
        }

        public bool VerifyClosingPeriodCreated()
        {
            var listClosingPeriods = _webDriver.FindElements(By.XPath(LIST_CLOSING_PERIODS));
            if(listClosingPeriods.Count() == 0)
            {
                return false;
            }
            return true;
        }
        public bool IsStartEndDateRespected(DateTime startDate, DateTime endDate, string dateFormat)
        {
            CultureInfo ci = dateFormat.Equals("dd/MM/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            var startDates = _webDriver.FindElements(By.XPath(START_DATES_GRID));
            var endDates = _webDriver.FindElements(By.XPath(END_DATES_GRID));
            Assert.AreEqual(startDates.Count(), endDates.Count());
            if (startDates.Count == 0 )
                return false;

            for (int i= 0;i<startDates.Count;i++)
            {
                try
                {
                    DateTime dateExclusionStart = DateTime.Parse(startDates[i].Text, new CultureInfo("en-US")).Date;
                    DateTime dateExclusionEnd = DateTime.Parse(endDates[i].Text, new CultureInfo("en-US")).Date;

                    if (DateTime.Compare(dateExclusionEnd, startDate) < 0)
                    {
                        return false;
                    }
                    else if (DateTime.Compare(dateExclusionStart, endDate) > 0)
                    {
                        return false;
                    }
                    else
                    {
                        continue;
                    }
                }
                catch
                {
                    return false;
                }
            }

            return true;
        }
    }
}
