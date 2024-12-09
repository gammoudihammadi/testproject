using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Flights.LpCart
{
    public class LpCartGeneralInformationPage : PageBase
    {

        // ____________________________________ Constantes __________________________________________

        private const string DATE_FROM = "datapicker-edit-lpcart-from";
        private const string DATE_TO = "datapicker-edit-lpcart-to";
        private const string NAME = "tb-new-lpcart-name";
        private const string COMMENT = "Comment";
        private const string AIRCRAFT = "drop-down-suppliers";
        // ____________________________________ Variables ___________________________________________

        [FindsBy(How = How.Id, Using = DATE_FROM)]
        private IWebElement _dateFrom;

        [FindsBy(How = How.Id, Using = DATE_TO)]
        private IWebElement _dateTo;

        [FindsBy(How = How.Id, Using = NAME)]
        private IWebElement _name;

        [FindsBy(How = How.Id, Using = COMMENT)]
        private IWebElement _comment;

        [FindsBy(How = How.Id, Using = AIRCRAFT)]
        private IWebElement _aircraft;


        public LpCartGeneralInformationPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        public void UpdateToDateTo()
        {
            _dateTo = WaitForElementIsVisible(By.Id(DATE_TO));
            _dateTo.SetValue(ControlType.DateTime, DateUtils.Now.AddDays(15));
            _dateTo.SendKeys(Keys.Tab);
            WaitForLoad();
        }

        public void UpdateFromToDate(DateTime from, DateTime to)
        {
            // Dating
            _dateFrom = WaitForElementIsVisible(By.Id(DATE_FROM));
            _dateFrom.SetValue(ControlType.DateTime, from);
            _dateFrom.SendKeys(Keys.Tab);

            Thread.Sleep(TimeSpan.FromSeconds(3));

            _dateTo = WaitForElementIsVisible(By.Id(DATE_TO));
            _dateTo.SetValue(ControlType.DateTime, to);
            _dateTo.SendKeys(Keys.Tab);

            Thread.Sleep(TimeSpan.FromSeconds(3));
        }

        public void UpdateFromDate(DateTime from)
        {
            // Dating
            _dateFrom = WaitForElementIsVisible(By.Id(DATE_FROM));
            _dateFrom.SetValue(ControlType.DateTime, from);
            _dateFrom.SendKeys(Keys.Tab);

            Thread.Sleep(TimeSpan.FromSeconds(3));
        }

        public void UpdateToDate(DateTime to)
        {
            // Dating
            _dateTo = WaitForElementIsVisible(By.Id(DATE_TO));
            _dateTo.SetValue(ControlType.DateTime, to);
            _dateTo.SendKeys(Keys.Tab);

            Thread.Sleep(TimeSpan.FromSeconds(3));
        }


        public List<string> GetSelectedAircrafts()
        {
            List<string> aircraftsSelected = new List<string>();
            var aircraftTags = _webDriver.FindElements(By.XPath("//*[@id=\"selected-aircrafts\"]/div[*]"));

            foreach(var tag in aircraftTags)
            {
                aircraftsSelected.Add(tag.Text);
            }

            return aircraftsSelected;
        }

        public void AddRoute(string route)
        {
            var routes = WaitForElementIsVisible(By.XPath("//*/input[contains(@class,'new-route tt-input')]"));
            routes.SetValue(ControlType.TextBox, route);
            var add = WaitForElementIsVisible(By.XPath("//*/div[contains(@class,'tt-selectable')]"));// //*/a[contains(@class,'add-tag-btn add-route-btn')]"));
            add.Click();
            WaitPageLoading();
            WaitForLoad();
        }

        public void AddAircraf(string aircraft)
        {
            WaitForLoad();
            _aircraft = WaitForElementIsVisible(By.XPath("//*/input[contains(@class,'new-aircraft') and contains(@class,'tt-input')]"));
            _aircraft.SetValue(ControlType.TextBox, aircraft);
            WaitForLoad();
            var choice = WaitForElementIsVisible(By.XPath("//*/div[@class='tt-suggestion tt-selectable']"));
            choice.Click();
            WaitForLoad();
        }

        public LpCartCartDetailPage ClickOnCarts()
        {
            Actions a = new Actions(_webDriver);
            var _cartTab = WaitForElementIsVisible(By.Id("hrefTabContentCarts"));
            a.MoveToElement(_cartTab).Perform();
            _cartTab.Click();
            WaitForLoad();

            return new LpCartCartDetailPage(_webDriver, _testContext);
        }

        public void UpdateLpCartName(string newName)
        {
            _name= WaitForElementIsVisible(By.Id(NAME));
            _name.SetValue(ControlType.TextBox, newName);
            WaitForLoad();
        }

        public void UpdateLpCartComment(string comment)
        {
            ScrollUntilElementIsVisible(By.Id(COMMENT));
            _comment = WaitForElementIsVisible(By.Id(COMMENT));
            _comment.SetValue(ControlType.TextBox, comment);
            Thread.Sleep(2000);
            WaitForLoad();
        }

        public string GetLpCartComment()
        {
            _comment = WaitForElementIsVisible(By.Id(COMMENT));
            return _comment.GetAttribute("value");
        }

        public string GetLpCartName()
        {
            _comment = WaitForElementIsVisible(By.Id(NAME));
            return _comment.GetAttribute("value");
        }
        public void ScrollUntilElementIsVisible(By by)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)_webDriver;
            var element = WaitForElementExists(by);
            if (element != null)
            {
                js.ExecuteScript("arguments[0].scrollIntoView(true);", element);
            }
        }
        public string GetLpCartDateFrom()
        {
            _dateFrom = WaitForElementIsVisible(By.Id(DATE_FROM));
            return _dateFrom.GetAttribute("value");

        }
        public string GetLpCartDateTo()
        {
            return WaitForElementExists(By.XPath("//*[@id=\"datapicker-edit-lpcart-to\"]")).GetAttribute("value");
        }

        public List<string> GetSelectedRoutes()
        {
            List<string> routesSelected = new List<string>();
            var routeTags = _webDriver.FindElements(By.XPath("//*[@id=\"tabContentDetails\"]/div[5]/div[2]/div/div[*]"));
            foreach (var tag in routeTags)
            {
                routesSelected.Add(tag.Text);
            }
            return routesSelected;
        }
    }
}
