using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Flights.TimeBlock
{
    public class BulkChangeModal : PageBase
    {
        public BulkChangeModal(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }
        public void Fill(string ensamblajeStartDay, string ensamblajeStart, string ensamblajeEndDay, string ensamblajeEnd, string corteStartDay, string corteStart, string corteEndDay, string corteEnd, bool major = false)
        {
            var _ensamblajeStartDay = WaitForElementIsVisible(By.Id("Workshops_0__StartDayBeforeEtd"));
            _ensamblajeStartDay.SetValue(ControlType.TextBox, ensamblajeStartDay);

            var _ensamblajeStart = WaitForElementIsVisible(By.Id("Workshops_0__StartBeforeEtd"));
            _ensamblajeStart.SendKeys(ensamblajeStart);

            var _ensamblajeEndDay= WaitForElementIsVisible(By.Id("Workshops_0__EndDayBeforeEtd"));
            _ensamblajeEndDay.SetValue(ControlType.TextBox, ensamblajeEndDay);

            var _ensamblajeEnd = WaitForElementIsVisible(By.Id("Workshops_0__EndBeforeEtd"));
            _ensamblajeEnd.SendKeys(ensamblajeEnd);

            var _corteStartDay = WaitForElementIsVisible(By.Id("Workshops_1__StartDayBeforeEtd"));
            _corteStartDay.SetValue(ControlType.TextBox, corteStartDay);

            var _corteStart = WaitForElementIsVisible(By.Id("Workshops_1__StartBeforeEtd"));
            _corteStart.SendKeys(corteStart);

            var _corteEndDay = WaitForElementIsVisible(By.Id("Workshops_1__EndDayBeforeEtd"));
            _corteEndDay.SetValue(ControlType.TextBox, corteEndDay);

            var _corteEnd = WaitForElementIsVisible(By.Id("Workshops_1__EndBeforeEtd"));
            _corteEnd.SendKeys(corteEnd);

            if (major)
            {
                var _major = WaitForElementExists(By.Id("Major"));
                _major.SetValue(ControlType.CheckBox, major);
            }
            WaitForLoad();
        }

        public void Submit()
        {
            var _validate = WaitForElementIsVisible(By.Id("validateButton"));
            _validate.Click();
            WaitPageLoading();
        }

        public void CheckFirstFlight(string ensamblajeStartDay, string ensamblajeStart, string ensamblajeEndDay, string ensamblajeEnd, string corteStartDay, string corteStart, string corteEndDay, string corteEnd, bool major = false)
        {
            var eStartDay = WaitForElementIsVisible(By.XPath("(//*[@id=\"workshopValue_WorkshopDayStart\"])[1]"));
            Assert.AreEqual(ensamblajeStartDay, eStartDay.GetAttribute("value"));

            var eStart = WaitForElementIsVisible(By.XPath("(//*[@id=\"workshopValue_WorkshopDateStart\"])[1]"));
            Assert.AreEqual(ensamblajeStart, eStart.GetAttribute("value"));
            Assert.IsTrue(eStart.GetAttribute("class").Contains("edited-input"), "pas de couleur");

            var eEndDay = WaitForElementIsVisible(By.XPath("(//*[@id=\"workshopValue_WorkshopDayEnd\"])[1]"));
            Assert.AreEqual(ensamblajeEndDay, eEndDay.GetAttribute("value"));

            var eEnd = WaitForElementIsVisible(By.XPath("(//*[@id=\"workshopValue_WorkshopDateEnd\"])[1]"));
            Assert.AreEqual(ensamblajeEnd, eEnd.GetAttribute("value"));
            Assert.IsTrue(eEnd.GetAttribute("class").Contains("edited-input"), "pas de couleur");

            var cStartDay = WaitForElementIsVisible(By.XPath("(//*[@id=\"workshopValue_WorkshopDayStart\"])[2]"));
            Assert.AreEqual(corteStartDay, cStartDay.GetAttribute("value"));

            var cStart = WaitForElementIsVisible(By.XPath("(//*[@id=\"workshopValue_WorkshopDateStart\"])[2]"));
            Assert.AreEqual(corteStart, cStart.GetAttribute("value"));
            Assert.IsTrue(cStart.GetAttribute("class").Contains("edited-input"), "pas de couleur");

            var cEndDay = WaitForElementIsVisible(By.XPath("(//*[@id=\"workshopValue_WorkshopDayEnd\"])[2]"));
            Assert.AreEqual(corteEndDay, cEndDay.GetAttribute("value"));

            var cEnd = WaitForElementIsVisible(By.XPath("(//*[@id=\"workshopValue_WorkshopDateEnd\"])[2]"));
            Assert.AreEqual(corteEnd, cEnd.GetAttribute("value"));
            Assert.IsTrue(cEnd.GetAttribute("class").Contains("edited-input"), "pas de couleur");

            var majorCheckBox = WaitForElementExists(By.Id("Items_Items_0__Important"));
            Assert.AreEqual(major, majorCheckBox.Selected);
        }
    }
}
