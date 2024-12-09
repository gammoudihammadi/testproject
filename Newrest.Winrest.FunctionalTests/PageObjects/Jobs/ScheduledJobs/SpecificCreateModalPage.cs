using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Jobs.ScheduledJobs
{
    public class SpecificCreateModalPage : PageBase
    {

        private const string CLOSE_X = "//*[@id=\"modal-1\"]/div/div/form/div[5]/button[1]";
        private const string BNT_CREATE = "last";
        private const string NAME = "Name";
        //private const string IS_ACTIVE = "/html/body/div[3]/div/div/div/div/form/div[2]/div[4]/div/div/input";
        private const string IS_ACTIVE = "IsActive";
        public SpecificCreateModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }
        [FindsBy(How = How.Id, Using = IS_ACTIVE)]
        private IWebElement _is_Active;

        [FindsBy(How = How.Id, Using = BNT_CREATE)]
        private IWebElement _btn_create;

        [FindsBy(How = How.Id, Using = NAME)]
        private IWebElement _name;

        public ScheduledSpecificPage FillFiled_CreateSpecific(string name, object value)
        {
            var _name = WaitForElementIsVisible(By.Id(NAME));
            _name.SetValue(ControlType.TextBox, name);

            _is_Active = WaitForElementExists(By.Id(IS_ACTIVE));
            _is_Active.SetValue(ControlType.CheckBox, value);

            _btn_create = WaitForElementToBeClickable(By.Id(BNT_CREATE));
            _btn_create.Click();
            WaitForLoad();
            return new ScheduledSpecificPage(_webDriver, _testContext);

        }
    }
}
