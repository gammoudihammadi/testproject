using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newrest.Winrest.FunctionalTests.Utils;
using DocumentFormat.OpenXml.InkML;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Jobs.Settings
{
    public class GuestMappingsPage : PageBase
    {
        private const string MAPPING = "//*[@id=\"tableListMenu\"]/tbody/tr[1]/td[3]/div/input";
    

        [FindsBy(How = How.XPath, Using = MAPPING)]
        private IWebElement _mapping;

        public GuestMappingsPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        public void EditMapping(string mapping)
        {
            _mapping = WaitForElementIsVisible(By.XPath(MAPPING));
            _mapping.Clear();
            _mapping.SetValue(ControlType.TextBox, mapping);

            WaitForLoad();
        }
        public string GetFirstMapping()
        {
            _mapping = WaitForElementIsVisible(By.XPath(MAPPING));
            return _mapping.GetAttribute("value");
        }
    }
}
