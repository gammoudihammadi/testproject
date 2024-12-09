using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Jobs.Settings
{
    public  class CegidPanel : PageBase
    {
        private const string SEARCH_BY_ORIGIN_TABLE = "SagePanelFiltersSelectedOriginTables_ms";
        private const string BUTTON_SELECTED = "#SagePanelFiltersSelectedOriginTables_ms.ui-state-active";
        //private const string CEGID_PANEL = "//*[@id=\"nav-tab\"]/li[7]/a";

        [FindsBy(How = How.Id, Using = SEARCH_BY_ORIGIN_TABLE)]
        private IWebElement _searchbyorigintable;

        [FindsBy(How = How.Id, Using = BUTTON_SELECTED)]
        private IWebElement _searchbyorigintable_buttonselected;
        public CegidPanel(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }
        public void searchbyoriginTable()
        {
            var filterButton = WaitForElementIsVisible(By.Id(SEARCH_BY_ORIGIN_TABLE));
            if (filterButton.GetAttribute("class").Contains("ui-state-active"))
            {
                Console.WriteLine("Button is already selected.");
            }
            else
            {
                filterButton.Click();
                WaitForElementIsVisible(By.CssSelector(BUTTON_SELECTED));
                Console.WriteLine("Button has been selected.");
            }
        }

       
    }

    }

