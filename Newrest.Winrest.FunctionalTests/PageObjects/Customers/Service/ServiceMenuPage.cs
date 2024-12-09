using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service
{
    public class ServiceMenuPage : PageBase
    {
        public ServiceMenuPage(IWebDriver _webDriver, TestContext _testContext) : base(_webDriver, _testContext)
        {
        }

        // __________________________________________ Constantes ______________________________________________

        private const string MENU_NAME = "//*[@id=\"tabContentDetails\"]/div/div[2]/div/div/div[1]/span";

        // __________________________________________ Variables _______________________________________________

        [FindsBy(How = How.XPath, Using = MENU_NAME)]
        private IWebElement _menuName;

        // __________________________________________ Méthodes ________________________________________________

        public string GetNameMenu()
        {
            _menuName = WaitForElementIsVisible(By.XPath(MENU_NAME));
            return _menuName.Text;
        }


    }
}
