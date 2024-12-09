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
    public class ServiceDeliveryPage : PageBase
    {
        public ServiceDeliveryPage(IWebDriver _webDriver, TestContext _testContext) : base(_webDriver, _testContext)
        {
        }

        // _________________________________________ Constantes _________________________________________________

        private const string DELIVERY_NAME = "//*[@id=\"tabContentDetails\"]/div/div[2]/div/div/div[1]/span";
        public const string CRAYON = "//*[@id=\"tabContentDetails\"]/div/div[2]/div/div/div[5]/div/div[2]/a";
        public const string CHECK_DELIVERIES_EXIST = "//*[@id=\"tabContentDetails\"]/div/div[2]/div/div";



        // _________________________________________ Variables __________________________________________________

        [FindsBy(How = How.XPath, Using = DELIVERY_NAME)]
        private IWebElement _deliveryName;

        [FindsBy(How = How.XPath, Using = CRAYON)]
        private IWebElement _crayon;

        [FindsBy(How = How.XPath, Using = CHECK_DELIVERIES_EXIST)]
        private IWebElement _checkDeliveriesExist;
        // _________________________________________ Méthodes ___________________________________________________

        public string GetDeliveryName()
        {
            _deliveryName = WaitForElementIsVisible(By.XPath(DELIVERY_NAME));
            return _deliveryName.Text;
        }

        public void ClickCrayon()
        {
            _crayon = WaitForElementIsVisible(By.XPath(CRAYON));
            _crayon.Click();
            WaitForLoad();

        }

        public bool CheckDeliveriesExist()
        {
            if (isElementVisible(By.XPath(CHECK_DELIVERIES_EXIST)))
            { 
                return true;
            }
            else 
            {
                return false; 
            }

        }
    }
}
