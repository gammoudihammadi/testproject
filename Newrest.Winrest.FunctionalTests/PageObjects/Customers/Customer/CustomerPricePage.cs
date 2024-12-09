using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.Customer
{
    public class CustomerPricePage : PageBase
    {
        private static string PRICE_NUMBER = "//*[contains(@class,'item-tab-row')]";

        private static string NEW_PRICE = "//*/a[text()='New price']";
        private static string NEW_PRICE_SITE = "drop-down-sites";
        private static string NEW_PRICE_PRICE = "drop-down-vip-prices";
        private static string NEW_PRICE_CREATE = "//*/button[text()='Create']";

        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";

        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;

        public CustomerPricePage(IWebDriver _webDriver, TestContext _testContext) : base(_webDriver, _testContext)
        {
        }

        public int PriceNumber()
        {
            return _webDriver.FindElements(By.XPath(PRICE_NUMBER)).Count;
        }

        public void AddPrice(string site, string price)
        {
            var buttonAdd = WaitForElementIsVisible(By.XPath(NEW_PRICE));
            buttonAdd.Click();
            WaitForLoad();

            var selectSite = WaitForElementIsVisible(By.Id(NEW_PRICE_SITE));
            new SelectElement(selectSite).SelectByText(site);
            WaitForLoad();

            var selectPrice = WaitForElementIsVisible(By.Id(NEW_PRICE_PRICE));
            new SelectElement(selectPrice).SelectByText(price);
            WaitForLoad();

            var createButton = WaitForElementIsVisible(By.XPath(NEW_PRICE_CREATE));
            createButton.Click();
            WaitForLoad();
        }

        public CustomerPage BackToList()
        {
            _backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            _backToList.Click();
            WaitForLoad();

            return new CustomerPage(_webDriver, _testContext);
        }

        public bool CheckSiteExists(string site)
        {
            var siteDiv = WaitForElementIsVisible(By.XPath("//*[@id=\"list-item-with-action\"]"));
            if (siteDiv.Text == "No prices found.")
                return false;
            else
            {
                string xpath = "//*[@id=\"list-item-with-action\"]/div/div[2]";
                var tableWithData = WaitForElementIsVisible(By.XPath(xpath));
                var allElems = tableWithData.FindElements(By.XPath("div"));
                int childrenCount = allElems.Count;

                if(childrenCount == 1)
                {
                    siteDiv = WaitForElementIsVisible(By.XPath("//*[@id=\"list-item-with-action\"]/div/div[2]/div/div/div[2]"));
                    return siteDiv.Text == site;
                }
                else
                {

                    for (int i = 1; i < childrenCount + 1; i++)
                    {
                        xpath = $"//*[@id=\"list-item-with-action\"]/div/div[2]/div[{i}]/div/div[2]";
                        siteDiv = WaitForElementIsVisible(By.XPath(xpath));
                        if (siteDiv.Text == site) return true;
                    }
                    return false;
                }
            }
        }
    }
}
