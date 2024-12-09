using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Flights.Trolleys
{
    public class ConfigurationModalPage : PageBase
    {
        public ConfigurationModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // __________________________________ Constantes _____________________________________

        private const string FIRST_ITEM = "/html/body/div[3]/div/div/div[4]/div[2]/div/table/tbody/tr[2]";


        public void SetValuesForService(string value)
        {
            var elements = _webDriver.FindElements(By.XPath("/html/body/div[3]/div/div/div[4]/div[2]/div/table/tbody/tr[2]/td[*]/input[@type='number']"));

            Actions action = new Actions(_webDriver);

            var firstItem = WaitForElementExists(By.XPath(FIRST_ITEM));
            action.MoveToElement(firstItem).Perform();

            foreach (var elm in elements)
            {
                //elm.Click();
                elm.SetValue(ControlType.TextBox, value);
            }
            WaitForLoad();
            Thread.Sleep(1000);
            IWebElement close;
            if(isElementVisible(By.XPath("//*[@id=\"modal-1\"]/div/div/div[5]/button")))
            {
                close = WaitForElementIsVisible(By.XPath("//*[@id=\"modal-1\"]/div/div/div[5]/button"));
            }
            else
            {
                close = WaitForElementIsVisible(By.XPath("//*[@id=\"modal-1\"]/div[5]/button"));
            }
            close.Click();

            WaitPageLoading();
        }

        public void DeleteValuesForService()
        {
            Actions action = new Actions(_webDriver);

            var firstItem = WaitForElementExists(By.XPath(FIRST_ITEM));
            action.MoveToElement(firstItem).Perform();

            var delete = WaitForElementIsVisible(By.XPath("//*[@id=\"configs-table\"]/tbody/tr[2]/td[6]/a/span"));
            delete.Click();

            var validate = WaitForElementIsVisible(By.Id("dataConfirmOK"));
            validate.Click();

            WaitForLoad();
            IWebElement close;

            if (isElementVisible(By.XPath("//*[@id=\"modal-1\"]/div/div/div[5]/button")))
            {
                close = WaitForElementIsVisible(By.XPath("//*[@id=\"modal-1\"]/div/div/div[5]/button"));
            }
            else
            {
                close = WaitForElementIsVisible(By.XPath("//*[@id=\"modal-1\"]/div[5]/button"));
            }
            close.Click();

            WaitPageLoading();
        }

    }
}
