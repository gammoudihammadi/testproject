using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.OutputForm
{
    public class OutputFormExpiry : PageBase
    {
        public OutputFormExpiry(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        public string GetExpiryName()
        {
            var ExpiryName = WaitForElementIsVisible(By.XPath("//*[@id=\"formSaveDates\"]/div[2]/div[1]/div[2]/p/b"));
            return ExpiryName.Text;
        }

        public string GetExpiryTotalQty()
        {
            var ExpiryTotalQty = WaitForElementIsVisible(By.XPath("//*[@id=\"formSaveDates\"]/div[2]/div[2]/div[2]/input"));
            return ExpiryTotalQty.GetAttribute("value");
        }

        public void CreateFirstRow(string value)
        {
            if (value == null)
            {
                var delete = WaitForElementIsVisible(By.XPath("//*[@id=\"rowExpiryDate_0\"]/div[5]/a"));
                delete.Click();
                return;
            }
            var firstQty = _webDriver.FindElements(By.Id("Quantity_0"));
            if (firstQty.Count == 0)
            {
                // pas de Quantity_0
                var newButton = WaitForElementIsVisible(By.Id("btn-create-new-row"));
                newButton.Click();
            }
            var ExpiryQty = WaitForElementIsVisible(By.Id("Quantity_0"));
            ExpiryQty.Clear();
            ExpiryQty.SendKeys(value);
            var ExpiryDate = WaitForElementIsVisible(By.Id("datepicker-expiration_0"));
            ExpiryDate.Clear();
            ExpiryDate.SendKeys(DateUtils.Now.Date.AddMonths(1).ToString("dd/MM/yyyy"));

        }

        public bool HasMessageError(string errorMessage)
        {
            var message = _webDriver.FindElements(By.XPath("//*[text()='" + errorMessage + "']"));
            return message.Count == 1;
        }

        public void Save()
        {
            var save = WaitForElementIsVisible(By.Id("btnSubmit"));
            save.Click();
        }

        public bool CheckGreenIcon()
        {
            var greenOn = WaitForElementIsVisible(By.XPath("(//*[contains(@name,'IconExpirationDate')])[1]"));
            return greenOn.GetAttribute("class").Contains("green-text");
        }
    }
}
