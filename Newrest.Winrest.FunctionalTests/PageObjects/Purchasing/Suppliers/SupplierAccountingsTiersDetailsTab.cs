using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Suppliers
{
    public class SupplierAccountingsTiersDetailsTab : PageBase
    {
        public SupplierAccountingsTiersDetailsTab(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {

        }

        public string GetTiersDetailId(int i)
        {
            string TierdDetailHREF = WaitForElementIsVisible(By.XPath("//*[@id='dat-container']/table/tbody/tr[" + (i + 2) + "]/td[6]/a[1]")).GetAttribute("href");
            return TierdDetailHREF.Substring(TierdDetailHREF.IndexOf("=") + 1);
        }

        public string GetTiersCode(int i)
        {
            return WaitForElementIsVisible(By.XPath("//*[@id='dat-container']/table/tbody/tr[" + (i + 2) + "]/td[1]")).Text;
        }

        public string GetTiersName(int i)
        {
            return WaitForElementIsVisible(By.XPath("//*[@id='dat-container']/table/tbody/tr[" + (i + 2) + "]/td[2]")).Text;
        }

        public string GetTiersAddressIdentifier(int i)
        {
            return WaitForElementIsVisible(By.XPath("//*[@id='dat-container']/table/tbody/tr[" + (i + 2) + "]/td[3]")).Text;
        }

        public string GetTiersAddress(int i)
        {
            return WaitForElementIsVisible(By.XPath("//*[@id='dat-container']/table/tbody/tr[" + (i + 2) + "]/td[4]")).Text;
        }

        public string GetTiersDomiciliationIdentifier(int i)
        {
            return WaitForElementIsVisible(By.XPath("//*[@id='dat-container']/table/tbody/tr[" + (i + 2) + "]/td[5]")).Text;
        }
    }
}
