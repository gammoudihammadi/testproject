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
    public class SupplierCertificationTab : PageBase
    {
        public SupplierCertificationTab(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {

        }

        public SupplierAccountingsTiersDetailsTab ClickOnAccountingsTiersDetailsTab()
        {
            var tab = WaitForElementExists(By.Id("hrefTabContentAccountingTiers"));
            tab.Click();
            WaitForLoad();
            return new SupplierAccountingsTiersDetailsTab(_webDriver, _testContext);
        }

        public string GetName(int i)
        {
            return WaitForElementIsVisible(By.XPath("//*[@id='form-certif_" + i + "']/div/div[1]")).Text;
        }

        public string GetIsCopyMandatory(int i)
        {
            return WaitForElementIsVisible(By.XPath("//*[@id='form-certif_" + i + "']/div/div[2]")).Text;
        }

        public string GetCopyGiven(int i)
        {
            return WaitForElementExists(By.Id("Certifications_" + i + "__CopyHasBeenGiven")).GetAttribute("value");
        }

        public string GetExpireDateCheck(int i)
        {
            return WaitForElementIsVisible(By.XPath("//*[@id='form-certif_" + i + "']/div/div[4]")).Text;
        }

        public string GetExpirationDate(int i)
        {
            return WaitForElementIsVisible(By.Id("expiration-date-picker_" + i)).GetAttribute("value");
        }
    }
}
