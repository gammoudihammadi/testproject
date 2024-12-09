using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Suppliers
{
    public class SupplierAccountingsTab : PageBase
    {
        public SupplierAccountingsTab(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {

        }

        private const string SITE = "//*[@id='dat-container']/table/tbody/tr[2]/td[1]";
        private const string THIRD_PART = "//*[@id='dat-container']/table/tbody/tr[2]/td[2]";
        private const string ACCOUNTING_ID = "//*[@id='dat-container']/table/tbody/tr[2]/td[5]";
        private const string THIRD_PART_DTI = "//*[@id='dat-container']/table/tbody/tr[2]/td[6]";
        private const string ACCOUNTING_ID_DTI = "//*[@id='dat-container']/table/tbody/tr[2]/td[7]";
        private const string SUPPLIER_ACCOUNTING_ID = "//*[@id='dat-container']/table/tbody/tr[2]/td[8]/a[1]";
        private const string ADDRESS_IDENTIFIER = "//*[@id='dat-container']/table/tbody/tr[2]/td[3]";
        private const string DOMICILIATION_IDENTIFIER = "//*[@id='dat-container']/table/tbody/tr[2]/td[4]";


        [FindsBy(How = How.XPath, Using = SITE)]
        private IWebElement _site;

        [FindsBy(How = How.XPath, Using = THIRD_PART)]
        private IWebElement _thirdPart;

        [FindsBy(How = How.XPath, Using = ACCOUNTING_ID)]
        private IWebElement _accountingId;

        [FindsBy(How = How.XPath, Using = THIRD_PART_DTI)]
        private IWebElement _thirdPartDti;

        [FindsBy(How = How.XPath, Using = ACCOUNTING_ID_DTI)]
        private IWebElement _accountingIdDti;

        [FindsBy(How = How.XPath, Using = SUPPLIER_ACCOUNTING_ID)]
        private IWebElement _supplierAccountingId;

        [FindsBy(How = How.XPath, Using = ADDRESS_IDENTIFIER)]
        private IWebElement _adressIdentifier;

        [FindsBy(How = How.XPath, Using = DOMICILIATION_IDENTIFIER)]
        private IWebElement _domiciliationIdentifier;
        



        public string GetSite()
        {
            _site = WaitForElementIsVisible(By.XPath(SITE));
            return _site.Text.Trim();
        }

        public string GetThirdPart()
        {
            _thirdPart = WaitForElementIsVisible(By.XPath(THIRD_PART));
            return _thirdPart.Text.Trim();
        }

        public string GetAccountingId()
        {
            _accountingId = WaitForElementIsVisible(By.XPath(ACCOUNTING_ID));
            return _accountingId.Text.Trim();
        }

        public string GetThirdPartDti()
        {
            _thirdPartDti = WaitForElementIsVisible(By.XPath(THIRD_PART_DTI));
            return _thirdPartDti.Text.Trim();
        }

        public string GetAccountingIdDti()
        {
            _accountingIdDti = WaitForElementIsVisible(By.XPath(ACCOUNTING_ID_DTI));
            return _accountingIdDti.Text.Trim();
        }

        public string GetSupplierAccountingId()
        {
            _supplierAccountingId = WaitForElementIsVisible(By.XPath(SUPPLIER_ACCOUNTING_ID));
            string hrefText = _supplierAccountingId.GetAttribute("href");
            var saId = _supplierAccountingId.GetAttribute("href").Substring(hrefText.IndexOf('=') + 1);
            return saId;
        }

        public string GetAddressIdentifier()
        {
            _adressIdentifier = WaitForElementIsVisible(By.XPath(ADDRESS_IDENTIFIER));
            return _adressIdentifier.Text.Trim();
        }

        public string GetDomiciliationidentifier()
        {
            _domiciliationIdentifier = WaitForElementIsVisible(By.XPath(DOMICILIATION_IDENTIFIER));
            return _domiciliationIdentifier.Text.Trim();
        }

        public SupplierCertificationTab ClickOnCertificationsTab()
        {
            var tab = WaitForElementExists(By.Id("hrefTabContentCertifications"));
            tab.Click();
            WaitForLoad();
            return new SupplierCertificationTab(_webDriver, _testContext);
        }
    }
}
