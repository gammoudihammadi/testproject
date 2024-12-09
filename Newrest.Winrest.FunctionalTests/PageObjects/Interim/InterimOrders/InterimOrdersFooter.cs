using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Interim.InterimOrders
{
    public class InterimOrdersFooter : PageBase
    {
        public InterimOrdersFooter(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        private const string TOTAL_FOOTER_PRICE = "//*[@id=\"tabContentDetails\"]/div/table[1]/tbody/tr[5]/td[4]";
        private const string TOTAL_FOOTER_GROSS_AMOUNT = "//*[@id=\"tabContentDetails\"]/div/table[1]/tbody/tr[4]/td[2]";
        private const string TOTAL_FOOTER_VAT_AMOUNT = "//*[@id=\"tabContentDetails\"]/div/table[1]/tbody/tr[4]/td[4]";
        private const string TAX_NAME = "//*[@id=\"tabContentDetails\"]/div/table[1]/tbody/tr[2]/td[1]";

        [FindsBy(How = How.XPath, Using = TOTAL_FOOTER_PRICE)]
        private IWebElement _totalFooterPrice;

        [FindsBy(How = How.XPath, Using = TOTAL_FOOTER_GROSS_AMOUNT)]
        private IWebElement _totalFooterGrossAmount;

        [FindsBy(How = How.XPath, Using = TOTAL_FOOTER_VAT_AMOUNT)]
        private IWebElement _totalFooterVatAmount;

        [FindsBy(How = How.XPath, Using = TAX_NAME)]
        private IWebElement _taxName;

        public decimal GetTotalFooterPrice()
        {
            _totalFooterPrice = WaitForElementIsVisible(By.XPath(TOTAL_FOOTER_PRICE));
            var priceText = _totalFooterPrice.Text;

            var cleanPriceText = priceText.Replace("€", "").Trim();

            return decimal.Parse(cleanPriceText);
        }

        public decimal GetTotalFooterGrossAmount()
        {
            _totalFooterGrossAmount = WaitForElementIsVisible(By.XPath(TOTAL_FOOTER_GROSS_AMOUNT));
            var priceText = _totalFooterGrossAmount.Text;

            var cleanPriceText = priceText.Replace("€", "").Trim();

            return decimal.Parse(cleanPriceText);
        }

        public decimal GetTotalFooterVatAmount()
        {
            _totalFooterVatAmount = WaitForElementIsVisible(By.XPath(TOTAL_FOOTER_VAT_AMOUNT));
            var priceText = _totalFooterVatAmount.Text;

            var cleanPriceText = priceText.Replace("€", "").Trim();

            return decimal.Parse(cleanPriceText);
        }
        public string GetTaxeName()
        {
            _taxName = WaitForElementIsVisible(By.XPath(TAX_NAME));
            return _taxName.Text.Trim(); ;
        }
    }
}
