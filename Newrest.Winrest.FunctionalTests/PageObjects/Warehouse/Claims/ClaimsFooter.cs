using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.Claims
{
    public class ClaimsFooter : PageBase
    {
        // Informations
        private const string TOTAL_SANCTIONS_NO_VAT = "//td[text()='TOTAL SANCTIONS']/../td[2]";
        private const string TOTAL_NO_VAT = "//td[text()='TOTAL GROSS AMOUNT']/../td[2]";
        private const string VAT = "//td[text()='TOTAL GROSS AMOUNT']/../td[4]";
        private const string TOTAL_WITH_VAT = "//td[text()='TOTAL CLAIM']/../td[4]";

        //_________________________________ Variables _______________________________________

        [FindsBy(How = How.Id, Using = TOTAL_NO_VAT)]
        private IWebElement _totalNoVat;

        [FindsBy(How = How.Id, Using = VAT)]
        private IWebElement _vat;

        [FindsBy(How = How.Id, Using = TOTAL_WITH_VAT)]
        private IWebElement _totalWithVat;

        public ClaimsFooter(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        public double GetTotalNoVat()
        {
            _totalNoVat = WaitForElementIsVisible(By.XPath(TOTAL_NO_VAT));
            double totalNoVat = ArrangeTarif(_totalNoVat.Text);
            if (isElementVisible(By.XPath(TOTAL_SANCTIONS_NO_VAT))) {
                var _totalSanction = WaitForElementIsVisible(By.XPath(TOTAL_SANCTIONS_NO_VAT));
                totalNoVat += ArrangeTarif(_totalSanction.Text);
            }

            return totalNoVat;
        }

        public string GetVat()
        {
            _vat = WaitForElementIsVisible(By.XPath(VAT));
            return _vat.Text;
        }

        public string GetTotalWithVat()
        {
            _totalWithVat = WaitForElementIsVisible(By.XPath(TOTAL_WITH_VAT));
            return _totalWithVat.Text;
        }

        public string GetTotalClaim()
        {
            _totalWithVat = WaitForElementIsVisible(By.XPath("//td[text()='TOTAL CLAIM']/../td[4]"));
            return _totalWithVat.Text;
        }

        public double ArrangeTarif(string tarif)
        {
            string tarifLinear = tarif.Replace("€", "").Replace(" ", "");
            double tarifDouble = Convert.ToDouble(tarifLinear, new NumberFormatInfo() { NumberDecimalSeparator = "," });
            return Math.Round(tarifDouble, 3);
        }
    }
}
