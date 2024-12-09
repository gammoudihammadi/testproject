using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Production.Production
{
    public class ProductionCustomerOrderTabPage : PageBase
    {
        public ProductionCustomerOrderTabPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // ______________________________________ Constantes _____________________________________________

        private const string CUSTOMERORDER_COLUMN_TITLE = "//*[@id=\"list-item-with-action\"]/div[1]/div/div/div[2]/table/thead/tr/th[1]";
        private const string CUSTOMERORDER_NAMES_COLUMN = "//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]/table/tbody/tr/td[1]";

        // ______________________________________ Variables _____________________________________________

        [FindsBy(How = How.XPath, Using = CUSTOMERORDER_COLUMN_TITLE)]
        private IWebElement _columnTitle;

        public bool IsResultDisplayedByCustomerOrder()
        {
            _columnTitle = WaitForElementIsVisible(By.XPath(CUSTOMERORDER_COLUMN_TITLE));
            if (_columnTitle.Text.Equals("Customer type"))
             return true;
            else return false;

        }

        public Dictionary<string, string> GetCustomerOrdersNumberAndCustomers()
        {
            int i = 0;

            // map : order number --> customer
            var mapOrderNumberCustomer = new Dictionary<string, string>();

            var columnCustomerOrdersNames = _webDriver.FindElements(By.XPath(CUSTOMERORDER_NAMES_COLUMN));

            foreach (var CO in columnCustomerOrdersNames)
            {
                // On limite le nombre de menus remontés à 5 pour ne pas surcharger le test
                if (i >= 6)
                    break;

                // Le numéro de CO est concaténé au nom du Customer, nous séparons les deux
                var coConcatenee = CO.Text;

                var indexCONumberStart = coConcatenee.IndexOf("°") + 1;
                var indexCONumberEnd = coConcatenee.IndexOf("-") - 1;
                var coLenght = indexCONumberEnd - indexCONumberStart;
                var orderNumber = coConcatenee.Substring(indexCONumberStart, coLenght);

                var indexCustomerStart = coConcatenee.IndexOf("-") + 2;
                var customerAndDeliveryName = coConcatenee.Substring(indexCustomerStart);

                mapOrderNumberCustomer.Add(orderNumber, customerAndDeliveryName);
                i++;
            }

            return mapOrderNumberCustomer;
        }
    }
}
