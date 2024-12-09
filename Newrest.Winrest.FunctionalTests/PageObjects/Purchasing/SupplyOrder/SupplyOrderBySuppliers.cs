using System;
using System.Globalization;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Purchasing
{
    public class SupplyOrderBySuppliers : PageBase
    {

        public SupplyOrderBySuppliers(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }
        // ________________________________________ Constantes ____________________________________________

        private const string SUPPLIERS_LIST = "//*[@id=\"tabContentDetails\"]/table/tbody/tr";
        //private const string SUPPLIER_NAME = "//*[@id=\"tabContentDetails\"]/table/tbody/tr[{0}]/td[1]";

        // ________________________________________ Variables ______________________________________________


        // ________________________________________ Méthodes _______________________________________________



        public bool HasSuppliers()
        {
            var suppliers = _webDriver.FindElements(By.XPath(SUPPLIERS_LIST));

            if(suppliers.Count < 2)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public bool IsSupplierPresent(string supplierName)
        {
            bool isSupplierPresent = false;

            var suppliers = _webDriver.FindElements(By.XPath(SUPPLIERS_LIST));

            if (suppliers.Count < 2)
            {
                isSupplierPresent = false;
            }
            else
            {

                foreach (var elm in suppliers)
                {
                    if (elm.Text.Contains(supplierName))
                    {
                        isSupplierPresent = true;
                        break;
                    }
                }
            }
            return isSupplierPresent;
        }

    }
}
