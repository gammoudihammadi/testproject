using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Interim.InterimOrders;
using Newrest.Winrest.FunctionalTests.PageObjects.Interim.InterimReceptions;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Accounting.SupplierInvoices
{
    public class InterimOrderPage : PageBase
    {
        public InterimOrderPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        public enum FilterType
        {
            Search
        }
        public void Filter(FilterType filterType, object value)
        {
            switch (filterType)
            {
                case FilterType.Search:
                    var _searchByName = WaitForElementIsVisible(By.Id("tbSearchPatternWithAutocomplete"));
                    _searchByName.SetValue(ControlType.TextBox, value);
                    _searchByName.SendKeys(Keys.Enter);
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);

            }
            WaitPageLoading();
            WaitForLoad();
        }

        public void WaitForLoad()
        {
            base.WaitForLoad();
        }


        public bool CheckDataExist()
        {
            var element = isElementExists(By.XPath("//*/form[contains(@action,'SaveRow')]"));
            return element;
        }

        public void CreateNewInterimOrder(string site, string supplierName)
        {
            var create4 = WaitForElementIsVisible(By.XPath("//*/button[text()='+']"));
            new Actions(_webDriver).MoveToElement(create4).Perform();
            var createNew2 = WaitForElementIsVisible(By.XPath("//*/a[text()='New interim order']"));
            createNew2.Click();
            WaitForLoad();

            var siteOrder = WaitForElementIsVisible(By.Id("SelectedSiteId"));
            siteOrder.SetValue(PageBase.ControlType.DropDownList, site);
            WaitForLoad();
            var supplierOrder = WaitForElementIsVisible(By.Id("drop-down-suppliers"));
            supplierOrder.SetValue(PageBase.ControlType.DropDownList, supplierName);


            var placeOrder = WaitForElementIsVisible(By.Id("SelectedSitePlaceId"));
            placeOrder.SetValue(PageBase.ControlType.DropDownList, "Produccion"); // ambiguite entre Producción ou  Produccion

            var create5 = WaitForElementIsVisible(By.Id("btn-submit-form-create-interim-order"));
            create5.Click();
        }
        public void ChangeQuatity()
        {
            var firstLine = WaitForElementExists(By.XPath("//*/form[contains(@action,'SaveRow')]"));

            new Actions(_webDriver).MoveToElement(firstLine).Click().Perform();
            var qtyOrder = WaitForElementIsVisible(By.Id("item_IodRowDto_Quantity"));
            qtyOrder.SetValue(PageBase.ControlType.TextBox, "5,000");

            WaitForLoad();
            bool hasError = false;
            string errorTitle = "";

            if (isElementVisible(By.XPath("//*/span[contains(@title,'Bad format key')]")))
            {
                var errorMessage = _webDriver.FindElement(By.XPath("//*/span[contains(@title,'Bad format key')]"));
                errorTitle = errorMessage.GetAttribute("title");
                hasError = true;

            }
            if (hasError)
            {
                // erreur
                Assert.Fail("Create interim reception from (i) " + errorTitle);
            }
        }
    }
}
