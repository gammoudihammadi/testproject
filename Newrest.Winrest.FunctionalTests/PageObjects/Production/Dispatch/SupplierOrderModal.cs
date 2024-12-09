using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Production.Dispatch
{
    public class SupplierOrderModal : PageBase
    {
        public SupplierOrderModal(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //__________________________________ Constantes ________________________________________

        public const string NUMBER = "tb-new-supplyorder-number";
        public const string SITE_PLACE = "SelectedSitePlaceId";
        public const string GENERATE_BTN = "btn-submit-form-create-supply-order";

        //__________________________________ Variables _________________________________________

        [FindsBy(How = How.Id, Using = NUMBER)]
        private IWebElement _number;

        [FindsBy(How = How.Id, Using = SITE_PLACE)]
        private IWebElement _sitePlace;

        [FindsBy(How = How.Id, Using = GENERATE_BTN)]
        private IWebElement _generate;

        //___________________________________ Méthodes _________________________________________

        public string GenerateSupplyOrder(string deliveryLocation)
        {
            _number = WaitForElementIsVisible(By.Id(NUMBER));

            _sitePlace = WaitForElementIsVisible(By.Id(SITE_PLACE));
            _sitePlace.SetValue(ControlType.DropDownList, deliveryLocation);
            WaitForLoad();

            return _number.GetAttribute("value");
        }

        public SupplyOrderItem Generate()
        {
            _generate = WaitForElementIsVisible(By.Id(GENERATE_BTN));
            _generate.Click();
            WaitForLoad();

            //Results are opened in a new tab, switch the driver to the newly created one
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            WaitForLoad();

            return new SupplyOrderItem(_webDriver, _testContext);
        }
    }


}
