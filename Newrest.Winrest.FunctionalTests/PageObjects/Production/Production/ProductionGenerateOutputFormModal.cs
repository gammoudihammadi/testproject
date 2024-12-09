using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.OutputForm;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Production.Production
{
    public class ProductionGenerateOutputFormModal : PageBase
    {
        public ProductionGenerateOutputFormModal(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // ______________________________________ Constantes _____________________________________________

        private const string FROM = "drop-down-places-from";
        private const string TO = "drop-down-places-to";
        private const string GENERATE = "btn-submit-create-output-form";


        // ______________________________________ Variables _____________________________________________

        [FindsBy(How = How.Id, Using = FROM)]
        private IWebElement _from;
        
        [FindsBy(How = How.Id, Using = TO)]
        private IWebElement _to;
        
        [FindsBy(How = How.Id, Using = GENERATE)]
        private IWebElement _generate;

        public void SetFromPlace(string value)
        {
            _from = WaitForElementIsVisible(By.Id(FROM));
            _from.SetValue(ControlType.DropDownList, value);
        }
        public void SetToPlace(string value)
        {
            _to = WaitForElementIsVisible(By.Id(TO));
            _to.SetValue(ControlType.DropDownList, value);
        }

        public OutputFormItem Generate()
        {
            _generate = WaitForElementIsVisible(By.Id(GENERATE));
            _generate.Click();

            //Results are opened in a new tab, switch the driver to the newly created one
            WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new OutputFormItem(_webDriver, _testContext);
        }
    }
}
