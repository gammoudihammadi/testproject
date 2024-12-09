using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.Production;
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
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Production.ProductionCO
{
    public class ProductionCOOutputFormPage : PageBase
    {
        public ProductionCOOutputFormPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }
        private const string FROM = "drop-down-places-from";
        private const string TO = "drop-down-places-to";
        private const string GENERATE = "btn-submit-create-output-form";
        private const string TELECHARGEMENT = "production-customerorderproduction-outputform-1";
        private const string OUTPUT_FROM_NO = "//*[@id=\"div-body\"]/div/div[1]/h1";
        private const string OUTPUT_FROM_NUMBER = "tb-new-outputform-number";
        private const string OUTPUT_FROM_NUMBER_ELEMENT= "//div[contains(@id, 'customer-order-production-detail')]//table/tbody/tr[2]/td[1]";
        private const string GENERAL_INFORMATION_TAB = "hrefTabContentInformations";

        [FindsBy(How = How.Id, Using = FROM)]
        private IWebElement _from;

        [FindsBy(How = How.Id, Using = TO)]
        private IWebElement _to;

        [FindsBy(How = How.Id, Using = TELECHARGEMENT)]
        private IWebElement _telechargement;

        [FindsBy(How = How.Id, Using = GENERATE)]
        private IWebElement _generate;

        [FindsBy(How = How.Id, Using = GENERAL_INFORMATION_TAB)]
        private IWebElement _generalInformationTab;

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
        public ProductionCOOutputFormPage Generate()
        {
            _generate = WaitForElementIsVisible(By.Id(GENERATE));
            _generate.Click();

            WaitForLoad();

            return new ProductionCOOutputFormPage(_webDriver, _testContext);
        }
        public string GetOutputFormNumber()
        {
            var outputFormNumberElement = _webDriver.FindElement(By.Id("tb-new-outputform-number"));
            return outputFormNumberElement.GetAttribute("value");
        }
        public string GetPCONumber()
        {
            var outputFormNumberElement = _webDriver.FindElement(By.XPath(OUTPUT_FROM_NUMBER_ELEMENT));
            string rawNumber = outputFormNumberElement.Text.Trim();
            int parsedNumber = int.Parse(rawNumber);
            return parsedNumber.ToString();
        }

        
        public bool VerifyIsOutputFormGenerated(string numberInitial)
        {
            var element = _webDriver.FindElement(By.XPath(OUTPUT_FROM_NO));
            string text = element.Text;
            return text == "OUTPUT FORM NO " + numberInitial.ToString();
        }
        public OutputFormGeneralInformation ClickOnGeneralInformationTab()
        {
            _generalInformationTab = WaitForElementIsVisible(By.Id(GENERAL_INFORMATION_TAB));
            _generalInformationTab.Click();
            WaitForLoad();

            return new OutputFormGeneralInformation(_webDriver, _testContext);
        }



    }

}

