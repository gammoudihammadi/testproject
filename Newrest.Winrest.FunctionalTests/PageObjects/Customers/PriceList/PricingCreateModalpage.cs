using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.PriceList
{
    public class PricingCreateModalpage : PageBase
    {
        public PricingCreateModalpage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //__________________________________ Constantes _______________________________________

        private const string NAME = "Name";
        private const string START_DATE = "StartDate";
        private const string END_DATE = "EndDate";
        private const string SITE = "dropdown-site";
        private const string IS_DEFAULT = "IsDefault";

        private const string CANCEL = "//*[@id=\"form-createdit-vipPrice\"]/div[3]/div/button[1]";
        private const string CREATE = "//*[@id=\"form-createdit-vipPrice\"]/div[3]/div/button[2]";

        private const string ERROR_MESSAGE_NAME = "//*[@id=\"createFormVipPrice\"]/div[1]/div/span";
        private const string ERROR_MESSAGE_ALREADY_EXIST = "//*[@id=\"createFormVipPrice\"]/p/span";

        //__________________________________ Variables _______________________________________

        [FindsBy(How = How.Id, Using = NAME)]
        private IWebElement _name;

        [FindsBy(How = How.Id, Using = START_DATE)]
        private IWebElement _startDate;

        [FindsBy(How = How.Id, Using = END_DATE)]
        private IWebElement _endDate;

        [FindsBy(How = How.Id, Using = SITE)]
        private IWebElement _site;

        [FindsBy(How = How.Id, Using = IS_DEFAULT)]
        private IWebElement _isDefault;

        [FindsBy(How = How.XPath, Using = CANCEL)]
        private IWebElement _cancel;

        [FindsBy(How = How.XPath, Using = CREATE)]
        private IWebElement _create;

        [FindsBy(How = How.XPath, Using = ERROR_MESSAGE_NAME)]
        private IWebElement _errorMessageName;

        [FindsBy(How = How.XPath, Using = ERROR_MESSAGE_ALREADY_EXIST)]
        private IWebElement _errorMessageAlreadyExist;

        // ___________________________________ Méthodes ______________________________________________

        public PriceListDetailsPage FillField_CreateNewPricing(string name, string site, DateTime startDateTime, DateTime endDateTime, bool isDefault = false, bool avoidAlreadyExist=false)
        {
            _name = WaitForElementIsVisible(By.Id(NAME));
            _name.SetValue(ControlType.TextBox, name);
            WaitForLoad();

            _startDate = WaitForElementIsVisible(By.Id(START_DATE));
            _startDate.SetValue(ControlType.DateTime, startDateTime);
            _startDate.SendKeys(Keys.Tab);
            WaitForLoad();

            _endDate = WaitForElementIsVisible(By.Id(END_DATE));
            _endDate.SetValue(ControlType.DateTime, endDateTime);
            _endDate.SendKeys(Keys.Tab);
            WaitForLoad();

            _site = WaitForElementIsVisible(By.Id(SITE));
            _site.SetValue(ControlType.DropDownList, site);
            WaitForLoad();

            if (isDefault)
            {
                _isDefault = WaitForElementExists(By.Id(IS_DEFAULT));
                _isDefault.SetValue(ControlType.CheckBox, isDefault);
                WaitForLoad();
            }

            CreatePricing();
            if(avoidAlreadyExist)
            {
                var dayOffset = 1;
                while (isElementVisible(By.XPath(ERROR_MESSAGE_ALREADY_EXIST)))
                {
                    _startDate = WaitForElementIsVisible(By.Id(START_DATE));
                    _startDate.SetValue(ControlType.DateTime, startDateTime.AddDays(dayOffset));
                    WaitForLoad();
                    dayOffset++;
                    CreatePricing();
                }
            }
     
            return new PriceListDetailsPage(_webDriver, _testContext);
        }

        public void CreatePricing()
        {
            // Click sur le bouton "Create"
            _create = WaitForElementToBeClickable(By.XPath(CREATE));
            _create.Click();
            WaitForLoad();
        }      

        public bool ErrorMessageNameRequired()
        {
            _errorMessageName = WaitForElementIsVisible(By.XPath(ERROR_MESSAGE_NAME));

            if (_errorMessageName.Text.Contains("The Name field is required"))
                return true;

            return false;
        }

        public bool ErrorMessageNameAlreadyExists()
        {

            _errorMessageName = WaitForElementExists(By.XPath(ERROR_MESSAGE_NAME));
            _errorMessageAlreadyExist = WaitForElementExists(By.XPath(ERROR_MESSAGE_ALREADY_EXIST));

            if (_errorMessageName.Text.Contains("Name already exists for the same site") 
                && _errorMessageAlreadyExist.Text.Contains("There is already an existing default pricing for this site within this period"))
                return true;

            return false;
        }

        public PriceListPage Cancel()
        {
            _cancel = WaitForElementIsVisible(By.XPath(CANCEL));
            _cancel.Click();
            WaitForLoad();

            return new PriceListPage(_webDriver, _testContext);
        }
    }
}
