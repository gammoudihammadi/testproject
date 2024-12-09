using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Flights.LpCart
{
    public class LpCartDuplicateLpCartModalPage : PageBase
    {

        // ____________________________________ Constantes __________________________________________

        private const string LPCART_SELECTED = "SelectLPCart";
        private const string CODE = "tb-new-lpcart-number";
        private const string NAME = "tb-new-lpcart-name";
        private const string SITE = "selectSite";
        private const string CUSTOMER = "selectCustomer";
        private const string AIRCRAFT = "drop-down-suppliers";
        private const string CREATE = "last";



        // ____________________________________ Variables ___________________________________________

        [FindsBy(How = How.Id, Using = LPCART_SELECTED)]
        private IWebElement _code;

        [FindsBy(How = How.Id, Using = CODE)]
        private IWebElement _lpCartSelected;

        [FindsBy(How = How.Id, Using = NAME)]
        private IWebElement _name;

        [FindsBy(How = How.Id, Using = SITE)]
        private IWebElement _site;

        [FindsBy(How = How.Id, Using = CUSTOMER)]
        private IWebElement _customer;

        [FindsBy(How = How.Id, Using = AIRCRAFT)]
        private IWebElement _aircraft;

        [FindsBy(How = How.Id, Using = CREATE)]
        private IWebElement _createBtn;



        public LpCartDuplicateLpCartModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }


        public void SetValuesForDuplication(string lpCartNameSelect, string lpCartCode, string lpCartName, string site, string customer, string aircraft)
        {
            // Selection du lpCart à dupliquer
            _lpCartSelected = WaitForElementIsVisible(By.Id(LPCART_SELECTED));
            _lpCartSelected.SetValue(ControlType.DropDownList, lpCartNameSelect);
            WaitForLoad();

            // Définition du nom du lpCode
            _code = WaitForElementIsVisible(By.Id(CODE));
            _code.SetValue(ControlType.TextBox, lpCartCode);

            // Définition du nom du lpCartName
            _name = WaitForElementIsVisible(By.Id(NAME));
            _name.SetValue(ControlType.TextBox, lpCartName);

            // Définition du site
            _site = WaitForElementIsVisible(By.Id(SITE));
            _site.SetValue(ControlType.DropDownList, site);
            WaitForLoad();

            // Définition du Customer
            _customer = WaitForElementIsVisible(By.Id(CUSTOMER));
            _customer.SetValue(ControlType.DropDownList, customer);
            WaitForLoad();

            _createBtn = WaitForElementToBeClickable(By.Id(CREATE));
            _createBtn.Click();
            WaitForLoad();

        }


    }
}
