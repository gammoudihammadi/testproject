using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.Customer
{
    public class CustomerBuyOnBoardPage : PageBase
    {
        public CustomerBuyOnBoardPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // ____________________________________ Constantes ____________________________________________

        // General
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";

        // Informations
        private const string CLIENT_KEY = "AirfiCustomerKey";
        private const string CLIENT_SECRET = "AirfiCustomerSecret";
        private const string AIRFI_IATA = "AirfiCustomerIata";
        private const string SMARTBAR_BTN = "//*[@id=\"bob-filter-form\"]/div/div[2]/div[1]/div/div[6]/div[1]/div/div/div/div/span[2]";
        private const string AIRFI_BTN = "//*[@id=\"bob-filter-form\"]/div/div[2]/div[1]/div/div[6]/div[2]/div/div/div/div/span[2]";
        private const string GUEST_TYPE_UNCHECK_ALL = "/html/body/div[12]/div/ul/li/a[@title= 'Uncheck all']";
        // ____________________________________ Variables _____________________________________________

        // General
        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;

        // Informations
        [FindsBy(How = How.Id, Using = CLIENT_KEY)]
        private IWebElement _clientKey;

        [FindsBy(How = How.Id, Using = CLIENT_SECRET)]
        private IWebElement _clientSecret;

        [FindsBy(How = How.Id, Using = AIRFI_IATA)]
        private IWebElement _airfiIata;

        [FindsBy(How = How.XPath, Using = SMARTBAR_BTN)]
        private IWebElement _smartbarBtn;

        [FindsBy(How = How.XPath, Using = AIRFI_BTN)]
        private IWebElement _airfiBtn;

        // ____________________________________ Méthodes ______________________________________________

        // General
        public CustomerPage BackToList()
        {
            _backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            _backToList.Click();
            WaitForLoad();

            return new CustomerPage(_webDriver, _testContext);
        }

        // Informations
        public void SwitchSmarBarAndAirfi()
        {   
            _smartbarBtn = WaitForElementIsVisible(By.XPath("//*/label[text()='SmarBar ?']/../div[contains(@class,'checkbox')]"));
            _smartbarBtn.Click();
            _airfiBtn = WaitForElementIsVisible(By.XPath("//*/label[text()='Airfi ?']/../div[contains(@class,'checkbox')]"));
            _airfiBtn.Click();

            // Temps de prise en compte de la modif
            Thread.Sleep(2000);
        }

        public void SetValueOnBob(string clientKey, string clientSecret, string airfi)
        {
            // Modification du mail du client
            Actions action = new Actions(_webDriver);
            _clientKey = WaitForElementIsVisible(By.Id(CLIENT_KEY));
            action.MoveToElement(_clientKey).Click().Perform();
            WaitForLoad();
            _clientKey.SetValue(ControlType.TextBox, clientKey);

            _clientSecret = WaitForElementIsVisible(By.Id(CLIENT_SECRET));
            _clientSecret.SetValue(ControlType.TextBox, clientSecret);

            _airfiIata = WaitForElementIsVisible(By.Id(AIRFI_IATA));
            _airfiIata.SetValue(ControlType.TextBox, airfi);

            // Temps d'enregistrement des valeurs
            Thread.Sleep(5000);
        }

        public List<string> GetValueOnBob()
        {
            List<string> values = new List<string>();

            _clientKey = WaitForElementIsVisible(By.Id(CLIENT_KEY));
            values.Add(_clientKey.GetAttribute("value"));


            _clientSecret = WaitForElementIsVisible(By.Id(CLIENT_SECRET));
            values.Add(_clientSecret.GetAttribute("value"));


            _airfiIata = WaitForElementIsVisible(By.Id(AIRFI_IATA));
            values.Add(_airfiIata.GetAttribute("value"));

            return values;
        }

        public void SelectGuestValueOnBob(string guestBob)
        {
            if (isElementVisible(By.XPath(String.Format("//*[@id=\"bgt-items\"]/div[*]/div[1]/span[text()='{0}']", guestBob))))
            {
                var searchGuest = WaitForElementIsVisible(By.XPath(String.Format("//*[@id=\"bgt-items\"]/div[*]/div[1]/span[text()='{0}']", guestBob)));
                var Guestname = searchGuest.Text;
            }
            else
            {
                var addGuest = WaitForElementIsVisible(By.XPath("//*[@id=\"btn-add-bob-guest-type\"]"));
                addGuest.Click();
                var gusetList = WaitForElementIsVisible(By.XPath("//*[@id=\"select-guest-type\"]"));
                gusetList.Click();
                var checkGuset = WaitForElementIsVisible(By.Id("select-guest-type"));
                checkGuset.SetValue(ControlType.DropDownList, guestBob);
                var saveGuest = WaitForElementIsVisible(By.XPath("//*[@id=\"modal-1\"]/div/div[3]/button[2]"));
                saveGuest.Click();
            }

            WaitForLoad();
            WaitPageLoading();

        }
        public bool OrganizationIsExist(string site, string organisation)
        {
            var elements = _webDriver.FindElements(By.XPath($"//*[@id=\"bob-filter-form-warehouses\"]/div/div/div/div/span[contains(text(),'{site}')]/../../div[2]/select/option"));


            foreach (var element in elements)
            {
                if (element.Text.Contains(organisation))
                {
                    return true;
                }
            }
            return false;
        }
    }
}
