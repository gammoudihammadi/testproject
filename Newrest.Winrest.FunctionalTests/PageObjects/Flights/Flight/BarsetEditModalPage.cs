using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Flights.Flight
{
    public class BarsetEditModalPage : PageBase
    {
        public BarsetEditModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // _________________________________________ Constantes ________________________________________
        private const string ADD_GUEST_BTN = "//a[text()='Add guest type']";
        private const string GUEST_NAME = "GuestTypeId";
        private const string CREATE_GUEST_BUTTON = "//*[@id=\"form-create-guest-type\"]/div[4]/button[2]";
        private const string ADD_MENU_BUTTON = "//*[@id=\"leg-details-with-services\"]/div/div/div/button/span";
        private const string ADD_SERVICE = "//a[text()='add service']";
        private const string CLOSE_BTN = "viewDetails-close";


        // ________________________________________ Variables ________________________________________________

        // General
        [FindsBy(How = How.XPath, Using = ADD_GUEST_BTN)]
        private IWebElement _addGuestBtn;

        [FindsBy(How = How.Id, Using = GUEST_NAME)]
        private IWebElement _guestName;

        [FindsBy(How = How.XPath, Using = CREATE_GUEST_BUTTON)]
        private IWebElement _createGuest;

        [FindsBy(How = How.XPath, Using = ADD_MENU_BUTTON)]
        private IWebElement _addServiceMenu;

        [FindsBy(How = How.XPath, Using = ADD_SERVICE)]
        private IWebElement _addService;

        public void AddGuestType(string guestType = null)
        {
            _addGuestBtn = WaitForElementIsVisible(By.XPath(ADD_GUEST_BTN));
            _addGuestBtn.Click();
            WaitForLoad();

            if (guestType != null)
            {
                _guestName = WaitForElementIsVisible(By.Id(GUEST_NAME));
                _guestName.SetValue(ControlType.DropDownList, guestType);
                WaitForLoad();
            }

            _createGuest = WaitForElementIsVisible(By.XPath(CREATE_GUEST_BUTTON));
            _createGuest.Click();
            WaitForLoad();
        }

        public void ShowAddServiceMenu()
        {
            WaitForLoad();
            Actions action = new Actions(_webDriver);

            _addServiceMenu = WaitForElementIsVisible(By.XPath("//*[@id=\"leg-details-with-services\"]/div/div/div/button/span"));
            action.MoveToElement(_addServiceMenu).Perform();
            WaitForLoad();
        }

        public void AddFirstServiceBarset()
        {
            ShowAddServiceMenu();

            _addService = WaitForElementIsVisible(By.XPath(ADD_SERVICE));
            _addService.Click();
            WaitForLoad();

        }

        public BarsetsPage CloseViewDetails()
        {
            WaitForLoad();
            var _closeBtn = WaitForElementIsVisible(By.Id(CLOSE_BTN));
            new Actions(_webDriver).MoveToElement(_closeBtn).Click().Perform();
            WaitForLoad();

            return new BarsetsPage(_webDriver, _testContext);
        }

    }
    }
