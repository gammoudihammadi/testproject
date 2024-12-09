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
using System.Threading;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Flights.LpCart
{
    public class LpCartDuplicateTrolleyModalPage : PageBase
    {
        // ____________________________________ Constantes __________________________________________

        private const string LPCART_DESTINATION = "SelectLPCarts_ms";
        private const string CONFIRM = "confirm";
        private const string SEARCH_CART = "/html/body/div[12]/div/div/label/input";
        private const string CHECK_BOX_TROLLEY = "/html/body/div[4]/div/div/div[1]/div/form/div[3]/table/tbody/tr[1]/td[1]/div/input[1]";
        private const string KEYWORD = "form-copy-from-tbSearchName";
        private const string TROLLEY_NAME = "//*[@id=\"table-copy-from-rn\"]/tbody/tr/td[2]";

        // __________________________________ Filtres ________________________________________

        [FindsBy(How = How.Id, Using = LPCART_DESTINATION)]
        private IWebElement _lpCartDestination;

        [FindsBy(How = How.XPath, Using = SEARCH_CART)]
        private IWebElement _searchCart;

        [FindsBy(How = How.XPath, Using = CHECK_BOX_TROLLEY)]
        private IWebElement _checkBoxTrolley;

        [FindsBy(How = How.XPath, Using = CONFIRM)]
        private IWebElement _confirmBtn;

        [FindsBy(How = How.Id, Using = KEYWORD)]
        private IWebElement _keyword;

        [FindsBy(How = How.Id, Using = TROLLEY_NAME)]
        private IWebElement _trolleyName;
        public LpCartDuplicateTrolleyModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        public void SetValuesForDuplication(string lpCartCode)
        {
            //Actions action = new Actions(_webDriver);

            //_lpCartDestination = WaitForElementExists(By.Id(LPCART_DESTINATION));
            //action.MoveToElement(_lpCartDestination).Perform();
            //_lpCartDestination.Click();

            //_searchCart = WaitForElementIsVisible(By.XPath(SEARCH_CART));
            ////_searchCart.SetValue(ControlType.TextBox, lpCartCode);
            ////478694
            //_searchCart.SendKeys(lpCartCode);
            //var valueToCheckCart = WaitForElementIsVisible(By.XPath("//span[text()='" + lpCartCode + " (current LPCart)" + "']"));
            //valueToCheckCart.SetValue(ControlType.CheckBox, true);
            //WaitForLoad();
            ComboBoxSelectById(new ComboBoxOptions(LPCART_DESTINATION, (string)$"{lpCartCode} (current LPCart)", false));
            WaitForLoad();
            // Check Trolley
            _checkBoxTrolley = WaitForElementExists(By.XPath(CHECK_BOX_TROLLEY));
            _checkBoxTrolley.SetValue(ControlType.CheckBox, true);

            _confirmBtn = WaitForElementToBeClickable(By.Id(CONFIRM));
            _confirmBtn.Click();
            WaitForLoad();

        }

        public void SetKeywordFilter(string value)
        {
            _keyword = WaitForElementExists(By.Id(KEYWORD));
            _keyword.SetValue(ControlType.TextBox, value);

            WaitPageLoading();//Temps d'attente en filtre Keyword 
            WaitForLoad();
            WaitPageLoading();
        }

        public bool IsGoodValueFilterKeyword(string value)
        {
            _trolleyName = WaitForElementIsVisible(By.XPath(TROLLEY_NAME));

            if (value != _trolleyName.Text)
            {
                return false;
            }
            return true;
        }
    }
}
