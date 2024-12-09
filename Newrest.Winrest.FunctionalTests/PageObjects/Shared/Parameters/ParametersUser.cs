using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Shared.Parameters
{
    public class ParametersUser : PageBase
    {
        private const string RADIO_BTN_SITE_CREATED = "//*[@id=\"{0}\"]";

        [FindsBy(How = How.XPath, Using = RADIO_BTN_SITE_CREATED)]
        private IWebElement _radioBtnSiteCreated;

        [FindsBy(How = How.XPath, Using = "//*[@id=\"tbSearchPattern\"]")]
        private IWebElement _searchUser;

        [FindsBy(How = How.XPath, Using = "//*[@id=\"82aebc96-5044-48b2-94b8-1323b58671d4\"]/td/h4")]
        private IWebElement _selectUser;

        public ParametersUser(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {

        }

        public void SearchAndSelectUser(string userName)
        {
            WaitForLoad();
            Actions action = new Actions(_webDriver);

            //Recherched du user
            action.MoveToElement(_searchUser).Perform();
            _searchUser.SetValue(ControlType.TextBox, userName);
            _searchUser.SendKeys(Keys.Tab);
            WaitForLoad();

            action.MoveToElement(_selectUser).Perform();
            _selectUser.Click();
            WaitForLoad();
        }

        public void GiveSiteRightsToUser(string id, bool hasPermission = true)
        {
            Thread.Sleep(1000);
            //Attente que la popup modale soit complètement chargée
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(30));
            Actions action = new Actions(_webDriver);

            var xpath = String.Format(RADIO_BTN_SITE_CREATED, id);
            IWebElement _siteCreated = _webDriver.FindElement(By.XPath(xpath));

            //Donne les droits au site créé
            action.MoveToElement(_siteCreated).Perform();
            _siteCreated.SetValue(ControlType.CheckBox, hasPermission);
            _siteCreated.SendKeys(Keys.Tab);
        }

        public void SelectUnselectAllSites(bool selectAll = true)
        {
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(30));
            Actions action = new Actions(_webDriver);
            var _selectUnselect = _webDriver.FindElement(By.XPath("//*[@id=\"selectAllSitesBtn\"][contains(.,' Select all')]"));

            if (selectAll)
            {

                if (!_selectUnselect.GetAttribute("class").Contains("selected"))
                {
                    action.MoveToElement(_selectUnselect).Perform();
                    _selectUnselect.Click();
                }
                else
                {
                    _selectUnselect.Click();
                    _selectUnselect.Click();

                }
            }
            else
            {
                if (_selectUnselect.GetAttribute("class").Contains("selected"))
                {
                    action.MoveToElement(_selectUnselect).Perform();
                    _selectUnselect.Click();
                }
                else
                {
                    _selectUnselect.Click();
                    _selectUnselect.Click();
                }
            }
            Thread.Sleep(2000);
            WaitForLoad();
        }


        public void ClickOnInformations()
        {
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(30));
            var _informationsBtn = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id=\"EditUserInformation\"]")));
            _informationsBtn.SendKeys(Keys.ArrowUp);
            _informationsBtn.Click();
            WaitForLoad();
        }
        public void ClickOnAffectedSite()
        {
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(30));
            var _affectedSitesBtn = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id=\"EditUserSites\"]")));
            _affectedSitesBtn.SendKeys(Keys.ArrowUp);
            _affectedSitesBtn.Click();
            WaitForLoad();
        }
    }
}
