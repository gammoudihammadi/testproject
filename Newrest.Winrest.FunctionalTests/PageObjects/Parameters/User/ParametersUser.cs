using DocumentFormat.OpenXml.Bibliography;
using iText.Commons.Utils;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Production;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Parameters.User
{
    public class ParametersUser : PageBase
    {
        // ___________________________________ Constantes ______________________________________________

        private const string RADIO_BTN_SITE_CREATED = "{0}";
        private const string SEARCH_USER = "tbSearchPattern";
        private const string SELECT_USER = "//*[@id=\"82aebc96-5044-48b2-94b8-1323b58671d4\"]/td/h4";
        private const string SELECT_UNSELECT = "selectAllSitesBtn";
        private const string AFFECTED_SITES = "EditUserSites";
        private const string SITE_TO_SELECT = "//label[contains(text(),'{0}')]";
        private const string PRODUCTION = "//*[@id=\"tabContentUsersRights\"]/div[2]/div[1]/table/tbody/tr[42]";
        private const string PRODUCTIONINLISTE = "/html/body/div[2]/div/div/div/div[2]/div[2]/div/div[3]/div/div[4]/div[1]/div/div[1]/table/tbody/tr";
        private const string PRODUCTIONMANAGEMENT = "/html/body/div[2]/div/div/div/div[2]/div[2]/div/div[3]/div/div[4]/div[2]/div/div[5]/div/div[1]/div/div[1]/table/tbody/tr/td";
        private const string RENAMEFAVORITE = "//*[@id=\"21210\"]";
        private const string RENAMEALLFAVORITE = "//*[@id=\"21211\"]";

        public const string USER_TAB = "//*[@id=\"div-body\"]/div/ul/li[1]/a";
        public const string ROLES_TAB = "//*[@id=\"div-body\"]/div/ul/li[2]/a";
        public const string RIGHTS_TAB = "//*[@id=\"div-body\"]/div/ul/li[3]/a";
        public const string AFFECTED_TO_TAB = "//*[@id=\"div-body\"]/div/ul/li[4]/a";
        public const string MOBILE_USAGE_TAB = "//*[@id=\"div-body\"]/div/ul/li[5]/a";
        public const string TABLET_USAGE_TAB = "//*[@id=\"div-body\"]/div/ul/li[6]/a";


        // ___________________________________ Variables _______________________________________________


        [FindsBy(How = How.XPath, Using = RENAMEFAVORITE)]
        private IWebElement _renamefavorite;

        [FindsBy(How = How.XPath, Using = RENAMEALLFAVORITE)]
        private IWebElement _renameallfavorite;
        [FindsBy(How = How.Id, Using = SEARCH_USER)]
        private IWebElement _searchUser;

        [FindsBy(How = How.XPath, Using = SELECT_USER)]
        private IWebElement _selectUser;
        [FindsBy(How = How.XPath, Using = PRODUCTION)]
        private IWebElement _production;
        [FindsBy(How = How.XPath, Using = PRODUCTIONINLISTE)]
        private IWebElement _productionlist;
        [FindsBy(How = How.XPath, Using = PRODUCTIONMANAGEMENT)]
        private IWebElement _productionmanagement;
        [FindsBy(How = How.Id, Using = SELECT_UNSELECT)]
        private IWebElement _selectUnselect;

        [FindsBy(How = How.Id, Using = AFFECTED_SITES)]
        private IWebElement _affectedSites;

        [FindsBy(How = How.XPath, Using = SITE_TO_SELECT)]
        private IWebElement _siteToSelect;

        // __________________________________ Méthodes __________________________________________________

        public ParametersUser(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {

        }

        public void SearchAndSelectUser(string userName)
        {
            //Recherche du user
            _searchUser = WaitForElementIsVisible(By.Id(SEARCH_USER));
            _searchUser.SetValue(ControlType.TextBox, userName);
            _searchUser.SendKeys(Keys.Enter);
            WaitForLoad();
            _selectUser = WaitForElementIsVisible(By.XPath(SELECT_USER));
            _selectUser.Click();
            WaitForLoad();
        }

        public void SearchAndSelectUserByClickingFirstRow(string userName)
        {
            _searchUser = WaitForElementIsVisible(By.Id(SEARCH_USER));
            _searchUser.SetValue(ControlType.TextBox, userName);
            _searchUser.SendKeys(Keys.Enter);
            WaitForLoad();
            _selectUser = WaitForElementIsVisible(By.XPath("/html/body/div[2]/div/div/div/div[1]/div[1]/div[4]/table/tbody/tr"));
            _selectUser.Click();
            WaitForLoad();
        }

        public void GiveSiteRightsToUser(string id, bool hasPermission = true,string SiteName =null)
        {

            
            var tableBody = _webDriver.FindElement(By.XPath("/html/body/div[2]/div/div/div/div[1]/div[2]/div/div/div/div/table/tbody"));

            var rows = tableBody.FindElements(By.TagName("tr"));

            foreach (var row in rows)
            {
                var secondTd = row.FindElement(By.XPath("td[2]"));
                var labelElement = secondTd.FindElement(By.XPath(".//label"));

                if (secondTd != null && SiteName.ToUpper() == labelElement.Text.ToUpper())
                {
                    labelElement.Click();
                    break;
                }
            }

            WaitPageLoading();
            WaitForLoad();
        }

        public void SelectUnselectAllSites(bool selectAll = true)
        {
            _selectUnselect = WaitForElementIsVisible(By.Id(SELECT_UNSELECT));
            WaitForLoad();

            if (selectAll)
            {

                if (_selectUnselect.Text.Equals("SelectAll"))
                {
                    _selectUnselect.Click();
                }
                else
                {
                    _selectUnselect.Click();
                    WaitPageLoading();
                    WaitForLoad();
                    _selectUnselect = WaitForElementIsVisible(By.Id(SELECT_UNSELECT));
                    _selectUnselect.Click();
                }
            }
            else
            {
                if (_selectUnselect.Text.Equals("UnselectAll"))
                {
                    _selectUnselect.Click();
                }
                else
                {
                    _selectUnselect.Click();
                    WaitPageLoading();
                    WaitForLoad();
                    _selectUnselect = WaitForElementIsVisible(By.Id(SELECT_UNSELECT));
                    _selectUnselect.Click();
                }
            }

            WaitPageLoading();
            WaitForLoad();
        }

        public void SelectOneSite(string siteName)
        {
            _siteToSelect = WaitForElementIsVisible(By.XPath(string.Format(SITE_TO_SELECT, siteName)));
            _siteToSelect.Click();
            WaitPageLoading();
            WaitForLoad();
        }

        public void ActivateSite(string siteName)
        {
            _siteToSelect = WaitForElementIsVisible(By.XPath(string.Format(SITE_TO_SELECT, siteName)));
            var id = _siteToSelect.GetAttribute("for");
            var inputRelated = WaitForElementExists(By.XPath(string.Format("//*[@id=\"{0}\"]", id)));
            if (!inputRelated.Selected)
            {
                _siteToSelect.Click();
                WaitPageLoading();
                WaitForLoad();
            }
        }

        public void ClickOnAffectedSite()
        {
            _affectedSites = WaitForElementIsVisible(By.Id(AFFECTED_SITES));
            _affectedSites.SendKeys(Keys.ArrowUp);
            _affectedSites.Click();
            WaitPageLoading();
            WaitForLoad();
        }

        public void ClickProduction()
        {
            _production = WaitForElementExists(By.XPath(PRODUCTION));
            _production.Click();
            WaitPageLoading();
            WaitForLoad();
        }
        public void ClickProductionInListe()
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)_webDriver;
            js.ExecuteScript("window.scrollTo(0, 0);");
            Thread.Sleep(1000); 
            IWebElement productionListElement = WaitForElementExists(By.XPath(PRODUCTIONINLISTE));
            productionListElement.Click();
            WaitPageLoading();
            WaitForLoad();
        }


        public void ClickProductionManagement()
        {
            _productionmanagement = WaitForElementExists(By.XPath(PRODUCTIONMANAGEMENT));
            _productionmanagement.Click();
            WaitPageLoading();
            WaitForLoad();
        }
        public void CocherRenameFavoritetrenameallfavorites()
        {
            _renamefavorite = WaitForElementExists(By.XPath(RENAMEFAVORITE));
            _renamefavorite.SetValue(ControlType.CheckBox, true);
            _renameallfavorite = WaitForElementExists(By.XPath(RENAMEALLFAVORITE));
            _renameallfavorite.SetValue(ControlType.CheckBox, true);
        }
        public void UncheckRenameFavoritetrenameallfavorites()
        {
            _renamefavorite = WaitForElementExists(By.XPath(RENAMEFAVORITE));
            _renamefavorite.SetValue(ControlType.CheckBox, false);
            _renameallfavorite = WaitForElementExists(By.XPath(RENAMEALLFAVORITE));
            _renameallfavorite.SetValue(ControlType.CheckBox, false);
        }
        public void SwitchTabs(string xPath)
        {
            IWebElement tab = WaitForElementIsVisible(By.XPath($"{xPath}"));
            tab.Click();
            WaitForLoad();
        }
        
    }
    
}
