using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Globalization;
using System.Linq;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.DeliveryRound
{
    public class DeliveryRoundGeneralInformationPage : PageBase
    {
        public DeliveryRoundGeneralInformationPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //__________________________________Constantes_________________________________________

        // General
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span";

        // Onglets
        private const string DELIVERY_TAB = "hrefTabContentRound";

        // Informations
        private const string DELIVERY_ROUND_NAME = "first";
        private const string SITE = "SiteId";
        private const string START_DATE = "StartDate";
        private const string END_DATE = "EndDate";
        private const string ACTIVATED_BTN = "IsActive";

        private const string ERROR_MESSAGE = "//*[@id=\"deliveryround-filter-form\"]/div/div[1]/div[1]/div/div[1]/div/span";


        //__________________________________Variables__________________________________________

        // General
        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;

        // Onglets
        [FindsBy(How = How.Id, Using = DELIVERY_TAB)]
        private IWebElement _deliveriesTab;

        // Informations
        [FindsBy(How = How.Id, Using = DELIVERY_ROUND_NAME)]
        private IWebElement _deliveryRoundName;

        [FindsBy(How = How.Id, Using = SITE)]
        private IWebElement _site;

        [FindsBy(How = How.Id, Using = START_DATE)]
        private IWebElement _startDate;

        [FindsBy(How = How.Id, Using = END_DATE)]
        private IWebElement _endDate;

        [FindsBy(How = How.Id, Using = ACTIVATED_BTN)]
        private IWebElement _activatedBtn;

        // __________________________________ Méthodes _________________________________________

        // General
        public DeliveryRoundPage BackToList()
        {
            _backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            _backToList.Click();
            WaitForLoad();

            return new DeliveryRoundPage(_webDriver, _testContext);
        }

        // Onglets
        public DeliveryRoundDeliveryPage ClickOnDeliveryTab()
        {
            _deliveriesTab = WaitForElementIsVisible(By.Id(DELIVERY_TAB));
            _deliveriesTab.Click();
            WaitForLoad();

            return new DeliveryRoundDeliveryPage(_webDriver, _testContext);
        }

        // Informations
        public void ClickActive(bool active)
        {
            _activatedBtn = WaitForElementExists(By.Id(ACTIVATED_BTN));
            _activatedBtn.SetValue(ControlType.CheckBox, active);
            WaitPageLoading();
            WaitForLoad();
        }

        public void UpdateGeneralInformation(string name, string site, DateTime startDateTime, DateTime endDateTime)
        {
            // Définition du nom
            _deliveryRoundName = WaitForElementIsVisible(By.Id(DELIVERY_ROUND_NAME));
            WaitPageLoading();
            _deliveryRoundName.SetValue(ControlType.TextBox, name);
            WaitPageLoading();
            WaitForLoad();

            // Définition du site
            _site = WaitForElementIsVisible(By.Id(SITE));
            _site.Click();
            var element = WaitForElementExists(By.XPath("//*[@id=\"SiteId\"]/option[@value = '" + site + "']"));
            WaitPageLoading();
            _site.SetValue(ControlType.DropDownList, element.Text);
            _site.Click();
            WaitPageLoading();
            WaitForLoad();

            // Définition de Start Date
            WaitPageLoading();
            WaitForLoad();
            _startDate = WaitForElementIsVisible(By.Id(START_DATE));
            WaitPageLoading();
            _startDate.SetValue(ControlType.DateTime, startDateTime);
            WaitPageLoading();
            WaitForLoad();

            // Définition de End Date
            _endDate = WaitForElementIsVisible(By.Id(END_DATE));
            WaitPageLoading();
            _endDate.SetValue(ControlType.DateTime, endDateTime);
            _endDate.SendKeys(Keys.Tab);
            WaitPageLoading();
            WaitForLoad();
        }

        public void UpdateDeliveryRoundName(string name)
        {
            // Définition du nom
            _deliveryRoundName = WaitForElementIsVisible(By.Id(DELIVERY_ROUND_NAME));
            _deliveryRoundName.SetValue(ControlType.TextBox, name);
            WaitPageLoading();
            WaitForLoad();
        }

        public string GetDeliveryRoundName()
        {
            _deliveryRoundName = WaitForElementIsVisible(By.Id(DELIVERY_ROUND_NAME));
            return _deliveryRoundName.GetAttribute("value");
        }

        public DateTime GetDeliveryRoundStartDate(string dateFormat)
        {
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            _startDate = WaitForElementIsVisible(By.Id(START_DATE));
            return DateTime.Parse(_startDate.GetAttribute("value"), ci);
        }

        public DateTime GetDeliveryRoundEndDate(string dateFormat)
        {
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            _endDate = WaitForElementIsVisible(By.Id(END_DATE));
            return DateTime.Parse(_endDate.GetAttribute("value"), ci);
        }

        public string GetDeliveryRoundSite()
        {
            var newSite = new SelectElement(_webDriver.FindElement(By.Id(SITE)));
            return newSite.AllSelectedOptions.FirstOrDefault()?.GetAttribute("value");
        }

        public bool GetErrorMessageAlreadyExisting()
        {
            if (isElementVisible(By.XPath(ERROR_MESSAGE)))
            {
                WaitForElementIsVisible(By.XPath(ERROR_MESSAGE));
                return true;
            }
            else
            {
                return false;
            }
        }

        public void SetDeliveryRoundEndDate(DateTime endDateTime)
        {
            // Définition de End Date
            _endDate = WaitForElementIsVisible(By.Id(END_DATE));
            _endDate.SetValue(ControlType.DateTime, endDateTime);
            _endDate.SendKeys(Keys.Tab);
            WaitPageLoading();
            WaitForLoad();
        }
        
    }
}