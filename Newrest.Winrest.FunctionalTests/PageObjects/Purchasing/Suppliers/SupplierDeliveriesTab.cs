using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Suppliers
{
    public class SupplierDeliveriesTab : PageBase
    {
        public SupplierDeliveriesTab(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {

        }

        private const string MIN_AMOUNT = "DeliveriesSites_1__MinimumOrder";
        private const string SHIPPING_COST = "DeliveriesSites_1__ShippingCost";
        private const string PREPA_DELAY = "DeliveriesSites_1__PrepaDelay";
        private const string COMMENT_SHOW = "//*[@id='header_9']/div/div[10]/div[1]/div/img";
        private const string COMMENT = "DeliveriesSites_1__Comment";
        private const string MONDAY = "//*[@id=\"header_9\"]/div/div[10]/div[1]/label[1]";
        private const string TUESDAY = "//*[@id=\"header_9\"]/div/div[10]/div[1]/label[2]";
        private const string WEDNESDAY = "//*[@id=\"header_9\"]/div/div[10]/div[1]/label[3]";
        private const string THURSDAY = "DeliveriesSites_1__Thursday";
        private const string FRIDAY = "DeliveriesSites_1__Friday";
        private const string SATURDAY = "DeliveriesSites_1__Saturday";
        private const string SUNDAY = "DeliveriesSites_1__Sunday";
        private const string DELIVERED_SITE_CHECKBOX = "//p[contains(text(), '{0}')]/../..//input[@class = 'checkbox-is-delivered']";
        private const string GO_TO_EDI_TAB = "hrefTabContentDetailsEDI";


        [FindsBy(How = How.Id, Using = MIN_AMOUNT)]
        private IWebElement _minAmount;

        [FindsBy(How = How.Id, Using = SHIPPING_COST)]
        private IWebElement _shippingCost;

        [FindsBy(How = How.Id, Using = PREPA_DELAY)]
        private IWebElement _prepaDelay;

        [FindsBy(How = How.XPath, Using = COMMENT_SHOW)]
        private IWebElement _commentShow;

        [FindsBy(How = How.Id, Using = COMMENT)]
        private IWebElement _comment;

        [FindsBy(How = How.XPath, Using = MONDAY)]
        private IWebElement _monday;

        [FindsBy(How = How.XPath, Using = TUESDAY)]
        private IWebElement _tuesday;

        [FindsBy(How = How.XPath, Using = WEDNESDAY)]
        private IWebElement _wednesday;

        [FindsBy(How = How.Id, Using = THURSDAY)]
        private IWebElement _thursday;

        [FindsBy(How = How.Id, Using = FRIDAY)]
        private IWebElement _friday;

        [FindsBy(How = How.Id, Using = SATURDAY)]
        private IWebElement _saturday;

        [FindsBy(How = How.Id, Using = SUNDAY)]
        private IWebElement _sunday;

        [FindsBy(How = How.XPath, Using = DELIVERED_SITE_CHECKBOX)]
        private IWebElement _deliveredSiteCheckbox;

        [FindsBy(How = How.Id, Using = GO_TO_EDI_TAB)]
        private IWebElement _goToEdiTab;

        public string GetMinAmount()
        {
            _minAmount = WaitForElementIsVisible(By.Id(MIN_AMOUNT));
            return _minAmount.GetAttribute("value");
        }

        public string GetShippingCost()
        {
            _shippingCost = WaitForElementIsVisible(By.Id(SHIPPING_COST));
            return _shippingCost.GetAttribute("value");
        }

        public string GetComment(string site)
        { 
                _commentShow = WaitForElementIsVisible(By.XPath(string.Format("//*[starts-with(@id, 'header_')]/div/div[1]/p[contains(text(), '{0}')]/../../div[10]/div[2]/div/img", site)));
                _commentShow.Click();
            
             _comment = WaitForElementIsVisible(By.Id(COMMENT));

            var text = _comment.Text;
                _commentShow = WaitForElementIsVisible(By.XPath("//*[@id=\"header_\"]/div/div[10]/div[2]/div/img"));
            
            _commentShow.Click();
            return text;
        }

        public string GetMonday()
        {
            _monday = WaitForElementExists(By.XPath(MONDAY));
            var MondayChecked = _monday.GetAttribute("checked");
            return MondayChecked == "true" ? "TRUE" : "FALSE";
        }

        public void SetDeliveryWeekDays()
        {
            _monday = WaitForElementExists(By.XPath(MONDAY));
            _monday.Click();
            _tuesday = WaitForElementExists(By.XPath(TUESDAY));
            _tuesday.Click();
            _wednesday = WaitForElementExists(By.XPath(WEDNESDAY));
            _wednesday.Click();
        }
        public string GetTuesday()
        {
            _tuesday = WaitForElementExists(By.XPath(TUESDAY));
            var TuesdayChecked = _tuesday.GetAttribute("checked");
            return TuesdayChecked == "true" ? "TRUE" : "FALSE";
        }

        public string GetWednesday()
        {
            _wednesday = WaitForElementExists(By.XPath(WEDNESDAY));
            var WednesdayChecked = _wednesday.GetAttribute("checked");
            return WednesdayChecked == "true" ? "TRUE" : "FALSE";
        }

        public string GetThursday()
        {
            _thursday = WaitForElementExists(By.Id(THURSDAY));
            var ThursdayChecked = _thursday.GetAttribute("checked");
            return ThursdayChecked == "true" ? "TRUE" : "FALSE";
        }

        public string GetFriday()
        {
            _friday = WaitForElementExists(By.Id(FRIDAY));
            var FridayChecked = _friday.GetAttribute("checked");
            return FridayChecked == "true" ? "TRUE" : "FALSE";
        }

        public string GetSaturday()
        {
            _saturday = WaitForElementExists(By.Id(SATURDAY));
            var SaturdayChecked = _saturday.GetAttribute("checked");
            return SaturdayChecked == "true" ? "TRUE" : "FALSE";
        }

        public string GetSunday()
        {
            _sunday = WaitForElementExists(By.Id(SUNDAY));
            var SundayChecked = _sunday.GetAttribute("checked");
            return SundayChecked == "true" ? "TRUE" : "FALSE";
        }

        public SupplierContactTab ClickOnContactsTab()
        {
            var tab = WaitForElementExists(By.Id("hrefTabContentContacts"));
            tab.Click();
            WaitForLoad();
            return new SupplierContactTab(_webDriver, _testContext);
        }

        public void SelectedDeliveredSites(string site, bool selected)
        {
            _deliveredSiteCheckbox = WaitForElementExists(By.XPath(string.Format(DELIVERED_SITE_CHECKBOX, site)));
            _deliveredSiteCheckbox.SetValue(PageBase.ControlType.CheckBox, selected);
            WaitForLoad();
        }
        public void SetAmountDeliveredSites(string site, string value)
        {
            var selectdeliveredSite= WaitForElementExists(By.XPath(string.Format(DELIVERED_SITE_CHECKBOX, site)));
            _minAmount = WaitForElementIsVisible(By.Id(MIN_AMOUNT));     
            _minAmount.SetValue(ControlType.TextBox, value);     
            WaitForLoad();
        }

        public void SetShippingCost(string site, string value)
        {
            var selectdeliveredSite = WaitForElementExists(By.XPath(string.Format(DELIVERED_SITE_CHECKBOX, site)));
            _shippingCost = WaitForElementIsVisible(By.Id(SHIPPING_COST));
            _shippingCost.SetValue(ControlType.TextBox, value);
            WaitForLoad();
        }

        public void SetPrepaDelay(string site, string value)
        {
            var selectdeliveredSite = WaitForElementExists(By.XPath(string.Format(DELIVERED_SITE_CHECKBOX, site)));
            _prepaDelay = WaitForElementIsVisible(By.Id(PREPA_DELAY));
            _prepaDelay.SetValue(ControlType.TextBox, value);
            WaitForLoad();
        }

        public bool CheckSelectedDeliveredSites(string site)
        {
            _deliveredSiteCheckbox = WaitForElementExists(By.XPath(string.Format(DELIVERED_SITE_CHECKBOX, site)));
            if (_deliveredSiteCheckbox.GetAttribute("Selected") == "true")
            {
                return true;
            }
            return false;
        }

        public string[]  GetAllColorsDeliveryDay(string site)
        {
            string[] deliveryday = new string[7];

            WaitForLoad();

            var monday = WaitForElementIsVisible(By.XPath(string.Format("//p[contains(text(), '{0}')]/../../div[10]/div/label[text()='Mon']", site)));

            deliveryday[0] = monday.GetCssValue("border");

            WaitForLoad();
          
            var tuesday = WaitForElementIsVisible(By.XPath(string.Format("//p[contains(text(), '{0}')]/../../div[10]/div/label[text()='Tue']", site)));

            deliveryday[1] = tuesday.GetCssValue("border");

            WaitForLoad();
            
            var wednesday = WaitForElementIsVisible(By.XPath(string.Format("//p[contains(text(), '{0}')]/../../div[10]/div/label[text()='Wed']", site)));

            deliveryday[2] = wednesday.GetCssValue("border");

            WaitForLoad();

            var thursday = WaitForElementIsVisible(By.XPath(string.Format("//p[contains(text(), '{0}')]/../../div[10]/div/label[text()='Thu']", site)));

            deliveryday[3] = thursday.GetCssValue("border");

            WaitForLoad();
            
            var friday = WaitForElementIsVisible(By.XPath(string.Format("//p[contains(text(), '{0}')]/../../div[10]/div/label[text()='Fri']", site)));

            deliveryday[4] = friday.GetCssValue("border");

            WaitForLoad();
            
            var saturday = WaitForElementIsVisible(By.XPath(string.Format("//p[contains(text(), '{0}')]/../../div[10]/div/label[text()='Sat']", site)));

            deliveryday[5] = saturday.GetCssValue("border");

            WaitForLoad();
            
            var sunday = WaitForElementIsVisible(By.XPath(string.Format("//p[contains(text(), '{0}')]/../../div[10]/div/label[text()='Sun']", site)));

            deliveryday[6] = sunday.GetCssValue("border");

            WaitForLoad();

            return deliveryday;

        }

        public SupplierEDITab GoToEdiTab()
        {
            _goToEdiTab = WaitForElementIsVisible(By.Id(GO_TO_EDI_TAB));
            _goToEdiTab.Click();
            WaitForLoad();
            return new SupplierEDITab(_webDriver, _testContext);
        }
    }
}
