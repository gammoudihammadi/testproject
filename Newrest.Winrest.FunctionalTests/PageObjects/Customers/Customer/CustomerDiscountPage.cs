using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.Customer
{
    public class CustomerDiscountPage : PageBase
    {

        public CustomerDiscountPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //_______________________________________________________Constantes____________________________________________________________

        // General
        private const string NEW_DISCOUNT = "//*[@id=\"tabContentDetails\"]/div/div/div/div/a[2]";
        private const string SITE = "drop-down-sites";
        private const string SERVICE_CATEGORY = "drop-down-categories";
        private const string DISCOUNT_RATE = "Discount";
        private const string DISCOUNT_INCREASE = "IsDiscountIncrease";

        private const string CREATE = "//button[text()='Create']";
        private const string UPDATE = "/html/body/div[4]/div/div/div/div/form/div[3]/button[2]";

        // Tableau
        private const string EDIT_DISCOUNT = "//*[@id=\"list-item-with-action\"]/div/div[2]/div/div/div[6]/div/a[2]";
        private const string DELETE_DISCOUNT = "//*[@id=\"list-item-with-action\"]/div/div[2]/div/div/div[6]/div/a[1]";
        private const string CONFIRM_DELETE = "/html/body/div[11]/div/div/div[3]/a[1]";
        private const string FIRST_DISCOUNT = "//*[@id=\"list-item-with-action\"]/div/div[2]/div/div/div[2]";
        private const string DISCOUNT = "//*[@id=\"list-item-with-action\"]/div/div[2]/div/div/div[4]";
        private const string INCREASE = "/html/body/div[3]/div/div/div[2]/div/div/div[2]/div/div/div[2]/div/div/div[5]/input";


        //_______________________________________________________Variables_____________________________________________________________

        // General
        [FindsBy(How = How.XPath, Using = NEW_DISCOUNT)]
        private IWebElement _newDiscount;

        [FindsBy(How = How.Id, Using = SITE)]
        private IWebElement _discountSite;

        [FindsBy(How = How.Id, Using = SERVICE_CATEGORY)]
        private IWebElement _discountServiceCategory;

        [FindsBy(How = How.Id, Using = DISCOUNT_RATE)]
        private IWebElement _discountRate;

        [FindsBy(How = How.Id, Using = DISCOUNT_INCREASE)]
        private IWebElement _discountIncrease;

        [FindsBy(How = How.XPath, Using = CREATE)]
        private IWebElement _createBtn;

        [FindsBy(How = How.XPath, Using = UPDATE)]
        private IWebElement _updateBtn;

        // Tableau
        [FindsBy(How = How.XPath, Using = EDIT_DISCOUNT)]
        private IWebElement _editDiscount;

        [FindsBy(How = How.XPath, Using = DELETE_DISCOUNT)]
        private IWebElement _deleteDiscount;

        [FindsBy(How = How.XPath, Using = CONFIRM_DELETE)]
        private IWebElement _confirmDelete;

        [FindsBy(How = How.XPath, Using = FIRST_DISCOUNT)]
        private IWebElement _firstDiscount;

        [FindsBy(How = How.XPath, Using = DISCOUNT)]
        private IWebElement _discount;

        [FindsBy(How = How.XPath, Using = INCREASE)]
        private IWebElement _increase;

        //_______________________________________________________Pages_________________________________________________________________

        public void CreateDiscount(string discountSite, string serviceCategory, string discountRate, bool isIncrease)
        {
            _newDiscount = WaitForElementIsVisible(By.XPath(NEW_DISCOUNT));
            _newDiscount.Click();
            WaitForLoad();

            _discountSite = WaitForElementIsVisible(By.Id(SITE));
            _discountSite.SetValue(ControlType.DropDownList, discountSite);
            WaitForLoad();

            _discountServiceCategory = WaitForElementIsVisible(By.Id(SERVICE_CATEGORY));
            _discountServiceCategory.SetValue(ControlType.DropDownList, serviceCategory);

            _discountRate = WaitForElementIsVisible(By.Id(DISCOUNT_RATE));
            _discountRate.SetValue(ControlType.TextBox, discountRate);

            if (isIncrease)
            {
                _discountIncrease = WaitForElementIsVisible(By.Id(DISCOUNT_INCREASE));
                _discountIncrease.Click();
            }

            _createBtn = WaitForElementIsVisible(By.XPath(CREATE));
            _createBtn.Click();
            WaitForLoad();
        }

        public void EditDiscount(string discountRate)
        {
            _editDiscount = WaitForElementIsVisible(By.XPath(EDIT_DISCOUNT));
            _editDiscount.Click();
            WaitForLoad();

            _discountRate = WaitForElementIsVisible(By.Id(DISCOUNT_RATE));
            _discountRate.SetValue(ControlType.TextBox, discountRate);

            _updateBtn = WaitForElementIsVisible(By.XPath(UPDATE));
            _updateBtn.Click();
            WaitForLoad();
        }

        // Tableau
        public bool IsDiscountDisplayed()
        {
            if (isElementVisible(By.XPath(FIRST_DISCOUNT)))
            {
                _firstDiscount = _webDriver.FindElement(By.XPath(FIRST_DISCOUNT));
                return _firstDiscount.Displayed;
            }
            else
            {
                return false;
            }
        }

        public bool IsIncrease()
        {
            bool valueBool = true;

            if (isElementVisible(By.XPath(INCREASE)))
            {
                _increase = _webDriver.FindElement(By.XPath(INCREASE));

                if (_increase.GetAttribute("checked") != "true")
                    valueBool = false;
            }
            else
            {
                valueBool = false;
            }
            return valueBool;
        }

        public string GetFirstDiscount()
        {
            _discount = WaitForElementIsVisible(By.XPath(DISCOUNT));
            return _discount.Text;
        }

        public void DeleteDiscount()
        {
            _deleteDiscount = WaitForElementIsVisible(By.XPath(DELETE_DISCOUNT));
            _deleteDiscount.Click();
            WaitForLoad();

            _confirmDelete = WaitForElementIsVisible(By.XPath(CONFIRM_DELETE));
            _confirmDelete.Click();
            WaitForLoad();
        }

    }
}
