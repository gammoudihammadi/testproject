using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Accounting.FreePrice
{
    public class FreePriceDetailsPage : PageBase
    {

        public FreePriceDetailsPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // __________________________________________ Constantes _________________________________________

        public const string SELLING_PRICE = "SellingPrice";
        public const string TAX_TYPE = "TaxTypeId";
        public const string IS_ACTIVE = "IsActive";
        public const string NAME = "Name";
        public const string FOREIGN_NAME = "CommercialName";
        public const string CODE = "Code";
        public const string WORK_SHOP = "/html/body/div[3]/div/div/div[2]/div/div/div/div/div/form/div[8]/div[2]/select";
        public const string WORK_SHOP_SELECTED = "/html/body/div[3]/div/div/div[2]/div/div/div/div/div/form/div[8]/div[2]/select/option[@selected='selected']";
        public const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";
        public const string SELLING_PRICE_VALUE = "/html/body/div[3]/div/div/div[2]/div/div/div/div/div/form/div[6]/div[2]/input";
        public const string TAX_TYPE_VALUE = "/html/body/div[3]/div/div/div[2]/div/div/div/div/div/form/div[7]/div[2]/select";
        public const string TAX_TYPE_SELECTED = "/html/body/div[3]/div/div/div[2]/div/div/div/div/div/form/div[7]/div[2]/select/option[@selected='selected']";
        public const string IS_ACTIVE_XPATH = "/html/body/div[3]/div/div/div[2]/div/div/div/div/div/form/div[9]/div[2]/div/input";

        // __________________________________________ Variables __________________________________________

        [FindsBy(How = How.Id, Using = SELLING_PRICE)]
        private IWebElement _sellingPrice;

        [FindsBy(How = How.Id, Using = TAX_TYPE)]
        private IWebElement _taxType;

        [FindsBy(How = How.Id, Using = IS_ACTIVE)]
        private IWebElement _isActive;

        [FindsBy(How = How.Id, Using = FOREIGN_NAME)]
        private IWebElement _foreignName;

        [FindsBy(How = How.Id, Using = CODE)]
        private IWebElement _code;

        [FindsBy(How = How.Id, Using = NAME)]
        private IWebElement _name;

        // __________________________________________ Méthodes ___________________________________________

        public string GetSellingPrice()
        {
            _sellingPrice = WaitForElementIsVisible(By.XPath(SELLING_PRICE_VALUE));
            return _sellingPrice.GetAttribute("value");
        }

        public void SetSellingPrice(string value)
        {
            _sellingPrice = WaitForElementIsVisible(By.Id(SELLING_PRICE));
            _sellingPrice.SetValue(ControlType.TextBox, value);

            WaitForLoad();
        }
        public void SetForeignName(string value)
        {
            _foreignName = WaitForElementIsVisible(By.Id(FOREIGN_NAME));
            _foreignName.SetValue(ControlType.TextBox ,value);

            WaitForLoad();
        }
        public void SetName(string value)
        {
            _name = WaitForElementIsVisible(By.Id(NAME));
            _name.SetValue(ControlType.TextBox, value);

            WaitForLoad();
        }
        public string GetCode()
        {
            _code = WaitForElementIsVisible(By.Id(CODE));
            return _code.GetAttribute("text");
        }

        public void SetTaxType(string value)
        {
            _taxType = WaitForElementIsVisible(By.Id(TAX_TYPE));
            _taxType.SetValue(ControlType.DropDownList, value);

            WaitForLoad();
        }

        public string GetTaxType()
        {
            _taxType = WaitForElementIsVisible(By.XPath(TAX_TYPE_SELECTED));
            return _taxType.GetAttribute("text");
        }
        public string GetWorkShop()
        {
            var workShop = WaitForElementIsVisible(By.XPath(WORK_SHOP_SELECTED));
            return workShop.GetAttribute("text");
        }
        public string GetForeignName()
        {
            var foreignName = WaitForElementIsVisible(By.Id(FOREIGN_NAME));
            return foreignName.GetAttribute("value");
        }

        public void SetActive(bool value)
        {
            _isActive = WaitForElementExists(By.Id(IS_ACTIVE));
            _isActive.SetValue(ControlType.CheckBox, value);

            WaitForLoad();
        }

        public bool IsActive()
        {
            _isActive = WaitForElementExists(By.Id(IS_ACTIVE));

            if (_isActive.Selected.Equals("true"))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public void EditDetailFreePrice(string foreignNameInput, string sellingPriceInput, string taxeTypeInput, string workShopInput, bool isActiveInput)
        {
            SetForeignName(foreignNameInput);
            SetSellingPrice(sellingPriceInput);
            SetTaxType(taxeTypeInput);

            var workshop = WaitForElementIsVisible(By.XPath(WORK_SHOP));
            workshop.SetValue(ControlType.DropDownList, workShopInput);
            SetActive(isActiveInput);
        }

        public void BackToList()
        {
            var backToListBtn = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            backToListBtn.Click();
        }

        public bool VerifyDataEdited(string foreignNameInput, string sellingPriceInput, string taxeTypeInput, string workShopInput, bool isActiveInput)
        {
            var foreignName = GetForeignName();
            var sellingPrice = GetSellingPrice();
            var taxeType = GetTaxType();
            var workshop = GetWorkShop();
            var active = IsActive();

            if(foreignName != foreignNameInput || sellingPrice != sellingPriceInput || taxeType != taxeTypeInput
                || workshop != workShopInput || isActiveInput != active)
            {
                return false;
            }
            return true;
        }
        public bool VerifyIsActiveFilter()
        {
            var checkbox = _webDriver.FindElement(By.XPath(IS_ACTIVE_XPATH));
            if(checkbox.GetAttribute("checked") != null)
            {
                return true;
            }
            return false;
        }
    }
}
