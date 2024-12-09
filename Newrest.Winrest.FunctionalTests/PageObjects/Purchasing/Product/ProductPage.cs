using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Product
{
    public class ProductPage : PageBase
    {
        // _____________________________________________ Constantes ________________________________________________
        private const string PLUS_BUTTON = "//*[@id=\"tabContentProductContainer\"]/div/div/div[2]/button";
        private const string NEW_PRODUCT = "purchasing-product-createbtn";
        private const string ITEM_INPUT = "item-input";
        private const string SERVICE_INPUT = "//*[@id=\"service-input\"]";
        private const string ADDNEW = "createNew";
        private const string ADD = "create";
        private const string CLOSE = "//*[@id=\"create-product-form\"]/div[3]/button[2]";
        private const string FROM = "create-product-form";
        private const string DELETE = "purchasing-product-delete1";
        private const string CONFIRMDELETE = "dataConfirmOK";
        private const string EDIT_SERVICE_BTN = "service-event";
        private const string EDIT_ITEM_BTN = "item-event";
        private const string ADD_NEW_PRODUCT = "purchasing-product-createbtn";

        // Filtres
        private const string RESET_FILTER_DEV = "ResetFilter";
        private const string SEARCH_FILTER = "SearchPatternWithAutocomplete";
        private const string FIRST_RESULT_SEARCH = "//*[@id=\"item-filter-form\"]/div[2]/span/div/div/div/strong[text()='{0}']";


        // ___________________________________________ Filtres __________________________________________________

        [FindsBy(How = How.Id, Using = ADD_NEW_PRODUCT)]
        private IWebElement _addNewProductButton;
        [FindsBy(How = How.XPath, Using = RESET_FILTER_DEV)]
        private IWebElement _resetFilterDev;
        [FindsBy(How = How.Id, Using = SEARCH_FILTER)]
        private IWebElement _searchFilter;
        [FindsBy(How = How.Id, Using = PLUS_BUTTON)]
        private IWebElement _plusButton;
        [FindsBy(How = How.Id, Using = NEW_PRODUCT)]
        private IWebElement _newProduct;
        [FindsBy(How = How.Id, Using = ITEM_INPUT)]
        private IWebElement _itemInput;
        [FindsBy(How = How.Id, Using = SERVICE_INPUT)]
        private IWebElement _serviceInput;
        [FindsBy(How = How.Id, Using = ADDNEW)]
        private IWebElement _add_new;
        [FindsBy(How = How.Id, Using = ADD)]
        private IWebElement _add;
        [FindsBy(How = How.Id, Using = CLOSE)]
        private IWebElement _close;
        [FindsBy(How = How.Id, Using = DELETE)]
        private IWebElement _delete;
        [FindsBy(How = How.Id, Using = CONFIRMDELETE)]
        private IWebElement _confirmDelete;
        [FindsBy(How = How.Id, Using = FROM)]
        private IWebElement _from;

        [FindsBy(How = How.Id, Using = EDIT_SERVICE_BTN)]
        private IWebElement _editServiceBtn;


        [FindsBy(How = How.Id, Using = EDIT_ITEM_BTN)]
        private IWebElement _editItemBtn;
        protected IWebDriver WebDriver
        {
            get => WebDriverFactory.Driver;
        }
        public ProductPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }
        public void ResetFilter()
        {
            _resetFilterDev = WaitForElementIsVisible(By.Id(RESET_FILTER_DEV));
            _resetFilterDev.Click();

            WaitForLoad();

            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                // pas de date
            }
        }
        public void Search(string value, bool ignoreAutoSuggest = false)
        {
            _searchFilter = WaitForElementIsVisible(By.Id(SEARCH_FILTER));
            _searchFilter.SetValue(ControlType.TextBox, value);
            if (ignoreAutoSuggest == false && isElementVisible(By.XPath(String.Format(FIRST_RESULT_SEARCH, value))))
            {
                _searchFilter.SendKeys(Keys.Tab);
            }
            WaitPageLoading();
            WaitForLoad();
            WaitPageLoading();
        }

        public void ClickNewProduct()
        {
            _plusButton = WaitForElementIsVisible(By.XPath(PLUS_BUTTON));
            _plusButton.Click();
            WaitForLoad();

            if(IsDev())
            {
                _newProduct = WaitForElementIsVisible(By.Id(NEW_PRODUCT));
                _newProduct.Click();
            }
            else
            {
                var newProduct = WaitForElementIsVisible(By.XPath("//*[@id=\"tabContentProductContainer\"]/div/div/div[2]/div/a"));
                newProduct.Click();
            }
            WaitForLoad();
        }
        public void ClickAddNewProduct()
        {
            _plusButton = WaitForElementIsVisible(By.XPath(PLUS_BUTTON));
            _plusButton.Click();
            _addNewProductButton = WaitForElementIsVisible(By.Id(ADD_NEW_PRODUCT));
            _addNewProductButton.Click();
        }

        public void CreateNewProduct(string item, string service)
        {
            WaitForLoad();

            _itemInput = WaitForElementIsVisible(By.Id(ITEM_INPUT));
            _itemInput.SendKeys(item);
            // //*[@id="create-product-form"]/div[2]/div/div[1]/div/div[2]/div[2]/div/span[1]/div/div/div[2]
            var firstItem = WaitForElementIsVisible(By.XPath($"//*[@id=\"create-product-form\"]/div[2]/div/div[1]/div/div[2]/div[2]/div/span[1]/div/div/div"));
            firstItem.Click();
            WaitForLoad(); 
            _serviceInput = WaitForElementIsVisible(By.XPath(SERVICE_INPUT));
            _serviceInput.SendKeys(service);
            WaitForLoad();

            var firstService = WaitForElementIsVisible(By.XPath("//*[@id=\"create-product-form\"]/div[2]/div/div[2]/div/div[2]/div[2]/div/span[1]/div/div/div"));
           
           firstService.Click();
            WaitForLoad();
        }

        public void ClickAddNew()
        {
            _add_new = WaitForElementIsVisible(By.Id(ADDNEW));
            _add_new.Click();
        }
        public void ClickClose()
        {
            _close = WaitForElementIsVisible(By.XPath(CLOSE));
            _close.Click();
        }
        public void ClickAdd()
        {
            _add = WaitForElementIsVisible(By.Id(ADD));
            _add.Click();
            WaitLoading();
            WaitPageLoading(); 
        }        
        public void DeleteFirstProduct()
        {
            if (IsDev())
            {
                _delete = WaitForElementIsVisible(By.Id("purchasing-product-delete1"));
            }
            else
            {
                _delete = WaitForElementIsVisible(By.Id(DELETE));
            }
            _delete.Click();
            WaitForLoad();
            _confirmDelete = WaitForElementIsVisible(By.Id(CONFIRMDELETE));
            _confirmDelete.Click();
            WaitForLoad();
        }
        
        public bool IsNewProductFormClosed()
        {
            return !isElementVisible(By.XPath(FROM));
        }

        public string GetFirstItemProduct()
        {
            WaitPageLoading();
            var firstItem = WaitForElementIsVisible(By.XPath("//*[@id=\"tabContentProductContainer\"]/table/tbody/tr/td[2]"));
            return firstItem.Text;
        }
        public string GetFirstServiceProduct()
        {
            var firstService = WaitForElementIsVisible(By.XPath("//*[@id=\"tabContentProductContainer\"]/table/tbody/tr/td[4]"));
            return firstService.Text;
        }

        public void ClickDeleteFirstProduct()
        {
            WaitPageLoading();  
            _delete = WaitForElementIsVisible(By.Id(DELETE));
            _delete.Click();
            WaitPageLoading();
        }

        public void CancelDeleteFirstProduct()
        {
            WaitLoading();
            var _cancelDelete = WaitForElementIsVisible(By.XPath("//*[@id=\"dataConfirmCancel\"]"));
            _cancelDelete.Click();
            WaitLoading();
        }
        public ServicePricePage OpenEditService()
        {
            _editServiceBtn = WaitForElementIsVisible(By.Id(EDIT_SERVICE_BTN));
            _editServiceBtn.Click();

            return new ServicePricePage(_webDriver, _testContext);
        }

        public ItemGeneralInformationPage OpenEditItem()
        {
            WaitLoading();
            _editItemBtn = WaitForElementIsVisible(By.Id(EDIT_ITEM_BTN));
            _editItemBtn.Click();
            WaitLoading();
            return new ItemGeneralInformationPage(_webDriver, _testContext);
        }
        public bool HasLink()
        {
            _editServiceBtn = WaitForElementIsVisible(By.Id(EDIT_SERVICE_BTN));
            var linkServiceButton = _editServiceBtn.GetAttribute("href");
            _editItemBtn = WaitForElementIsVisible(By.Id(EDIT_ITEM_BTN));
            var linkItemButton = _editItemBtn.GetAttribute("href");

            return linkItemButton.Contains("/Purchasing/Item/Detail?id=") && linkServiceButton.Contains("/Customers/Service/Detail?id=");
        }
        public int CheckTotalNumberProduct()
        {
            WaitLoading() ;
            WaitPageLoading(); 
            var _totalNumber = WaitForElementExists(By.XPath("//*[@id=\"tabContentProductContainer\"]/div/h1/span"));
          int nombre = Int32.Parse(_totalNumber.Text);
            return nombre; 
        }
    }
}
