using DocumentFormat.OpenXml.VariantTypes;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System.Security.Policy;
using System.Threading;
using System;
using System.Data;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using DocumentFormat.OpenXml.Bibliography;
using iText.Layout;
using System.Text.RegularExpressions;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.Catalogs
{
    public class CatalogPage : PageBase
    {
        public CatalogPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {

        }
        //_______________________________________________________Constantes____________________________________________________________

        private const string SEARCH_FILTER = "tbSearchPattern";
        private const string RESET_FILTER_DEV = "ResetFilter";
        private const string PRODUCT_TAB = "PRODUCTS";
        private const string PLUS_BTN = "/html/body/div[2]/div/div[2]/ul/li[4]/div/button";
        private const string NEW_PRODUCT = "ProductCreateBtn";
        private const string PROD_NAME = "Name";
        private const string SELECT_ITEM = "select-item";
        private const string SAVE_PROD = "/html/body/div[3]/div/div/div/div[3]/button[2]";
        private const string SELECT_CUSTOMER = "listbox-customers";
        private const string FIRST_ITEM = "//*[@id=\"list-item-with-action\"]/div[2]/div/div/div[2]";
        private const string ITEM = "//*[@id=\"create-product-form\"]/div[2]/div/div[5]/div[2]/div/div[1]/div/div[1]/span/div/div/div[1]";
        private const string PROD_ITEM_NAME = "/html/body/div[3]/div/div[2]/div/div/div/form/div/div/div/div/div[1]/div/input";
        private const string DELETE_BTN = "/ html/body/div[2]/div/div[2]/div/div/div[2]/div/div/div[3]/a/span";
        private const string CONFIRM_BTN = "dataConfirmOK";
        private const string PLUS_CATALOG = "//*[@id=\"paramPurchasingTab\"]/li[4]/div/button";
        private const string NEW_CATALOG = "CatalogCreateBtn";
        private const string SAVE_CATALOG = "btn-submit-new-catalog";
        private const string START_DATE = "datepicker-catalog-start";
        private const string END_DATE = "datepicker-catalog-end";
        private const string FIRST_CATALOG = "//*[@id=\"customers-catalog-catalogupdate-1 \"]/div[2]";
        private const string CATEGORIE_TAB = "CategoryReadAllFromCatalog";
        private const string NEW_CATEGORIE = "//*[@id=\"xxx\"]/div/div[1]/div[2]/div/a[2]";
        private const string CATEGORIE_NAME = "CategoryName";
        private const string SELECT_CATALOG_TYPE = "dropdown-category-catalog-pos-type";
        private const string SAVE_CATEGORY = "//*[@id=\"modal-1\"]/div/div/form/div[3]/button[2]";
        private const string DETAIL_TAB = "//*[@id=\"nav-tab\"]/li[3]/a";
        private const string ADD_ITEM = "//*[@id=\"btn-open-add-item-modal\"]/div/a[2]";
        private const string SELECT_DETAIL_TYPE = "dropdown-detail-types";
        private const string SELECT_PRODUCT = "dropdown-detail-products";
        private const string SELECT_TAX = "dropdown-detail-taxtypes";
        private const string SELECT_COLOR = "dropdown-airfi-button-color";
        private const string SAVE_CATEG_PROD = "add-new-catalog-item-submit-btn";
        private const string CLICK_CATEGORY = "/html/body/div[3]/div/div[2]/div/div/div[2]/div[1]/nav/div/a";
        private const string ITEMS = "/html/body/div[3]/div/div[2]/div/div/div[2]/div[2]/div/div/table/tbody/tr[*]";
        public const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";
        public const string CLOSE = "//*[@id=\"create-catalog-category\"]/div[3]/button[1]";
        public const string DELETE_ITEM = "/html/body/div[3]/div/div[2]/div/div/div[2]/div[2]/div/div/table/tbody/tr/td[9]/a[1]";
        public const string CATALOG = "//*[@id=\"customers-catalog-catalogupdate-1 \"]/div[2]";
        public const string DELETE_CATALOG = "//*[@id=\"customers-catalog-catalogdelete-1\"]/span";
        public const string FILTER_PROD = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[4]/input";
        //_______________________________________________________Variables_____________________________________________________________

        [FindsBy(How = How.Id, Using = SEARCH_FILTER)]
        private IWebElement _searchFilter;

        [FindsBy(How = How.Id, Using = RESET_FILTER_DEV)]
        private IWebElement _resetFilterDev;

        [FindsBy(How = How.Id, Using = PRODUCT_TAB)]
        private IWebElement _productTab;

        [FindsBy(How = How.Id, Using = NEW_PRODUCT)]
        private IWebElement _newProduct;

        [FindsBy(How = How.XPath, Using = PLUS_BTN)]
        private IWebElement _plusButton;

        [FindsBy(How = How.Id, Using = PROD_NAME)]
        private IWebElement _ProductName;

        [FindsBy(How = How.Id, Using = SELECT_ITEM)]
        private IWebElement _selectItem;

        [FindsBy(How = How.XPath, Using = SAVE_PROD)]
        private IWebElement _save;

        [FindsBy(How = How.Id, Using = SELECT_CUSTOMER)]
        private IWebElement _selectCustomer;

        [FindsBy(How = How.XPath, Using = ITEM)]
        private IWebElement _Item;

        [FindsBy(How = How.XPath, Using = FIRST_ITEM)]
        private IWebElement _firstItem;

        [FindsBy(How = How.XPath, Using = PROD_ITEM_NAME)]
        private IWebElement _ProdItemName;

        [FindsBy(How = How.XPath, Using = PLUS_CATALOG)]
        private IWebElement _plusCatalog;

        [FindsBy(How = How.Id, Using = NEW_CATALOG)]
        private IWebElement _newCatalog;

        [FindsBy(How = How.Id, Using = SAVE_CATALOG)]
        private IWebElement _saveCatalog;

        [FindsBy(How = How.Id, Using = START_DATE)]
        private IWebElement _startDate;

        [FindsBy(How = How.Id, Using = END_DATE)]
        private IWebElement _endDate;

        [FindsBy(How = How.XPath, Using = FIRST_CATALOG)]
        private IWebElement _firstCatalog;

        [FindsBy(How = How.Id, Using = CATEGORIE_TAB)]
        private IWebElement _categorieTab;

        [FindsBy(How = How.XPath, Using = NEW_CATEGORIE)]
        private IWebElement _newCategorie;

        [FindsBy(How = How.Id, Using = CATEGORIE_NAME)]
        private IWebElement _categorieName;

        [FindsBy(How = How.Id, Using = SELECT_CATALOG_TYPE)]
        private IWebElement _selectCatalogType;

        [FindsBy(How = How.XPath, Using = SAVE_CATEGORY)]
        private IWebElement _saveCategorie;

        [FindsBy(How = How.XPath, Using = DETAIL_TAB)]
        private IWebElement _detailTab;

        [FindsBy(How = How.XPath, Using = ADD_ITEM)]
        private IWebElement _addItem;

        [FindsBy(How = How.Id, Using = SELECT_DETAIL_TYPE)]
        private IWebElement _selectdetailType;

        [FindsBy(How = How.Id, Using = SELECT_PRODUCT)]
        private IWebElement _selectProd;

        [FindsBy(How = How.Id, Using = SELECT_TAX)]
        private IWebElement _selectTax;

        [FindsBy(How = How.Id, Using = SELECT_COLOR)]
        private IWebElement _selectColor;

        [FindsBy(How = How.Id, Using = SAVE_CATEG_PROD)]
        private IWebElement _saveCategProd;

        [FindsBy(How = How.XPath, Using = CLICK_CATEGORY)]
        private IWebElement _clickCategory;

        [FindsBy(How = How.XPath, Using = ITEMS)]
        private IWebElement _items;

        [FindsBy(How = How.XPath, Using = CLOSE)]
        private IWebElement _close;

        [FindsBy(How = How.XPath, Using = FILTER_PROD)]
        private IWebElement _filterProd;

        //_______________________________________ Méthodes_____________________________________________________________________________


        public enum FilterType
        {
            Search,
            Customer

        }
        public void Filter(FilterType filterType, object value)
        {
            switch (filterType)
            {
                case FilterType.Search:
                    WaitPageLoading();
                    _searchFilter = WaitForElementIsVisible(By.Id(SEARCH_FILTER));
                    _searchFilter.SetValue(ControlType.TextBox, value);
                    WaitPageLoading();
                    WaitForLoad();
                    break;

            }
            WaitForLoad();
        }

        public void FilterProd(FilterType filterType, object value)
        {
            switch (filterType)
            {
                case FilterType.Search:
                    WaitPageLoading();
                    _filterProd = WaitForElementIsVisible(By.XPath(FILTER_PROD));
                    _filterProd.SetValue(ControlType.TextBox, value);
                    WaitPageLoading();
                    WaitForLoad();
                    break;

            }
            WaitForLoad();
        }
        public void ResetFilters()
        {
            WaitPageLoading();
            if (isElementVisible(By.Id(RESET_FILTER_DEV)))
            {
                _resetFilterDev = WaitForElementIsVisible(By.Id(RESET_FILTER_DEV));
                _resetFilterDev.Click();
            }

            WaitForLoad();
            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                // pas de date
            }
        }

        public CatalogPage GoToProductPage()
        {
            _productTab = WaitForElementToBeClickable(By.Id(PRODUCT_TAB));
            _productTab.Click();

            WaitPageLoading();
            WaitForLoad();
            return new CatalogPage(_webDriver, _testContext);
        }

        public CatalogPage ClickProductCreatePage()
        {
            WaitForLoad();
            _plusButton = WaitForElementIsVisible(By.XPath(PLUS_BTN));
            var actions = new Actions(_webDriver);
            actions.MoveToElement(_plusButton).Perform();
            WaitForLoad();

            _newProduct = WaitForElementIsVisible(By.Id(NEW_PRODUCT));
            _newProduct.Click();
            WaitForLoad();

            return new CatalogPage(_webDriver, _testContext);
        }

        public CatalogPage FillField_CreatNewProduct(string ProdName)
        {
            if (ProdName != null)
            {
                _ProductName = WaitForElementIsVisible(By.Id(PROD_NAME));
                _ProductName.SetValue(ControlType.TextBox, ProdName);
                WaitForLoad();
            }
            var s = WaitForElementIsVisible(By.Id(SELECT_CUSTOMER));
            s.SetValue(ControlType.TextBox, "SMARTWINGS, A.S. (TVS)");
            WaitForLoad();

            if (isElementExists(By.Id(SELECT_CUSTOMER)))
            {
                var selectFirstItemC = WaitForElementIsVisible(By.Id(SELECT_CUSTOMER));
                selectFirstItemC.Click();
                WaitForLoad();
            }

            var f = WaitForElementIsVisible(By.Id(SELECT_ITEM));
            f.SetValue(ControlType.TextBox, "A");
            WaitPageLoading();

            if (isElementExists(By.Id(SELECT_ITEM)))
            {
                var selectFirstItem = WaitForElementIsVisible(By.XPath(ITEM));
                selectFirstItem.Click();
                WaitForLoad();
            }


            _save = WaitForElementIsVisible(By.Id("btn-submit-new-product"));
            _save.Click();

            return new CatalogPage(_webDriver, _testContext);
        }

        public void SelectFirstItemProduct()
        {
            _firstItem = WaitForElementIsVisible(By.XPath(FIRST_ITEM));
            _firstItem.Click();
            WaitForLoad();
        }

        public bool GeneralInfoNotEmpty(string ProdName)
        {
            _ProdItemName = WaitForElementIsVisible(By.Id("Name"));
            var value = _ProdItemName.GetAttribute("value");
            WaitForLoad();
            if (String.IsNullOrEmpty(value) || !string.Equals(value, ProdName))
            {

                return false;


            }

            else return true;
        }

        public void DeleteFirstProd()
        {

            Actions actions = new Actions(_webDriver);
            IWebElement rowElement = WaitForElementIsVisible(By.XPath(FIRST_ITEM));
            actions.MoveToElement(rowElement).Perform();
            if (isElementExists(By.XPath(DELETE_BTN)))
            {
                var DeleteBtn = WaitForElementIsVisible(By.XPath(DELETE_BTN));
                DeleteBtn.Click();

                var confirmBtn = WaitForElementIsVisible(By.Id(CONFIRM_BTN));
                confirmBtn.Click();

                WaitForLoad();
            }
        }
        public CatalogPage ClickCatalogCreatePage()
        {
            WaitForLoad();
            _plusCatalog = WaitForElementIsVisible(By.XPath(PLUS_CATALOG));
            var actions = new Actions(_webDriver);
            actions.MoveToElement(_plusCatalog).Perform();
            WaitForLoad();

            _newCatalog = WaitForElementIsVisible(By.Id(NEW_CATALOG));
            _newCatalog.Click();
            WaitForLoad();

            return new CatalogPage(_webDriver, _testContext);
        }
        public CatalogPage FillField_CreatNewCatalog(string CatalogName, DateTime DateFrom, DateTime DateTo)
        {
            if (CatalogName != null)
            {
                _ProductName = WaitForElementIsVisible(By.Id(PROD_NAME));
                _ProductName.SetValue(ControlType.TextBox, CatalogName);
                WaitForLoad();
            }
            var s = WaitForElementIsVisible(By.Id(SELECT_CUSTOMER));
            s.SetValue(ControlType.TextBox, "SMARTWINGS, A.S. (TVS)");
            WaitForLoad();

            if (isElementExists(By.Id(SELECT_CUSTOMER)))
            {
                var selectFirstItemC = WaitForElementIsVisible(By.Id(SELECT_CUSTOMER));
                selectFirstItemC.Click();
                WaitForLoad();
            }

            // Définition de Start Date
            _startDate = WaitForElementIsVisible(By.Id(START_DATE));
            _startDate.SetValue(ControlType.DateTime, DateFrom);
            _startDate.SendKeys(Keys.Tab);

            // Définition de End Date
            _endDate = WaitForElementIsVisible(By.Id(END_DATE));
            _endDate.SetValue(ControlType.DateTime, DateTo);
            _endDate.SendKeys(Keys.Tab);

            _saveCatalog = WaitForElementIsVisible(By.Id(SAVE_CATALOG));
            _saveCatalog.Click();

            return new CatalogPage(_webDriver, _testContext);
        }
        public CatalogPage SelectFirstItemCatalog()
        {
            _firstCatalog = WaitForElementIsVisible(By.XPath(FIRST_CATALOG));
            _firstCatalog.Click();
            WaitForLoad();


            return new CatalogPage(_webDriver, _testContext);
        }
        public CatalogPage GoToCategoriePage()
        {
            _categorieTab = WaitForElementToBeClickable(By.Id(CATEGORIE_TAB));
            _categorieTab.Click();

            WaitPageLoading();
            WaitForLoad();
            return new CatalogPage(_webDriver, _testContext);
        }
        public CatalogPage FillField_CreatNewCategorie(string CategorieName)
        {
            _newCategorie = WaitForElementIsVisible(By.XPath(NEW_CATEGORIE));
            _newCategorie.Click();
            WaitForLoad();

            if (CategorieName != null)
            {
                _categorieName = WaitForElementIsVisible(By.Id(CATEGORIE_NAME));
                _categorieName.SetValue(ControlType.TextBox, CategorieName);
                WaitForLoad();
            }
            var s = WaitForElementIsVisible(By.Id(SELECT_CATALOG_TYPE));
            s.SetValue(ControlType.TextBox, "Catering");
            WaitForLoad();

            if (isElementExists(By.Id(SELECT_CATALOG_TYPE)))
            {
                var selectFirstItemC = WaitForElementIsVisible(By.Id(SELECT_CATALOG_TYPE));
                selectFirstItemC.Click();
                WaitForLoad();
            }



            _saveCategorie = WaitForElementIsVisible(By.XPath(SAVE_CATEGORY));
            _saveCategorie.Click();

            return new CatalogPage(_webDriver, _testContext);
        }

        public CatalogPage GoToDetailPage()
        {
            _detailTab = WaitForElementToBeClickable(By.XPath(DETAIL_TAB));
            _detailTab.Click();

            WaitPageLoading();
            WaitForLoad();
            return new CatalogPage(_webDriver, _testContext);
        }

        public CatalogPage FillField_CreatNewCatalogProd(string ProdName)
        {
            _addItem = WaitForElementIsVisible(By.XPath(ADD_ITEM));
            _addItem.Click();
            WaitForLoad();

            var s = WaitForElementIsVisible(By.Id(SELECT_DETAIL_TYPE));
            s.SetValue(ControlType.TextBox, "Product");
            WaitForLoad();

            if (isElementExists(By.Id(SELECT_DETAIL_TYPE)))
            {
                var selectFirstItemC = WaitForElementIsVisible(By.Id(SELECT_DETAIL_TYPE));
                selectFirstItemC.Click();
                WaitForLoad();
            }

            var d = WaitForElementIsVisible(By.Id(SELECT_PRODUCT));
            d.SetValue(ControlType.TextBox, ProdName);
            WaitForLoad();

            if (isElementExists(By.Id(SELECT_PRODUCT)))
            {
                var selectFirstItemprod = WaitForElementIsVisible(By.Id(SELECT_PRODUCT));
                selectFirstItemprod.Click();
                WaitForLoad();
            }

            var t = WaitForElementIsVisible(By.Id(SELECT_TAX));
            t.SetValue(ControlType.TextBox, "6-EXO BOB");
            WaitForLoad();

            if (isElementExists(By.Id(SELECT_TAX)))
            {
                var selectFirstItemt = WaitForElementIsVisible(By.Id(SELECT_TAX));
                selectFirstItemt.Click();
                WaitForLoad();
            }

            var co = WaitForElementIsVisible(By.Id(SELECT_COLOR));
            co.SetValue(ControlType.TextBox, "Green");
            WaitForLoad();

            if (isElementExists(By.Id(SELECT_COLOR)))
            {
                var selectFirstItemt = WaitForElementIsVisible(By.Id(SELECT_COLOR));
                selectFirstItemt.Click();
                WaitForLoad();
            }

            _saveCategProd = WaitForElementIsVisible(By.Id(SAVE_CATEG_PROD));
            _saveCategProd.Click();


            _close = WaitForElementIsVisible(By.XPath(CLOSE));
            _close.Click();

            return new CatalogPage(_webDriver, _testContext);
        }
        public bool ItemIsVerifiedShow()
        {
            _addItem = WaitForElementIsVisible(By.XPath(ADD_ITEM));

            if (_addItem != null)
            {
                return true;
            }

            else return false;


        }
        public bool ItemIsVerifiedShowWithProd()
        {

            _clickCategory = WaitForElementIsVisible(By.XPath(CLICK_CATEGORY));
            _clickCategory.Click();
            WaitForLoad();

            var _nb = _webDriver.FindElements(By.XPath(ITEMS));
            if (_nb.Count != 0)
            {
                return true;
            }
            else return false;
        }
        public void BackToList()
        {
            var backToListBtn = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            backToListBtn.Click();
        }
        public void DeleteFirstCatalog()
        {
            Actions actions = new Actions(_webDriver);
            IWebElement rowElement = WaitForElementIsVisible(By.XPath(CATALOG));
            actions.MoveToElement(rowElement).Perform();

            if (isElementExists(By.XPath(DELETE_CATALOG)))
            {
                var DeleteBtn = WaitForElementIsVisible(By.XPath(DELETE_CATALOG));
                DeleteBtn.Click();

                var confirmBtn = WaitForElementIsVisible(By.Id(CONFIRM_BTN));
                confirmBtn.Click();

                WaitForLoad();
            }
        }

        public void deleteItem()
        {
            _clickCategory = WaitForElementIsVisible(By.XPath(CLICK_CATEGORY));
            _clickCategory.Click();
            WaitForLoad();

            Actions actions = new Actions(_webDriver);
            IWebElement rowElement = WaitForElementIsVisible(By.XPath(ITEMS));
            actions.MoveToElement(rowElement).Perform();

            var _nb = _webDriver.FindElements(By.XPath(ITEMS));

            if (_nb.Count > 0)
            {
                foreach (var line in _nb)
                {
                    if (isElementExists(By.XPath(DELETE_ITEM)))
                    {
                        var DeleteBtn = WaitForElementIsVisible(By.XPath(DELETE_ITEM));
                        DeleteBtn.Click();

                        var confirmBtn = WaitForElementIsVisible(By.Id(CONFIRM_BTN));
                        confirmBtn.Click();

                        WaitForLoad();
                    }
                }
            }
        }

        public CatalogPage FillField_CreatNewProduct_SpecialCaractere(string ProdName, string SpecialCaractere)
        {
            if (ProdName != null)
            {
                _ProductName = WaitForElementIsVisible(By.Id(PROD_NAME));
                _ProductName.SetValue(ControlType.TextBox, ProdName);
                WaitForLoad();
            }
            var s = WaitForElementIsVisible(By.Id(SELECT_CUSTOMER));
            s.SetValue(ControlType.TextBox, "SMARTWINGS, A.S. (TVS)");
            WaitForLoad();

            if (isElementExists(By.Id(SELECT_CUSTOMER)))
            {
                var selectFirstItemC = WaitForElementIsVisible(By.Id(SELECT_CUSTOMER));
                selectFirstItemC.Click();
                WaitForLoad();
            }

            var f = WaitForElementIsVisible(By.Id(SELECT_ITEM));
            f.SetValue(ControlType.TextBox, SpecialCaractere);
            WaitForLoad();
            f.SendKeys(Keys.Enter);

            if (isElementExists(By.Id(SELECT_ITEM)))
            {
                var selectFirstItem = WaitForElementIsVisible(By.XPath(ITEM));
                selectFirstItem.Click();
                WaitForLoad();
            }


            _save = WaitForElementIsVisible(By.Id("btn-submit-new-product"));
            _save.Click();

            return new CatalogPage(_webDriver, _testContext);
        }
        public bool SearchProductByName(string SpecialCaractere)
        {
            if (string.IsNullOrEmpty(SpecialCaractere))
            {
                throw new ArgumentException("Item name cannot be null or empty.");
            }
            var productSearchField = WaitForElementIsVisible(By.Id(SELECT_ITEM));
            productSearchField.SetValue(ControlType.TextBox, SpecialCaractere);
            WaitForLoad();
            productSearchField.SendKeys(Keys.Enter);  // Trigger search by pressing Enter
            WaitForLoad();
            if (isElementExists(By.Id(SELECT_ITEM)))
            {
                var selectFirstItem = WaitForElementIsVisible(By.XPath(ITEM)); // Adjust ITEM to the correct XPath for the first item
                string itemText = selectFirstItem.Text;
                string pattern = @"[^a-zA-Z0-9 ]";
                bool containsSpecialCharacter = Regex.IsMatch(itemText, pattern);

                return containsSpecialCharacter;
            }

            return false;
        }
        public bool IsIngredientPresent(string ingredient)
        {
            return isElementExists(By.XPath($"//*[@id='list-item-with-action']/div[2]/div/div/div[2]/table/tbody/tr/td[3][contains(text(), '{ingredient}')]"));
        }

    }
}