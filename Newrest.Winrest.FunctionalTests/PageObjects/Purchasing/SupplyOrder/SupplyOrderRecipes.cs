using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Recipes;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
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

namespace Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.SupplyOrder
{
    public class SupplyOrderRecipes : PageBase
    {
        public SupplyOrderRecipes(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // ____________________________________________ Constantes ________________________________________________

        private const string EXTENDED_BTN = "//*[@id=\"div-body\"]/div/div[1]/div/div[1]/button";
        private const string REFRESH = "btn-supply-orders-refresh";

        private const string ADD_RECIPE = "IngredientName";

        private const string FIRST_RECIPE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div";
        private const string FIRST_RECIPE_QTY = "item_RelatedSupplyOrderRecipeVM_SupplyOrderRecipeDetail_NeededQuantity";
        private const string FIRST_RECIPE_PRICE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div/div/div/form/div[1]/div[8]/span";
        private const string EDIT_FIRST_RECIPE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div/div/div/form/div[1]/div[9]/div/a[1]/span";
        private const string DELETE_FIRST_RECIPE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div/div/div/form/div[1]/div[9]/div/a[2]/span";

        // Filtres
        private const string RESET_FILTER_DEV = "ResetFilter";
        private const string RESET_FILTER_PATCH = "//*[@id=\"formSearchRecipes\"]/div[1]/a";

        private const string SEARCH_FILTER = "SearchPatternWithAutocomplete";
        private const string FIRST_RESULT_SEARCH = "//*[@id=\"formSearchRecipes\"]/div[2]/span/div/div/div/strong[text()='{0}']";
        private const string SORTBY_FILTER = "cbSortBy";

        private const string ITEMS_LIST = "editable-row";
        private const string RECIPE_LINE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[{0}]";
        private const string RECIPE_NAME = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[{0}]/div/div/form/div[1]/div[2]/span";
        private const string RECIPE_NAME2 = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[2]/span";
        private const string RECIPE_PORTION = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[{0}]/div/div/form/div[1]/div[5]/span";
        private const string ITEMS = "hrefTabContentItems";

        // ____________________________________________ Variables _________________________________________________

        [FindsBy(How = How.XPath, Using = EXTENDED_BTN)]
        private IWebElement _extendedBtn;

        [FindsBy(How = How.Id, Using = REFRESH)]
        private IWebElement _refresh;

        [FindsBy(How = How.Id, Using = ADD_RECIPE)]
        private IWebElement _addRecipe;

        // Tableau recipes
        [FindsBy(How = How.XPath, Using = FIRST_RECIPE)]
        private IWebElement _firstRecipe;

        [FindsBy(How = How.XPath, Using = FIRST_RECIPE_QTY)]
        private IWebElement _firstRecipeQty;

        [FindsBy(How = How.XPath, Using = FIRST_RECIPE_PRICE)]
        private IWebElement _firstRecipePrice;

        [FindsBy(How = How.XPath, Using = EDIT_FIRST_RECIPE)]
        private IWebElement _editFirstRecipe;

        [FindsBy(How = How.XPath, Using = DELETE_FIRST_RECIPE)]
        private IWebElement _deleteFirstRecipe;

        // Filtres
        [FindsBy(How = How.Id, Using = RESET_FILTER_DEV)]
        private IWebElement _resetFilterDev;
        
        [FindsBy(How = How.XPath, Using = RESET_FILTER_PATCH)]
        private IWebElement _resetFilterPatch;

        [FindsBy(How = How.Id, Using = SEARCH_FILTER)]
        private IWebElement _searchFilter;

        [FindsBy(How = How.Id, Using = SORTBY_FILTER)]
        private IWebElement _sortByFilter;

        public enum FilterSupplyRecipeType
        {
            SearchByName,
            SortBy
        }

        public void Filter(FilterSupplyRecipeType FilterSupplyItemType, object value)
        {

            switch (FilterSupplyItemType)
            {
                case FilterSupplyRecipeType.SearchByName:
                    _searchFilter = WaitForElementIsVisible(By.Id(SEARCH_FILTER));
                    _searchFilter.SetValue(ControlType.TextBox, value);

                    try
                    {
                        var firstResultSearch = _webDriver.FindElement(By.XPath(String.Format(FIRST_RESULT_SEARCH, value)));
                        firstResultSearch.Click();
                    }
                    catch
                    {
                     //item non trouvé
                    }
                break;

                case FilterSupplyRecipeType.SortBy:
                    _sortByFilter = WaitForElementIsVisible(By.Id(SORTBY_FILTER));
                    _sortByFilter.Click();
                    var element = WaitForElementIsVisible(By.XPath("//option[contains(@value,'" + value + "')]"));
                    _sortByFilter.SetValue(ControlType.DropDownList, element.Text);
                    _sortByFilter.Click();
                    break;
            }

            WaitPageLoading();
            Thread.Sleep(1500);
        }

        public void ResetFilter()
        {
            try
            {
                _resetFilterDev = WaitForElementIsVisible(By.Id(RESET_FILTER_DEV));
                _resetFilterDev.Click();
            }
            catch
            {
                _resetFilterPatch = WaitForElementIsVisible(By.XPath(RESET_FILTER_PATCH));
                _resetFilterPatch.Click();
            }
            
            WaitForLoad();
            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                // pas de date
            }
        }

        public bool VerifyName(string name)
        {
            bool valueBool = true;

            var tot = _webDriver.FindElements(By.ClassName(ITEMS_LIST)).Count;
            if (tot > 100)
            {
                tot = 100;
            }
            if (tot == 0)
                return false;

            for (int i = 1; i <= tot; i++)
            {
                try
                {
                    IWebElement line = _webDriver.FindElement(By.XPath(string.Format(RECIPE_LINE,i)));
                    line.Click();

                    var element = WaitForElementIsVisible(By.XPath(string.Format(RECIPE_NAME, i)));

                    if (!element.Text.Contains(name))
                    { valueBool = false; }

                }
                catch
                {
                    valueBool = false;
                }
            }
            return valueBool;
        }



        public bool IsSortedByName()
        {
            bool valueBool = true;
            var ancientName = "";

            var tot = _webDriver.FindElements(By.ClassName(ITEMS_LIST)).Count;
            if (tot > 100)
            {
                tot = 100;
            }
            if (tot == 0)
                return false;

            for (int i = 1; i <= tot; i++)
            {
                try
                {
                    IWebElement line = _webDriver.FindElement(By.XPath(string.Format(RECIPE_LINE, i)));
                    line.Click();

                    IWebElement Name = _webDriver.FindElement(By.XPath(string.Format(RECIPE_NAME, i)));
                    string name = Name.Text;

                    if (string.Compare(ancientName, name) > 0)
                    { valueBool = false; }

                    ancientName = name;
                }
                catch
                {
                    valueBool = false;
                }
            }
            return valueBool;
        }

        public bool IsSortedByPortion()
        {
            bool valueBool = true;
            var ancientName = "";

            var tot = _webDriver.FindElements(By.ClassName(ITEMS_LIST)).Count;
            if (tot > 100)
            {
                tot = 100;
            }
            if (tot == 0)
                return false;

            for (int i = 1; i <= tot; i++)
            {
                try
                {
                    IWebElement line = _webDriver.FindElement(By.XPath(string.Format(RECIPE_LINE, i)));
                    line.Click();

                    IWebElement element = _webDriver.FindElement(By.XPath(string.Format(RECIPE_PORTION, i)));
                    string name = element.Text;

                    if (string.Compare(ancientName, name) > 0)
                    { valueBool = false; }

                    ancientName = name;
                }
                catch
                {
                    valueBool = false;
                }
            }
            return valueBool;
        }

        // ____________________________________________ Méthodes __________________________________________________

        public override void ShowExtendedMenu()
        {
            WaitForElementExists(By.XPath(EXTENDED_BTN));
            var actions = new Actions(_webDriver);
            actions.MoveToElement(_extendedBtn).Perform();
            WaitForLoad();
        }

        public void Refresh()
        {
            ShowExtendedMenu();

            _refresh = WaitForElementIsVisible(By.Id(REFRESH));
            _refresh.Click();
            WaitForLoad();
        }

        public void AddRecipe(string recipeName)
        {
            _addRecipe = WaitForElementIsVisible(By.Id(ADD_RECIPE));
            _addRecipe.SetValue(ControlType.TextBox, recipeName);
            WaitPageLoading();

            var selectRecipe = WaitForElementIsVisible(By.XPath("//span[text()='" + recipeName + "']"));
            selectRecipe.Click();
            WaitForLoad();
        }

        public string GetFirstNeededQty()
        {
            _firstRecipe = WaitForElementIsVisible(By.XPath(FIRST_RECIPE));
            _firstRecipe.Click();

            _firstRecipeQty = WaitForElementIsVisible(By.Id(FIRST_RECIPE_QTY));
            return _firstRecipeQty.GetAttribute("value");
        }

        public void SetFirstRecipeNeededQty(string quantity)
        {
            _firstRecipeQty = WaitForElementIsVisible(By.Id(FIRST_RECIPE_QTY));
            _firstRecipeQty.SetValue(ControlType.TextBox, quantity);

            // Temps de prise en compte de la valeur
            Thread.Sleep(2000);
        }

        public string GetFirstPrice()
        {
            _firstRecipe = WaitForElementIsVisible(By.XPath(FIRST_RECIPE));
            _firstRecipe.Click();

            _firstRecipePrice = WaitForElementIsVisible(By.XPath(FIRST_RECIPE_PRICE));
            return _firstRecipePrice.Text;
        }

        public bool IsFirstRecipeAdded()
        {
            try
            {
                _webDriver.FindElement(By.XPath(FIRST_RECIPE));
                return true;
            }
            catch
            {
                return false;
            }
        }

        public void DeleteFirstRecipe()
        {
            _firstRecipe = WaitForElementIsVisible(By.XPath(FIRST_RECIPE));
            _firstRecipe.Click();

            _deleteFirstRecipe = WaitForElementIsVisible(By.XPath(DELETE_FIRST_RECIPE));
            _deleteFirstRecipe.Click();
            WaitForLoad();
        }

        public RecipesVariantPage EditFirstRecipe()
        {
            _firstRecipe = WaitForElementIsVisible(By.XPath(FIRST_RECIPE));
            _firstRecipe.Click();

            _editFirstRecipe = WaitForElementIsVisible(By.XPath(EDIT_FIRST_RECIPE));
            _editFirstRecipe.Click();
            WaitForLoad();

            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new RecipesVariantPage(_webDriver, _testContext);
        }
        public void AddRecipeSO(string recipeName)
        {
            _addRecipe = WaitForElementIsVisible(By.Id(ADD_RECIPE));
            _addRecipe.SetValue(ControlType.TextBox, recipeName);
            var selectRecipe = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div[3]/div/div/div/div[3]/div[2]/table/tbody/tr[1]/td[1]"));
            selectRecipe.Click();
            WaitForLoad();
        }
        public bool IsFirstRecipeDeleted()
        {
            try
            {
                _webDriver.FindElement(By.XPath(FIRST_RECIPE));
                return false;
            }
            catch
            {
                return true;
            }
        }
        public SupplyOrderItem ClickOnItems()
        {
            var items = WaitForElementIsVisible(By.Id(ITEMS));
            items.Click();
            WaitForLoad();
            return new SupplyOrderItem(_webDriver, _testContext);
        }
    }
}
