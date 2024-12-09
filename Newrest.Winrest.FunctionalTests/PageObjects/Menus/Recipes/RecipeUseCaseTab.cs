using DocumentFormat.OpenXml.VariantTypes;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Menus.Recipes
{
    public class RecipeUseCaseTab : PageBase
    {
        private const string DATE_FROM_FILTER = "date-picker-start";
        private const string DATE_TO_FILTER = "date-picker-end";
        private const string SEARCH_FILTER = "SearchPatternWithAutocomplete";
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";
        
        private const string SELECT_MENU_CHECKBOX = "//*/span[text()='{0}']/../../../../../div/div/input";
        private const string SELECT_FIRST_MENU_CHECKBOX = "(//*[@id='recipeUseCase_IsSelected'])[1]";
        private const string FIRST_MENU_NAME = "//*[@id='tabContentItemContainer']/div[2]/div/div[2]/li[1]/div/div[2]/div/div[4]/p/span"; 

        private const string REPLACE_BY_ANOTHER_RECIPE = "//*[@id='tabContentItemContainer']/div[1]/div/a[1]";
        private const string REPLACE_BY_RECIPE = "searchVM_SearchPattern";
        private const string REPLACE_BY_RECIPE_BUTTON = "//*[@id='list-item-with-action']/div/div/div/div/table/tbody/tr/td[2]/div/div/a[2]";
        private const string REPLACE_BY_VALIDATE = "dataConfirmok";
        private const string REPLACE_BY_CLOSE = "//*[@id=\"modal-1\"]//button[contains(text(), 'Close')]";
        private const string FIRST_RECIPE = "/html/body/div[3]/div/div[2]/div/div/div[5]/div/div/div[2]/div[2]/div/div[2]/li[1]/div/div[1]/div/input[1]";
        private const string SELECT_ALL = "//*[@id=\"tabContentItemContainer\"]/div[1]/div/a[4]";
        private const string CANCEL = "//*[@id=\"modal-1\"]/div[1]/button";

        [FindsBy(How = How.Id, Using = DATE_FROM_FILTER)]
        private IWebElement _fromDate;

        [FindsBy(How = How.Id, Using = SEARCH_FILTER)]
        private IWebElement _searchFilter;

        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;

        [FindsBy(How = How.Id, Using = DATE_TO_FILTER)]
        private IWebElement _toDate;

        public RecipeUseCaseTab(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        public enum FilterType
        {
            DateFrom,
            DateTo,
            SearchMenu
        }


        public void Filter(FilterType filterType, object value)
        {
            switch (filterType)
            {
                case FilterType.DateFrom:
                    _fromDate = WaitForElementIsVisible(By.Id(DATE_FROM_FILTER));
                    _fromDate.SetValue(ControlType.DateTime, value);
                    _fromDate.SendKeys(Keys.Tab);
                    break;
                case FilterType.DateTo:
                    _toDate = WaitForElementIsVisible(By.Id(DATE_TO_FILTER));
                    _toDate.SetValue(ControlType.DateTime, value);
                    _toDate.SendKeys(Keys.Tab);
                    break;
                case FilterType.SearchMenu:
                    _searchFilter = WaitForElementIsVisible(By.Id(SEARCH_FILTER));
                    _searchFilter.SetValue(ControlType.TextBox, value);
                    break;

            }

            WaitPageLoading();
            WaitForLoad();
        }

        public RecipesPage BackToList()
        {
            _backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            _backToList.Click();

            WaitForLoad();

            return new RecipesPage(_webDriver, _testContext);
        }

        public string GetFirstMenuName()
        {
            var recipeMenuSelect = WaitForElementIsVisible(By.XPath(FIRST_MENU_NAME));
            return recipeMenuSelect.Text;
        }

        public void SelectFirstMenuCheckBox()
        {
            var recipeSelect = WaitForElementExists(By.XPath(SELECT_FIRST_MENU_CHECKBOX));
            recipeSelect.SetValue(PageBase.ControlType.CheckBox, true);
        }

        public void ReplaceByAnotherRecipe(string recipe)
        {
            //Cliquer sur "Replace By Another Recipe
            var replace = WaitForElementIsVisible(By.XPath(REPLACE_BY_ANOTHER_RECIPE));
            replace.Click();
            WaitForLoad();
            // SEARCH A RECIPE TO REPLACE "*ARROZ CON CONEJO PMI"
            // RecipeForTestMenu
            var replaceWith = WaitForElementIsVisible(By.Id(REPLACE_BY_RECIPE));
            replaceWith.SetValue(PageBase.ControlType.TextBox, recipe);
            WaitForLoad();
            var buttonReplace = WaitForElementIsVisible(By.XPath(REPLACE_BY_RECIPE_BUTTON));
            buttonReplace.Click();
            WaitPageLoading();

            var validate = WaitForElementIsVisible(By.Id(REPLACE_BY_VALIDATE));
            validate.Click();
            WaitPageLoading();

            var closeXPath = isElementVisible(By.XPath(REPLACE_BY_CLOSE)) ? REPLACE_BY_CLOSE : CANCEL;
            var close = WaitForElementIsVisible(By.XPath(closeXPath));
            close.Click();
            WaitForLoad();
        }

        public void SelectMenuCheckBox(string menu)
        {
            // //*/span[text()='Menu-14/11/2022-1134223950']
            // /html/body/div[3]/div/div[2]/div/div/div[5]/div/div/div[2]/div[2]/div/div[2]/li[1]/div/div[1]/div/input
            // /html/body/div[3]/div/div[2]/div/div/div[5]/div/div/div[2]/div[2]/div/div[2]/li[1]/div/div[2]/div/div[4]/p/span
            var recipeSelect = WaitForElementExists(By.XPath(string.Format(SELECT_MENU_CHECKBOX,menu)));
            recipeSelect.SetValue(PageBase.ControlType.CheckBox, true);
        }
        public void SelectAll()
        {
            WaitForLoad();
            var selectall = _webDriver.FindElement(By.XPath(SELECT_ALL));
            selectall.Click();
            WaitForLoad();
                }
        public void ClickOnFirstRecipe()
        {
            var firstrecipe = _webDriver.FindElement(By.XPath(FIRST_RECIPE));
            firstrecipe.Click();
            WaitForLoad();
        }
    }
    
}
