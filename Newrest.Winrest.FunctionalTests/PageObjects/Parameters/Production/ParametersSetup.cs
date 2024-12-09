using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Linq;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Production
{
    public class ParametersSetup : PageBase
    {
        // _____________________________________________ Constantes _______________________________________

        private const string PARAMETERS_SETUP_TAB = "SettingsSetupTabNav";

        // PAX per recipe variant
        private const string PAX_PER_RECIPE_VARIANT_TAB = "RecipesVariants_ListItem";
        private const string SEARCH_MENU_FILTER = "SearchPattern";
        private const string DROPDOWN_MENU_BUTTON = "//*[@id=\"tabContentParameters\"]/div[1]/div[2]/div[1]/div/div[1]/button";
        private const string SELECT_ALL = "//*[@id=\"tabContentParameters\"]/div[1]/div[2]/div[1]/div/div[1]/div/a[2]";
        private const string SELECT_ONE_RECIPE = "//label[contains(text(),'{0}')]/../..//span[contains(text(),'{1}')]/../..//parent::label";
        private const string PAX_GASTRO_PACKAGING_RECIPE_VARIANT = "//label[contains(text(),'{0}')]/../..//span[contains(text(),'{1}')]/../..//table";
        private const string DROPDOWN_ADD_BUTTON = "//*[@id=\"tabContentParameters\"]/div[1]/div[2]/div[1]/div/div[2]/button";
        private const string MULTIPLE_UPDATE_BUTTON = "btn-massive-update";

        //MASSIVE UPDATE MODAL
        private const string IS_FOODPACK = "//*[@id=\"FoodPackUseCases_{0}__FoodPackId\"]/option[@selected = 'selected']";
        private const string ADD_FOODPACK_BUTTON = "//*[@id=\"modal-1\"]/div/div/div/div/div/form/div[2]/div[*]/div[2]/div/div/a[2]";
        private const string FOODPACK_SELECT = "//*[@id=\"modal-1\"]/div/div/div/div/div/form/div[2]/div[{0}]/div[1]/select";
        private const string FOODPACK_VALUE_INPUT = "//*[@id=\"modal-1\"]/div/div/div/div/div/form/div[2]/div[{0}]/div[2]/input";
        private const string FOODPACK_GROUPED_BY_DELIVERY_CHECKBOX = "//*[@id=\"modal-1\"]/div/div/div/div/div/form/div[2]/div[{0}]/div[3]/label";
        private const string FOODPACK_DROPDOWN2 = "//*[@id=\"modal-1\"]/div/div/div/div/div/form/div[2]/div[4]/div[1]/select";
        private const string FOODPACK_VALUE_INPUT2 = "//*[@id=\"modal-1\"]/div/div/div/div/div/form/div[2]/div[4]/div[2]/input";
        private const string UPDATE_BUTTON = "btn-popup-update";

        // PAX per recipe
        private const string PAX_PER_RECIPE_TAB = "Recipes_ListItem";
        private const string PAX_GASTRO_PACKAGING_RECIPE = "//label[contains(text(),'{0}')]/../..//table";

        // ____________________________________________ Variables _________________________________________

        [FindsBy(How = How.Id, Using = PARAMETERS_SETUP_TAB)]
        private IWebElement _parametersSetupTab;

        // PAX per recipe variant
        [FindsBy(How = How.Id, Using = PAX_PER_RECIPE_VARIANT_TAB)]
        private IWebElement _paxPerRecipeVariant_tab;

        [FindsBy(How = How.Id, Using = SEARCH_MENU_FILTER)]
        private IWebElement _searchMenuFilter;

        [FindsBy(How = How.XPath, Using = DROPDOWN_MENU_BUTTON)]
        private IWebElement _dropdownMenuButton;

        [FindsBy(How = How.XPath, Using = SELECT_ALL)]
        private IWebElement _selectAll;

        [FindsBy(How = How.XPath, Using = SELECT_ONE_RECIPE)]
        private IWebElement _selectOneRecipe;

        [FindsBy(How = How.XPath, Using = PAX_GASTRO_PACKAGING_RECIPE_VARIANT)]
        private IWebElement _pax_gastro_packaging_recipe_variant;

        [FindsBy(How = How.XPath, Using = DROPDOWN_ADD_BUTTON)]
        private IWebElement _dropdownAddButton;

        [FindsBy(How = How.Id, Using = MULTIPLE_UPDATE_BUTTON)]
        private IWebElement _multipleUpdateButton;

        //MASSIVE UPDATE MODAL

        [FindsBy(How = How.XPath, Using = ADD_FOODPACK_BUTTON)]
        private IWebElement _addFoodPackButton;

        [FindsBy(How = How.XPath, Using = FOODPACK_SELECT)]
        private IWebElement _foodPackDropdown;

        [FindsBy(How = How.XPath, Using = FOODPACK_VALUE_INPUT)]
        private IWebElement _foodPackValueInput;

        [FindsBy(How = How.XPath, Using = FOODPACK_GROUPED_BY_DELIVERY_CHECKBOX)]
        private IWebElement _foodPackGroupeByDelivery;

        [FindsBy(How = How.XPath, Using = FOODPACK_DROPDOWN2)]
        private IWebElement _foodPackDropdown2;

        [FindsBy(How = How.XPath, Using = FOODPACK_VALUE_INPUT2)]
        private IWebElement _foodPackValueInput2;
        [FindsBy(How = How.Id, Using = UPDATE_BUTTON)]
        private IWebElement _updateButton;

        // PAX per recipe
        [FindsBy(How = How.Id, Using = PAX_PER_RECIPE_TAB)]
        private IWebElement _paxPerRecipe_tab;

        [FindsBy(How = How.XPath, Using = PAX_GASTRO_PACKAGING_RECIPE)]
        private IWebElement _pax_gastro_packaging_recipe;

        // ___________________________________________ Méthodes ___________________________________________

        public ParametersSetup(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {

        }

        //
        public bool CheckSetupTabSelected()
        {
            _parametersSetupTab = WaitForElementIsVisible(By.Id(PARAMETERS_SETUP_TAB));
            string parametersSetupTabClass = _parametersSetupTab.GetAttribute("class");
            if (parametersSetupTabClass == "menu-item active") return true;
            return false;
        }

        // Got tot Tab PAX per recipe variant
        public void GoToTab_PAXPerRecipeVariant()
        {
            _paxPerRecipeVariant_tab = WaitForElementToBeClickable(By.Id(PAX_PER_RECIPE_VARIANT_TAB));
            _paxPerRecipeVariant_tab.Click();
            // async
            WaitPageLoading();
        }

        // Got tot Tab PAX per recipe
        public void GoToTab_PAXPerRecipe()
        {
            _paxPerRecipe_tab = WaitForElementToBeClickable(By.Id(PAX_PER_RECIPE_TAB));
            _paxPerRecipe_tab.Click();
            WaitForLoad();
        }

        public enum FilterType
        {
            SearchRecipe
        }

        //Filter on Tabs PAX per recipe variant & PAX per recipe
        public void Filter(FilterType filterType, object value)
        {
            switch (filterType)
            {
                case FilterType.SearchRecipe:
                    _searchMenuFilter = WaitForElementIsVisible(By.Id(SEARCH_MENU_FILTER));
                    _searchMenuFilter.SetValue(ControlType.TextBox, value);
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);
            }

            WaitPageLoading();
        }


        public void SelectAll()
        {
            _dropdownMenuButton = WaitForElementIsVisible(By.XPath(DROPDOWN_MENU_BUTTON));
            WaitPageLoading();
            _dropdownMenuButton.Click();
            WaitForLoad();

            _selectAll = WaitForElementIsVisible(By.XPath(SELECT_ALL));
            WaitPageLoading();
            _selectAll.Click();
            WaitForLoad();
        }
        public void SelectRecipe(string recipe, string variant)
        {
            _selectOneRecipe = WaitForElementIsVisible(By.XPath(string.Format((SELECT_ONE_RECIPE), recipe, variant)));
            var actions = new Actions(_webDriver);
            actions.MoveToElement(_selectOneRecipe).Click().Build().Perform();
            WaitForLoad();
        }

        public void ClickOnMultipleUpdate()
        {
            _dropdownAddButton = WaitForElementIsVisible(By.XPath(DROPDOWN_ADD_BUTTON));
            _dropdownAddButton.Click();
            WaitForLoad();

            _multipleUpdateButton = WaitForElementIsVisible(By.Id(MULTIPLE_UPDATE_BUTTON));
            _multipleUpdateButton.Click();
            WaitPageLoading();

        }

        public string GetFoodPackPaxGastroPackaging(string recipe, string variant)
        {
            if (variant != null)
            {
                _pax_gastro_packaging_recipe_variant = WaitForElementIsVisible(By.XPath(string.Format(PAX_GASTRO_PACKAGING_RECIPE_VARIANT, recipe, variant)));
                return _pax_gastro_packaging_recipe_variant.Text;
            }
            else
            {
                _pax_gastro_packaging_recipe = WaitForElementIsVisible(By.XPath(string.Format(PAX_GASTRO_PACKAGING_RECIPE, recipe)));
                return _pax_gastro_packaging_recipe.Text;
            }
        }
        public void AddFoodPackPerRecipe(string foodPack, string foodPackValue, bool groupedByDelivery = false)
        {
            int compteur = 0;
            bool isFound = false;
            WaitForLoad();

            //limité à 4 ajouts, possibilité d'augmenter
            while (compteur <= 3 && !isFound)
            {
                var elementPresent = _webDriver.FindElements(By.XPath(string.Format(IS_FOODPACK, compteur))).Count;

                if (elementPresent == 0)
                {
                    isFound = true;
                    _addFoodPackButton = WaitForElementIsVisible(By.XPath("//*[@id=\"modal-1\"]/div/div/div/form/div[2]/div[*]/div[2]/div/div/a[2]"), nameof(ADD_FOODPACK_BUTTON));
                    _addFoodPackButton.Click();
                    WaitForLoad();

                    var foodpackDiv = 3 + compteur;
                    _foodPackDropdown = WaitForElementIsVisible(By.XPath(string.Format("//*[@id=\"modal-1\"]/div/div/div/form/div[2]/div[{0}]/div[1]/select", foodpackDiv)), nameof(FOODPACK_SELECT));

                    _foodPackDropdown.SetValue(ControlType.DropDownList, foodPack);
                    WaitForLoad();

                    _foodPackValueInput = WaitForElementIsVisible(By.XPath(string.Format("//*[@id=\"modal-1\"]/div/div/div/form/div[2]/div[{0}]/div[2]/input", foodpackDiv)), nameof(FOODPACK_VALUE_INPUT));

                    _foodPackValueInput.SetValue(ControlType.TextBox, foodPackValue);
                    WaitForLoad();

                    if (groupedByDelivery == true)
                    {
                        _foodPackGroupeByDelivery = WaitForElementIsVisible(By.XPath(string.Format("//*[@id=\"modal-1\"]/div/div/div/form/div[2]/div[{0}]/div[3]/label", foodpackDiv)), nameof(FOODPACK_GROUPED_BY_DELIVERY_CHECKBOX));

                        var actions = new Actions(_webDriver);
                        actions.MoveToElement(_foodPackGroupeByDelivery).Click().Build().Perform();
                        WaitForLoad();
                    }
                }
                else
                {
                    if (_webDriver.FindElement(By.XPath(string.Format(IS_FOODPACK, compteur))).Text == foodPack)
                    {
                        isFound = true;
                    }
                    else
                    {
                        compteur++;
                    }
                }
            }

            _updateButton = WaitForElementIsVisible(By.Id(UPDATE_BUTTON));
            _updateButton.Click();
            WaitForLoad();
        }

        public void AddPaxPerRecipe(string recipe, string variant, string foodPack, string foodPackValue, bool groupedByDelivery = false)
        {
            Filter(ParametersSetup.FilterType.SearchRecipe, recipe);
            var paxGastroPackaging = GetFoodPackPaxGastroPackaging(recipe, variant);
            if (!new[] { foodPack, foodPackValue }.Any(item => paxGastroPackaging.Contains(item)))
            {
                SelectRecipe(recipe, variant);
                ClickOnMultipleUpdate();
                AddFoodPackPerRecipe(foodPack, foodPackValue, groupedByDelivery);
                SelectRecipe(recipe, variant);
                paxGastroPackaging = GetFoodPackPaxGastroPackaging(recipe, variant);
                Assert.IsTrue(new[] { foodPack, foodPackValue }.Any(item => paxGastroPackaging.Contains(item)), "Dans Setup, l'ajout de pax par packaging de la recette {0} ne s'est pas fait.", recipe);
            }
        }

        public void RemoveFoodPackPerRecipe(string foodpack)
        {
            if (isElementExists(By.XPath(string.Format("//*[text()='{0}' and @selected='selected']", foodpack))))
            {
                var remove = WaitForElementIsVisible(By.XPath(string.Format("//*[text()='{0}' and @selected='selected']/parent::select/parent::div/parent::div/div[4]/a", foodpack)));
                remove.Click();
                WaitForLoad();
            }
        }
    }
}