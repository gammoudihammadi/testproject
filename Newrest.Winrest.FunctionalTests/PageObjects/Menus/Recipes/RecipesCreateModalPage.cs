using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Menus.Recipes
{
    public class RecipesCreateModalPage : PageBase
    {
        // __________________________________________ Constantes _________________________________________________

        private const string RECIPE_NAME = "Name";
        private const string RECIPE_COMMERCIALNAME1 = "CommercialName1";
        private const string RECIPE_TYPE = "RecipeTypeId";
        private const string COOKING_MODE = "CookingModeId";
        private const string WORKSHOP = "WorkshopId";
        private const string DIETARY_TYPE = "DietaryTypeId";
        private const string ACTIVATED = "IsActive";
        private const string NUMBER_OF_PORTIONS = "NumberOfPortions";
        private const string SAVE = "//*[@id=\"recipe-create-form\"]/div/div[2]/div/button[2]";
        private const string ERROR_MESSAGE = "//*[@id=\"recipe-create-form\"]/div/div[1]/div[1]/div[1]/div[1]/div/span";

        // __________________________________________ Variables __________________________________________________

        [FindsBy(How = How.Id, Using = RECIPE_NAME)]
        private IWebElement _recipeName; 
        
        [FindsBy(How = How.Id, Using = RECIPE_COMMERCIALNAME1)]
        private IWebElement _commercialName1;

        [FindsBy(How = How.Id, Using = RECIPE_TYPE)]
        private IWebElement _recipeType;

        [FindsBy(How = How.Id, Using = DIETARY_TYPE)]
        private IWebElement _dietaryType;

        [FindsBy(How = How.Id, Using = COOKING_MODE)]
        private IWebElement _cookingMode;

        [FindsBy(How = How.Id, Using = WORKSHOP)]
        private IWebElement _workshop;

        [FindsBy(How = How.Id, Using = ACTIVATED)]
        private IWebElement _activated;

        [FindsBy(How = How.Id, Using = NUMBER_OF_PORTIONS)]
        private IWebElement _numberOfPortions;

        [FindsBy(How = How.XPath, Using = SAVE)]
        private IWebElement _save;

        [FindsBy(How = How.XPath, Using = ERROR_MESSAGE)]
        private IWebElement _errorMessage;

        // __________________________________________ Méthodes ___________________________________________________

        public RecipesCreateModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        public RecipeGeneralInformationPage FillField_CreateNewRecipe(string name, string recipeType, string nbPortions, bool isActive = true, string dietaryType = null, string cookingMode = null, string workshop = null, string commercialName = null)
        {
            _recipeName = WaitForElementIsVisibleNew(By.Id(RECIPE_NAME));
            _recipeName.SetValue(ControlType.TextBox, name);

            _recipeType = WaitForElementIsVisibleNew(By.Id(RECIPE_TYPE));
            _recipeType.SetValue(ControlType.DropDownList, recipeType);

            if(dietaryType != null)
            {
                _dietaryType = WaitForElementIsVisibleNew(By.Id(DIETARY_TYPE));
                _dietaryType.SetValue(ControlType.DropDownList, dietaryType);
            }

            _activated = WaitForElementExists(By.Id(ACTIVATED));
            _activated.SetValue(ControlType.CheckBox, isActive);

            _numberOfPortions = WaitForElementIsVisibleNew(By.Id(NUMBER_OF_PORTIONS));
            _numberOfPortions.SetValue(ControlType.TextBox, nbPortions);

            if (cookingMode != null)
            {
                _cookingMode = WaitForElementIsVisibleNew(By.Id(COOKING_MODE));
                _cookingMode.SetValue(ControlType.DropDownList, cookingMode);
            }

            if (workshop != null)
            {
                _workshop = WaitForElementIsVisibleNew(By.Id(WORKSHOP));
                _workshop.SetValue(ControlType.DropDownList, workshop);
            }

            //Define CommercialName1
            if (!string.IsNullOrEmpty(commercialName))
            {
                _commercialName1 = WaitForElementIsVisibleNew(By.Id(RECIPE_COMMERCIALNAME1));
                _commercialName1.SetValue(ControlType.TextBox, commercialName);
            }

            // Click sur le bouton "Create"
            _save = WaitForElementToBeClickable(By.XPath(SAVE));
            _save.Click();
            WaitPageLoading();

            return new RecipeGeneralInformationPage(_webDriver, _testContext);
        }

        public bool ErrorMessageNameAlreadyExists()
        {
            try
            {
                _errorMessage = _webDriver.FindElement(By.XPath(ERROR_MESSAGE));

                return true;
            }
            catch
            {
                return false;
            }
        }

    }
}
