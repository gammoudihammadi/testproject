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
    public class RecipeImportIngredientPage : PageBase
    {
        public RecipeImportIngredientPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // ____________________________________________ Constantes _________________________________________________

        private const string RECIPE_NAME = "//*[@id=\"scrollable-dropdown-menu\"]/span/input[2]";
        private const string SELECTED_RECIPE = "//*[@id=\"scrollable-dropdown-menu\"]/span/div/div/div[contains(text(),'{0}')]";

        private const string ERASE_INGREDIENT = "EraseIngredients";
        private const string NEXT = "btn-next";
        private const string VARIANT = "//*[@id=\"RecipeToVariantChoice\"]/label[text()='{0}']";
        private const string CREATE = "btn-create";
        private const string OK = "//*[@id=\"modal-result\"]/div[3]/button";
        private const string VARIANT_DATA = "//*[@id=\"RecipeToVariantChoice\"]/label";
        // ____________________________________________ Variables __________________________________________________

        [FindsBy(How = How.XPath, Using = RECIPE_NAME)]
        private IWebElement _recipeName;

        [FindsBy(How = How.Id, Using = ERASE_INGREDIENT)]
        private IWebElement _eraseIngredient;

        [FindsBy(How = How.Id, Using = NEXT)]
        private IWebElement _nextBtn;

        [FindsBy(How = How.XPath, Using = VARIANT)]
        private IWebElement _variant;

        [FindsBy(How = How.Id, Using = CREATE)]
        private IWebElement _create;

        [FindsBy(How = How.XPath, Using = OK)]
        private IWebElement _okBtn;

        // ____________________________________________ Méthodes ___________________________________________________

        public void FillFields_ImportIngredients(string recipeName, string variant, bool eraseIngredient)
        {
            _recipeName = WaitForElementIsVisible(By.XPath(RECIPE_NAME));
            _recipeName.SetValue(ControlType.TextBox, recipeName);
            var selectedRecipe = WaitForElementIsVisible(By.XPath(String.Format(SELECTED_RECIPE, recipeName)));
            selectedRecipe.Click();

            _eraseIngredient = WaitForElementExists(By.Id(ERASE_INGREDIENT));
            _eraseIngredient.SetValue(ControlType.CheckBox, eraseIngredient);

            _nextBtn = WaitForElementToBeClickable(By.Id(NEXT));
            _nextBtn.Click();
            WaitForLoad();

            _variant = WaitForElementIsVisible(By.XPath(string.Format(VARIANT, variant.Replace("+", "-"))));
            _variant.Click();

            _create = WaitForElementToBeClickable(By.Id(CREATE));
            _create.Click();
            WaitForLoad();
        }

        public bool FillFields_ImportIngredientsError(string recipeName)
        {
            _recipeName = WaitForElementIsVisible(By.XPath(RECIPE_NAME));
            _recipeName.SetValue(ControlType.TextBox, recipeName);

            // Temps de mise à jour de l'affichage (moins long que WaitForElementisVisible)
            WaitForLoad();

            try
            {
                var selectedRecipe = _webDriver.FindElement(By.XPath(String.Format(SELECTED_RECIPE, recipeName)));
                selectedRecipe.Click();
                return true;
            }
            catch
            {
                return false;
            }

        }


        public RecipesVariantPage ConfirmImport()
        {
            _okBtn = WaitForElementToBeClickable(By.XPath(OK));
            _okBtn.Click();
            WaitForLoad();

            // Temps de prise en compte des changements
            WaitPageLoading();

            return new RecipesVariantPage(_webDriver, _testContext);
        }

        public bool IsImportOK()
        {
            try
            {
                _okBtn = WaitForElementIsVisible(By.XPath(OK));
                return true;
            }
            catch
            {
                return false;
            }
        }
        public void FillFields_ImportIngredientSub(string recipeName, string variant, bool eraseIngredient)
        {
            _recipeName = WaitForElementIsVisible(By.XPath(RECIPE_NAME));
            _recipeName.SetValue(ControlType.TextBox, recipeName);
            WaitForLoad();
            var selectedRecipe = WaitForElementIsVisible(By.XPath(String.Format(SELECTED_RECIPE, recipeName)));
            selectedRecipe.Click();

            _eraseIngredient = WaitForElementExists(By.Id(ERASE_INGREDIENT));
            _eraseIngredient.SetValue(ControlType.CheckBox, eraseIngredient);

            _nextBtn = WaitForElementToBeClickable(By.Id(NEXT));
            _nextBtn.Click();
            WaitForLoad();

            _variant = WaitForElementIsVisible(By.XPath(string.Format(VARIANT_DATA, variant.Replace("+", "-"))));
            _variant.Click();

            _create = WaitForElementToBeClickable(By.Id(CREATE));
            _create.Click();
            WaitForLoad();
        }
    }
}
