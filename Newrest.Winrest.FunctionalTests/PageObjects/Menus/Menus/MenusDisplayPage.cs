using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Menus.Menus
{
    public class MenusDisplayPage : PageBase
    {
        public MenusDisplayPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // ______________________________________ Constantes _____________________________________________
        private const string RECIPE_NAME = ".icFyaD";
        private const string RECIPE_POINTS_PRICE = ".cBtlYI";
        private const string RECIPE_SALES_PRICE = ".cUmbsU";
        private const string RECIPE_ENERGY = ".cKCRin";

        [FindsBy(How = How.CssSelector, Using = RECIPE_NAME)]
        private IWebElement _recipeName;

        [FindsBy(How = How.CssSelector, Using = RECIPE_POINTS_PRICE)]
        private IWebElement _recipePointsPrice;

        [FindsBy(How = How.CssSelector, Using = RECIPE_SALES_PRICE)]
        private IWebElement _recipeSalesPrice;

        [FindsBy(How = How.CssSelector, Using = RECIPE_ENERGY)]
        private IWebElement _recipeEnergy;

        public bool CheckHasRecipeName()
        {
            _recipeName = WaitForElementIsVisible(By.CssSelector(RECIPE_NAME));
            return !string.IsNullOrEmpty(_recipeName.Text);
        }

        public bool CheckHasRecipePointsPrice()
        {
            return isElementExists(By.CssSelector(RECIPE_POINTS_PRICE));
        }

        public bool CheckHasRecipeSalesPrice()
        {
            return isElementExists(By.CssSelector(RECIPE_SALES_PRICE));
        }

        public bool CheckHasRecipeEnergy()
        {
            return isElementExists(By.CssSelector(RECIPE_ENERGY));
        }
    }
}
