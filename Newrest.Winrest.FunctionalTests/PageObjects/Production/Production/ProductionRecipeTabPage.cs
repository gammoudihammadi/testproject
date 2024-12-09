using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Production.Production
{

    public class ProductionRecipeTabPage : PageBase
    {
        public ProductionRecipeTabPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // ______________________________________ Constantes _____________________________________________
        private const string MENU_TAB = "//*[@id=\"itemTabTab\"]/li[1]";

        private const string FOLD_UNFOLD_ALL = "//*[@id=\"tabContentItemContainer\"]/div[1]/div/div/div/a[1]";
        private const string RECIPENAME_COLUMN_TITLE = "//*[@id=\"list-item-with-action\"]/div[1]/div/div/div[2]/table/thead/tr/th[1]";
        private const string RECIPENAMES_COLUMN = "//*[@id=\"list-item-with-action\"]/div[*]/div/div/div[2]/table/tbody/tr/td[1]";
        private const string FIRST_RECIPE_NAME = "//*[@id=\"list-item-with-action\"]/div[2]/div[1]/div/div[2]/table/tbody/tr/td[1]";
        private const string FIRST_RECIPE_FOODPACK = "//*[@id=\"list-item-with-action\"]/div[2]/div[1]/div/div[2]/table/tbody/tr/td[3]";
        private const string RECIPE_FOODPACK = "//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]/table/tbody/tr/td[1][contains(text(),'{0}')]/../td[3]";
        private const string CURRENT_RECIPE_QTY = "//*[@id=\"list-item-with-action\"]/div[{0}]/div/div/div[2]/table/tbody/tr/td[5]";
        private const string CURRENT_RECIPE_PAX = "//*[@id=\"list-item-with-action\"]/div[{0}]/div/div/div[2]/table/tbody/tr/td[4]";
        private const string FIRST_RECIPE_ITEM_NAME = "//*[@id=\"list-item-with-action\"]/div[{0}]/div[2]/div/div/table/tbody/tr/td[1]/span";
        private const string CURRENT_RECIPE_ITEM_NETWEIGHT = "//*[@id=\"list-item-with-action\"]/div[{0}]/div[2]/div/div/table/tbody/tr/td[2]/span[2]";

        // ______________________________________ Variables _____________________________________________
        [FindsBy(How = How.XPath, Using = MENU_TAB)]
        private IWebElement _menuTab;

        [FindsBy(How = How.XPath, Using = FOLD_UNFOLD_ALL)]
        private IWebElement _foldUnfoldAll;

        [FindsBy(How = How.XPath, Using = RECIPENAME_COLUMN_TITLE)]
        private IWebElement _columnTitle;

        [FindsBy(How = How.XPath, Using = FIRST_RECIPE_NAME)]
        private IWebElement _firstRecipeName;

        [FindsBy(How = How.XPath, Using = FIRST_RECIPE_FOODPACK)]
        private IWebElement _firstRecipeFoodpack;

        [FindsBy(How = How.XPath, Using = RECIPE_FOODPACK)]
        private IWebElement _recipeFoodpack;

        [FindsBy(How = How.XPath, Using = CURRENT_RECIPE_QTY)]
        private IWebElement _currentRecipeQty;

        [FindsBy(How = How.XPath, Using = CURRENT_RECIPE_PAX)]
        private IWebElement _currentRecipePAX;

        [FindsBy(How = How.XPath, Using = FIRST_RECIPE_ITEM_NAME)]
        private IWebElement _firstRecipeItemName;

        [FindsBy(How = How.XPath, Using = CURRENT_RECIPE_ITEM_NETWEIGHT)]
        private IWebElement _currentRecipeNetweight;

        //ONGLETS
        public ProductionSearchResultsMenuTabPage GoToProductionMenuTab()
        {
            _menuTab = WaitForElementIsVisible(By.XPath(MENU_TAB));
            _menuTab.Click();
            WaitForLoad();
            return new ProductionSearchResultsMenuTabPage(_webDriver, _testContext);
        }
        public bool IsResultDisplayedByRecipe()
        {
            _columnTitle = WaitForElementIsVisible(By.XPath(RECIPENAME_COLUMN_TITLE));
            _firstRecipeName = WaitForElementIsVisible(By.XPath(FIRST_RECIPE_NAME));

            if (_columnTitle.Text.Equals("Recipe") && _firstRecipeName.Text.Contains("Recipe"))
            {
                return true;

            }
            else
            {
                return false;
            }
        }

        public Dictionary<string, string> GetRecipesNamesAndQtiesToProduce()
        {
            int i = 0;

            Dictionary<string, string> recipesNamesAndQtiesToProduce = new Dictionary<string, string>();

            var columnRecipesNames = _webDriver.FindElements(By.XPath(RECIPENAMES_COLUMN));

            foreach (var recipe in columnRecipesNames)
            {
                // On limite le nombre de menus remontés à 3 pour ne pas surcharger le test
                if (i >= 3)
                    break;

                // Le type de recette est concaténé au nom de la recette, nous conservons uniquement le nom de la recette
                var recipeConcatenee = recipe.Text;
                var indexRecipeNameStart = recipeConcatenee.IndexOf("-") + 2;

                var recipeName = recipeConcatenee.Substring(indexRecipeNameStart);

                _currentRecipeQty = _webDriver.FindElement(By.XPath(String.Format(CURRENT_RECIPE_QTY, i + 2)));

                recipesNamesAndQtiesToProduce.Add(recipeName, _currentRecipeQty.Text);
                i++;
            }

            return recipesNamesAndQtiesToProduce;
        }

        public Dictionary<string, string> GetRecipesPAX()
        {
            int i = 0;

            Dictionary<string, string> recipesPAX = new Dictionary<string, string>();

            var columnRecipesNames = _webDriver.FindElements(By.XPath(RECIPENAMES_COLUMN));

            foreach (var recipe in columnRecipesNames)
            {
                // On limite le nombre de menus remontés à 3 pour ne pas surcharger le test
                if (i >= 3)
                    break;

                // Le type de recette est concaténé au nom de la recette, nous conservons uniquement le nom de la recette
                var recipeConcatenee = recipe.Text;
                var indexRecipeNameStart = recipeConcatenee.IndexOf("-") + 2;

                var recipeName = recipeConcatenee.Substring(indexRecipeNameStart);

                _currentRecipePAX = _webDriver.FindElement(By.XPath(String.Format(CURRENT_RECIPE_PAX, i + 2)));

                recipesPAX.Add(recipeName, _currentRecipePAX.Text);
                i++;
            }

            return recipesPAX;
        }

        public string GetFirstRecipeName()
        {
            _firstRecipeName = WaitForElementIsVisible(By.XPath(FIRST_RECIPE_NAME));

            var recipeConcatenee = _firstRecipeName.Text;

            var indexRecipeNameStart = recipeConcatenee.IndexOf("-") + 2;

            var recipeName = recipeConcatenee.Substring(indexRecipeNameStart);

            return recipeName;
        }

        public string GetFirstRecipeFoodPack()
        {
            _firstRecipeFoodpack = WaitForElementIsVisible(By.XPath(FIRST_RECIPE_FOODPACK));

            return _firstRecipeFoodpack.Text;
        }
        public string GetRecipeFoodPack(string recipe)
        {
            _recipeFoodpack = WaitForElementIsVisible(By.XPath(string.Format(RECIPE_FOODPACK, recipe)));

            return _recipeFoodpack.Text;
        }

        public Dictionary<string, string> GetRecipesItemsNetWeight()
        {
            int i = 0;

            Dictionary<string, string> recipesNamesAndQtiesToProduce = new Dictionary<string, string>();

            FoldUnfoldAll();

            var columnRecipesNames = _webDriver.FindElements(By.XPath(RECIPENAMES_COLUMN));

            foreach (var recipe in columnRecipesNames)
            {
                // On limite le nombre de menus remontés à 3 pour ne pas surcharger le test
                if (i >= 3)
                    break;

                // Le type de recette est concaténé au nom de la recette, nous conservons uniquement le nom de la recette
                var recipeConcatenee = recipe.Text;
                var indexRecipeNameStart = recipeConcatenee.IndexOf("-") + 2;

                var recipeName = recipeConcatenee.Substring(indexRecipeNameStart);

                _currentRecipeNetweight = _webDriver.FindElement(By.XPath(String.Format(CURRENT_RECIPE_ITEM_NETWEIGHT, i + 2)));

                recipesNamesAndQtiesToProduce.Add(recipeName, _currentRecipeNetweight.Text);
                i++;
            }

            return recipesNamesAndQtiesToProduce;
        }

        public bool FoldUnfoldAll()
        {
            ShowExtendedMenu();
            _foldUnfoldAll = WaitForElementIsVisible(By.XPath(FOLD_UNFOLD_ALL));
            _foldUnfoldAll.Click();
            //Temps d'affichage du résultat
            WaitPageLoading();
            if (isElementVisible(By.XPath("//*[@id=\"list-item-with-action\"]/div[2]/div[2]/div/div/table/tbody/tr/td[1]/span")))
            {
                _firstRecipeItemName = WaitForElementIsVisible(By.XPath("//*[@id=\"list-item-with-action\"]/div[2]/div[2]/div/div/table/tbody/tr/td[1]/span"));
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
