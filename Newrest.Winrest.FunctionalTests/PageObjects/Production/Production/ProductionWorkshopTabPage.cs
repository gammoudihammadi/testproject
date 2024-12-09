using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Production.Production
{
    public class MappedRecipeAndQuantity
    {
        public string recipeName { get; set; }
        public string qtyToProduce { get; set; }
    }

    public class ProductionWorkshopTabPage : PageBase
    {
        public ProductionWorkshopTabPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext) 
        { 
        }

        // ______________________________________ Constantes _____________________________________________

        private const string WORKSHOPNAME_COLUMN_TITLE = "//*[@id=\"list-item-with-action\"]/div[1]/div/div/div[2]/table/thead/tr/th[1]";
        private const string WORKSHOPNAME_COLUMN = "//*[@id=\"list-item-with-action\"]/div[{0}]/div[1]/div/div[2]/table/tbody/tr/td[1]";
        private const string WORKSHOPNAMES_COLUMN = "//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]/table/tbody/tr/td[1]";
        private const string WORKSHOPNAMES_COLUMN_TOGGLE_BTN = "//*[@id=\"list-item-with-action\"]/div[{0}]/div[1]/div/div[1]";
        private const string WORKSHOPNAMES_COLUMN_CURRENT_RECIPE = "/html/body/div[3]/div/div[2]/div[2]/div/div/div/div[{0}]/div[2]/div/div/table/tbody/tr[1]/td[2]";
        private const string WORKSHOPNAMES_COLUMN_RECIPENAME_COLUMN = "/html/body/div[3]/div/div[2]/div[2]/div/div/div/div[{0}]/div[2]/div/div/table/tbody/tr[*]/td[2]";

        private const string WORKSHOPNAMES_COLUMN_CURRENT_RECIPE_TOGGLEBTN = "/html/body/div[3]/div/div[2]/div[2]/div/div/div/div[{0}]/div[2]/div/div/table/tbody/tr[1]/td[1]/div/div/div/div";
        private const string WORKSHOPNAMES_COLUMN_CURRENT_QTYTOPRODUCE = "/html/body/div[3]/div/div[2]/div[2]/div/div/div/div[{0}]/div[2]/div/div/table/tbody/tr[2]/td/div/div/div/table/tbody/tr/td[4]";
        private const string FIRST_WORKSHOP_NAME = "//*[@id=\"list-item-with-action\"]/div[2]/div[1]/div/div[2]/table/tbody/tr/td[1]";
        private const string FIRST_WORKSHOP_TOGGLE_BTN = "//*[@id=\"list-item-with-action\"]/div[2]/div[1]/div/div[1]/span";
        private const string FIRST_WORKSHOP_RECIPE = "/html/body/div[3]/div/div[2]/div[2]/div/div/div/div[2]/div[2]/div/div/table/tbody/tr[1]/td[2]";

        // ______________________________________ Variables _____________________________________________

        [FindsBy(How = How.XPath, Using = WORKSHOPNAME_COLUMN_TITLE)]
        private IWebElement _columnTitle;

        [FindsBy(How = How.XPath, Using = WORKSHOPNAME_COLUMN_TITLE)]
        private IWebElement _workshopNameColumn;

        [FindsBy(How = How.XPath, Using = FIRST_WORKSHOP_NAME)]
        private IWebElement _firstWorkshopName;
        
        [FindsBy(How = How.XPath, Using = WORKSHOPNAMES_COLUMN_TOGGLE_BTN)]
        private IWebElement _workshopNameColumnCurrentToggleBtn;
        
        [FindsBy(How = How.XPath, Using = WORKSHOPNAMES_COLUMN_CURRENT_RECIPE)]
        private IWebElement _workshopNameColumnCurrentRecipe;
        
        [FindsBy(How = How.XPath, Using = WORKSHOPNAMES_COLUMN_RECIPENAME_COLUMN)]
        private IWebElement _workshopNameColumnRecipeColmun;
        
        [FindsBy(How = How.XPath, Using = WORKSHOPNAMES_COLUMN_CURRENT_RECIPE_TOGGLEBTN)]
        private IWebElement _workshopNameColumnCurrentRecipeToggleBtn;
        
        [FindsBy(How = How.XPath, Using = WORKSHOPNAMES_COLUMN_CURRENT_QTYTOPRODUCE)]
        private IWebElement _workshopNameColumnCurrentQtyToProduce;
        
        [FindsBy(How = How.XPath, Using = FIRST_WORKSHOP_TOGGLE_BTN)]
        private IWebElement _firstWorkshopToggleBtn;
        
        [FindsBy(How = How.XPath, Using = FIRST_WORKSHOP_RECIPE)]
        private IWebElement _firstWorkshopRecipe;

        public bool IsResultDisplayedByWorkshop(string value)
        {
            _columnTitle = WaitForElementIsVisible(By.XPath(WORKSHOPNAME_COLUMN_TITLE));
            _firstWorkshopName = WaitForElementIsVisible(By.XPath(FIRST_WORKSHOP_NAME));

            bool result = false;

            int tot = CheckTotalNumber();

            if (tot == 0)
                result = false;

            for (int i = 1; i <= tot; i++)
            {
                try
                {
                    _workshopNameColumn = _webDriver.FindElement(By.XPath(String.Format(WORKSHOPNAME_COLUMN, i + 1)));
                    if (_workshopNameColumn.Text.Equals(value) && _columnTitle.Text.Equals("Workshop")) 
                    {
                        result = true;
                        break;
                    }
                        
                }
                catch
                {
                    result = false;
                }
            }

            return result;
        }

        public string GetFirstWorkshopName()
        {
            _firstWorkshopName = WaitForElementIsVisible(By.XPath(FIRST_WORKSHOP_NAME));
            return _firstWorkshopName.Text;
        }

        public List<string> GetWorkshopsNames()
        {
            int i = 0;

            // map : workshop --> recette
            var mapWorkshopRecipe = new List<string>();

            var columnWorkshopsNames = _webDriver.FindElements(By.XPath(WORKSHOPNAMES_COLUMN));

            foreach (var workshop in columnWorkshopsNames)
            {
                // On limite le nombre de menus remontés à 15 pour ne pas surcharger le test
                if (i >= 15)
                    break;

                _workshopNameColumnCurrentToggleBtn = _webDriver.FindElement(By.XPath(String.Format(WORKSHOPNAMES_COLUMN_TOGGLE_BTN, i + 2)));
                _workshopNameColumnCurrentToggleBtn.Click();
                //animation
                WaitForLoad();
                
                var workshopNameColumnCurrentRecipe = _webDriver.FindElements(By.XPath(String.Format("/html/body/div[3]/div/div[2]/div[2]/div/div/div/div[{0}]/div[2]/div/div/table/tbody/tr[*]/td[2]", i + 2)));
                foreach(var recipeName in workshopNameColumnCurrentRecipe)
                {
                    mapWorkshopRecipe.Add(recipeName.Text);
                  
                }
                i++;
            }

            return mapWorkshopRecipe;
        }

        // OUTPUT : <workshopName, MappedRecipeAndQuantity<recipeName, qtyToProduce>>
        public Dictionary<string, MappedRecipeAndQuantity> GetWorkshopNamesAndMappedRecipesQties()
        {
            int i = 0;

            Dictionary<string, MappedRecipeAndQuantity> workshopNamesAndMappedRecipesQties = new Dictionary<string, MappedRecipeAndQuantity>();

            var columnWorkshopsNames = _webDriver.FindElements(By.XPath(WORKSHOPNAMES_COLUMN));

            foreach (var workshop in columnWorkshopsNames)
            {
                // On limite le nombre de menus remontés à 5 pour ne pas surcharger le test
                if (i >= 4)
                    break;

                _workshopNameColumnCurrentToggleBtn = _webDriver.FindElement(By.XPath(String.Format(WORKSHOPNAMES_COLUMN_TOGGLE_BTN, i + 2)));
                _workshopNameColumnCurrentToggleBtn.Click();

                WaitForLoad();

                _workshopNameColumnCurrentRecipeToggleBtn = _webDriver.FindElement(By.XPath(String.Format(WORKSHOPNAMES_COLUMN_CURRENT_RECIPE_TOGGLEBTN, i + 2)));
                _workshopNameColumnCurrentRecipeToggleBtn.Click();

                _workshopNameColumnCurrentRecipe = _webDriver.FindElement(By.XPath(String.Format(WORKSHOPNAMES_COLUMN_CURRENT_RECIPE, i + 2)));

                _workshopNameColumnCurrentQtyToProduce = _webDriver.FindElement(By.XPath(String.Format(WORKSHOPNAMES_COLUMN_CURRENT_QTYTOPRODUCE, i + 2)));

                MappedRecipeAndQuantity mappedRecipeAndQuantity = new MappedRecipeAndQuantity();

                var qtyString = _workshopNameColumnCurrentQtyToProduce.Text;

                if (qtyString.Length >= 6)
                {
                    var startQty = qtyString.IndexOf(":") + 2;
                    mappedRecipeAndQuantity.qtyToProduce = qtyString.Substring(startQty);
                }
                else
                {
                    mappedRecipeAndQuantity.qtyToProduce = qtyString;
                }

                mappedRecipeAndQuantity.recipeName = _workshopNameColumnCurrentRecipe.Text;

                workshopNamesAndMappedRecipesQties.Add(workshop.Text, mappedRecipeAndQuantity);

                i++;
            }

            return workshopNamesAndMappedRecipesQties;
        }


        public string GetFirstWorkshopRecipe()
        {
            _firstWorkshopToggleBtn = _webDriver.FindElement(By.XPath(FIRST_WORKSHOP_TOGGLE_BTN));
            _firstWorkshopToggleBtn.Click();

            _firstWorkshopRecipe = _webDriver.FindElement(By.XPath(FIRST_WORKSHOP_RECIPE));
            return _firstWorkshopRecipe.Text;
        }
    }
}
