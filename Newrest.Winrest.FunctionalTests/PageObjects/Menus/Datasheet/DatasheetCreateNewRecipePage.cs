using DocumentFormat.OpenXml.Drawing.Charts;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Menus.Datasheet
{
    public class DatasheetCreateNewRecipePage : PageBase
    {

        public DatasheetCreateNewRecipePage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // ____________________________________________ Constantes ___________________________________________

        private const string COMMERCIAL_NAME = "Name";
        private const string COMMERCIAL_NAME2 = "CommercialName2";

        private const string COOKING_MODE = "CookingModeId";
        private const string WORKSHOP = "WorkshopId";
        private const string CREATE_RECIPE = "//*[@id=\"recipe-create-form\"]/div/div[2]/div/button[2]";
        private const string INGREDIENT_NAME = "IngredientName";
        private const string FIRST_INGREDIENT = "//*[@id=\"ingredient-list\"]/table/tbody/tr[1]/td[2]";
        private const string CLOSE_BTN = "btn-from-datasheet-close-modal";
        private const string CLOSE_WINDOW = "//*[@id=\"modal-1\"]/div[1]/button/span";
        private const string Close_Btn_Sous_Reciep = "//*[@id=\"modal-1\"]/div[1]/button";
        private const string FIRSTLINE = "/html/body/div[3]/div/div[2]/div[3]/div/div/div[2]/div/div/div/div/div[2]/div[1]/ul/li";
        private const string EDITFIRSTLINE = "/html/body/div[3]/div/div[2]/div[3]/div/div/div[2]/div/div/div/div/div[2]/div[1]/ul/li/div/div[1]/div[1]/div[11]/div/a[2]";
        private const string DELETESOUSRECIEP = "//*[@id=\"menus-datasheet-datasheetrecipe-deletervds-edit-1\"]";
        private const string TOTAL_PORTION = "recipe_WeightByPortion";

        // ____________________________________________ Variables ____________________________________________

        [FindsBy(How = How.Id, Using = COMMERCIAL_NAME)]
        private IWebElement _commercialName;

        [FindsBy(How = How.XPath, Using = Close_Btn_Sous_Reciep)]
        private IWebElement _closebtnsousreciep;

        [FindsBy(How = How.XPath, Using = DELETESOUSRECIEP)]
        private IWebElement _deletesousreciep;

        [FindsBy(How = How.XPath, Using = EDITFIRSTLINE)]
        private IWebElement _edit;
        [FindsBy(How = How.XPath, Using = FIRSTLINE)]
        private IWebElement _firstline;
        [FindsBy(How = How.Id, Using = COMMERCIAL_NAME2)]
        private IWebElement _commercialName2;
        [FindsBy(How = How.Id, Using = COOKING_MODE)]
        private IWebElement _cookingMode;

        [FindsBy(How = How.Id, Using = WORKSHOP)]
        private IWebElement _workshop;

        [FindsBy(How = How.XPath, Using = CREATE_RECIPE)]
        private IWebElement _createRecipe;

        [FindsBy(How = How.Id, Using = INGREDIENT_NAME)]
        private IWebElement _ingredientName;

        [FindsBy(How = How.XPath, Using = FIRST_INGREDIENT)]
        private IWebElement _firstIngredient;

        [FindsBy(How = How.Id, Using = CLOSE_BTN)]
        private IWebElement _closeBtn;

        [FindsBy(How = How.XPath, Using = CLOSE_WINDOW)]
        private IWebElement _closeWindow;

        // ____________________________________________ Méthodes _____________________________________________

        public DatasheetDetailsPage FillFields_AddNewRecipeToDatasheet(string commercialName, string ingredient = null, string cookingMode = null, string workshop = null)
        {
            _commercialName = WaitForElementIsVisible(By.Id(COMMERCIAL_NAME));
            _commercialName.SetValue(ControlType.TextBox, commercialName);
            _commercialName.SendKeys(Keys.Tab);

            if (cookingMode != null)
            {
                _cookingMode = WaitForElementIsVisible(By.Id(COOKING_MODE));
                _cookingMode.SetValue(ControlType.DropDownList, cookingMode);
            }

            if (workshop != null)
            {
                _workshop = WaitForElementIsVisible(By.Id(WORKSHOP));
                _workshop.SetValue(ControlType.DropDownList, workshop);
            }

            _createRecipe = WaitForElementToBeClickable(By.XPath(CREATE_RECIPE));
            _createRecipe.Click();
            WaitPageLoading();

            if (ingredient != null)
            {
                AddIngredient(ingredient);
            }
            if (isElementExists(By.Id(CLOSE_BTN)))
            {
                _closeBtn = WaitForElementIsVisible(By.Id(CLOSE_BTN));
                _closeBtn.Click();
            }
            else
            {
                WaitPageLoading();
                _closeBtn = WaitForElementIsVisible(By.Id("btn-from-datasheet-close-modal-sub"));
                _closeBtn.Click();
            }


            WaitPageLoading();
            WaitForLoad();

            return new DatasheetDetailsPage(_webDriver, _testContext);
        }
        public void Closebtn()
        {
            WaitPageLoading();

            _closebtnsousreciep = WaitForElementIsVisible(By.XPath(Close_Btn_Sous_Reciep));
            _closebtnsousreciep.Click();
        }

        public void Editfirstline()
        {
            _edit = WaitForElementIsVisible(By.XPath(EDITFIRSTLINE));
            try
            {
                _edit.Click();
            }
            catch (StaleElementReferenceException ex) {
                _edit = WaitForElementIsVisible(By.XPath(EDITFIRSTLINE));
                _edit.Click();
            }
            
        }
        public void FIRSTLINECLICK()
        {
            WaitForLoad();
            _firstline = WaitForElementIsVisible(By.XPath(FIRSTLINE));
            try
            {
                _firstline.Click();
            }
            catch (StaleElementReferenceException ex)
            {
                _firstline = WaitForElementIsVisible(By.XPath(FIRSTLINE));
                _firstline.Click();
            }

            
        }

        public DatasheetDetailsPage FillFields_AddNewRecipeToDatasheetCommercial2(string commercialName, string commercialName2, string ingredient = null, string cookingMode = null, string workshop = null)
        {
            _commercialName = WaitForElementIsVisible(By.Id(COMMERCIAL_NAME));
            _commercialName.SetValue(ControlType.TextBox, commercialName);
            _commercialName.SendKeys(Keys.Tab);

            _commercialName2 = WaitForElementIsVisible(By.Id(COMMERCIAL_NAME2));
            _commercialName2.SetValue(ControlType.TextBox, commercialName2);
            _commercialName2.SendKeys(Keys.Tab);
            if (cookingMode != null)
            {
                _cookingMode = WaitForElementIsVisible(By.Id(COOKING_MODE));
                _cookingMode.SetValue(ControlType.DropDownList, cookingMode);
            }

            if (workshop != null)
            {
                _workshop = WaitForElementIsVisible(By.Id(WORKSHOP));
                _workshop.SetValue(ControlType.DropDownList, workshop);
            }

            _createRecipe = WaitForElementToBeClickable(By.XPath(CREATE_RECIPE));
            _createRecipe.Click();
            WaitForLoad();

            if (ingredient != null)
            {
                AddIngredient(ingredient);
            }
            if (isElementExists(By.Id(CLOSE_BTN)))
            {
                _closeBtn = WaitForElementIsVisible(By.Id(CLOSE_BTN));
                _closeBtn.Click();
            }
            else
            {
                _closeBtn = WaitForElementIsVisible(By.Id("btn-from-datasheet-close-modal-sub"));
                _closeBtn.Click();
            }


            WaitPageLoading();
            WaitForLoad();

            return new DatasheetDetailsPage(_webDriver, _testContext);
        }

        public void AddIngredient(string ingredient)
        {
            _ingredientName = WaitForElementIsVisible(By.Id(INGREDIENT_NAME));
            _ingredientName.SetValue(ControlType.TextBox, ingredient);
            WaitPageLoading();
            WaitForLoad();
           
            var searchButton = WaitForElementIsVisible(By.Id("ingredient-search-button-id"));
            searchButton.Click();
            WaitForLoad();
           
            _firstIngredient = WaitForElementIsVisible(By.XPath(FIRST_INGREDIENT));
            _firstIngredient.Click();
            WaitPageLoading();
            WaitForLoad();
        }

        public DatasheetDetailsPage CloseWindow()
        {
            WaitPageLoading();
            _closeWindow = WaitForElementIsVisible(By.XPath(CLOSE_WINDOW));
            _closeWindow.Click();
            WaitPageLoading();

            return new DatasheetDetailsPage(_webDriver, _testContext);
        }
    
        public float GetTotalNetWeight()
        {
            WaitPageLoading();
            var first_line = WaitForElementIsVisible(By.Id("datasheet-datasheetrecipevariantdetail-recipe-1"));
            first_line.Click();
            var total_NetWeight = WaitForElementIsVisible(By.XPath("/html/body/div[4]/div/div/div[2]/div[2]/div[2]/div/div/div[1]/div[2]/ul/li/div/div/div/form/table[1]/tbody/tr/td[4]/div/div[2]/input"));
            string total_weight = total_NetWeight.GetAttribute("value");
            float parsedWeight;
            if (float.TryParse(total_weight, out parsedWeight))
            {
                return parsedWeight;
            }
            else
            {throw new FormatException();
            }
        }

        public double GetTotalPortion(string decimalSeparator)
        {
            WaitPageLoading();
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            var totalPortion = WaitForElementIsVisible(By.Id(TOTAL_PORTION));
            var total = totalPortion.GetAttribute("value");

            return Convert.ToDouble(total, ci);
        }
        public float GetFirstNetWeight()
        {
            WaitPageLoading();
            var first_line = WaitForElementIsVisible(By.Id("menus-datasheet-datasheetrecipevariantdetail-item-1"));
            first_line.Click(); 
            var fisrt_NetWeight = WaitForElementIsVisible(By.XPath("/html/body/div[4]/div/div/div[2]/div[2]/div[2]/div/div/div[1]/div[2]/ul/li/ul/li[1]/div/div/div/form/table[1]/tbody/tr/td[4]/div/div[2]/input"));
            string total_weight = fisrt_NetWeight.GetAttribute("value");
            float parsedWeight;
            if (float.TryParse(total_weight, out parsedWeight))
            {
                return parsedWeight;
            }
            else
            {
                throw new FormatException();
            }
        }  
        public float GetSecondNetWeight()
        {
            WaitPageLoading();
            var second_line = WaitForElementIsVisible(By.XPath("//*[@id=\"datasheet-datasheetrecipevariantdetail-item-2\"]"));
            second_line.Click();
            var fisrt_NetWeight = WaitForElementIsVisible(By.XPath("/html/body/div[4]/div/div/div[2]/div[2]/div[2]/div/div/div[1]/div[2]/ul/li/ul/li[2]/div/div/div/form/table[1]/tbody/tr/td[4]/div/div[2]/input"));
            string total_weight = fisrt_NetWeight.GetAttribute("value");
            float parsedWeight;
            if (float.TryParse(total_weight, out parsedWeight))
            {
                return parsedWeight;
            }
            else
            {
                throw new FormatException();
            }
        }
        public void ChangeFirstNetWeight(string  NewWeight)
        {
            var first_line = WaitForElementIsVisible(By.XPath("//*[@id=\"datasheet-datasheetrecipevariantdetail-item-1\"]"));
            first_line.Click();
            var fisrt_NetWeight = WaitForElementIsVisible(By.XPath("/html/body/div[4]/div/div/div[2]/div[2]/div[2]/div/div/div[1]/div[2]/ul/li/ul/li[1]/div/div/div/form/table[1]/tbody/tr/td[4]/div/div[2]/input"));
            fisrt_NetWeight.Clear(); 
            fisrt_NetWeight.SendKeys(NewWeight);
            WaitPageLoading();
        }
    }

    }

