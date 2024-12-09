using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.VariantTypes;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Datasheet;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Item;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Security.Policy;
using System.Text.RegularExpressions;
using UglyToad.PdfPig.Fonts.TrueType.Tables;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Menus.Recipes
{
    public class RecipeGeneralInformationPage : PageBase
    {
        // ________________________________________ Constantes __________________________________________

        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";
        private const string VARIANT_IN_MENU = "//h4[text()='{0}']";

        // Summary
        private const string SUMMARY = "//*[@id=\"recipe-site-variant-container\"]/div/ul/li[1]/a";
        private const string SUMMARY_FIRST_VARIANT = "//*[@id=\"summary\"]/div[2]/table/tbody/tr/td[1]/b";
        private const string VARIANT_TOTAL = "//*[@id=\"summary\"]/div[1]/div/div[2]";
        private const string VARIANT_NAME = "//*[@id=\"recipe-variant-list\"]/tbody/tr/td/h4[text()=\"{0}\"]/../div";

        private const string VARIANT_SITE = "//*[@id=\"summary\"]/div[2]/table/tbody/tr[{0}]/td[2]/b";

        // Edit Information        
        private const string EDIT_INFO_DEV = "editInformationTab";
        private const string EDIT_INFO_PATCH = "menus-datasheet-editdatasheetrecipe-edit-1"; //"//*[@id=\"recipe-site-variant-container\"]/div/ul/li[2]/a";
        private const string SELECT_FIRST_ITEM = "//*[@id=\"dataItem1264894\"]/div/div[1]/div[1]";

        private const string RECIPE_NAME = "Name";
        private const string RECIPE_COMM_NAME1 = "CommercialName1";
        private const string RECIPE_COMM_NAME2 = "CommercialName2";
        private const string COOKING_MODE_ID = "CookingModeId";
        private const string COOKING_MODE_SELECTED = "//*[@id=\"CookingModeId\"]/option[@selected = 'selected']";
        private const string RECIPE_TYPE_ID = "RecipeTypeId";
        private const string RECIPE_TYPE_SELECTED = "//*[@id=\"RecipeTypeId\"]/option[@selected = 'selected']";
        private const string RECIPE_DIETARY_ID = "DietaryTypeId";
        private const string RECIPE_DIETARY_SELECTED = "//*[@id=\"DietaryTypeId\"]/option[@selected = 'selected']";
        private const string WORKSHOP_ID = "WorkshopId";
        private const string WORKSHOP_SELECTED = "//*[@id=\"WorkshopId\"]/option[@selected = 'selected']";
        private const string NUMBER_PORTIONS = "NumberOfPortions";
        private const string IS_ACTIVE = "IsActive";
        private const string RECIPE_CONSERVE_MODE = "ConserveMode";
        private const string RECIPE_CONSERVE_TEMPERATURE = "ConserveTemperature";
        private const string RECIPE_INFORMATION_COMMENT = "Comment";
        private const string FIRST_RECIPE_NAME = "menus-datasheet-datasheet-recipe-name-1";

        // Select variants
        private const string SEARCH_VARIANT_SITE = "SearchVM_Site";
        private const string SEARCH_VARIANT_NAME = "SearchVM_Variant";
        private const string VALIDATE_VARIANT = "Validate_SiteVariants";
        private const string FIRST_VARIANT = "//*[@id=\"recipe-variant-list\"]/tr/td/h4";
        private const string FIRST_VARIANT_LIST = "//*[@id=\"recipe-variant-list\"]/tbody/tr/td/h4";
        private const string VARIANT_WITH_SITE = "//*[starts-with(@id,\"link-recipe-variant-\")]";


        // Use case tab
        public const string USE_CASE_TAB = "useCaseTab";
        public const string RECIPE_INFORMATION_CLOSE = "btn-from-datasheet-close-modal";

        public const string RECIPE_NUMBERE = "//*[@id=\"tabContentItemContainer\"]/div[1]/h1/span[1]";
        public const string KEYWORDS = "hrefTabContentVariantDetailsKeyword";
        public const string INTOLERANCE_TAB = "hrefTabContentVariantDetailsIntolerance";

        // Keywords Tab
        public const string KEYWORDS_LIST = "//*[@id=\"tabContentDetails\"]/div/table/tbody/tr[*]/td";
        public const string FIRSTVARIANT = "//*[@id=\"recipe-siteVariant-container\"]/div/div/form/table/tbody/tr[2]";
        public const string VALIDATE = "//*[@id=\"Validate_SiteVariants\"]";


        public const string SELECT_FIRST_VARIANT = "SiteVariants_0__IsSelected";
        public const string FIRST_KEYWORD = "//*[@id=\"tabContentDetails\"]/div/table/tbody/tr[2]/td";


        // ________________________________________ Variables ___________________________________________

        [FindsBy(How = How.Id, Using = INTOLERANCE_TAB)]
        private IWebElement _intoleranceTab;
        [FindsBy(How = How.XPath, Using = FIRSTVARIANT)]
        private IWebElement _selectfirstvariant;
        [FindsBy(How = How.XPath, Using = VALIDATE)]
        private IWebElement _validate;
        [FindsBy(How = How.Id, Using = SELECT_FIRST_VARIANT)]
        private IWebElement _selectFirstVariant;

        [FindsBy(How = How.XPath, Using = RECIPE_NUMBERE)]
        private IWebElement _recipesnumber;

        [FindsBy(How = How.XPath, Using = KEYWORDS)]
        private IWebElement _keywords;


        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;

        // Summary
        [FindsBy(How = How.XPath, Using = SUMMARY)]
        private IWebElement _summary;

        [FindsBy(How = How.XPath, Using = SUMMARY_FIRST_VARIANT)]
        private IWebElement _summaryFirstVariant;

        [FindsBy(How = How.XPath, Using = VARIANT_TOTAL)]
        private IWebElement _variantTotal;

        // Edit Information
        [FindsBy(How = How.Id, Using = EDIT_INFO_DEV)]
        private IWebElement _editInformationDev;

        [FindsBy(How = How.Id, Using = EDIT_INFO_PATCH)]
        private IWebElement _editInformationPatch;

        [FindsBy(How = How.Id, Using = RECIPE_NAME)]
        private IWebElement _recipeName;

        [FindsBy(How = How.XPath, Using = RECIPE_TYPE_SELECTED)]
        private IWebElement _recipeType;

        [FindsBy(How = How.XPath, Using = COOKING_MODE_SELECTED)]
        private IWebElement _cookingMode;

        [FindsBy(How = How.XPath, Using = WORKSHOP_SELECTED)]
        private IWebElement _workshop;

        [FindsBy(How = How.Id, Using = NUMBER_PORTIONS)]
        private IWebElement _nbPortions;

        [FindsBy(How = How.Id, Using = IS_ACTIVE)]
        private IWebElement _isActive;

        [FindsBy(How = How.Id, Using = RECIPE_CONSERVE_MODE)]
        private IWebElement _conserveMode;

        [FindsBy(How = How.Id, Using = RECIPE_CONSERVE_TEMPERATURE)]
        private IWebElement _conserveTemperature;

        [FindsBy(How = How.Id, Using = RECIPE_INFORMATION_COMMENT)]
        private IWebElement _recipeInformationComment;
        [FindsBy(How = How.Id, Using = FIRST_RECIPE_NAME)]
        private IWebElement _firstRecipeName;

        // Select variants
        [FindsBy(How = How.Id, Using = SEARCH_VARIANT_SITE)]
        private IWebElement _searchVariantBySite;

        [FindsBy(How = How.Id, Using = VALIDATE_VARIANT)]
        private IWebElement _validateVariant;

        [FindsBy(How = How.XPath, Using = FIRST_VARIANT)]
        private IWebElement _firstVariant;

        [FindsBy(How = How.Id, Using = USE_CASE_TAB)]
        private IWebElement _useCaseTab;


        [FindsBy(How = How.Id, Using = RECIPE_INFORMATION_CLOSE)]
        private IWebElement _close;
        // ________________________________________ Méthodes ____________________________________________

        public RecipeGeneralInformationPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        public RecipesPage BackToList()
        {
            _backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            _backToList.Click();

            return new RecipesPage(_webDriver, _testContext);
        }

        public bool IsVariantInAsideMenu(List<string> variants)
        {
            bool isPresent = true;

            if (variants != null)
            {
                foreach (string variant in variants)
                {
                    try
                    {
                        string newVariant = variant.Replace("+", "-");
                        WaitForElementIsVisible(By.XPath(String.Format(VARIANT_IN_MENU, newVariant)));
                    }
                    catch
                    {
                        isPresent = false;
                        break;
                    }
                }
            }
            else
            {
                isPresent = false;
            }

            return isPresent;
        }

        // Summary
        public void ClickOnSummary()
        {
            _summary = WaitForElementIsVisible(By.XPath(SUMMARY));
            _summary.Click();
            WaitForLoad();
        }

        public RecipesVariantPage ClickOnFirstVariant()
        {
            _summaryFirstVariant = WaitForElementIsVisibleNew(By.XPath(SUMMARY_FIRST_VARIANT));
            _summaryFirstVariant.Click();
            WaitForLoad();

            return new RecipesVariantPage(_webDriver, _testContext);
        }

        public int GetVariantTotal(string decimalSeparator)
        {
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");
            _variantTotal = WaitForElementIsVisible(By.XPath(VARIANT_TOTAL));
            var variantNumber = Regex.Replace(_variantTotal.Text, @"\D", "");
            WaitPageLoading();

            return Convert.ToInt32(variantNumber, ci);
        }

        public bool VerifySiteVariant(int nbVariants, string site)
        {
            bool valueBool = true;

            for (int i = 1; i <= nbVariants; i++)
            {
                IWebElement variantSite = _webDriver.FindElement(By.XPath(String.Format(VARIANT_SITE, i)));

                if (variantSite.Text.Contains(site))
                {
                    valueBool = true;
                    break;
                }
                else
                {
                    valueBool = false;
                }
            }

            return valueBool;
        }


        // Edit Information
        public void ClickOnEditInformation()
        {
            _editInformationDev = WaitForElementIsVisible(By.Id(EDIT_INFO_DEV));
            _editInformationDev.Click();
            WaitPageLoading();
        }

        public void UpdateRecipe(string updateValue, string nbPortions)
        {
            _recipeName = WaitForElementIsVisible(By.Id(RECIPE_NAME));
            _recipeName.SetValue(ControlType.TextBox, updateValue);

            _nbPortions = WaitForElementIsVisible(By.Id(NUMBER_PORTIONS));
            _nbPortions.SetValue(ControlType.TextBox, nbPortions);

            // Temps de prise en compte des modifs
            WaitPageLoading();
        }

        public void ClickOnActivatedCheckbox()
        {
            _isActive = WaitForElementExists(By.Id(IS_ACTIVE));
            _isActive.Click();

            // Temps de prise en compte du changement
            WaitPageLoading();
        }

        public string GetRecipeName()
        {
            _recipeName = WaitForElementExists(By.Id(RECIPE_NAME));
            return _recipeName.GetAttribute("value");
        }

        public void SetCommercialName1(string commName1)
        {
            var _commName1 = WaitForElementExists(By.Id(RECIPE_COMM_NAME1));
            _commName1.SetValue(ControlType.TextBox, commName1);
            WaitForLoad();
        }

        public string GetCommercialName1()
        {
            var _commName1 = WaitForElementExists(By.Id(RECIPE_COMM_NAME1));
            return _commName1.GetAttribute("value");
        }

        public void SetCommercialName2(string commName2)
        {
            var _commName2 = WaitForElementExists(By.Id(RECIPE_COMM_NAME2));
            _commName2.SetValue(ControlType.TextBox, commName2);
            WaitForLoad();
        }

        public string GetCommercialName2()
        {
            var _commName2 = WaitForElementExists(By.Id(RECIPE_COMM_NAME2));
            return _commName2.GetAttribute("value");
        }

        public int GetNumberOfPortions()
        {
            _nbPortions = WaitForElementIsVisible(By.Id(NUMBER_PORTIONS));
            return int.Parse(_nbPortions.GetAttribute("value"));
        }

        public void SetRecipeType(string recipeTypeName)
        {
            DropdownListSelectById(RECIPE_TYPE_ID, recipeTypeName);
        }

        public string GetRecipeType()
        {
            _recipeType = WaitForElementIsVisible(By.XPath(RECIPE_TYPE_SELECTED));
            return _recipeType.Text.Trim();
        }

        public void SetWorkshop(string workshopName)
        {
            DropdownListSelectById(WORKSHOP_ID, workshopName);
        }

        public string GetWorkshop()
        {
            _workshop = WaitForElementIsVisible(By.XPath(WORKSHOP_SELECTED));
            return _workshop.Text.Trim();
        }

        public void SetCookingMode(string cookingMode)
        {
            DropdownListSelectById(COOKING_MODE_ID, cookingMode);
        }

        public string GetCookingMode()
        {
            _cookingMode = WaitForElementIsVisible(By.XPath(COOKING_MODE_SELECTED));
            return _cookingMode.Text;
        }

        public void SetDietaryType(string dietaryTypeName)
        {
            DropdownListSelectById(RECIPE_DIETARY_ID, dietaryTypeName);
        }

        public string GetDietaryType()
        {
            var _dietaryType = WaitForElementIsVisible(By.XPath(RECIPE_DIETARY_SELECTED));
            return _dietaryType.Text;
        }

        public string GetConserveMode()
        {
            _conserveMode = WaitForElementIsVisible(By.Id(RECIPE_CONSERVE_MODE));
            return _conserveMode.GetAttribute("value");
        }

        public void SetConserveMode(string conserveMode)
        {
            _conserveMode = WaitForElementIsVisible(By.Id(RECIPE_CONSERVE_MODE));
            _conserveMode.SetValue(ControlType.TextBox, conserveMode);
            WaitForLoad();
        }

        public string GetConserveTemperature()
        {
            _conserveTemperature = WaitForElementIsVisible(By.Id(RECIPE_CONSERVE_TEMPERATURE));
            return _conserveTemperature.GetAttribute("value");
        }

        public void SetConserveTemperature(string conserveTemperature)
        {
            _conserveTemperature = WaitForElementIsVisible(By.Id(RECIPE_CONSERVE_TEMPERATURE));
            _conserveTemperature.SetValue(ControlType.TextBox, conserveTemperature);
            WaitForLoad();
        }

        public string GetComment()
        {
            var _comment = WaitForElementIsVisible(By.Id(RECIPE_INFORMATION_COMMENT));
            return _comment.GetAttribute("value");
        }

        public void SetComment(string comment)
        {
            var _comment = WaitForElementIsVisible(By.Id(RECIPE_INFORMATION_COMMENT));
            _comment.SetValue(ControlType.TextBox, comment);
            WaitForLoad();
        }

        // Select variants
        public void AddVariantWithSite(string site, string variant)
        {
        	_searchVariantBySite = WaitForElementIsVisibleNew(By.Id(SEARCH_VARIANT_SITE));
        	_searchVariantBySite.SetValue(ControlType.TextBox, site);
            WaitForLoad();
            WaitPageLoading();
            
            var selectedVariant = WaitForElementIsVisibleNew(By.XPath("//td[text()='" + variant + "']"));
            selectedVariant.Click();
            WaitForLoad();
                
            if (isElementVisible(By.Id(VALIDATE_VARIANT)))
            {
                _validateVariant = WaitForElementToBeClickable(By.Id(VALIDATE_VARIANT));
                Actions action = new Actions(_webDriver);
                action.MoveToElement(_validateVariant).Click().Perform();
                _validateVariant.Click();
                WaitForLoad();
            }
        }
        public void AddSite(string site)
        {
            WaitPageLoading();
            _searchVariantBySite = WaitForElementIsVisible(By.Id(SEARCH_VARIANT_SITE));
            _searchVariantBySite.SetValue(ControlType.TextBox, site);
            WaitPageLoading();
            WaitForLoad();
            var selectVariant = WaitForElementIsVisible(By.Id("recipe-select-all-text"));
            selectVariant.Click();  // Set checkbox value to true
            WaitForLoad();
            var validate = WaitForElementIsVisible(By.Id("Validate_SiteVariants"));
            validate.Click();
            WaitForLoad();
        }
        public RecipesVariantPage SelectFirstVariantFromList()
        {
            WaitPageLoading();
            WaitForLoad();
            _firstVariant = !isElementVisible(By.XPath(FIRST_VARIANT_LIST)) ? WaitForElementIsVisibleNew(By.XPath(FIRST_VARIANT)) : WaitForElementIsVisibleNew(By.XPath(FIRST_VARIANT_LIST));
            WaitPageLoading(); 
            _firstVariant.Click();
            WaitForLoad();

            return new RecipesVariantPage(_webDriver, _testContext);
        }

        public bool IsFirstVariantIsVisible()
        {
            WaitPageLoading();
            if (isElementVisible(By.XPath(FIRST_VARIANT)))
            {
                return true;

            }
            else if (isElementVisible(By.XPath(FIRST_VARIANT_LIST)))
            {
                return true;
            }
            WaitForLoad();

            return false;
        }
        public void SelectFirstVariant()
        {
            _selectfirstvariant = WaitForElementIsVisible(By.XPath(FIRSTVARIANT));
            _selectfirstvariant.Click();
            WaitForLoad();
        }
        public void Validate()
        {
            _validate = WaitForElementIsVisible(By.XPath(VALIDATE));
            _validate.Click();
            WaitPageLoading();
            WaitForLoad();
        }


        public RecipesVariantPage SelectVariant(string variant)
        {

            var variantToClick = WaitForElementIsVisible(By.XPath(String.Format(VARIANT_IN_MENU, variant)));
            variantToClick.Click();
            WaitForLoad();

            return new RecipesVariantPage(_webDriver, _testContext);
        }
        public RecipesVariantPage SearchVariant(string variant, string site)
        {
            _searchVariantBySite = WaitForElementIsVisible(By.Id(SEARCH_VARIANT_SITE));
            _searchVariantBySite.SetValue(ControlType.TextBox, site);
            WaitPageLoading();

            _searchVariantBySite = WaitForElementIsVisible(By.Id(SEARCH_VARIANT_NAME));
            _searchVariantBySite.SetValue(ControlType.TextBox, variant);
            WaitPageLoading();



            return new RecipesVariantPage(_webDriver, _testContext);
        }

        public RecipesVariantPage SelectVariantWithSite(string site)
        {
            bool isVariants = false;

            var variants = _webDriver.FindElements(By.XPath(VARIANT_WITH_SITE));

            foreach (var elm in variants)
            {
                if (elm.Text.Contains(site))
                {
                    elm.Click();
                    isVariants = true;
                    WaitForLoad();
                    break;
                }
            }

            if (!isVariants)
            {
                throw new Exception("Aucune variante n'a été créé pour la recette.");
            }

            return new RecipesVariantPage(_webDriver, _testContext);
        }

        public void AddMultipleVariants(string site, List<String> variants)
        {
            if (variants != null)
            {
                _searchVariantBySite = WaitForElementIsVisible(By.Id(SEARCH_VARIANT_SITE));
                _searchVariantBySite.SetValue(ControlType.TextBox, site);
                WaitPageLoading();

                foreach (String variant in variants)
                {
                    var selectedVariant1 = WaitForElementIsVisible(By.XPath("//td[text()='" + variant + "']"));
                    selectedVariant1.Click();
                }

                _validateVariant = WaitForElementIsVisible(By.Id(VALIDATE_VARIANT));
                _validateVariant.Click();
                WaitForLoad();
            }
        }

        public string GetRecipeInformationComment()
        {
            _recipeInformationComment = WaitForElementIsVisible(By.Id(RECIPE_INFORMATION_COMMENT));
            return _recipeInformationComment.GetAttribute("value");
        }
        public void SetRecipeInformationComment(string comment)
        {
            _recipeInformationComment = WaitForElementIsVisible(By.Id(RECIPE_INFORMATION_COMMENT));
            _recipeInformationComment.SetValue(ControlType.TextBox, comment);
            WaitPageLoading();
        }

        public RecipeUseCaseTab ClickOnUseCaseTab()
        {
            _useCaseTab = WaitForElementIsVisible(By.Id(USE_CASE_TAB));
            _useCaseTab.Click();
            WaitForLoad();
            return new RecipeUseCaseTab(_webDriver, _testContext);
        }
        public void ClickFirstFlight()
        {
            var firstflight = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div[2]/div/div/div[1]/div[2]/table/tbody/tr[1]/td[2]"));
            firstflight.Click();
        }
        public void AddIngredient(string itemname)
        {
            var ingredientInput = WaitForElementIsVisible(By.Id("IngredientName"));
            ingredientInput.SendKeys(itemname);
            WaitPageLoading();

            var firstingredient = WaitForElementIsVisible(By.XPath("//*[@id=\"ingredient-list\"]/table/tbody/tr[1]/td[2]"));
            firstingredient.Click();
            WaitPageLoading();
        }

        public bool IsIngredientTextColorOrange()
        {
            var ingredientRow = WaitForElementExists(By.XPath("//*[@id=\"menu-recipe-datasheet-datasheetrecipevariantdetail-item-1\"]/td[1]/span"));
            string color = ingredientRow.GetCssValue("color");
            // The expected orange color in RGB format (equivalent to #ff9900)
            string expectedColor = "rgba(255, 153, 0, 1)"; // RGB equivalent of #ff9900
            return color == expectedColor;
        }
        public string GetFirstIngredient()
        {
            WaitForLoad();
            if (isElementVisible(By.XPath("//*/table[@class='display-zone']/tbody/tr/td[1]/span")))
            {
                var firstIngredient = WaitForElementIsVisible(By.XPath("//*/table[@class='display-zone']/tbody/tr/td[1]/span"));
                return firstIngredient.Text;
            }
            return null;
        }
        public string GetRecipeNameFlight()
        {
            _recipeName = WaitForElementExists(By.Id(RECIPE_NAME));
            var name = _recipeName.GetAttribute("value");

            _close = WaitForElementToBeClickable(By.Id(RECIPE_INFORMATION_CLOSE));
            _close.Click();
            WaitForLoad();

            return name;
        }

        public void Go_To_New_Navigate()
        {
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.Close();
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[0]);

        }

        public int CheckTotalRecipesNumber()
        {
            WaitForLoad();
            _recipesnumber = WaitForElementExists(By.XPath(RECIPE_NUMBERE));
            int nombre = Int32.Parse(_recipesnumber.Text);
            return nombre;
        }

        public ItemIntolerancePage ClickOnIntolerancePage()
        {
            _intoleranceTab = WaitForElementIsVisible(By.Id(INTOLERANCE_TAB));
            _intoleranceTab.Click();
            WaitForLoad();

            return new ItemIntolerancePage(_webDriver, _testContext);
        }
        public void ClickOnFirstVariant(string variant)
        {
            var variantName = WaitForElementIsVisible(By.XPath(string.Format(VARIANT_NAME, variant)));
            variantName.Click();
            WaitForLoad();
        }
        public List<string> GetKeywords()
        {
            List<string> keywords = new List<string>();
            var keywordsList = _webDriver.FindElements(By.XPath(KEYWORDS_LIST));
            foreach (var keyword in keywordsList)
            {
                keywords.Add(keyword.Text);
            }
            return keywords;
        }
        public bool VerifyKeywordsValues(List<string> keywordsMenu, List<string> keywordsRecipe)
        {
            foreach (var keyword in keywordsMenu)
            {
                if (!keywordsRecipe.Contains(keyword))
                {
                    return false;
                }
            }
            return true;
        }

        public DatasheetDetailsPage ClickOnDetailsFromDatasheet()
        {

            var details = WaitForElementIsVisible(By.XPath("//*/tr[@data-target='#datasheet-central-Pane'][2]"));
            details.Click();
            WaitPageLoading();
            WaitForLoad();
            //FIXME page vide
            var recipeTab = WaitForElementIsVisible(By.Id("hrefTabContentRecipes"));
            recipeTab.Click();
            WaitForLoad();
            return new DatasheetDetailsPage(_webDriver, _testContext);
        }

        public RecipeKeywordPage ClickOnKeywordPage()
        {
            _keywords = WaitForElementIsVisible(By.Id(KEYWORDS));
            _keywords.Click();
            WaitForLoad();

            return new RecipeKeywordPage(_webDriver, _testContext);
        }
        public void ClickOnFirstRecipe()
        {

            WaitPageLoading();
            _firstRecipeName = WaitForElementExists(By.Id(FIRST_RECIPE_NAME));
            _firstRecipeName.Click();
            WaitForLoad();
        }

        public void SelectVariantRecipe()
        {
            var variantElement = _webDriver.FindElement(By.XPath("//*[@id=\"recipe-variant-list\"]/tr/td/div"));
            variantElement.Click();
        }

        public void AddIngredients(string searchQuery, int numberOfIngredients)
        {
            WaitPageLoading();
            for (int i = 1; i <= numberOfIngredients; i++)
            {
                var searchBox = _webDriver.FindElement(By.XPath("//*[@id=\"IngredientName\"]"));

                if (!searchBox.Displayed || !searchBox.Enabled)
                {
                    IJavaScriptExecutor js = (IJavaScriptExecutor)_webDriver;
                    js.ExecuteScript("document.getElementById('IngredientName').focus();");
                }

                searchBox.Click();
                searchBox.Clear();
                searchBox.SendKeys(searchQuery);

                WaitPageLoading();

                // Sélectionner l'ingrédient correspondant
                var ingredient = _webDriver.FindElement(By.XPath($"//*[@id='ingredient-list']/table/tbody/tr[{i}]"));

                ingredient.Click();

                WaitPageLoading();
            }
        }
        public List<(string, int)> GetIngredientNames()
        {
            WaitPageLoading();
            // Sélectionner tous les éléments <li> dans la liste des ingrédients (à partir du deuxième <li>)
            var ingredients = _webDriver.FindElements(By.XPath(".//li[position() >= 1]//table[2]//td[1]"));
            var ingredientNames = new List<(string, int)>();

            // Parcourir chaque élément <li> pour récupérer le nom de l'
            var i = 1;
            foreach (var ingredient in ingredients)
            {

                string ingredientName = ingredient.Text.ToUpper().Replace(" ", "");
                ingredientNames.Add((ingredientName, i));
                i++;
            }

            return ingredientNames;
        }

        public bool IsKeywordElementPresent()
        {
            WaitPageLoading();
            try
            {
                IWebElement keywordElement = _webDriver.FindElement(By.XPath(FIRST_KEYWORD));
                return keywordElement != null;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }

        public bool IsKeywordTextCorrect(string expectedKeyword)
        {
            IWebElement keywordElement = _webDriver.FindElement(By.XPath(FIRST_KEYWORD));
            string actualKeyword = keywordElement.Text.Trim();
            return actualKeyword == expectedKeyword;
        }

        public bool IsKeywordDisplayed()
        {
            WaitPageLoading();
            IWebElement keywordElement = _webDriver.FindElement(By.XPath(FIRST_KEYWORD));
            return keywordElement.Displayed;
        }

        public bool IsRecipeActivated()
        {
            _isActive = WaitForElementExists(By.Id(IS_ACTIVE));
            if (_isActive.GetAttribute("checked") != "true")
                return false;
            else
                return true;
        }


    }
}
