using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.EMMA;
using iText.Commons.Utils;
using iText.StyledXmlParser.Jsoup.Select;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Datasheet;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.CustomerOrder;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;
using System.Windows.Forms.VisualStyles;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Menus.Recipes
{
    public class RecipesVariantPage : PageBase
    {
        // ___________________________________ Constantes _____________________________________________

        // Général
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";
        private const string GENERAL_INFORMATION = "-1";
        private const string GENERAL_INFORMATION_TAB = "hrefTabContentVariantDetailsInfos";
        private const string RECIPE_PRICE = "recipe-display-price";
        private const string RECIPE_NAME = "//h1";
        private const string IMPORT = "//*[@id=\"recipe-site-variant-container\"]/div[1]/div/a[1]";
        private const string IMPORTSUBRECIPE = "//*[@id=\"modal-1\"]/div[2]/div[1]/a";
        private const string COPY_VALUES_FOR_SITE = "//*[@id=\"recipe-site-variant-container\"]/div[1]/div/a[2]";
        private const string DUPLICATE_VARIANT = "//*[@id=\"recipe-site-variant-container\"]/div[1]/div/a[3]";
        private const string PRINT = "//*[@id=\"recipe-site-variant-container\"]/div[1]/div/a[4]";
        private const string CONFIRM_PRINT = "print";
        private const string ITEM = "nameRecipeH1";
        private const string FIRST_ITEM_NAME = "//*[@id=\"datasheet-datasheetrecipevariantdetail-item-1\"]/td[1]/span";
        // Onglets
        private const string ITEMS_TAB = "hrefTabContentVariantDetails";
        private const string PROCEDURE_TAB = "hrefTabContentVariantDetailsProcedure";
        private const string DIETETIC_TAB = "hrefTabContentVariantDetailsDietetic";
        private const string INTOLERANCE_TAB = "hrefTabContentVariantDetailsIntolerance";
        private const string COMPOSITION_TAB = "hrefTabContentVariantDetailsComposition";
        private const string PICTURE_TAB = "hrefTabContentVariantDetailsPicture";
        private const string KEYWORD_TAB = "hrefTabContentVariantDetailsKeyword";

        // Ajout ingredient
        private const string ITEM_NAME = "IngredientName";
        private const string ITEM_CHECK_BOX = "ItemsOnly";
        private const string FIRST_ITEM = "//*[@id=\"ingredient-list\"]/table/tbody/tr[1]/td[2]";
        private const string ERROR_MESSAGE = "//*[@id=\"ingredient-list\"]/p";

        // Tableau des ingrédients display-zone ou edit-zone ? la recipe est en haut et l'item en bas...
        private const string INGREDIENT = "menu-recipe-datasheet-datasheetrecipevariantdetail_item-1";
        private const string ALL_INGREDIENT = "//*[starts-with(@id,\"Datasheet-DatasheetRecipeVariantDetail_item\")]/td[1]/span";
        private const string INGREDIENTLIGNES = "//*[@id=\"recipe-variant-detail-list\"]/div[2]/ul/li[*]";
        private const string FIRST_INGREDIENT = "//*[@id=\"menu-recipe-datasheet-datasheetrecipevariantdetail_item-1\"]"; 
        private const string FIRST_INGREDIENT_IN_DEV = "//*[@id=\"recipe-variant-detail-list\"]/div[2]/ul/li[1]";
        private const string SUB_RECIPE = "datasheet-datasheetrecipevariantdetail-recipe-1";
        private const string SUB_RECIPE_PATCH = "menu-recipe-datasheet-datasheetrecipevariantdetail_recipe-1";
        private const string SECOND_SUB_RECIPE = "datasheet-datasheetrecipevariantdetail-recipe-2";
        private const string ALL_SUB_RECIPE = "//*[starts-with(@id,\"menu-recipe-datasheet-datasheetrecipevariantdetail-recipe-\")]";
        private const string INGREDIENT_NAME = "//*[starts-with(@id,\"rvd-\")]/table[2]/tbody/tr[*]/td[1]/span";
        private const string INGREDIENT_NET_WEIGHT = "recipe_NetWeight";
        private const string INGREDIENT_NET_QTY = "recipe_NetQuantity";
        private const string INGREDIENT_YIELD = "recipe_Yield";
        private const string INGREDIENT_YIELD_DATA = " //input[@class='input-number yield']";
        private const string NET_WEIGHT = "recipe_NetWeight";
        private const string FIRST_lINE = "datasheet-datasheetrecipevariantdetail-item-1";  

        private const string INGREDIENT_GROSS_QTY = "recipe_Quantity";
        private const string INGREDIENT_PRICE = "recipe_PriceEdit";
        private const string INGREDIENT_PRICEMODAL = "//*[starts-with(@id,\"rvd-\")]/table[1]/tbody/tr/td[8]";
        private const string INGREDIENT_WEIGHT = ".display-zone td:nth-child(3) span:first-child";
        private const string LIST_KEYWORD = "//*[@id='tabContentDetails']/div/table/tbody/tr[*]/td";  
        // display-zone ou edit-zone ? la recipe est en haut et l'item en bas...

        private const string EDIT_INGREDIENT = "//span[text()='{0}']/../../td[13]/div/a/span[@class=\"fas fa-pencil-alt\"]";

        private const string EDIT_FIRST_INGREDIENT_EDIT = "//table[@class='edit-zone']//a[starts-with(@href,'/Purchasing/Item/Detail')]/span";
        private const string EDIT_FIRST_INGREDIENT_DISPLAY = "//table[normalize-space(@class)='display-zone']//a[starts-with(@href,'/Purchasing/Item/Detail')]/span";
        private const string EDIT_FIRST_SUB_RECIPE_EDIT = "//table[@class='edit-zone']//a[starts-with(@href,'/Menus/Recipe/Details')]/span"; 
        private const string EDIT_FIRST_SUB_RECIPE = "menus-recipevariantdetail-detail-display-recipe-1";
        private const string EDIT_FIRST_SUB_RECIPE_DATASHEET = "//*[starts-with(@id,\"menus-datasheet-datasheetrecipe-subrecipe-display\")]";
        private const string EDIT_FIRST_SUB_RECIPE_DISPLAY = "//table[normalize-space(@class)='display-zone']//a[starts-with(@href,'/Menus/Recipe/Details')]/span";

        private const string DELETE_FIRST_INGREDIENT_EDIT = "//table[@class='edit-zone']//a[starts-with(@href,'/Menus/Recipe/DeleteRecipeVariantDetailSiteWithDetailSite')]/span";
        private const string DELETE_FIRST_INGREDIENT_DISPLAY = "//table[normalize-space(@class)='display-zone']//a[starts-with(@href,'/Menus/Recipe/DeleteRecipeVariantDetailSiteWithDetailSite')]/span";
        private const string DELETE_INGREDIENT_DISPLAY = "//table[normalize-space(@class)='display-zone']//tr[@data-type = 'Item']//span[contains(text(), '{0}')]/../../td/div/a[starts-with(@href,'/Menus/Recipe/DeleteRecipeVariantDetailSiteWithDetailSite')]";
        private const string DELETE_ING = "//table[normalize-space(@class)='display-zone']//tr[@data-type = 'Item']//span[contains(text(), '{0}')]/../../td[13]/div/a[contains(@class,\"itemDetailDeleteBtn\")]";
        private const string CONFIRM_DELETE = "dataConfirmOK";

        private const string COMMENT_FIRST_INGREDIENT_EDIT = "//table[@class='edit-zone']//a[starts-with(@href,'/Menus/Recipe/Comment')]/span";
        private const string COMMENT_FIRST_INGREDIENT_DISPLAY = "//table[normalize-space(@class)='display-zone']//a[starts-with(@href,'/Menus/Recipe/Comment')]/span";

        private const string PRICES_INGREDIENTS = "//table[2]/tbody/tr/td[8]";
        private const string COMMENT_AREA = "Comment";
        private const string VALIDATE_COMMENT = "//*[@id=\"item-details-comment\"]/div[2]/button[2]";

        // Bas tableau ingrédients
        private const string SALES_PRICE = "recipe_SalesPrice";
        private const string POINTS_PRICE = "recipe_PointsPrice";
        private const string TOTAL_PORTION = "recipe_WeightByPortion";
        private const string TOTAL_FOR_1_PORTION = "//*[@id=\"recipe-variant-detail-list\"]/div[3]/div/div[4]/p/span";
        private const string PRICE_PORTION = "//*[@id=\"recipe-variant-detail-list\"]/div[3]/div/div[8]/p/span[1]";
        private const string PRICE_FOR_1_PORTION = "//*[@id=\"recipe-variant-detail-list\"]/div[3]/div/div[8]/p/span[2]";
        private const string INGREDIENT_COLUMN = "//*/table[normalize-space(@class)='display-zone']//tr[@data-type = 'Item']/td/span[contains(text(), '{0}')]"; 
        private const string INGREDIENT_FIRST = "/html/body/div[3]/div/div[2]/div[3]/div[2]/div/div/div[1]/div[2]/ul/li[1]/div/div/div/form/table[2]/tbody/tr/td[1]/span";

        // Procédure
        private const string PROCEDURE = "Procedure";

        // Dietetic
        private const string RECIPE_DIETETIC_ENERGY_KCAL = "dietetic_EnergyKCal";
        private const string RECIPE_DIETETIC_COMPUTED_MANUAL_SWITCH = ".bootstrap-switch-id-dietetic_IsComputed";
        private const string RECIPE_DIETETIC_COMPUTED_MANUAL_INPUT = "dietetic_IsComputed";
        private const string RECIPE_DIETETIC_COMPUTED_MANUAL_CONFIRM = "dataConfirmOK";
        private const string RECIPE_DIETETIC_MANUAL_DATA = "manual-data";

        // Composition
        private const string COMPOSITION = "compo_Composition";
        
        // Picture
        private const string UPLOAD_PICTURE = "FileSent";
        private const string DELETE_PICTURE = "//*[@id=\"tabContentDetails\"]/div/div/div/div/form/a";
        private const string DELETE_PICTURE_LINK = "//*[@id=\"tabContentDetails\"]/div/div/form/a";
        private const string PICTURE = "//*[@id=\"tabContentDetails\"]/div/div/form/img";

        // Autre
        private const string CLOSE_WINDOW = "btn-from-datasheet-close-modal";
        private const string CLOSE_SUB_POP_UP = "//*[@id=\"modal-close-btn-sub\"]/span";


        private const string SUB_RECIPE_TO_DELETE = "//table[2]/tbody/tr/td[1]/span[contains(text(), '{0}')]";
        private const string SUB_RECIPE_DELETE_BUTTON_DEV = "//table[1]/tbody/tr/td[1]/span[contains(text(), '{0}')]/../..//span[contains(@class, 'fas fa-trash-alt')]";
        private const string SUB_RECIPE_DELETE_BUTTON_PATCH = "//table[1]/tbody/tr/td[1]/span[contains(text(), '{0}')]/../..//span[contains(@class, 'glyphicon glyphicon-trash')]";
        private const string VARIANTS = "//a[contains(@id, 'link-recipe-variant-site')]";
        private const string CONFIRM_DELETE_SUBRECIPE = "//*[@id=\"dataConfirmOK\"]";
        private const string RECIPE_IS_COMPUTED = ".bootstrap-switch-handle-on.bootstrap-switch-primary";
        private const string WORKSHOPS = "//select[@id=\"recipe_WorkshopId\"]/option[@selected]";
        private const string SEARCHCLICK = "//*[@id=\"ingredient-search-button-id\"]";
        private const string TOTAL_PRICE_ROWS = "/html/body/div[4]/div/div/div[2]/div[2]/div[2]/div/div/div[1]/div[2]/ul/li[*]/div/div/div/form/table[2]/tbody/tr/td[8]";
        private const string INGREDIENT_NAME_ITEM = "/html/body/div[3]/div/div[2]/div[3]/div[2]/div/div/div[1]/div[2]/ul/li/div/div/div/form/table[2]/tbody/tr/td[1]/span";
        private const string FIRST_INGREDIENT_ITEM = "/html/body/div[3]/div/div[2]/div[3]/div[2]/div/div/div[1]/div[2]/ul/li/div/div/div/form/table[1]/tbody/tr";
        private const string DELETE_INGREDIENT = "menus-recipevariantdetail-delete-display-1"; 
        private const string CLICKSECONDLINE = "//*[@id=\"menu-recipe-datasheet-datasheetrecipevariantdetail-item-2\"]";
        // ___________________________________ Variables ______________________________________________

        // Général
        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;
        [FindsBy(How = How.XPath, Using = CLICKSECONDLINE)]
        private IWebElement _clicksecondline;
        [FindsBy(How = How.XPath, Using = SEARCHCLICK)]
        private IWebElement _searchclick;

        [FindsBy(How = How.Id, Using = GENERAL_INFORMATION)]
        private IWebElement _generalInformation;

        [FindsBy(How = How.Id, Using = RECIPE_PRICE)]
        private IWebElement _recipePrice;

        [FindsBy(How = How.XPath, Using = RECIPE_NAME)]
        private IWebElement _recipeName;

        [FindsBy(How = How.XPath, Using = IMPORT)]
        private IWebElement _import;

        [FindsBy(How = How.XPath, Using = COPY_VALUES_FOR_SITE)]
        private IWebElement _copyValuesForAllSites;

        [FindsBy(How = How.XPath, Using = DUPLICATE_VARIANT)]
        private IWebElement _duplicateVariant;

        [FindsBy(How = How.XPath, Using = PRINT)]
        private IWebElement _print;

        [FindsBy(How = How.Id, Using = CONFIRM_PRINT)]
        private IWebElement _confirmPrint;
        
        [FindsBy(How = How.XPath, Using = ITEM)]
        private IWebElement _firstItem;


        // Onglets
        [FindsBy(How = How.Id, Using = ITEMS_TAB)]
        private IWebElement _itemsTab;

        [FindsBy(How = How.Id, Using = KEYWORD_TAB)]
        private IWebElement _keyWordTab;

        [FindsBy(How = How.Id, Using = PROCEDURE_TAB)]
        private IWebElement _procedureTab;

        [FindsBy(How = How.Id, Using = DIETETIC_TAB)]
        private IWebElement _dieteticTab;

        [FindsBy(How = How.Id, Using = INTOLERANCE_TAB)]
        private IWebElement _intoleranceTab;

        [FindsBy(How = How.Id, Using = COMPOSITION_TAB)]
        private IWebElement _compositionTab;

        [FindsBy(How = How.Id, Using = PICTURE_TAB)]
        private IWebElement _pictureTab;


        // Ajout ingrédient
        [FindsBy(How = How.Id, Using = ITEM_NAME)]
        private IWebElement _itemName;

        [FindsBy(How = How.Id, Using = ITEM_CHECK_BOX)]
        private IWebElement _itemOnlyCheckBox;

        [FindsBy(How = How.XPath, Using = FIRST_ITEM)]
        private IWebElement _firstItemName;

        [FindsBy(How = How.XPath, Using = ERROR_MESSAGE)]
        private IWebElement _errorMessage;


        // Tableau ingrédient
        [FindsBy(How = How.Id, Using = INGREDIENT)]
        private IWebElement _firstIngredient;

        [FindsBy(How = How.Id, Using = FIRST_INGREDIENT)]
        private IWebElement _clickOnFirstIngredient;

        [FindsBy(How = How.XPath, Using = SUB_RECIPE)]
        private IWebElement _subRecipe;

        [FindsBy(How = How.Id, Using = INGREDIENT_NET_WEIGHT)]
        private IWebElement _netWeight;

        [FindsBy(How = How.Id, Using = INGREDIENT_NET_QTY)]
        private IWebElement _netQty;

        [FindsBy(How = How.Id, Using = INGREDIENT_YIELD)]
        private IWebElement _yield;

        [FindsBy(How = How.Id, Using = INGREDIENT_GROSS_QTY)]
        private IWebElement _grossQty;

        [FindsBy(How = How.Id, Using = INGREDIENT_PRICE)]
        private IWebElement _price;

        [FindsBy(How = How.CssSelector, Using = INGREDIENT_WEIGHT)]
        private IWebElement _weight;

        [FindsBy(How = How.XPath, Using = EDIT_INGREDIENT)]
        private IWebElement _editIngredient;

        [FindsBy(How = How.XPath, Using = EDIT_FIRST_INGREDIENT_EDIT)]
        private IWebElement _editFirstIngredient;

        [FindsBy(How = How.XPath, Using = EDIT_FIRST_SUB_RECIPE)]
        private IWebElement _editFirstSubIngredient;

        [FindsBy(How = How.XPath, Using = EDIT_FIRST_SUB_RECIPE_DATASHEET)]
        private IWebElement _EditFirstSubRecipeDatasheet;

        [FindsBy(How = How.XPath, Using = DELETE_FIRST_INGREDIENT_EDIT)]
        private IWebElement _deleteFirstIngredient;

        [FindsBy(How = How.XPath, Using = CONFIRM_DELETE)]
        private IWebElement _confirmDelete;

        [FindsBy(How = How.XPath, Using = COMMENT_FIRST_INGREDIENT_EDIT)]
        private IWebElement _commentFirstIngredient;

        [FindsBy(How = How.XPath, Using = COMMENT_AREA)]
        private IWebElement _comment;

        [FindsBy(How = How.XPath, Using = VALIDATE_COMMENT)]
        private IWebElement _validateComment;

        //sub-recipe
        [FindsBy(How = How.XPath, Using = SUB_RECIPE_TO_DELETE)]
        private IWebElement _subRecipeToDelete;

        [FindsBy(How = How.XPath, Using = SUB_RECIPE_DELETE_BUTTON_DEV)]
        private IWebElement _subRecipeDeleteButtonDev;

        [FindsBy(How = How.XPath, Using = SUB_RECIPE_DELETE_BUTTON_PATCH)]
        private IWebElement _subRecipeDeleteButtonPatch;

        [FindsBy(How = How.XPath, Using = CONFIRM_DELETE_SUBRECIPE)]
        private IWebElement _confirmDelteSubrecipe;

        // Bas tableau ingrédients
        [FindsBy(How = How.Id, Using = SALES_PRICE)]
        private IWebElement _salesPrice;

        // Bas tableau ingrédients
        [FindsBy(How = How.Id, Using = POINTS_PRICE)]
        private IWebElement _pointsPrice;

        [FindsBy(How = How.Id, Using = TOTAL_PORTION)]
        private IWebElement _totalPortion;

        [FindsBy(How = How.XPath, Using = TOTAL_FOR_1_PORTION)]
        private IWebElement _totalFor1Portion;

        [FindsBy(How = How.XPath, Using = PRICE_PORTION)]
        private IWebElement _pricePortion;

        [FindsBy(How = How.XPath, Using = PRICE_FOR_1_PORTION)]
        private IWebElement _priceFor1Portion;

        // Procédure
        [FindsBy(How = How.Id, Using = PROCEDURE)]
        private IWebElement _procedure;

        // Dietetic
        [FindsBy(How = How.Id, Using = RECIPE_DIETETIC_ENERGY_KCAL)]
        private IWebElement _recipeDieteticEnergyKCal;

        [FindsBy(How = How.CssSelector, Using = RECIPE_DIETETIC_COMPUTED_MANUAL_SWITCH)]
        private IWebElement _recipeDieteticComputedManualSwitch;

        [FindsBy(How = How.Id, Using = RECIPE_DIETETIC_COMPUTED_MANUAL_INPUT)]
        private IWebElement _recipeDieteticComputedManualInput;

        [FindsBy(How = How.Id, Using = RECIPE_DIETETIC_COMPUTED_MANUAL_CONFIRM)]
        private IWebElement _recipeDieteticComputedManualConfirm;

        [FindsBy(How = How.Id, Using = RECIPE_DIETETIC_MANUAL_DATA)]
        private IWebElement _recipeDieteticManualData;

        // Composition
        [FindsBy(How = How.Id, Using = COMPOSITION)]
        private IWebElement _composition;

        // Picture
        [FindsBy(How = How.Id, Using = UPLOAD_PICTURE)]
        private IWebElement _uploadPicture;

        [FindsBy(How = How.XPath, Using = DELETE_PICTURE_LINK)]
        private IWebElement _deletePicture;

        // Autre
        [FindsBy(How = How.Id, Using = CLOSE_WINDOW)]
        private IWebElement _closeWindow;

        [FindsBy(How = How.Id, Using = RECIPE_IS_COMPUTED)]
        private IWebElement _IsComputed;

        [FindsBy(How = How.XPath, Using = IMPORTSUBRECIPE)]
        private IWebElement _importSubRecipe;

        [FindsBy(How = How.XPath, Using = WORKSHOPS)]
        private IWebElement _workshops;

        [FindsBy(How = How.XPath, Using = INGREDIENT_NAME_ITEM)]
        private IWebElement _ingredientNameItem;

        [FindsBy(How = How.XPath, Using = FIRST_INGREDIENT_ITEM)]
        private IWebElement _firstIngredientItem;

        [FindsBy(How = How.Id, Using = DELETE_INGREDIENT)]
        private IWebElement _deleteIngredient;
        
        // ___________________________________ Méthodes _______________________________________________

        public RecipesVariantPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // Général
        public RecipesPage BackToList()
        {
            _backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            _backToList.Click();

            WaitForLoad();

            return new RecipesPage(_webDriver, _testContext);
        }
        public void Searchclick()
        {
            _searchclick = WaitForElementIsVisible(By.XPath(SEARCHCLICK));
            _searchclick.Click();
            WaitForLoad();

         }

        public RecipeGeneralInformationPage ClickOnGeneralInformation()
        {
            _generalInformation = WaitForElementIsVisible(By.Id(GENERAL_INFORMATION));
            _generalInformation.Click();
            WaitForLoad();

            return new RecipeGeneralInformationPage(_webDriver, _testContext);
        }

        public RecipeGeneralInformationPage ClickOnGeneralInformationFromRecipeDatasheet()
        {
            int index = 10;
            while(index>0 && !isElementExists(By.Id(GENERAL_INFORMATION_TAB)))
            {
                WaitLoading();
                if (isElementExists(By.Id(GENERAL_INFORMATION_TAB)))
                {
                    break;
                }
                index--;
            }
            var _generalInformationTab = WaitForElementIsVisible(By.Id(GENERAL_INFORMATION_TAB));
            _generalInformationTab.Click();

            return new RecipeGeneralInformationPage(_webDriver, _testContext);
        }


        public double GetRecipePrice(string currency, string decimalSeparator)
        {
            WaitForLoad();
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            _recipePrice = WaitForElementIsVisibleNew(By.Id(RECIPE_PRICE));
            var analyseText = _recipePrice.Text;

            var format = (NumberFormatInfo)ci.NumberFormat.Clone();
            format.CurrencySymbol = currency;
            var mynumber = Decimal.Parse(analyseText, NumberStyles.Currency, format);

            return Convert.ToDouble(mynumber, ci);
        }

        public string GetRecipeName()
        {
            try
            {
                _recipeName = WaitForElementIsVisible(By.XPath(RECIPE_NAME));
                return _recipeName.Text;
            }
            catch
            {
                return "";
            }
        }

        public RecipeImportIngredientPage ImportIngredient()
        {
            _import = WaitForElementIsVisible(By.XPath(IMPORT));
            _import.Click();
            WaitForLoad();

            return new RecipeImportIngredientPage(_webDriver, _testContext);
        }

        public void CopyValuesForAllSites()
        {
            _copyValuesForAllSites = WaitForElementIsVisible(By.XPath(COPY_VALUES_FOR_SITE));
            _copyValuesForAllSites.Click();
            WaitForLoad();
        }

        public RecipeDuplicateVariantPage DuplicateVariant()
        {
            _duplicateVariant = WaitForElementIsVisible(By.XPath(DUPLICATE_VARIANT));
            _duplicateVariant.Click();
            WaitForLoad();

            return new RecipeDuplicateVariantPage(_webDriver, _testContext);
        }

        public PrintReportPage Print(bool versionPrint)
        {           
            _print = WaitForElementIsVisible(By.XPath(PRINT));
            _print.Click();
            WaitForLoad();

            _confirmPrint = WaitForElementIsVisible(By.Id(CONFIRM_PRINT));
            _confirmPrint.Click();
            WaitForLoad();

            if (versionPrint)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-pdf']"));
                ClickPrintButton();
            }

            //Results are opened in a new tab, switch the driver to the newly created one
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new PrintReportPage(_webDriver, _testContext);
        }

        // Onglets
        public void ClickOnItemTab()
        {
            _itemsTab = WaitForElementIsVisible(By.Id(ITEMS_TAB));
            _itemsTab.Click();
            WaitForLoad();
        }

        public void ClickOnProcedureTab()
        {
            _procedureTab = WaitForElementIsVisible(By.Id(PROCEDURE_TAB));
            _procedureTab.Click();
            WaitForLoad();
        }

        public void ClickOnDieteticTab()
        {
            _dieteticTab = WaitForElementIsVisible(By.Id(DIETETIC_TAB));
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            javaScriptExecutor.ExecuteScript("window.scrollTo(0,0)");
            _dieteticTab.Click();
            WaitForLoad();
        }

        public RecipeIntolerancePage ClickOnIntoleranceTab()
        {
            try
            {
                _intoleranceTab = WaitForElementIsVisible(By.Id(INTOLERANCE_TAB));
                _intoleranceTab.Click();
            }
            catch
            {
                _intoleranceTab = WaitForElementIsVisible(By.Id("hrefTabContentVariantDetailsIntolerance-sub"));
                _intoleranceTab.Click();
            }
            WaitForLoad();

            return new RecipeIntolerancePage(_webDriver, _testContext);
        }

        public void ClickOnCompositionTab()
        {
            _compositionTab = WaitForElementIsVisible(By.Id(COMPOSITION_TAB));
            _compositionTab.Click();
            WaitForLoad();
        }

        public void ClickOnPictureTab()
        {
            _pictureTab = WaitForElementIsVisible(By.Id(PICTURE_TAB));
            _pictureTab.Click();
            WaitForLoad();
        }

        public RecipeKeywordPage ClickOnKeyWordTab()
        {
            _keyWordTab = WaitForElementIsVisible(By.Id(KEYWORD_TAB));
            _keyWordTab.Click();
            WaitForLoad();
            return new RecipeKeywordPage(_webDriver, _testContext);

        }



        public bool CheckKeywordDuplicate(string keyword)
        {
            var elements = _webDriver.FindElements(By.XPath(LIST_KEYWORD));
            int count = 0;
            foreach (var element in elements)
            {
                if (element.Text.Contains(keyword))
                {
                    count++;
                }

                if (count > 1)
                {
                    return false;
                }
            }

            return true;
        }


        // Ajout Ingrédient
        public void AddIngredient(string ingredient)
        {
            if (IsDev())
            {
                _itemOnlyCheckBox = WaitForElementExists(By.Id(ITEM_CHECK_BOX));
                _itemOnlyCheckBox.SetValue(ControlType.CheckBox, true);

                _itemName = WaitForElementIsVisibleNew(By.Id(ITEM_NAME));
                _itemName.SetValue(ControlType.TextBox, ingredient);
                WaitForLoad();
                WaitPageLoading();
                _firstItem = WaitForElementIsVisibleNew(By.XPath(FIRST_ITEM));
                _firstItem.Click();
                WaitForLoad();
            }
            else
            {
                _itemOnlyCheckBox = WaitForElementExists(By.Id(ITEM_CHECK_BOX));
            _itemOnlyCheckBox.SetValue(ControlType.CheckBox, true);

            _itemName = WaitForElementIsVisible(By.Id(ITEM_NAME));
            _itemName.SetValue(ControlType.TextBox, ingredient);
            WaitForLoad();
             _firstItem = WaitForElementIsVisible(By.XPath(FIRST_ITEM));
             _firstItem.Click();
            WaitForLoad();
            }
        }
        public void AjouterIngredient(string ingredient)
        {
 
            _itemName = WaitForElementIsVisible(By.Id(ITEM_NAME));
            _itemName.SetValue(ControlType.TextBox, ingredient);
            WaitPageLoading();
            WaitForLoad();
            Searchclick();
            _firstItem = WaitForElementIsVisible(By.XPath(FIRST_ITEM));
            _firstItem.Click();
            WaitPageLoading();
            WaitForLoad();
        }


        public bool AddIngredientError(string ingredient)
        {
            bool isAdded = true;

            _itemOnlyCheckBox = WaitForElementExists(By.Id(ITEM_CHECK_BOX));
            _itemOnlyCheckBox.SetValue(ControlType.CheckBox, true);
            WaitForLoad();

            _itemName = WaitForElementIsVisible(By.Id(ITEM_NAME));
            _itemName.SetValue(ControlType.TextBox, ingredient);
            WaitPageLoading();
            WaitForLoad();
            try
            {
                _errorMessage = WaitForElementIsVisible(By.XPath(ERROR_MESSAGE));
                if (_errorMessage.Text == "no ingredient found")
                    isAdded = false;
            }
            catch
            {
                isAdded = true;
            }

            return isAdded;
        }

        public void AddSubRecipe(string recipeName)
        {
            WaitPageLoading();
            _itemOnlyCheckBox = WaitForElementExists(By.Id(ITEM_CHECK_BOX));
            _itemOnlyCheckBox.SetValue(ControlType.CheckBox, false);
            WaitForLoad();

            _itemName = WaitForElementIsVisible(By.Id(ITEM_NAME));
            _itemName.SetValue(ControlType.TextBox, recipeName);
            WaitForLoad();

            _firstItem = WaitForElementIsVisible(By.XPath(FIRST_ITEM));
            _firstItem.Click();
            WaitPageLoading();
            WaitForLoad();
        }

        public void AddSubRecipeDatasheet(string recipeName)
        {
            var _itemOnly = WaitForElementIsVisible(By.Id("SelectedSearchOption"));
            _itemOnly.SetValue(ControlType.DropDownList, "Recipes");
            WaitForLoad();
            _itemName = WaitForElementIsVisible(By.Id(ITEM_NAME));
            _itemName.SetValue(ControlType.TextBox, recipeName);
            WaitForLoad();

            var boutonSearch = WaitForElementIsVisible(By.Id("ingredient-search-button-id"));
            boutonSearch.Click();
            WaitPageLoading();
            WaitForLoad();

            _firstItem = WaitForElementIsVisible(By.XPath(FIRST_ITEM));
            _firstItem.Click();
            WaitPageLoading();
            WaitForLoad();
        }

        public bool AddSubRecipeError(string recipeName)
        {
            WaitPageLoading();
            _itemOnlyCheckBox = WaitForElementExists(By.Id(ITEM_CHECK_BOX));
            _itemOnlyCheckBox.SetValue(ControlType.CheckBox, false);

            _itemName = WaitForElementIsVisible(By.Id(ITEM_NAME));
            _itemName.SetValue(ControlType.TextBox, recipeName);
            WaitPageLoading();

            if (isElementVisible(By.XPath(ERROR_MESSAGE)))
            {
                _errorMessage = WaitForElementIsVisible(By.XPath(ERROR_MESSAGE));
                if (_errorMessage.Text == "no ingredient found")
                {
                    return false;
                }
            }
            return true;

        }

        // Tableau ingrédients
        public void HighlightFirstIngredient()
        {
            Actions action = new Actions(_webDriver);

            if (IsDev())
            {
                _firstIngredient = WaitForElementIsVisible(By.Id(INGREDIENT));
                action.MoveToElement(_firstIngredient).Perform();
            }
            else
            {
                _firstIngredient = WaitForElementIsVisible(By.XPath("//*[@id='menu-recipe-recipe-variant-details-display-zone']/tbody/tr/td[1]/span"));
                action.MoveToElement(_firstIngredient).Perform();
            }
        }
        public void HighlightIngredient(string item)
        {
            Actions action = new Actions(_webDriver);
            IWebElement ingredient;
            try
            {
                ingredient = WaitForElementIsVisible(By.XPath(string.Format(INGREDIENT_COLUMN, item)));
            }
            catch
            {
                ingredient = WaitForElementIsVisible(By.XPath(string.Format("//*/table[@class='edit-zone']//tr[@data-type = 'Item']/td/span[contains(text(), '{0}')]", item)));
            }
            action.MoveToElement(ingredient).Perform();
        }
        public void HighlightFirstSubRecipe()
        {
            Actions action = new Actions(_webDriver);
          
                WaitPageLoading();
                _subRecipe = WaitForElementIsVisible(By.Id(SUB_RECIPE_PATCH));
                action.MoveToElement(_subRecipe).Perform();
          
        }

        public void HighlightFirstSubRecipeDatasheet()
        {
            Actions action = new Actions(_webDriver);
                _subRecipe = WaitForElementIsVisible(By.Id(SUB_RECIPE));
                action.MoveToElement(_subRecipe).Perform();
        }

        public void ClickFirstSubRecipeDatasheet()
        {
            _subRecipe = WaitForElementIsVisible(By.Id(SUB_RECIPE));
            _subRecipe.Click();
        }
        public void ClickOnFirstIngredient()
        {
         
            if (IsDev())
            {
                var elementId = FIRST_INGREDIENT_IN_DEV;
                _clickOnFirstIngredient = WaitForElementIsVisible(By.XPath(elementId));
            }
            else
            {
                var elementId = FIRST_INGREDIENT;
                _clickOnFirstIngredient = WaitForElementIsVisible(By.XPath(elementId));
            }
            WaitPageLoading();
              _clickOnFirstIngredient.Click();

            WaitForLoad();
        }


        public void ClickOnFirstIngredientDatasheet()
        {
            WaitForLoad();
            _clickOnFirstIngredient = WaitForElementIsVisible(By.XPath(FIRST_INGREDIENT_IN_DEV));
            _clickOnFirstIngredient.Click();
            WaitForLoad();
        }

        public string GetFirstIngredient()
        {
            var getFirstIngredient = WaitForElementIsVisible(By.XPath(INGREDIENT_FIRST));
            return getFirstIngredient.Text;
        }
        public double GetNetWeight(string decimalSeparator)
        {
            WaitPageLoading();
            WaitForLoad();
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            var elements = _webDriver.FindElements(By.Id(INGREDIENT_NET_WEIGHT));

            foreach (var elm in elements)
            {
                if (elm.Displayed)
                {
                    _netWeight = elm;
                    break; 
                }
            }
            string weight = _netWeight.GetAttribute("value");
            return Convert.ToDouble(weight, ci);

        }

        public void SetNetWeight(string value)
        {
            WaitPageLoading();
            _netWeight = WaitForElementIsVisible(By.Id(INGREDIENT_NET_WEIGHT));
            _netWeight.SetValue(ControlType.TextBox, value);
            WaitPageLoading();
            WaitForLoad() ;
        }

        public double GetNetQuantity(string decimalSeparator)
        {
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            _netQty = WaitForElementIsVisible(By.Id(INGREDIENT_NET_QTY));
            var qty = _netQty.GetAttribute("value");

            return Convert.ToDouble(qty, ci);
        }

        public void SetNetQuantity(string value)
        {
            _netQty = WaitForElementIsVisible(By.Id(INGREDIENT_NET_QTY));
            _netQty.SetValue(ControlType.TextBox, value);

            WaitPageLoading();
        }

        public double GetYield(string decimalSeparator)
        {
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            _yield = WaitForElementIsVisible(By.Id(INGREDIENT_YIELD));
            var yield = _yield.GetAttribute("value");

            return Convert.ToDouble(yield, ci);
        }

        public void SetYield(string value)
        {
            WaitPageLoading();
            _yield = WaitForElementIsVisible(By.Id("recipe_Yield"));
            _yield.Click();
            var _yieldInput = WaitForElementIsVisible(By.XPath("//*[@id=\"recipe_Yield\"]"));
            _yieldInput.SetValue(ControlType.TextBox, value);

            WaitPageLoading();
        }

        public double GetGrossQty(string decimalSeparator)
        {
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            _grossQty = WaitForElementExists(By.Id(INGREDIENT_GROSS_QTY));
            var grossQty = _grossQty.GetAttribute("value");

            return Convert.ToDouble(grossQty, ci);
        }

        public void SetGrossQty(string value)
        {
            _grossQty = WaitForElementIsVisible(By.Id(INGREDIENT_GROSS_QTY));
            _grossQty.SetValue(ControlType.TextBox, value);

            WaitPageLoading();
        }

        public string GetPrice()
        {
            _price = WaitForElementExists(By.Id(INGREDIENT_PRICE));
            return _price.Text;
        }

        public double GetPrice(string decimalSeparator, string currency)
        {
            WaitPageLoading();
            WaitForLoad();
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");
            _price = WaitForElementExists(By.XPath("//*/table[@class='display-zone']/tbody/tr/td[8]"));
            
            var price = _price.GetAttribute("innerText").Replace(currency, "").Trim();
            price = price.Replace(" ", "").Trim();
            return Convert.ToDouble(price, ci);
        }


        public double GetPriceSouRecette(string decimalSeparator, string currency)
        {
            WaitPageLoading();
            WaitForLoad();
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");
            _price = WaitForElementExists(By.XPath("//*[@id=\"menus-datasheet-datasheetrecipevariantdetail-recipe-1\"]/td[8]"));

            var price = _price.GetAttribute("innerText").Replace(currency, "").Trim();
            price = price.Replace(" ", "").Trim();
            return Convert.ToDouble(price, ci);
        }

        public double GetTotalPriceFromAllRows(string decimalSeparator, string currency)
        {
            WaitPageLoading();
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            var elements = _webDriver.FindElements(By.XPath(TOTAL_PRICE_ROWS));

            double totalSum = 0;

            foreach (var elm in elements)
            {
                var price = elm.GetAttribute("innerText").Replace(currency, "").Trim();
                price = price.Replace(" ", "").Trim();
                var finalPrice = Convert.ToDouble(price, ci);

                totalSum += finalPrice;
            }

            return totalSum;
        }

        public double GetModalPrice(string decimalSeparator, string currency)
        {
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");
            _price = WaitForElementIsVisible(By.XPath(INGREDIENT_PRICEMODAL));
            var price = _price.Text.Replace(currency, "").Trim();
            price = price.Replace(" ", "").Trim();
            return Convert.ToDouble(price, ci);
        }

        public double GetWeight(string decimalSeparator)
        {
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            _weight = WaitForElementIsVisible(By.CssSelector(INGREDIENT_WEIGHT));
            var weight = _weight.Text;

            return Convert.ToDouble(weight, ci);
        }

        public ItemGeneralInformationPage EditFirstIngredient(bool edit)
        {
            HighlightFirstIngredient();

            if (edit)
            {
                _editFirstIngredient = WaitForElementIsVisible(By.XPath(EDIT_FIRST_INGREDIENT_EDIT));
            }
            else
            {
                _editFirstIngredient = WaitForElementIsVisible(By.XPath(EDIT_FIRST_INGREDIENT_DISPLAY));
            }
            _editFirstIngredient.Click();
            WaitForLoad();

            //Results are opened in a new tab, switch the driver to the newly created one
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(30));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            WaitForLoad();

            return new ItemGeneralInformationPage(_webDriver, _testContext);
        }
        public ItemGeneralInformationPage EditIngredient(string itemName)
        {
            WaitForLoad();

            HighlightIngredient(itemName);

            try
            {
                _editIngredient = WaitForElementIsVisible(By.XPath(string.Format(EDIT_INGREDIENT, itemName)));
            }
            catch
            {
                _editIngredient = WaitForElementIsVisible(By.XPath(string.Format("//table[@class='edit-zone']/tbody/tr/td[1][@title=\"{0}\"]/..//a[starts-with(@href,'/Purchasing/Item/Detail')]/span", itemName)));
            }
            _editIngredient.Click();
            WaitForLoad();

            //Results are opened in a new tab, switch the driver to the newly created one
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(30));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            if (_webDriver.WindowHandles.Count > 2)
            {
                _webDriver.SwitchTo().Window(_webDriver.WindowHandles[2]);
            }
            else
            {
                _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);
            }

            WaitForLoad();

            return new ItemGeneralInformationPage(_webDriver, _testContext);
        }
        public RecipesVariantPage EditFirstSubRecipe(bool edit)
        {
            HighlightFirstSubRecipe();
            if (edit)
            {
                _editFirstIngredient = WaitForElementIsVisible(By.XPath(EDIT_FIRST_SUB_RECIPE_EDIT));
            }
            else
            {

                _editFirstSubIngredient = WaitForElementIsVisible(By.Id(EDIT_FIRST_SUB_RECIPE));
             
            }

            _editFirstSubIngredient.Click();
            WaitForLoad();

            //Results are opened in a new tab, switch the driver to the newly created one
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(30));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            WaitForLoad();

            return new RecipesVariantPage(_webDriver, _testContext);
        }

        public RecipesVariantPage EditFirstSubRecipeDatasheet()
        {
            HighlightFirstSubRecipeDatasheet();
            _EditFirstSubRecipeDatasheet = WaitForElementIsVisible(By.XPath(EDIT_FIRST_SUB_RECIPE_DATASHEET));
            _EditFirstSubRecipeDatasheet.Click();
            WaitForLoad();
            return new RecipesVariantPage(_webDriver, _testContext);
        }

        public void DeleteFirstIngredient(bool edit)
        {
            HighlightFirstIngredient();
            if (edit)
            {
                _deleteFirstIngredient = WaitForElementIsVisible(By.XPath(DELETE_FIRST_INGREDIENT_EDIT));
            }
            else
            {
                _deleteFirstIngredient = WaitForElementIsVisible(By.XPath(DELETE_FIRST_INGREDIENT_DISPLAY));
            }
            _deleteFirstIngredient.Click();
            WaitForLoad();

            _confirmDelete = WaitForElementIsVisible(By.Id(CONFIRM_DELETE));
            _confirmDelete.Click();
            WaitForLoad();
            WaitPageLoading();
        }
        public void DeleteIngredient(string item)
        {
            HighlightIngredient(item);
            var deletetIngredient = WaitForElementExists(By.XPath(string.Format(DELETE_ING, item)));
            Actions a = new Actions(_webDriver);
            a.MoveToElement(deletetIngredient).Perform();
           
            deletetIngredient.Click();
            WaitForLoad();

            _confirmDelete = WaitForElementIsVisible(By.Id(CONFIRM_DELETE));
            _confirmDelete.Click();
            WaitForLoad();
            WaitPageLoading();
        }

        public void AddCommentToIngredient(string comment, bool edit)
        {
            HighlightFirstIngredient();
            if (edit)
            {
                _commentFirstIngredient = WaitForElementIsVisible(By.XPath(COMMENT_FIRST_INGREDIENT_EDIT));
            }
            else
            {
                _commentFirstIngredient = WaitForElementIsVisible(By.XPath(COMMENT_FIRST_INGREDIENT_DISPLAY));
            }
            _commentFirstIngredient.Click();
            WaitForLoad();

            _comment = WaitForElementIsVisible(By.Id(COMMENT_AREA));
            _comment.SetValue(ControlType.TextBox, comment);

            _validateComment = WaitForElementToBeClickable(By.XPath(VALIDATE_COMMENT));
            _validateComment.Click();
            WaitForLoad();
        }

        public string GetComment(bool edit)
        {
            HighlightFirstIngredient();

            if (edit)
            {
                _commentFirstIngredient = WaitForElementIsVisible(By.XPath(COMMENT_FIRST_INGREDIENT_EDIT));
            }
            else
            {
                _commentFirstIngredient = WaitForElementIsVisible(By.XPath(COMMENT_FIRST_INGREDIENT_DISPLAY));
            }
            _commentFirstIngredient.Click();
            WaitForLoad();

            _comment = WaitForElementIsVisible(By.Id(COMMENT_AREA));
            return _comment.Text;
        }

        public void CloseCommentPopup()
        {
            var cancel = WaitForElementIsVisible(By.XPath("//button[contains(text(), 'Cancel')]"));
            cancel.Click();
            WaitForLoad();
        }


        public void CloseSubRecipePopUp()
        {
            var cancel = WaitForElementIsVisible(By.XPath(CLOSE_SUB_POP_UP));
            cancel.Click();
            WaitForLoad();
        }
        public bool IsIngredientDisplayed()
        {
            bool isDisplayed;
            WaitForLoad();
            if (IsDev())
            {
                if (isElementVisible(By.Id(INGREDIENT)))
                {
                    _webDriver.FindElement(By.Id(INGREDIENT));
                    isDisplayed = true;
                }
                else
                {
                    isDisplayed = false;
                }
            }
            else
            {
                if (isElementVisible(By.XPath("//*/table[*][contains(@class,'display-zone')]/tbody/tr/td[1]/span")))
                {
                    _webDriver.FindElement(By.XPath("//*/table[*][contains(@class,'display-zone')]/tbody/tr/td[1]/span"));
                    isDisplayed = true;
                }
                else
                {
                    isDisplayed = false;
                }
            }

            return isDisplayed;
        }

        public int GetNumberOfIngredient()
        {
            WaitPageLoading();
            if (IsDev())
            {
                return _webDriver.FindElements(By.XPath(INGREDIENTLIGNES)).Count;
            }
            else
            {
                return _webDriver.FindElements(By.XPath("//*/table[*][contains(@class,'display-zone')]/tbody/tr/td[1]/span")).Count;
            }
        }

        public int CalculIngredientLignes()
        {
            if (IsDev())
            {
                return _webDriver.FindElements(By.XPath(INGREDIENTLIGNES)).Count;
            }
            else
            {
                return _webDriver.FindElements(By.XPath("//*/table[*][contains(@class,'display-zone')]/tbody/tr/td[1]/span")).Count;
            }
        }

        public string GetIngredientName()
        {
            return _webDriver.FindElement(By.XPath(INGREDIENT_NAME)).Text;
        }

        public List<string> GetIngredients()
        {
            List<string> listIngredients = new List<string>();
            var ingredients = _webDriver.FindElements(By.XPath(INGREDIENT_NAME));

            foreach (var ingredient in ingredients)
            {
                listIngredients.Add(ingredient.Text.Trim());
            }

            return listIngredients;
        }

        public bool IsSubRecipeDisplayed()
        {
            bool isDisplayed;
            WaitPageLoading();
            try
            {
                if (IsDev())
                {
                    _webDriver.FindElement(By.XPath(ALL_SUB_RECIPE));
                }
                else
                {
                    if (SolveVisible("//*/tr[@data-type='Recipe']") == null) {
                        return false;
                    }
                    else
                    {
                        return true;
                    }
                }
                isDisplayed = true;

            }
            catch
            {
                isDisplayed = false;
            }

            return isDisplayed;
        }

        // Bas tableau ingrédients
        public double GetSalesPrice(string decimalSeparator)
        {
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            _salesPrice = WaitForElementIsVisible(By.Id(SALES_PRICE));
            var sales = _salesPrice.GetAttribute("value");

            return Convert.ToDouble(sales, ci);
        }

        public void SetSalesPrice(string value)
        {
            _salesPrice = WaitForElementIsVisible(By.Id(SALES_PRICE));
            _salesPrice.SetValue(ControlType.TextBox, value);

            WaitPageLoading();
        }

        public void SetPointsPrice(string vakue)
        {
            _pointsPrice = WaitForElementIsVisible(By.Id(POINTS_PRICE));
            _pointsPrice.SetValue(ControlType.TextBox, vakue);

            WaitPageLoading();
        }

        public double GetTotalPortion(string decimalSeparator)
        {
            WaitPageLoading();
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            _totalPortion = WaitForElementIsVisible(By.Id(TOTAL_PORTION));
            var total = _totalPortion.GetAttribute("value");

            return Convert.ToDouble(total, ci);
        }

        public void SetTotalPortion(string value)
        {
            _totalPortion = WaitForElementIsVisible(By.Id(TOTAL_PORTION));
            _totalPortion.SetValue(ControlType.TextBox, value);

            WaitPageLoading();
        }

        public string GetTotalFor1Portion()
        {
            _totalFor1Portion = WaitForElementIsVisible(By.XPath(TOTAL_FOR_1_PORTION));
            return _totalFor1Portion.Text;
        }

        public string GetPricePortion()
        {
            WaitPageLoading();
            _pricePortion = WaitForElementIsVisible(By.XPath("//*[@id='recipe-variant-detail-list']/div[3]/div/div[5]/div/div[2]/div[2]/p/span[1]"));
            return _pricePortion.Text;
        }

        public string GetPriceFor1Portion()
        {
            _priceFor1Portion = WaitForElementIsVisible(By.XPath("//*[@id='recipe-variant-detail-list']/div[3]/div/div[5]/div/div[2]/div[2]/p/span[2]"));
            return _priceFor1Portion.Text;
        }

        // Onglet Dietetic
        public void SetEnergyKCal(string value)
        {
            _recipeDieteticEnergyKCal = WaitForElementIsVisible(By.Id(RECIPE_DIETETIC_ENERGY_KCAL));
            _recipeDieteticEnergyKCal.SetValue(ControlType.TextBox, value);

            // Prise en compte des modifs
            WaitPageLoading();
        }

        public string GetEnergyKCal()
        {
            _recipeDieteticEnergyKCal = WaitForElementIsVisible(By.Id(RECIPE_DIETETIC_ENERGY_KCAL));
            return _recipeDieteticEnergyKCal.Text;
        }

        // Onglet Procédure
        public void SetProcedure(string value)
        {
            _procedure = WaitForElementIsVisible(By.Id(PROCEDURE));
            _procedure.SetValue(ControlType.TextBox, value);

            // Prise en compte des modifs
            WaitPageLoading();
        }

        public string GetProcedure()
        {
            _procedure = WaitForElementIsVisible(By.Id(PROCEDURE));
            return _procedure.Text;
        }

        // Onglet Composition
        public void SetComposition(string value)
        {
            _composition = WaitForElementIsVisible(By.Id(COMPOSITION));
            _composition.SetValue(ControlType.TextBox, value);

            // Prise en compte des modifs
            WaitPageLoading();
        }

        public string GetComposition()
        {
            _composition = WaitForElementIsVisible(By.Id(COMPOSITION));
            return _composition.Text;
        }

        // Onglet picture
        public void AddPicture(string imagePath)
        {
            _uploadPicture = WaitForElementIsVisible(By.Id(UPLOAD_PICTURE));
            _uploadPicture.SendKeys(imagePath);
            WaitForLoad();

            WaitForElementIsVisible(By.XPath(DELETE_PICTURE));
        }

        public bool IsPictureAdded()
        {
            try
            {
                WaitForElementIsVisible(By.XPath(PICTURE));
                return true;
            }
            catch
            {
                return false;
            }
        }

        public void DeletePicture()
        {
            _deletePicture = WaitForElementExists(By.XPath(DELETE_PICTURE_LINK));
            _deletePicture.Click();
            WaitForLoad();

            _confirmDelete = WaitForElementIsVisible(By.Id(CONFIRM_DELETE));
            _confirmDelete.Click();
            WaitForLoad();
        }

        public string GetUrlPicture()
        {
            var picture = WaitForElementIsVisible(By.XPath(PICTURE));
            return picture.GetAttribute("src");
        }

        // Autre
        public DatasheetDetailsPage CloseWindow()
        {
            WaitPageLoading();
            _closeWindow = WaitForElementIsVisible(By.Id(CLOSE_WINDOW));
            _closeWindow.Click();
            WaitPageLoading();

            return new DatasheetDetailsPage(_webDriver, _testContext);
        }
        
        public string GetFirstItemName()
        {
            _firstItemName = WaitForElementIsVisible(By.Id(ITEM));
            return _firstItemName.Text.Trim();
        }

        public string GetFirstItemNameSousRecipe()
        {
            var firstItemName = WaitForElementIsVisible(By.XPath(FIRST_ITEM_NAME));
            return firstItemName.Text.Trim();
        }

        public void DeleteSubRecipe(string recipeName)
        {
            _subRecipeToDelete = WaitForElementExists(By.XPath(string.Format(SUB_RECIPE_TO_DELETE, recipeName)));
            Actions action = new Actions(_webDriver);
            action.MoveToElement(_subRecipeToDelete).Click().Perform();

            _subRecipeDeleteButtonDev = WaitForElementIsVisible(By.XPath(string.Format(SUB_RECIPE_DELETE_BUTTON_DEV, recipeName)));
            _subRecipeDeleteButtonDev.Click();

            _confirmDelteSubrecipe = WaitForElementIsVisible(By.XPath(CONFIRM_DELETE_SUBRECIPE));
            _confirmDelteSubrecipe.Click();
        }

        public void CliqueSurVariantSite(string site)
        {
            var listVariants = _webDriver.FindElements(By.XPath(VARIANTS));
            foreach (var variant in listVariants)
            {
                if(variant.Text.Contains(site))
                {
                    variant.Click();
                }
            }
        }
        public bool verifyPricesIngredients(string price)
        {
            WaitPageLoading();
            IEnumerable<IWebElement> pricesIngredient;
            pricesIngredient = _webDriver.FindElements(By.XPath("//table[2]/tbody/tr/td[8]"));
            
            foreach(var priceIngredient in pricesIngredient)
            {
                // euros + espace
                if(priceIngredient.Text.Substring(2) != price)
                {
                    return false;
                }
            }
            return true;
        }

        public bool verifyPricesIngredient(double price, string currency)
        {
            WaitPageLoading();
            CultureInfo ci = price.ToString().Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");
            var pricesIngredient = WaitForElementIsVisible(By.XPath("//table[2]/tbody/tr/td[8]"));
            var prices = pricesIngredient.Text.Replace(currency, "").Trim();
            prices = prices.Replace(" ", "").Trim();
            // euros + espace
            if (Convert.ToDouble(price, ci) != price)
            {
                return false;
            }

            return true;
        }
        public bool verifyNetWeightIngredients(string netWeight)
        {
            WaitPageLoading();
            string decimalSeparator = ",";
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");
            var netWeightsIngredient = _webDriver.FindElements(By.Id("recipe_NetWeight"));
            foreach (var netWeightIngredient in netWeightsIngredient)
            {
                var weight = netWeightIngredient.GetAttribute("value");
                if (Convert.ToDouble(weight, ci).ToString() != netWeight)
                {
                    return false;
                }
            }
            return true;
        }
        public bool verifyYieldIngredients(string yield)
        {
            WaitPageLoading();
            var netYieldsIngredients = _webDriver.FindElements(By.XPath("//table[2]/tbody/tr/td[6]/span[1]"));
            foreach (var netYieldsIngredient in netYieldsIngredients)
            {
                var weight = netYieldsIngredient.Text;
                if (weight != yield)
                {
                    return false;
                }
            }
            return true;
        }
        public bool verifyIsExistIngredients(string name)
        {
            WaitPageLoading();
            var ingredients = _webDriver.FindElements(By.XPath("//table[2]//tr[@data-type=\"Item\"]/td[1]/span"));
            List<string> list = new List<string>();
            foreach (var ingredient in ingredients)
            {
                list.Add(ingredient.Text);
            }

            if (!list.Contains(name))
            {
                return false;
            }
            return true;
        }
        
        public bool IsDieteticComputed()
        {
            _recipeDieteticComputedManualInput = WaitForElementExists(By.Id(RECIPE_DIETETIC_COMPUTED_MANUAL_INPUT));
            return _recipeDieteticComputedManualInput.Selected;
        }

        public bool IsMalualDataVisible()
        {
            _recipeDieteticManualData = WaitForElementExists(By.Id(RECIPE_DIETETIC_MANUAL_DATA));
            return _recipeDieteticManualData.Displayed;
        }

        public void ChangeDieteticComputedManual()
        {
            _recipeDieteticComputedManualSwitch = WaitForElementIsVisible(By.CssSelector(RECIPE_DIETETIC_COMPUTED_MANUAL_SWITCH));
            _recipeDieteticComputedManualSwitch.Click();

            _recipeDieteticComputedManualConfirm = WaitForElementIsVisible(By.Id(RECIPE_DIETETIC_COMPUTED_MANUAL_CONFIRM));
            _recipeDieteticComputedManualConfirm.Click();
            WaitForLoad();
        }
        public RecipeImportIngredientPage ImportIngredientSubRecipe()
        {
            _importSubRecipe = WaitForElementIsVisible(By.XPath(IMPORTSUBRECIPE));
            _importSubRecipe.Click();
            WaitForLoad();

            return new RecipeImportIngredientPage(_webDriver, _testContext);
        }
        public void SetYieldData(string value)
        {
            _yield = WaitForElementIsVisible(By.XPath(INGREDIENT_YIELD_DATA));
            _yield.Click();
            var _yieldInput = WaitForElementIsVisible(By.XPath("//*[@id=\"recipe_Yield\"]"));
            _yieldInput.SetValue(ControlType.TextBox, value);

            WaitPageLoading();
        }
        public List<string> GetWorkShops()
        {
            List<string> workShopList = new List<string>();
            var workShops = _webDriver.FindElements(By.XPath(WORKSHOPS));

            foreach (var ingredient in workShops)
            {
                workShopList.Add(ingredient.GetAttribute("text").Trim());
            }

            return workShopList;
        }
        public void Go_To_New_Navigate(int ongletChrome = 1)
        {
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > ongletChrome);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[ongletChrome]);
        }

        public string AjouterIngredient2(string ingredient)
        {

            _itemName = WaitForElementIsVisible(By.Id(ITEM_NAME));
            _itemName.SetValue(ControlType.TextBox, ingredient);
            WaitPageLoading();
            WaitForLoad();
            Searchclick();
            _firstItem = WaitForElementIsVisible(By.XPath(FIRST_ITEM));
            var text = _firstItem.Text;
            _firstItem.Click();
            WaitPageLoading();
            WaitForLoad();

            
            string[] parts = text.Split('\r');
            string result = parts[0];
            return result;
        }

        public bool VerifiedSameItemComposition(string newComposition)
        {
            WaitForLoad();
            _ingredientNameItem = WaitForElementIsVisible(By.XPath(INGREDIENT_NAME_ITEM));
            WaitForLoad();

            if (_ingredientNameItem.Text == newComposition)
            {
                return true;
            }
            else return false;

        }

        public void ClickOnDeleteIngredient()
        {            
            Actions actions = new Actions(_webDriver);
            IWebElement rowElement = WaitForElementIsVisible(By.Id(INGREDIENT));
            actions.MoveToElement(rowElement).Perform();
            if (isElementExists(By.Id(DELETE_INGREDIENT)))
            {
                var DeleteBtn = WaitForElementIsVisible(By.Id(DELETE_INGREDIENT));
                DeleteBtn.Click();

                var confirmBtn = WaitForElementIsVisible(By.Id(CONFIRM_DELETE));
                confirmBtn.Click();

                WaitForLoad();
            }


        }

        public bool VerifiedDeleteIngredient(string newComposition)
        {
            WaitForLoad();
            if (isElementExists(By.Id(COMPOSITION)))
            {
                _composition = WaitForElementIsVisible(By.Id(COMPOSITION));

                WaitForLoad();
                var compositionText = _composition.Text;
                if (!_composition.Text.Contains(newComposition))
                {
                    return true;
                }
                else return false;
            }


            return false;


        }
        public void ClickFirstLine()
        {
            WaitForLoad();
            var firstline = WaitForElementIsVisible(By.Id(FIRST_lINE));
            firstline.Click();
            WaitForLoad();
        }


        public void EditNetWeight(string netWeight)
        {
            ClickFirstLine();
            WaitForLoad(); 
            var netWeightInput = WaitForElementIsVisible(By.Id(NET_WEIGHT));
            netWeightInput.Clear();
            netWeightInput.SetValue(ControlType.TextBox, netWeight);
            WaitPageLoading();
        }
         public double GetGrossQuantity(string decimalSeparator)
        {
            var ingredients = WaitForElementExists(By.XPath("//*[@id='recipe-variant-detail-list']/div[2]/ul"));
            double sommeGrossQuantity = 0;
            var ulElements = ingredients.FindElements(By.TagName("li"));
            for(int i=0; i<ulElements.Count; i++)
            {
                CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

                _grossQty = WaitForElementExists(By.XPath("/html/body/div[4]/div/div/div[2]/div[2]/div[2]/div/div/div[1]/div[2]/ul/li[1]/div/div/div/form/table[2]/tbody/tr/td[7]/span[1]"));
                var grossQty = _grossQty.Text;

                sommeGrossQuantity += Convert.ToDouble(grossQty, ci);

            }
            return sommeGrossQuantity;


         }
   

        public bool VerifyWorkshop()
            {
                try
                {
                    IWebElement dropdown = _webDriver.FindElement(By.XPath("//select[@id='recipe_WorkshopId']"));

                    SelectElement select = new SelectElement(dropdown);

                    string selectedOptionText = select.SelectedOption.Text;

                    int dropdownWidth = dropdown.Size.Width;

                    int textWidth = GetTextWidth( selectedOptionText, dropdown.GetCssValue("font-family"), dropdown.GetCssValue("font-size"));

                    if (textWidth <= dropdownWidth)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                catch (NoSuchElementException)
                {
                    return false;
                }
        }

            private int GetTextWidth( string text, string fontFamily, string fontSize)
            {
        
                return text.Length * 8; 
            }
        // Add this function in your DatasheetDetailPage class
        public bool IsGreenIconDisplayedForRecipe(string recipeName)
        {
            WaitPageLoading();

            // Locate the row for the recipe by name and check if the green icon is present
            var recipeRow = _webDriver.FindElement(By.XPath($"//p[contains(text(), '{recipeName}')]"));
            var greenIcon = _webDriver.FindElement(By.XPath($"//li[contains(@class, 'editable-row') and contains(@id, 'dataItem')][.//p[contains(text(), '{recipeName}')]]//img[@src='/images/icons/icon-picture.png' and @title='Has a picture']"));

            // Vérifier si l'icône est bien affichée et de couleur verte
            if (greenIcon.Displayed)
            {
                var iconColor = greenIcon.GetCssValue("color");
                return iconColor == "rgba(0, 128, 0, 1)";
            }
            return true;

        }

    }
}
