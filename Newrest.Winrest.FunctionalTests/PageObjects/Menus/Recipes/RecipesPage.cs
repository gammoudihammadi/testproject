using DocumentFormat.OpenXml.Spreadsheet;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Menus;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Menus.Recipes
{
    public class RecipesPage : PageBase
    {
        // ______________________________________ Constantes _____________________________________________

        private const string PLUS_BTN = "//*[@id=\"tabContentRecipeContainer\"]/div[1]/div/div[2]/button";
        private const string EXTENDED_BTN = "//*[@id=\"tabContentRecipeContainer\"]/div[1]/div/div[1]/button";
        private const string NEW_RECIPE = "//*[@id=\"recipe-createBtn\"]";
        private const string DUPLICATE_RECIPE = "recipe-duplicatetoothersites";        
        private const string FIRST_RECIPE = "//*[@id=\"list-item-with-action\"]/div[2]/div/div/div[2]/table/tbody/tr/td[2]";
        private const string FOLD_ALL = "//*[@id=\"unfoldBtn\"][@class=\"unfold-all-btn unfold-all-btn-auto unfolded\"]";
        private const string UNFOLD_ALL = "//*[@id=\"unfoldBtn\"][@class=\"unfold-all-btn unfold-all-btn-auto\"]";
        private const string CONTENT = "//*[starts-with(@id,\"content_\")]";
        private const string EXPORT = "exportBtn";
        private const string EXPORT_ROBOT = "exportRobotBtn";
        private const string EXPORT_SALES_PRICE = "exportSalesPriceBtn";
        private const string EXPORT_ALLERGEN = "IncludeAllergens";
        private const string CONFIRM_EXPORT = "btn-export-recipes";
        private const string IMPORT_SALES_PRICE = "//*[@id=\"tabContentRecipeContainer\"]/div[1]/div/div[1]/div/a[4]";
        private const string RESET_FILTER_DEV = "ResetFilter";
        private const string RESET_FILTER_PATCH = "//*[@id=\"recipe-filter-form\"]/div[1]/a";
        private const string MASSIVE_DELETE = "//*[@id=\"tabContentRecipeContainer\"]/div[1]/div/div[1]/div/a[6]";
        private const string MASSIVE_PRINT = "Recipe_MassivePrintBtn";
        private const string SEARCH_FILTER = "SearchPattern";
        private const string SORTBY_FILTER = "cbSortBy";
        private const string SHOW_WITHOUT_VARIANT_FILTER = "ShowRecipesWithoutVariants";
        private const string SEARCH_RECIPES_BTN = "SearchRecipesBtn";
        private const string SHOW_ALL_FILTER_DEV = "ShowAll";
        private const string SHOW_ALL_FILTER_PATCH = "//*[@id=\"ShowActive\"][@value=\"All\"]";

        private const string SHOW_ACTIVE_FILTER_DEV = "ShowOnlyActive";
        private const string SHOW_ACTIVE_FILTER_PATCH = "//*[@id=\"ShowActive\"][@value=\"ActiveOnly\"]";

        private const string SHOW_INACTIVE_FILTER_DEV = "ShowOnlyInactive";
        private const string SHOW_INACTIVE_FILTER_PATCH = "//*[@id=\"ShowActive\"][@value=\"InactiveOnly\"]";

        private const string GUEST_FILTER = "SelectedGuestTypesIds_ms";
        private const string UNCHECK_ALL_GUEST = "/html/body/div[10]/div/ul/li[2]/a";
        private const string SEARCH_INPUT_GUEST = "/html/body/div[10]/div/div/label/input";
        private const string MEAL_FILTER = "SelectedMealTypesIds_ms";
        private const string UNCHECK_ALL_MEAL = "/html/body/div[11]/div/ul/li[2]/a";
        private const string SEARCH_INPUT_MEAL = "/html/body/div[11]/div/div/label/input";
        private const string INTOLERANCE_FILTER = "SelectedIntolerancesIds_ms";
        private const string UNCHECK_ALL_INTOLERANCE = "/html/body/div[12]/div/ul/li[2]/a";
        private const string SEARCH_INPUT_INTOLERANCE = "/html/body/div[12]/div/div/label/input";
        private const string RECIPE_TYPE_FILTER = "SelectedRecipeTypesIds_ms";
        private const string UNCHECK_ALL_RECIPE = "/html/body/div[13]/div/ul/li[2]/a";
        private const string SEARCH_INPUT_RECIPE = "/html/body/div[13]/div/div/label/input";
        private const string DIETARY_TYPE_FILTER = "SelectedDietaryTypesIds_ms";
        private const string UNCHECK_ALL_DIETARY = "/html/body/div[14]/div/ul/li[2]/a";
        private const string SEARCH_INPUT_DIETARY = "/html/body/div[14]/div/div/label/input";
        
        private const string NAME = "//*[@id=\"list-item-with-action\"]/div[{0}]/div[1]/div/div[2]/table/tbody/tr/td[2]";
        private const string NAME1 = "//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]/table/tbody/tr/td[2]";
        private const string PORTION = "//*[@id=\"list-item-with-action\"]/div[{0}]/div[1]/div/div[2]/table/tbody/tr/td[3]";
        private const string PORTION1 = "//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]/table/tbody/tr/td[3]";
        private const string VARIANT = "//*[@id=\"list-item-with-action\"]/div[{0}]/div[1]/div/div[2]/table/tbody/tr/td[4]";
        private const string VARIANT1 = "//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]/table/tbody/tr/td[4]";
        private const string INACTIVE = "//*[@id=\"list-item-with-action\"]/div[{0}]/div[1]/div/div[2]/table/tbody/tr/td[1]/img[@alt='Inactive']";
        private const string FIRSTVARIANT = "menus-recipes-recipe-variants-name-1"; 
        private const string RECIPE = "//*[@id=\"list-item-with-action\"]/div[2]/div[1]/div/div[2]/table/tbody/tr";
        private const string PRINTBUTTON = "header-print-button";
        private const string CLEAR = "//*[starts-with(@id,\"popover\")]/div[2]/div/a[2]"; 

        // ______________________________________ Variables _____________________________________________

        [FindsBy(How = How.XPath, Using = PLUS_BTN)]
        private IWebElement _plusButton;

        [FindsBy(How = How.XPath, Using = EXTENDED_BTN)]
        private IWebElement _extendedButton;

        [FindsBy(How = How.XPath, Using = NEW_RECIPE)]
        private IWebElement _createNewRecipe;

        [FindsBy(How = How.Id, Using = DUPLICATE_RECIPE)]
        private IWebElement _duplicateRecipesVariant;

        [FindsBy(How = How.XPath, Using = FIRST_RECIPE)]
        private IWebElement _firstRecipe;
        
        [FindsBy(How = How.XPath, Using = RECIPE)]
        private IWebElement _recipe;


        [FindsBy(How = How.XPath, Using = UNFOLD_ALL)]
        private IWebElement _unfoldAll;

        [FindsBy(How = How.XPath, Using = FOLD_ALL)]
        private IWebElement _foldAll;

        [FindsBy(How = How.Id, Using = EXPORT)]
        private IWebElement _export;
        [FindsBy(How = How.Id, Using = PRINTBUTTON)]
        private IWebElement _printButton;

        [FindsBy(How = How.Id, Using = EXPORT_ALLERGEN)]
        private IWebElement _exportAllergen;

        [FindsBy(How = How.Id, Using = CONFIRM_EXPORT)]
        private IWebElement _confirmExport;

        [FindsBy(How = How.Id, Using = EXPORT_SALES_PRICE)]
        private IWebElement _exportSalesPrice;

        [FindsBy(How = How.Id, Using = EXPORT_ROBOT)]
        private IWebElement _exportRobot;

        [FindsBy(How = How.XPath, Using = IMPORT_SALES_PRICE)]
        private IWebElement _importSalesPrice;

        [FindsBy(How = How.XPath, Using = MASSIVE_DELETE)]
        private IWebElement _massiveDelete;

        [FindsBy(How = How.Id, Using = MASSIVE_PRINT)]
        private IWebElement _massivePrint;


        // ______________________________________ Filtres _____________________________________________

        [FindsBy(How = How.Id, Using = RESET_FILTER_DEV)]
        private IWebElement _resetFilterDev;
        
        [FindsBy(How = How.XPath, Using = RESET_FILTER_PATCH)]
        private IWebElement _resetFilterPatch;

        [FindsBy(How = How.Id, Using = SEARCH_FILTER)]
        private IWebElement _searchFilter;

        [FindsBy(How = How.Id, Using = SORTBY_FILTER)]
        private IWebElement _sortBy;

        [FindsBy(How = How.Id, Using = SHOW_WITHOUT_VARIANT_FILTER)]
        private IWebElement _showWithoutVariants;

        [FindsBy(How = How.Id, Using = SHOW_ALL_FILTER_DEV)]
        private IWebElement _showAllDev;
        
        [FindsBy(How = How.XPath, Using = SHOW_ALL_FILTER_PATCH)]
        private IWebElement _showAllPatch;

        [FindsBy(How = How.Id, Using = SHOW_ACTIVE_FILTER_DEV)]
        private IWebElement _showActiveDev;
        
        [FindsBy(How = How.XPath, Using = SHOW_ACTIVE_FILTER_PATCH)]
        private IWebElement _showActivePatch;

        [FindsBy(How = How.Id, Using = SHOW_INACTIVE_FILTER_DEV)]
        private IWebElement _showInactiveDev;
        
        [FindsBy(How = How.XPath, Using = SHOW_INACTIVE_FILTER_PATCH)]
        private IWebElement _showInactivePatch;

        [FindsBy(How = How.Id, Using = GUEST_FILTER)]
        private IWebElement _guestTypeFilter;

        [FindsBy(How = How.XPath, Using = UNCHECK_ALL_GUEST)]
        private IWebElement _uncheckAllGuest;

        [FindsBy(How = How.XPath, Using = SEARCH_INPUT_GUEST)]
        private IWebElement _searchInputGuest;

        [FindsBy(How = How.Id, Using = MEAL_FILTER)]
        private IWebElement _mealTypeFilter;

        [FindsBy(How = How.XPath, Using = UNCHECK_ALL_MEAL)]
        private IWebElement _uncheckAllMeal;

        [FindsBy(How = How.XPath, Using = SEARCH_INPUT_MEAL)]
        private IWebElement _searchInputMeal;

        [FindsBy(How = How.Id, Using = INTOLERANCE_FILTER)]
        private IWebElement _intoleranceFilter;

        [FindsBy(How = How.XPath, Using = UNCHECK_ALL_INTOLERANCE)]
        private IWebElement _uncheckAllIntolerance;

        [FindsBy(How = How.XPath, Using = SEARCH_INPUT_INTOLERANCE)]
        private IWebElement _searchInputIntolerance;

        [FindsBy(How = How.Id, Using = RECIPE_TYPE_FILTER)]
        private IWebElement _recipeTypeFilter;

        [FindsBy(How = How.XPath, Using = UNCHECK_ALL_RECIPE)]
        private IWebElement _uncheckAllRecipeType;

        [FindsBy(How = How.XPath, Using = SEARCH_INPUT_RECIPE)]
        private IWebElement _searchInputRecipeType;

        [FindsBy(How = How.Id, Using = DIETARY_TYPE_FILTER)]
        private IWebElement _dietaryTypeFilter;

        [FindsBy(How = How.XPath, Using = UNCHECK_ALL_DIETARY)]
        private IWebElement _uncheckAllDietary;

        [FindsBy(How = How.XPath, Using = SEARCH_INPUT_DIETARY)]
        private IWebElement _searchInputDietary;
        [FindsBy(How = How.XPath, Using = CLEAR)]
        private IWebElement _clearButton;


        public enum FilterType
        {
            SearchRecipe,
            SortBy,
            ShowWithoutVariants,
            ShowAll,
            ShowActive,
            ShowInactive,
            TypeOfGuest,
            TypeOfMeal,
            Intolerance,
            RecipeType,
            Dietarytype
        }

        public void Filter(FilterType filterType, object value)
        {
            Actions action = new Actions(_webDriver);

            switch (filterType)
            {
                case FilterType.SearchRecipe:
                     _searchFilter = WaitForElementIsVisibleNew(By.Id("SearchPattern"));
                    _searchFilter.SetValue(ControlType.TextBox, value);
                    break;
                case FilterType.SortBy:
                    _sortBy = WaitForElementExists(By.Id(SORTBY_FILTER));
                    _sortBy.Click();
                    var element = WaitForElementIsVisibleNew(By.XPath("//option[contains(@value,'" + value + "')]"));
                    _sortBy.SetValue(ControlType.DropDownList, element.Text);
                    _sortBy.Click();
                    break;
                case FilterType.ShowWithoutVariants:
                    _showWithoutVariants = WaitForElementExists(By.Id(SHOW_WITHOUT_VARIANT_FILTER));
                    action.MoveToElement(_showWithoutVariants).Perform();
                    _showWithoutVariants.SetValue(ControlType.CheckBox, value);
                    break;
                case FilterType.ShowAll:
                    try
                    {
                        _showAllDev = WaitForElementExists(By.Id(SHOW_ALL_FILTER_DEV));
                        action.MoveToElement(_showAllDev).Perform();
                        _showAllDev.SetValue(ControlType.RadioButton, value);
                    }
                    catch
                    {
                        _showAllPatch = WaitForElementExists(By.XPath(SHOW_ALL_FILTER_PATCH));
                        action.MoveToElement(_showAllPatch).Perform();
                        _showAllPatch.SetValue(ControlType.RadioButton, value);
                    }
                    break;
                case FilterType.ShowActive:
                    _showActiveDev = WaitForElementExists(By.Id(SHOW_ACTIVE_FILTER_DEV));
                    action.MoveToElement(_showActiveDev).Perform();
                    _showActiveDev.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.ShowInactive:
                    _showInactiveDev = WaitForElementExists(By.Id(SHOW_INACTIVE_FILTER_DEV));
                    action.MoveToElement(_showInactiveDev).Perform();
                    _showInactiveDev.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.TypeOfGuest:
                    _guestTypeFilter = WaitForElementIsVisibleNew(By.Id(GUEST_FILTER));
                    _guestTypeFilter.Click();

                    _uncheckAllGuest = WaitForElementIsVisibleNew(By.XPath(UNCHECK_ALL_GUEST));
                    _uncheckAllGuest.Click();

                    _searchInputGuest = WaitForElementIsVisibleNew(By.XPath(SEARCH_INPUT_GUEST));
                    _searchInputGuest.SetValue(ControlType.TextBox, value);

                    var guestTypeToCheck = WaitForElementIsVisibleNew(By.XPath("//input[@name='multiselect_SelectedGuestTypesIds']/../span[text()='" + value + "']"));
                    guestTypeToCheck.SetValue(ControlType.CheckBox, true);

                    _guestTypeFilter.Click();
                    break;
                case FilterType.TypeOfMeal:
                    _mealTypeFilter = WaitForElementIsVisibleNew(By.Id(MEAL_FILTER));
                    _mealTypeFilter.Click();

                    _uncheckAllMeal = WaitForElementIsVisibleNew(By.XPath(UNCHECK_ALL_MEAL));
                    _uncheckAllMeal.Click();

                    _searchInputMeal = WaitForElementIsVisibleNew(By.XPath(SEARCH_INPUT_MEAL));
                    _searchInputMeal.SetValue(ControlType.TextBox, value);

                    var mealTypeToCheck = WaitForElementIsVisibleNew(By.XPath("//input[@name='multiselect_SelectedMealTypesIds']/../span[text()='" + value + "']"));
                    mealTypeToCheck.SetValue(ControlType.CheckBox, true);

                    _mealTypeFilter.Click();
                    break;
                case FilterType.Intolerance:
                    _intoleranceFilter = WaitForElementIsVisibleNew(By.Id(INTOLERANCE_FILTER));
                    _intoleranceFilter.Click();

                    _uncheckAllIntolerance = WaitForElementIsVisibleNew(By.XPath(UNCHECK_ALL_INTOLERANCE));
                    _uncheckAllIntolerance.Click();

                    _searchInputIntolerance = WaitForElementIsVisibleNew(By.XPath(SEARCH_INPUT_INTOLERANCE));
                    _searchInputIntolerance.SetValue(ControlType.TextBox, value);

                    var intoleranceToCheck = WaitForElementIsVisibleNew(By.XPath("//span[text()='" + value + "']"));
                    intoleranceToCheck.SetValue(ControlType.CheckBox, true);

                    _intoleranceFilter.Click();
                    break;
                case FilterType.RecipeType:
                    _recipeTypeFilter = WaitForElementIsVisibleNew(By.Id(RECIPE_TYPE_FILTER));
                    _recipeTypeFilter.Click();

                    _uncheckAllRecipeType = WaitForElementIsVisibleNew(By.XPath(UNCHECK_ALL_RECIPE));
                    _uncheckAllRecipeType.Click();

                    _searchInputRecipeType = WaitForElementIsVisibleNew(By.XPath(SEARCH_INPUT_RECIPE));
                    _searchInputRecipeType.SetValue(ControlType.TextBox, value);

                    var recipeTypeToCheck = WaitForElementIsVisibleNew(By.XPath("//span[text()='" + value + "']"));
                    recipeTypeToCheck.SetValue(ControlType.CheckBox, true);

                    _recipeTypeFilter.Click();
                    break;
                case FilterType.Dietarytype:
                    _dietaryTypeFilter = WaitForElementIsVisibleNew(By.Id(DIETARY_TYPE_FILTER));
                    _dietaryTypeFilter.Click();

                    _uncheckAllDietary = WaitForElementIsVisibleNew(By.XPath(UNCHECK_ALL_DIETARY));
                    _uncheckAllDietary.Click();

                    _searchInputDietary = WaitForElementIsVisibleNew(By.XPath(SEARCH_INPUT_DIETARY));
                    _searchInputDietary.SetValue(ControlType.TextBox, value);

                    var dietaryTypeToCheck = WaitForElementIsVisibleNew(By.XPath("//span[text()='" + value + "']"));
                    dietaryTypeToCheck.SetValue(ControlType.CheckBox, true);

                    _dietaryTypeFilter.Click();
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);

            }

            WaitPageLoading();
            Thread.Sleep(2500);
        }

        public void ResetFilter()
        {
            try
            {
                _resetFilterDev = WaitForElementIsVisibleNew(By.Id(RESET_FILTER_DEV));
                _resetFilterDev.Click();
            }
            catch
            {
                _resetFilterPatch = WaitForElementIsVisibleNew(By.XPath(RESET_FILTER_PATCH));
                _resetFilterPatch.Click();
            }
            WaitForLoad();

            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                // pas de date
            }
        }

        public bool IsSortedByName()
        {
            bool valueBool = true;
            var ancienName = "";
            int tot;

            if (CheckTotalNumber() > 100)
            {
                tot = 100;
            }
            else
            {
                tot = CheckTotalNumber();
            }
            if (tot == 0)
                return false;


            var elements = _webDriver.FindElements(By.XPath(NAME1));

            foreach (var element in elements)
            {
                try
                {

                    string name = element.Text;

                    if (String.Compare(ancienName, name) > 0)
                    { valueBool = false; }

                    ancienName = name;
                }
                catch
                {
                    valueBool = false;
                }
            }

            return valueBool;
        }

        public bool IsSortedByPortions()
        {
            bool valueBool = true;
            var ancientNumber = int.MaxValue;
            int tot;

            if (CheckTotalNumber() > 100)
            {
                tot = 100;
            }
            else
            {
                tot = CheckTotalNumber();
            }
            if (tot == 0)
                return false;


            var elements = _webDriver.FindElements(By.XPath(PORTION1));

            foreach (var element in elements)
            {
                try
                {
                    //IWebElement element = _webDriver.FindElement(By.XPath(xpath));
                    int portion = 0;
                    portion = element.Text.IndexOf("portions");
                    if (portion == -1)
                    {
                        portion = element.Text.IndexOf("KG");
                    }
                    int number = int.Parse(element.Text.Substring(0, portion));

                    if (number > ancientNumber)
                    {
                        valueBool = false;
                    }

                    ancientNumber = number;
                }
                catch
                {
                    valueBool = false;
                }
            }

            return valueBool;
        }

        public bool IsWithVariant()
        {
            bool valueBool = true;
            int tot;

            if (CheckTotalNumber() > 100)
            {
                tot = 100;
            }
            else
            {
                tot = CheckTotalNumber();
            }

            if (tot == 0)
                return false;


            var elements = _webDriver.FindElements(By.XPath(VARIANT1));

            foreach (var element in elements)
            {

                try
                {

                    int number = int.Parse(element.Text);

                    if (element.Text.Equals("0"))
                    {
                        valueBool = false;
                        break;
                    }
                }
                catch
                {
                    valueBool = false;
                }
            }


            return valueBool;
        }

        public bool IsWithoutVariant()
        {
            bool valueBool = true;
            int tot;

            if (CheckTotalNumber() > 100)
            {
                tot = 100;
            }
            else
            {
                tot = CheckTotalNumber();
            }

            if (tot == 0)
                return false;


            var elements = _webDriver.FindElements(By.XPath(VARIANT1));

            foreach (var element in elements)
            {

                try
                {

                    int number = int.Parse(element.Text);

                    if (!element.Text.Equals("0"))
                    {
                        valueBool = false;
                        break;
                    }
                }
                catch
                {
                    valueBool = false;
                }
            }

            return valueBool;
        }

        public bool CheckStatus(bool active)
        {
            bool isActive = false;
            int tot = CheckTotalNumber() > 100 ? 100 : CheckTotalNumber();

            if (tot == 0)
                return false;

            for (int i = 1; i <= tot; i++)
            {
                try
                {
                    _webDriver.FindElement(By.XPath(String.Format(INACTIVE, i + 1)));

                    if (active)
                        return false;
                }
                catch
                {
                    isActive = true;
                    if (!active)
                        return true;
                }
            }
            return isActive;
        }

        public bool VerifyVariant(List<string> variants, string value)
        {
            bool valueBool = true;

            if(variants != null)
            {
                foreach(string variant in variants)
                {
                    if (!variant.Contains(value))
                    {
                        valueBool = false;
                        break;
                    }
                }
            }
            else
            {
                valueBool = false;
            }

            return valueBool;
        }


        public Boolean IsExist()
        {
            bool valueBool = false;

            if (CheckTotalNumber() > 0)
            {
                valueBool = true;
            }

            return valueBool;
        }

        // ______________________________________ Méthodes ____________________________________________

        public RecipesPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        public override void ShowPlusMenu()
        {
            Actions action = new Actions(_webDriver);
            _plusButton = WaitForElementIsVisibleNew(By.XPath(PLUS_BTN));
            action.MoveToElement(_plusButton).Perform();
            WaitForLoad();
        }

        public override void ShowExtendedMenu()
        {
            Actions action = new Actions(_webDriver);
            _extendedButton = WaitForElementIsVisible(By.XPath(EXTENDED_BTN));
            action.MoveToElement(_extendedButton).Perform();
        }

        public RecipesCreateModalPage CreateNewRecipe()
        {
            ShowPlusMenu();    
            WaitForLoad();
            _createNewRecipe = WaitForElementIsVisibleNew(By.XPath(NEW_RECIPE));
            _createNewRecipe.Click();
            WaitPageLoading();
            WaitForLoad();

            return new RecipesCreateModalPage(_webDriver, _testContext);
        }

        public DuplicateRecipesVariantModalPage DuplicateRecipesVariant()
        {
            ShowPlusMenu();
            _duplicateRecipesVariant = WaitForElementIsVisible(By.Id(DUPLICATE_RECIPE));
            _duplicateRecipesVariant.Click();
            WaitForLoad();

            return new DuplicateRecipesVariantModalPage(_webDriver, _testContext);
        }

        public RecipeGeneralInformationPage SelectFirstRecipe()
        {
            WaitForLoad();
            _firstRecipe = WaitForElementIsVisibleNew(By.XPath(FIRST_RECIPE));
            _firstRecipe.Click();
            WaitForLoad();

            return new RecipeGeneralInformationPage(_webDriver, _testContext);
        }

        public string GetFirstRecipeName()
        {
            if (isElementVisible(By.XPath(FIRST_RECIPE)))
            {
                _firstRecipe = WaitForElementExists(By.XPath(FIRST_RECIPE));
                return _firstRecipe.Text;
            }
            else
            {
                return "";
            }
            
        }

        public void UnfoldAll()
        {
            ShowExtendedMenu();

            _unfoldAll = WaitForElementIsVisible(By.XPath(UNFOLD_ALL));
            _unfoldAll.Click();
            WaitForLoad();
        }

        public Boolean IsUnfoldAll()
        {
            bool valueBool = false;

            var content = WaitForElementIsVisible(By.XPath(CONTENT));

            // Temps nécessaire pour que l'élément change de classe
            WaitPageLoading();
            if (content.GetAttribute("class") == "panel-collapse collapse in" || content.GetAttribute("class") == "panel-collapse collapse show")
                valueBool = true;

            return valueBool;
        }

        public void FoldAll()
        {
            ShowExtendedMenu();

            _foldAll = WaitForElementIsVisible(By.XPath(FOLD_ALL));
            _foldAll.Click();
            WaitForLoad();
        }

        public Boolean IsFoldAll()
        {
            bool valueBool = false;

            var content = WaitForElementExists(By.XPath(CONTENT));

            // Temps nécessaire pour que l'élément change de classe
            WaitPageLoading();

            if (content.GetAttribute("class") == "panel-collapse collapse")
                valueBool = true;

            return valueBool;
        }

        public void Export(bool versionPrint, bool allergen = false)
        {
            ShowExtendedMenu();
            _export = WaitForElementIsVisible(By.Id(EXPORT));
            _export.Click();
            WaitForLoad();

            WaitForElementToBeClickable(By.Id(CONFIRM_EXPORT));

            _exportAllergen = WaitForElementExists(By.Id(EXPORT_ALLERGEN));
            _exportAllergen.SetValue(ControlType.CheckBox, allergen);

            _confirmExport = WaitForElementIsVisible(By.Id(CONFIRM_EXPORT));
            _confirmExport.Click();
            WaitForLoad();

            if (versionPrint)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));
                ClickPrintButton();
            }

            WaitForDownload();
            Close();
        }

        public void ExportSalesPrice(bool versionPrint)
        {
            ShowExtendedMenu();

            _exportSalesPrice = WaitForElementIsVisible(By.Id(EXPORT_SALES_PRICE));
            _exportSalesPrice.Click();
            WaitForLoad();

            if (versionPrint)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));
                ClickPrintButton();
            }

            WaitForDownload();
            Close();
            ClosePrintButton();
        }

        public bool ImportSalesPrice(string filename)
        {
            ShowExtendedMenu();

            _importSalesPrice = WaitForElementIsVisible(By.XPath(IMPORT_SALES_PRICE));
            _importSalesPrice.Click();
            WaitForLoad();

            var _importDatasheet = WaitForElementIsVisible(By.Id("fileSent"));
            _importDatasheet.SendKeys(filename);
            WaitForLoad();

            var checkButton = WaitForElementIsVisible(By.XPath("//*/button[text()='Check file']"));
            checkButton.Click();
            WaitForLoad();

            if(isElementVisible(By.ClassName("green-text")))
            {
                return true;
            }
            return false;
        }

        public void ClodeImportModal()
        {
            var cancelButton = WaitForElementIsVisible(By.XPath("//*/button[text()='Cancel']"));
            cancelButton.Click();
            WaitForLoad();
        }

        public void ExportRobot(bool versionPrint)
        {
            ShowExtendedMenu();

            _exportRobot = WaitForElementIsVisible(By.Id(EXPORT_ROBOT));
            _exportRobot.Click();
            WaitForLoad();

            _confirmExport = WaitForElementIsVisible(By.Id(CONFIRM_EXPORT));
            _confirmExport.Click();
            WaitForLoad();

            if (versionPrint)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));
                ClickPrintButton();
            }

            WaitForDownload();
            Close();
        }

        public FileInfo GetExportExcelFile(FileInfo[] taskFiles, bool isSalesPrice, bool robot)
        {
            List<FileInfo> correctDownloadFiles = new List<FileInfo>();
            foreach (var file in taskFiles)
            {
                //  Test REGEX
                if (IsExcelFileCorrect(file.Name, isSalesPrice, robot))
                {
                    correctDownloadFiles.Add(file);
                }
            }

            if (correctDownloadFiles.Count <= 0)
            {
                throw new Exception(MessageErreur.FICHIER_NON_TROUVE);
            }

            var time = correctDownloadFiles[0].LastWriteTimeUtc;
            var correctFile = correctDownloadFiles[0];

            correctDownloadFiles.ForEach(file =>
            {
                if (time < file.LastWriteTimeUtc)
                {
                    time = file.LastWriteTimeUtc;
                    correctFile = file;
                }
            });

            return correctFile;
        }

        public bool IsExcelFileCorrect(string filePath, bool salesPrice, bool robot)
        {
            Match m = null;
            //Recipes 2020-05-26 14-17-48.xlsx
            string mois = "(?:0[1-9]|1[0-2])";         // mois
            string space = "(\\s)";                    // Espace
            string annee = "\\d{4}";                   // annee YYYY
            string jour = "[0-3]\\d";                  // jour
            string heure = "(?:0[0-9]|1[0-9]|2[0-3])"; // heure
            string minutes = "[0-5]\\d";               // minutes
            string secondes = "[0-5]\\d";              // secondes

            Regex r = new Regex("^Recipes" + space + annee + "-" + mois + "-" + jour + space + heure + "-" + minutes + "-" + secondes + ".xlsx$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Regex rsp = new Regex("^Recipes Sales Price" + space + annee + "-" + mois + "-" + jour + space + heure + "-" + minutes + "-" + secondes + ".xlsx$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Regex rb = new Regex("^RecipeRobot" + "_" + annee + "-" + mois + "-" + jour + space + heure + "-" + minutes + "-" + secondes + ".xlsx$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            if(salesPrice)
            {
                m = rsp.Match(filePath);
            }
            else if(robot)
            {
                m = rb.Match(filePath);
            }
            else
            {
                m = r.Match(filePath);
            }
            

            return m.Success;
        }
        public void UpdateFileExported(string filePath, string sheetName, FileInfo correctDownloadedFile)
        {
            // On récupère les 3 ids dans le fichier Excel
            var ids = OpenXmlExcel.GetValuesInList("Recipe Id", "Recipes Sales Price", correctDownloadedFile.FullName);
            var id1 = ids[0];
            var id2 = ids[1];
            string id3 = "";
            if (CheckTotalNumber()>2)
            {
                id3 = ids[2];
            }

            //On modifie l'Action pour les 3 datasheets
            OpenXmlExcel.WriteDataInCell("Action", "Recipe Id", id1, "Recipes Sales Price", filePath, "M", CellValues.SharedString);
            OpenXmlExcel.WriteDataInCell("Action", "Recipe Id", id2, "Recipes Sales Price", filePath, "M", CellValues.SharedString);
            if (CheckTotalNumber() > 2)
            {
                OpenXmlExcel.WriteDataInCell("Action", "Recipe Id", id3, "Recipes Sales Price", filePath, "M", CellValues.SharedString);
            }
        }
        public void MassiveDeleteRecipes(string recipeName, string site, string recipeType)
        {
            ShowExtendedMenu();

            _massiveDelete = WaitForElementIsVisible(By.XPath(MASSIVE_DELETE));
            _massiveDelete.Click();
            WaitForLoad();

            var searchPattern = WaitForElementIsVisible(By.XPath("//*[@id=\"formMassiveDeleteRecipe\"]/div/div[1]/div/input[@id=\"SearchPattern\"]"));
            searchPattern.SendKeys(recipeName);
            ComboBoxSelectById(new ComboBoxOptions("SelectedSiteIds_ms", site, false));
            ComboBoxSelectById(new ComboBoxOptions("SelectedRecipeTypes_ms", recipeType, false));

            var searchRecipesbtn = WaitForElementIsVisible(By.Id(SEARCH_RECIPES_BTN));
            searchRecipesbtn.Click();
            WaitForLoad();

            if(isElementVisible(By.XPath("//*[@id=\"tableRecipes\"]/tbody/tr")))
            {
                var checkInput = WaitForElementIsVisible(By.XPath("//*[@id=\"tableRecipes\"]/tbody/tr/td[1]"));
                checkInput.Click();
                WaitForLoad();
            }

            var deleteBtn = WaitForElementIsVisible(By.Id("deleteRecipeBtn"));
            deleteBtn.Click();
            WaitForLoad();
            WaitLoading(); 

            var deleteConfirmBtn = WaitForElementIsVisible(By.Id("dataConfirmOK"));
            deleteConfirmBtn.Click();
            WaitLoading();
            WaitForLoad();

            var okButton = WaitForElementIsVisible(By.XPath("//*/button[text()='Ok']"));
            okButton.Click();
            WaitPageLoading();
        }
        public void MassivePrint()
        {
            WaitLoading();
            ShowExtendedMenu();
            WaitLoading();
            _massivePrint = WaitForElementExists(By.Id(MASSIVE_PRINT));
            _massivePrint.Click();
            WaitLoading();

            var numberOfportions = WaitForElementIsVisible(By.Id("NumberOfPortions"));
            numberOfportions.Clear();
            numberOfportions.SendKeys("1");

            var printbtn = WaitForElementIsVisible(By.Id("print"));
            printbtn.Click();
            WaitLoading();
        }

        public bool IsFileExistInPrintList()
        {
            try
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-archive']"));
                return true;
            }
            catch { return false; }
        }

        public RecipeMassiveDeletePopup OpenMassiveDeletePopup()
        {
            WaitPageLoading();
            ShowExtendedMenu();

            var massiveDelete = WaitForElementIsVisible(By.XPath(MASSIVE_DELETE));
            massiveDelete.Click();
            WaitPageLoading();
            return new RecipeMassiveDeletePopup(_webDriver, _testContext);
        }
        public string GetFirstVariant()
        {
            var variant = WaitForElementExists(By.Id(FIRSTVARIANT));
            return variant.Text;
        }
        public void Go_To_New_Navigate()
        {
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);
        }

        public string GetSearchValue()
        {
            var _searchInput = WaitForElementIsVisible(By.Id(SEARCH_FILTER));
            return _searchInput.Text;
        }

        public RecipeGeneralInformationPage SelectRecipe()
        {
            WaitForLoad();
            _recipe = WaitForElementIsVisible(By.XPath(RECIPE));
            _recipe.Click();
            WaitForLoad();

            return new RecipeGeneralInformationPage(_webDriver, _testContext);
        }
        private void WaitForModalToClose()
        {
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(10));
            wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.InvisibilityOfElementLocated(By.ClassName("modal-backdrop")));
        }

        public PrintReportPage ClickPrinterIcon()
        {
            WaitForLoad();
            WaitForModalToClose(); // Wait for the modal to close
            _printButton = _webDriver.FindElement(By.Id(PRINTBUTTON));
            _printButton.Click();
            return new PrintReportPage(_webDriver, _testContext);
        }


        public void ImportFile()
        {
            var _importFileBtn = WaitForElementToBeClickable(By.Id("check-file-btn-ImportSalesPrice"));
            _importFileBtn.Click();
            WaitForLoad();

            // Temps de prise en compte des changements
            WaitPageLoading();
        }

        public void CloseAfterImport()
        {
            var _closeImportBtn = WaitForElementToBeClickable(By.Id("Close-ImportSalesPrice"));
            _closeImportBtn.Click();
            WaitForLoad();

            // Temps de prise en compte des changements
            WaitPageLoading();
        }
        public int GetTotalRcipes()
        {
            var numbers = WaitForElementIsVisible(By.XPath("/html/body/div[2]/div/div[2]/div/div[1]/h1/span"));
            return Int32.Parse(numbers.Text);
        }

    }
}
