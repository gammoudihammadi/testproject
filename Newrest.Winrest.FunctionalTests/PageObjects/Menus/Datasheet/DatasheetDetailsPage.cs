using DocumentFormat.OpenXml.Bibliography;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Recipes;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Menus.Datasheet
{
    public class DatasheetDetailsPage : PageBase
    {
        public DatasheetDetailsPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // ______________________________________________ Constantes ____________________________________________

        // Général
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";
        private const string GENERAL_INFORMATION = "-1";
        private const string PRINT = "PrintButton";
        private const string CONFIRM_PRINT = "print_btn";

        /*************************************** Onglets *******************************************/
        // Recipe
        private const string RECIPE_TAB = "hrefTabContentRecipes";

        private const string CREATE_NEW_RECIPE_DEV = "createNewRecipe";
        private const string CREATE_NEW_RECIPE_PATCH = "/html/body/div[3]/div/div[2]/div[3]/div/div/div[2]/div/div/div/div/div[2]/div[1]/ul/li/div/div[1]";

        private const string SHOW_DETAILS_DEV = "showDetails";
        private const string SHOW_DETAILS_PATCH = "//*[@id=\"datasheet-list\"]/div[1]/div/a[2]/span[@class='glyphicon glyphicon-eye-open']";

        private const string HIDE_DETAILS_DEV = "//*[@id=\"showDetails\"]/span[@class='fas fa-eye-slash']";
        private const string HIDE_DETAILS_PATCH = "//*[@id=\"datasheet-list\"]/div[1]/div/a[2]/span[@class='glyphicon glyphicon-eye-close']";
        private const string ADD_RECIPE = "RecipeName";
        private const string SELECTED_RECIPE = "//*[@id=\"recipe-list-result\"]/table/tbody/tr[1]/td[2]";
        private const string ADD_RECIPE_ERROR_MSG = "//*[@id=\"recipe-list-result\"]/p"; 
        private const string ADD_RECIPE_FROM_DATASHEET = "DatasheetName";
        private const string FIRST_RECIPE_FROM_DATASHEET = "//*[@id=\"datasheet-list-result\"]/table/tbody/tr[*]//p[contains(text(),'{0}')]";
        private const string SAVE_ADD_RECIPE_FROM_DATASHEET = "btn-import-recipe";
        private const string ADD_RECIPE_DATASHEET_ERROR_MSG = "datasheet-list-result";
        private const string ITEM_RECIPE = "/html/body/div[3]/div/div[2]/div[3]/div/div/div[2]/div/div/div/div/div[2]/ul/li[{0}]/div/div[2]/div[1]/div[4]/span";
        private const string NONE = "/html/body/div[4]/div/div/div/div/form/div[2]/div/table/thead/tr/th[1]/a[2]";

        private const string SELECT_SITES = "sites-detail-filter";

        private const string ADDED_RECIPE = "//*[@id='datasheet-list']/div[2]/div[1]/ul/li[@class='editable-row']//div[contains(@class,'display-zone')]/div[contains(@class,'row')]";
        private const string FIRST_RECIPE_NAME = "menus-datasheet-datasheet-recipe-name-1";

        private const string RECIPE_DETAILS = "//*[starts-with(@id,\"dataItem\")]/div/div[2]/div[2]";
        private const string ADD_SUBRECIPE_DEV = "//span[@class = 'fas fa-plus']";
        private const string ADD_SUBRECIPE_PATCH = "menus-datasheet-createdatasheetrecipe-edit-1";
        private const string ADDED_SUBRECIPE = "//*[starts-with(@id,\"dataItem\")]/div/div[2]/div[2]/div[2]/div[1]/img";
        private const string DELETE_RECIPE_DEV = "//span[@class='fas fa-trash-alt']";
        private const string DELETE_RECIPE_PATCH = "/html/body/div[3]/div/div[2]/div[3]/div/div/div[2]/div/div/div/div/div[2]/div[1]/ul/li/div/div[1]/div[1]/div[11]/div/a[3]";
        private const string CONFIRM_DELETE_RECIPE = "dataConfirmOK";
        private const string EDIT_RECIPE_DEV = "menus-datasheet-editdatasheetrecipe-edit-1";
        private const string EDIT_SUB_RECIPE_DEV = "/html/body/div[3]/div/div[2]/div[3]/div/div/div[2]/div/div/div/div/div[2]/div[1]/ul/li/div/div[2]/div[2]/div[2]/div[8]/div/a[1]";
        private const string EDIT_RECIPE_PATCH = "//*[starts-with(@id,\"menus-datasheet-editdatasheetrecipe-edit\")]/span";
        private const string ADD_PICTURE_DEV = "//span[@class = 'fas fa-image']";
        private const string ADD_PICTURE_PATCH = "//span[@class = 'glyphicon glyphicon-picture']";
        private const string DUPLICATE_RECIPE_DEV = "//*[starts-with(@id,\"dataItem\")]/div/div[1]/div[1]/div[11]/div/a[4]/span[@class = 'fas fa-star']";
        private const string DUPLICATE_RECIPE_PATCH = "//*[starts-with(@id,\"dataItem\")]/div/div[1]/div[1]/div[11]/div/a[4]/span[@class = 'glyphicon glyphicon-star']";
        private const string DUPLICATE_ERROR_MODALE = "//*[@id=\"dataAlertModal\"]/div/div";

        private const string PICTURE_ADDED_TO_RECIPE_PATCH = "//*[starts-with(@id,\"dataItem\")]/div/div[2]/div[1]/div[2]/img";
        private const string PICTURE_ADDED_TO_RECIPE_DEV = "//*[starts-with(@id,\"dataItem\")]/div/div[2]/div[1]/div[3]/img";
        private const string NET_WEIGHT = "/html/body/div[3]/div/div[2]/div[3]/div/div/div[2]/div/div/div/div/div[2]/div[1]/ul/li/div/div[1]/div[1]/div[10]/span[1]";
        private const string PORTION = ".col-md-1.align-price .priceRecipe";
        private const string COEF_RECIPE = "//*[starts-with(@id,\"inputDetailCoeff-\")]";
        private const string METHOD_RECIPE = "//*[starts-with(@id,\"dataItem\")]/div/div[1]/div[1]/div[6]/select";
        private const string METHOD_RECIPE_SELECTED = "//*[starts-with(@id,\"dataItem\")]/div/div[1]/div[1]/div[6]/select/option";
        private const string PAX_RECIPE = "//*[starts-with(@id,\"dataItem\")]/div/div[1]/div[1]/div[9]/span[2]";
        private const string POPUP_SCALE_QTY = "/html/body/div[4]/div/div/div/div/form/div[2]/table/tbody/tr[2]/td[2]/input";
        private const string SCALE_SAVE = "last";
        private const string SUB_RECIPES = "/html/body/div[3]/div/div[2]/div[3]/div/div/div[2]/div/div/div/div/div[2]/div[1]/ul/li/div/div[2]/div[2]/div[*]/div[1]";

        private const string RECIPE_PRICE = ".display-zone .priceResult";
        private const string TOTAL_COSTING = ".totalCostingText";

        // Intolerance
        private const string INTOLERANCE_TAB = "hrefTabContentAllergen";
        private const string VALIDATE_ALLERGEN = "bootstrap-switch-on";
        private const string NOT_VALIDATE_ALLERGEN = "bootstrap-switch-off";
        private const string ALLERGEN_VALIDATOR = "AllergenQualityValidatedBy";
        private const string ALLERGEN_VALIDATION_DATE = "AllergenQualityValidatedDate";

        // Picture
        private const string PICTURE_TAB = "hrefTabContentPicture";
        private const string UPLOAD_PICTURE = "FileSent";
        private const string PICTURE = "//*[@id=\"datasheet-detail-container\"]/div/div/div/div/form/img";
        private const string DELETE_PICTURE_ADDED = "//*[@id=\"datasheet-detail-container\"]/div/div/div/div/form/a/span";
        private const string DELETE_PICTURE = "//*[@id=\"datasheet-detail-container\"]/div/div/form/a/span";
        private const string CONFIRM_DELETE_PICTURE = "dataConfirmOK";

        // Use case
        private const string USE_CASE_TAB = "hrefTabContentUseCase";
        private const string SERVICE_EDIT = "//*[@id=\"use-case\"]/tbody/tr[*]/td[2][normalize-space(text())='{0}']/parent::*/td[6]/a/span";
        private const string SERVICE_LINE = "//*[@id=\"use-case\"]/tbody/tr[*]/td[2][normalize-space(text())='{0}']";
        private const string DS_USE_CASES = "//*[@id=\"use-case\"]/tbody/tr[*]/td[1]";

        private const string CLICKFIRSTLINERECIPES = "//*[@id=\"dsdTable\"]/ul";
        private const string EDITBUTTON = "//*[@id=\"menus-datasheet-editdatasheetrecipe-edit-1\"]";
        private const string INGREDIENTNAME = "//*[@id=\"datasheet-datasheetrecipevariantdetail-item-1\"]/td[1]/span";
        private const string PICTUREADDED = "//*[@id=\"tabContentDetails\"]/div/div/form/img";
        private const string RECIPE_ANOTHER_SITE = "//*[contains(@id,\"menus-datasheet-datasheet-recipe-name\")]";
        private const string RECIPE_TYPE_SELECTED = "//*[@id=\"RecipeTypeId\"]/option[@selected = 'selected']";
        private const string WORKSHOP_SELECTED = "//*[@id=\"WorkshopId\"]/option[@selected = 'selected']";
        private const string COOKING_MODE_SELECTED = "//*[@id=\"CookingModeId\"]/option[@selected = 'selected']";
        private const string OPENPRINT = "//*[@id=\"PrintButton\"]";
        private const string PRICE = "/html/body/div[3]/div/div[2]/div[3]/div/div/div[2]/div/div/div/div/div[2]/div[1]/ul/li/div/div[2]/div[1]/div[9]/span[2]";



        // ______________________________________________ Variables _____________________________________________

        // Général
        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;
        [FindsBy(How = How.XPath, Using = OPENPRINT)]
        private IWebElement _openprint;
        [FindsBy(How = How.XPath, Using = CLICKFIRSTLINERECIPES)]
        private IWebElement _clickfirstline;

        [FindsBy(How = How.Id, Using = GENERAL_INFORMATION)]
        private IWebElement _generalInformationTab;

        [FindsBy(How = How.Id, Using = PRINT)]
        private IWebElement _print;

        [FindsBy(How = How.Id, Using = CONFIRM_PRINT)]
        private IWebElement _confirmPrint;

        /******************************** Onglets **************************************/
        // Recipes
        [FindsBy(How = How.Id, Using = RECIPE_TAB)]
        private IWebElement _recipeTab;

        [FindsBy(How = How.XPath, Using = ITEM_RECIPE)]
        private IWebElement _recipeType;

        [FindsBy(How = How.Id, Using = CREATE_NEW_RECIPE_DEV)]
        private IWebElement _createNewRecipeDev;

        [FindsBy(How = How.XPath, Using = CREATE_NEW_RECIPE_PATCH)]
        private IWebElement _createNewRecipePatch;

        [FindsBy(How = How.Id, Using = SHOW_DETAILS_DEV)]
        private IWebElement _showDetailsDev;

        [FindsBy(How = How.XPath, Using = SHOW_DETAILS_PATCH)]
        private IWebElement _showDetailsPatch;

        [FindsBy(How = How.XPath, Using = HIDE_DETAILS_DEV)]
        private IWebElement _hideDetailsDev;

        [FindsBy(How = How.XPath, Using = HIDE_DETAILS_PATCH)]
        private IWebElement _hideDetailsPatch;

        [FindsBy(How = How.Id, Using = ADD_RECIPE)]
        private IWebElement _addRecipe;

        [FindsBy(How = How.XPath, Using = SELECTED_RECIPE)]
        private IWebElement _selectedRecipe;

        [FindsBy(How = How.XPath, Using = ADD_RECIPE_ERROR_MSG)]
        private IWebElement _addRecipeErrorMsg;

        [FindsBy(How = How.Id, Using = ADD_RECIPE_FROM_DATASHEET)]
        private IWebElement _addRecipeFromDatasheet;

        [FindsBy(How = How.XPath, Using = FIRST_RECIPE_FROM_DATASHEET)]
        private IWebElement _firstRecipeFromDatasheet;

        [FindsBy(How = How.Id, Using = SAVE_ADD_RECIPE_FROM_DATASHEET)]
        private IWebElement _saveAddRecipeFromDatasheet;

        [FindsBy(How = How.Id, Using = ADD_RECIPE_DATASHEET_ERROR_MSG)]
        private IWebElement _addRecipeDatasheetErrorMsg;

        [FindsBy(How = How.Id, Using = FIRST_RECIPE_NAME)]
        private IWebElement _firstRecipeName;

        [FindsBy(How = How.XPath, Using = RECIPE_DETAILS)]
        private IWebElement _recipeDetails;

        [FindsBy(How = How.XPath, Using = ADD_SUBRECIPE_DEV)]
        private IWebElement _addSubrecipe;        
        
        [FindsBy(How = How.Id, Using = ADD_SUBRECIPE_PATCH)]
        private IWebElement _addSubrecipePatch;


        [FindsBy(How = How.XPath, Using = DELETE_RECIPE_DEV)]
        private IWebElement _deleteRecipeDev;

        [FindsBy(How = How.XPath, Using = DELETE_RECIPE_PATCH)]
        private IWebElement _deleteRecipePatch;

        [FindsBy(How = How.Id, Using = CONFIRM_DELETE_RECIPE)]
        private IWebElement _confirmDeleteRecipe;

        [FindsBy(How = How.Id, Using = EDIT_RECIPE_DEV)]
        private IWebElement _editRecipeDev;

        [FindsBy(How = How.XPath, Using = EDIT_RECIPE_PATCH)]
        private IWebElement _editRecipePatch;

        [FindsBy(How = How.XPath, Using = ADD_PICTURE_DEV)]
        private IWebElement _addPicturedev;

        [FindsBy(How = How.XPath, Using = ADD_PICTURE_PATCH)]
        private IWebElement _addPicturepatch;

        [FindsBy(How = How.XPath, Using = PICTURE_ADDED_TO_RECIPE_PATCH)]
        private IWebElement _recipePicturepatch;

        [FindsBy(How = How.XPath, Using = PICTURE_ADDED_TO_RECIPE_DEV)]
        private IWebElement _recipePicturedev;

        [FindsBy(How = How.XPath, Using = DUPLICATE_RECIPE_DEV)]
        private IWebElement _duplicateRecipeDev;

        [FindsBy(How = How.XPath, Using = DUPLICATE_RECIPE_PATCH)]
        private IWebElement _duplicateRecipePatch;

        [FindsBy(How = How.XPath, Using = NET_WEIGHT)]
        private IWebElement _netWeight;

        [FindsBy(How = How.XPath, Using = PORTION)]
        private IWebElement _portion;

        [FindsBy(How = How.XPath, Using = COEF_RECIPE)]
        private IWebElement _coefRecipe;

        [FindsBy(How = How.XPath, Using = METHOD_RECIPE)]
        private IWebElement _methodRecipe;

        [FindsBy(How = How.XPath, Using = PAX_RECIPE)]
        private IWebElement _paxRecipe;

        [FindsBy(How = How.XPath, Using = POPUP_SCALE_QTY)]
        private IWebElement _popupScaleQty;

        [FindsBy(How = How.Id, Using = SCALE_SAVE)]
        private IWebElement _scaleSave;

        [FindsBy(How = How.CssSelector, Using = RECIPE_PRICE)]
        private IWebElement _recipePrice;

        [FindsBy(How = How.CssSelector, Using = TOTAL_COSTING)]
        private IWebElement _totalCosting;

        [FindsBy(How = How.Id, Using = SELECT_SITES)]
        private IWebElement _selectSites;

        // Intolerance
        [FindsBy(How = How.Id, Using = INTOLERANCE_TAB)]
        private IWebElement _intoleranceTab;

        [FindsBy(How = How.ClassName, Using = VALIDATE_ALLERGEN)]
        private IWebElement _validateAllergen;

        [FindsBy(How = How.ClassName, Using = NOT_VALIDATE_ALLERGEN)]
        private IWebElement _notValidateAllergen;


        // Picture
        [FindsBy(How = How.Id, Using = PICTURE_TAB)]
        private IWebElement _pictureTab;

        [FindsBy(How = How.Id, Using = UPLOAD_PICTURE)]
        private IWebElement _uploadPicture;

        [FindsBy(How = How.XPath, Using = DELETE_PICTURE)]
        private IWebElement _deletePicture;

        [FindsBy(How = How.Id, Using = CONFIRM_DELETE_PICTURE)]
        private IWebElement _confirmDeletePicture;


        // Use case
        [FindsBy(How = How.Id, Using = USE_CASE_TAB)]
        private IWebElement _useCaseTab;

        [FindsBy(How = How.XPath, Using = PICTUREADDED)]
        private IWebElement _pictureAdded;

        [FindsBy(How = How.XPath, Using = RECIPE_ANOTHER_SITE)]
        private IWebElement _recipeAnotherSite;
        [FindsBy(How = How.XPath, Using = RECIPE_TYPE_SELECTED)]
        private IWebElement _recipeTypeSelected;
        [FindsBy(How = How.XPath, Using = WORKSHOP_SELECTED)]
        private IWebElement _workshop;
        [FindsBy(How = How.XPath, Using = COOKING_MODE_SELECTED)]
        private IWebElement _cookingMode;
        // ______________________________________________ Méthodes ______________________________________________

        // Général
        public DatasheetPage BackToList()
        {
            _backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            _backToList.Click();
            WaitForLoad();

            return new DatasheetPage(_webDriver, _testContext);
        }

        public void OpenPrint()
        {
            _openprint = WaitForElementIsVisible(By.XPath(OPENPRINT));
            _openprint.Click();
            WaitForLoad();

         }

        public void Firstreciepclick()
        {
            _clickfirstline = WaitForElementIsVisible(By.XPath(CLICKFIRSTLINERECIPES));
            _clickfirstline.Click();
            WaitForLoad();

        }
        public void EditButton()
        {
            _backToList = WaitForElementIsVisible(By.XPath(EDITBUTTON));
            _backToList.Click();
            WaitForLoad();

        }
        public DatasheetGeneralInformationPage ClickOnGeneralInformation()
        {
            _generalInformationTab = WaitForElementIsVisible(By.Id(GENERAL_INFORMATION));
            _generalInformationTab.Click();
            WaitForLoad();

            return new DatasheetGeneralInformationPage(_webDriver, _testContext);
        }

        public PrintReportPage Print(bool versionPrint)
        {
            _print = WaitForElementIsVisible(By.Id(PRINT));
            _print.Click();
            WaitForLoad();

            _confirmPrint = WaitForElementToBeClickable(By.Id(CONFIRM_PRINT));
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

        /************************************************* Onglets **********************************************/

        // Recipe
        public void ClickOnRecipesTab()
        {
            _recipeTab = WaitForElementIsVisible(By.Id(RECIPE_TAB));
            _recipeTab.Click();
            // secousse
            Thread.Sleep(2000);
            WaitForLoad();
        }

        public void ShowDetails()
        {
            WaitPageLoading();
            WaitForLoad();
            _showDetailsDev = _webDriver.FindElement(By.Id(SHOW_DETAILS_DEV));
            _showDetailsDev.Click();
        }

        public void HideDetails()
        {
            WaitForLoad();
            _hideDetailsDev = WaitForElementIsVisible(By.XPath(HIDE_DETAILS_DEV));
            _hideDetailsDev.Click();

            // Temps que l'action se réalise
            WaitPageLoading();
            WaitForLoad();
        }

        public bool IsRecipeDetailsVisible()
        {
            WaitForLoad();
            var recipeDetails = SolveVisible("//*/div[@class='recipe-details']");
            return recipeDetails != null;
        }

        public DatasheetCreateNewRecipePage CreateNewRecipe()
        {
            _createNewRecipeDev = WaitForElementIsVisible(By.Id(CREATE_NEW_RECIPE_DEV));
            _createNewRecipeDev.Click();

            WaitForLoad();

            return new DatasheetCreateNewRecipePage(_webDriver, _testContext);
        }

        public void AddRecipe(string recipeName)
        {
            WaitPageLoading();
            WaitForLoad();

            Actions action = new Actions(_webDriver);
            _addRecipe = WaitForElementIsVisible(By.Id(ADD_RECIPE));
            _addRecipe.SetValue(ControlType.TextBox, recipeName);
            Thread.Sleep(1000);
            WaitForLoad();
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(10));
            _selectedRecipe = WaitForElementToBeClickable(By.XPath("//*[@id='recipe-list-result']/table/tbody/tr[1]/td[1][@class='click-to-add']"));
            _selectedRecipe.Click();

            //_selectedRecipe = WaitForElementIsVisible(By.XPath("//*[@id=\"recipe-list-result\"]/table/tbody/tr[1]/td[1][@class='click-to-add']"));
            //action.MoveToElement(_selectedRecipe).Perform();
            //_selectedRecipe.Click();
            WaitPageLoading();
            WaitForLoad();
        }

        public void AddRecipeFromDatasheet(string datasheetName)
        {
            _addRecipeFromDatasheet = WaitForElementIsVisible(By.Id(ADD_RECIPE_FROM_DATASHEET));
            _addRecipeFromDatasheet.SetValue(ControlType.TextBox, datasheetName);
            WaitPageLoading();

            _firstRecipeFromDatasheet = WaitForElementIsVisible(By.XPath(String.Format(FIRST_RECIPE_FROM_DATASHEET, datasheetName)));
            _firstRecipeFromDatasheet.Click();
            WaitPageLoading();

            _saveAddRecipeFromDatasheet = WaitForElementIsVisible(By.Id(SAVE_ADD_RECIPE_FROM_DATASHEET));
            _saveAddRecipeFromDatasheet.Click();
            WaitForLoad();
        }
        public void AddRecipeFromDatasheetToshowRecipe(string datasheetName)
        {
            _addRecipeFromDatasheet = WaitForElementIsVisible(By.Id(ADD_RECIPE_FROM_DATASHEET));
            _addRecipeFromDatasheet.SetValue(ControlType.TextBox, datasheetName);
            WaitPageLoading();

            _firstRecipeFromDatasheet = WaitForElementIsVisible(By.XPath(String.Format(FIRST_RECIPE_FROM_DATASHEET, datasheetName)));
            _firstRecipeFromDatasheet.Click();
            WaitPageLoading();

        }

        public bool IsRecipeAdded()
        {
            WaitPageLoading();
            //WaitForLoad();
            if (isElementVisible(By.XPath(ADDED_RECIPE)))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public bool AddRecipeError(string recipeName)
        {
            bool isAdded = true;
            WaitPageLoading();
            _addRecipe = WaitForElementIsVisible(By.Id(ADD_RECIPE));
            _addRecipe.SetValue(ControlType.TextBox, recipeName);
            WaitPageLoading();
            WaitPageLoading();
            WaitPageLoading();

            if (isElementVisible(By.XPath(ADD_RECIPE_ERROR_MSG)))
            {
                _addRecipeErrorMsg = WaitForElementIsVisible(By.XPath(ADD_RECIPE_ERROR_MSG));
                WaitForLoad();
                return _addRecipeErrorMsg.Text != "No recipe found";
               
            }
            else
            {
                return true;
            }

        //return _addRecipeErrorMsg.Text != "No recipe found";
        }

        public bool AddRecipeFromDatasheetError(string datasheetName)
        {
            bool isAdded = true;

            _addRecipeFromDatasheet = WaitForElementIsVisible(By.Id(ADD_RECIPE_FROM_DATASHEET));
            _addRecipeFromDatasheet.SetValue(ControlType.TextBox, datasheetName);
            WaitForLoad();
            WaitPageLoading();

            if (isElementVisible(By.Id(ADD_RECIPE_DATASHEET_ERROR_MSG)))
            {
                _addRecipeDatasheetErrorMsg = WaitForElementIsVisible(By.Id(ADD_RECIPE_DATASHEET_ERROR_MSG));
                WaitForLoad();
                if (_addRecipeDatasheetErrorMsg.Text == "No datasheet found")
                    isAdded = false;
            }
            else
            {
                isAdded = true;
            }

            return isAdded;
        }

        // Tableau recettes
        public string GetFirstRecipeName()
        {

            _firstRecipeName = WaitForElementIsVisible(By.Id(FIRST_RECIPE_NAME));

            return _firstRecipeName.Text;

        }
        public string getingredientname()
        {
            _firstRecipeName = WaitForElementIsVisible(By.XPath(INGREDIENTNAME));
            return _firstRecipeName.Text;
        }

        public void ClickOnFirstRecipe()
        {

            WaitPageLoading();
            _firstRecipeName = WaitForElementExists(By.Id(FIRST_RECIPE_NAME));
            _firstRecipeName.Click();
            WaitForLoad();
        }

        public DatasheetCreateNewRecipePage AddSubrecipeForFirstRecipe()
        {
            ClickOnFirstRecipe();
            if (IsDev())
            {
                _addSubrecipe = WaitForElementIsVisible(By.XPath(ADD_SUBRECIPE_DEV));
                _addSubrecipe.Click();
            }
            else
            {
                _addSubrecipePatch = WaitForElementIsVisible(By.Id(ADD_SUBRECIPE_PATCH));
                _addSubrecipePatch.Click();
            }
            WaitForLoad();
            return new DatasheetCreateNewRecipePage(_webDriver, _testContext);
        }

        public bool IsSubrecipeAdded()
        {

            bool isSubrecipeAdded = false;
            if (IsDev())
            {
                WaitPageLoading();
                if (isElementVisible(By.XPath(ADDED_SUBRECIPE)))
                {
                    var element = _webDriver.FindElement(By.XPath(ADDED_SUBRECIPE));
                    if (element.GetAttribute("src").Contains("/images/icons/icon-subrecipe.png"))
                        isSubrecipeAdded = true;
                }
                else
                {
                    isSubrecipeAdded = false;
                }
            }
            else
            {
                WaitPageLoading();
                WaitForLoad();
                if (SolveVisible("//*/div[@class='row']/div[1]/img[@src='/images/icons/icon-subrecipe.png']") != null)
                {
                    isSubrecipeAdded = true;
                }
            }

            return isSubrecipeAdded;
        }

        public RecipesVariantPage EditRecipe()
        {
            ClickOnFirstRecipe();
            WaitForLoad();
            if (IsDev()) _editRecipeDev = WaitForElementIsVisible(By.Id(EDIT_RECIPE_DEV));
            else _editRecipeDev = WaitForElementIsVisible(By.XPath(EDIT_RECIPE_PATCH));
            //else _editRecipeDev = WaitForElementIsVisible(By.Id("inflightrawmaterials-production-service-detail-1"));
            _editRecipeDev.Click();
            WaitPageLoading();
            WaitLoading();
            return new RecipesVariantPage(_webDriver, _testContext);
        }
        public RecipesVariantPage EditFirstRecipe()
        {
            WaitPageLoading();
            _editRecipeDev = WaitForElementExists(By.Id("menus-datasheeteditdatasheetrecipe-display-1"));
            new Actions(_webDriver).MoveToElement(_editRecipeDev).Click().Perform();
            WaitPageLoading();
            return new RecipesVariantPage(_webDriver, _testContext);
        }
        public RecipesVariantPage EditFirstRecipeItem()
        {
            if (IsDev())
                _editRecipeDev = WaitForElementExists(By.Id("menus-datasheeteditdatasheetrecipe-display-1"));
            else
                _editRecipeDev = WaitForElementExists(By.XPath("//*[starts-with(@id,\"inflightrawmaterials-production-service-detail-\")]"));
            new Actions(_webDriver).MoveToElement(_editRecipeDev).Click().Perform();
            WaitPageLoading();
            return new RecipesVariantPage(_webDriver, _testContext);
        }

        public void DeleteRecipe()
        {
            WaitPageLoading();
            WaitForLoad();
            ClickOnFirstRecipe();
            if (IsDev()) _deleteRecipeDev = WaitForElementIsVisible(By.XPath(DELETE_RECIPE_DEV));
            else _deleteRecipeDev = WaitForElementIsVisible(By.XPath(DELETE_RECIPE_PATCH));

            _deleteRecipeDev.Click();
            WaitForLoad();

            _confirmDeleteRecipe = WaitForElementIsVisible(By.Id(CONFIRM_DELETE_RECIPE));
            _confirmDeleteRecipe.Click();
            WaitForLoad();
        }

        public RecipesVariantPage AddPictureToRecipe()
        {
            ClickOnFirstRecipe();
            _addPicturedev = WaitForElementIsVisible(By.XPath(ADD_PICTURE_DEV));
            _addPicturedev.Click();
            WaitForLoad();

            return new RecipesVariantPage(_webDriver, _testContext);
        }

        public bool IsPictureAddedToRecipe()
        {
            WaitForLoad();
            _recipePicturedev = WaitForElementIsVisible(By.XPath(PICTURE_ADDED_TO_RECIPE_DEV));
            if (_recipePicturedev.GetAttribute("alt").Equals("Has a picture"))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public RecipesVariantPage DuplicateRecipe()
        {
            ClickOnFirstRecipe();
            _duplicateRecipeDev = WaitForElementIsVisible(By.XPath(DUPLICATE_RECIPE_DEV));
            _duplicateRecipeDev.Click();
            WaitForLoad();

            //Results are opened in a new tab, switch the driver to the newly created one
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            WaitForLoad();

            return new RecipesVariantPage(_webDriver, _testContext);
        }

        public void DuplicateRecipeError()
        {
            ClickOnFirstRecipe();
            _duplicateRecipeDev = WaitForElementIsVisible(By.XPath(DUPLICATE_RECIPE_DEV));
            _duplicateRecipeDev.Click();
            WaitForLoad();
        }

        public bool IsPopUpDuplicateErrorVisible()
        {
            if (isElementVisible(By.XPath(DUPLICATE_ERROR_MODALE)))
            {
                _webDriver.FindElement(By.XPath(DUPLICATE_ERROR_MODALE));
                return true;
            }
            else
            {
                return false;
            }
        }

        public double GetNetWeight(string decimalSeparator)
        {
            WaitPageLoading();
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            _netWeight = WaitForElementIsVisible(By.XPath(NET_WEIGHT));
            var total = _netWeight.Text;

            return Convert.ToDouble(total, ci);
        }

        public string GetCoef()
        {
            _coefRecipe = WaitForElementIsVisible(By.XPath(COEF_RECIPE));
            return _coefRecipe.GetAttribute("value");
        }

        public bool SetCoef(string coef)
        {
            if (isElementVisible(By.XPath(COEF_RECIPE)))
            {
                _coefRecipe = WaitForElementIsVisible(By.XPath(COEF_RECIPE));
                if (_coefRecipe.GetAttribute("disabled") == "true")
                {
                    return false;
                }
                _coefRecipe.SetValue(ControlType.TextBox, coef);
                WaitPageLoading();
                return true;
            }
            else
            {
                return false;
            }

        }

        public string GetMethod()
        {
            var elements = _webDriver.FindElements(By.XPath(METHOD_RECIPE_SELECTED));
            foreach (var element in elements)
            {
                if (element.GetAttribute("selected") != null)
                {
                    return element.Text;
                }
            }

            return null;
        }

        public void SetMethod(string method)
        {

            _methodRecipe = WaitForElementIsVisible(By.XPath(METHOD_RECIPE));
            _methodRecipe.SetValue(ControlType.DropDownList, method);
            WaitLoading();
        }

        public string GetPAXRecipe()
        {
            _paxRecipe = WaitForElementIsVisible(By.XPath(PAX_RECIPE));
            return _paxRecipe.GetAttribute("innerText");
        }

        public void FillPopUpScale()
        {
            _popupScaleQty = WaitForElementIsVisible(By.XPath(POPUP_SCALE_QTY));
            _popupScaleQty.SendKeys("1");
            WaitForLoad();

            _scaleSave = WaitForElementIsVisible(By.Id(SCALE_SAVE));
            _scaleSave.Click();
            WaitForLoad();
        }

        // Intolerance
        public DatasheetIntolerancePage ClickOnIntoleranceTab()
        {
            _intoleranceTab = WaitForElementIsVisible(By.Id(INTOLERANCE_TAB));
            _intoleranceTab.Click();
            return new DatasheetIntolerancePage(_webDriver, _testContext);
        }

        public void ValidateAllergen()
        {
            // On clique sur le bouton 'Not Validated' pour passer à l'état Validé
            _notValidateAllergen = WaitForElementIsVisible(By.ClassName(NOT_VALIDATE_ALLERGEN));
            _notValidateAllergen.Click();
            WaitForLoad();
        }

        public void DevalidateAllergen()
        {
            // On clique sur le bouton 'Validated' pour passer à l'état non validé
            _validateAllergen = WaitForElementIsVisible(By.ClassName(VALIDATE_ALLERGEN));
            _validateAllergen.Click();
            WaitForLoad();
        }

        public bool IsAllergenValidated()
        {
            bool isValidated = true;

            try
            {
                if (isElementVisible(By.Id(ALLERGEN_VALIDATOR)))
                {
                    _webDriver.FindElement(By.Id(ALLERGEN_VALIDATOR));
                    _webDriver.FindElement(By.Id(ALLERGEN_VALIDATION_DATE));
                }
                else
                {
                    //PATCH 24Aug
                    _webDriver.FindElement(By.Id("AllergenValidatedBy"));
                    _webDriver.FindElement(By.Id("AllergenValidatedDate"));
                }
            }
            catch
            {
                isValidated = false;
            }

            return isValidated;
        }

        // Picture
        public void ClickOnPictureTab()
        {
            _pictureTab = WaitForElementIsVisible(By.Id(PICTURE_TAB));
            _pictureTab.Click();
            WaitForLoad();
        }

        public void AddPicture(string imagePath)
        {
            _uploadPicture = WaitForElementIsVisible(By.Id(UPLOAD_PICTURE));
            _uploadPicture.SendKeys(imagePath);
            WaitForLoad();
            WaitForElementIsVisible(By.XPath(DELETE_PICTURE_ADDED));
        }

        public bool IsPictureAdded()
        {
            WaitForLoad();
            if (isElementVisible(By.XPath(PICTURE)))
            {
                _webDriver.FindElement(By.XPath(PICTURE));
                return true;
            }
            else
            {
                return false;
            }
        }

        public void DeletePicture()
        {
            _deletePicture = WaitForElementIsVisible(By.XPath(DELETE_PICTURE));
            _deletePicture.Click();
            WaitForLoad();

            _confirmDeletePicture = WaitForElementIsVisible(By.Id(CONFIRM_DELETE_PICTURE));
            _confirmDeletePicture.Click();
            WaitForLoad();
        }

        // Use Case
        public void ClickOnUseCaseTab()
        {
            _useCaseTab = WaitForElementIsVisible(By.Id(USE_CASE_TAB));
            _useCaseTab.Click();
            WaitForLoad();
        }

        public void ClickOnDetailsTab()
        {
            var detailsTab = WaitForElementIsVisible(By.Id("hrefTabContentRecipes"));
            detailsTab.Click();
            WaitForLoad();
        }

        public ServicePricePage SelectService(string serviceName)
        {
            Actions action = new Actions(_webDriver);

            var serviceLine = _webDriver.FindElement(By.XPath(String.Format(SERVICE_LINE, serviceName)));

            action.MoveToElement(serviceLine).Perform();
            serviceLine.Click();

            var serviceEdit = WaitForElementIsVisible(By.XPath(String.Format(SERVICE_EDIT, serviceName)));
            serviceEdit.Click();
            WaitForLoad();

            //Results are opened in a new tab, switch the driver to the newly created one
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            WaitForLoad();

            return new ServicePricePage(_webDriver, _testContext);
        }
        public RecipesVariantPage EditSubRcipe()
        {
            IWebElement editSubRecipe;
            editSubRecipe = WaitForElementIsVisible(By.XPath(EDIT_SUB_RECIPE_DEV));
            editSubRecipe.Click();

            return new RecipesVariantPage(_webDriver, _testContext);
        }
        public void SelectSite(string site)
        {
            WaitPageLoading();
            if (isElementVisible(By.Id(SELECT_SITES)))
            {
                _selectSites = WaitForElementIsVisible(By.Id(SELECT_SITES));
                _selectSites.SetValue(ControlType.DropDownList, site);
                WaitPageLoading();
            }
            else
            {
                ComboBoxSelectById(new ComboBoxOptions(SELECT_SITES, site));
            }
        }

        public double GetRecipePrice(string decimalSeparator)
        {
            // Récupération du type de séparateur (, ou . selon les pays)
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            _recipePrice = WaitForElementIsVisible(By.CssSelector(RECIPE_PRICE));
            return double.Parse(_recipePrice.Text, ci);
        }

        public double GetTotalCosting(string decimalSeparator)
        {
            // Récupération du type de séparateur (, ou . selon les pays)
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            _totalCosting = WaitForElementIsVisible(By.CssSelector(TOTAL_COSTING));
            return double.Parse(_totalCosting.Text, ci);
        }
        public RecipesVariantPage EditRecipe(string recipeName)
        {
            IReadOnlyCollection<IWebElement> recipes = _webDriver.FindElements(By.Id(FIRST_RECIPE_NAME));
            IReadOnlyCollection<IWebElement> editrecipes = _webDriver.FindElements(By.Id(EDIT_RECIPE_DEV));
            int i = 0;
            foreach (IWebElement recipe in recipes)
            {
                string recipeText = recipe.Text;
                if (recipeText == recipeName)
                {
                    recipe.Click();
                    break;
                }
                i++;
            }
            if (editrecipes.Count > 0)
                editrecipes.ElementAt(i).Click();
            WaitForLoad();

            WaitPageLoading();
            return new RecipesVariantPage(_webDriver, _testContext);
        }
        public string GetPortion()
        {
            _portion = WaitForElementIsVisible(By.CssSelector(PORTION));
            return _portion.Text;
        }
        public int CountUseCases()
        {
            var useCases = _webDriver.FindElements(By.XPath(DS_USE_CASES));
            return useCases.Count();
        }
        public void Go_To_New_Navigate(int ongletChrome = 1)
        {
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > ongletChrome);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[ongletChrome]);
        }

        public bool isPopupClosed()
        {
            if (isElementVisible(By.Id("tabContentDetails-sub")))
            {
                return false;
            }
            if (isElementVisible(By.Id("recipe-variant-detail-list")))
            {
                return false;
            }
            return true;
        }
        public void CloseDetailModal()
        {
            var _close = WaitForElementIsVisible(By.XPath("//*[@id=\"modal-1\"]/div[1]/button"));
            _close.Click();
        }
        public bool IsNewSubRecipeAdded(string subRecipe)
        {
            WaitPageLoading();
            var elements = _webDriver.FindElements(By.XPath(SUB_RECIPES));

            return elements.Any(elm => elm.Text.Trim().Equals(subRecipe, StringComparison.OrdinalIgnoreCase));
        }


        public bool IsPictureExist()
        {
            _pictureAdded = WaitForElementIsVisible(By.XPath(PICTUREADDED));
            WaitForLoad();

            if (_pictureAdded != null)
            {

                return true;
            }
            else
            {
                return false;
            }
        }

        public bool CheckRecipeExist(string recipeName)
        {
            var recipes = _webDriver.FindElements(By.XPath(RECIPE_ANOTHER_SITE));
            return recipes.Any(r => r.Text == recipeName);
        }

        public List<(string, int)> GetIngredientNamesDatasheet()
        {
            var jsExecutor = (IJavaScriptExecutor)_webDriver;

            var lineDetails = WaitForElementIsVisible(By.XPath("//*/li/div/div[2]/div[1]/div[2]"));
            lineDetails.Click();
            var toggleButton = WaitForElementExists(By.XPath("//span[@class='fas fa-angle-right']"));
            toggleButton.Click();
            WaitPageLoading();
            var ingredients = _webDriver.FindElements(By.XPath("//*/li/div/div[1]/div[2]/div[position() >= 1]/div[1]"));
            var ingredientNamesDatasheet = new List<(string, int)>();

            // Parcourir chaque élément <li> pour récupérer le nom de l'ingrédient
            var i = 1;
            foreach (var ingredient in ingredients)
            {
                string ingredientName = ingredient.Text.ToUpper().Replace(" ", ""); ;
                ingredientNamesDatasheet.Add((ingredientName, i));
                i++;
            }
            return ingredientNamesDatasheet;
        }
        public string GetRecipeType()
        {
            _recipeTypeSelected = WaitForElementIsVisible(By.XPath(RECIPE_TYPE_SELECTED));
            return _recipeTypeSelected.Text;
        }
        public string GetWorkshop()
        {
            _workshop = WaitForElementIsVisible(By.XPath(WORKSHOP_SELECTED));
            return _workshop.Text;
        }
        public string GetCookingMode()
        {
            _cookingMode = WaitForElementIsVisible(By.XPath(COOKING_MODE_SELECTED));
            return _cookingMode.Text;
        }
        public void CloseWindow(string originalWindow)
        {
            
            IList<string> allWindows = _webDriver.WindowHandles;

            // Switch to the new window
            foreach (string window in allWindows)
            {
                if (window != originalWindow)
                {
                    _webDriver.SwitchTo().Window(window);
                    break;
                }
            }
            _webDriver.Close();

            // Switch back to the original window
            _webDriver.SwitchTo().Window(originalWindow);
        }
        public bool IsTextEqualNone()
        {
            try
            {
                var element = _webDriver.FindElement(By.XPath(NONE));
                return element.Displayed && element.Text.Trim().Equals("none", StringComparison.OrdinalIgnoreCase);
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }
        public string GetPrice()
        {
            WaitForLoad();

            var element = WaitForElementIsVisible(By.XPath(PRICE));

            return element.Text;
        }


        public void CloseDetail()
        {
            var _close = WaitForElementIsVisible(By.Id("btn-from-datasheet-close-modal"));
            _close.Click();
        }

    }
}
