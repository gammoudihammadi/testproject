using DocumentFormat.OpenXml.Bibliography;
using iText.Commons.Utils;
using iText.StyledXmlParser.Jsoup.Nodes;
using Microsoft.VisualStudio.TestTools.UnitTesting;
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

namespace Newrest.Winrest.FunctionalTests.PageObjects.Menus.Menus
{
    public class MenusDayViewPage : PageBase
    {
        public MenusDayViewPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // ______________________________________ Constantes _____________________________________________

        // Général
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";
        private const string GENERAL_INFORMATION = "-1";
        private const string WEEK_VIEW = "//*/a[contains(@class, 'btn-go-to-week')]";
        private const string EXPORT_TO_OTHER_MENU = "//*/a[contains(@id, 'exportDataSheet')]";
        private const string ADD_MULTIPLE_MENUS_DESTINATIONS = "//*[@id=\"btn-add-mutiple-menus-destinations\"]";

        //Add multiple menu destinations
        private const string SITE = "collapseMenusDestinationsSelectedSitesIdsFilter";
        private const string MENU = "collapseMenuDestinationSelectedMenuIdsFilter";
        private const string SAVE = "//*[@id=\"btn-add-mutiple-destination\"]";
        private const string EXPORT = "//*[@id=\"last\"]";
        private const string OK_BUTTON = "//*[@id=\"ExportSelectForm\"]/div[2]/button";
        private const string PRINT = "//*[@id=\"div-body\"]/div/div[2]/div[1]/div/div/div/a[3]";

        // Menu latéral
        private const string SHOW_MENU_DAY = "datepicker-show-menu-day";
        private const string DAYS_PLANNING = "//*[starts-with(@id,\"MenuDay_\")]/td/h4";
        private const string PLAN_VIEW = "displayMenuCheckbox";
        private const string PRICER_DEVICES = "//*[@id=\"menu-central-Pane\"]/a[2]";
        private const string UPDATE_MENU_BTN = "//*[@id=\"btn-update-menu\"]";

        // Add recipe
        private const string ADD_RECIPE = "RecipeName";
        private const string KIOSK_IMAGE = "/html/body/div[3]/div/div[2]/div[3]/div/div[2]/div[2]/ul/li/div/div[2]/div[4]/span[starts-with(@id,\"kiosque\")]";

        private const string SELECTED_KIOSK_IMAGE = "//*[starts-with(@id,\"popover\")]/div[2]/div/div/div/div/img[1]";
        private const string SELECTED_KIOSK = "//*[starts-with(@id,\"popover\")]/div[2]/div/div/div/div/h3[1]";
        private const string KIOSK_NAME = "//*[starts-with(@id,\"dataItem\")]/div/div[1]/div[3]/span[text()='{0}']";

        // Tableau recettes
        private const string DAY_BEFORE = "previous-date-button";
        private const string DAY_DISPLAY = "menu-day-display";
        private const string DAY_AFTER = "next-date-button";
        private const string THEORICAL_PAX_DAY = "TheoricalPaxNumberDay";
        private const string DUPLICATE_DAY = "//a[contains(text(), 'Duplicate this day')]";
        private const string CONFIRM_DUPLICATE = "//*[@id=\"modal-1\"]/div[1]/div/div[2]/div/form/div[2]/button[2]";

        private const string RECIPE_NAME = "//*[starts-with(@id,\"dataItem\")]/div/div[2]/div[1]/div[2]/p[1]";
        private const string FIRST_RECIPE_NAME_ONDISPLAYVIEW = "/html/body/div[3]/div/div[2]/div[3]/div/div[2]/div[2]/ul/li[1]/div/div[2]/div[1]/div[2]/p[1]";
        private const string RECIPE_NAME_ONDISPLAYVIEW = "/html/body/div[3]/div/div[2]/div[3]/div/div[2]/div[2]/ul/li/div/div[2]";
        private const string RECIPE_NAME_INPUT_ON_DISPLAY_VIEW = "//*[starts-with(@id,\"dataItem\")]/div/div[1]/div[1]/div[2]/input";

        private const string THIS_RECIPE_NAME = "//p[contains(text(),'{0}')]";
        private const string METHOD = "//*[starts-with(@id,\"dataItem\")]/div/div[2]/div[3]/span";
        private const string METHOD_SELECTED = "//*[starts-with(@id,\"dataItem\")]/div/div[1]/div[3]/select";
        private const string COEF = "//*[starts-with(@id,\"dataItem\")]/div/div[2]/div[4]/span";
        private const string PORTION = "//*[starts-with(@id,\"dataItem\")]/div/div[2]/div[5]/span[2]";
        private const string PAX = "//*[starts-with(@id,\"dataItem\")]/div/div[2]/div[6]/span[2]";
        private const string PRIXTOTAL = "//*[@id=\"datasheet-list\"]/div[2]/div/div[2]/div/span[2]";
        private const string COEF_SELECTED = "//*[starts-with(@id,\"inputDetailCoeff-\")]";
        private const string EDIT_FIRST_RECIPE = "//*[starts-with(@id,\"dataItem\")]/div/div[1]/div[9]/div/a[1]/span";
        private const string DELETE_FIRST_RECIPE = "//span[contains(@class,'glyphicon-trash')]";
        private const string CONFIRM_DELETE = "dataConfirmOK";
        private const string MENU_INFORMATIONS = "/html/body/div[3]/div/div[2]/div[2]/h2/span[2]";
        private const string ALL_METHOD = "//div[contains(@class,\"display-zone\")]/div[3][contains(@class,\"item-detail-col\")]";
        private const string VARIANTMENU_XPATH = "//*[@id=\"createFormMenu\"]/div[7]/div";

        
        // Display menu
        private const string DISPLAY_MENU = ".btn-display-menu";
        private const string DISPLAY_MENU_CHECK_SERVICE = "SelectedDeliveriesIds";
        private const string DISPLAY_MENU_VALIDATE = "display-menu-generateUrl-btn";

        private const string ELEMENTS_TO_ADD_LIST = "//*[@id=\"recipe-list-result\"]/table//tr";
        private const string FILTER_DIETARY_TYPE = "//*[@id=\"SelectedDietaryTypes_ms\"]";
        private const string SEARCH_DIETARY_TYPE_INPUT = "/html/body/div[12]/div/div/label/input";
        private const string SELECT_FIRST_SERACH_RESULT = "//*[@id=\"ui-multiselect-1-SelectedDietaryTypes-option-1\"]";
        // Right Menu -- Filter --
        private const string LIST_RESULT_RECIPE_NAME = "/html/body/div[3]/div/div[2]/div[3]/div/div[3]/div[2]/table/tbody/tr[*]/td[2]/p[1]";
        private const string RECIPE_TYPES_FILTER = "SelectedRecipeTypes_ms";
        private const string SEARCH_RECIPE_TYPES = "/html/body/div[11]/div/div/label/input";
        private const string UNCHECK_ALL_RECIPE_TYPES = "/html/body/div[11]/div/ul/li[2]/a";
        private const string CHECK_ALL_RECIPE_TYPES = "/html/body/div[11]/div/ul/li[1]/a";
        private const string CHECK_ALL_ALLERGEN = "//*[starts-with(@id, 'popover')]/div[2]/div/div/div/table/tbody/tr[*]/td[*]/div/figure/figcaption[text()='Egg']";
        private const string ALLERGEN_ETOILE = "//*[@id=\"menu-dataSheet-display-id\"]/div[2]/ul/li//div/div[2]/div[5]/span/i";
        private const string KIOSQUE_IMAGE = "//*[@id=\"menu-dataSheet-display-id\"]/div[2]/ul/li//div/div[1]/div[4]/span/i"; 
        private const string KEYWORD = "//*[@id=\"menu-dataSheet-display-id\"]/div[2]/ul/li//div/div[1]/div[6]/span/i";
        private const string ICON_KEYWORDS_RECIPE = "/html/body/div[3]/div/div[2]/div[3]/div/div[2]/div[2]/ul/li[1]/div/div[1]/div[6]/span";
        private const string KEYWORDS_RECIPE = "//*[starts-with(@id, 'popover')]/div[2]/div/div/div/table/tbody/tr[*]/td[*]/div/figure[@class='captcha checked']";
        private const string KEYWORD_CHECKED = "//*[starts-with(@id, 'popover')]/div[2]/div/div/div/table/tbody/tr/td/div/figure/figcaption";
        private const string KEYWORD_CHECK = "//*[starts-with(@id,'popover')]/div[2]/div/div/div/table/tbody/tr[*]/td[*]/div/figure/figcaption[text()='{0}']/..";
        private const string KEYWORDS = "//*[starts-with(@id,\"keywords\")]";
        private const string ALLERGEN_TEXT = "//*[starts-with(@id, 'popover')]/div[2]/div/div/div/table/tbody/tr[1]/td[2]/div/figure/figcaption";
        private const string CHECK_ALLERGEN = "//*[starts-with(@id, 'popover')]/div[2]/div/div/div/table/tbody/tr[1]/td[2]/div/figure/figcaption";


        private const string LIST_DAYS = "/html/body/div[3]/div/div[1]/table[2]/tr[*]";

        private const string COEF_VALID = "//*[@id=\"qtyChangeAll\"]";

        private const string CONFIRMATIONCOEF = "//*[@id=\"btnApplyQuantityChange\"]";

        private const string CONFIRMATIONMESSAGE = "//*[@id=\"dataConfirmModal\"]/div/div/div[2]";

         private const string BTN_CONFIRMATIONMESSAGE = "//*[@id=\"dataConfirmOK\"]";

        private const string CHECK_FIRST_ALLERGEN = "//*[starts-with(@id, 'popover')]/div[2]/div/div/div/table/tbody/tr[1]/td[1]/div/figure/img";
        private const string CHECK_FIRST_KIOSQUE = "//*[starts-with(@id,\"popover\")]/div[2]/div/div/div/div/img[1]";
        private const string RECIPE_NAME_OFFDISPLAYVIEW = "/html/body/div[3]/div/div[2]/div[3]/div/div[2]/div[1]/ul/li/div/div[2]/div[1]/div[2]/p[1]";
        private const string MESSAGE_DISPLAY = "menuDisplay-list-text";
        private const string KEYWORD_DISPLAY = "keywords290690";
        private const string MESSAGEPLANVIEW = "datasheet-list-text";
        private const string FIRSTLINEPLANVIEW = "/html/body/div[3]/div/div[2]/div[3]/div/div[2]/div[1]/ul/li/div/div[2]";
        private const string FIRSTLINEDISPLAYVIEW = "/html/body/div[3]/div/div[2]/div[3]/div/div[2]/div[2]/ul/li/div/div[2]";



        // ______________________________________ Variables ______________________________________________

        // Général
        [FindsBy(How = How.XPath, Using = BTN_CONFIRMATIONMESSAGE)]
        private IWebElement _btnconfirmationmessage;

        [FindsBy(How = How.XPath, Using = CONFIRMATIONMESSAGE)]
        private IWebElement _confirmationmessage;

        [FindsBy(How = How.XPath, Using = CONFIRMATIONCOEF)]
        private IWebElement _confirmationcoef;

        [FindsBy(How = How.XPath, Using = COEF_VALID)]
        private IWebElement _coefvalid;


        [FindsBy(How = How.XPath, Using = KEYWORD_CHECK)]
        private IWebElement  _checkkeyword;

        [FindsBy(How = How.XPath, Using = CHECK_FIRST_KIOSQUE)]
        private IWebElement _checkkiosque;

        [FindsBy(How = How.XPath, Using = KIOSQUE_IMAGE)]
        private IWebElement _kiosqueimage;

        [FindsBy(How = How.XPath, Using = KEYWORD)]
        private IWebElement _keyword;

        [FindsBy(How = How.XPath, Using = CHECK_ALL_ALLERGEN)]
        private IWebElement _checkallallergen;

        [FindsBy(How = How.XPath, Using = ALLERGEN_ETOILE)]
        private IWebElement _allergenetoile;

        [FindsBy(How = How.XPath, Using = PORTION)]
        private IWebElement _recipePorton;

        [FindsBy(How = How.XPath, Using = PAX)]
        private IWebElement _recipePax;

        [FindsBy(How = How.XPath, Using = PRIXTOTAL)]
        private IWebElement _recipePrixtotal;


        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;

        [FindsBy(How = How.Id, Using = GENERAL_INFORMATION)]
        private IWebElement _generalInformation;

        [FindsBy(How = How.XPath, Using = WEEK_VIEW)]
        private IWebElement _weekView;

        // Menu latéral
        [FindsBy(How = How.Id, Using = SHOW_MENU_DAY)]
        private IWebElement _showMenuDay;

        // Add recipe
        [FindsBy(How = How.Id, Using = ADD_RECIPE)]
        private IWebElement _addRecipe;

        [FindsBy(How = How.XPath, Using = KIOSK_IMAGE)]
        private IWebElement _kioskImage;

        [FindsBy(How = How.XPath, Using = SELECTED_KIOSK_IMAGE)]
        private IWebElement _selectedKioskImage;
        [FindsBy(How = How.XPath, Using = SELECTED_KIOSK)]
        private IWebElement _selectedKiosk;

        [FindsBy(How = How.XPath, Using = KIOSK_NAME)]
        private IWebElement _kioskName;

        [FindsBy(How = How.XPath, Using = KEYWORD_CHECKED)]
        private IWebElement _keywordCheck;

        [FindsBy(How = How.XPath, Using = KEYWORDS)]
        private IWebElement _keywordS;

        //Change Recipe Method
        private const string METHOD_RECIPE = "//*[@id=\"methodChangeAll\"]";
        private const string METHOD_RECIPE_CHANGE_BUTTON = "//*[@id=\"btnApplyMethodChange\"]";

        // Tableau recettes
        [FindsBy(How = How.Id, Using = DAY_BEFORE)]
        private IWebElement _dayBefore;

        [FindsBy(How = How.Id, Using = DAY_DISPLAY)]
        private IWebElement _dayDisplay;

        [FindsBy(How = How.Id, Using = DAY_AFTER)]
        private IWebElement _dayAfter;

        [FindsBy(How = How.Id, Using = THEORICAL_PAX_DAY)]
        private IWebElement _theoricalPAXDay;

        [FindsBy(How = How.XPath, Using = DUPLICATE_DAY)]
        private IWebElement _duplicateDay;

        [FindsBy(How = How.XPath, Using = CONFIRM_DUPLICATE)]
        private IWebElement _confirmDuplicate;

        [FindsBy(How = How.XPath, Using = RECIPE_NAME)]
        private IWebElement _firstRecipeName;

        [FindsBy(How = How.XPath, Using = RECIPE_NAME)]
        private IWebElement _variantmenu;
        
        [FindsBy(How = How.XPath, Using = FIRST_RECIPE_NAME_ONDISPLAYVIEW)]
        private IWebElement _firstRecipeNameDisplayView;

        [FindsBy(How = How.XPath, Using = METHOD)]
        private IWebElement _recipeMethod;

        [FindsBy(How = How.XPath, Using = METHOD_SELECTED)]
        private IWebElement _recipeMethodSelected;

        [FindsBy(How = How.XPath, Using = COEF)]
        private IWebElement _recipeCoef;

        [FindsBy(How = How.XPath, Using = COEF_SELECTED)]
        private IWebElement _recipeCoefSelected;

        [FindsBy(How = How.XPath, Using = EDIT_FIRST_RECIPE)]
        private IWebElement _editFirstRecipe;

        [FindsBy(How = How.XPath, Using = DELETE_FIRST_RECIPE)]
        private IWebElement _deleteFirstRecipe;

        [FindsBy(How = How.Id, Using = CONFIRM_DELETE)]
        private IWebElement _confirmDelete;

        [FindsBy(How = How.Id, Using = PLAN_VIEW)]
        private IWebElement _planView;

        [FindsBy(How = How.XPath, Using = PRICER_DEVICES)]
        private IWebElement _pricerDevices;

        [FindsBy(How = How.XPath, Using = UPDATE_MENU_BTN)]
        private IWebElement _updateMenuBtn;

        [FindsBy(How = How.XPath, Using = PRINT)]
        private IWebElement _print;

        [FindsBy(How = How.CssSelector, Using = DISPLAY_MENU)]
        private IWebElement _displayMenu;

        [FindsBy(How = How.Id, Using = DISPLAY_MENU_CHECK_SERVICE)]
        private IWebElement _displayMenuCheckService;

        [FindsBy(How = How.Id, Using = DISPLAY_MENU_VALIDATE)]
        private IWebElement _displayMenuValidate;

        // Right Filter

        [FindsBy(How = How.XPath, Using = LIST_RESULT_RECIPE_NAME)]
        private IWebElement _recipeNameListResult;

        [FindsBy(How = How.Id, Using = RECIPE_TYPES_FILTER)]
        private IWebElement _recipe_types_filter;

        [FindsBy(How = How.XPath, Using = SEARCH_RECIPE_TYPES)]
        private IWebElement _searchRecipeType;

        [FindsBy(How = How.XPath, Using = CHECK_ALL_RECIPE_TYPES)]
        private IWebElement _checkAllRecipeTypes;

        [FindsBy(How = How.XPath, Using = CHECK_ALL_RECIPE_TYPES)]
        private IWebElement _unCheckAllRecipeTypes;

        [FindsBy(How = How.XPath, Using = METHOD_RECIPE)]
        private IWebElement _allMethodRecipe;

        [FindsBy(How = How.XPath, Using = ICON_KEYWORDS_RECIPE)]
        private IWebElement _iconKeywords;

        [FindsBy(How = How.XPath, Using = METHOD_RECIPE_CHANGE_BUTTON)]
        private IWebElement _allMethodRecipeChangeButton;

        [FindsBy(How = How.XPath, Using = CHECK_FIRST_ALLERGEN)]
        private IWebElement _checkFirstallergen;

        [FindsBy(How = How.XPath, Using = RECIPE_NAME_OFFDISPLAYVIEW)]
        private IWebElement _recipeNameOffDisplayView;

        [FindsBy(How = How.Id, Using = MESSAGE_DISPLAY)]
        private IWebElement _messageDisplay;

        [FindsBy(How = How.Id, Using = KEYWORD_DISPLAY)]
        private IWebElement _keywordDisplay;

        [FindsBy(How = How.Id, Using = MESSAGEPLANVIEW)]
        private IWebElement _messagePlanView;

        [FindsBy(How = How.Id, Using = FIRSTLINEPLANVIEW)]
        private IWebElement _firstLinePlanView;

        [FindsBy(How = How.Id, Using = FIRSTLINEDISPLAYVIEW)]
        private IWebElement _firstLineDisplayView;


        // ______________________________________ Méthodes _______________________________________________

        // Général

        public MenusPage BackToList()
        {
            _backToList = WaitForElementIsVisibleNew(By.XPath(BACK_TO_LIST));
            _backToList.Click();
            WaitForLoad();

            return new MenusPage(_webDriver, _testContext);
        }

        public MenusGeneralInformationPage ClickOnGeneralInformation()
        {
            _generalInformation = WaitForElementIsVisible(By.Id(GENERAL_INFORMATION));
            _generalInformation.Click();
            WaitForLoad();

            return new MenusGeneralInformationPage(_webDriver, _testContext);
        }

        public MenusWeekViewPage SwitchToWeekView()
        {
            ShowExtendedMenu();

            _weekView = WaitForElementIsVisible(By.XPath(WEEK_VIEW));
            _weekView.Click();
            WaitForLoad();

            return new MenusWeekViewPage(_webDriver, _testContext);
        }

        // Menu latéral
        public void SetShowMenuDay(DateTime date)
        {
            _showMenuDay = WaitForElementIsVisible(By.Id(SHOW_MENU_DAY));
            _showMenuDay.SetValue(ControlType.DateTime, date);
            _showMenuDay.SendKeys(Keys.Enter);
            WaitForLoad();
        }

        public bool SetMenuPlanningOfMonth(DateTime date)
        {
            bool isDayFound = false;

            Actions action = new Actions(_webDriver);
            var dateFormat = _showMenuDay.GetAttribute("data-date-format");
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            var menuDays = _webDriver.FindElements(By.XPath(DAYS_PLANNING));

            foreach (var element in menuDays)
            {
                string dayText;
                dayText = element.Text;
                dayText = dayText.Substring("Day ".Length, 10);
                DateTime day = DateTime.Parse(dayText, ci).Date;

                if (day.Equals(date.Date))
                {
                    action.MoveToElement(element).Perform();
                    element.Click();
                    WaitForLoad();
                    isDayFound = true;
                    break;
                }
            }

            return isDayFound;
        }
        public void ChangeRecipeMethod(string method)
        {
            _allMethodRecipe = WaitForElementIsVisible(By.XPath(METHOD_RECIPE));
            _allMethodRecipe.SetValue(ControlType.DropDownList, method);
            _allMethodRecipeChangeButton.Click();
             WaitPageLoading();
            _confirmDelete = WaitForElementIsVisible(By.Id(CONFIRM_DELETE));
            _confirmDelete.Click();
            WaitForLoad();
        }

        public bool CheckMethodChanges(string method)
        {
            var firstRecipe = WaitForElementExists(By.XPath("/html/body/div[3]/div/div[2]/div[3]/div/div[2]/div[1]/ul/li[1]/div/div[2]/div[3]/span"));
            var secondRecipe = WaitForElementExists(By.XPath("/html/body/div[3]/div/div[2]/div[3]/div/div[2]/div[1]/ul/li[2]/div/div[2]/div[3]/span"));
            if (!string.IsNullOrEmpty(firstRecipe.Text) && !string.IsNullOrEmpty(secondRecipe.Text) && secondRecipe.Text.Equals(method) && firstRecipe.Text.Equals(method)) { return true; }
            return false;
        }
        // Add recipe
        public void AddRecipe(string recipeName)
        {
            _addRecipe = WaitForElementIsVisibleNew(By.Id(ADD_RECIPE));
            _addRecipe.SetValue(ControlType.TextBox, recipeName);
            WaitForLoad();

            var selectRecipe = WaitForElementIsVisibleNew(By.XPath("//p[contains(@class,'click-to-add') and (contains(text(), '" + recipeName + "'))]"));
            TryToClickOnElement(selectRecipe);
            WaitForLoad();
        }

        // Tableau recettes
        public void ClickOnDayBefore()
        {
            _dayBefore = WaitForElementIsVisible(By.Id(DAY_BEFORE));
            _dayBefore.Click();
            WaitForLoad();
        }

        public void ClickOnDayAfter()
        {
            _dayAfter = WaitForElementIsVisible(By.Id(DAY_AFTER));
            _dayAfter.Click();
            WaitForLoad();
        }

        public DateTime GetDayDisplayed()
        {
            var dateFormat = _showMenuDay.GetAttribute("data-date-format");
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            _dayDisplay = WaitForElementIsVisible(By.Id(DAY_DISPLAY));

            string date;
            string day = _dayDisplay.Text;
            date = day.Substring("Day ".Length, 10);
            return DateTime.Parse(date, ci).Date;
        }

        public void SetTheoricalPAXDay(string value)
        {
            _theoricalPAXDay = WaitForElementIsVisible(By.Id(THEORICAL_PAX_DAY));
            _theoricalPAXDay.ClearElement();
            _theoricalPAXDay.SendKeys(value);
            WaitPageLoading();
        }

        public string GetTheoricalPAXDay()
        {
            _theoricalPAXDay = WaitForElementIsVisible(By.Id(THEORICAL_PAX_DAY));
            return _theoricalPAXDay.GetAttribute("value");
        }

        public void DuplicateMenuDay(DateTime day)
        {
            _duplicateDay = WaitForElementIsVisible(By.XPath(DUPLICATE_DAY));
            _duplicateDay.Click();
            WaitForLoad();

            var date = WaitForElementIsVisible(By.Id("date-picker-start"));
            date.SetValue(ControlType.DateTime, day);
            WaitForLoad();
            if (isElementVisible(By.XPath(CONFIRM_DUPLICATE)))
            {
                _confirmDuplicate = WaitForElementIsVisible(By.XPath(CONFIRM_DUPLICATE));
            }
            else
            {
                _confirmDuplicate = WaitForElementIsVisible(By.XPath("//*[@id=\"modal-1\"]/div[2]/div/form/div[2]/button[2]"));
            }
            _confirmDuplicate.Click();
            WaitForLoad();
        }

        public void ClickOnFirstRecipe()
        {
            _firstRecipeName = WaitForElementIsVisible(By.XPath(RECIPE_NAME));
            _firstRecipeName.Click();
            WaitForLoad();
        }

        public void ClickOnKioskImage()
        {
            _kioskImage = WaitForElementIsVisible(By.XPath(KIOSK_IMAGE));
            _kioskImage.Click();
            WaitForLoad();
        }
        public string SelectKiosk()
        {
            _selectedKioskImage = WaitForElementIsVisible(By.XPath(SELECTED_KIOSK_IMAGE));
            _selectedKiosk = WaitForElementIsVisible(By.XPath(SELECTED_KIOSK));
            string kiosqueName = _selectedKiosk.Text;
            _selectedKioskImage.Click();
            return kiosqueName;
        }
        public bool VerifyKioskName(string kiosqueName)
        {
            var kiosqueNameLower = kiosqueName.ToLower();
            var kiosqueNameRecipe = WaitForElementIsVisible(By.XPath("//div/div[2][contains(@class, 'display-zone')]/div[3][@style]/span"));
            var kiosqueNameRecipeLower = kiosqueNameRecipe.Text.ToLower();
            if (kiosqueNameRecipeLower.Equals(kiosqueNameLower))

                return true;


            return false;
        }

        public RecipesVariantPage EditFirstRecipe()
        {
            ClickOnFirstRecipe();

            _editFirstRecipe = WaitForElementIsVisible(By.XPath(EDIT_FIRST_RECIPE));
            _editFirstRecipe.Click();
            WaitForLoad();

            //Results are opened in a new tab, switch the driver to the newly created one
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(30));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            WaitForLoad();

            return new RecipesVariantPage(_webDriver, _testContext);
        }

        public void DeleteFirstRecipe()
        {
            ClickOnFirstRecipe();
            if (isElementVisible(By.XPath(DELETE_FIRST_RECIPE)))
            {
                _deleteFirstRecipe = WaitForElementIsVisible(By.XPath(DELETE_FIRST_RECIPE));
            }
            else
            {
                _deleteFirstRecipe = WaitForElementIsVisible(By.XPath("//a[contains (@class, 'btn  detailDeleteBtn active-post-link')]/span[contains(@class,'fas fa-trash-alt')]"));
            }
            _deleteFirstRecipe.Click();
            WaitForLoad();

            _confirmDelete = WaitForElementIsVisible(By.Id(CONFIRM_DELETE));
            _confirmDelete.Click();
            WaitForLoad();
        }

        public string GetRecipeName()
        {
            _firstRecipeName = WaitForElementIsVisible(By.XPath(RECIPE_NAME));
            return _firstRecipeName.Text;
        }

        public bool IsRecipeDisplayed()
        {
            bool isDisplayed;

            if (isElementVisible(By.XPath(RECIPE_NAME)))
            {
                isDisplayed = true;
            }
            else
            {
                isDisplayed = false;
            }

            return isDisplayed;
        }
        public bool IsThisRecipeDisplayed(string recipeName)
        {
            bool isDisplayed;

            if (isElementVisible(By.XPath(String.Format(THIS_RECIPE_NAME, recipeName))))
            {
                isDisplayed = true;
            }
            else
            {
                isDisplayed = false;
            }

            return isDisplayed;
        }
        public void SetRecipeMethod(string method)
        {
            _recipeMethodSelected = WaitForElementIsVisible(By.XPath(METHOD_SELECTED));
            _recipeMethodSelected.SetValue(ControlType.DropDownList, method);

            // Temps de prise en compte de la modification
            WaitPageLoading();
        }

        public string GetRecipeMethod()
        {
            _recipeMethod = WaitForElementExists(By.XPath(METHOD));
            return _recipeMethod.Text;
        }

        public void SetRecipeCoef(string coef)
        {
            _recipeCoefSelected = WaitForElementIsVisible(By.XPath(COEF_SELECTED));
            var actions = new Actions(_webDriver);
            actions.MoveToElement(_recipeCoefSelected).Perform();

            try
            {
                _recipeCoefSelected.SetValue(ControlType.TextBox, coef);
            }
            catch
            {
                _recipeCoefSelected.Clear();
                _recipeCoefSelected.SendKeys(coef);
            }


            // Temps de prise en compte de la modification
            WaitPageLoading();
        }

        public string GetRecipeCoef()
        {
            _recipeCoef = WaitForElementExists(By.XPath(COEF));
            return _recipeCoef.Text;
        }
        public string GetPortion()
        {
            _recipeCoef = WaitForElementExists(By.XPath(PORTION));
            return _recipePorton.Text;
        }
        public string GetPAX()
        {
            _recipeCoef = WaitForElementExists(By.XPath(PAX));
            return _recipePax.Text;
        }

        public double GetPAX(string decimalSeparator)
        {
            WaitPageLoading();
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            _recipeCoef = WaitForElementExists(By.XPath(PAX));
            var total = _recipePax.Text;

            return Convert.ToDouble(total, ci);
        }
        public string GetPRIXXTOTAL()
        {
            _recipePrixtotal = WaitForElementExists(By.XPath(PRIXTOTAL));
            return _recipePrixtotal.Text;
        }

        public double GetPRIXXTOTAL(string decimalSeparator)
        {
            WaitPageLoading();
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            _recipePrixtotal = WaitForElementExists(By.XPath(PRIXTOTAL));
            var total = _recipePrixtotal.Text;

            return Convert.ToDouble(total, ci);
        }



        public string GetSite()
        {
            var menuInformations = WaitForElementIsVisible(By.XPath(MENU_INFORMATIONS));
            var site = menuInformations.Text.Substring(menuInformations.Text.Length - 3);
            return site;
        }

        public bool IsGuestTypeExist(string guestType)
        {
            var menuInfos = WaitForElementIsVisible(By.XPath(MENU_INFORMATIONS));
            if (menuInfos.Text.Contains(guestType))
                return true;
            return false;
        }
        public void ExportToOtherMenu(string site, string menuName)
        {
            Actions action = new Actions(_webDriver);
            ShowExtendedMenu();

            var _exportToOtherMenu = WaitForElementIsVisible(By.XPath(EXPORT_TO_OTHER_MENU));
            _exportToOtherMenu.Click();
            WaitForLoad();

            var _addMultipleMenusDest = WaitForElementIsVisible(By.XPath(ADD_MULTIPLE_MENUS_DESTINATIONS));
            action.MoveToElement(_addMultipleMenusDest).Perform();
            _addMultipleMenusDest.Click();
            WaitPageLoading();
            WaitForLoad();

            ComboBoxSelectById(new ComboBoxOptions(MENU, menuName, false));

            ComboBoxSelectById(new ComboBoxOptions(SITE, site, false));

            var _saveBtn = WaitForElementIsVisible(By.XPath(SAVE));
            _saveBtn.Click();
            WaitForLoad();

            var _export = WaitForElementIsVisible(By.XPath(EXPORT));
            _export.Click();
            WaitForLoad();

            var _okBtn = WaitForElementIsVisible(By.XPath(OK_BUTTON));
            _okBtn.Click();
            WaitForLoad();
        }
        public int CheckTotalRecipe()
        {
            var _totalRecipe = 0;
            var listRecipes = _webDriver.FindElements(By.XPath("//*[@id=\"datasheet-list\"]/ul/li"));
            _totalRecipe = listRecipes.Count;
            return _totalRecipe;
        }
        public bool VerifyrecipeExist(string recipe)
        {
            var listRecipes = _webDriver.FindElements(By.XPath("//*[@id=\"datasheet-list\"]/ul/li/div/div[@class =\"display-zone col\"]/div/div[2]/p[1]"));
            foreach (var i in listRecipes)
            {
                if (i.Text != recipe)
                {
                    return false;
                }
            }
            return true;
        }
        public void ClickPlanView()
        {
            _planView = WaitForElementIsVisible(By.Id(PLAN_VIEW));
            _planView.Click();
            WaitForLoad();
        }
        public bool VerifyPlanViewChanged()
        {
            if (isElementVisible(By.XPath("//*[@id=\"menuDisplay-list\"]/div[1]/div/div")))
            {
                return true;
            }
            return false;
        }
        public void ClickPricerDevices()
        {
            _pricerDevices = WaitForElementIsVisible(By.XPath(PRICER_DEVICES));
            _pricerDevices.Click();
            WaitForLoad();
        }
        public void UpdateMenuFromPricerDevices()
        {
            _updateMenuBtn = WaitForElementIsVisible(By.XPath(UPDATE_MENU_BTN));
            _updateMenuBtn.Click();
            WaitPageLoading();
        }
        public PrintReportPage Print(bool newVersionPrint)
        {
            ShowExtendedMenu();
            _print = WaitForElementIsVisible(By.XPath(PRINT));
            _print.Click();
            WaitForLoad();

            var _printBtn = WaitForElementIsVisible(By.XPath("//*[@id=\"print\"]"));
            _printBtn.Click();
            WaitForLoad();

            if (newVersionPrint)
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

        public MenusDisplayPage OpenDisplayMenu()
        {
            _displayMenu = WaitForElementIsVisible(By.CssSelector(DISPLAY_MENU));
            _displayMenu.Click();

            WaitForLoad();

            _displayMenuCheckService = WaitForElementIsVisible(By.Id(DISPLAY_MENU_CHECK_SERVICE));
            _displayMenuCheckService.Click();
            _displayMenuValidate = WaitForElementIsVisible(By.Id(DISPLAY_MENU_VALIDATE));
            _displayMenuValidate.Click();

            //Menu is opened in a new tab, switch the driver to the newly created one
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new MenusDisplayPage(_webDriver, _testContext);
        }

        public void SelectDietaryType(string dietaryType)
        {
            ComboBoxSelectById(new ComboBoxOptions("SelectedDietaryTypes_ms", dietaryType, true));
            WaitForLoad();
        }

        public int GetNumberOfRecipesDisplayed()
        {
            WaitForLoad();
            var _elementsToAddList = _webDriver.FindElements(By.XPath(ELEMENTS_TO_ADD_LIST));
            WaitForLoad();

            return _elementsToAddList.Count;
        }

        public void searchRecipe(string recipeName)
        {
            Actions action = new Actions(_webDriver);

            var _searchRecipe = WaitForElementIsVisible(By.Id(ADD_RECIPE));
            action.MoveToElement(_searchRecipe).Perform();
            _searchRecipe.SetValue(ControlType.TextBox, recipeName);
            WaitForLoad();
        }
        public enum FilterType
        {
            SearchByRecipe,
            RecipeTypes,
            DietaryTypes
        }
        public void Filter(FilterType filterType, object value)
        {
            Actions action = new Actions(_webDriver);

            switch (filterType)
            {
                case FilterType.SearchByRecipe:
                    _addRecipe = WaitForElementIsVisible(By.Id(ADD_RECIPE));
                    _addRecipe.SetValue(ControlType.TextBox, value);
                    break;
                case FilterType.RecipeTypes:
                    _recipe_types_filter = WaitForElementExists(By.Id(RECIPE_TYPES_FILTER));
                    action.MoveToElement(_recipe_types_filter).Perform();
                    _recipe_types_filter.Click();

                    _unCheckAllRecipeTypes = _webDriver.FindElement(By.XPath(UNCHECK_ALL_RECIPE_TYPES));
                    _unCheckAllRecipeTypes.Click();

                    IWebElement selectedRecipeTypes;
                    _searchRecipeType = WaitForElementIsVisible(By.XPath(SEARCH_RECIPE_TYPES));
                    _searchRecipeType.SetValue(ControlType.TextBox, value);

                    selectedRecipeTypes = _webDriver.FindElement(By.XPath("//html/body/div[11]/ul/li[11]/label/span[text()='" + value + "']"));
                    selectedRecipeTypes.SetValue(ControlType.CheckBox, true);

                    _recipe_types_filter.Click();
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);
            }

            WaitPageLoading();
            WaitForLoad();
        }
        public List<string> GetAllRecipeNameResultFiltredByRecipe()
        {
            List<string> resultFilter = new List<string>();
            var results = _webDriver.FindElements(By.XPath(LIST_RESULT_RECIPE_NAME));
            WaitForLoad();
            foreach (var result in results)
            {
                resultFilter.Add(result.Text);
            }
            return resultFilter;
        }
        public int NbAllergenCheked()
        {
            var checkallallergenh = _webDriver.FindElements(By.XPath("//*[starts-with(@id,\"popover\")]/div[2]/div/div/div/table/tbody/tr[*]/td[*]/div/figure/img[@class='captcha checked']"));
            return checkallallergenh.Count();
        }
        public bool IsFirstAllergenCheked()
        {
            if (NbAllergenCheked() >= 1)
            {
                _checkFirstallergen = WaitForElementIsVisible(By.XPath(CHECK_FIRST_ALLERGEN));
                if (_checkFirstallergen.GetAttribute("class").Equals("captcha checked"))
                {
                    return true;
                }
            }
            return false;
        }

        public void CheckFirstAllergenEtoile()
        {
            WaitPageLoading();
            Actions actions = new Actions(_webDriver);
            _checkFirstallergen = WaitForElementIsVisible(By.XPath(CHECK_FIRST_ALLERGEN));
            actions.MoveToElement(_checkFirstallergen).Perform();
            _checkFirstallergen.Click();
            WaitPageLoading();
        }

        public void ClickOnKeyword()
       {
            Actions actions = new Actions(_webDriver);
            _keyword = WaitForElementIsVisible(By.XPath(KEYWORDS));
            actions.MoveToElement(_keyword).Perform();
            _keyword.Click();
            WaitForLoad();
       }
       public void ClickOnAllergenEtoile()
       {
            Actions actions = new Actions(_webDriver);
            _allergenetoile = WaitForElementIsVisible(By.XPath(ALLERGEN_ETOILE));
            actions.MoveToElement(_allergenetoile).Perform();
            _allergenetoile.Click();
            WaitForLoad();
            WaitPageLoading();
        }
       public void ClickOnKiosque()
       {
            Actions actions = new Actions(_webDriver);
            _kiosqueimage = WaitForElementIsVisible(By.XPath(KIOSQUE_IMAGE));
            actions.MoveToElement(_kiosqueimage).Perform();
            _kiosqueimage.Click();
            WaitForLoad();
       }

       public void CheckFirstKiosque()
       {
            Actions actions = new Actions(_webDriver);
            _checkkiosque = WaitForElementIsVisible(By.XPath(CHECK_FIRST_KIOSQUE));
            actions.MoveToElement(_checkkiosque).Perform();   
            _checkkiosque.Click();
            WaitForLoad();
       }

       public void CheckKeyword(string keyword)
       {
            Actions actions = new Actions(_webDriver);
            _checkkeyword = WaitForElementIsVisible(By.XPath(string.Format(KEYWORD_CHECK, keyword)));
            _checkkeyword.Click();
       }
       
       public void ResetAllAllergenCkeked()
       {
                  
        var checkallallergenh = _webDriver.FindElements(By.XPath("//*[starts-with(@id,\"popover\")]/div[2]/div/div/div/table/tbody/tr[*]/td[*]/div/figure/img[@class='captcha checked']"));

            if(checkallallergenh.Count()>0)
            {
                IJavaScriptExecutor js = (IJavaScriptExecutor)_webDriver;

                foreach (var element in checkallallergenh)
                {
                    js.ExecuteScript("arguments[0].className = 'captcha';", element);
                }
            }
        }
       public void ResetAllKeywordCkeked()
        {

            var checkallkeyword = _webDriver.FindElements(By.XPath("//*[starts-with(@id,\"popover\")]/div[2]/div/div/div/table/tbody/tr[*]/td[*]/div/figure[@class='captcha checked']"));

            if (checkallkeyword.Count() > 0)
            {
                IJavaScriptExecutor js = (IJavaScriptExecutor)_webDriver;

                foreach (var element in checkallkeyword)
                {
                    js.ExecuteScript("arguments[0].className = 'captcha';", element);
                }
            }
        }        
       public void ResetAllKiosqueCkeked()
        {

            var checkallKiosque = _webDriver.FindElements(By.XPath("//*[starts-with(@id,\"popover\")]/div[2]/div/div/div/div/img[@class='captcha checked']"));

            if (checkallKiosque.Count() > 0)
            {
                IJavaScriptExecutor js = (IJavaScriptExecutor)_webDriver;

                foreach (var element in checkallKiosque)
                {
                    js.ExecuteScript("arguments[0].className = 'captcha';", element);
                }
            }
        }
        
        public bool CheckAllergenchecked()
        {
            var checkallkeyword = _webDriver.FindElements(By.XPath("//*[starts-with(@id,\"popover\")]/div[2]/div/div/div/table/tbody/tr/td/div/figure[img[@class='captcha checked'] and figcaption[text()='Egg']]"));
            var lkeywordchecked = _webDriver.FindElements(By.XPath("//*[starts-with(@id,\"popover\")]/div[2]/div/div/div/table/tbody/tr/td/div/figure/img[@class='captcha checked']"));


            if ( checkallkeyword.Count() == 1 && lkeywordchecked.Count() == 1)
            {
                return true;
            }

            return false;
        }
        public bool CheckKeywordchecked(string keyword)
        {
                _checkkeyword = WaitForElementIsVisible(By.XPath(string.Format(KEYWORD_CHECK, keyword)));
                if (_checkkeyword.GetAttribute("class").Equals("captcha checked"))
                {
                    return true;
                }

                return false;

        }
        public string GetkeywordChecked()
        {
            var keywordchecked = _webDriver.FindElement(By.XPath(KEYWORD_CHECKED));
            if (keywordchecked.GetAttribute("class").Equals("captcha checked"))
            {
                return keywordchecked.Text;
            }
            return "";

        }

        public bool IsFirstkiosqueCheked()
        {
                _checkFirstallergen = WaitForElementIsVisible(By.XPath(CHECK_FIRST_KIOSQUE));
                if (_checkFirstallergen.GetAttribute("class").Equals("captcha checked"))
                {
                    return true;
                }
            return false;
        }
        public string GetRecipeNameOnDisplayView()
        {
            _firstRecipeNameDisplayView = WaitForElementIsVisible(By.XPath(FIRST_RECIPE_NAME_ONDISPLAYVIEW));
            return _firstRecipeNameDisplayView.Text;
        }
        public void ClickOnFirstRecipeOnDispalyView()
        {
            var firstRecipeName = WaitForElementIsVisible(By.XPath(RECIPE_NAME_ONDISPLAYVIEW));
            firstRecipeName.Click();
            WaitForLoad();  
        }
        public void SetFirstRecipeOnDispalyView(string updateRecipeName)
        {
            var inputRecipeName = WaitForElementIsVisible(By.XPath(RECIPE_NAME_INPUT_ON_DISPLAY_VIEW));
            inputRecipeName.SetValue(PageBase.ControlType.TextBox,updateRecipeName);
            inputRecipeName.SendKeys(Keys.Enter);
            WaitPageLoading();
        }

        public int GetDaysMenuPlanningOfMonthCount()
        {
            var _listOfDays = _webDriver.FindElements(By.XPath(LIST_DAYS));
            return _listOfDays.Count;
        }
        public bool CheckDaysDispoAreDaysSelected(int daySelected)
        {
            var daysCount = GetDaysMenuPlanningOfMonthCount();
            for (int i = 0; i < daysCount; i++) {
                var gggg = GetDayDisplayed().Day.ToString();
                if (GetDayDisplayed().Date.DayOfWeek.ToString() != Enum.GetName(typeof(DayOfWeek), daySelected))
                {
                    return false;
                }

                this.ClickOnDayAfter();
            }
            return true;
        }
        public void ClickOnKeyWords()
        {
            Actions actions = new Actions(_webDriver);
            _iconKeywords = WaitForElementIsVisible(By.XPath(ICON_KEYWORDS_RECIPE));
            actions.MoveToElement(_iconKeywords).Perform();
            _iconKeywords.Click();
            WaitForLoad();
        }

        public List<string> GetKeyWordsRecipe()
        {
            ClickOnKeyWords();
            List<string> result = new List<string>();
            var elements = _webDriver.FindElements(By.XPath(KEYWORDS_RECIPE));
            foreach (var element in elements)
            {
                result.Add(element.Text.Trim());
            }
            return result;
        }

        public void SetSetCoeffMassiveCoef(string coef)
        {
            _coefvalid = WaitForElementIsVisible(By.XPath(COEF_VALID));
            var actions = new Actions(_webDriver);
            actions.MoveToElement(_coefvalid).Perform();
            _coefvalid.SetValue(ControlType.TextBox, coef);
            WaitForLoad();
        }

        public void ClickValidatee()
        {
            _confirmationcoef = WaitForElementIsVisible(By.XPath(CONFIRMATIONCOEF));
            _confirmationcoef.Click();
            WaitForLoad();
        }
        public string GetMessageConfirmationMassive()
        {
            _confirmationmessage = WaitForElementIsVisible(By.XPath(CONFIRMATIONMESSAGE));
           return _confirmationmessage.Text;
        }
         public void ClickConfirmationMassive()
        {
            _btnconfirmationmessage = WaitForElementIsVisible(By.XPath(BTN_CONFIRMATIONMESSAGE));
            _btnconfirmationmessage.Click();
            WaitForLoad();
        }

        public List<string> GetAllCoefMassive()
        {
            List<string> result = new List<string>();

            var checkallkeyword = _webDriver.FindElements(By.XPath("//*[starts-with(@id,\"dataItem\")]/div/div[2]/div[4]/span"));

            if (checkallkeyword.Count() > 0)
            {
                foreach (var element in checkallkeyword)
                {
                    result.Add(element.Text);
                }
            }
            return result;
        }

        public string CheckDefaultAllergenchecked()
        {
             var allergencheked = _webDriver.FindElements(By.XPath("//*[starts-with(@id,\"popover\")]/div[2]/div/div/div/table/tbody/tr/td/div/figure/img[@class='captcha checked']"));

            if (allergencheked.Count() == 1)
            {
                foreach (var element in allergencheked)
                {                    
                    var h3Element = element.FindElement(By.XPath("//*[starts-with(@id,\"popover\")]/div[2]/div/div/div/table/tbody/tr/td/div/figure/img[@class='captcha checked']/../figcaption"));
                    return h3Element.Text;
                }
            }
            return "";
        }

        public List<string> GetAllMethod()
        {
            List<string> allMethodResult= new List<string>();

            var allmethod = _webDriver.FindElements(By.XPath(ALL_METHOD));

            if (allmethod != null)
            {
                foreach (var element in allmethod)
                {
                    allMethodResult.Add(element.Text.Trim());
                }
            }
            
            return allMethodResult;
        }

        public string GetMenuVariant()
        {
            WaitPageLoading();
            var _variantmenu = WaitForElementExists(By.XPath(VARIANTMENU_XPATH));

            return _variantmenu.Text;
        }

        public bool ClickOnFirstRecipeOFFDispalyView()
        {

            var firstRecipeName = WaitForElementIsVisible(By.XPath(RECIPE_NAME_OFFDISPLAYVIEW));

            if (firstRecipeName != null)
            {

                firstRecipeName.Click();
                WaitForLoad();

                return true;

            }
            else return false;
        }
      public void ClickOnKeywordDisplayView()
        {
            Actions actions = new Actions(_webDriver);
            var _keywordDisplay = WaitForElementIsVisible(By.Id(KEYWORD_DISPLAY));
            actions.MoveToElement(_keywordDisplay).Perform();
            _keywordDisplay.Click();
            WaitForLoad();
        }
        public void ClickOnDisplayView()
        {
            Actions actions = new Actions(_webDriver);
            var _keywordDisplay = WaitForElementIsVisible(By.Id(PLAN_VIEW));
            actions.MoveToElement(_keywordDisplay).Perform();
            _keywordDisplay.Click();
            WaitForLoad();
        }
        public void ClickOnFirstMenuRecipe()
        {
            var firstMenuRecipeName = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div[2]/div[3]/div/div[2]/div[2]/ul/li/div/div[2]/div[1]/div[2]/p[1]"));
            firstMenuRecipeName.Click();
            WaitForLoad();
        }
        public bool MessageIsDisplayView()
        {
            _messageDisplay = WaitForElementIsVisible(By.Id(MESSAGE_DISPLAY));
            if (_messageDisplay != null)
            {
                WaitForLoad();
                return true;

            }
            else return false;
        }
        public bool ClickOnKiosquesIcon()
        {
            Actions actions = new Actions(_webDriver);
            var kiosquesIcon = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div[2]/div[3]/div/div[2]/div[2]/ul/li/div/div[1]/div[4]/span"));
            actions.MoveToElement(kiosquesIcon).Perform();
            var originalPosition = kiosquesIcon.Location;
            kiosquesIcon.Click();
            WaitForLoad();
            var closeKiosquesIcon = WaitForElementIsVisible(By.XPath("/html/body/div[13]/h3/button"));
            actions.MoveToElement(closeKiosquesIcon).Perform();
            closeKiosquesIcon.Click();
            var newPosition = kiosquesIcon.Location;
            return originalPosition.Equals(newPosition);
        }

        public bool ClickOnAllergensIcon()
        {
            Actions actions = new Actions(_webDriver);
            var allergensIcon = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div[2]/div[3]/div/div[2]/div[2]/ul/li/div/div[1]/div[5]/span"));
            actions.MoveToElement(allergensIcon).Perform();
            var originalPosition = allergensIcon.Location;
            allergensIcon.Click();
            WaitForLoad();
            var closeIcon = WaitForElementIsVisible(By.XPath("/html/body/div[13]/h3/button"));
            actions.MoveToElement(closeIcon).Perform();
            closeIcon.Click();
            var newPosition = allergensIcon.Location;
            return originalPosition.Equals(newPosition);
        }

        public bool ClickOnKeywordsIcon()
        {
            Actions actions = new Actions(_webDriver);
            var keywordsIcon = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div[2]/div[3]/div/div[2]/div[2]/ul/li/div/div[1]/div[6]/span"));
            actions.MoveToElement(keywordsIcon).Perform();
            var originalPosition = keywordsIcon.Location;
            keywordsIcon.Click();
            WaitForLoad();
            var newPosition = keywordsIcon.Location;
            return originalPosition.Equals(newPosition);
        }
        public bool MessageIsPlanView()
        {
            _messageDisplay = WaitForElementIsVisible(By.Id(MESSAGEPLANVIEW));
            if (_messageDisplay != null)
            {
                WaitForLoad();
                return true;

            }
            else return false;
        }
        public List<string> GetTheFirstLineOfDisplayPage()
        {
            List<string> firstLine = new List<string>();
            IList<IWebElement> list = _webDriver.FindElements(By.XPath(FIRSTLINEDISPLAYVIEW));

            foreach (IWebElement element in list)
            {
                firstLine.Add(element.Text);
            }

            return firstLine;
        }
        public List<string> GetTheFirstLineOfPlanPage()
        {
            List<string> firstLine = new List<string>();
            IList<IWebElement> list = _webDriver.FindElements(By.XPath(FIRSTLINEPLANVIEW));

            foreach (IWebElement element in list)
            {
                firstLine.Add(element.Text);
            }

            return firstLine;
        }
    }
}
