using DocumentFormat.OpenXml.Office2010.ExcelAc;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.Production;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Production
{
    public class ParametersProduction : PageBase
    {
        private int allergenRowIndex = -1;

        // _____________________________________________ Constantes _______________________________________

        // Group
        private const string GROUP_TAB = "//*[@id=\"paramProdTab\"]/li[*]/a[text()='Group']";
        private const string GROUP_LINE = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[2][normalize-space(text())='{0}']";
        private const string ADD_GROUP = "//*[@id=\"tabContentParameters\"]/div[1]/a";
        private const string GROUP_CODE = "first";
        private const string GROUP_NAME = "Name";
        private const string SAVE_GROUP = "last";
        private const string ISPRODUCT_GROUP = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[2][contains(text(),'{0}')]/../td[8]/div/input[1]";
        private const string ISFOOD_GROUP = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[2][contains(text(),'{0}')]/../td[9]/div/input[1]";
        private const string EDIT_GROUP = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[2][contains(text(),'{0}')]/../td[11]/a[contains(@href, 'Edit')]";
        private const string GROUP_TABLE = "//*[@id=\"tabContentParameters\"]/table/tbody";
        // SubGroup
        private const string SUBGROUP_TAB = "//*[@id=\"paramProdTab\"]/li[*]/a[text()='SubGroup']";
        private const string NEW_SUBGROUP_BTN = "//*[@id=\"tabContentParameters\"]/div[1]/a";
        private const string SUB_GROUP_CODE = "first";
        private const string SUB_GROUP_NAME = "Name";
        private const string SUB_GROUP_SAVE_BTN = "btnSave";
        private const string LIST_SUB_GROUP = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[2]";
        private const string GROUP_SELECT = "//*[@id=\"ItemSubGroupModal\"]/div/div/div/div/form/div[2]/div[3]/div/select";


        // Inventory
        private const string INVENTORY_TAB = "//*[@id=\"paramProdTab\"]/li[*]/a[text()='Inventory']";
        private const string INVENTORY_DATE = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[{0}]/td[3]/a/span";

        // Guest
        private const string GUEST_TAB = "//*[@id=\"paramProdTab\"]/li[*]/a[text()='Guest']";
        private const string ADD_GUEST = "//*[@id=\"tabContentParameters\"]/div[1]/a";
        private const string GUEST_EDIT = "//td[1][normalize-space(text())='{0}']/..//span[contains(@class, 'glyphicon glyphicon-pencil')]";
        private const string GUEST_EDIT_2 = "//td[1][normalize-space(text())='{0}']/..//span[contains(@class, 'fas fa-pencil-alt')]";
        private const string GUEST_DELETE = "//*[@id=\"tabContentParameters\"]/div[2]/div/table/tbody/tr[*]/td[1][normalize-space(text())='{0}']/parent::*/td[11]/a[2]/span";
        private const string GUEST_LINE = "//*[@id=\"tabContentParameters\"]/div[2]/div/table/tbody/tr[*]/td[1][normalize-space(text())='{0}']";
        private const string TYPE_GUEST = "first";
        private const string GUEST_HAS_ALLERGY = "IsAllergy";
        private const string GUEST_IS_ALLERGIC = "item_IsAllergy";
        private const string GUEST_ORDER = "Order";
        private const string SAVE_GUEST = "last";
        private const string GUEST_ALLERGY_LINE = "//*[@id=\"GuestTypeModal\"]/div/div/div/div/form/ul/li[*]/label[text()='{0}']";
        private const string SAVE_ALLERGY = "//*[@id=\"GuestTypeModal\"]/div/div/div/div/form/div[2]/button[2]";
        private const string CONFIRM_DELETE_GUEST = "first";
        private const string ALLERGEN_ALL_LINE = "//*[@id=\"modal-1\"]/div/div/form/ul/li[*]/label";
        private const string ALLERGEN_LINE = "//*[@id=\"modal-1\"]/div/div/form/ul/li[*]/label[text()='{0}']";
        private const string DEFINE_FIRST_ALLERGENS = "//*[@id=\"tabContentParameters\"]/div[2]/div/table/tbody/tr[2]/td[8]/a";
        private const string SAVE_ALLEERGY_GUEST = "//*[@id=\"modal-1\"]/div/div/form/div[2]/button[2]";


        // Variant
        private const string VARIANT_TAB = "//*[@id=\"paramProdTab\"]/li[*]/a[text()='Variant']";
        private const string PLUS_MENU_VARIANT = "//*[@id=\"tabContentParameters\"]/div[1]/div/div[2]/button";
        private const string ADD_VARIANT = "//*[@id=\"tabContentParameters\"]/div[1]/div/div[2]/div/a";
        private const string VARIANT_GUEST = "first";
        private const string VARIANT_MEAL = "SiteVariant_Variant_MealTypeId";
        private const string VARIANT_SITE = "SiteVariant_SiteId";
        private const string SAVE_VARIANT = "last";

        private const string FILTER_SITE = "SelectedSites_ms";
        private const string SITE_SEARCH = "/html/body/div[10]/div/div/label/input";
        private const string UNCHECK_ALL_SITE = "/html/body/div[10]/div/ul/li[2]/a";

        private const string FILTER_GUEST = "SelectedGuestTypes_ms";
        private const string GUEST_SEARCH = "/html/body/div[11]/div/div/label/input";
        private const string UNCHECK_ALL_GUEST = "/html/body/div[11]/div/ul/li[2]/a";

        private const string FILTER_MEAL = "SelectedMealTypes_ms";
        private const string MEAL_SEARCH = "/html/body/div[12]/div/div/label/input";
        private const string UNCHECK_ALL_MEAL = "/html/body/div[12]/div/ul/li[2]/a";

        private const string DELETE_VARIANT = "//*[@id=\"tableVariants\"]/tbody/tr[2]/td[4]/a[2]";
        private const string CONFIRM_DELETE_VARIANT = "DeleteVariant";
        private const string CONFIRM_DELETE_2 = "dataConfirmOK";

        private const string VARIANTS = "//*[@id=\"tableVariants\"]/tbody/tr[*]";
        private const string VARIANTS_LIST = "//*[@id=\"tableVariants\"]/tbody/tr[*]/td[2]";
        private const string FIRST_GUEST_VARIANT = "//*[@id=\"tableVariants\"]/tbody/tr[2]/td[2]";
        private const string FIRST_MEAL_VARIANT = "//*[@id=\"tableVariants\"]/tbody/tr[2]/td[3]";

        //Meal
        private const string MEAL_TAB = "//*[@id=\"paramProdTab\"]/li[*]/a[text()='Meal']";
        private const string NEW_MEAL = "//*[@id=\"tabContentParameters\"]/div[1]/a";
        private const string MEAL_NAME = "first";
        private const string MEAL_COMMERCIAL_NAME = "CommercialName1";
        private const string MEAL_ORDER = "Order";
        private const string MEAL_HOUR = "//*[@id=\"MealTypeModal\"]/div/div/div/div/form/div[2]/div[5]/div/input[2]";
        private const string SAVE_MEAL = "last";
        private const string MEALS = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]";

        //Foodpack
        private const string FOODPACK_TAB = "//*[@id=\"paramProdTab\"]/li[*]/a[text()='Food pack']";
        private const string NEW_FOODPACK = "//*[@id=\"tabContentParameters\"]/div[1]/a";
        private const string ADD_NEW_FOODPACK = "SaveAndNew";
        //private const string DELETE_FOODPACK = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[3]/td[{0}]/../td[4]/a[2]/span";
        private const string SAVE_FOODPACK = "last";
        private const string FOODPACK_NAME = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[contains(text(), '{0}')]";
        private const string INPUT_NEW_FOODPACK_NAME = "first";
        private const string FOODPACK_EDIT_BUTTON = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[contains(text(),'{0}')]/../td[5]/a[1]/span";
        private const string FOODPACK_ADD_FOODPACK_BUTTON = "//*[@id=\"bulkContainer\"]/div[2]/div[2]/div/div/a[2]";
        private const string FOODPACK_ADD_FOODPACK_SELECT = "//*[@id=\"bulkContainer\"]/div[2]/div[2]/select";
        private const string FOODPACK_ADD_FOODPACK_VALUE = "//*[@id=\"bulkContainer\"]/div[2]/div[3]/input";
        private const string FOODPACK_ADD_WARMING_MODE = "WarmingMode";
        private const string FOODPACK_ADD_WARMING_TEMPERATURE = "WarmingTemperature";
        private const string FOODPACK_BULK_EQUIVALENCE = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[contains(text(),'{0}')]/../td[4]";
        private const string FOODPACK_WARMING_MODE = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[contains(text(),'{0}')]/../td[2]";
        private const string FOODPACK_WARMING_TEMPERATURE = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[contains(text(),'{0}')]/../td[3]";

        //Recipe Type Tab
        private const string RECIPE_TYPE_TAB = "//*[@id=\"paramProdTab\"]/li[*]/a[text()='Recipe type']";
        private const string RECIPE_TYPE_NAME = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[contains(text(), '{0}')]";
        private const string ADD_RECIPE_TYPE_BUTTON = "//*[@id=\"tabContentParameters\"]/div[1]/a";
        private const string RECIPE_TYPE_TYPE_INPUT = "first";
        private const string RECIPE_TYPE_ORDER_INPUT = "Order";
        private const string RECIPE_TYPE_SAVE_BUTTON = "last";


        //Workshop Tab
        private const string WORKSHOP_TAB = "/html/body/div[2]/div/ul/li[14]/a";
        private const string WORKSHOPS = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[1]";
        private const string CREATE_WORKSHOP = "//*[@id=\"tabContentParameters\"]/div[1]/a";
        private const string CREATE_INPUT = "first";
        private const string ORDER_INPUT = "//*[@id=\"Order\"]";
        private const string CONFIRM_SAVE = "//*[@id=\"last\"]";
        private const string WORKSHOP_TYPE = "//*[@id=\"Type\"]";

        //Keywords Tab
        private const string KEYWORD_TAB = "//*/a[text()='Keywords']";

        //Allergen Tab 
        private const string ALLERGEN_TAB = "//*[@id=\"paramProdTab\"]/li[7]/a";
        private const string NEW_ALLERGEN_BTN = "//*[@id=\"tabContentParameters\"]/div[1]/div/a";
        private const string NAME_INPUT = "//*[@id=\"first\"]";
        private const string ENGNAME_INPUT = "//*[@id=\"EnglishName\"]";
        private const string SAVE_BTN = "//*[@id=\"last\"]";
        private const string ALLERGEN_LINETABLE = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]label[text()='{0}']\r\n";

        

        private const string CREATE_KEYWORD = "//*[@id=\"btn-add-keywordvalue\"]";
        private const string VALUE_INPUT = "//*[@id=\"first\"]";
        private const string DISPLAY_EAT = "//*[@id=\"model-name-input\"]";
        private const string SAVE = "last";


        private const string KEYWORD_LINE = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[2][normalize-space(text())='{0}']";
        private const string KEYWORD_DELETE = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[2][normalize-space(text())='{0}']/parent::*/td[5]/a[2]/span";
        private const string CONFIRM_DELETE_KEYWORD = "first";
        private const string CHECK_KEYWORD = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[2][text()='{0}']";
        private const string CLOSE_ICON_DELETE_KEYWORD = "//*[@id=\"modal-1\"]/div/button";
        private const string OK_DELETE_KEYWORD = "//*[@id=\"tabDeleteKeywordValueContainer\"]/pre";
        private const string COOKING_MODE = "//*[@id=\"paramProdTab\"]/li[9]/a";
        // ____________________________________________ Variables _________________________________________

        //Keyword

        [FindsBy(How = How.XPath, Using = KEYWORD_TAB)]
        private IWebElement _keywordtab;

        [FindsBy(How = How.XPath, Using = ALLERGEN_TAB)]
        private IWebElement _allergentab;

        [FindsBy(How = How.XPath, Using = NEW_ALLERGEN_BTN)]
        private IWebElement _newallergen;

        [FindsBy(How = How.XPath, Using = NAME_INPUT)]
        private IWebElement _nameinputallergen;

        [FindsBy(How = How.XPath, Using = ENGNAME_INPUT)]
        private IWebElement _engnameinputallergen;

        [FindsBy(How = How.XPath, Using = SAVE_BTN)]
        private IWebElement _saveallergen;

        [FindsBy(How = How.XPath, Using = ALLERGEN_LINETABLE)]
        private IWebElement _allergenLinetr
;

        [FindsBy(How = How.XPath, Using = CREATE_KEYWORD)]
        private IWebElement _createkeyword;

        [FindsBy(How = How.Id, Using = VALUE_INPUT)]
        private IWebElement _valueinputkeyword;

        [FindsBy(How = How.Id, Using = DISPLAY_EAT)]
        private IWebElement _displayeat;

        [FindsBy(How = How.Id, Using = SAVE)]
        private IWebElement _savekeyword;


        // Group
        [FindsBy(How = How.XPath, Using = GROUP_TAB)]
        private IWebElement _group_tab;

        [FindsBy(How = How.XPath, Using = GROUP_TABLE)]
        private IWebElement _groupTable;

        [FindsBy(How = How.XPath, Using = ADD_GROUP)]
        private IWebElement _addGroupBtn;

        [FindsBy(How = How.Id, Using = GROUP_CODE)]
        private IWebElement _groupCode;

        [FindsBy(How = How.Id, Using = GROUP_NAME)]
        private IWebElement _groupName;

        [FindsBy(How = How.Id, Using = SAVE_GROUP)]
        private IWebElement _saveGroup;

        // SubGroup
        [FindsBy(How = How.XPath, Using = SUBGROUP_TAB)]
        private IWebElement _subGroup_tab;

        [FindsBy(How = How.Id, Using = SUB_GROUP_CODE)]
        private IWebElement _subGroup_code;

        [FindsBy(How = How.Id, Using = SUB_GROUP_NAME)]
        private IWebElement _subGroup_name;

        [FindsBy(How = How.XPath, Using = NEW_SUBGROUP_BTN)]
        private IWebElement _subGroup_btn;

        [FindsBy(How = How.Id, Using = SUB_GROUP_SAVE_BTN)]
        private IWebElement _subGroup_savebtn;

        // Inventory
        [FindsBy(How = How.XPath, Using = INVENTORY_TAB)]
        private IWebElement _inventory_tab;

        [FindsBy(How = How.XPath, Using = INVENTORY_DATE)]
        private IWebElement _editInventoryDate;

        // Guest
        [FindsBy(How = How.XPath, Using = GUEST_TAB)]
        private IWebElement _guest_tab;

        [FindsBy(How = How.XPath, Using = ADD_GUEST)]
        private IWebElement _addGuestBtn;

        [FindsBy(How = How.Id, Using = TYPE_GUEST)]
        private IWebElement _typeOfGuest;

        [FindsBy(How = How.Id, Using = GUEST_HAS_ALLERGY)]
        private IWebElement _guestHasAllergy;

        [FindsBy(How = How.Id, Using = GUEST_IS_ALLERGIC)]
        private IWebElement _guestIsAllergic;

        [FindsBy(How = How.Id, Using = GUEST_ORDER)]
        private IWebElement _guestOrder;

        [FindsBy(How = How.Id, Using = SAVE_GUEST)]
        private IWebElement _saveGuest;

        [FindsBy(How = How.XPath, Using = SAVE_ALLERGY)]
        private IWebElement _saveAllergy;

        [FindsBy(How = How.Id, Using = CONFIRM_DELETE_GUEST)]
        private IWebElement _confirmDeleteGuest;

        // Variant
        [FindsBy(How = How.XPath, Using = VARIANT_TAB)]
        private IWebElement _variant_tab;

        [FindsBy(How = How.XPath, Using = PLUS_MENU_VARIANT)]
        private IWebElement _plusMenuVariant;

        [FindsBy(How = How.XPath, Using = ADD_VARIANT)]
        private IWebElement _addVariantBtn;

        [FindsBy(How = How.Id, Using = VARIANT_GUEST)]
        private IWebElement _variantGuest;

        [FindsBy(How = How.Id, Using = VARIANT_MEAL)]
        private IWebElement _variantMeal;

        [FindsBy(How = How.Id, Using = VARIANT_SITE)]
        private IWebElement _variantSite;

        [FindsBy(How = How.Id, Using = SAVE_VARIANT)]
        private IWebElement _saveVariant;

        [FindsBy(How = How.Id, Using = FILTER_SITE)]
        private IWebElement _siteFilter;

        [FindsBy(How = How.XPath, Using = UNCHECK_ALL_SITE)]
        private IWebElement _uncheckAllSite;

        [FindsBy(How = How.XPath, Using = SITE_SEARCH)]
        private IWebElement _searchSite;

        [FindsBy(How = How.Id, Using = FILTER_GUEST)]
        private IWebElement _guestFilter;

        [FindsBy(How = How.XPath, Using = UNCHECK_ALL_GUEST)]
        private IWebElement _uncheckAllGuests;

        [FindsBy(How = How.XPath, Using = GUEST_SEARCH)]
        private IWebElement _searchGuest;

        [FindsBy(How = How.Id, Using = FILTER_MEAL)]
        private IWebElement _mealFilter;

        [FindsBy(How = How.XPath, Using = UNCHECK_ALL_MEAL)]
        private IWebElement _uncheckAllMeal;

        [FindsBy(How = How.XPath, Using = MEAL_SEARCH)]
        private IWebElement _searchMeal;

        [FindsBy(How = How.XPath, Using = DELETE_VARIANT)]
        private IWebElement _deleteVariant;

        [FindsBy(How = How.Id, Using = CONFIRM_DELETE_VARIANT)]
        private IWebElement _confirmDeleteVariant;

        [FindsBy(How = How.Id, Using = CONFIRM_DELETE_2)]
        private IWebElement _confirmDelete2;

        // MEAL
        [FindsBy(How = How.XPath, Using = MEAL_TAB)]
        private IWebElement _meal_tab;

        [FindsBy(How = How.XPath, Using = NEW_MEAL)]
        private IWebElement _new_meal;

        [FindsBy(How = How.Id, Using = MEAL_NAME)]
        private IWebElement _meal_name;

        [FindsBy(How = How.Id, Using = MEAL_COMMERCIAL_NAME)]
        private IWebElement _meal_commercial_name;

        [FindsBy(How = How.Id, Using = MEAL_ORDER)]
        private IWebElement _meal_order;

        [FindsBy(How = How.XPath, Using = MEAL_HOUR)]
        private IWebElement _meal_hour;

        [FindsBy(How = How.Id, Using = SAVE_MEAL)]
        private IWebElement _saveMeal;

        //FOODPACK
        [FindsBy(How = How.XPath, Using = FOODPACK_TAB)]
        private IWebElement _foodpack_tab;

        [FindsBy(How = How.XPath, Using = NEW_FOODPACK)]
        private IWebElement _new_foodpack;

        [FindsBy(How = How.Id, Using = ADD_NEW_FOODPACK)]
        private IWebElement _add_new_foodpack;

        [FindsBy(How = How.Id, Using = INPUT_NEW_FOODPACK_NAME)]
        private IWebElement _input_new_foodpack_name;

        [FindsBy(How = How.XPath, Using = FOODPACK_EDIT_BUTTON)]
        private IWebElement _foodpack_edit_button;

        [FindsBy(How = How.XPath, Using = FOODPACK_ADD_FOODPACK_BUTTON)]
        private IWebElement _foodpack_add_foodpack_button;

        [FindsBy(How = How.XPath, Using = FOODPACK_ADD_FOODPACK_SELECT)]
        private IWebElement _foodpack_add_foodpack_select;

        [FindsBy(How = How.XPath, Using = FOODPACK_ADD_FOODPACK_VALUE)]
        private IWebElement _foodpack_add_foodpack_value;

        [FindsBy(How = How.XPath, Using = FOODPACK_ADD_WARMING_MODE)]
        private IWebElement _foodpack_addWarmingMode;

        [FindsBy(How = How.XPath, Using = FOODPACK_ADD_WARMING_TEMPERATURE)]
        private IWebElement _foodpack_addWarmingTemperature;

        [FindsBy(How = How.Id, Using = SAVE_FOODPACK)]
        private IWebElement _save_foodpack;

        [FindsBy(How = How.XPath, Using = FOODPACK_BULK_EQUIVALENCE)]
        private IWebElement _bulk_equivalence;

        [FindsBy(How = How.Id, Using = FOODPACK_WARMING_MODE)]
        private IWebElement _warmingMode;

        [FindsBy(How = How.Id, Using = FOODPACK_WARMING_TEMPERATURE)]
        private IWebElement _warmingTemperature;

        //RECIPE TYPE
        [FindsBy(How = How.XPath, Using = RECIPE_TYPE_TAB)]
        private IWebElement _recipe_type_tab;

        [FindsBy(How = How.XPath, Using = RECIPE_TYPE_NAME)]
        private IWebElement _recipe_type_name;

        [FindsBy(How = How.XPath, Using = ADD_RECIPE_TYPE_BUTTON)]
        private IWebElement _add_recipe_type_button;

        [FindsBy(How = How.Id, Using = RECIPE_TYPE_TYPE_INPUT)]
        private IWebElement _recipe_type_type_input;

        [FindsBy(How = How.Id, Using = RECIPE_TYPE_ORDER_INPUT)]
        private IWebElement _recipe_type_order_input;

        [FindsBy(How = How.Id, Using = RECIPE_TYPE_SAVE_BUTTON)]
        private IWebElement _recipe_type_save_button;

        [FindsBy(How = How.XPath, Using = KEYWORD_LINE)]
        private IWebElement _keywordLine;

        [FindsBy(How = How.XPath, Using = KEYWORD_DELETE)]
        private IWebElement _keywordDelete;

        [FindsBy(How = How.Id, Using = CONFIRM_DELETE_KEYWORD)]
        private IWebElement _confirmDeleteKeyword;

        [FindsBy(How = How.XPath, Using = CLOSE_ICON_DELETE_KEYWORD)]
        private IWebElement _closeIconkeywordDelete;

        [FindsBy(How = How.XPath, Using = OK_DELETE_KEYWORD)]
        private IWebElement _okkeywordDelete;

        [FindsBy(How = How.XPath, Using = COOKING_MODE)]
        private IWebElement _cookingMode_tab;
        // ___________________________________________ Méthodes ___________________________________________

        public ParametersProduction(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {

        }

        // Group
        public void GoToTab_Group()
        {
            _group_tab = WaitForElementToBeClickable(By.XPath(GROUP_TAB));
            _group_tab.Click();
            WaitForLoad();
        }

        public void GoToTab_SubGroup()
        {
            _subGroup_tab = WaitForElementToBeClickable(By.XPath(SUBGROUP_TAB));
            _subGroup_tab.Click();
            WaitForLoad();
        }

        public void AddNewGroup(string groupName, string groupCode)
        {
            _addGroupBtn = WaitForElementIsVisible(By.XPath(ADD_GROUP));
            _addGroupBtn.Click();
            WaitForLoad();

            _groupCode = WaitForElementIsVisible(By.Id(GROUP_CODE));
            _groupCode.SetValue(ControlType.TextBox, groupCode);

            _groupName = WaitForElementIsVisible(By.Id(GROUP_NAME));
            _groupName.SetValue(ControlType.TextBox, groupName);

            _saveGroup = WaitForElementIsVisible(By.Id(SAVE_GROUP));
            _saveGroup.Click();
            WaitForLoad();
        }
        public void AddNewSubGroup(string groupName, string groupCode)
        {
            //if (IsDev())
            //{
            _subGroup_btn = WaitForElementIsVisible(By.XPath("//*[@id=\"tabContentParameters\"]/div[1]/div/div[2]/a"));
            //}
            //else
            //{
            //_subGroup_btn = WaitForElementIsVisible(By.XPath(NEW_SUBGROUP_BTN));
            //}
            _subGroup_btn.Click();
            WaitForLoad();

            _subGroup_code = WaitForElementIsVisible(By.Id(SUB_GROUP_CODE));
            _subGroup_code.SetValue(ControlType.TextBox, groupCode);

            _subGroup_name = WaitForElementIsVisible(By.Id(SUB_GROUP_NAME));
            _subGroup_name.SetValue(ControlType.TextBox, groupName);

            _subGroup_savebtn = WaitForElementIsVisible(By.Id(SUB_GROUP_SAVE_BTN));
            _subGroup_savebtn.Click();
            WaitForLoad();
        }

        public bool IsGroupPresent(string groupName)
        {
            try
            {
                _webDriver.FindElement(By.XPath(String.Format(GROUP_LINE, groupName)));
            }
            catch
            {
                return false;
            }

            return true;
        }

        // Inventory
        public void GoToTab_Inventory()
        {
            _inventory_tab = WaitForElementToBeClickable(By.XPath(INVENTORY_TAB));
            _inventory_tab.Click();
            WaitForLoad();
        }

        public bool SetInventoryValue(int jour)
        {
            // Recherche de la ligne correspondant au jour en cours
            var inventoryModale = FindDate(jour);

            // On coche la case de l'inventory pour pouvoir exécuter le test
            return inventoryModale.SetAllowed();
        }

        public void RemoveInventoryValue(int jour)
        {
            // Recherche de la ligne correspondant au jour en cours
            var inventoryModale = FindDate(jour);

            // On coche la case de l'inventory pour pouvoir exécuter le test
            inventoryModale.RemoveAllowed();
        }

        public ParametersProductionModalPage FindDate(int jour)
        {
            try
            {
                // On click sur l'icône Editer de la ligne correspondante au jour en cours
                int date = jour + 1;

                _editInventoryDate = WaitForElementIsVisible(By.XPath(String.Format(INVENTORY_DATE, date.ToString())));
                _editInventoryDate.Click();
                WaitForLoad();

                return new ParametersProductionModalPage(_webDriver, _testContext);
            }
            catch
            {
                throw new Exception("Le jour en cours n'a pas été trouvé dans la liste des dates dans l'application.");
            }
        }

        // Guest
        public void GoToTab_Guest()
        {
            _guest_tab = WaitForElementToBeClickable(By.XPath(GUEST_TAB));
            _guest_tab.Click();
            WaitPageLoading();
            WaitForLoad();
        }
        public bool AddNewGuest(string guestName, bool hasAllergy, string order)
        {
            _addGuestBtn = WaitForElementIsVisible(By.XPath(ADD_GUEST));
            _addGuestBtn.Click();
            WaitForLoad();

            _typeOfGuest = WaitForElementIsVisible(By.Id(TYPE_GUEST));
            _typeOfGuest.SetValue(ControlType.TextBox, guestName);

            _guestHasAllergy = WaitForElementExists(By.Id(GUEST_HAS_ALLERGY));
            _guestHasAllergy.SetValue(ControlType.CheckBox, hasAllergy);

            _guestOrder = WaitForElementIsVisible(By.Id(GUEST_ORDER));
            _guestOrder.SetValue(ControlType.TextBox, order);

            _saveGuest = WaitForElementIsVisible(By.Id(SAVE_GUEST));
            _saveGuest.Click();
            WaitForLoad();

            return true;
        }

        public bool IsGuestPresent(string guestName)
        {
            try
            {
                _webDriver.FindElement(By.XPath(String.Format(GUEST_LINE, guestName)));
            }
            catch
            {
                return false;
            }

            return true;
        }

        public void SetAllergy(string allergy)
        {
            var guestAllergy = WaitForElementIsVisible(By.XPath(String.Format(GUEST_ALLERGY_LINE, allergy)));
            guestAllergy.Click();

            _saveAllergy = WaitForElementIsVisible(By.XPath(SAVE_ALLERGY));
            _saveAllergy.Click();
            WaitForLoad();
        }

        public void DeleteGuest(string guestName)
        {
            Actions action = new Actions(_webDriver);

            var guestLine = WaitForElementIsVisible(By.XPath(String.Format(GUEST_LINE, guestName)));
            action.MoveToElement(guestLine).Perform();

            var guestDelete = WaitForElementIsVisible(By.XPath(String.Format(GUEST_DELETE, guestName)));
            guestDelete.Click();
            WaitForLoad();

            _confirmDeleteGuest = WaitForElementIsVisible(By.Id(CONFIRM_DELETE_GUEST));
            _confirmDeleteGuest.Click();
            WaitForLoad();
        }
        public void EditGuest(string guestName)
        {
            Actions action = new Actions(_webDriver);

            var guestLine = WaitForElementIsVisible(By.XPath(String.Format(GUEST_LINE, guestName)));
            action.MoveToElement(guestLine).Perform();
            IWebElement editGuest;
            if (isElementVisible(By.XPath(String.Format(GUEST_EDIT, guestName))))
            {
                editGuest = WaitForElementIsVisible(By.XPath(String.Format(GUEST_EDIT, guestName)));
            }
            else
            {
                editGuest = WaitForElementIsVisible(By.XPath(String.Format(GUEST_EDIT_2, guestName)));
            }
            editGuest.Click();
            WaitForLoad();
        }

        public void AddAllergenToGuest(string allergen)
        {
            _guestHasAllergy = WaitForElementExists(By.Id(GUEST_HAS_ALLERGY));
            _guestHasAllergy.SetValue(ControlType.CheckBox, true);

            _saveGuest = WaitForElementIsVisible(By.Id(SAVE_GUEST));
            _saveGuest.Click();
            WaitForLoad();

            SetAllergy(allergen);
        }
        public void AddIsClinicToGuest()
        {
            _guestHasAllergy = WaitForElementExists(By.Id("IsClinic"));
            _guestHasAllergy.SetValue(ControlType.CheckBox, true);

            _saveGuest = WaitForElementIsVisible(By.Id(SAVE_GUEST));
            _saveGuest.Click();
            WaitForLoad();
        }
        public void AddGuestTypeTypeToGuest(string type)
        {
            _guestHasAllergy = WaitForElementExists(By.Id("GuestTypeType"));
            _guestHasAllergy.SetValue(ControlType.DropDownList, type);

            _saveGuest = WaitForElementIsVisible(By.Id(SAVE_GUEST));
            _saveGuest.Click();
            WaitForLoad();
        }

        public bool IsErrorMessageAllergen()
        {
            try
            {
                var _errorMeaageAllergen = WaitForElementIsVisible(By.XPath("//*[@id=\"GuestTypeModal\"]/div/div/div/div/div/div/form/ul[@class='errorMessage']"));
                return true;
            }
            catch
            {
                return false;
            }
        }

        // Variant
        public void GoToTab_Variant()
        {
            WaitForLoad();
            _variant_tab = WaitForElementToBeClickable(By.XPath(VARIANT_TAB));
            _variant_tab.Click();
            WaitPageLoading();
        }
        public bool AddNewVariant(string guestName, string meal, string site)
        {
            Actions action = new Actions(_webDriver);

            _plusMenuVariant = WaitForElementIsVisible(By.XPath(PLUS_MENU_VARIANT));
            action.MoveToElement(_plusMenuVariant).Perform();

            _addVariantBtn = WaitForElementIsVisible(By.XPath(ADD_VARIANT));
            _addVariantBtn.Click();
            WaitForLoad();

            _variantGuest = WaitForElementIsVisible(By.Id(VARIANT_GUEST));
            _variantGuest.SetValue(ControlType.DropDownList, guestName);
            _variantGuest.SendKeys(Keys.Tab);

            _variantMeal = WaitForElementExists(By.Id(VARIANT_MEAL));
            _variantMeal.SetValue(ControlType.DropDownList, meal);
            _variantMeal.SendKeys(Keys.Tab);

            _variantSite = WaitForElementIsVisible(By.Id(VARIANT_SITE));
            _variantSite.SetValue(ControlType.DropDownList, site);
            _variantSite.SendKeys(Keys.Tab);

            _saveVariant = WaitForElementIsVisible(By.Id(SAVE_VARIANT));
            _saveVariant.Click();
            WaitForLoad();

            return true;
        }
        public void DeleteVariant(string guestName)
        {
            bool isResult = true;

            FilterGuestType(guestName);

            try
            {
                _deleteVariant = _webDriver.FindElement(By.XPath(DELETE_VARIANT));
                _deleteVariant.Click();
                WaitForLoad();
            }
            catch
            {
                isResult = false;
            }

            if (isResult)
            {
                _confirmDeleteVariant = WaitForElementIsVisible(By.Id(CONFIRM_DELETE_VARIANT));
                _confirmDeleteVariant.Click();
                WaitForLoad();

                _confirmDelete2 = WaitForElementIsVisible(By.Id(CONFIRM_DELETE_2));
                _confirmDelete2.Click();
                WaitForLoad();
            }
        }
        public bool IsVariantPresent(string guestName, string meal)
        {

            var variants = _webDriver.FindElements(By.XPath(VARIANTS));

            foreach (var elm in variants)
            {
                if (elm.Text.Contains(guestName) && elm.Text.Contains(meal))
                    return true;
            }

            return false;
        }
        public void filterVariant(string guestName, string meal, string site)
        {

        }

        public void FilterSite(string site)
        {
            ComboBoxSelectById(new ComboBoxOptions(FILTER_SITE, site,false));
            WaitPageLoading();

        }

        public void FilterGuestType(string guestName)
        {
            ComboBoxSelectById(new ComboBoxOptions(FILTER_GUEST, guestName,false));
            WaitPageLoading();
        }

        public void FilterMealType(string meal)
        {
            ComboBoxSelectById(new ComboBoxOptions(FILTER_MEAL, meal, false));
        }

        //Meal
        public bool IsMealPresent(string meal)
        {

            var meals = _webDriver.FindElements(By.XPath(MEALS));

            foreach (var elm in meals)
            {
                if (elm.Text.Contains(meal))
                    return true;
            }

            return false;
        }

        public void GoToTab_Meal()
        {
            _meal_tab = WaitForElementToBeClickable(By.XPath(MEAL_TAB));
            _meal_tab.Click();
            WaitForLoad();
        }

        public void CreateNewMeal(string mealName, string mealCommercialName, string mealOrder, string mealHour)
        {
            _new_meal = WaitForElementToBeClickable(By.XPath(NEW_MEAL));
            _new_meal.Click();

            _meal_name = WaitForElementToBeClickable(By.Id(MEAL_NAME));
            _meal_name.SetValue(ControlType.TextBox, mealName);

            _meal_commercial_name = WaitForElementToBeClickable(By.Id(MEAL_COMMERCIAL_NAME));
            _meal_commercial_name.SetValue(ControlType.TextBox, mealCommercialName);

            _meal_order = WaitForElementToBeClickable(By.Id(MEAL_ORDER));
            _meal_order.SetValue(ControlType.TextBox, mealOrder);

            _meal_hour = WaitForElementIsVisible(By.XPath(MEAL_HOUR));

            _meal_hour.SendKeys(Keys.ArrowRight);
            _meal_hour.SendKeys(Keys.ArrowRight);
            _meal_hour.SendKeys(Keys.Backspace);
            _meal_hour.SendKeys(Keys.Backspace);
            _meal_hour.SendKeys(mealHour);

            _saveMeal = WaitForElementToBeClickable(By.Id(SAVE_MEAL));
            _saveMeal.Click();
        }

        //Foodpack Tab
        public void GoToTab_Foodpack()
        {
            _foodpack_tab = WaitForElementToBeClickable(By.XPath(FOODPACK_TAB));
            _foodpack_tab.Click();
            WaitForLoad();
        }
        public bool CheckFoodPackExists(string foodpackName)
        {
            bool foodpackExist = false;

            if (isElementVisible(By.XPath(String.Format(FOODPACK_NAME, foodpackName))))
            {
                foodpackExist = true;
            }
            else
            {
                //foodpack not present - bool still false
            }

            return foodpackExist;
        }

        public void CreateNewFoodPack(string foodpackName)
        {
            _new_foodpack = WaitForElementToBeClickable(By.XPath(NEW_FOODPACK));
            _new_foodpack.Click();

            _input_new_foodpack_name = WaitForElementToBeClickable(By.Id(INPUT_NEW_FOODPACK_NAME));
            _input_new_foodpack_name.SetValue(ControlType.TextBox, foodpackName);

            _add_new_foodpack = WaitForElementToBeClickable(By.Id(SAVE_FOODPACK));
            _add_new_foodpack.Click();
            WaitForLoad();

            Assert.IsTrue(CheckFoodPackExists(foodpackName), "Le Food pack {0} n'a pas été créé", foodpackName);
        }
        public void AddBulkEquivalenceToFoodPack(string foodpackName, string foodpackToAdd, string valueToAdd)
        {
            _foodpack_edit_button = WaitForElementToBeClickable(By.XPath(String.Format(FOODPACK_EDIT_BUTTON, foodpackName)));
            _foodpack_edit_button.Click();
            WaitForLoad();

            _foodpack_add_foodpack_button = WaitForElementToBeClickable(By.XPath(FOODPACK_ADD_FOODPACK_BUTTON));
            _foodpack_add_foodpack_button.Click();
            WaitForLoad();

            _foodpack_add_foodpack_select = WaitForElementToBeClickable(By.XPath(FOODPACK_ADD_FOODPACK_SELECT));
            _foodpack_add_foodpack_select.SetValue(ControlType.DropDownList, foodpackToAdd);
            WaitForLoad();

            _foodpack_add_foodpack_value = WaitForElementToBeClickable(By.XPath(FOODPACK_ADD_FOODPACK_VALUE));
            _foodpack_add_foodpack_value.SetValue(ControlType.TextBox, valueToAdd);
            WaitForLoad();

            _save_foodpack = WaitForElementToBeClickable(By.Id(SAVE_FOODPACK));
            _save_foodpack.Click();
            WaitForLoad();
        }
        public string GetBulkEquivalence(string foodpackName)
        {
            _bulk_equivalence = WaitForElementExists(By.XPath(String.Format(FOODPACK_BULK_EQUIVALENCE, foodpackName)));
            return _bulk_equivalence.Text;
        }
        public void AddWarmingModeToFoodPack(string foodpackName, string warmingModeValue)
        {
            _foodpack_edit_button = WaitForElementToBeClickable(By.XPath(String.Format(FOODPACK_EDIT_BUTTON, foodpackName)));
            _foodpack_edit_button.Click();
            WaitForLoad();

            _foodpack_addWarmingMode = WaitForElementIsVisible(By.Id(FOODPACK_ADD_WARMING_MODE), nameof(FOODPACK_ADD_WARMING_MODE));
            _foodpack_addWarmingMode.SetValue(ControlType.TextBox, warmingModeValue);

            _save_foodpack = WaitForElementToBeClickable(By.Id(SAVE_FOODPACK));
            _save_foodpack.Click();
            WaitForLoad();
        }

        public string GetWarmingMode(string foodpackName)
        {
            _warmingMode = WaitForElementExists(By.XPath(String.Format(FOODPACK_WARMING_MODE, foodpackName)));
            return _warmingMode.Text;
        }
        public void AddWarmingTemperatureToFoodPack(string foodpackName, string warmingTemperatureValue)
        {
            _foodpack_edit_button = WaitForElementToBeClickable(By.XPath(String.Format(FOODPACK_EDIT_BUTTON, foodpackName)));
            _foodpack_edit_button.Click();
            WaitForLoad();

            _foodpack_addWarmingTemperature = WaitForElementIsVisible(By.Id(FOODPACK_ADD_WARMING_TEMPERATURE), nameof(FOODPACK_ADD_WARMING_TEMPERATURE));
            _foodpack_addWarmingTemperature.SetValue(ControlType.TextBox, warmingTemperatureValue);

            _save_foodpack = WaitForElementToBeClickable(By.Id(SAVE_FOODPACK));
            _save_foodpack.Click();
            WaitForLoad();
        }

        public string GetWarmingTemperature(string foodpackName)
        {
            _warmingTemperature = WaitForElementExists(By.XPath(String.Format(FOODPACK_WARMING_TEMPERATURE, foodpackName)));
            return _warmingTemperature.Text;
        }

        //Recipe Type Tab
        public void GoToTab_RecipeType()
        {
            _recipe_type_tab = WaitForElementToBeClickable(By.XPath(RECIPE_TYPE_TAB));
            _recipe_type_tab.Click();
        }
        public bool CheckRecipeTypeExists(string recipeType)
        {
            bool recypeTypeExist = false;

            try
            {
                _recipe_type_name = WaitForElementExists(By.XPath(String.Format(RECIPE_TYPE_NAME, recipeType)));
                recypeTypeExist = true;
            }
            catch
            {
                //recype type not present - bool still false
            }

            return recypeTypeExist;
        }
        public void CreateNewRecipeType(string recipeType, string order)
        {
            _add_recipe_type_button = WaitForElementToBeClickable(By.XPath(ADD_RECIPE_TYPE_BUTTON));
            _add_recipe_type_button.Click();

            _recipe_type_type_input = WaitForElementExists(By.Id(RECIPE_TYPE_TYPE_INPUT));
            _recipe_type_type_input.SetValue(ControlType.TextBox, recipeType);

            _recipe_type_order_input = WaitForElementExists(By.Id(RECIPE_TYPE_ORDER_INPUT));
            _recipe_type_order_input.SetValue(ControlType.TextBox, order);

            _recipe_type_save_button = WaitForElementToBeClickable(By.Id(RECIPE_TYPE_SAVE_BUTTON));
            _recipe_type_save_button.Click();
            WaitForLoad();

            Assert.IsTrue(CheckRecipeTypeExists(recipeType), "Le recipe type {0} n'a pas été créé", recipeType);
        }

        public void WorkshopTab()
        {
            var workshopTab = WaitForElementIsVisible(By.XPath(WORKSHOP_TAB));
            workshopTab.Click();
        }
        public bool VerifyWorkshopExist(params string[] values)
        {
            WaitForLoad();
            var listWorkshops = _webDriver.FindElements(By.XPath(WORKSHOPS));

            foreach (var value in values)
            {
                var list = listWorkshops.Where(w => w.Text == value);
                if (list.Count() == 0)
                {
                    return false;
                }
            }
            return true;
        }
        public void CreateNewWorkshop(string workshop)
        {
            Random rnd = new Random();

            var btn = WaitForElementExists(By.XPath(CREATE_WORKSHOP));
            btn.Click();
            var input = WaitForElementExists(By.Id(CREATE_INPUT));
            input.SendKeys(workshop);
            input.SendKeys(Keys.Tab);
            var workshopType = WaitForElementIsVisible(By.XPath(WORKSHOP_TYPE));
            workshopType.SendKeys(Keys.Tab);
            var order = WaitForElementIsVisible(By.XPath(ORDER_INPUT));
            order.SendKeys((rnd.Next()).ToString());
            var save = WaitForElementExists(By.XPath(CONFIRM_SAVE));
            save.Click();
        }
        public bool isGroupProduct(string groupName)
        {
            IWebElement isproductGroup;
            isproductGroup = WaitForElementExists(By.XPath(string.Format(ISPRODUCT_GROUP, groupName)));
            return isproductGroup.Selected;
        }
        public bool isGroupFood(string groupName)
        {
            var isproductFood = WaitForElementExists(By.XPath(string.Format(ISFOOD_GROUP, groupName)));
            return isproductFood.Selected;
        }
        public void EditGroup(string groupName, bool isProduct, bool isFood)
        {
            var editGroup = WaitForElementExists(By.XPath(string.Format(EDIT_GROUP, groupName)));
            editGroup.Click();
            WaitForLoad();

            if (isProduct)
            {
                var isProductCheck = WaitForElementExists(By.XPath("//*[@id=\"IsProduct\"]"));
                isProductCheck.Click();
                WaitForLoad();
            }
            if (isFood)
            {
                var isFoodCheck = WaitForElementExists(By.XPath("//*[@id=\"IsFood\"]"));
                isFoodCheck.Click();
                WaitForLoad();
            }

            CloseEditModal();
        }
        public void CloseEditModal()
        {
            var save = WaitForElementIsVisible(By.XPath("//*[@id=\"last\"]"));
            save.Click();
            WaitForLoad();
        }
        public List<string> GetVariantListOfGuests()
        {
            var _listOfGuests = _webDriver.FindElements(By.XPath(VARIANTS_LIST));
            var _listOfGuestsNames = _listOfGuests.Select(guest => guest.Text).ToList();
            return _listOfGuestsNames;

        }

        public string GetFirstVariant()
        {
            var _firsGuest = WaitForElementIsVisible(By.XPath(FIRST_GUEST_VARIANT));
            var _firsMeal = WaitForElementIsVisible(By.XPath(FIRST_MEAL_VARIANT));
              var guestName =_firsGuest.Text;
              var mealName = _firsMeal.Text;
            return $"{guestName} - {mealName}"; ;
        }

        //Keyword Tab
        public void GoToTab_Keyword()
        {
            _keywordtab = WaitForElementToBeClickable(By.XPath(KEYWORD_TAB));
            _keywordtab.Click();
            WaitForLoad();
        }

        public void GoToTab_Allergen()
        { 
                _allergentab = WaitForElementToBeClickable(By.XPath(ALLERGEN_TAB));        
                _allergentab.Click();
            WaitForLoad();
        }

        public void CreateNewAllergen(string AllergenName, string NameEnglish)
        {
            _newallergen = WaitForElementToBeClickable(By.XPath(NEW_ALLERGEN_BTN));
            _newallergen.Click();
            WaitForLoad();

            _nameinputallergen = WaitForElementIsVisibleNew(By.XPath(NAME_INPUT));
            _nameinputallergen.SetValue(ControlType.TextBox, AllergenName);
            WaitForLoad();

            _engnameinputallergen = WaitForElementIsVisibleNew(By.XPath(ENGNAME_INPUT));
            _engnameinputallergen.SetValue(ControlType.TextBox, NameEnglish);
            WaitForLoad();

            _saveallergen = WaitForElementToBeClickable(By.XPath(SAVE_BTN));
            _saveallergen.Click();
            WaitForLoad();
        }

        public void ClickEditAllergen(string AllergenName)
        {
            Actions action = new Actions(_webDriver);

            var allergenRow = WaitForElementExists(By.XPath($"//*[@id='tabContentParameters']/table/tbody/tr[td[contains(text(), '{AllergenName}')]]"));

            action.MoveToElement(allergenRow).Perform();

            // Localisation du bouton d'édition correspondant à l'allergène
            var allergenBox = allergenRow.FindElement(By.XPath(".//a[contains(@class, 'btn btn-default')]"));

            if (allergenBox != null && allergenBox.Displayed)
            {
                allergenBox.Click();
            }
            else
            {
                throw new NoSuchElementException("Le bouton d'édition pour l'allergène n'a pas été trouvé.");
            }

            // Attente du chargement de la page après le clic
            WaitPageLoading();
        }

        public void ClickDeleteAllergen(string allergenName)
        {
            Actions action = new Actions(_webDriver);

            // Locate the row that contains the allergen name
            var allergenRow = WaitForElementExists(By.XPath($"//*[@id='tabContentParameters']/table/tbody/tr[td[contains(text(), '{allergenName}')]]"));

            // Ensure the row was found and move the mouse to it
            action.MoveToElement(allergenRow).Perform();

            // Find the delete button within the row (the exact XPath for the delete button)
            var deleteButton = allergenRow.FindElement(By.XPath(".//td[4]/a[2]"));

            if (deleteButton != null && deleteButton.Displayed)
            {
                // Cliquer sur le bouton de suppression
                deleteButton.Click();
            }
            else
            {
                throw new NoSuchElementException($"Le bouton de suppression pour l'allergène '{allergenName}' n'a pas été trouvé.");
            }

            WaitPageLoading();
        }

        public void ChangeAllergenName(string newAllergenName)
        {
            var allergenNameField = _webDriver.FindElement(By.Id("first"));
            allergenNameField.Clear();
            allergenNameField.SendKeys(newAllergenName);
            var saveButton = _webDriver.FindElement(By.Id("last")); 
            saveButton.Click();
        }

        public void AddKeyword(string keywordName)
        {
            _createkeyword = WaitForElementToBeClickable(By.XPath(CREATE_KEYWORD));
            _createkeyword.Click();
            WaitForLoad();

            _valueinputkeyword = WaitForElementIsVisible(By.XPath(VALUE_INPUT));
            _valueinputkeyword.SetValue(ControlType.TextBox, keywordName);
            WaitForLoad();

            _displayeat = WaitForElementExists(By.XPath(DISPLAY_EAT));
            _displayeat.SetValue(ControlType.CheckBox, true);

            _savekeyword = WaitForElementToBeClickable(By.Id(SAVE));
            _savekeyword.Click();
            WaitForLoad();
        }

        public bool CheckKeyword(string keywordName)
        {
            return isElementVisible(By.XPath(String.Format(CHECK_KEYWORD, keywordName))) ? true : false;        
        }

        public void UncheckAllAllergen(string allergy)
        {
            var guestAllergy = WaitForElementIsVisible(By.XPath(DEFINE_FIRST_ALLERGENS));
            guestAllergy.Click();
            WaitPageLoading();

            var checkallallergenh = _webDriver.FindElements(By.XPath(ALLERGEN_ALL_LINE));


            foreach (var element in checkallallergenh)
            {
                var allergenBox = _webDriver.FindElement(By.Id(element.GetAttribute("for")));
                if (allergenBox.GetAttribute("checked") == "true")
                    allergenBox.Click();
            }
            WaitPageLoading();
            var saveChanges = WaitForElementIsVisible(By.XPath(SAVE_ALLEERGY_GUEST));
            var html = _webDriver.FindElement(By.TagName("html"));
            html.SendKeys(Keys.PageDown);
            saveChanges.Click();
            WaitPageLoading();
        }

        public void AddAllergen(string allergen)
        {
            WaitPageLoading();
            Actions action = new Actions(_webDriver);

            var guestAllergy = WaitForElementToBeClickable(By.XPath(DEFINE_FIRST_ALLERGENS));
            action.MoveToElement(guestAllergy).Click().Perform();
            WaitPageLoading();


            var _allergenLine = WaitForElementExists(By.XPath(String.Format(ALLERGEN_LINE, allergen)));
            action.MoveToElement(_allergenLine).Perform();

            var allergenBox = _webDriver.FindElement(By.Id(_allergenLine.GetAttribute("for")));
            if (allergenBox.GetAttribute("checked") != "true")
            {
                allergenBox.Click();
            }

            var saveChanges = WaitForElementIsVisible(By.XPath(SAVE_ALLEERGY_GUEST));
            var html = _webDriver.FindElement(By.TagName("html"));
            html.SendKeys(Keys.PageDown);
            saveChanges.Click();
        }
        public bool DeleteKeyword(string keyword, string OkToDelete)
        {
            Actions action = new Actions(_webDriver);

            var KeywordLine = WaitForElementIsVisible(By.XPath(String.Format(KEYWORD_LINE, keyword)));
            action.MoveToElement(KeywordLine).Perform();

            var KeywordDelete = WaitForElementIsVisible(By.XPath(String.Format(KEYWORD_DELETE, keyword)));
            KeywordDelete.Click();
            WaitForLoad();

            _confirmDeleteKeyword = WaitForElementIsVisible(By.Id(CONFIRM_DELETE_KEYWORD));
            _confirmDeleteKeyword.Click();
            WaitForLoad();

            var KeywordMsgDelete = WaitForElementIsVisible(By.XPath(OK_DELETE_KEYWORD));

            if (OkToDelete != null && KeywordMsgDelete.Text == OkToDelete)
            {

                _closeIconkeywordDelete = WaitForElementIsVisible(By.XPath(CLOSE_ICON_DELETE_KEYWORD));
                _closeIconkeywordDelete.Click();
                WaitForLoad();
                return true;
            }

            else
                _closeIconkeywordDelete = WaitForElementIsVisible(By.XPath(CLOSE_ICON_DELETE_KEYWORD));
                _closeIconkeywordDelete.Click();
                WaitForLoad();
                return false;

        }
        public bool IsGroupExist(string groupName)
        {
            if (!isElementVisible(By.XPath(string.Format(GROUP_TABLE, groupName))))
            {
                return false;
            }
            return true;
        }
        public void EnsureSubGroup(string subgroup, string group)
        {
            var listSubGroup = _webDriver.FindElements(By.XPath(LIST_SUB_GROUP));
            bool trouve = false;
            foreach (var ligne in listSubGroup)
            {
                if (ligne.Text == subgroup)
                {
                    trouve = true;
                }
            }
            if (!trouve)
            {
                var add = WaitForElementIsVisible(By.XPath(NEW_SUBGROUP_BTN));
                add.Click();
                var code = WaitForElementIsVisible(By.Id(SUB_GROUP_CODE));
                code.SetValue(PageBase.ControlType.TextBox, subgroup);
                var name = WaitForElementIsVisible(By.Id(SUB_GROUP_NAME));
                name.SetValue(PageBase.ControlType.TextBox, subgroup);
                var groupSelect = WaitForElementIsVisible(By.XPath(GROUP_SELECT));
                groupSelect.SetValue(PageBase.ControlType.DropDownList, group);
                var save = WaitForElementIsVisible(By.Id(SUB_GROUP_SAVE_BTN));
                save.Click();
            }
        }
        public void GoToTab_CookingMode()
        {
            _cookingMode_tab = WaitForElementToBeClickable(By.XPath(COOKING_MODE));
            _cookingMode_tab.Click();
            WaitForLoad();
        }

        public bool isVacuumPackedCooking(string cookingMode)
        {
            var findCookingMode = _webDriver.FindElement(By.XPath(string.Format("//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[contains(text(),'{0}')]/../td[2]/div/input", cookingMode)));
            var cookingModeSelected = findCookingMode.Selected;
            return cookingModeSelected;
        }
    }
}