using DocumentFormat.OpenXml.Bibliography;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Threading;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Production.Production
{
    public class MappedDetailsByMenu
    {
        public string menuPAX { get; set; }
        public string recipeName { get; set; }
        public string serviceName { get; set; }
        public string deliveryName { get; set; }
        public string foodPack { get; set; }
        public string recipePAX { get; set; }
        public string qtyToProduce { get; set; }
        public string itemName { get; set; }
        public string itemNetWeight { get; set; }
    }

    public class ProductionSearchResultsMenuTabPage : PageBase
    {
        public ProductionSearchResultsMenuTabPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // ______________________________________ Constantes _____________________________________________


        // ONGLETS

        private const string RECIPE_TAB = "//*[@id=\"itemTabTab\"]/li[2]";
        private const string CUSTOMER_TAB = "//*[@id=\"itemTabTab\"]/li[3]";
        private const string WORKSHOP_TAB = "//*[@id=\"itemTabTab\"]/li[4]";
        private const string CUSTOMER_ORDER_TAB = "//*[@id=\"itemTabTab\"]/li[5]";


        // FILTRES

        private const string RESET_FILTER = "//*[@id=\"item-filter-form\"]/div[1]/a";
        private const string SEARCH_FILTER = "SearchPattern";
        private const string DATE_FILTER = "StartDate";
        private const string DATE_LABEL_TO_CLICK = "//*[@id=\"item-filter-form\"]/h3[1]";
        private const string SITE_FILTER = "SiteId";

        private const string SHOW_ITEMS_FILTER = "//*[@id=\"item-filter-form\"]/div[6]/div";
        private const string SHOW_PROCEDURE_FILTER = "WithProcedure";
        private const string SHOW_PAX_IS_ZERO_FILTER = "ShowPAXIsZero";

        private const string CUSTOMER_TYPES_FILTER = "SelectedCustomerTypesIds_ms";
        private const string CUSTOMER_TYPES_FILTER_UNCHECK_ALL = "/html/body/div[11]/div/ul/li[2]/a/span[2]";
        private const string CUSTOMER_TYPES_FILTER_SEARCH = "/html/body/div[11]/div/div/label/input";

        private const string CUSTOMERS_FILTER = "SelectedCustomersIds_ms";
        private const string CUSTOMERS_FILTER_UNCHECK_ALL = "/html/body/div[12]/div/ul/li[2]/a/span[2]";
        private const string CUSTOMERS_FILTER_SEARCH = "/html/body/div[12]/div/div/label/input";

        private const string WORKSHOPS_FILTER = "//*[@id=\"SelectedWorkshopsIds_ms\"]";
        private const string WORKSHOPS_FILTER_UNCHECK_ALL = "/html/body/div[13]/div/ul/li[2]/a/span[2]";
        private const string WORKSHOPS_FILTER_SEARCH = "/html/body/div[13]/div/div/label/input";

        private const string MEALTYPES_FILTER = "SelectedMealTypesIds_ms";
        private const string MEALTYPES_FILTER_UNCHECK_ALL = "/html/body/div[14]/div/ul/li[2]/a/span[2]";
        private const string MEALTYPES_FILTER_SEARCH = "/html/body/div[14]/div/div/label/input";

        private const string ITEM_GROUPS_FILTER = "//*[@id=\"SelectedItemGroupsIds_ms\"]";
        private const string ITEM_GROUPS_FILTER_UNCHECK_ALL = "/html/body/div[15]/div/ul/li[2]/a/span[2]";
        private const string ITEM_GROUPS_FILTER_SEARCH = "/html/body/div[15]/div/ul/li[2]/a/span[2]";

        private const string GUESTTYPES_FILTER = "SelectedGuestTypesIds_ms";
        private const string GUESTTYPES_FILTER_UNCHECK_ALL = "/html/body/div[17]/div/ul/li[2]/a/span[2]";
        private const string GUESTTYPES_FILTER_SEARCH = "/html/body/div[17]/div/div/label/input";

        private const string DELIVERY_ROUNDS_FILTER = "SelectedDeliveryRoundsIds_ms";
        private const string DELIVERY_ROUNDS_FILTER_UNCHECK_ALL = "/html/body/div[18]/div/ul/li[2]/a/span[2]";
        private const string DELIVERY_ROUNDS_FILTER_SEARCH = "/html/body/div[18]/div/div/label/input";

        private const string FOODPACK_TYPES_FILTER = "SelectedFoodPackTypesIds_ms";
        private const string FOODPACK_TYPES_FILTER_UNCHECK_ALL = "/html/body/div[19]/div/ul/li[2]/a/span[2]";
        private const string FOODPACK_TYPES_FILTER_SEARCH = "/html/body/div[19]/div/div/label/input";

        private const string DELIVERIES_FILTER = "SelectedDeliveriesIds_ms";
        private const string DELIVERIES_FILTER_UNCHECK_ALL = "/html/body/div[20]/div/ul/li[2]/a/span[2]";
        private const string DELIVERIES_FILTER_SEARCH = "/html/body/div[20]/div/div/label/input";


        // PRINT
        private const string BTN_PRINT_ALLOTMENT = "printAllotment";
        private const string BTN_PRINT_ALLOTMENT_NON_PACKAGED = "printAllotmentNC";
        private const string BTN_PRINT_ALLOTMENT_GLOBAL_NON_PACKAGED = "printAllotmentNCGlobal";
        private const string BTN_PRINT_HACCP = "printHACCP";
        private const string BTN_PRINT_RAWMATERIALS = "printItemGroup";
        private const string BTN_PRINT_PRODUCTION = "printProd";
        private const string BTN_PRINT_PRODUCTION_VALIDATE = "validatePrint";


        // TABLEAU BY MENU
        private const string NUMBER_OF_PRODUCTION = "//*[@id=\"tabContentItemContainer\"]/div[1]/h1/span";
        private const string FIRST_MENU_NAME = "//*[@id=\"list-item-with-action\"]/div[2]/div[1]/div/div[2]/table/tbody/tr/td[1]";
        private const string FIRST_MENU_VARIANT = "//*[@id=\"list-item-with-action\"]/div[2]/div[1]/div/div[2]/table/tbody/tr/td[2]";
        private const string TOGGLE_FIRST_MENU_DISPLAYED = "//*[@id=\"list-item-with-action\"]/div[2]/div[1]/div/div[1]";
        private const string FIRST_MENU_SITE = "//*[@id=\"content_56\"]/div/div/table/tbody/tr[1]/td[3]";
        private const string FIRST_MENU_RECIPE = "/html/body/div[3]/div/div[2]/div[2]/div/div/div/div[2]/div[2]/div/div/table/tbody/tr[1]/td[2]";
        private const string FIRST_MENU_DELIVERY = "/html/body/div[3]/div/div[2]/div[2]/div/div/div/div[2]/div[2]/div/div/table/tbody/tr/td[4]";
        private const string FIRST_MENU_FOODPACK = "/html/body/div[3]/div/div[2]/div[2]/div/div/div/div[2]/div[2]/div/div/table/tbody/tr/td[5]";
        private const string COLLAPSE_ITEM = "//*/div[@class='row']/div/table/tbody/tr/td[1]";
        private const string ITEM_COLLAPSED = "//*/div[@class='panel-body']/div/table/tbody/tr/td[2]";
        private const string FIRST_CUSTOMER_DISPLAYED = "//*[@id=\"list-item-with-action\"]/div[2]/div[1]/div/div[2]/table/tbody/tr/td[1]";

        private const string FOLD_UNFOLD_ALL = "//*[@id=\"tabContentItemContainer\"]/div[1]/div/div/div/a[1]";
        private const string GENERATE_OUTPUT_FORM = "btn-generate-output-form";

        private const string COLUMN_MENUS_NAMES = "//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]/table/tbody/tr/td[1]";
        private const string COLUMN_MENUS_TOGGLEBUTTON = "//*[@id=\"list-item-with-action\"]/div[{0}]/div[1]/div/div[1]";
        private const string  COLUMN_MENUS_PAX = "/html/body/div[3]/div/div[2]/div[2]/div/div/div/div[{0}]/div/div/div[2]/table/tbody/tr/td[4]";

        private const string COLUMN_MENUS_RECIPE_NAME = "/html/body/div[3]/div/div[2]/div[2]/div/div/div/div[{0}]/div[2]/div/div/table/tbody/tr/td[2]";
        private const string COLUMN_MENUS_SERVICE_NAME = "/html/body/div[3]/div/div[2]/div[2]/div/div/div/div[{0}]/div[2]/div/div/table/tbody/tr/td[3]";
        private const string COLUMN_MENUS_DELIVERY = "/html/body/div[3]/div/div[2]/div[2]/div/div/div/div[{0}]/div[2]/div/div/table/tbody/tr/td[4]";
        private const string COLUMN_MENUS_RECIPE_FOODPACK = "/html/body/div[3]/div/div[2]/div[2]/div/div/div/div[{0}]/div[2]/div/div/table/tbody/tr/td[5]";
        private const string COLUMN_MENUS_RECIPE_PAX = "/html/body/div[3]/div/div[2]/div[2]/div/div/div/div[{0}]/div[2]/div/div/table/tbody/tr/td[6]";
        private const string COLUMN_MENUS_QTYPRODUCE = "/html/body/div[3]/div/div[2]/div[2]/div/div/div/div[{0}]/div[2]/div/div/table/tbody/tr/td[8]";

        private const string COLUMN_MENUS_ITEM_NAME = "/html/body/div[3]/div/div[2]/div[2]/div/div/div/div[{0}]/div[2]/div/div/table/tbody/tr[2]/td/div/div/div/table/tbody/tr/td[1]/span";
        private const string COLUMN_MENUS_ITEM_NETWEIGHT = "/html/body/div[3]/div/div[2]/div[2]/div/div/div/div[{0}]/div[2]/div/div/table/tbody/tr[2]/td/div/div/div/table/tbody/tr/td[2]/span[2]";
        private const string FIRST_LINE = "/html/body/div[3]/div/div[2]/div[2]/div/div/div/div[2]/div[2]/div";

        private const string PAGINATION1 = "//*[@id=\"page-size-selector\"]/option[1]";
        private const string PAGINATION2 = "//*[@id=\"page-size-selector\"]/option[2]";
        private const string PAGINATION3 = "//*[@id=\"page-size-selector\"]/option[3]";
        private const string PAGINATION4 = "//*[@id=\"page-size-selector\"]/option[4]";
        private const string PAGINATION5 = "//*[@id=\"page-size-selector\"]/option[5]";

        // ______________________________________ Variables _____________________________________________

        // ONGLETS

        [FindsBy(How = How.XPath, Using = CUSTOMER_TAB)]
        private IWebElement _customerTab;

        [FindsBy(How = How.XPath, Using = RECIPE_TAB)]
        private IWebElement _recipeTab;

        [FindsBy(How = How.XPath, Using = WORKSHOP_TAB)]
        private IWebElement _workshopTab;

        [FindsBy(How = How.XPath, Using = CUSTOMER_ORDER_TAB)]
        private IWebElement _customerOrderTab;


        // FILTRES

        [FindsBy(How = How.XPath, Using = RESET_FILTER)]
        private IWebElement _resetFilter;

        [FindsBy(How = How.Id, Using = SEARCH_FILTER)]
        private IWebElement _searchFilter;

        [FindsBy(How = How.Id, Using = DATE_FILTER)]
        private IWebElement _dateFilter;
        
        [FindsBy(How = How.XPath, Using = DATE_LABEL_TO_CLICK)]
        private IWebElement _dateLabelToClick;

        [FindsBy(How = How.Id, Using = SITE_FILTER)]
        private IWebElement _selectSite;


        [FindsBy(How = How.XPath, Using = SHOW_ITEMS_FILTER)]
        private IWebElement _showItems;
        
        [FindsBy(How = How.Id, Using = SHOW_PROCEDURE_FILTER)]
        private IWebElement _showProcedure;
        
        [FindsBy(How = How.Id, Using = SHOW_PAX_IS_ZERO_FILTER)]
        private IWebElement _showPaxIsZero;


        [FindsBy(How = How.Id, Using = CUSTOMER_TYPES_FILTER)]
        private IWebElement _customerTypesFilter;

        [FindsBy(How = How.XPath, Using = CUSTOMER_TYPES_FILTER_UNCHECK_ALL)]
        private IWebElement _customerTypesFilterUncheckAll;
        
        [FindsBy(How = How.XPath, Using = CUSTOMER_TYPES_FILTER_SEARCH)]
        private IWebElement _customerTypesFilterSearch;

        [FindsBy(How = How.Id, Using = CUSTOMERS_FILTER)]
        private IWebElement _customersFilter;

        [FindsBy(How = How.XPath, Using = CUSTOMERS_FILTER_UNCHECK_ALL)]
        private IWebElement _customersFilterUncheckAll;

        [FindsBy(How = How.XPath, Using = CUSTOMERS_FILTER_SEARCH)]
        private IWebElement _customersFilterSearch;

        [FindsBy(How = How.XPath, Using = WORKSHOPS_FILTER)]
        private IWebElement _workshopsFilter;

        [FindsBy(How = How.XPath, Using = WORKSHOPS_FILTER_UNCHECK_ALL)]
        private IWebElement _workshopsFilterUncheckAll;

        [FindsBy(How = How.XPath, Using = WORKSHOPS_FILTER_SEARCH)]
        private IWebElement _workshopsFilterSearch;

        [FindsBy(How = How.Id, Using = MEALTYPES_FILTER)]
        private IWebElement _mealtypesFilter;

        [FindsBy(How = How.XPath, Using = MEALTYPES_FILTER_UNCHECK_ALL)]
        private IWebElement _mealtypesFilterUncheckAll;

        [FindsBy(How = How.XPath, Using = MEALTYPES_FILTER_SEARCH)]
        private IWebElement _mealtypesFilterSearch;


        [FindsBy(How = How.XPath, Using = ITEM_GROUPS_FILTER)]
        private IWebElement _itemGroupsFilter;

        [FindsBy(How = How.XPath, Using = ITEM_GROUPS_FILTER_UNCHECK_ALL)]
        private IWebElement _itemGroupsFilterUncheckAll;

        [FindsBy(How = How.XPath, Using = ITEM_GROUPS_FILTER_SEARCH)]
        private IWebElement _itemGroupsFilterSearch;


        [FindsBy(How = How.Id, Using = GUESTTYPES_FILTER)]
        private IWebElement _guestTypesFilter;

        [FindsBy(How = How.XPath, Using = GUESTTYPES_FILTER_UNCHECK_ALL)]
        private IWebElement _guestTypesFilterUncheckAll;

        [FindsBy(How = How.XPath, Using = GUESTTYPES_FILTER_SEARCH)]
        private IWebElement _guestTypesFilterSearch;


        [FindsBy(How = How.Id, Using = DELIVERY_ROUNDS_FILTER)]
        private IWebElement _deliveryRoundsFilter;

        [FindsBy(How = How.XPath, Using = DELIVERY_ROUNDS_FILTER_UNCHECK_ALL)]
        private IWebElement _deliveryRoundsFilterUncheckAll;

        [FindsBy(How = How.XPath, Using = DELIVERY_ROUNDS_FILTER_SEARCH)]
        private IWebElement _deliveryRoundsFilterSearch;

        [FindsBy(How = How.Id, Using = FOODPACK_TYPES_FILTER)]
        private IWebElement _foodpackTypesFilter;

        [FindsBy(How = How.XPath, Using = FOODPACK_TYPES_FILTER_UNCHECK_ALL)]
        private IWebElement _foodpackTypesFilterUncheckAll;

        [FindsBy(How = How.XPath, Using = FOODPACK_TYPES_FILTER_SEARCH)]
        private IWebElement _foodpackTypesFilterSearch;

        [FindsBy(How = How.Id, Using = DELIVERIES_FILTER)]
        private IWebElement _deliveriesFilter;

        [FindsBy(How = How.XPath, Using = DELIVERIES_FILTER_UNCHECK_ALL)]
        private IWebElement _deliveriesFilterUncheckAll;

        [FindsBy(How = How.XPath, Using = DELIVERIES_FILTER_SEARCH)]
        private IWebElement _deliveriesFilterSearch;

        // PRINT

        [FindsBy(How = How.Id, Using = BTN_PRINT_ALLOTMENT)]
        private IWebElement _printAllotment;
        
        [FindsBy(How = How.Id, Using = BTN_PRINT_ALLOTMENT_NON_PACKAGED)]
        private IWebElement _printAllotmentNonPackaged;
        
        [FindsBy(How = How.Id, Using = BTN_PRINT_ALLOTMENT_GLOBAL_NON_PACKAGED)]
        private IWebElement _printAllotmentGlobalNonPackaged;
        
        [FindsBy(How = How.Id, Using = BTN_PRINT_HACCP)]
        private IWebElement _printHaccp;
        
        [FindsBy(How = How.Id, Using = BTN_PRINT_RAWMATERIALS)]
        private IWebElement _printRawMaterials;
        
        [FindsBy(How = How.Id, Using = BTN_PRINT_PRODUCTION)]
        private IWebElement _printProduction;
        
        [FindsBy(How = How.Id, Using = BTN_PRINT_PRODUCTION_VALIDATE)]
        private IWebElement _printProductionValidate;

        // TABLEAU BY MENU

        [FindsBy(How = How.XPath, Using = NUMBER_OF_PRODUCTION)]
        private IWebElement _numberOfProduction;

        [FindsBy(How = How.XPath, Using = FIRST_MENU_NAME)]
        private IWebElement _firstMenuName;
        
        [FindsBy(How = How.XPath, Using = FIRST_MENU_RECIPE)]
        private IWebElement _firstMenuRecipe;

        [FindsBy(How = How.XPath, Using = TOGGLE_FIRST_MENU_DISPLAYED)]
        private IWebElement _toggleFirstMenu;

        [FindsBy(How = How.XPath, Using = FIRST_MENU_SITE)]
        private IWebElement _firstMenuSite;
        
        [FindsBy(How = How.XPath, Using = FIRST_MENU_DELIVERY)]
        private IWebElement _firstMenuDelivery;
        
        [FindsBy(How = How.XPath, Using = FIRST_MENU_FOODPACK)]
        private IWebElement _firstMenuFoodpack;
        
        [FindsBy(How = How.XPath, Using = FIRST_MENU_VARIANT)]
        private IWebElement _firstMenuVariant;

        [FindsBy(How = How.XPath, Using = COLLAPSE_ITEM)]
        private IWebElement _collapseItem;
        
        [FindsBy(How = How.XPath, Using = ITEM_COLLAPSED)]
        private IWebElement _itemCollapsed;
        
        [FindsBy(How = How.XPath, Using = FOLD_UNFOLD_ALL)]
        private IWebElement _foldUnfoldAll;
        
        [FindsBy(How = How.Id, Using = GENERATE_OUTPUT_FORM)]
        private IWebElement _generateOutputForm;
        
        [FindsBy(How = How.XPath, Using = COLUMN_MENUS_TOGGLEBUTTON)]
        private IWebElement _menusColumnCurrentToggleButton;

        [FindsBy(How = How.XPath, Using = COLUMN_MENUS_PAX)]
        private IWebElement _menusColumnCurrentPAX;

        [FindsBy(How = How.XPath, Using = COLUMN_MENUS_RECIPE_NAME)]
        private IWebElement _menusColumnCurrentRecipeName;

        [FindsBy(How = How.XPath, Using = COLUMN_MENUS_SERVICE_NAME)]
        private IWebElement _menusColumnCurrentServiceName;
        
        [FindsBy(How = How.XPath, Using = COLUMN_MENUS_DELIVERY)]
        private IWebElement _menusColumnCurrentDelivery;

        [FindsBy(How = How.XPath, Using = COLUMN_MENUS_RECIPE_FOODPACK)]
        private IWebElement _menusColumnCurrentRecipeFoodPack;

        [FindsBy(How = How.XPath, Using = COLUMN_MENUS_RECIPE_PAX)]
        private IWebElement _menusColumnCurrentRecipePAX;

        [FindsBy(How = How.XPath, Using = COLUMN_MENUS_QTYPRODUCE)]
        private IWebElement _menusColumnCurrentQtyProduce;

        [FindsBy(How = How.XPath, Using = COLUMN_MENUS_ITEM_NAME)]
        private IWebElement _menusColumnCurrentItemName;

        [FindsBy(How = How.XPath, Using = COLUMN_MENUS_ITEM_NETWEIGHT)]
        private IWebElement _menusColumnCurrentItemNetWeight;

        [FindsBy(How = How.XPath, Using = PAGINATION1)]
        private IWebElement _pagination1;
        [FindsBy(How = How.XPath, Using = PAGINATION2)]
        private IWebElement _pagination2;
        [FindsBy(How = How.XPath, Using = PAGINATION3)]
        private IWebElement _pagination3;
        [FindsBy(How = How.XPath, Using = PAGINATION4)]
        private IWebElement _pagination4;
        [FindsBy(How = How.XPath, Using = PAGINATION5)]
        private IWebElement _pagination5;
        // ONGLETS
        public ProductionRecipeTabPage GoToProductionRecipeTab()
        {
            _recipeTab = WaitForElementIsVisible(By.XPath(RECIPE_TAB));
            _recipeTab.Click();
            WaitForLoad();
            return new ProductionRecipeTabPage(_webDriver, _testContext);
        }
        public ProductionCustomerTabPage GoToProductionCustomerTab()
        {
            _customerTab = WaitForElementIsVisible(By.XPath(CUSTOMER_TAB));
            _customerTab.Click();
            WaitForLoad();
            return new ProductionCustomerTabPage(_webDriver, _testContext);
        }

        public ProductionWorkshopTabPage GoToProductionWorkshopTab()
        {
            _workshopTab = WaitForElementIsVisible(By.XPath(WORKSHOP_TAB));
            _workshopTab.Click();
            WaitForLoad();
            return new ProductionWorkshopTabPage(_webDriver, _testContext);
        }
        public ProductionCustomerOrderTabPage GoToProductionCustomerOrderTab()
        {
            _customerOrderTab = WaitForElementIsVisible(By.XPath(CUSTOMER_ORDER_TAB));
            _customerOrderTab.Click();
            WaitForLoad();
            return new ProductionCustomerOrderTabPage(_webDriver, _testContext);
        }

        // PRINTS

        public enum PrintType
        {
            AllotmentReport,
            NonPackaged,
            GlobalNonPackaged,
            Production,
            RawMaterials,
            Haccp
        }

        public PrintReportPage Print(PrintType printType, bool printValue)
        {
            ShowExtendedMenu();

            switch (printType)
            {
                case PrintType.AllotmentReport:
                    _printAllotment = WaitForElementIsVisible(By.Id(BTN_PRINT_ALLOTMENT));
                    _printAllotment.Click();
                    WaitForLoad();
                    Thread.Sleep(2000);
                    var _printButton = WaitForElementIsVisible(By.Id("printButton"));
                    _printButton.Click();
                    WaitForLoad();

                    break;
                case PrintType.NonPackaged:
                    _printAllotmentNonPackaged = WaitForElementIsVisible(By.Id(BTN_PRINT_ALLOTMENT_NON_PACKAGED));
                    _printAllotmentNonPackaged.Click();
                    WaitForLoad();
                    break;
                case PrintType.GlobalNonPackaged:
                    _printAllotmentGlobalNonPackaged = WaitForElementIsVisible(By.Id(BTN_PRINT_ALLOTMENT_GLOBAL_NON_PACKAGED));
                    _printAllotmentGlobalNonPackaged.Click();
                    WaitForLoad();
                    break;
                case PrintType.Production:
                    _printProduction = WaitForElementIsVisible(By.Id(BTN_PRINT_PRODUCTION));
                    _printProduction.Click();
                    WaitForLoad();
                    _printProductionValidate = WaitForElementIsVisible(By.Id(BTN_PRINT_PRODUCTION_VALIDATE));
                    _printProductionValidate.Click();
                    WaitForLoad();
                    break;
                case PrintType.RawMaterials:
                    _printRawMaterials = WaitForElementIsVisible(By.Id(BTN_PRINT_RAWMATERIALS));
                    _printRawMaterials.Click();
                    WaitForLoad();
                    break;
                case PrintType.Haccp:
                    _printHaccp = WaitForElementIsVisible(By.Id(BTN_PRINT_HACCP));
                    _printHaccp.Click();
                    WaitForLoad();
                    break;
            }

            if (printValue)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-pdf']"));
                ClickPrintButton();
            }

            //Results are opened in a new tab, switch the driver to the newly created one
            WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new PrintReportPage(_webDriver, _testContext);
        }


        // FILTRES
        public enum FilterType
        { 
            Search,
            Date,
            DateTo,
            Site,
            ShowItems,
            ShowProcedure,
            ShowPaxIsZero,
            CustomerTypes,
            Customers,
            Workshops,
            MealTypes,
            ItemGroups,
            GuestTypes,
            DeliveryRounds,
            FoodpackTypes,
            Deliveries
        }

        public ProductionPage ResetFilter()
        {
            _resetFilter = WaitForElementIsVisible(By.XPath(RESET_FILTER));
            _resetFilter.Click();
            WaitForLoad();

            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                Filter(FilterType.Date, DateUtils.Now.ToString("dd/MM/yyyy"));
            }

            return new ProductionPage(_webDriver, _testContext);
        }

        public void Filter(FilterType filterType, object value)
        {
            Actions action = new Actions(_webDriver);

            switch (filterType)
            {
                case FilterType.Search:
                    _searchFilter = WaitForElementIsVisible(By.Id(SEARCH_FILTER));
                    _searchFilter.SetValue(ControlType.TextBox, value);
                    WaitForLoad();
                    break;
                case FilterType.Date:
                    _dateFilter = WaitForElementIsVisible(By.Id(DATE_FILTER));
                    action.MoveToElement(_dateFilter).Perform();
                    _dateFilter.SetValue(ControlType.TextBox, value);
                    WaitForLoad();
                    _dateLabelToClick = WaitForElementIsVisible(By.XPath("//*/td[@class='active day']"));
                    _dateLabelToClick.Click();
                    break;
                case FilterType.Site:
                    _selectSite = WaitForElementIsVisible(By.Id(SITE_FILTER));
                    _selectSite.SetValue(ControlType.DropDownList, value);
                    WaitForLoad();
                    break;
                case FilterType.ShowItems:
                    _showItems = WaitForElementIsVisible(By.XPath(SHOW_ITEMS_FILTER));
                    bool checkBoxInitialState;
                    var showItemscheckboxElement = WaitForElementIsVisible(By.XPath("//*[@id=\"item-filter-form\"]/div[6]/div/span"));
                    var style = showItemscheckboxElement.GetCssValue("background-color");
                    if (style == "rgba(255, 255, 255, 1)") //==false
                    {
                        checkBoxInitialState = false;
                        if (value.ToString() != checkBoxInitialState.ToString())
                        {
                            _showItems.Click();
                        }
                    }
                    else
                    {
                        checkBoxInitialState = true;
                        if (value.ToString() != checkBoxInitialState.ToString())
                        {
                            _showItems.Click();
                        }
                    }
                    WaitForLoad();
                    break;
                case FilterType.ShowProcedure:
                    _showProcedure = WaitForElementIsVisible(By.Id(SHOW_PROCEDURE_FILTER));
                    action.MoveToElement(_showProcedure).Perform();
                    _showProcedure.SetValue(ControlType.CheckBox, value);
                    WaitForLoad();
                    break;
                case FilterType.ShowPaxIsZero:
                    _showPaxIsZero = WaitForElementIsVisible(By.Id(SHOW_PAX_IS_ZERO_FILTER));
                    action.MoveToElement(_showPaxIsZero).Perform();
                    _showPaxIsZero.SetValue(ControlType.CheckBox, value);
                    WaitForLoad();
                    break;
                case FilterType.CustomerTypes:
                    _customerTypesFilter = WaitForElementIsVisible(By.Id(CUSTOMER_TYPES_FILTER));
                    _customerTypesFilter.Click();

                    _customerTypesFilterUncheckAll = WaitForElementIsVisible(By.XPath(CUSTOMER_TYPES_FILTER_UNCHECK_ALL));
                    _customerTypesFilterUncheckAll.Click();

                    _customerTypesFilterSearch = WaitForElementIsVisible(By.XPath(CUSTOMER_TYPES_FILTER_SEARCH));
                    _customerTypesFilterSearch.SetValue(ControlType.TextBox, value);

                    var _customerTypeToCheck = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    _customerTypeToCheck.SetValue(ControlType.CheckBox, true);
                    WaitForLoad();
                    break;
                case FilterType.Customers:
                    _customersFilter = WaitForElementIsVisible(By.Id(CUSTOMERS_FILTER));
                    _customersFilter.Click();

                    _customersFilterUncheckAll = WaitForElementIsVisible(By.XPath(CUSTOMERS_FILTER_UNCHECK_ALL));
                    _customersFilterUncheckAll.Click();

                    _customersFilterSearch = WaitForElementIsVisible(By.XPath(CUSTOMERS_FILTER_SEARCH));
                    _customersFilterSearch.SetValue(ControlType.TextBox, value);

                    var _customerToCheck = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    _customerToCheck.SetValue(ControlType.CheckBox, true);
                    WaitForLoad();
                    break;
                case FilterType.Workshops:
                    _workshopsFilter = WaitForElementIsVisible(By.XPath(WORKSHOPS_FILTER));
                    _workshopsFilter.Click();

                    _workshopsFilterUncheckAll = WaitForElementIsVisible(By.XPath(WORKSHOPS_FILTER_UNCHECK_ALL));
                    _workshopsFilterUncheckAll.Click();

                    _workshopsFilterSearch = WaitForElementIsVisible(By.XPath(WORKSHOPS_FILTER_SEARCH));
                    _workshopsFilterSearch.SetValue(ControlType.TextBox, value);

                    var _workshopToCheck = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    _workshopToCheck.SetValue(ControlType.CheckBox, true);
                    WaitForLoad();
                    break;
                case FilterType.MealTypes:
                    _mealtypesFilter = WaitForElementIsVisible(By.Id(MEALTYPES_FILTER));
                    _mealtypesFilter.Click();

                    _mealtypesFilterUncheckAll = WaitForElementIsVisible(By.XPath(MEALTYPES_FILTER_UNCHECK_ALL));
                    _mealtypesFilterUncheckAll.Click();

                    _mealtypesFilterSearch = WaitForElementIsVisible(By.XPath(MEALTYPES_FILTER_SEARCH));
                    _mealtypesFilterSearch.SetValue(ControlType.TextBox, value);

                    var _mealtypeToCheck = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    _mealtypeToCheck.SetValue(ControlType.CheckBox, true);
                    WaitForLoad();
                    break;
                case FilterType.ItemGroups:
                    _itemGroupsFilter = WaitForElementIsVisible(By.XPath(ITEM_GROUPS_FILTER));
                    _itemGroupsFilter.Click();

                    _itemGroupsFilterUncheckAll = WaitForElementIsVisible(By.XPath(ITEM_GROUPS_FILTER_UNCHECK_ALL));
                    _itemGroupsFilterUncheckAll.Click();

                    _itemGroupsFilterSearch = WaitForElementIsVisible(By.XPath(ITEM_GROUPS_FILTER_SEARCH));
                    _itemGroupsFilterSearch.SetValue(ControlType.TextBox, value);

                    var _itemGroupToCheck = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    _itemGroupToCheck.SetValue(ControlType.CheckBox, true);
                    WaitForLoad();
                    break;
                case FilterType.GuestTypes:
                    _guestTypesFilter = WaitForElementExists(By.Id(GUESTTYPES_FILTER));
                    action.MoveToElement(_guestTypesFilter).Perform();
                    ComboBoxSelectById(new ComboBoxOptions(GUESTTYPES_FILTER, (string)value));
                    break;
                case FilterType.DeliveryRounds:
                    _deliveryRoundsFilter = WaitForElementExists(By.Id(DELIVERY_ROUNDS_FILTER));
                    action.MoveToElement(_deliveryRoundsFilter).Perform();
                    ComboBoxSelectById(new ComboBoxOptions(DELIVERY_ROUNDS_FILTER, (string)value));
                    break;
                case FilterType.FoodpackTypes:
                    _foodpackTypesFilter = WaitForElementExists(By.Id(FOODPACK_TYPES_FILTER));
                    action.MoveToElement(_foodpackTypesFilter).Perform();
                    ComboBoxSelectById(new ComboBoxOptions(FOODPACK_TYPES_FILTER, (string)value));
                    break;
                case FilterType.Deliveries:
                    _deliveriesFilter = WaitForElementExists(By.Id(DELIVERIES_FILTER));
                    action.MoveToElement(_deliveriesFilter).Perform();
                    ComboBoxSelectById(new ComboBoxOptions(DELIVERIES_FILTER, (string)value));
                    break;
            }
            WaitPageLoading();
            WaitForLoad();
        }


        // TABLEAU BY MENU

        public bool IsResultsDisplayed()
        {
            Thread.Sleep(2000);//wait table
            _numberOfProduction = WaitForElementIsVisible(By.XPath(NUMBER_OF_PRODUCTION));
            if (_numberOfProduction.Text.Equals("0"))
            {
                return false;
            }
            else 
            { 
                return true;
            }
        }

        public bool ToggleFirstMenu()
        {
            try {
                _toggleFirstMenu = WaitForElementIsVisible(By.XPath(TOGGLE_FIRST_MENU_DISPLAYED));
                _toggleFirstMenu.Click();
                return true;
            } catch
            {
                return false;
            }
        }

        // Retourne vrai si les éléments du tableau sont dépliés
        // Retourne faux si les éléments du tableau sont repliés
        public bool FoldUnfoldAll()
        {
            ShowExtendedMenu();
            _foldUnfoldAll = WaitForElementIsVisible(By.XPath(FOLD_UNFOLD_ALL));
            _foldUnfoldAll.Click();
            //Temps d'affichage du résultat
            WaitPageLoading();

            return isElementVisible(By.XPath(FIRST_LINE));
        }

        public ProductionGenerateOutputFormModal GenerateOutputForm()
        {
            ShowExtendedMenu();
            _generateOutputForm = WaitForElementIsVisible(By.Id(GENERATE_OUTPUT_FORM));
            _generateOutputForm.Click();
            return new ProductionGenerateOutputFormModal(_webDriver, _testContext);
        }

        // GETTERS

        public string GetNumberOfDisplayedResults()
        {
            _numberOfProduction = WaitForElementIsVisible(By.XPath(NUMBER_OF_PRODUCTION));
            return _numberOfProduction.Text;
        }

        public string GetFirstMenuName()
        {
            _firstMenuName = WaitForElementIsVisible(By.XPath(FIRST_MENU_NAME));
            return _firstMenuName.Text;
        }

        public string GetFirstMenuVariant()
        {
            _firstMenuVariant = WaitForElementIsVisible(By.XPath(FIRST_MENU_VARIANT));
            return _firstMenuVariant.Text;
             
        }
        public string GetMealTypes(string mealType)
        {
            return mealType.Substring(mealType.IndexOf("-") + 1).Trim();
        }

        public string GetDeliveryOfFirstMenu()
        {
            if (this.ToggleFirstMenu())
            {
                _firstMenuDelivery = WaitForElementIsVisible(By.XPath(FIRST_MENU_DELIVERY));
                return _firstMenuDelivery.Text;
            } else
            {
                return "";
            }
        }

        public string GetFirstMenuItem()
        {
            if (isElementVisible(By.XPath( FIRST_MENU_RECIPE)))
            {
                _collapseItem = WaitForElementIsVisible(By.XPath(COLLAPSE_ITEM));
                _collapseItem.Click();
                _itemCollapsed = WaitForElementIsVisible(By.XPath(ITEM_COLLAPSED));
                return _itemCollapsed.Text;
            } else
            {
                return "";
            }
        }

        public string GetFoodpackOfFirstMenu()
        {
            if (this.ToggleFirstMenu())
            {
                _firstMenuFoodpack = WaitForElementIsVisible(By.XPath(FIRST_MENU_FOODPACK));
                return _firstMenuFoodpack.Text;
            }
            else
            {
                return "";
            }
        }

        public bool IsSiteOfFirstProductionDisplayedCorrect(string site)
        {
            _toggleFirstMenu = WaitForElementIsVisible(By.XPath(TOGGLE_FIRST_MENU_DISPLAYED));
            _toggleFirstMenu.Click();
            _firstMenuSite = WaitForElementIsVisible(By.XPath(FIRST_MENU_SITE));

            if (_firstMenuSite.Text.Contains(site))
            {
                return true;
            } else
            {
                return false;
            }
        }

        public bool IsFilterShowApplied()
        {
            _toggleFirstMenu = WaitForElementIsVisible(By.XPath(TOGGLE_FIRST_MENU_DISPLAYED));
            _toggleFirstMenu.Click();

            _collapseItem = WaitForElementIsVisible(By.XPath(COLLAPSE_ITEM));
            
            try
            {
                _collapseItem.Click();
                return true;
            } 
            catch
            {
                return false;
            }
        }

        public Dictionary<string, MappedDetailsByMenu> GetMenusNamesAndQty()
        {
            Thread.Sleep(2000);
            int tot = CheckTotalNumber();
            Assert.IsFalse(tot == 0, "Le résultat de Production n'est pas affiché, il peut y avoir un problème dans le paramétrage.");
                
            int i = 0;

            Dictionary<string, MappedDetailsByMenu> menusDetails = new Dictionary<string, MappedDetailsByMenu>();

            ShowExtendedMenu();
            var foldOrUnfold = _webDriver.FindElement(By.XPath("//*[@id=\"tabContentItemContainer\"]/div[1]/div/div/div/a[1]"));
            if (foldOrUnfold.Text != "Fold All")
            {
                FoldUnfoldAll();
                WaitForLoad();
            }

            WaitForLoad();

            var columnMenusNames = _webDriver.FindElements(By.XPath(COLUMN_MENUS_NAMES));

            foreach (var menu in columnMenusNames)
            {
                // On limite le nombre de menus remontés à 5 pour ne pas surcharger le test
                if (i >= 4)
                    break;

                var mappedDeliveryAndQty = new MappedDetailsByMenu();

                //Menu PAX
                _menusColumnCurrentPAX = _webDriver.FindElement(By.XPath(String.Format(COLUMN_MENUS_PAX, i + 2)));
                mappedDeliveryAndQty.menuPAX = _menusColumnCurrentPAX.Text;

                //Menu Recipe Name
                _menusColumnCurrentRecipeName = _webDriver.FindElement(By.XPath(String.Format(COLUMN_MENUS_RECIPE_NAME, i + 2)));
                mappedDeliveryAndQty.recipeName = _menusColumnCurrentRecipeName.Text;

                //Menu Service Name
                _menusColumnCurrentServiceName = _webDriver.FindElement(By.XPath(String.Format(COLUMN_MENUS_SERVICE_NAME, i + 2)));
                mappedDeliveryAndQty.serviceName = _menusColumnCurrentServiceName.Text;

                //Menu Delivery Name
                _menusColumnCurrentDelivery = _webDriver.FindElement(By.XPath(String.Format(COLUMN_MENUS_DELIVERY, i + 2)));
                mappedDeliveryAndQty.deliveryName = _menusColumnCurrentDelivery.Text;

                //Menu Recipe Foodpack

                _menusColumnCurrentRecipeFoodPack = _webDriver.FindElement(By.XPath(String.Format(COLUMN_MENUS_RECIPE_FOODPACK, i + 2)));
                mappedDeliveryAndQty.foodPack = _menusColumnCurrentRecipeFoodPack.Text;

                //Menu Recipe Foodpack
                _menusColumnCurrentRecipePAX = _webDriver.FindElement(By.XPath(String.Format(COLUMN_MENUS_RECIPE_PAX, i + 2)));
                mappedDeliveryAndQty.recipePAX = _menusColumnCurrentRecipePAX.Text;

                //Menu Recipe QUantity to produce
                _menusColumnCurrentQtyProduce = _webDriver.FindElement(By.XPath(String.Format(COLUMN_MENUS_QTYPRODUCE, i + 2)));
                var qtyString = _menusColumnCurrentQtyProduce.Text;
                if (qtyString.Length >= 6)
                {
                    var startQty = qtyString.IndexOf(":") + 2;
                    mappedDeliveryAndQty.qtyToProduce = qtyString.Substring(startQty);
                }
                else
                {
                    mappedDeliveryAndQty.qtyToProduce = qtyString;
                }

                //Menu Recipe Item Name
                try
                {
                    _menusColumnCurrentItemName = _webDriver.FindElement(By.XPath(String.Format(COLUMN_MENUS_ITEM_NAME, i + 2)));
                    mappedDeliveryAndQty.itemName = _menusColumnCurrentItemName.Text;

                    //Menu Recipe Item Net Weight
                    _menusColumnCurrentItemNetWeight = _webDriver.FindElement(By.XPath(String.Format(COLUMN_MENUS_ITEM_NETWEIGHT, i + 2)));
                    mappedDeliveryAndQty.itemNetWeight = _menusColumnCurrentItemNetWeight.Text;
                }
                catch
                {
                    //continue
                }

                menusDetails.Add(menu.Text, mappedDeliveryAndQty);
                i++;
            }

            return menusDetails;
        }

            // INPUT : Dictionary<recipeName, qtyToProduce>
            // OUTPUT : Dictionary<delivery, qtyToProduce>
            public Dictionary<string, string> GetMappedDeliveriesAndQtiesFromRecipes(Dictionary<string, string> recipesNamesAndQties)
        {
            int tot = CheckTotalNumber();
            Assert.IsFalse(tot == 0, "Le résultat de Production n'est pas affiché, il peut y avoir un problème dans le paramétrage.");

            int i = 0;

            Dictionary<string, string> deliveriesAndQties = new Dictionary<string, string>();

            var columnMenusNames = _webDriver.FindElements(By.XPath(COLUMN_MENUS_NAMES));

            foreach (var menu in columnMenusNames)
            {
                // On limite le nombre de recettes remontées à 5 pour ne pas surcharger le test
                if (i >= 4)
                    break;

                _menusColumnCurrentToggleButton = _webDriver.FindElement(By.XPath(String.Format(COLUMN_MENUS_TOGGLEBUTTON, i + 2)));
                _menusColumnCurrentToggleButton.Click();

                WaitForLoad();

                _menusColumnCurrentDelivery = _webDriver.FindElement(By.XPath(String.Format(COLUMN_MENUS_DELIVERY, i + 2)));

                _menusColumnCurrentRecipeName = _webDriver.FindElement(By.XPath(String.Format(COLUMN_MENUS_RECIPE_NAME, i + 2)));

                foreach(var recipeNameAndQty in recipesNamesAndQties)
                {
                    if (recipeNameAndQty.Key.Equals(_menusColumnCurrentRecipeName.Text))
                    {
                        deliveriesAndQties.Add(_menusColumnCurrentDelivery.Text, recipeNameAndQty.Value);
                    }
                }

                i++;
            }

            return deliveriesAndQties;
        }

        public Dictionary<string, MappedDetailsByMenu> GetCustomersDeliveriesAndQties()
        {
            int tot = CheckTotalNumber();
            Assert.IsFalse(tot == 0, "Le résultat de Production n'est pas affiché, il peut y avoir un problème dans le paramétrage.");

            int i = 0;

            Dictionary<string, MappedDetailsByMenu> menusNamesAndQty = new Dictionary<string, MappedDetailsByMenu>();

            var columnMenusNames = _webDriver.FindElements(By.XPath(COLUMN_MENUS_NAMES));

            foreach (var menu in columnMenusNames)
            {
                // On limite le nombre de menus remontés à 5 pour ne pas surcharger le test
                if (i >= 15)
                    break;

                _menusColumnCurrentToggleButton = _webDriver.FindElement(By.XPath(String.Format(COLUMN_MENUS_TOGGLEBUTTON, i + 2)));
                _menusColumnCurrentToggleButton.Click();

                var mappedDeliveryAndQty = new MappedDetailsByMenu();

                _menusColumnCurrentQtyProduce = _webDriver.FindElement(By.XPath(String.Format(COLUMN_MENUS_QTYPRODUCE, i + 2)));

                var qtyString = _menusColumnCurrentQtyProduce.Text;

                if (qtyString.Length >= 6)
                {
                    var startQty = qtyString.IndexOf(":") + 2;
                    mappedDeliveryAndQty.qtyToProduce = qtyString.Substring(startQty);

                }
                else
                {
                    mappedDeliveryAndQty.qtyToProduce = qtyString;
                }

                _menusColumnCurrentDelivery = _webDriver.FindElement(By.XPath(String.Format(COLUMN_MENUS_DELIVERY, i + 2)));

                mappedDeliveryAndQty.deliveryName = _menusColumnCurrentDelivery.Text;

                menusNamesAndQty.Add(menu.Text, mappedDeliveryAndQty);
                i++;
            }

            return menusNamesAndQty;
        }

        public string GetFirstMenuRecipe()
        {
            if (this.ToggleFirstMenu())
            {
                _firstMenuRecipe = WaitForElementIsVisible(By.XPath(FIRST_MENU_RECIPE));
                return _firstMenuRecipe.Text;
            } else
            {
                return "";
            }
        }

        // INPUT : menu name & recipe name - OUTPUT : recipe pax
        public string GetMenuRecipePAX(string menu, string recipe)
        {
            ShowExtendedMenu();
            var foldOrUnfold = _webDriver.FindElement(By.XPath("//*[@id=\"tabContentItemContainer\"]/div[1]/div/div/div/a[1]"));
            if (foldOrUnfold.Text != "Fold All")
            {
                FoldUnfoldAll();
                WaitForLoad();
            }

            var pax = _webDriver.FindElement(By.XPath(String.Format("//td[1][contains(text(),'{0}')]//ancestor::div[@class='panel panel-default']//td[2][contains(text(),'{1}')]/..//span[contains(text(),'PAX')]/parent::td", menu, recipe)));
            return pax.Text;
        }

        // INPUT : menu name & recipe name - OUTPUT : recipe foodpack
        public string GetMenuRecipeFoodpack(string menu, string recipe)
        {
            ShowExtendedMenu();
            var foldOrUnfold = _webDriver.FindElement(By.XPath("//*[@id=\"tabContentItemContainer\"]/div[1]/div/div/div/a[1]"));
            if (foldOrUnfold.Text != "Fold All")
            {
                FoldUnfoldAll();
                WaitForLoad();
            }

            var foopack = _webDriver.FindElement(By.XPath(String.Format("//td[1][contains(text(),'{0}')]//ancestor::div[@class='panel panel-default']//td[2][contains(text(),'{1}')]/..//span[contains(text(),'Food Pack')]/parent::td", menu, recipe)));
            return foopack.Text;
        }

        // INPUT : menu name & recipe name - OUTPUT : recipe quantity to produce
        public string GetMenuRecipeQuantityToProduce(string menu, string recipe)
        {
            ShowExtendedMenu();
            var foldOrUnfold = _webDriver.FindElement(By.XPath("//*[@id=\"tabContentItemContainer\"]/div[1]/div/div/div/a[1]"));
            if (foldOrUnfold.Text != "Fold All")
            {
                FoldUnfoldAll();
                WaitForLoad();
            }

            var foopack = _webDriver.FindElement(By.XPath(String.Format("//td[1][contains(text(),'{0}')]//ancestor::div[@class='panel panel-default']//td[2][contains(text(),'{1}')]/..//span[contains(text(),'Qty to produce')]/parent::td", menu, recipe)));
            return foopack.Text;
        }

        // INPUT : menu name & recipe name - OUTPUT : recipe net weight
        public string GetMenuRecipeNetWeight(string menu, string recipe)
        {
            ShowExtendedMenu();
            var foldOrUnfold = _webDriver.FindElement(By.XPath("//*[@id=\"tabContentItemContainer\"]/div[1]/div/div/div/a[1]"));
            if (foldOrUnfold.Text != "Fold All")
            {
                FoldUnfoldAll();
                WaitForLoad();
            }

            var recipeNetWeight = _webDriver.FindElement(By.XPath(String.Format("//td[1][contains(text(),'{0}')]//ancestor::div[@class='panel panel-default']//td[2][contains(text(),'{1}')]/../following-sibling::tr[1]//span[contains(text(),'Net Weight')]/parent::td/span[2]", menu, recipe)));
            return recipeNetWeight.Text;
        }

        public string Round2decimals(string weight)
        {
            // entrant :
            // 1 KG => 1 KG
            // 0,999 KG => 1 KG
            // 0.555 KG => 0,56 KG
            string number = weight.Replace(" KG","").Replace(",",".");
            Console.WriteLine(weight + "->" + number);
            Decimal numberDouble = Decimal.Parse(number.Replace(".", ","));
            Decimal numberDecimal2 = Decimal.Round(numberDouble,2);
            string result = numberDecimal2.ToString().Replace(".",",") + " KG";
            Console.WriteLine(weight + " => " + result);
            return result;
        }

        public bool Pagination()
        {
            var _paginate1 = WaitForElementExists(By.XPath(PAGINATION2));
            if (_paginate1 == null) return false;
            _paginate1.Click();
            WaitForLoad();

            var _paginate2 = WaitForElementExists(By.XPath(PAGINATION2));
            if (_paginate2 == null) return false;
            _paginate2.Click();
            WaitForLoad();

            var _paginate3 = WaitForElementExists(By.XPath(PAGINATION3));
            if (_paginate3 == null) return false;
            _paginate3.Click();
            WaitForLoad();

            var _paginate4 = WaitForElementExists(By.XPath(PAGINATION4));
            if (_paginate4 == null) return false;
            _paginate4.Click();
            WaitForLoad();


            var _paginate5 = WaitForElementExists(By.XPath(PAGINATION5));
            if (_paginate5 == null) return false;
            _paginate5.Click();
            WaitForLoad();

            return true;
        }

        public string GetFilterValue(FilterType filterType)
        {
            switch (filterType)
            {
                case FilterType.Date:
                    var fromDateInput = _webDriver.FindElement(By.Id("StartDate"));
                    return fromDateInput.GetAttribute("value");

                case FilterType.DateTo:
                    var toDateInput = _webDriver.FindElement(By.Id("EndDate")); 
                    return toDateInput.GetAttribute("value");

                default:
                    throw new ArgumentException("Invalid filter type");
            }
        }

        public void ScrollToMealTypeFilter()
        {
            var filter = _webDriver.FindElement(By.Id(MEALTYPES_FILTER));
            ScrollToElement(filter);
        }

    }
}