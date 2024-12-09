using DocumentFormat.OpenXml.Office2010.ExcelAc;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.ProductionCO;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.ProductionManagement;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Threading;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Production.Production
{
    public class ProductionPage : PageBase
    {
        public ProductionPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // ______________________________________ Constantes _____________________________________________

        private const string BTN_VALIDATE = "//*[@id=\"btnValidate\"]";
        private const string INPUT_DATE = "StartDate";
        private const string SELECT_SITE = "SiteId";

        //GETTERS

        private const string SELECTED_CUSTOMERS = "//*[@id=\"SelectedCustomerTypesIds_ms\"]/span[2]";


        // FILTRES

        private const string SEARCH_FILTER = "//*[@id=\"SearchPattern\"]";
        private const string DATE_FILTER = "StartDate";
        private const string DATETO_FILTER = "EndDate";

        private const string DATE_LABEL_TO_CLICK = "//*[@id=\"production-filter-form\"]/div/div[1]/div[1]/label";
        //html/body/div[19]/div[1]/table/tbody/tr[5]/td[4]
        private const string SITE_FILTER = "SiteId";

        private const string SHOW_ITEMS_FILTER = "WithItems";
        private const string SHOW_PROCEDURE_FILTER = "WithProcedure";
        private const string SHOW_PAX_IS_ZERO_FILTER = "ShowPAXIsZero";

        private const string CUSTOMER_TYPES_FILTER = "SelectedCustomerTypesIds_ms";
        private const string CUSTOMER_TYPES_FILTER_UNCHECK_ALL = "/html/body/div[12]/div/ul/li[2]/a/span[2]";
        private const string CUSTOMER_TYPES_FILTER_SEARCH = "/html/body/div[12]/div/div/label/input";
        private const string CUSTOMER_TYPES_FILTER_SEARCH_FIRST_ELEMENT = "/html/body/div[12]/ul/li[1]/label/span";

        private const string CUSTOMERS_FILTER = "SelectedCustomersIds_ms";
        private const string CUSTOMERS_FILTER_UNCHECK_ALL = "/html/body/div[13]/div/ul/li[2]/a/span[2]";
        private const string CUSTOMERS_FILTER_SEARCH = "/html/body/div[13]/div/div/label/input";
        private const string CUSTOMERS_FILTER_SEARCH_FIRST_ELEMENT = "/html/body/div[13]/ul/li[1]/label/span";

        private const string WORKSHOPS_FILTER = "SelectedWorkshopsIds_ms";
        private const string WORKSHOPS_FILTER_UNCHECK_ALL = "/html/body/div[11]/div/ul/li[2]/a/span[2]";
        private const string WORKSHOPS_FILTER_SEARCH = "/html/body/div[11]/div/div/label/input";
        private const string WORKSHOPS_FILTER_SEARCH_FIRST_ELEMENT = "/html/body/div[11]/ul/li[1]/label/span";

        private const string MEALTYPES_FILTER = "SelectedMealTypesIds_ms";
        private const string MEALTYPES_FILTER_UNCHECK_ALL = "/html/body/div[10]/div/ul/li[2]/a/span[2]";
        private const string MEALTYPES_FILTER_SEARCH = "/html/body/div[10]/div/div/label/input";
        private const string MEALTYPES_FILTER_SEARCH_FIRST_ELEMENT = "/html/body/div[10]/ul/li[1]/label/span";

        private const string ITEM_GROUPS_FILTER = "SelectedItemGroupsIds_ms";
        private const string ITEM_GROUPS_FILTER_UNCHECK_ALL = "/html/body/div[14]/div/ul/li[2]/a/span[2]";
        private const string ITEM_GROUPS_FILTER_SEARCH = "/html/body/div[14]/div/ul/li[2]/a/span[2]";
        private const string ITEM_GROUPS_FILTER_SEARCH_FIRST_ELEMENT = "/html/body/div[14]/ul/li[1]/label/span";

        private const string GUESTTYPES_FILTER = "SelectedGuestTypesIds_ms";
        private const string GUESTTYPES_FILTER_UNCHECK_ALL = "/html/body/div[16]/div/ul/li[2]/a/span[2]";
        private const string GUESTTYPES_FILTER_SEARCH = "/html/body/div[16]/div/div/label/input";
        private const string GUESTTYPES_FILTER_SEARCH_FIRST_ELEMENT = "/html/body/div[16]/ul/li[1]/label/span";

        private const string DELIVERY_ROUNDS_FILTER = "SelectedDeliveryRoundsIds_ms";
        private const string DELIVERY_ROUNDS_FILTER_UNCHECK_ALL = "/html/body/div[17]/div/ul/li[2]/a/span[2]";
        private const string DELIVERY_ROUNDS_FILTER_SEARCH = "/html/body/div[17]/div/div/label/input";
        private const string DELIVERY_ROUNDS_FILTER_SEARCH_FIRST_ELEMENT = "/html/body/div[17]/ul/li[1]/label/span";

        private const string FOODPACK_TYPES_FILTER = "SelectedFoodPackTypesIds_ms";
        private const string FOODPACK_TYPES_FILTER_UNCHECK_ALL = "/html/body/div[15]/div/ul/li[2]/a/span[2]";
        private const string FOODPACK_TYPES_FILTER_SEARCH = "/html/body/div[15]/div/div/label/input";
        private const string FOODPACK_TYPES_FILTER_SEARCH_FIRST_ELEMENT = "/html/body/div[15]/ul/li[1]/label/span";

        private const string DELIVERIES_FILTER = "SelectedDeliveriesIds_ms";
        private const string DELIVERIES_FILTER_UNCHECK_ALL = "/html/body/div[18]/div/ul/li[2]/a/span[2]";
        private const string DELIVERIES_FILTER_SEARCH = "/html/body/div[18]/div/div/label/input";
        private const string DELIVERIES_FILTER_SEARCH_FIRST_ELEMENT = "/html/body/div[18]/ul/li[1]/label/span";
        private const string BTN_PRINT_PRODUCTION = "printProd";
        private const string BYMENUBUTTON = "Menu";
        private const string BYRECIPEBUTTON = "Recipe";
        private const string BYCUSTOMERBUTTON = "Customer";
        private const string BYWORKSHOPBUTTON = "Workshop";
        private const string BYCUSTOMERORDERBUTTON = "Order";
        private const string BTN_PRINT_PRODUCTION_VALIDATE = "validatePrint";
        private const string PRODUCTION_PAGE = "//*[@id=\"tabContentItemContainer\"]/div[1]/h1";

        private const string MENUCHECKBOX = "//*[@id=\"Menu\"]";
        private const string RECIPECHECKBOX = "//*[@id=\"Recipe\"]";
        private const string CUSTOMERCHECKBOX = "//*[@id=\"Customer\"]";
        private const string ITEMGROUPECHECKBOX = "//*[@id=\"ItemGroup\"]";
        private const string WORKSHOPCHECKBOX = "//*[@id=\"Workshop\"]";
        private const string CUSTOMERORDERCHECKBOX = "//*[@id=\"Order\"]";

        private const string PRODUCTIONBUTTON = "//*[@id=\"printProd\"]";

        private const string COMMENTSECTION = "//*[@id=\"WorkshopComment\"]";

        private const string VALIDATEBUTTON = "//*[@id=\"validatePrint\"]";
        // ______________________________________ Variables _____________________________________________

        [FindsBy(How = How.XPath, Using = BTN_VALIDATE)]
        private IWebElement _btnValidate;

        [FindsBy(How = How.Id, Using = INPUT_DATE)]
        private IWebElement _inputDate;

        [FindsBy(How = How.Id, Using = SELECT_SITE)]
        private IWebElement _selectSite;

        //GETTERS

        [FindsBy(How = How.XPath, Using = SELECTED_CUSTOMERS)]
        private IWebElement _selectedCustomers;

        // FILTRES

        [FindsBy(How = How.XPath, Using = SEARCH_FILTER)]
        private IWebElement _searchFilter;

        [FindsBy(How = How.Id, Using = DATE_FILTER)]
        private IWebElement _dateFilter;

        [FindsBy(How = How.Id, Using = DATETO_FILTER)]
        private IWebElement _datetoFilter;

        [FindsBy(How = How.XPath, Using = DATE_LABEL_TO_CLICK)]
        private IWebElement _dateLabelToClick;

        [FindsBy(How = How.Id, Using = SITE_FILTER)]
        private IWebElement _siteFilter;


        [FindsBy(How = How.Id, Using = SHOW_ITEMS_FILTER)]
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

        [FindsBy(How = How.XPath, Using = CUSTOMER_TYPES_FILTER_SEARCH_FIRST_ELEMENT)]
        private IWebElement _customerTypesFilterSearchFirstElement;


        [FindsBy(How = How.Id, Using = CUSTOMERS_FILTER)]
        private IWebElement _customersFilter;

        [FindsBy(How = How.XPath, Using = CUSTOMERS_FILTER_UNCHECK_ALL)]
        private IWebElement _customersFilterUncheckAll;

        [FindsBy(How = How.XPath, Using = CUSTOMERS_FILTER_SEARCH)]
        private IWebElement _customersFilterSearch;

        [FindsBy(How = How.XPath, Using = CUSTOMERS_FILTER_SEARCH_FIRST_ELEMENT)]
        private IWebElement _customersFilterSearchFirstElement;


        [FindsBy(How = How.Id, Using = WORKSHOPS_FILTER)]
        private IWebElement _workshopsFilter;

        [FindsBy(How = How.XPath, Using = WORKSHOPS_FILTER_UNCHECK_ALL)]
        private IWebElement _workshopsFilterUncheckAll;

        [FindsBy(How = How.XPath, Using = WORKSHOPS_FILTER_SEARCH)]
        private IWebElement _workshopsFilterSearch;

        [FindsBy(How = How.XPath, Using = WORKSHOPS_FILTER_SEARCH_FIRST_ELEMENT)]
        private IWebElement _workshopsFilterSearchResultFirstElement;


        [FindsBy(How = How.Id, Using = MEALTYPES_FILTER)]
        private IWebElement _mealtypesFilter;

        [FindsBy(How = How.XPath, Using = MEALTYPES_FILTER_UNCHECK_ALL)]
        private IWebElement _mealtypesFilterUncheckAll;

        [FindsBy(How = How.XPath, Using = MEALTYPES_FILTER_SEARCH)]
        private IWebElement _mealtypesFilterSearch;

        [FindsBy(How = How.XPath, Using = MEALTYPES_FILTER_SEARCH_FIRST_ELEMENT)]
        private IWebElement _mealtypesFilterSearchFirstElement;


        [FindsBy(How = How.Id, Using = ITEM_GROUPS_FILTER)]
        private IWebElement _itemGroupsFilter;

        [FindsBy(How = How.XPath, Using = ITEM_GROUPS_FILTER_UNCHECK_ALL)]
        private IWebElement _itemGroupsFilterUncheckAll;

        [FindsBy(How = How.XPath, Using = ITEM_GROUPS_FILTER_SEARCH)]
        private IWebElement _itemGroupsFilterSearch;

        [FindsBy(How = How.XPath, Using = ITEM_GROUPS_FILTER_SEARCH_FIRST_ELEMENT)]
        private IWebElement _itemGroupsFilterSearchFirstElement;


        [FindsBy(How = How.Id, Using = GUESTTYPES_FILTER)]
        private IWebElement _guestTypesFilter;

        [FindsBy(How = How.XPath, Using = GUESTTYPES_FILTER_UNCHECK_ALL)]
        private IWebElement _guestTypesFilterUncheckAll;

        [FindsBy(How = How.XPath, Using = GUESTTYPES_FILTER_SEARCH)]
        private IWebElement _guestTypesFilterSearch;

        [FindsBy(How = How.XPath, Using = GUESTTYPES_FILTER_SEARCH_FIRST_ELEMENT)]
        private IWebElement _guestTypesFilterSearchFirstElement;


        [FindsBy(How = How.Id, Using = DELIVERY_ROUNDS_FILTER)]
        private IWebElement _deliveryRoundsFilter;

        [FindsBy(How = How.XPath, Using = DELIVERY_ROUNDS_FILTER_UNCHECK_ALL)]
        private IWebElement _deliveryRoundsFilterUncheckAll;

        [FindsBy(How = How.XPath, Using = DELIVERY_ROUNDS_FILTER_SEARCH)]
        private IWebElement _deliveryRoundsFilterSearch;

        [FindsBy(How = How.XPath, Using = DELIVERY_ROUNDS_FILTER_SEARCH_FIRST_ELEMENT)]
        private IWebElement _deliveryRoundsFilterSearchFirstElement;


        [FindsBy(How = How.Id, Using = FOODPACK_TYPES_FILTER)]
        private IWebElement _foodpackTypesFilter;

        [FindsBy(How = How.XPath, Using = FOODPACK_TYPES_FILTER_UNCHECK_ALL)]
        private IWebElement _foodpackTypesFilterUncheckAll;

        [FindsBy(How = How.XPath, Using = FOODPACK_TYPES_FILTER_SEARCH)]
        private IWebElement _foodpackTypesFilterSearch;

        [FindsBy(How = How.XPath, Using = FOODPACK_TYPES_FILTER_SEARCH_FIRST_ELEMENT)]
        private IWebElement _foodpackTypesFilterSearchFirstElement;


        [FindsBy(How = How.Id, Using = DELIVERIES_FILTER)]
        private IWebElement _deliveriesFilter;

        [FindsBy(How = How.XPath, Using = DELIVERIES_FILTER_UNCHECK_ALL)]
        private IWebElement _deliveriesFilterUncheckAll;

        [FindsBy(How = How.XPath, Using = DELIVERIES_FILTER_SEARCH)]
        private IWebElement _deliveriesFilterSearch;

        [FindsBy(How = How.XPath, Using = DELIVERIES_FILTER_SEARCH_FIRST_ELEMENT)]
        private IWebElement _deliveriesFilterSearchFirstElement;
        [FindsBy(How = How.Id, Using = BTN_PRINT_PRODUCTION)]
        private IWebElement _printProduction;

        [FindsBy(How = How.Id, Using = BTN_PRINT_PRODUCTION_VALIDATE)]
        private IWebElement _printProductionValidate;
        [FindsBy(How = How.XPath, Using = PRODUCTION_PAGE)]
        private IWebElement _productionPage;

        [FindsBy(How = How.XPath, Using = MENUCHECKBOX)]
        private IWebElement _menuCheckBox;

        [FindsBy(How = How.XPath, Using = RECIPECHECKBOX)]
        private IWebElement _recipeCheckBox;

        [FindsBy(How = How.XPath, Using = CUSTOMERCHECKBOX)]
        private IWebElement _customerCheckBox;

        [FindsBy(How = How.XPath, Using = ITEMGROUPECHECKBOX)]
        private IWebElement _itemGroupeCheckBox;

        [FindsBy(How = How.XPath, Using = WORKSHOPCHECKBOX)]
        private IWebElement _workShopCheckBox;


        [FindsBy(How = How.XPath, Using = CUSTOMERORDERCHECKBOX)]
        private IWebElement _customerOrderCheckBox;

        [FindsBy(How = How.XPath, Using = COMMENTSECTION)]
        private IWebElement _commentSection;





        // FILTRES
        public enum FilterType
        {
            SearchText,
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

        public enum PrintType
        {
            Menu,
            Recipe,
            Customer,
            ItemGroupe,
            Workshop,
            CustomerOrder,

        }

        public void Filter(FilterType filterType, object value)
        {
            switch (filterType)
            {
                case FilterType.SearchText:
                    _searchFilter = WaitForElementIsVisible(By.XPath(SEARCH_FILTER));
                    _searchFilter.SetValue(ControlType.TextBox, value);
                    WaitForLoad();
                    break;
                case FilterType.Date:
                    _dateFilter = WaitForElementIsVisible(By.Id(DATE_FILTER));
                    _dateFilter.Clear();
                    _dateFilter.SetValue(ControlType.TextBox, value);
                    WaitForLoad();
                    _dateFilter.SendKeys(Keys.Enter);
                    break;
                case FilterType.DateTo:
                    _datetoFilter = WaitForElementIsVisible(By.Id(DATETO_FILTER));
                    _datetoFilter.Clear();
                    _datetoFilter.SetValue(ControlType.TextBox, value);
                    WaitForLoad();
                    _datetoFilter.SendKeys(Keys.Enter);
                    break;
                case FilterType.Site:
                    _siteFilter = WaitForElementIsVisible(By.Id(SITE_FILTER));
                    _siteFilter.SetValue(ControlType.DropDownList, value);
                    WaitForLoad();
                    break;
                case FilterType.ShowItems:
                    _showItems = WaitForElementExists(By.Id(SHOW_ITEMS_FILTER));
                    _showItems.SetValue(ControlType.CheckBox, value);
                    WaitForLoad();
                    break;
                case FilterType.ShowProcedure:
                    _showProcedure = WaitForElementIsVisible(By.Id(SHOW_PROCEDURE_FILTER));
                    _showProcedure.SetValue(ControlType.CheckBox, value);
                    WaitForLoad();
                    break;
                case FilterType.ShowPaxIsZero:
                    _showPaxIsZero = WaitForElementIsVisible(By.Id(SHOW_PAX_IS_ZERO_FILTER));
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
                    _workshopsFilter = WaitForElementIsVisible(By.Id(WORKSHOPS_FILTER));
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
                    _itemGroupsFilter = WaitForElementIsVisible(By.Id(ITEM_GROUPS_FILTER));
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
                    ComboBoxSelectById(new ComboBoxOptions(GUESTTYPES_FILTER, (string)value, false));
                    break;
                case FilterType.DeliveryRounds:
                    ComboBoxSelectById(new ComboBoxOptions(DELIVERY_ROUNDS_FILTER, (string)value, false));
                    break;
                case FilterType.FoodpackTypes:
                    ComboBoxSelectById(new ComboBoxOptions(FOODPACK_TYPES_FILTER, (string)value, false));
                    break;
                case FilterType.Deliveries:
                    _deliveriesFilter = WaitForElementIsVisible(By.Id(DELIVERIES_FILTER));
                    _deliveriesFilter.Click();

                    _deliveriesFilterUncheckAll = WaitForElementIsVisible(By.XPath(DELIVERIES_FILTER_UNCHECK_ALL));
                    _deliveriesFilterUncheckAll.Click();

                    _deliveriesFilterSearch = WaitForElementIsVisible(By.XPath(DELIVERIES_FILTER_SEARCH));
                    _deliveriesFilterSearch.SetValue(ControlType.TextBox, value);

                    var _deliveryToCheck = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    _deliveryToCheck.SetValue(ControlType.CheckBox, true);
                    WaitForLoad();
                    break;
            }
        }

        public ProductionSearchResultsMenuTabPage Validate()
        {
            _btnValidate = WaitForElementIsVisible(By.XPath(BTN_VALIDATE));
            _btnValidate.Click();
            return new ProductionSearchResultsMenuTabPage(_webDriver, _testContext);
        }

        // GETTERS
        public string GetSelectedSite()
        {
            _siteFilter = WaitForElementIsVisible(By.Id(SITE_FILTER));
            return _siteFilter.Text;
        }

        public string GetSelectedCustomers()
        {
            _selectedCustomers = WaitForElementIsVisible(By.XPath(SELECTED_CUSTOMERS));
            return _selectedCustomers.Text;
        }

        public string GetSelectedDate()
        {
            _dateFilter = WaitForElementIsVisible(By.Id(DATE_FILTER));
            return _dateFilter.GetAttribute("value");
        }
        public IWebElement GetByMenuButton()
        {
            var byMenyButton = _webDriver.FindElement(By.ClassName(BYMENUBUTTON));
            return byMenyButton;
        }
        public IWebElement GetByRecipeButton()
        {
            var byRecipeButton = _webDriver.FindElement(By.ClassName(BYRECIPEBUTTON));
            return byRecipeButton;
        }
        public IWebElement GetByCustomerButton()
        {
            var byCustomerButton = _webDriver.FindElement(By.ClassName(BYCUSTOMERBUTTON));
            return byCustomerButton;
        }
        public IWebElement GetByWorkshopButton()
        {
            var byWorkshopButton = _webDriver.FindElement(By.ClassName(BYWORKSHOPBUTTON));
            return byWorkshopButton;
        }
        public IWebElement GetByCustomerOrder()
        {
            var byCustomerOrder = _webDriver.FindElement(By.ClassName(BYCUSTOMERORDERBUTTON));
            return byCustomerOrder;
        }
        public IWebElement GetCheckBox()
        {
            _showItems = WaitForElementExists(By.Id(SHOW_ITEMS_FILTER));
            return _showItems;
        }

        public void Print(PrintType printType, bool value , string Comment = "")
        {
            /*Actions a = new Actions(_webDriver);
            var etcButton = WaitForElementIsVisible(By.XPath("//*[@id='div-body']/div/div[1]/div/div[1]/button/parent::div"));
            a.MoveToElement(etcButton).Perform();*/
            var ProductionButton = WaitForElementIsVisible(By.XPath(PRODUCTIONBUTTON));
            ProductionButton.Click();
            WaitForLoad();
            switch (printType)
            {
                case PrintType.Menu:
                    _menuCheckBox = WaitForElementIsVisible(By.XPath(MENUCHECKBOX));
                    _menuCheckBox.SetValue(ControlType.CheckBox, value);
                    WaitForLoad();
                    break;
                case PrintType.Recipe:
                    _recipeCheckBox = WaitForElementIsVisible(By.XPath(RECIPECHECKBOX));
                    _recipeCheckBox.SetValue(ControlType.CheckBox, value);
                    WaitForLoad();
                    break;
                case PrintType.Customer:
                    _customerCheckBox = WaitForElementIsVisible(By.XPath(CUSTOMERCHECKBOX));
                    _customerCheckBox.SetValue(ControlType.CheckBox, value);
                    WaitForLoad();
                    break;
                case PrintType.ItemGroupe:
                    _itemGroupeCheckBox = WaitForElementIsVisible(By.XPath(ITEMGROUPECHECKBOX));
                    _itemGroupeCheckBox.SetValue(ControlType.CheckBox, value);
                    WaitForLoad();
                    break;
                case PrintType.Workshop:
                    _workShopCheckBox = WaitForElementIsVisible(By.XPath(WORKSHOPCHECKBOX));
                    _workShopCheckBox.SetValue(ControlType.CheckBox, value);
                    WaitForLoad();
                    break;
                case PrintType.CustomerOrder:
                    _customerOrderCheckBox = WaitForElementIsVisible(By.XPath(CUSTOMERCHECKBOX));
                    _customerOrderCheckBox.SetValue(ControlType.CheckBox, value);
                    WaitForLoad();
                    break;
            }
            if(printType== PrintType.Workshop)
            {
                _commentSection = WaitForElementIsVisible(By.XPath(COMMENTSECTION));
                _commentSection.SetValue(ControlType.TextBox, Comment);
                WaitForLoad();
            }
            var validateButton = WaitForElementIsVisible(By.XPath(VALIDATEBUTTON));
            validateButton.Click();

        }

        public PrintReportPage ClickPrint()
        {
             IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-pdf']"));
             ClickPrintButton();

                
             WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(30));
             wait.Until((driver) => driver.WindowHandles.Count > 1);
             _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new PrintReportPage(_webDriver, _testContext);

        }
        public bool CheckCommentExist(string comment, List<string> wordsList)
        {
            string[] words = comment.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            int currentIndex = 0;

            foreach (var word in words)
            {
                currentIndex = wordsList.IndexOf(word, currentIndex);
                if (currentIndex == -1)
                {
                    return false;
                }
                currentIndex++;
            }

            return true;
        }

        public bool IsPageLoadedWithoutError()
        {
            if (isElementVisible(By.XPath(PRODUCTION_PAGE)))
            {
                _productionPage = _webDriver.FindElement(By.XPath(PRODUCTION_PAGE));
                return true;
            }
            else
            {
                return false;
            }
        }

     
        public FilterAndFavoritesPage ClickOnResetSearch()
        {
            var reset = WaitForElementIsVisible(By.XPath("//*[@id=\"favorite-filter-form\"]/div[1]/div/div/span[2]/button"));
            reset.Click();
            WaitForLoad();
            return new FilterAndFavoritesPage(_webDriver, _testContext);
        }

    }
}
