using DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentAssertions.Common;
using iText.StyledXmlParser.Jsoup.Select;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Datasheet;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Production.ProductionManagement
{
    public class QuantityAdjustmentsPage : PageBase
    {

        public QuantityAdjustmentsPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //__________________________________Constantes_____________________________________

        // Général
        private const string EXTENDED_MENU = "//*[@id=\"div-body\"]/div/div[2]/div[1]/div/div[1]/button";
        private const string RESET_QUANTITIES = "btnResetQuantities";
        private const string RAZ_QUANTITIES = "btnRazQuantities";
        private const string PREVIOUS_DATE = "prev-date";
        private const string NEXT_DATE = "next-date";
        private const string NEXT = "inflightrawmaterials-quantityadjustments-detailpopupcontent-nextcustomer";
        private const string PREVIOUS = "inflightrawmaterials-quantityadjustments-detail-popupcontent-previouscustomer";

        // Onglets
        private const string RESULT_TAB = "//*[@id=\"hrefTabContentItemContainer\"][text()='Results']";

        // Tableau
        private const string DATE = "//*[@id=\"list-item-with-action\"]/div/table/thead/tr/th[@class=\"day-col\"]";
        private const string SERVICES = "//*[@id=\"list-item-with-action\"]/div/table/tbody/tr";
        private const string ICON_DATASHEET = "//*[@id=\"list-item-with-action\"]/div/table/tbody/tr[*]/td[1]/span";
        private const string SERVICE_NAME = "inflightrawmaterials-production-servicename-detail-1";
        private const string SERVICE_CUSTOMER = "//*[@id=\"list-item-with-action\"]/div/table/tbody/tr[*]/td[3]/div";
        private const string SERVICE_QUANTITY = "//*[@id=\"list-item-with-action\"]/div/table/tbody/tr[*]/td[contains(text(),'{0}')]/../td[6]/span[1]";
        private const string QUANTITY_INPUT = "//*[@id=\"list-item-with-action\"]/div/table/tbody/tr[*]/td[contains(text(),'{0}')]/../td[6]/input[6]";
        private const string QUANTITY_INPUT_WIDHOUT_SERVICE = "/html/body/div[3]/div/div/div[3]/div[2]/div/table/tbody/tr/td[6]/input[7]";
        private const string EDIT_FIRST_SERVICE = "//*[@id=\"list-item-with-action\"]/div/table/tbody/tr[1]/td[2]/a";
        private const string FIRST_SERVICE = "//*[@id=\"list-item-with-action\"]/div/table/tbody/tr[1]/td[2]";
        private const string POPUP_ICON = "//*[@id=\"list-item-with-action\"]/div/table/tbody/tr[1]/td[8]/a/span";
        //private const string POPUP_ICON2 = "inflightrawmaterials-production-quantityadjustments-detail-1";
        private const string POPUP_ICON2 = "//*[@id=\"list-item-with-action\"]/div/table/tbody/tr[*]/td[contains(text(),'{0}')]/../td[10]/a";
        // Filtres
        private const string BACK = "//*[@id=\"item-filter-form\"]/div[1]/a";

        private const string SITE = "SiteId";
        private const string SITE_SELECTED = "//*[@id=\"SiteId\"]/option[@selected = 'selected']";
        private const string SORTBY = "cbSortBy";

        private const string FILTER_DATE_FROM = "ProdDateFrom";
        private const string FILTER_DATE_TO = "ProdDateTo";

        private const string PRODMAN_FILTER_START_TIME_HOUR = "//*[@id=\"item-filter-form\"]/div[7]/div/input[2]";
        private const string PRODMAN_FILTER_START_TIME_MIN = "//*[@id=\"item-filter-form\"]/div[7]/div/input[3]";
        private const string PRODMAN_FILTER_END_TIME_HOUR = "//*[@id=\"item-filter-form\"]/div[8]/div/input[2]";
        private const string PRODMAN_FILTER_END_TIME_MIN = "//*[@id=\"item-filter-form\"]/div[8]/div/input[3]";

        private const string SHOW_WITHOUT_DATASHEET_ONLY = "ShowWithoutDatasheetOnly";
        private const string SHOW_HIDDEN_ARTICLE = "ShowHiddenArticles";
        private const string SHOW_VALIDATE_FLIGHT = "ShowValidatedFlightsOnly";

        private const string SHOW_NORMAL_AND_VACUUM_PROD = "includeVacuumPackedAll";
        private const string SHOW_VACUUM_PROD = "includeVacuumPackedOnly";
        private const string SHOW_NORMAL_PROD = "IncludeNovacuumPacked";

        private const string WORKSHOP_FILTER = "SelectedWorkshops_ms";
        private const string UNSELECT_WORKSHOP = "/html/body/div[10]/div/ul/li[2]/a/span[2]";
        private const string SEARCH_WORKSHOP = "/html/body/div[10]/div/div/label/input";
        private const string WORKSHOP_SELECTED = "/html/body/div[10]/ul/li[*]/label/input[@checked='checked']/../span";

        private const string CUSTOMER_FILTER = "SelectedCustomers_ms";
        private const string UNSELECT_CUSTOMER = "/html/body/div[11]/div/ul/li[2]/a";
        private const string SEARCH_CUSTOMER = "/html/body/div[11]/div/div/label/input";
        private const string CUSTOMER_SELECTED = "/html/body/div[11]/ul/li[*]/label/input[@checked='checked']/../span";

        private const string GUEST_TYPE_FILTER = "SelectedGuestTypes_ms";
        private const string UNSELECT_GUEST_TYPE = "/html/body/div[12]/div/ul/li[2]/a";
        private const string SEARCH_GUEST_TYPE = "/html/body/div[12]/div/div/label/input";
        private const string GUEST_TYPE_SELECTED = "/html/body/div[12]/ul/li[*]/label/input[@checked='checked']/../span";

        private const string SERVICE_CATEGORIE_FILTER = "SelectedServiceCategories_ms";
        private const string UNSELECT_CATEGORIE_SERVICE = "/html/body/div[13]/div/ul/li[2]/a";
        private const string SEARCH_CATEGORIE_SERVICE = "/html/body/div[13]/div/div/label/input";
        private const string SERVICE_CATEGORIE_SELECTED = "/html/body/div[13]/ul/li[*]/label/input[@checked='checked']/../span";

        private const string SERVICE_FILTER = "SelectedServices_ms";
        private const string UNSELECT_SERVICE = "/html/body/div[14]/div/ul/li[2]/a";
        private const string SEARCH_SERVICE = "/html/body/div[14]/div/div/label/input";
        private const string SERVICE_SELECTED = "/html/body/div[14]/ul/li[*]/label/input[@checked='checked']/../span";

        private const string RECIPE_TYPE_FILTER = "SelectedRecipeTypes_ms";
        private const string UNSELECT_RECIPE_TYPE = "/html/body/div[15]/div/ul/li[2]/a";
        private const string SEARCH_RECIPE_TYPE = "/html/body/div[15]/div/div/label/input";
        private const string RECIPE_TYPE_SELECTED = "/html/body/div[15]/ul/li[*]/label/input[@checked='checked']/../span";

        private const string ITEM_GROUP_FILTER = "SelectedItemGroups_ms";
        private const string UNSELECT_ITEM_GROUP = "/html/body/div[16]/div/ul/li[2]/a";
        private const string SEARCH_ITEM_GROUP = "/html/body/div[16]/div/div/label/input";
        private const string ITEM_GROUP_SELECTED = "/html/body/div[16]/ul/li[*]/label/input[@checked='checked']/../span";
        private const string FILTRED_TAB_SERVICES = "//*[@id=\"list-item-with-action\"]/div/table/tbody/tr/td[2]";
        private const string QUANTITY_ADJ_TABLE_ITEMS = "//*[@id=\"list-item-with-action\"]/div/table/tbody/tr[*]/td[2]";
        private const string SERVICE_CONSULT_FLESH = "//a[@class='btn-with-icon btn btn-open-qty-details']";
        private const string LISTE_VOL_QTYDETAILS = "//div[1]/table/tbody/tr[*]/td[2]/div/a[@class='btn-line']";
        private const string FLIGHT = "/html/body/div[3]/div/div/div[3]/div[2]/div/table/tbody/tr/td[1]";
        private const string SELECTED_GROUPBY = "//*[@id=\"DdlSelectedGroupBy\"]";

        private const string REFRESH_SERVICE_BTN = "//*[@id=\"list-item-with-action\"]/div/table/tbody[1]/tr[@class='line-selected']/td[6]/a";
        private const string SELECT_SERVICE_LINE = "/html/body/div[3]/div/div/div[3]/div[2]/div/table/tbody/tr/td[1]";
        private const string SERVICE_QUANTITY_LINE = "//input[@class='input-number text-select-on-click flight-adjusted-qty-input']";
        private const string ITEM_SUB_GROUP_FILTER = "SelectedItemSubGroups_ms";
        private const string UNSELECT_ITEM_SUB_GROUP = "/html/body/div[17]/div/ul/li[2]/a";
        private const string SEARCH_ITEM_SUB_GROUP = "/html/body/div[17]/div/div/label/input";
        private const string ITEM_SUB_GROUP_SELECTED = "/html/body/div[17]/ul/li[*]/ul/li/label/input[@checked='checked']/../span";
        private const string SERVICE_ROW = "//*[@id=\"list-item-with-action\"]/div/table/tbody/tr[{0}]";

        //private const string TOTAL_SERVICE = "//*[@id=\"list-item-with-action\"]/div/table/tbody/tr[1]/td[4]";
        private const string TOTAL_SERVICE = "//*[@id=\"list-item-with-action\"]/div/table/tbody/tr[*]/td[contains(text(),'{0}')]/../td[4]";

        private const string COLTOTAL = "//*[@id=\"list-item-with-action\"]/div/table/tbody/tr[*]/td[5]";
        private const string NOMBRE_PICTO_VERT = "//*[@id=\"list-item-with-action\"]/div/table/tbody/tr[*]/td[1]/a";
        private const string FIRST_QUANTITY = "inflightrawmaterials-production-service-row-detail-1";
        private const string FIRST_ICON_POP_UP = "inflightrawmaterials-production-quantityadjustments-detail-1"; 


        //__________________________________Variables______________________________________

        // General
        [FindsBy(How = How.XPath, Using = POPUP_ICON)]
        public IWebElement _poppupicon;

        [FindsBy(How = How.XPath, Using = EXTENDED_MENU)]
        public IWebElement _extendedMenu;

        [FindsBy(How = How.Id, Using = RESET_QUANTITIES)]
        public IWebElement _resetQuantities;

        [FindsBy(How = How.Id, Using = RAZ_QUANTITIES)]
        public IWebElement _razQuantities;

        [FindsBy(How = How.Id, Using = PREVIOUS_DATE)]
        public IWebElement _previousDate;

        [FindsBy(How = How.Id, Using = NEXT_DATE)]
        public IWebElement _nextDate;

        [FindsBy(How = How.Id, Using = PREVIOUS)]
        public IWebElement _previous;

        [FindsBy(How = How.Id, Using = NEXT)]
        public IWebElement _next;

        // Onglets
        [FindsBy(How = How.XPath, Using = RESULT_TAB)]
        public IWebElement _resultTab;

        // Tableau
        [FindsBy(How = How.XPath, Using = SERVICES)]
        public IWebElement _services;

        [FindsBy(How = How.Id, Using = SERVICE_NAME)]
        public IWebElement _serviceName;

        [FindsBy(How = How.XPath, Using = QUANTITY_INPUT)]
        public IWebElement _inputQty;

        [FindsBy(How = How.XPath, Using = SERVICE_QUANTITY)]
        public IWebElement _quantity;

        [FindsBy(How = How.XPath, Using = FILTRED_TAB_SERVICES)]
        public IWebElement _filtred_tab_services;


        [FindsBy(How = How.XPath, Using = SERVICE_CONSULT_FLESH)]
        public IWebElement _serviceConsultFlesh;

        [FindsBy(How = How.XPath, Using = TOTAL_SERVICE)]
        public IWebElement _totalService;

        [FindsBy(How = How.XPath, Using = COLTOTAL)]
        public IWebElement _colTotal;

        [FindsBy(How = How.Id, Using = FIRST_QUANTITY)]
        public IWebElement _fistQuantity;
        
        // __________________________________ Filtres __________________________________________

        [FindsBy(How = How.XPath, Using = BACK)]
        private IWebElement _back;

        [FindsBy(How = How.Id, Using = SITE)]
        private IWebElement _site;

        [FindsBy(How = How.XPath, Using = SORTBY)]
        private IWebElement _sortBy;

        [FindsBy(How = How.Id, Using = FILTER_DATE_FROM)]
        public IWebElement _dateFromFilter;

        [FindsBy(How = How.Id, Using = FILTER_DATE_TO)]
        public IWebElement _dateToFilter;

        [FindsBy(How = How.Id, Using = PRODMAN_FILTER_START_TIME_HOUR)]
        public IWebElement _filterStartTimeHour;

        [FindsBy(How = How.Id, Using = PRODMAN_FILTER_START_TIME_MIN)]
        public IWebElement _filterStartTimeMin;

        [FindsBy(How = How.Id, Using = PRODMAN_FILTER_END_TIME_HOUR)]
        public IWebElement _filterEndTimeHour;

        [FindsBy(How = How.Id, Using = PRODMAN_FILTER_END_TIME_MIN)]
        public IWebElement _filterEndTimeMin;

        [FindsBy(How = How.Id, Using = SHOW_WITHOUT_DATASHEET_ONLY)]
        private IWebElement _showWithoutDatasheetOnly;

        [FindsBy(How = How.Id, Using = SHOW_HIDDEN_ARTICLE)]
        public IWebElement _showHiddenArticle;

        [FindsBy(How = How.Id, Using = SHOW_VALIDATE_FLIGHT)]
        public IWebElement _showValidateFlight;

        [FindsBy(How = How.XPath, Using = SHOW_NORMAL_AND_VACUUM_PROD)]
        private IWebElement _showNormalAndVacuumProd;

        [FindsBy(How = How.Id, Using = SHOW_VACUUM_PROD)]
        private IWebElement _showVacuumProd;

        [FindsBy(How = How.Id, Using = SHOW_NORMAL_PROD)]
        private IWebElement _showNormalProd;

        [FindsBy(How = How.XPath, Using = WORKSHOP_FILTER)]
        private IWebElement _workShopFilter;

        [FindsBy(How = How.XPath, Using = UNSELECT_WORKSHOP)]
        private IWebElement _unselectAllWorkShop;

        [FindsBy(How = How.XPath, Using = SEARCH_WORKSHOP)]
        private IWebElement _searchWorkShop;

        [FindsBy(How = How.XPath, Using = CUSTOMER_FILTER)]
        private IWebElement _customerFilter;

        [FindsBy(How = How.XPath, Using = UNSELECT_CUSTOMER)]
        private IWebElement _unselectAllCustomer;

        [FindsBy(How = How.XPath, Using = SEARCH_CUSTOMER)]
        private IWebElement _searchCustomer;

        [FindsBy(How = How.XPath, Using = GUEST_TYPE_FILTER)]
        private IWebElement _guestTypeFilter;

        [FindsBy(How = How.XPath, Using = UNSELECT_GUEST_TYPE)]
        private IWebElement _unselectAllGuestType;

        [FindsBy(How = How.XPath, Using = SEARCH_GUEST_TYPE)]
        private IWebElement _searchGuestType;

        [FindsBy(How = How.XPath, Using = SERVICE_CATEGORIE_FILTER)]
        private IWebElement _serviceCategoryFilter;

        [FindsBy(How = How.XPath, Using = UNSELECT_CATEGORIE_SERVICE)]
        private IWebElement _unselectAllServiceCategory;

        [FindsBy(How = How.XPath, Using = SEARCH_CATEGORIE_SERVICE)]
        private IWebElement _searchCategoryService;

        [FindsBy(How = How.XPath, Using = SERVICE_FILTER)]
        private IWebElement _serviceFilter;

        [FindsBy(How = How.XPath, Using = UNSELECT_SERVICE)]
        private IWebElement _unselectAllService;

        [FindsBy(How = How.XPath, Using = SEARCH_SERVICE)]
        private IWebElement _searchService;

        [FindsBy(How = How.XPath, Using = UNSELECT_RECIPE_TYPE)]
        private IWebElement _unselectAllRecipeType;

        [FindsBy(How = How.XPath, Using = RECIPE_TYPE_FILTER)]
        private IWebElement _recipeTypeFilter;

        [FindsBy(How = How.XPath, Using = SEARCH_RECIPE_TYPE)]
        private IWebElement _searchRecipeType;

        [FindsBy(How = How.XPath, Using = ITEM_GROUP_FILTER)]
        private IWebElement _itemGroupFilter;

        [FindsBy(How = How.XPath, Using = UNSELECT_ITEM_GROUP)]
        private IWebElement _unselectAllItemGroup;

        [FindsBy(How = How.XPath, Using = SEARCH_ITEM_GROUP)]
        private IWebElement _searchItemGroup;

        [FindsBy(How = How.XPath, Using = EDIT_FIRST_SERVICE)]
        private IWebElement _editFirstService;

        [FindsBy(How = How.XPath, Using = FIRST_SERVICE)]
        private IWebElement _firstService;

        [FindsBy(How = How.XPath, Using = ITEM_SUB_GROUP_FILTER)]
        private IWebElement _itemSubGroupFilter;

        [FindsBy(How = How.XPath, Using = UNSELECT_ITEM_SUB_GROUP)]
        private IWebElement _unselectAllItemSubGroup;

        [FindsBy(How = How.XPath, Using = SEARCH_ITEM_SUB_GROUP)]
        private IWebElement _searchItemSubGroup;

        [FindsBy(How = How.XPath, Using = POPUP_ICON2)]
        public IWebElement _poppupicon2;

        public enum FilterType
        {
            Site,
            SortBy,
            DateFrom,
            DateTo,
            StartTime,
            EndTime,
            ShowHiddenArticles,
            ShowValidateFlight,
            ShowWithoutDatasheet,
            ShowNormalAndVacuumProd,
            ShowVacuumProd,
            ShowNormalProd,
            Workshops,
            Customers,
            GuestType,
            ServicesCategorie,
            Services,
            RecipeType,
            ItemGroups,
            ItemSubGroups
        }

        public void Filter(FilterType filterType, object value)
        {
            Actions action = new Actions(_webDriver);
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;

            switch (filterType)
            {
                case FilterType.Site:
                    _site = WaitForElementIsVisible(By.Id(SITE));
                    _site.SetValue(ControlType.DropDownList, value);
                    break;

                case FilterType.SortBy:
                    _sortBy = WaitForElementIsVisible(By.Id(SORTBY));
                    _sortBy.Click();
                    var element = WaitForElementIsVisible(By.XPath("//option[contains(@value,'" + value + "')]"));
                    _sortBy.SetValue(ControlType.DropDownList, element.Text);
                    _sortBy.Click();
                    break;

                case FilterType.DateFrom:
                    _dateFromFilter = WaitForElementIsVisible(By.Id(FILTER_DATE_FROM));
                    _dateFromFilter.SetValue(ControlType.DateTime, value);
                    _dateFromFilter.SendKeys(Keys.Enter);
                    WaitForLoad();
                    break;

                case FilterType.DateTo:
                    _dateToFilter = WaitForElementIsVisible(By.Id(FILTER_DATE_TO));
                    _dateToFilter.SetValue(ControlType.DateTime, value);
                    _dateToFilter.SendKeys(Keys.Enter);
                    WaitForLoad();
                    break;

                case FilterType.StartTime:
                    _filterStartTimeHour = WaitForElementIsVisible(By.XPath(PRODMAN_FILTER_START_TIME_HOUR), nameof(PRODMAN_FILTER_START_TIME_HOUR));
                    _filterStartTimeHour.ClearElement();
                    _filterStartTimeHour.SetValue(ControlType.TextBox, value.ToString().Substring(0, 2));
                    _filterStartTimeHour.SendKeys(Keys.Tab);

                    _filterStartTimeMin = WaitForElementIsVisible(By.XPath(PRODMAN_FILTER_START_TIME_MIN), nameof(PRODMAN_FILTER_START_TIME_MIN));
                    _filterStartTimeMin.ClearElement();
                    _filterStartTimeMin.SetValue(ControlType.TextBox, value.ToString().Substring(value.ToString().LastIndexOf(':') + 1));
                    _filterStartTimeMin.SendKeys(Keys.Tab);
                    WaitForLoad();
                    break;

                case FilterType.EndTime:
                    _filterEndTimeHour = WaitForElementIsVisible(By.XPath(PRODMAN_FILTER_END_TIME_HOUR), nameof(PRODMAN_FILTER_END_TIME_HOUR));
                    _filterEndTimeHour.ClearElement();
                    _filterEndTimeHour.SetValue(ControlType.TextBox, value.ToString().Substring(0, 2));
                    _filterEndTimeHour.SendKeys(Keys.Tab);

                    _filterEndTimeMin = WaitForElementIsVisible(By.XPath(PRODMAN_FILTER_END_TIME_MIN), nameof(PRODMAN_FILTER_END_TIME_MIN));
                    _filterEndTimeMin.ClearElement();
                    _filterEndTimeMin.SetValue(ControlType.TextBox, value.ToString().Substring(value.ToString().LastIndexOf(':') + 1));
                    _filterEndTimeMin.SendKeys(Keys.Tab);
                    WaitForLoad();
                    break;

                case FilterType.ShowHiddenArticles:
                    _showHiddenArticle = WaitForElementExists(By.Id(SHOW_HIDDEN_ARTICLE));
                    _showHiddenArticle.SetValue(ControlType.CheckBox, true);
                    WaitForLoad();
                    break;

                case FilterType.ShowValidateFlight:
                    _showValidateFlight = WaitForElementExists(By.Id(SHOW_VALIDATE_FLIGHT));
                    _showValidateFlight.SetValue(ControlType.CheckBox, true);
                    WaitForLoad();
                    break;

                case FilterType.ShowWithoutDatasheet:
                    _showWithoutDatasheetOnly = WaitForElementExists(By.Id(SHOW_WITHOUT_DATASHEET_ONLY));
                    _showWithoutDatasheetOnly.SetValue(ControlType.CheckBox, true);
                    WaitForLoad();
                    break;

                case FilterType.ShowNormalAndVacuumProd:
                    _showNormalAndVacuumProd = WaitForElementExists(By.Id(SHOW_NORMAL_AND_VACUUM_PROD));
                    action.MoveToElement(_showNormalAndVacuumProd).Perform();
                    _showNormalAndVacuumProd.SetValue(ControlType.RadioButton, value);
                    WaitForLoad();
                    break;

                case FilterType.ShowVacuumProd:
                    _showVacuumProd = WaitForElementExists(By.Id(SHOW_VACUUM_PROD));
                    action.MoveToElement(_showVacuumProd).Perform();
                    _showVacuumProd.SetValue(ControlType.RadioButton, value);
                    WaitForLoad();
                    break;

                case FilterType.ShowNormalProd:
                    _showNormalProd = WaitForElementExists(By.Id(SHOW_NORMAL_PROD));
                    action.MoveToElement(_showNormalProd).Perform();
                    _showNormalProd.SetValue(ControlType.RadioButton, value);
                    WaitForLoad();
                    break;

                case FilterType.Workshops:
                    _workShopFilter = WaitForElementIsVisible(By.Id(WORKSHOP_FILTER));
                    _workShopFilter.Click();

                    // On décoche toutes les options
                    _unselectAllWorkShop = WaitForElementIsVisible(By.XPath(UNSELECT_WORKSHOP));
                    _unselectAllWorkShop.Click();
                    WaitPageLoading();

                    _searchWorkShop = WaitForElementIsVisible(By.XPath(SEARCH_WORKSHOP));
                    _searchWorkShop.SetValue(ControlType.TextBox, value);

                    var resultWorkShop = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    resultWorkShop.SetValue(ControlType.CheckBox, true);

                    _workShopFilter.Click();
                    break;

                case FilterType.Customers:
                    ComboBoxSelectById(new ComboBoxOptions(CUSTOMER_FILTER, (string)value));
                    break;

                case FilterType.GuestType:
                    _guestTypeFilter = WaitForElementIsVisible(By.Id(GUEST_TYPE_FILTER));
                    _guestTypeFilter.Click();

                    // On décoche toutes les options
                    _unselectAllGuestType = WaitForElementIsVisible(By.XPath(UNSELECT_GUEST_TYPE));
                    _unselectAllGuestType.Click();
                    WaitPageLoading();

                    _searchGuestType = WaitForElementIsVisible(By.XPath(SEARCH_GUEST_TYPE));
                    _searchGuestType.SetValue(ControlType.TextBox, value);

                    var resultGuest = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    resultGuest.SetValue(ControlType.CheckBox, true);

                    _guestTypeFilter.Click();
                    break;

                case FilterType.ServicesCategorie:
                    //go down in the page
                    _workShopFilter = WaitForElementIsVisible(By.Id(WORKSHOP_FILTER));
                    javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", _workShopFilter);
                    WaitForLoad();
                    WaitPageLoading();
                    _serviceCategoryFilter = WaitForElementIsVisible(By.Id(SERVICE_CATEGORIE_FILTER));
                    _serviceCategoryFilter.Click();
                    WaitForLoad();
                    WaitPageLoading();

                    // On décoche toutes les options
                    _unselectAllServiceCategory = WaitForElementIsVisible(By.XPath(UNSELECT_CATEGORIE_SERVICE));
                    WaitForLoad();
                    WaitPageLoading();
                    _unselectAllServiceCategory.Click();
                    WaitPageLoading();
                    _searchCategoryService = WaitForElementIsVisible(By.XPath(SEARCH_CATEGORIE_SERVICE));
                    _searchCategoryService.SetValue(ControlType.TextBox, value);
                    WaitPageLoading();
                    var serviceCat = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    serviceCat.SetValue(ControlType.CheckBox, true);

                    _serviceCategoryFilter.Click();
                    WaitForLoad();
                    break;

                case FilterType.Services:
                    //go down in the page
                    _workShopFilter = WaitForElementIsVisible(By.Id(WORKSHOP_FILTER));
                    javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", _workShopFilter);

                    _serviceFilter = WaitForElementIsVisible(By.Id(SERVICE_FILTER));
                    _serviceFilter.Click();

                    // On décoche toutes les options
                    _unselectAllService = WaitForElementIsVisible(By.XPath(UNSELECT_SERVICE));
                    _unselectAllService.Click();
                    WaitPageLoading();

                    _searchService = WaitForElementIsVisible(By.XPath(SEARCH_SERVICE));
                    _searchService.SetValue(ControlType.TextBox, value);

                    var service = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    WaitPageLoading();
                    service.SetValue(ControlType.CheckBox, true);
                    WaitPageLoading();
                    _serviceFilter.Click();
                    WaitForLoad();
                    break;

                case FilterType.RecipeType:
                    //go down in the page
                    _workShopFilter = WaitForElementIsVisible(By.Id(WORKSHOP_FILTER));
                    javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", _workShopFilter);

                    _recipeTypeFilter = WaitForElementExists(By.Id(RECIPE_TYPE_FILTER));
                    action.MoveToElement(_recipeTypeFilter).Perform();
                    _recipeTypeFilter.Click();

                    // On décoche toutes les options
                    _unselectAllRecipeType = WaitForElementIsVisible(By.XPath(UNSELECT_RECIPE_TYPE));
                    _unselectAllRecipeType.Click();
                    WaitPageLoading();

                    _searchRecipeType = WaitForElementIsVisible(By.XPath(SEARCH_RECIPE_TYPE));
                    _searchRecipeType.SetValue(ControlType.TextBox, value);

                    var resultRecipe = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    resultRecipe.SetValue(ControlType.CheckBox, true);

                    _recipeTypeFilter.Click();
                    WaitForLoad();
                    break;

                case FilterType.ItemSubGroups:
                    _itemSubGroupFilter = WaitForElementExists(By.Id(ITEM_SUB_GROUP_FILTER));
                    action.MoveToElement(_itemSubGroupFilter).Perform();
                    _itemSubGroupFilter.Click();

                    // On décoche toutes les options
                    _unselectAllItemSubGroup = WaitForElementIsVisible(By.XPath(UNSELECT_ITEM_SUB_GROUP));
                    _unselectAllItemSubGroup.Click();
                    WaitPageLoading();

                    _searchItemSubGroup = WaitForElementIsVisible(By.XPath(SEARCH_ITEM_SUB_GROUP));
                    _searchItemSubGroup.SetValue(ControlType.TextBox, value);
                    Thread.Sleep(2000);

                    var resultSubItem = WaitForElementIsVisible(By.XPath("/html/body/div[17]/ul"));
                    resultSubItem.Click();

                    _itemSubGroupFilter.Click();
                    WaitForLoad();
                    break;
                case FilterType.ItemGroups:
                    _itemGroupFilter = WaitForElementExists(By.Id(ITEM_GROUP_FILTER));
                    action.MoveToElement(_itemGroupFilter).Perform();
                    _itemGroupFilter.Click();

                    // On décoche toutes les options
                    _unselectAllItemGroup = WaitForElementIsVisible(By.XPath(UNSELECT_ITEM_GROUP));
                    _unselectAllItemGroup.Click();
                    WaitPageLoading();

                    _searchItemGroup = WaitForElementIsVisible(By.XPath(SEARCH_ITEM_GROUP));
                    _searchItemGroup.SetValue(ControlType.TextBox, value);
                    Thread.Sleep(2000);

                    var resultItem = WaitForElementIsVisible(By.XPath("/html/body/div[16]/ul"));
                    resultItem.Click();

                    _itemGroupFilter.Click();
                    WaitForLoad();
                    break;

                default:
                    break;
            }

            WaitPageLoading();
            WaitPageLoading();
        }

        public FilterAndFavoritesPage Back()
        {
            _back = WaitForElementIsVisible(By.XPath(BACK));
            _back.Click();
            WaitForLoad();

            return new FilterAndFavoritesPage(_webDriver, _testContext);
        }

        public string GetSite()
        {
            WaitForLoad(); 
            _site = WaitForElementIsVisible(By.XPath(SITE_SELECTED));
            return _site.Text;
        }

        public DateTime GetDateFrom()
        {
            _dateFromFilter = WaitForElementIsVisible(By.Id(FILTER_DATE_FROM));

            var dateFormat = _dateFromFilter.GetAttribute("data-date-format");
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            return DateTime.Parse(_dateFromFilter.GetAttribute("value"), ci);
        }

        public DateTime GetDateTo()
        {
            _dateToFilter = WaitForElementIsVisible(By.Id(FILTER_DATE_TO));

            var dateFormat = _dateToFilter.GetAttribute("data-date-format");
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            return DateTime.Parse(_dateToFilter.GetAttribute("value"), ci);
        }

        public string GetWorkshop()
        {
            _workShopFilter = WaitForElementIsVisible(By.Id(WORKSHOP_FILTER));
            _workShopFilter.Click();

            var workshop = _webDriver.FindElement(By.XPath(WORKSHOP_SELECTED));
            return workshop.Text;
        }

        public string GetCustomer()
        {
            _customerFilter = WaitForElementIsVisible(By.Id(CUSTOMER_FILTER));
            _customerFilter.Click();

            var customer = _webDriver.FindElement(By.XPath(CUSTOMER_SELECTED));
            return customer.Text;
        }

        public string GetGuestType()
        {
            _guestTypeFilter = WaitForElementIsVisible(By.Id(GUEST_TYPE_FILTER));
            _guestTypeFilter.Click();

            var guestType = _webDriver.FindElement(By.XPath(GUEST_TYPE_SELECTED));
            return guestType.Text;
        }

        public string GetServiceCategorie()
        {
            //go down in the page
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            _workShopFilter = WaitForElementIsVisible(By.Id(WORKSHOP_FILTER));
            javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", _workShopFilter);

            _serviceCategoryFilter = WaitForElementIsVisible(By.Id(SERVICE_CATEGORIE_FILTER));
            _serviceCategoryFilter.Click();

            var serviceCategorie = _webDriver.FindElement(By.XPath(SERVICE_CATEGORIE_SELECTED));
            return serviceCategorie.Text;
        }

        public string GetService()
        {
            //go down in the page
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            _workShopFilter = WaitForElementIsVisible(By.Id(WORKSHOP_FILTER));
            javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", _workShopFilter);

            _serviceFilter = WaitForElementIsVisible(By.Id(SERVICE_FILTER));
            _serviceFilter.Click();

            var service = _webDriver.FindElement(By.XPath(SERVICE_SELECTED));
            return service.Text;
        }

        public string GetRecipeType()
        {
            //go down in the page
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            _workShopFilter = WaitForElementIsVisible(By.Id(WORKSHOP_FILTER));
            javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", _workShopFilter);

            _recipeTypeFilter = WaitForElementIsVisible(By.Id(RECIPE_TYPE_FILTER));
            _recipeTypeFilter.Click();

            var recipeType = _webDriver.FindElement(By.XPath(RECIPE_TYPE_SELECTED));
            return recipeType.Text;
        }

        public string GetItemGroup()
        {
            Actions action = new Actions(_webDriver);

            _itemGroupFilter = WaitForElementExists(By.Id(ITEM_GROUP_FILTER));
            action.MoveToElement(_itemGroupFilter).Perform();
            _itemGroupFilter.Click();

            var itemGroup = _webDriver.FindElement(By.XPath(ITEM_GROUP_SELECTED));
            return itemGroup.Text;
        }

        public bool IsNormalAndVacuumProduction()
        {
            _showNormalAndVacuumProd = _webDriver.FindElement(By.Id(SHOW_NORMAL_AND_VACUUM_PROD));
            return _showNormalAndVacuumProd.GetAttribute("checked").Equals("true");
        }

        public bool IsVacuumProduction()
        {
            _showVacuumProd = _webDriver.FindElement(By.Id(SHOW_VACUUM_PROD));
            return _showVacuumProd.GetAttribute("checked").Equals("true");
        }

        public bool IsNormalProduction()
        {
            _showNormalProd = _webDriver.FindElement(By.Id(SHOW_NORMAL_PROD));
            return _showNormalProd.GetAttribute("checked").Equals("true");
        }

        public bool IsSortedByService()
        {

            var ancientName = "";
            int compteur = 1;
            ReadOnlyCollection<IWebElement> services;

            if (IsDev())
            {
                services = _webDriver.FindElements(By.Id(SERVICE_NAME));
            }
            else
            {
                services = _webDriver.FindElements(By.XPath("//*[@id='list-item-with-action']/div/table/tbody/tr[*]/td[2]"));
            }

            if (services.Count == 0)
                return false;

            foreach (var elm in services)
            {
                if (compteur == 1)
                    ancientName = elm.Text.Trim();

                if (string.Compare(ancientName, elm.Text.Trim()) > 0)
                    return false;

                ancientName = elm.Text.Trim();
                compteur++;
            }

            return true;
        }

        public bool IsSortedByCustomer()
        {

            var ancientName = "";
            int compteur = 1;

            var customers = _webDriver.FindElements(By.XPath(SERVICE_CUSTOMER));

            if (customers.Count == 0)
                return false;

            foreach (var elm in customers)
            {
                if (compteur == 1)
                    ancientName = elm.GetAttribute("title");

                if (string.Compare(ancientName, elm.GetAttribute("title")) > 0)
                    return false;

                ancientName = elm.GetAttribute("title");
                compteur++;
            }

            return true;
        }

        public bool VerifyCustomer(string customerName)
        {

            var customers = _webDriver.FindElements(By.XPath(SERVICE_CUSTOMER));

            if (customers.Count == 0)
                return false;

            foreach (var elm in customers)
            {
                if (!elm.GetAttribute("title").Equals(customerName))
                    return false;
            }

            return true;
        }


        public bool VerifyDateFrom(DateTime value)
        {
            var dateFrom = _webDriver.FindElements(By.XPath(DATE)).FirstOrDefault();

            if (dateFrom != null)
            {
                // On modifie le format de la date pour qu'elle soit compatible avec un format universel
                // entrée : 07-Mar (Thu)
                // sortie : {7,Mar,Thu}
                string[] date = FormatDate(dateFrom.Text);

                // Tuesday, March 5, 2024
                var formatDate = value.ToString("D", CultureInfo.CreateSpecificCulture("en-US"));

                if (formatDate.Contains(date[0]) && formatDate.Contains(date[1]) && formatDate.Contains(date[2]))
                {
                    return true;
                }

            }
            return false;
        }
        public bool VerifyDateTo(DateTime value)
        {
            var dateFrom = _webDriver.FindElements(By.XPath(DATE)).LastOrDefault();

            if (dateFrom != null)
            {
                // On modifie le format de la date pour qu'elle soit compatible avec un format universel
                // "16-Dec (Sat)"
                //string[] date = FormatDate(dateFrom.Text);
                string[] date = FormatDate(dateFrom.Text);

                var formatDate = value.ToString("D", CultureInfo.CreateSpecificCulture("en-US"));

                if (formatDate.Contains(date[0]) && formatDate.Contains(date[1]) && formatDate.Contains(date[2]))
                {
                    return true;
                }

            }
            return false;
        }

        public bool IsWithoutDatasheet()
        {
            var elements = _webDriver.FindElements(By.XPath(ICON_DATASHEET));

            if (elements.Count == 0)
                return false;

            foreach (var element in elements)
            {
                string classDatasheet = element.GetAttribute("class");

                if (!classDatasheet.Equals("btn-icon-recipe btn-icon Warning"))
                    return false;
            }

            return true;
        }

        public string GetItemSubGroup()
        {
            Actions action = new Actions(_webDriver);

            _itemSubGroupFilter = WaitForElementExists(By.Id(ITEM_SUB_GROUP_FILTER));
            action.MoveToElement(_itemSubGroupFilter).Perform();
            _itemSubGroupFilter.Click();

            var itemSubGroup = _webDriver.FindElement(By.XPath(ITEM_SUB_GROUP_SELECTED));
            return itemSubGroup.Text;
        }
        public string GetNameItemSubGroup()
        {
            WaitForLoad();
            var resultSubItem = WaitForElementIsVisible(By.XPath("/html/body/div[17]/ul"));
            return resultSubItem.Text;
        }


        //___________________________________ Méthodes _________________________________________

        // General
        public override void ShowExtendedMenu()
        {
            Actions actions = new Actions(_webDriver);
            _extendedMenu = WaitForElementExists(By.XPath(EXTENDED_MENU));
            actions.MoveToElement(_extendedMenu).Perform();
        }

        public void ResetAllQuantities()
        {
            ShowExtendedMenu();

            _resetQuantities = WaitForElementIsVisible(By.Id(RESET_QUANTITIES));
            _resetQuantities.Click();
            WaitPageLoading();
            WaitForLoad();
        }

        public void RazAllQuantities()
        {
            ShowExtendedMenu();

            _razQuantities = WaitForElementIsVisible(By.Id(RAZ_QUANTITIES));
            _razQuantities.Click();
            WaitForLoad();

            Thread.Sleep(1000);
        }

        public void ClickOnPreviousDate()
        {
            _previousDate = WaitForElementIsVisible(By.Id(PREVIOUS_DATE));
            _previousDate.Click();
            WaitPageLoading();
        }

        public void ClickOnNextDate()
        {
            _nextDate = WaitForElementIsVisible(By.Id(NEXT_DATE));
            _nextDate.Click();
            WaitPageLoading();
        }

        // Onglets
        public ResultPage GoToResultPage()
        {
            _resultTab = WaitForElementIsVisible(By.XPath(RESULT_TAB));
            _resultTab.Click();
            WaitForLoad();
            WaitPageLoading();
            WaitForLoad();
            return new ResultPage(_webDriver, _testContext);
        }

        // Tableau

        public int CountServices()
        {
            WaitForLoad();
            return _webDriver.FindElements(By.XPath(SERVICES)).Count;
        }

        public void SelectService(string serviceName)
        {
            _services = WaitForElementIsVisible(By.XPath(String.Format(SERVICE_QUANTITY, serviceName)));
            _services.Click();
            WaitForLoad();
            WaitPageLoading();
        }

        public string getTotalService(string serviceName)
        {

            _totalService = WaitForElementIsVisible(By.XPath(String.Format(TOTAL_SERVICE, serviceName)));
            string Total = _totalService.Text.Trim();
            WaitForLoad();
            return Total;
        }



        public void ClickFirstRow()
        {
            IWebElement firstRow;
            if (IsDev())
            {
                firstRow = WaitForElementIsVisible(By.Id("inflightrawmaterials-production-service-row-detail-1"));
            }
            else
            {
                firstRow = WaitForElementIsVisible(By.XPath("//*[@id=\"list-item-with-action\"]/div/table/tbody/tr[1]"));
            }
            new Actions(_webDriver).MoveToElement(firstRow).Perform();
            // icone au milieu de la zone de texte "nom du service"
            //_firstService = WaitForElementIsVisible(By.XPath(FIRST_SERVICE));
            //if (IsDev())
            //{
            _firstService = WaitForElementIsVisible(By.Id("inflightrawmaterials-production-service-detail-1"));
            _firstService.Click();
            WaitForLoad();
            //}
            //else
            //{
            //    _firstService = WaitForElementIsVisible(By.XPath("//*[@id=\"list-item-with-action\"]/div/table/tbody/tr/td[2]"));
            //    _firstService.Click();
            //    WaitForLoad();
            //}

        }
        public ServicePricePage EditFirstService()
        {
            Actions actions = new Actions(_webDriver);
            ClickFirstRow();
            WaitPageLoading();

            return new ServicePricePage(_webDriver, _testContext);
        }

        public void ConsultService(string service_name)
        {
            SelectService(service_name);
            _serviceConsultFlesh = WaitForElementIsVisible(By.XPath(SERVICE_CONSULT_FLESH));
            _serviceConsultFlesh.Click();
            WaitForLoad();
        }

        public string GetServiceName()
        {
            _serviceName = WaitForElementIsVisible(By.Id(SERVICE_NAME));
            return _serviceName.Text.Trim();
        }

        public bool IsServicePresent(string serviceName)
        {
            WaitPageLoading();
            WaitForLoad();
            ReadOnlyCollection<IWebElement> elements = null;
            if (IsDev()) elements = _webDriver.FindElements(By.Id(SERVICE_NAME));
            else elements = _webDriver.FindElements(By.XPath("//*[@id='list-item-with-action']/div/table/tbody/tr[*]/td[2]"));
            WaitForLoad();
            if (elements.Count == 0) return false;
            foreach (var element in elements)
                if (element.Text.Trim().Equals(serviceName))
                    return true;
            WaitPageLoading();
            WaitForLoad(); 
            return false;
        }

        public void SetQuantity(string value, string serviceName)
        {

            _inputQty = WaitForElementIsVisible(By.XPath(String.Format(QUANTITY_INPUT, serviceName)));
            _inputQty.SetValue(ControlType.TextBox, value);
            WaitForLoad();
            WaitPageLoading();
        }

        public void SetQuantityWithSelect (string value, string serviceName)
        {
            _services = WaitForElementIsVisible(By.XPath(String.Format(SERVICE_QUANTITY, serviceName)));
            _services.Click();
            WaitForLoad();
            WaitPageLoading();
            Actions action = new Actions(_webDriver);

            _inputQty = WaitForElementIsVisible(By.XPath(String.Format(QUANTITY_INPUT, serviceName)));
            action.MoveToElement(_inputQty).Perform();
            _inputQty.SetValue(ControlType.TextBox, value);
            WaitForLoad();
            WaitPageLoading();
        }


        public void SetQuantityAfterConsult(string value)
        {
            var lineselected = WaitForElementIsVisible(By.XPath(SELECT_SERVICE_LINE));
            lineselected.Click();
            _inputQty = WaitForElementIsVisible(By.XPath(SERVICE_QUANTITY_LINE));
            WaitForLoad();
            _inputQty.Click();
            _inputQty.SetValue(ControlType.TextBox, value);
            WaitForLoad();
            WaitPageLoading();
        }

        public bool IsQuantityUpdatedAfterConsult(string value)
        {
            try
            {
                _inputQty = _webDriver.FindElement(By.XPath(String.Format(SERVICE_QUANTITY_LINE)));
                var x = _inputQty.GetAttribute("value");
                var y = _inputQty.GetCssValue("background-color");
                if (_inputQty.GetAttribute("value").Equals(value) && _inputQty.GetCssValue("background-color").Equals("rgba(135, 206, 235, 1)"))
                {
                    return true;
                }
            }
            catch
            {
                return false;
            }

            return false;
        }

        public double GetQuantity(string decimalSeparator, string serviceName)
        {
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");
            WaitPageLoading(); 
            _quantity = WaitForElementIsVisible(By.XPath(String.Format(SERVICE_QUANTITY, serviceName)));
            return double.Parse(_quantity.Text, ci);
        }

        public bool IsQuantityUpdated(string value, string serviceName)
        {
            try
            {
                _inputQty = _webDriver.FindElement(By.XPath(String.Format(QUANTITY_INPUT, serviceName)));

                if (_inputQty.GetAttribute("value").Equals(value) && _inputQty.GetCssValue("background-color").Equals("rgba(143, 188, 143, 1)"))
                {
                    return true;
                }
            }
            catch
            {
                return false;
            }

            return false;
        }

        public List<string> GetCustomerList()
        {
            HashSet<string> customerSet = new HashSet<string>();

            var customers = _webDriver.FindElements(By.XPath(SERVICE_CUSTOMER));

            foreach (var elm in customers)
            {
                customerSet.Add(elm.GetAttribute("title").Trim());
            }

            return customerSet.ToList();
        }

        private string[] FormatDate(string date)
        {
            // entrée : 07-Mar (Thu)
            // sortie : {7,Mar,Thu}
            string newDateFormat = date.Replace("(", "").Replace(")", "").Replace(" ", "-");
            string[] infoDate = newDateFormat.Split('-');

            if (infoDate[0] != null && infoDate[0].StartsWith("0"))
            {
                infoDate[0] = infoDate[0].Remove(0, 1);
            }

            return infoDate;
        }
        public void ClickTabQtyAdj()
        {
            var qtyAdjTab = WaitForElementIsVisible(By.Id("hrefTabContentItemContainer"));
            qtyAdjTab.Click();
            WaitForLoad();
        }
        private string GetResFiltredTabServices()
        {
            var tabFiltredServices = WaitForElementExists(By.XPath(FILTRED_TAB_SERVICES));
            return tabFiltredServices.Text;
        }
        public bool IsFilterServiceApplied(string filtredService)
        {
            string resultFiltred = GetResFiltredTabServices();

            return filtredService == resultFiltred;
        }

        public void OpenNewQtyDetailsPopUp(string serviceName = null)
        {
            if (!string.IsNullOrEmpty(serviceName))
            {
                _services = WaitForElementIsVisible(By.XPath(String.Format(SERVICE_QUANTITY, serviceName)));
                _services.Click();
                IWebElement screenIcon = WaitForElementIsVisible(By.XPath(String.Format(POPUP_ICON2, serviceName)));
                screenIcon.Click();
                WaitForLoad();
            }
            else
            {

                IWebElement screenIcon = WaitForElementIsVisible(By.XPath(POPUP_ICON));
                screenIcon.Click();
                WaitForLoad();

            }

        }
        public void OpenFirstQtyDetailsPopUp()
        {
                _services = WaitForElementIsVisible(By.Id(FIRST_QUANTITY));
                _services.Click();
                IWebElement screenIcon = WaitForElementIsVisible(By.Id(FIRST_ICON_POP_UP));
                screenIcon.Click();
                WaitForLoad();
        }
        public bool AreAllFlightsDifferent()
        {
            string lastflight = "";
            HashSet<string> volListe = new HashSet<string>();

            var volElements = _webDriver.FindElements(By.XPath(LISTE_VOL_QTYDETAILS));
            WaitForLoad();
            for (var i = 0; i < volElements.Count; i++)
            {
                var vols = _webDriver.FindElements(By.XPath(LISTE_VOL_QTYDETAILS));
                WaitForLoad();
                vols[i].Click();
                WaitForLoad();
                var flightElement = WaitForElementExists(By.XPath(FLIGHT));
                if (flightElement != null)
                {
                    var currentFlight = flightElement.Text;
                    if (!currentFlight.Equals(lastflight))
                    {
                        volListe.Add(currentFlight);
                        lastflight = currentFlight;
                    }

                }

            }

            return volElements.Count == volListe.Count;
        }
        public string GetSelectedGroupBy()
        {
            var selectedGroup = _webDriver.FindElement(By.XPath(SELECTED_GROUPBY));
            return selectedGroup.GetAttribute("value");
        }
        public string ClickRefreshAndGetTheQty()
        {
            Actions action = new Actions(_webDriver);

            var flightName = WaitForElementIsVisible(By.XPath(SELECT_SERVICE_LINE));
            flightName.Click();
            _inputQty = WaitForElementIsVisible(By.XPath(QUANTITY_INPUT_WIDHOUT_SERVICE));
            _inputQty.Click();

            var refresh = WaitForElementIsVisible(By.XPath(REFRESH_SERVICE_BTN));
            action.MoveToElement(refresh).Perform();
            refresh.Click();

            return _inputQty.GetAttribute("value");
        }
        public string GetStartTime()
        {
            var hour = WaitForElementExists(By.XPath(PRODMAN_FILTER_START_TIME_HOUR)).GetAttribute("value");
            var minute = WaitForElementExists(By.XPath(PRODMAN_FILTER_START_TIME_MIN)).GetAttribute("value");
            return $"{hour}:{minute}";

        }
        public string GetEndTime()
        {
            var hour = WaitForElementExists(By.XPath(PRODMAN_FILTER_END_TIME_HOUR)).GetAttribute("value");
            var minute = WaitForElementExists(By.XPath(PRODMAN_FILTER_END_TIME_MIN)).GetAttribute("value");
            return $"{hour}:{minute}";

        }
        public string GetFirstFlightNo()
        {
            string lastflight = "";
            HashSet<string> volListe = new HashSet<string>();

            var volElements = _webDriver.FindElements(By.XPath(LISTE_VOL_QTYDETAILS));
            WaitForLoad();
            for (var i = 0; i < volElements.Count; i++)
            {
                var vols = _webDriver.FindElements(By.XPath(LISTE_VOL_QTYDETAILS));
                WaitForLoad();
                vols[i].Click();
                WaitForLoad();
                var flightElement = WaitForElementExists(By.XPath(FLIGHT));
                if (flightElement != null)
                {
                    var currentFlight = flightElement.Text;
                    if (!currentFlight.Equals(lastflight))
                    {
                        volListe.Add(currentFlight);
                        lastflight = currentFlight;
                    }

                }

            }

            return lastflight;
        }
        public void ClickOnPrevious()
        {
            _previous = WaitForElementIsVisible(By.Id(PREVIOUS));
            _previous.Click();
            WaitPageLoading();
        }
        public bool IsNextAvailable()
        {
            if (IsDev()) return isElementExists(By.Id(NEXT));
            else return isElementExists(By.Id("inflightrawmaterials-quantityadjustments-detailpopupcontent-nextcustomer"));
        }
        public void ClickOnNext()
        {
            if (IsDev()) _next = WaitForElementIsVisible(By.Id(NEXT));
            else _next = WaitForElementIsVisible(By.XPath("//*[@id=\"qtyAdjustmentsDetailBody\"]/div[1]/table/tbody/tr/td[3]/a"));
            _next.Click();
            WaitPageLoading();
        }
        public ServicePricePage ClickFirstRowGoPrice()
        {
            var actions = new Actions(_webDriver);
            _firstService = WaitForElementIsVisible(By.XPath(FIRST_SERVICE));
            actions.MoveToElement(_firstService).Perform();

            _firstService.Click();

            return new ServicePricePage(_webDriver, _testContext);
        }

        public string GetFirstServiceName()
        {
            WaitForLoad();
            _editFirstService = WaitForElementIsVisible(By.Id("inflightrawmaterials-production-servicename-detail-1"));
            return _editFirstService.Text;
        }
        public string GetFlightNoWithLignNumber(int index)
        {
            string flightNo = string.Empty;
            var vols = _webDriver.FindElements(By.XPath(LISTE_VOL_QTYDETAILS));

            WaitForLoad();
            vols[index].Click();

            WaitForLoad();
            var flightElement = WaitForElementExists(By.XPath(FLIGHT));

            if (flightElement != null)
            {
                flightNo = flightElement.Text.Trim();
            }

            return flightNo;
        }
        public string GetFlightFromPopUp()
        {
            string flightNo = string.Empty;

            WaitForLoad();
            var flightElement = WaitForElementExists(By.XPath("//*[@id=\"list-item-with-action\"]/div/table/tbody/tr/td[1]"));

            if (flightElement != null)
            {
                flightNo = flightElement.Text.Trim();
            }

            return flightNo;
        }
        public IEnumerable<string> GetAllServiceName()
        {

            var serviceElements = _webDriver.FindElements(By.Id(SERVICE_NAME));
            return serviceElements.Select(el => el.Text.Trim());
        }
        public void DisplayRow(int rowNumber)
        {
            Actions actions = new Actions(_webDriver);
            var row = WaitForElementExists(By.XPath(string.Format(SERVICE_ROW, rowNumber)));
            actions.MoveToElement(row).Perform();
            WaitForLoad();

        }
        public bool VerifyQtyTotal(int Total)
        {
            /*
             * WARNING : This method only handles the calculation for the first page and does not account for scenarios with multiple pages.
             * TODO:
             * 1- Retrieve the Total Page Count: Determine how many pages are available by checking the API response or data source for pagination details.
             * 2- Iterate Through All Pages: Create a loop to iterate through each page, retrieving and processing the data on each one.
             * 3- Accumulate the Results: Instead of calculating the total for just the first page, aggregate the totals from each page as you iterate.
             */
            int somme = 0;
            ReadOnlyCollection<IWebElement> ColonneTotal = _webDriver.FindElements(By.XPath(COLTOTAL));
            if (ColonneTotal.Count == 0)
                return false;
            foreach (var c in ColonneTotal)
                if (!String.IsNullOrEmpty(c.Text))
                {
                    int valeur = int.Parse(c.Text);
                    somme += valeur;
                }
            return Total == somme;
        }

        public bool VerifierPictoVertExist()
        {
            var pictovert = _webDriver.FindElements(By.XPath(NOMBRE_PICTO_VERT));

            if (pictovert.Count == 0)
                return false;

            return true;
        }
        public string GetFirstServiceNameWithDataSheet()
        {
            var servicenamewithdatasheet = _webDriver.FindElements(By.XPath(NOMBRE_PICTO_VERT)).FirstOrDefault();
            if (servicenamewithdatasheet != null)
            {
                var nextTd = servicenamewithdatasheet.FindElement(By.XPath("./ancestor::td/following-sibling::td[1]"));

                return nextTd.Text;

            }

            return "";
        }
        public DatasheetPage ClickFirstPictoVert()
        {
            var pictovert = _webDriver.FindElements(By.XPath(NOMBRE_PICTO_VERT)).FirstOrDefault();
            if (pictovert != null)
            {
                pictovert.Click();
            }
            return new DatasheetPage(_webDriver, _testContext);
        }

        public string GetClorFirstServiceNameWithDataSheet()
        {
            var servicenamewithdatasheet = _webDriver.FindElements(By.XPath(NOMBRE_PICTO_VERT)).FirstOrDefault();
            if (servicenamewithdatasheet != null)
            {
                IJavaScriptExecutor js = (IJavaScriptExecutor)_webDriver;
                string color = (string)js.ExecuteScript("return window.getComputedStyle(arguments[0], null).getPropertyValue('color');", servicenamewithdatasheet);
                return color;
            }
            return "";
        }
        public List<string> GetServicesName()
        {

            var serviceElements = _webDriver.FindElements(By.XPath("//*[@id=\"list-item-with-action\"]/div/table/tbody/tr[*]/td[2]"));
            return serviceElements.Select(el => el.Text.Trim()).ToList();
        }
        public ServicePricePage EditService(string serviceName)
        {
            var editServiceIcon = _webDriver.FindElement(By.XPath(string.Format("//*[@id=\"list-item-with-action\"]/div/table/tbody/tr[*]/td[contains(text(),'{0}')]/./a", serviceName)));
            Actions actions = new Actions(_webDriver);
            actions.MoveToElement(editServiceIcon).Click().Perform();
            WaitPageLoading();

            return new ServicePricePage(_webDriver, _testContext);
        }
       public String  GetFirstServiceICAO()
        {
            var customerICAO  = WaitForElementIsVisible(By.Id("inflightrawmaterials-production-service-detail-customercode-1"));
            return customerICAO.Text ; 
        }
        public string GetCustomerFilterNumber ()
        {
            _customerFilter = WaitForElementIsVisible(By.Id(CUSTOMER_FILTER));

            return _customerFilter.Text;
        }
        public bool VerifyIfNewWindowIsOpened()
        {

            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            if (_webDriver.WindowHandles.Count == 1)
            {
                return false;
            }
            return true;
        }
        public void  ReturnTothePastQty()
        {
            _customerFilter = WaitForElementIsVisible(By.XPath("//*[@id=\"production-inflightrawmaterials-reset-flightquantity-0-1\"]/span"));
            _customerFilter.Click(); 
        }
        public bool IsQtyElementAdjusted()
        {
            var qty_elements = _webDriver.FindElements(By.Id("Details_Items_0__Details_0__AdjustedTotalQuantity"));
            if (qty_elements.Count > 0)
            {
                var classDatasheet = qty_elements[0].GetAttribute("class");
                return classDatasheet.Contains("adjusted"); 
            }
            
                return false; 
        }
    
    }
}
