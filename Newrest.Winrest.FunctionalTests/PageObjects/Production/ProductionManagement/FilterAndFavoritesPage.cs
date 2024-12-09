using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
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
using System.Text.RegularExpressions;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Production.ProductionManagement
{
    public class FilterAndFavoritesPage : PageBase
    {

        public FilterAndFavoritesPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //__________________________________Constantes_____________________________________

        private const string RESET_FILTER = "//*[@id=\"production-filter-form\"]/div/div[1]/a";

        private const string SITE = "SiteId";
        private const string FILTER_DATE_FROM = "ProdDateFrom";
        private const string FILTER_DATE_TO = "ProdDateTo";

        private const string FILTER_START_TIME_HOUR = "//div[6]/div/input[2]";
        private const string FILTER_START_TIME_MIN = "//div[6]/div/input[3]";
        private const string FILTER_END_TIME_HOUR = "//div[7]/div/input[2]";
        private const string FILTER_END_TIME_MIN = "//div[7]/div/input[3]";

        private const string WORKSHOP_FILTER = "SelectedWorkshops_ms";
        private const string UNSELECT_WORKSHOP = "/html/body/div[10]/div/ul/li[2]/a";
        private const string SEARCH_WORKSHOP = "/html/body/div[10]/div/div/label/input";

        private const string CUSTOMER_FILTER = "SelectedCustomers_ms";
        private const string UNSELECT_CUSTOMER = "/html/body/div[11]/div/ul/li[2]/a";
        private const string SEARCH_CUSTOMER = "/html/body/div[11]/div/div/label/input";

        private const string GUEST_TYPE_FILTER = "SelectedGuestTypes_ms";
        private const string UNSELECT_GUEST_TYPE = "/html/body/div[12]/div/ul/li[2]/a";
        private const string SEARCH_GUEST_TYPE = "/html/body/div[12]/div/div/label/input";

        private const string SERVICE_CATEGORIE_FILTER = "SelectedServiceCategories_ms";
        private const string UNSELECT_CATEGORIE_SERVICE = "/html/body/div[13]/div/ul/li[2]/a";
        private const string SEARCH_CATEGORIE_SERVICE = "/html/body/div[13]/div/div/label/input";

        private const string SERVICE_FILTER = "SelectedServices_ms";
        private const string UNSELECT_SERVICE = "/html/body/div[14]/div/ul/li[2]/a";
        private const string SEARCH_SERVICE = "/html/body/div[14]/div/div/label/input";
        private const string SERVICE_SELECTED = "//*[@id=\"SelectedItemGroups_ms\"]/span[2]";

        private const string RECIPE_TYPE_FILTER = "SelectedRecipeTypes_ms";
        private const string UNSELECT_RECIPE_TYPE = "/html/body/div[15]/div/ul/li[2]/a";
        private const string SEARCH_RECIPE_TYPE = "/html/body/div[15]/div/div/label/input";

        private const string ITEM_GROUP_FILTER = "SelectedItemGroups_ms";
        private const string UNSELECT_ITEM_GROUP = "/html/body/div[16]/div/ul/li[2]/a";
        private const string SEARCH_ITEM_GROUP = "/html/body/div[16]/div/div/label/input";
        private const string ITEM_GROUP_SELECTED = "//*[@id=\"SelectedItemGroups_ms\"]/span[2]";
        private const string ITEM_GROUP_SELECTION = "//*[@id=\"SelectedItemGroups_ms\"]/span[contains(text(), 'selected')]";

        private const string QTY_ADJUST = "Mode_QtyAdjustments";
        private const string RAW_MATERIAL_BY_GROUP = "Mode_RawMaterials";
        private const string RAW_MATERIAL_BY_WORKSHOP = "Mode_RawMaterialsWorkshop";
        private const string RAW_MATERIAL_BY_SUPPLIER = "Mode_RawMaterialsSupplier";
        private const string RAW_MATERIAL_BY_RECIPE = "Mode_RawMaterialsRecipe";
        private const string RAW_MATERIAL_BY_CUSTOMER = "Mode_RawMaterialsCustomer";

        private const string SHOW_NORMAL_AND_VACUUM_PROD = "includeVacuumPackedAll";
        private const string SHOW_VACUUM_PROD = "includeVacuumPackedOnly";
        private const string SHOW_NORMAL_PROD = "IncludeNovacuumPacked";

        private const string GROUP_BY = "SelectedGroupBy";

        private const string FAVORITE_APPLY = "//*[@id=\"favorite-filter-form\"]/div[2]/div/div/div[1]/a[2]/span";
        private const string FAVORITE_EDIT_FAV = "//*[@id=\"favorite-filter-form\"]/div[2]/div/div/div[1]/a[1]/span";
        private const string FAVORITE_SEARCH = "//*[@id=\"searchPattern\"]";
        private const string FAVORITE = "//*[@id=\"favorite-filter-form\"]/div[2]/div/div/div[@class='raw-favorite clickable-favorite']/span[text()='{0}']";
        private const string FAVORITE_NAME = "Name";
        private const string SAVE_FAVORITE = "last";
        private const string DELETE_FAVORITE = "//*[@id=\"favorite-filter-form\"]/div[2]/div/div/div[@class='raw-favorite clickable-favorite']/span[text()='{0}']/../button/span";
        private const string DELETE_FAVORITECROSS = "//*[@id=\"favorite-filter-form\"]/div[2]/div/div/div[@class='raw-favorite clickable-favorite']/span[text()='{0}']/../button";
        private const string CONFIRM_DELETE_FAVORITE = "dataConfirmOK";

        private const string DONE = "//*[@id=\"btnFilter\"]/input";
        private const string MAKE_FAVORITE = "//*[@id=\"btnFilter\"]/div/a[2]";

        private const string ITEM_SUB_GROUP_FILTER = "SelectedItemSubGroups_ms";
        private const string UNSELECT_ITEM_SUB_GROUP = "/html/body/div[117]/div/ul/li[2]/a";
        private const string SEARCH_ITEM_SUB_GROUP = "/html/body/div[17]/div/div/label/input";
        private const string ITEM_SUB_GROUP_SELECTED = "//*[@id=\"SelectedItemSubGroups\"]/span[2]";
        private const string SELECTED_SITE_ID = "SelectedSiteId";
        private const string DATE_FROM = "datapicker-new-supply-order-start";
        private const string DATE_TO = "datapicker-new-supply-order-end";
        private const string DELIVERY_DATE = "datapicker-new-supply-order-delivery";
        private const string DELIVERY_LOCATION = "SelectedSitePlaceId";
        private const string ACTIVATED = "SupplyOrder_IsActive";
        private const string SUBMIT = "btn-submit-form-create-supply-order";
        private const string PERFILL_QUANTITIES_FROM_PRODUCTION_MANAGEMENT = "//*[@id=\"checkBoxPrefillRawMaterials\"]";
        private const string NEW_SUPPLY_ORDER = "GenerateSupplyOrderBtn";
        private const string VALIDATE = "btn-validate-supply-order";
        private const string VALIDATION_BUTTON = "//*[@id=\"div-body\"]/div/div[1]/div/div[2]/button";
        private const string MESSAGE_VALIDATE = "tb-new-supplyorder-number";
        private const string FOOTER_PACKETS = "SelectedFoodPackets_ms";
        private const string FAVORITE_FILTER = "//*[@id=\"favorite-filter-form\"]/div[2]/div/div/div/span";
        private const string EXISTANCEFAVORITE = "//*[@id=\"favorite-filter-form\"]/div[2]/div/div/div[1]";
        private const string EDITBUTTON = "//*[@id=\"favorite-filter-form\"]/div[2]/div/div/div[1]/a[1]";
        private const string EDITINPUT = "//*[@id=\"Name\"]";
        private const string MAKEFAVORITE = "//*[@id=\"addOrEditFavoriteBtn\"]";
        private const string UNCHECK_ALL = "/html/body/div[17]/div/ul/li[2]/a/span[1]";
        //__________________________________Variables______________________________________
        [FindsBy(How = How.XPath, Using = EDITBUTTON)]
        private IWebElement _editbutton;
        [FindsBy(How = How.XPath, Using = MAKEFAVORITE)]
        private IWebElement _makefavorite;
        [FindsBy(How = How.XPath, Using = RESET_FILTER)]
        private IWebElement _resetFilter;

        [FindsBy(How = How.Id, Using = SITE)]
        private IWebElement _site;

        [FindsBy(How = How.Id, Using = FILTER_DATE_FROM)]
        public IWebElement _dateFrom;

        [FindsBy(How = How.Id, Using = FILTER_DATE_TO)]
        public IWebElement _dateTo;

        [FindsBy(How = How.Id, Using = FILTER_START_TIME_HOUR)]
        public IWebElement _filterStartTimeHour;

        [FindsBy(How = How.Id, Using = FILTER_START_TIME_MIN)]
        public IWebElement _filterStartTimeMin;

        [FindsBy(How = How.Id, Using = FILTER_END_TIME_HOUR)]
        public IWebElement _filterEndTimeHour;

        [FindsBy(How = How.Id, Using = FILTER_END_TIME_MIN)]
        public IWebElement _filterEndTimeMin;


        [FindsBy(How = How.XPath, Using = WORKSHOP_FILTER)]
        private IWebElement _workShopFilter;

        [FindsBy(How = How.XPath, Using = UNSELECT_WORKSHOP)]
        private IWebElement _unselectAllWorkShop;

        [FindsBy(How = How.XPath, Using = SEARCH_WORKSHOP)]
        private IWebElement _searchWorkShop;

        [FindsBy(How = How.XPath, Using = CUSTOMER_FILTER)]
        private IWebElement _customerFilter;

        [FindsBy(How = How.XPath, Using = UNSELECT_CUSTOMER)]
        private IWebElement _unselectAllCust;

        [FindsBy(How = How.XPath, Using = SEARCH_CUSTOMER)]
        private IWebElement _searchCustomer;

        [FindsBy(How = How.XPath, Using = GUEST_TYPE_FILTER)]
        private IWebElement _guestFilter;

        [FindsBy(How = How.XPath, Using = UNSELECT_GUEST_TYPE)]
        private IWebElement _unselectAllGuest;

        [FindsBy(How = How.XPath, Using = SEARCH_GUEST_TYPE)]
        private IWebElement _searchGuest;

        [FindsBy(How = How.XPath, Using = SERVICE_CATEGORIE_FILTER)]
        private IWebElement _serviceCategorieFilter;

        [FindsBy(How = How.XPath, Using = UNSELECT_CATEGORIE_SERVICE)]
        private IWebElement _unselectAllServiceCategorie;

        [FindsBy(How = How.XPath, Using = SEARCH_CATEGORIE_SERVICE)]
        private IWebElement _searchCategorieService;

        [FindsBy(How = How.XPath, Using = SERVICE_FILTER)]
        private IWebElement _serviceFilter;

        [FindsBy(How = How.XPath, Using = UNSELECT_SERVICE)]
        private IWebElement _unselectAllService;

        [FindsBy(How = How.XPath, Using = SEARCH_SERVICE)]
        private IWebElement _searchService;

        [FindsBy(How = How.XPath, Using = RECIPE_TYPE_FILTER)]
        private IWebElement _recipeTypeFilter;

        [FindsBy(How = How.XPath, Using = UNSELECT_RECIPE_TYPE)]
        private IWebElement _unselectAllRecipeType;

        [FindsBy(How = How.XPath, Using = SEARCH_RECIPE_TYPE)]
        private IWebElement _searchRecipeType;

        [FindsBy(How = How.XPath, Using = ITEM_GROUP_FILTER)]
        private IWebElement _itemGroupFilter;

        [FindsBy(How = How.XPath, Using = UNSELECT_ITEM_GROUP)]
        private IWebElement _unselectAllItemGroup;

        [FindsBy(How = How.XPath, Using = SEARCH_ITEM_GROUP)]
        private IWebElement _searchItemGroup;

        [FindsBy(How = How.XPath, Using = ITEM_GROUP_SELECTION)]
        private IWebElement _itemGroupsSelection;

        [FindsBy(How = How.Id, Using = QTY_ADJUST)]
        private IWebElement _qtyAdjustements;

        [FindsBy(How = How.Id, Using = RAW_MATERIAL_BY_GROUP)]
        private IWebElement _rawMaterialByGroup;

        [FindsBy(How = How.Id, Using = RAW_MATERIAL_BY_WORKSHOP)]
        private IWebElement _rawMaterialByWorkshop;

        [FindsBy(How = How.Id, Using = RAW_MATERIAL_BY_SUPPLIER)]
        private IWebElement _rawMaterialBySupplier;

        [FindsBy(How = How.Id, Using = RAW_MATERIAL_BY_RECIPE)]
        private IWebElement _rawMaterialByRecipe;

        [FindsBy(How = How.Id, Using = RAW_MATERIAL_BY_CUSTOMER)]
        private IWebElement _rawMaterialByCustomer;

        [FindsBy(How = How.Id, Using = SHOW_NORMAL_AND_VACUUM_PROD)]
        private IWebElement _showNormalAndVacuumProd;

        [FindsBy(How = How.Id, Using = SHOW_VACUUM_PROD)]
        private IWebElement _showVacuumProd;

        [FindsBy(How = How.Id, Using = SHOW_NORMAL_PROD)]
        private IWebElement _showNormalProd;

        [FindsBy(How = How.Id, Using = GROUP_BY)]
        private IWebElement _groupBy;

        [FindsBy(How = How.XPath, Using = DONE)]
        private IWebElement _done;

        [FindsBy(How = How.XPath, Using = FAVORITE_SEARCH)]
        private IWebElement _favorite_search; 

        [FindsBy(How = How.XPath, Using = FAVORITE_APPLY)]
        private IWebElement _favorite_apply;

        [FindsBy(How = How.XPath, Using = FAVORITE_EDIT_FAV)]
        private IWebElement _favorite_edit;

        [FindsBy(How = How.XPath, Using = FAVORITE)]
        private IWebElement _favorite;

        [FindsBy(How = How.XPath, Using = DELETE_FAVORITE)]
        private IWebElement _deleteFavorite;

        [FindsBy(How = How.Id, Using = CONFIRM_DELETE_FAVORITE)]
        private IWebElement _confirmDeleteFavorite;

        [FindsBy(How = How.Id, Using = FAVORITE_NAME)]
        private IWebElement _favoriteName;

        [FindsBy(How = How.Id, Using = SAVE_FAVORITE)]
        private IWebElement _saveFavorite;

        [FindsBy(How = How.XPath, Using = MAKE_FAVORITE)]
        private IWebElement _makeFavorite;

        [FindsBy(How = How.Id, Using = ITEM_SUB_GROUP_FILTER)]
        private IWebElement _itemSubGroup;

        [FindsBy(How = How.XPath, Using = UNSELECT_ITEM_SUB_GROUP)]
        private IWebElement _unselectAllItemSubGroup;

        [FindsBy(How = How.XPath, Using = SEARCH_ITEM_SUB_GROUP)]
        private IWebElement _searchItemSubGroup;
        [FindsBy(How = How.Id, Using = DATE_TO)]
        public IWebElement _dateto;
        [FindsBy(How = How.Id, Using = DATE_FROM)]
        public IWebElement _datefrom;
        [FindsBy(How = How.Id, Using = SELECTED_SITE_ID)]
        public IWebElement _sites;
        [FindsBy(How = How.Id, Using = DELIVERY_DATE)]
        public IWebElement _deliveryDate;
        [FindsBy(How = How.Id, Using = DELIVERY_LOCATION)]
        public IWebElement _deliveryLocation;
        [FindsBy(How = How.Id, Using = ACTIVATED)]
        public IWebElement _activated;
        [FindsBy(How = How.Id, Using = SUBMIT)]
        public IWebElement _submit;
        [FindsBy(How = How.XPath, Using = PERFILL_QUANTITIES_FROM_PRODUCTION_MANAGEMENT)]
        private IWebElement _perfillmanagement;
        [FindsBy(How = How.Id, Using = NEW_SUPPLY_ORDER)]
        private IWebElement _createNewSupplyOrder;
        [FindsBy(How = How.Id, Using = VALIDATE)]
        private IWebElement _validate;
        [FindsBy(How = How.XPath, Using = VALIDATION_BUTTON)]
        private IWebElement _validationButton;
        [FindsBy(How = How.XPath, Using = MESSAGE_VALIDATE)]
        private IWebElement _messagevalidate;
        [FindsBy(How = How.Id, Using = FOOTER_PACKETS)]
        private IWebElement _footerpackets;
        [FindsBy(How = How.XPath, Using = FAVORITE_FILTER)]
        private IWebElement _favoriteFilter;



        //___________________________________Pages_________________________________________
        public enum FilterType
        {
            Site,
            DateFrom,
            DateTo,
            StartTime,
            EndTime,
            Workshops,
            Customers,
            GuestType,
            ServicesCategorie,
            Service,
            RecipeType,
            ItemGroups,
            FoodPacket,
            QtyAdjustements,
            RawMaterialByGroup,
            RawMaterialByWorkshop,
            RawMaterialBySupplier,
            RawMaterialByRecipe,
            RawMaterialByCustomer,
            ShowNormalAndVacuumProd,
            ShowVacuumProd,
            ShowNormalProd,
            GroupBy,
            ItemSubGroups
        }

        public void Filter(FilterType filterType, object value)
        {
            switch (filterType)
            {

                case FilterType.Site:
                    _site = WaitForElementIsVisible(By.Id(SITE));
                    _site.SetValue(ControlType.DropDownList, value);
                    break;

                case FilterType.DateFrom:
                    WaitForElementIsVisible(By.Id(FILTER_DATE_FROM));
                    _dateFrom.SetValue(ControlType.DateTime, value);
                    _dateFrom.SendKeys(Keys.Tab);
                    break;

                case FilterType.DateTo:
                    WaitForElementIsVisible(By.Id(FILTER_DATE_TO));
                    _dateTo.SetValue(ControlType.DateTime, value);
                    _dateTo.SendKeys(Keys.Tab);
                    break;

                case FilterType.StartTime:
                    _filterStartTimeHour = WaitForElementIsVisible(By.XPath(FILTER_START_TIME_HOUR), nameof(FILTER_START_TIME_HOUR));
                    _filterStartTimeHour.ClearElement();
                    _filterStartTimeHour.SetValue(ControlType.TextBox, value.ToString().Substring(0, 2));
                    _filterStartTimeHour.SendKeys(Keys.Tab);

                    _filterStartTimeMin = WaitForElementIsVisible(By.XPath(FILTER_START_TIME_MIN), nameof(FILTER_START_TIME_MIN));
                    _filterStartTimeMin.ClearElement();
                    _filterStartTimeMin.SetValue(ControlType.TextBox, value.ToString().Substring(value.ToString().LastIndexOf(':') + 1));
                    _filterStartTimeMin.SendKeys(Keys.Tab);
                    break;

                case FilterType.EndTime:
                    _filterEndTimeHour = WaitForElementIsVisible(By.XPath(FILTER_END_TIME_HOUR), nameof(FILTER_END_TIME_HOUR));
                    _filterEndTimeHour.ClearElement();
                    _filterEndTimeHour.SetValue(ControlType.TextBox, value.ToString().Substring(0, 2));
                    _filterEndTimeHour.SendKeys(Keys.Tab);

                    _filterEndTimeMin = WaitForElementIsVisible(By.XPath(FILTER_END_TIME_MIN), nameof(FILTER_END_TIME_MIN));
                    _filterEndTimeMin.ClearElement();
                    _filterEndTimeMin.SetValue(ControlType.TextBox, value.ToString().Substring(value.ToString().LastIndexOf(':') + 1));
                    _filterEndTimeMin.SendKeys(Keys.Tab);
                    break;


                case FilterType.Workshops:
                    ComboBoxSelectById(new ComboBoxOptions(WORKSHOP_FILTER, (string)value, false));
                    break;

                case FilterType.Customers:
                    ComboBoxSelectById(new ComboBoxOptions(CUSTOMER_FILTER, (string)value, false));
                    break;

                case FilterType.GuestType:
                    ComboBoxSelectById(new ComboBoxOptions(GUEST_TYPE_FILTER, (string)value, false));
                    break;

                case FilterType.ServicesCategorie:
                    ComboBoxSelectById(new ComboBoxOptions(SERVICE_CATEGORIE_FILTER, (string)value, false));
                    break;


                case FilterType.Service:
                    ComboBoxSelectById(new ComboBoxOptions(SERVICE_FILTER, (string)value, false));
                    break;

                case FilterType.RecipeType:
                    ComboBoxSelectById(new ComboBoxOptions(RECIPE_TYPE_FILTER, (string)value, false));
                    break;

                case FilterType.ItemGroups:
                    ComboBoxSelectById(new ComboBoxOptions(ITEM_GROUP_FILTER, (string)value, false));
                    break;
                case FilterType.ItemSubGroups:
                    ComboBoxSelectById(new ComboBoxOptions(ITEM_SUB_GROUP_FILTER, (string)value, false));
                    break;
                case FilterType.FoodPacket:
                    ComboBoxSelectById(new ComboBoxOptions(FOOTER_PACKETS, (string)value, false));
                    break;

                case FilterType.QtyAdjustements:
                    _qtyAdjustements = WaitForElementExists(By.Id(QTY_ADJUST));
                    _qtyAdjustements.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterType.RawMaterialByGroup:
                    _rawMaterialByGroup = WaitForElementExists(By.Id(RAW_MATERIAL_BY_GROUP));
                    _rawMaterialByGroup.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterType.RawMaterialByWorkshop:
                    _rawMaterialByWorkshop = WaitForElementExists(By.Id(RAW_MATERIAL_BY_WORKSHOP));
                    _rawMaterialByWorkshop.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterType.RawMaterialBySupplier:
                    _rawMaterialBySupplier = WaitForElementExists(By.Id(RAW_MATERIAL_BY_SUPPLIER));
                    _rawMaterialBySupplier.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterType.RawMaterialByRecipe:
                    _rawMaterialByRecipe = WaitForElementExists(By.Id(RAW_MATERIAL_BY_RECIPE));
                    _rawMaterialByRecipe.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterType.RawMaterialByCustomer:
                    _rawMaterialByCustomer = WaitForElementExists(By.Id(RAW_MATERIAL_BY_CUSTOMER));
                    _rawMaterialByCustomer.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterType.ShowNormalAndVacuumProd:
                    _showNormalAndVacuumProd = WaitForElementExists(By.Id(SHOW_NORMAL_AND_VACUUM_PROD));
                    _showNormalAndVacuumProd.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterType.ShowVacuumProd:
                    _showVacuumProd = WaitForElementExists(By.Id(SHOW_VACUUM_PROD));
                    _showVacuumProd.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterType.ShowNormalProd:
                    _showNormalProd = WaitForElementExists(By.Id(SHOW_NORMAL_PROD));
                    _showNormalProd.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.GroupBy:
                    _groupBy = WaitForElementIsVisible(By.Id(GROUP_BY));
                    _groupBy.Click();
                    var element = WaitForElementIsVisible(By.XPath("//option[contains(@value,'" + value + "')]"));
                    _groupBy.SetValue(ControlType.DropDownList, element.Text);
                    _groupBy.Click();
                    break;

                default:
                    break;
            }

            Thread.Sleep(2500);
        }

        public void ResetFilter()
        {
            _resetFilter = WaitForElementIsVisible(By.XPath(RESET_FILTER));
            _resetFilter.Click();
            WaitForLoad();
            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                Filter(FilterType.DateTo, DateUtils.Now);
            }
        }

        public string CheckTAllItemGroupsSelected()
        {
            _itemGroupsSelection = WaitForElementIsVisible(By.XPath(ITEM_GROUP_SELECTION));
            Regex r = new Regex("([0-9]+) of ([0-9]+) Item Groups selected");
            Match m = r.Match(_itemGroupsSelection.Text);
            string select1 = m.Groups[1].Value;
            string select2 = m.Groups[2].Value;
            if (select1 == select2)
            {
                return "All";
            }
            else
            {
                return _itemGroupsSelection.Text;
            }
        }

        public string GetNbService()
        {
            var service = _webDriver.FindElement(By.XPath(SERVICE_SELECTED));
            return service.Text;
        }
        public bool ExistanceFavorite()
        {
             var elements = _webDriver.FindElements(By.XPath(EXISTANCEFAVORITE));
              return elements.Count > 0;
        }

        public bool EditInput()
        {
            var elements = _webDriver.FindElements(By.XPath(EDITINPUT));
            return elements.Count > 0;
        }

        public string GetNbItemGroup()
        {
            var itemGroup = _webDriver.FindElement(By.XPath(ITEM_GROUP_SELECTED));
            return itemGroup.Text;
        }

        public bool IsNormalProduction()
        {
            _showNormalProd = _webDriver.FindElement(By.Id(SHOW_NORMAL_PROD));

            try
            {
                string check = _showNormalProd.GetAttribute("checked");
                return check.Equals("true");
            }
            catch
            {
                return false;
            }
        }

        public bool IsRawMatByWorkshop()
        {
            _rawMaterialByWorkshop = _webDriver.FindElement(By.Id(RAW_MATERIAL_BY_WORKSHOP));

            try
            {
                string check = _rawMaterialByWorkshop.GetAttribute("checked");
                return check.Equals("true");
            }
            catch
            {
                return false;
            }
        }

        public QuantityAdjustmentsPage DoneToQtyAjustement()
        {
            _done = WaitForElementIsVisible(By.XPath(DONE));
            _done.Click();
            WaitForLoad();

            return new QuantityAdjustmentsPage(_webDriver, _testContext);
        }

        public void ClickEditButton()
        {
            var actions = new Actions(_webDriver);
            var elementToHover = WaitForElementIsVisible(By.XPath("//*[@id=\"favorite-filter-form\"]/div[2]/div/div/div[1]"));
            actions.MoveToElement(elementToHover).Perform();

            _editbutton = WaitForElementIsVisible(By.XPath("//*[@id=\"favorite-filter-form\"]/div[2]/div/div/div[1]/a[1]/span"));
            _editbutton.Click();
            WaitForLoad();

         }
        public void ClickMakeFavorite()
        {
            _makefavorite = WaitForElementIsVisible(By.XPath(MAKEFAVORITE));
            _makefavorite.Click();
            WaitForLoad();

        }

        public ResultPage DoneToResults()
        {
            _done = WaitForElementIsVisible(By.XPath(DONE));
            _done.Click();
            WaitForLoad();
            return new ResultPage(_webDriver, _testContext);
        }

        public void MakeFavorite(string favoriteName)
        {
            _makeFavorite = WaitForElementIsVisible(By.XPath(MAKE_FAVORITE));
            _makeFavorite.Click();
            WaitForLoad();

            _favoriteName = WaitForElementIsVisible(By.Id(FAVORITE_NAME));
            _favoriteName.SetValue(ControlType.TextBox, favoriteName);
            WaitForLoad();

            _saveFavorite = WaitForElementIsVisible(By.Id(SAVE_FAVORITE));
            _saveFavorite.Click();
            WaitPageLoading();//too long to save
            WaitForLoad();
        }

        public void DeleteFavorite(string favoriteName)
        {

            Actions action = new Actions(_webDriver);
            var favorite = _webDriver.FindElements(By.XPath(String.Format(FAVORITE, favoriteName))).Count;

            if (favorite > 0)
            {
                _favorite = WaitForElementIsVisible(By.XPath(String.Format(FAVORITE, favoriteName)));
                action.MoveToElement(_favorite).Perform();

                var deleteFavorite = WaitForElementIsVisible(By.XPath(String.Format(DELETE_FAVORITECROSS, favoriteName)));

                ((IJavaScriptExecutor)(IWebDriver)_webDriver).ExecuteScript(
                 "arguments[0].removeAttribute('class','class')", deleteFavorite);

                _deleteFavorite = WaitForElementIsVisible(By.XPath(String.Format(DELETE_FAVORITE, favoriteName)));
                _deleteFavorite.Click();
                WaitForLoad();

                _confirmDeleteFavorite = WaitForElementIsVisible(By.Id(CONFIRM_DELETE_FAVORITE));
                _confirmDeleteFavorite.Click();
                WaitForLoad();
            }
        }

        public void SelectFavorite(string favoriteName)
        {
            _favorite = WaitForElementIsVisible(By.XPath(String.Format(FAVORITE, favoriteName)));
            _favorite.Click();
            WaitPageLoading();

            Thread.Sleep(2500);
        }

        public bool IsFavoritePresent(string favoriteName)
        {
            var favorite = _webDriver.FindElements(By.XPath(String.Format(FAVORITE, favoriteName))).Count;

            if (favorite == 0)
                return false;

            return true;
        }

        public void SetFavoriteText(string favoriteName)
        {
            _favorite_search = WaitForElementIsVisible(By.XPath(FAVORITE_SEARCH));
            _favorite_search.SetValue(ControlType.TextBox, favoriteName);
            WaitPageLoading();
            WaitForLoad();
        }

        public ResultPage SelectFavoriteName(string favoriteName)
        {
            TryToClickOnElement(_favorite_apply);
            WaitPageLoading();
            WaitForLoad();

            return  new ResultPage(_webDriver, _testContext);
        }
    


        public EditFavoriteModal EditFavoriteName(string favoriteName)
        {
            TryToClickOnElement(_favorite_edit);
            return new EditFavoriteModal(_webDriver, _testContext);
        }

        public string GetFavoriteNameText()
        {
            WaitPageLoading();
            Thread.Sleep(2500);
            _favoriteFilter = _webDriver.FindElement(By.XPath(FAVORITE_FILTER));
            string text = _favoriteFilter.Text;

            return text;
        }

        public DateTime getDateFilterFrom()
        {
            var dateFormat = _dateFrom.GetAttribute("value");
            DateTime date;
            if (DateTime.TryParseExact(dateFormat, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
            {
                return date.Date;
            }
            else
            {

                throw new FormatException("Invalid date format");
            }
        }


        public DateTime getDateFilterTo()
        {
            var dateFormat = _dateTo.GetAttribute("value");
            DateTime date;
            if (DateTime.TryParseExact(dateFormat, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
            {
                return date.Date;
            }
            else
            {
                throw new FormatException("Invalid date format");
            }
        }

        public string GetFilterValue(string Id)
        {

            IWebElement selectElement = _webDriver.FindElement(By.Id(Id));
            SelectElement select = new SelectElement(selectElement);
            IWebElement selectedOption = select.SelectedOption;

            return selectedOption.Text;
        }

        public List<string> GetSelectedFilters(string Id)
        {
            var listGuesTyp = new List<string>();

            var selectElement = _webDriver.FindElement(By.Id(Id));

            WaitPageLoading();
            Thread.Sleep(1500);
            var options = selectElement.FindElements(By.TagName("option"));

            IJavaScriptExecutor js = (IJavaScriptExecutor)_webDriver;
            var selectedOptions = options.Where(option => option.Selected).ToList();

            foreach (var option in selectedOptions)
            {
                string optionText = (string)js.ExecuteScript("return arguments[0].text;", option);          
                listGuesTyp.Add(optionText);
            }

            return listGuesTyp;
        }
        public void FillPrincipalField_GenerateSupplyOrder(string site, DateTime from, DateTime to, DateTime deliveryDate, string deliveryLocation, bool isActive = true)
        {
            //Attente que la popup modale soit complètement chargée
            _sites = WaitForElementIsVisible(By.Id(SELECTED_SITE_ID));
            _sites.SetValue(ControlType.DropDownList, site);
            WaitForLoad();

            _datefrom = WaitForElementIsVisible(By.Id(DATE_FROM));
            _datefrom.SetValue(ControlType.DateTime, from);
            _datefrom.SendKeys(Keys.Tab);

            _dateto = WaitForElementIsVisible(By.Id(DATE_TO));
            _dateto.SetValue(ControlType.DateTime, to);
            _dateto.SendKeys(Keys.Tab);

            if (deliveryDate <= from)
            {
                _deliveryDate = WaitForElementIsVisible(By.Id(DELIVERY_DATE));
                _deliveryDate.SetValue(ControlType.DateTime, deliveryDate);
                _deliveryDate.SendKeys(Keys.Tab);
                // WaitPageLoading();
            }

            _deliveryLocation = WaitForElementIsVisible(By.Id(DELIVERY_LOCATION));
            _deliveryLocation.SetValue(ControlType.DropDownList, deliveryLocation);

            _activated = WaitForElementExists(By.Id(ACTIVATED));
            _activated.SetValue(ControlType.CheckBox, isActive);
        }

        public SupplyOrderItem Submit()
        {
            _submit = WaitForElementToBeClickable(By.Id(SUBMIT));
            _submit.Click();
            WaitForLoad();

            return new SupplyOrderItem(_webDriver, _testContext);
        }
        public void Check_Perfill_Quantities_From_Production_Management()
        {
            _perfillmanagement = WaitForElementExists(By.XPath(PERFILL_QUANTITIES_FROM_PRODUCTION_MANAGEMENT));
            _perfillmanagement.SetValue(PageBase.ControlType.CheckBox, true);
            WaitForLoad();
        }
        public CreateSupplyOrderModalPage CreateNewSupplyOrder()
        {
            ShowPlusMenu();

            _createNewSupplyOrder = WaitForElementIsVisible(By.Id(NEW_SUPPLY_ORDER));
            _createNewSupplyOrder.Click();
            WaitForLoad();

            return new CreateSupplyOrderModalPage(_webDriver, _testContext);
        }
        public string GetNewSupplierOrderNumber()
        {

             _messagevalidate = WaitForElementIsVisible(By.Id(MESSAGE_VALIDATE));
            return _messagevalidate.GetAttribute("value"); ;
        }
        public void UnselectAllItemsFromSubGroup()
        {
            _itemSubGroup = WaitForElementIsVisible(By.Id(ITEM_SUB_GROUP_FILTER));
            _itemSubGroup.Click();
            WaitForLoad();
            var _unselectAllItemSubGroup = WaitForElementExists(By.XPath(UNCHECK_ALL));
            _unselectAllItemSubGroup.Click();
            WaitForLoad();
            _itemSubGroup.Click();
            WaitForLoad();
        }
        public FilterAndFavoritesPage ClickPageUneFavo()
        {
            var pageUneFavoButton = WaitForElementIsVisible(By.XPath("//*[@id=\"favorite-filter-form\"]/nav/ul/li[2]/a"));
            pageUneFavoButton.Click();
            WaitForLoad();
            return new FilterAndFavoritesPage(_webDriver, _testContext);
        }
        public void ClickdeuxemePageFavo()
        {
            var pageUneFavoButton = WaitForElementIsVisible(By.XPath("//*[@id=\"favorite-filter-form\"]/nav/ul/li[3]/a"));
            pageUneFavoButton.Click();
            WaitForLoad();
        }


        public List<string> GetFavoritesList()
        {
            var favoriteElements = _webDriver.FindElements(By.ClassName("raw-favorite"));
            List<string> favoritesList = new List<string>();
            foreach (var favorite in favoriteElements)
            {
                var favoriteNumber = favorite.FindElement(By.TagName("span")).Text;
                favoritesList.Add(favoriteNumber);
            }
            return favoritesList;
        }
        public bool AreFavoritesDisplayedCorrectly()
        {
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(10));
            wait.Until(drv => drv.FindElements(By.CssSelector(".raw-favorite.clickable-favorite")).Count > 0);

            var favoriteElements = _webDriver.FindElements(By.CssSelector(".raw-favorite.clickable-favorite"));
            var actions = new Actions(_webDriver);

            foreach (var favoriteElement in favoriteElements)
            {
                // Move the cursor over each favorite element to trigger any hover-based display
                actions.MoveToElement(favoriteElement).Perform();
                var deleteButton = favoriteElement.FindElements(By.CssSelector("button.close")).FirstOrDefault();
                if (deleteButton == null || !deleteButton.Displayed || !deleteButton.Enabled)
                {
                    return false;
                }
                var editFavoriteButton = favoriteElement.FindElements(By.CssSelector("a.btn-edit-favorite")).FirstOrDefault();
                if (editFavoriteButton == null || !editFavoriteButton.Displayed || !editFavoriteButton.Enabled)
                {
                    return false;
                }
                var navigateButton = favoriteElement.FindElements(By.CssSelector("a[title='Get favorite and go to results']")).FirstOrDefault();
                if (navigateButton == null || !navigateButton.Displayed || !navigateButton.Enabled)
                {
                    return false;
                }

                // Wait for favorite number to be visible and non-empty
                var favoriteNumber = favoriteElement.FindElements(By.CssSelector("span")).FirstOrDefault();
                if (favoriteNumber == null || !favoriteNumber.Displayed || string.IsNullOrEmpty(favoriteNumber.Text))
                {
                    return false;
                }
            }

            return true;
        }


    }
}