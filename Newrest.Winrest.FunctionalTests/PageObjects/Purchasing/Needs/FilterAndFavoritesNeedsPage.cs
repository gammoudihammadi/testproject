using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Text.RegularExpressions;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Needs
{
    public class FilterAndFavoritesNeedsPage : PageBase
    {

        public FilterAndFavoritesNeedsPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //__________________________________Constantes_____________________________________

        private const string RESET_FILTER = "//*[@id=\"production-filter-form\"]/div/div[1]/a";

        private const string SITE = "SiteId";
        private const string FILTER_DATE_FROM = "ProdDateFrom";
        private const string FILTER_DATE_TO = "ProdDateTo";

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
        private const string SERVICE_SELECTED = "//*[@id=\"SelectedServices_ms\"]/span[2]";

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

        private const string FAVORITE = "//*[@id=\"favorite-filter-form\"]/div[2]/div/div/div[@class='raw-favorite clickable-favorite']/span[text()='{0}']";
        private const string FAVORITE_NAME = "Name";
        private const string SAVE_FAVORITE = "last";
        private const string DELETE_FAVORITE = "//*[@id=\"favorite-filter-form\"]/div[2]/div/div/div[@class='raw-favorite clickable-favorite']/span[text()='{0}']/../button";
        private const string DELETE_FAVORITECROSS = "//*[@id=\"favorite-filter-form\"]/div[2]/div/div/div[@class='raw-favorite clickable-favorite']/span[text()='{0}']/../button";
        private const string CONFIRM_DELETE_FAVORITE = "dataConfirmOK";

        private const string DONE = "//*[@id=\"btnFilter\"]/input";
        private const string MAKE_FAVORITE = "//*[@id=\"btnFilter\"]/div/a[2]";

        private const string FOOD_PACKETS = "SelectedFoodPackets_ms";
        private const string UNCHECK_ALL_FOOD_PACKETS = "/html/body/div[18]/div/ul/li[2]/a";

        //__________________________________Variables______________________________________

        [FindsBy(How = How.XPath, Using = RESET_FILTER)]
        private IWebElement _resetFilter;

        [FindsBy(How = How.Id, Using = SITE)]
        private IWebElement _site;

        [FindsBy(How = How.Id, Using = FILTER_DATE_FROM)]
        public IWebElement _dateFrom;

        [FindsBy(How = How.Id, Using = FILTER_DATE_TO)]
        public IWebElement _dateTo;

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

        [FindsBy(How = How.XPath, Using = SERVICE_SELECTED)]
        private IWebElement _nbServiceSelected;

        [FindsBy(How = How.XPath, Using = ITEM_GROUP_SELECTED)]
        private IWebElement _nbItemGroupSelected;

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

        [FindsBy(How = How.Id, Using = FOOD_PACKETS)]
        private IWebElement _foodPacket;

        [FindsBy(How = How.XPath, Using = UNCHECK_ALL_FOOD_PACKETS)]
        private IWebElement _foodPacketUncheckAll;


        //___________________________________Pages_________________________________________
        public enum FilterType
        {
            Site,
            DateFrom,
            DateTo,
            Workshops,
            Customers,
            GuestType,
            ServicesCategorie,
            Service,
            RecipeType,
            ItemGroups,
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
            FoodPacket
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

                case FilterType.Workshops:
                    _workShopFilter = WaitForElementIsVisible(By.Id(WORKSHOP_FILTER));
                    _workShopFilter.Click();

                    // On décoche toutes les options
                    _unselectAllWorkShop = WaitForElementIsVisible(By.XPath(UNSELECT_WORKSHOP));
                    _unselectAllWorkShop.Click();

                    _searchWorkShop = WaitForElementIsVisible(By.XPath(SEARCH_WORKSHOP));
                    _searchWorkShop.SetValue(ControlType.TextBox, value);

                    var resultWorkShop = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    resultWorkShop.SetValue(ControlType.CheckBox, true);

                    _workShopFilter.Click();
                    break;

                case FilterType.Customers:
                    _customerFilter = WaitForElementIsVisible(By.Id(CUSTOMER_FILTER));
                    _customerFilter.Click();

                    // On décoche toutes les options
                    _unselectAllCust = WaitForElementIsVisible(By.XPath(UNSELECT_CUSTOMER));
                    _unselectAllCust.Click();

                    _searchCustomer = WaitForElementIsVisible(By.XPath(SEARCH_CUSTOMER));
                    _searchCustomer.SetValue(ControlType.TextBox, value);

                    var resultCustom = WaitForElementIsVisible(By.XPath("//span[contains(text(),'" + value + "')]"));
                    resultCustom.SetValue(ControlType.CheckBox, true);

                    _customerFilter.Click();
                    break;

                case FilterType.GuestType:
                    _guestFilter = WaitForElementIsVisible(By.Id(GUEST_TYPE_FILTER));
                    _guestFilter.Click();

                    // On décoche toutes les options
                    _unselectAllGuest = WaitForElementIsVisible(By.XPath(UNSELECT_GUEST_TYPE));
                    _unselectAllGuest.Click();

                    _searchGuest = WaitForElementIsVisible(By.XPath(SEARCH_GUEST_TYPE));
                    _searchGuest.SetValue(ControlType.TextBox, value);

                    var resultGuest = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    resultGuest.SetValue(ControlType.CheckBox, true);

                    _guestFilter.Click();
                    break;

                case FilterType.ServicesCategorie:
                    _serviceCategorieFilter = WaitForElementIsVisible(By.Id(SERVICE_CATEGORIE_FILTER));
                    _serviceCategorieFilter.Click();

                    // On décoche toutes les options
                    _unselectAllServiceCategorie = WaitForElementIsVisible(By.XPath(UNSELECT_CATEGORIE_SERVICE));
                    _unselectAllServiceCategorie.Click();

                    _searchCategorieService = WaitForElementIsVisible(By.XPath(SEARCH_CATEGORIE_SERVICE));
                    _searchCategorieService.SetValue(ControlType.TextBox, value);

                    var serviceCat = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    serviceCat.SetValue(ControlType.CheckBox, true);

                    _serviceCategorieFilter.Click();
                    break;


                case FilterType.Service:
                    _serviceFilter = WaitForElementIsVisible(By.Id(SERVICE_FILTER));
                    _serviceFilter.Click();

                    // On décoche toutes les options
                    _unselectAllService = WaitForElementIsVisible(By.XPath(UNSELECT_SERVICE));
                    _unselectAllService.Click();

                    _searchService = WaitForElementIsVisible(By.XPath(SEARCH_SERVICE));
                    _searchService.SetValue(ControlType.TextBox, value);

                    var resultService = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    resultService.SetValue(ControlType.CheckBox, true);

                    _serviceFilter.Click();
                    break;

                case FilterType.RecipeType:
                    _recipeTypeFilter = WaitForElementIsVisible(By.Id(RECIPE_TYPE_FILTER));
                    _recipeTypeFilter.Click();

                    // On décoche toutes les options
                    _unselectAllRecipeType = WaitForElementIsVisible(By.XPath(UNSELECT_RECIPE_TYPE));
                    _unselectAllRecipeType.Click();

                    _searchRecipeType = WaitForElementIsVisible(By.XPath(SEARCH_RECIPE_TYPE));
                    _searchRecipeType.SetValue(ControlType.TextBox, value);

                    var resultRecipe = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    resultRecipe.SetValue(ControlType.CheckBox, true);

                    _recipeTypeFilter.Click();
                    break;

                case FilterType.ItemGroups:
                    _itemGroupFilter = WaitForElementIsVisible(By.Id(ITEM_GROUP_FILTER));
                    _itemGroupFilter.Click();

                    // On décoche toutes les options
                    _unselectAllItemGroup = WaitForElementIsVisible(By.XPath(UNSELECT_ITEM_GROUP));
                    _unselectAllItemGroup.Click();

                    _searchItemGroup = WaitForElementIsVisible(By.XPath(SEARCH_ITEM_GROUP));
                    _searchItemGroup.SetValue(ControlType.TextBox, value);
                    Thread.Sleep(1500);

                    var resultItem = WaitForElementIsVisible(By.XPath("/html/body/div[16]/ul"));
                    resultItem.Click();

                    _itemGroupFilter.Click();
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
                case FilterType.FoodPacket:
                    ComboBoxSelectById(new ComboBoxOptions("SelectedFoodPackets_ms", (string)value, false));
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
            Regex r = new Regex("([0-9]+) of ([0-9]+) selected");
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

        public string GetFoodPacket()
        {
            var foodPacket = WaitForElementIsVisible(By.XPath("//*/button[@id='SelectedFoodPackets_ms']/span[2]"));
            return foodPacket.Text;
        }

        public void SelectFoodPacket(string value)
        {
            if (value == null)
            {
                _foodPacket = WaitForElementIsVisible(By.Id(FOOD_PACKETS));
                _foodPacket.Click();
                _foodPacketUncheckAll = WaitForElementIsVisible(By.XPath(UNCHECK_ALL_FOOD_PACKETS));
                _foodPacketUncheckAll.Click();
                // on referme
                _foodPacket.Click();
            }
            else
            {
                ComboBoxSelectById(new ComboBoxOptions(FOOD_PACKETS, null, false));
            }
        }

        public string GetNbServiceSelected()
        {
            _nbServiceSelected = WaitForElementIsVisible(By.XPath(SERVICE_SELECTED));
            return _nbServiceSelected.Text;
        }

        public string GetNbItemGroupSelected()
        {
            _nbItemGroupSelected = WaitForElementIsVisible(By.XPath(ITEM_GROUP_SELECTED));
            return _nbItemGroupSelected.Text;
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

        public QuantityAdjustmentsNeedsPage DoneToQtyAjustement()
        {
            _done = WaitForElementIsVisible(By.XPath(DONE));
            _done.Click();
            WaitForLoad();
            return new QuantityAdjustmentsNeedsPage(_webDriver, _testContext);
        }

        public ResultPageNeeds DoneToResults()
        {
            _done = WaitForElementIsVisible(By.XPath(DONE));
            _done.Click();
            WaitForLoad();
            Thread.Sleep(1000);
            return new ResultPageNeeds(_webDriver, _testContext);
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

                ((IJavaScriptExecutor)_webDriver).ExecuteScript(
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
            WaitForLoad();

            Thread.Sleep(2500);
        }

        public bool IsFavoritePresent(string favoriteName)
        {
            var favorite = _webDriver.FindElements(By.XPath(String.Format(FAVORITE, favoriteName))).Count;

            if (favorite == 0)
                return false;

            return true;
        }
        public string CheckNbrItemGroupsSelected()
        {
            _itemGroupsSelection = WaitForElementIsVisible(By.XPath(ITEM_GROUP_SELECTION));
            Regex regex = new Regex(@"(\d+) of (\d+)");
            Match match = regex.Match(_itemGroupsSelection.Text);
            string selectedItemGroups = match.Groups[1].Value;
            string totItemGroups = match.Groups[2].Value;
            return selectedItemGroups;
        }
        public string CheckNbrItemGroupsSelectedFromResults()
        {
            _itemGroupsSelection = WaitForElementIsVisible(By.Id("SelectedItemGroups_ms"));
            Regex regex = new Regex(@"(\d+) of (\d+)");
            Match match = regex.Match(_itemGroupsSelection.Text);
            string selectedItemGroups = match.Groups[1].Value;
            string totItemGroups = match.Groups[2].Value;
            return selectedItemGroups;
        }
        public string CheckNbrRecipeTypeSelected()
        {
            _itemGroupsSelection = WaitForElementIsVisible(By.Id("SelectedRecipeTypes_ms"));
            Regex regex = new Regex(@"(\d+) of (\d+)");
            Match match = regex.Match(_itemGroupsSelection.Text);
            string selectedItemGroups = match.Groups[1].Value;
            string totItemGroups = match.Groups[2].Value;
            return selectedItemGroups;
        }
        public string CheckNbrWorkshopsSelected()
        {
            _itemGroupsSelection = WaitForElementIsVisible(By.Id("SelectedWorkshops_ms"));
            Regex regex = new Regex(@"(\d+) of (\d+)");
            Match match = regex.Match(_itemGroupsSelection.Text);
            string selectedItemGroups = match.Groups[1].Value;
            string totItemGroups = match.Groups[2].Value;
            return selectedItemGroups;
        }
    }
}
