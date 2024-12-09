using Microsoft.VisualStudio.TestTools.UnitTesting;
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
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Needs

{
    public class QuantityAdjustmentsNeedsPage : PageBase
    {

        public QuantityAdjustmentsNeedsPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //__________________________________Constantes_____________________________________

        // Général
        private const string EXTENDED_MENU = "//*[@id=\"div-body\"]/div/div[2]/div[1]/div/div[1]/button";
        private const string RESET_QUANTITIES = "btnResetQuantities";
        private const string RAZ_QUANTITIES = "btnRazQuantities";
        private const string PREVIOUS_DATE = "prev-date";
        private const string NEXT_DATE = "next-date";

        // Onglets
        private const string RESULT_TAB = "//*[@id=\"hrefTabContentItemContainer\"][text()='Results']";

        // Tableau
        private const string DATE = "//*[@id=\"list-item-with-action\"]/div/table/thead/tr/th[@class='day-col']";
        private const string SERVICES = "//*[@id=\"list-item-with-action\"]/div/table/tbody/tr";
        private const string ICON_DATASHEET = "//*[@id=\"list-item-with-action\"]/div/table/tbody/tr[*]/td[1]/span";
        private const string SERVICE_NAME = "//*[@id=\"list-item-with-action\"]/div/table/tbody/tr[*]/td[2]";
        private const string SERVICE_CUSTOMER = "//*[@id=\"list-item-with-action\"]/div/table/tbody/tr[*]/td[3]/div";
        private const string SERVICE_QUANTITY = "//*[@id=\"list-item-with-action\"]/div/table/tbody/tr[*]/td[contains(text(),'{0}')]/../td[6]/span[1]";
        private const string QUANTITY_INPUT = "//*[@id=\"list-item-with-action\"]/div/table/tbody/tr[*]/td[contains(text(),'{0}')]/../td[6]/input[6]";

        // Filtres
        private const string BACK = "//*[@id=\"item-filter-form\"]/div[1]/a";

        private const string SITE = "SiteId";
        private const string SITE_SELECTED = "//*[@id=\"SiteId\"]/option[@selected = 'selected']";
        private const string SORTBY = "cbSortBy";

        private const string FILTER_DATE_FROM = "ProdDateFrom";
        private const string FILTER_DATE_TO = "ProdDateTo";

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

        //__________________________________Variables______________________________________

        // General
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

        // Onglets
        [FindsBy(How = How.XPath, Using = RESULT_TAB)]
        public IWebElement _resultTab;

        // Tableau
        [FindsBy(How = How.XPath, Using = SERVICES)]
        public IWebElement _services;

        [FindsBy(How = How.XPath, Using = SERVICE_NAME)]
        public IWebElement _serviceName;

        [FindsBy(How = How.Id, Using = QUANTITY_INPUT)]
        public IWebElement _inputQty;

        [FindsBy(How = How.XPath, Using = SERVICE_QUANTITY)]
        public IWebElement _quantity;

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


        public enum FilterType
        {
            Site,
            SortBy,
            DateFrom,
            DateTo,
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
            ItemGroups
        }

        public void Filter(FilterType filterType, object value)
        {
            Actions action = new Actions(_webDriver);
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
                    break;

                case FilterType.ShowVacuumProd:
                    _showVacuumProd = WaitForElementExists(By.Id(SHOW_VACUUM_PROD));
                    action.MoveToElement(_showVacuumProd).Perform();
                    _showVacuumProd.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterType.ShowNormalProd:
                    _showNormalProd = WaitForElementExists(By.Id(SHOW_NORMAL_PROD));
                    action.MoveToElement(_showNormalProd).Perform();
                    _showNormalProd.SetValue(ControlType.RadioButton, value);
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
                    _customerFilter = WaitForElementIsVisible(By.Id(CUSTOMER_FILTER));
                    _customerFilter.Click();

                    // On décoche toutes les options
                    _unselectAllCustomer = WaitForElementIsVisible(By.XPath(UNSELECT_CUSTOMER));
                    _unselectAllCustomer.Click();
                    WaitPageLoading();

                    _searchCustomer = WaitForElementIsVisible(By.XPath(SEARCH_CUSTOMER));
                    _searchCustomer.SetValue(ControlType.TextBox, value);

                    var resultCustom = WaitForElementIsVisible(By.XPath("//span[contains(text(),'" + value + "')]"));
                    resultCustom.SetValue(ControlType.CheckBox, true);

                    _customerFilter.Click();
                    break;

                case FilterType.GuestType:
                    ComboBoxSelectById(new ComboBoxOptions(GUEST_TYPE_FILTER, (string)value));
                    break;

                case FilterType.ServicesCategorie:
                    _serviceCategoryFilter = WaitForElementIsVisible(By.Id(SERVICE_CATEGORIE_FILTER));
                    _serviceCategoryFilter.Click();
                    WaitPageLoading();

                    // On décoche toutes les options
                    _unselectAllServiceCategory = WaitForElementIsVisible(By.XPath(UNSELECT_CATEGORIE_SERVICE));
                    _unselectAllServiceCategory.Click();

                    _searchCategoryService = WaitForElementIsVisible(By.XPath(SEARCH_CATEGORIE_SERVICE));
                    _searchCategoryService.SetValue(ControlType.TextBox, value);

                    var serviceCat = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    serviceCat.SetValue(ControlType.CheckBox, true);

                    _serviceCategoryFilter.Click();
                    break;

                case FilterType.Services:
                    action.MoveToElement(WaitForElementExists(By.Id(SERVICE_FILTER))).Perform();
                    ComboBoxSelectById(new ComboBoxOptions(SERVICE_FILTER, (string)value));
                    action.MoveToElement(WaitForElementExists(By.XPath(BACK))).Perform();

                    WaitForLoad();
                    break;

                case FilterType.RecipeType:
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
                    break;

                default:
                    break;
            }

            WaitPageLoading();
            Thread.Sleep(2500);
        }

        public FilterAndFavoritesNeedsPage Back()
        {
            _back = WaitForElementIsVisible(By.XPath(BACK));
            _back.Click();
            WaitForLoad();

            return new FilterAndFavoritesNeedsPage(_webDriver, _testContext);
        }

        public string GetSite()
        {
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

        public List<string> GetCustomers()
        {
            _customerFilter = WaitForElementIsVisible(By.Id(CUSTOMER_FILTER));
            _customerFilter.Click();

            var customers = _webDriver.FindElements(By.XPath(CUSTOMER_SELECTED));
            return customers.Select(customer => customer.Text).ToList();
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
            _serviceCategoryFilter = WaitForElementIsVisible(By.Id(SERVICE_CATEGORIE_FILTER));
            _serviceCategoryFilter.Click();
            IWebElement serviceCategorie = _webDriver.FindElement(By.XPath(SERVICE_CATEGORIE_SELECTED));
            return serviceCategorie.Text;
        }

        public string GetService()
        {
            _serviceFilter = WaitForElementIsVisible(By.Id(SERVICE_FILTER));
            _serviceFilter.Click();

            var service = _webDriver.FindElement(By.XPath(SERVICE_SELECTED));
            return service.Text;
        }

        public string GetRecipeType()
        {
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

            var services = _webDriver.FindElements(By.XPath(SERVICE_NAME));

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

            IReadOnlyCollection<IWebElement> customers = _webDriver.FindElements(By.XPath(SERVICE_CUSTOMER));

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
                string[] date = FormatDate(dateFrom.Text);

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

            var dateTo = _webDriver.FindElements(By.XPath(DATE)).LastOrDefault();

            if (dateTo != null)
            {
                // On modifie le format de la date pour qu'elle soit compatible avec un format universel
                string[] date = FormatDate(dateTo.Text);

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
            WaitForLoad();

            Thread.Sleep(2000);//Too long to refresh
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
        public ResultPageNeeds GoToResultPage()
        {
            _resultTab = WaitForElementIsVisible(By.XPath(RESULT_TAB));
            _resultTab.Click();
            WaitPageLoading();
            WaitForLoad();

            return new ResultPageNeeds(_webDriver, _testContext);
        }

        // Tableau

        public int CountServices()
        {
            return _webDriver.FindElements(By.XPath(SERVICES)).Count();
        }

        public void SelectService(string serviceName)
        {
            _services = WaitForElementIsVisible(By.XPath(String.Format(SERVICE_QUANTITY, serviceName)));
            _services.Click();
        }

        public string GetServiceName()
        {
            _serviceName = WaitForElementIsVisible(By.XPath(SERVICE_NAME));
            return _serviceName.Text.Trim();
        }

        public bool IsServicePresent(string serviceName)
        {
            var elements = _webDriver.FindElements(By.XPath(SERVICE_NAME));

            if (elements.Count == 0)
                return false;

            foreach (var element in elements)
            {
                if (element.Text.Equals(serviceName))
                    return true;
            }

            return false;
        }

        public void SetQuantity(string value, string serviceName)
        {
            _inputQty = WaitForElementIsVisible(By.XPath(String.Format(QUANTITY_INPUT, serviceName)));
            _inputQty.SetValue(ControlType.TextBox, value);
            WaitForLoad();
            Thread.Sleep(2000);//too long to save new value
        }

        public double GetQuantity(string decimalSeparator, string serviceName)
        {
            WaitForLoad();
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

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
            string newDateFormat = date.Replace("(", "").Replace(")", "").Replace(" ", "-");
            string[] infoDate = newDateFormat.Split('-');

            if (infoDate[0] != null && infoDate[0].StartsWith("0"))
            {
                infoDate[0] = infoDate[0].Remove(0, 1);
            }

            return infoDate;
        }
    }
}
