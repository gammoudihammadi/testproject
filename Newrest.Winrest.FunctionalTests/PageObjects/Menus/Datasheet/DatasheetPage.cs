using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Recipes;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.Setup;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Menus.Datasheet
{
    public class DatasheetPage : PageBase
    {
        public DatasheetPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // _________________________________________ Constantes ____________________________________________

        // Général
        private const string PLUS_BUTTON = "//*[@id=\"tabContentItemContainer\"]/div/div/div[2]/button";
        private const string NEW_DATASHEET_DEV = "New datasheet";
        private const string NEW_DATASHEET_PATCH = "menus-datasheet-newdatasheetBtn"; 

        private const string EXTENDED_BUTTON = "//*[@id=\"tabContentItemContainer\"]/div/div/div[1]/button";
        private const string PRINT = "//*[@id=\"tabContentItemContainer\"]/div/div/div[1]/div/a[1]";
        private const string CONFIRM_PRINT = "print_btn";

        private const string EXPORT_WMS = "exportBtnToTXT";
        private const string EXPORT_WMS_CUSTOMER = "WMSExportCustomerId";
        private const string CONFIRM_EXPORT_WMS = "btnExport";
        private const string EXPORT_DATASHEET = "exportDatasheet";
        private const string IMPORT_DATASHEET = "//*/a[text()='Import Datasheets']";

        // Tableau datasheet
        private const string FIRST_DATASHEET = "menus-datasheet-edit-name-1";
        private const string DATASHEET_ID = "//*[@id=\"tableListMenu\"]/tbody/tr[{0}]/td[2]/a";
        private const string DATASHEET_NAME = "//*[@id=\"tableListMenu\"]/tbody/tr[{0}]/td[3]/a";
        private const string DATASHEET_NAME1 = "//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[3]/a";
        private const string DATASHEET_INACTIVE = "//*[@id=\"tableListMenu\"]/tbody/tr[{0}]/td[1]/img[@alt='Inactive']";
        private const string DATASHEET_USE_CASE = "//*[@id=\"tableListMenu\"]/tbody/tr[{0}]/td[6]";
        private const string DATASHEET_USE_CASE1 = "//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[6]";
        private const string DATASHEET_ALLERGEN = "//*[@id=\"tableListMenu\"]/tbody/tr[{0}]/td[8]/img";
        private const string DATASHEET_PICTURE = "//*[@id=\"tableListMenu\"]/tbody/tr[{0}]/td[9]/img";
        private const string DATASHEET_PICTURE1 = "//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[9]/img";

        // Filtres
        private const string RESET_FILTER = "//*[@id=\"item-filter-form\"]/div[1]/a";

        private const string DATASHEET_FILTER = "SearchPattern";
        private const string DATASHEETNAME_FILTER = "/html/body/div[3]/div/div/div[2]/div/form/div/div[1]/div/input";
        private const string CUSTOMER_CODE_FILTER = "CustomerSearch";
        private const string SERVICE_NAME_FILTER = "SearchService";
        private const string SORTBY_FILTER = "cbSortBy";
        private const string DATE_FROM_FILTER = "date-picker-start";
        private const string DATE_TO_FILTER = "date-picker-end";

        private const string SHOW_ALL_DEV = "ShowAll";
        private const string SHOW_ALL_PATCH = "//*[@id=\"ShowActive\"][@value='All']";

        private const string SHOW_ACTIVE_DEV = "ShowOnlyActive";
        private const string SHOW_ACTIVE_PATCH = "//*[@id=\"ShowActive\"][@value='ActiveOnly']";

        private const string SHOW_INACTIVE_DEV = "ShowOnlyInactive";
        private const string SHOW_INACTIVE_PATCH = "//*[@id=\"ShowActive\"][@value='InactiveOnly']";


        private const string US_SHOW_ALL_DEV = "ShowAllUseCase";
        private const string US_SHOW_ALL_PATCH = "//*[@id=\"UseCaseDisplay\"][@value='ShowAll']";

        private const string US_AFFECTED_DEV = "ShowWithUseCase";
        private const string US_AFFECTED_PATCH = "//*[@id=\"UseCaseDisplay\"][@value='ShowOnlyWithUseCase']";

        private const string US_NOT_AFFECTED_DEV = "ShowWithoutUseCase";
        private const string US_NOT_AFFECTED_PATCH = "//*[@id=\"UseCaseDisplay\"][@value='ShowOnlyWithoutUseCase']";

        private const string US_EXPIRED_DEV = "ShowExpiredUseCase";
        private const string US_EXPIRED_PATCH = "//*[@id=\"UseCaseDisplay\"][@value='ExpiredOnly']";


        private const string ALLERGEN_SHOW_ALL_DEV = "ShowAllAllergen";
        private const string ALLERGEN_SHOW_ALL_PATCH = "//*[@id=\"AllergenValidatedDisplay\"][@value='ShowAll']";

        private const string ALLERGEN_VALIDATED_DEV = "ShowAllergenValidated";
        private const string ALLERGEN_VALIDATED_PATCH = "//*[@id=\"AllergenValidatedDisplay\"][@value='ShowOnlyAllergenValidated']";

        private const string ALLERGEN_NOT_VALIDATED_DEV = "ShowAllergenNotValidated";
        private const string ALLERGEN_NOT_VALIDATED_PATCH = "//*[@id=\"AllergenValidatedDisplay\"][@value='ShowOnlyAllergenNotValidated']";


        private const string CUSTOMERS_FILTER = "SelectedCustomers_ms";
        private const string UNCHECK_CUSTOMERS = "/html/body/div[11]/div/ul/li[2]/a";
        private const string CUSTOMERS_SEARCH = "/html/body/div[10]/div/div/label/input";

        private const string CATEGORIES_FILTER = "SelectedCategories_ms";
        private const string UNCHECK_CATEGORIES = "/html/body/div[12]/div/ul/li[2]/a";
        private const string CATEGORIES_SEARCH = "/html/body/div[11]/div/div/label/input";

        private const string SITES_FILTER = "SelectedSites_ms";
        private const string UNCHECK_SITES = "/html/body/div[12]/div/ul/li[2]/a";
        private const string SITES_SEARCH = "/html/body/div[12]/div/div/label/input";

        private const string GUEST_TYPE_FILTER = "SelectedGuestTypes_ms";
        private const string UNCHECK_GUEST_TYPE = "/html/body/div[13]/div/ul/li[2]/a";
        private const string GUEST_TYPE_SEARCH = "/html/body/div[13]/div/div/label/input";

        private const string WORKSHOP_FILTER = "SelectedRecipesWorkshops_ms";
        private const string UNCHECK_WORKSHOP = "/html/body/div[14]/div/ul/li[2]/a";
        private const string WORKSHOP_SEARCH = "/html/body/div[14]/div/div/label/input";

        private const string COOKING_MODE_FILTER = "SelectedRecipesCookingModes_ms";
        private const string UNCHECK_COOKING_MODE = "/html/body/div[15]/div/ul/li[2]/a";
        private const string COOKING_MODE_SEARCH = "/html/body/div[15]/div/div/label/input";



        private const string EXTEND_MENU = "//*[@id=\"tabContentItemContainer\"]/div/div/div[1]/button";
        private const string MASSIVE_DELETE = "//*[@id=\"tabContentItemContainer\"]/div/div/div[1]/div/a[2]";
        private const string SITE_BTN = "//*[@id=\"SelectedSiteIds_ms\"]";
        private const string MASSIVE_DELETE_SITE_FILTER = "/html/body/div[19]/ul/li[*]/label/span[text()='{0}']/../input";
        private const string CUSTOMER_BTN = "//*[@id=\"SelectedCustomersForDeletion_ms\"]";
        private const string MASSIVE_DELETE_CUSTOMER_FILTER = "/html/body/div[20]/ul/li[*]/label/span[text()='{0}']/../input";
        private const string MASSIVE_DELETE_SEARCH = "//*[@id=\"SearchDatasheetsBtn\"]";
        private const string MASSIVE_DELETE_SELECTALL = "//*[@id=\"selectAll\"]";
        private const string MASSIVE_DELETE_DELETE = "//*[@id=\"deleteDatasheetBtn\"]";
        private const string MASSIVE_DELETE_CONFIRM_DELETE = "//*[@id=\"dataConfirmOK\"]";
        private const string GUEST_TYPE_BTN = "//*[@id=\"SelectedGuestTypesForDeletion_ms\"]";
        private const string MASSIVE_DELETE_GUEST_TYPE_FILTER = "/html/body/div[17]/ul/li[*]/label/span[text()='{0}']/../input";
        private const string MASSIVE_DELETE_BY_DATASHEET_NAME = "/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/table/tbody/tr[*]/td[2]/div[text()='{0}']/../../td[1]/input[1]";
        private const string PAGE_SIZE = "/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/nav/select";
        private const string PAGE_SIZE_TO_SET = "/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/nav/select/option[@value='{0}']";
        private const string OK_BTN_PATCH = "//*[@id=\"modal-1\"]/div/div/div[3]/button";
        private const string OK_BTN_DEV = "//*[@id=\"modal-1\"]/div[3]/button";
        private const string ROWS_FOR_PAGINATION = "//*[@id=\"tabContentItemContainer\"]/table/tbody/tr[*]/td[2]";
        
        // _________________________________________ Variables _____________________________________________

        // Général
        [FindsBy(How = How.XPath, Using = PLUS_BUTTON)]
        private IWebElement _plusButton;

        [FindsBy(How = How.LinkText, Using = NEW_DATASHEET_DEV)]
        private IWebElement _createNewDatasheet;

        [FindsBy(How = How.XPath, Using = EXTENDED_BUTTON)]
        private IWebElement _extendedButton;

        [FindsBy(How = How.XPath, Using = PRINT)]
        private IWebElement _print;

        [FindsBy(How = How.Id, Using = CONFIRM_PRINT)]
        private IWebElement _confirmPrint;

        [FindsBy(How = How.Id, Using = EXPORT_WMS)]
        private IWebElement _exportWMS;

        [FindsBy(How = How.Id, Using = EXPORT_WMS_CUSTOMER)]
        private IWebElement _exportWMSCustomer;

        [FindsBy(How = How.Id, Using = CONFIRM_EXPORT_WMS)]
        private IWebElement _confirmExportWMS;

        // Tableau datasheet
        [FindsBy(How = How.XPath, Using = FIRST_DATASHEET)]
        private IWebElement _firstDatasheet;

        // _________________________________________ Filtres _______________________________________________

        [FindsBy(How = How.XPath, Using = RESET_FILTER)]
        private IWebElement _resetFilter;

        [FindsBy(How = How.Id, Using = DATASHEET_FILTER)]
        private IWebElement _datasheetFilter;

        [FindsBy(How = How.Id, Using = CUSTOMER_CODE_FILTER)]
        private IWebElement _customerCodeFilter;

        [FindsBy(How = How.Id, Using = SERVICE_NAME_FILTER)]
        private IWebElement _serviceNameFilter;

        [FindsBy(How = How.Id, Using = SORTBY_FILTER)]
        private IWebElement _sortByFilter;

        [FindsBy(How = How.Id, Using = DATE_FROM_FILTER)]
        private IWebElement _dateFromFilter;

        [FindsBy(How = How.Id, Using = DATE_TO_FILTER)]
        private IWebElement _dateToFilter;

        [FindsBy(How = How.Id, Using = SHOW_ALL_DEV)]
        private IWebElement _showAllDev;
        
        [FindsBy(How = How.XPath, Using = SHOW_ALL_PATCH)]
        private IWebElement _showAllPatch;

        [FindsBy(How = How.Id, Using = SHOW_ACTIVE_DEV)]
        private IWebElement _showActiveDev;
        
        [FindsBy(How = How.XPath, Using = SHOW_ACTIVE_PATCH)]
        private IWebElement _showActivePatch;

        [FindsBy(How = How.Id, Using = SHOW_INACTIVE_DEV)]
        private IWebElement _showInactiveDev;
        
        [FindsBy(How = How.XPath, Using = SHOW_INACTIVE_PATCH)]
        private IWebElement _showInactivePatch;

        [FindsBy(How = How.Id, Using = US_SHOW_ALL_DEV)]
        private IWebElement _useCaseShowAllFilterDev;
        
        [FindsBy(How = How.XPath, Using = US_SHOW_ALL_PATCH)]
        private IWebElement _useCaseShowAllFilterPatch;

        [FindsBy(How = How.Id, Using = US_AFFECTED_DEV)]
        private IWebElement _useCaseAffectedFilterDev;
        
        [FindsBy(How = How.XPath, Using = US_AFFECTED_PATCH)]
        private IWebElement _useCaseAffectedFilterPatch;

        [FindsBy(How = How.Id, Using = US_NOT_AFFECTED_DEV)]
        private IWebElement _useCaseNotAffectedFilterDev;
        
        [FindsBy(How = How.XPath, Using = US_NOT_AFFECTED_PATCH)]
        private IWebElement _useCaseNotAffectedFilterPatch;

        [FindsBy(How = How.Id, Using = US_EXPIRED_DEV)]
        private IWebElement _useCaseExpiredFilterDev;
        
        [FindsBy(How = How.XPath, Using = US_EXPIRED_PATCH)]
        private IWebElement _useCaseExpiredFilterPatch;

        [FindsBy(How = How.Id, Using = ALLERGEN_SHOW_ALL_DEV)]
        private IWebElement _allergenShowAllFilterDev;
        
        [FindsBy(How = How.XPath, Using = ALLERGEN_SHOW_ALL_PATCH)]
        private IWebElement _allergenShowAllFilterPatch;

        [FindsBy(How = How.Id, Using = ALLERGEN_VALIDATED_DEV)]
        private IWebElement _allergenValidatedFilterDev;
        
        [FindsBy(How = How.XPath, Using = ALLERGEN_VALIDATED_PATCH)]
        private IWebElement _allergenValidatedFilterPatch;

        [FindsBy(How = How.Id, Using = ALLERGEN_NOT_VALIDATED_DEV)]
        private IWebElement _allergenNotValidatedFilterDev;
        
        [FindsBy(How = How.XPath, Using = ALLERGEN_NOT_VALIDATED_PATCH)]
        private IWebElement _allergenNotValidatedFilterPatch;

        [FindsBy(How = How.Id, Using = CUSTOMERS_FILTER)]
        private IWebElement _customersFilter;

        [FindsBy(How = How.XPath, Using = UNCHECK_CUSTOMERS)]
        private IWebElement _uncheckCustomers;

        [FindsBy(How = How.XPath, Using = CUSTOMERS_SEARCH)]
        private IWebElement _customersSearch;

        [FindsBy(How = How.Id, Using = CATEGORIES_FILTER)]
        private IWebElement _categoriesFilter;

        [FindsBy(How = How.XPath, Using = UNCHECK_CATEGORIES)]
        private IWebElement _uncheckCategories;

        [FindsBy(How = How.XPath, Using = CATEGORIES_SEARCH)]
        private IWebElement _categoriesSearch;

        [FindsBy(How = How.Id, Using = SITES_FILTER)]
        private IWebElement _sitesFilter;

        [FindsBy(How = How.XPath, Using = UNCHECK_SITES)]
        private IWebElement _uncheckSites;

        [FindsBy(How = How.XPath, Using = SITES_SEARCH)]
        private IWebElement _sitesSearch;

        [FindsBy(How = How.Id, Using = GUEST_TYPE_FILTER)]
        private IWebElement _guestTypesFilter;

        [FindsBy(How = How.XPath, Using = UNCHECK_GUEST_TYPE)]
        private IWebElement _uncheckGuestTypes;

        [FindsBy(How = How.XPath, Using = GUEST_TYPE_SEARCH)]
        private IWebElement _guestTypesSearch;

        [FindsBy(How = How.Id, Using = WORKSHOP_FILTER)]
        private IWebElement _workshopsFilter;

        [FindsBy(How = How.XPath, Using = UNCHECK_WORKSHOP)]
        private IWebElement _uncheckWorkshops;

        [FindsBy(How = How.XPath, Using = WORKSHOP_SEARCH)]
        private IWebElement _workshopsSearch;

        [FindsBy(How = How.Id, Using = COOKING_MODE_FILTER)]
        private IWebElement _cookingModesFilter;

        [FindsBy(How = How.XPath, Using = UNCHECK_COOKING_MODE)]
        private IWebElement _uncheckCookingModes;

        [FindsBy(How = How.XPath, Using = COOKING_MODE_SEARCH)]
        private IWebElement _cookingModesSearch;


        public enum FilterType
        {
            DatasheetName,
            CustomerCode,
            ServiceName,
            SortBy,
            DateFrom,
            DateTo,
            ShowAll,
            ShowActive,
            ShowInactive,
            UseCaseShowAll,
            UseCaseAffected,
            UseCaseNotAffected,
            UseCaseExpired,
            AllergenShowAll,
            AllergenValidated,
            AllergenNotValidated,
            Customers,
            Categories,
            Sites,
            GuestTypes,
            Workshops,
            CookingModes
        }

        public enum ImportType
        {
            ImportDatasheet
        }


        public void Filter(FilterType filterType, object value)
        {
            Actions action = new Actions(_webDriver);

            switch (filterType)
            {
                case FilterType.DatasheetName:
                    _datasheetFilter = WaitForElementIsVisible(By.Id(DATASHEET_FILTER));
                    _datasheetFilter.SetValue(ControlType.TextBox, value);
                    break;
                case FilterType.CustomerCode:
                    _customerCodeFilter = WaitForElementIsVisible(By.Id(CUSTOMER_CODE_FILTER));
                    _customerCodeFilter.SetValue(ControlType.TextBox, value);
                    break;
                case FilterType.ServiceName:
                    _serviceNameFilter = WaitForElementIsVisible(By.Id(SERVICE_NAME_FILTER));
                    _serviceNameFilter.SetValue(ControlType.TextBox, value);
                    break;
                case FilterType.SortBy:
                    _sortByFilter = WaitForElementExists(By.Id(SORTBY_FILTER));
                    _sortByFilter.Click();
                    var element = WaitForElementIsVisible(By.XPath("//option[contains(@value,'" + value + "')]"));
                    _sortByFilter.SetValue(ControlType.DropDownList, element.Text);
                    _sortByFilter.Click();
                    break;
                case FilterType.DateFrom:
                    _dateFromFilter = WaitForElementIsVisible(By.Id(DATE_FROM_FILTER));
                    _dateFromFilter.SetValue(ControlType.DateTime, value);
                    _dateFromFilter.SendKeys(Keys.Tab);
                    break;
                case FilterType.DateTo:
                    _dateToFilter = WaitForElementIsVisible(By.Id(DATE_TO_FILTER));
                    _dateToFilter.SetValue(ControlType.DateTime, value);
                    _dateToFilter.SendKeys(Keys.Tab);
                    break;
                case FilterType.ShowAll:
                    if(isElementVisible(By.Id(SHOW_ALL_DEV)))
                    {
                        _showAllDev = WaitForElementExists(By.Id(SHOW_ALL_DEV));
                        action.MoveToElement(_showAllDev).Perform();
                        _showAllDev.SetValue(ControlType.RadioButton, value);
                    }
                    else
                    {
                        _showAllPatch = WaitForElementExists(By.XPath(SHOW_ALL_PATCH));
                        action.MoveToElement(_showAllPatch).Perform();
                        _showAllPatch.SetValue(ControlType.RadioButton, value);
                    }
                    break;
                case FilterType.ShowActive:
                    if(isElementVisible(By.Id(SHOW_ACTIVE_DEV)))
                    {
                        _showActiveDev = WaitForElementExists(By.Id(SHOW_ACTIVE_DEV));
                        action.MoveToElement(_showActiveDev).Perform();
                        _showActiveDev.SetValue(ControlType.RadioButton, value);
                    }
                    else
                    {
                        _showActivePatch = WaitForElementExists(By.XPath(SHOW_ACTIVE_PATCH));
                        action.MoveToElement(_showActivePatch).Perform();
                        _showActivePatch.SetValue(ControlType.RadioButton, value);
                    }
                    break;
                case FilterType.ShowInactive:
                    if(isElementVisible(By.Id(SHOW_INACTIVE_DEV)))
                    {
                        _showInactiveDev = WaitForElementExists(By.Id(SHOW_INACTIVE_DEV));
                        action.MoveToElement(_showInactiveDev).Perform();
                        _showInactiveDev.SetValue(ControlType.RadioButton, value);
                    }
                    else
                    {
                        _showInactivePatch = WaitForElementExists(By.XPath(SHOW_INACTIVE_PATCH));
                        action.MoveToElement(_showInactivePatch).Perform();
                        _showInactivePatch.SetValue(ControlType.RadioButton, value);
                    }
                    break;
                case FilterType.UseCaseShowAll:
                    if(isElementVisible(By.Id(US_SHOW_ALL_DEV)))
                    {
                        _useCaseShowAllFilterDev = WaitForElementExists(By.Id(US_SHOW_ALL_DEV));
                        action.MoveToElement(_useCaseShowAllFilterDev).Perform();
                        _useCaseShowAllFilterDev.SetValue(ControlType.RadioButton, value);
                    }
                    else
                    {
                        _useCaseShowAllFilterPatch = WaitForElementExists(By.XPath(US_SHOW_ALL_PATCH));
                        action.MoveToElement(_useCaseShowAllFilterPatch).Perform();
                        _useCaseShowAllFilterPatch.SetValue(ControlType.RadioButton, value);
                    }
                    break;
                case FilterType.UseCaseAffected:
                    if(isElementVisible(By.Id(US_AFFECTED_DEV)))
                    {
                        _useCaseAffectedFilterDev = WaitForElementExists(By.Id(US_AFFECTED_DEV));
                        action.MoveToElement(_useCaseAffectedFilterDev).Perform();
                        _useCaseAffectedFilterDev.SetValue(ControlType.RadioButton, value);
                    }
                    else
                    {
                        _useCaseAffectedFilterPatch = WaitForElementExists(By.XPath(US_AFFECTED_PATCH));
                        action.MoveToElement(_useCaseAffectedFilterPatch).Perform();
                        _useCaseAffectedFilterPatch.SetValue(ControlType.RadioButton, value);
                    }
                    break;
                case FilterType.UseCaseNotAffected:
                    _useCaseNotAffectedFilterDev = WaitForElementExists(By.Id(US_NOT_AFFECTED_DEV));
                    action.MoveToElement(_useCaseNotAffectedFilterDev).Perform();
                    _useCaseNotAffectedFilterDev.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.UseCaseExpired:
                    try
                    {
                        _useCaseExpiredFilterDev = WaitForElementExists(By.Id(US_EXPIRED_DEV));
                        action.MoveToElement(_useCaseExpiredFilterDev).Perform();
                        _useCaseExpiredFilterDev.SetValue(ControlType.RadioButton, value);
                    }
                    catch
                    {
                        _useCaseExpiredFilterPatch = WaitForElementExists(By.XPath(US_EXPIRED_PATCH));
                        action.MoveToElement(_useCaseExpiredFilterPatch).Perform();
                        _useCaseExpiredFilterPatch.SetValue(ControlType.RadioButton, value);
                    }
                    break;
                case FilterType.AllergenShowAll:
                    try
                    {
                        _allergenShowAllFilterDev = WaitForElementExists(By.Id(ALLERGEN_SHOW_ALL_DEV));
                        action.MoveToElement(_allergenShowAllFilterDev).Perform();
                        _allergenShowAllFilterDev.SetValue(ControlType.RadioButton, value);
                    }
                    catch
                    {
                        _allergenShowAllFilterPatch = WaitForElementExists(By.XPath(ALLERGEN_SHOW_ALL_PATCH));
                        action.MoveToElement(_allergenShowAllFilterPatch).Perform();
                        _allergenShowAllFilterPatch.SetValue(ControlType.RadioButton, value);
                    }
                    break;
                case FilterType.AllergenValidated:
                    try
                    {
                        _allergenValidatedFilterDev = WaitForElementExists(By.Id(ALLERGEN_VALIDATED_DEV));
                        action.MoveToElement(_allergenValidatedFilterDev).Perform();
                        _allergenValidatedFilterDev.SetValue(ControlType.RadioButton, value);
                    }
                    catch
                    {
                        _allergenValidatedFilterPatch = WaitForElementExists(By.XPath(ALLERGEN_VALIDATED_PATCH));
                        action.MoveToElement(_allergenValidatedFilterPatch).Perform();
                        _allergenValidatedFilterPatch.SetValue(ControlType.RadioButton, value);
                    }
                    break;
                case FilterType.AllergenNotValidated:
                    try
                    {
                        _allergenNotValidatedFilterDev = WaitForElementExists(By.Id(ALLERGEN_NOT_VALIDATED_DEV));
                        action.MoveToElement(_allergenNotValidatedFilterDev).Perform();
                        _allergenNotValidatedFilterDev.SetValue(ControlType.RadioButton, value);
                    }
                    catch
                    {
                        _allergenNotValidatedFilterPatch = WaitForElementExists(By.XPath(ALLERGEN_NOT_VALIDATED_PATCH));
                        action.MoveToElement(_allergenNotValidatedFilterPatch).Perform();
                        _allergenNotValidatedFilterPatch.SetValue(ControlType.RadioButton, value);
                    }
                    break;
                case FilterType.Customers:
                    _customersFilter = WaitForElementExists(By.Id(CUSTOMERS_FILTER));
                    action.MoveToElement(_customersFilter).Perform();
                    if (value == null)
                    {
                        _customersFilter.Click();
                        _uncheckCustomers = WaitForElementIsVisible(By.XPath(UNCHECK_CUSTOMERS));
                        _uncheckCustomers.Click();
                        // on referme
                        _customersFilter.Click();
                    }
                    else
                    {
                        ComboBoxSelectById(new ComboBoxOptions(CUSTOMERS_FILTER, (string)value));
                    }
                    break;
                case FilterType.Categories:
                    _categoriesFilter = WaitForElementExists(By.Id(CATEGORIES_FILTER));
                    action.MoveToElement(_categoriesFilter).Perform();
                    if (value == null)
                    {
                        _categoriesFilter.Click();
                        _uncheckCategories = WaitForElementIsVisible(By.XPath(UNCHECK_CATEGORIES));
                        _uncheckCategories.Click();
                        // on referme
                        _categoriesFilter.Click();
                    }
                    else
                    {
                        ComboBoxSelectById(new ComboBoxOptions(CATEGORIES_FILTER, (string)value));
                    }
                    break;
                case FilterType.Sites:
                    _sitesFilter = WaitForElementExists(By.Id(SITES_FILTER));
                    action.MoveToElement(_sitesFilter).Perform();
                    ComboBoxSelectById(new ComboBoxOptions(SITES_FILTER, (string)value));
                    break;
                case FilterType.GuestTypes:
                    _guestTypesFilter = WaitForElementExists(By.Id(GUEST_TYPE_FILTER));
                    action.MoveToElement(_guestTypesFilter).Perform();
                    ComboBoxSelectById(new ComboBoxOptions(GUEST_TYPE_FILTER, (string)value));
                    break;
                case FilterType.Workshops:
                    _workshopsFilter = WaitForElementExists(By.Id(WORKSHOP_FILTER));
                    action.MoveToElement(_workshopsFilter).Perform();
                    _workshopsFilter.Click();

                    _uncheckWorkshops = WaitForElementIsVisible(By.XPath(UNCHECK_WORKSHOP));
                    _uncheckWorkshops.Click();

                    if (value != null)
                    {
                        _workshopsSearch = WaitForElementIsVisible(By.XPath(WORKSHOP_SEARCH));
                        _workshopsSearch.SetValue(ControlType.TextBox, value);

                        var workshopToCheck = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                        workshopToCheck.SetValue(ControlType.CheckBox, true);
                    }
                    _workshopsFilter.Click();
                    break;
                case FilterType.CookingModes:
                    _cookingModesFilter = WaitForElementExists(By.Id(COOKING_MODE_FILTER));
                    action.MoveToElement(_cookingModesFilter).Perform();
                    _cookingModesFilter.Click();

                    _uncheckCookingModes = WaitForElementIsVisible(By.XPath(UNCHECK_COOKING_MODE));
                    _uncheckCookingModes.Click();

                    if (value != null)
                    {
                        _cookingModesSearch = WaitForElementIsVisible(By.XPath(COOKING_MODE_SEARCH));
                        _cookingModesSearch.SetValue(ControlType.TextBox, value);
                        Thread.Sleep(1500);

                        var cookingModeToCheck = WaitForElementIsVisible(By.XPath("/html/body/div[15]/ul"));
                        cookingModeToCheck.Click();
                    }
                    _cookingModesFilter.Click();
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);
            }

            WaitPageLoading();
            WaitForLoad();
        }

        public void ResetFilter()
        {
            _webDriver.Navigate().Refresh();
            _resetFilter = WaitForElementIsVisible(By.Id("ResetFilter"));
            _resetFilter.Click();
            WaitForLoad();
            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                // DateTo à null par défaut
            }
        }

        public bool IsSortedByName()
        {
            bool valueBool = true;
            var ancienName = "";
            int tot;
            int i = 1;

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


            var elements = _webDriver.FindElements(By.XPath(DATASHEET_NAME1));

            foreach (var element in elements)
            {
                try
                {
                    string name = element.Text;

                    if (i == 0)
                        ancienName = name;

                    if (String.Compare(ancienName, name) > 0)
                    { valueBool = false; }

                    ancienName = name;
                }
                catch
                {
                    valueBool = false;
                }
                i++;
            }
            return valueBool;
        }

        public bool IsSortedByID()
        {
            bool valueBool = true;
            int ancienID = int.MaxValue;
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

            for (int i = 0; i < tot; i++)
            {
                try
                {
                    IWebElement element = _webDriver.FindElement(By.XPath(string.Format(DATASHEET_ID, i + 1)));
                    int id = int.Parse(element.Text.Substring(element.Text.IndexOf("°") + 1));

                    if (i == 0)
                        ancienID = id;

                    if (ancienID > id)
                    { valueBool = false; }

                    ancienID = id;
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

            for (int i = 0; i < tot; i++)
            {
                try
                {
                    _webDriver.FindElement(By.XPath(String.Format(DATASHEET_INACTIVE, i + 1)));

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

        public bool CheckAllergenValidation(bool validated)
        {
            bool isValidated = false;
            int tot = CheckTotalNumber() > 100 ? 100 : CheckTotalNumber();

            if (tot == 0)
                return false;

            for (int i = 0; i < tot; i++)
            {
                try
                {
                    var element = _webDriver.FindElement(By.XPath(String.Format(DATASHEET_ALLERGEN, i + 1)));

                    if (validated && element.GetAttribute("alt").Equals("Allergen(s) validated"))
                        isValidated = true;
                    else if (validated && !element.GetAttribute("alt").Equals("Allergen(s) validated"))
                        return false;
                    else if (!validated && element.GetAttribute("alt").Equals("Allergen(s) validated"))
                        return true;
                }
                catch
                {
                    return false;
                }
            }
            return isValidated;
        }

        public bool IsUseCaseAffected()
        {
            bool valueBool = false;
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
            var elements = _webDriver.FindElements(By.XPath(DATASHEET_USE_CASE1));
            foreach (var element in elements)
            {
                try
                {
                    if (element.Text != "0")
                        valueBool = true;
                }
                catch
                {
                    valueBool = false;
                }
            }

            return valueBool;
        }

        public bool IsPictureLinked()
        {
            bool valueBool = false;
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

            var elements = _webDriver.FindElements(By.XPath(DATASHEET_PICTURE1));
            foreach (var element in elements)
            {
                try
                {
                    if (element.GetAttribute("alt").Equals("Has a picture"))
                        valueBool = true;
                }
                catch
                {
                    valueBool = false;
                }
            }

            return valueBool;
        }

        // _______________________________________________ Méthodes ____________________________________________________

        // Général
        public override void ShowPlusMenu()
        {
            Actions action = new Actions(_webDriver);
            _plusButton = WaitForElementIsVisible(By.XPath(PLUS_BUTTON));
            action.MoveToElement(_plusButton).Perform();
        }

        public DatasheetCreateModalPage CreateNewDatasheet()
        {
            WaitForLoad();
            LoadingPage();
            ShowPlusMenu();
            if (IsDev())
            {
                _createNewDatasheet = WaitForElementIsVisible(By.LinkText(NEW_DATASHEET_DEV));
            }
            else
            {
                _createNewDatasheet = WaitForElementIsVisible(By.Id(NEW_DATASHEET_PATCH));

            }

            _createNewDatasheet.Click();
            WaitForLoad();

            return new DatasheetCreateModalPage(_webDriver, _testContext);
        }

        public override void ShowExtendedMenu()
        {
            Actions action = new Actions(_webDriver);
            _extendedButton = WaitForElementIsVisible(By.XPath(EXTENDED_BUTTON));
            action.MoveToElement(_extendedButton).Perform();
        }

        public PrintReportPage PrintDatasheet(bool versionPrint)
        {
            ShowExtendedMenu();
            _print = WaitForElementIsVisible(By.XPath(PRINT));
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

        public void ExportWMSDatasheet(string customer)
        {
            ShowExtendedMenu();
            _exportWMS = WaitForElementIsVisible(By.Id(EXPORT_WMS));
            _exportWMS.Click();
            WaitForLoad();

            _exportWMSCustomer = WaitForElementIsVisible(By.Id(EXPORT_WMS_CUSTOMER));
            _exportWMSCustomer.SetValue(ControlType.DropDownList, customer);

            _confirmExportWMS = WaitForElementToBeClickable(By.Id(CONFIRM_EXPORT_WMS));
            _confirmExportWMS.Click();
            WaitForLoad();
        }

        public void ExportDatasheet()
        {
            ShowExtendedMenu();
            var _exportDatasheet = WaitForElementIsVisible(By.Id(EXPORT_DATASHEET));
            _exportDatasheet.Click();
            WaitForLoad();
            IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));
            ClickPrintButton();

            WaitForDownload();
            Close();
        }

        public void ImportDatasheet(ImportType importType, string filename)
        {
            ShowExtendedMenu();

            var _importExtendedButton = WaitForElementIsVisible(By.XPath(IMPORT_DATASHEET));
            _importExtendedButton.Click();
            WaitForLoad();

            switch (importType)
            {
                case ImportType.ImportDatasheet:
                    var _importDatasheet = WaitForElementIsVisible(By.Id("fileSent"));
                    _importDatasheet.SendKeys(filename);
                    WaitForLoad();
                    var checkButton = WaitForElementIsVisible(By.XPath("//*/button[text()='Check file']"));
                    checkButton.Click();
                    WaitForLoad();
                    var verification = WaitForElementIsVisible(By.XPath("//*/span[@class='green-text']"));
                    Assert.IsTrue(verification.Text.Contains("line"));
                    var import = WaitForElementIsVisible(By.XPath("//*/button[text()='Import File']"));
                    import.Click();
                    WaitPageLoading();
                    WaitForLoad();
                    var success = WaitForElementIsVisible(By.XPath("//*/p[@class='green-text']"));
                    Assert.AreEqual("Import of datasheets done successfully.", success.Text);
                    WaitForLoad();
                    break;
            }
        }



        // Tableau datasheet
        public string GetFirstDatasheetName()
        {
            if (IsDev())
            {
                _firstDatasheet = WaitForElementIsVisible(By.Id(FIRST_DATASHEET));
            }
            else
            {
                _firstDatasheet = WaitForElementIsVisible(By.XPath("//*[@id=\"tableListMenu\"]/tbody/tr/td[3]/a/b"));
            }
            
            return _firstDatasheet.Text;
        }

        public DatasheetDetailsPage SelectFirstDatasheet()
        {
            if (IsDev())
            {
                _firstDatasheet = WaitForElementIsVisible(By.Id(FIRST_DATASHEET));
            }
            else
            {
                _firstDatasheet = WaitForElementIsVisible(By.XPath("//*[@id=\"tableListMenu\"]/tbody/tr/td[3]/a/b"));
            }
            _firstDatasheet.Click();
            WaitForLoad();

            return new DatasheetDetailsPage(_webDriver, _testContext);
        }

        public void MassiveDelete(string site,string guestType)
        {
            Actions action = new Actions(_webDriver);
            var extendMenu =  WaitForElementIsVisible(By.XPath(EXTEND_MENU));
            action.MoveToElement(extendMenu).Perform();
            var massiveDelete = WaitForElementIsVisible(By.XPath(MASSIVE_DELETE));
            massiveDelete.Click();

            ComboBoxSelectById(new ComboBoxOptions("SelectedSiteIds_ms", site, false));
            
            var guestBtn = WaitForElementIsVisible(By.XPath(GUEST_TYPE_BTN));
            guestBtn.Click();
            var guestToSelect = WaitForElementIsVisible(By.XPath(string.Format(MASSIVE_DELETE_GUEST_TYPE_FILTER, guestType)));
            guestToSelect.SetValue(ControlType.CheckBox, true);
            guestBtn.Click();

            var search = WaitForElementIsVisible(By.XPath(MASSIVE_DELETE_SEARCH));
            search.Click();
            WaitForLoad();
            var selectAll = WaitForElementIsVisible(By.XPath(MASSIVE_DELETE_SELECTALL));
            selectAll.Click();
            var delete = WaitForElementIsVisible(By.XPath(MASSIVE_DELETE_DELETE));
            delete.Click();
            var confirmDelete = WaitForElementIsVisible(By.XPath(MASSIVE_DELETE_CONFIRM_DELETE));
            confirmDelete.Click();

            var ok = WaitForElementIsVisible(By.XPath(OK_BTN_DEV));
            ok.Click();
        }

        public DatasheetMassiveDeletePopup OpenMassiveDeletePopup()
        {
            Actions action = new Actions(_webDriver);
            var extendMenu = WaitForElementIsVisible(By.XPath(EXTEND_MENU));
            action.MoveToElement(extendMenu).Perform();
            var massiveDelete = WaitForElementIsVisible(By.XPath(MASSIVE_DELETE));
            massiveDelete.Click();

            return new DatasheetMassiveDeletePopup(_webDriver, _testContext);
        }

        public bool VerifierMassiveDelete(string site, string guestType)
        {
            Filter(FilterType.Sites, site);
            Filter(FilterType.GuestTypes, guestType);

            if (CheckTotalNumber() == 0)
            {
                return true;
            }
            return false;
        }

        public FileInfo GetExportExcelFile(FileInfo[] taskFiles)
        {
            List<FileInfo> correctDownloadFiles = new List<FileInfo>();

            foreach (var file in taskFiles)
            {
                //  Test REGEX
                if (IsExcelFileCorrect(file.Name))
                {
                    correctDownloadFiles.Add(file);
                }
            }

            if (correctDownloadFiles.Count == 0)
            {
                return null;
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

        public bool IsExcelFileCorrect(string filePath)
        {
            string mois = "(?:0[1-9]|1[0-2])";         // mois
            //string space = "(\\s)";                    // Espace
            string annee = "\\d{4}";                   // annee YYYY
            string jour = "[0-3]\\d";                  // jour
            //string heure = "(?:0[0-9]|1[0-9]|2[0-3])"; // heure
            //string minutes = "[0-5]\\d";               // minutes
            //string secondes = "[0-5]\\d";              // secondes

            // Datasheet_2024-03-18.xlsx
            Regex r = new Regex("^Datasheet_" + annee + "-" + mois + "-" + jour + ".xlsx$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = r.Match(filePath);

            return m.Success;
        }
        public void Go_To_New_Navigate()
        {
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);
            WaitPageLoading();
        }
        public string GetServiceNameInFilterDataSheet()
        {
          _serviceNameFilter = WaitForElementIsVisible(By.Id(SERVICE_NAME_FILTER));
          return _serviceNameFilter.GetAttribute("value");                
        }
        public void EnterDatasheetName(string datasheetName)
        {
            WaitPageLoading();
            var datasheetNameInput = _webDriver.FindElement(By.XPath(DATASHEETNAME_FILTER)); 

            datasheetNameInput.Clear();

            datasheetNameInput.SendKeys(datasheetName);
        }

        public List<string> GetAllNameDatasheet()
        {
            List<string> result = new List<string>();          
           var elements = _webDriver.FindElements(By.XPath(DATASHEET_NAME1));

            foreach (var element in elements)
            {                 
                result.Add(element.Text);                 
            }
            return result;
        }
        public PrintReportPage PrintItemGenerique(SetupFilterAndFavoritesPage setupFilterAndFavoritesPage)
        {
            // Lancement du Print
            PrintReportPage reportPage;
            reportPage = setupFilterAndFavoritesPage.PrintDeliveryRoundResults();
            // le zip se télécharge directement, alors que le pdf s'ouvre dans un nouveau onglet
            var isReportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();
            setupFilterAndFavoritesPage.ClickPrintButton();

            //Assert
            Assert.IsTrue(isReportGenerated, "L'application n'a pas pu générer le fichier attendu.");


            return reportPage;
        }
        public int GetTotalDataSheets()
        {
            WaitPageLoading ();
            var numbers = WaitForElementIsVisible(By.XPath("/html/body/div[2]/div/div[2]/div/div/h1/span"));
            return Int32.Parse(numbers.Text);
        }
    }
}
