using DocumentFormat.OpenXml.Bibliography;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Menus;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Recipes;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.Extensions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Production.Dispatch
{
    public class DispatchPage : PageBase
    {
        public DispatchPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //__________________________________Constantes_____________________________________

        // General
        private const string EXTENDED_MENU = "//*[@id=\"tabContentItemContainer\"]/div[1]/div/div[2]/button";
        private const string SUPPLIER_ORDER_MENU = "//*[@id=\"tabContentItemContainer\"]/div[1]/div/div[2]/div/a[1]";
        private const string EXPORT = "exportBtn";
        private const string EXPORT_VALIDATION = "print-export";
        private const string PRINT = "Print turnover";
        private const string PRINT_VALIDATION = "printButton";

        private const string VALIDATE_BTN = "//*[@id=\"tabContentItemContainer\"]/div[1]/div/div[1]/button";
        private const string VALIDATE_ALL = "//*[@id=\"validate-all-btn\"]";
        private const string UNVALIDATE_ALL = "//*[@id=\"tabContentItemContainer\"]/div[1]/div/div[1]/div/a[2]";
        private const string VALIDATE_SUNDAY = "btnSun";

        private const string PREVIOUS_WEEK_FILTER = "//*[@id=\"tabContentItemContainer\"]/div[2]/h2/a[1]";
        private const string NEXT_WEEK_FILTER = "//*[@id=\"tabContentItemContainer\"]/div[2]/h2/a[2]";
        private const string CALENDAR = "calendar-dispatch";

        // Onglets
        private const string PREVISIONAL_QUANTITY = "tabPrevisional";
        private const string QUANTITY_TO_PRODUCE = "tabProduction";
        private const string QUANTITY_TO_INVOICE = "/html/body/div[2]/div/div[2]/div[3]/ul/li[3]/a";

        // Filtres
        private const string RESET_FILTER = "//*[@id=\"item-filter-form\"]/div[1]/a";
        private const string SEARCH = "SearchPattern";
        private const string SORTBY = "cbSortBy";
        private const string SITE = "SiteId";
        private const string SHOW_ALL = "//*[@id=\"ValidatedFilter\"][@value='ShowAll']";
        private const string VALIDATE_ONLY = "//*[@id=\"ValidatedFilter\"][@value='ShowValidatedOnly']";
        private const string NOT_VALIDATE_ONLY = "//*[@id=\"ValidatedFilter\"][@value='ShowNotValidatedOnly']";
        private const string CUSTOMER_TYPES_FILTER = "SelectedCustomerTypes_ms";
        private const string SEARCH_CUSTOMER_TYPE = "/html/body/div[10]/div/div/label/input";
        private const string SEARCH_CUSTOMER_TYPE1 = "/html/body/div[11]/div/div/label/input";
        private const string UNSELECT_CUSTOMER_TYPE = "/html/body/div[10]/div/ul/li[2]/a";
        private const string UNSELECT_CUSTOMER_TYPE1 = "/html/body/div[11]/div/ul/li[2]/a";
        private const string CUSTOMER_FILTER = "SelectedCustomers_ms";
        private const string SEARCH_CUSTOMER = "/html/body/div[11]/div/div/label/input";
        private const string SEARCH_CUSTOMER1 = "/html/body/div[12]/div/div/label/input";
        private const string UNSELECT_CUSTOMER = "/html/body/div[11]/div/ul/li[2]/a";
        private const string UNSELECT_CUSTOMER1 = "/html/body/div[12]/div/ul/li[2]/a/span[2]";
        private const string SERVICE_CATEGORIES_FILTER = "SelectedServiceCategories_ms";
        private const string SEARCH_SERVICE_CATEGORIES = "/html/body/div[12]/div/div/label/input";
        private const string SEARCH_SERVICE_CATEGORIES1 = "/html/body/div[12]/div/div/label/input";
        private const string UNSELECT_SERVICE_CATEGORIES = "/html/body/div[12]/div/ul/li[2]/a";
        private const string UNSELECT_SERVICE_CATEGORIES1 = "/html/body/div[12]/div/ul/li[2]/a";
        private const string DELIVERY_FILTER = "SelectedFlightDeliveries_ms";
        private const string SEARCH_DELIVERY = "/html/body/div[13]/div/div/label/input";
        private const string SEARCH_DELIVERY1 = "/html/body/div[14]/div/div/label/input";
        private const string UNSELECT_DELIVERIES = "/html/body/div[13]/div/ul/li[2]/a";
        private const string UNSELECT_DELIVERIES1 = "/html/body/div[14]/div/ul/li[2]/a";
        private const string DELIVERY_ROUND_FILTER = "SelectedDeliveryRounds_ms";
        private const string SEARCH_DELIVERY_ROUND = "/html/body/div[14]/div/div/label/input";
        private const string SEARCH_DELIVERY_ROUND1 = "/html/body/div[15]/div/div/label/input";
        private const string UNSELECT_DELIVERY_ROUND = "/html/body/div[14]/div/ul/li[2]/a";
        private const string UNSELECT_DELIVERY_ROUND1 = "/html/body/div[15]/div/ul/li[2]/a";

        // Tableau jours
        private const string PREVISIONAL_QUANTITY_ROW = "//*[@id=\"dispatchTable\"]/tbody/tr[*]/td[4]/a[contains(text(), '{0}')]/../..";
        private const string MONDAY_QTY = "//*[@id=\"dispatchTable\"]/tbody/tr[*]/td[4]/a[contains(text(), '{0}')]/../../td[6]/span[2]/input";
        private const string TUESDAY_QTY = "//*[@id=\"dispatchTable\"]/tbody/tr[*]/td[4]/a[contains(text(), '{0}')]/../../td[7]/span[2]/input";
        private const string WEDNESDAY_QTY = "//*[@id=\"dispatchTable\"]/tbody/tr[*]/td[4]/a[contains(text(), '{0}')]/../../td[8]/span[2]/input";
        private const string THURSDAY_QTY = "//*[@id=\"dispatchTable\"]/tbody/tr[*]/td[4]/a[contains(text(), '{0}')]/../../td[9]/span[2]/input";
        private const string FRIDAY_QTY = "//*[@id=\"dispatchTable\"]/tbody/tr[*]/td[4]/a[contains(text(), '{0}')]/../../td[10]/span[2]/input";
        private const string SATURDAY_QTY = "//*[@id=\"dispatchTable\"]/tbody/tr[*]/td[4]/a[contains(text(), '{0}')]/../../td[11]/span[2]/input";
        private const string SUNDAY_QTY = "//*[@id=\"dispatchTable\"]/tbody/tr[*]/td[4]/a[contains(text(), '{0}')]/../../td[12]/span[2]/input";

        private const string MENUS_COUNT = "//*[@id=\"dispatchTable\"]/tbody/tr[*]/td[4]/a[contains(text(), '{0}')]/../../td[5]/div/span";
        private const string MENUS_COUNT_BUTTON = "//div[@class='menu-count-many btn menus ']";
        private const string RELATED_MENUS_POPUP = "/html/body/div[6]/div/div/div[1]/h4[contains(text(), 'Related menus')]";
        private const string RELATED_MENU_NAME = "MenuFilter";
        private const string RELATED_MENU_DATE_SELECTED = "menu-day-display";
        private const string RELATED_MENU_RECIPE_NAME = "//p[contains(@class, 'recipe-name')]";
        private const string RELATED_MENU_RECIPE_EDIT_BUTTON = "//span[contains(@class, 'pencil')]";
        private const string RELATED_MENU_RECIPE_METHOD = "//select[contains(@class, 'methodQtySelector form-control')]";
        private const string RELATED_MENU_GO_TO_MENU_LINK = "hrefGoToMenu";

        private const string TABLE_INPUTS = "//*[@id=\"dispatchTable\"]/tbody/tr[{0}]/td[{1}]/span[1]/input";
        private const string CLICK_CUSTOMER = "//*[@id=\"SelectedCustomers_ms\"]/span[2]";
        private const string UNCHECKALL_CUSTOMER = "/html/body/div[12]/div/ul/li[2]/a/span[2]";
        private const string CLICK_DELIVERY = "//*[@id=\"SelectedFlightDeliveries_ms\"]/span[2]";
        private const string UNCHECKALL_DELIVERY = "/html/body/div[14]/div/ul/li[2]/a/span[2]";
        private const string LISTDELIVERY_PREVISIONALQTY = "//*[@id=\"dispatchTable\"]/tbody/tr[*]/td[3]";
        private const string LISTDELIVERY_QTY_TO_PRODUCE = "//*[@id=\"dispatchTable\"]/tbody/tr[*]/td[3]";
        private const string LISTDELIVERY_QTY_TO_INVOICE = "//*[@id=\"dispatchTable\"]/tbody/tr[*]/td[3]";

        //__________________________________ Variables ______________________________________

        // général
        [FindsBy(How = How.XPath, Using = EXTENDED_MENU)]
        private IWebElement _extendedBtn;

        [FindsBy(How = How.XPath, Using = SUPPLIER_ORDER_MENU)]
        private IWebElement _supplierOrderBtn;

        [FindsBy(How = How.Id, Using = EXPORT)]
        private IWebElement _export;

        [FindsBy(How = How.Id, Using = EXPORT_VALIDATION)]
        private IWebElement _exportValidation;

        [FindsBy(How = How.LinkText, Using = PRINT)]
        private IWebElement _print;

        [FindsBy(How = How.Id, Using = PRINT_VALIDATION)]
        private IWebElement _printValidation;

        [FindsBy(How = How.XPath, Using = VALIDATE_BTN)]
        private IWebElement _validateButton;

        [FindsBy(How = How.XPath, Using = VALIDATE_ALL)]
        private IWebElement _validateAll;

        [FindsBy(How = How.XPath, Using = UNVALIDATE_ALL)]
        private IWebElement _unValidate;


        [FindsBy(How = How.Id, Using = VALIDATE_SUNDAY)]
        private IWebElement _validateSunday;

        // Onglets
        [FindsBy(How = How.Id, Using = PREVISIONAL_QUANTITY)]
        private IWebElement _previsionalQuantity;

        [FindsBy(How = How.Id, Using = QUANTITY_TO_PRODUCE)]
        private IWebElement _quantityToProduce;

        [FindsBy(How = How.XPath, Using = QUANTITY_TO_INVOICE)]
        private IWebElement _quantityToInvoice;

        // Autres       

        [FindsBy(How = How.XPath, Using = NEXT_WEEK_FILTER)]
        private IWebElement _nextWeek;

        [FindsBy(How = How.XPath, Using = PREVIOUS_WEEK_FILTER)]
        private IWebElement _previousWeek;

        [FindsBy(How = How.Id, Using = CALENDAR)]
        private IWebElement _calendar;


        //__________________________________ Filtres ________________________________________

        [FindsBy(How = How.XPath, Using = RESET_FILTER)]
        private IWebElement _resetFilter;

        [FindsBy(How = How.Id, Using = SEARCH)]
        private IWebElement _search;

        [FindsBy(How = How.Id, Using = SORTBY)]
        private IWebElement _sortBy;

        [FindsBy(How = How.Id, Using = SITE)]
        private IWebElement _site;

        [FindsBy(How = How.XPath, Using = SHOW_ALL)]
        private IWebElement _showAll;

        [FindsBy(How = How.XPath, Using = VALIDATE_ONLY)]
        private IWebElement _validatedOnly;

        [FindsBy(How = How.XPath, Using = NOT_VALIDATE_ONLY)]
        private IWebElement _notValidatedOnly;

        [FindsBy(How = How.XPath, Using = CUSTOMER_TYPES_FILTER)]
        private IWebElement _customerTypesFilter;

        [FindsBy(How = How.XPath, Using = SEARCH_CUSTOMER_TYPE)]
        private IWebElement _searchCustomerType;

        [FindsBy(How = How.XPath, Using = UNSELECT_CUSTOMER_TYPE)]
        private IWebElement _unselectAllCustomerTypes;

        [FindsBy(How = How.XPath, Using = CUSTOMER_FILTER)]
        private IWebElement _customerFilter;

        [FindsBy(How = How.XPath, Using = SEARCH_CUSTOMER)]
        private IWebElement _searchCustomer;

        [FindsBy(How = How.XPath, Using = UNSELECT_CUSTOMER)]
        private IWebElement _unselectAllCustomers;

        [FindsBy(How = How.XPath, Using = SERVICE_CATEGORIES_FILTER)]
        private IWebElement _serviceCategoriesFilter;

        [FindsBy(How = How.XPath, Using = SEARCH_SERVICE_CATEGORIES)]
        private IWebElement _searchServiceCategories;

        [FindsBy(How = How.XPath, Using = UNSELECT_SERVICE_CATEGORIES)]
        private IWebElement _unselectAllServiceCategories;

        [FindsBy(How = How.XPath, Using = DELIVERY_FILTER)]
        private IWebElement _deliveryFilter;

        [FindsBy(How = How.XPath, Using = SEARCH_DELIVERY)]
        private IWebElement _searchDeliveries;

        [FindsBy(How = How.XPath, Using = UNSELECT_DELIVERIES)]
        private IWebElement _unselectAllDeliveries;

        [FindsBy(How = How.XPath, Using = DELIVERY_ROUND_FILTER)]
        private IWebElement _deliveryRoundFilter;

        [FindsBy(How = How.XPath, Using = SEARCH_DELIVERY_ROUND)]
        private IWebElement _searchDeliveryRound;

        [FindsBy(How = How.XPath, Using = UNSELECT_DELIVERY_ROUND)]
        private IWebElement _unselectAllDeliveryRound;

        // Tableau jours

        [FindsBy(How = How.XPath, Using = PREVISIONAL_QUANTITY_ROW)]
        private IWebElement _previsionalQuantityRow;

        [FindsBy(How = How.XPath, Using = MONDAY_QTY)]
        private IWebElement _monday;

        [FindsBy(How = How.XPath, Using = TUESDAY_QTY)]
        private IWebElement _tuesday;

        [FindsBy(How = How.XPath, Using = WEDNESDAY_QTY)]
        private IWebElement _wednesday;

        [FindsBy(How = How.XPath, Using = THURSDAY_QTY)]
        private IWebElement _thursday;

        [FindsBy(How = How.XPath, Using = FRIDAY_QTY)]
        private IWebElement _friday;

        [FindsBy(How = How.XPath, Using = SATURDAY_QTY)]
        private IWebElement _saturday;

        [FindsBy(How = How.XPath, Using = SUNDAY_QTY)]
        private IWebElement _sunday;

        // Related Menus

        [FindsBy(How = How.XPath, Using = MENUS_COUNT)]
        private IWebElement _menusCount;

        [FindsBy(How = How.XPath, Using = MENUS_COUNT_BUTTON)]
        private IWebElement _menusCountButton;

        [FindsBy(How = How.XPath, Using = MENUS_COUNT_BUTTON)]
        private IWebElement _relatedMenu;

        [FindsBy(How = How.Id, Using = RELATED_MENU_DATE_SELECTED)]
        private IWebElement _relatedMenuDate;

        [FindsBy(How = How.Id, Using = RELATED_MENU_RECIPE_EDIT_BUTTON)]
        private IWebElement _relatedMenuRecipeEditButton;

        [FindsBy(How = How.XPath, Using = RELATED_MENU_RECIPE_METHOD)]
        private IWebElement _relatedMenuRecipeMethod;

        [FindsBy(How = How.Id, Using = RELATED_MENU_GO_TO_MENU_LINK)]
        private IWebElement _goToMenu;

        //[FindsBy(How = How.XPath, Using = TABLE_INPUTS)]
        //private IWebElement _tableInput;
        public enum FilterType
        {
            Search,
            SortBy,
            Site,
            ShowAll,
            ValidateOnly,
            NotValidateOnly,
            CustomersTypes,
            Customers,
            ServiceCategories,
            Deliveries,
            DeliveryRounds
        }

        public void Filter(FilterType filterType, object value)
        {
            Actions action = new Actions(_webDriver);

            switch (filterType)
            {
                case FilterType.Search:
                    _search = WaitForElementIsVisible(By.Id(SEARCH));
                    _search.SetValue(ControlType.TextBox, value);
                    break;
                case FilterType.SortBy:
                    _sortBy = WaitForElementIsVisible(By.Id(SORTBY));
                    _sortBy.Click();
                    var element = WaitForElementIsVisible(By.XPath("//option[contains(@value,'" + value + "')]"));
                    _sortBy.SetValue(ControlType.DropDownList, element.Text);
                    _sortBy.Click();
                    break;
                case FilterType.Site:
                    _site = WaitForElementIsVisible(By.Id(SITE));
                    _site.SetValue(ControlType.DropDownList, value);
                    break;
                case FilterType.ShowAll:
                    if (isElementVisible(By.Id("ValidatedFilterAll")))
                    {
                        var showAll = WaitForElementExists(By.Id("ValidatedFilterAll"));
                        showAll.SetValue(ControlType.RadioButton, value);
                    }
                    else
                    {
                        _showAll = WaitForElementExists(By.XPath(SHOW_ALL));
                        _showAll.SetValue(ControlType.RadioButton, value);
                    }

                    break;
                case FilterType.ValidateOnly:
                    if (isElementVisible(By.Id("ValidatedFilterValidated")))
                    {
                        var validatedOnly = WaitForElementExists(By.Id("ValidatedFilterValidated"));
                        validatedOnly.SetValue(ControlType.RadioButton, value);

                    }
                    else
                    {
                        _validatedOnly = WaitForElementExists(By.XPath(VALIDATE_ONLY));
                        _validatedOnly.SetValue(ControlType.RadioButton, value);
                    }

                    break;
                case FilterType.NotValidateOnly:

                    if (isElementVisible(By.Id("ValidatedFilterNotValidated")))
                    {
                        var notValidatedOnly = WaitForElementExists(By.Id("ValidatedFilterNotValidated"));
                        notValidatedOnly.SetValue(ControlType.RadioButton, value);

                    }
                    else
                    {
                        _notValidatedOnly = WaitForElementExists(By.XPath(NOT_VALIDATE_ONLY));
                        _notValidatedOnly.SetValue(ControlType.RadioButton, value);
                    }

                    break;
                case FilterType.CustomersTypes:
                    _customerTypesFilter = WaitForElementIsVisible(By.Id(CUSTOMER_TYPES_FILTER));
                    _customerTypesFilter.Click();

                    // On décoche toutes les options
                    _unselectAllCustomerTypes = WaitForElementIsVisible(By.XPath(UNSELECT_CUSTOMER_TYPE1));
                    _unselectAllCustomerTypes.Click();

                    _searchCustomerType = WaitForElementIsVisible(By.XPath(SEARCH_CUSTOMER_TYPE1));
                    _searchCustomerType.SetValue(ControlType.TextBox, value);
                    WaitForLoad();

                    var valueToCheck = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    valueToCheck.SetValue(ControlType.CheckBox, true);

                    _customerTypesFilter.Click();
                    break;
                case FilterType.Customers:
                    ComboBoxSelectById(new ComboBoxOptions(CUSTOMER_FILTER, (string)value));
                    break;
                case FilterType.ServiceCategories:
                    ComboBoxSelectById(new ComboBoxOptions(SERVICE_CATEGORIES_FILTER, (string)value));

                    break;
                case FilterType.Deliveries:
                    ComboBoxSelectById(new ComboBoxOptions(DELIVERY_FILTER, (string)value));

                    break;
                case FilterType.DeliveryRounds:
                    ComboBoxSelectById(new ComboBoxOptions(DELIVERY_ROUND_FILTER, (string)value));

                    break;

                default:
                    break;
            }

            WaitPageLoading();
            Thread.Sleep(2000);
        }

        public void ResetFilter()
        {
            _resetFilter = WaitForElementIsVisible(By.XPath(RESET_FILTER));
            _resetFilter.Click();
            WaitForLoad();
            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                // pas de date
            }
        }

        // _____________________________________________ Méthodes _____________________________________________________

        public override void ShowExtendedMenu()
        {
            _extendedBtn = WaitForElementExists(By.XPath(EXTENDED_MENU));
            var actions = new Actions(_webDriver);
            actions.MoveToElement(_extendedBtn).Perform();
        }

        public void ShowExtendedMenu_Dispatch()
        {
            _extendedBtn = WaitForElementExists(By.XPath("//*[@id=\"tabContentItemContainer\"]/div/div[1]/div/div[2]/button"));
            var actions = new Actions(_webDriver);
            actions.MoveToElement(_extendedBtn).Perform();
        }



        public SupplierOrderModal ClickGenerateSupplierOrder()
        {

            ShowExtendedMenu();

            _supplierOrderBtn = WaitForElementIsVisible(By.XPath(SUPPLIER_ORDER_MENU));
            _supplierOrderBtn.Click();
            WaitForLoad();

            return new SupplierOrderModal(_webDriver, _testContext); ;
        }

        public void ExportExcel(bool versionPrint)
        {
            ShowExtendedMenu();

            _export = WaitForElementIsVisible(By.Id(EXPORT));
            _export.Click();
            WaitForLoad();

            _exportValidation = WaitForElementIsVisible(By.Id(EXPORT_VALIDATION));
            _exportValidation.Click();
            WaitForLoad();

            if (versionPrint)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));
            }

            WaitForDownload();
            Close();
        }
        public void ExportExcel_Dispatch(bool versionPrint)
        {
            ShowExtendedMenu_Dispatch();

            _export = WaitForElementIsVisible(By.Id(EXPORT));
            _export.Click();
            WaitForLoad();

            _exportValidation = WaitForElementIsVisible(By.Id(EXPORT_VALIDATION));
            _exportValidation.Click();
            WaitForLoad();

            if (versionPrint)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));
            }

            WaitForDownload();
            Close();
        }

        public FileInfo GetExcelFile(FileInfo[] taskFiles)
        {
            List<FileInfo> correctDownloadFiles = new List<FileInfo>();

            foreach (var file in taskFiles)
            {
                if (IsExcelFileCorrect(file.Name))
                {
                    correctDownloadFiles.Add(file);
                }
            }

            if (correctDownloadFiles.Count <= 0)
            {
                throw new Exception(MessageErreur.FICHIER_NON_TROUVE);
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
            string annee = "\\d{4}";                   // annee YYYY
            string jour = "[0-3]\\d";                  // jour
            string heure = "(?:0[0-9]|1[0-9]|2[0-3])"; // heure
            string minutes = "[0-5]\\d";               // minutes
            string secondes = "[0-5]\\d";              // minutes

            Regex reg = new Regex("^Dispatch_Export" + "_" + annee + "-" + mois + "-" + jour + " " + heure + "-" + minutes + "-" + secondes + ".xlsx$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = reg.Match(filePath);

            return m.Success;
        }

        public PrintReportPage PrintTurnover(bool printValue)
        {

            ShowExtendedMenu();

            _print = WaitForElementIsVisible(By.LinkText(PRINT));
            _print.Click();
            WaitForLoad();

            _printValidation = WaitForElementIsVisible(By.Id(PRINT_VALIDATION));
            _printValidation.Click();
            WaitForLoad();

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

        public override void ShowValidationMenu()
        {
            _validateButton = WaitForElementExists(By.XPath(VALIDATE_BTN));
            var actions = new Actions(_webDriver);
            actions.MoveToElement(_validateButton).Perform();
        }

        public void ValidateAll()
        {
            ShowValidationMenu();
            WaitForLoad();
            WaitPageLoading();
            _validateAll = WaitForElementIsVisible(By.XPath(VALIDATE_ALL));
            _validateAll.Click();
            WaitPageLoading();
        }

        public void UnValidateAll()
        {
            ShowValidationMenu();
            WaitForLoad();
            WaitPageLoading();
            _unValidate = WaitForElementIsVisible(By.XPath(UNVALIDATE_ALL));
            WaitPageLoading();
            _unValidate.Click();
            WaitForLoad();
        }

        public void ValidateSunday()
        {
            ShowValidationMenu();

            _validateSunday = WaitForElementIsVisible(By.Id(VALIDATE_SUNDAY));
            _validateSunday.Click();
            WaitForLoad();
        }

        // Onglets
        public PrevisionalQtyPage ClickPrevisonalQuantity()
        {
            _previsionalQuantity = WaitForElementIsVisible(By.Id(PREVISIONAL_QUANTITY));
            _previsionalQuantity.Click();
            WaitPageLoading();
            WaitForLoad();

            return new PrevisionalQtyPage(_webDriver, _testContext);
        }
        public void AddQuantityOnPrevisonalQuantity(string service, string quantity)
        {
            _previsionalQuantityRow = WaitForElementIsVisible(By.XPath(string.Format(PREVISIONAL_QUANTITY_ROW, service)));
            _previsionalQuantityRow.Click();

            int today = (int)DateUtils.Now.DayOfWeek;
            if (today == 0)
            {
                today = 7;
            }
            WaitPageLoading();
            if (isElementVisible(By.Id("close-btn")))
            {
                WaitPageLoading();
                var closeBtn = WaitForElementIsVisible(By.Id("close-btn"));
                closeBtn.Click();
                WaitPageLoading();
            }
            _monday = WaitForElementIsVisible(By.XPath(string.Format(MONDAY_QTY, service)));
            if (_monday.Enabled)
            {
                WaitPageLoading();
                _monday.SetValue(ControlType.TextBox, quantity);
            }
            else if (today <= 1)
            {
                throw new Exception("Monday , previous week ?");
            }

            _tuesday = WaitForElementIsVisible(By.XPath(string.Format(TUESDAY_QTY, service)));
            if (_tuesday.Enabled)
            {
                WaitPageLoading();
                _tuesday.SetValue(ControlType.TextBox, quantity);
            }
            else if (today <= 2)
            {
                throw new Exception("Tuesday , previous week ?");
            }
            WaitPageLoading();
            _wednesday = WaitForElementIsVisible(By.XPath(string.Format(WEDNESDAY_QTY, service)));
            if (_wednesday.Enabled)
            {

                WaitPageLoading();
                _wednesday.SetValue(ControlType.TextBox, quantity);
            }
            else if (today <= 3)
            {
                throw new Exception("Wednesday , previous week ?");
            }

            _thursday = WaitForElementIsVisible(By.XPath(string.Format(THURSDAY_QTY, service)));
            if (_thursday.Enabled)
            {
                WaitPageLoading();
                _thursday.SetValue(ControlType.TextBox, quantity);
            }
            else if (today <= 4)
            {
                throw new Exception("Thursday , previous week ?");
            }

            _friday = WaitForElementIsVisible(By.XPath(string.Format(FRIDAY_QTY, service)));
            if (_friday.Enabled)
            {
                WaitPageLoading();
                _friday.SetValue(ControlType.TextBox, quantity);
            }
            else if (today <= 5)
            {
                throw new Exception("Friday , previous week ?");
            }

            _saturday = WaitForElementIsVisible(By.XPath(string.Format(SATURDAY_QTY, service)));
            if (_saturday.Enabled)
            {
                WaitPageLoading();
                _saturday.SetValue(ControlType.TextBox, quantity);
            }
            else if (today <= 6)
            {
                throw new Exception("Saturday , previous week ?");
            }

            _sunday = WaitForElementIsVisible(By.XPath(string.Format(SUNDAY_QTY, service)));
            if (_sunday.Enabled)
            {
                WaitPageLoading();
                _sunday.SetValue(ControlType.TextBox, quantity);
            }
            else if (today <= 7)
            {
                throw new Exception("Sunday , previous week ?");
            }
            WaitLoading();
            WaitPageLoading();
            //Thread.Sleep(2000);
        }

        public string GetNumberMenusAssociated(string service)
        {
            WaitPageLoading();
            _previsionalQuantityRow = WaitForElementIsVisible(By.XPath(string.Format(PREVISIONAL_QUANTITY_ROW, service)));
            WaitForLoad();
            _previsionalQuantityRow.Click();
            WaitPageLoading();
            _menusCount = WaitForElementIsVisible(By.XPath(string.Format(MENUS_COUNT, service)));

            if (isElementVisible(By.Id("close-btn")))
            {
                WaitPageLoading();
                var closeBtn = WaitForElementIsVisible(By.Id("close-btn"));
                closeBtn.Click();
                WaitPageLoading();
            }
            return _menusCount.Text;

        }
        public bool GoToRelatedMenusPopup()
        {
            _menusCountButton = WaitForElementIsVisible(By.XPath(MENUS_COUNT_BUTTON), nameof(MENUS_COUNT_BUTTON));
            _menusCountButton.Click();
            WaitPageLoading();

            if (isElementVisible(By.XPath(RELATED_MENUS_POPUP)))
            {
                return true;
            }
            return false;
        }

        public string GetRelatedMenu()
        {
            _relatedMenu = WaitForElementIsVisible(By.Id(RELATED_MENU_NAME));
            return _relatedMenu.Text;
        }

        public void ClickOnRelatedMenuOfToday()
        {
            var dayButton = WaitForElementIsVisible(By.Id(DateUtils.Now.DayOfWeek.ToString()));
            dayButton.Click();
            WaitForLoad();
        }
        public string GetRelatedMenuDate()
        {
            _relatedMenuDate = WaitForElementIsVisible(By.Id(RELATED_MENU_DATE_SELECTED), nameof(RELATED_MENU_DATE_SELECTED));
            return _relatedMenuDate.Text;
        }
        public string GetRelatedMenuRecipeName()
        {
            var relatedMenuRecipesName = _webDriver.FindElements(By.XPath(RELATED_MENU_RECIPE_NAME));
            return relatedMenuRecipesName[1].Text;
        }

        public string CheckRelatedMenuRecipe()
        {
            var relatedMenuRecipesName = _webDriver.FindElements(By.XPath(RELATED_MENU_RECIPE_NAME));
            relatedMenuRecipesName[1].Click();

            _relatedMenuRecipeEditButton = WaitForElementIsVisible(By.XPath(RELATED_MENU_RECIPE_EDIT_BUTTON), nameof(RELATED_MENU_RECIPE_EDIT_BUTTON));
            _relatedMenuRecipeEditButton.Click();
            WaitForLoad();

            //nouvel onglet
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles.Last());

            var recipeVariantPage = new RecipesVariantPage(_webDriver, _testContext);
            var recipeName = recipeVariantPage.GetRecipeName();

            recipeVariantPage.Close();

            return recipeName;
        }
        public void SetRelatedMenuRecipeMethod(string method)
        {
            WaitForLoad();
            _relatedMenuRecipeMethod = WaitForElementIsVisible(By.XPath(RELATED_MENU_RECIPE_METHOD));
            _relatedMenuRecipeMethod.SetValue(ControlType.DropDownList, method);
            WaitForLoad();//pas suiffisant
            Thread.Sleep(1000);
        }
        public MenusDayViewPage GoToRelatedMenuLink()
        {
            _goToMenu = WaitForElementIsVisible(By.Id(RELATED_MENU_GO_TO_MENU_LINK));
            _goToMenu.Click();
            WaitForLoad();

            //nouvel onglet
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles.Last());

            return new MenusDayViewPage(_webDriver, _testContext);
        }

        public QuantityToProducePage ClickQuantityToProduce()
        {
            _quantityToProduce = WaitForElementIsVisible(By.Id(QUANTITY_TO_PRODUCE));
            _quantityToProduce.Click();
            WaitPageLoading();
            WaitForLoad();

            return new QuantityToProducePage(_webDriver, _testContext); ;
        }

        public QuantityToInvoicePage ClickQuantityToInvoice()
        {
            _quantityToInvoice = WaitForElementIsVisible(By.XPath(QUANTITY_TO_INVOICE));
            _quantityToInvoice.Click();
            WaitPageLoading();
            WaitForLoad();//pas suffisant
            return new QuantityToInvoicePage(_webDriver, _testContext); ;
        }

        // Autres

        public void ClickNextWeek()
        {
            _nextWeek = WaitForElementIsVisible(By.XPath(NEXT_WEEK_FILTER));
            _nextWeek.Click();
            WaitForLoad();
            WaitPageLoading();
        }

        public void ClickPreviousWeek()
        {
            _previousWeek = WaitForElementIsVisible(By.XPath(PREVIOUS_WEEK_FILTER));
            _previousWeek.Click();
            WaitForLoad();
        }

        public string GetDateTimeFormat()
        {
            _calendar = _webDriver.FindElement(By.Id(CALENDAR));
            return _calendar.GetAttribute("data-date-format");
        }

        public ServicePricePage ShowService()
        {
            var linkService = WaitForElementIsVisible(By.XPath("//*[@id='dispatchTable']/tbody/tr[2]/td[4]/a"));
            linkService.Click();
            // nouveau onglet ouvert !!!
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);
            WaitForLoad();

            return new ServicePricePage(_webDriver, _testContext);

        }
        public void ClosePrintPopoverIfVisible()
        {
            var popoverPrintVisible = isElementVisible(By.XPath("//*[contains(@id,'popover')]"));
            if (popoverPrintVisible)
            {
                var btn = WaitForElementIsVisible(By.Id("header-print-button"));
                btn.Click();
            }
        }
        public void Import(string excelPath)
        {
            Actions actions = new Actions(_webDriver);
            ClosePrintPopoverIfVisible();
            ShowExtendedMenu_Dispatch();
            _export = WaitForElementIsVisible(By.Id("importBtn"));
            _export.Click();
            WaitForLoad();
            IWebElement _fileToImport = WaitForElementIsVisible(By.XPath("//*[@id=\"fileSent\"]"));
            _fileToImport.SendKeys(excelPath);
            IWebElement _checkFile = WaitForElementIsVisible(By.XPath("//*[@id=\"ImportFileForm\"]/div[3]/button[2]"));
            _checkFile.Click();
            WaitForLoad();


        }
        public bool VerifyImportFile()
        {
            IWebElement _import = WaitForElementIsVisible(By.XPath("//*[@id=\"ImportFileForm\"]/div[1]/h4"));
            string import_text = _import.Text;
            bool import_text_equals = import_text.ToString().Equals("Import : Loading Excel sheets");
            return import_text_equals;
        }

        public void ClickOnFirstInput()
        {
            WaitForLoad();
            var tableInput = WaitForElementIsVisible(By.XPath(string.Format(TABLE_INPUTS, 4, 6)));
            tableInput.Click();
            WaitForLoad();
        }
        public bool ArrowKeyboardClick(string arrow)
        {
            try
            {
                IWebElement focusedElement = _webDriver.SwitchTo().ActiveElement();
                Actions action = new Actions(_webDriver);
                switch (arrow)
                {
                    case "left":
                        action.SendKeys(Keys.ArrowLeft).Perform();
                        break;
                    case "right":
                        action.SendKeys(Keys.ArrowRight).Perform();
                        break;
                    case "up":
                        action.SendKeys(Keys.ArrowUp).Perform();
                        break;
                    case "down":
                        action.SendKeys(Keys.ArrowDown).Perform();
                        break;
                    default:
                        throw new Exception("Erreur lors de la navigation avec les touches du clavier");
                }
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Erreur lors de la navigation avec les touches du clavier: " + ex.Message);
                return false;
            }
        }
        public List<string> GetTotalListDeliveryPrevisionalQty()
        {
            List<string> liste = new List<string>();
            var dispatchList = _webDriver.FindElements(By.XPath(LISTDELIVERY_PREVISIONALQTY)).ToList();
            WaitLoading();
            liste = dispatchList.Select(x => x.Text).ToList();

            return liste;

        }
        public List<string> GetTotalListDeliveryQtyToProduce()
        {
            List<string> liste = new List<string>();
            var dispatchListQtyToProduce = _webDriver.FindElements(By.XPath(LISTDELIVERY_QTY_TO_PRODUCE)).ToList();
            WaitLoading();
            liste = dispatchListQtyToProduce.Select(x => x.Text).ToList();
            //foreach (var tableDeliveryQtyToProduce in dispatchListQtyToProduce)
            //{
            //    liste.Add(tableDeliveryQtyToProduce.Text);
            //}
            return liste;

        }
        public List<string> GetTotalListDeliveryQtyToInvoice()
        {
            List<string> liste = new List<string>();
            var dispatchListQtyToInvoice = _webDriver.FindElements(By.XPath(LISTDELIVERY_QTY_TO_INVOICE)).ToList();
            WaitLoading();
            liste = dispatchListQtyToInvoice.Select(x => x.Text).ToList();
            //foreach (var tableDeliveryQtyToInvoice in dispatchListQtyToInvoice)
            //{
            //    WaitLoading();
            //    liste.Add(tableDeliveryQtyToInvoice.Text);
            //}
            return liste;

        }
        public bool VerifyDispatchInDeliveries(List<string> deliverys, List<string> listOfDispatch)
        {
            foreach (var dispatch in listOfDispatch)
            {
                if (!deliverys.Contains(dispatch))
                {
                    return false;
                }
            }
            return true;
        }
        public void FilterCustomersUnCheckAll()
        {
            var customers = WaitForElementIsVisible(By.XPath(CLICK_CUSTOMER));
            customers.Click();
            var uncheckAll = WaitForElementIsVisible(By.XPath(UNCHECKALL_CUSTOMER));
            uncheckAll.Click();
            customers.Click();
        }
        public void FilterDeliverysUnCheckAll()
        {
            var delivery = WaitForElementIsVisible(By.XPath(CLICK_DELIVERY));
            delivery.Click();
            var uncheckAll = WaitForElementIsVisible(By.XPath(UNCHECKALL_DELIVERY));
            uncheckAll.Click();
            delivery.Click();
        }
        public void FilterCustomers(List<string> customerNames)
        {
            // Parcourt chaque nom dans la liste de customers
            foreach (var customerName in customerNames.Distinct())
            {
                ComboBoxSelectMultipleById(new ComboBoxOptions(CUSTOMER_FILTER, customerName, false));
            }
        }
        public void FilterDeliverys(List<string> deliverys)
        {
            // Parcourt chaque nom dans la liste de customers
            foreach (var deleverie in deliverys.Distinct())
            {
                ComboBoxSelectMultipleById(new ComboBoxOptions(DELIVERY_FILTER, deleverie, false));
            }
        }
        public void ComboBoxSelectMultipleById(ComboBoxOptions cbOpt)
        {
            // Sélectionne le champ du filtre
            IWebElement input = WaitForElementIsVisible(By.Id(cbOpt.XpathId));
            input.Click();
            WaitForLoad();

            bool selectionWasModified = false;

            if (cbOpt.SelectionValue != null)
            {
                var searchVisible = SolveVisible("//*/input[@type='search']");
                Assert.IsNotNull(searchVisible);

                // Entrez le nom du client dans le champ de recherche
                _webDriver.ExecuteJavaScript("arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('input'));", searchVisible, cbOpt.SelectionValue);

                Thread.Sleep(1000);
                WaitForLoad();

                // Vérifie si l'élément est déjà coché
                var select = SolveVisible("//*/label[contains(@for, 'ui-multiselect')]/span[contains(text(),'" + cbOpt.SelectionValue + "')]");
                Assert.IsNotNull(select, "Pas de sélection de " + cbOpt.SelectionValue);

                // Vérifie l'état de la case (déjà cochée ou non)
                bool isAlreadyChecked = select.GetAttribute("class").Contains("checked");

                if (!isAlreadyChecked)
                {
                    select.Click();
                    selectionWasModified = true;
                }
            }

            if (selectionWasModified)
            {
                if (cbOpt.IsUsedInFilter)
                {
                    WaitPageLoading();
                    WaitForLoad();
                }
            }

            input = WaitForElementIsVisible(By.Id(cbOpt.XpathId));

            try
            {
                input.SendKeys(Keys.Enter);
            }
            catch
            {
                input.Click();
            }

            WaitForLoad();
        }

        public void FilterDeliverysUnCheckAllDelivery()
        {
            var delivery = WaitForElementIsVisible(By.XPath(CLICK_DELIVERY));
            delivery.Click();
            var uncheckAll = WaitForElementIsVisible(By.XPath("/html/body/div[15]/div/ul/li[2]/a/span[2]"));
            uncheckAll.Click();
            delivery.Click();
        }
        public void FilterCustomersUnCheckAllCustomer()
        {
            var customers = WaitForElementIsVisible(By.XPath(CLICK_CUSTOMER));
            customers.Click();
            var uncheckAll = WaitForElementIsVisible(By.XPath("/html/body/div[14]/div/ul/li[2]/a/span[2]"));
            uncheckAll.Click();
            customers.Click();
        }
    }
}
