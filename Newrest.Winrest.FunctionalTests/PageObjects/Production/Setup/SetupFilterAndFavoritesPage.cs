using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System;
using OpenQA.Selenium.Support.UI;
using System.Collections.Generic;
using DocumentFormat.OpenXml.ExtendedProperties;
using UglyToad.PdfPig.Content;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.ProductionManagement;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Production.Setup
{
    public class SetupFilterAndFavoritesPage : PageBase
    {

        public SetupFilterAndFavoritesPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //__________________________________ Constantes _____________________________________
        private const string BTN_VALIDATE = "btnValidate";

        // Filtres
        private const string RESET_FILTER = "//a[contains (text(), 'Reset Filter')]";
        private const string SEARCH_DELIVERY_ROUND_NAME_INPUT = "SearchPattern";
        private const string SETUP_FILTER_DATE = "StartDate";
        private const string SITES = "SiteId";
        private const string SEARCH_RECIPE_NAME_INPUT = "SearchRecipeName";
        private const string CUSTOMER_TYPES_FILTER = "SelectedCustomerTypesIds_ms";
        private const string CUSTOMER_TYPES_FILTER_UNCHECK_ALL = "/html/body/div[10]/div/ul/li[2]/a/span[2]";
        private const string CUSTOMER_TYPES_FILTER_SEARCH = "/html/body/div[10]/div/div/label/input";
        private const string CUSTOMERS_FILTER = "SelectedCustomersIds_ms";
        private const string CUSTOMERS_FILTER_UNCHECK_ALL = "/html/body/div[13]/div/ul/li[2]/a/span[2]";
        private const string CUSTOMERS_FILTER_SEARCH = "/html/body/div[13]/div/div/label/input";
        private const string DELIVERIES_FILTER = "SelectedDeliveriesIds_ms";
        private const string DELIVERIES_FILTER_UNCHECK_ALL = "/html/body/div[14]/div/ul/li[2]/a/span[2]";
        private const string DELIVERIES_FILTER_SEARCH = "/html/body/div[14]/div/div/label/input";
        private const string MEALTYPES_FILTER = "//*[@id=\"SelectedMealTypesIds_ms\"]";
        private const string MEALTYPES_FILTER_UNCHECK_ALL = "/html/body/div[12]/div/ul/li[2]/a/span[2]";
        private const string MEALTYPES_FILTER_SEARCH = "/html/body/div[12]/div/div/label/input";
        private const string WORKSHOPS_FILTER = "//*[@id=\"SelectedWorkshopsIds_ms\"]";
        private const string WORKSHOPS_FILTER_UNCHECK_ALL = "/html/body/div[13]/div/ul/li[2]/a/span[2]";
        private const string WORKSHOPS_FILTER_SEARCH = "/html/body/div[13]/div/div/label/input";
        private const string DONE_BUTTON = "/html/body/div[2]/div/div[1]/div/form/div/div[4]/div/input";
        private const string DELIVERIES_FILTER_LIST = "/html/body/div[2]/div/div[2]/div[2]/div/div/div[2]/div[*]/div[1]/div/div[2]/table/tbody/tr/td";
        private const string DELIVERY_MENU_FILTER = "/html/body/div[2]/div/div[2]/div[1]/div/div/button";
        private const string DELIVERY_NOTE_BY_RECIPES = "/html/body/div[2]/div/div[2]/div[1]/div/div/div/a[4]";
        private const string DELIVERY_NOTE_VALORIZED = "/html/body/div[2]/div/div[2]/div[1]/div/div/div/a[7]";
        private const string DELIVERY_NOTE = "//*/a[text()='Delivery Notes (singles)']";
        private const string DELIVERY_ROUND = "/html/body/div[2]/div/div[2]/div[1]/div/div/div/a[2]";
        private const string DELIVERY_NOTE_BY_SERVICE = "/html/body/div[2]/div/div[2]/div[1]/div/div/div/a[6]";
        private const string MODAL_PRINT_SETUP = "/html/body/div[3]/div/div";
        private const string VALIDATE_BUTTON = "/html/body/div[3]/div/div/div[3]/button[2]";
        private const string PRINT = "btn-printreportpopup";
        private const string FOOD_PACK_REPORT = "/html/body/div[2]/div/div[2]/div[1]/div/div/div/a[8]";
        private const string DELIVERYROUND = "printSelection_10";
        private const string DELIVERY_ROUND_SINGLE = "printSelection_21";
        private const string DELIVERYROUNDBYRECIPES = "printSelection_12";
        private const string DELIVERYBUTTON = "SelectedDeliveriesIds_ms";
        private const string DELIVERYSELEMENTS = "label.ui-corner-all";
        private const string TEXTNOTHINGSELECTED = "span.ui-multiselect-open + span";
        private const string SELECT_DELIVERY_ROUND = "deliveryRound-filter";
        private const string OPTION_DELIVERY = "/html/body/div[2]/div/div[2]/div[2]/div/div/div[1]/select/option[2]";
        private const string PROVIDERS_NAMES = "//*[@id=\"tableListMenu\"]/thead/tr/th[2]";
        private const string FILTER_NAME = "//*[@id=\"favorite-filter-form\"]/div[2]/div/div/div/span";
        private const string NAME = "Name";
        private const string MAKE_FAVORITE = "//*[@id=\"btnFilter\"]/div/a[2]";
        private const string FAVORITE_NAME = "Name";
        private const string SAVE_FAVORITE = "last";
        private const string FAVORITE_SEARCH = "//*[@id=\"searchPattern\"]";
        private const string FAVORITE = "//*[@id=\"favorite-filter-form\"]/div[2]/div/div/div[@class='raw-favorite clickable-favorite']/span[text()='{0}']";
        private const string FAVORITE_EDIT_FAV = "//*[@id=\"favorite-filter-form\"]/div[2]/div/div/div[1]/a[1]/span";
        private const string FAVORITE_FILTER = "//*[@id=\"favorite-filter-form\"]/div[2]/div/div/div/span";
        private const string DELETE_FAVORITE = "//*[@id=\"favorite-filter-form\"]/div[2]/div/div/div[@class='raw-favorite clickable-favorite']/span[text()='{0}']/../button/span";
        private const string DELETE_FAVORITECROSS = "//*[@id=\"favorite-filter-form\"]/div[2]/div/div/div[@class='raw-favorite clickable-favorite']/span[text()='{0}']/../button";
        private const string CONFIRM_DELETE_FAVORITE = "dataConfirmOK";
        private const string SELECTED_FAVORITE = "//*[@id=\"favorite-filter-form\"]/div[2]/div/div/div/span";
        private const string FAVORITE_RESULT_FAV = "//*[@id=\"favorite-filter-form\"]/div[2]/div/div/div[1]/a[2]/span";
        private const string FAVORITE_RES_FILTRE = "//span[contains(text(),'{0}')]";
        //__________________________________ Variables _____________________________________
        [FindsBy(How = How.Id, Using = BTN_VALIDATE)]
        private IWebElement _btnValidate;
        [FindsBy(How = How.Id, Using = DELIVERYBUTTON)]
        private IWebElement _deliveryButton;
        [FindsBy(How = How.CssSelector, Using = DELIVERYSELEMENTS)]
        private IWebElement _deliverysElement;
        [FindsBy(How = How.CssSelector, Using = TEXTNOTHINGSELECTED)]
        private IWebElement _textNothingSelected;
        //Filters
        [FindsBy(How = How.XPath, Using = RESET_FILTER)]
        private IWebElement _resetFilter;

        [FindsBy(How = How.Id, Using = SEARCH_DELIVERY_ROUND_NAME_INPUT)]
        private IWebElement _searchDeliveryRoundNameInput;

        [FindsBy(How = How.Id, Using = SETUP_FILTER_DATE)]
        private IWebElement _setupFilterDate;

        [FindsBy(How = How.Id, Using = SITES)]
        private IWebElement _sites;

        [FindsBy(How = How.Id, Using = SEARCH_RECIPE_NAME_INPUT)]
        private IWebElement _searchRecipeNameInput;

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

        [FindsBy(How = How.Id, Using = DELIVERIES_FILTER)]
        private IWebElement _deliveriesFilter;

        [FindsBy(How = How.XPath, Using = DELIVERIES_FILTER_UNCHECK_ALL)]
        private IWebElement _deliveriesFilterUncheckAll;

        [FindsBy(How = How.XPath, Using = DELIVERIES_FILTER_SEARCH)]
        private IWebElement _deliveriesFilterSearch;

        [FindsBy(How = How.XPath, Using = MEALTYPES_FILTER)]
        private IWebElement _mealtypesFilter;

        [FindsBy(How = How.XPath, Using = MEALTYPES_FILTER_UNCHECK_ALL)]
        private IWebElement _mealtypesFilterUncheckAll;

        [FindsBy(How = How.XPath, Using = MEALTYPES_FILTER_SEARCH)]
        private IWebElement _mealtypesFilterSearch;

        [FindsBy(How = How.XPath, Using = WORKSHOPS_FILTER)]
        private IWebElement _workshopsFilter;

        [FindsBy(How = How.XPath, Using = WORKSHOPS_FILTER_UNCHECK_ALL)]
        private IWebElement _workshopsFilterUncheckAll;

        [FindsBy(How = How.XPath, Using = WORKSHOPS_FILTER_SEARCH)]
        private IWebElement _workshopsFilterSearch;

        [FindsBy(How = How.Id, Using = PRINT)]
        private IWebElement _print;

        [FindsBy(How = How.XPath, Using = FILTER_NAME)]
        private IWebElement _nameFilter;

        [FindsBy(How = How.Id, Using = NAME)]
        private IWebElement _name;

        [FindsBy(How = How.XPath, Using = FAVORITE_EDIT_FAV)]
        private IWebElement _favorite_edit;

        [FindsBy(How = How.XPath, Using = FAVORITE_RESULT_FAV)]
        private IWebElement _favorite_result;
        //__________________________________ Methods _____________________________________

        // FILTRES

        public enum FilterType
        {
            SearchDeliveryRoundName,
            StartDate,
            Sites,
            RecipeName,
            CustomerTypes,
            Customers,
            Deliveries,
            MealTypes,
            Workshops
        }
        public void ResetFilters()
        {
            _resetFilter = WaitForElementIsVisible(By.XPath(RESET_FILTER));
            _resetFilter.Click();
            WaitForLoad();
            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                // pas de date de fin
            }
        }
        public void Filter(FilterType filterType, object value)
        {
            Actions action = new Actions(_webDriver);

            switch (filterType)
            {
                case FilterType.SearchDeliveryRoundName:
                    _searchDeliveryRoundNameInput = WaitForElementIsVisible(By.Id(SEARCH_DELIVERY_ROUND_NAME_INPUT));
                    _searchDeliveryRoundNameInput.SetValue(ControlType.TextBox, value);
                    WaitForLoad();
                    break;
                case FilterType.StartDate:
                    _setupFilterDate = WaitForElementIsVisible(By.Id(SETUP_FILTER_DATE));
                    _setupFilterDate.SetValue(ControlType.DateTime, value);
                    _setupFilterDate.SendKeys(Keys.Tab);
                    WaitForLoad();
                    break;
                case FilterType.Sites:
                    _sites = WaitForElementIsVisible(By.Id(SITES));
                    _sites.SetValue(ControlType.DropDownList, value);
                    break;
                case FilterType.RecipeName:
                    _searchRecipeNameInput = WaitForElementIsVisible(By.Id(SEARCH_RECIPE_NAME_INPUT));
                    _searchRecipeNameInput.SetValue(ControlType.TextBox, value);
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
                    _customerTypesFilter = WaitForElementIsVisible(By.Id(CUSTOMERS_FILTER));
                    _customerTypesFilter.Click();

                    _customerTypesFilterUncheckAll = WaitForElementIsVisible(By.XPath(CUSTOMERS_FILTER_UNCHECK_ALL));
                    _customerTypesFilterUncheckAll.Click();

                    _customerTypesFilterSearch = WaitForElementIsVisible(By.XPath(CUSTOMERS_FILTER_SEARCH));
                    _customerTypesFilterSearch.SetValue(ControlType.TextBox, value);

                    var _customerToCheck = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    _customerToCheck.SetValue(ControlType.CheckBox, true);
                    WaitForLoad();
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
                case FilterType.MealTypes:
                    _mealtypesFilter = WaitForElementExists(By.XPath(MEALTYPES_FILTER));
                    action.MoveToElement(_mealtypesFilter).Perform();
                    _mealtypesFilter.Click();

                    _mealtypesFilterUncheckAll = WaitForElementIsVisible(By.XPath(MEALTYPES_FILTER_UNCHECK_ALL));
                    _mealtypesFilterUncheckAll.Click();

                    _mealtypesFilterSearch = WaitForElementIsVisible(By.XPath(MEALTYPES_FILTER_SEARCH));
                    _mealtypesFilterSearch.SetValue(ControlType.TextBox, value);

                    var _mealtypeToCheck = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    _mealtypeToCheck.SetValue(ControlType.CheckBox, true);
                    WaitForLoad();
                    break;
                case FilterType.Workshops:
                    _workshopsFilter = WaitForElementExists(By.XPath(WORKSHOPS_FILTER));
                    action.MoveToElement(_workshopsFilter).Perform();
                    _workshopsFilter.Click();

                    _workshopsFilterUncheckAll = WaitForElementIsVisible(By.XPath(WORKSHOPS_FILTER_UNCHECK_ALL));
                    _workshopsFilterUncheckAll.Click();

                    _workshopsFilterSearch = WaitForElementIsVisible(By.XPath(WORKSHOPS_FILTER_SEARCH));
                    _workshopsFilterSearch.SetValue(ControlType.TextBox, value);

                    var _workshopToCheck = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    _workshopToCheck.SetValue(ControlType.CheckBox, true);
                    WaitForLoad();
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);

            }

            WaitPageLoading();
        }

        public enum PrintType
        {
            DeliveryRound,
            DeliveryRoundByRecipes,
            DeliveryNoteByRecipes,
            DeliveryNoteByServices,
            DeliveryNoteValorized,
            FoodPackReport,
            FoodPackGroupByDelivery,
            Export,
            Parameters
        }
        public SetupPage Validate()
        {
            if (isElementVisible(By.Id(BTN_VALIDATE)))
            {
                _btnValidate = WaitForElementIsVisible(By.Id(BTN_VALIDATE));
                _btnValidate.Click();
            }
            return new SetupPage(_webDriver, _testContext);
        }

        public SetupDeliveryRoundTabPage GetDeliveryRound()
        {
            //FIXME vue le return, c'est plutot un ClickOnDone()
            var doneBtn = WaitForElementIsVisible(By.Id("btnValidate"));
            doneBtn.Click();
            return new SetupDeliveryRoundTabPage(_webDriver, _testContext);
        }

        public void GoPrintDeliveryNote()
        {
            ShowExtendedMenu();
            var deliveryNoteByRecipesBtn = WaitForElementIsVisible(By.XPath(DELIVERY_NOTE));
            deliveryNoteByRecipesBtn.Click();
            WaitForLoad();
            var validateButton = WaitForElementIsVisible(By.XPath(VALIDATE_BUTTON));
            validateButton.Click();
            WaitForLoad();
        }

        public void GoPrintDeliveryValorized()
        {
            ShowExtendedMenu();
            var deliveryValorizedBtn = WaitForElementIsVisible(By.XPath(DELIVERY_NOTE_VALORIZED));
            deliveryValorizedBtn.Click();
            WaitForLoad();
            var modalPrintSetup = WaitForElementIsVisible(By.XPath(MODAL_PRINT_SETUP));
            modalPrintSetup.Click();
            WaitForLoad();
            var validateButton = WaitForElementIsVisible(By.XPath(VALIDATE_BUTTON));
            validateButton.Click();
            WaitForLoad();
        }

        public void GoPrintDeliveryNoteByService()
        {
            ShowExtendedMenu();
            var deliveryNoteByService = WaitForElementIsVisible(By.XPath(DELIVERY_NOTE_BY_SERVICE));
            deliveryNoteByService.Click();
            WaitForLoad();

            var modalPrintSetup = WaitForElementIsVisible(By.XPath(MODAL_PRINT_SETUP));
            modalPrintSetup.Click();
            WaitForLoad();

            var validateButton = WaitForElementIsVisible(By.XPath(VALIDATE_BUTTON));
            validateButton.Click();
            WaitForLoad();
        }

        public void GoPrintDeliveryNoteByRecipes()
        {
            ShowExtendedMenu();
            var deliveryNoteByService = WaitForElementIsVisible(By.XPath(DELIVERY_NOTE_BY_RECIPES));
            deliveryNoteByService.Click();
            WaitForLoad();
            var validateButton = WaitForElementIsVisible(By.XPath(VALIDATE_BUTTON));
            validateButton.Click();
            WaitForLoad();
        }

        public PrintReportPage PrintArchive(SetupFilterAndFavoritesPage setupFilterAndFavoritesPage)
        {
            // Lancement du Print
            PrintReportPage reportPage;
            // cas export Production Setup "Delivery Notes (singles)"
            reportPage = setupFilterAndFavoritesPage.PrintDeliveryRoundArchive();
            WaitForDownload();
            Close();
            return reportPage;
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

        public PrintReportPage PrintDeliveryRoundResults()
        {
            IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-pdf']"));
            ClickPrintButton();
            WaitForLoad();

            //Results are opened in a new tab, switch the driver to the newly created one
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new PrintReportPage(_webDriver, _testContext);
        }

        public PrintReportPage PrintDeliveryRoundArchive()
        {
            IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-archive']"));
            ClickPrintButton();
            WaitForLoad();

            //Results are opened in a new tab, switch the driver to the newly created one
            //var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            //wait.Until((driver) => driver.WindowHandles.Count > 1);
            //_webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);
            WaitPageLoading();
            return new PrintReportPage(_webDriver, _testContext);
        }

        public void GoPrintFoodPackReport()
        {
            ShowExtendedMenu();
            var foodPackReport = WaitForElementIsVisible(By.XPath(FOOD_PACK_REPORT));
            foodPackReport.Click();
            WaitForLoad();
            var validateButton = WaitForElementIsVisible(By.XPath(VALIDATE_BUTTON));
            validateButton.Click();
            WaitForLoad();
        }
        public void GoPrintDeliveryRound()
        {
            ShowExtendedMenu();
            var deliveryRoundReport = WaitForElementIsVisible(By.Id(DELIVERYROUND));
            deliveryRoundReport.Click();
            WaitForLoad();
            var validateButton = WaitForElementIsVisible(By.XPath(VALIDATE_BUTTON));
            validateButton.Click();
            WaitForLoad();
        }
        public void GoPrintDeliveryRoundSingle()
        {
            ShowExtendedMenu();
            var deliveryRoundReport = WaitForElementIsVisible(By.Id(DELIVERY_ROUND_SINGLE));
            deliveryRoundReport.Click();
            WaitForLoad();
            var validateButton = WaitForElementIsVisible(By.XPath(VALIDATE_BUTTON));
            validateButton.Click();
            WaitForLoad();
        }
        public void GoPrintDeliveryRoundByRecipes()
        {
            ShowExtendedMenu();
            var deliveryRoundByRecipe = WaitForElementIsVisible(By.Id(DELIVERYROUNDBYRECIPES));
            deliveryRoundByRecipe.Click();
            WaitForLoad();
            var validateButton = WaitForElementIsVisible(By.XPath(VALIDATE_BUTTON));
            validateButton.Click();
            WaitForLoad();
        }
        public List<string> GetDeliveries()
        {
            _deliveryButton = _webDriver.FindElement(By.Id(DELIVERYBUTTON));
            _deliveryButton.Click();
            var deliveriesElements = _webDriver.FindElements(By.CssSelector(DELIVERYSELEMENTS));
            List<string> selectedDeliveries = new List<string>();
            foreach (var deliveryElement in deliveriesElements)
            {
                var checkbox = deliveryElement.FindElement(By.TagName("input"));
                var spanText = deliveryElement.FindElement(By.TagName("span")).Text;
                if (checkbox.Selected)
                {
                    selectedDeliveries.Add(spanText);
                }
            }
            _deliveryButton.Click();

            return selectedDeliveries;
        }

        public string GetDeliveriesDropdownState()
        {
            _deliveryButton = _webDriver.FindElement(By.Id(DELIVERYBUTTON));
            _deliveryButton.Click();
            var buttonTextElement = _deliveryButton.FindElement(By.CssSelector(TEXTNOTHINGSELECTED));
            return buttonTextElement.Text;
        }

        public PrintReportPage PrintItemGenerique2(SetupFilterAndFavoritesPage setupFilterAndFavoritesPage)
        {
            // Lancement du Print
            PrintReportPage reportPage;
            reportPage = setupFilterAndFavoritesPage.PrintDeliveryRoundResults();
            // le zip se télécharge directement, alors que le pdf s'ouvre dans un nouveau onglet
            var isReportGenerated = reportPage.IsReportGenerated();
            //reportPage.Close();
            //setupFilterAndFavoritesPage.ClickPrintButton();

            //Assert
            Assert.IsTrue(isReportGenerated, "L'application n'a pas pu générer le fichier attendu.");


            return reportPage;
        }

        public bool RetourLigneEffectuer(List<string> wordStrings, string[] result)
        {

            string Cordren = wordStrings[0];
            string Information = wordStrings[4];
            string date = wordStrings[5];

            string CordrenResult = result[0];
            string InformationResult = result[1];
            string dateResult = result[2];

            if ((Cordren == CordrenResult) && (Information == InformationResult) && (date == dateResult))
            {
                return true;
            }
            else return false;

        }
        public void SelectDeliveryRound()
        {
            var select = WaitForElementIsVisible(By.Id(SELECT_DELIVERY_ROUND));
            select.Click();
            var option = WaitForElementIsVisible(By.XPath(OPTION_DELIVERY));
            option.Click();

        }
        public void ClearDownloadFolder(string directory)
        {
            foreach (string file in Directory.GetFiles(directory))
            {
                File.Delete(file);
            }
        }
        public List<string> GetNameProvidersList()
        {
            var NameProvidersListId = new List<string>();

            var NameProvidersList = _webDriver.FindElements(By.XPath(PROVIDERS_NAMES));

            foreach (var fileFlowProviders in NameProvidersList)
            {
                NameProvidersListId.Add(fileFlowProviders.Text);
            }

            return NameProvidersListId;
        }

        public void MakeFavorite(string favoriteName)
        {
            WebDriverWait wait = new WebDriverWait(_webDriver,TimeSpan.FromSeconds(5));
            var _makeFavorite = WaitForElementIsVisible(By.XPath("*//a[text()=\"Make Favorite\"]"));
            _makeFavorite.Click();

            var _favoriteName = WaitForElementIsVisible(By.Id(FAVORITE_NAME));
            _favoriteName.Clear();
            _favoriteName.SendKeys(favoriteName);

            var _saveFavorite = WaitForElementIsVisible(By.Id(SAVE_FAVORITE));
            _saveFavorite.Click();
            //too long to save
            wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.InvisibilityOfElementLocated(By.XPath("//*[@id=\"modal-1\"]")));
        }
        public void SetFavoriteText(string favoriteName)
        {
            WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(10));
            var favorite_search = _webDriver.FindElement(By.XPath(FAVORITE_SEARCH));
            favorite_search.Clear();
            favorite_search.SetValue(ControlType.TextBox,favoriteName);
            wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath($"//*[@id=\"searchPattern\"][@value='{favoriteName}']")));
            LoadingPage();
        }
        public void SelectFavorite()
        {

            var _favoriteName = WaitForElementIsVisible(By.XPath(SELECTED_FAVORITE));
            _favoriteName.Click();
            WaitPageLoading();
        }
        public bool IsFavoritePresent(string favoriteName)
        {
            var favorite = _webDriver.FindElements(By.XPath(String.Format(FAVORITE_RES_FILTRE, favoriteName))).Count;

            if (favorite == 0)
                return false;

            return true;
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
            var _favoriteFilter = _webDriver.FindElement(By.XPath(FAVORITE_FILTER));
            string text = _favoriteFilter.Text;

            return text;
        }

        public void DeleteFavorite(string favoriteName)
        {

            Actions action = new Actions(_webDriver);
            var favorite = _webDriver.FindElements(By.XPath(String.Format(FAVORITE, favoriteName))).Count;

            if (favorite > 0)
            {
                var _favorite = WaitForElementIsVisible(By.XPath(String.Format(FAVORITE, favoriteName)));
                action.MoveToElement(_favorite).Perform();

                var deleteFavorite = WaitForElementIsVisible(By.XPath(String.Format(DELETE_FAVORITECROSS, favoriteName)));

                ((IJavaScriptExecutor)(IWebDriver)_webDriver).ExecuteScript(
                 "arguments[0].removeAttribute('class','class')", deleteFavorite);

                var _deleteFavorite = WaitForElementIsVisible(By.XPath(String.Format(DELETE_FAVORITE, favoriteName)));
                _deleteFavorite.Click();
                WaitForLoad();

                var _confirmDeleteFavorite = WaitForElementIsVisible(By.Id(CONFIRM_DELETE_FAVORITE));
                _confirmDeleteFavorite.Click();
                WaitForLoad();
            }
        }
        public ResultModal ResultFavoriteName(string favoriteName)
        {
            TryToClickOnElement(_favorite_result);
            return new ResultModal(_webDriver, _testContext);
        }
        public string GetSitesAfterReset()
        {
            var sitesDropdown = WaitForElementIsVisible(By.Id(SITES));
            return sitesDropdown.Text;
        }
        public string GetCustomerType()
        {
            // Attendre que le filtre CustomerTypes soit visible
            var customerTypesFilter = WaitForElementIsVisible(By.Id(CUSTOMER_TYPES_FILTER));
            // Récupérer et retourner le texte du filtre
            return customerTypesFilter.Text;
        }
        public string GetSelectedCustomerType()
        {
            var customerTypesFilter = WaitForElementIsVisible(By.Id(CUSTOMER_TYPES_FILTER));
            customerTypesFilter.Click();
            var selectedOption = WaitForElementIsVisible(By.XPath("/html/body/div[10]/ul/li[1]/label"));
            return selectedOption.Text;
        }






    }
}