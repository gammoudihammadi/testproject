using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.PriceList
{
    public class PriceListDetailsPage : PageBase
    {
        public PriceListDetailsPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //__________________________________Constantes_______________________________________

        // General
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";
        private const string PERIOD = "//*[@id=\"tabContentItemContainer\"]/div[2]/h3";
        private const string EXPORT = "exportBtn";
        private const string IMPORT = "//*[@id=\"tabContentItemContainer\"]/div[1]/div/a[2]";
        private const string PRICE_NAME = "//*[@id=\"tabContentItemContainer\"]/div[1]/h1";

        private const string ADD_NEW_PERIOD = "//*[@id=\"tabContentItemContainer\"]/div[1]/div/div/a[2]";
        private const string PERIOD_START_DATE = "StartDate";
        private const string PERIOD_END_DATE = "EndDate";
        private const string CREATE_PERIOD = "//*[@id=\"form-createdit-vipPrice\"]/div[3]/div/button[2]";
        private const string ERROR_MESSAGE_PERIOD = "//*[@id=\"createFormVipPrice\"]/p/span";
        private const string PERIOD_DROPDOWN = "period-list";
        private const string PERIOD_LIST = "//*[@id=\"period-list\"]/option";
        private const string SELECTED_PERIOD = "//*[@id=\"period-list\"]/option[contains(text(),'{0}')]";
        private const string FOREIGN_SWITCH = "//*[@id=\"tabContentItemContainer\"]/div[3]/div";
        private const string POPUP_CANCEL_BUTTON = "//*[@id=\"form-createdit-vipPrice\"]/div[3]/div/button[1]";

        // Onglets
        private const string GENERAL_INFORMATION = "//*[@id=\"-1\"]/td/h4";

        // Add item
        private const string SHOW_ITEMS_ONLY = "ItemsOnly";
        private const string ADD_ITEM = "IngredientName";
        private const string SELECTED_ITEM = "//*[@id=\"ingredient-list\"]/table/tbody/tr[@data-name = '{0}']/td[1]";

        // Tableau
        private const string ITEM_NAME = "//*[@id=\"price-list-table\"]/tbody/tr[*]/td[2]";
        private const string INITIAL_PRICE = "//*[@id=\"price-list-table\"]/tbody/tr[*]/td[5]/div[1]/span";
        private const string INITIAL_PRICE_SELECTED = "//*[@id=\"price-list-table\"]/tbody/tr[*]/td[5]/div[2]/div[2]/input";
        private const string MARKUP = "//*[@id=\"price-list-table\"]/tbody/tr[2]/td[7]/span[2]";
        private const string COEFF = "//*[@id=\"price-list-table\"]/tbody/tr[*]/td[contains(text(),'{0}')]";
        private const string COEFF_INPUT = "//*[@id=\"price-list-table\"]/tbody/tr[*]/td[contains(text(),'{0}')]/../td[6]/input";
        private const string PROD_COMMENT = "//*[@id=\"price-list-table\"]/tbody/tr[2]/td[11]/a";
        private const string BILLING_COMMENT = "//*[@id=\"price-list-table\"]/tbody/tr[2]/td[12]/a";
        private const string COMMENT = "Comment";
        private const string SAVE_COMMENT = "//*[@id=\"vipPriceDetailCommentForm\"]/div[2]/button[2]";

        private const string DELETE_ITEM = "//*[@id=\"price-list-table\"]/tbody/tr[2]/td[13]/a";
        private const string CONFIRM_DELETE = "dataConfirmOK";

        // Filtres
        private const string RESET_FILTER = "//*[@id=\"item-filter-form\"]/div[1]/a";
        private const string SEARCH_FILTER = "SearchPattern";
        private const string SORTBY_FILTER = "cbSortBy";
        private const string SHOW_WITHOUT_PRICE_FILTER = "ShowWithoutPriceOnly";
        private const string SHOW_NEGATIVE_MARKUP_FILTER = "ShowNegativeMarkupOnly";
        private const string ELEMENT_TYPE = "//*[@id=\"item-filter-form\"]/div[6]/a";
        private const string ELEMENT_TYPE_ALL_FILTER = "all";
        private const string ELEMENT_TYPE_ITEMS_FILTER = "items";
        private const string ELEMENT_TYPE_RECIPES_FILTER = "recipes";
        private const string ELEMENT_TYPE_FREE_PRICES_FILTER = "freeprices";

        //__________________________________ Variables _______________________________________

        // Général
        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;

        [FindsBy(How = How.XPath, Using = PERIOD)]
        private IWebElement _period;

        [FindsBy(How = How.Id, Using = EXPORT)]
        private IWebElement _export;

        [FindsBy(How = How.XPath, Using = IMPORT)]
        private IWebElement _import;

        [FindsBy(How = How.XPath, Using = ADD_NEW_PERIOD)]
        private IWebElement _addNewPeriod;

        [FindsBy(How = How.Id, Using = PERIOD_START_DATE)]
        public IWebElement _periodStartDate;

        [FindsBy(How = How.Id, Using = PERIOD_END_DATE)]
        public IWebElement _periodEndDate;

        [FindsBy(How = How.XPath, Using = CREATE_PERIOD)]
        private IWebElement _createPeriod;

        [FindsBy(How = How.XPath, Using = ERROR_MESSAGE_PERIOD)]
        private IWebElement _errorMessagePeriod;

        [FindsBy(How = How.Id, Using = PERIOD_DROPDOWN)]
        private IWebElement _periodDropDown;

        [FindsBy(How = How.XPath, Using = PRICE_NAME)]
        private IWebElement _priceName; 
        
        [FindsBy(How = How.XPath, Using = FOREIGN_SWITCH)]
        private IWebElement _foreignSwitch;   
        
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement _commercialName;  
        
        [FindsBy(How = How.XPath, Using = POPUP_CANCEL_BUTTON)]
        private IWebElement _cancelButton;

        // Onglets
        [FindsBy(How = How.XPath, Using = GENERAL_INFORMATION)]
        private IWebElement _generalInformationBtn;

        // Add item
        [FindsBy(How = How.Id, Using = SHOW_ITEMS_ONLY)]
        private IWebElement _showItemsOnly;

        [FindsBy(How = How.XPath, Using = ADD_ITEM)]
        private IWebElement _addItem;

        [FindsBy(How = How.XPath, Using = SELECTED_ITEM)]
        private IWebElement _selectedItem;


        // Tableau

        [FindsBy(How = How.XPath, Using = ITEM_NAME)]
        private IWebElement _firstItemName;

        [FindsBy(How = How.XPath, Using = INITIAL_PRICE)]
        private IWebElement _initialPrice;

        [FindsBy(How = How.XPath, Using = INITIAL_PRICE_SELECTED)]
        private IWebElement _initialPriceSelected;

        [FindsBy(How = How.XPath, Using = MARKUP)]
        private IWebElement _markup;

        [FindsBy(How = How.XPath, Using = COEFF_INPUT)]
        private IWebElement _coeff;

        [FindsBy(How = How.XPath, Using = PROD_COMMENT)]
        private IWebElement _addProductionComment;

        [FindsBy(How = How.XPath, Using = BILLING_COMMENT)]
        private IWebElement _addBillingComment;

        [FindsBy(How = How.Id, Using = COMMENT)]
        private IWebElement _comment;

        [FindsBy(How = How.XPath, Using = COMMENT)]
        private IWebElement _validateComment;

        [FindsBy(How = How.XPath, Using = DELETE_ITEM)]
        private IWebElement _deleteItem;

        [FindsBy(How = How.Id, Using = CONFIRM_DELETE)]
        private IWebElement _confirmDelete;

        // __________________________________ Filtres _______________________________________

        [FindsBy(How = How.XPath, Using = RESET_FILTER)]
        private IWebElement _resetFilter;

        [FindsBy(How = How.Id, Using = SEARCH_FILTER)]
        private IWebElement _searchFilter;

        [FindsBy(How = How.Id, Using = SORTBY_FILTER)]
        private IWebElement _sortByFilter;

        [FindsBy(How = How.XPath, Using = SHOW_WITHOUT_PRICE_FILTER)]
        private IWebElement _showWithoutPriceOnly;

        [FindsBy(How = How.XPath, Using = SHOW_NEGATIVE_MARKUP_FILTER)]
        private IWebElement _showNegativeMarkup;

        [FindsBy(How = How.XPath, Using = ELEMENT_TYPE)]
        private IWebElement _elementType;

        [FindsBy(How = How.XPath, Using = ELEMENT_TYPE_ALL_FILTER)]
        private IWebElement _elementTypeAll;

        [FindsBy(How = How.XPath, Using = ELEMENT_TYPE_ITEMS_FILTER)]
        private IWebElement _elementTypeItems;

        [FindsBy(How = How.XPath, Using = ELEMENT_TYPE_RECIPES_FILTER)]
        private IWebElement _elementTypeRecipes;

        [FindsBy(How = How.XPath, Using = ELEMENT_TYPE_RECIPES_FILTER)]
        private IWebElement _elementTypeFreeRecipes;

        public enum FilterType
        {
            Search,
            SortBy,
            ShowWithoutPriceOnly,
            ShowNegativeMarkup,
            ElementTypeItems,
            ElementTypeRecipes,
            ElementTypeFreePrices,
            ElementTypeAll
        }

        public void Filter(FilterType filterType, object value)
        {
            switch (filterType)
            {
                case FilterType.Search:
                    _searchFilter = WaitForElementIsVisible(By.Id(SEARCH_FILTER));
                    _searchFilter.SetValue(ControlType.TextBox, value);
                    break;
                case FilterType.SortBy:
                    _sortByFilter = WaitForElementIsVisible(By.Id(SORTBY_FILTER));
                    _sortByFilter.Click();
                    var element = WaitForElementIsVisible(By.XPath("//option[contains(@value,'" + value + "')]"));
                    _sortByFilter.SetValue(ControlType.DropDownList, element.Text);
                    _sortByFilter.Click();
                    break;
                case FilterType.ShowWithoutPriceOnly:
                    _showWithoutPriceOnly = WaitForElementExists(By.Id(SHOW_WITHOUT_PRICE_FILTER));
                    _showWithoutPriceOnly.SetValue(ControlType.CheckBox, value);
                    break;
                case FilterType.ShowNegativeMarkup:
                    _showNegativeMarkup = WaitForElementExists(By.Id(SHOW_NEGATIVE_MARKUP_FILTER));
                    _showNegativeMarkup.SetValue(ControlType.CheckBox, value);
                    break;
                case FilterType.ElementTypeAll:
                    _elementTypeAll = WaitForElementExists(By.Id(ELEMENT_TYPE_ALL_FILTER));
                    _elementTypeAll.SetValue(ControlType.RadioButton, value);
                    WaitForLoad();
                    break;
                case FilterType.ElementTypeItems:
                    _elementTypeItems = WaitForElementExists(By.Id(ELEMENT_TYPE_ITEMS_FILTER));
                    _elementTypeItems.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.ElementTypeRecipes:
                    _elementTypeRecipes = WaitForElementExists(By.Id(ELEMENT_TYPE_RECIPES_FILTER));
                    _elementTypeRecipes.SetValue(ControlType.RadioButton, value);
                    WaitForLoad();
                    break;
                case FilterType.ElementTypeFreePrices:
                    _elementTypeFreeRecipes = WaitForElementExists(By.Id(ELEMENT_TYPE_FREE_PRICES_FILTER));
                    _elementTypeFreeRecipes.SetValue(ControlType.RadioButton, value);
                    WaitForLoad();
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);

            }

            WaitPageLoading();
            WaitForLoad();
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

        public void ShowElementTypesFilter()
        {
            _elementType = WaitForElementIsVisible(By.XPath(ELEMENT_TYPE));

            if (_elementType.GetAttribute("class").Equals("filterCollapseButton collapsed"))
            {
                _elementType.Click();
                WaitPageLoading();
            }
        }

        // _____________________________________________ Méthodes _____________________________________________

        // Général
        public PriceListPage BackToList()
        {
            _backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            _backToList.Click();
            WaitForLoad();

            return new PriceListPage(_webDriver, _testContext);
        }

        public string GetPriceName()
        {
            _priceName = WaitForElementIsVisible(By.XPath(PRICE_NAME));

            var text = _priceName.Text.Replace("PRICE : ", "");

            return text.Trim();
        }

        public string GetPeriod()
        {
            _period = WaitForElementIsVisible(By.XPath(PERIOD));
            return _period.Text;
        }

        public int CountPeriod()
        {
            return _webDriver.FindElements(By.XPath(PERIOD_LIST)).Count;
        }

        public void ExportPriceList(bool versionPrint)
        {
            _export = WaitForElementIsVisible(By.Id(EXPORT));
            _export.Click();
            WaitForLoad();

            if (versionPrint)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));
                ClickPrintButton();
            }

            WaitForDownload();
            Close();
        }

        public FileInfo GetExportExcelFile(FileInfo[] taskFiles)
        {
            List<FileInfo> correctDownloadFiles = new List<FileInfo>();

            foreach (var file in taskFiles)
            {
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
            string space = "(\\s)";                    // Espace
            string annee = "\\d{4}";                   // annee YYYY
            string jour = "[0-3]\\d";                  // jour
            string heure = "(?:0[0-9]|1[0-9]|2[0-3])"; // heure
            string minutes = "[0-5]\\d";               // minutes
            string secondes = "[0-5]\\d";              // secondes

            //VipPrices_Export_2020-01-28 15-02-06.xlsx
            Regex r = new Regex("^VipPrices_Export_" + annee + "-" + mois + "-" + jour + space + heure + "-" + minutes + "-" + secondes + ".xlsx$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = r.Match(filePath);

            return m.Success;
        }

        public PriceListImportPage PriceListImportPage()
        {
            _import = WaitForElementIsVisible(By.XPath(IMPORT));
            _import.Click();
            WaitForLoad();

            return new PriceListImportPage(_webDriver, _testContext);
        }


        // Onglets
        public PriceListGeneralInformationPage ClickOnGeneralInformation()
        {
            _generalInformationBtn = WaitForElementIsVisible(By.XPath(GENERAL_INFORMATION));
            _generalInformationBtn.Click();
            WaitForLoad();

            return new PriceListGeneralInformationPage(_webDriver, _testContext);
        }

        // Add item
        public void AddItem(string itemName, bool isIngredient)
        {
            _showItemsOnly = WaitForElementExists(By.Id(SHOW_ITEMS_ONLY));
            _showItemsOnly.SetValue(ControlType.CheckBox, isIngredient);
            WaitForLoad();

            _addItem = WaitForElementIsVisible(By.Id(ADD_ITEM));
            _addItem.SetValue(ControlType.TextBox, itemName);
            WaitPageLoading();
            WaitForLoad();

            _selectedItem = WaitForElementIsVisible(By.XPath(string.Format(SELECTED_ITEM, itemName)));
            _selectedItem.Click();
            WaitForLoad();

            Thread.Sleep(1500);
        }

        public void AddCommercialNameByRow(string commName, int row)
        {
            string rowXpath = $"/html/body/div[3]/div/div[2]/div/div[4]/div[1]/div/div/table/tbody/tr[{row+1}]";
            var rowToClick = WaitForElementToBeClickable(By.XPath(rowXpath));
            rowToClick.Click();
            string fullXpath = $"/html/body/div[3]/div/div[2]/div/div[4]/div[1]/div/div/table/tbody/tr[{row+1}]/td[2]/input";
            var _commNameInput = WaitForElementIsVisible(By.XPath(fullXpath));
            _commNameInput.SetValue(ControlType.TextBox, commName);
            WaitForLoad();
        }


        // Tableau
        public int CountItems()
        {
            return _webDriver.FindElements(By.XPath(ITEM_NAME)).Count;
        }

        public bool IsItemAdded(string itemName)
        {
            var items = _webDriver.FindElements(By.XPath(ITEM_NAME));

            if (items.Count == 0)
                return false;

            foreach (var elm in items)
            {
                if (elm.Text.Contains(itemName))
                    return true;
            }

            return false;
        }

        public string GetItemName()
        {
            _firstItemName = WaitForElementIsVisible(By.XPath(ITEM_NAME));
            return _firstItemName.Text.Trim();
        }

        public void UpdateItem(string initPrice, int ligne)
        {
            var items = _webDriver.FindElements(By.XPath(ITEM_NAME));
            items[ligne].Click();
            WaitForLoad();

            _initialPriceSelected = SolveVisible(INITIAL_PRICE_SELECTED);
            _initialPriceSelected.SetValue(ControlType.TextBox, initPrice);
            WaitForLoad();

            _initialPriceSelected.SendKeys(Keys.ArrowRight);
            Thread.Sleep(2000);
            _webDriver.Navigate().Refresh();
        }

        public double GetInitialPrice(string decimalSeparator)
        {
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            _initialPrice = WaitForElementExists(By.XPath(INITIAL_PRICE));

            return double.Parse(_initialPrice.Text, ci);
        }

        public double GetMarkup(string decimalSeparator)
        {
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            _markup = WaitForElementIsVisible(By.XPath(MARKUP));
            return double.Parse(_markup.Text, ci);
        }

        public void AddCoef(string value, string itemName)
        {
            var firstItem = WaitForElementIsVisible(By.XPath(string.Format(COEFF, itemName)));
            firstItem.Click();
            WaitForLoad();

            _coeff = WaitForElementIsVisible(By.XPath(string.Format(COEFF_INPUT, itemName)));
            _coeff.SetValue(ControlType.TextBox, value);
            _coeff.SendKeys(Keys.Tab);
            // calcul "Price excl. VAT" selon "Initial Price"
            Thread.Sleep(10000);
            WaitForLoad();
        }

        public void AddProductionComment(string comment)
        {
            _addProductionComment = WaitForElementIsVisible(By.XPath(PROD_COMMENT));
            _addProductionComment.Click();
            WaitForLoad();

            _comment = WaitForElementIsVisible(By.Id(COMMENT));
            _comment.SetValue(ControlType.TextBox, comment);

            _validateComment = WaitForElementIsVisible(By.XPath(SAVE_COMMENT));
            _validateComment.Click();
            WaitForLoad();
        }

        public string GetProductionComment()
        {
            _addProductionComment = WaitForElementIsVisible(By.XPath(PROD_COMMENT));
            _addProductionComment.Click();
            WaitForLoad();

            _comment = WaitForElementIsVisible(By.Id(COMMENT));
            return _comment.Text;
        }

        public void AddBillingComment(string comment)
        {
            _addBillingComment = WaitForElementIsVisible(By.XPath(BILLING_COMMENT));
            _addBillingComment.Click();
            WaitForLoad();

            _comment = WaitForElementIsVisible(By.Id(COMMENT));
            _comment.SetValue(ControlType.TextBox, comment);

            _validateComment = WaitForElementIsVisible(By.XPath(SAVE_COMMENT));
            _validateComment.Click();
            WaitForLoad();
        }

        public string GetBillingComment()
        {
            _addBillingComment = WaitForElementIsVisible(By.XPath(BILLING_COMMENT));
            _addBillingComment.Click();
            WaitForLoad();

            _comment = WaitForElementIsVisible(By.Id(COMMENT));
            return _comment.Text;
        }

        public void DeleteItem()
        {
            var firstItem = WaitForElementIsVisible(By.XPath(ITEM_NAME));
            firstItem.Click();

            _deleteItem = WaitForElementToBeClickable(By.XPath(DELETE_ITEM));
            _deleteItem.Click();
            WaitForLoad();

            _confirmDelete = WaitForElementIsVisible(By.Id(CONFIRM_DELETE));
            _confirmDelete.Click();
            WaitForLoad();
        }

        public void AddNewPeriod(DateTime fromDate, DateTime toDate)
        {
            _addNewPeriod = WaitForElementIsVisible(By.XPath(ADD_NEW_PERIOD));
            _addNewPeriod.Click();
            WaitForLoad();

            _periodStartDate = WaitForElementIsVisible(By.Id(PERIOD_START_DATE));
            _periodStartDate.SetValue(ControlType.DateTime, fromDate);
            _periodStartDate.SendKeys(Keys.Tab);

            _periodEndDate = WaitForElementIsVisible(By.Id(PERIOD_END_DATE));
            _periodEndDate.SetValue(ControlType.DateTime, toDate);
            _periodEndDate.SendKeys(Keys.Tab);

            _createPeriod = WaitForElementToBeClickable(By.XPath(CREATE_PERIOD));
            _createPeriod.Click();
            WaitForLoad();
        }

        public void ChangePeriod(string startDate)
        {
            _periodDropDown = WaitForElementIsVisible(By.Id(PERIOD_DROPDOWN));
            _periodDropDown.Click();
            var selectedPeriod = WaitForElementIsVisible(By.XPath(string.Format(SELECTED_PERIOD, startDate)));
            _periodDropDown.SetValue(ControlType.DropDownList, selectedPeriod.Text);
        }

        public bool ErrorMessageNewPeriod(string message)
        {
            _errorMessagePeriod = WaitForElementIsVisible(By.XPath(ERROR_MESSAGE_PERIOD));

            if (_errorMessagePeriod.Text.Contains(message))
                return true;

            return false;
        }

        public void ClickOnForeign()
        {
            _foreignSwitch = WaitForElementIsVisible(By.XPath(FOREIGN_SWITCH));
            _foreignSwitch.Click();
            WaitForLoad();
        }

        public string GetNameByRow(int row)
        {
            string xpath = $"//*[@id=\"price-list-table\"]/tbody/tr[{row + 1}]/td[2]";
            _commercialName = WaitForElementIsVisible(By.XPath(xpath));
            string name = _commercialName.Text.Replace("N : ", "");
            int indexOfCN = name.IndexOf("\r");
            string res = name.Substring(0, indexOfCN);
            return res;
        } 
        
        public string GetCommercialNameByRow(int row)
        {
            //NOT TESTED
            string xpath = $"//*[@id=\"price-list-table\"]/tbody/tr[{row + 1}]/td[2]";
            _commercialName = WaitForElementIsVisible(By.XPath(xpath));
            string name = _commercialName.Text.Replace("N : ", "");

            string CNstr = "CN : ";
            int indexOfCN = name.IndexOf(CNstr);
            string res = name.Substring(indexOfCN + CNstr.Length, name.Length);
            return res;
        }

        public void ClosePeriodPopup()
        {
            _cancelButton = WaitForElementIsVisible(By.XPath(POPUP_CANCEL_BUTTON));
            _cancelButton.Click();
        }
        public string GetCommercialNameByRowNumber(int row)
        {
            //NOT TESTED
             var commercialName = WaitForElementIsVisible(By.XPath(string.Format("//*[@id=\"price-list-table\"]/tbody/tr[{0}]/td[2]/span",row+1)));
            return commercialName.Text;
        }
        public bool GetTotalTabs()
        {
            IList<string> allWindows = _webDriver.WindowHandles;

            return allWindows.Count > 1;
        }

    }
}
