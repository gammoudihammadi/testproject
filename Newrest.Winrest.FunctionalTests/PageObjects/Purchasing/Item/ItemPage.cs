using ClosedXML.Excel;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Accounting.FreePrice;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Item;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Drawing;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Purchasing
{
    public class ItemPage : PageBase
    {
        public ItemPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // _____________________________________________ Constantes ________________________________________________

        // Général
        private const string CREATE_NEW_ITEM = "New item";

        private const string FOLD_ALL = "//*[@id=\"unfoldBtn\"][@class=\"unfold-all-btn unfold-all-btn-auto unfolded\"]";
        private const string UNFOLD_ALL = "//*[@id=\"unfoldBtn\"][@class=\"unfold-all-btn unfold-all-btn-auto\"]";
        private const string EXPORT_BTN = "exportBtn";
        private const string CONFIRM_EXPORT_BTN = "btn-export";
        private const string IMPORT = "btn-import-items";
        private const string PRINT_THEORICAL_REPORT = "//*/a[text()='Print Theorical Report']";
        private const string ITEM_QTY_ITEM = "//*[@id=\"detailsItemContainer\"]/div/table/tbody/tr/td[6]";
        private const string ITEM_UNIT = "//*[@id=\"detailsItemContainer\"]/div/table/tbody/tr/td[7]";
        private const string SEARCH_ITEMSO = "SearchPatternWithAutocomplete";
        private const string FIRST_ITEM_NUMBER = "//*[@id=\"div-copy-from-items\"]/table/tbody/tr[2]/td[2]";
        private const string POPUP_DEACTIVATED = "formId";
        private const string CANCEL_BTN = "//*[@id=\"formId\"]/div[3]/button[1]";
        private const string SAVE_BTN = "//*[@id=\"formId\"]/div[3]/button[2]";
        private const string BACK_BTN = "//*[@id=\"ItemBtn\"]/span[2]";


        // Tableau items
        private const string PURCHASE_TAB = "//*[@id=\"itemTabTab\"]/li[1]";
        private const string STORAGE_TAB = "//*[@id=\"itemTabTab\"]/li[2]";
        private const string INTOLERANCE_TAB = "hrefTabContentIntolerance";


        private const string STORAGE_ITEM = "hrefTabContentStorage";
        private const string FIRST_ITEM = "//*[@id=\"purchasing-item-detail-1 \"]/div[2]/table/tbody/tr/td[2]";
        private const string FIRST_ITEM_NAME = "//*[@id=\"purchasing-item-detail-1 \"]/div[2]/table/tbody/tr/td[2]";
        private const string FIRST_ITEM_COL = "//*[@id=\"list-item-with-action\"]/div[2]/div[1]/div";
        private const string CONTENT = "//*[starts-with(@id,\"content_\")]";
        private const string FIRST_ITEM_UNIT = "//*[@id=\"purchasing-item-detail-1 \"]/div[2]/table/tbody/tr/td[4]";
        private const string USECASE_TAB = "//*[@id=\"hrefTabContentWeightRecipe\"]";
        private const string FIRST_ITEM_UNIT_PATCH = "//*[@id=\"purchasing-item-detail-1 \"]/div[2]/table/tbody/tr/td[4]";
        private const string FIRST_ITEM_ACTION = "//*[@id=\"list-item-with-action\"]/div[2]/div[1]/div/div[2]/table/tbody/tr/td[2]";
        private const string FIRST_ITEM_GROUP = "/html/body/div[2]/div/div[2]/div[2]/div/div/div/div[2]/div[1]/div/div[2]/table/tbody/tr/td[3]";
        private const string GENERAL_INFORMATION_TAB = "/html/body/div[3]/div/div[2]/ul/li[1]/a\r\n";

        // Filtres
        private const string RESET_FILTER_DEV = "ResetFilter";
        private const string SEARCH_FILTER = "SearchPatternWithAutocomplete";
        private const string FIRST_RESULT_SEARCH = "//*[@id=\"item-filter-form\"]/div[2]/span/div/div/div/strong[text()='{0}']";
        private const string KEYWORD_FILTER = "SelectedKeywords_ms";
        private const string SEARCH_KEYWORD = "/html/body/div[10]/div/div/label/input";
        private const string UNCHECK_ALL_KEYWORD = "/html/body/div[10]/div/ul/li[2]/a";
        private const string SEARCH_SUPPLIER_REF_FILTER = "SearchRefSupplier";
        private const string SORT_BY = "cbSortBy";
        private const string SHOW_ALL_DEV = "ShowAll";
        private const string SHOW_ONLY_ACTIVE_DEV = "ShowOnlyActive";
        private const string SHOW_ONLY_INACTIVE_DEV = "ShowOnlyInactive";
        private const string ACTIVE_NOT_PURCHASABLE_DEV = "ActiveNotPurchasable";
        private const string SITE_FILTER = "SelectedSites_ms";
        private const string SEARCH_SITE = "/html/body/div[11]/div/div/label/input";
        private const string UNCHECK_ALL_SITES = "/html/body/div[11]/div/ul/li[2]/a";
        private const string SUPPLIER_FILTER = "SelectedSuppliers_ms";
        private const string ALLERGENS_FILTER = "SelectedAllergens_ms";
        private const string SEARCH_ALLERGENS = "/html/body/div[13]/div/div/label/input";
        private const string UNCHECK_ALL_ALLERGENS = "/html/body/div[13]/div/ul/li[2]/a";
        private const string GROUP_FILTER = "SelectedGroups_ms";
        private const string SUBGROUP_FILTER = "SelectedSubGroups_ms";

        private const string NAME = "//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]/table/tbody/tr/td[2]";
        private const string GROUP = "//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]/table/tbody/tr/td[3]";
        private const string INACTIVE = "//*[@id=\"list-item-with-action\"]/div[{0}]/div[1]/div/div[2]/table/tbody/tr/td[1]/img[@alt='Inactive']";

        private const string FIRST_PACKAGING = "/html/body/div[3]/div/div[2]/div[2]/div/div[2]/div[1]/div[2]/div/table/tbody/tr";
        private const string LIMIT_QTY = "MinimumAlertThreshold";
        private const string STORAGE = "hrefTabContentStorage";
        private const string STORAGE_THRESHOLD_FILTER = "ShowStorageAlertThresholdOnly";
        private const string MASSIVE_DELETE = "//*[@id=\"btn-massive-delete-items\"]";
        private const string LABEL_TAB = "//*[@id=\"hrefTabContentComposition\"]";
        private const string COMMENT_FIELD = "//*[@id=\"Composition\"]";
        private const string VALIDATE_BUTTON = "//*[@id=\"ValidateBtn\"]";

        private const string DOWNLOADS_FIRST_ROW = "//*[starts-with(@id, \"printData\")]";
        private const string INGREDIENTNAME = "//*[@id=\"tabContentItemContainer\"]/div[2]/div/ul/li[1]/div/div/div/form/div[3]/div[2]/div[2]/p/span[1]";
        private const string SELECTALL = "//*[@id=\"selectallBtn\"]";
        private const string REPLACE = "//*[@id=\"tabContentItemContainer\"]/div[1]/div/div[1]/a[1]";
        private const string INPUTE_REPLACE_ITEM = "//*[@id=\"searchVM_SearchPatternWithAutocomplete\"]";
        private const string FIRSTLINE = "//*[@id=\"search-replacement-item-form\"]/div[2]/div[1]/span/div/div/div[1]";
        private const string REPLACEWITHITEM = "//*[@id=\"list-item-with-action\"]/div[1]/div[1]/div/div[2]/table/tbody/tr/td[2]/div/div/a[2]";
        private const string CONFIRMBUTTON = "//*[@id=\"dataConfirmOK\"]";
        private const string SUCCESS = "//*[@id=\"modal-1\"]/div[2]/p";
        private const string CLOSE = "//*[@id=\"modal-1\"]/div[3]/button";
        private const string PRODQUANTITYERROR = "//*[@id=\"Heading0\"]/div[1]/div/div[2]/p/b";
        private const string PRODWEIGHT = "//*[@id=\"Heading1\"]/div[1]/div/div[2]/p/b";
        private const string STORAGEQUANTITY = "//*[@id=\"Heading2\"]/div[1]/div/div[2]/p/b";
        private const string CANCEL = "//*[@id=\"ImportFileForm\"]/div[3]/button";
        private const string CLICK_DETAIL = "/html/body/div[2]/div/div[2]/div[2]/div/div/div/div[2]/div[1]/div/div[1]";
        private const string UNIT_PRICE = "/html/body/div[2]/div/div[2]/div[2]/div/div/div/div[2]/div[2]/div/table/tbody[2]/tr/td[7]";
        private const string DEACTIVATED = "btn-deactivate-items";
        private const string YIELDCHOICEITEM = "yieldChoiceItem";
        private const string FIRST_ITEMSO = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[2]";
        private const string FIRST_I = "purchasing-item-detail-row-1";



        // _____________________________________________ Variables _______________________________________________

        // General
        [FindsBy(How = How.XPath, Using = DOWNLOADS_FIRST_ROW)]
        private IWebElement _downloadFirstRowButton;
        [FindsBy(How = How.XPath, Using = CANCEL)]
        private IWebElement _cancel;
        [FindsBy(How = How.XPath, Using = REPLACEWITHITEM)]
        private IWebElement _replacewithitem;
        [FindsBy(How = How.XPath, Using = CLOSE)]
        private IWebElement _close;
        [FindsBy(How = How.XPath, Using = CONFIRMBUTTON)]
        private IWebElement _confirmbutton;
        [FindsBy(How = How.XPath, Using = REPLACE)]
        private IWebElement _replace;
        [FindsBy(How = How.XPath, Using = FIRSTLINE)]
        private IWebElement _firstline;
        [FindsBy(How = How.XPath, Using = INPUTE_REPLACE_ITEM)]
        private IWebElement _inputreplaceitem;
        [FindsBy(How = How.XPath, Using = SELECTALL)]
        private IWebElement _selectall;
        [FindsBy(How = How.XPath, Using = FIRST_ITEM_UNIT)]
        private IWebElement _firstItemUnit;
        [FindsBy(How = How.LinkText, Using = CREATE_NEW_ITEM)]
        private IWebElement _createNewItem;

        [FindsBy(How = How.XPath, Using = FOLD_ALL)]
        private IWebElement _foldAll;

        [FindsBy(How = How.XPath, Using = UNFOLD_ALL)]
        private IWebElement _unfoldAll;

        [FindsBy(How = How.Id, Using = EXPORT_BTN)]
        private IWebElement _export;

        [FindsBy(How = How.Id, Using = CONFIRM_EXPORT_BTN)]
        private IWebElement _confirmDownload;

        [FindsBy(How = How.Id, Using = IMPORT)]
        private IWebElement _import;

        [FindsBy(How = How.XPath, Using = PRINT_THEORICAL_REPORT)]
        private IWebElement _printTheoricalReport;

        // Tableau items
        [FindsBy(How = How.XPath, Using = PURCHASE_TAB)]
        private IWebElement _purchaseTab;

        [FindsBy(How = How.XPath, Using = STORAGE_TAB)]
        private IWebElement _storageTab;

        [FindsBy(How = How.XPath, Using = FIRST_ITEM)]
        private IWebElement _firstItem;

        [FindsBy(How = How.XPath, Using = CONTENT)]
        private IWebElement _content;
        [FindsBy(How = How.XPath, Using = FIRST_ITEM_UNIT_PATCH)]
        private IWebElement _firstItemUnit_Patch;

        [FindsBy(How = How.XPath, Using = FIRST_ITEM_GROUP)]
        private IWebElement _firstItemGroup;

        [FindsBy(How = How.Id, Using = INTOLERANCE_TAB)]
        private IWebElement _intoleranceTab;

        [FindsBy(How = How.XPath, Using = CLICK_DETAIL)]
        private IWebElement _clickDetail;

        [FindsBy(How = How.XPath, Using = UNIT_PRICE)]
        private IWebElement _unitPrice;

        // ___________________________________________ Filtres __________________________________________________

        [FindsBy(How = How.XPath, Using = RESET_FILTER_DEV)]
        private IWebElement _resetFilterDev;

        [FindsBy(How = How.Id, Using = SEARCH_FILTER)]
        private IWebElement _searchFilter;

        [FindsBy(How = How.Id, Using = KEYWORD_FILTER)]
        private IWebElement _keywordFilter;

        [FindsBy(How = How.XPath, Using = UNCHECK_ALL_KEYWORD)]
        private IWebElement _uncheckAllKeyword;

        [FindsBy(How = How.XPath, Using = SEARCH_KEYWORD)]
        private IWebElement _searchKeyword;

        [FindsBy(How = How.Id, Using = SEARCH_SUPPLIER_REF_FILTER)]
        private IWebElement _searchSupplierRefFilter;

        [FindsBy(How = How.Id, Using = SORT_BY)]
        private IWebElement _sortBy;

        [FindsBy(How = How.Id, Using = SHOW_ALL_DEV)]
        private IWebElement _showAllDev;

        [FindsBy(How = How.Id, Using = SHOW_ONLY_ACTIVE_DEV)]
        private IWebElement _showOnlyActiveDev;

        [FindsBy(How = How.Id, Using = SHOW_ONLY_INACTIVE_DEV)]
        private IWebElement _showOnlyInactiveDev;

        [FindsBy(How = How.Id, Using = ACTIVE_NOT_PURCHASABLE_DEV)]
        private IWebElement _activeNotPurchsableDev;

        [FindsBy(How = How.Id, Using = SITE_FILTER)]
        private IWebElement _siteFilter;

        [FindsBy(How = How.XPath, Using = UNCHECK_ALL_SITES)]
        private IWebElement _uncheckAllSites;

        [FindsBy(How = How.XPath, Using = SEARCH_SITE)]
        private IWebElement _searchSite;

        [FindsBy(How = How.Id, Using = ALLERGENS_FILTER)]
        private IWebElement _allergensFilter;

        [FindsBy(How = How.XPath, Using = UNCHECK_ALL_ALLERGENS)]
        private IWebElement _uncheckAllAllergens;

        [FindsBy(How = How.XPath, Using = SEARCH_ALLERGENS)]
        private IWebElement _searchAllergen;

        [FindsBy(How = How.Id, Using = GROUP_FILTER)]
        private IWebElement _groupFilter;

        [FindsBy(How = How.XPath, Using = FIRST_ITEMSO)]
        private IWebElement _firstItemso;

        [FindsBy(How = How.XPath, Using = MASSIVE_DELETE)]
        private IWebElement _massiveDelete;
        [FindsBy(How = How.XPath, Using = FIRST_ITEM_ACTION)]
        private IWebElement _firstItemAction;
        public enum FilterType
        {
            Search,
            Keyword,
            SupplierRef,
            SortBy,
            ShowAll,
            ShowOnlyActive,
            ShowOnlyInactive,
            ActiveNotPurchasable,
            Site,
            Supplier,
            Allergens,
            Group,
            SubGroup,
            ShowStorageInAlertThreshold
        }

        public void Filter(FilterType filterType, object value, bool clearAllPreviousSelection = true, bool ignoreAutoSuggest = false)
        {
            Actions action = new Actions(_webDriver);
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;

            switch (filterType)
            {
                case FilterType.Search:
                    _searchFilter = WaitForElementIsVisible(By.Id(SEARCH_FILTER));
                    //_searchFilter.Clear();
                    _searchFilter.SetValue(ControlType.TextBox, value);
                    if (ignoreAutoSuggest == false && isElementVisible(By.XPath(String.Format(FIRST_RESULT_SEARCH, value))))
                    {
                        _searchFilter.SendKeys(Keys.Tab);
                    }
                    WaitForLoad();

                    break;
                case FilterType.Keyword:
                    _keywordFilter = WaitForElementIsVisible(By.Id(KEYWORD_FILTER));
                    _keywordFilter.Click();

                    _uncheckAllKeyword = WaitForElementIsVisible(By.XPath(UNCHECK_ALL_KEYWORD));
                    _uncheckAllKeyword.Click();

                    _searchKeyword = WaitForElementIsVisible(By.XPath(SEARCH_KEYWORD));
                    _searchKeyword.SetValue(ControlType.TextBox, value);

                    var keywordToCheck = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    keywordToCheck.SetValue(ControlType.CheckBox, true);

                    _keywordFilter.Click();
                    WaitForLoad();
                    break;
                case FilterType.SupplierRef:
                    _searchSupplierRefFilter = WaitForElementIsVisible(By.Id(SEARCH_SUPPLIER_REF_FILTER));
                    _searchSupplierRefFilter.SetValue(ControlType.TextBox, value);
                    break;
                case FilterType.SortBy:
                    _sortBy = WaitForElementIsVisible(By.Id(SORT_BY));
                    _sortBy.Click();
                    var element = WaitForElementIsVisible(By.XPath("//*[@id=\"cbSortBy\"]/option[@value='" + value + "']"));
                    _sortBy.SetValue(ControlType.DropDownList, element.Text);
                    _sortBy.Click();
                    break;
                case FilterType.ShowAll:
                    _showAllDev = WaitForElementExists(By.Id(SHOW_ALL_DEV));
                    _showAllDev.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.ShowOnlyActive:
                    _showOnlyActiveDev = WaitForElementExists(By.Id(SHOW_ONLY_ACTIVE_DEV));
                    _showOnlyActiveDev.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.ShowOnlyInactive:
                    _showOnlyInactiveDev = WaitForElementExists(By.Id(SHOW_ONLY_INACTIVE_DEV));
                    _showOnlyInactiveDev.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.ActiveNotPurchasable:
                    _activeNotPurchsableDev = WaitForElementExists(By.Id(ACTIVE_NOT_PURCHASABLE_DEV));
                    _activeNotPurchsableDev.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.Site:
                    if (value is bool)
                    {
                        ComboBoxOptions cbOpt = new ComboBoxOptions(SITE_FILTER, null) { ClickCheckAllAtStart = true };
                        ComboBoxSelectById(cbOpt);
                    }
                    else
                    {
                        ComboBoxOptions cbOpt = new ComboBoxOptions(SITE_FILTER, (string)value) { ClickUncheckAllAtStart = clearAllPreviousSelection };
                        ComboBoxSelectById(cbOpt);
                    }
                    break;
                case FilterType.Supplier:
                    if (value is bool)
                    {
                        ComboBoxOptions cbOpt = new ComboBoxOptions(SUPPLIER_FILTER, null) { ClickCheckAllAtStart = true };
                        ComboBoxSelectById(cbOpt);
                    }
                    else
                    {
                        ComboBoxOptions cbOpt = new ComboBoxOptions(SUPPLIER_FILTER, (string)value) { ClickCheckAllAtStart = false, ClickUncheckAllAtStart = true };
                        ComboBoxSelectById(cbOpt);
                    }
                    break;
                case FilterType.Allergens:
                    ScrollUntilElementIsInView(By.Id(ALLERGENS_FILTER));
                    ComboBoxSelectById(new ComboBoxOptions(ALLERGENS_FILTER, (string)value));
                    break;
                case FilterType.Group:
                    //scroll down
                    action.MoveToElement(WaitForElementExists(By.Id(GROUP_FILTER))).Perform();
                    _groupFilter = WaitForElementIsVisible(By.Id(GROUP_FILTER));
                    javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", _groupFilter);

                    if (value is bool)
                    {
                        ComboBoxOptions cbOpt = new ComboBoxOptions(GROUP_FILTER, null) { ClickCheckAllAtStart = true };
                        ComboBoxSelectById(cbOpt);
                    }
                    else
                    {
                        ComboBoxOptions cbOpt = new ComboBoxOptions(GROUP_FILTER, (string)value) { ClickUncheckAllAtStart = true };
                        ComboBoxSelectById(cbOpt);
                    }
                    break;
                case FilterType.SubGroup:
                    // scroll down
                    action.MoveToElement(WaitForElementExists(By.Id(SUBGROUP_FILTER))).Perform();
                    ComboBoxSelectById(new ComboBoxOptions(SUBGROUP_FILTER, (string)value));
                    action.MoveToElement(WaitForElementExists(By.Id(RESET_FILTER_DEV))).Perform();
                    break;
                case FilterType.ShowStorageInAlertThreshold:
                    var input = WaitForElementExists(By.Id(STORAGE_THRESHOLD_FILTER));
                    input.SetValue(ControlType.CheckBox, value);
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);

            }
            WaitPageLoading();
            WaitForLoad();
        }

        public void DoubleFilterSite(string site1, string site2)
        {
            ComboBoxSelectById(new ComboBoxOptions(SITE_FILTER, site1));
            ComboBoxOptions cbOpt = new ComboBoxOptions(SITE_FILTER, site2) { ClickUncheckAllAtStart = false };
            ComboBoxSelectById(cbOpt);
        }

        public void ResetFilter()
        {
            _resetFilterDev = WaitForElementIsVisibleNew(By.Id(RESET_FILTER_DEV));
            _resetFilterDev.Click();

            WaitForLoad();

            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                // pas de date
            }
        }

        public bool IsSortedByName()
        {
            var listeNames = _webDriver.FindElements(By.XPath(NAME));

            if (listeNames.Count == 0)
                return false;

            var ancientName = listeNames[0].Text;

            foreach (var elm in listeNames)
            {
                if (elm != null && !string.IsNullOrEmpty(elm.Text) && ancientName != null)
                {
                    if (elm.Text.CompareTo(ancientName) >= 0)
                    {
                        ancientName = elm.Text;
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    return false;
                }
            }

            return true;
        }

        public bool IsSortedByItemGroup()
        {
            var listeGroups = _webDriver.FindElements(By.XPath(GROUP));

            if (listeGroups.Count == 0)
                return false;

            var ancientGroup = listeGroups[0].Text;

            foreach (var elm in listeGroups)
            {
                if (elm != null && !string.IsNullOrEmpty(elm.Text) && ancientGroup != null)
                {
                    if (elm.Text.CompareTo(ancientGroup) >= 0)
                    {
                        ancientGroup = elm.Text;
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    return false;
                }
            }

            return true;
        }

        public bool VerifyGroup(string group)
        {
            var customers = _webDriver.FindElements(By.XPath(GROUP));

            if (customers.Count == 0)
                return false;

            foreach (var elm in customers)
            {
                if (elm.Text != group)
                {
                    return false;
                }
            }
            return true;
        }

        public bool CheckStatus(bool active)
        {
            bool isActive = false;
            int tot = CheckTotalNumber() > 100 ? 100 : CheckTotalNumber();

            if (tot == 0)
                return false;

            for (int i = 1; i <= tot; i++)
            {
                if (isElementVisible(By.XPath(String.Format(INACTIVE, i + 1))))
                {
                    _webDriver.FindElement(By.XPath(String.Format(INACTIVE, i + 1)));

                    if (active)
                        return false;
                }
                else
                {
                    isActive = true;
                    if (!active)
                        return true;
                }
            }
            return isActive;
        }

        // ___________________________________________ Méthodes _________________________________________

        // General
        public ItemCreateModalPage ItemCreatePage()
        {
            ShowPlusMenu();

            _createNewItem = WaitForElementIsVisibleNew(By.LinkText(CREATE_NEW_ITEM));
            _createNewItem.Click();
            WaitForLoad();
            Thread.Sleep(1000); //lenteur ouverture popup

            return new ItemCreateModalPage(_webDriver, _testContext);
        }

        public void FoldAll()
        {
            WaitPageLoading();
            ShowExtendedMenu();
            WaitForLoad();
            _foldAll = WaitForElementIsVisible(By.XPath(FOLD_ALL));
            _foldAll.Click();
            WaitPageLoading();
        }
        public void SelectALL()
        {

            _selectall = WaitForElementIsVisible(By.XPath(SELECTALL));
            _selectall.Click();
            WaitForLoad();
        }

        public void ReplaceWithAnOtherItem()
        {

            _replace = WaitForElementIsVisible(By.XPath(REPLACE));
            _replace.Click();
            WaitForLoad();
        }
        public void InputReplaceWithAnOtherItem(string value)
        {

            _inputreplaceitem = WaitForElementIsVisible(By.XPath(INPUTE_REPLACE_ITEM));
            _inputreplaceitem.SetValue(ControlType.TextBox, value);
            _inputreplaceitem.Click();

        }
        public void FirstLineClick()
        {
            _firstline = WaitForElementIsVisible(By.XPath(FIRSTLINE));
            _firstline.Click();
            WaitForLoad();
        }

        public void Replaceitem()
        {
            _replacewithitem = WaitForElementIsVisible(By.XPath(REPLACEWITHITEM));
            _replacewithitem.Click();
            WaitForLoad();
        }
        public bool ConfirmeButton()
        {
            int i = 20;
            _confirmbutton = WaitForElementIsVisible(By.XPath(CONFIRMBUTTON));
            _confirmbutton.Click();
            while (i > 0)
            {
                if (isElementVisible(By.XPath(SUCCESS)))
                {
                    return true;

                }
                else
                {
                    WaitPageLoading();
                    i--;
                }
            }
            return false;
        }

        public void CloseButton()
        {
            _close = WaitForElementIsVisible(By.XPath(CLOSE));
            _close.Click();
            WaitForLoad();
        }
        public void CancelButton()
        {
            _cancel = WaitForElementIsVisible(By.XPath(CANCEL));
            _cancel.Click();
            WaitForLoad();
        }
        public bool IsFoldAll()
        {
            _content = WaitForElementExists(By.XPath(CONTENT));

            // Temps nécessaire pour que l'élément change de classe
            Thread.Sleep(1000);

            if (_content.GetAttribute("class") == "panel-collapse collapse")
                return true;

            return false;
        }

        public void UnfoldAll()
        {
            ShowExtendedMenu();
            WaitPageLoading();
            WaitForLoad();
            _unfoldAll = WaitForElementIsVisible(By.XPath(UNFOLD_ALL));
            _unfoldAll.Click();
            WaitPageLoading();
        }

        public Boolean IsUnfoldAll()
        {
            _content = WaitForElementIsVisible(By.XPath(CONTENT));

            // Temps nécessaire pour que l'élément change de classe
            Thread.Sleep(1000);
            if (_content.GetAttribute("class") == "panel-collapse collapse show")
                return true;

            return false;
        }

        public void Export(bool versionPrint)
        {
            ShowExtendedMenu();

            _export = WaitForElementIsVisibleNew(By.Id(EXPORT_BTN));
            _export.Click();
            WaitForLoad();

            if (IsDev())
            {
                _confirmDownload = WaitForElementToBeClickable(By.Id("btn-export"));
            }
            else
            {
                _confirmDownload = WaitForElementToBeClickable(By.Id(CONFIRM_EXPORT_BTN));
            }
            _confirmDownload.Click();
            WaitPageLoading();

            if (versionPrint)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));
                ClosePrintButton();
            }

            WaitForDownload();
            Close();
        }

        public FileInfo GetExportExcelFile(FileInfo[] taskFiles, bool isStorage = false)
        {
            List<FileInfo> correctDownloadFiles = new List<FileInfo>();
            foreach (FileInfo file in taskFiles)
                if (IsExcelFileCorrect(file.Name, isStorage))
                    correctDownloadFiles.Add(file);
            if (correctDownloadFiles.Count <= 0) return null;
            DateTime time = correctDownloadFiles[0].LastWriteTimeUtc;
            FileInfo correctFile = correctDownloadFiles[0];
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
        public bool IsExcelFileCorrect(string filePath, bool isStorage = false)
        {
            // "PurchaseBook 2019-11-22 12-27-20.xlsx"
            // "StorageBook 2023-03-15 15-20-38.xlsx"
            string mois = "(?:0[1-9]|1[0-2])";         // mois
            string space = "(\\s)";                    // Espace
            string annee = "\\d{4}";                   // annee YYYY
            string jour = "[0-3]\\d";                  // jour
            string heure = "(?:0[0-9]|1[0-9]|2[0-3])"; // heure
            string minutes = "[0-5]\\d";               // minutes
            string secondes = "[0-5]\\d";              // secondes
            string pre = isStorage ? "StorageBook" : "PurchaseBook";
            Regex r = new Regex("^" + pre + space + annee + "-" + mois + "-" + jour + space + heure + "-" + minutes + "-" + secondes + ".xlsx$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = r.Match(filePath);

            return m.Success;
        }

        // Function to wait for the Excel file to be ready for reading
        public XLWorkbook WaitForExcelFile(string path)
        {
            int retryCount = 5;
            int retryDelay = 2000;

            while (retryCount > 0)
            {
                try
                {
                    return new XLWorkbook(path);
                }
                catch (IOException)
                {
                    // Wait before retrying
                    Thread.Sleep(retryDelay);
                    retryCount--;
                }
            }
            throw new IOException($"Le fichier Excel '{path}' n'a pas pu être ouvert après plusieurs tentatives.");
        }

        public ItemImportPage Import()
        {
            ShowExtendedMenu();

            //WaitPageLoading();
            WaitForLoad();
            try
            {
                _import = WaitForElementIsVisible(By.Id(IMPORT));
                _import.Click();
                WaitPageLoading();
                WaitForLoad();
            }
            catch (UnhandledAlertException ex)
            {
                IAlert alert = _webDriver.SwitchTo().Alert();
                Console.WriteLine("Alert text: " + alert.Text);
                alert.Dismiss();

                _import = WaitForElementIsVisible(By.Id(IMPORT));
                _import.Click();
                WaitPageLoading();
                WaitForLoad();
            }
            return new ItemImportPage(_webDriver, _testContext);

        }

        public void ClickOnPurchaseTab()
        {
            _purchaseTab = WaitForElementIsVisible(By.XPath(PURCHASE_TAB));
            _purchaseTab.Click();
            WaitForLoad();
        }

        public ItemStoragePage ClickOnStorage()
        {
            _storageTab = WaitForElementIsVisible(By.XPath(STORAGE_TAB));
            _storageTab.Click();
            WaitForLoad();
            return new ItemStoragePage(_webDriver, _testContext);

        }
        public ItemIntolerancePage ClickOnIntolerance()
        {
            _intoleranceTab = WaitForElementIsVisibleNew(By.Id(INTOLERANCE_TAB));
            _intoleranceTab.Click();
            WaitForLoad();
            return new ItemIntolerancePage(_webDriver, _testContext);

        }
        public ItemStoragePage ClickOnStorageItem()
        {
            _storageTab = WaitForElementIsVisible(By.Id(STORAGE_ITEM));
            _storageTab.Click();
            WaitForLoad();
            return new ItemStoragePage(_webDriver, _testContext);

        }

        // Tableau items
        public ItemGeneralInformationPage ClickOnFirstItem()
        {
            WaitPageLoading();
            WaitForLoad();
            _firstItem = WaitForElementExists(By.XPath(FIRST_ITEM));
            _firstItem.Click();
            WaitForLoad();

            return new ItemGeneralInformationPage(_webDriver, _testContext);
        }

        public void SetProdQty(string newProdQty)
        {
            var itemRow = WaitForElementIsVisible(By.XPath("//*[@id=\"item_SodRowDto_Quantity\"]"));
            itemRow.Clear();
            itemRow.SendKeys(newProdQty);
        }

        public bool VerifyFDNotRounded()
        {
            WaitPageLoading();
            var itemRow = WaitForElementIsVisible(By.XPath("//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[1]/div[9]"));

            var fAndDValue = WaitForElementIsVisible(By.XPath("//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[1]/div[9]/span")).Text;
            var prodQtyValue = itemRow.FindElement(By.XPath("//*[@id=\"item_SodRowDto_Quantity\"]")).GetAttribute("value");

            decimal fAndD = Convert.ToDecimal(fAndDValue);
            decimal prodQty = Convert.ToDecimal(prodQtyValue);

            return Math.Round(fAndD) != prodQty;
        }


        public string GetFirstItemName()
        {
            WaitPageLoading();
            if (isElementVisible(By.XPath(FIRST_ITEM_NAME)))
            {
                var firstItemName = WaitForElementIsVisible(By.XPath(FIRST_ITEM_NAME));
                return firstItemName.Text;
            }
            else
            {
                return string.Empty;
            }
        }
        public string GetreciepnametName()
        {

            _firstItem = WaitForElementIsVisible(By.XPath(INGREDIENTNAME));
            return _firstItem.Text;

        }
        public bool VerifyProdQty()
        {
            if (isElementExists(By.XPath(PRODQUANTITYERROR)))
            {
                return true;

            }
            else
            {
                return false;

            }
        }
        public bool VerifyProdWeight()
        {

            if (isElementExists(By.XPath(PRODWEIGHT)))
            {
                return true;

            }
            else
            {
                return false;

            }
        }
        public bool VerifyStorageQty()
        {

            if (isElementExists(By.XPath(STORAGEQUANTITY)))
            {
                return true;

            }
            else
            {
                return false;

            }
        }
        public PrintReportPage PrintTheoricalReport(string site)
        {
            ShowExtendedMenu();
            _printTheoricalReport = WaitForElementIsVisible(By.XPath(PRINT_THEORICAL_REPORT));
            _printTheoricalReport.Click();
            WaitForLoad();

            var selectSite = WaitForElementIsVisible(By.Id("SelectedSite"));
            var select = new SelectElement(selectSite);
            select.SelectByText(site);
            WaitForLoad();

            var printButton = WaitForElementIsVisible(By.Id("printButton"));
            printButton.Click();
            WaitForLoad();

            IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-pdf']"));
            ClickPrintButton();

            //Results are opened in a new tab, switch the driver to the newly created one
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new PrintReportPage(_webDriver, _testContext);
        }
        public BestPriceModal MenuBestPrice()
        {
            ShowExtendedMenu();
            var bestPrice = WaitForElementIsVisible(By.XPath("//*[@id='tabContentItemContainer']/div[1]/div/div[1]/div/a[5]"));
            bestPrice.Click();
            WaitForLoad();
            return new BestPriceModal(_webDriver, _testContext);
        }
        public void ClickOnFirstPackagingAndSetLimitQty(string limitQty)
        {
            var firstPackaging = WaitForElementIsVisible(By.XPath(FIRST_PACKAGING));
            firstPackaging.Click();
            var limitQtyInput = WaitForElementIsVisible(By.Id(LIMIT_QTY));
            limitQtyInput.Clear();
            limitQtyInput.SendKeys(limitQty);
        }
        public void ClickOnStorageTab()
        {
            var storagetab = WaitForElementIsVisible(By.Id(STORAGE));
            storagetab.Click();
        }

        public List<string> GetItemsNames()
        {
            List<String> liste = new List<string>();
            var _items = _webDriver.FindElements(By.XPath("//*/td[@class='primary-information']"));
            foreach (var item in _items)
            {
                liste.Add(item.Text);
            }
            return liste;
        }

        public ItemGeneralInformationPage ClickOnItem(int itemOffset)
        {
            var item = WaitForElementIsVisible(By.XPath("(//*/td[@class='primary-information'])[" + (itemOffset + 1) + "]"));
            item.Click();
            return new ItemGeneralInformationPage(_webDriver, _testContext);
        }
        public int SUMAllergensCommunes2Items(IEnumerable<string> list1, IEnumerable<string> list2)
        {
            IEnumerable<string> distinctStrings = list1.Union(list2);
            return distinctStrings.Count();
        }

        public MassiveDeleteModal MenuMassiveDelete()
        {
            ShowExtendedMenu();
            if (IsDev())
            {
                _massiveDelete = WaitForElementIsVisibleNew(By.XPath(MASSIVE_DELETE));
                _massiveDelete.Click();
            }
            else
            {
                _massiveDelete = WaitForElementIsVisibleNew(By.XPath("//*/a[text()='Massive delete']"));
                WaitForLoad();
                _massiveDelete.Click();
            }

            return new MassiveDeleteModal(_webDriver, _testContext);
        }
        public string GetFirstItemUnit()
        {
            if (IsDev())
            {
                _firstItemUnit = WaitForElementIsVisible(By.XPath(FIRST_ITEM_UNIT));
                return _firstItemUnit.Text;
            }
            else
            {
                _firstItemUnit_Patch = WaitForElementIsVisible(By.XPath(FIRST_ITEM_UNIT_PATCH));
                return _firstItemUnit_Patch.Text;
            }

        }
        public ItemGeneralInformationPage ClickOnFirstItems()
        {
            var _firstItem = WaitForElementIsVisible(By.XPath(FIRST_ITEM_COL));
            _firstItem.Click();
            WaitForLoad();
            return new ItemGeneralInformationPage(_webDriver, _testContext);
        }
        public void DownloadFirstFile()
        {
            _downloadFirstRowButton = WaitForElementIsVisible(By.XPath(DOWNLOADS_FIRST_ROW));
            _downloadFirstRowButton.Click();
        }

        public void ClickOnLabelTab()
        {
            var labelTab = _webDriver.FindElement(By.XPath(LABEL_TAB));
            labelTab.Click();
            WaitForLoad();
        }

        public void ClickOnUseCase()
        {
            var UseCaseTab = _webDriver.FindElement(By.XPath(USECASE_TAB));
            UseCaseTab.Click();
            WaitForLoad();
        }
        public void ClickOngeneralInformation()
        {
            var generalInformation = _webDriver.FindElement(By.XPath(GENERAL_INFORMATION_TAB));
            generalInformation.Click();
            WaitForLoad();
        }

        public void EnterComment(string comment)
        {
            var commentField = _webDriver.FindElement(By.XPath(COMMENT_FIELD));
            commentField.Clear();
            commentField.SendKeys(comment);
        }

        public void ClickOnValidateButton()
        {
            var validateButton = _webDriver.FindElement(By.XPath(VALIDATE_BUTTON));
            validateButton.Click();
            WaitForLoad();
        }

        public string GetCommentText()
        {
            var commentElement = _webDriver.FindElement(By.XPath(COMMENT_FIELD));

            return commentElement.Text;
        }
        public void SearchItem(string itemName)
        {
            WaitPageLoading();
            var searchBox = _webDriver.FindElement(By.Id(SEARCH_ITEMSO));
            searchBox.SendKeys(itemName);
            searchBox.SendKeys(Keys.Enter);
        }

        public string GetItemQuantity()
        {
            WaitPageLoading();
            var itemQtyElement = _webDriver.FindElement(By.XPath(ITEM_QTY_ITEM));
            ScrollToElement(itemQtyElement);
            return itemQtyElement.Text.Trim();
        }

        public string GetItemUnit()
        {
            WaitPageLoading();
            var itemUnitElement = _webDriver.FindElement(By.XPath(ITEM_UNIT));
            ScrollToElement(itemUnitElement);
            return itemUnitElement.Text.Trim();
        }

        public bool ReplacedSuccess()
        {
            WaitPageLoading();

            if (isElementVisible(By.XPath(SUCCESS)))
            {
                return true;

            }
            else { return false; }
        }
        public bool IsAllResultsConatinsSameSite(string colmunnName, string sheetName, string fileName, string siteName)
        {
            int counter = 0;
            var listofSite = OpenXmlExcel.GetValuesInList(colmunnName, sheetName, fileName);
            if (listofSite.Any())
            {
                counter = listofSite.Where(c => c.Equals(siteName)).Count();

            }
            return counter == listofSite.Count();

        }
        public string GetFirstItem()
        {
            if (isElementVisible(By.XPath(FIRST_ITEM_ACTION)))
            {
                _firstItemAction = WaitForElementIsVisible(By.XPath(FIRST_ITEM_ACTION));
                return _firstItemAction.Text;
            }
            else
            {
                return "";
            }
        }
        public string GetFirstItemGroup()
        {
            _firstItemGroup = WaitForElementIsVisible(By.XPath(FIRST_ITEM_GROUP));
            return _firstItemGroup.Text;
        }
        public double GetTotalColFromExcel(List<string> colonneThVal, List<string> colonnePlace, string place, string decimalSeparator)
        {
            double thValB = 0;
            CultureInfo ci = decimalSeparator.Equals(",") ? CultureInfo.GetCultureInfo("fr-FR") : CultureInfo.GetCultureInfo("en-US");

            for (int i = 0; i < colonnePlace.Count; i++)
            {
                if (colonnePlace[i].ToUpper() == place.ToUpper())
                {
                    string sanitizedValue = colonneThVal[i];

                    if (double.TryParse(sanitizedValue, NumberStyles.Any, ci, out double parsedValue))
                    {
                        thValB += (double)parsedValue;
                    }
                }
            }
            return thValB;
        }
        public string GetFirstItemNumber()
        {
            WaitPageLoading();
            if (isElementVisible(By.XPath(FIRST_ITEM_NUMBER)))
            {
                var firstItemNumber = WaitForElementIsVisible(By.XPath(FIRST_ITEM_NUMBER));
                return firstItemNumber.Text;
            }
            else
            {
                return string.Empty;
            }
        }
        public bool IsItemExist(string itemName)
        {
            if (!isElementVisible(By.XPath(string.Format(FIRST_ITEM_ACTION, itemName))))
            {
                return false;
            }
            return true;
        }

        public Point GetCheckBoxPosition(IWebElement itemDetails, string checkBoxName)
        {
            return itemDetails.FindElement(By.XPath($"//input[@type='checkbox'][@name='{checkBoxName}']")).Location;
        }

        public Point GetLabelPosition(IWebElement itemDetails, string labelText)
        {
            return itemDetails.FindElement(By.XPath($"//th[contains(text(), '{labelText}')]")).Location;
        }

        public bool AreCheckBoxesAligned(int roundToUsY, int purchasableY, int unavailableY)
        {
            return roundToUsY == purchasableY && roundToUsY == unavailableY;
        }

        public IWebElement GetItemDetailsContainer()
        {
            return WaitForElementExists(By.XPath("//*[@id=\"detailsItemContainer\"]/div/table/tbody/tr"));
        }

        public ItemGeneralInformationPage ClickOnItemDetails()
        {
            var _clickDetail = WaitForElementIsVisible(By.XPath(CLICK_DETAIL));
            _clickDetail.Click();
            WaitForLoad();
            return new ItemGeneralInformationPage(_webDriver, _testContext);
        }

        public bool GetPriceCurrency(string decimalSeparator, string currency)
        {
            WaitPageLoading();
            WaitForLoad();
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");
            _unitPrice = WaitForElementExists(By.XPath(UNIT_PRICE));

            var price = _unitPrice.GetAttribute("innerText").Replace(currency, "").Trim();
            var currencyVisible = _unitPrice.GetAttribute("innerText").Replace(price, "").Trim();
            if (currencyVisible == currency)
            {
                return true;
            }
            else return false;

        }
        public void DeactivateNotPurchasable()
        {
            ShowExtendedMenu();

            WaitPageLoading();
            WaitForLoad();
            try
            {
                _import = WaitForElementIsVisible(By.Id(DEACTIVATED));
                _import.Click();
                WaitPageLoading();
                WaitForLoad();
            }
            catch (UnhandledAlertException ex)
            {
                IAlert alert = _webDriver.SwitchTo().Alert();
                Console.WriteLine("Alert text: " + alert.Text);
                alert.Dismiss();

                _import = WaitForElementIsVisible(By.Id(DEACTIVATED));
                _import.Click();
                WaitPageLoading();
                WaitForLoad();
            }

        }
        public bool IsItemsDeactivationReportPopupVisible()
        {
            try
            {
                var popup = _webDriver.FindElement(By.Id("formId"));
                return popup.Displayed;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }

        public bool AreCancelAndSaveButtonsVisible()
        {
            WaitPageLoading();
            try
            {
                var popup = _webDriver.FindElement(By.Id(POPUP_DEACTIVATED));

                var cancelButton = popup.FindElement(By.XPath(CANCEL_BTN));
                var saveButton = popup.FindElement(By.XPath(SAVE_BTN));

                return cancelButton.Displayed && saveButton.Displayed;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }
        public void ClickFlecheAGauche()
        {
            // Wait for the element to be clickable
            var flecheElement = _webDriver.FindElement(By.XPath("/html/body/div[2]/div/div[2]/div[2]/div/div/div/div[2]/div[1]/div/div[1]"));
            flecheElement.Click();
            WaitForLoad();
        }
        //public bool ArePackagingsGrouped(string site, string supplier)
        //{
        //    var packagingRows = _webDriver.FindElements(By.XPath("//*[@id='content_72798']/div/table/tbody[1]/tr/th[8]"));
        //    foreach (var row in packagingRows)
        //    {
        //        var rowSite = row.FindElement(By.XPath("//*[@id='content_72798']/div/table/tbody[1]/tr/th[2]")).Text;
        //        var rowSupplier = row.FindElement(By.XPath("//*[@id='content_72798']/div/table/tbody[1]/tr/th[3]")).Text;
        //        if (rowSite != site || rowSupplier != supplier)
        //        {
        //            return false;
        //        }
        //    }
        //    return true;
        //}
        public string GetSite()
        {
            var site = WaitForElementIsVisible(By.XPath("/html/body/div[2]/div/div[2]/div[2]/div/div/div/div[2]/div[2]/div/table/tbody[2]/tr[1]/td[2]/span"));
            return site.Text;
        }
        public string GetSupplier1()
        {
            var site = WaitForElementIsVisible(By.XPath("/html/body/div[2]/div/div[2]/div[2]/div/div/div/div[2]/div[2]/div/table/tbody[2]/tr[1]/td[3]"));
            return site.Text;
        }
        public string GetSupplier()
        {
            var site = WaitForElementIsVisible(By.XPath("/html/body/div[2]/div/div[2]/div[2]/div/div/div/div[2]/div[2]/div/table/tbody[2]/tr[2]/td[3]"));
            return site.Text;
        }
        public void clickOnItemCheck()
        {
            var itemCheckbox = WaitForElementIsVisible(By.Id(YIELDCHOICEITEM));
            itemCheckbox.Click();
        }

        public void ClickOnFirstItemSupplyOrder()
        {
            WaitPageLoading();
            WaitForLoad();
            _firstItemso = WaitForElementExists(By.XPath(FIRST_ITEMSO));
            _firstItemso.Click();
            WaitForLoad();

        }

        public ItemGeneralInformationPage ClickFirstItem()
        {
            var firstItem = WaitForElementIsVisible(By.Id(FIRST_I));
            firstItem.Click();
            WaitForLoad();
            return new ItemGeneralInformationPage(_webDriver, _testContext);
        }
        public bool IsElementVisible()
        {
            // Attente d'un élément visible à partir du XPath spécifié
            var element = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div/div/div/form/div[2]/table/tbody/tr/td/ul/li"));

            if (element != null && element.Displayed)
            {
                return true;  // Si l'élément est trouvé et visible
            }
            else
            {
                return false;  // Si l'élément n'est pas trouvé ou n'est pas visible
            }
        }

        public void BackToList()
        {
            var BackToListBtn = _webDriver.FindElement(By.XPath(BACK_BTN));
            BackToListBtn.Click();
            WaitForLoad();
        }
        public string GetSupplierFromIndex()
        {
            var site = WaitForElementIsVisible(By.XPath("/html/body/div[2]/div/div[2]/div[2]/div/div/div/div[2]/div[2]/div/table/tbody[2]/tr/td[3]"));
            return site.Text;
        }
        public List<string> GetAllSites()
        {
            var tableSites = _webDriver.FindElements(By.XPath("/html/body/div[2]/div/div[2]/div[2]/div/div/div/div[2]/div[2]/div/table/tbody[2]/tr[*]/td[2]/span"));
            List<string> liste = new List<string>();
            foreach (var tableSite in tableSites)
            {
                liste.Add(tableSite.Text);
            }
            return liste;
        }
        public List<string> GetAllSitesAfterClickFlech()
        {
            var tableSites = _webDriver.FindElements(By.XPath("/html/body/div[2]/div/div[2]/div[2]/div/div/div/div[2]/div[2]/div/table/tbody[2]/tr[*]/td[2]/span"));
            List<string> liste = new List<string>();
            foreach (var tableSite in tableSites)
            {
                liste.Add(tableSite.Text);
            }
            return liste;
        }
        public bool ItemDetailsIsVisible()
        {
            return isElementVisible(By.XPath("//*[starts-with(@id,\"content_\")]/div/table"));
        }
    }
}
