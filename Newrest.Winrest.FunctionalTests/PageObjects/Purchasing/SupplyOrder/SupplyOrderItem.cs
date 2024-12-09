using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.SupplyOrder;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.OutputForm;
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
using UglyToad.PdfPig;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Purchasing
{
    public class SupplyOrderItem : PageBase
    {

        public SupplyOrderItem(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }
        // ______________________________________ Constantes _______________________________________________

        // General
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";
        private const string EXTENDED_BUTTON = "//*[@id=\"div-body\"]/div/div[1]/div/div/button";
        private const string REFRESH = "btn-supply-orders-refresh";
        private const string EXPORT = "btn-export-excel";
        private const string IMPORT = "//*/a[@id='btn-export-excel']/following-sibling::a[1]";
        private const string ERASE_QUANTITY = "EraseQuantity";
        private const string CONFIRM_ERASE_QUANTITY = "dataConfirmOK";
        private const string PRINT = "btn-print-current-view";
        private const string MODAL_PRINT_BUTTON = "printButton";
        private const string VALIDATE_PO = "btn-validate-purchase-order";

        private const string COMPARE = "btn-compare";
        private const string ERASE_COMPARE = "btn-compare";
        private const string TEST_PACKAGING = "//*[@id=\"test\"]/div[2]/div/form/div[2]/div[*]";

        private const string VALIDATE_MENU = "//*[@id=\"div-body\"]/div/div[1]/div/div[2]/button";
        private const string VALIDATE = "btn-validate-supply-order";
        private const string CONFIRM_VALIDATE = "btn-popup-validate";

        private const string GENERATE_PO = "//*[@id=\"div-body\"]/div/div[1]/div/div/div/a[2]";
        private const string GENERATE_OF = "//*[@id=\"div-body\"]/div/div[1]/div/div/div/a[3]";
        private const string SELECT_SUPPLIERS = "generate-select-all-suppliers";
        private const string DELIVERY_DATE = "//*[starts-with(@id,\"datapicker-supplier-delivery-date-\")]";
        private const string ACTIVE_DAY_DEV = "//td[@class='day'][text()='{0}']";
        private const string GENERATE = "btn-generate";
        private const string VALIDATE_BUTTON = "//*[@id=\"div-body\"]/div/div[1]/div/div[2]/button";
        private const string QTEINPUTPO = "//*[@id=\"listOfItems\"]/div[2]/div[2]/div/div/form/div[2]/div[7]";
        // Onglets
        private const string GENERAL_INFORMATIONS = "hrefTabContentInformations";
        private const string RECIPES = "hrefTabContentRecipes";
        private const string BY_SUPPLIERS = "hrefTabContentBySupplier";

        // Tableau page Items
        private const string GROUP = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/span";
        private const string ITEM = "/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/form/div[2]";
        private const string ITEM_NAME = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[2]/div[3]/span";
        private const string ITEM_NAME2 = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span";
        private const string ITEM_NAME_VALIDATED = "//*/div[contains(@class,'supply-order-item-row-display')]/div[3]/span";
        private const string ITEM_SUPPLIER = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[2]/div[5]/span";
        private const string PACKING = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[2]/div[6]/span";
        private const string PACKING_VALIDATED = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div/div[6]/span";
        private const string QUANTITY_INPUT = "item_SodRowDto_Quantity";
        private const string QUANTITY_INPUT_PO = "item_PodRowDto_Quantity";
        private const string QUANTITY_INPUT_XPATH = "//div[contains(@class,\"edit-zone\")]//span[contains(@title,'{0}')]/../../div[9]/input";
        private const string ITEM_QUANTITY = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[2]/div[9]/span";
        private const string ITEM_QUANTITY_REFRESH = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[2]/div[9]/span";
        private const string ITEM_QUANTITY_VALIDATED = "//*/div[contains(@class,'supply-order-item-row-display')]/div[9]/span";
        private const string STORAGE_VALUE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[2]/div[12]/span";
        private const string PROD_QTY_VALUE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div/div[10]/span"; 
        private const string STORAGE_VALUE_FIRSTITEM = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[2]/div[12]/span";
        private const string LINE_ITEM = "/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/form/div[2]/div[3]/span";
        private const string EDIT_ITEM = "/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/form/div[1]/div[9]/input";
        private const string QTY = "/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/form/div[2]/div[9]/span";

        //private const string COMPARE_CHECK_BOX = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[contains(@title,'{0}')]/../../div[14]/input";
        private const string COMPARE_CHECK_BOX = "//*[@id='list-item-with-action']/div/div[2]/div[*]/div/div/form/div[contains(@class,'display-zone')]/div[3]/span[contains(@title,'{0}')]/../../div[14]/input";
        private const string EXTENDED_MENU = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[contains(@title,'{0}')]/../../div[16]/div/a/span";
        private const string DELETE_ICON = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[contains(@title,'{0}')]/../../div[16]/div/ul/li[*]/a/span[@class=\"glyphicon glyphicon-trash glyph-span\"]";
        private const string COMMENT_ICON = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[contains(@title,'{0}')]/../../div[16]/div/ul/li[*]/a/span[@class=\"glyphicon glyphicon-comment glyph-span\"]";
        private const string COMMENT = "myDetailComment";
        private const string SAVE_COMMENT = "//*[@id=\"modal-1\"]/div/div/div/div[2]/div/form/div[2]/button[2]";
        private const string OTHER_PACKAGING = "//span[contains(@title,'{0}')]/../../div[16]/div/ul/li[*]/a/span[contains(@class, 'shopping-cart')]";
        private const string PACKAGING_NAME = "//*[@id=\"test\"]/div[2]/div/form/div[2]/div[{0}]/div[2]";
        private const string OTHER_PACKAGING_CHECK_BOX = "Packagings_{0}__IsSelected";
        private const string SAVE_PACKAGING = "//*[@id=\"test\"]/div[2]/div/form/div[3]/button[2]";
        private const string ITEM_EDIT = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[contains(@title,'{0}')]/../../div[16]/div/ul/li[*]/a/span[@class='glyphicon glyphicon glyphicon-pencil glyph-span']";
        private const string ERROR_MODAL = "//*[@id=\"dataAlertModal\"]/div/div";


        private const string ITEM_PROD_QUANTITY = "item_SodRowDto_Quantity";
        private const string EXTENDED_MENU_ITEM = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[1]/div[16]/div/a";
        private const string EDIT_ITEM_NEW = "//*[@id=\"detail-item-event\"]/span";
        // Filtres
        private const string RESET_FILTER_DEV = "ResetFilter";
        private const string RESET_FILTER_PATCH = "//*[@id=\"formSearchItems\"]/div[1]/a";

        private const string SEARCH_BY_NAME_FILTER = "tbSearchPatternWithAutocomplete";
        private const string FIRST_RESULT_SEARCH = "//*[@id=\"formSearchItems\"]/div[2]/span/div/div/div/strong[text()='{0}']";
        private const string KEYWORD_FILTER = "ItemIndexVMSelectedKeywords_ms";
        private const string UNCHECK_ALL_KEYWORD = "/html/body/div[11]/div/ul/li[2]/a";
        private const string SEARCH_KEYWORD = "/html/body/div[11]/div/div/label/input";
        private const string SHOW_ITEMS_NOT_SUPPLIED = "ShowNotSupplied";
        private const string SORT_BY_FILTER = "ItemIndexVM_SelectedSort";
        private const string SUPPLIERS_FILTER = "ItemIndexVMSelectedSuppliers_ms";
        private const string UNSELECT_ALL_SUPPLIERS = "/html/body/div[12]/div/ul/li[2]/a";
        private const string SEARCH_SUPPLIERS = "/html/body/div[12]/div/div/label/input";
        private const string GROUPS_FILTER = "ItemIndexVMSelectedGroups_ms";
        private const string UNSELECT_ALL_GROUPS = "/html/body/div[13]/div/ul/li[2]/a";
        private const string SEARCH_GROUPS = "/html/body/div[13]/div/div/label/input";
        private const string UNCHECK_ALL_GROUP = "/html/body/div[12]/div/ul/li[2]/a";
        private const string UNCHECK_ALL_SUB_GROUP = "/html/body/div[14]/div/ul/li[2]/a/span[2]";
        private const string SEARCH_GROUP = "/html/body/div[12]/div/div/label/input";
        private const string SEARCH_SUB_GROUP = "ItemIndexVMSelectedSubGroups_ms";
        private const string SEARCH_SUB_GROUP_TEXT = "/html/body/div[13]/div/div/label/input";
        private const string INGREDIENT = "IngredientName";
        private const string FIRST_INGREDIENT_RESULT = "//*[@id=\"ingredient-list\"]/table/tbody/tr[1]/td[2]";
        private const string SECOND_INGREDIENT_RESULT = "//*[@id=\"ingredient-list\"]/table/tbody/tr[2]/td[2]";
        private const string THIRD_INGREDIENT_RESULT = "//*[@id=\"ingredient-list\"]/table/tbody/tr[3]/td[2]";
        private const string SORT_BY = "cbSortBy";
        private const string SORT_BY_PORTIONS = "//*[@id=\"cbSortBy\"]/option[2]";
        private const string ALL_PORTIONS = "/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[*]/div/div/form/div/div[3]/span";
        private const string DOWNLOAD_ICON = "header-print-button";
        private const string DOWNLOAD_FILE = "/html/body/div[15]/div[2]/table/tbody/tr[2]/td[4]/a";
        private const string VALIDATE_MESSAGE = "//*[@id=\"modal-1\"]/div/div/div/form/div[2]/p[1]/b";
        private const string VALIDATE_BUTTON_MESSAGE = "//*[@id=\"btn-popup-validate\"]";
        private const string SELECT_FIRST_ITEM_SUPPLYORDER = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[2]";
        private const string ALL_TOLAL_VAT = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[2]/div[11]/span";
        private const string NO_ITEM_MESSAGE = "/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/p[text()=\"No item\"]";
        private const string ITEM_FD_QUANTITY = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[2]/div[9]";
        private const string CONFIRMVALIDATE = "/html/body/div[4]/div/div/div/div/div/form/div[2]/div[2]/div[2]/button[2]";

        private const string NUMBER = "//*[@id=\"tb-new-supplyorder-number\"]";
        private const string FIRSTITEM = "/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/form/div[2]";

        private const string FIRST_VALUE_F_D = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[2]/div[9]/span";
        private const string SUPPLY_ORDERS_NUMBER  = "//*[@id=\"div-body\"]/div/div[1]/h1";
        private const string PRICE_VALUE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[2]/div[8]/span"; 


        // _____________________________________ Variables _________________________________________________

        // General
        [FindsBy(How = How.XPath, Using = FIRST_VALUE_F_D)]
        private IWebElement _first_valueFD;

        [FindsBy(How = How.XPath, Using = PRICE_VALUE)]
        private IWebElement _price_value ;

        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;
        [FindsBy(How = How.XPath, Using = STORAGE_VALUE_FIRSTITEM)]
        private IWebElement _storagevaluefirstitem;
        [FindsBy(How = How.XPath, Using = FIRSTITEM)]
        private IWebElement _firstitemclick;
        [FindsBy(How = How.XPath, Using = NUMBER)]
        private IWebElement _number;
        [FindsBy(How = How.XPath, Using = EXTENDED_BUTTON)]
        private IWebElement _extendedBtn;
        [FindsBy(How = How.XPath, Using = CONFIRMVALIDATE)]
        private IWebElement _confirmvalidate;
        [FindsBy(How = How.Id, Using = REFRESH)]
        private IWebElement _refresh;

        [FindsBy(How = How.Id, Using = EXPORT)]
        private IWebElement _export;

        [FindsBy(How = How.XPath, Using = IMPORT)]
        private IWebElement _import;

        [FindsBy(How = How.Id, Using = ERASE_QUANTITY)]
        private IWebElement _eraseQuantity;

        [FindsBy(How = How.Id, Using = CONFIRM_ERASE_QUANTITY)]
        private IWebElement _confirmEraseQuantity;

        [FindsBy(How = How.Id, Using = PRINT)]
        private IWebElement _print;

        [FindsBy(How = How.Id, Using = MODAL_PRINT_BUTTON)]
        private IWebElement _modalPrintButton;

        [FindsBy(How = How.Id, Using = COMPARE)]
        private IWebElement _compare;

        [FindsBy(How = How.Id, Using = ERASE_COMPARE)]
        private IWebElement _eraseCompare;

        [FindsBy(How = How.XPath, Using = VALIDATE_MENU)]
        private IWebElement _validateMenu;

        [FindsBy(How = How.Id, Using = VALIDATE)]
        private IWebElement _validate;

        [FindsBy(How = How.Id, Using = CONFIRM_VALIDATE)]
        private IWebElement _confirmValidate;

        [FindsBy(How = How.XPath, Using = GENERATE_PO)]
        private IWebElement _generatePurchaseOrder;

        [FindsBy(How = How.XPath, Using = SELECT_SUPPLIERS)]
        private IWebElement _selectAllSuppliers;

        [FindsBy(How = How.XPath, Using = DELIVERY_DATE)]
        private IWebElement _deliveryDate;

        [FindsBy(How = How.Id, Using = GENERATE)]
        private IWebElement _generate;


        [FindsBy(How = How.XPath, Using = VALIDATE_BUTTON)]
        private IWebElement _validateButton;

        [FindsBy(How = How.Id, Using = VALIDATE_PO)]
        private IWebElement _validatePO;
        // Onglets
        [FindsBy(How = How.Id, Using = GENERAL_INFORMATIONS)]
        private IWebElement _generalInformation;

        [FindsBy(How = How.Id, Using = RECIPES)]
        private IWebElement _recipes;

        [FindsBy(How = How.Id, Using = BY_SUPPLIERS)]
        private IWebElement _bySuppliers;


        // Tableau page Items
        [FindsBy(How = How.XPath, Using = GROUP)]
        private IWebElement _group;


        [FindsBy(How = How.XPath, Using = ITEM)]
        private IWebElement _firstItem;

        [FindsBy(How = How.XPath, Using = ITEM_NAME)]
        private IWebElement _firstItemName;

        [FindsBy(How = How.XPath, Using = ITEM_NAME_VALIDATED)]
        private IWebElement _firstItemNameValidated;

        [FindsBy(How = How.XPath, Using = ITEM_SUPPLIER)]
        private IWebElement _itemSupplier;

        [FindsBy(How = How.XPath, Using = PACKING)]
        private IWebElement _packing;

        [FindsBy(How = How.XPath, Using = PACKING_VALIDATED)]
        private IWebElement _packingValidated;
        [FindsBy(How = How.XPath, Using = QTEINPUTPO)]
        private IWebElement _qteInputPO;


        [FindsBy(How = How.Id, Using = QUANTITY_INPUT)]
        private IWebElement _quantityInput;
        [FindsBy(How = How.Id, Using = QUANTITY_INPUT_PO)]
        private IWebElement _quantityInputPO;

        [FindsBy(How = How.XPath, Using = ITEM_QUANTITY)]
        private IWebElement _itemQuantity;

        [FindsBy(How = How.XPath, Using = ITEM_QUANTITY_VALIDATED)]
        private IWebElement _itemQuantityValidated;

        [FindsBy(How = How.XPath, Using = STORAGE_VALUE)]
        private IWebElement _storageValue;

        [FindsBy(How = How.XPath, Using = COMMENT_ICON)]
        private IWebElement _commentItem;

        [FindsBy(How = How.Id, Using = COMMENT)]
        private IWebElement _comment;

        [FindsBy(How = How.XPath, Using = SAVE_COMMENT)]
        private IWebElement _saveComment;

        [FindsBy(How = How.XPath, Using = DELETE_ICON)]
        private IWebElement _deleteItem;

        [FindsBy(How = How.XPath, Using = OTHER_PACKAGING)]
        private IWebElement _otherPackaging;

        [FindsBy(How = How.XPath, Using = SAVE_PACKAGING)]
        private IWebElement _savePackaging;

        [FindsBy(How = How.XPath, Using = EXTENDED_MENU)]
        private IWebElement _extendedMenu;

        [FindsBy(How = How.XPath, Using = ITEM_EDIT)]
        private IWebElement _editItem;

        [FindsBy(How = How.XPath, Using = EXTENDED_MENU_ITEM)]
        private IWebElement _extendedMenuItem;

        [FindsBy(How = How.XPath, Using = EDIT_ITEM_NEW)]
        private IWebElement _editItemNew;

        // ________________________________________________ Filtres ___________________________________________________

        [FindsBy(How = How.Id, Using = RESET_FILTER_DEV)]
        private IWebElement _resetFilterDev;

        [FindsBy(How = How.XPath, Using = RESET_FILTER_PATCH)]
        private IWebElement _resetFilterPatch;

        [FindsBy(How = How.Id, Using = SEARCH_BY_NAME_FILTER)]
        private IWebElement _searchByName;

        [FindsBy(How = How.Id, Using = KEYWORD_FILTER)]
        private IWebElement _keywordFilter;

        [FindsBy(How = How.XPath, Using = UNCHECK_ALL_KEYWORD)]
        private IWebElement _uncheckAllKeyword;

        [FindsBy(How = How.XPath, Using = SEARCH_KEYWORD)]
        private IWebElement _searchKeyword;

        [FindsBy(How = How.Id, Using = SHOW_ITEMS_NOT_SUPPLIED)]
        private IWebElement _showItemsNotSupplied;

        [FindsBy(How = How.Id, Using = SORT_BY_FILTER)]
        private IWebElement _sortBy;

        [FindsBy(How = How.Id, Using = SUPPLIERS_FILTER)]
        private IWebElement _suppliers;

        [FindsBy(How = How.XPath, Using = UNSELECT_ALL_SUPPLIERS)]
        private IWebElement _unselectAllSuppliers;

        [FindsBy(How = How.XPath, Using = SEARCH_SUPPLIERS)]
        private IWebElement _searchSuppliers;

        [FindsBy(How = How.Id, Using = GROUPS_FILTER)]
        private IWebElement _groups;

        [FindsBy(How = How.XPath, Using = UNSELECT_ALL_GROUPS)]
        private IWebElement _unselectAllGroups;

        [FindsBy(How = How.XPath, Using = SEARCH_GROUPS)]
        private IWebElement _searchGroups;

        [FindsBy(How = How.XPath, Using = UNCHECK_ALL_GROUP)]
        private IWebElement _uncheckAllGroup;

        [FindsBy(How = How.XPath, Using = SEARCH_GROUP)]
        private IWebElement _searchGroup;

        [FindsBy(How = How.Id, Using = SEARCH_SUB_GROUP)]
        private IWebElement _filterbysubgrp;


        [FindsBy(How = How.Id, Using = VALIDATE_MESSAGE)]
        private IWebElement _validateMessage;
        [FindsBy(How = How.XPath, Using = VALIDATE_BUTTON_MESSAGE)]
        private IWebElement _validate_Button_Message;

        [FindsBy(How = How.XPath, Using = SELECT_FIRST_ITEM_SUPPLYORDER)]
        private IWebElement _select_first_item_supplyorder;

        
        public enum FilterSupplyItemType
        {
            ByName,
            ByKeyword,
            ShowItemsNotSupplied,
            SortBy,
            Suppliers,
            Groups,
            BySubGroup

        }

        public void Filter(FilterSupplyItemType FilterSupplyItemType, object value)
        {
            Actions action = new Actions(_webDriver);

            switch (FilterSupplyItemType)
            {
                case FilterSupplyItemType.ByName:
                    _searchByName = WaitForElementIsVisible(By.Id(SEARCH_BY_NAME_FILTER));
                    _searchByName.SetValue(ControlType.TextBox, value);

                    try
                    {
                        var firstResultSearch = _webDriver.FindElement(By.XPath(String.Format(FIRST_RESULT_SEARCH, value)));
                        firstResultSearch.Click();
                    }
                    catch
                    {
                        // item non trouvé
                    }

                    break;

                case FilterSupplyItemType.ByKeyword:
                    _keywordFilter = WaitForElementIsVisible(By.Id(KEYWORD_FILTER));
                    _keywordFilter.Click();

                    _uncheckAllKeyword = WaitForElementIsVisible(By.XPath(UNCHECK_ALL_KEYWORD));
                    _uncheckAllKeyword.Click();

                    _searchKeyword = WaitForElementIsVisible(By.XPath(SEARCH_KEYWORD));
                    _searchKeyword.SetValue(ControlType.TextBox, value);

                    var keywordToCheck = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    keywordToCheck.SetValue(ControlType.CheckBox, true);

                    _keywordFilter.Click();
                    break;

                case FilterSupplyItemType.ShowItemsNotSupplied:
                    _showItemsNotSupplied = WaitForElementExists(By.Id(SHOW_ITEMS_NOT_SUPPLIED));
                    _showItemsNotSupplied.SetValue(ControlType.CheckBox, value);
                    break;

                case FilterSupplyItemType.SortBy:
                    var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
                    javaScriptExecutor.ExecuteScript("window.scrollBy(0, -350)", "");
                    _sortBy = WaitForElementExists(By.Id(SORT_BY_FILTER));
                    _sortBy.Click();

                    var element = WaitForElementIsVisible(By.XPath("//option[contains(@value,'" + value + "')]"));
                    _sortBy.SetValue(ControlType.DropDownList, element.Text);
                    _sortBy.Click();
                    break;

                case FilterSupplyItemType.Suppliers:
                    _suppliers = WaitForElementIsVisible(By.Id(SUPPLIERS_FILTER));
                    _suppliers.Click();

                    _unselectAllSuppliers = WaitForElementIsVisible(By.XPath(UNSELECT_ALL_SUPPLIERS));
                    _unselectAllSuppliers.Click();

                    _searchSuppliers = WaitForElementIsVisible(By.XPath(SEARCH_SUPPLIERS));
                    _searchSuppliers.SetValue(ControlType.TextBox, value);


                    var valueToCheck = WaitForElementIsVisible(By.XPath("//label//span[text()='" + value + "']"));
                    valueToCheck.Click();

                    _suppliers.Click();
                    break;

                case FilterSupplyItemType.Groups:
                    _groups = WaitForElementIsVisible(By.Id(GROUPS_FILTER));
                    _groups.Click();

                    _unselectAllGroups = WaitForElementIsVisible(By.XPath(UNSELECT_ALL_GROUPS));
                    _unselectAllGroups.Click();
                    if (value is List<string>)
                    {
                        List<string> values = (List<string>)value;
                        foreach (string val in values)
                        {
                            _searchGroups = WaitForElementIsVisible(By.XPath(SEARCH_GROUPS));
                            _searchGroups.SetValue(ControlType.TextBox, val);

                            var groupToCheck = WaitForElementIsVisible(By.XPath("//label//span[text()='" + val + "']"));
                            groupToCheck.SetValue(ControlType.CheckBox, true);
                        }
                    }
                    else
                    {
                        _searchGroups = WaitForElementIsVisible(By.XPath(SEARCH_GROUPS));
                        _searchGroups.SetValue(ControlType.TextBox, value);

                        var groupToCheck = WaitForElementIsVisible(By.XPath("//label//span[text()='" + value + "']"));
                        groupToCheck.SetValue(ControlType.CheckBox, true);

                    }

                    _groups.Click();
                    break;
                case FilterSupplyItemType.BySubGroup:
                    FilterUncheckAllSubGroup();
                    _filterbysubgrp = WaitForElementIsVisible(By.XPath(SEARCH_SUB_GROUP_TEXT));
                    _filterbysubgrp.SetValue(ControlType.TextBox, value);
                    Thread.Sleep(1000);
                    var subgrptocheck = WaitForElementIsVisible(By.XPath("//label//span[text()='" + value + "']"));
                    subgrptocheck.Click();
                    break;
            }

            WaitPageLoading();
            Thread.Sleep(1500);
        }
        public void FilterUncheckAllSubGroup()
        {
            _searchGroup = WaitForElementIsVisible(By.Id(SEARCH_SUB_GROUP));
            _searchGroup.Click();

            _uncheckAllGroup = WaitForElementIsVisible(By.XPath(UNCHECK_ALL_SUB_GROUP));
            _uncheckAllGroup.Click();
        }
        public void ResetFilter()
        {
            if (IsDev()) {
                try
                {
                    _resetFilterDev = WaitForElementIsVisible(By.Id(RESET_FILTER_DEV));
                    _resetFilterDev.Click();
                }
                catch
                {
                    _resetFilterPatch = WaitForElementIsVisible(By.XPath(RESET_FILTER_PATCH));
                    _resetFilterPatch.Click();
                }
            }
            else
            {
                _resetFilterPatch = WaitForElementIsVisible(By.XPath(RESET_FILTER_PATCH));
                _resetFilterPatch.Click();
            }

            WaitForLoad();

            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                // pas de date
            }
        }

        public bool VerifyName(string value)
        {
            var elements = _webDriver.FindElements(By.XPath(ITEM_NAME));

            if (elements.Count == 0)
                return false;

            foreach (var elm in elements)
            {
                if (!elm.Text.Contains(value))
                {
                    return false;
                }
            }
            return true;
        }

        public bool VerifySupplied()
        {
            var elements = _webDriver.FindElements(By.XPath(ITEM_QUANTITY));

            if (elements.Count == 0)
                return false;

            int compteur = 1;

            while (compteur <= elements.Count)
            {
                foreach (var elm in elements)
                {
                    if (elm.Text.Equals("0"))
                    {
                        return false;
                    }

                    compteur++;
                }
            }
            return true;
        }

        public bool IsSortedByName()
        {
            var ancientName = "";
            int compteur = 1;

            var elements = _webDriver.FindElements(By.XPath(ITEM_NAME));

            if (elements.Count == 0)
                return false;

            foreach (var elm in elements)
            {
                if (compteur == 1)
                    ancientName = elm.GetAttribute("InnerText");

                if (string.Compare(ancientName, elm.GetAttribute("InnerText")) > 0)
                    return false;

                ancientName = elm.GetAttribute("InnerText");
                compteur++;
            }

            return true;
        }

        public bool IsSortedBySupplier()
        {
            var ancientSupplier = "";
            int compteur = 1;

            var elements = _webDriver.FindElements(By.XPath(ITEM_SUPPLIER));

            if (elements.Count == 0)
                return false;

            foreach (var elm in elements)
            {
                if (compteur == 1)
                    ancientSupplier = elm.GetAttribute("title");

                if (string.Compare(ancientSupplier, elm.GetAttribute("title")) > 0)
                    return false;

                ancientSupplier = elm.GetAttribute("title");
                compteur++;
            }

            return true;
        }

        public bool IsSortedByGroup()
        {
            var ancientGroup = "";
            int compteur = 1;

            var elements = _webDriver.FindElements(By.XPath(GROUP));

            if (elements.Count == 0)
                return false;

            foreach (var elm in elements)
            {
                if (compteur == 1)
                    ancientGroup = elm.Text;

                if (string.Compare(ancientGroup, elm.Text) > 0)
                    return false;

                ancientGroup = elm.Text;
                compteur++;
            }

            return true;
        }

        public bool IsSortedByPortions()
        {
            var ancientPortion = "";
            int compteur = 1;
            var noEmptyElements = new List<string>();
            var elements = _webDriver.FindElements(By.XPath(ALL_PORTIONS));
            foreach(var elm in elements)
            {
                if(!elm.Text.Equals(""))
                {
                    noEmptyElements.Add(elm.Text);
                }
            }

            if (noEmptyElements.Count == 0)
                return false;

            foreach (var elm in noEmptyElements)
            {
                if (compteur == 1)
                    ancientPortion = elm;

                if (string.Compare(ancientPortion, elm) > 0)
                    return false;

                ancientPortion = elm;
                compteur++;
            }

            return true;
        }
        public Boolean VerifySupplier(string supplier)
        {
            var elements = _webDriver.FindElements(By.XPath(ITEM_SUPPLIER));

            if (elements.Count == 0)
                return false;

            foreach (var elm in elements)
            {
                if (!elm.Text.Equals(supplier))
                {
                    return false;
                }
            }

            return true;
        }

        public Boolean VerifyGroup(string group)
        {
            var boolValue = true;
            var elements = _webDriver.FindElements(By.XPath(ITEM_NAME));

            if (elements.Count == 0)
                return false;

            foreach (var elm in elements)
            {
                elm.Click();

                var itemGeneralInformationPage = EditItem(elm.GetAttribute("title"));
                var groupName = itemGeneralInformationPage.GetGroupName();
                itemGeneralInformationPage.Close();

                // On ferme l'onglet ouvert
                if (!groupName.Equals(group))
                {
                    return false;
                }
            }
            return boolValue;
        }

        public Boolean VerifyKeyword(string keyword)
        {
            var boolValue = true;
            var elements = _webDriver.FindElements(By.XPath(ITEM_NAME));

            if (elements.Count == 0)
            {
                return false;
            }

            foreach (var elm in elements)
            {
                // clic sur la ligne
                //elm.Click();

                var itemGeneralInformationPage = EditItem(elm.GetAttribute("title").Trim());
                var itemKeywordTab = itemGeneralInformationPage.ClickOnKeywordItem();

                var itemKeyword = itemKeywordTab.IsKeywordPresent(keyword);

                itemGeneralInformationPage.Close();

                if (!itemKeyword)
                    return false;
            }

            return boolValue;
        }

        // ____________________________________________ Méthodes ______________________________________________________



        // General
        public SupplyOrderPage BackToList()
        {
            WaitPageLoading();
            _backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            _backToList.Click();
            WaitForLoad();

            return new SupplyOrderPage(_webDriver, _testContext);
        }
        public void FIRSTITEMCLICK()
        {
            RefreshPage();
            _firstitemclick = WaitForElementIsVisible(By.XPath(FIRSTITEM));
            _firstitemclick.Click();
            WaitForLoad();

         }

        public override void ShowExtendedMenu()
        {
            var actions = new Actions(_webDriver);
            _extendedBtn = WaitForElementIsVisible(By.XPath(EXTENDED_BUTTON));
            actions.MoveToElement(_extendedBtn).Perform();
            WaitForLoad();
        }

        public void Refresh()
        {
            ShowExtendedMenu();

            _refresh = WaitForElementIsVisible(By.Id(REFRESH));
            _refresh.Click();
            WaitForLoad();
        }

        public void ExportItems(bool versionPrint)
        {
            ShowExtendedMenu();
            _export = WaitForElementIsVisible(By.Id(EXPORT));
            _export.Click();

            if (versionPrint)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));
                ClosePrintButton();
            }

            WaitForDownload();
            Close();
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
            // "SupplyForm MAD 9370 2019-10-16 14-36-15.xlsx"
            string site = "(?:[A-Z]{3})";
            string number = "(\\d+)";
            string mois = "(?:0[1-9]|1[0-2])";         // mois
            string space = "(\\s)";                    // Espace
            string annee = "\\d{4}";                   // annee YYYY
            string jour = "[0-3]\\d";                  // jour
            string heure = "(?:0[0-9]|1[0-9]|2[0-3])"; // heure
            string minutes = "[0-5]\\d";               // minutes
            string secondes = "[0-5]\\d";

            Regex r = new Regex("^SupplyForm" + space + site + space + number + space + annee + "-" + mois + "-" + jour + space + heure + "-" + minutes + "-" + secondes + ".xlsx$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = r.Match(filePath);

            return m.Success;
        }

        public SupplyOrderImport Import()
        {
            ShowExtendedMenu();

            _import = WaitForElementIsVisible(By.XPath(IMPORT));
            _import.Click();
            WaitForLoad();

            return new SupplyOrderImport(_webDriver, _testContext);
        }

        public void EraseQuantity()
        {
            ShowExtendedMenu();

            _eraseQuantity = WaitForElementIsVisible(By.Id(ERASE_QUANTITY));
            _eraseQuantity.Click();
            WaitForLoad();

            _confirmEraseQuantity = WaitForElementIsVisible(By.Id(CONFIRM_ERASE_QUANTITY));
            _confirmEraseQuantity.Click();
            WaitForLoad();
        }

        public PrintReportPage Print(bool versionPrint)
        {
            ShowExtendedMenu();
            _print = WaitForElementIsVisible(By.Id(PRINT));
            _print.Click();
            WaitForLoad();

            if (isElementVisible(By.XPath("//h4[contains(text(), 'Print')]"))) //new modal for include prices on report
            {
                _modalPrintButton = WaitForElementIsVisible(By.Id(MODAL_PRINT_BUTTON));
                _modalPrintButton.Click();
                WaitForLoad();
            }

            if (versionPrint)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-pdf']"));
                ClickPrintButton();
            }

            //Results are opened in a new tab, switch the driver to the newly created one
            WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(30));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new PrintReportPage(_webDriver, _testContext);
        }

        public void Compare()
        {
            ShowExtendedMenu();

            _compare = WaitForElementIsVisible(By.Id(COMPARE));
            _compare.Click();
            WaitPageLoading();
        }

        public void EraseCompare()
        {
            ShowValidationMenu();

            ShowExtendedMenu();

            _eraseCompare = WaitForElementIsVisible(By.Id(ERASE_COMPARE));
            _eraseCompare.Click();
            WaitForLoad();
            WaitPageLoading();
        }

        public void Validate()
        {
            var actions = new Actions(_webDriver);
            actions.MoveToElement(_validateMenu).Perform();
            WaitForLoad();
            _validate = WaitForElementIsVisible(By.Id(VALIDATE));
            _validate.Click();
            WaitForLoad();

            _confirmValidate = WaitForElementIsVisible(By.Id(CONFIRM_VALIDATE));
            _confirmValidate.Click();
            WaitPageLoading();
            WaitForLoad();
        }
        public void VerifypopupValidate()
        {
            ShowValidationMenu();

            _validate = WaitForElementIsVisible(By.Id(VALIDATE));
            _validate.Click();
            WaitForLoad();
        }
        public void ConfirmValidate()
        {
            ShowValidationMenu();

            _validate = WaitForElementIsVisible(By.XPath(CONFIRMVALIDATE));
            _validate.Click();
            WaitForLoad();
        }

        public string GetMessageTextValidate()
        {
            try
            {
                _validateMessage = WaitForElementIsVisible(By.XPath(VALIDATE_MESSAGE));
                return _validateMessage.Text;
            }
            catch
            {
                return "";
            }
        }
        public string GetNumber()
        {
          
                _number = WaitForElementIsVisible(By.XPath(NUMBER));
            return _number.GetAttribute("value");



        }
        public string GetSupplyOrderFormNumber()
        {
            var number = WaitForElementExists(By.XPath(SUPPLY_ORDERS_NUMBER));
            string text = number.Text.Trim();
            var match = Regex.Match(text, @"\b\d{6}\b");
            return match.Value;
        }
        public string GetMessageButtonValidate()
        {
            try
            {
                _validate_Button_Message = WaitForElementIsVisible(By.XPath(VALIDATE_BUTTON_MESSAGE));
                return _validate_Button_Message.Text;
            }
            catch
            {
                return "";
            }
        }

        public PurchaseOrdersPage GeneratePurchaseOrder()
        {
            ShowExtendedMenu();

            _generatePurchaseOrder = WaitForElementIsVisible(By.XPath(GENERATE_PO));
            _generatePurchaseOrder.Click();
            WaitForLoad();

            _selectAllSuppliers = WaitForElementIsVisible(By.Id(SELECT_SUPPLIERS));
            _selectAllSuppliers.Click();
            WaitForLoad();

            _deliveryDate = WaitForElementIsVisible(By.XPath(DELIVERY_DATE));
            _deliveryDate.Click();
            var activeDay = WaitForElementIsVisible(By.XPath(String.Format(ACTIVE_DAY_DEV, DateUtils.Now.Day)));
            activeDay.Click();
            WaitForLoad();

            _generate = WaitForElementToBeClickable(By.Id(GENERATE));
            _generate.Click();
            // barre de progression
            Thread.Sleep(10000);
            WaitForLoad();

            return new PurchaseOrdersPage(_webDriver, _testContext);
        }

        public OutputFormItem GenerateOutputForm()
        {
            ShowExtendedMenu();

            _generatePurchaseOrder = WaitForElementIsVisible(By.XPath(GENERATE_OF));
            _generatePurchaseOrder.Click();
            WaitForLoad();

            var selectFrom = WaitForElementIsVisible(By.Id("drop-down-places-from"));
            SelectElement sFrom = new SelectElement(selectFrom);
            sFrom.Options[1].Click();

            var selectTo = WaitForElementIsVisible(By.Id("drop-down-places-to"));
            SelectElement sTo = new SelectElement(selectTo);
            sTo.Options[2].Click();

            var _create = WaitForElementToBeClickable(By.Id("btn-submit-create-output-form"));
            _create.Click();
            WaitForLoad();
            return new OutputFormItem(_webDriver, _testContext);
        }

        // Onglets
        public SupplyOrderGeneralInformation ClickOnGeneralInformation()
        {
            _generalInformation = WaitForElementIsVisible(By.Id(GENERAL_INFORMATIONS));
            _generalInformation.Click();
            WaitForLoad();

            return new SupplyOrderGeneralInformation(_webDriver, _testContext);
        }

        public SupplyOrderRecipes ClickOnRecipes()
        {
            _recipes = WaitForElementIsVisible(By.Id(RECIPES));
            _recipes.Click();
            WaitForLoad();
            return new SupplyOrderRecipes(_webDriver, _testContext);
        }

        public SupplyOrderBySuppliers ClickOnBySuppliers()
        {
            _bySuppliers = WaitForElementIsVisible(By.Id(BY_SUPPLIERS));
            _bySuppliers.Click();
            WaitForLoad();

            return new SupplyOrderBySuppliers(_webDriver, _testContext);
        }

        // Tableau items
        public void SelectFirstItem()
        {
            var noItemMessage = isElementVisible(By.XPath(NO_ITEM_MESSAGE));
            Assert.IsFalse(noItemMessage, "there is no item showing");
            _firstItem = WaitForElementIsVisible(By.XPath(ITEM));
            _firstItem.Click();
            WaitForLoad();
        }
        public void SelectFirstItemNewTab()
        {
            var noItemMessage = isElementVisible(By.XPath("/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/p[text()=\"No item\"]"));
            Assert.IsFalse(noItemMessage, "there is no item showing");
            _select_first_item_supplyorder = WaitForElementIsVisible(By.XPath(SELECT_FIRST_ITEM_SUPPLYORDER));
            _select_first_item_supplyorder.Click();
            WaitForLoad();
        }

        public string GetFirstItemName()
        {
            var noItemMessage = isElementVisible(By.XPath(NO_ITEM_MESSAGE));
            Assert.IsFalse(noItemMessage, "there is no item showing");
            if (isElementVisible(By.XPath(ITEM_NAME)))
            {
                _firstItemName = WaitForElementIsVisible(By.XPath(ITEM_NAME));
            }
            else
            {
                _firstItemName = WaitForElementIsVisible(By.XPath(ITEM_NAME2));
            }
            string itemName = _firstItemName.Text;
            int parenthese = itemName.IndexOf('(');
            if (parenthese != -1)
            {
                itemName = itemName.Substring(0, parenthese);
                itemName = itemName.TrimEnd(' ');
            }
            return itemName;
        }
        public string GetFirstItemNameValidated()
        {
            var noItemMessage = isElementVisible(By.XPath("/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/p[text()=\"No item\"]"));
            Assert.IsFalse(noItemMessage, "there is no item showing");
            _firstItemNameValidated = WaitForElementIsVisible(By.XPath(ITEM_NAME_VALIDATED));
            return _firstItemNameValidated.Text;
        }

        public List<String> GetItemNames()
        {
            List<String> itemNames = new List<String>();

            var items = _webDriver.FindElements(By.XPath(ITEM_NAME));

            foreach (var item in items)
            {
                itemNames.Add(item.GetAttribute("title"));
            }

            return itemNames;
        }

        public bool IsGroupDisplayActive()
        {
            try
            {
                _group = _webDriver.FindElement(By.XPath(GROUP));
                return true;
            }
            catch
            {
                return false;
            }
        }

        public string GetFirstItemSupplier()
        {
            _itemSupplier = WaitForElementIsVisible(By.XPath(ITEM_SUPPLIER));
            return _itemSupplier.Text;
        }

        public string GetFirstPacking()
        {
            _packing = WaitForElementIsVisible(By.XPath(PACKING));
            return _packing.Text;
        }

        public string GetFirstPackingValueValidated()
        {
            string packingValue;
            _packingValidated = WaitForElementIsVisible(By.XPath(PACKING_VALIDATED));
            //try to get first value in (), if not get value in text
            if (_packingValidated.Text.Contains("(") && _packingValidated.Text.Contains(")"))
            {
                int pFrom = _packingValidated.Text.IndexOf("(") + "(".Length;
                int pTo = _packingValidated.Text.LastIndexOf(")");
                string packingValueWithText = _packingValidated.Text.Substring(pFrom, pTo - pFrom);
                packingValue = Regex.Match(packingValueWithText, @"\d+").Value;
            }
            else
            {
                packingValue = Regex.Match(_packingValidated.Text, @"\d+").Value;
            }
            return packingValue;
        }


        public string GetFirstItemQty()
        {
            WaitPageLoading();
            var noItemMessage = isElementVisible(By.XPath("/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/p[text()=\"No item\"]"));
            Assert.IsFalse(noItemMessage, "there is no item showing");
            _itemQuantity = WaitForElementExists(By.XPath(ITEM_QUANTITY));
            return _itemQuantity.GetAttribute("innerText");
        }
        public string GetFirstItemQtyValidated()
        {
            var noItemMessage = isElementVisible(By.XPath("/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/p[text()=\"No item\"]"));
            Assert.IsFalse(noItemMessage, "there is no item showing");
            _itemQuantityValidated = WaitForElementExists(By.XPath(ITEM_QUANTITY_VALIDATED));
            return _itemQuantityValidated.Text;
        }


        public void AddQuantity(string itemName, string qty)
        {
            _quantityInput = WaitForElementIsVisible(By.Id(string.Format(QUANTITY_INPUT, itemName)));
            _quantityInput.SetValue(ControlType.TextBox, qty);
            WaitPageLoading();

        }
        public void AddQuantityPO(string itemName, string qty)
        {
            _quantityInputPO = WaitForElementIsVisible(By.Id(string.Format(QUANTITY_INPUT_PO, itemName)));
            _quantityInputPO.SetValue(ControlType.TextBox, qty);
            WaitPageLoading();

        }
        public bool IsValidationBlockedByRedBanner()
        {
            var redBanner = WaitForElementIsVisible(By.ClassName("red-banner-class"));
            return redBanner != null && redBanner.Displayed;
        }

        public List<string> GetAllTotalVat()
        {
            WaitPageLoading();

            List<string> itemTotalVat = new List<string>();
            var items = _webDriver.FindElements(By.XPath(ALL_TOLAL_VAT));

            foreach (var item in items)
            {
                var itemTotalVatText = item.Text.Replace("€", "").Trim();
                itemTotalVat.Add(itemTotalVatText);
            }

            return itemTotalVat;
        }
        public double GetTotalVat()
        {
            WaitPageLoading(); 

            var item = _webDriver.FindElement(By.Id("total-price-span"));

            string itemTotalVatText = item.Text.Replace("€", "").Trim(); 

            if (double.TryParse(itemTotalVatText, out double vat))
            {
                return vat; 
            }
            return 0.0; 
        
        }

        public List<string> GetAllStorage()
        {
            WaitPageLoading();

            List<string> itemStorage = new List<string>();
            var items = _webDriver.FindElements(By.XPath(STORAGE_VALUE));

            foreach (var item in items)
            {
                var itemStorageText = item.Text;
                itemStorage.Add(itemStorageText);
            }

            return itemStorage;
        }

        public List<double> GetAllProdQty()
        {
            WaitPageLoading();

            List<double> prodQtyStorage = new List<double>();
            var qtys = _webDriver.FindElements(By.XPath(PROD_QTY_VALUE));

            foreach (var qty in qtys)
            {
                if (double.TryParse(qty.Text, out double qtyValue))
                {
                    prodQtyStorage.Add(qtyValue);
                }
             
            }

            return prodQtyStorage;
        }
        public List<double> GetPriceQuantityMultiplication(List<Tuple<double, double>> priceQuantityPairs)
        {
            List<double> multiplicationResults = new List<double>();

            foreach (var pair in priceQuantityPairs)
            {
                // Multiply price (Item1) by quantity (Item2)
                double result = pair.Item1 * pair.Item2;
                multiplicationResults.Add(result);
            }

            return multiplicationResults;
        }


        public List<double> GetAllPrice()
        {
            WaitPageLoading();

            List<double> pricesStorage = new List<double>();
            var prices = _webDriver.FindElements(By.XPath("//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[2]/div[7]/span"));

            foreach (var price in prices)
            {
              
                string cleanedPrice = price.Text.Replace("€", "").Trim();
                cleanedPrice = cleanedPrice.Replace(",", ".");
                if (double.TryParse(cleanedPrice, NumberStyles.Any, CultureInfo.InvariantCulture, out double priceValue))
                {
                    pricesStorage.Add(priceValue);
                }
            }

            return pricesStorage;
        }


        public string GetFirstStorageValue()
        {
            _storageValue = WaitForElementExists(By.XPath(STORAGE_VALUE_FIRSTITEM));
            return _storageValue.GetAttribute("innerText");
        }

        public ItemGeneralInformationPage EditItem(string itemName)
        {
            var noItemMessage = isElementVisible(By.XPath(NO_ITEM_MESSAGE));
            Assert.IsFalse(noItemMessage, "there is no item showing");
             _webDriver.FindElement(By.XPath(string.Format("//*[contains(@class,'display-zone')]//*[@title='{0}']", itemName))).Click();
            //_firstItem = WaitForElementIsVisible(By.XPath(ITEM));
            //_firstItem.Click();
            WaitForLoad();
            _extendedMenu = WaitForElementIsVisible(By.XPath(string.Format("//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[contains(@title,'{0}')]/../../div[15]/div/a/span", itemName)));

            var actions = new Actions(_webDriver);
            actions.MoveToElement(_extendedMenu).Perform();
            _extendedMenu.Click();

            _editItem = WaitForElementIsVisible(By.XPath(string.Format("//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[contains(@title,'{0}')]/../../div[15]/div/ul/li[*]/a/span[contains(@class,'pencil')]", itemName)));
            _editItem.Click();
            WaitForLoad();

            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            WaitForLoad();

            return new ItemGeneralInformationPage(_webDriver, _testContext);
        }
        public ItemGeneralInformationPage EditItemForProdMan (string itemName)
        {
            var noItemMessage = isElementVisible(By.XPath(NO_ITEM_MESSAGE));
            Assert.IsFalse(noItemMessage, "there is no item showing");
            _firstItem = WaitForElementIsVisible(By.XPath(ITEM));
            _firstItem.Click();
            WaitForLoad();
            _extendedMenu = _webDriver.FindElement(By.XPath(string.Format("//span[@title='{0}']/../../div[15]/div/a/span", itemName)));

            var actions = new Actions(_webDriver);
            actions.MoveToElement(_extendedMenu).Perform();
            _extendedMenu.Click();

            _editItem = _webDriver.FindElement(By.XPath(string.Format("//span[@title='{0}']/../..//span[contains(@class,'pencil')]", itemName))); _editItem.Click();
            WaitForLoad();

            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            WaitForLoad();

            return new ItemGeneralInformationPage(_webDriver, _testContext);
        }

        public ItemGeneralInformationPage EditItemNew(string itemName)
        {
            _extendedMenuItem = WaitForElementIsVisible(By.XPath(EXTENDED_MENU_ITEM));
            _extendedMenuItem.Click();

            _editItemNew = WaitForElementIsVisible(By.XPath(EDIT_ITEM_NEW));
            _editItemNew.Click();
            WaitForLoad();

            //itemGeneralInformationPage.Go_To_New_Navigate(2);

            return new ItemGeneralInformationPage(_webDriver, _testContext);
        }

        public void AddComment(string itemName, string comment)
        {
            _extendedMenu = WaitForElementIsVisible(By.XPath(string.Format("//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[contains(@title,'{0}')]/../../div[15]/div/a/span", itemName)));
            _extendedMenu.Click();

            _commentItem = WaitForElementIsVisible(By.XPath(string.Format("//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[contains(@title,'{0}')]/../../div[15]/div/ul/li[*]/a/span[contains(@class,'message')]", itemName)));
            _commentItem.Click();
            WaitForLoad();

            _comment = WaitForElementIsVisible(By.Id(COMMENT));
            _comment.SetValue(ControlType.TextBox, comment);

            _saveComment = WaitForElementToBeClickable(By.XPath("//*[@id=\"modal-1\"]/div/div[2]/div/form/div[2]/button[2]"));

            _saveComment.Click();
            WaitForLoad();
        }

        public string GetComment(string itemName)
        {
            _extendedMenu = WaitForElementIsVisible(By.XPath(string.Format("//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[contains(@title,'{0}')]/../../div[15]/div/a/span", itemName)));
            _extendedMenu.Click();

            _commentItem = WaitForElementIsVisible(By.XPath(string.Format("//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[contains(@title,'{0}')]/../../div[15]/div/ul/li[*]/a/span[contains(@class,'message')]", itemName)));
            _commentItem.Click();
            WaitForLoad();

            try
            {
                _comment = WaitForElementExists(By.Id(COMMENT));
                return _comment.Text;
            }
            catch
            {
                return "";
            }
        }

        public void DeleteFirstItem(string itemName)
        {
            _extendedMenu = WaitForElementIsVisible(By.XPath(string.Format("//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[contains(@title,'{0}')]/../../div[15]/div/a/span", itemName)));
            _extendedMenu.Click();

            _deleteItem = WaitForElementIsVisible(By.XPath(string.Format("//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[contains(@title,'{0}')]/../../div[15]/div/ul/li[*]/a/span[contains(@class,'trash')]", itemName)));
            _deleteItem.Click();

            WaitForLoad();
        }

        public void Fill_Compare(int nbLignes)
        {
            var elements = _webDriver.FindElements(By.XPath(ITEM_NAME));

            int nbMax = elements.Count <= nbLignes ? elements.Count : nbLignes;
            int compteur = 1;
            int compteur2 = 1;
            foreach (var elm in elements)
            {
                if (compteur <= nbMax)
                {
                    elm.Click();
                    var quantities = _webDriver.FindElements(By.XPath(string.Format(QUANTITY_INPUT_XPATH, elm.GetAttribute("title"))));
                    quantities.Where(e=>e.Displayed).FirstOrDefault().SetValue(ControlType.TextBox, "10");

                    // plus de disquette
                    // Thread.Sleep(2000);
                    WaitForLoad();


                    if (compteur + 1 == nbMax + 1)
                    {
                        elements[compteur + 1].Click();
                    }
                    compteur++;
                }


            }
            foreach (var elm in elements)
            {
                if (compteur2 <= nbMax)
                {

                    if (isElementVisible(By.XPath(string.Format(COMPARE_CHECK_BOX, elm.GetAttribute("title")))))
                    {
                        var compareCheckBox = WaitForElementIsVisible(By.XPath(string.Format(COMPARE_CHECK_BOX, elm.GetAttribute("title"))));
                        compareCheckBox.Click();
                    }
                    else
                    {
                        var compareCheckBox = WaitForElementIsVisible(By.XPath(string.Format("//*[@id='list-item-with-action']/div/div[2]/div[*]/div/div/form/div[contains(@class,'edit-zone')]//span[@class='item-detail-col-name-span' and @title=\"{0}\"]/../../div[14]/input[@class='cbCompare']", elm.GetAttribute("title"))));
                        compareCheckBox.Click();
                    }
                    WaitForLoad();

                    compteur2++;
                }
            }
        }

        public bool IsCompareActivated(string itemName)
        {
            try
            {
                _webDriver.FindElement(By.XPath(string.Format(COMPARE_CHECK_BOX, itemName)));
                return true;
            }
            catch
            {
                return false;
            }
        }

        public bool IsErrorPopUpVisible()
        {
            return isElementVisible(By.XPath(ERROR_MODAL));
        }

        public int GetNbCompareLines()
        {
            return _webDriver.FindElements(By.XPath(ITEM_NAME)).Count;
        }

        public bool SelectOtherPackaging(string itemName, string packagingName)
        {
            _extendedMenu = WaitForElementIsVisible(By.XPath(string.Format("//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[contains(@title,'{0}')]/../../div[15]/div/a/span", itemName)));
            _extendedMenu.Click();

            _otherPackaging = WaitForElementIsVisible(By.XPath(string.Format("//span[contains(@title,'{0}')]/../../div[15]/div/ul/li[*]/a/span[contains(@class, 'shopping-cart')]", itemName)));
            _otherPackaging.Click();
            WaitForLoad();

            return SelectPackaging(packagingName);
        }

        public bool SelectPackaging(string packagingName)
        {
            bool isFound = false;
            var nbPackaging = _webDriver.FindElements(By.XPath(TEST_PACKAGING)).Count;

            for (int i = 1; i < nbPackaging; i++)
            {
                var packaging = WaitForElementIsVisible(By.XPath(string.Format(PACKAGING_NAME, i + 1)));
                if (packaging.Text.StartsWith(packagingName))
                {
                    var selectBox = WaitForElementExists(By.Id(string.Format(OTHER_PACKAGING_CHECK_BOX, i - 1)));
                    selectBox.Click();
                    isFound = true;
                    break;
                }
            }

            _savePackaging = WaitForElementIsVisible(By.XPath(SAVE_PACKAGING));
            _savePackaging.Click();
            WaitPageLoading();

            return isFound;
        }
        public bool VerifySubGroupFiltre(string itemname)
        {
            var firstitemName = GetFirstItemName();
            if (firstitemName == itemname)
            {
                return true;
            }
            return false;
        }
        public void AddReceipe(string receipeName)
        {
            WaitForLoad();
            var ingredient = WaitForElementIsVisible(By.Id(INGREDIENT));
            ingredient.SendKeys(receipeName);
            ClickFirtThreeIngredientsResults();
        }
        public void ClickFirtThreeIngredientsResults()
        {
            var firstResult = WaitForElementIsVisible(By.XPath(FIRST_INGREDIENT_RESULT));
            firstResult.Click();
            WaitForLoad();
            var secondResult = WaitForElementIsVisible(By.XPath(SECOND_INGREDIENT_RESULT));
            secondResult.Click();
            WaitForLoad();
            var thirdResult = WaitForElementIsVisible(By.XPath(THIRD_INGREDIENT_RESULT));
            thirdResult.Click();
            WaitForLoad();
        }

        public void SortByPortions()
        {
            var sortBy = WaitForElementIsVisible(By.Id(SORT_BY));
            sortBy.Click();
            var sortByPortions = WaitForElementIsVisible(By.XPath(SORT_BY_PORTIONS));
            sortByPortions.Click();
        }

        public void ClickDownloadButton()
        {
            var DwnBtn = WaitForElementIsVisible(By.Id(DOWNLOAD_ICON));
            DwnBtn.Click();
            WaitForLoad();
            var DownloadFile = WaitForElementExists(By.XPath(DOWNLOAD_FILE));
            DownloadFile.Click();
        }
        public bool VerifyPdf(string trouve, string itemName)
        {
            // GNOCCI DE CALABAZA
            string [] itemNameSplit = itemName.Split(' ');
            bool[] checks = new bool[itemNameSplit.Length];
            for (int i = 0;i<checks.Length;i++)
            {
                checks[i] = false;
            }
            PdfDocument document = PdfDocument.Open(trouve);
            List<string> mots = new List<string>();

            foreach (var p in document.GetPages())
            {
                foreach (var mot in p.GetWords())
                {
                    mots.Add(mot.Text);
                }
            }
            foreach (var mot in mots)
            {
                for (int j = 0; j < itemNameSplit.Length; j++)
                {
                    if (mot == itemNameSplit[j])
                    {
                        checks[j] = true;
                    }

                }
            }
            foreach (bool check in checks)
            {
                if (!check) return false;
            }
            return true;
        }
        public bool VerifyDetailChangeLine(string itemName)
        {
            var qte = "15";
            var i = 2;
            _quantityInput = WaitForElementIsVisible(By.Id(QUANTITY_INPUT));
            _quantityInput.Click();
            AddQuantity(itemName, qte);
            WaitForLoad();
            _quantityInput.SendKeys(Keys.Enter);

            var autreLigne = WaitForElementIsVisible(By.XPath("//*/div[contains(@class,'selected') and contains(@class,'item-tab-row')]/div/div/form/div[contains(@class,'edit-zone')]/div[9]/input"));

            Assert.AreNotEqual(qte, autreLigne.GetAttribute("value"));
            return true;
        }
        public void Go_To_New_Navigate()
        {
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);
            WaitPageLoading();
        }

        public void SetFirstQuantityRefresh(string itemName, string qty)
        {
            _quantityInput = WaitForElementIsVisible(By.XPath("//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[2]/div[3]/span"));
            _quantityInput.Click();

            _quantityInput = WaitForElementIsVisible(By.XPath("//*[@id=\"item_SodRowDto_Quantity\"]"));
            _quantityInput.Click();
            _quantityInput.SetValue(ControlType.TextBox, qty);

            Thread.Sleep(2000);
            WaitForLoad();
            // pas de disquette
            //WaitForElementIsVisible(By.XPath(String.Format(SAVE_ICON, itemName)));
        }
        public void SetfirstQuantity( string qty)
        {

            _quantityInput = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/form/div[1]/div[9]/input"));
            _quantityInput.Click();
            _quantityInput.SetValue(ControlType.TextBox, qty);

            Thread.Sleep(2000);
            WaitForLoad();
            // pas de disquette
            //WaitForElementIsVisible(By.XPath(String.Format(SAVE_ICON, itemName)));
        }
        public void RefreshPage()
        {
            _webDriver.Navigate().Refresh();

        }
        public string GetFirstItemQtyRefresh()
        {
            WaitPageLoading();
            var noItemMessage = isElementVisible(By.XPath("/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/p[text()=\"No item\"]"));
            Assert.IsFalse(noItemMessage, "there is no item showing");
            _itemQuantity = WaitForElementExists(By.XPath(ITEM_QUANTITY_REFRESH));
            return _itemQuantity.GetAttribute("innerText");
        }

        public void OpenNewTab()
        {
            var actualUrl = _webDriver.Url;
            ((IJavaScriptExecutor)_webDriver).ExecuteScript("window.open();");

            // Switch to the new tab
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            _webDriver.Navigate().GoToUrl(actualUrl);
        }

        public double GetProdQuantitySum()
        {
            WaitForLoad(); 
            var physicalQuantities = _webDriver.FindElements(By.Id(ITEM_PROD_QUANTITY));
            double sum = 0;
            foreach (var physQty in physicalQuantities)
            {
                sum += double.Parse(Regex.Replace(physQty.GetAttribute("value"), @"\s+", "").ToString());
            }
            return sum;
        }
        public string GetWarningValidationMessage()
        {
            _validateMessage = WaitForElementIsVisible(By.XPath(VALIDATE_MESSAGE));
            return _validateMessage.Text;
        }
        public void ClickOnIgnoreAndValidate()
        {
            _validate_Button_Message = WaitForElementIsVisible(By.XPath(VALIDATE_BUTTON_MESSAGE));
            _validate_Button_Message.Click();
        }
        public void EditQuantityRefresh(string itemName, string qty)
        {
            _quantityInput = WaitForElementIsVisible(By.Id(ITEM_PROD_QUANTITY));
            _quantityInput.Click();
            _quantityInput.SetValue(ControlType.TextBox, qty);
        }

        public string GetProdQuantity()
        {
            _firstItem = WaitForElementIsVisible(By.XPath(ITEM));
            _firstItem.Click();
            var prodQuantity = WaitForElementIsVisible(By.Id(ITEM_PROD_QUANTITY));
            return prodQuantity.GetAttribute("value");
        }

        public bool IsFDQuantityNotRounded()
        {
            _firstItem = WaitForElementIsVisible(By.XPath(ITEM));
            _firstItem.Click();
            var itemFDQuantity = WaitForElementIsVisible(By.XPath(ITEM_FD_QUANTITY));
            var result = itemFDQuantity.Text.Trim();

            if (result.Contains(','))
            {
                var decimalPart = result.Split(',')[1];
                return decimalPart.Length > 0;
            }
            return false;
        }
        public int GetQuantityPO()
        {
            _qteInputPO = _webDriver.FindElement(By.XPath(QTEINPUTPO));
            string quantityText = _qteInputPO.Text.Trim();

            // Check if the text is not empty and is a valid integer, otherwise return 0
            return int.TryParse(quantityText, out int quantity) ? quantity : 0;
        }

        public void ValidateApresUpdate()
        {
            var actions = new Actions(_webDriver);
            actions.MoveToElement(_validateMenu).Perform();
            WaitPageLoading();

            _validatePO = WaitForElementIsVisible(By.Id(VALIDATE_PO));
            _validatePO.Click();
            WaitForLoad();

            _confirmValidate = WaitForElementIsVisible(By.Id(CONFIRM_VALIDATE));
            _confirmValidate.Click();
            WaitForLoad();
        }
        public bool CheckValidation(bool expectedCondition)
        {
            WaitForLoad();
            var redBanner = _webDriver.FindElements(By.XPath("//*[contains(@class, 'red-banner-class')]"));
            return expectedCondition ? redBanner.Count == 0 : redBanner.Count > 0;
        }

        public void EditFirstItem(string newValue)
        {
            WaitPageLoading();
            WaitForLoad();

            var element = WaitForElementExists(By.XPath(LINE_ITEM));
            element.Click();
            WaitForLoad();
            WaitPageLoading();
            var inputField = WaitForElementIsVisible(By.XPath(EDIT_ITEM));
            inputField.Clear();
            WaitForLoad();
            WaitPageLoading();

            inputField.SendKeys(newValue);
            WaitPageLoading();
            WaitForLoad();
        }
        public string GetProdQty()
        {WaitForLoad();
            var element = WaitForElementIsVisible(By.XPath(QTY));

            return element.GetAttribute("value") ?? element.Text;
        }

        public void SwitchToNewWindow(string originalWindow)
        {

            IList<string> allWindows = _webDriver.WindowHandles;
            // Switch to the new window
            foreach (string window in allWindows)
            {
                if (window != originalWindow)
                {
                    _webDriver.SwitchTo().Window(window);
                    break;
                }
                else
                {
                    _webDriver.Close();
                }
            }
           
        }
        public string GetFirstItemFD()
        {
            _first_valueFD = WaitForElementExists(By.XPath(FIRST_VALUE_F_D));
            return _first_valueFD.Text;
        }
        public string GetPrice_KG()
        {
            _price_value = WaitForElementIsVisible(By.XPath(PRICE_VALUE)) ;
            string price = _price_value.Text.Trim();
            return price; 
        }
    }
}
