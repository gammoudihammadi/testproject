using DocumentFormat.OpenXml.Bibliography;
using iText.StyledXmlParser.Jsoup.Nodes;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Item;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.Events;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Purchasing
{
    public class ItemGeneralInformationPage : PageBase
    {
        public ItemGeneralInformationPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // ____________________________________________ Constantes ________________________________________________

        // General
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";

        private const string PRINT_BTN = "//*[@id=\"div-body\"]/div/div[1]/div/a";
        private const string TYPE_REPORT_PRINT = "SelectedReportType";
        private const string FROM_DATE_PRINT = "datepicker-print-from";
        private const string TO_DATE_PRINT = "datepicker-print-to";
        private const string PRINT_SITE_FILTER = "PrintSelectedSites_ms";
        private const string UNCHECK_PRINT_SITES = "/html/body/div[12]/div/ul/li[2]/a";
        private const string PRINT_SITES_SEARCH = "/html/body/div[12]/div/div/label/input";
        private const string PRINT_CONFIRMATION = "//*[@id=\"item-print-form\"]/div/div[3]/button[2]";
        private const string CHECK_PRINT_SITES = "/html/body/div[12]/div/ul/li[1]/a/span[2]";

        // Onglets
        private const string GENERAL_INFORMATION_TAB = "hrefTabContentItem";
        private const string ITEM_TAB = "hrefTabContentArticles";
        private const string USE_CASE_TAB = "hrefTabContentWeightRecipe";
        private const string DIETETIC_TAB = "hrefTabContentDietetic";
        private const string INTOLERANCE_TAB = "hrefTabContentIntolerance";
        private const string GENERALINFO_TAB = "hrefTabContentItem";

        private const string LABEL_TAB = "hrefTabContentComposition";
        private const string LAST_RECEIPT_NOTES_TAB = "//li/a[text() = 'Last receipt notes']";
        private const string KEYWORD_TAB = "hrefTabContentKeyword";
        private const string PICTURE_TAB = "//li/a[text() = 'Picture']";
        private const string LAST_ORDERS_NOT_RECEIVED_TAB = "hrefTabContentLastOrdersNotReceived";
        private const string LAST_OUTPUT_FORMS_TAB = "hrefTabContentLastOutputforms";
        private const string STORAGE_TAB = "hrefTabContentStorage";
        // General Informations
        private const string ITEM_NAME = "Name";
        private const string GROUP_NAME = "ddl_group";
        private const string GROUP_VALUE = "//*[@id=\"ddl_group\"]/option[@selected=\"selected\"]";
        private const string VAT_VALUE = "//*[@id=\"TaxTypeId\"]/option[@selected=\"selected\"]";
        private const string WORKSHOP = "WorkshopId";
        private const string WORKSHOP_VALUE = "//*[@id=\"WorkshopId\"]/option[@selected=\"selected\"]";
        private const string KEYWORD = "//*[@id=\"div-keywords\"]/div[1]/div/div/input";
        private const string ADD_KEYWORD = "//*[@id=\"div-keywords\"]/div[1]/div/div/a";
        private const string TAX_TYPE = "TaxTypeId";
        private const string PROD_UNIT = "ItemUnitProductionTypeId";
        private const string PROD_UNIT_VALUE = "//*[@id=\"ItemUnitProductionTypeId\"]/option[@selected=\"selected\"]";
        private const string ITEM_COMMERCIALNAME1 = "CommercialName1";
        private const string ITEM_COMMERCIALNAME2 = "CommercialName2";
        private const string ITEM_REFERENCE = "Reference";
        private const string IS_SANITIZATION_CHECKBOX = "IsSanitization";
        private const string ITEM_WEIGHTINGRAM = "ItemUnitProductionWeight";
        private const string DEACTIVATE_PRICE_BTN = "btn-popup-send-all";
        private const string DEACTIVATE_PRICE_SITES_SELECTION = "SelectedItemSites_ms";
        private const string DEACTIVATE_PRICE_SEARCH_SITES = "/html/body/div[11]/div/div/label/input";
        private const string DEACTIVATE_PRICE_UNSELECT_ALL_SITES = "/html/body/div[11]/div/ul/li[2]/a/span[2]";
        private const string DEACTIVATE_PRICE_SITE_TO_CHECK = "//label/span[text()='{0}']";
        private const string DEACTIVATE_PRICE_SUPPLIERS_SELECTION = "SelectedItemSuppliers_ms";
        private const string DEACTIVATE_PRICE_SEARCH_SUPPLIERS = "/html/body/div[12]/div/div/label/input";
        private const string DEACTIVATE_PRICE_UNSELECT_ALL_SUPPLIERS = "/html/body/div[12]/div/ul/li[2]/a/span[2]";
        private const string DEACTIVATE_PRICE_SUPPLIER_TO_CHECK = "//label/span[text()='{0}']";
        private const string DEACTIVATE_PRICE_PACKAGINGS_SELECTION = "SelectedItemPackagings_ms";
        private const string DEACTIVATE_PRICE_SEARCH_PACKAGINGS = "/html/body/div[13]/div/div/label/input";
        private const string DEACTIVATE_PRICE_UNSELECT_ALL_PACKAGINGS = "/html/body/div[13]/div/ul/li[2]/a/span[2]";
        private const string DEACTIVATE_PRICE_PACKAGING_TO_CHECK = "//label/span[contains(text(), '{0}')]";
        private const string DEACTIVATE_PRICE_VALIDATION_BTN = "btn-popup-deactive-price";

        private const string UNPURCHASABLE_ITEM_BTN = "//*[@id=\"btn-popup-send-all\"][text()='set Unpurchasable  Items']";
        private const string UNPURCHASABLE_SITE_FILTER = "SelectedSites_ms";
        private const string UNPURCHASABLE_SITE_UNCHECK_ALL = "/html/body/div[11]/div/ul/li[2]/a";
        private const string UNPURCHASABLE_SITE_INPUT = "/html/body/div[11]/div/div/label/input";
        private const string UNPURCHASABLE_SUPPLIER_FILTER = "SelectedSuppliers_ms";
        private const string UNPURCHASABLE_SUPPLIER_UNCHECK_ALL = "/html/body/div[12]/div/ul/li[2]/a";
        private const string UNPURCHASABLE_SUPPLIER_INPUT = "/html/body/div[12]/div/div/label/input";
        private const string UNPURCHASABLE_PACKAGING_FILTER = "SelectedPackaging_ms";
        private const string UNPURCHASABLE_PACKAGING_UNCHECK_ALL = "/html/body/div[13]/div/ul/li[2]/a";
        private const string UNPURCHASABLE_PACKAGING_INPUT = "/html/body/div[13]/div/div/label/input";
        private const string CONFIRM_UNPURCHASABLE_ITEM = "btn-set-unpurchasable";
        private const string IS_SANITISATION_CHECKBOX = "//input[@id=\"IsSanitization\"]/..";
        private const string IS_THAWING_CHECKBOX = "//input[@id=\"IsThawing\"]/..";
        private const string IS_HACCP_RECORD_EXCLUDED_CHECKBOX = "//input[@id=\"IsHACCPRecordExcluded\"]/..";
        private const string IS_CSR_CHECKBOX = "//input[@id=\"IsCsr\"]/..";
        private const string IS_VEGETABLE_CHECKBOX = "//input[@id=\"IsVegetable\"]/..";
        private const string IS_CHECKED = "//input[@id=\"{0}\" and @checked = 'checked']";
        private const string ITEM_IS_THAWING_CHECKBOX = "IsThawing";
        private const string ITEM_IS_HACCP_RECORD_EXCLUDED_CHECKBOX = "IsHACCPRecordExcluded";
        private const string ITEM_IS_VEGETABLE_CHECKBOX = "IsVegetable";
        private const string ITEM_IS_CSR_CHECKBOX = "IsCsr";    
        // Packaging
        private const string NEW_PACKAGING = "btn-add-item-detail";

        private const string PACKAGING_LINES = "//*[@id=\"detailsItemContainer\"]/div/table/tbody/tr[*]";

        private const string SEARCH_PACKAGING = "PaginatedDetails_SearchPattern";
        private const string FIRST_PACKAGING = "//*[@id=\"detailsItemContainer\"]/div/table/tbody/tr[1]";
        private const string FIRST_PACKAGING_SUPPLIER = "//*[@id=\"detailsItemContainer\"]/div/table/tbody/tr[1]/td[3]/span";
        private const string ITEM_SUPPLIER_SITE = "//*[@id=\"detailsItemContainer\"]/div/table/tbody/tr[*]/td[1]/span[contains(text(),'{0}')]/../../td[3]/span";
        private const string ITEM_PACKAGING_NAME = "//*[@id=\"detailsItemContainer\"]/div/table/tbody/tr[*]/td[5]/span";
        private const string FIRST_PACKAGING_SUPPLIER_REF = "//*[@id=\"detailsItemContainer\"]/div/table/tbody/tr[1]/td[4]/span";
        private const string FIRST_PACKAGING_INFORMATION = "//*[@id=\"detailsItemContainer\"]/div/table/tbody/tr[1]/td[5]/span";
        private const string FIRST_PACKAGING_STORAGE_QTY = "//*[@id=\"detailsItemContainer\"]/div/table/tbody/tr[1]/td[6]/span";
        private const string FIRST_PACKAGING_STORAGE_UNIT = "//*[@id=\"detailsItemContainer\"]/div/table/tbody/tr[1]/td[7]/span";
        private const string FIRST_PACKAGING_PROD_QTY = "//*[@id=\"detailsItemContainer\"]/div/table/tbody/tr[1]/td[8]/div/span";
        private const string FIRST_PACKAGING_PROD_QTY_CASE = "//*[@id=\"detailsItemContainer\"]/div/table/tbody/tr[1]/td[8]/div";
        private const string FIRST_PACKAGING_UNIT_PRICE = "item-detail-unitprice-1";
        private const string FIRST_PACKAGING_UNIT_PRICE_PATCH = "Item-Detail-UnitPrice-1";
        private const string FIRST_PACKAGING_PACK_PRICE = "item-detail-packingprice-1";
        private const string FIRST_PACKAGING_PACK_PRICE_PATCH = "Item-Detail-PackingPrice-1";
        private const string FIRST_PACKAGING_LIMIT_QTY = "//*[@id=\"detailsItemContainer\"]/div/table/tbody/tr[1]/td[12]/div";
        private const string FIRST_PACKAGING_YIELD = "//*[@id=\"detailsItemContainer\"]/div/table/tbody/tr[1]/td[13]/div";
        private const string FIRST_PACKAGING_SITE = "//*[@id=\"detailsItemContainer\"]/div/table/tbody/tr[1]/td[1]/span";
        private const string PURCHASABLE_PACKAGING_ITEM = "/html/body/div[3]/div/div[2]/div[2]/div/div[2]/div[1]/div[2]/div/table/tbody/tr/td[15]/div/input";
        private const string PURCHASABLE_SITE = "/html/body/div[11]/ul";
        private const string PURCHASABLE_SUPPLIER = "/html/body/div[12]/ul";
        private const string PURCHASABLE_PACKAGING = "/html/body/div[13]/ul";
        private const string SHOW_ACTIVE_INACTIVE_ALL = "//*[@id=\"detailsItemContainer\"]/div/table/thead/tr/th[17]/a";
        private const string CHECHBOX = "//*[@id='detail_IsMainSupplier']";

        private const string PACKAGING_EXTENDED_BTN = "/html/body/div[3]/div/div[2]/div[2]/div/div[2]/div[1]/div[2]/div/table/tbody/tr[{0}]/td[17]/span[2]/a";

        private const string DUPLICATE_PACKAGING = "//*[starts-with(@id,\"duplicate-btn-\")]/span";
        private const string DUPLICATE_SITE_FILTER = "SelectedDestinationSitesIds_ms";
        private const string DUPLICATE_UNCHECK_ALL = "/html/body/div[12]/div/ul/li[2]/a/span[2]";
        private const string DUPLICATE_SEARCH_SITE = "/html/body/div[12]/div/div/label/input";
        private const string VALIDATE_DUPLICATE = "//*[@id=\"modal-1\"]/div/div/div[2]/div/form/div[2]/button[2]";
        private const string DUPLICATED_LINE = "//*[@id=\"detailsItemContainer\"]/div/table/tbody/tr[2]";

        private const string ADD_BAR_CODE = "//*[starts-with(@id,\"bar-code-btn-\")]/span";
        private const string BAR_CODE_INPUT = "//*[@id=\"add-barcode\"]/div[2]/div[1]/input";
        private const string VALIDATE_BAR_CODE = "//*[@id=\"add-barcode\"]/div[2]/div[2]/button";
        private const string BAR_CODE = "//*[@id=\"{0}\"]/div[1]/div";

        private const string COMMENT_ITEM = "//*[starts-with(@id,\"comment-btn-\")]/span";

        private const string DELETE_ITEM = "//*[starts-with(@id,\"delete-btn-\")]/span";
        private const string CONFIRM_DELETE = "dataConfirmOK";
        private const string DELETE_ERROR_MESSAGE = "//*[@id=\"dataAlertModal\"]/div/div/div[2]/p";

        private const string PACKAGING_SITE = "//*[@id=\"detailsItemContainer\"]/div/table/tbody/tr[*]/td[1]";
        private const string PACKAGING_MAIN = "//*[@id=\"detailsItemContainer\"]/div/table/tbody/tr[*]/td[3]/span[text()='{0}']/../../td[2]/div/input";
        private const string PACKAGING_SUPPLIER = "//*[@id=\"detailsItemContainer\"]/div/table/tbody/tr[*]/td[3]/span";
        private const string PACKAGING_PURCHASABLE = "//*[@id=\"detailsItemContainer\"]/div/table/tbody/tr[*]/td[3]/span[text()='{0}']/../../td[1]/span[text()='{1}']/../../td[15]/div/input";

        private const string DISQUETTE_SAVED = "//*[@id=\"detailsItemContainer\"]/div/table/tbody/tr[*]/td[17]/span[1]/span[@class='save-state btn-icon btn-icon-status-save']";
        private const string NAME_EDIT_MENU = "Name";


        private const string GROUPE_COMBO = "ddl_group";
        private const string SUBGROUP_COMBO = "/html/body/div[3]/div/div[2]/div[2]/div/div[1]/div/div/form/div[1]/div[1]/div[1]/div[5]/div/fieldset/select";


        private const string INTOLERANCE = "/html/body/div[3]/div/div[2]/div[2]/div/div/div/form/table/tbody/tr[3]";
        private const string SELECTED_SUBGRPNAME = "//*[@id=\"ddl_subgroup\"]/option[@selected=\"selected\"]";
        private const string USE_CASES_COUNTER = "//*[@id=\"tabContentItemContainer\"]/div[1]/h1/span[1]";
        private const string SELECT_ALL_BTN = "//*[@id=\"tabContentItemContainer\"]/div[1]/div/a[4]";
        private const string MULTIPLE_UPDATE = "//*[@id=\"tabContentItemContainer\"]/div[1]/div/div[2]/a[2]";
        private const string FIELD_TO_UPDATE = "FieldToUpdate";
        private const string FIELD_TO_UPDATE_OPTION = "//*[@id=\"FieldToUpdate\"]/option[text()='{0}']";
        private const string FIELD_TO_UPDATE_VALUE = "UpdatedValue";
        private const string UPDATE = "//*[@id=\"ReplacementForm\"]/div[3]/a[2]";
        private const string CONFIRM_UPDATE = "//*[@id=\"dataConfirmOK\"]";
        private const string CLOSE = "//*[@id=\"modal-1\"]/div/div/div[2]/div[2]/button";

        private const string WORKSHOPS = "//*[@id=\"tabContentItemContainer\"]/div[2]/div/ul/li[*]/div/div/div/form/div[3]/div[2]/div[10]/span[1]";
        private const string YIELDS = "//*[@id=\"tabContentItemContainer\"]/div[2]/div/ul/li[*]/div/div/div/form/div[3]/div[2]/div[9]/span[1]";
        private const string GROSS_QTYS = "//*[@id=\"tabContentItemContainer\"]/div[2]/div/ul/li[*]/div/div/div/form/div[3]/div[2]/div[8]/span[1]";
        private const string NET_QTYS = "//*[@id=\"tabContentItemContainer\"]/div[2]/div/ul/li[*]/div/div/div/form/div[3]/div[2]/div[7]/span[1]";
        private const string NET_WEIGHTS = "//*[@id=\"tabContentItemContainer\"]/div[2]/div/ul/li[*]/div/div/div/form/div[3]/div[2]/div[6]/span[1]";
        private const string SELECTED_WORKSHOP = "SelectedWorkshop";
        private const string WORKSHOP_TO_SELECT = "//*[@id=\"SelectedWorkshop\"]/option[text()='{0}']";
        private const string BEST_PRICE = " //*[@id=\"btn-bestPrice\"]";
        private const string LAST_RECEIPT_NOTES = "//a[@id='hrefTabContentLastOrdersNotReceived' and contains(@class, 'active')]";
        private const string FIRSTEXTENDEDMENU = "/html/body/div[3]/div/div[2]/div[2]/div/div[2]/div[1]/div[2]/div/table/tbody/tr/td[17]/span[2]/a/span";
        private const string SHOWFIRSTPACKAGING = "/html/body/div[3]/div/div[2]/div[2]/div/div[2]/div[1]/div[2]/div/table/tbody/tr/td[17]/span[2]/ul";
        private const string CONFIRM_SET_UNPURCHASABLE_REPORT_FOR_SUPPLIER = "//*[@id=\"DeactivateItemsReportForm\"]/div[2]/button";
        private const string WEIGHTONGENERALINFORMATION = "//*[@id=\"ItemUnitProductionWeight\"]";
        private const string REF_SUPPLIER = "RefSupplier";
        private const string SUPPLIER = "RefSupplier";


        // ____________________________________________ Variables ___________________________________________________________

        //General
        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;
        [FindsBy(How = How.XPath, Using = FIRSTEXTENDEDMENU)]
        private IWebElement _firstextendedmenu;
        [FindsBy(How = How.XPath, Using = SHOWFIRSTPACKAGING)]
        private IWebElement _showfirstpackaging;

        [FindsBy(How = How.XPath, Using = PRINT_BTN)]
        private IWebElement _printBtn;

        [FindsBy(How = How.Id, Using = TYPE_REPORT_PRINT)]
        private IWebElement _typeReportPrint;

        [FindsBy(How = How.Id, Using = FROM_DATE_PRINT)]
        private IWebElement _fromDatePrint;

        [FindsBy(How = How.Id, Using = TO_DATE_PRINT)]
        private IWebElement _toDatePrint;

        [FindsBy(How = How.Id, Using = PRINT_SITE_FILTER)]
        private IWebElement _printSiteFilter;

        [FindsBy(How = How.XPath, Using = UNCHECK_PRINT_SITES)]
        private IWebElement _uncheckPrintSites;

        [FindsBy(How = How.XPath, Using = PRINT_SITES_SEARCH)]
        private IWebElement _printSiteSearch;

        [FindsBy(How = How.XPath, Using = PRINT_CONFIRMATION)]
        private IWebElement _printConfirmation;

        [FindsBy(How = How.XPath, Using = IS_SANITISATION_CHECKBOX)]
        private IWebElement _IsSanitization;

        [FindsBy(How = How.XPath, Using = IS_THAWING_CHECKBOX)]
        private IWebElement _IsThawing;

        [FindsBy(How = How.XPath, Using = IS_HACCP_RECORD_EXCLUDED_CHECKBOX)]
        private IWebElement _IsHACCPRecordExcluded;

        [FindsBy(How = How.XPath, Using = IS_CSR_CHECKBOX)]
        private IWebElement _IsCsr;

        [FindsBy(How = How.XPath, Using = IS_VEGETABLE_CHECKBOX)]
        private IWebElement _IsVegetable;

        [FindsBy(How = How.XPath, Using = ITEM_IS_CSR_CHECKBOX)]
        private IWebElement _IsCSR;
        // Onglets

        [FindsBy(How = How.Id, Using = ITEM_TAB)]
        private IWebElement _itemTabBtn;

        [FindsBy(How = How.Id, Using = USE_CASE_TAB)]
        private IWebElement _useCaseTab;

        [FindsBy(How = How.Id, Using = DIETETIC_TAB)]
        private IWebElement _dieteticTab;

        [FindsBy(How = How.Id, Using = INTOLERANCE_TAB)]
        private IWebElement _intoleranceTab;

        [FindsBy(How = How.Id, Using = GENERALINFO_TAB)]
        private IWebElement _generalinfoTab;

        [FindsBy(How = How.Id, Using = LABEL_TAB)]
        private IWebElement _labelTab;

        [FindsBy(How = How.XPath, Using = LAST_RECEIPT_NOTES_TAB)]
        private IWebElement _lastReceiptNotesTab;

        [FindsBy(How = How.XPath, Using = LAST_ORDERS_NOT_RECEIVED_TAB)]
        private IWebElement _lastOrdersNotReceivedTab;

        [FindsBy(How = How.Id, Using = LAST_OUTPUT_FORMS_TAB)]
        private IWebElement _lastOutputFormsTab;

        [FindsBy(How = How.Id, Using = STORAGE_TAB)]
        private IWebElement _storageTab;

        [FindsBy(How = How.Id, Using = KEYWORD_TAB)]
        private IWebElement _keywordItem;

        [FindsBy(How = How.XPath, Using = PICTURE_TAB)]
        private IWebElement _pictureTab;


        // General Informations
        [FindsBy(How = How.Id, Using = ITEM_NAME)]
        private IWebElement _itemName;

        [FindsBy(How = How.Id, Using = GROUP_NAME)]
        private IWebElement _groupName;

        [FindsBy(How = How.XPath, Using = GROUP_VALUE)]
        private IWebElement _groupValue;


        [FindsBy(How = How.XPath, Using = VAT_VALUE)]
        private IWebElement _vatValue;

        [FindsBy(How = How.Id, Using = WORKSHOP)]
        private IWebElement _workshop;

        [FindsBy(How = How.XPath, Using = KEYWORD)]
        private IWebElement _keyword;

        [FindsBy(How = How.XPath, Using = ADD_KEYWORD)]
        private IWebElement _addKeyword;

        [FindsBy(How = How.Id, Using = TAX_TYPE)]
        private IWebElement _taxType;

        [FindsBy(How = How.Id, Using = PROD_UNIT)]
        private IWebElement _prodUnit; 

        [FindsBy(How = How.Id, Using = ITEM_COMMERCIALNAME2)]
        private IWebElement _itemCommercialName2; 

        [FindsBy(How = How.Id, Using = ITEM_REFERENCE)]
        private IWebElement _itemReference;
        
        [FindsBy(How = How.Id, Using = ITEM_WEIGHTINGRAM)]
        private IWebElement _itemWeightInGram;

        [FindsBy(How = How.Id, Using = ITEM_COMMERCIALNAME1)]
        private IWebElement _itemCommercialName1; 

        [FindsBy(How = How.Id, Using = IS_SANITIZATION_CHECKBOX)]
        private IWebElement _checkboxIsSantization;

        [FindsBy(How = How.Id, Using = DEACTIVATE_PRICE_BTN)]
        private IWebElement _deactivatePriceBtn;

        [FindsBy(How = How.Id, Using = DEACTIVATE_PRICE_SITES_SELECTION)]
        private IWebElement _deactivatePriceSiteSelection;

        [FindsBy(How = How.XPath, Using = DEACTIVATE_PRICE_SEARCH_SITES)]
        private IWebElement _deactivatPriceUncheckAllSites;

        [FindsBy(How = How.XPath, Using = DEACTIVATE_PRICE_UNSELECT_ALL_SITES)]
        private IWebElement _deactivatePriceSearchSite;

        [FindsBy(How = How.XPath, Using = DEACTIVATE_PRICE_SITE_TO_CHECK)]
        private IWebElement _deactivatePriceSiteToCheck;

        [FindsBy(How = How.Id, Using = DEACTIVATE_PRICE_SUPPLIERS_SELECTION)]
        private IWebElement _deactivatePriceSuppliersSelection;

        [FindsBy(How = How.XPath, Using = DEACTIVATE_PRICE_SEARCH_SUPPLIERS)]
        private IWebElement _deactivatPriceUncheckAllSuppliers;

        [FindsBy(How = How.XPath, Using = DEACTIVATE_PRICE_UNSELECT_ALL_SUPPLIERS)]
        private IWebElement _deactivatePriceSearchSupplier;

        [FindsBy(How = How.XPath, Using = DEACTIVATE_PRICE_SUPPLIER_TO_CHECK)]
        private IWebElement _deactivatePriceSupplierToCheck;

        [FindsBy(How = How.Id, Using = DEACTIVATE_PRICE_PACKAGINGS_SELECTION)]
        private IWebElement _deactivatePricePackagingsSelection;

        [FindsBy(How = How.XPath, Using = DEACTIVATE_PRICE_SEARCH_PACKAGINGS)]
        private IWebElement _deactivatPriceUncheckAllPackagings;

        [FindsBy(How = How.XPath, Using = DEACTIVATE_PRICE_UNSELECT_ALL_PACKAGINGS)]
        private IWebElement _deactivatePriceSearchPackaging;

        [FindsBy(How = How.XPath, Using = DEACTIVATE_PRICE_PACKAGING_TO_CHECK)]
        private IWebElement _deactivatePricePackagingToCheck;

        [FindsBy(How = How.Id, Using = DEACTIVATE_PRICE_VALIDATION_BTN)]
        private IWebElement _deactivatePriceValidation;

        [FindsBy(How = How.XPath, Using = UNPURCHASABLE_ITEM_BTN)]
        private IWebElement _unpurchasableItemBtn;

        [FindsBy(How = How.XPath, Using = UNPURCHASABLE_SITE_FILTER)]
        private IWebElement _unpurchasableSiteFilter;

        [FindsBy(How = How.XPath, Using = UNPURCHASABLE_SITE_UNCHECK_ALL)]
        private IWebElement _unpurchasableSiteUncheckAll;

        [FindsBy(How = How.XPath, Using = UNPURCHASABLE_SITE_INPUT)]
        private IWebElement _unpurchasableSiteInput;

        [FindsBy(How = How.XPath, Using = UNPURCHASABLE_SUPPLIER_FILTER)]
        private IWebElement _unpurchasableSupplierFilter;

        [FindsBy(How = How.XPath, Using = UNPURCHASABLE_SUPPLIER_UNCHECK_ALL)]
        private IWebElement _unpurchasableSupplierUncheckAll;

        [FindsBy(How = How.XPath, Using = UNPURCHASABLE_SUPPLIER_INPUT)]
        private IWebElement _unpurchasableSupplierInput;

        [FindsBy(How = How.XPath, Using = UNPURCHASABLE_PACKAGING_FILTER)]
        private IWebElement _unpurchasablePackagingFilter;

        [FindsBy(How = How.XPath, Using = UNPURCHASABLE_PACKAGING_UNCHECK_ALL)]
        private IWebElement _unpurchasablePackagingUncheckAll;

        [FindsBy(How = How.XPath, Using = UNPURCHASABLE_PACKAGING_INPUT)]
        private IWebElement _unpurchasablePackagingInput;

        [FindsBy(How = How.Id, Using = CONFIRM_UNPURCHASABLE_ITEM)]
        private IWebElement _confirmUnpurchasable;

        [FindsBy(How = How.Id, Using = ITEM_IS_THAWING_CHECKBOX)]
        private IWebElement _itemIsThawing;

        // Packaging
        [FindsBy(How = How.Id, Using = SEARCH_PACKAGING)]
        private IWebElement _searchPackaging;

        [FindsBy(How = How.XPath, Using = PACKAGING_EXTENDED_BTN)]
        private IWebElement _packagingExtendedMenu;

        [FindsBy(How = How.Id, Using = NEW_PACKAGING)]
        private IWebElement _newPackaging;

        [FindsBy(How = How.XPath, Using = FIRST_PACKAGING)]
        private IWebElement _firstPackaging;

        [FindsBy(How = How.XPath, Using = FIRST_PACKAGING_SUPPLIER)]
        private IWebElement _firstPackagingSupplier;

        [FindsBy(How = How.XPath, Using = ITEM_PACKAGING_NAME)]
        private IWebElement _PackagingName;

        [FindsBy(How = How.XPath, Using = FIRST_PACKAGING_SUPPLIER_REF)]
        private IWebElement _firstPackagingSupplierRef;

        [FindsBy(How = How.XPath, Using = FIRST_PACKAGING_INFORMATION)]
        private IWebElement _firstPackagingInformation;

        [FindsBy(How = How.XPath, Using = FIRST_PACKAGING_STORAGE_QTY)]
        private IWebElement _firstPackagingStorageQty;

        [FindsBy(How = How.XPath, Using = FIRST_PACKAGING_STORAGE_UNIT)]
        private IWebElement _firstPackagingStorageUnit;

        [FindsBy(How = How.XPath, Using = FIRST_PACKAGING_PROD_QTY)]
        private IWebElement _firstPackagingProdQty;

        [FindsBy(How = How.Id, Using = FIRST_PACKAGING_UNIT_PRICE)]
        private IWebElement _firstPackagingUnitPrice;

        [FindsBy(How = How.Id, Using = FIRST_PACKAGING_PACK_PRICE)]
        private IWebElement _firstPackagingPackPrice;

        [FindsBy(How = How.XPath, Using = FIRST_PACKAGING_LIMIT_QTY)]
        private IWebElement _firstPackagingLimitQty;

        [FindsBy(How = How.XPath, Using = FIRST_PACKAGING_YIELD)]
        private IWebElement _firstPackagingYield;

        [FindsBy(How = How.XPath, Using = FIRST_PACKAGING_SITE)]
        private IWebElement _firstPackagingSite;



        [FindsBy(How = How.XPath, Using = PURCHASABLE_PACKAGING_ITEM)]
        private IWebElement _purchasablePackagingItem;

        [FindsBy(How = How.XPath, Using = SHOW_ACTIVE_INACTIVE_ALL)]
        private IWebElement _showActiveInactiveAll;

        [FindsBy(How = How.XPath, Using = DUPLICATE_PACKAGING)]
        private IWebElement _duplicatePackaging;

        [FindsBy(How = How.Id, Using = DUPLICATE_SITE_FILTER)]
        private IWebElement _duplicatePackagingSite;

        [FindsBy(How = How.XPath, Using = DUPLICATE_UNCHECK_ALL)]
        private IWebElement _duplicateUncheckAll;

        [FindsBy(How = How.XPath, Using = DUPLICATE_SEARCH_SITE)]
        private IWebElement _duplicateSearchSite;

        [FindsBy(How = How.XPath, Using = VALIDATE_DUPLICATE)]
        private IWebElement _validateDuplicate;

        [FindsBy(How = How.XPath, Using = ADD_BAR_CODE)]
        private IWebElement _addBarCode;

        [FindsBy(How = How.XPath, Using = BAR_CODE_INPUT)]
        private IWebElement _barCodeInput;

        [FindsBy(How = How.XPath, Using = VALIDATE_BAR_CODE)]
        private IWebElement _validateBarCode;

        [FindsBy(How = How.XPath, Using = COMMENT_ITEM)]
        private IWebElement _commentItem;

        [FindsBy(How = How.XPath, Using = DELETE_ITEM)]
        private IWebElement _deleteItem;

        [FindsBy(How = How.XPath, Using = CONFIRM_DELETE)]
        private IWebElement _confirmDelete;

        [FindsBy(How = How.XPath, Using = DELETE_ERROR_MESSAGE)]
        private IWebElement _deleteErrorMessage;

        [FindsBy(How = How.XPath, Using = BEST_PRICE)]
        private IWebElement _bestPrice;
        [FindsBy(How = How.XPath, Using = LAST_RECEIPT_NOTES)]
        private IWebElement _last_Receipt_notes;
        [FindsBy(How = How.XPath, Using = CONFIRM_SET_UNPURCHASABLE_REPORT_FOR_SUPPLIER)]
        private IWebElement _confirm_set_unpurchasable_report_for_supplier;

        public enum isValueChecked
        {
            IsSanitization,
            IsThawing,
            IsHACCPRecordExcluded,
            IsCsr,
            IsVegetable
        }

        public void Check_Is(isValueChecked isValue, bool value)
        {

            switch (isValue)
            {
                case isValueChecked.IsSanitization:
                    _IsSanitization = WaitForElementIsVisible(By.XPath(IS_SANITISATION_CHECKBOX));
                    IWebElement inputElement = _IsSanitization.FindElement(By.TagName("input"));
                    inputElement.SetValue(ControlType.CheckBox, value);
                    break;

                case isValueChecked.IsThawing:
                    _IsThawing = WaitForElementIsVisible(By.XPath(IS_THAWING_CHECKBOX));
                    _IsThawing.SetValue(ControlType.CheckBox, value);
                    break;

                case isValueChecked.IsHACCPRecordExcluded:
                    _IsHACCPRecordExcluded = WaitForElementIsVisible(By.XPath(IS_HACCP_RECORD_EXCLUDED_CHECKBOX));
                    _IsHACCPRecordExcluded.SetValue(ControlType.CheckBox, value);
                    break;

                case isValueChecked.IsCsr:
                    _IsCsr = WaitForElementIsVisible(By.XPath(IS_CSR_CHECKBOX));
                    _IsCsr.SetValue(ControlType.CheckBox, value);
                    break;

                case isValueChecked.IsVegetable:
                    _IsVegetable = WaitForElementIsVisible(By.XPath(IS_VEGETABLE_CHECKBOX));
                    _IsVegetable.SetValue(ControlType.CheckBox, value);
                    break;
            }
        }

        // _____________________________________________ Méthodes ___________________________________________________

        // General
        public ItemPage BackToList()
        {
            _backToList = WaitForElementIsVisibleNew(By.XPath(BACK_TO_LIST));
            _backToList.Click();
            WaitForLoad();

            return new ItemPage(_webDriver, _testContext);
        }

        public void showfirstpackaging()
        {
            _firstextendedmenu = WaitForElementIsVisible(By.XPath(FIRSTEXTENDEDMENU));
            _firstextendedmenu.Click();
            WaitForLoad();
            if (isElementVisible(By.XPath(SHOWFIRSTPACKAGING)))
            {
                _showfirstpackaging = WaitForElementIsVisible(By.XPath(SHOWFIRSTPACKAGING));
                _showfirstpackaging.Click();
                WaitForLoad();
            }
        }
        public PrintReportPage Print(bool versionPrint, string site)
        {
            _printBtn = WaitForElementToBeClickable(By.XPath(PRINT_BTN));
            _printBtn.Click();
            WaitForLoad();

            _typeReportPrint = WaitForElementIsVisible(By.Id(TYPE_REPORT_PRINT));
            _typeReportPrint.SetValue(ControlType.DropDownList, "Stock movements of item");

            _fromDatePrint = WaitForElementIsVisible(By.Id(FROM_DATE_PRINT));
            _fromDatePrint.SetValue(ControlType.DateTime, DateUtils.Now);
            _fromDatePrint.SendKeys(Keys.Tab);

            _toDatePrint = WaitForElementIsVisible(By.Id(TO_DATE_PRINT));
            _toDatePrint.SetValue(ControlType.DateTime, DateUtils.Now.AddDays(+20));
            _toDatePrint.SendKeys(Keys.Tab);

            _printSiteFilter = WaitForElementIsVisible(By.Id(PRINT_SITE_FILTER));
            _printSiteFilter.Click();

            _uncheckPrintSites = WaitForElementIsVisible(By.XPath(UNCHECK_PRINT_SITES));
            _uncheckPrintSites.Click();

            _printSiteSearch = WaitForElementIsVisible(By.XPath(PRINT_SITES_SEARCH));
            _printSiteSearch.SetValue(ControlType.TextBox, site);

            var valueToCheck = WaitForElementIsVisible(By.XPath("//span[text() = '" + site + " - " + site + "']"));
            valueToCheck.Click();

            _printSiteFilter.Click();

            _printConfirmation = WaitForElementToBeClickable(By.XPath(PRINT_CONFIRMATION));
            _printConfirmation.Click();
            WaitPageLoading();
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

        // Onglets
        public ItemGeneralInformationPage ClickOnGeneralInformationPage()
        {
            WaitForLoad();
            _generalinfoTab = WaitForElementIsVisible(By.Id(GENERALINFO_TAB));
            _generalinfoTab.Click();
            WaitForLoad();

            return new ItemGeneralInformationPage(_webDriver, _testContext);
        }
        public ItemIntolerancePage ClickOnIntolerancePage()
        {
            _intoleranceTab = WaitForElementIsVisible(By.Id(INTOLERANCE_TAB));
            _intoleranceTab.Click();
            WaitForLoad();

            return new ItemIntolerancePage(_webDriver, _testContext);
        }

        public ItemKeywordPage ClickOnKeywordItem()
        {
            _keywordItem = WaitForElementIsVisible(By.Id(KEYWORD_TAB));
            _keywordItem.Click();
            WaitForLoad();

            return new ItemKeywordPage(_webDriver, _testContext);
        }

        public ItemUseCasePage ClickOnUseCasePage()
        {
            _useCaseTab = WaitForElementIsVisible(By.Id(USE_CASE_TAB));
            _useCaseTab.Click();
            WaitForLoad();

            return new ItemUseCasePage(_webDriver, _testContext);
        }

        public ItemLabelPage ClickOnLabelPage()
        {
            _labelTab = WaitForElementIsVisible(By.Id(LABEL_TAB));
            _labelTab.Click();
            WaitForLoad();

            return new ItemLabelPage(_webDriver, _testContext);
        }

        public ItemLastReceiptNotesPage ClickOnLastReceiptNotesPage()
        {
            _lastReceiptNotesTab = WaitForElementIsVisible(By.XPath(LAST_RECEIPT_NOTES_TAB));
            _lastReceiptNotesTab.Click();
            WaitForLoad();

            return new ItemLastReceiptNotesPage(_webDriver, _testContext);
        }

        public ItemPicturePage ClickOnPicturePage()
        {
            _pictureTab = WaitForElementIsVisible(By.XPath(PICTURE_TAB));
            _pictureTab.Click();
            WaitForLoad();

            return new ItemPicturePage(_webDriver, _testContext);
        }

        public ItemDieteticPage ClickOnDieteticPage()
        {
            _dieteticTab = WaitForElementIsVisible(By.Id(DIETETIC_TAB));
            _dieteticTab.Click();
            WaitForLoad();

            return new ItemDieteticPage(_webDriver, _testContext);
        }

        public ItemLastOrdersNotReceivedPage ClickOnLastOrdersNotReceived()
        {
            _lastOrdersNotReceivedTab = WaitForElementIsVisible(By.Id(LAST_ORDERS_NOT_RECEIVED_TAB));
            _lastOrdersNotReceivedTab.Click();
            WaitForLoad();

            return new ItemLastOrdersNotReceivedPage(_webDriver, _testContext);
        }

        public ItemLastOutputFormsPage ClickOnLastOutputForms()
        {
            _lastOutputFormsTab = WaitForElementIsVisible(By.Id(LAST_OUTPUT_FORMS_TAB));
            _lastOutputFormsTab.Click();
            WaitForLoad();

            return new ItemLastOutputFormsPage(_webDriver, _testContext);
        }


        public ItemStoragePage ClickOnStorage()
        {
            _storageTab = WaitForElementIsVisible(By.Id(STORAGE_TAB));
            _storageTab.Click();
            WaitForLoad();

            return new ItemStoragePage(_webDriver, _testContext);
        }

        // General Informations
        public void SetName(string newName)
        {
            WaitPageLoading();
            _itemName = WaitForElementIsVisible(By.Id(ITEM_NAME));
            WaitPageLoading();
            _itemName.SetValue(ControlType.TextBox, newName);
            WaitPageLoading();

            Thread.Sleep(2000);
        }

        public string GetItemName()
        {
            WaitPageLoading();
            _itemName = WaitForElementIsVisible(By.Id(ITEM_NAME));
            return _itemName.GetAttribute("value");
        }
        public string GetItemCommercialName1()
        {
            WaitPageLoading();
            _itemCommercialName1 = WaitForElementIsVisible(By.Id(ITEM_COMMERCIALNAME1));
            return _itemCommercialName1.GetAttribute("value");
        }
        public string GetItemCommercialName2()
        {
            WaitPageLoading();
            _itemCommercialName2 = WaitForElementIsVisible(By.Id(ITEM_COMMERCIALNAME2));
            return _itemCommercialName2.GetAttribute("value");
        }
        public string GetItemReference()
        {
            WaitPageLoading();
            _itemReference = WaitForElementIsVisible(By.Id(ITEM_REFERENCE));
            return _itemReference.GetAttribute("value");
        }

        public string GetItemWeightInGram()
        {
            WaitPageLoading();
            _itemWeightInGram = WaitForElementIsVisible(By.Id(ITEM_WEIGHTINGRAM));
            return _itemWeightInGram.GetAttribute("value");
        }
        public bool GetCheckboxIsSantization()
        {
            WaitPageLoading();
            _checkboxIsSantization = WaitForElementExists(By.Id(IS_SANITIZATION_CHECKBOX));
            return _checkboxIsSantization.Selected;
        }
        public bool GetCheckboxIsThawing()
        {
            WaitPageLoading();
            _IsThawing = WaitForElementExists(By.Id(ITEM_IS_THAWING_CHECKBOX));
            return _IsThawing.Selected;
        }
        public bool GetCheckboxNoHACCPRecord()
        {
            WaitPageLoading();
            _IsHACCPRecordExcluded = WaitForElementExists(By.Id(ITEM_IS_HACCP_RECORD_EXCLUDED_CHECKBOX));
            return _IsHACCPRecordExcluded.Selected;
        }
        public bool GetCheckboxIsVegetable()
        {
            WaitPageLoading();
            _IsVegetable = WaitForElementExists(By.Id(ITEM_IS_VEGETABLE_CHECKBOX));
            return _IsVegetable.Selected;
        }
        public bool GetCheckboxisCsr()
        {
            WaitPageLoading();
            _IsCSR = WaitForElementExists(By.Id(ITEM_IS_CSR_CHECKBOX));
            return _IsCSR.Selected;
        }
        public void SetGroupName(string name)
        {
            var grpCombo = WaitForElementIsVisible(By.Id(GROUPE_COMBO));
            grpCombo.SetValue(ControlType.DropDownList, name);
            WaitForLoad();
        }
        public void SetSubgroupName(string subGrpName)
        {
            var subGrpCombo = WaitForElementIsVisible(By.Id("ddl_subgroup"));
            subGrpCombo.SetValue(ControlType.DropDownList, subGrpName);
            WaitForLoad();
        }
        public string GetGroupName()
        {
            _groupValue = WaitForElementIsVisible(By.XPath(GROUP_VALUE));
            return _groupValue.Text.Trim();
        }
        public string GetWorkshop()
        {
            _workshop = WaitForElementIsVisible(By.XPath(WORKSHOP_VALUE));
            return _workshop.Text.Trim();
        }
        public string GetVatName()
        {
            _vatValue = WaitForElementIsVisible(By.XPath(VAT_VALUE));
            return _vatValue.Text.Trim();
        }
        public string GetProdUnit()
        {
            _prodUnit = WaitForElementExists(By.XPath(PROD_UNIT_VALUE));
            return _prodUnit.Text.Trim();
        }

        public void UpdateItem(string name, string group, string workshop, string taxType, string prodUnit)
        {
            SetItemName(name);
            SetItemGroup(group);
            SetItemWorkshop(workshop);
            SetItemVAT(taxType);
            SetItemProdUnit(prodUnit);
        }
        public void SetItemName(string name)
        {
            _itemName = WaitForElementIsVisible(By.Id(ITEM_NAME));
            _itemName.SetValue(ControlType.TextBox, name);
            WaitForLoad();
            WaitPageLoading();
        }
        public void SetItemGroup(string group)
        {
            _groupName = WaitForElementIsVisible(By.Id(GROUP_NAME));
            _groupName.SetValue(ControlType.DropDownList, group);
            WaitForLoad();
            WaitPageLoading();
        }
        public void SetItemWorkshop(string workshop)
        {
            _workshop = WaitForElementIsVisible(By.Id(WORKSHOP));
            _workshop.SetValue(ControlType.DropDownList, workshop);
            WaitForLoad();
            WaitPageLoading();
        }
        public void SetItemVAT(string taxType)
        {
            _taxType = WaitForElementIsVisible(By.Id(TAX_TYPE));
            _taxType.SetValue(ControlType.DropDownList, taxType);
            WaitForLoad();
            WaitPageLoading();
        }
        public void SetItemCommercialName1(string commercialName1)
        {
            _itemCommercialName1 = WaitForElementIsVisible(By.Id(ITEM_COMMERCIALNAME1));
            _itemCommercialName1.SetValue(ControlType.TextBox, commercialName1);
            WaitForLoad();
            WaitPageLoading();
        }
        public void SetItemCommercialName2(string commercialName2)
        {
            _itemCommercialName2 = WaitForElementIsVisible(By.Id(ITEM_COMMERCIALNAME2));
            _itemCommercialName2.SetValue(ControlType.TextBox, commercialName2);
            WaitForLoad();
            WaitPageLoading();
        }
        public void SetItemReference(string reference)
        {
            _itemReference = WaitForElementIsVisible(By.Id(ITEM_REFERENCE));
            _itemReference.SetValue(ControlType.TextBox, reference);
            WaitForLoad();
            WaitPageLoading();
        }

        public void SetItemProdUnit(string prodUnit)
        {
            _prodUnit = WaitForElementIsVisible(By.Id(PROD_UNIT));
            _prodUnit.SetValue(ControlType.DropDownList, prodUnit);
            WaitForLoad();
            WaitPageLoading();
        }
        public void SetItemWeightInGram(string WeightInGram)
        {
            _itemWeightInGram = WaitForElementIsVisible(By.Id(ITEM_WEIGHTINGRAM));
            _itemWeightInGram.SetValue(ControlType.TextBox, WeightInGram);
            WaitForLoad();
            WaitPageLoading();
        }
        public void SetCheckoxIsSanitization(bool IsSanitization)
        {
            _checkboxIsSantization = WaitForElementExists(By.Id(IS_SANITIZATION_CHECKBOX));
            _checkboxIsSantization.SetValue(PageBase.ControlType.CheckBox, IsSanitization);
            WaitForLoad();
            WaitPageLoading();
        }
        public void SetCheckoxIsThawing(bool IsThawing)
        {
            _IsThawing = WaitForElementIsVisible(By.XPath(IS_THAWING_CHECKBOX));
            _IsThawing.SetValue(PageBase.ControlType.CheckBox, IsThawing);
            WaitForLoad();
            WaitPageLoading();
        }
        public void SetCheckoxNoHACCPRecord(bool NoHACCPRecord)
        {
            _IsHACCPRecordExcluded = WaitForElementIsVisible(By.XPath(IS_HACCP_RECORD_EXCLUDED_CHECKBOX));
            _IsHACCPRecordExcluded.SetValue(PageBase.ControlType.CheckBox, NoHACCPRecord);
            WaitForLoad();
            WaitPageLoading();
        }
        public void SetCheckoxIsVegetable(bool IsVegetable)
        {
            _IsVegetable = WaitForElementIsVisible(By.XPath(IS_VEGETABLE_CHECKBOX));
            _IsVegetable.SetValue(PageBase.ControlType.CheckBox, IsVegetable);
            WaitForLoad();
            WaitPageLoading();
        }
        public void AddKeyword(string keyword)
        {
            int nbKeyword = _webDriver.FindElements(By.XPath("//span[text()='" + keyword + "']")).Count;

            if (nbKeyword == 0)
            {
                _keyword = WaitForElementExists(By.XPath(KEYWORD));
                _keyword.SetValue(ControlType.TextBox, keyword);

                _addKeyword = WaitForElementIsVisible(By.XPath(ADD_KEYWORD));
                _addKeyword.Click();
                WaitForLoad();
            }
        }

        public bool IsKeywordPresent(string keyword)
        {
            int nbKeyword = _webDriver.FindElements(By.XPath("//span[text()='" + keyword + "']")).Count;

            if (nbKeyword == 0)
                return false;

            return true;
        }

        public void DeactivatePrice(string site, string supplier, string packaging)
        {
            _deactivatePriceBtn = WaitForElementIsVisibleNew(By.Id(DEACTIVATE_PRICE_BTN));
            _deactivatePriceBtn.Click();
            WaitForLoad();

            try
            {
                // Site selection
                ComboBoxSelectById(new ComboBoxOptions(DEACTIVATE_PRICE_SITES_SELECTION, site, false));

                // Supplier selection
                ComboBoxSelectById(new ComboBoxOptions(DEACTIVATE_PRICE_SUPPLIERS_SELECTION, supplier, false));

                // Packaging selection
                ComboBoxSelectById(new ComboBoxOptions(DEACTIVATE_PRICE_PACKAGINGS_SELECTION, packaging, false));
            }
            catch
            {
                //normal comportement patch 2021.1229.1-P12
            }

            _deactivatePriceValidation = WaitForElementIsVisible(By.Id(DEACTIVATE_PRICE_VALIDATION_BTN));
            _deactivatePriceValidation.Click();
            WaitForLoad();
        }


        // Packaging
        public ItemCreateNewPackagingModalPage NewPackaging()
        {

            _newPackaging = WaitForElementIsVisibleNew(By.Id(NEW_PACKAGING));
            ((IJavaScriptExecutor)_webDriver).ExecuteScript("arguments[0].scrollIntoView(true);", _newPackaging);
            _newPackaging.Click();
            WaitPageLoading();
            WaitForLoad();
            return new ItemCreateNewPackagingModalPage(_webDriver, _testContext);
        }

        public void SearchPackaging(string value)
        {
            _searchPackaging = WaitForElementIsVisible(By.Id(SEARCH_PACKAGING));
            _searchPackaging.SetValue(ControlType.TextBox, value);
            WaitPageLoading();
            Thread.Sleep(2000); //lenteurs chargement filtre
            WaitForLoad();
        }
        public void SearchPackagingProdMan(string value)
        {
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            javaScriptExecutor.ExecuteScript("window.scrollBy(0, window.innerHeight);");
            _searchPackaging = WaitForElementIsVisible(By.Id(SEARCH_PACKAGING));
            _searchPackaging.SetValue(ControlType.TextBox, value);
            WaitPageLoading();
            Thread.Sleep(2000); //lenteurs chargement filtre
            WaitForLoad();
        }

        public bool ScanPackaging(string site, string supplier)
        {
            WaitForLoad();
            bool trouve = isElementExists(By.XPath(string.Format("//*/tbody/tr/td[1]/span[text()='{0}']/../../td[3]/span[text()='{1}']", site, supplier)));
            WaitForLoad();
            return trouve;
        }

        public bool IsPackagingVisible()
        {
            if (isElementVisible(By.XPath(FIRST_PACKAGING)))
            {
                _firstPackaging = _webDriver.FindElement(By.XPath(FIRST_PACKAGING));
                return true;
            }
            else
            {
                return false;
            }
        }

        public int CountPackaging()
        {
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            //scroll down to see packaging
            javaScriptExecutor.ExecuteScript("window.scrollBy(0, window.innerHeight);");
            return _webDriver.FindElements(By.XPath(PACKAGING_LINES)).Count;
        }

        public bool VerifyPackagingSite(string siteName)
        {
            var nbPackaging = _webDriver.FindElements(By.XPath(PACKAGING_SITE));

            if (nbPackaging.Count == 0)
                return false;

            foreach (var elm in nbPackaging)
            {
                if (elm.Text.Equals(siteName))
                {
                    return true;
                }
            }
            return false;
        }

        public bool VerifyPackagingSupplier(string supplier)
        {
            var nbPackaging = _webDriver.FindElements(By.XPath(PACKAGING_SUPPLIER));

            if (nbPackaging.Count == 0)
                return false;

            foreach (var elm in nbPackaging)
            {
                if (elm.Text.Equals(supplier))
                {
                    return true;
                }
            }
            return false;
        }

        public string GetPackagingSupplierBySite(string site)
        {
            var supplier = _webDriver.FindElement(By.XPath(String.Format(ITEM_SUPPLIER_SITE, site)));
            return supplier.Text.Trim();
        }

        public bool IsMainChecked(string supplier)
        {
            var nbPackaging = _webDriver.FindElements(By.XPath(string.Format(PACKAGING_MAIN, supplier)));

            if (nbPackaging.Count == 0)
                return false;

            foreach (var elm in nbPackaging)
            {

                if (elm.GetAttribute("checked") == "true")
                {
                    return true;
                }
            }

            return false;
        }

        public bool IsPurchasable(string site, string supplier)
        {
            var nbPackaging = _webDriver.FindElements(By.XPath(string.Format(PACKAGING_PURCHASABLE, supplier, site)));

            if (nbPackaging.Count == 0)
                return false;

            foreach (var elm in nbPackaging)
            {
                if (elm.GetAttribute("checked") == "true")
                {
                    return true;
                }

            }

            return false;
        }

        public string GetFirstPackagingSupplier()
        {
            _firstPackagingSupplier = WaitForElementIsVisible(By.XPath(FIRST_PACKAGING_SUPPLIER));
            WaitForLoad();
            return _firstPackagingSupplier.GetAttribute("innerText");
        }
        public string GetFirstPackagingSite()
        {
            _firstPackagingSite = WaitForElementIsVisible(By.XPath(FIRST_PACKAGING_SITE));
            WaitForLoad();
            return _firstPackagingSite.GetAttribute("innerText");
        }

        public string GetFirstPackagingSupplierRef()
        {
            _firstPackagingSupplierRef = WaitForElementIsVisible(By.XPath(FIRST_PACKAGING_SUPPLIER_REF));
            return _firstPackagingSupplierRef.GetAttribute("innerText");
        }

        public string GetFirstPackagingInformation()
        {
            _firstPackagingInformation = WaitForElementIsVisible(By.XPath(FIRST_PACKAGING_INFORMATION));
            return _firstPackagingInformation.GetAttribute("innerText");
        }

        public decimal GetFirstPackagingStorageQty(string decimalSeparator)
        {
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            _firstPackagingStorageQty = _webDriver.FindElement(By.XPath(FIRST_PACKAGING_STORAGE_QTY));
            decimal val = Convert.ToDecimal(_firstPackagingStorageQty.GetAttribute("innerText"), ci);

            return Math.Round(val, 2);
        }

        public string GetFirstPackagingStorageQtyForProdMan()
        {
           
            _firstPackagingStorageQty = _webDriver.FindElement(By.XPath(FIRST_PACKAGING_STORAGE_QTY));


            return _firstPackagingStorageQty.Text;
        }

        public string GetFirstPackagingStorUnit()
        {
            _firstPackagingStorageUnit = _webDriver.FindElement(By.XPath(FIRST_PACKAGING_STORAGE_UNIT));
            return _firstPackagingStorageUnit.GetAttribute("innerText");
        }

        public string GetFirstPackagingProdQty()
        {
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            javaScriptExecutor.ExecuteScript("window.scrollBy(0, window.innerHeight);");
            _firstPackagingProdQty = _webDriver.FindElement(By.XPath(FIRST_PACKAGING_PROD_QTY));
            return _firstPackagingProdQty.Text;
        }
        public string GetFirstPackagingProdQtyProdMan()
        {
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            javaScriptExecutor.ExecuteScript("window.scrollBy(0, window.innerHeight);");
            _firstPackagingProdQty = _webDriver.FindElement(By.XPath(FIRST_PACKAGING_PROD_QTY_CASE));
            return _firstPackagingProdQty.Text;
        }


        public string GetFirstPackagingUnitPrice()
        {
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            javaScriptExecutor.ExecuteScript("window.scrollBy(0, window.innerHeight);");
            _firstPackagingUnitPrice = WaitForElementIsVisible(By.Id(FIRST_PACKAGING_UNIT_PRICE));
            return _firstPackagingUnitPrice.GetAttribute("innerText");
        }
        public string GetFirstPackagingUnitPriceWithClick()
        {
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            javaScriptExecutor.ExecuteScript("window.scrollBy(0, window.innerHeight);");
            _firstPackagingUnitPrice.Click();
            _firstPackagingUnitPrice = WaitForElementIsVisible(By.Id("PackingPrice"));
          
            return _firstPackagingUnitPrice.GetAttribute("value");
        }

        public string GetFirstPackagingPackPrice()
        {
            _firstPackagingPackPrice = WaitForElementExists(By.Id(FIRST_PACKAGING_PACK_PRICE));

            return _firstPackagingPackPrice.GetAttribute("innerText");
        }

        public string GetFirstPackagingLimitQty()
        {
            _firstPackagingLimitQty = _webDriver.FindElement(By.XPath(FIRST_PACKAGING_LIMIT_QTY));
            return _firstPackagingLimitQty.GetAttribute("innerText");
        }

        public string GetFirstPackagingYield()
        {
            var firstPackagingLimitQty = _webDriver.FindElement(By.XPath(FIRST_PACKAGING_YIELD));
            return firstPackagingLimitQty.GetAttribute("innerText");
        }

        public void SetPurchasable(bool isActivate = true)
        {
            _firstPackaging = WaitForElementToBeClickable(By.XPath(FIRST_PACKAGING));
            _firstPackaging.Click();
            WaitForLoad();

            _purchasablePackagingItem = _webDriver.FindElement(By.XPath(PURCHASABLE_PACKAGING_ITEM));
            _purchasablePackagingItem.SetValue(ControlType.CheckBox, isActivate);

            // save-state btn-icon btn-icon-status-save
            WaitForElementIsVisible(By.XPath("//*[@id='detailsItemContainer']/div/table/tbody/tr/td[17]/span[1]/span[contains(@class,'btn-icon-status-save')]"));
        }
        public void ClickOnItem()
        {
            var item = WaitForElementIsVisible(By.XPath("//*[@id='detailsItemContainer']/div/table/tbody/tr"));
            item.Click();
            WaitForLoad();

        }
        public void UncheckIsPurchasable()
        {
            var checkBox = WaitForElementExists(By.XPath("/html/body/div[3]/div/div[2]/div[2]/div/div[2]/div[1]/div[2]/div/table/tbody/tr/td[15]/div/input"));
            checkBox.Click();
            WaitForLoad();
        }






        public void SetUnpurchasableItem(string site, string supplier, string packaging)
        {
            _unpurchasableItemBtn = WaitForElementIsVisible(By.XPath(UNPURCHASABLE_ITEM_BTN));
            _unpurchasableItemBtn.Click();
            WaitForLoad();

            // Filtre Site
            ComboBoxSelectById(new ComboBoxOptions(UNPURCHASABLE_SITE_FILTER, site, false));

            // Filtre supplier
            ComboBoxSelectById(new ComboBoxOptions(UNPURCHASABLE_SUPPLIER_FILTER, supplier, false));

            // Filtre Packaging
            ComboBoxSelectById(new ComboBoxOptions(UNPURCHASABLE_PACKAGING_FILTER, packaging, false));

            // Validation
            _confirmUnpurchasable = WaitForElementIsVisible(By.Id(CONFIRM_UNPURCHASABLE_ITEM));
            _confirmUnpurchasable.Click();
            WaitForLoad();
        }

        public void ShowAllPackaging()
        {
            _showActiveInactiveAll = WaitForElementExists(By.XPath(SHOW_ACTIVE_INACTIVE_ALL));
            _showActiveInactiveAll.Click();
            WaitForLoad();
        }


        public void DuplicateItem(string site)
        {
            _packagingExtendedMenu = WaitForElementIsVisible(By.XPath(String.Format(PACKAGING_EXTENDED_BTN, 1)));
            _packagingExtendedMenu.Click();
            WaitForLoad();

            _duplicatePackaging = WaitForElementIsVisible(By.XPath(DUPLICATE_PACKAGING));
            _duplicatePackaging.Click();
            WaitForLoad();

            ComboBoxSelectById(new ComboBoxOptions(DUPLICATE_SITE_FILTER, site, false));

            _validateDuplicate = WaitForElementIsVisible(By.XPath("//*/button[@value='Create']"));
            _validateDuplicate.Click();
            WaitForLoad();
        }

        public void AddBarCode(string code)
        {
            _packagingExtendedMenu = WaitForElementIsVisible(By.XPath(String.Format(PACKAGING_EXTENDED_BTN, 1)));
            _packagingExtendedMenu.Click();

            _addBarCode = WaitForElementIsVisible(By.XPath(ADD_BAR_CODE));
            _addBarCode.Click();
            WaitForLoad();

            _barCodeInput = WaitForElementIsVisible(By.XPath(BAR_CODE_INPUT));
            _barCodeInput.SetValue(ControlType.TextBox, code);

            _validateBarCode = WaitForElementIsVisible(By.XPath(VALIDATE_BAR_CODE));
            _validateBarCode.Click();
            WaitForLoad();
        }

        public bool IsCodeBarAdded(string code)
        {
            if (isElementVisible(By.XPath(String.Format(BAR_CODE, code))))
            {
                _webDriver.FindElement(By.XPath(String.Format(BAR_CODE, code)));
                return true;
            }
            else
            {
                return false;
            }
        }

        public ItemsCommentItem ClickOnCommentItem()
        {
            _packagingExtendedMenu = WaitForElementIsVisible(By.XPath(String.Format(PACKAGING_EXTENDED_BTN, 1)));
            _packagingExtendedMenu.Click();

            _commentItem = WaitForElementIsVisible(By.XPath(COMMENT_ITEM));
            _commentItem.Click();
            WaitForLoad();

            return new ItemsCommentItem(_webDriver, _testContext);
        }

        public void ClickOnDeleteItem()
        {
            _packagingExtendedMenu = WaitForElementIsVisibleNew(By.XPath(String.Format(PACKAGING_EXTENDED_BTN, 1)));
            _packagingExtendedMenu.Click();

            _deleteItem = WaitForElementIsVisibleNew(By.XPath(DELETE_ITEM));
            _deleteItem.Click();
            WaitLoading();
        }

        public void ConfirmDelete()
        {
            _confirmDelete = WaitForElementIsVisibleNew(By.Id(CONFIRM_DELETE));
            _confirmDelete.Click();
            WaitLoading();
            _webDriver.Navigate().Refresh();
        }

        public bool ErrorMessageCannotBeDeleted(string supplier, string site)
        {
            WaitLoading();
            _deleteErrorMessage = WaitForElementIsVisible(By.XPath(DELETE_ERROR_MESSAGE));

            string expectedMessage = "This item cannot be deleted for the following reason(s): - this item has a positive stock";
            string expectedMessage2 = string.Format("This item cannot be deleted for the following reason(s): - Item [{0}] - [{1}] has a positive stock.", supplier, site);
            if (_deleteErrorMessage.Text == expectedMessage || _deleteErrorMessage.Text == expectedMessage2)
            {
                return true;
            }

            return false;

        }

        public string GetItemNameFromEditMenu()
        {
            var name = WaitForElementIsVisible(By.Id(NAME_EDIT_MENU));
            return name.GetAttribute("value");
        }

        public string CheckIntoleranceIfNotChecked()
        {
            var allergName = WaitForElementIsVisible(By.XPath(INTOLERANCE + "/td[2]/label"));
            var allergChecked = WaitForElementExists(By.XPath(INTOLERANCE + "/td[1]/div[1]/input"));
            if (!allergChecked.Selected)
            {
                allergChecked.Click();
                return allergName.Text;
            }

            return allergName.Text;
        }
        public string getSubGroupeName()
        {
            var subname = WaitForElementIsVisible(By.XPath(SELECTED_SUBGRPNAME));
            return subname.Text;
        }

        public SupplierItem ClickOnItemsTab()
        {
            _itemTabBtn = WaitForElementIsVisible(By.Id(ITEM_TAB));
            _itemTabBtn.Click();
            WaitForLoad();
            return new SupplierItem(_webDriver, _testContext);
        }

        public void SetActivated(bool active)
        {
            var activated = WaitForElementExists(By.Id("IsActive"));
            activated.SetValue(ControlType.CheckBox, active);
            WaitForLoad();
        }
        public int GetNumberOfUseCases()
        {
            var counter = WaitForElementIsVisible(By.XPath(USE_CASES_COUNTER)).Text;
            var counterInt = int.Parse(counter);
            return counterInt;
        }

        public void SelectAll()
        {
            var selectAll = WaitForElementIsVisible(By.XPath(SELECT_ALL_BTN));
            selectAll.Click();
        }
        public void MultipleUpdate(string field, string value)
        {
            ((IJavaScriptExecutor)_webDriver).ExecuteScript("window.scrollTo(0, 0)");
            var updateAll = WaitForElementIsVisible(By.XPath(MULTIPLE_UPDATE));
            updateAll.Click();
            var fieldToUpdateInput = WaitForElementIsVisible(By.Id(FIELD_TO_UPDATE));
            fieldToUpdateInput.Click();
            var fieltToSelect = WaitForElementIsVisible(By.XPath(string.Format(FIELD_TO_UPDATE_OPTION, field)));
            fieltToSelect.Click();
            if (field.Equals("Workshop"))
            {
                var input = WaitForElementIsVisible(By.Id(SELECTED_WORKSHOP));
                var valueToSelect = WaitForElementIsVisible(By.XPath(string.Format(WORKSHOP_TO_SELECT, value)));
                valueToSelect.Click();
            }
            else
            {
                var valueInput = WaitForElementIsVisible(By.Id(FIELD_TO_UPDATE_VALUE));
                valueInput.Clear();
                valueInput.SendKeys(value);
            }

            var update = WaitForElementIsVisible(By.XPath(UPDATE));
            update.Click();
            var confirm = WaitForElementIsVisible(By.XPath(CONFIRM_UPDATE));
            confirm.Click();
            WaitForLoad();

            var close = WaitForElementIsVisible(By.XPath("//*/div[@class=\"modal-footer\"]/button"));
            close.Click();

        }
        public enum Fields
        {
            Yield,
            NetWeight,
            NetQuantity,
            GrossQuantity,
            Workshop
        }
        public bool VerifyMultipleUpdate(Fields field, string value)
        {
            PageSize("100");
            switch (field)
            {
                case Fields.Yield:
                    var listy = _webDriver.FindElements(By.XPath(YIELDS));
                    foreach (var item in listy)
                    {
                        if (item.Text != value)
                            return false;
                    }
                    break;
                case Fields.NetWeight:
                    var listnw = _webDriver.FindElements(By.XPath(NET_WEIGHTS));
                    foreach (var item in listnw)
                    {
                        if (item.Text != value)
                            return false;
                    }
                    break;
                case Fields.NetQuantity:
                    var listnq = _webDriver.FindElements(By.XPath(NET_QTYS));
                    foreach (var item in listnq)
                    {
                        if (item.Text != value)
                            return false;
                    }
                    break;
                case Fields.GrossQuantity:
                    var listgq = _webDriver.FindElements(By.XPath(GROSS_QTYS));
                    foreach (var item in listgq)
                    {
                        if (item.Text != value)
                            return false;
                    }
                    break;
                case Fields.Workshop:
                    var listw = _webDriver.FindElements(By.XPath(WORKSHOPS));
                    foreach (var item in listw)
                    {
                        if (item.Text != value)
                            return false;
                    }
                    break;
            }
            return true;


        }

        public void SelectFirstPackaging()
        {
            var FirstItem = WaitForElementIsVisible(By.XPath(FIRST_PACKAGING));
            FirstItem.Click();
        }

        public string GetPrice()
        {
            WaitForLoad();
            var FirstItemPrice = WaitForElementIsVisible(By.XPath("//*[@id=\"UnitPrice\"]"));
            return FirstItemPrice.GetAttribute("value");
        }

        public void ModifPrice(string newprice)
        {
            WaitForLoad();
            var FirstItemPrice = WaitForElementIsVisible(By.XPath("//*[@id=\"UnitPrice\"]"));
            FirstItemPrice.SetValue(ControlType.TextBox, newprice);

            WaitForElementIsVisible(By.XPath("//*[@id='detailsItemContainer']/div/table/tbody/tr/td[17]/span[1]/span[contains(@class,'btn-icon-status-save')]"));
        }
        public void SetPackagingPrice(string newprice)
        {
            var FirstItemPrice = WaitForElementIsVisible(By.Id("PackingPrice"));
            FirstItemPrice.SetValue(ControlType.TextBox, newprice);

            WaitForElementIsVisible(By.XPath("//*[@id='detailsItemContainer']/div/table/tbody/tr/td[17]/span[1]/span[contains(@class,'save-state btn-icon')]"));
        }

        public bool isChecked(string inputId)
        {
            if (!isElementVisible(By.XPath(string.Format("//input[@id=\"{0}\" and @checked = 'checked']/..", inputId))))
            {
                return false;
            }
            return true;
        }
        public bool FindPackagingName(string name)
        {
            if (!IsPackagingVisible())
            {
                return false;
            }
            IReadOnlyList<IWebElement> packs = _webDriver.FindElements(By.XPath(ITEM_PACKAGING_NAME));

            _PackagingName = WaitForElementIsVisible(By.XPath(ITEM_PACKAGING_NAME));
            foreach (IWebElement pack in packs)
            {
                if (pack.GetAttribute("innerText") == name) ;
                return true;
            }
            return false;
        }
        /**
         * param ongletChrome : mettre 2 si c'est un troisième onglet
         */
        public void Go_To_New_Navigate(int ongletChrome = 1)
        {
            WaitPageLoading();
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > ongletChrome);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[ongletChrome]);
            WaitPageLoading();
        }

        public void ScrollDown()
        {

            IJavaScriptExecutor js = (IJavaScriptExecutor)_webDriver;
            js.ExecuteScript("window.scrollTo(0, document.body.scrollHeight)");



        }
        public BestPriceModal ClickOnBestPrice()
        {
            _bestPrice = WaitForElementIsVisible(By.XPath(BEST_PRICE));
            _bestPrice.Click();
            WaitForLoad();

            return new BestPriceModal(_webDriver, _testContext);
        }
        public bool IsLastOrdersNotReceivedTabOpened()
        {
            _last_Receipt_notes = _webDriver.FindElement(By.XPath(LAST_RECEIPT_NOTES));
            return _last_Receipt_notes.Displayed;
        }
        public void Close()
        {
            _webDriver.Close();
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles.Last());
        }
        public void ConfirmSetUnpurchasableReportForSupplier()
        {
            _confirm_set_unpurchasable_report_for_supplier = WaitForElementIsVisible(By.XPath(CONFIRM_SET_UNPURCHASABLE_REPORT_FOR_SUPPLIER));
            _confirm_set_unpurchasable_report_for_supplier.Click();
            WaitForLoad();
        }
        public bool IsExpireDateColumnVisible()
        {
            try
            {
                // Localiser l'élément de la colonne "Expire date"
                var expireDateColumn = _webDriver.FindElement(By.XPath("//th[contains(text(),'Expire date')]"));
                // Vérifier si l'élément est visible
                return expireDateColumn.Displayed;
            }
            catch (NoSuchElementException)
            {
                return false; // La colonne n'est pas visible ou n'existe pas
            }
        }
        public bool IsPackageExist(string PackageName)
        {
            if (!isElementVisible(By.XPath(string.Format(PACKAGING_LINES, PackageName))))
            {
                return false;
            }
            return true;
        }
        public string GetFirstPackagingAveragePrice()
        {
            var _firstPackagingAveragePrice = WaitForElementExists(By.XPath("//*[@id=\"detailsItemContainer\"]/div/table/tbody/tr/td[11]"));

            return _firstPackagingAveragePrice.GetAttribute("innerText");
        }
        public bool IsCheckboxChecked()
        {
            IWebElement checkbox = _webDriver.FindElement(By.XPath(CHECHBOX));

            bool isChecked = checkbox.Selected;

            return isChecked;
        }
        public string GetWeightInGram()
        {
            var weightongeneralinformation = WaitForElementIsVisible(By.XPath(WEIGHTONGENERALINFORMATION));
            return weightongeneralinformation.GetAttribute("value");

        }
        public bool VerifySupplierRef(string supplierRef)
        {
            var result = _webDriver.FindElements(By.XPath("/html/body/div[3]/div/div[2]/div[2]/div/div[2]/div[1]/div[2]/div/table/tbody/tr[*]/td[4]/span"));
            foreach (var element in result)
            {
                if (element.Text != supplierRef)
                    return false;
            }
            return true;
        }
        public void  GetUnitPrice()
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)_webDriver;
            js.ExecuteScript("window.scrollTo(0, document.body.scrollHeight)");
            var unitPrice = _webDriver.FindElement(By.Id("UnitPrice"));
         //   return unitPrice.Text.Trim();
        }
        public List<string> GetAllSites()
        {
            var tableSitesPackaging = _webDriver.FindElements(By.XPath("/html/body/div[3]/div/div[2]/div[2]/div/div[2]/div[1]/div[2]/div/table/tbody/tr[*]/td[1]/span"));
            List<string> liste = new List<string>();
            foreach (var tableSite in tableSitesPackaging)
            {
                liste.Add(tableSite.Text);
            }
            return liste;
        }
        public PrintReportPage printHistory(bool Print)
        {
            _printBtn = WaitForElementToBeClickable(By.XPath(PRINT_BTN));
            _printBtn.Click();
            WaitForLoad();

            _typeReportPrint = WaitForElementIsVisible(By.Id(TYPE_REPORT_PRINT));
            _typeReportPrint.SetValue(ControlType.DropDownList, "Price history of item");

            _fromDatePrint = WaitForElementIsVisible(By.Id(FROM_DATE_PRINT));
            _fromDatePrint.SetValue(ControlType.DateTime, DateUtils.Now);
            _fromDatePrint.SendKeys(Keys.Tab);

            _toDatePrint = WaitForElementIsVisible(By.Id(TO_DATE_PRINT));
            _toDatePrint.SetValue(ControlType.DateTime, DateUtils.Now.AddDays(+20));
            _toDatePrint.SendKeys(Keys.Tab);

            _printSiteFilter = WaitForElementIsVisible(By.Id(PRINT_SITE_FILTER));
            _printSiteFilter.Click();

            _uncheckPrintSites = WaitForElementIsVisible(By.XPath(CHECK_PRINT_SITES));
            _uncheckPrintSites.Click();

            _printConfirmation = WaitForElementToBeClickable(By.XPath(PRINT_CONFIRMATION));
            _printConfirmation.Click();
            WaitPageLoading();
            WaitForLoad();

            if (Print)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-pdf']"));
                ClickPrintButton();
            }

            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new PrintReportPage(_webDriver, _testContext);
        }
        public string GetRefSupplier()
        {
            var element = WaitForElementIsVisible(By.Id(SUPPLIER));
            //return element.Text;
            return element.GetAttribute("value");

        }
      
        public void SetSupplierRef(string newRefSupplier)
        {
            var FirstItemPrice = WaitForElementIsVisible(By.Id(REF_SUPPLIER));

            FirstItemPrice.SetValue(ControlType.TextBox, newRefSupplier);
        }
    }
}