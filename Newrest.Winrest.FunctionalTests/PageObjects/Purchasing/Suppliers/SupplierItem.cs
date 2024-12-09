using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Suppliers;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Policy;
using System.Text.RegularExpressions;
using System.Threading;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;
using Keys = OpenQA.Selenium.Keys;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Purchasing
{
    public class SupplierItem : PageBase
    {
        public SupplierItem(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //_________________________________ CONSTANTES ____________________________________________________

        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";

        // General Informations

        private const string GENERAL_INFORMATION_TAB = "hrefTabContentInformations";
        private const string SANITARY_NUMBER = "SanitaryNumber";
        private const string SIRET_FIELD = "IdExternal";
        private const string DATE_END_CONTRACT = "date-picker-supplier-end-contract";
        private const string LAST_AUDIT_DATE = "date-picker-supplier-last-audit";
        private const string NEXT_AUDIT_DATE = "date-picker-supplier-next-audit";
        private const string ACTIVATE_SUPPLIER = "IsActive";
        private const string DEACTIVATE_SUPPLIER = "btn-deactivate";
        private const string HEADER_PRINT_BUTTON = "header-print-button";
        private const string SITE_FC = "siteSource";
        private const string IBAN_FIELD = "//*[@id=\"IBANNumber\"]";
        private const string AUTO_CLOSE = "IsAutoClose";


        // Items

        private const string ITEM_TAB = "hrefTabContentArticles";
        private const string NEW_ITEM = "//*[@id=\"tabContentDetails\"]/div/div[2]/div/div[1]/div[2]/div/a[2]";
        private const string ITEM_NAME = "//*[@id=\"list-item-with-action\"]/div/div[1]/div/div[2]/table/tbody/tr/td[2]";
        private const string ITEM_ELEMENT = "//*[@id=\"list-item-with-action\"]/div/div[1]/div/div[2]/table/tbody/tr/td[*]";
        private const string EXPORT_ITEMS = "//*[@id=\"tabContentDetails\"]//a[contains(text(),'Export')]";
        private const string CONFIRM_EXPORT_ITEMS = "btnExport";
        private const string DEACTIVATE_ITEMS = "deactivate-items-btn";
        private const string UNPURCHASABLE_ITEMS = "unpurchasable-items-btn";
        private const string CHECK_ALL_SITES_ON_ITEMS = "//div[contains(@style, 'display')]//span[text()='Check all']";
        private const string IMPORT_ITEMS = "//*[@id=\"tabContentDetails\"]//a[contains(text(),'Import')]";
        private const string UNFOLD_BTN = "//*[@id=\"tabContentDetails\"]/div/div[2]/div/div[1]/div[1]/a";
        private const string SITE_SEARCH_ITEMS = "SelectedSitesToUpdate_ms";
        private const string FILTER_ITEM_SEARCH_BY_NAME = "tbSearchPatternWithAutocomplete";
        private const string FIRST_RESULT_SEARCH = "//*[@id=\"formSearchItems\"]/div[2]/span/div/div/div/strong[text()='{0}']";
        private const string FIRST_ITEM_NAME = "/html/body/div[3]/div/div/div/div/div[2]/div/div[2]/div/div[1]/div[1]/div/div[2]/table/tbody/tr/td[2]";
        private const string FAST_COPY_PRICE = "/html/body/div[3]/div/div/div/div/div[2]/div/div[1]/div[2]/a[1]";
        private const string SET_SITE = "/html/body/div[4]/div/div/div/div/form/div[2]/div[1]/div[1]/select";
        private const string FILTER_BY_NAME = "/html/body/div[4]/div/div/div/div/form/div[2]/div[1]/div[3]/div[2]/input";
        private const string COPY_BTN = "//*[@id=\"FastCopyPriceForm\"]/div[3]/button[2]";
        private const string OK_BTN = "/html/body/div[4]/div/div/div[3]/button";
        private const string CHECK_FIRST_SITE = "/html/body/div[4]/div/div/div/div/form/div[2]/div[1]/div[4]/table/tbody/tr[3]/td[1]/div/span";
        private const string SELECT_ALL = "/html/body/div[4]/div/div/div/div/form/div[2]/div[1]/div[3]/div[1]/a[1]";
        private const string COPY_PRICES = "//*[@id=\"FastCopyPriceForm\"]/div[3]/button[2]";
        private const string ITEM_EDITED = "//*[@id=\"modal-1\"]/div[2]/div/ul/li";
        private const string COPY_ALL_PRICES_FROM = "siteSource";
        private const string OK_COPY = "//*[@id=\"modal-1\"]/div[3]/button";
        private const string SELECTALL = "select-button";
        private const string COPY = "//*[@id=\"FastCopyPriceForm\"]/div[3]/button[2]";
        private const string OK = "/html/body/div[4]/div/div/div[3]/button";

        // filter itemsTab
        private const string FILTER_SORT_BY = "cbSortBy";

        private const string RESET_FILTER_ITEMS_DEV = "ResetFilter";
        private const string RESET_FILTER_ITEMS_PATCH = "//*[@id=\"formSearchItems\"]/div[1]/a";

        private const string CONTENT = "//*[starts-with(@id,\"content_\")]";
        private const string BTN_OK = "//*[@id=\"DeactivateItemsReportForm\"]/div[2]/button";
        private const string UNFOLD_BTN_ALL = "//*[@id=\"unfoldBtn\"]/span[@class='glyphicon glyphicon-resize-small']";
        private const string UNFOLD_BTN_GLYPHY_VERTICAL = "//*[@id=\"unfoldBtn\"]/span[@class='glyphicon glyphicon-resize-vertical']";
        private const string FIRST_ITEM_CONTACT = "//*[@id=\"list-item-with-action\"]/div/div[1]/div";

        //Deliveries

        private const string DELIVERY_TAB = "hrefTabContentDelivery";
        private const string MONDAY = "//*/label[text()='Mon']";
        private const string TUESDAY = "//*/label[text()='Tue']";
        private const string WEDNESDAY = "//*/label[text()='Wed']";
        private const string THURSDAY = "//*/label[text()='Thu']";
        private const string FRIDAY = "//*/label[text()='Fri']";
        private const string SATURDAY = "//*/label[text()='Sat']";
        private const string SUNDAY = "//*/label[text()='Sun']";
        private const string SHOW_COMMENT_PANEL = "//*/img[contains(@src,'comment')]";
        private const string COMMENT_PANEL = "DeliveriesSites_0__Comment";

        // Contacts

        private const string CONTACTS_TAB = "hrefTabContentContacts";
        private const string CONTACT_FOLD_BTN = "//*[@id=\"tabContentDetails\"]/div/div[1]/a";
        private const string NEW_CONTACT_BTN = "//*[@id=\"tabContentDetails\"]/div/div[1]/div/a[2]";
        private const string CONTACT_NAME = "Contact_Name";
        private const string CONTACT_MAIL = "Contact_Mail";
        private const string CONTACT_SITE_SELECTOR = "//*[@id=\"SelectedSites_ms\"]/span[2]";
        private const string CONTACT_CHECK_ALL = "/html/body/div[12]/ul/li[15]/label/span";
        private const string CONTACT_VALIDATE_BTN = "//*[@id=\"modal-1\"]/div/div/div[2]/div/form/div[2]/button[2]";
        private const string CONTACT_VALIDATE_BTN_DEV = "//*[@id=\"modal-1\"]/div[2]/div/form/div[2]/button[2]";
        private const string FULL_NAME_CONTACT = "//*[@id=\"list-item-with-action\"]/div/div[1]/div/div[2]/span";
        private const string CONTACT_LIST = "//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]";
        private const string EDIT_CONTACT = "//*[@id=\"list-item-with-action\"]/div[{0}]/div[1]/div/div[3]/a[1]";
        private const string EDIT_CONTACT_FOR_CLAIM = "//*[@id=\"list-item-with-action\"]/div[1]/div[1]/div/div[3]/a[1]";
        private const string DETAILS_CONTACTS = "/html/body/div[3]/div/div/div/div/div[2]/div/div[*]/div[2]";
        private const string ROW_CONTACTS = "//*[@id=\"list-item-with-action\"]/div/div[1]";

        // Accounting

        private const string ACCOUNTING_TAB = "hrefTabContentAccounting";

        private const string NEW_ACCOUNTING = "//*[@id=\"dat-container\"]/div/div/a[2]";
        private const string SITE_SEARCH_ACCOUNTING = "SelectedSites_ms";
        private const string CHECK_ALL_SITES = "/html/body/div[12]/div/ul/li[1]/a/span[1]";
        private const string ACCOUNTING_CONFIRM_BTN = "last";
        private const string ACCOUNTING = "//*[@id=\"dat-container\"]/table/tbody/tr[*]/td[1][contains(@title, '{0}')]";
        private const string EDIT_ACCOUNTING = "//*[@id=\"dat-container\"]/table/tbody/tr[2]/td[1][contains(@title, '{0}')]/../td[8]/a";
        private const string THIRD_PART = "SupplierAccounting_IdAnalytical";
        private const string ACCOUNTING_ID = "SupplierAccounting_IdAccounting";
        private const string THIRD_PART_DTI = "SupplierAccounting_IdAnalyticalDueToInvoice";
        private const string ACCOUNTING_ID_DTI = "SupplierAccounting_IdAccountingDueToInvoice";

        private const string ACCOUNTING_TIERS_DETAILS_TAB = "hrefTabContentAccountingTiers";
        private const string TIERS_CODE = "TiersCode";
        private const string TIERS_NAME = "TiersName";
        private const string TIERS_ADDRESS_IDENTIFIER = "TiersAddressIdentifier";
        private const string TIERS_ADDRESS = "TiersAddress";
        private const string TIERS_DOMICILIATION_IDENTIFIER = "TiersDomiciliationIdentifier";
        private const string ACCOUNTING_TIERS_DETAILS_LINE = "//td[contains(text(), '{0}')]";
        private const string EDIT_ACCOUNTING_TIERS_DETAILS = "//td[contains(text(), '{0}')]/..//span[@class='glyphicon glyphicon glyphicon-pencil']";
        private const string DELETE_ACCOUNTING_TIERS_DETAILS = "//td[contains(text(), '{0}')]/..//span[@class='glyphicon glyphicon-trash']";

        // Certifications

        private const string CERTIFICATION_TAB = "hrefTabContentCertifications";
        private const string NEW_CERTIFICATION = "//*[@id=\"dat-container\"]/div[1]/div/a[2]";
        private const string CERTIFICATION_ID = "SelectedCertificationId";
        private const string CERTIFICATION_EXPIRATION_DATE = "expiration-date";
        private const string ADD_CERTIFICATION = "//*[@id=\"add-certif-form\"]/div[2]/button[2]";
        private const string DELETE_CERTIFICATION = "//*[@id=\"form-certif_0\"]/div/div[6]/a/span";
        private const string DELETE_CERTIFICATION_TARGET = "//*/div[text()='{0}']/../div[6]/a";
        private const string DELETE_CERTIFICATION_TARGET_CONFIRM = "dataConfirmOK";
        // EDI
        private const string EDI_TAB = "hrefTabContentDetailsEDI";
        private const string ID_EDI = "IdEdi";
        private const string REPOSITORY = "EdiFileSFTP";

        private const string ACTIVE_NOT_PURCHASABLE = "ItemsIndexModel_ShowOnlyActiveNotPurchasable";
        private const string PACKAGINGS_PURCHASABLE = "/html/body/div[3]/div/div[2]/div[2]/div/div[2]/div[1]/div[2]/div/table/tbody/tr[*]/td[15]/div/input";
        private const string FIRST_ITEM = "/html/body/div[3]/div/div/div/div/div[2]/div/div[2]/div/div[1]/div[1]/div/div[2]/table/tbody/tr/td[1]";
        private const string SELECTED_GRP = "ddl_group";
        private const string ITEM_PACKAGING_SITE = "/html/body/div[3]/div/div[2]/div[2]/div/div[2]/div[1]/div[2]/div/table/tbody/tr/td[1]/span";
        private const string FILTER_BY_SUBGROUP = "/html/body/div[3]/div/div/div/div/div[1]/div/div/div/form/div[3]/div[4]/div/div/div[1]/button";
        private const string UNCHECK_ALL_SUBGROUPS = "/html/body/div[12]/div/ul/li[2]/a/span[2]";
        private const string FILTER_BY_SUBGROUP_INPUT = "/html/body/div[11]/div/div/label/input";
        private const string FIRST_SUBGRP = "/html/body/div[11]/ul/li/ul/li/label";
        private const string SELECTED_SUBGRP = "//*[@id=\"ddl_subgroup\"]/option[@selected=\"selected\"]";
        private const string ITEMS = "//*[@id=\"list-item-with-action\"]/div/div/div/div[2]/table/tbody/tr[*]/td[2]";
        private const string EXPIRATION_DATE = "expiration-date-picker_0";
        private const string ITEM_GROUP = "//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]/table/tbody/tr/td[3]";

        //searchField
        private const string SEARCH_FIELD = "tbSearchPatternWithAutocomplete";
        //Storage
        private const string STORAGE_TAB = "hrefTabContentDetailsStorage";
        private const string UNSELECT_ALL_BTN = "/html/body/div[3]/div/div/div/div/div[2]/div[1]/div/a[1]";
        private const string SELECTALL_BTN = "/html/body/div[3]/div/div/div/div/div[2]/div[1]/div/a[2]";

        private const string LIST_ITEMS = "//*[@id=\"list-item-with-action\"]/div[*]";
        private const string SETUNPURCHAISINGITEM = "//*[@id=\"unpurchasable-items-btn\"]";
        private const string VERIFDEACTIVATED = "//*[@id=\"formId\"]/div[2]/table/tbody/tr[2]/td[2]";
        private const string UNFOLDALL = "/html/body/div[3]/div/div/div/div/div[2]/div/div[1]/div[1]/a";
        private const string FLECHEFORDELTE = "/html/body/div[3]/div/div/div/div/div[2]/div/div[2]/div/div/div[1]/div/div[1]";
        private const string DELETEBUTTON = "/html/body/div[3]/div/div/div/div/div[2]/div/div[2]/div/div/div[2]/div/table/tbody/tr[2]/td[6]/a";
        private const string CONFIRMBUTTONPOPUP = "//*[@id=\"DeactivateItemsReportForm\"]/div[2]/button";


        private const string CONTACT_ELEMENT = "/html/body/div[3]/div/div/div/div/div[2]/div/div/div[1]/div/div[2]";
        private const string CONFIRM_BUTTON = "//*[@id=\"dataConfirmOK\"]";
        private const string SELECT_ALL_CHECKBOXES_COPYPRICE = "/html/body/div[4]/div/div/div/div/form/div[2]/div[1]/div[4]/table/tbody/tr[*]/td[1]";
        private const string REVERSE_VAT = "/html/body/div[3]/div/div/div/div/div/form/div/div[9]/div[3]/div[1]/div[3]/input";
        private const string TIME_LIMIT = "/html/body/div[3]/div/div/div/div[1]/div/div[2]/div[1]/div/div[8]/input";
        private const string FIRST_MIN_AMOUNT = "DeliveriesSites_0__MinimumOrder";
        private const string MIN_AMOUNT = "/html/body/div[3]/div/div/div/div[1]/div/div[2]/div[1]/div/div[5]/input";
        private const string FILTER_NAME = "SearchPattern";
        private const string SUPPLIER_NAME = "/html/body/div[3]/div/div/div/div/div[2]/table/tbody/tr[2]/td[2]";
        private const string FIRST_SHIPPING_COST = "DeliveriesSites_0__ShippingCost";
        // ___________________________________________ VARIABLES _______________________________________________


        [FindsBy(How = How.XPath, Using = SETUNPURCHAISINGITEM)]
        private IWebElement _setunpurchaisingitem;
        [FindsBy(How = How.XPath, Using = CONFIRMBUTTONPOPUP)]
        private IWebElement _confirmbuttonpopup;
        [FindsBy(How = How.XPath, Using = DELETEBUTTON)]
        private IWebElement _deletebutton;
        [FindsBy(How = How.XPath, Using = FLECHEFORDELTE)]
        private IWebElement _fleche;

        [FindsBy(How = How.XPath, Using = UNFOLDALL)]
        private IWebElement _unfoldall;
        [FindsBy(How = How.XPath, Using = LIST_ITEMS)]
        private IWebElement _listeitems;

        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;

        // General Informations

        [FindsBy(How = How.Id, Using = GENERAL_INFORMATION_TAB)]
        private IWebElement _generalInformationTab;

        [FindsBy(How = How.Id, Using = SANITARY_NUMBER)]
        private IWebElement _sanitaryNumber;

        [FindsBy(How = How.Id, Using = SIRET_FIELD)]
        private IWebElement _siretField;

        [FindsBy(How = How.Id, Using = DATE_END_CONTRACT)]
        private IWebElement _dateEndContract;

        [FindsBy(How = How.Id, Using = LAST_AUDIT_DATE)]
        private IWebElement _lastAuditDate;

        [FindsBy(How = How.Id, Using = NEXT_AUDIT_DATE)]
        private IWebElement _nextAuditDate;

        [FindsBy(How = How.Id, Using = DEACTIVATE_SUPPLIER)]
        private IWebElement _deactivateSupplier;

        [FindsBy(How = How.Id, Using = HEADER_PRINT_BUTTON)]
        private IWebElement _headerPrintButton;

        // Items

        [FindsBy(How = How.Id, Using = ITEM_TAB)]
        private IWebElement _itemTabBtn;

        [FindsBy(How = How.XPath, Using = NEW_ITEM)]
        private IWebElement _createNewItem;

        [FindsBy(How = How.XPath, Using = ITEM_NAME)]
        private IWebElement _itemName;

        [FindsBy(How = How.XPath, Using = EXPORT_ITEMS)]
        private IWebElement _exportItems;

        [FindsBy(How = How.Id, Using = CONFIRM_EXPORT_ITEMS)]
        private IWebElement _confirmExportItems;

        [FindsBy(How = How.Id, Using = DEACTIVATE_ITEMS)]
        private IWebElement _deactivateItems;

        [FindsBy(How = How.Id, Using = UNPURCHASABLE_ITEMS)]
        private IWebElement _unpurchasableItems;

        [FindsBy(How = How.XPath, Using = CHECK_ALL_SITES_ON_ITEMS)]
        private IWebElement _checkAllSitesOnItems;

        [FindsBy(How = How.XPath, Using = IMPORT_ITEMS)]
        private IWebElement _importItems;

        [FindsBy(How = How.XPath, Using = UNFOLD_BTN)]
        private IWebElement _unfoldBtn;

        [FindsBy(How = How.Id, Using = FILTER_ITEM_SEARCH_BY_NAME)]
        private IWebElement _searchItemByName;

        [FindsBy(How = How.Id, Using = RESET_FILTER_ITEMS_DEV)]
        private IWebElement _resetFilterItemsDev;

        [FindsBy(How = How.XPath, Using = RESET_FILTER_ITEMS_PATCH)]
        private IWebElement _resetFilterItemsPatch;

        [FindsBy(How = How.Id, Using = SITE_SEARCH_ITEMS)]
        private IWebElement _siteDropdownList;

        [FindsBy(How = How.Id, Using = ACCOUNTING_CONFIRM_BTN)]
        private IWebElement _confirmDelete;

        [FindsBy(How = How.XPath, Using = BTN_OK)]
        private IWebElement _okBtn;

        [FindsBy(How = How.XPath, Using = ITEM_GROUP)]
        private IWebElement _itemGroup;

        [FindsBy(How = How.XPath, Using = FAST_COPY_PRICE)]
        private IWebElement _fast_copy_price;

        [FindsBy(How = How.XPath, Using = COPY_BTN)]
        private IWebElement _copy_btn;

        [FindsBy(How = How.XPath, Using = SELECT_ALL)]
        private IWebElement _select_All;

        [FindsBy(How = How.XPath, Using = COPY_PRICES)]
        private IWebElement _copy_prices;

        [FindsBy(How = How.XPath, Using = ITEM_EDITED)]
        private IWebElement _item_edited;

        [FindsBy(How = How.Id, Using = COPY_ALL_PRICES_FROM)]
        private IWebElement _copyAllPricesFrom;

        [FindsBy(How = How.XPath, Using = OK_COPY)]
        private IWebElement _okCopy;


        //Deliveries

        [FindsBy(How = How.Id, Using = DELIVERY_TAB)]
        private IWebElement _deliveriesTab;

        [FindsBy(How = How.XPath, Using = MONDAY)]
        private IWebElement _monday;

        [FindsBy(How = How.XPath, Using = TUESDAY)]
        private IWebElement _tuesday;

        [FindsBy(How = How.XPath, Using = WEDNESDAY)]
        private IWebElement _wednesday;

        [FindsBy(How = How.XPath, Using = THURSDAY)]
        private IWebElement _thursday;

        [FindsBy(How = How.XPath, Using = FRIDAY)]
        private IWebElement _friday;

        [FindsBy(How = How.XPath, Using = SATURDAY)]
        private IWebElement _saturday;

        [FindsBy(How = How.XPath, Using = SUNDAY)]
        private IWebElement _sunday;

        [FindsBy(How = How.XPath, Using = SHOW_COMMENT_PANEL)]
        private IWebElement _showCommentPanel;

        [FindsBy(How = How.Id, Using = COMMENT_PANEL)]
        private IWebElement _commentPanel;

        [FindsBy(How = How.Id, Using = FIRST_MIN_AMOUNT)]
        private IWebElement _firstMinAmount;

        [FindsBy(How = How.Id, Using = FIRST_SHIPPING_COST)]
        private IWebElement _firstShippingCost;

        // Contacts

        [FindsBy(How = How.Id, Using = CONTACTS_TAB)]
        private IWebElement _contactsTab;

        [FindsBy(How = How.XPath, Using = CONTACT_FOLD_BTN)]
        private IWebElement _contactFoldBtn;

        [FindsBy(How = How.XPath, Using = NEW_CONTACT_BTN)]
        private IWebElement _newContactBtn;

        [FindsBy(How = How.Id, Using = CONTACT_NAME)]
        private IWebElement _contactName;

        [FindsBy(How = How.Id, Using = CONTACT_MAIL)]
        private IWebElement _contactMail;

        [FindsBy(How = How.XPath, Using = CONTACT_SITE_SELECTOR)]
        private IWebElement _siteSelector;

        [FindsBy(How = How.XPath, Using = CONTACT_CHECK_ALL)]
        private IWebElement _contactCheckAll;

        [FindsBy(How = How.XPath, Using = CONTACT_VALIDATE_BTN)]
        private IWebElement _contactValidateBtn;

        [FindsBy(How = How.XPath, Using = FULL_NAME_CONTACT)]
        private IWebElement _fullNameContact;

        [FindsBy(How = How.XPath, Using = DETAILS_CONTACTS)]
        private IWebElement _details_contact;

        // Accounting

        [FindsBy(How = How.Id, Using = ACCOUNTING_TAB)]
        private IWebElement _acccountingTab;

        [FindsBy(How = How.XPath, Using = NEW_ACCOUNTING)]
        private IWebElement _newAccountingBtn;

        [FindsBy(How = How.Id, Using = SITE_SEARCH_ACCOUNTING)]
        private IWebElement _siteSearchAccounting;

        [FindsBy(How = How.XPath, Using = CHECK_ALL_SITES)]
        private IWebElement _checkAllSites;

        [FindsBy(How = How.Id, Using = ACCOUNTING_CONFIRM_BTN)]
        private IWebElement _accountingCreateBtn;

        [FindsBy(How = How.XPath, Using = ACCOUNTING)]
        private IWebElement _accounting;

        [FindsBy(How = How.Id, Using = THIRD_PART)]
        private IWebElement _thirdPart;

        [FindsBy(How = How.Id, Using = ACCOUNTING_ID)]
        private IWebElement _accountingId;

        [FindsBy(How = How.Id, Using = THIRD_PART_DTI)]
        private IWebElement _thirdPartDTI;

        [FindsBy(How = How.Id, Using = ACCOUNTING_ID_DTI)]
        private IWebElement _accountingIdDTI;

        [FindsBy(How = How.Id, Using = TIERS_CODE)]
        private IWebElement _tiersCodeInput;

        [FindsBy(How = How.Id, Using = TIERS_NAME)]
        private IWebElement _tiersNameInput;

        [FindsBy(How = How.Id, Using = TIERS_ADDRESS_IDENTIFIER)]
        private IWebElement _tiersAddressIdInput;

        [FindsBy(How = How.Id, Using = TIERS_DOMICILIATION_IDENTIFIER)]
        private IWebElement _tiersDomiciliationInput;

        [FindsBy(How = How.Id, Using = TIERS_ADDRESS)]
        private IWebElement _tiersAddressInput;

        // Certifications

        [FindsBy(How = How.Id, Using = CERTIFICATION_TAB)]
        private IWebElement _certificationTab;

        [FindsBy(How = How.XPath, Using = NEW_CERTIFICATION)]
        private IWebElement _newCertification;

        [FindsBy(How = How.Id, Using = CERTIFICATION_ID)]
        private IWebElement _certificationId;

        [FindsBy(How = How.Id, Using = CERTIFICATION_EXPIRATION_DATE)]
        private IWebElement _certificationExpirationDate;

        [FindsBy(How = How.XPath, Using = ADD_CERTIFICATION)]
        private IWebElement _addCertification;

        [FindsBy(How = How.XPath, Using = DELETE_CERTIFICATION)]
        private IWebElement _delCertification;

        // Storage
        [FindsBy(How = How.Id, Using = STORAGE_TAB)]
        private IWebElement _storageTab;

        [FindsBy(How = How.XPath, Using = SELECTALL_BTN)]
        private IWebElement _selectAll_btn;

        [FindsBy(How = How.XPath, Using = UNSELECT_ALL_BTN)]
        private IWebElement _unselectAll_btn;

        // Storage

        [FindsBy(How = How.Id, Using = SITE_FC)]
        private IWebElement _sitefastcopy;

        // __________________________________________ UTILITAIRE _________________________________________

        public SuppliersPage BackToList()
        {
            _backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            _backToList.Click();
            WaitForLoad();
            return new SuppliersPage(_webDriver, _testContext);
        }

        public void ClickOkPopup()
        {
            _confirmbuttonpopup = WaitForElementIsVisible(By.XPath(CONFIRMBUTTONPOPUP));
            _confirmbuttonpopup.Click();
            WaitForLoad();
        }

        // __________________________________________ METHODES ___________________________________________

        // General Informations

        public SupplierGeneralInfoTab ClickOnGeneralInformationTab()
        {
            _generalInformationTab = WaitForElementIsVisible(By.Id(GENERAL_INFORMATION_TAB));
            _generalInformationTab.Click();
            WaitForLoad();
            return new SupplierGeneralInfoTab(_webDriver, _testContext);
        }

        public void CompletGeneralInformations()
        {
            _sanitaryNumber = WaitForElementIsVisible(By.Id(SANITARY_NUMBER));
            _sanitaryNumber.SetValue(ControlType.TextBox, "Sanitary-" + new Random().Next().ToString());

            _siretField = WaitForElementIsVisible(By.Id(SIRET_FIELD));
            _siretField.SetValue(ControlType.TextBox, new Random().Next().ToString());

            _dateEndContract = WaitForElementIsVisible(By.Id(DATE_END_CONTRACT));
            _dateEndContract.SetValue(ControlType.DateTime, DateUtils.Now.AddYears(10));

            _lastAuditDate = WaitForElementIsVisible(By.Id(LAST_AUDIT_DATE));
            _lastAuditDate.SetValue(ControlType.DateTime, DateUtils.Now);

            _nextAuditDate = WaitForElementIsVisible(By.Id(NEXT_AUDIT_DATE));
            _nextAuditDate.SetValue(ControlType.DateTime, DateUtils.Now.AddMonths(6));

            _nextAuditDate.SendKeys(Keys.PageUp);
            //Ajout Thread.Sleep() avant d'éxecuter la prochaine étape
            Thread.Sleep(2000);
        }

        public void DeactivateSupplier()
        {
            ShowExtendedMenu();

            _deactivateSupplier = WaitForElementExists(By.Id(DEACTIVATE_SUPPLIER));
            _deactivateSupplier.Click();
            WaitForLoad();
            // modal
            var ok = WaitForElementIsVisible(By.Id("dataAlertCancel"));
            ok.Click();
            WaitForLoad();

            //Thread.Sleep() pour tenir compte du nouveau changement
            //Thread.Sleep(2000);

        }

        public string SanitaryNumber()
        {
            _sanitaryNumber = WaitForElementIsVisible(By.Id(SANITARY_NUMBER));
            return _sanitaryNumber.GetAttribute("value");
        }

        public string SiretNumber()
        {
            _siretField = WaitForElementIsVisible(By.Id(SIRET_FIELD));
            return _siretField.GetAttribute("value");
        }

        public string LastAuditDate()
        {
            _lastAuditDate = WaitForElementIsVisible(By.Id(LAST_AUDIT_DATE));
            return _lastAuditDate.GetAttribute("value");
        }

        public string NextAuditDate()
        {
            _nextAuditDate = WaitForElementIsVisible(By.Id(NEXT_AUDIT_DATE));
            return _nextAuditDate.GetAttribute("value");
        }

        // Items
        public void GetItemsTab()
        {
            _itemTabBtn = WaitForElementIsVisibleNew(By.Id(ITEM_TAB));
            _itemTabBtn.Click();
            WaitForLoad();
        }

        public ItemCreateModalPage ItemCreatePage()
        {
            _createNewItem = WaitForElementIsVisibleNew(By.XPath(NEW_ITEM));
            _createNewItem.Click();

            return new ItemCreateModalPage(_webDriver, _testContext);
        }

        public string getItemName()
        {
            try
            {
                _itemName = WaitForElementExists(By.XPath(ITEM_NAME));
                return _itemName.GetAttribute("innerText");
            }
            catch
            {
                return "";
            }
        }


        public bool IsItemPresent()
        {

            var itemName = _webDriver.FindElements(By.XPath(ITEM_ELEMENT));

            if (itemName.Count == 1)
            {
                return true;
            }

            return false;

        }

        public int NombreItem()
        {

            var itemName = _webDriver.FindElements(By.XPath("//*[@id=\"list-item-with-action\"]/div/div[1]/div/div[2]/table/tbody/tr"));
            return itemName.Count;

        }

        public void ExportItems()
        {
            _exportItems = WaitForElementIsVisible(By.XPath(EXPORT_ITEMS));
            _exportItems.Click();
            WaitForLoad();
            _confirmExportItems = WaitForElementIsVisible(By.Id(CONFIRM_EXPORT_ITEMS));
            _confirmExportItems.Click();
            WaitPageLoading();

            IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));
            ClickPrintButton();

            WaitForDownload();
            Close();
            ClosePrintButton();
        }

        public FileInfo GetExportExcelFileItem(FileInfo[] taskFiles)
        {
            List<FileInfo> correctDownloadFiles = new List<FileInfo>();

            foreach (var file in taskFiles)
            {
                //  Test REGEX
                if (IsSupplierItemExcelFileCorrect(file.Name))
                {
                    correctDownloadFiles.Add(file);
                }
            }

            if (correctDownloadFiles.Count <= 0)
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

        public bool IsSupplierItemExcelFileCorrect(string filePath)
        {
            // "PurchaseBook 2019-10-29 16-16-40.xlsx";

            string re1 = "((?:[a-z][a-z]+))";   // Word 1
            string re2 = "(\\s+)";  // White Space 1
            string re3 = "((?:(?:[1]{1}\\d{1}\\d{1}\\d{1})|(?:[2]{1}\\d{3}))[-:\\/.](?:[0]?[1-9]|[1][012])[-:\\/.](?:(?:[0-2]?\\d{1})|(?:[3][01]{1})))(?![\\d])";   // YYYYMMDD 1
            string re4 = "(\\s+)";  // White Space 2
            string re5 = "(\\d)";   // Any Single Digit 1
            string re6 = "(\\d)";   // Any Single Digit 2
            string re7 = "([-+]\\d+)";  // Integer Number 1
            string re8 = "([-+]\\d+)";  // Integer Number 2
            string re9 = "(.)"; // Any Single Character 1
            string re10 = "((?:[a-z][a-z]+))";  // Word 2

            Regex r = new Regex(re1 + re2 + re3 + re4 + re5 + re6 + re7 + re8 + re9 + re10, RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = r.Match(filePath);

            return m.Success;
        }

        public SupplierImport Import()
        {
            _importItems = WaitForElementIsVisible(By.XPath(IMPORT_ITEMS));
            _importItems.Click();
            WaitForLoad();

            //Léger Sleep sur l'import le temps que le fichier soit entièrement téléchargé
            Thread.Sleep(2000);

            return new SupplierImport(_webDriver, _testContext);
        }

        public void SetUnpurchasableItems()
        {
            _unpurchasableItems = WaitForElementIsVisible(By.Id(UNPURCHASABLE_ITEMS));
            _unpurchasableItems.Click();
            WaitForLoad();

            _siteDropdownList = WaitForElementIsVisible(By.Id(SITE_SEARCH_ITEMS));
            _siteDropdownList.Click();

            _checkAllSitesOnItems = WaitForElementIsVisible(By.XPath(CHECK_ALL_SITES_ON_ITEMS), nameof(CHECK_ALL_SITES_ON_ITEMS));
            _checkAllSitesOnItems.Click();
            WaitForLoad();

            _confirmDelete = WaitForElementIsVisible(By.Id(ACCOUNTING_CONFIRM_BTN));
            _confirmDelete.Click();
            WaitForLoad();

            _okBtn = WaitForElementIsVisible(By.XPath(BTN_OK));
            _okBtn.Click();
            WaitForLoad();
        }

        public void DeactivateItems()
        {
            _deactivateItems = WaitForElementIsVisible(By.Id(DEACTIVATE_ITEMS));
            _deactivateItems.Click();
            WaitForLoad();

            _siteDropdownList = WaitForElementIsVisible(By.Id(SITE_SEARCH_ITEMS));
            _siteDropdownList.Click();
            WaitForLoad();

            _checkAllSitesOnItems = WaitForElementIsVisible(By.XPath(CHECK_ALL_SITES_ON_ITEMS), nameof(CHECK_ALL_SITES_ON_ITEMS));
            _checkAllSitesOnItems.Click();
            WaitForLoad();

            _confirmDelete = WaitForElementIsVisible(By.Id(ACCOUNTING_CONFIRM_BTN));
            _confirmDelete.Click();
            WaitForLoad();

            _okBtn = WaitForElementIsVisible(By.XPath("//*[@id='formId']/div[3]/button[2]"));
            _okBtn.Click();
            WaitPageLoading();
            WaitForLoad();
        }


        public bool VerifDeactivateItems()
        {
            _deactivateItems = WaitForElementIsVisible(By.Id(DEACTIVATE_ITEMS));
            _deactivateItems.Click();
            WaitForLoad();

            _siteDropdownList = WaitForElementIsVisible(By.Id(SITE_SEARCH_ITEMS));
            _siteDropdownList.Click();
            WaitForLoad();

            _checkAllSitesOnItems = WaitForElementIsVisible(By.XPath(CHECK_ALL_SITES_ON_ITEMS), nameof(CHECK_ALL_SITES_ON_ITEMS));
            _checkAllSitesOnItems.Click();
            WaitForLoad();

            _confirmDelete = WaitForElementIsVisible(By.Id(ACCOUNTING_CONFIRM_BTN));
            _confirmDelete.Click();
            WaitForLoad();

            if (isElementVisible(By.XPath(VERIFDEACTIVATED)))
            {

                return true;
            }
            else
            {
                return false;
            }
            _okBtn = WaitForElementIsVisible(By.XPath("//*[@id='formId']/div[3]/button[2]"));
            _okBtn.Click();

        }
        public enum FilterType
        {
            Search,
            // itemsTab
            SortBy,
            ShowOnlyActive,
            ShowOnlyInactive,
            ActiveNotPurchasable,
            ByGroup,
            BySite,
            BySubGroupe
        }

        public void Filter(FilterType filterType, object value)
        {

            switch (filterType)
            {
                case FilterType.Search:
                    _searchItemByName = WaitForElementIsVisible(By.Id(FILTER_ITEM_SEARCH_BY_NAME));
                    _searchItemByName.SetValue(ControlType.TextBox, value);
                    WaitPageLoading();

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
                case FilterType.SortBy:
                    var cbSortBy = WaitForElementIsVisible(By.Id(FILTER_SORT_BY));
                    var selectSortBy = new SelectElement(cbSortBy);
                    selectSortBy.SelectByText((string)value);
                    break;
                case FilterType.ShowOnlyActive:
                    var showActive = WaitForElementIsVisible(By.Id("ItemsIndexModel_ShowActive"));
                    showActive.Click();
                    // pourquoi ?
                    PageUp();
                    break;
                case FilterType.ShowOnlyInactive:
                    var showInactive = WaitForElementIsVisible(By.Id("ItemsIndexModel_ShowOnlyInactive"));
                    showInactive.Click();
                    // pourquoi ?
                    PageUp();
                    break;
                case FilterType.ActiveNotPurchasable:
                    var activeNotPurchasable = WaitForElementIsVisible(By.Id(ACTIVE_NOT_PURCHASABLE));
                    activeNotPurchasable.Click();
                    WaitForLoad();
                    break;
                case FilterType.ByGroup:
                    ComboBoxSelectById(new ComboBoxOptions("ItemsIndexModelSelectedGroups_ms", (string)value));
                    break;
                case FilterType.BySite:
                    ComboBoxSelectById(new ComboBoxOptions("ItemsIndexModelSelectedSites_ms", (string)value));
                    break;
                case FilterType.BySubGroupe:
                    ComboBoxSelectById(new ComboBoxOptions("ItemsIndexModelSelectedSubGroups_ms", (string)value));
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);

            }

            WaitPageLoading();
            Thread.Sleep(1500);
        }

        public void PageUp()
        {
            IWebElement html = null;
            html = _webDriver.FindElement(By.TagName("html"));
            html.SendKeys(Keys.PageUp);
            html.SendKeys(Keys.PageUp);
            html.SendKeys(Keys.PageUp);
            html.SendKeys(Keys.PageUp);
        }

        public void ResetFilter()
        {
            try
            {
                _resetFilterItemsDev = WaitForElementIsVisibleNew(By.Id(RESET_FILTER_ITEMS_DEV));
                _resetFilterItemsDev.Click();
            }
            catch
            {
                _resetFilterItemsPatch = WaitForElementIsVisibleNew(By.XPath(RESET_FILTER_ITEMS_PATCH));
                _resetFilterItemsPatch.Click();
            }

            WaitForLoad();

            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                // pas de date
            }
        }

        public Boolean IsUnfoldAll()
        {
            bool valueBool = false;
            try
            {
                WaitForElementExists(By.XPath(UNFOLD_BTN_ALL));

                // Temps nécessaire pour que l'élément change de classe
                Thread.Sleep(1000);
                var content = WaitForElementIsVisible(By.XPath(CONTENT));
                if (content.GetAttribute("class") == "panel-collapse collapse in")
                    valueBool = true;
            }
            catch
            {
                valueBool = false;
            }

            return valueBool;
        }

        public Boolean IsFoldAll()
        {
            bool valueBool = false;
            try
            {
                WaitForElementExists(By.XPath("//*[@id=\"unfoldBtn\"]/span[@class='fas fa-arrows-alt-v']"));


                // Temps nécessaire pour que l'élément change de classe
                Thread.Sleep(1000);
                var content = WaitForElementExists(By.XPath(CONTENT));
                if (content.GetAttribute("class") == "panel-collapse collapse")
                    valueBool = true;
            }
            catch
            {
                valueBool = false;
            }

            return valueBool;
        }

        public void UnFoldorFoldItems()
        {
            _unfoldBtn = WaitForElementIsVisible(By.XPath(UNFOLD_BTN));
            _unfoldBtn.Click();
            WaitForLoad();
            WaitPageLoading();
        }

        public void ClickOnFirstCopyPrice()
        {
            _fast_copy_price = WaitForElementIsVisible(By.XPath(FAST_COPY_PRICE));
            _fast_copy_price.Click();
            WaitForLoad();
        }
        public void ClickOnSelectAllForCopyPrice()
        {
            _select_All = WaitForElementIsVisible(By.XPath(SELECT_ALL));
            _select_All.Click();
            WaitForLoad();
        }
        public void ScrollUntilCopyIsVisible()
        {
            ScrollUntilElementIsInView(By.XPath(COPY_PRICES));
        }
        public void ClickInCopyBtn()
        {
            _copy_prices = WaitForElementIsVisible(By.XPath(COPY_PRICES));
            _copy_prices.Click();
            WaitForLoad();
        }
        public string CopyPopUpMessage()
        {
            _item_edited = WaitForElementIsVisible(By.XPath(ITEM_EDITED));
            WaitForLoad();
            return _item_edited.Text;

        }

        //Deliveries

        public void ClickToDeliveryBtn()
        {
            _deliveriesTab = WaitForElementIsVisible(By.Id(DELIVERY_TAB));
            _deliveriesTab.Click();
            WaitForLoad();
        }

        public void CompletDeliveriesDaysAndComment()
        {
            ClickToDeliveryBtn();

            // Select days for ALL SITES
            _monday = WaitForElementIsVisible(By.XPath(MONDAY));
            _monday.Click();

            _tuesday = WaitForElementIsVisible(By.XPath(TUESDAY));
            _tuesday.Click();

            _wednesday = WaitForElementIsVisible(By.XPath(WEDNESDAY));
            _wednesday.Click();

            _thursday = WaitForElementIsVisible(By.XPath(THURSDAY));
            _thursday.Click();

            _friday = WaitForElementIsVisible(By.XPath(FRIDAY));
            _friday.Click();

            _saturday = WaitForElementIsVisible(By.XPath(SATURDAY));
            _saturday.Click();
            WaitForLoad();

            // Comment line
            ClickOnDeliveryComment();

            _commentPanel = WaitForElementIsVisible(By.Id(COMMENT_PANEL));
            _commentPanel.SetValue(ControlType.TextBox, "This is a comment");

            //Ajout Thread.sleep() avant de passer à la prochaine étape
            Thread.Sleep(2000);

            _webDriver.Navigate().Refresh();
            WaitForLoad();
        }

        public void SelectAllDeliveryweekdays()
        {
            ClickToDeliveryBtn();

            var listsite = _webDriver.FindElements(By.XPath(LIST_ITEMS));
            // ligne 2 c'est "ALL SITES" donc on le zap
            for (int i = 3; i <= listsite.Count(); i++)
            {
                WaitForLoad();

                var monday = WaitForElementIsVisible(By.XPath("//*[@id=\"list-item-with-action\"]/div[" + i + "]/div/div/div[10]/div/label[text()='Mon']"));
                monday.Click();

                WaitForLoad();

                var tuesday = WaitForElementIsVisible(By.XPath("//*[@id=\"list-item-with-action\"]/div[" + i + "]/div/div/div[10]/div/label[text()='Tue']"));
                tuesday.Click();

                WaitForLoad();

                var wednesday = WaitForElementIsVisible(By.XPath("//*[@id=\"list-item-with-action\"]/div[" + i + "]/div/div/div[10]/div/label[text()='Wed']"));
                wednesday.Click();

                WaitForLoad();

                var thursday = WaitForElementIsVisible(By.XPath("//*[@id=\"list-item-with-action\"]/div[" + i + "]/div/div/div[10]/div/label[text()='Thu']"));
                thursday.Click();

                WaitForLoad();

                var friday = WaitForElementIsVisible(By.XPath("//*[@id=\"list-item-with-action\"]/div[" + i + "]/div/div/div[10]/div/label[text()='Fri']"));
                friday.Click();

                WaitForLoad();

                var saturday = WaitForElementIsVisible(By.XPath("//*[@id=\"list-item-with-action\"]/div[" + i + "]/div/div/div[10]/div/label[text()='Sat']"));
                saturday.Click();

                WaitForLoad();

                var sunday = WaitForElementIsVisible(By.XPath("//*[@id=\"list-item-with-action\"]/div[" + i + "]/div/div/div[10]/div/label[text()='Sun']"));
                sunday.Click();

                WaitForLoad();

            }

            //Ajout Thread.sleep() avant de passer à la prochaine étape
            Thread.Sleep(2000);
            PageUp();
            _contactsTab = WaitForElementIsVisible(By.Id("hrefTabContentContacts"));
            new Actions(_webDriver).MoveToElement(_contactsTab).Click().Perform();

        }

        public bool AreAllDeliveryweekdays()
        {

            _deliveriesTab = WaitForElementIsVisible(By.Id(DELIVERY_TAB));
            new Actions(_webDriver).MoveToElement(_deliveriesTab).Click().Perform();
            Thread.Sleep(5000);

            bool valuebool = true;

            var colorGreen = "rgb(104, 162, 4)";

            var listsite = _webDriver.FindElements(By.XPath(LIST_ITEMS));
            // ligne 2 c'est "ALL SITES" donc on le zap
            for (int i = 3; i <= listsite.Count(); i++)
            {
                WaitForLoad();

                var monday = WaitForElementIsVisible(By.XPath("//*[@id=\"list-item-with-action\"]/div[" + i + "]/div/div/div[10]/div/label[text()='Mon']"));

                WaitForLoad();

                var tuesday = WaitForElementIsVisible(By.XPath("//*[@id=\"list-item-with-action\"]/div[" + i + "]/div/div/div[10]/div/label[text()='Tue']"));

                WaitForLoad();

                var wednesday = WaitForElementIsVisible(By.XPath("//*[@id=\"list-item-with-action\"]/div[" + i + "]/div/div/div[10]/div/label[text()='Wed']"));

                WaitForLoad();

                var thursday = WaitForElementIsVisible(By.XPath("//*[@id=\"list-item-with-action\"]/div[" + i + "]/div/div/div[10]/div/label[text()='Thu']"));

                WaitForLoad();

                var friday = WaitForElementIsVisible(By.XPath("//*[@id=\"list-item-with-action\"]/div[" + i + "]/div/div/div[10]/div/label[text()='Fri']"));

                WaitForLoad();

                var saturday = WaitForElementIsVisible(By.XPath("//*[@id=\"list-item-with-action\"]/div[" + i + "]/div/div/div[10]/div/label[text()='Sat']"));

                WaitForLoad();

                var sunday = WaitForElementIsVisible(By.XPath("//*[@id=\"list-item-with-action\"]/div[" + i + "]/div/div/div[10]/div/label[text()='Sun']"));

                WaitForLoad();

                if (!monday.GetCssValue("border").Contains(colorGreen)
                    || !tuesday.GetCssValue("border").Contains(colorGreen)
                    || !wednesday.GetCssValue("border").Contains(colorGreen)
                    || !thursday.GetCssValue("border").Contains(colorGreen)
                    || !friday.GetCssValue("border").Contains(colorGreen)
                    || !saturday.GetCssValue("border").Contains(colorGreen)
                    || !sunday.GetCssValue("border").Contains(colorGreen)
                  )
                {
                    valuebool = false;
                    break;
                }
            }

            return valuebool;
        }

        public void AddDeliveryDays(bool monday = false, bool tuesday = false, bool wednesday = false, bool thursday = false, bool friday = false)
        {
            ClickToDeliveryBtn();

            // Select days for ALL SITES
            if (monday)
            {
                _monday = WaitForElementIsVisible(By.XPath(MONDAY));
                _monday.Click();
            }

            if (tuesday)
            {
                _tuesday = WaitForElementIsVisible(By.XPath(TUESDAY));
                _tuesday.Click();
                Thread.Sleep(1000);//too long to save
            }

            if (wednesday)
            {
                _wednesday = WaitForElementIsVisible(By.XPath(WEDNESDAY));
                _wednesday.Click();
            }

            if (thursday)
            {
                _thursday = WaitForElementIsVisible(By.XPath(THURSDAY));
                _thursday.Click();
            }

            if (friday)
            {
                _friday = WaitForElementIsVisible(By.XPath(FRIDAY));
                _friday.Click();
            }
        }

        public void ClickOnDeliveryComment()
        {
            _showCommentPanel = WaitForElementIsVisible(By.XPath(SHOW_COMMENT_PANEL));
            _showCommentPanel.Click();
            WaitForLoad();
        }

        public bool IsDeliveryCommentFilled(string comment)
        {
            ClickOnDeliveryComment();

            _commentPanel = WaitForElementIsVisible((By.Id(COMMENT_PANEL)));
            if (_commentPanel.Text == comment)
                return true;

            return false;
        }

        public bool AreDaysAdded()
        {
            bool valueBool = false;

            _monday = WaitForElementIsVisible(By.XPath(MONDAY));
            _tuesday = WaitForElementIsVisible(By.XPath(TUESDAY));
            _wednesday = WaitForElementIsVisible(By.XPath(WEDNESDAY));
            _thursday = WaitForElementIsVisible(By.XPath(THURSDAY));
            _friday = WaitForElementIsVisible(By.XPath(FRIDAY));
            _saturday = WaitForElementIsVisible(By.XPath(SATURDAY));
            _sunday = WaitForElementIsVisible(By.XPath(SUNDAY));

            if (_monday.Displayed & _tuesday.Displayed & _wednesday.Displayed &
                _thursday.Displayed & _friday.Displayed & _saturday.Displayed)
            {
                valueBool = true;
            }

            WaitForLoad();
            return valueBool;
        }


        // Contacts

        public void GoToContactTab()
        {
            _contactsTab = WaitForElementIsVisible(By.Id(CONTACTS_TAB));
            _contactsTab.Click();
            WaitForLoad();
        }

        public string AddContact(string contactEmail)
        {
            string nameGenerated = "Contact" + new Random().Next().ToString();

            GoToContactTab();

            // Unfold all / Fold all
            _contactFoldBtn = WaitForElementIsVisible(By.XPath(CONTACT_FOLD_BTN));
            _contactFoldBtn.Click();
            WaitForLoad();
            _contactFoldBtn.Click();

            // New concat modal
            _newContactBtn = WaitForElementIsVisible(By.XPath(NEW_CONTACT_BTN));
            _newContactBtn.Click();

            // Completed the contact form
            // Name
            _contactName = WaitForElementIsVisible(By.Id(CONTACT_NAME));
            _contactName.SetValue(ControlType.TextBox, nameGenerated);

            _contactMail = WaitForElementIsVisible(By.Id(CONTACT_MAIL));
            _contactMail.SetValue(ControlType.TextBox, contactEmail);

            // Site
            ComboBoxSelectById(new ComboBoxOptions("SelectedSites_ms", "MAD - MAD", false));


            // Validate
            _contactValidateBtn = WaitForElementIsVisible(By.XPath("//*[@id=\"modal-1\"]/div[2]/div/form/div[2]/button[2]"));

            _contactValidateBtn.Click();
            WaitForLoad();

            return nameGenerated;
        }

        public string GetFirstContact()
        {
            _fullNameContact = WaitForElementIsVisible(By.XPath(FULL_NAME_CONTACT));
            return _fullNameContact.Text.Split('-')[0];
        }

        public void SearchAndEditContact(string supplier, string site, string contactEmail)
        {
            var listContacts = _webDriver.FindElements(By.XPath(CONTACT_LIST));

            if (listContacts.Count == 0)
            {
                AddContact(contactEmail);
            }
            else
            {
                bool isFound = false;
                int i = 1;

                foreach (var contact in listContacts)
                {

                    if (contact.Text.Contains(supplier) && contact.Text.Contains(site))
                    {
                        Actions action = new Actions(_webDriver);
                        action.MoveToElement(contact).Perform();

                        var editContact = WaitForElementIsVisible(By.XPath(String.Format(EDIT_CONTACT, i)));
                        editContact.Click();
                        WaitForLoad();

                        EditContact(contactEmail);
                        isFound = true;
                        break;
                    }

                    i++;
                }

                if (!isFound)
                {
                    AddContact(contactEmail);
                }
            }
        }

        public void EditContact(string contactEmail)
        {
            _contactMail = WaitForElementIsVisible(By.Id(CONTACT_MAIL));

            if (_contactMail.GetAttribute("value").Equals(""))
            {
                _contactMail.SetValue(ControlType.TextBox, contactEmail);
            }
            _contactValidateBtn = WaitForElementIsVisible(By.XPath("//*[@id=\"modal-1\"]/div[2]/div/form/div[2]/button[2]"));

            _contactValidateBtn.Click();
            WaitForLoad();

        }

        public void EditContactForClaim(string contactEmail)
        {

            var contact = _webDriver.FindElement(By.XPath(FIRST_ITEM_CONTACT));

            Actions action = new Actions(_webDriver);
            action.MoveToElement(contact).Perform();

            var editContact = WaitForElementIsVisible(By.XPath(EDIT_CONTACT_FOR_CLAIM));
            editContact.Click();

            _contactMail = WaitForElementIsVisible(By.Id(CONTACT_MAIL));
            _contactMail.SetValue(ControlType.TextBox, contactEmail);

            _contactValidateBtn = WaitForElementIsVisible(By.XPath(CONTACT_VALIDATE_BTN_DEV));


            _contactValidateBtn.Click();
            WaitForLoad();

        }

        // Accounting

        public void ClickOnAccountingTab()
        {
            _acccountingTab = WaitForElementIsVisible(By.Id(ACCOUNTING_TAB));
            _acccountingTab.Click();
            WaitPageLoading();
            WaitForLoad();
        }

        public void ClickOnAccountingTiersDetailsTab()
        {
            _acccountingTab = WaitForElementIsVisible(By.Id(ACCOUNTING_TIERS_DETAILS_TAB));
            _acccountingTab.Click();
            WaitPageLoading();
            WaitForLoad();
        }

        public SupplierItem ClickOnItemsTab()
        {
            var _itemsTab = WaitForElementIsVisibleNew(By.Id("hrefTabContentArticles"));
            _itemsTab.Click();
            WaitPageLoading();
            WaitForLoad();
            return new SupplierItem(_webDriver, _testContext);
        }

        public bool IsAccountingPresent(string site)
        {
            try
            {
                _accounting = _webDriver.FindElement(By.XPath(string.Format(ACCOUNTING, site)));
                return true;
            }
            catch
            {
                return false;
            }
        }

        public void SearchAccounting(string site)
        {
            var editAccounting = WaitForElementIsVisible(By.XPath(string.Format(EDIT_ACCOUNTING, site)));
            editAccounting.Click();
            WaitPageLoading();
        }

        public void CreateNewAccounting(string thirdPart = null, string accountingId = null, string thirdPartDTI = null, string accountingIdDTI = null, string site = null)
        {
            _newAccountingBtn = WaitForElementIsVisible(By.XPath(NEW_ACCOUNTING));
            _newAccountingBtn.Click();
            WaitForLoad();

            if (thirdPart != null)
            {
                _thirdPart = WaitForElementIsVisible(By.Id(THIRD_PART));
                _thirdPart.SetValue(ControlType.TextBox, thirdPart);
                WaitForLoad();
            }

            if (accountingId != null)
            {
                _accountingId = WaitForElementIsVisible(By.Id(ACCOUNTING_ID));
                _accountingId.SetValue(ControlType.TextBox, accountingId);
                WaitForLoad();
            }

            if (thirdPartDTI != null)
            {
                _thirdPartDTI = WaitForElementIsVisible(By.Id(THIRD_PART_DTI));
                _thirdPartDTI.SetValue(ControlType.TextBox, thirdPartDTI);
                WaitForLoad();
            }

            if (accountingIdDTI != null)
            {
                _accountingIdDTI = WaitForElementIsVisible(By.Id(ACCOUNTING_ID_DTI));
                _accountingIdDTI.SetValue(ControlType.TextBox, accountingIdDTI);
                WaitForLoad();
            }

            // Site
            if (site != null)
            {
                ComboBoxSelectById(new ComboBoxOptions(SITE_SEARCH_ACCOUNTING, site, false));
            }
            else
            {
                _siteSearchAccounting = WaitForElementIsVisible(By.Id(SITE_SEARCH_ACCOUNTING));
                _siteSearchAccounting.Click();
                WaitForLoad();

                _checkAllSites = WaitForElementIsVisible(By.XPath(CHECK_ALL_SITES));
                _checkAllSites.Click();
                WaitForLoad();
            }

            //Try to create
            _accountingCreateBtn = WaitForElementIsVisible(By.Id(ACCOUNTING_CONFIRM_BTN));
            _accountingCreateBtn.Click();
            WaitForLoad();
        }

        public void EditAccounting(string thirdPart = null, string accountingId = null, string thirdPartDTI = null, string accountingIdDTI = null, string site = null)
        {
            if (thirdPart != null)
            {
                _thirdPart = WaitForElementIsVisible(By.Id(THIRD_PART));
                _thirdPart.SetValue(ControlType.TextBox, thirdPart);
                WaitForLoad();
            }

            if (accountingId != null)
            {
                _accountingId = WaitForElementIsVisible(By.Id(ACCOUNTING_ID));
                _accountingId.SetValue(ControlType.TextBox, accountingId);
                WaitForLoad();
            }

            if (thirdPartDTI != null)
            {
                _thirdPartDTI = WaitForElementIsVisible(By.Id(THIRD_PART_DTI));
                _thirdPartDTI.SetValue(ControlType.TextBox, thirdPartDTI);
                WaitForLoad();
            }

            if (accountingIdDTI != null)
            {
                _accountingIdDTI = WaitForElementIsVisible(By.Id(ACCOUNTING_ID_DTI));
                _accountingIdDTI.SetValue(ControlType.TextBox, accountingIdDTI);
                WaitForLoad();
            }

            if (site != null)
            {
                ComboBoxSelectById(new ComboBoxOptions(SITE_SEARCH_ACCOUNTING, site, false));
            }

            //Try to create
            _accountingCreateBtn = WaitForElementIsVisible(By.Id(ACCOUNTING_CONFIRM_BTN));
            _accountingCreateBtn.Click();
            WaitForLoad();
        }

        public void CreateNewAccountingTiersDetails(string tiersCode = null, string tiersName = null, string tiersAddressId = null, string tiersAddress = null, string tiersDomiciliation = null)
        {
            _newAccountingBtn = WaitForElementIsVisible(By.XPath(NEW_ACCOUNTING));
            _newAccountingBtn.Click();
            WaitForLoad();

            if (tiersCode != null)
            {
                _tiersCodeInput = WaitForElementIsVisible(By.Id(TIERS_CODE));
                _tiersCodeInput.SetValue(ControlType.TextBox, tiersCode);
                WaitForLoad();
            }

            if (tiersName != null)
            {
                _tiersNameInput = WaitForElementIsVisible(By.Id(TIERS_NAME));
                _tiersNameInput.SetValue(ControlType.TextBox, tiersName);
                WaitForLoad();
            }

            if (tiersAddressId != null)
            {
                _tiersAddressIdInput = WaitForElementIsVisible(By.Id(TIERS_ADDRESS_IDENTIFIER));
                _tiersAddressIdInput.SetValue(ControlType.TextBox, tiersAddressId);
                WaitForLoad();
            }

            if (tiersAddress != null)
            {
                _tiersAddressInput = WaitForElementIsVisible(By.Id(TIERS_ADDRESS));
                _tiersAddressInput.SetValue(ControlType.TextBox, tiersAddress);
                WaitForLoad();
            }

            if (tiersDomiciliation != null)
            {
                _tiersDomiciliationInput = WaitForElementIsVisible(By.Id(TIERS_DOMICILIATION_IDENTIFIER));
                _tiersDomiciliationInput.SetValue(ControlType.TextBox, tiersDomiciliation);
                WaitForLoad();
            }
            //Try to create
            _accountingCreateBtn = WaitForElementIsVisible(By.Id(ACCOUNTING_CONFIRM_BTN));
            _accountingCreateBtn.Click();
            WaitPageLoading();
        }

        public void EditAccountingTiersDetails(string tiersCodeToModify, string tiersCode = null, string tiersName = null, string tiersAddressId = null, string tiersAddress = null, string tiersDomiciliation = null)
        {
            var editAccounting = WaitForElementIsVisible(By.XPath(string.Format("//td[contains(text(), '{0}')]/..//span[contains(@class,'pencil')]", tiersCodeToModify)));
            editAccounting.Click();

            WaitForLoad();

            if (tiersCode != null)
            {
                _tiersCodeInput = WaitForElementIsVisible(By.Id(TIERS_CODE));
                _tiersCodeInput.SetValue(ControlType.TextBox, tiersCode);
                WaitForLoad();
            }

            if (tiersName != null)
            {
                _tiersNameInput = WaitForElementIsVisible(By.Id(TIERS_NAME));
                _tiersNameInput.SetValue(ControlType.TextBox, tiersName);
                WaitForLoad();
            }

            if (tiersAddressId != null)
            {
                _tiersAddressIdInput = WaitForElementIsVisible(By.Id(TIERS_ADDRESS_IDENTIFIER));
                _tiersAddressIdInput.SetValue(ControlType.TextBox, tiersAddressId);
                WaitForLoad();
            }

            if (tiersAddress != null)
            {
                _tiersAddressInput = WaitForElementIsVisible(By.Id(TIERS_ADDRESS));
                _tiersAddressInput.SetValue(ControlType.TextBox, tiersAddress);
                WaitForLoad();
            }

            if (tiersDomiciliation != null)
            {
                _tiersDomiciliationInput = WaitForElementIsVisible(By.Id(TIERS_DOMICILIATION_IDENTIFIER));
                _tiersDomiciliationInput.SetValue(ControlType.TextBox, tiersDomiciliation);
                WaitForLoad();
            }

            //Try to create
            _accountingCreateBtn = WaitForElementIsVisible(By.Id(ACCOUNTING_CONFIRM_BTN));
            _accountingCreateBtn.Click();
            WaitPageLoading();
        }


        // Certifications

        public void GetCertificationTab()
        {
            _certificationTab = WaitForElementIsVisible(By.Id(CERTIFICATION_TAB));
            _certificationTab.Click();
            WaitForLoad();
        }

        public void ClickOnEDITab()
        {
            var ediTab = WaitForElementIsVisible(By.Id(EDI_TAB));
            ediTab.Click();
            WaitForLoad();
        }
        public void FillEdiData(string idedi, string repo)
        {
            var edi_id = WaitForElementIsVisible(By.Id(ID_EDI));
            edi_id.Clear();
            edi_id.SendKeys(idedi);
            WaitForLoad();
            var repository = WaitForElementIsVisible(By.Id(REPOSITORY));
            repository.Clear();
            repository.SendKeys(repo);
            WaitForLoad();
        }
        public string GetIdEdi()
        {
            var idEdi = WaitForElementIsVisible(By.Id(ID_EDI));
            var value = idEdi.GetAttribute("value");
            return value;
        }
        public string GetRepository()
        {
            var repository = WaitForElementIsVisible(By.Id(REPOSITORY));
            var value = repository.GetAttribute("value");
            return value;
        }

        public bool AddNewCertification(string certification)
        {
            try
            {
                _newCertification = WaitForElementIsVisible(By.XPath(NEW_CERTIFICATION));
                _newCertification.Click();
                WaitForLoad();

                _certificationId = WaitForElementIsVisible(By.Id(CERTIFICATION_ID));
                _certificationId.SetValue(ControlType.DropDownList, certification);

                _certificationExpirationDate = WaitForElementIsVisible(By.Id(CERTIFICATION_EXPIRATION_DATE));
                _certificationExpirationDate.SetValue(ControlType.DateTime, DateUtils.Now.AddMonths(6));

                _addCertification = WaitForElementIsVisible(By.XPath(ADD_CERTIFICATION));
                _addCertification.Click();
                WaitForLoad();
                return true;
            }
            catch
            {
                return false;
            }
        }

        public bool DelFirstCertification()
        {
            try
            {
                _delCertification = WaitForElementIsVisible(By.XPath(DELETE_CERTIFICATION));
                _delCertification.Click();
                WaitForLoad();
                return true;
            }
            catch
            {
                return false;
            }
        }

        public void DelCertification(string certification)
        {
            if (isElementVisible(By.XPath(string.Format(DELETE_CERTIFICATION_TARGET, certification))))
            {
                var del = WaitForElementIsVisible(By.XPath(string.Format(DELETE_CERTIFICATION_TARGET, certification)));
                del.Click();
                WaitForLoad();
                var delConfirm = WaitForElementIsVisible(By.Id(DELETE_CERTIFICATION_TARGET_CONFIRM));
                delConfirm.Click();
            }
        }

        public string GetSanitaryNumber()
        {
            return WaitForElementIsVisible(By.Id("SanitaryNumber")).GetAttribute("value");
        }

        public string GetSIRET()
        {
            return WaitForElementIsVisible(By.Id("IdExternal")).GetAttribute("value");
        }

        public string GetPaymentTerm()
        {
            SelectElement se = new SelectElement(WaitForElementIsVisible(By.Id("PaymentTypeId")));
            return se.SelectedOption.Text;
        }

        public string GetEndOfContract()
        {
            return WaitForElementIsVisible(By.Id("date-picker-supplier-end-contract")).GetAttribute("value");
        }

        public string GetComment()
        {
            return WaitForElementIsVisible(By.Id("Comment")).GetAttribute("value");
        }

        public string GetAddress1()
        {
            return WaitForElementIsVisible(By.Id("Address1")).GetAttribute("value");
        }

        public string GetAddress2()
        {
            return WaitForElementIsVisible(By.Id("Address2")).GetAttribute("value");
        }

        public string GetZipCode()
        {
            return WaitForElementIsVisible(By.Id("ZipCode")).GetAttribute("value");
        }

        public string GetCity()
        {
            return WaitForElementIsVisible(By.Id("City")).GetAttribute("value");
        }

        public string GetCountry()
        {
            SelectElement se2 = new SelectElement(WaitForElementIsVisible(By.Id("drop-down-countries")));
            return se2.SelectedOption.Text;
        }

        public string GetCurrency()
        {
            SelectElement se3 = new SelectElement(WaitForElementIsVisible(By.Id("drop-down-currencies")));
            return se3.SelectedOption.Text;
        }

        public string GetLevelOfRisk()
        {
            return WaitForElementIsVisible(By.Id("LevelOfRisk")).GetAttribute("value");
        }

        public string GetIntentedUse()
        {
            return WaitForElementIsVisible(By.Id("IntendedUse")).GetAttribute("value");
        }

        public string GetRiskassessmentscore()
        {
            string valeur = WaitForElementIsVisible(By.XPath("//*/label[text()='Risk assessment score']/../div/p")).Text.Trim();
            if (valeur.StartsWith("Low"))
            {
                return "Low";
            }
            else if (valeur.StartsWith("Medium"))
            {
                // non testé
                return "Medium";
            }
            else
            {
                // non testé
                return "Not defined";
            }
        }

        public string GetQuantityVolume()
        {
            return WaitForElementIsVisible(By.Id("QuantityPerVolume")).GetAttribute("value");
        }

        public string GetLastAuditDate()
        {
            WaitForLoad();
            var test = WaitForElementIsVisible(By.Id("date-picker-supplier-last-audit"));
            return WaitForElementIsVisible(By.Id("date-picker-supplier-last-audit")).GetAttribute("value");
        }
        public string GetNextAuditDate()
        {
            return WaitForElementIsVisible(By.Id("date-picker-supplier-next-audit")).GetAttribute("value");
        }

        public SupplierDeliveriesTab ClickOnDeliveriesTab()
        {

            var tab = WaitForElementIsVisible(By.Id("hrefTabContentDelivery"));
            tab.Click();
            WaitForLoad();
            return new SupplierDeliveriesTab(_webDriver, _testContext);
        }
        public string SelectFirstItem()
        {
            var itemName = WaitForElementIsVisible(By.XPath(FIRST_ITEM_NAME));
            var value = itemName.Text;
            itemName.Click();
            WaitForLoad();
            return value;
        }
        public SupplierGeneralInfoTab SelectItem()
        {
            var itemName = WaitForElementIsVisible(By.XPath(FIRST_ITEM_NAME));
            itemName.Click();
            WaitForLoad();
            return new SupplierGeneralInfoTab(_webDriver, _testContext);

        }
        public bool AtleastOnePackagingIsNotPurchasable()
        {
            var packagingsPurchasableState = _webDriver.FindElements(By.XPath(PACKAGINGS_PURCHASABLE));

            foreach (var element in packagingsPurchasableState)
            {
                if (element.Selected is false)
                    return true;
            }
            return false;

        }
        public string GetFirstItemGroupName()
        {
            WaitForLoad();
            var firstItem = WaitForElementIsVisible(By.XPath(FIRST_ITEM));
            firstItem.Click();
            var selectedGrp = WaitForElementIsVisible(By.Id(SELECTED_GRP));
            return new SelectElement(selectedGrp).SelectedOption.Text;
        }
        public string GetFirstItemSite()
        {
            var firstItem = WaitForElementIsVisible(By.XPath(FIRST_ITEM));
            firstItem.Click();
            var selectedGrp = WaitForElementIsVisible(By.XPath(ITEM_PACKAGING_SITE));
            return selectedGrp.Text;
        }

        public List<string> GetPackagingItemSite()
        {
            WaitForLoad();
            var firstItem = WaitForElementIsVisibleNew(By.XPath(FIRST_ITEM));
            firstItem.Click();
            WaitForLoad();
            var sitesPackaging = _webDriver.FindElements(By.XPath("//*/tbody/tr[*]/td[1]/span[@class='readonly-input']"));
            return sitesPackaging.Select(p => p.Text).ToList<string>();
        }

        public string GetFirstItemSubGrp()
        {
            var firstItem = WaitForElementIsVisible(By.XPath(FIRST_ITEM));
            firstItem.Click();
            var selectedSubGrp = WaitForElementIsVisible(By.XPath(SELECTED_SUBGRP));
            return selectedSubGrp.Text;
        }
        public bool VerifyFilterSearch(string itemname)
        {
            var table = _webDriver.FindElements(By.XPath(ITEMS));
            var numberOfItems = table.Count;
            if (table.FirstOrDefault().Text == itemname)
            {
                return true;
            }
            return false;
        }
        public void ModifyExpirationDateCertification(string date)
        {
            var expiryDate = WaitForElementIsVisible(By.Id(EXPIRATION_DATE));
            expiryDate.Clear();
            expiryDate.SendKeys(date);
            Thread.Sleep(1000);
        }
        public string GetFirstExpirationDate()
        {
            var expiryDate = WaitForElementIsVisible(By.Id(EXPIRATION_DATE));
            return expiryDate.GetAttribute("value");
        }

        public void DeleteAccounting(string site)
        {
            WaitForLoad();
            if (isElementVisible(By.XPath("//*/td[@title='" + site + "']/../td[@class='actions']/a[2]")))
            {
                var linkDel = WaitForElementIsVisible(By.XPath("//*/td[@title='" + site + "']/../td[@class='actions']/a[2]"));
                linkDel.Click();
                WaitForLoad();
                var confirmDel = WaitForElementIsVisible(By.Id("dataConfirmOK"));
                confirmDel.Click();
                WaitForLoad();
            }
        }

        public void DeleteAccountingTiersDetails(string tiersCode)
        {
            WaitForLoad();
            // bug nouvelle ligne à la place d'édition de la ligne
            var compteur = 10;
            while (compteur > 0 && isElementVisible(By.XPath(string.Format(ACCOUNTING_TIERS_DETAILS_LINE, tiersCode))))
            {
                var linkDel = WaitForElementIsVisible(By.XPath(string.Format("//td[contains(text(), '{0}')]/..//span[contains(@class,'trash')]", tiersCode)));
                linkDel.Click();

                WaitForLoad();
                var confirmDel = WaitForElementIsVisible(By.Id("dataConfirmOK"));
                confirmDel.Click();
                WaitPageLoading();
                WaitForLoad();
                compteur--;
            }
        }

        public int AccountingCount()
        {
            int count = _webDriver.FindElements(By.XPath("//*[@id=\"dat-container\"]/table/tbody/tr[*]")).Count;
            WaitForLoad();
            return count - 1;
        }

        public bool VerifyResetFilter()
        {
            var searchField = WaitForElementIsVisible(By.Id(SEARCH_FIELD));
            if (String.IsNullOrEmpty(searchField.Text))
            {
                return true;
            }
            return false;
        }

        public string getSupplierName()
        {
            var supplierName = WaitForElementIsVisible(By.XPath("//*/h1"));
            return supplierName.Text.Substring("SUPPLIER : ".Length);
        }

        public bool IsItemsPage()
        {
            // bouton "New Item" présent, donc on est bien dans l'onglet Items
            return isElementVisible(By.XPath("//*/a[text()='New item']"));
        }

        public void SetSite(string value)
        {
            ComboBoxSelectById(new ComboBoxOptions("SelectedSites_ms", value, false));
        }
        public void FastCopyPrice(string copyFromsite, string copyToSite)
        {
            _fast_copy_price = WaitForElementIsVisible(By.XPath(FAST_COPY_PRICE));
            _fast_copy_price.Click();

            var setSite = WaitForElementExists(By.XPath(SET_SITE));
            setSite.SetValue(ControlType.DropDownList, copyFromsite);
            WaitForLoad();

            var filterByName = WaitForElementExists(By.XPath(FILTER_BY_NAME));
            filterByName.SetValue(ControlType.TextBox, copyToSite);
            WaitForLoad();
            WaitPageLoading();
            var checkSite = WaitForElementIsVisible(By.XPath("/html/body/div[4]/div/div/div/div/form/div[2]/div[1]/div[3]/div[1]/a[1]"));
            //checkSite.Click();
            checkSite.SetValue(ControlType.CheckBox, true);
            WaitPageLoading();
            _copy_btn = WaitForElementIsVisible(By.XPath(COPY_BTN));
            _copy_btn.Click();
            WaitPageLoading();
            var okBtn = WaitForElementIsVisible(By.XPath(OK_BTN));
            okBtn.Click();

        }


        public void SearchKeywords(string site)
        {
            var searchkeywrod = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div[2]/div[2]/div/div/div/form/div[2]/span/input"));
            searchkeywrod.SetValue(ControlType.TextBox, site);
        }
        public bool CountKEYWORD()
        {
            var list = _webDriver.FindElements(By.XPath("//*[@id='div-keywords-details']/table/tbody/tr"));

            if (list.Count > 0)
            {
                return true;
            }

            return false;
        }


        public void ClickKeyword()
        {
            var keyword = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div[2]/ul/li[3]/a"));
            //checkSite.Click();
            keyword.Click();
        }
        public List<string> GetListGroupName()
        {
            var groups = _webDriver.FindElements(By.XPath(ITEM_GROUP));

            var listGroup = new List<string>();

            foreach (var item in groups)
            {
                listGroup.Add(item.Text);
            }
            return listGroup;
        }
        public List<string> GetSubGroupFirstItem()
        {
            var groups = _webDriver.FindElements(By.XPath(ITEM_GROUP));

            var listGroup = new List<string>();

            foreach (var item in groups)
            {
                listGroup.Add(item.Text);
            }
            return listGroup;
        }
        public bool DetailsContactsIsVisible()
        {

            var _details_contact = WaitForElementExists(By.XPath(DETAILS_CONTACTS));
            return _details_contact.Displayed;

        }
        public bool RowContactsIsVisible()
        {
            return isElementVisible(By.XPath(ROW_CONTACTS));
        }
        public void ClickFoldAllButton()
        {
            var foldAllButton = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div/div/div/div[1]/a[contains(text(),'Fold All')]"));
            foldAllButton.Click();
            WaitPageLoading();
        }
        public void ClickUnfoldAllButton()
        {
            var unfoldAllButton = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div/div/div/div[1]/a[contains(text(),'Unfold All')]"));
            unfoldAllButton.Click();
            WaitPageLoading();
        }
        //Storage Tab
        public void GoToStorageTab()
        {
            _storageTab = WaitForElementIsVisible(By.Id(STORAGE_TAB));
            _storageTab.Click();
            WaitPageLoading();
        }

        public void FilterBySitesOnStorageTab(string sites)
        {
            ComboBoxSelectById(new ComboBoxOptions("SelectedSites_ms", sites, false));
            WaitForLoad();
        }
        public bool AreFiltredSitesStorageTabCorrect(string expectedResult)
        {
            var filtredItems = _webDriver.FindElements(By.XPath("//*[@id=\"table-itemDetailsStorage\"]/tbody/tr[*]/td[3]"));
            foreach (var item in filtredItems)
            {
                if (!item.Text.Contains(expectedResult))
                {
                    return false;
                }

            }
            return true;
        }
        public void ClickUnSelectAll()
        {
            //click UnSelectAll On Tab Storage
            _unselectAll_btn = WaitForElementIsVisible(By.XPath(UNSELECT_ALL_BTN));
            _unselectAll_btn.Click();
            WaitPageLoading();
            WaitLoading();
        }
        public void UnfoldAllForDelete()
        {
            _unfoldall = WaitForElementIsVisible(By.XPath(UNFOLDALL));
            _unfoldall.Click();
            WaitPageLoading();

            _fleche = WaitForElementIsVisible(By.XPath(FLECHEFORDELTE));
            _fleche.Click();
            WaitPageLoading();
        }

        public void DeleteButton()
        {
            Actions actions = new Actions(_webDriver);
            var line = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div/div/div/div[2]/div/div[2]/div/div/div[2]/div/table/tbody/tr[2]"));
            actions.MoveToElement(line).Perform();

            _deletebutton = WaitForElementIsVisible(By.XPath(DELETEBUTTON));
            _deletebutton.Click();
            WaitPageLoading();


        }

        public void ClickSelectAll()
        {
            //click SelectAll On Tab Storage
            _selectAll_btn = WaitForElementIsVisible(By.XPath(SELECTALL_BTN));
            _selectAll_btn.Click();
            WaitPageLoading();
            WaitLoading();
        }
        public bool AreAllItemsSelected()
        {
            var items = _webDriver.FindElements(By.XPath("/html/body/div[3]/div/div/div/div/div[2]/table/tbody/tr[*]/td[1]/input[1]"));
            if (items.Count == 0)
            {
                return false;
            }

            foreach (var item in items)
            {
                if (!item.Selected)
                {
                    return false;
                }

            }
            return true;
        }
        public bool AreAllItemsUnSelected()
        {
            var items = _webDriver.FindElements(By.XPath("/html/body/div[3]/div/div/div/div/div[2]/table/tbody/tr[*]/td[1]/input[1]"));
            if (items.Count == 0)
            {
                return false;
            }

            foreach (var item in items)
            {
                if (item.Selected)
                {
                    return false;
                }

            }
            return true;
        }
        public bool IsExported()
        {
            IWebElement _exported = WaitForElementExists(By.Id(HEADER_PRINT_BUTTON));
            WaitForLoad();
            _exported.Click();
            IWebElement _exportedFirstLine = _webDriver.FindElement(By.XPath("//table[contains(@class,'print-table')]/tbody/tr[2]/td[2]"));
            return _exportedFirstLine.Text.Contains("PurchaseBook");
        }

        public int certificationsCount()
        {
            var certifications = _webDriver.FindElements(By.XPath("/html/body/div[3]/div/div/div/div/div[2]/div/div[*]/div[*]"));


            return certifications.Count();
        }
        public bool ClickOnFirstItems()
        {
            if (isElementVisible(By.XPath("/html/body/div[3]/div/div/div/div/div[2]/div/div[2]/div/div[1]/div[1]/div")))
            {
                var _itemsTab = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div/div/div/div[2]/div/div[2]/div/div[1]/div[1]/div"));
                _itemsTab.Click();
                return true;
            }
            return false;
        }

        public int ItemsCount()
        {
            if (IsDev())
            {
                if (isElementVisible(By.XPath("/html/body/div[*]/div/div/div/div/div[*]/div/div[*]/div/div[*]/div[*]/div[@class='row']")))
                {
                    var _itemsTab = _webDriver.FindElements(By.XPath("/html/body/div[*]/div/div/div/div/div[*]/div/div[*]/div/div[*]/div[*]/div[@class='row']"));
                    return _itemsTab.Count;
                }
            }
            else
            {

                if (isElementVisible(By.XPath("/html/body/div[3]/div/div/div/div/div[2]/div/div[2]/div/div[*]/div[1]/div/div[2]/table/tbody/tr")))
                {
                    var _itemsTab = _webDriver.FindElements(By.XPath("/html/body/div[3]/div/div/div/div/div[2]/div/div[2]/div/div[*]/div[1]/div/div[2]/table/tbody/tr"));
                    return _itemsTab.Count;
                }
            }

            return (int)decimal.Zero;
        }

        public void AddPackaging(string site = "ACE", string packaging = "KG", string storageQuantity = "5", string storageUnit = "KG", string ProdQuantity = "58", string supplier = "SupplierTestDeleteItem")
        {
            //click create new packaging
            var addPackagingButton = WaitForElementIsVisible(By.Id("btn-add-item-detail"));
            addPackagingButton.Click();

            //adding site
            var siteSelect = WaitForElementIsVisible(By.Id("SelectedSites_ms"));
            siteSelect.Click();
            var searchSite = WaitForElementIsVisible(By.XPath("/html/body/div[12]/div/div/label/input"));
            searchSite.SetValue(ControlType.TextBox, site);
            var UncheckAllSites = WaitForElementIsVisible(By.XPath("/html/body/div[12]/div/ul/li[2]/a/span[2]"));
            UncheckAllSites.Click();
            var selectSite = WaitForElementIsVisible(By.XPath("/html/body/div[12]/ul/li[2]/label/span"));
            selectSite.Click();
            siteSelect.Click();

            //adding packaging
            var packagingDropDown = WaitForElementIsVisible(By.Id("PackagingUnitListVM_ItemUnitPackagingTypeId"));
            packagingDropDown.Click();
            packagingDropDown.SetValue(ControlType.DropDownList, packaging);

            //adding storage quantity
            var storageQuantityInput = WaitForElementIsVisible(By.Id("NewPackaging_ItemUnitStorageTypeValue"));
            storageQuantityInput.SetValue(ControlType.TextBox, storageQuantity);

            //adding storage Unit
            var storageUnitDropDown = WaitForElementIsVisible(By.Id("PackagingUnitListVM_ItemUnitStorageTypeId"));
            storageUnitDropDown.SetValue(ControlType.DropDownList, storageUnit);

            //adding production quantity
            var productionQuantityInput = WaitForElementIsVisible(By.Id("NewPackaging_ItemUnitProductionQuantity"));
            productionQuantityInput.SetValue(ControlType.TextBox, ProdQuantity);

            //adding supplier
            var supplierDropDown = WaitForElementIsVisible(By.Id("NewPackaging_SupplierId"));
            supplierDropDown.Click();
            supplierDropDown.SetValue(ControlType.TextBox, supplier);

            //click create
            if (IsDev())
            {
                var createButton = WaitForElementIsVisible(By.Id("btn-create-new-packaging"));
                createButton.Click();
            }
            else
            {
                var createButton = WaitForElementIsVisible(By.XPath("//*/button[text()='Create']"));
                createButton.Click();
            }
        }

        public void DeactivateItems(string site = "AGP")
        {
            var deactivateButton = WaitForElementIsVisible(By.Id("deactivate-items-btn"));
            deactivateButton.Click();

            var sites = WaitForElementIsVisible(By.Id("SelectedSitesToUpdate_ms"));
            sites.Click();
            var siteBox = WaitForElementIsVisible(By.XPath("/html/body/div[14]/div/div/label/input"));

            siteBox.SetValue(ControlType.TextBox, site);
            siteBox.Click();
            sites.Click();

            var deleteButton = WaitForElementIsVisible(By.Id("last"));
            deleteButton.Click();

            var saveButton = WaitForElementIsVisible(By.XPath("/html/body/div[4]/div/div/div/div[3]/button[2]"));
            saveButton.Click();

        }

        public void BTN_SetUnpurchasableItems()
        {
            _setunpurchaisingitem = WaitForElementIsVisible(By.XPath(SETUNPURCHAISINGITEM));
            _setunpurchaisingitem.Click();
            WaitForLoad();

        }
        public void Set_Site_SetUnpurchasableItems(string site)
        {
            ComboBoxSelectById(new ComboBoxOptions("SelectedSitesToUpdate_ms", site, false));
            WaitForLoad();
        }

        public void Click_SetUnpurchasableItems()
        {
            var clickSetUnpurchasableItems = WaitForElementIsVisible(By.XPath("//*[@id=\"last\"]"));
            clickSetUnpurchasableItems.Click();
            WaitForLoad();
        }
        public void Click_OK_SetUnpourchasableItems()
        {
            var clickOKSetUnpurchasableItems = WaitForElementIsVisible(By.XPath("//*[@id=\"DeactivateItemsReportForm\"]/div[2]/button"));
            clickOKSetUnpurchasableItems.Click();
            WaitForLoad();
        }

        public IEnumerable<string> GetItemNames()
        {
            var itemNames = _webDriver.FindElements(By.XPath("//div[@class=\"panel panel-default\"]/div/div/div[2]/table//td[2]"));
            return itemNames.Select(i => i.Text);
        }

        public void SetSiteForFastCopy()
        {
            _sitefastcopy = WaitForElementIsVisible(By.Id(SITE_FC));
            _sitefastcopy.Click();
            WaitForLoad();
        }
        public ItemGeneralInformationPage ClickFirstItemName()
        {

            _itemName = WaitForElementExists(By.XPath(ITEM_NAME));
            _itemName.Click();

            return new ItemGeneralInformationPage(_webDriver, _testContext);

        }


        public void ClickOKforCopyPopUp()
        {
            _okCopy = WaitForElementIsVisible(By.XPath(OK_COPY));
            _okCopy.Click();
            WaitForLoad();
        }

        public void CopyAllPricesFromSIte(string site)
        {
            _copyAllPricesFrom.SetValue(ControlType.DropDownList, site);

            WaitForLoad();


        }

        public string CheckForSupplierExist(string supplierType)
        {
            var supplierTypes = WaitForElementExists(By.Id("drop-down-suppliertypes"));
            supplierTypes.SetValue(PageBase.ControlType.DropDownList, supplierType);
            var supplierID = WaitForElementExists(By.XPath("//*/input[@name='Id' and @type!='hidden']"));
            var supplierNumber = supplierID.GetAttribute("value");
            return supplierNumber;

        }
        public IReadOnlyList<IWebElement> GetContactList()
        {
            return _webDriver.FindElements(By.XPath(CONTACT_ELEMENT));
        }
        public void DeleteAllContacts()
        {
            WaitForLoad();

            var contactElements = _webDriver.FindElements(By.XPath(CONTACT_ELEMENT));

            while (contactElements.Count > 0)
            {
                var contactElement = contactElements[0];
                Actions action = new Actions(_webDriver);
                action.MoveToElement(contactElement).Perform();
                var deleteButton = WaitForElementIsVisible(By.XPath("//*[@id=\"purchasing-supplier-contact-1\"]"));
                deleteButton.Click();
                var moveToDelete = WaitForElementIsVisible(By.XPath("//*[@id=\"purchasing-supplier-delete-contact-1\"]/span"));
                moveToDelete.Click();
                var confirmButton = WaitForElementExists(By.XPath(CONFIRM_BUTTON));
                confirmButton.Click();
                WaitForLoad();
                contactElements = _webDriver.FindElements(By.XPath(CONTACT_ELEMENT));
            }
        }
        public SupplierContactTab ClickOnNewContactButton()
        {
            var newContactButton = WaitForElementIsVisible(By.XPath("//*[@id=\"tabContentDetails\"]/div/div[1]/div/a[2]"));
            newContactButton.Click();
            WaitForLoad();
            return new SupplierContactTab(_webDriver, _testContext);
        }
        public void FillIBANField(string iban)
        {
            WaitPageLoading();
            // Localiser le champ IBAN et le remplir
            var ibanField = _webDriver.FindElement(By.XPath(IBAN_FIELD));
            ibanField.Clear();
            ibanField.SendKeys(iban);
            WaitForLoad();
            WaitPageLoading();

        }
        public string GetSavedIBAN()
        {
            WaitPageLoading();
            //récupérer la valeur
            var ibanField = _webDriver.FindElement(By.XPath(IBAN_FIELD));
            return ibanField.GetAttribute("value");
        }

        public void ClickOnFastCopyPrice()
        {

            var fastCopyPrice = WaitForElementIsVisible(By.Id("FastCopyPriceBtn"));
            fastCopyPrice.Click();
        }
        public enum FilterTypePopup
        {
            Search,
            Site
        }



        public void FilterPopup(FilterTypePopup filterTypePopup, object value)
        {

            switch (filterTypePopup)
            {
                case FilterTypePopup.Search:
                    _searchItemByName = WaitForElementIsVisible(By.Id("siteSearch"));
                    _searchItemByName.SetValue(ControlType.TextBox, value);
                    WaitPageLoading();
                    break;
                case FilterTypePopup.Site:
                    var cbSortBy = WaitForElementIsVisible(By.Id("siteSource"));
                    var selectSortBy = new SelectElement(cbSortBy);
                    selectSortBy.SelectByText((string)value);
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterTypePopup), filterTypePopup, null);

            }

            WaitPageLoading();
        }
        public void SelectAll()
        {

            var selectAll = WaitForElementIsVisible(By.Id(SELECTALL));
            selectAll.Click();
        }
        public void CopyPriceFromSupplier()
        {
            SelectAll();
            WaitForLoad();
            var copy = WaitForElementIsVisible(By.XPath(COPY));
            copy.Click();
            WaitPageLoading();
            WaitPageLoading();


            var ok = WaitForElementIsVisible(By.XPath(OK));
            ok.Click();
            WaitForLoad();
        }

        public string AddContactWithoutSite(string contactEmail)
        {
            string nameGenerated = "Contact" + new Random().Next().ToString();

            GoToContactTab();

            // Unfold all / Fold all
            _contactFoldBtn = WaitForElementIsVisible(By.XPath(CONTACT_FOLD_BTN));
            _contactFoldBtn.Click();
            WaitForLoad();
            _contactFoldBtn.Click();

            // New concat modal
            _newContactBtn = WaitForElementIsVisible(By.XPath(NEW_CONTACT_BTN));
            _newContactBtn.Click();

            // Completed the contact form
            // Name
            _contactName = WaitForElementIsVisible(By.Id(CONTACT_NAME));
            _contactName.SetValue(ControlType.TextBox, nameGenerated);

            _contactMail = WaitForElementIsVisible(By.Id(CONTACT_MAIL));
            _contactMail.SetValue(ControlType.TextBox, contactEmail);


            // Validate
            _contactValidateBtn = WaitForElementIsVisible(By.XPath("//*[@id=\"modal-1\"]/div[2]/div/form/div[2]/button[2]"));

            _contactValidateBtn.Click();
            WaitForLoad();

            return nameGenerated;
        }
        public SupplierContactTab ClickOnContactTab()
        {
            var ContactTab = WaitForElementIsVisible(By.Id(CONTACTS_TAB));
            ContactTab.Click();
            WaitLoading();
            return new SupplierContactTab(_webDriver, _testContext);
        }

        public List<string> GetCheckedKnownSupplier()
        {
            System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> checkedElements = _webDriver.FindElements(By.XPath("/html/body/div[11]/ul/li/label/input[@checked]"));

            return checkedElements
                .Select(element =>
                {
                    string spanText = element.FindElement(By.XPath("..")).FindElement(By.TagName("span")).Text;
                    int lastSpaceIndex = spanText.LastIndexOf(' ');
                    return lastSpaceIndex >= 0 ? spanText.Substring(lastSpaceIndex + 1) : spanText;
                })
                .ToList();
        }

        public string GetItemGroupNameSelected()
        {
            WaitForLoad();
            var selectedGrp = WaitForElementIsVisible(By.Id(SELECTED_GRP));
            return new SelectElement(selectedGrp).SelectedOption.Text;
        }
        public string GetName()
        {
            var nameField = WaitForElementIsVisible(By.XPath("//*[@id=\"supplier-items-details-1 \"]/td[2]"));
            return nameField.Text;
        }
        public string GetItemGroup()
        {
            var nameField = WaitForElementIsVisible(By.XPath("//*[@id=\"supplier-items-details-1 \"]/td[3]"));
            return nameField.Text;
        }
        public void UnfoldAll()
        {
            var unfoldAllButton = WaitForElementIsVisibleNew(By.Id("unfoldBtn"));
            unfoldAllButton.Click();
            WaitPageLoading();
        }
        public List<string> GetSiteItem()
        {
            WaitForLoad();
            var sitesPackagingItem = _webDriver.FindElements(By.XPath("/html/body/div[3]/div/div/div/div/div[2]/div/div[2]/div/div/div[2]/div/table/tbody/tr[*]/td[1]"));
            return sitesPackagingItem.Select(p => p.Text).ToList<string>();
        }
        public bool VerifySiteAfterFilterExist(string siteACE, List<string> sitesAfterFilter)
        {
            foreach (var site in sitesAfterFilter)
            {
                if (site != siteACE)
                {
                    return false;
                }
            }
            return true;
        }
        public void SelectCheckBoxes(int numberOfCheckBoxes)
        {
            var checkboxes = _webDriver.FindElements(By.XPath(SELECT_ALL_CHECKBOXES_COPYPRICE));
            for (int i = 0; i < numberOfCheckBoxes; i++)
            {
                if (checkboxes.Count > i)
                {
                    var checkbox = checkboxes[i];
                    if (!checkbox.Selected)
                    {
                        checkbox.Click();
                    }
                }
                else
                {
                    break;
                }
            }
            WaitPageLoading();
        }


        public void CheckAutoClose()
        {
            WaitPageLoading();
            var autoClose = _webDriver.FindElement(By.Id(AUTO_CLOSE));
            autoClose.Click();
            WaitPageLoading();
            WaitPageLoading();
        }
        public bool GetCheckAutoCloseStatus()
        {
            WaitPageLoading();
            var autoClose = _webDriver.FindElement(By.Id(AUTO_CLOSE));
            return autoClose.Selected;
        }
        public void checkReverseVAT()
        {
            var reverse = WaitForElementExists(By.XPath(REVERSE_VAT));
            reverse.Click();
            WaitForLoad();
        }

        public bool IsReverseVatChecked()
        {
        
            var reverse = WaitForElementExists(By.XPath(REVERSE_VAT));
             string cheked = reverse.GetAttribute("value");
            return (cheked == "True"); 
            
        }
        public string GetStartTime()
        {
            var timeLimit = WaitForElementIsVisible(By.XPath(TIME_LIMIT));
            return timeLimit.GetAttribute("value");
        }
        public void EditStartTime(string updateTime)
        {
            var timeLimit = WaitForElementIsVisible(By.XPath(TIME_LIMIT));
            timeLimit.Clear();
            timeLimit.SendKeys(updateTime);
            WaitLoading();
        }
        public string GetMinAmount()
        {

            _firstMinAmount = WaitForElementIsVisible(By.Id(FIRST_MIN_AMOUNT));
           
            return _firstMinAmount.GetAttribute("value");
        
        }
        public void EditMinAmount(string minAmount)
        {

            _firstMinAmount = WaitForElementIsVisible(By.Id(FIRST_MIN_AMOUNT));
            _firstMinAmount.Clear();
            _firstMinAmount.SendKeys(minAmount);
            WaitLoading();
        }

        public void FilterByNameOnStorageTab(string startName)
        {
            var searchName  = WaitForElementIsVisible(By.Id(FILTER_NAME));
            searchName.SetValue(ControlType.TextBox, startName);
            WaitPageLoading();
        }
        public string GetItemNameSupplier()
        {
            try
            {
                var name = WaitForElementExists(By.XPath(SUPPLIER_NAME));
                return name.GetAttribute("innerText");
            }
            catch
            {
                return "";
            }
        }
        public void EditShippingCost(string shippingCostToSet)
        {
            _firstShippingCost = WaitForElementIsVisible(By.Id(FIRST_SHIPPING_COST));
            _firstShippingCost.Clear();
            _firstShippingCost.SendKeys(shippingCostToSet);
            WaitLoading();
        }
        public string GetShippingCost()
        {

            _firstShippingCost = WaitForElementIsVisible(By.Id(FIRST_SHIPPING_COST));

            return _firstShippingCost.GetAttribute("value");

        }
    }
}