using DocumentFormat.OpenXml.Bibliography;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.Claims;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Accounting.SupplierInvoices
{
    public class SupplierInvoicesItem : PageBase
    {

        public SupplierInvoicesItem(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }
        //_________________________CONSTANTES____________________________________________________

        // Menu général
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";

        private const string EXTENDED_BUTTON = "Extended-Menu";
        private const string REFRESH = "btn-supplier-invoice-refresh";
        private const string EXPORT_SAGE = "btn-export-for-sage";
        private const string CONFIRM_EXPORT_SAGE = "btn-popup-validate";
        private const string DOWNLOAD_SAGE = "btn-download-file";
        private const string CANCEL_SAGE = "btn-cancel-popup";
        private const string SET_INTEGRATION_DATE = "ShowIntegrationDate";
        private const string SET_DATE = "datapicker-integration-date";
        private const string ENABLE_SAGE = "btn-enable-export-for-sage";

        private const string NOT_VERIFIED = "IsNotVerified-btn";
        private const string IS_VERIFIED = "IsVerified-btn";
        private const string VALIDATE = "btn-validate-supplier-invoice";
        private const string CONFIRM_VALIDATE = "btn-popup-validate";

        // Onglets
        private const string GENERAL_INFORMATION_TAB = "hrefTabContentInformations";
        private const string FOOTER_TAB = "hrefTabContentFooter";
        private const string ACCOUNTING_TAB = "hrefTabContentExportSageWriting";
        private const string GENERAL_INFORMATIONS_TAB = "/html/body/div[3]/div/div/div[3]/ul/li[1]/a";

        // Ajout item
        private const string CREATE_NEW_ROW = "btn-add-item-detail";
        private const string ITEM_NAME = "//*[@id=\"form-add-supplier-invoice-row-with-item\"]/div[1]/div[1]/div/span[1]/input[2]";
        private const string ITEM_QUANTITY = "//*[@id=\"form-add-supplier-invoice-row-with-item\"]//*[@id=\"Quantityinput\"]";
        private const string ITEM_PRICE = "//*[@id=\"form-add-supplier-invoice-row-with-item\"]//*[@id=\"PackingPriceInput\"]";
        private const string TAX_TYPE = "drop-down-taxes";
        private const string SAVE_NEW_ROW = "btn-save";
        private const string CLOSE_ADD_ROW_MODAL = "//*[@id=\"form-add-supplier-invoice-row-with-item\"]/div[2]/button[1]";
        private const string ERROR_MESSAGE_ITEM_REQUIRED = "//*[@id=\"form-add-supplier-invoice-row-with-item\"]/div[1]/div[1]/div/span";
        private const string ERROR_MESSAGE_TAX_REQUIRED = "//*[@id=\"form-add-supplier-invoice-row-with-item\"]/div[1]/div[4]/div/span";
        private const string SI_PRICE = "packagingPriceInput";
        private const string SI_QTY = "quantityInput";
        private const string TOTAL_VAT = "//*[@id=\"total-price-span\"]";
        private const string RN_CHECKBOX = "/html/body/div[5]/div/div/div[2]/div/form/div[1]/div[5]/div/input[2]";

        // Tableau items
        private const string ITEM_NUMBER = "editable-row";
        private const string ITEMS_NAME = "//*[starts-with(@id,\"item-tab-row_\")]/div/div/form/div[2]/div[4]/span[text()='{0}']";
        private const string FIRST_GROUP = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[1]/div/div/span";
        private const string FIRST_ITEM = "//*[starts-with(@id,\"item-tab-row_\")]/div/div/form/div[2]";
        private const string ITEM_QTY = "//*[starts-with(@id,\"item-tab-row_\")]/div/div/form/div[1]/div[4]/span[text()='{0}']/../../div[7]/input";
        private const string VAT_RATE = "dropdown-taxType";
        private const string DN_PRICE = "//*[starts-with(@id,\"item-tab-row_\")]/div/div/form/div[1]/div[3]/span[text()='{0}']/../../div[10]/input";
        private const string DN_QUANTITY = "//*[starts-with(@id,\"item-tab-row_\")]/div/div/form/div[1]/div[3]/span[text()='{0}']/../../div[11]/input";
        private const string ITEM_TAX_BASE_AMOUNT = "//*[starts-with(@id,\"item-tab-row_\")]/div/div/form/div[2]/div[7]";
        private const string ITEM_TOTAL_PRICE = "//*[starts-with(@id,\"item-tab-row_\")]/div/div/form/div[2]/div[3]/span[text()='{0}']/../../div[7]";
        private const string TOTAL_PRICE = "//form/div[2]//span[@name='ItemTotalExclVat']";
        private const string SUPPLIER_INVOICE_TOTAL_DN = "total-delivery-note-amount-span";
        private const string ITEM_QUANTITY_TABLEAU = "//*[starts-with(@id,\"item-tab-row_\")]/div/div/form/div[2]/div[4]/span[text()='{0}']/../../div[7]/span";
        private const string EDITCLAIM = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[text()='{0}']/../../div[13]/a";
        private const string NEWCLAIM = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[2]/div[3]/span[text()='{0}']/../../div[13]/a";
        private const string ITEM_VAT_AMOUNT_TABLEAU = "//*[starts-with(@id,\"item-tab-row_\")]/div/div/form/div[2]/div[4]/span[text()='{0}']/../../div[10]/span";
        private const string ITEM_VAT_RATE_TABLEAU = "//span[text()='{0}']/../..//span[@name='ItemTaxName']";
        private const string ITEM_EXTEND_MENU = "//*[@id=\"item-tab-row_241971\"]/div/div/form/div[1]/div[15]/a";
        private const string DELETE_BUTTON = "//span[text()='{0}']/../..//a[@class='mini-btn open-item-packaging']";
        private const string CHECK_IF_ITEM_DELETED = "//div[@class = 'display-zone display-zone-unique']//span[contains(text(), '{0}')]";
        private const string ICON_SAVED = "//span[@class='glyphicon glyphicon-floppy-saved']";
        // Filtres
        private const string RESET_FILTER = "//*[@id=\"formSearchDetails\"]/div[1]/a";
        private const string RECEIPT_NOTE_NUMBER = "SearchReceiptNoteNumber";
        private const string PURCHASE_ORDER_NUMBER = "SearchPurchaseOrderNumber";
        private const string GROUP_FILTER = "ItemIndexVMSelectedGroups_ms";
        private const string GROUP_FILTER_SEARCH = "/html/body/div[11]/div/div/label/input";
        private const string UNSELECT_ALL_SITES = "/html/body/div[11]/div/ul/li[2]/a";
        private const string FIRST_LINE_ADD_CLAIM = "/html/body/div[3]/div/div/div[3]/div/div/div/div[2]/div/div[2]/div/div[2]/div[2]/div/div/form/div[1]/div[14]/a/span";
        private const string FIRST_CHECK_CLAIM_TYPE = "IsChecked_0";
        private const string CLAIM_COMMENT = "Comment";
        private const string CLAIM_SAVE_BTN = "btn-valid-claim";
        private const string FIRST_SUPPLIER_INVOICE_LINE = "/html/body/div[3]/div/div/div[3]/div/div/div/div[2]/div/div[2]/div/div[2]/div[2]/div/div/form/div[2]";
        private const string CLAIM_NUBMER = "//*[@id=\"form-create-supplier-invoice\"]/div/div[8]/div/div/div/ul/li/a";
        private const string CLAIMS_TAB = "hrefTabContentClaims";
        private const string FIRST_CLAIM = "/html/body/div[3]/div/div/div[3]/div/div/div/div/div[1]/div[2]/div/div";
        private const string TRASH_BTN = "/html/body/div[3]/div/div/div[3]/div/div/div/div/div[1]/div[2]/div/div/div[15]/a/span";
        private const string NO_ITEM_FOUND_IN_CLAIMS = "/html/body/div[3]/div/div/div[3]/div/div/div/div/p";
        private const string CLAIMS_COUNT = "//*[@id=\"div-body\"]/div/div[2]/div[1]/h1/span";
        private const string UNCHECK_ALL_SUB_GROUP = "/html/body/div[12]/div/ul/li[2]/a/span[2]";
        private const string SEARCH_SUB_GROUP = "ItemIndexVMSelectedSubGroups_ms";
        private const string SEARCH_SUB_GROUP_TEXT = "/html/body/div[11]/div/div/label/input";
        private const string SI_PRICE_POP_UP = "PackingPriceInput";
        private const string FIRSTQUANTITY = "/html/body/div[3]/div/div/div[3]/div/div/div/div[2]/div/div[2]/div/div[2]/div[2]/div/div/form/div[2]/div[7]/span";
        private const string ITEMS_TAB = "hrefTabContentItems";


        private const string SEARCH_ITEM = "//*[@id=\"form-add-supplier-invoice-row-with-item\"]/div[1]/div[1]/div/span[1]/div/div/div[1]";
        private const string NOT_VERIFIED_SHOW = "//*[@id=\"accounting-supplierinvoice-details-1\"]/div/div/i";
        private const string VERIFIED_SHOW = "//*[@id=\"accounting-supplierinvoice-details-1\"]/div/div/i";
        private const string ICON_COMMENT_GREEN = "//*[@id=\"tabContentServiceContainer\"]/div[1]/div[2]/div/div/div[12]/span";
        private const string IS_ACTIVE = "SupplierInvoice_IsActive";
        private const string REINVOICE_IN_CUSTOMER_INVOICE = "btn-reinvoice-in-customer-invoice";
        private const string REINVOICE_IN_CUSTOMER_INVOICE_VERIFY_BUTTON = "btnCreateInvoice";
        private const string REINVOICE_IN_CUSTOMER_INVOICE_GENERATE_BUTTON = "btnCreateInvoice";
        private const string REINVOICE_IN_CUSTOMER_INVOICE_CANCEL_BUTTON = "//*[@id=\"form-reinvoice-data\"]/div[5]/button[1]";
        
        //_________________________VARIABLES____________________________________________________

        // Menu général
        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;
        [FindsBy(How = How.XPath, Using = FIRSTQUANTITY)]
        private IWebElement _firstquantity;

        [FindsBy(How = How.Id, Using = EXTENDED_BUTTON)]
        private IWebElement _extendedButton1;

        [FindsBy(How = How.Id, Using = REFRESH)]
        private IWebElement _refresh;
        [FindsBy(How = How.Id, Using = ITEMS_TAB)]
        private IWebElement _itemsTab;

        [FindsBy(How = How.Id, Using = EXPORT_SAGE)]
        private IWebElement _exportSage;

        [FindsBy(How = How.Id, Using = CONFIRM_EXPORT_SAGE)]
        private IWebElement _confirmExportSage;

        [FindsBy(How = How.Id, Using = DOWNLOAD_SAGE)]
        private IWebElement _downloadSage;

        [FindsBy(How = How.Id, Using = CANCEL_SAGE)]
        private IWebElement _cancelSage;

        [FindsBy(How = How.Id, Using = SET_INTEGRATION_DATE)]
        private IWebElement _setIntegrationDate;

        [FindsBy(How = How.Id, Using = SET_DATE)]
        private IWebElement _setDate;

        [FindsBy(How = How.Id, Using = ENABLE_SAGE)]
        private IWebElement _enableSage;

        [FindsBy(How = How.Id, Using = NOT_VERIFIED)]
        private IWebElement _notVerified;

        [FindsBy(How = How.Id, Using = VALIDATE)]
        private IWebElement _validate;

        [FindsBy(How = How.Id, Using = CONFIRM_VALIDATE)]
        private IWebElement _confirmValidate;

        // Onglets
        [FindsBy(How = How.Id, Using = GENERAL_INFORMATION_TAB)]
        private IWebElement _generalInformationTab;

        [FindsBy(How = How.Id, Using = FOOTER_TAB)]
        private IWebElement _footerTab;

        [FindsBy(How = How.Id, Using = ACCOUNTING_TAB)]
        private IWebElement _accountingTab;

        // Ajout item
        [FindsBy(How = How.Id, Using = CREATE_NEW_ROW)]
        private IWebElement _createNewRow;

        [FindsBy(How = How.XPath, Using = ITEM_NAME)]
        private IWebElement _itemName;

        [FindsBy(How = How.XPath, Using = ITEM_QUANTITY)]
        private IWebElement _itemQuantity;

        [FindsBy(How = How.XPath, Using = ITEM_PRICE)]
        private IWebElement _itemPrice;

        [FindsBy(How = How.Id, Using = TAX_TYPE)]
        private IWebElement _itemTaxType;

        [FindsBy(How = How.XPath, Using = SAVE_NEW_ROW)]
        private IWebElement _saveNewRow;

        [FindsBy(How = How.XPath, Using = CLOSE_ADD_ROW_MODAL)]
        private IWebElement _closeAddRowModal;

        [FindsBy(How = How.XPath, Using = ERROR_MESSAGE_ITEM_REQUIRED)]
        private IWebElement _errorMessageItemRequired;

        [FindsBy(How = How.XPath, Using = ERROR_MESSAGE_TAX_REQUIRED)]
        private IWebElement _errorMessageTaxRequired;

        // Tableau items
        [FindsBy(How = How.XPath, Using = FIRST_GROUP)]
        private IWebElement _firstGroup;

        [FindsBy(How = How.XPath, Using = FIRST_ITEM)]
        private IWebElement _firstItem;

        [FindsBy(How = How.XPath, Using = ITEMS_NAME)]
        private IWebElement _item;

        [FindsBy(How = How.XPath, Using = ITEM_QTY)]
        private IWebElement _itemQty;

        [FindsBy(How = How.XPath, Using = DN_PRICE)]
        private IWebElement _dnPrice;

        [FindsBy(How = How.XPath, Using = DN_QUANTITY)]
        private IWebElement _dnQuantity;

        [FindsBy(How = How.XPath, Using = ITEM_QUANTITY_TABLEAU)]
        private IWebElement _itemQuantityTableau;

        [FindsBy(How = How.XPath, Using = ITEM_VAT_AMOUNT_TABLEAU)]
        private IWebElement _itemVATAmountTableau;

        [FindsBy(How = How.XPath, Using = ITEM_VAT_RATE_TABLEAU)]
        private IWebElement _itemVATRateTableau;

        [FindsBy(How = How.Id, Using = VAT_RATE)]
        private IWebElement _vatRate;

        [FindsBy(How = How.XPath, Using = ITEM_TAX_BASE_AMOUNT)]
        private IWebElement _itemTaxBaseAmount;

        [FindsBy(How = How.XPath, Using = ITEM_TOTAL_PRICE)]
        private IWebElement _itemTotalPrice;

        [FindsBy(How = How.Id, Using = SUPPLIER_INVOICE_TOTAL_DN)]
        private IWebElement _totalDN;

        // Filtres
        [FindsBy(How = How.XPath, Using = RESET_FILTER)]
        private IWebElement _resetFilter;

        [FindsBy(How = How.Id, Using = RECEIPT_NOTE_NUMBER)]
        private IWebElement _receiptNoteNumber;

        [FindsBy(How = How.Id, Using = PURCHASE_ORDER_NUMBER)]
        private IWebElement _purchaseOrderNumber;

        [FindsBy(How = How.Id, Using = GROUP_FILTER)]
        private IWebElement _groupFilter;

        [FindsBy(How = How.XPath, Using = GROUP_FILTER_SEARCH)]
        private IWebElement _groupFilterSearch;

        [FindsBy(How = How.XPath, Using = UNSELECT_ALL_SITES)]
        private IWebElement _unselectAllSites;

        [FindsBy(How = How.Id, Using = SI_PRICE)]
        private IWebElement _siPackPrice;

        [FindsBy(How = How.Id, Using = SI_QTY)]
        private IWebElement _siQty;

        [FindsBy(How = How.XPath, Using = TOTAL_VAT)]
        private IWebElement _siTotalVat;

        [FindsBy(How = How.Id, Using = SI_PRICE_POP_UP)]
        private IWebElement _siPackPricePopUp;


        [FindsBy(How = How.XPath, Using = NOT_VERIFIED_SHOW)]
        private IWebElement _notVerifiedShow;

        [FindsBy(How = How.XPath, Using = VERIFIED_SHOW)]
        private IWebElement _VerifiedShow;

        [FindsBy(How = How.XPath, Using = REINVOICE_IN_CUSTOMER_INVOICE)]
        private IWebElement _reinvoiceInCustomerInvoice;
        // __________________________________________Filters _____________________________________________________

        public enum FilterItemType
        {
            ByReceiptNoteNumber,
            ByPurchaseOrderNumber,
            ByGroup,
            BySubGrp
        }

        public void Filter(FilterItemType FilterItemType, object value)
        {
            switch (FilterItemType)
            {
                case FilterItemType.ByReceiptNoteNumber:
                    _receiptNoteNumber = WaitForElementIsVisible(By.Id(RECEIPT_NOTE_NUMBER));
                    _receiptNoteNumber.SetValue(ControlType.TextBox, value);
                    break;
                case FilterItemType.ByPurchaseOrderNumber:
                    _purchaseOrderNumber = WaitForElementIsVisible(By.Id(PURCHASE_ORDER_NUMBER));
                    _purchaseOrderNumber.SetValue(ControlType.TextBox, value);
                    break;
                case FilterItemType.ByGroup:
                    _groupFilter = WaitForElementIsVisible(By.Id(GROUP_FILTER));
                    _groupFilter.Click();

                    _unselectAllSites = _webDriver.FindElement(By.XPath(UNSELECT_ALL_SITES));
                    _unselectAllSites.Click();

                    _groupFilterSearch = WaitForElementIsVisible(By.XPath(GROUP_FILTER_SEARCH));
                    _groupFilterSearch.SetValue(ControlType.TextBox, value);
                    WaitForLoad();
                    var groupFilterValue = WaitForElementIsVisible(By.XPath(string.Format("//*/span[text()='{0}']/parent::label/input", value)));
                    groupFilterValue.Click();
                    _groupFilter.Click();
                    break;
                case FilterItemType.BySubGrp:
                    FilterUncheckAllSubGroup();
                    var filterBySubgrp = WaitForElementIsVisible(By.XPath(SEARCH_SUB_GROUP_TEXT));
                    filterBySubgrp.SetValue(ControlType.TextBox, value);
                    WaitForLoad();
                    var subgrptocheck = WaitForElementIsVisible(By.XPath("//label//span[text()='" + value + "']"));
                    subgrptocheck.Click();
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterItemType), FilterItemType, null);
            }

            WaitPageLoading();
            WaitForLoad();
        }

        public bool IsItemFiltered(string itemName)
        {
            return isElementVisible(By.XPath(String.Format("//*[starts-with(@id,\"item-tab-row_\")]/div/div/form/div[2]/div[4]/span[contains(text(),'{0}')]", itemName)));
        }

        public void FilterUncheckAllSubGroup()
        {
            var seachGrp = WaitForElementIsVisible(By.Id(SEARCH_SUB_GROUP));
            seachGrp.Click();

            var uncheckallgrp = WaitForElementIsVisible(By.XPath(UNCHECK_ALL_SUB_GROUP));
            uncheckallgrp.Click();
        }

        public bool IsFilteredByGroup(string group)
        {
            if (isElementVisible(By.XPath((FIRST_GROUP))))
            {
                _firstGroup = WaitForElementIsVisible(By.XPath(FIRST_GROUP));
                if (_firstGroup.Text.Equals(group))
                    return true;
                else
                    return false;
            }
            else
            {
                return false;
            }
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

        //___________________________________________ Méthodes _________________________________________

        // Menu Général
        public SupplierInvoicesPage BackToList()
        {
            _backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            _backToList.Click();
            WaitForLoad();

            return new SupplierInvoicesPage(_webDriver, _testContext);
        }

        public override void ShowExtendedMenu()
        {
            var actions = new Actions(_webDriver);

            _extendedButton1 = WaitForElementIsVisible(By.Id(EXTENDED_BUTTON));
            actions.MoveToElement(_extendedButton1).Perform();
        }

        public void Refresh()
        {
            ShowExtendedMenu();

            _refresh = WaitForElementIsVisible(By.Id(REFRESH));
            _refresh.Click();
            WaitForLoad();
        }

        public string ManualExportSageError(bool printValue, bool isMessage)
        {
            string errorMessage = "";

            ShowExtendedMenu();

            _exportSage = WaitForElementIsVisible(By.Id(EXPORT_SAGE));
            _exportSage.Click();
            WaitForLoad();

            _confirmExportSage = WaitForElementIsVisible(By.Id(CONFIRM_EXPORT_SAGE));
            _confirmExportSage.Click();
            WaitForLoad();

            if (!printValue)
            {
                if (isElementVisible(By.Id(DOWNLOAD_SAGE)))
                {
                    _downloadSage = _webDriver.FindElement(By.Id(DOWNLOAD_SAGE));
                    _downloadSage.Click();
                    WaitForLoad();

                    // On ferme la pop-up
                    _cancelSage = WaitForElementIsVisible(By.Id(CANCEL_SAGE));
                    _cancelSage.Click();
                    WaitForLoad();
                }
                else
                {
                    _cancelSage = WaitForElementIsVisible(By.Id(CANCEL_SAGE));
                    _cancelSage.Click();
                    WaitForLoad();
                }
            }
            else
            {
                errorMessage = IsFileInError(By.CssSelector("[class='fa fa-info-circle']"), isMessage);
                ClickPrintButton();
            }

            return errorMessage;
        }

        public void ManualExportSage(bool printValue, bool integrationDate = false)
        {
            ShowExtendedMenu();

            _exportSage = WaitForElementIsVisible(By.Id(EXPORT_SAGE));
            _exportSage.Click();
            WaitForLoad();

            if (integrationDate)
            {
                var firstDayOftheMonth = new DateTime(DateUtils.Now.Year, DateUtils.Now.Month, 1);

                _setIntegrationDate = WaitForElementExists(By.Id(SET_INTEGRATION_DATE));
                _setIntegrationDate.SetValue(ControlType.CheckBox, true);
                WaitForLoad();

                _setDate = WaitForElementExists(By.Id(SET_DATE));
                _setDate.SetValue(ControlType.DateTime, firstDayOftheMonth);
                WaitForLoad();
            }

            _confirmExportSage = WaitForElementIsVisible(By.Id(CONFIRM_EXPORT_SAGE));
            _confirmExportSage.Click();
            WaitForLoad();

            if (!printValue)
            {
                if (isElementVisible(By.Id(DOWNLOAD_SAGE)))
                {
                    _downloadSage = _webDriver.FindElement(By.Id(DOWNLOAD_SAGE));
                    _downloadSage.Click();
                    WaitForLoad();

                    // On ferme la pop-up
                    _cancelSage = WaitForElementIsVisible(By.Id(CANCEL_SAGE));
                    _cancelSage.Click();
                    WaitForLoad();
                }
                else
                {
                    _cancelSage = WaitForElementIsVisible(By.Id(CANCEL_SAGE));
                    _cancelSage.Click();
                    WaitForLoad();
                }
            }
            else
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-alt']"));
                ClickPrintButton();
            }

            WaitForDownload();
            Close();
        }

        public FileInfo GetExportSAGEFile(FileInfo[] taskFiles)
        {
            List<FileInfo> correctDownloadFiles = new List<FileInfo>();

            foreach (var file in taskFiles)
            {
                //  Test REGEX
                if (IsSAGEFileCorrect(file.Name))
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

        public bool IsSAGEFileCorrect(string filePath)
        {
            // "supplier-invoices 2019-12-09 13-20-29.txt"

            string mois = "(?:0[1-9]|1[0-2])";         // mois
            string space = "(\\s)";                    // Espace
            string annee = "\\d{4}";                   // annee YYYY
            string jour = "[0-3]\\d";                  // jour
            string heure = "(?:0[0-9]|1[0-9]|2[0-3])"; // heure
            string minutes = "[0-5]\\d";               // minutes
            string secondes = "[0-5]\\d";              // secondes


            Regex r = new Regex("^supplier-invoices" + space + annee + "-" + mois + "-" + jour + space + heure + "-" + minutes + "-" + secondes + ".txt$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = r.Match(filePath);

            return m.Success;
        }

        public bool CanClickOnSAGE()
        {
            ShowExtendedMenu();

            _exportSage = WaitForElementExists(By.Id(EXPORT_SAGE));
            if (_exportSage.GetAttribute("disabled") != null)
                return false;

            return true;
        }

        public bool CanClickOnEnableSAGE()
        {
            ShowExtendedMenu();

            _enableSage = WaitForElementIsVisible(By.Id(ENABLE_SAGE));
            if (_enableSage.GetAttribute("disabled") != null)
                return false;

            return true;
        }

        public void ClickOnEnableSAGE()
        {
            ShowExtendedMenu();

            _enableSage = WaitForElementIsVisible(By.Id(ENABLE_SAGE));
            _enableSage.Click();
            WaitForLoad();
        }

        public void ValidateSupplierInvoice()
        {
            // menu dynamique (Validate grisé puis normal)
            ShowValidationMenu();

            _validate = WaitForElementIsVisible(By.Id(VALIDATE));
            _validate.Click();
            WaitLoading();

            _confirmValidate = WaitForElementIsVisible(By.Id(CONFIRM_VALIDATE));
            _confirmValidate.Click();
            WaitPageLoading();
        }

        public void ReinvoiceInCustomerInvoice()
        {
            // menu dynamique (Validate grisé puis normal)
            ShowExtendedMenu();

            _reinvoiceInCustomerInvoice = WaitForElementIsVisible(By.Id(REINVOICE_IN_CUSTOMER_INVOICE));
            _reinvoiceInCustomerInvoice.Click();
            WaitLoading();
            var _buttonVerify = WaitForElementIsVisible(By.Id(REINVOICE_IN_CUSTOMER_INVOICE_VERIFY_BUTTON));
            _buttonVerify.Click();
            WaitPageLoading();
            var _buttonGenerate = WaitForElementIsVisible(By.Id(REINVOICE_IN_CUSTOMER_INVOICE_GENERATE_BUTTON));
            _buttonGenerate.Click();
            WaitPageLoading();
            var _cancel = WaitForElementIsVisible(By.XPath(REINVOICE_IN_CUSTOMER_INVOICE_CANCEL_BUTTON));
            _cancel.Click();
            WaitPageLoading();

        }


        public void SetVerified()
        {
            ShowValidationMenu();

            _notVerified = WaitForElementIsVisible(By.Id(NOT_VERIFIED));
            _notVerified.Click();

            var confirm = WaitForElementIsVisible(By.Id("btn-popup-validate"));
            confirm.Click();

            WaitForLoad();
        }

        public bool IsVerified()
        {
            ShowValidationMenu();
            if (isElementVisible(By.Id(NOT_VERIFIED)))
            {
                _notVerified = WaitForElementIsVisible(By.Id(NOT_VERIFIED));

            }
            else
            {
                _notVerified = WaitForElementIsVisible(By.Id(IS_VERIFIED));

            }

            return _notVerified.Text.Equals(" Verified");
        }


        // Onglets
        public SupplierInvoicesGeneralInformation ClickOnGeneralInformation()
        {
            _generalInformationTab = WaitForElementIsVisible(By.Id(GENERAL_INFORMATION_TAB));
            _generalInformationTab.Click();
            WaitForLoad();

            return new SupplierInvoicesGeneralInformation(_webDriver, _testContext);
        }

        public SupplierInvoicesFooterPage ClickOnFooter()
        {
            _footerTab = WaitForElementIsVisible(By.Id(FOOTER_TAB));
            _footerTab.Click();
            WaitForLoad();

            return new SupplierInvoicesFooterPage(_webDriver, _testContext);
        }
        public SupplierInvoicesFooterPage ClickOnItems()
        {
            _itemsTab = WaitForElementIsVisible(By.Id(ITEMS_TAB));
            _itemsTab.Click();
            WaitForLoad();

            return new SupplierInvoicesFooterPage(_webDriver, _testContext);
        }

        public SupplierInvoicesAccounting ClickOnAccounting()
        {
            _accountingTab = WaitForElementIsVisible(By.Id(ACCOUNTING_TAB));
            _accountingTab.Click();
            WaitForLoad();

            return new SupplierInvoicesAccounting(_webDriver, _testContext);
        }

        // Ajout item
        public void AddNewItem(string item, string qty, string taxType = null, string price = "20")
        {
            _createNewRow = WaitForElementIsVisible(By.Id(CREATE_NEW_ROW));
            _createNewRow.Click();
            WaitForLoad();

            //input with search bar
            _itemName = WaitForElementIsVisible(By.XPath(ITEM_NAME));
            _itemName.SetValue(ControlType.TextBox, item);
            WaitForLoad();

            var searchItemValue = WaitForElementIsVisible(By.XPath(SEARCH_ITEM));
            searchItemValue.Click();
            WaitForLoad();

            _itemName.SendKeys(Keys.Tab);
            WaitForLoad();

            _itemQuantity = WaitForElementIsVisible(By.XPath(ITEM_QUANTITY));
            _itemQuantity.SetValue(ControlType.TextBox, qty);
            // si pas de price alors pas de Footer
            WaitForLoad();

            _itemPrice = WaitForElementIsVisible(By.XPath(ITEM_PRICE));
            _itemPrice.SetValue(ControlType.TextBox, price);
            WaitForLoad();

            if (taxType != null)
            {
                _itemTaxType = WaitForElementIsVisible(By.Id(TAX_TYPE));
                _itemTaxType.SetValue(ControlType.DropDownList, taxType);
                WaitForLoad();
            }

            _saveNewRow = WaitForElementIsVisible(By.Id(SAVE_NEW_ROW));
            _saveNewRow.Click();
            WaitPageLoading();
            WaitForLoad();
        }
        public void AddNewItemForRN(string item, string qty, string taxType = null, string price = "20")
        {
            _createNewRow = WaitForElementIsVisible(By.Id(CREATE_NEW_ROW));
            _createNewRow.Click();
            WaitForLoad();

            //input with search bar
            _itemName = WaitForElementIsVisible(By.XPath(ITEM_NAME));
            _itemName.SetValue(ControlType.TextBox, item);
            WaitForLoad();

            var searchItemValue = WaitForElementIsVisible(By.XPath(SEARCH_ITEM));
            searchItemValue.Click();
            WaitForLoad();

            _itemName.SendKeys(Keys.Tab);
            WaitForLoad();

            _itemQuantity = WaitForElementIsVisible(By.XPath(ITEM_QUANTITY));
            _itemQuantity.SetValue(ControlType.TextBox, qty);
            // si pas de price alors pas de Footer
            WaitForLoad();

            _itemPrice = WaitForElementIsVisible(By.XPath(ITEM_PRICE));
            _itemPrice.SetValue(ControlType.TextBox, price);
            WaitForLoad();

            if (taxType != null)
            {
                _itemTaxType = WaitForElementIsVisible(By.Id(TAX_TYPE));
                _itemTaxType.SetValue(ControlType.DropDownList, taxType);
                WaitForLoad();
            }

         
            WaitPageLoading();
            WaitForLoad();
        }
        public void CheckRN()
        {
            var Element = _webDriver.FindElement(By.XPath(RN_CHECKBOX));

            Element.Click();



        }
        public bool IsRNChecked()
        {
            var radioButton = _webDriver.FindElement(By.XPath(RN_CHECKBOX));
            return radioButton.Selected;
        }


        public void AddNewItemError()
        {
            _createNewRow = WaitForElementIsVisible(By.Id(CREATE_NEW_ROW));
            _createNewRow.Click();
            WaitForLoad();

            _saveNewRow = WaitForElementIsVisible(By.Id(SAVE_NEW_ROW));
            _saveNewRow.Click();
            WaitForLoad();
        }

        public void CloseNewItemModal()
        {
            _closeAddRowModal = WaitForElementIsVisible(By.XPath(CLOSE_ADD_ROW_MODAL));
            _closeAddRowModal.Click();
            WaitForLoad();

            // Temps de fermeture de la fenêtre : ne pas enlever

            WaitPageLoading();
        }

        public bool ErrorMessageItemRequired()
        {
            _errorMessageItemRequired = _webDriver.FindElement(By.XPath(ERROR_MESSAGE_ITEM_REQUIRED));
            return _errorMessageItemRequired.Displayed;
        }

        public bool ErrorMessageTaxRequired()
        {
            _errorMessageTaxRequired = _webDriver.FindElement(By.XPath(ERROR_MESSAGE_TAX_REQUIRED));
            return _errorMessageTaxRequired.Displayed;
        }

        // Tableau items
        public int GetNumberOfItems()
        {
            WaitPageLoading();
            int itemsNumber = 0;
            var itemsAdded = isElementVisible(By.ClassName(ITEM_NUMBER));
            if (itemsAdded)
            {
                itemsNumber = _webDriver.FindElements(By.ClassName(ITEM_NUMBER)).Count;
            }
            return itemsNumber;
        }

        public void SelectFirstItem()
        {
            _firstItem = WaitForElementIsVisible(By.XPath(FIRST_ITEM));
            _firstItem.Click();
        }

        public void SelectItem(string itemName)
        {
            _item = WaitForElementIsVisible(By.XPath(String.Format("//*[starts-with(@id,\"item-tab-row_\")]/div/div/form/div[2]/div[4]/span[text()='{0}']", itemName)));

            _item.Click();
        }

        public void SetItemQuantity(string itemName, string quantity)
        {
            _itemQty = WaitForElementIsVisible(By.XPath(string.Format(ITEM_QTY, itemName)));
            _itemQty.SetValue(ControlType.TextBox, quantity);
            WaitForLoad();

            // Temps d'attente obligatoire pour la prise en compte de l'item
            if (isElementVisible(By.XPath(ICON_SAVED)))
            {
                WaitForElementIsVisible(By.XPath(ICON_SAVED));
            }
            WaitPageLoading();
        }

        public void SetVatRate(string vatRate)
        {
            _vatRate = WaitForElementIsVisible(By.Id(VAT_RATE));
            _vatRate.SetValue(ControlType.DropDownList, vatRate);
            WaitPageLoading();

            WaitForLoad();
        }
        public string GetSIPrice()
        {
            var _siPrice = WaitForElementIsVisible(By.Id("packagingPriceInput"));
            return _siPrice.Text;
        }


        public void SetDNPrice(string dnPrice, string itemName)
        {
            _dnPrice = WaitForElementIsVisible(By.XPath("//span[text()='" + itemName + "']/../../div[11]/input"));

            _dnPrice.SetValue(ControlType.TextBox, dnPrice);
            WaitForLoad();
        }
        public void SetDNQuantity(string dnQuantity, string itemName)
        {
            _dnQuantity = WaitForElementIsVisible(By.XPath("//span[text()='" + itemName + "']/../../div[12]/input"));

            _dnQuantity.SetValue(ControlType.TextBox, dnQuantity);
            WaitForLoad();
            Thread.Sleep(2000);
        }

        public double GetItemTaxBaseAmount(string currency, string decimalSeparator)
        {
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");
            _itemTaxBaseAmount = WaitForElementIsVisible(By.XPath("//*[starts-with(@id,\"item-tab-row_\")]/div/div/form/div[2]/div[8]"));

            string element = _itemTaxBaseAmount.Text.Replace(currency, "").Trim();

            return Math.Round(double.Parse(element, ci), 2);
        }

        public double GetItemTotalPrice(string itemName, string currency, string decimalSeparator)
        {
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            _itemTotalPrice = WaitForElementIsVisible(By.XPath(string.Format("//*[starts-with(@id,\"item-tab-row_\")]/div/div/form/div[2]/div[4]/span[text()='{0}']/../../div[8]", itemName)));

            string element = _itemTotalPrice.Text.Replace(currency, "").Trim();

            return Math.Round(double.Parse(element, ci), 2);
        }

        public bool IsItemAdded(string itemName)
        {
            var itemAdded = isElementVisible(By.XPath(string.Format("//div[@class = 'display-zone display-zone-unique']//span[contains(text(), '{0}')]", itemName)));
            return itemAdded;
        }
        public bool IsItemDeleted(string itemName)
        {
            var itemexist = isElementVisible(By.XPath(string.Format(CHECK_IF_ITEM_DELETED, itemName)));
            if (itemexist)
            {
                return false;
            }
            return true;
        }
        public int GetItemVATAmount(string itemName)
        {

            _itemVATAmountTableau = WaitForElementIsVisible(By.XPath(string.Format(ITEM_VAT_AMOUNT_TABLEAU, itemName)));
            return Int16.Parse(_itemVATAmountTableau.Text.Trim(new Char[] { ' ', '€', '.' }));
        }
        public string GetItemVATRate(string itemName)
        {
            _itemVATRateTableau = WaitForElementIsVisible(By.XPath(string.Format(ITEM_VAT_RATE_TABLEAU, itemName)));
            return _itemVATRateTableau.Text;
        }

        public string GetItemQuantity(string itemName)
        {

            _itemQuantityTableau = WaitForElementIsVisible(By.XPath(string.Format(ITEM_QUANTITY_TABLEAU, itemName)));
            return _itemQuantityTableau.Text.Trim();
        }
        public double GetInvoiceTotalPrice(string currency, string decimalSeparator)
        {
            WaitForLoad();
            double montant = 0;

            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            ReadOnlyCollection<IWebElement> itemTotalPrices;
            itemTotalPrices = _webDriver.FindElements(By.XPath("//form/div[2]//div[8]"));

            if (itemTotalPrices.Count != 0)
            {
                foreach (var totalprice in itemTotalPrices)
                {
                    string element = totalprice.Text.Replace(currency, "").Trim();
                    element = element.Replace(" ", "").Trim();
                    if (!string.IsNullOrEmpty(element))
                    {
                        montant += double.Parse(element, CultureInfo.InvariantCulture);
                    }

                }
            }

            return montant;
        }

        public double GetInvoiceTotalDN(string currency, string decimalSeparator)
        {
            double total = 0;
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            _totalDN = WaitForElementIsVisible(By.Id(SUPPLIER_INVOICE_TOTAL_DN), nameof(SUPPLIER_INVOICE_TOTAL_DN));

            if (_totalDN.Text != null)
            {
                string element = _totalDN.Text.Replace(currency, "").Trim();
                total = Math.Round(double.Parse(element, ci), 2);
            }

            return total;
        }

        public ClaimEditClaimForm EditClaimForm(string itemName)
        {
            IWebElement editClaim;
            // ligne en mode édition
            editClaim = WaitForElementIsVisible(By.XPath(String.Format("//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[4]/span[text()='{0}']/../../div[14]/a", itemName)));

            editClaim.Click();
            WaitForLoad();
            return new ClaimEditClaimForm(_webDriver, _testContext);
        }
        public ClaimEditClaimForm EditFirstClaim()
        {
            var editFirstClaim = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div/div[3]/div/div/div/div/div[1]/div[2]/div/div/div[14]/a/span"));
            editFirstClaim.Click();
            return new ClaimEditClaimForm(_webDriver, _testContext);
        }
        public void ClickFirstItem()
        {
            var first_item = WaitForElementIsVisible(By.XPath(FIRST_ITEM));
            first_item.Click();
        }
        public void DeleteItem(string itemName)
        {
            var btn = WaitForElementIsVisible(By.XPath(String.Format(DELETE_BUTTON, itemName)));
            Actions act = new Actions(_webDriver);
            act.DoubleClick(btn).Perform();
        }
        public void AddFirstSupplierInvoiceDetailClaim(string comment)
        {
            var firstLine = WaitForElementIsVisible(By.XPath(FIRST_SUPPLIER_INVOICE_LINE));
            firstLine.Click();

            var firstLineMegaphone = WaitForElementIsVisible(By.XPath(FIRST_LINE_ADD_CLAIM));
            firstLineMegaphone.Click();

            AddClaim(comment);
            WaitForLoad();
        }
        public void AddClaim(string commentText)
        {
            var firstCheck = WaitForElementIsVisible(By.Id(FIRST_CHECK_CLAIM_TYPE));
            firstCheck.SetValue(ControlType.CheckBox, true);
            var comment = WaitForElementIsVisible(By.Id(CLAIM_COMMENT));
            comment.SendKeys(commentText);
            var btnSave = WaitForElementIsVisible(By.Id(CLAIM_SAVE_BTN));
            btnSave.Click();
        }
        public string GetClaimNumber()
        {
            var claimNumber = WaitForElementIsVisible(By.Id("tb-new-supplier-invoice-number"));
            return claimNumber.GetAttribute("value");
        }
        public string GetfirstQuantity()
        {
            _firstquantity = WaitForElementIsVisible(By.XPath(FIRSTQUANTITY));
            return _firstquantity.Text;
        }
        public void ClickonClaims()
        {
            var claimsTab = WaitForElementIsVisible(By.Id(CLAIMS_TAB));
            claimsTab.Click();
        }
        public void DeleteClaim()
        {
            var firstClaim = WaitForElementIsVisible(By.XPath(FIRST_CLAIM));
            firstClaim.Click();
            var trashBtn = WaitForElementIsVisible(By.XPath(TRASH_BTN));
            trashBtn.Click();
            WaitForLoad();
        }
        public bool VerifySupplierClaimsIsEmpty()
        {
            return isElementExists(By.XPath(NO_ITEM_FOUND_IN_CLAIMS));
        }
        public bool ClaimExistInWarehouseClaims(string number)
        {
            var page = GoToWarehouse_ClaimsPage();
            page.Filter(ClaimsPage.FilterType.ByNumber, number);
            var claimsCount = WaitForElementIsVisible(By.XPath(CLAIMS_COUNT));
            if (claimsCount.Text == "0")
            {
                return false;
            }
            return true;
        }
        public bool ClaimExist(string number)
        {
            if (!VerifySupplierClaimsIsEmpty() && ClaimExistInWarehouseClaims(number))
            {
                return true;
            }
            return false;
        }
        public void ClickOnGeneralInformationTab()
        {
            var tab = WaitForElementIsVisible(By.XPath(GENERAL_INFORMATIONS_TAB));
            tab.Click();
            WaitForLoad();
        }
        public bool CheckNoDetails()
        {
            var noDetailsMsg = isElementExists(By.XPath("/html/body/div[3]/div/div/div[3]/div/div/div/div[2]/div/div[2]/div/div[2]/p/i"));
            return noDetailsMsg;
        }
        public bool VerifyItemExist(string itemname)
        {
            var contains = isElementExists(By.XPath("/html/body/div[3]/div/div/div[3]/div/div/div/div[2]/div/div[2]/div/div[2]/div[2]/div/div/form/div[2]/div[4]/span"));
            if (contains)
            {
                var item = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div/div[3]/div/div/div/div[2]/div/div[2]/div/div[2]/div[2]/div/div/form/div[2]/div[4]/span"));
                if (item.Text == itemname)
                {
                    return true;
                }
            }
            return false;

        }

        public string GetSI_PackPrice()
        {
            var _siPackPrice = WaitForElementIsVisible(By.Id(SI_PRICE));

            return _siPackPrice.GetAttribute("value");
        }
        public string GetSI_Qty()
        {

            var _siQty = WaitForElementIsVisible(By.Id(SI_QTY));
            string qty = _siQty.GetAttribute("value");
            WaitForLoad();
            return qty;
        }
        public void SetSI_Qty(string qty)
        {
            var _siQty = WaitForElementIsVisible(By.Id(SI_QTY));
            _siQty.SetValue(ControlType.TextBox, qty);
            WaitPageLoading();
            WaitPageLoading();


        }
        public string GetSI_Total_Vat()
        {
            var _siTotalVat = WaitForElementIsVisible(By.XPath(TOTAL_VAT));
            return _siTotalVat.Text;
        }

        public void AddNewItemSiPriceAuto(string item, string qty, string taxType = null)
        {
            //input with search bar
            _itemName = WaitForElementIsVisible(By.XPath(ITEM_NAME));
            _itemName.SetValue(ControlType.TextBox, item);
            WaitForLoad();

            var searchItemValue = WaitForElementIsVisible(By.XPath("//div[text()='" + item + "']"));
            searchItemValue.Click();
            WaitForLoad();

            _itemName.SendKeys(Keys.Tab);
            WaitForLoad();

            _itemQuantity = WaitForElementIsVisible(By.XPath(ITEM_QUANTITY));
            _itemQuantity.SetValue(ControlType.TextBox, qty);
            // si pas de price alors pas de Footer
            WaitForLoad();

            if (taxType != null)
            {
                _itemTaxType = WaitForElementIsVisible(By.Id(TAX_TYPE));
                _itemTaxType.SetValue(ControlType.DropDownList, taxType);
                WaitForLoad();
            }
        }
        public void SubmitBtn()
        {
            _saveNewRow = WaitForElementIsVisible(By.Id(SAVE_NEW_ROW));
            _saveNewRow.Click();
            WaitPageLoading();
            WaitForLoad();
        }
        public void ShowBtnPlus()
        {
            _createNewRow = WaitForElementIsVisible(By.Id(CREATE_NEW_ROW));
            _createNewRow.Click();
            WaitPageLoading();
            WaitForLoad();
        }
        public string GetSI_PackPriceAdded()
        {
            _siPackPricePopUp = WaitForElementIsVisible(By.Id(SI_PRICE_POP_UP));

            return _siPackPricePopUp.GetAttribute("value");
        }

        public bool ItemIsVerifiedShow()
        {
            var _notVerifiedShow = isElementVisible(By.XPath("//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[1]//i[contains(@class, 'fa-eye-slash')]"));
            if (_notVerifiedShow)
            {
                return true;
            }

            else return false;


        }

        public bool ItemIsNotVerifiedShow()
        {
            var _verifiedShow = isElementVisible(By.XPath("//*[@id='list-item-with-action']/table/tbody/tr[*]/td[1]//i[@class='fa-solid fa-eye']"));
            if (_verifiedShow)
            {
                return true;
            }

            else return false;
        }

        public string GetQuantityInput()
        {
            var qtyElement = WaitForElementIsVisible(By.Id("quantityInput"));

            return qtyElement.GetAttribute("value");
        }

        public bool IsCommentGreen()
        {
            if (isElementExists(By.XPath(ICON_COMMENT_GREEN)))
            {
                return true;
            }
            return false;
        }
        public bool IsSupplierInvoiceActivated()
        {
            var isActive = WaitForElementExists(By.Id(IS_ACTIVE));
            return isActive.GetAttribute("checked") == "true";
             
        }
    }
}