using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Accounting.SupplierInvoices;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.Claims;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.ReceiptNotes
{
    public class ReceiptNotesItem : PageBase
    {

        public ReceiptNotesItem(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }
        //______________________________________ Constantes ____________________________________________________

        // General 
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";
        private const string RECEPT_NOTE_NUMBER = "//h1";
        private const string EXTENDED_BUTTON = "//*[@id=\"div-body\"]/div/div[1]/div/div/button";
        private const string REFRESH = "btn-receipt-notes-refresh";
        private const string PRINT = "print_receipt_note";
        private const string CONFIRM_PRINT = "printButton";
        private const string SEND_BY_MAIL = "btn-send-by-email-receipt-note";
        private const string SEND_BY_MAIL_POPUP = "//*[@id=\"modal-1\"]";
        private const string EMAIL_RECEIVER = "ToAddresses";
        private const string SEND_BUTTON = "btn-init-async-send-mail";
        private const string GENERATE_SI = "btn-generate-associated-supplier-invoice";
        private const string GENERATE_OF = "//*[contains(following-sibling::text(),'Generate Output Form')]/parent::a";
        private const string INVOICE_NUMBER = "tb-new-supplier-invoice-number";
        private const string CREATE_FROM = "check-box-create-from";
        private const string SUBMIT_SI = "btn-submit-form-create-supplier-invoice";
        private const string EXPORT_FOR_SAGE = "btn-export-for-sage";
        private const string CONFIRM_EXPORT_SAGE = "btn-validate-export";
        private const string ENABLE_EXPORT_SAGE = "btn-enable-export-for-sage";
        private const string ADD_NEW_CLAIM = "//span[contains(text(), '{0}')]/../..//a[contains(@id, 'btnAddEditClaim_')]";
        private const string EXPIRY_DATE = "//*[@id=\"itemForm_0\"]/div[1]/div[15]/a";
        private const string CANCEL_RN = "btn-cancel-rn";
        private const string CONFIRM_CANCEL_RN_DEV = "btn-popup-validate";
        private const string CONFIRM_CANCEL_RN = "dataConfirmOK";
        private const string EDIT_ITEM_DEV = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[starts-with(@title,'{0}')]/../../div[17]/ul/li[*]/../../ul/li[*]/a/span[@class = \"fas  fa-pencil-alt glyph-span\"]";
        private const string ADD_COMMENT_BTN_1 = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[@title='{0}']/../../div[15]/a";
        private const string VALIDATE = "btn-validate-receipt-note";
        private const string CONFIRM_VALIDATE = "btn-popup-validate";

        private const string HREF_LINK_TO_CANCELLATION = "hrefLinkToCancellation";
        private const string HREF_LINK_TO_CANCELLED_RN = "hrefLinkToCancelledRN";

        // Onglets
        private const string GENERAL_INFORMATION_TAB = "hrefTabContentInformations";
        private const string ITEMS_TAB = "hrefTabContentItems";
        private const string QUALITY_CHECKS_TAB = "hrefTabContentChecks";
        private const string FOOTER_TAB = "hrefTabContentFooter";
        private const string ACCOUNTING_TAB = "hrefTabContentExportSageWriting";

        // Tableau
        private const string GROUP = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/span";
        private const string ITEM = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[2]";
        private const string FIRST_ITEM_NAME = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[2]/div[3]/span";
        private const string OTHER_ITEM_NAME = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[{0}]/div/div/form/div[2]/div[3]/span";
        private const string ITEM_NAME = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[2]/div[3]/span[@title='{0}']";
        private const string ITEM_RECEIVED_QUANTITY = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[2]/div[7]/span";
        private const string ITEM_PRICE = "//*[starts-with(@id,\"itemForm_\")]/div[2]/div[9]/span";//"//*[@id='itemForm_0']/div[1]/div[11]/span";
        private const string FORM_ITEM_PACKAGING = "//*[starts-with(@id,\"itemForm_\")]/div[1]/div[5]/span";
        private const string FORM_ITEM_PACK_PRICE = "//*[@id=\"itemForm_0\"]/div[2]/div[6]/span";
        private const string FORM_ITEM_RECEIVED = "//*[@id=\"itemForm_0\"]/div[2]/div[8]/span";
        private const string RECEIVED_QUANTITY_INPUT = "//div[@class='edit-zone edit-zone-like-display']/div/span[@class='item-detail-col-name-span' and starts-with(@title,'{0}')]/../../div[8]/input[@id=\"item_RndRowDto_ReceivedQuantity\"]";
        private const string CLAIM_INPUT = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[@title='{0}']/../../div[9]/div/input";


        private const string FORM_COMMENT_ICON = "//*[@id=\"tabContentServiceContainer\"]/div[1]/div[2]/div/div/div[12]/span/i";
        private const string FORM_PICTURE_ICON = "//*[@id=\"tabContentServiceContainer\"]/div[1]/div[2]/div/div/div[13]/span/i";
        private const string FORM_CLAIM_ICON = "//*/form[@id=\"itemForm_0\"]/div[1]/div[13]/div/a/span";
        private const string FORM_EXPIRY_ICON = "//*/form[@id=\"itemForm_0\"]/div[1]/div[15]/a/span[1]";
        private const string FORM_TEMPERATURE_ICON = "//*/form[@id=\"itemForm_0\"]/div[1]/div[13]/div/a/span";
        private const string TEMPERATURE_ICON = "//*[@name='IconTemperature']";
        private const string TEMPERATURE_BUTTON = "//span[@title='AGITADOR TRANSPARENTE']/../..//span[@name = 'IconTemperature']";
        private const string TEMPERATURE_INPUT = "Temperature";
        private const string TEMPERATURE_DUPLICATE = "CopyTemperature";
        private const string TEMPERATURE_SAVE = "btnSubmit";
        private const string ITEM_FORMULA_PRICE = "//*[@id='row0']//input[contains(@id,'rn-detail-price-')]";
        private const string ITEM_FORMULA_QTY = "//*[@id='row0']//input[contains(@id,'rn-detail-qty-')]";
        private const string ITEM_FORMULA_TOTAL = "//*[@id=\"itemForm_0\"]/div[1]/div[12]/span";
        private const string ITEM_FORMULA_TOTAL_SUM = "dn-total-price-span";

        private const string ITEM_EXPIRY_DATE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[2]/div[11]/span";

        private const string EXTENDED_MENU = "//span[@title='{0}']/../..//a[@class='mini-btn open-item-packaging']";
        private const string COMMENT_ICON = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[@title='{0}']/../../div[15]/ul/li[*]/a/span[@class=\"glyphicon glyphicon-comment glyph-span\"]";
        private const string COMMENT = "Comment";
        private const string SAVE_COMMENT = "//*[@id=\"modal-1\"]/div/div/div/div[2]/div/form/div[2]/button[2]";
        private const string SAVE_COMMENT_2 = "//*[@id=\"formSaveComment\"]/div[2]/button[2]";

        private const string ITEM_EDIT = "//span[@title='{0}']/../..//span[@class='fas  fa-pencil-alt glyph-span']";
        private const string EDITCLAIM = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[contains(@title,'{0}')]/../../div[17]/ul/li[3]/a";
        private const string ETC_BUTTON = "//*[@id=\"itemForm_0\"]/div[1]/div[15]/a";
        private const string OPTIONS_ICON = "//*[@id=\"itemForm_0\"]/div[1]/div[17]/a";
        private const string ALLERGENS_BTN = "//div[contains(@class, 'display-zone')]//span[@title='{0}']/../..//span[@name='IconAllergens']";
        private const string ALLERGENS_LIST = "/html/body/div[4]/div/div/div/div[2]/div/div/div/div/ul";

        // Filtres
        private const string RESET_FILTER = "ResetFilter";
        private const string FILTER_NAME = "tbSearchPatternWithAutocomplete";
        private const string FIRST_RESULT_SEARCH = "//*[@id=\"formSearchItems\"]/div[2]/span/div/div/div/strong[text()='{0}']";
        private const string FILTER_NOT_RECEIVED = "ShowNullQuantity";
        private const string FILTER_KEYWORD = "ItemIndexVMSelectedKeywords_ms";
        private const string UNSELECT_ALL_KEYWORD = "/html/body/div[11]/div/ul/li[2]/a";
        private const string SEARCH_KEYWORD = "/html/body/div[11]/div/div/label/input";
        private const string FILTER_GROUP = "ItemIndexVMSelectedGroups_ms";
        private const string SEARCH_GROUP = "/html/body/div[12]/div/div/label/input";
        private const string UNSELECT_ALL = "/html/body/div[12]/div/ul/li[2]/a/span[2]";
        private const string FILTER_SUBGROUP = "ItemIndexVMSelectedSubGroups_ms";
        private const string UNSELECT_ALL_SUBGROUPS = "/html/body/div[13]/div/ul/li[2]/a/span[2]";
        private const string EDIT_FIRST_ROW = "/html/body/div[3]/div/div[3]/div[1]/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/form/div[1]/div[17]/ul/li[2]/a";
        private const string SANCTION_AMOUNT = "SanctionAmount";
        private const string BTN_SAVE = "btn-valid-claim";
        private const string ITEMNAME_CLAIMS = "//*[@id=\"tabContentServiceContainer\"]/div[1]/div[2]/div/div/div[1]/span";

        //_____________________________________ Variables ____________________________________________________

        // Général
        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;

        [FindsBy(How = How.XPath, Using = EXTENDED_BUTTON)]
        private IWebElement _extendedButton;

        [FindsBy(How = How.Id, Using = REFRESH)]
        private IWebElement _refresh;

        [FindsBy(How = How.XPath, Using = RECEPT_NOTE_NUMBER)]
        private IWebElement _receiptNoteNb;

        [FindsBy(How = How.Id, Using = PRINT)]
        private IWebElement _print;

        [FindsBy(How = How.Id, Using = CONFIRM_PRINT)]
        private IWebElement _confirmPrint;

        [FindsBy(How = How.Id, Using = SEND_BY_MAIL)]
        private IWebElement _sendByEmailBtn;

        [FindsBy(How = How.Id, Using = EMAIL_RECEIVER)]
        private IWebElement _emailReceiver;

        [FindsBy(How = How.Id, Using = SEND_BUTTON)]
        private IWebElement _sendButton;

        [FindsBy(How = How.Id, Using = GENERATE_SI)]
        private IWebElement _generateSI;

        [FindsBy(How = How.Id, Using = GENERATE_OF)]
        private IWebElement _generateOF;

        [FindsBy(How = How.Id, Using = INVOICE_NUMBER)]
        private IWebElement _invoiceNumber;

        [FindsBy(How = How.Id, Using = CREATE_FROM)]
        private IWebElement _createFrom;

        [FindsBy(How = How.Id, Using = SUBMIT_SI)]
        private IWebElement _submitSI;

        [FindsBy(How = How.Id, Using = EXPORT_FOR_SAGE)]
        private IWebElement _exportForSage;

        [FindsBy(How = How.Id, Using = CONFIRM_EXPORT_SAGE)]
        private IWebElement _confirmExportSage;

        [FindsBy(How = How.Id, Using = ENABLE_EXPORT_SAGE)]
        private IWebElement _enableExportSage;

        [FindsBy(How = How.Id, Using = VALIDATE)]
        private IWebElement _validate;

        [FindsBy(How = How.Id, Using = CONFIRM_VALIDATE)]
        private IWebElement _confirmValidate;

        [FindsBy(How = How.XPath, Using = ADD_NEW_CLAIM)]
        private IWebElement _addNewClaim;

        [FindsBy(How = How.Id, Using = CANCEL_RN)]
        private IWebElement _cancelRN_Btn;

        [FindsBy(How = How.Id, Using = CONFIRM_CANCEL_RN)]
        private IWebElement _cancelRN_ConfirmBtn;

        // Onglets
        [FindsBy(How = How.Id, Using = GENERAL_INFORMATION_TAB)]
        private IWebElement _generalInformationTab;

        [FindsBy(How = How.Id, Using = QUALITY_CHECKS_TAB)]
        private IWebElement _qualityChecksTab;

        [FindsBy(How = How.Id, Using = FOOTER_TAB)]
        private IWebElement _footerTab;

        [FindsBy(How = How.Id, Using = ACCOUNTING_TAB)]
        private IWebElement _accountingTab;

        // Tableau
        [FindsBy(How = How.XPath, Using = GROUP)]
        private IWebElement _group;

        [FindsBy(How = How.XPath, Using = ITEM)]
        private IWebElement _item;

        [FindsBy(How = How.XPath, Using = FIRST_ITEM_NAME)]
        private IWebElement _itemName;

        [FindsBy(How = How.XPath, Using = ITEM_RECEIVED_QUANTITY)]
        private IWebElement _receivedQty;

        [FindsBy(How = How.XPath, Using = RECEIVED_QUANTITY_INPUT)]
        private IWebElement _receivedQtyInput;

        [FindsBy(How = How.XPath, Using = ITEM_PRICE)]
        private IWebElement _itemPrice;

        [FindsBy(How = How.XPath, Using = FORM_ITEM_PACKAGING)]
        private IWebElement _itemPackaging;

        [FindsBy(How = How.XPath, Using = FORM_ITEM_PACK_PRICE)]
        private IWebElement _itemPackPrice;

        [FindsBy(How = How.XPath, Using = FORM_ITEM_RECEIVED)]
        private IWebElement _itemReceived;

        [FindsBy(How = How.XPath, Using = ITEM_FORMULA_PRICE)]
        private IWebElement _itemFormulaPrice;

        [FindsBy(How = How.XPath, Using = ITEM_FORMULA_QTY)]
        private IWebElement _itemFormulaQty;

        [FindsBy(How = How.XPath, Using = ITEM_FORMULA_TOTAL)]
        private IWebElement _itemFormulaTotal;

        [FindsBy(How = How.XPath, Using = ITEM_FORMULA_TOTAL_SUM)]
        private IWebElement _itemFormulaTotalSum;

        [FindsBy(How = How.XPath, Using = CLAIM_INPUT)]
        private IWebElement _claimInput;

        [FindsBy(How = How.XPath, Using = TEMPERATURE_BUTTON)]
        private IWebElement _temperatureButton;

        [FindsBy(How = How.Id, Using = TEMPERATURE_INPUT)]
        private IWebElement _temperatureInput;

        [FindsBy(How = How.Id, Using = TEMPERATURE_SAVE)]
        private IWebElement _temperatureSave;

        [FindsBy(How = How.Id, Using = TEMPERATURE_ICON)]
        private IWebElement _temperatureIcon;

        [FindsBy(How = How.Id, Using = TEMPERATURE_DUPLICATE)]
        private IWebElement _temperatureDuplicate;

        [FindsBy(How = How.XPath, Using = FORM_TEMPERATURE_ICON)]
        private IWebElement _temperatureGreen;

        [FindsBy(How = How.XPath, Using = FORM_COMMENT_ICON)]
        private IWebElement _commentGreen;

        [FindsBy(How = How.XPath, Using = FORM_PICTURE_ICON)]
        private IWebElement _pictureGreen;

        [FindsBy(How = How.XPath, Using = FORM_CLAIM_ICON)]
        private IWebElement _claimGreen;

        [FindsBy(How = How.XPath, Using = FORM_EXPIRY_ICON)]
        private IWebElement _expiryGreen;

        [FindsBy(How = How.XPath, Using = ITEM_EXPIRY_DATE)]
        private IWebElement _expiryDate;

        [FindsBy(How = How.XPath, Using = EXTENDED_MENU)]
        private IWebElement _extendedMenu;

        [FindsBy(How = How.XPath, Using = COMMENT_ICON)]
        private IWebElement _addComment;

        [FindsBy(How = How.Id, Using = COMMENT)]
        private IWebElement _comment;

        [FindsBy(How = How.XPath, Using = SAVE_COMMENT)]
        private IWebElement _saveComment;

        [FindsBy(How = How.XPath, Using = ITEM_EDIT)]
        private IWebElement _editItem;

        [FindsBy(How = How.XPath, Using = ETC_BUTTON)]
        private IWebElement _etcButton;

        [FindsBy(How = How.Id, Using = ALLERGENS_BTN)]
        private IWebElement _allergensBtn;

        [FindsBy(How = How.Id, Using = ALLERGENS_LIST)]
        private IWebElement _allergensList;

        //_______________________________________  Filter _____________________________________________________________

        [FindsBy(How = How.XPath, Using = RESET_FILTER)]
        private IWebElement _resetFilter;

        [FindsBy(How = How.Id, Using = FILTER_NAME)]
        private IWebElement _searchByNameFilter;

        [FindsBy(How = How.Id, Using = FILTER_NOT_RECEIVED)]
        private IWebElement _showNotReceived;

        [FindsBy(How = How.Id, Using = SANCTION_AMOUNT)]
        private IWebElement _sanctionAmountInput;

        [FindsBy(How = How.Id, Using = BTN_SAVE)]
        private IWebElement _clickSave;
        public enum FilterItemType
        {
            SearchByName,
            ShowNotReceived,
            SearchByKeyword,
            ByGroup,
            BySubGroup
        }

        public void Filter(FilterItemType FilterItemType, object value)
        {
            switch (FilterItemType)
            {
                case FilterItemType.SearchByName:
                    WaitForElementIsVisibleNew(By.Id(FILTER_NAME));
                    var itemName = extractItemName((string)value);

                    _searchByNameFilter.SetValue(ControlType.TextBox, itemName);

                    if (isElementVisible(By.XPath(String.Format(FIRST_RESULT_SEARCH, itemName))))
                    {
                        var firstResultSearch = _webDriver.FindElement(By.XPath(String.Format(FIRST_RESULT_SEARCH, itemName)));
                        firstResultSearch.Click();
                    }
                    else
                    {
                        return;
                    }
                    break;
                case FilterItemType.SearchByKeyword:
                    ComboBoxSelectById(new ComboBoxOptions(FILTER_KEYWORD, (string)value));
                    break;
                case FilterItemType.ShowNotReceived:
                    _showNotReceived = WaitForElementExists(By.Id(FILTER_NOT_RECEIVED));
                    _showNotReceived.SetValue(ControlType.CheckBox, value);
                    break;
                case FilterItemType.ByGroup:
                    ComboBoxSelectById(new ComboBoxOptions(FILTER_GROUP, (string)value));
                    break;
                case FilterItemType.BySubGroup:
                    ComboBoxSelectById(new ComboBoxOptions(FILTER_SUBGROUP, (string)value));
                    break;
            }

            WaitPageLoading();
            WaitForLoad();
        }

        public void ResetFilters()
        {
            _resetFilter = WaitForElementIsVisibleNew(By.Id(RESET_FILTER));
            _resetFilter.Click();
            WaitForLoad();

            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                // pas de date
            }
        }

        public bool VerifyName(string value)
        {
            var elements = _webDriver.FindElements(By.XPath(FIRST_ITEM_NAME));

            if (elements.Count == 0)
                return false;

            foreach (var elm in elements)
            {
                if (elm.Text != value)
                {
                    return false;
                }
            }
            return true;
        }

        public bool VerifyKeyword(string keyword)
        {
            var boolValue = true;
            var elements = _webDriver.FindElements(By.XPath(FIRST_ITEM_NAME));

            if (elements.Count == 0)
                return false;

            int nbMax = elements.Count <= 8 ? elements.Count : 8;
            int compteur = 1;

            while (compteur <= nbMax)
            {
                foreach (var elm in elements)
                {
                    // clic sur la ligne
                    elm.Click();

                    var itemGeneralInformationPage = EditItem(elm.GetAttribute("title"));
                    var itemKeywordTab = itemGeneralInformationPage.ClickOnKeywordItem();

                    var itemKeyword = itemKeywordTab.IsKeywordPresent(keyword);

                    itemGeneralInformationPage.Close();

                    if (!itemKeyword)
                        return false;

                    compteur++;
                }
            }

            return boolValue;
        }

        public Boolean VerifyReceived()
        {
            var elements = _webDriver.FindElements(By.XPath(ITEM_RECEIVED_QUANTITY));

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

        public Boolean VerifyGroup(string group)
        {
            var boolValue = true;
            var elements = _webDriver.FindElements(By.XPath(FIRST_ITEM_NAME));

            if (elements.Count == 0)
                return false;

            int nbMax = elements.Count <= 8 ? elements.Count : 8;
            int compteur = 1;

            while (compteur <= nbMax)
            {
                foreach (var elm in elements)
                {
                    // clic sur la ligne
                    WaitForLoad();
                    ((IJavaScriptExecutor)_webDriver).ExecuteScript("arguments[0].scrollIntoView(true);", elm);
                    elm.Click();
                    WaitPageLoading();
                    WaitForLoad();

                    var itemGeneralInformationPage = EditItem(elm.GetAttribute("title"));
                    WaitForLoad();
                    var groupName = itemGeneralInformationPage.GetGroupName();
                    WaitForLoad();
                    itemGeneralInformationPage.Close();

                    // On ferme l'onglet ouvert
                    if (!groupName.Equals(group))
                        return false;
                    else compteur++;
                }
            }
            return boolValue;
        }


        // __________________________________  Méthodes _________________________________________________

        // Général
        public ReceiptNotesPage BackToList()
        {
            WaitPageLoading();
            WaitForLoad();
            _backToList = WaitForElementToBeClickable(By.XPath(BACK_TO_LIST));
            WaitForLoad();
            _backToList.Click();
            WaitPageLoading();

            return new ReceiptNotesPage(_webDriver, _testContext);
        }

        public override void ShowExtendedMenu()
        {
            // menu dynamique
            Thread.Sleep(2000);
            var actions = new Actions(_webDriver);
            _extendedButton = WaitForElementIsVisibleNew(By.XPath(EXTENDED_BUTTON));
            actions.MoveToElement(_extendedButton).Perform();
            WaitForLoad();
        }

        public void Refresh()
        {
            ShowExtendedMenu();

            _refresh = WaitForElementIsVisibleNew(By.Id(REFRESH));
            _refresh.Click();
            WaitForLoad();
        }
        public string GetReceiptNoteNumber()
        {
            _receiptNoteNb = WaitForElementIsVisible(By.XPath(RECEPT_NOTE_NUMBER), nameof(RECEPT_NOTE_NUMBER));
            return _receiptNoteNb.Text.Substring(_receiptNoteNb.Text.IndexOf('°') + 1);
        }

        public PrintReportPage PrintReceiptNote(bool printValue)
        {
            // Temps d'attente pour que le bouton soit disponible
            ShowExtendedMenu();

            _print = WaitForElementIsVisible(By.Id(PRINT));
            _print.Click();
            WaitForLoad();
            WaitPageLoading();
            _confirmPrint = WaitForElementIsVisible(By.Id(CONFIRM_PRINT));
            _confirmPrint.Click();
            WaitForLoad();
            WaitPageLoading();
            if (printValue)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-pdf']"));
            }

            //Results are opened in a new tab, switch the driver to the newly created one
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new PrintReportPage(_webDriver, _testContext);
        }

        public void SendByMail()
        {
            WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            ShowExtendedMenu();
            //WaitForLoad();
            _sendByEmailBtn = WaitForElementIsVisible(By.Id(SEND_BY_MAIL));
            _sendByEmailBtn.Click();
            WaitForLoad();
            _emailReceiver = WaitForElementIsVisible(By.Id(EMAIL_RECEIVER));
            _emailReceiver.SetValue(ControlType.TextBox, "test@mail.com");
            WaitForLoad();
            _sendButton = WaitForElementIsVisible(By.Id(SEND_BUTTON));
            _sendButton.Click();
             wait.Until(d => !isElementVisible(By.XPath(SEND_BY_MAIL_POPUP)));
            WaitPageLoading();
        }

        public ReceiptNotesItem CancelReceiptNote()
        {
            ShowExtendedMenu();

            _cancelRN_Btn = WaitForElementIsVisible(By.Id(CANCEL_RN));
            _cancelRN_Btn.Click();

            _cancelRN_ConfirmBtn = WaitForElementIsVisible(By.Id(CONFIRM_CANCEL_RN_DEV));

            _cancelRN_ConfirmBtn.Click();

            WaitForLoad();

            return new ReceiptNotesItem(_webDriver, _testContext);
        }

        public SupplierInvoicesItem GenerateSupplierInvoice(string supplierInvoiceNumber)
        {
            // Completed the modal
            ShowExtendedMenu();

            _generateSI = WaitForElementIsVisible(By.Id(GENERATE_SI));
            _generateSI.Click();
            WaitForLoad();

            _invoiceNumber = WaitForElementIsVisible(By.Id(INVOICE_NUMBER));
            _invoiceNumber.SetValue(ControlType.TextBox, supplierInvoiceNumber);
            WaitForLoad();

            _createFrom = _webDriver.FindElement(By.Id(CREATE_FROM));
            _createFrom.SetValue(ControlType.CheckBox, true);
            WaitForLoad();

            _submitSI = WaitForElementIsVisible(By.Id(SUBMIT_SI));
            _submitSI.Click();
            WaitForLoad();

            //Results are opened in a new tab, switch the driver to the newly created one
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(30));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new SupplierInvoicesItem(_webDriver, _testContext);
        }

        public void ExportResultsForSage(bool printValue)
        {
            ShowExtendedMenu();

            _exportForSage = WaitForElementIsVisible(By.Id(EXPORT_FOR_SAGE));
            _exportForSage.Click();
            WaitForLoad();

            _confirmExportSage = WaitForElementIsVisible(By.Id(CONFIRM_EXPORT_SAGE));
            _confirmExportSage.Click();
            WaitForLoad();

            if (printValue)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-alt']"));
                ClickPrintButton();
            }

            WaitForDownload();
            Close();

            if (printValue)
            {
                ClosePrintButton();
            }
        }

        public FileInfo GetExportForSageFile(FileInfo[] taskFiles)
        {
            List<FileInfo> correctDownloadFiles = new List<FileInfo>();

            foreach (var file in taskFiles)
            {
                //  Test REGEX
                if (IsSageFileCorrect(file.Name))
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

        public bool IsSageFileCorrect(string filePath)
        {
            // ReceiptNote 2020-07-29 10-17-17.txt

            string mois = "(?:0[1-9]|1[0-2])";         // mois
            string space = "(\\s)";                    // Espace
            string annee = "\\d{4}";                   // annee YYYY
            string jour = "[0-3]\\d";                  // jour
            string heure = "(?:0[0-9]|1[0-9]|2[0-3])"; // heure
            string minutes = "[0-5]\\d";               // minutes
            string secondes = "[0-5]\\d";              // secondes


            Regex r = new Regex("^ReceiptNote" + space + annee + "-" + mois + "-" + jour + space + heure + "-" + minutes + "-" + secondes + ".txt$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = r.Match(filePath);

            return m.Success;
        }

        public bool CanClickOnSAGE()
        {
            ShowExtendedMenu();

            _exportForSage = WaitForElementExists(By.Id(EXPORT_FOR_SAGE));
            if (_exportForSage.GetAttribute("disabled") != null)
                return false;

            return true;
        }

        public bool CanClickOnEnableSAGE()
        {
            ShowExtendedMenu();

            _enableExportSage = WaitForElementIsVisible(By.Id(ENABLE_EXPORT_SAGE));
            if (_enableExportSage.GetAttribute("disabled") != null)
                return false;

            return true;
        }

        public void ClickOnEnableSAGE()
        {
            ShowExtendedMenu();

            _enableExportSage = WaitForElementIsVisible(By.Id(ENABLE_EXPORT_SAGE));
            _enableExportSage.Click();
            WaitForLoad();
        }

        public void Validate()
        {
            ShowValidationMenu();
            WaitForLoad();
            _validate = WaitForElementIsVisibleNew(By.Id(VALIDATE));
            _validate.Click();
            WaitForLoad();
            _confirmValidate = WaitForElementIsVisibleNew(By.Id(CONFIRM_VALIDATE));
            _confirmValidate.Click();
            WaitForLoad();
            WaitLoading();
        }

        public ClaimsItem ValidateClaim()
        {
            ShowValidationMenu();

            _validate = WaitForElementIsVisibleNew(By.Id(VALIDATE));
            _validate.Click();
            WaitForLoad();

            _confirmValidate = WaitForElementIsVisibleNew(By.Id(CONFIRM_VALIDATE));
            _confirmValidate.Click();

            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            WaitForLoad();

            return new ClaimsItem(_webDriver, _testContext);
        }

        // Onglets
        public ReceiptNotesGeneralInformation ClickOnGeneralInformationTab()
        {
            _generalInformationTab = WaitForElementIsVisible(By.Id(GENERAL_INFORMATION_TAB));
            _generalInformationTab.Click();
            WaitForLoad();

            return new ReceiptNotesGeneralInformation(_webDriver, _testContext);
        }

        public ReceiptNotesItem ClickOnItemsTab()
        {
            WaitPageLoading();
            var itemTab = WaitForElementIsVisible(By.Id(ITEMS_TAB));
            itemTab.Click();
            WaitForLoad();

            return new ReceiptNotesItem(_webDriver, _testContext);
        }

        public ReceiptNotesQualityChecks ClickOnChecksTab()
        {
            _qualityChecksTab = WaitForElementIsVisible(By.Id(QUALITY_CHECKS_TAB));
            _qualityChecksTab.Click();
            WaitForLoad();

            return new ReceiptNotesQualityChecks(_webDriver, _testContext);
        }

        public ReceiptNotesFooterPage ClickOnFooterTab()
        {
            _footerTab = WaitForElementIsVisible(By.Id(FOOTER_TAB));
            _footerTab.Click();
            WaitForLoad();

            return new ReceiptNotesFooterPage(_webDriver, _testContext);
        }

        public ReceiptNotesAccountingPage ClickOnAccountingTab()
        {
            _accountingTab = WaitForElementIsVisible(By.Id(ACCOUNTING_TAB));
            _accountingTab.Click();
            WaitForLoad();

            return new ReceiptNotesAccountingPage(_webDriver, _testContext);
        }

        // Tableau
        public void SelectFirstItem()
        {
            _item = WaitForElementIsVisibleNew(By.XPath(ITEM));
            _item.Click();
            WaitForLoad();
        }

        /// <summary>
        /// Sélection de la ligne de détail RN
        /// </summary>
        /// <param name="offset">Le numéro de la ligne, commençant à 1</param>
        public void SelectItemRow(int offset)
        {
            _item = WaitForElementIsVisible(By.Id(String.Format("row{0}", offset - 1)));
            _item.Click();
            WaitForLoad();
        }

        public void SelectItem(string itemName)
        {
            _item = WaitForElementIsVisible(By.XPath(String.Format(ITEM_NAME, itemName)));
            _item.Click();
            WaitForLoad();
        }
        public string GetItem(string itemName)
        {
            _itemName = WaitForElementExists(By.XPath(String.Format("//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[2]/div[3]/span[@title='{0}']", itemName)));
            return _itemName.Text;
        }
        public string GetFirstItemName()
        {
            WaitPageLoading();
            _itemName = WaitForElementExists(By.XPath(FIRST_ITEM_NAME));
            WaitPageLoading();
            return _itemName.GetAttribute("title");
        }

        /// <summary>
        /// Returns the item name of the unselected row number. /!\ The row MUST be unselected.
        /// </summary>
        /// <param name="offset">The row number, starts at 1.</param>
        /// <returns></returns>
        public string GetItemName(int offset)
        {
            _itemName = WaitForElementExists(By.XPath(string.Format(OTHER_ITEM_NAME, offset + 1)));
            return _itemName.GetAttribute("title");
        }

        public string extractItemName(string itemName)
        {
            if (itemName.Contains("("))
            {
                // PLATO PLASTICO NEUTRAL 150x90 (PLG00014)
                return itemName.Substring(0, itemName.IndexOf("(") - 1);
            }
            return itemName;
        }


        public string GetFirstGroupName()
        {
            _group = WaitForElementExists(By.XPath(GROUP));
            return _group.Text;
        }

        public bool IsGroupDisplayActive()
        {
            if (isElementVisible(By.XPath(GROUP)))
            {
                _group = _webDriver.FindElement(By.XPath(GROUP));
                return true;
            }
            else
            {
                return false;
            }
        }

        public void AddReceived(string itemName, string quantity)
        {
           WaitPageLoading();
            _receivedQtyInput = WaitForElementExists(By.XPath(string.Format(RECEIVED_QUANTITY_INPUT, itemName)));
            WaitPageLoading();
            WaitLoading();
            _receivedQtyInput.SetValue(ControlType.TextBox, quantity);
            WaitLoading();
        }
        public void AddReceived_receipt(string itemName, string quantity)
        {
            WaitPageLoading();
            _receivedQtyInput = WaitForElementExists(By.Id("item_RndRowDto_ReceivedQuantity"));
            WaitPageLoading();
            WaitLoading();
            _receivedQtyInput.SetValue(ControlType.TextBox, quantity);
            WaitLoading();
        }
        public void AddReicedQuantity(string itemName, string quantity)
        {
            WaitPageLoading();
            _receivedQtyInput = WaitForElementExists(By.XPath(string.Format("//*[starts-with(@id,\"itemForm_\")]/div[1]/div[3]/span[contains(text(),'{0}')]/../../div[8]/input", itemName)));
            _receivedQtyInput.SetValue(ControlType.TextBox, quantity);
            WaitPageLoading();
        }

        public string GetReceived()
        {
            if (IsDev()) _receivedQty = WaitForElementExists(By.XPath(ITEM_RECEIVED_QUANTITY));
            else _receivedQty = WaitForElementExists(By.XPath("//*[starts-with(@id,'itemForm_')]/div[2]/div[8]/span"));
            return _receivedQty.Text;
        }

        public bool IsModifiableReceivedValue(string itemName)
        {
            if (isElementVisible(By.XPath(string.Format(RECEIVED_QUANTITY_INPUT, itemName))))
            {
                _receivedQtyInput = WaitForElementIsVisible(By.XPath(string.Format(RECEIVED_QUANTITY_INPUT, itemName)));
                return true;
            }
            else
            {
                return false;
            }
        }

        public void CheckClaim(string itemName)
        {
            _claimInput = _webDriver.FindElement(By.XPath(String.Format(CLAIM_INPUT, itemName)));
            _claimInput.SetValue(ControlType.CheckBox, true);
            Thread.Sleep(2000);
        }

        public void AddTemperature(int temp)
        {
            _temperatureButton = WaitForElementIsVisible(By.XPath(TEMPERATURE_ICON));
            _temperatureButton.Click();


            _temperatureInput = WaitForElementIsVisible(By.Id("temperatureInput"));

            _temperatureInput.SetValue(ControlType.TextBox, temp.ToString());

            _temperatureSave = WaitForElementIsVisible(By.Id(TEMPERATURE_SAVE));
            _temperatureSave.Click();

            WaitForLoad();
        }

        public void AddComment(string comment)
        {
            IWebElement menuButton = WaitForElementIsVisible(By.XPath("//*[starts-with(@id,\"itemForm_\")]/div[1]/div[17]/a"));
            menuButton.Click();
            WaitForLoad();
            _addComment = WaitForElementIsVisible(By.XPath("//*[starts-with(@id,\"item-detail-\")]/li[1]/a"));
            _addComment.Click();
            WaitForLoad();
            _comment = WaitForElementIsVisible(By.Id(COMMENT));
            _comment.SetValue(ControlType.TextBox, comment);
            _saveComment = WaitForElementToBeClickable(By.XPath(SAVE_COMMENT_2));
            _saveComment.Click();
            WaitForLoad();
        }

        public ReceiptNotesEditClaim AddClaim()
        {
            //ADD_NEW_CLAIM
            _addNewClaim = WaitForElementIsVisible(By.Id("btnAddEditClaim_0"));
            _addNewClaim.Click();
            WaitPageLoading();
            return new ReceiptNotesEditClaim(_webDriver, _testContext);
        }

        public string GetComment(string itemName)
        {
            //_addComment = WaitForElementIsVisible(By.XPath(string.Format("//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[@title='{0}']/../../div[15]/a", itemName)));
            //_addComment.Click();
            //WaitForLoad();
            IWebElement menuButton = WaitForElementIsVisible(By.XPath("//*[starts-with(@id,\"itemForm_\")]/div[1]/div[17]/a"));
            menuButton.Click();
            WaitForLoad();
            _addComment = WaitForElementIsVisible(By.XPath("//*[starts-with(@id,\"item-detail-\")]/li[1]/a"));
            _addComment.Click();
            WaitForLoad();
            string comment = "";
            if (isElementVisible(By.Id(COMMENT)))
            {
                _comment = WaitForElementExists(By.Id(COMMENT));
                comment = _comment.Text;
            }
            IWebElement cancel = WaitForElementIsVisible(By.XPath("//*/button[text()='Cancel']"));
            cancel.Click();
            return comment;
        }

        public ItemGeneralInformationPage EditItem(string itemName)
        {
            _extendedMenu = WaitForElementIsVisible(By.XPath(string.Format("//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[starts-with(@title,'{0}')]/../../div[17]/a", itemName)));
            _extendedMenu.Click();
            WaitForLoad();
            _editItem = WaitForElementIsVisible(By.XPath(string.Format(EDIT_ITEM_DEV, itemName)));
            _editItem.Click();
            WaitForLoad();

            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            WaitForLoad();

            return new ItemGeneralInformationPage(_webDriver, _testContext);
        }


        public ClaimEditClaimForm EditClaimForm(string itemName)
        {

            var extendMenuItem = WaitForElementIsVisible(By.XPath(String.Format(EDITCLAIM, itemName)));
            extendMenuItem.Click();
            WaitForLoad();


            return new ClaimEditClaimForm(_webDriver, _testContext);

        }

        public ReceiptNoteExpiry ShowFirstExpiryDate()
        {
            //modifier PhysQty change l'emplacement
            _expiryDate = WaitForElementIsVisible(By.XPath(EXPIRY_DATE));
            _expiryDate.Click();

            WaitForLoad();

            return new ReceiptNoteExpiry(_webDriver, _testContext);
        }

        public ReceiptNoteToOuputForm GenerateOutputForm()
        {
            // Completed the modal
            ShowExtendedMenu();

            _generateOF = WaitForElementIsVisible(By.XPath(GENERATE_OF));
            _generateOF.Click();
            WaitForLoad();

            return new ReceiptNoteToOuputForm(_webDriver, _testContext);
        }

        public string GetPrice()
        {
            _itemPrice = WaitForElementIsVisible(By.XPath(ITEM_PRICE));
            return _itemPrice.Text;
        }

        public bool TemperatureGreen()
        {
            WaitForLoad();
            _temperatureIcon = WaitForElementExists(By.XPath(TEMPERATURE_ICON));
            return _temperatureIcon.GetAttribute("class").Contains("green-text");
        }

        public bool FormTemperatureGreen()
        {
            _temperatureGreen = WaitForElementIsVisible(By.XPath(FORM_TEMPERATURE_ICON));
            return _temperatureGreen.GetAttribute("class").Contains("green-text");
        }

        public bool FormClaimGreen()
        {
            _claimGreen = WaitForElementIsVisible(By.XPath(FORM_CLAIM_ICON));
            return _claimGreen.GetAttribute("class").Contains("green-text");
        }

        public bool FormCommentGreen()
        {
            _commentGreen = WaitForElementIsVisible(By.XPath(FORM_COMMENT_ICON));
            return _commentGreen.GetAttribute("class").Contains("green-text");
        }

        public bool FormPictureGreen()
        {
            _pictureGreen = WaitForElementIsVisible(By.XPath(FORM_PICTURE_ICON));
            return _pictureGreen.GetAttribute("class").Contains("green-text");
        }

        public bool FormExpiryGreen()
        {
            _expiryGreen = WaitForElementIsVisible(By.XPath(FORM_EXPIRY_ICON));
            return _expiryGreen.GetAttribute("class").Contains("green-text");
        }

        public bool TemperatureAllGreen()
        {
            WaitForLoad();
            var allTemperature = _webDriver.FindElements(By.XPath(TEMPERATURE_ICON));
            foreach (var temperature in allTemperature)
            {
                if (!temperature.GetAttribute("class").Contains("green-text"))
                {
                    return false;
                }
            }
            return true;
        }

        public void TemperatureIconClick(int temperature, bool duplicateAll = false)
        {
            _temperatureIcon = WaitForElementIsVisible(By.XPath(TEMPERATURE_ICON));
            _temperatureIcon.Click();
            _temperatureInput = WaitForElementIsVisible(By.Id("temperatureInput"));

            _temperatureInput.Clear();
            _temperatureInput.SendKeys(temperature.ToString());
            if (duplicateAll)
            {
                _temperatureDuplicate = WaitForElementIsVisible(By.Id(TEMPERATURE_DUPLICATE));
                _temperatureDuplicate.Click();
            }
            _temperatureSave = WaitForElementIsVisible(By.Id(TEMPERATURE_SAVE));
            _temperatureSave.Click();
            WaitForLoad();
        }

        public double GetDNPrice(string decimalSeparator)
        {
            _itemFormulaPrice = WaitForElementIsVisible(By.XPath(ITEM_FORMULA_PRICE));
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");
            return double.Parse(_itemFormulaPrice.GetAttribute("value"), ci);
        }

        public double GetDNQty(string decimalSeparator)
        {
            _itemFormulaQty = WaitForElementIsVisible(By.XPath(ITEM_FORMULA_QTY));
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");
            return double.Parse(_itemFormulaQty.GetAttribute("value"), ci);
        }

        public double GetDNTotal(string decimalSeparator)
        {
            _itemFormulaTotal = WaitForElementIsVisible(By.XPath(ITEM_FORMULA_TOTAL));
            // "€ 1 026,0700"
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");
            return double.Parse(_itemFormulaTotal.Text.Replace("€", "").Replace(" ", ""), ci);
        }

        public double GetDNTotalSum(string decimalSeparator)
        {
            _itemFormulaTotalSum = WaitForElementIsVisible(By.Id(ITEM_FORMULA_TOTAL_SUM));
            // "€ 1 026,0700"
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");
            return double.Parse(_itemFormulaTotalSum.Text.Replace("€", "").Replace(" ", ""), ci);
        }

        public void DeleteItem()
        {
            _etcButton = WaitForElementIsVisible(By.XPath("//*[@id=\"itemForm_0\"]/div[1]/div[17]/a"));
            _etcButton.Click();
            var deleteButton = WaitForElementIsVisible(By.XPath("//span[contains(@class, 'trash')]"));
            deleteButton.Click();
        }

        public string GetFormPackaging()
        {
            _itemPackaging = WaitForElementIsVisible(By.XPath(FORM_ITEM_PACKAGING));
            return _itemPackaging.Text;
        }

        public double GetFormPackPrice(string decimalSeparator)
        {
            _itemPackPrice = WaitForElementIsVisible(By.XPath(FORM_ITEM_PACK_PRICE));
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");
            return double.Parse(_itemPackPrice.Text.Replace("€", "").Replace(" ", ""), ci);
        }

        public string GetFormReceived()
        {
            _itemReceived = WaitForElementExists(By.XPath(FORM_ITEM_RECEIVED));
            return _itemReceived.GetAttribute("value");
        }

        public bool ExistsLinkToCancelledRN()
        {
            if (isElementVisible(By.Id(HREF_LINK_TO_CANCELLED_RN)))
            {
                var linkToCancelledRN_Href = WaitForElementToBeClickable(By.Id(HREF_LINK_TO_CANCELLED_RN));
                return true;
            }
            else
            {
                return false;
            }
        }

        public void SetFirstReceivedQuantity(string qty)
        {
            SelectFirstItem();
            WaitForLoad();
            IWebElement receivedInput;
            receivedInput = WaitForElementIsVisible(By.Id("item_RndRowDto_ReceivedQuantity"));
            receivedInput.Clear();
            receivedInput.SendKeys(qty);
            Thread.Sleep(2000);
            WaitForLoad();
        }

        public bool VerifyDetailChangeLine()
        {
            Random random = new Random();
            string rnd = random.Next(0, 200).ToString();
            string input;
            string nextInput;
            bool result = false;
            var receivedInputs = _webDriver.FindElements(By.Id("item_RndRowDto_ReceivedQuantity"));
            for (var i = 0; i < 6; i = i + 1)
            {
                var receivedInput = receivedInputs[i];

                receivedInput.SendKeys(rnd);
                input = receivedInput.GetAttribute("value");
                receivedInput.SendKeys(Keys.ArrowDown);
                nextInput = receivedInputs[i + 1].GetAttribute("value");

                if (input == nextInput)
                {
                    result = false;
                }
                result = true;
            }
            return result;
        }
        public void Go_To_New_Navigate()
        {
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);
        }
        public void OpenModalToEditFirstItem()
        {
            var options = WaitForElementIsVisible(By.XPath(OPTIONS_ICON));
            options.Click();
            WaitForLoad();
            var edit_icon = WaitForElementIsVisible(By.XPath(EDIT_FIRST_ROW));
            edit_icon.Click();
            WaitForLoad();
        }
        public double GetRNTotal(string decimalSeparator, string currency)
        {
            WaitForLoad();
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            _itemFormulaTotal = WaitForElementIsVisible(By.Id(ITEM_FORMULA_TOTAL_SUM));
            var sanctionAmountText = _itemFormulaTotal.Text;

            var format = (NumberFormatInfo)ci.NumberFormat.Clone();
            format.CurrencySymbol = currency;
            var mynumber = Decimal.Parse(sanctionAmountText, NumberStyles.Currency, format);

            return Convert.ToDouble(mynumber, ci);
        }

        public bool IsAllergenIconGreen(string itemName)
        {
            _allergensBtn = WaitForElementIsVisible(By.XPath(string.Format(ALLERGENS_BTN, itemName)));
            IWebElement icon = _allergensBtn.FindElement(By.XPath(string.Format(ALLERGENS_BTN, itemName)));
            //var icon = _allergensBtn.FindElement(By.TagName("span"));
            string iconClass = icon.GetAttribute("class");
            return iconClass.Contains("green") ? true : false;
        }

        public List<string> GetAllergens(string itemName)
        {
            List<string> item_allergens = new List<string>();
            _allergensBtn = WaitForElementIsVisible(By.XPath(string.Format("//div[contains(@class, 'display-zone')]//span[@title='{0}']/../..//span[@name='IconAllergens']", itemName)));
            _allergensBtn.Click();

            _allergensList = WaitForElementIsVisible(By.XPath(string.Format(ALLERGENS_LIST)));
            var allergensInList = _allergensList.FindElements(By.TagName("li"));
            if (allergensInList.Count > 0)
            {
                foreach (IWebElement allergen in allergensInList)
                {
                    //if (allergen.FindElement(By.TagName("img")) == null) continue;
                    if (allergen.FindElements(By.TagName("img")).Count != 0) item_allergens.Add(allergen.Text);
                   
                }
            }
            return item_allergens;
        }

        public ItemGeneralInformationPage EditItemGeneralInformation(string itemName)
        {
            WaitPageLoading();
            _extendedMenu = WaitForElementIsVisible(By.XPath(string.Format(EXTENDED_MENU, itemName)));
            _extendedMenu.Click();
            WaitPageLoading();
            _editItem = WaitForElementIsVisible(By.XPath(string.Format(ITEM_EDIT, itemName)));
            _editItem.Click();
            WaitForLoad();

            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            WaitForLoad();

            return new ItemGeneralInformationPage(_webDriver, _testContext);
        }

        public bool IsIconDisplayed()
        {
            var iconElement = _webDriver.FindElement(By.XPath("//*[@id=\"itemForm_0\"]/div[2]/div[1]/span"));
            return iconElement.Displayed;
        }

        public string GetFirstItemNameText()
        {
            _itemName = WaitForElementExists(By.XPath(FIRST_ITEM_NAME));
            return _itemName.Text;
        }
        public bool FormMegaphoneGreen()
        {
            var _megaphoneGreen = WaitForElementIsVisible(By.XPath("//*[@id=\"btnAddEditClaim_0\"]/span"));
            return _megaphoneGreen.GetAttribute("class").Contains("green-text");
        }
        public double GetDNTotalItem(string decimalSeparator)
        {
            _itemFormulaTotal = WaitForElementIsVisible(By.XPath("//*[@id=\"itemForm_0\"]/div[1]/div[12]/span"));
            // "€ 1 026,0700"
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");
            return double.Parse(_itemFormulaTotal.Text.Replace("€", "").Replace(" ", ""), ci);
        }
        public void ShowThreePoint()
        {
            var _threePoint = WaitForElementIsVisible(By.XPath("//*[@id=\"itemForm_0\"]/div[1]/div[17]/a/span"));
            _threePoint.Click();
            WaitPageLoading();
            WaitForLoad();
        }
        public void ClickBtnMegaphone()
        {
            var _clickIconMegaphone = WaitForElementIsVisibleNew(By.XPath("//*[@id=\"itemForm_0\"]/div[2]/div[13]/div/a/span"));
            _clickIconMegaphone.Click();
            WaitPageLoading();
            WaitForLoad();
        }
        public void AddSanctionAmount(string quantity)
        {
            WaitPageLoading();
            _sanctionAmountInput = WaitForElementExists(By.Id(SANCTION_AMOUNT));
            _sanctionAmountInput.SetValue(ControlType.TextBox, quantity);
            WaitPageLoading();
            _clickSave = WaitForElementIsVisibleNew(By.Id(BTN_SAVE));
            _clickSave.Click();
        }
        public string GetItemNameClaimsText()
        {
            var _itemNameClaims = WaitForElementExists(By.XPath(ITEMNAME_CLAIMS));
            return _itemNameClaims.Text;
        }
        public string GetReceivedClaimsText()
        {
            var _itemNameClaims = WaitForElementExists(By.XPath("//*[@id=\"tabContentServiceContainer\"]/div[1]/div[2]/div/div/div[4]/span"));
            return _itemNameClaims.Text;
        }
    }
}