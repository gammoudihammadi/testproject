using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.ReceiptNotes;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using DocumentFormat.OpenXml.Drawing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using UglyToad.PdfPig.Content;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Purchasing

{
    public class PurchaseOrderItem : PageBase
    {
        public PurchaseOrderItem(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // _____________________________________________ Constantes _________________________________________________

        // General
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";

        private const string VALIDATE_BUTTON = "//*[@id=\"div-body\"]/div/div[1]/div/div[2]/button";
        private const string VALIDATE_PO = "btn-validate-purchase-order";
        private const string VALIDATE_ERROR = "#modal-1 .errors-panel";
        private const string CONFIRM_VALIDATE = "btn-popup-validate";
        private const string APPROVE_PO = "btn-approve-po";
        private const string CONFIRM_APPROVE = "btn-popup-validate";
        private const string CANCEL_POPUP_BUTTON = "btn-cancel-popup";

        private const string EXTENDED_BUTTON = "//*[@id=\"div-body\"]/div/div[1]/div/div/button";
        private const string PRINT_PURCHASE_ORDER = "btn_print_po";
        private const string CONFIRM_PRINT = "printButton";

        private const string GENERATE_RN = "btn-generate-receipt-note";
        private const string DELIVERY_ORDER_NUMBER = "ReceiptNote_DeliveryOrderNumber";
        private const string RECEIPT_NOTE_NUMBER = "tb-new-receipt-note-number";
        private const string CREATE_RN = "btn-submit-form-create-receipt-note";

        private const string SEND_MAIL_BTN = "btn-send-by-email-purchase-order";
        private const string EMAIL_TO = "ToAddresses";
        private const string SEND = "btn-popup-send";
        private const string SEND_TO_EDI = "btn-send-to-edi";
        private const string NAMEFIELD = "//*[@id=\"popup0\"]/div[3]";


        // General information
        private const string GENERAL_INFORMATION = "hrefTabContentInformations";
        private const string PO_RECEIPT_STATUS = "PurchaseOrder_ReceiptStatus";
        private const string PO_STATUS = "PurchaseOrder_Status";

        // Items
        private const string GROUP = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/span";
        private const string ITEM = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[2]";
        private const string ITEM_NAME = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[2]/div[3]/span";
        private const string ITEM_NAME_VALIDATED = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div/div[3]/span";
        private const string QUANTITY_INPUT = "item_RelatedPurchaseOrderItemVM_PurchaseOrderItemDetail_Quantity";
        private const string QUANTITY_VALUE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div/div[7]/span";
        private const string SAVE_ICON = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[13]/span[@class='glyphicon glyphicon-floppy-saved']";
        private const string EXTENDED_MENU = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[@title='{0}']/../../div[14]/div/a/span";
        private const string ITEM_EDIT = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[@title='{0}']/../../div[14]/div/ul/li[*]/a/span[@class='glyphicon glyphicon glyphicon-pencil glyph-span']";
        private const string COMMENT_ICON = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[@title='{0}']/../../div[14]/div/ul/li[*]/a/span[@class='glyphicon glyphicon-comment glyph-span']";
        private const string COMMENT = "Comment";
        private const string SAVE_COMMENT = "//*[@id=\"modal-1\"]/div/div/div/div[2]/div/form/div[2]/button[2]";
        private const string ITEM_FORMULA_TOTAL_SUM = "total-price-span";
        private const string LIST_PROD_QUANTITY_LINE = "/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[{0}]/div/div/form/div[1]/div[7]/input";
        private const string ALLERGENS_BTN = "/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div[1]/div/div[2]/div[2]/div/div/form/div[2]/div[13]/a";
        private const string ALLERGENS_LIST = "/html/body/div[4]/div/div/div/div[2]/div/div/div/div/ul";

        // Filtres
        private const string RESET_FILTER_DEV = "ResetFilter";
        private const string RESET_FILTER_PATCH = "//*[@id=\"formSearchItems\"]/div[1]/a";

        private const string SEARCH_BY_NAME_FILTER = "tbSearchPatternWithAutocomplete";
        private const string FIRST_RESULT_SEARCH = "//*[@id=\"formSearchItems\"]/div[2]/span/div/div/div/strong[text()='{0}']";
        private const string KEYWORD_FILTER = "ItemIndexVMSelectedKeywords_ms";
        private const string UNCHECK_ALL_KEYWORD = "/html/body/div[11]/div/ul/li[2]/a";
        private const string SEARCH_KEYWORD = "/html/body/div[11]/div/div/label/input";
        private const string SHOW_NOT_PURCHASED_FILTER = "ShowNotPurchased";
        private const string GROUP_FILTER = "ItemIndexVMSelectedGroups_ms";
        private const string UNCHECK_ALL_GROUP = "/html/body/div[12]/div/ul/li[2]/a";
        private const string UNCHECK_ALL_SUB_GROUP = "/html/body/div[13]/div/ul/li[2]/a";
        private const string SEARCH_GROUP = "/html/body/div[12]/div/div/label/input";
        private const string SEARCH_SUB_GROUP = "ItemIndexVMSelectedSubGroups_ms";
        private const string SEARCH_SUB_GROUP_TEXT = "/html/body/div[12]/div/div/label/input";
        private const string SEND_BY_EMAIL = "btn-send-by-email-purchase-order";
        private const string TO_EMAIL = "ToAddresses";
        private const string SEND_EMAIL_CONFIRM = "btn-popup-send";
        private const string PROD_QUANTITY_INPUT = "item_PodRowDto_Quantity";
        private const string DISPLAYED_PACKING = "//*[@id=\"listOfItems\"]/div[2]/div[2]/div/div/form/div/div[5]/span";
        private const string ORDERED_QTE = "//*[@id=\"listOfItems\"]/div[2]/div[2]/div/div/form/div/div[12]/span";
        private const string SEND_EMAIL_CONFIRMP = "btn-init-async-send-mail";
        private const string CC_ADDRESS = "CcAddresses";
        private const string CLICK_COMMENT = "/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div[1]/div/div[2]/div[2]/div/div/form/div/div[14]/a/span";

        // ____________________________________________ Variables ____________________________________________________

        // général
        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;
        [FindsBy(How = How.XPath, Using = ORDERED_QTE)]
        private IWebElement _orderedQte;


        [FindsBy(How = How.XPath, Using = VALIDATE_BUTTON)]
        private IWebElement _validateButton;

        [FindsBy(How = How.Id, Using = VALIDATE_PO)]
        private IWebElement _validatePO;

        [FindsBy(How = How.CssSelector, Using = VALIDATE_ERROR)]
        private IWebElement _validateError;

        [FindsBy(How = How.Id, Using = CONFIRM_VALIDATE)]
        private IWebElement _confirmValidate;

        [FindsBy(How = How.Id, Using = APPROVE_PO)]
        private IWebElement _approvePO;

        [FindsBy(How = How.Id, Using = CONFIRM_APPROVE)]
        private IWebElement _confirmApprove;

        [FindsBy(How = How.Id, Using = CANCEL_POPUP_BUTTON)]
        private IWebElement _cancelPopupButton;

        [FindsBy(How = How.XPath, Using = EXTENDED_BUTTON)]
        private IWebElement _extendedButton;

        [FindsBy(How = How.Id, Using = PRINT_PURCHASE_ORDER)]
        private IWebElement _print;

        [FindsBy(How = How.Id, Using = CONFIRM_PRINT)]
        private IWebElement _confirmPrint;

        [FindsBy(How = How.Id, Using = GENERATE_RN)]
        private IWebElement _generateReceiptNote;

        [FindsBy(How = How.Id, Using = DELIVERY_ORDER_NUMBER)]
        private IWebElement _deliveryOrderNumber;

        [FindsBy(How = How.Id, Using = RECEIPT_NOTE_NUMBER)]
        private IWebElement _receiptNoteNumber;

        [FindsBy(How = How.Id, Using = CREATE_RN)]
        private IWebElement _createRN;

        [FindsBy(How = How.Id, Using = SEND_MAIL_BTN)]
        private IWebElement _sendByMailBtn;

        [FindsBy(How = How.Id, Using = EMAIL_TO)]
        private IWebElement _emailTo;

        [FindsBy(How = How.Id, Using = SEND)]
        private IWebElement _sendBtn;

        [FindsBy(How = How.Id, Using = SEND_TO_EDI)]
        private IWebElement _sendBtnToEdi;

        // General information
        [FindsBy(How = How.Id, Using = GENERAL_INFORMATION)]
        private IWebElement _generalInformation;

        [FindsBy(How = How.Id, Using = PO_RECEIPT_STATUS)]
        private IWebElement _receiptStatusPO;

        [FindsBy(How = How.Id, Using = PO_STATUS)]
        private IWebElement _statusPO;

        // Items
        [FindsBy(How = How.XPath, Using = GROUP)]
        private IWebElement _group;

        [FindsBy(How = How.XPath, Using = ITEM)]
        private IWebElement _item;

        [FindsBy(How = How.XPath, Using = ITEM_NAME)]
        private IWebElement _itemName;

        [FindsBy(How = How.XPath, Using = ITEM_NAME_VALIDATED)]
        private IWebElement _itemNameValidated;

        [FindsBy(How = How.Id, Using = QUANTITY_INPUT)]
        private IWebElement _quantityInput;

        [FindsBy(How = How.XPath, Using = QUANTITY_VALUE)]
        private IWebElement _quantityValue;

        [FindsBy(How = How.XPath, Using = EXTENDED_MENU)]
        private IWebElement _extendedMenu;

        [FindsBy(How = How.XPath, Using = ITEM_EDIT)]
        private IWebElement _editItem;

        [FindsBy(How = How.XPath, Using = COMMENT_ICON)]
        private IWebElement _commentItem;

        [FindsBy(How = How.Id, Using = COMMENT)]
        private IWebElement _comment;

        [FindsBy(How = How.XPath, Using = SAVE_COMMENT)]
        private IWebElement _saveComment;

        [FindsBy(How = How.Id, Using = ITEM_FORMULA_TOTAL_SUM)]
        private IWebElement _itemFormulaTotalSum;

        [FindsBy(How = How.Id, Using = ALLERGENS_BTN)]
        private IWebElement _allergensBtn;

        [FindsBy(How = How.Id, Using = ALLERGENS_LIST)]
        private IWebElement _allergensList;

        // _____________________________________ Filtres ____________________________________________

        [FindsBy(How = How.Id, Using = RESET_FILTER_DEV)]
        private IWebElement _resetFilterDev;

        [FindsBy(How = How.XPath, Using = RESET_FILTER_PATCH)]
        private IWebElement _resetFilterPatch;

        [FindsBy(How = How.Id, Using = SEARCH_BY_NAME_FILTER)]
        private IWebElement _searchByNameFilter;

        [FindsBy(How = How.Id, Using = KEYWORD_FILTER)]
        private IWebElement _keywordFilter;

        [FindsBy(How = How.XPath, Using = UNCHECK_ALL_KEYWORD)]
        private IWebElement _uncheckAllKeyword;

        [FindsBy(How = How.XPath, Using = SEARCH_KEYWORD)]
        private IWebElement _searchKeyword;

        [FindsBy(How = How.Id, Using = SHOW_NOT_PURCHASED_FILTER)]
        private IWebElement _showItemNotPurchasedFilter;

        [FindsBy(How = How.Id, Using = GROUP_FILTER)]
        private IWebElement _groupFilter;

        [FindsBy(How = How.XPath, Using = UNCHECK_ALL_GROUP)]
        private IWebElement _uncheckAllGroup;

        [FindsBy(How = How.XPath, Using = SEARCH_GROUP)]
        private IWebElement _searchGroup;

        [FindsBy(How = How.Id, Using = SEARCH_SUB_GROUP)]
        private IWebElement _filterbysubgrp;

        [FindsBy(How = How.Id, Using = PROD_QUANTITY_INPUT)]
        private IWebElement _prodQuantityInput;

        [FindsBy(How = How.XPath, Using = DISPLAYED_PACKING)]
        private IWebElement _displayedPacking;

        [FindsBy(How = How.XPath, Using = CLICK_COMMENT)]
        private IWebElement _clickComment;

        public enum FilterItemType
        {
            ByName,
            ByKeyword,
            ShowItemsNotPurchased,
            ByGroup,
            BySubGroup
        }

        public void Filter(FilterItemType FilterItemType, object value)
        {
            Actions action = new Actions(_webDriver);
            switch (FilterItemType)
            {
                case FilterItemType.ByName:
                    _searchByNameFilter = WaitForElementIsVisible(By.Id(SEARCH_BY_NAME_FILTER));
                    _searchByNameFilter.SetValue(ControlType.TextBox, value);

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
                case FilterItemType.ByKeyword:
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
                case FilterItemType.ShowItemsNotPurchased:
                    _showItemNotPurchasedFilter = WaitForElementExists(By.Id(SHOW_NOT_PURCHASED_FILTER));
                    action.MoveToElement(_showItemNotPurchasedFilter).Perform();
                    _showItemNotPurchasedFilter.SetValue(ControlType.CheckBox, value);
                    break;
                case FilterItemType.ByGroup:
                    FilterUncheckAllGroup();

                    _searchGroup = WaitForElementIsVisible(By.XPath(SEARCH_GROUP));
                    _searchGroup.SetValue(ControlType.TextBox, value);
                    Thread.Sleep(1000);

                    var valueToCheck = WaitForElementIsVisible(By.XPath("//label//span[text()='" + value + "']"));
                    valueToCheck.Click();

                    _groupFilter.Click();
                    break;
                case FilterItemType.BySubGroup:
                    FilterUncheckAllSubGroup();
                    _filterbysubgrp = WaitForElementIsVisible(By.XPath(SEARCH_SUB_GROUP_TEXT));
                    _filterbysubgrp.SetValue(ControlType.TextBox, value);
                    Thread.Sleep(1000);
                    var subgrptocheck = WaitForElementIsVisible(By.XPath("//label//span[text()='" + value + "']"));
                    subgrptocheck.Click();
                    break;
            }
            WaitPageLoading();
            Thread.Sleep(2000);
        }

        public void FilterUncheckAllGroup()
        {
            _groupFilter = WaitForElementIsVisible(By.Id(GROUP_FILTER));
            _groupFilter.Click();

            _uncheckAllGroup = WaitForElementIsVisible(By.XPath(UNCHECK_ALL_GROUP));
            _uncheckAllGroup.Click();
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
                if (elm.Text.Trim() != value.Trim())
                {
                    return false;
                }
            }
            return true;
        }

        public Boolean VerifyKeyword(string keyword)
        {
            var boolValue = true;
            var elements = _webDriver.FindElements(By.XPath(ITEM_NAME));

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

                    try
                    {
                        var itemGeneralInformationPage = EditItem(elm.GetAttribute("title"));
                        var itemKeywordTab = itemGeneralInformationPage.ClickOnKeywordItem();

                        var itemKeyword = itemKeywordTab.IsKeywordPresent(keyword);
                        itemGeneralInformationPage.Close();

                        if (!itemKeyword)
                            return false;
                    }
                    catch
                    {
                        boolValue = false;
                    }
                    compteur++;
                }
            }

            return boolValue;
        }

        public double GetTotalSum()
        {
            _itemFormulaTotalSum = WaitForElementIsVisible(By.Id(ITEM_FORMULA_TOTAL_SUM));
            // "€ 1 026,0700"
            return Convert.ToDouble(_itemFormulaTotalSum.Text.Replace("€", "").Replace(" ", ""), new NumberFormatInfo() { NumberDecimalSeparator = "," });
        }


        public Boolean VerifyGroup(string group)
        {
            var boolValue = true;
            var elements = _webDriver.FindElements(By.XPath(ITEM_NAME));

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

                    try
                    {
                        var itemGeneralInformationPage = EditItem(elm.GetAttribute("title"));
                        var groupName = itemGeneralInformationPage.GetGroupName();
                        itemGeneralInformationPage.Close();

                        // On ferme l'onglet ouvert
                        if (!groupName.Equals(group))
                            return false;
                    }
                    catch
                    {
                        boolValue = false;
                    }
                    compteur++;
                }
            }
            return boolValue;
        }

        public bool VerifyPurchased()
        {
            var elements = _webDriver.FindElements(By.XPath(QUANTITY_VALUE));

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

        // ___________________________________________ Méthodes _____________________________________________________

        public PurchaseOrdersPage BackToList()
        {
            _backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            _backToList.Click();
            WaitForLoad();

            return new PurchaseOrdersPage(_webDriver, _testContext);
        }

        public void ShowValidateMenu()
        {
            var actions = new Actions(_webDriver);

            _validateButton = WaitForElementIsVisible(By.XPath(VALIDATE_BUTTON));
            actions.MoveToElement(_validateButton).Perform();
            WaitForLoad();
        }

        public void Validate()
        {
            WaitPageLoading();
            ShowValidateMenu();

            _validatePO = WaitForElementIsVisible(By.Id(VALIDATE_PO));
            _validatePO.Click();
            WaitForLoad();

            _confirmValidate = WaitForElementIsVisible(By.Id(CONFIRM_VALIDATE));
            _confirmValidate.Click();
            WaitForLoad();
        }

        public bool ValidateHasError()
        {
            ShowValidateMenu();

            _validatePO = WaitForElementIsVisible(By.Id(VALIDATE_PO));
            _validatePO.Click();
            WaitForLoad();

            _cancelPopupButton = WaitForElementIsVisible(By.Id(CANCEL_POPUP_BUTTON));
            var result = isElementExists(By.CssSelector(VALIDATE_ERROR));
            _cancelPopupButton.Click();
            return result;
        }

        public void Approve()
        {
            ShowValidateMenu();

            _approvePO = WaitForElementIsVisible(By.Id(APPROVE_PO));
            _approvePO.Click();
            WaitForLoad();

            _confirmApprove = WaitForElementIsVisible(By.Id(CONFIRM_APPROVE));
            _confirmApprove.Click();
            WaitForLoad();
        }

        public override void ShowExtendedMenu()
        {
            var actions = new Actions(_webDriver);

            _extendedButton = WaitForElementIsVisible(By.XPath(EXTENDED_BUTTON));
            actions.MoveToElement(_extendedButton).Perform();
        }

        public PrintReportPage PrintDetails(bool versionPrint)
        {
            ShowExtendedMenu();

            _print = WaitForElementIsVisible(By.Id(PRINT_PURCHASE_ORDER));
            _print.Click();
            WaitForLoad();

            _confirmPrint = WaitForElementIsVisible(By.Id(CONFIRM_PRINT));
            _confirmPrint.Click();
            WaitForLoad();

            if (versionPrint)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-pdf']"));
                ClosePrintButton();
            }

            //Results are opened in a new tab, switch the driver to the newly created one
            WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new PrintReportPage(_webDriver, _testContext);
        }

        public string GenerateReceiptNote(bool withNumber)
        {
            var deliveryOrderNumber = new Random().Next();
            var receiptNoteID = "";

            ShowExtendedMenu();

            _generateReceiptNote = WaitForElementIsVisible(By.Id(GENERATE_RN));
            _generateReceiptNote.Click();

            _deliveryOrderNumber = WaitForElementIsVisible(By.Id(DELIVERY_ORDER_NUMBER));
            _deliveryOrderNumber.SetValue(ControlType.TextBox, deliveryOrderNumber.ToString());

            if (withNumber)
            {
                _receiptNoteNumber = WaitForElementIsVisible(By.Id(RECEIPT_NOTE_NUMBER));
                receiptNoteID = _receiptNoteNumber.GetAttribute("value");
            }

            return receiptNoteID;
        }

        public ReceiptNotesItem ValidateReceiptNoteCreation()
        {
            WaitForLoad();
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            _createRN = WaitForElementIsVisible(By.Id(CREATE_RN));
            javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", _createRN);
            WaitForLoad();
            _createRN.Click();
            WaitForLoad();

            return new ReceiptNotesItem(_webDriver, _testContext);
        }

        public void SendPODetailsByMail(string email)
        {
            ShowExtendedMenu();

            _sendByMailBtn = WaitForElementIsVisible(By.Id(SEND_MAIL_BTN));
            _sendByMailBtn.Click();
            WaitForLoad();

            _emailTo = WaitForElementIsVisible(By.Id(EMAIL_TO));
            _emailTo.SetValue(ControlType.TextBox, email);

            _sendBtn = WaitForElementIsVisible(By.Id(SEND));
            _sendBtn.Click();
            WaitForLoad();

            // Temps de fermeture de la fenêtre
            Thread.Sleep(2000);
        }

        // General Informations

        public PurchaseOrderGeneralInformation ClickOnGeneralInformation()
        {

            _generalInformation = WaitForElementIsVisible(By.Id(GENERAL_INFORMATION));
            _generalInformation.Click();
            WaitForLoad();

            return new PurchaseOrderGeneralInformation(_webDriver, _testContext);
        }

        public void ChangeReceiptStatus(string value)
        {
            _receiptStatusPO = WaitForElementIsVisible(By.Id("ReceiptStatus"));
            _receiptStatusPO.SetValue(ControlType.DropDownList, value);

            Thread.Sleep(2000);
        }

        public void ChangeStatus(string value)
        {
            _statusPO = WaitForElementIsVisible(By.Id("PurchaseOrderStatus"));
            _statusPO.SetValue(ControlType.DropDownList, value);
            Thread.Sleep(2000);
        }

        // Items

        public void SelectFirstItemPo()
        {
            _item = WaitForElementIsVisible(By.XPath(ITEM));
            _item.Click();
            WaitForLoad();
        }

        public string GetFirstItemName()
        {
            try
            {
                WaitForLoad();
                _itemName = _webDriver.FindElement(By.XPath(ITEM_NAME));
                // title : Item Test OF from SO (Item Test OF from SO_SupplierRef)
                // text : Item Test OF from SO
                return _itemName.GetAttribute("innerText").Trim(); // GetAttribute("title");
            }
            catch
            {
                return "";
            }

        }

        public string GetFirstItemNameValidated()
        {
            _itemNameValidated = _webDriver.FindElement(By.XPath(ITEM_NAME_VALIDATED));
            return _itemNameValidated.GetAttribute("innerText").Trim(); // GetAttribute("title");
        }
        public string GetDeliveryDate()
        {
            var dateD = _webDriver.FindElement(By.Id("datapicker-new-purchase-order-delivery"));
            return dateD.GetAttribute("value").Trim();
        }



        public string getorderedQty()
        {
            _orderedQte = _webDriver.FindElement(By.XPath(ORDERED_QTE));
            return _orderedQte.Text;//GetAttribute("innerText").Trim(); // GetAttribute("title");
        }

        public string GetFirstGroupName()
        {
            try
            {
                _group = _webDriver.FindElement(By.XPath(GROUP));
                var grpSubGrpName = _group.Text;
                var grpName = grpSubGrpName.Split('/')[0];
                grpName = grpName.Trim();
                return grpName;
            }
            catch
            {
                return "";
            }

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

        public void AddQuantity(string quantity)
        {
            _quantityInput = WaitForElementIsVisible(By.Id("item_PodRowDto_Quantity"));
            _quantityInput.SetValue(ControlType.TextBox, quantity);
            // manque la disquette
            //if(isElementVisible(By.XPath(SAVE_ICON)))
            //{
            //WaitForElementIsVisible(By.XPath(SAVE_ICON));
            //}
            Thread.Sleep(2000);
            WaitForLoad();
        }

        public string GetQuantity()
        {
            _quantityValue = WaitForElementIsVisible(By.XPath(QUANTITY_VALUE));

            return _quantityValue.Text;
        }

        public void AddComment(string itemName, string comment)
        {
            _extendedMenu = WaitForElementIsVisible(By.XPath(string.Format(EXTENDED_MENU, itemName)));
            _extendedMenu.Click();

            _commentItem = WaitForElementIsVisible(By.XPath(string.Format("//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[@title='{0}']/../../div[14]/div/ul/li[*]/a/span[contains(@class,'message')]", itemName)));

            _commentItem.Click();
            WaitForLoad();

            _comment = WaitForElementIsVisible(By.Id(COMMENT));
            _comment.SetValue(ControlType.TextBox, comment);

            _saveComment = WaitForElementToBeClickable(By.XPath("//*[@id=\"modal-1\"]/div/div[2]/div/form/div[2]/button[2]"));

            _saveComment.Click();
            WaitForLoad();
        }
        public void selectCommentItem()
        {
            _clickComment = WaitForElementIsVisible(By.XPath(CLICK_COMMENT));
            _clickComment.Click();
            WaitForLoad();
        }
        public bool checkSizeTextComment()
        {
            IWebElement element = _webDriver.FindElement(By.XPath("//*[@id='SavePORowCommentForm']/div[2]"));
            string fontSize = element.GetCssValue("font-size");   //D'après Monitoring 'FontSize' faut égal à '12px' Mais elle est égal à '14px'
            if (fontSize == "12px") return true;
            return false;

        }

        public string GetComment(string itemName)
        {
            _extendedMenu = WaitForElementIsVisible(By.XPath(string.Format(EXTENDED_MENU, itemName)));
            _extendedMenu.Click();

            _commentItem = WaitForElementIsVisible(By.XPath(string.Format("//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[@title='{0}']/../../div[14]/div/ul/li[*]/a/span[contains(@class,'message')]", itemName)));

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

        public ItemGeneralInformationPage EditItem(string itemName)
        {
            _extendedMenu = WaitForElementIsVisible(By.XPath(string.Format(EXTENDED_MENU, itemName)));
            _extendedMenu.Click();
            _editItem = WaitForElementIsVisible(By.XPath(string.Format("//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[@title='{0}']/../../div[14]/div/ul/li[*]/a/span[contains(@class,'pencil')]", itemName)));

            _editItem.Click();
            WaitForLoad();

            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            WaitForLoad();

            return new ItemGeneralInformationPage(_webDriver, _testContext);
        }

        public string GetPurchaseOrderNumber()
        {
            var _purchaseOrderNumber = WaitForElementIsVisible(By.XPath("//*[@id=\"div-body\"]/div/div[1]/h1"));
            return _purchaseOrderNumber.Text.Substring("PURCHASE ORDER No°".Length);
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
        public bool VerifyDetailChangeLine()
        {
            var qte = "15";
            var i = 2;
            var lineQuantity = WaitForElementIsVisible(By.XPath(string.Format(LIST_PROD_QUANTITY_LINE, i)));
            _quantityInput = WaitForElementIsVisible(By.Id("item_PodRowDto_Quantity"));
            _quantityInput.Click();
            AddQuantity(qte);
            WaitForLoad();
            _quantityInput.SendKeys(Keys.Enter);

            if (isElementVisible(By.XPath("//*/div[contains(@class,'editable-row') and contains(@class,'selected')]//div[7]/input[@type!='hidden']")))
            {
                lineQuantity = WaitForElementIsVisible(By.XPath("//*/div[contains(@class,'editable-row') and contains(@class,'selected')]//div[7]/input[@type!='hidden']"));
                if (lineQuantity.GetAttribute("value") != qte)
                {
                    return true;
                }
                return false;
            }
            else
            {
                return false;
            }
        }

        public bool IsAllergenIconGreen(string itemName)
        {
            _allergensBtn = WaitForElementIsVisible(By.XPath(string.Format(ALLERGENS_BTN, itemName)));
            var icon = _allergensBtn.FindElement(By.TagName("span"));
            var iconClass = icon.GetAttribute("class");
            return iconClass.Contains("green") ? true : false;
        }

        public List<string> GetAllergens(string itemName)
        {
            List<string> item_allergens = new List<string>();
            _allergensBtn = WaitForElementIsVisible(By.XPath(string.Format(ALLERGENS_BTN, itemName)));
            _allergensBtn.Click();

            _allergensList = WaitForElementIsVisible(By.XPath(string.Format(ALLERGENS_LIST)));
            var allergensInList = _allergensList.FindElements(By.TagName("li"));
            if (allergensInList.Count > 0)
            {
                foreach (var allergen in allergensInList)
                {
                    // Check if <li> contains an image, if not move to next iteration.
                    // The allergen won't added to the list ans so the test won't pass.
                    if (allergen.FindElement(By.TagName("img")) == null) continue;
                    item_allergens.Add(allergen.Text);
                }
            }
            return item_allergens;
        }
        public PurchaseOrderItem SendByEmail(string email)
        {
            ShowExtendedMenu();
            var sendBtn = WaitForElementIsVisible(By.Id(SEND_BY_EMAIL));
            sendBtn.Click();
            var to = WaitForElementIsVisible(By.Id(TO_EMAIL));
            to.Clear();
            to.SendKeys(email);
            ConfirmEmailSend();
            //WaitForLoad();

            return new PurchaseOrderItem(_webDriver, _testContext);
        }
        public void ConfirmEmailSend()
        {
            var confirmBtn = WaitForElementIsVisible(By.Id(SEND_EMAIL_CONFIRMP));
            confirmBtn.Click();

            int i = 10;
            while (i > 0)
            {
                WaitPageLoading();
                i--;
                if (!isElementExists(By.Id(CC_ADDRESS)))
                {
                    break;
                }
            }
        }

        public void SendToEdi()
        {

            ShowExtendedMenu();

            _sendBtnToEdi = WaitForElementIsVisible(By.Id(SEND_TO_EDI));
            _sendBtnToEdi.Click();
            Thread.Sleep(2000);
            var ClickOk = WaitForElementIsVisible(By.Id("dataAlertCancel"));
            ClickOk.Click();
            Thread.Sleep(2000);


        }

        public void ChangeProdQuantity(string quantity)
        {
            _prodQuantityInput = WaitForElementIsVisible(By.Id(PROD_QUANTITY_INPUT));
            _prodQuantityInput.SetValue(ControlType.TextBox, quantity);
            WaitPageLoading();
            WaitForLoad();
        }

        public string GetPacking()
        {
            WaitPageLoading();
            _displayedPacking = _webDriver.FindElement(By.XPath(DISPLAYED_PACKING));
            return _displayedPacking.Text.Trim();
        }

        public bool ValidateIsEnabled()
        {
            ShowValidateMenu();
            _validatePO = WaitForElementIsVisible(By.Id(VALIDATE_PO));

            var disabledAttribute = _validatePO.GetAttribute("disabled");

            return string.IsNullOrEmpty(disabledAttribute);
        }
        public int GetItemsCount()
        {
            // Localiser les éléments 
            var items = _webDriver.FindElements(By.XPath("//*[@id=\"listOfItems\"]/div[2]/div[2]"));

            return items.Count;
        }

        public string GetItemNameValue()
        {
            WaitPageLoading();
            var nameField = _webDriver.FindElement(By.XPath(NAMEFIELD));
            return nameField.Text;
        }

    }
}
