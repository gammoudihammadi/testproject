using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using iText.StyledXmlParser.Jsoup.Nodes;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Accounting.SupplierInvoices;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Security.Claims;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.Claims
{
    public class ClaimsItem : PageBase
    {

        public ClaimsItem(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //___________________________________________ Constantes ____________________________________________________

        // General 
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";
        private const string EXTENDED_BUTTON = "//*[@id=\"div-body\"]/div/div[1]/div/div/button";
        private const string EXTENDED_BUTTON_INDEX = "//*[@id=\"div-body\"]/div/div[2]/div[1]/div/div[1]/button";
        private const string REFRESH = "btn-receipt-notes-refresh";
        private const string PRINT_DEV = "print_claim_note";
        private const string PRINT = "print_receipt_note";
        private const string PRINT_CLAIM_RESULT = "//*[@id=\"btn-print-claim-notes-report\"]";
        private const string VALIDATE_PRINT = "printButton";
        private const string SEND_BY_MAIL_BTN = "btn-send-by-email-receipt-note";
        private const string SEND_BY_MAIL_BTN_DEV = "btn-send-by-email-claim-note";
        private const string EMAIL_RECEIVER = "ToAddresses";
        private const string SEND_MAIL_BTN = "btn-init-async-send-mail";
        private const string GENERATE_SI = "btn-generate-associated-supplier-invoice";
        private const string INVOICE_NUMBER = "tb-new-supplier-invoice-number";
        private const string CREATE_FROM = "check-box-create-from";
        private const string CONFIRM_SI = "btn-submit-form-create-supplier-invoice";
        private const string SAVE = "btn-valid-claim";

        private const string VALIDATE = "btn-validate-claim-note";
        private const string VALIDATE_CLAIM = "btn-validate-claim-note";
        private const string CONFIRM_VALIDATE = "btn-popup-validate";
        private const string SEND_MAIL_POPUP = "//*[@id=\"modal-1\"]";

        // Onglets
        private const string GENERAL_INFORMATION = "hrefTabContentInformations";
        private const string CLAIMS = "hrefTabContentChecks";
        private const string FOOTER = "hrefTabContentFooter";

        // Tableau claims items
        private const string GROUP = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/span";
        private const string ITEM = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[2]";
        private const string ITEM2 = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[@id='row{0}']/div/div/form/div[2]";
        private const string ICON_EDIT = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[1]/span/img";
        private const string FIRST_ITEM_NAME = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[2]/div[3]/span";
        private const string SECOND_ITEM_NAME = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[{0}]/div/div/form/div[2]/div[3]/span";
        private const string ITEM_NAME = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[2]/div[3]/span[@title='{0}']";
        private const string ITEM_PRICE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[{0}]/div/div/form/div[2]/div[5]/span";
        private const string ITEM_PRICE_FORM = "//*[@id=\"itemForm_{0}\"]/div[1]/div[5]/div/div[2]/input";
        private const string ITEM_DN_PRICE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[{0}]/div/div/form/div[2]/div[7]/span";
        private const string ITEM_DN_PRICE_FORM = "//*[@id=\"itemForm_{0}\"]/div[1]/div[7]/div/div[2]/input";
        private const string PRICE_INPUT = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[@title='{0}']/../../div[5]/div[1]/div[2]/input";

        private const string ITEM_QUANTITY = "//*[@id=\"itemForm_{0}\"]/div[2]/div[6]/span";
        private const string ITEM_QUANTITY_FORM = "//*[@id=\"itemForm_{0}\"]/div[1]/div[6]/input";
        private const string ITEM_DN_QUANTITY = "//*[@id=\"itemForm_{0}\"]/div[2]/div[8]/span";
        private const string ITEM_DN_QUANTITY_FORM = "//*[@id=\"itemForm_{0}\"]/div[1]/div[8]/input";
        private const string ITEM_PACKAGING = "//*[@id=\"itemForm_0\"]/div[2]/div[4]/span";
        private const string ITEM_SANCTION = "//*/form[@id='itemForm_0']/div[1]/div[10]/div/div[2]/input";
        private const string CLAIM_AMOUNT = "//*/form[@id='itemForm_{0}']/div[2]/div[9]/span";
        private const string DECR_STOCK_CHECKBOX = "item_CndRowDto_DecreaseStock";
        private const string DECR_STOCK_QTY = "//*/form[@id='itemForm_0']/div[1]/div[14]/input";
        private const string FORM_QTY = "//*/form[@id='itemForm_0']/div[1]/div[6]/input";
        private const string QUANTITY_INPUT_DN = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[@title='{0}']/../../div[8]/input";
        private const string QUANTITY_INPUT = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[@title='{0}']/../../div[6]/input";
        private const string QUANTITY_INPUT_RN = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[@title='{0}']/../../div[8]/input";
        private const string ITEM_TYPE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[2]/div[11]";
        private const string TYPE_SELECT = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[@title='{0}']/../../div[7]/select";
        private const string TYPE_SELECT_RN = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[@title='{0}']/../../div[9]/select";
        private const string SAVE_ICON = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[@title='{0}']/../../div[16]/span";
        private const string TOTAL_SANCTION = "total-sanction-span";
        private const string TOTAL_CLAIM = "total-price-span";
        private const string RECEIVED_QTY_LINE = "/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[{0}]/div/div/form/div[1]/div[6]/input";

        private const string EXTENDED_MENU = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[@title='{0}']/../../div[10]/a/span";
        private const string EXTENDED_MENU_ITEM = "//span[@title='{0}']/../..//a[contains(@class, 'mini-btn')]";
        private const string EXTENDED_MENU_ITEMBIS = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[@title='{0}']/../../div[12]/div/a";
        private const string EXTENDED_MENU_FORM = "(//*/form/div[1]/div[15]/a/span)[1]";
        private const string EDIT_ITEM_FORM = "(//*/form/div[1]/div[15]/ul/li[3]/a/span)[1]";
        private const string PLUS_BTN = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[@title='{0}']/../../div[14]/a/../ul/li[3]/a/span";
        private const string PLUS_BTNBIS = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[@title='{0}']/../../div[12]/a/../ul/li[3]/a/span";
        private const string COMMENT_ICON_DEV = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[@title='{0}']/../../div[15]/ul/li[*]/a/span[contains(@class,'message')]";
        private const string COMMENT_ICON_PATCH = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[@title='{0}']/../../div[15]/ul/li[*]/a/span[@class=\"glyphicon glyphicon-comment glyph-span\"]";
        private const string COMMENT = "Comment";
        private const string SAVE_COMMENT_PATCH = "//*[@id=\"modal-1\"]/div/div/div/div[2]/div/form/div[2]/button[2]";
        private const string SAVE_COMMENT_DEV = "//*[@id=\"modal-1\"]/div/div[2]/div/form/div[2]/button[2]";
        private const string DELETE_ICON_DEV = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[@title='{0}']/../../div[15]/ul/li[*]/a/span[@class=\"fas fa-trash-alt glyph-span\"]";
        private const string DELETE_ICON_PATCH = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[@title='{0}']/../../div[15]/ul/li[*]/a/span[@class=\"glyphicon glyphicon-trash glyph-span\"]";
        //private const string EDIT_ITEM = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[@title='{0}']/../../div[10]/ul/li[*]/a/span[@class=\"glyphicon glyphicon glyphicon-pencil glyph-span\"]";
        private const string EDIT_ITEM = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[@title='{0}']/../../div[10]/a/../ul/li[5]/a/span";
        private const string EDIT_ITEM_INF = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[@title='{0}']/../../div[15]/a/../ul/li[3]/a/span";
        private const string PICTURE_ICON = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[@title='{0}']/../../div[10]/ul/li[*]/a/span[@class=\"glyphicon glyphicon-picture glyph-span\"]";
        private const string UPLOAD_PICTURE = "FileSent";
        private const string DELETE_PICTURE = "//*[@id=\"modal-1\"]/div/div/div/div[2]/div/form/div/a";
        private const string CLOSE_PICTURE = "//*[@id=\"modal-1\"]/div/div/div/div[3]/button";
        private const string ALLERGENS_BTN = "/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/form/div[2]/div[15]/a";
        private const string ALLERGENS_LIST = "/html/body/div[4]/div/div/div/div[2]/div/div/div/div/ul";

        // Filtres
        private const string RESET_FILTER = "//*[@id=\"formSearchItems\"]/div[1]/a";
        private const string FILTER_SEARCH_BY_NAME = "tbSearchPatternWithAutocomplete";
        private const string FIRST_RESULT_SEARCH = "//*[@id=\"formSearchItems\"]/div[2]/span/div/div/div[1]/strong[text()='{0}']";
        private const string KEYWORD_FILTER = "ItemIndexVMSelectedKeywords_ms";
        private const string UNSELECT_ALL_KEYWORD = "/html/body/div[11]/div/ul/li[2]/a";
        private const string SEARCH_KEYWORD = "/html/body/div[11]/div/div/label/input";
        private const string FILTER_NOT_CLAIMED = "ShowNullQuantity";
        private const string FILTER_GROUP = "ItemIndexVMSelectedGroups_ms";
        private const string UNSELECT_ALL_GROUP = "/html/body/div[12]/div/ul/li[2]/a";
        private const string SEARCH_GROUP = "/html/body/div[12]/div/div/label/input";
        private const string SHOW_ITEMS_NOT_CLAIMED = "//*[@id=\"formSearchItems\"]/div[3]/div[2]/div";

        private const string MEGAPHONE = "//*[@id=\"btnAddEditClaim_0\"]/div/a/span";

        private const string CC_MAIL = "CcAddresses";
        private const string ITEMS_SUB_MENU = "hrefTabContentItems";
        private const string GENERAL_INFORMATION_SUB_MENU = "//*[@id=\"hrefTabContentInformations\"]";
        private const string CLAIM_ITEM = "//*[@id=\"itemForm_{0}\"]/div[2]/div[3]/span";
        private const string FILE_SENT = "add-attachment-input";

        //___________________________________________ Variables ________________________________________________________________

        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;

        [FindsBy(How = How.XPath, Using = EXTENDED_BUTTON)]
        private IWebElement _extendedButton;

        [FindsBy(How = How.XPath, Using = EXTENDED_BUTTON_INDEX)]
        private IWebElement _extendedButtonIndex;

        [FindsBy(How = How.Id, Using = REFRESH)]
        private IWebElement _refresh;

        [FindsBy(How = How.Id, Using = PRINT)]
        private IWebElement _print;

        [FindsBy(How = How.Id, Using = PRINT_DEV)]
        private IWebElement _printDev;

        [FindsBy(How = How.XPath, Using = PRINT_CLAIM_RESULT)]
        private IWebElement _printClaimResult;

        [FindsBy(How = How.Id, Using = VALIDATE_PRINT)]
        private IWebElement _validatePrint;

        [FindsBy(How = How.Id, Using = SEND_BY_MAIL_BTN)]
        private IWebElement _sendByEmailBtn;

        [FindsBy(How = How.Id, Using = SEND_BY_MAIL_BTN_DEV)]
        private IWebElement _sendByEmailBtnDev;

        [FindsBy(How = How.Id, Using = EMAIL_RECEIVER)]
        private IWebElement _emailReceiver;

        [FindsBy(How = How.Id, Using = SEND_MAIL_BTN)]
        private IWebElement _sendEmail;

        [FindsBy(How = How.Id, Using = GENERATE_SI)]
        private IWebElement _generateSI;

        [FindsBy(How = How.Id, Using = INVOICE_NUMBER)]
        private IWebElement _invoiceNumber;

        [FindsBy(How = How.Id, Using = CREATE_FROM)]
        private IWebElement _createFrom;

        [FindsBy(How = How.Id, Using = CONFIRM_SI)]
        private IWebElement _validateSI;

        [FindsBy(How = How.Id, Using = VALIDATE)]
        private IWebElement _validate;

        [FindsBy(How = How.Id, Using = CONFIRM_VALIDATE)]
        private IWebElement _confirmValidate;

        // Onglets
        [FindsBy(How = How.Id, Using = GENERAL_INFORMATION)]
        private IWebElement _generalInformation;

        [FindsBy(How = How.Id, Using = FOOTER)]
        private IWebElement _footer;

        // Tableau items
        [FindsBy(How = How.XPath, Using = GROUP)]
        private IWebElement _group;

        [FindsBy(How = How.XPath, Using = ITEM)]
        private IWebElement _item;

        [FindsBy(How = How.XPath, Using = FIRST_ITEM_NAME)]
        private IWebElement _itemName;

        [FindsBy(How = How.XPath, Using = PRICE_INPUT)]
        private IWebElement _priceInput;

        [FindsBy(How = How.XPath, Using = ITEM_PRICE)]
        private IWebElement _price;

        [FindsBy(How = How.XPath, Using = ITEM_DN_PRICE)]
        private IWebElement _priceDN;

        [FindsBy(How = How.Name, Using = QUANTITY_INPUT)]
        private IWebElement _quantityInput;

        [FindsBy(How = How.XPath, Using = ITEM_QUANTITY)]
        private IWebElement _quantity;

        [FindsBy(How = How.XPath, Using = ITEM_DN_QUANTITY)]
        private IWebElement _quantityDN;

        [FindsBy(How = How.XPath, Using = ITEM_PACKAGING)]
        private IWebElement _packaging;

        [FindsBy(How = How.XPath, Using = ITEM_SANCTION)]
        private IWebElement _sanctionInput;

        [FindsBy(How = How.XPath, Using = CLAIM_AMOUNT)]
        private IWebElement _claimAmount;

        [FindsBy(How = How.Id, Using = DECR_STOCK_CHECKBOX)]
        private IWebElement _decrStock;

        [FindsBy(How = How.XPath, Using = DECR_STOCK_QTY)]
        private IWebElement _decrQty;

        [FindsBy(How = How.XPath, Using = FORM_QTY)]
        private IWebElement _qty;

        [FindsBy(How = How.Name, Using = TYPE_SELECT)]
        private IWebElement _typeSelect;

        [FindsBy(How = How.XPath, Using = ITEM_TYPE)]
        private IWebElement _type;

        [FindsBy(How = How.XPath, Using = EXTENDED_MENU)]
        private IWebElement _extendedMenu;

        [FindsBy(How = How.XPath, Using = EXTENDED_MENU_FORM)]
        private IWebElement _extendedMenuForm;

        [FindsBy(How = How.XPath, Using = EDIT_ITEM_FORM)]
        private IWebElement _editItemForm;

        [FindsBy(How = How.XPath, Using = COMMENT_ICON_DEV)]
        private IWebElement _addCommentDev;

        [FindsBy(How = How.XPath, Using = COMMENT_ICON_PATCH)]
        private IWebElement _addCommentPatch;

        [FindsBy(How = How.Id, Using = COMMENT)]
        private IWebElement _comment;

        [FindsBy(How = How.XPath, Using = SAVE_COMMENT_DEV)]
        private IWebElement _saveCommentDev;

        [FindsBy(How = How.XPath, Using = SAVE_COMMENT_PATCH)]
        private IWebElement _saveCommentPatch;

        [FindsBy(How = How.XPath, Using = DELETE_ICON_DEV)]
        private IWebElement _deleteDev;

        [FindsBy(How = How.XPath, Using = DELETE_ICON_PATCH)]
        private IWebElement _deletePatch;

        [FindsBy(How = How.XPath, Using = EDIT_ITEM)]
        private IWebElement _editItem;

        [FindsBy(How = How.XPath, Using = DELETE_PICTURE)]
        private IWebElement _deletePicture;

        [FindsBy(How = How.Id, Using = UPLOAD_PICTURE)]
        private IWebElement _uploadPicture;

        [FindsBy(How = How.XPath, Using = CLOSE_PICTURE)]
        private IWebElement _closePicture;

        [FindsBy(How = How.XPath, Using = PICTURE_ICON)]
        private IWebElement _addPicture;

        [FindsBy(How = How.Id, Using = TOTAL_SANCTION)]
        private IWebElement _totalSanction;

        [FindsBy(How = How.Id, Using = TOTAL_CLAIM)]
        private IWebElement _totalClaim;

        [FindsBy(How = How.Id, Using = ALLERGENS_BTN)]
        private IWebElement _allergensBtn;

        [FindsBy(How = How.Id, Using = ALLERGENS_LIST)]
        private IWebElement _allergensList;

        [FindsBy(How = How.XPath, Using = CLAIM_ITEM)]
        private IWebElement _claimItem;

        //_______________________________________________ Filtres _____________________________________________________________

        [FindsBy(How = How.XPath, Using = RESET_FILTER)]
        private IWebElement _resetFilter;

        [FindsBy(How = How.Id, Using = FILTER_SEARCH_BY_NAME)]
        private IWebElement _searchByNameFilter;

        [FindsBy(How = How.Id, Using = KEYWORD_FILTER)]
        private IWebElement _keywordFilter;

        [FindsBy(How = How.XPath, Using = UNSELECT_ALL_KEYWORD)]
        private IWebElement _unselectAllKeyword;

        [FindsBy(How = How.XPath, Using = SEARCH_KEYWORD)]
        private IWebElement _searchKeyword;

        [FindsBy(How = How.Id, Using = FILTER_NOT_CLAIMED)]
        private IWebElement _showNotClaimedFilter;

        [FindsBy(How = How.Id, Using = FILTER_GROUP)]
        private IWebElement _groupFilter;

        [FindsBy(How = How.XPath, Using = UNSELECT_ALL_GROUP)]
        private IWebElement _unselectAllGroup;

        [FindsBy(How = How.XPath, Using = SEARCH_GROUP)]
        private IWebElement _searchGroup;

        [FindsBy(How = How.Id, Using = SHOW_ITEMS_NOT_CLAIMED)]
        private IWebElement _showItemsNotClaimed;

        [FindsBy(How = How.Id, Using = CC_MAIL)]
        private IWebElement _ccMail;

        public enum FilterItemType
        {
            SearchByName,
            SearchByKeyword,
            ShowNotClaimed,
            ByGroup

        }

        public void Filter(FilterItemType FilterItemType, object value)
        {
            Actions action = new Actions(_webDriver);

            switch (FilterItemType)
            {
                case FilterItemType.SearchByName:
                    _searchByNameFilter = WaitForElementIsVisible(By.Id(FILTER_SEARCH_BY_NAME));
                    action.MoveToElement(_searchByNameFilter).Perform();
                    _searchByNameFilter.SetValue(ControlType.TextBox, value);
                    WaitPageLoading();
                    // choix "combobox" de recherche
                    if (isElementVisible(By.XPath(String.Format(FIRST_RESULT_SEARCH, value))))
                    {

                        var firstResultSearch = WaitForElementIsVisible(By.XPath(String.Format(FIRST_RESULT_SEARCH, value)));
                        firstResultSearch.Click();
                    }
                    else
                    {
                        // deuxième recherche...
                        return;
                    }
                    break;
                case FilterItemType.SearchByKeyword:
                    _keywordFilter = WaitForElementIsVisible(By.Id(KEYWORD_FILTER));
                    _keywordFilter.Click();

                    _unselectAllKeyword = WaitForElementIsVisible(By.XPath(UNSELECT_ALL_KEYWORD));
                    _unselectAllKeyword.Click();

                    _searchKeyword = WaitForElementIsVisible(By.XPath(SEARCH_KEYWORD));
                    _searchKeyword.SetValue(ControlType.TextBox, value);

                    var keywordToCheck = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    keywordToCheck.SetValue(ControlType.CheckBox, true);

                    _keywordFilter.Click();
                    break;
                case FilterItemType.ShowNotClaimed:
                    _showNotClaimedFilter = WaitForElementExists(By.Id(FILTER_NOT_CLAIMED));
                    action.MoveToElement(_showNotClaimedFilter).Perform();
                    _showNotClaimedFilter.SetValue(ControlType.CheckBox, value);
                    break;
                case FilterItemType.ByGroup:
                    _groupFilter = WaitForElementIsVisible(By.Id(FILTER_GROUP));
                    _groupFilter.Click();

                    _unselectAllGroup = WaitForElementIsVisible(By.XPath(UNSELECT_ALL_GROUP));
                    _unselectAllGroup.Click();

                    _searchGroup = WaitForElementIsVisible(By.XPath(SEARCH_GROUP));
                    _searchGroup.SetValue(ControlType.TextBox, value);
                    Thread.Sleep(2000);

                    var groupFilterValue = WaitForElementIsVisible(By.XPath("/html/body/div[12]/ul"));
                    groupFilterValue.Click();

                    _groupFilter.Click();
                    break;
            }

            WaitPageLoading();
            Thread.Sleep(2000);
        }

        public void ResetFilters()
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

        public bool IsSortedByName()
        {
            var ancientName = "";
            int compteur = 1;

            var elements = _webDriver.FindElements(By.XPath(FIRST_ITEM_NAME));

            if (elements.Count == 0)
                return false;

            foreach (var elm in elements)
            {
                if (compteur == 1)
                    ancientName = elm.GetAttribute("title");

                if (string.Compare(ancientName, elm.GetAttribute("title")) > 0)
                    return false;

                ancientName = elm.GetAttribute("title");
                compteur++;
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

                    try
                    {
                        var itemGeneralInformationPage = EditItemGeneralInformation(elm.GetAttribute("title"));
                        //var itemKeyword = itemGeneralInformationPage.IsKeywordPresent(keyword);
                        var itemKeywordTab = itemGeneralInformationPage.ClickOnKeywordItem();

                        var itemKeyword = itemKeywordTab.IsKeywordPresent(keyword);
                        //var itemKeyword = itemGeneralInformationPage.IsKeywordPresent(keyword);
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

        public Boolean IsClaimed()
        {
            var elements = _webDriver.FindElements(By.XPath(string.Format(ITEM_QUANTITY, 0)));

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
                    elm.Click();

                    try
                    {
                        var itemGeneralInformationPage = EditItemGeneralInformation(elm.GetAttribute("title"));
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

        //___________________________________________ Méthodes _____________________________________________________

        // General       
        public ClaimsPage BackToList()
        {
            _backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            _backToList.Click();
            WaitForLoad();
            return new ClaimsPage(_webDriver, _testContext);
        }

        public override void ShowExtendedMenu()
        {
            WaitForElementIsVisible(By.XPath(EXTENDED_BUTTON));
            var actions = new Actions(_webDriver);
            actions.MoveToElement(_extendedButton).Perform();
            WaitForLoad();
        }

        public void Refresh()
        {
            ShowExtendedMenu();

            if (isElementVisible(By.Id("btn-claim-notes-refresh")))
            {
                _refresh = WaitForElementIsVisible(By.Id("btn-claim-notes-refresh"));
            }
            else
            {
                // claim from RN
                _refresh = WaitForElementIsVisible(By.Id("btn-receipt-notes-refresh"));
            }
            _refresh.Click();
            WaitPageLoading();
            WaitForLoad();
        }


        public PrintReportPage PrintClaimItems(bool printValue)
        {
            ShowExtendedMenu();

            // Click on "print"
            if (IsDev())
            {
                _printDev = WaitForElementIsVisible(By.Id(PRINT_DEV));
                _printDev.Click();
                WaitForLoad();
            }
            else
            {
                _print = WaitForElementIsVisible(By.Id(PRINT));
                _print.Click();
                WaitForLoad();
            }

            _validatePrint = WaitForElementIsVisible(By.Id(VALIDATE_PRINT));
            _validatePrint.Click();
            WaitForLoad();

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

        public PrintReportPage PrintClaimResults(bool printValue)
        {
            _extendedButtonIndex = WaitForElementIsVisible(By.XPath(EXTENDED_BUTTON_INDEX));
            var actions = new Actions(_webDriver);
            actions.MoveToElement(_extendedButtonIndex).Perform();
            WaitForLoad();

            // Click on "print"
            _printClaimResult = WaitForElementIsVisible(By.XPath(PRINT_CLAIM_RESULT));
            _printClaimResult.Click();
            WaitForLoad();

            _validatePrint = WaitForElementIsVisible(By.Id(VALIDATE_PRINT));
            _validatePrint.Click();
            WaitForLoad();

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
            ClosePrintButton();
            ShowExtendedMenu();

            if (IsDev())
            {
                _sendByEmailBtnDev = WaitForElementIsVisible(By.Id(SEND_BY_MAIL_BTN_DEV));
                _sendByEmailBtnDev.Click();
                WaitForLoad();
            }
            else
            {
                _sendByEmailBtn = WaitForElementIsVisible(By.Id(SEND_BY_MAIL_BTN));
                _sendByEmailBtn.Click();
                WaitForLoad();
            }

            _emailReceiver = WaitForElementIsVisible(By.Id(EMAIL_RECEIVER));
            _emailReceiver.SetValue(ControlType.TextBox, "test@mail.com");
            WaitForLoad();

            _sendEmail = WaitForElementIsVisible(By.Id(SEND_MAIL_BTN));
            _sendEmail.Click();
            WaitForLoad();
            WaitPageLoading();
            wait.Until(d => !isElementVisible(By.XPath(SEND_MAIL_POPUP)));
        }


        public SupplierInvoicesItem GenerateSupplierInvoice(string supplierInvoiceNumber)
        {
            // Completed the modal
            ShowExtendedMenu();

            _generateSI = WaitForElementIsVisible(By.Id(GENERATE_SI));
            _generateSI.Click();
            //WaitForLoad();

            _invoiceNumber = WaitForElementIsVisible(By.Id(INVOICE_NUMBER));
            _invoiceNumber.SetValue(ControlType.TextBox, supplierInvoiceNumber);

            _createFrom = _webDriver.FindElement(By.Id(CREATE_FROM));
            _createFrom.SetValue(ControlType.CheckBox, true);
            WaitForLoad();

            _validateSI = WaitForElementIsVisible(By.Id(CONFIRM_SI));
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", _validateSI);
            _validateSI.Click();

            //Results are opened in a new tab, switch the driver to the newly created one
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(30));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new SupplierInvoicesItem(_webDriver, _testContext);
        }

        public void Validate()
        {
            ShowValidationMenu();
            WaitForLoad();
            _validate = WaitForElementIsVisible(By.Id(VALIDATE));
            _validate.Click();
            // animation
            //Thread.Sleep(2000);
            WaitForLoad();


            _confirmValidate = WaitForElementIsVisible(By.Id(CONFIRM_VALIDATE));
            _confirmValidate.Click();
            WaitForLoad();
        }
        public void Validate_CLAIM()
        {
            ShowValidationMenu();

            _validate = WaitForElementIsVisible(By.Id(VALIDATE_CLAIM));
            _validate.Click();
            // animation
            Thread.Sleep(2000);
            WaitForLoad();

            _confirmValidate = WaitForElementIsVisible(By.Id(CONFIRM_VALIDATE));
            _confirmValidate.Click();
            WaitForLoad();
        }

        // Onglets
        public ClaimsGeneralInformation ClickOnGeneralInformation()
        {
            _generalInformation = WaitForElementIsVisibleNew(By.Id(GENERAL_INFORMATION));
            _generalInformation.Click();
            WaitForLoad();

            return new ClaimsGeneralInformation(_webDriver, _testContext);
        }

        public ClaimsChecks ClickOnChecks()
        {
            var _checks = WaitForElementIsVisible(By.Id(CLAIMS));
            _checks.Click();
            WaitForLoad();

            return new ClaimsChecks(_webDriver, _testContext);
        }

        // Tableau claims items
        public void SelectFirstItem()
        {
            _item = WaitForElementIsVisible(By.XPath(ITEM));
            _item.Click();
            WaitForLoad();
        }

        public void SelectSecondItem()
        {
            _item = WaitForElementIsVisible(By.XPath(string.Format(ITEM2, 1)));
            _item.Click();
            WaitForLoad();
        }

        public void SelectItem(string itemName)
        {
            _item = WaitForElementIsVisible(By.XPath(String.Format(ITEM_NAME, itemName)));
            _item.Click();

            WaitForLoad();
        }

        public string GetFirstItemName()
        {
            _itemName = WaitForElementExists(By.XPath(FIRST_ITEM_NAME));
            return _itemName.Text;
        }

        public string GetSecondItemName()
        {
            _itemName = WaitForElementExists(By.XPath(string.Format(SECOND_ITEM_NAME, "4")));
            // DeleteItem utilise @title
            return _itemName.GetAttribute("title");
        }

        public string GetFirstGroupName()
        {
            if (isElementVisible(By.XPath(GROUP)))
            {
                _group = _webDriver.FindElement(By.XPath(GROUP));
                return _group.Text;
            }
            else
            {
                return "";
            }
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

        public void AddPrice(string itemName, int price)
        {
            _priceInput = WaitForElementIsVisible(By.XPath(String.Format(PRICE_INPUT, itemName)));
            _priceInput.SetValue(ControlType.TextBox, price.ToString());
            Thread.Sleep(2000);
            WaitForLoad();
        }

        public string GetPrice(int offset = 2)
        {
            _price = WaitForElementExists(By.XPath(string.Format(ITEM_PRICE, offset)));
            return _price.Text;
        }

        public void SetPrice(string price, int offset = 0)
        {
            _price = WaitForElementExists(By.XPath(string.Format(ITEM_PRICE_FORM, offset)));
            _price.Clear();
            _price.SendKeys(price);
            WaitForLoad();
        }



        public void SetDNPrice(string price, int offset = 1)
        {
            _priceDN = WaitForElementExists(By.XPath(string.Format(ITEM_DN_PRICE_FORM, offset)));
            _priceDN.Clear();
            _priceDN.SendKeys(price);
            _priceDN.SendKeys(Keys.Enter);
            WaitForLoad();
        }

        public string GetDNPrice(int offset = 2)
        {
            _priceDN = WaitForElementExists(By.XPath(string.Format(ITEM_DN_PRICE, offset)));
            return _priceDN.Text;
        }

        public void AddQuantityAndType(string itemName, int qty)
        {
            // remove disquette - plus de disquette
            //IJavaScriptExecutor js = (IJavaScriptExecutor)_webDriver;
            //js.ExecuteScript("$(\"span[name='IconSavedVisible']\").addClass(\"hidden\");");

            _quantityInput = WaitForElementIsVisible(By.XPath(String.Format(QUANTITY_INPUT_DN, itemName)));
            _quantityInput.SetValue(ControlType.TextBox, "3");
            // plus de disquette
            //WaitForElementIsVisible(By.XPath(String.Format(SAVE_ICON, itemName)));

            // remove disquette - plus de disquette
            //js = (IJavaScriptExecutor)_webDriver;
            //js.ExecuteScript("$(\"span[name='IconSavedVisible']\").addClass(\"hidden\");");

            var receivedQty = WaitForElementIsVisible(By.XPath(String.Format(QUANTITY_INPUT, itemName)));
            receivedQty.SetValue(ControlType.TextBox, qty.ToString());

            Thread.Sleep(2000);

            // plus de disquette
            //WaitForElementIsVisible(By.XPath(String.Format(SAVE_ICON, itemName)));

            WaitForLoad();
        }

        public ClaimEditClaimForm EditClaimForm(string itemName)
        {
            //try
            //{
            //    var extendMenuItem = WaitForElementIsVisible(By.XPath(String.Format(EXTENDED_MENU_ITEM, itemName)));
            //    extendMenuItem.Click();
            //    WaitForLoad();

            //    var plusBtn = WaitForElementIsVisible(By.XPath(String.Format(PLUS_BTN, itemName)));
            //    plusBtn.Click();
            //    WaitForLoad();
            //}
            //catch 
            //{
            WaitForLoad();
            var extendMenuItem = WaitForElementIsVisible(By.XPath(String.Format(EXTENDED_MENU_ITEMBIS, itemName)));
            extendMenuItem.Click();
            WaitForLoad();

            //var plusBtn = WaitForElementIsVisible(By.XPath(String.Format(PLUS_BTNBIS, itemName)));
            //plusBtn.Click();
            //WaitForLoad();


            //}

            return new ClaimEditClaimForm(_webDriver, _testContext);

        }

        public void AddQuantityAndTypeReceiptNote(string itemName, int qty, string type)
        {

            _quantityInput = WaitForElementIsVisible(By.XPath(String.Format(QUANTITY_INPUT_RN, itemName)));
            _quantityInput.SetValue(ControlType.TextBox, qty.ToString());

            //Carl : Commenté suite a la nouvelle version 
            //_typeSelect = WaitForElementIsVisible(By.XPath(String.Format(TYPE_SELECT_RN, itemName)));
            //_typeSelect.SetValue(ControlType.DropDownList, type);

            //WaitForElementIsVisible(By.XPath(ICON_EDIT));
        }

        public string GetQuantity(int offset = 0)
        {
            _quantity = WaitForElementExists(By.XPath(string.Format(ITEM_QUANTITY, offset)));
            return _quantity.Text;
        }

        public void SetQuantity(string qty, int offset = 0)
        {
            _quantity = WaitForElementExists(By.XPath(string.Format(ITEM_QUANTITY_FORM, offset)));
            _quantity.Clear();
            _quantity.SendKeys(qty);
            WaitForLoad();
        }


        public void SetDNQuantity(string qty, int offset = 0)
        {
            _quantityDN = WaitForElementExists(By.XPath(string.Format(ITEM_DN_QUANTITY_FORM, offset)));
            _quantityDN.Clear();
            _quantityDN.SendKeys(qty);
            WaitForLoad();
        }

        public string GetDNQuantity(int offset = 0)
        {
            _quantityDN = WaitForElementExists(By.XPath(string.Format(ITEM_DN_QUANTITY, offset)));
            return _quantityDN.Text;
        }

        public string GetClaimType()
        {
            _type = WaitForElementExists(By.XPath(ITEM_TYPE));
            var title = _type.GetAttribute("title");
            return title;
        }

        public string GetPackaging()
        {
            _packaging = WaitForElementExists(By.XPath(ITEM_PACKAGING));
            return _packaging.Text;
        }

        public void AddComment(string itemName, string comment)
        {
            _extendedMenu = WaitForElementIsVisible(By.XPath(string.Format(EXTENDED_MENU_ITEM, itemName)));
            _extendedMenu.Click();

            _addCommentDev = WaitForElementIsVisible(By.XPath(string.Format(COMMENT_ICON_DEV, itemName)));
            _addCommentDev.Click();
            WaitForLoad();

            _comment = WaitForElementIsVisible(By.Id(COMMENT));
            _comment.SetValue(ControlType.TextBox, comment);

            _saveCommentDev = WaitForElementToBeClickable(By.XPath(SAVE_COMMENT_DEV));
            _saveCommentDev.Click();

            WaitForLoad();
        }
        public string GetCommentFromEditMenu()
        {
            var megaphone = WaitForElementIsVisible(By.XPath(MEGAPHONE));
            megaphone.Click();
            var comment = WaitForElementIsVisible(By.Id(COMMENT));
            var saveButton = WaitForElementIsVisible(By.Id(SAVE));
            saveButton.Click();

            return comment.Text;
        }
        public string GetComment(string itemName)
        {
            _extendedMenu = WaitForElementIsVisible(By.XPath(string.Format(EXTENDED_MENU_ITEM, itemName)));
            _extendedMenu.Click();

            _addCommentDev = WaitForElementIsVisible(By.XPath(string.Format(COMMENT_ICON_DEV, itemName)));
            _addCommentDev.Click();
            WaitForLoad();


            if (isElementVisible(By.Id(COMMENT)))
            {
                _comment = WaitForElementExists(By.Id(COMMENT));
                return _comment.Text;
            }
            else
            {
                return "";
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
            WaitForLoad();
            _allergensList = WaitForElementIsVisible(By.XPath(string.Format(ALLERGENS_LIST)));
            var allergensInList = _allergensList.FindElements(By.TagName("li"));
            if (allergensInList.Count > 0)
                foreach (IWebElement allergen in allergensInList)
                    if (allergen.FindElements(By.TagName("img")).Count != 0) item_allergens.Add(allergen.Text);
            return item_allergens;
        }

        public void DeleteItem(string itemName)
        {
            _extendedMenu = WaitForElementIsVisible(By.XPath(string.Format(EXTENDED_MENU_ITEM, itemName)));
            _extendedMenu.Click();

            _deleteDev = WaitForElementIsVisible(By.XPath(string.Format(DELETE_ICON_DEV, itemName)));
            _deleteDev.Click();


            WaitForLoad();
        }

        public ItemGeneralInformationPage EditItem()
        {
            // via form (cad ligne sélectionnée)
            _extendedMenuForm = WaitForElementIsVisible(By.XPath(EXTENDED_MENU_FORM));
            _extendedMenuForm.Click();

            _editItemForm = WaitForElementIsVisible(By.XPath(EDIT_ITEM_FORM));
            _editItemForm.Click();
            WaitForLoad();

            // nouveau onglet !!!
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);
            WaitForLoad();

            return new ItemGeneralInformationPage(_webDriver, _testContext);
        }

        public ItemGeneralInformationPage EditItem(string itemName)
        {
            _extendedMenu = WaitForElementIsVisible(By.XPath(string.Format(EXTENDED_MENU_ITEM, itemName)));
            _extendedMenu.Click();

            _editItem = WaitForElementIsVisible(By.XPath(string.Format(EDIT_ITEM, itemName)));
            _editItem.Click();
            WaitForLoad();

            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            WaitForLoad();

            return new ItemGeneralInformationPage(_webDriver, _testContext);
        }

        public ItemGeneralInformationPage EditItemGeneralInformation(string itemName)
        {
            _extendedMenu = WaitForElementIsVisible(By.XPath(string.Format(EXTENDED_MENU_ITEM, itemName)));
            _extendedMenu.Click();

            _editItem = WaitForElementIsVisible(By.XPath(string.Format(EDIT_ITEM_INF, itemName)));
            _editItem.Click();
            WaitForLoad();

            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            WaitForLoad();

            return new ItemGeneralInformationPage(_webDriver, _testContext);
        }

        //public void AddPicture(string itemName, string imagePath)
        //{
        //    _extendedMenu = WaitForElementIsVisible(By.XPath(string.Format(EXTENDED_MENU, itemName)));
        //    _extendedMenu.Click();

        //    _addPicture = WaitForElementIsVisible(By.XPath(string.Format(PICTURE_ICON, itemName)));
        //    _addPicture.Click();
        //    WaitForLoad();

        //    _uploadPicture = WaitForElementIsVisible(By.Id(UPLOAD_PICTURE));
        //    _uploadPicture.SendKeys(imagePath);
        //    WaitPageLoading();
        //    Thread.Sleep(3000);

        //    _closePicture = WaitForElementIsVisible(By.XPath(CLOSE_PICTURE));
        //    _closePicture.Click();
        //    WaitForLoad();
        //}

        public bool HavePicture(string itemName)
        {
            _extendedMenu = WaitForElementIsVisible(By.XPath(string.Format(EXTENDED_MENU, itemName)));
            _extendedMenu.Click();

            _addPicture = WaitForElementIsVisible(By.XPath(string.Format(PICTURE_ICON, itemName)));
            _addPicture.Click();
            WaitForLoad();

            if (isElementVisible(By.XPath(DELETE_PICTURE)))
            {
                _deletePicture = _webDriver.FindElement(By.XPath(DELETE_PICTURE));
                return _deletePicture.Displayed;
            }
            else
            {
                return false;
            }
        }

        public void ShowItemsNotClaimed()
        {
            //Cette methode est deja ecrite
            _showItemsNotClaimed = WaitForElementToBeClickable(By.XPath(SHOW_ITEMS_NOT_CLAIMED));
            _showItemsNotClaimed.SetValue(ControlType.CheckBox, true);
        }

        public bool IsClaimV3Version()
        {
            if (isElementVisible(By.Id("hrefTabContentChecks")))
            {
                _webDriver.FindElement(By.Id("hrefTabContentChecks"));
                return true;
            }
            else
            {
                return false;
            }
        }

        public ClaimsFooter ClickOnFooter()
        {
            _footer = WaitForElementIsVisible(By.Id(FOOTER));
            _footer.Click();
            WaitForLoad();

            return new ClaimsFooter(_webDriver, _testContext);
        }

        public void ClickDecrStockCheckBox()
        {
            _decrStock = WaitForElementExists(By.XPath("//*[@id[starts-with(., 'btn-item-decrease-stock')]]"));
            _decrStock.SetValue(ControlType.CheckBox, true);
            WaitForLoad();
        }

        public void SetDecrQty(string qty)
        {
            _decrQty = WaitForElementIsVisible(By.XPath(DECR_STOCK_QTY));
            _decrQty.Clear();
            _decrQty.SendKeys(qty);
            WaitForLoad();
        }

        public bool IsDecrStock()
        {
            _decrStock = WaitForElementExists(By.XPath("//*[@id[starts-with(., 'btn-item-decrease-stock')]]\r\n"));
            return _decrStock.Selected;
        }

        public string GetDecrQty()
        {
            _decrQty = WaitForElementIsVisible(By.XPath(DECR_STOCK_QTY));
            return _decrQty.GetAttribute("value");
        }

        public string GetFormQty()
        {
            _qty = WaitForElementIsVisible(By.XPath(FORM_QTY));
            return _qty.GetAttribute("value");
        }

        public string GetClaimTotalAmount(int offset = 0)
        {
            _claimAmount = WaitForElementIsVisible(By.XPath(string.Format(CLAIM_AMOUNT, offset)));
            return _claimAmount.Text;
        }

        public void AddSanction(string sanction)
        {
            _sanctionInput = WaitForElementIsVisible(By.XPath(ITEM_SANCTION));
            _sanctionInput.Clear();
            _sanctionInput.SendKeys(sanction);
            WaitForLoad();
            // propagation de la valeur dans Claim attachée.
            Thread.Sleep(2000);
        }

        public string GetSanction()
        {
            _sanctionInput = WaitForElementIsVisible(By.XPath(ITEM_SANCTION));
            return _sanctionInput.GetAttribute("value");
        }

        public string GetTotalSanction()
        {
            _totalSanction = WaitForElementIsVisible(By.Id(TOTAL_SANCTION));
            return _totalSanction.Text;
        }

        public string GetTotalClaim()
        {
            _totalClaim = WaitForElementIsVisible(By.Id(TOTAL_CLAIM));
            return _totalClaim.Text;
        }
        public void SetReceviedQuantity(string quantity)
        {
            IWebElement receviedQuantityInput = WaitForElementIsVisible(By.Id("item_CndRowDto_ClaimRNQuantity"));
            receviedQuantityInput.Clear();
            receviedQuantityInput.SetValue(ControlType.TextBox, quantity);
            WaitForLoad();
        }
        public void SetReceivedQty(string receivedQty)
        {
            var receivedQtyLine = WaitForElementIsVisible(By.XPath(string.Format("/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/form/div[2]/div[6]")));
            receivedQtyLine.Click();

            var receivedQtyInput = WaitForElementIsVisible(By.XPath(string.Format(RECEIVED_QTY_LINE, 2)));
            receivedQtyInput.SetValue(ControlType.TextBox, receivedQty);
            WaitForLoad();
        }
        public bool VerifyDetailChangeLine(string receivedQty)
        {
            const string RECEIVED_QTY_LINE_BASE = "//*[@id='itemForm_{0}']/div[2]/div[5]";

            var i = 2;
            var receivedQtyLine = WaitForElementIsVisible(By.XPath(string.Format("/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/form/div[2]/div[6]")));
            receivedQtyLine.Click();

            var receivedQtyInput = WaitForElementIsVisible(By.XPath(string.Format(RECEIVED_QTY_LINE, 2)));
            receivedQtyInput.SetValue(ControlType.TextBox, receivedQty);

            receivedQtyInput.SendKeys(Keys.Enter);
            for (i = i + 1; i > 2; i++)
            {
                if (isElementExists(By.XPath(string.Format(RECEIVED_QTY_LINE_BASE, i))))
                {
                    receivedQtyInput = WaitForElementIsVisible(By.XPath(string.Format(RECEIVED_QTY_LINE_BASE, i)));
                    if (receivedQtyInput.GetAttribute("value") != receivedQty)
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
            return false;
        }
        public void ClickClaim()
        {
            var claimbutton = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/form/div[1]/div[12]/div/a"));
            claimbutton.Click();
            WaitForLoad();
        }
        public void FillClaim(string comment)
        {
            var claimCheck = WaitForElementIsVisible(By.Id("IsChecked_0"));
            claimCheck.Click();
            WaitForLoad();
            var claimComment = WaitForElementIsVisible(By.Id("Comment"));
            claimComment.SetValue(ControlType.TextBox, comment);

            var saveButton = WaitForElementIsVisible(By.Id("btn-valid-claim"));
            saveButton.Click();
        }

        public bool VerifyAdressAndSendByMail(string mail, string AdressManager)
        {
            ClosePrintButton();
            ShowExtendedMenu();

            if (IsDev())
            {
                _sendByEmailBtnDev = WaitForElementIsVisible(By.Id(SEND_BY_MAIL_BTN_DEV));
                _sendByEmailBtnDev.Click();
                WaitForLoad();
            }
            else
            {
                _sendByEmailBtn = WaitForElementIsVisible(By.Id(SEND_BY_MAIL_BTN));
                _sendByEmailBtn.Click();
                WaitForLoad();
            }

            _emailReceiver = WaitForElementIsVisible(By.Id(EMAIL_RECEIVER));
            _emailReceiver.SetValue(ControlType.TextBox, "test@mail.com");
            WaitForLoad();


            _ccMail = WaitForElementIsVisible(By.Id(CC_MAIL));
            string CCADRESS = _ccMail.GetAttribute("value");
            WaitForLoad();

            string MailUser = CCADRESS.Split(';')[0];
            string AdrManager = string.Empty;
            try
            {

                AdrManager = CCADRESS.Split(';')[1].Trim();
            }
            catch
            {
                AdrManager = string.Empty;
            }
            if ((mail.Trim() == MailUser.Trim()) && (AdressManager.Trim() == AdrManager.Trim()))
            {


                _sendEmail = WaitForElementIsVisible(By.Id(SEND_MAIL_BTN));
                _sendEmail.Click();

                int i = 5;
                while (i > 0)
                {
                    WaitPageLoading();
                    i--;
                    if (!isElementExists(By.Id(FILE_SENT)))
                    {
                        break;
                    }
                }

                return true;
            }
            else
            {
                return false;
            }
        }
        public void ClickItemsSubMenu()
        {
            ((IJavaScriptExecutor)_webDriver).ExecuteScript("window.scrollTo(0, 0);");

            var items = WaitForElementIsVisible(By.Id(ITEMS_SUB_MENU));
            ((IJavaScriptExecutor)_webDriver).ExecuteScript("arguments[0].scrollIntoView(true);", items);
            ((IJavaScriptExecutor)_webDriver).ExecuteScript("arguments[0].click();", items);
            WaitPageLoading();
            WaitForLoad();
        }
        public void ClickGeneralInformationSubMenu()
        {
            var generalinformation = WaitForElementIsVisible(By.XPath(GENERAL_INFORMATION_SUB_MENU));
            generalinformation.Click();
            WaitPageLoading();
            WaitForLoad();

        }

        public string GetItemForClaim(int row)
        {
            _claimItem = WaitForElementIsVisible(By.XPath((string.Format(CLAIM_ITEM, row))));
            return _claimItem.Text.Trim();
        }

        public bool GetActivateFromGeneralInformation()
        {
            WaitPageLoading();
            var activate= WaitForElementExists(By.Id("ClaimNote_IsActive"));
            bool result = activate.Selected;
            return result;
        }

        public void SetActivateFromGeneralInformation(bool isActivate)
        {
            var activate = WaitForElementExists(By.Id("ClaimNote_IsActive"));
            activate.SetValue(ControlType.CheckBox, isActivate);
        }
        
        public void SetCommentFromGeneralInformation(string comment)
        {
            var commentElement = WaitForElementExists(By.Id("ClaimNote_Comment"));
            commentElement.Clear();
            commentElement.SendKeys(comment);
        }

        public string GetCommentFromGeneralInformation()
        {
            WaitPageLoading();
            var commentElement = WaitForElementExists(By.Id("ClaimNote_Comment"));
            string result = commentElement.GetAttribute("value");
            return result;
        }

        public void SetStatusFromGeneralInformation(string choice)
        {
            var statusElement = WaitForElementExists(By.Id("ClaimNote_Status"));
            var selectElement = new SelectElement(statusElement);
            selectElement.SelectByValue(choice); 
        }

        public string GetStatusFromGeneralInformation()
        {
            var statusElement = WaitForElementExists(By.Id("ClaimNote_Status"));
            var selectElement = new SelectElement(statusElement);
            string selectedText = selectElement.SelectedOption.Text;
            return selectedText;
        }
        public void CloseModalComment()
        {
            var _closeModalComment = WaitForElementIsVisible(By.XPath("//*[@id=\"modal-1\"]/div/div[2]/div/form/div[2]/button[1]"));
            _closeModalComment.Click();
            WaitPageLoading();
            WaitForLoad();
        }
    }
}