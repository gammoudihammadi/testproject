using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using iText.StyledXmlParser.Jsoup.Helper;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.Trolleys;
using Newrest.Winrest.FunctionalTests.PageObjects.Interim.InterimReceptions;
using Newrest.Winrest.FunctionalTests.PageObjects.Jobs.ScheduledJobs;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Media.Animation;


namespace Newrest.Winrest.FunctionalTests.PageObjects.Interim.InterimOrders
{
    public class InterimOrdersItem : PageBase
    {
        public InterimOrdersItem(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //___________________________________________ Constantes ____________________________________________________

        // General 
        private DateTime _validationDate;
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";
        private const string GO_TO_GENERAL_INFORMATION = "hrefTabContentInformations";
        private const string ITEM_INTERIM_ORDER = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[2]";
        private const string GETQTYITEM = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[2]/div[8]";
        private const string GETQTYITEM_1 = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div/div[8]/span";
        private const string QUANTITE = "item_IodRowDto_Quantity";
        private const string COMMENT_BTN = "/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/form/div[1]/div[11]/div/a[3]/span";
        private const string COMMENT = "Comment";
        private const string SAVE_COMMENT = "//*[@id=\"modal-1\"]/div/div[2]/div/form/div[2]/button[2]";
        private const string ICON_COMMENT = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[1]/div[11]/div/a[2]/span";
        private const string NAMES_COLUMNS = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div/div[3]/span";
        private const string MSG_COMMENT = "//*[@id=\"modal-1\"]/div/p";
        private const string CLICKSHOWMENU = "//*[@id=\"div-body\"]/div/div[1]/div/div/button";
        private const string GENERATE_INTERIM_RECEPTION = "btn-generate-interim-reception";
        private const string GETPACKAGINGPRICEITEM = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[2]/div[7]/span";
        private const string GETTOTALVATITEM = "/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/form/div[2]/div[9]/span";
        // Filtres
        private const string FILTER_NAME = "tbSearchPatternWithAutocomplete";
        //SUBGROUP
        private const string SUBGROUPS_FILTER = "ItemIndexVMSelectedSubGroups_ms";
        private const string UNSELECT_ALL_GROUPS = "/html/body/div[13]";
        private const string SUBGROUPS_SEARCH = "/html/body/div[12]/div/div/label/input";
        private const string FILTER_GROUP = "ItemIndexVMSelectedGroups_ms";
        private const string LIST_GROUP = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[*]";
        private const string SUBGROUPS_NAME = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[1]/div/div/span";
        private const string SUBGROUPS_TO_CHECK = "ui-multiselect-2-ItemIndexVMSelectedSubGroups-option-0";
        private const string CHECK_BUTTON = "//*[@id=\"div-body\"]/div/div[1]/div/div[2]/button";
        private const string VALIDATE = "btn-validate-interim-order";
        private const string VALIDATE_POP_UP = "btn-popup-validate";
        private const string FIRST_LINE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[2]";
        //*[@id="list-item-with-action"]/div/div[2]/div[2]/div/div/form/div[2]
        private const string POUBELLEICON = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[1]/div[11]/div/a[1]";
        private const string QUANTITY = "item_IodRowDto_Quantity";
        private const string CRAYONICON = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[1]/div[11]/div/a[3]";
        private const string ITEMTAB = "//*[@id=\"ItemTabNav\"]/a";
        private const string DATE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div/div[4]";
        private const string LIST_ITEM = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div/div[3]/span";
        private const string ITEM = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[2]/div[3]/span";
        private const string PRICE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div/div[7]/span";
        private const string TOTALPEICE = "total-price-span";
        private const string DETAIL_BUTTON = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[1]/td[5]";
        private const string DETAIL_BUTTONS = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[2]/td[5]";
        private const string ORDERNUMBER = "//*[@id=\"div-body\"]/div/div[1]";
        private const string COPIED_INTERIM_ORDER_NAME = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[2]/div[3]";
        private const string COPIED_INTERIM_ORDER_SUPPLIER = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[2]/div[5]";
        private const string COPIED_INTERIM_ORDER_PACKING = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[2]/div[6]";
        private const string COPIED_INTERIM_ORDER_PACKAGING_PRICE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[2]/div[7]";
        private const string COPIED_INTERIM_ORDER_QUANTITY = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[2]/div[8]";
        private const string COPIED_INTERIM_ORDER_TOTAL_PRICE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[2]/div[9]";
        private const string ORIGIN_INTERIM_ORDER_NAME = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[2]/div[3]/span";
        private const string ORIGIN_INTERIM_ORDER_NAME_SELECT = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div/div[3]/span";
        private const string ORIGIN_INTERIM_ORDER_SUPPLIER = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[2]/div[5]";
        private const string ORIGIN_INTERIM_ORDER_SUPPLIER_SELECT = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div/div[5]";
        private const string ORIGIN_INTERIM_ORDER_PACKING = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[2]/div[6]";
        private const string ORIGIN_INTERIM_ORDER_PACKING_SELECT = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div/div[6]"; 
        private const string ORIGIN_INTERIM_ORDER_PACKAGING_PRICE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[2]/div[7]";
        private const string ORIGIN_INTERIM_ORDER_PACKAGING_PRICE_SELECT = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div/div[7]"; 
        private const string ORIGIN_INTERIM_ORDER_QUANTITY = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[2]/div[8]/span";
        private const string ORIGIN_INTERIM_ORDER_QUANTITY_SELECT = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div/div[8]";
        private const string ORIGIN_INTERIM_ORDER_TOTAL_PRICE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[2]/div[9]";
        private const string ORIGIN_INTERIM_ORDER_TOTAL_PRICE_SELECT = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div/div[9]";
        private const string CONFIRM_SEND = "btn-popup-send";
        private const string CANCEL = "btn-cancel-popup";
        private const string EXTENDED_BUTTON = "Extended-Menu";
        private const string SEND_INVOICES_MAIL = "btn-send-by-email-interim-order";
        private const string RESET_FILTER = "//*[@id=\"formSearchItems\"]/div[1]/a";
        private const string KEYWORD = "ItemIndexVMSelectedKeywords_ms";
        private const string LIST_SITE_FILTER = "/html/body/div[11]/ul/li[*]/label/input";
        private const string FOOTER = "hrefTabContentFooter";
        private const string MODAL_FORM = "//*[@id=\"modal-1\"]/div/div/div/form";
        private const string ATTACHED_FILE = "//*[@id=\"modal-1\"]/div/div/div/form/div[2]/div/div[4]/div/div/div/div/div/span";
        private const string GROUPE_COMBO_DEFAULT_VALUE = "//*[@id='ItemIndexVMSelectedGroups_ms']/span[2]";
        private const string SUBGROUP_COMBO_DEFAULT_VALUE = "//*[@id=\"ItemIndexVMSelectedSubGroups_ms\"]/span[2]";
        private const string ATTACHEMENTS_FILE="//*[@id=\"modal-1\"]/div/div/div/form/div[2]/div/div[4]/div/div/div/div/div/span";
        private const string SEND_BY_MAIL_POPUP = "//*[@id=\"modal-1\"]/div/div/div/form";
        private const string TOTALVAT_ITEM = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[2]/div[9]/span";
        private const string SELECTED_PROD_QTY_INTERIM_ORDER_ITEM = "/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/form/div/div[8]/span";
        private const string SELECTED_TOTAL_VAT_INTERIM_ORDER_ITEM = "/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/form/div/div[9]/span";
        private const string ITEMS = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form";

        public const string TOADRESSES = "ToAddresses";
        public const string SEND = "btn-init-async-send-mail";
        public const string TEXT = "//*[@id=\"modal-1\"]/div/div/div/form/div[2]/div/div[6]/div/div/div/div/div[3]/div[2]";
        public const string ADD_COMMENT = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[1]/div[11]/div/a[2]/span";

        //___________________________________________ Variables ________________________________________________________________
        [FindsBy(How = How.Id, Using = TOADRESSES)]
        private IWebElement _toAdresses;
        [FindsBy(How = How.Id, Using = SEND)]
        private IWebElement _sendByEnterimOrder;
        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;
        [FindsBy(How = How.XPath, Using = ADD_COMMENT)]
        private IWebElement _commentaddbtn;

        [FindsBy(How = How.XPath, Using = FIRST_LINE)]
        private IWebElement _firstLine;
        [FindsBy(How = How.XPath, Using = CRAYONICON)]
        private IWebElement _crayonIcon;

        [FindsBy(How = How.XPath, Using = TEXT)]
        private IWebElement _text;
        [FindsBy(How = How.Id, Using = GO_TO_GENERAL_INFORMATION)]
        private IWebElement _goToGeneralInformation;

        [FindsBy(How = How.XPath, Using = ITEM_INTERIM_ORDER)]
        private IWebElement _itemInterimOrder;

        [FindsBy(How = How.XPath, Using = GETQTYITEM)]
        private IWebElement _getQtyItem;

        [FindsBy(How = How.XPath, Using = GETPACKAGINGPRICEITEM)]
        private IWebElement _getPackagingPriceItem;

        [FindsBy(How = How.XPath, Using = GETTOTALVATITEM)]
        private IWebElement _getTotalVATItem;

        [FindsBy(How = How.Id, Using = QUANTITE)]
        private IWebElement _quantite;

        [FindsBy(How = How.XPath, Using = COMMENT_BTN)]
        private IWebElement _commentbtn;

        [FindsBy(How = How.Id, Using = COMMENT)]
        private IWebElement _comment;

        [FindsBy(How = How.XPath, Using = SAVE_COMMENT)]
        private IWebElement _saveBtn;

        [FindsBy(How = How.XPath, Using = ICON_COMMENT)]
        private IWebElement _iconComment;

        [FindsBy(How = How.XPath, Using = MSG_COMMENT)]
        private IWebElement _msgComment;

        [FindsBy(How = How.XPath, Using = CLICKSHOWMENU)]
        private IWebElement _clickshowmenu;

        [FindsBy(How = How.Id, Using = GENERATE_INTERIM_RECEPTION)]
        private IWebElement _generateInterimReception;

        [FindsBy(How = How.XPath, Using = DATE)]
        private IWebElement _date;

        [FindsBy(How = How.XPath, Using = ITEM)]
        private IWebElement _item;

        [FindsBy(How = How.XPath, Using = ORDERNUMBER)]
        private IWebElement _orderNumber;
        //_______________________________________  Filter _____________________________________________________________

        [FindsBy(How = How.Id, Using = FILTER_NAME)]
        private IWebElement _searchByNameFilter;
        [FindsBy(How = How.Id, Using = FILTER_GROUP)]
        private IWebElement _searchByFilter;

        //SUBGROUP
        [FindsBy(How = How.Id, Using = SUBGROUPS_FILTER)]
        public IWebElement _subgroupsFilter;
        [FindsBy(How = How.XPath, Using = UNSELECT_ALL_GROUPS)]
        public IWebElement _unselectAllsubGroups;
        [FindsBy(How = How.XPath, Using = SUBGROUPS_SEARCH)]
        public IWebElement _searchsubGroups;

        [FindsBy(How = How.Id, Using = SUBGROUPS_TO_CHECK)]
        private IWebElement _subgroupToCheck;

        [FindsBy(How = How.Id, Using = VALIDATE)]
        private IWebElement _validate;

        [FindsBy(How = How.Id, Using = VALIDATE_POP_UP)]
        private IWebElement _validatePopUp;

        [FindsBy(How = How.XPath, Using = SEND_INVOICES_MAIL)]
        private IWebElement _checkButton;

        [FindsBy(How = How.XPath, Using = COPIED_INTERIM_ORDER_NAME)]
        private IWebElement _copiedInterimOrderName;

        [FindsBy(How = How.XPath, Using = COPIED_INTERIM_ORDER_SUPPLIER)]
        private IWebElement _copiedInterimOrderSupplier;

        [FindsBy(How = How.XPath, Using = COPIED_INTERIM_ORDER_PACKING)]
        private IWebElement _copiedInterimOrderPacking;

        [FindsBy(How = How.XPath, Using = COPIED_INTERIM_ORDER_PACKAGING_PRICE)]
        private IWebElement _copiedInterimOrderPackagingPrice;

        [FindsBy(How = How.XPath, Using = COPIED_INTERIM_ORDER_QUANTITY)]
        private IWebElement _copiedInterimOrderQuantity;

        [FindsBy(How = How.XPath, Using = COPIED_INTERIM_ORDER_TOTAL_PRICE)]
        private IWebElement _copiedInterimOrderTotalPrice;

        [FindsBy(How = How.XPath, Using = ORIGIN_INTERIM_ORDER_NAME)]
        private IWebElement _originInterimOrderName;

        [FindsBy(How = How.XPath, Using = ORIGIN_INTERIM_ORDER_SUPPLIER)]
        private IWebElement _originInterimOrderSupplier;

        [FindsBy(How = How.XPath, Using = ORIGIN_INTERIM_ORDER_PACKING)]
        private IWebElement _originInterimOrderPacking;

        [FindsBy(How = How.XPath, Using = ORIGIN_INTERIM_ORDER_PACKAGING_PRICE)]
        private IWebElement _originInterimOrderPackagingPrice;

        [FindsBy(How = How.XPath, Using = ORIGIN_INTERIM_ORDER_QUANTITY)]
        private IWebElement _originInterimOrderQuantity;

        [FindsBy(How = How.XPath, Using = ORIGIN_INTERIM_ORDER_TOTAL_PRICE)]
        private IWebElement _originInterimOrderTotalPrice;

        [FindsBy(How = How.Id, Using = TOTALPEICE)]
        private IWebElement _totalPrice;

        [FindsBy(How = How.XPath, Using = DETAIL_BUTTON)]
        private IWebElement _detail_Button;

        [FindsBy(How = How.Id, Using = CONFIRM_SEND)]
        private IWebElement _confirmSendMail;

        [FindsBy(How = How.Id, Using = CANCEL)]
        private IWebElement _cancel;

        [FindsBy(How = How.Id, Using = EXTENDED_BUTTON)]
        private IWebElement _extendedButton;

        [FindsBy(How = How.XPath, Using = RESET_FILTER)]
        private IWebElement _resetFilter;

        [FindsBy(How = How.Id, Using = KEYWORD)]
        private IWebElement _keyword;

        [FindsBy(How = How.Id, Using = FOOTER)]
        private IWebElement _footer;
        [FindsBy(How = How.XPath, Using = MODAL_FORM)]
        private IWebElement _modalForm;
        [FindsBy(How = How.XPath, Using = ATTACHED_FILE)]
        private IWebElement _attachedfile;

        [FindsBy(How = How.XPath, Using = SELECTED_TOTAL_VAT_INTERIM_ORDER_ITEM)]
        private IWebElement _selectedTotalVatInterimOrderItem;
        //___________________________________________ Méthodes _____________________________________________________

        // General       
        public enum FilterItemType
        {
            SearchByNameRef,
            SearchByGroup,
            Keyword,
            subGroup
        }
        public void Filter(FilterItemType FilterItemType, object value)
        {
            switch (FilterItemType)
            {
                case FilterItemType.subGroup:
                    ComboBoxSelectById(new ComboBoxOptions(SUBGROUPS_FILTER, (string)value));
                    break;
                case FilterItemType.SearchByNameRef:
                    WaitForElementIsVisible(By.Id(FILTER_NAME));
                    _searchByNameFilter.SetValue(ControlType.TextBox, value);
                    break;

                case FilterItemType.SearchByGroup:
                    ComboBoxSelectById(new ComboBoxOptions(FILTER_GROUP, (string)value));
                    break;

                case FilterItemType.Keyword:
                    ComboBoxSelectById(new ComboBoxOptions(KEYWORD, (string)value));
                    break;

            }

            WaitPageLoading();
            WaitForLoad();
        }


        public object GetFilterValue(FilterItemType filterType)
        {
            switch (filterType)
            {
                case FilterItemType.SearchByGroup:
                    var _searchByFilter = WaitForElementExists(By.XPath(FILTER_GROUP));
                    return _searchByFilter.Selected;
                case FilterItemType.Keyword:
                    var _keyword = WaitForElementExists(By.Id(KEYWORD));
                    return _keyword.Selected;

                case FilterItemType.subGroup:
                    var _subgroupsFilter = WaitForElementExists(By.Id(SUBGROUPS_FILTER));
                    return _subgroupsFilter.Selected;
            }

            return null;
        }

        public InterimOrdersPage BackToList()
        {
            WaitPageLoading(); 
            _backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            _backToList.Click();
            WaitForLoad();
            return new InterimOrdersPage(_webDriver, _testContext);
        }

        public InterimOrdersGeneralInformation GoToGeneralInformation()
        {
            _goToGeneralInformation = WaitForElementIsVisible(By.Id(GO_TO_GENERAL_INFORMATION));
            _goToGeneralInformation.Click();
            WaitForLoad();

            return new InterimOrdersGeneralInformation(_webDriver, _testContext);
        }
        public void SetQty(string qty)
        {

            if (!isElementVisible(By.Id(QUANTITE)))
            {
                _itemInterimOrder = WaitForElementIsVisible(By.XPath(ITEM_INTERIM_ORDER));
                _itemInterimOrder.Click();
            }
            _quantite = WaitForElementIsVisible(By.Id(QUANTITE));
            _quantite.SetValue(PageBase.ControlType.TextBox, qty);
            WaitForLoad();
            LoadingPage();
            WaitPageLoading();
        }

        public enum FilterType
        {
            subGroup
        }

        public void Filter(FilterType filterType, object value)
        {
            Actions action = new Actions(_webDriver);

            switch (filterType)
            {
                case FilterType.subGroup:
                    _subgroupsFilter = WaitForElementIsVisible(By.Id(SUBGROUPS_FILTER));
                    _subgroupsFilter.Click();

                    _unselectAllsubGroups = WaitForElementIsVisible(By.XPath(UNSELECT_ALL_GROUPS));
                    _unselectAllsubGroups.Click();

                    _searchsubGroups = WaitForElementIsVisible(By.XPath(SUBGROUPS_SEARCH));
                    _searchsubGroups.SetValue(ControlType.TextBox, value);
                    WaitForLoad();

                    _subgroupToCheck = WaitForElementExists(By.Id(SUBGROUPS_TO_CHECK));
                    _subgroupToCheck.Click();

                    _subgroupsFilter.Click();
                    break;



                default:
                    throw new ArgumentOutOfRangeException(nameof(filterType), filterType, "Invalid filter type provided.");
            }

            WaitPageLoading();
            WaitForLoad();
        }

        public bool FindRowItemGroup(string subgrpname)
        {
            var itemrow = _webDriver.FindElements(By.XPath(SUBGROUPS_NAME));
            foreach (var item in itemrow)
            {
                if (item.Text.Equals(subgrpname))
                {
                    return true;
                }
            }
            return false;

        }
        public void ClickFirstItem()
        {
            _itemInterimOrder = WaitForElementIsVisible(By.XPath(ITEM_INTERIM_ORDER));
            _itemInterimOrder.Click();
            WaitForLoad();
        }

        public string GetQtyItem()
        {
            _getQtyItem = WaitForElementIsVisible(By.XPath(GETQTYITEM));
            return _getQtyItem.Text.Trim();

        }
        public bool HasPricePackaging()
        {
            _getPackagingPriceItem = WaitForElementIsVisible(By.XPath(GETPACKAGINGPRICEITEM));
            var priceText = _getPackagingPriceItem.Text.Trim().Replace("€", "");
            if (decimal.Parse(priceText) > 0)
            {
                return true;
            }
            else
            {
                return false;

            }

        }

        public bool HasPrice()
        {
            bool itemExist = isElementVisible(By.XPath(PRICE));
            return itemExist;

        }

        public bool ExistQtyItem()
        {
            var test = isElementVisible(By.XPath(GETQTYITEM_1));
            return test;

        }
        public string GetPackagingPriceItem()
        {
            _getPackagingPriceItem = WaitForElementIsVisible(By.XPath(GETPACKAGINGPRICEITEM));
            return _getPackagingPriceItem.Text.Trim().Substring(2).Replace(",", "");


        }

        public void OpenModalAddComment()
        {
            _commentbtn = WaitForElementIsVisible(By.XPath(COMMENT_BTN));
            _commentbtn.Click();
            WaitForLoad(); 
        }
        public void AddComment()
        {
            _commentaddbtn = WaitForElementIsVisible(By.XPath(ADD_COMMENT));
            _commentaddbtn.Click();
            WaitForLoad();
        }
        public void SetComment(string comment)
        {
            if (isElementVisible(By.Id(COMMENT)))
            {
                _comment = WaitForElementIsVisible(By.Id(COMMENT));
                _comment.SetValue(PageBase.ControlType.TextBox, comment);
                _comment.SendKeys(Keys.Enter);
                _saveBtn = WaitForElementIsVisible(By.XPath(SAVE_COMMENT));
                _saveBtn.Click();
                WaitForLoad();
            }

        }
        public bool isCommentIconGreen()
        {
            int expectedR = 0;
            int expectedG = 128;
            int expectedB = 0;

            var _iconComment = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/form/div[1]/div[11]/div/a[2]/span"));
            var iconColor = _iconComment.GetCssValue("color");

            var colorMatch = System.Text.RegularExpressions.Regex.Match(iconColor, @"rgba?\((\d+),\s*(\d+),\s*(\d+)");

            if (colorMatch.Success)
            {
                int r = int.Parse(colorMatch.Groups[1].Value);
                int g = int.Parse(colorMatch.Groups[2].Value);
                int b = int.Parse(colorMatch.Groups[3].Value);

                return r == expectedR && g == expectedG && b == expectedB;
            }

            return false;
        }


        public List<string> GetAllNamesResultPaged()
        {
            var names = new List<string>();

            var elements = _webDriver.FindElements(By.XPath(LIST_ITEM));
            if (elements == null)
            {
                return new List<string>();
            }
            foreach (var element in elements)
            {
                var trimmedText = element.Text.Trim();
                var nameWithoutParentheses = Regex.Replace(trimmedText, @"\s*\(.*?\)\s*", "");
                names.Add(nameWithoutParentheses);
            }
            return names;
        }
        public string GetCommentMsgItem()
        {
            _msgComment = WaitForElementIsVisible(By.XPath(MSG_COMMENT));
            return _msgComment.Text.Trim();

        }
        public void ShowExtendedMenu()
        {
            _clickshowmenu = WaitForElementIsVisible(By.XPath(CLICKSHOWMENU));
            _clickshowmenu.Click();
            WaitForLoad();
        }
        public InterimOrdersCreateModalPage GoToGenerateInterimReception()
        {
            _generateInterimReception = WaitForElementIsVisible(By.Id(GENERATE_INTERIM_RECEPTION));
            _generateInterimReception.Click();
            WaitForLoad();


            return new InterimOrdersCreateModalPage(_webDriver, _testContext);
        }

        public InterimReceptionsCreateModalPage GenerateIntreimReception()
        {
            _generateInterimReception = WaitForElementExists(By.Id(GENERATE_INTERIM_RECEPTION));
            _generateInterimReception.Click();
            WaitForLoad();


            return new InterimReceptionsCreateModalPage(_webDriver, _testContext);
        }
        public int GetNumberOfDisplayedRows()
        {

            var elements = _webDriver.FindElements(By.XPath(LIST_GROUP));
            return elements.Count;
        }

        public string GetTotalVATItem()
        {
            _getTotalVATItem = WaitForElementIsVisible(By.XPath(TOTALVAT_ITEM));
            return _getTotalVATItem.Text.Trim().Substring(2).Replace(",", "");

        }


        public void Validate()
        {
            _checkButton = WaitForElementExists(By.XPath(CHECK_BUTTON));
            _checkButton.Click();
            _validate = WaitForElementExists(By.Id(VALIDATE));
            _validate.Click();
            _validatePopUp = WaitForElementExists(By.Id(VALIDATE_POP_UP));
            _validatePopUp.Click();
            SetValidationDate();
            WaitForLoad();
            LoadingPage();
            WaitPageLoading();
        }
        public string GetDate()
        {
            _date = WaitForElementIsVisible(By.XPath(DATE));
            return _date.Text.Trim();
        }
        public void ClickOnItem()
        {
            WaitLoading();
            _firstLine = _webDriver.FindElement(By.XPath(FIRST_LINE));
            _firstLine.Click();
        }
        public void ClickOnPoubelle()
        {
            var PoubelleElement = _webDriver.FindElement(By.XPath(POUBELLEICON));

            PoubelleElement.Click();
            WaitPageLoading();
        }

        public void ClickOnCrayon()
        {
            _crayonIcon = _webDriver.FindElement(By.XPath(CRAYONICON));
            _crayonIcon.Click();
            WaitPageLoading();
        }

        public ItemGeneralInformationPage EditButton()
        {

            _itemInterimOrder = WaitForElementIsVisible(By.XPath(ITEM_INTERIM_ORDER));
            _itemInterimOrder.Click();
            var CrayonIcon = _webDriver.FindElement(By.XPath(CRAYONICON));
            CrayonIcon.Click();
            WaitPageLoading();
            return new ItemGeneralInformationPage(_webDriver, _testContext);


        }

        public void SetQuantity()
        {
            var Quantity = _webDriver.FindElement(By.Id(QUANTITE));

            Quantity.SetValue(ControlType.TextBox, "100");
        }

        public string GetQuantity()
        {
            var Quantity = _webDriver.FindElement(By.Id(QUANTITE));

            return Quantity.GetAttribute("value");
        }

        public bool verifyOpenedItem()
        {
            if (isElementExists(By.XPath(ITEMTAB)))
            {
                return true;
            }
            return false;
        }
        public void ValidatePopUp()
        {
            _checkButton = WaitForElementIsVisible(By.XPath(CHECK_BUTTON));
            _checkButton.Click();
            _validate = WaitForElementIsVisible(By.Id(VALIDATE));
            _validate.Click();
            WaitForLoad();

        }
        public bool IsModalVisible()
        {

            _modalForm = _webDriver.FindElement(By.XPath(MODAL_FORM));
            return _modalForm.Displayed;
        }

        public string ReturnItemName()
        {
           var item = WaitForElementIsVisible(By.XPath(ITEM));
            return item.Text;
        }

        public string GetFirstItemName()
        {
            _item = WaitForElementIsVisible(By.XPath(ITEM));
            return _item.Text;
        }

        public string GetCopiedInterimOrderName()
        {
            _copiedInterimOrderName = WaitForElementIsVisible(By.XPath(COPIED_INTERIM_ORDER_NAME));
            return _copiedInterimOrderName.Text;
        }

        public string GetCopiedInterimOrderSupplier()
        {
            _copiedInterimOrderSupplier = WaitForElementIsVisible(By.XPath(COPIED_INTERIM_ORDER_SUPPLIER));
            return _copiedInterimOrderSupplier.Text;
        }

        public string GetCopiedInterimOrderPacking()
        {
            _copiedInterimOrderPacking = WaitForElementIsVisible(By.XPath(COPIED_INTERIM_ORDER_PACKING));
            return _copiedInterimOrderPacking.Text;
        }

        public decimal GetCopiedInterimOrderPackagingPrice()
        {
            _copiedInterimOrderPackagingPrice = WaitForElementIsVisible(By.XPath(COPIED_INTERIM_ORDER_PACKAGING_PRICE));
            var priceText = _copiedInterimOrderPackagingPrice.Text;

            var cleanPriceText = priceText.Replace("€", "").Trim();

            return decimal.Parse(cleanPriceText);
        }

        public int GetCopiedInterimOrderQuantity()
        {
            _copiedInterimOrderQuantity = WaitForElementIsVisible(By.XPath(COPIED_INTERIM_ORDER_QUANTITY));
            return int.Parse(_copiedInterimOrderQuantity.Text);
        }

        public decimal GetCopiedInterimOrderTotalPrice()
        {
            _copiedInterimOrderTotalPrice = WaitForElementIsVisible(By.XPath(COPIED_INTERIM_ORDER_TOTAL_PRICE));
            var priceText = _copiedInterimOrderTotalPrice.Text;

            var cleanPriceText = priceText.Replace("€", "").Trim();
            WaitPageLoading(); 
            return decimal.Parse(cleanPriceText);
        }

        public string GetOriginInterimOrderName()
        {
            WaitForLoad();
            _originInterimOrderName = WaitForElementIsVisible(By.XPath(ORIGIN_INTERIM_ORDER_NAME));
            return _originInterimOrderName.Text;
        }
        public string GetOriginInterimOrderNameSelectsFirstInterim()
        {
            WaitForLoad();
            _originInterimOrderName = WaitForElementIsVisible(By.XPath(ORIGIN_INTERIM_ORDER_NAME_SELECT));
            return _originInterimOrderName.Text;
        }

        public string GetOriginInterimOrderSupplier()
        {
            WaitForLoad();
            _originInterimOrderSupplier = WaitForElementIsVisible(By.XPath(ORIGIN_INTERIM_ORDER_SUPPLIER));
            return _originInterimOrderSupplier.Text;
        }
        public string GetOriginInterimOrderSupplierSelectsFirstInterim()
        {
            WaitForLoad();
            _originInterimOrderSupplier = WaitForElementIsVisible(By.XPath(ORIGIN_INTERIM_ORDER_SUPPLIER_SELECT));
            return _originInterimOrderSupplier.Text;
        }

        public string GetOriginInterimOrderPacking()
        {
            _originInterimOrderPacking = WaitForElementIsVisible(By.XPath(ORIGIN_INTERIM_ORDER_PACKING));
            return _originInterimOrderPacking.Text;
        }
        public string GetOriginInterimOrderPackingSelectsFirstInterim()
        {
            _originInterimOrderPacking = WaitForElementIsVisible(By.XPath(ORIGIN_INTERIM_ORDER_PACKING_SELECT));
            return _originInterimOrderPacking.Text;
        }

        public decimal GetOriginInterimOrderPackagingPrice()
        {
            _originInterimOrderPackagingPrice = WaitForElementExists(By.XPath(ORIGIN_INTERIM_ORDER_PACKAGING_PRICE));
            var priceText = _originInterimOrderPackagingPrice.Text;

            var cleanPriceText = priceText.Replace("€", "").Trim();

            return decimal.Parse(cleanPriceText);
        }
        public decimal GetOriginInterimOrderPackagingPriceSelectsFirstInterim()
        {
            _originInterimOrderPackagingPrice = WaitForElementExists(By.XPath(ORIGIN_INTERIM_ORDER_PACKAGING_PRICE_SELECT));
            var priceText = _originInterimOrderPackagingPrice.Text;

            var cleanPriceText = priceText.Replace("€", "").Trim();

            return decimal.Parse(cleanPriceText);
        }

        public int GetOriginInterimOrderQuantity()
        {
            _originInterimOrderQuantity = WaitForElementExists(By.XPath(ORIGIN_INTERIM_ORDER_QUANTITY));
            return int.Parse(_originInterimOrderQuantity.Text);
        }
        public int GetOriginInterimOrderQuantitySelectsFirstInterim()
        {
            _originInterimOrderQuantity = WaitForElementExists(By.XPath(ORIGIN_INTERIM_ORDER_QUANTITY_SELECT));
            return int.Parse(_originInterimOrderQuantity.Text);
        }
        public decimal GetOriginInterimOrderTotalPrice()
        {
            _originInterimOrderTotalPrice = WaitForElementExists(By.XPath(ORIGIN_INTERIM_ORDER_TOTAL_PRICE));
            var priceText = _originInterimOrderTotalPrice.Text;

            var cleanPriceText = priceText.Replace("€", "").Trim();

            return decimal.Parse(cleanPriceText);
        }
        public decimal GetOriginInterimOrderTotalPriceSelectsFirstInterim()
        {
            _originInterimOrderTotalPrice = WaitForElementExists(By.XPath(ORIGIN_INTERIM_ORDER_TOTAL_PRICE_SELECT));
            var priceText = _originInterimOrderTotalPrice.Text;

            var cleanPriceText = priceText.Replace("€", "").Trim();

            return decimal.Parse(cleanPriceText);
        }
        private void SetValidationDate()
        {
            _validationDate = DateTime.Now;
        }
        public string GetFormattedValidationDate()
        {
            return _validationDate.ToString("yyyy-MM-dd");
        }
        public DateTime GetValidationDate()
        {
            return _validationDate;
        }
        public void ScrollUntilResetFilterIsVisible()
        {
            WaitForLoad();
            ScrollUntilElementIsInView(By.XPath(RESET_FILTER));
            WaitForLoad();
        }
        public void ResetFilters()
        {
            _resetFilter = WaitForElementIsVisible(By.XPath(RESET_FILTER));
            _resetFilter.Click();
            WaitForLoad();
        }
        public int GetNumberSelectedSiteFilter()
        {
            var listSitesSelectedFilters = _webDriver.FindElements(By.XPath(LIST_SITE_FILTER));
            var numberSitesSelectedSite = listSitesSelectedFilters
               .Where(p => p.Selected == true).Count();

            return numberSitesSelectedSite;
        }
        public string GetGroupSelectionStatus()
        {
            var _getGroupSelectionStatus = WaitForElementIsVisible(By.XPath(GROUPE_COMBO_DEFAULT_VALUE));
            return _getGroupSelectionStatus.Text.Trim();
        }

        public string GetsubGroupSelectionStatus()
        {
            var _grtsubgroupSelectionXPath = WaitForElementIsVisible(By.XPath(SUBGROUP_COMBO_DEFAULT_VALUE));
            return _grtsubgroupSelectionXPath.Text.Trim();
        }
        public string GetCopiedInterimOrderNO()
        {
            WaitForLoad();
            var copiedInterimOrderName = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div[1]/h1"));
            string fullOrderName = copiedInterimOrderName.Text;
            string numericPart = Regex.Match(fullOrderName, @"\d+").Value;
            return numericPart;
        }

        public void SendByMail()
        {
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            ShowExtendedMenu();
            WaitPageLoading();
            _checkButton = WaitForElementIsVisible(By.Id(SEND_INVOICES_MAIL));
            _checkButton.Click();
            WaitPageLoading();
        }
        public void SendByEmailEnterimOrder(string userEmail, string text = null)
        {
            WaitForLoad();
            _toAdresses = WaitForElementIsVisible(By.Id(TOADRESSES));
            _toAdresses.SetValue(ControlType.TextBox, userEmail);
            WaitForLoad();
            if (text != null)
            {
                var textArea = WaitForElementIsVisible(By.XPath("//div[@class='note-editable' and @contenteditable='true']"));
                textArea.Clear();
                textArea.SendKeys(text);
            }
            if (IsDev())
            {
                _sendByEnterimOrder = WaitForElementIsVisible(By.Id(SEND));
            }
            else
            {
                _sendByEnterimOrder = WaitForElementIsVisible(By.Id("btn-init-async-send-mail"));
            }
            _sendByEnterimOrder.Click();
            WaitPageLoading();
            WaitForLoad();
        }
        public bool IsEmailPopupDisplayed()
        {
            _modalForm = _webDriver.FindElement(By.XPath(MODAL_FORM));
            return _modalForm.Displayed;
        }
        public string GetAttachedFiles()
        {
            _attachedfile = WaitForElementIsVisible(By.XPath(ATTACHED_FILE));
            return _attachedfile.Text;

        }

        public decimal GetTotalItemPrice()
        {
            _totalPrice = WaitForElementIsVisible(By.Id(TOTALPEICE));
            var priceText = _totalPrice.Text;

            var cleanPriceText = priceText.Replace("€", "").Trim();

            return decimal.Parse(cleanPriceText);
        }
        public InterimOrdersFooter GoToFooter()
        {
            var footer = WaitForElementIsVisible(By.Id(FOOTER));
            footer.Click();

            WaitForLoad();
            return new InterimOrdersFooter(_webDriver, _testContext);
        }
        public float GetTotal()
        {
            _totalPrice = WaitForElementIsVisible(By.Id(TOTALPEICE));
            string price =_totalPrice.Text.Trim();
            string cleanedPrice = Regex.Replace(price, @"[^\d,]", "");

            cleanedPrice = cleanedPrice.Replace(',', '.');
            if (decimal.TryParse(cleanedPrice, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal result))
            {
                return (float)result;
            }
            throw new FormatException("Total price is not Parsed ");


        }
        public string ReturnInterimOrderNumber()
        {
            _orderNumber = WaitForElementIsVisible(By.XPath(ORDERNUMBER));
            string orderNumber =_orderNumber.Text.Trim();
            Match match = Regex.Match(orderNumber, @"\d+");

            // Check if a match was found and return the numeric part
            if (match.Success)
            {
                return match.Value;
            }
            return orderNumber;


        }
        public int GetSelectedProdQtyInterimOrderItem()
        {
            WaitPageLoading();
            var prodQty = WaitForElementExists(By.XPath(SELECTED_PROD_QTY_INTERIM_ORDER_ITEM));
            return int.Parse(prodQty.Text);
        }
        public decimal GetSelectedTotalVatInterimOrderItem()
        {
            WaitPageLoading();
            _selectedTotalVatInterimOrderItem = WaitForElementExists(By.XPath(SELECTED_TOTAL_VAT_INTERIM_ORDER_ITEM));
            var totalVat = _selectedTotalVatInterimOrderItem.Text;

            var cleanTotalVatText = totalVat.Replace("€", "").Trim();

            return decimal.Parse(cleanTotalVatText);
        }
        public bool HasItems()
        {
            var itemRows = _webDriver.FindElements(By.XPath(ITEMS)); 

            return itemRows.Count > 0;
        }
       
    }
}

