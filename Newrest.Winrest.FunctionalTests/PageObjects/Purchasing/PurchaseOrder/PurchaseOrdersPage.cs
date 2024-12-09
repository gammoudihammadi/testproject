using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
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
using System.Security.Policy;
using System.Text.RegularExpressions;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Purchasing
{
    public class PurchaseOrdersPage : PageBase
    {

        public PurchaseOrdersPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // _________________________________________ Constantes ______________________________________________

        // Général
        private const string NEW_PURCHASE_ORDER = "New purchase order";

        private const string EXPORT = "btn-export-excel";
        private const string PRINT_LIST_PO = "Print purchase orders";
        private const string PRINT_PO = "btn_print_po";
        private const string PRINT_WITH_PRICES_CHECKBOX = "//label[contains(text(), 'Include prices on purchase order print.')]/../div";
        private const string CONFIRM_PRINT = "printButton";
        private const string EXPORT_WMS = "exportBtnToTXT";
        private const string CONFIRM_EXPORT_WMS = "btnExport";
        private const string SEND_RESULTS_BY_MAIL = "btn-send-all-by-email"; 
        private const string CONFIRM_SEND_MAIL = "btn-init-async-send-mail";

        // Tableau
        private const string FIRST_PO = "//*[@id=\"list-item-with-action\"]/table/tbody/tr";
        private const string FIRST_PO_NUMBER = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[1]/td[2]";
        private const string ALL_PO_NUMBERS = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[2]";
        private const string VALIDATION = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[{0}]/td[1]/img[@alt='Valid']";
        private const string INACTIVE = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[{0}]/td[1]/img[@alt='Inactive']";
        private const string NUMBER = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[2]";

        private const string VERIFY_SITE = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[3]";
        private const string VERIFY_SUPPLIER = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[4]";
        private const string SENT_EDI = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[{0}]/td[5]/div[1]/i";
        private const string SENDING_BY_MAIL = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[{0}]/td[5]/div[2]/span";
        private const string NOT_RECEIVED = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[{0}]/td[5]/div[3]/img";
        private const string DATE = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[6]";

        // Filtres
        private const string RESET_FILTER_DEV = "ResetFilter";

        private const string SEARCH_FILTER = "SearchNumber";
        private const string SO_NUMBER_FILTER = "SearchNumberSecondary";
        private const string SUPPLIER_FILTER = "SelectedSuppliers_ms";
        private const string UNCHECK_ALL_SUPPLIER = "/html/body/div[10]/div/ul/li[2]/a";
        private const string SEARCH_SUPPLIER = "/html/body/div[10]/div/div/label/input";
        private const string DATE_FROM = "date-picker-start";
        private const string DATE_TO = "date-picker-end";

        private const string VALIDATION_DATE_DEV = "SearchByValidationDate";

        private const string SHOW_FILTER = "cbSortByShowOrder";
        private const string APPROVED_BY = "cbFilterApproval";

        private const string SHOW_ALL_DEV = "ShowAll";
        private const string SHOW_ACTIVE_DEV = "ShowOnlyActive";
        private const string SHOW_INACTIVE_DEV = "ShowOnlyInactive";

        private const string RECEIPT_STATUS_FILTER = "cbReceiptStatus";
        private const string SORTBY_FILTER = "cbSortBy";
        private const string TO_SEND_BY_MAIL = "ShowSpecificItems";
        private const string STATUS_ALL = "status-all";
        private const string STATUS_OPENED = "status-opened";
        private const string STATUS_CLOSED = "status-closed";
        private const string STATUS_CANCELLED = "status-cancelled";
        private const string SITE_FILTER = "SelectedSites_ms";
        private const string UNCHECK_ALL_SITE = "/html/body/div[11]/div/ul/li[2]/a";
        private const string SEARCH_SITE = "/html/body/div[11]/div/div/label/input";

        private const string FIRST_PURCHASE_ORDER_DELETE = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[1]/td[9]/a/span";
        private const string NUMBERS_LIST = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[2]";
        private const string CONFIRM_DELETE = "dataConfirmOK";
        private const string FIRST_LINE = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[1]";
        private const string FIRST_ITEM = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]";
        private const string QUANTITY = "item_RelatedPurchaseOrderItemVM_PurchaseOrderItemDetail_Quantity";
        private const string TOTAL_VAT = "//*[@id=\"total-price-span\"]";
        private const string FIRST_ITEM_NUMBER = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[1]/div[3]/span";
        private const string ACTIONS_BUTTONS = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[1]/div[14]/div/a/span";
        private const string DELETE = "delete-event";
        private const string PEN_BTN = "edit-item-event";
        private const string EXTENDED_MENU_BTN = "//*[@id=\"div-body\"]/div/div[1]/div/div";
        private const string SEND_BY_EMAIL = "btn-send-by-email-purchase-order";
        private const string SEND_EMAIL_CONFIRM = "btn-popup-send";
        private const string TO_EMAIL = "ToAddresses";
        private const string ENVELOPPE = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[1]/td[5]/div[2]/span";
        private const string FOOTER = "hrefTabContentFooter";
        private const string TOTAL_GROSS_AMOUNT = "//*[@id=\"tabContentDetails\"]/div/table[1]/tbody/tr[4]/td[2]";
        private const string ROWS_FOR_PAGINATION = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[2]";
        private const string FIRST_ITEM_FILTRED = "//*[@id=\"purchasing-purchaseorders-details-1\"]";
        private const string ITEM_NAME = "//*[@id=\"listOfItems\"]/div[2]/div[2]/div/div/form/div/div[3]";
        private const string GO_TO_ITEM = "//*[@id=\"ItemTabNav\"]";
        public const string SEARCHITEMNAME = "//*[@id=\"SearchPatternWithAutocomplete\"]";
        public const string ITEM = "//*[@id=\"purchasing-item-detail-1 \"]/div[2]/table";
        public const string REFERENCE = "//*[@id=\"Reference\"]";
        public const string GROUP = "SelectedGroups_ms";
        public const string SUBGROUP = "SelectedSubGroups_ms";
        public const string FIRST_PO_SITE = "/html/body/div[2]/div/div[2]/div[2]/table/tbody/tr[1]/td[3]";
        public const string FIRStNUMBER = "//*[@id=\"purchasing-purchaseorders-details-1\"]/td[2]";
        public const string PRINCTLIST = "//*[@id=\"header-print-button\"]";
        public const string CLEARCLICK = "/html/body/div[14]/div[2]/div/a[2]";
        private const string SEND_EMAIL_CONFIRMP = "btn-init-async-send-mail";
        private const string FILE_SENT = "FileSent";
        private const string NEXT_SENT = "purchaseorder-sendmail-next-btn";
        private const string MODAL_FORM = "//*[@id=\"modal-1\"]/div/div/div/form";
        private const string DELIVERY_LOCATION_FILTER = "SelectedSitePlaces_ms";
        private const string UNCHECK_ALL_DELIVERY_LOCATION = "/html/body/div[12]/div/ul/li[2]/a/span[2]";
        private const string SEARCH_DELIVERY_LOCATION = "/html/body/div[12]/div/div/label/input";
        private const string DELIVERY_LOCATION = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[5]";

        // _________________________________________ Variables _______________________________________________

        // Général 
        [FindsBy(How = How.Id, Using = APPROVED_BY)]
        private IWebElement _approvedby;
        [FindsBy(How = How.XPath, Using = FIRST_ITEM_FILTRED)]
        private IWebElement _firstitemfiltred;
        [FindsBy(How = How.XPath, Using = CLEARCLICK)]
        private IWebElement _clearclick;
        [FindsBy(How = How.XPath, Using = PRINCTLIST)]
        private IWebElement _printlist;
        [FindsBy(How = How.XPath, Using = FIRStNUMBER)]
        private IWebElement _firstnumber;
        [FindsBy(How = How.XPath, Using = ITEM)]
        private IWebElement _item;
        [FindsBy(How = How.XPath, Using = GO_TO_ITEM)]
        private IWebElement _gototime;

        [FindsBy(How = How.XPath, Using = ITEM_NAME)]
        private IWebElement _itemname;
        [FindsBy(How = How.Id, Using = PRINT_PO)]
        private IWebElement _printpo;

        [FindsBy(How = How.LinkText, Using = NEW_PURCHASE_ORDER)]
        private IWebElement _createNewPurchaseOrder;

        [FindsBy(How = How.LinkText, Using = PRINT_LIST_PO)]
        private IWebElement _printResults;

        [FindsBy(How = How.LinkText, Using = PRINT_WITH_PRICES_CHECKBOX)]
        private IWebElement _printWithPrices;

        [FindsBy(How = How.Id, Using = CONFIRM_PRINT)]
        private IWebElement _confirmPrint;

        [FindsBy(How = How.Id, Using = EXPORT)]
        private IWebElement _export;

        [FindsBy(How = How.Id, Using = EXPORT_WMS)]
        private IWebElement _exportWMS;

        [FindsBy(How = How.Id, Using = CONFIRM_EXPORT_WMS)]
        private IWebElement _confirmExportWMS;

        [FindsBy(How = How.Id, Using = SEND_RESULTS_BY_MAIL)]
        private IWebElement _sendResultsByMail;

        [FindsBy(How = How.Id, Using = CONFIRM_SEND_MAIL)]
        private IWebElement _confirmSendMail;

        // Tableau
        [FindsBy(How = How.XPath, Using = FIRST_PO)]
        private IWebElement _firstPO;

        [FindsBy(How = How.XPath, Using = FIRST_PO_NUMBER)]
        private IWebElement _firstPONumber;     
        [FindsBy(How = How.XPath, Using = FIRST_PO_SITE)]
        private IWebElement _firstPOSite;
        //__________________________________Filters_______________________________________

        [FindsBy(How = How.Id, Using = RESET_FILTER_DEV)]
        private IWebElement _resetFilterDev;

        [FindsBy(How = How.Id, Using = SEARCH_FILTER)]
        private IWebElement _searchByNumber;
        [FindsBy(How = How.Id, Using = SEARCHITEMNAME)]
        private IWebElement _searchByName;

        [FindsBy(How = How.Id, Using = SO_NUMBER_FILTER)]
        private IWebElement _searchBySupplyOrderNumber;

        [FindsBy(How = How.Id, Using = SUPPLIER_FILTER)]
        private IWebElement _supplierFilter;

        [FindsBy(How = How.XPath, Using = UNCHECK_ALL_SUPPLIER)]
        private IWebElement _uncheckAllSupplier;

        [FindsBy(How = How.XPath, Using = SEARCH_SUPPLIER)]
        private IWebElement _searchSupplier;

        [FindsBy(How = How.Id, Using = DATE_FROM)]
        public IWebElement _dateFrom;

        [FindsBy(How = How.Id, Using = DATE_TO)]
        public IWebElement _dateTo;

        [FindsBy(How = How.Id, Using = VALIDATION_DATE_DEV)]
        private IWebElement _byValidationDateDev;

        [FindsBy(How = How.Id, Using = SHOW_FILTER)]
        private IWebElement _filterShow;

        [FindsBy(How = How.Id, Using = SHOW_ALL_DEV)]
        private IWebElement _showAllDev;

        [FindsBy(How = How.Id, Using = SHOW_ACTIVE_DEV)]
        private IWebElement _showActiveDev;

        [FindsBy(How = How.Id, Using = SHOW_INACTIVE_DEV)]
        private IWebElement _showInactiveDev;

        [FindsBy(How = How.Id, Using = RECEIPT_STATUS_FILTER)]
        private IWebElement _receiptStatus;

        [FindsBy(How = How.Id, Using = SORTBY_FILTER)]
        private IWebElement _sortBy;

        [FindsBy(How = How.Id, Using = TO_SEND_BY_MAIL)]
        private IWebElement _toSendByMail;

        [FindsBy(How = How.Id, Using = STATUS_ALL)]
        private IWebElement _statusAllFilter;

        [FindsBy(How = How.Id, Using = STATUS_OPENED)]
        private IWebElement _statusOpenedFilter;

        [FindsBy(How = How.Id, Using = STATUS_CLOSED)]
        private IWebElement _statusClosedFilter;

        [FindsBy(How = How.Id, Using = STATUS_CANCELLED)]
        private IWebElement _statusCancelledFilter;

        [FindsBy(How = How.Id, Using = SITE_FILTER)]
        private IWebElement _siteFilter;

        [FindsBy(How = How.XPath, Using = UNCHECK_ALL_SITE)]
        private IWebElement _uncheckAllSite;

        [FindsBy(How = How.XPath, Using = SEARCH_SITE)]
        private IWebElement _searchSite;
        [FindsBy(How = How.Id, Using = GROUP)]
        private IWebElement _searchGroup;

        [FindsBy(How = How.Id, Using = DELIVERY_LOCATION_FILTER)]
        private IWebElement _deliveryLocationFilter;

        [FindsBy(How = How.XPath, Using = UNCHECK_ALL_DELIVERY_LOCATION)]
        private IWebElement _uncheckAllDeliveryLocation;

        [FindsBy(How = How.XPath, Using = SEARCH_DELIVERY_LOCATION)]
        private IWebElement _searchDeliveryLocation;
        public enum FilterType
        {
            ByNumber,
            BySupplierOrderNumber,
            Supplier,
            DateFrom,
            DateTo,
            ByValidationDate,
            FilterShow,
            tobeapprovedby,
            ShowAll,
            ShowActive,
            ShowInactive,
            ReceiptStatus,
            SortBy,
            ToSendByMail,
            StatusAll,
            Opened,
            Closed,
            Cancelled,
            Site,
            ByName,
            Group,
            DeliveryLocation
        }

        public void Filter(FilterType FilterType, object value)
        {
            Actions action = new Actions(_webDriver);

            switch (FilterType)
            {
                case FilterType.ByNumber:

                    _searchByNumber = WaitForElementIsVisible(By.Id(SEARCH_FILTER));
                    action.MoveToElement(_searchByNumber).Perform();
                    _searchByNumber.SetValue(ControlType.TextBox, value);
                    break;
                case FilterType.BySupplierOrderNumber:
                    _searchBySupplyOrderNumber = WaitForElementIsVisible(By.Id(SO_NUMBER_FILTER));
                    _searchBySupplyOrderNumber.SetValue(ControlType.TextBox, value);
                    break;
                case FilterType.Supplier:
                    _supplierFilter = WaitForElementExists(By.Id(SUPPLIER_FILTER));
                    _supplierFilter.Click();

                    _uncheckAllSupplier = WaitForElementIsVisible(By.XPath(UNCHECK_ALL_SUPPLIER));
                    _uncheckAllSupplier.Click();

                    _searchSupplier = WaitForElementIsVisible(By.XPath(SEARCH_SUPPLIER));
                    _searchSupplier.SetValue(ControlType.TextBox, value);

                    var supplierToCheck = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    supplierToCheck.SetValue(ControlType.CheckBox, true);

                    _supplierFilter.Click();
                    break;
                case FilterType.DateFrom:
                    _dateFrom = WaitForElementIsVisible(By.Id(DATE_FROM));
                    _dateFrom.SetValue(ControlType.DateTime, value);
                    _dateFrom.SendKeys(Keys.Tab);
                    break;
                case FilterType.DateTo:
                    _dateTo = WaitForElementIsVisible(By.Id(DATE_TO));
                    _dateTo.SetValue(ControlType.DateTime, value);
                    _dateTo.SendKeys(Keys.Tab);
                    break;
                case FilterType.ByValidationDate:
                        _byValidationDateDev = WaitForElementExists(By.Id(VALIDATION_DATE_DEV));
                        action.MoveToElement(_byValidationDateDev).Perform();
                        _byValidationDateDev.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.FilterShow:
                    _filterShow = WaitForElementIsVisible(By.Id(SHOW_FILTER));
                    _filterShow.SetValue(ControlType.DropDownList, value);
                    break;

                case FilterType.tobeapprovedby:
                    _approvedby = WaitForElementIsVisible(By.Id(APPROVED_BY));
                    _approvedby.SetValue(ControlType.DropDownList, value);
                    break;

                case FilterType.ShowAll:
                        _showAllDev = WaitForElementExists(By.Id(SHOW_ALL_DEV));
                        action.MoveToElement(_showAllDev).Perform();
                        _showAllDev.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.ShowActive:
                        _showActiveDev = WaitForElementExists(By.Id(SHOW_ACTIVE_DEV));
                        action.MoveToElement(_showActiveDev).Perform();
                        _showActiveDev.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.ShowInactive:
                        _showInactiveDev = WaitForElementExists(By.Id(SHOW_INACTIVE_DEV));
                        action.MoveToElement(_showInactiveDev).Perform();
                        _showInactiveDev.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.ReceiptStatus:
                    _receiptStatus = WaitForElementIsVisible(By.Id(RECEIPT_STATUS_FILTER));
                    action.MoveToElement(_receiptStatus).Perform();
                    _receiptStatus.Click();
                    var element = WaitForElementIsVisible(By.XPath("//*[@id=\"cbReceiptStatus\"]/option[@value='" + value + "']"));
                    _receiptStatus.SetValue(ControlType.DropDownList, element.Text);
                    _receiptStatus.Click();
                    break;
                case FilterType.SortBy:
                    _sortBy = WaitForElementIsVisible(By.Id(SORTBY_FILTER));
                    action.MoveToElement(_sortBy).Perform();
                    _sortBy.SetValue(ControlType.DropDownList, value);
                    break;
                case FilterType.ToSendByMail:
                    _toSendByMail = WaitForElementExists(By.Id(TO_SEND_BY_MAIL));
                    action.MoveToElement(_toSendByMail).Perform();
                    _toSendByMail.SetValue(ControlType.CheckBox, value);
                    break;
                case FilterType.StatusAll:
                    _statusAllFilter = WaitForElementExists(By.Id(STATUS_ALL));
                    action.MoveToElement(_statusAllFilter).Perform();
                    _statusAllFilter.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.Opened:
                    _statusOpenedFilter = WaitForElementExists(By.Id(STATUS_OPENED));
                    action.MoveToElement(_statusOpenedFilter).Perform();
                    _statusOpenedFilter.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.Closed:
                    _statusClosedFilter = WaitForElementExists(By.Id(STATUS_CLOSED));
                    action.MoveToElement(_statusClosedFilter).Perform();
                    _statusClosedFilter.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.Cancelled:
                    _statusCancelledFilter = WaitForElementExists(By.Id(STATUS_CANCELLED));
                    action.MoveToElement(_statusCancelledFilter).Perform();
                    _statusCancelledFilter.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.Site:
                    _siteFilter = WaitForElementExists(By.Id(SITE_FILTER));
                    action.MoveToElement(_siteFilter).Perform();
                    _siteFilter.Click();

                    _uncheckAllSite = WaitForElementIsVisible(By.XPath(UNCHECK_ALL_SITE));
                    _uncheckAllSite.Click();

                    _searchSite = WaitForElementIsVisible(By.XPath(SEARCH_SITE));
                    _searchSite.SetValue(ControlType.TextBox, value);

                    var valueToCheck = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    valueToCheck.SetValue(ControlType.CheckBox, true);

                    _siteFilter.Click();
                    break;
                case FilterType.Group:
                    _searchGroup = WaitForElementExists(By.Id(GROUP));
                    action.MoveToElement(_searchGroup).Perform();
                    _searchGroup.Click();

                    _uncheckAllSite = WaitForElementIsVisible(By.XPath(UNCHECK_ALL_SITE));
                    _uncheckAllSite.Click();

                    _searchSite = WaitForElementIsVisible(By.XPath(SEARCH_SITE));
                    _searchSite.SetValue(ControlType.TextBox, value);

                    //var valueToCheck = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    //valueToCheck.SetValue(ControlType.CheckBox, true);

                    _siteFilter.Click();
                    break;

                case FilterType.DeliveryLocation:
                    _deliveryLocationFilter = WaitForElementExists(By.Id(DELIVERY_LOCATION_FILTER));
                    action.MoveToElement(_deliveryLocationFilter).Perform();
                    _deliveryLocationFilter.Click();

                    _uncheckAllDeliveryLocation = WaitForElementIsVisible(By.XPath(UNCHECK_ALL_DELIVERY_LOCATION));
                    _uncheckAllDeliveryLocation.Click();

                    _searchDeliveryLocation = WaitForElementIsVisible(By.XPath(SEARCH_DELIVERY_LOCATION));
                    _searchDeliveryLocation.SetValue(ControlType.TextBox, value);

                    var valueToCheckDelivery = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    valueToCheckDelivery.SetValue(ControlType.CheckBox, true);

                    _deliveryLocationFilter.Click();
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), FilterType, null);
            }
            WaitPageLoading();
            Thread.Sleep(2000);
        }
        public void FilterItem(FilterType FilterType, object value)
        {
            Actions action = new Actions(_webDriver);

            switch (FilterType)
            {
                case FilterType.ByName:
                    _searchByName = WaitForElementIsVisible(By.XPath(SEARCHITEMNAME));
                    action.MoveToElement(_searchByName).Perform();
                    _searchByName.SetValue(ControlType.TextBox, value);
                    break;
  

                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), FilterType, null);
            }
            WaitPageLoading();
            Thread.Sleep(2000);
        }

        public void ClearDateFrom()
        {
            _dateFrom = WaitForElementIsVisible(By.Id(DATE_FROM));
            _dateFrom.ClearElement();
        }

        public void ResetFilters()
        {
            _resetFilterDev = WaitForElementIsVisible(By.Id(RESET_FILTER_DEV));
            _resetFilterDev.Click();
            WaitForLoad();

            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                Filter(FilterType.DateTo, DateUtils.Now);
            }
        }

        public bool IsSortedByNumber()
        {
            bool valueBool = true;
            var ancientNumber = int.MaxValue;
            int tot;

            if (CheckTotalNumber() > 100)
            {
                tot = 100;
            }
            else
            {
                tot = CheckTotalNumber();
            }
            if (tot == 0)
                return false;


            var elements = _webDriver.FindElements(By.XPath(NUMBER));

            foreach (var element in elements)
            {

                try
                {
                    int number = int.Parse(element.Text);

                    if (number > ancientNumber)
                    { valueBool = false; }

                    ancientNumber = number;
                }
                catch 
                {
                    valueBool = false;
                }
               
            }

            return valueBool;
        }
        public bool IsSortedByDate()
        {
            // Take the date format from the datepicker element and use it to format the date column to avoid date errors
            var dateFormat = _dateFrom.GetAttribute("data-date-format");
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            bool valueBool = true;
            DateTime ancientDate = DateTime.MinValue;
            int tot;

            if (CheckTotalNumber() > 100)
            {
                tot = 100;
            }
            else
            {
                tot = CheckTotalNumber();
            }
            if (tot == 0)
                return false;

            IReadOnlyCollection<IWebElement> elements;
           
                elements = _webDriver.FindElements(By.XPath("//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[7]"));
            

            foreach (var (element, index) in elements.Select((value, i) => (value, i)))
            {

                try
                {
                    string dateText = element.Text;
                    DateTime date = DateTime.Parse(dateText, ci);

                    if (index == 0)
                    {
                        ancientDate = date;
                    }

                    if (DateTime.Compare(ancientDate.Date, date) < 0)
                    { valueBool = false; }

                    ancientDate = date;
                }
                catch 
                {
                    valueBool = false;
                }
           
            }
            return valueBool;
        }

        public bool CheckValidation(bool validated)
        {
            bool isValidated = true;
            int tot = CheckTotalNumber() > 100 ? 100 : CheckTotalNumber();

            if (tot == 0)
                return false;

            for (int i = 0; i < tot; i++)
            {
                try
                {
                        _webDriver.FindElement(By.XPath(string.Format("//*[@id='list-item-with-action']/table/tbody/tr/td[1]/div/i[contains(@class,'circle-check')]", i + 1)));
                    

                    if (!validated)
                        return true;
                }
                catch
                {
                    isValidated = false;

                    if (validated)
                        return false;
                }
            }

            return isValidated;
        }

        public void ClearAllPrintCache()
        {
            _printlist = WaitForElementIsVisible(By.XPath(PRINCTLIST));
            _printlist.Click();
            WaitForLoad();
            _clearclick = WaitForElementIsVisible(By.XPath(CLEARCLICK));
            _clearclick.Click();
        }
        public void ClosePrintList()
        {
            _printlist = WaitForElementIsVisible(By.XPath(PRINCTLIST));
            _printlist.Click();
            WaitForLoad();
        
        }
        public Boolean IsSentByMail()
        {
            bool valueBool = true;
            int tot;

            if (CheckTotalNumber() > 100)
            {
                tot = 100;
            }
            else
            {
                tot = CheckTotalNumber();
            }
            if (tot == 0)
                return false;

   
            for (int i = 0; i < tot; i++)
            {
                try
                {
                    
                        _webDriver.FindElement(By.XPath(string.Format("//*[@id='list-item-with-action']/table/tbody/tr[{0}]/td[6]/div[2]/div/i[contains(@class,'envelope')]", i + 1)));
                    

                }
                catch
                {
                    valueBool = false;
                }
            }
            return valueBool;
        }

        public Boolean VerifySite(string value)
        {
            bool valueBool = true;
            int tot;

            if (CheckTotalNumber() > 100)
            {
                tot = 100;
            }
            else
            {
                tot = CheckTotalNumber();
            }
            if (tot == 0)
                return false;

            var elements = _webDriver.FindElements(By.XPath(VERIFY_SITE));

            foreach (var element in elements)
            {
                if (element.Text != value)
                {
                    valueBool = false;
                }
            }

            return valueBool;
        }

        public bool VerifySupplier(string value)
        {
            bool valueBool = true;
            int tot;

            if (CheckTotalNumber() > 100)
            {
                tot = 100;
            }
            else
            {
                tot = CheckTotalNumber();
            }
            if (tot == 0)
                return false;

            var elements = _webDriver.FindElements(By.XPath(VERIFY_SUPPLIER));

            foreach (var element in elements)
            {
                if (element.Text != value)
                {
                    valueBool = false;
                }
            }

            return valueBool;
        }

        public bool IsDelivered()
        {
            bool valueBool = false;
            int tot;

            if (CheckTotalNumber() > 100)
            {
                tot = 100;
            }
            else
            {
                tot = CheckTotalNumber();
            }
            if (tot == 0)
                return false;


            for (int i = 0; i < tot; i++)
            {
                try
                {
                    IWebElement element;

                    element = _webDriver.FindElement(By.XPath(string.Format("//*[@id=\"list-item-with-action\"]/table/tbody/tr[{0}]/td[6]/div[2]/img", i + 1)));

                    if (element.GetAttribute("title") == "Delivered")
                    {
                        valueBool = true;
                    }
                }
                catch
                {
                    valueBool = false;
                }
            }
            return valueBool;
        }

        public bool IsPartiallyDelivered()
        {
            bool valueBool = false;
            int tot;

            if (CheckTotalNumber() > 100)
            {
                tot = 100;
            }
            else
            {
                tot = CheckTotalNumber();
            }
            if (tot == 0)
                return false;

            for (int i = 0; i < tot; i++)
            {
                try
                {
                    IWebElement element;
                    element = _webDriver.FindElement(By.XPath(string.Format("//*[@id=\"list-item-with-action\"]/table/tbody/tr[{0}]/td[6]/div[2]/img[@title = 'Partially Delivered']", i + 1)));
                    
                    if (element.GetAttribute("title") == "Partially Delivered")
                    {
                        valueBool = true;
                    }
                }
                catch
                {
                    valueBool = false;
                }
            }
            return valueBool;
        }

        public bool CheckStatus(bool active)
        {
            bool isActive = false;
            int tot = CheckTotalNumber() > 100 ? 100 : CheckTotalNumber();

            if (tot == 0)
                return false;

            for (int i = 0; i < tot; i++)
            {
                try
                {
                        _webDriver.FindElement(By.XPath(String.Format("//*[@id='list-item-with-action']/table/tbody/tr[1]/td[1]/div/i[contains(@class,'circle-xmark')]", i + 1)));
                    

                    if (active)
                        return false;
                }
                catch
                {
                    isActive = true;
                    if (!active)
                        return true;
                }
            }
            return isActive;
        }

        public bool IsSentByEDI()
        {
            bool valueBool = false;
            int tot;

            if (CheckTotalNumber() > 100)
            {
                tot = 100;
            }
            else
            {
                tot = CheckTotalNumber();
            }
            if (tot == 0)
                return false;

            for (int i = 0; i < tot; i++)
            {
                try
                {
                    var element = _webDriver.FindElement(By.XPath(string.Format(SENT_EDI, i + 1))).GetAttribute("title").ToString();
                    if (element == "Sent")
                        return true;
                }
                catch
                {
                    valueBool = false;
                }
            }
            return valueBool;
        }

        public bool IsDateRespected(DateTime fromDate, DateTime toDate)
        {
            // Take the date format from the datepicker element and use it to format the date column to avoid date errors
            var dateFormat = _dateFrom.GetAttribute("data-date-format");
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            var valueBool = true;
            int tot;

            if (CheckTotalNumber() > 100)
            {
                tot = 100;
            }
            else
            {
                tot = CheckTotalNumber();
            }
            if (tot == 0)
                return false;
            IReadOnlyCollection<IWebElement> elements;
           
                elements = _webDriver.FindElements(By.XPath("//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[7]"));
            
            foreach (var element in elements)
            {
                try
                {
                    string dateText = element.Text;
                    DateTime date = DateTime.Parse(dateText, ci).Date;

                    if (DateTime.Compare(date, fromDate) < 0 || DateTime.Compare(date, toDate) > 0)
                    { valueBool = false; }
                }
                catch
                {
                    valueBool = false;
                }
            }

            return valueBool;
        }   
        public string DeleteFirstPurchaseOrder()
        {
            var number = GetFirstPurchaseOrderNumber();
            IWebElement firstPurchaseOrder;
           
                firstPurchaseOrder = WaitForElementIsVisible(By.XPath("//*[@id=\"list-item-with-action\"]/table/tbody/tr[1]/td[10]/a/span"));
           
            firstPurchaseOrder.Click();
            var confirmBtn = WaitForElementIsVisible(By.Id(CONFIRM_DELETE));
            confirmBtn.Click();
            WaitForLoad();
            return number;
        }

        // ___________________________________ Méthodes _____________________________________________

        // Général
        public CreatePurchaseOrderModalPage CreateNewPurchaseOrder()
        {
            ShowPlusMenu();

            _createNewPurchaseOrder = WaitForElementIsVisible(By.LinkText(NEW_PURCHASE_ORDER));
            _createNewPurchaseOrder.Click();
            WaitForLoad();

            return new CreatePurchaseOrderModalPage(_webDriver, _testContext);
        }

        public PrintReportPage PrintResults(bool versionPrint)
        {
            ShowExtendedMenu();

            _printResults = WaitForElementIsVisible(By.LinkText(PRINT_LIST_PO));
            _printResults.Click();
            WaitForLoad();

            _printWithPrices = WaitForElementIsVisible(By.XPath(PRINT_WITH_PRICES_CHECKBOX));
            _printWithPrices.SetValue(ControlType.CheckBox, true);

            _confirmPrint = WaitForElementIsVisible(By.Id(CONFIRM_PRINT));
            _confirmPrint.Click();
            WaitForLoad();

            if (versionPrint)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-pdf']"));
                ClickPrintButton();
            }

            //Results are opened in a new tab, switch the driver to the newly created one
            WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new PrintReportPage(_webDriver, _testContext);
        }

        public PrintReportPage PrintResultsNew(bool versionPrint)
        {
            ShowExtendedMenu();

            _printpo = WaitForElementIsVisible(By.Id(PRINT_PO));
            _printpo.Click();
            WaitForLoad();

            _printWithPrices = WaitForElementIsVisible(By.XPath(PRINT_WITH_PRICES_CHECKBOX));
            _printWithPrices.SetValue(ControlType.CheckBox, true);

            _confirmPrint = WaitForElementIsVisible(By.Id(CONFIRM_PRINT));
            _confirmPrint.Click();
            WaitForLoad();

            if (versionPrint)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-pdf']"));
                ClickPrintButton();
            }

            //Results are opened in a new tab, switch the driver to the newly created one
            WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new PrintReportPage(_webDriver, _testContext);
        }

        public void Export(bool versionPrint)
        {
            ShowExtendedMenu();

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
            //Purchase orders2020-01-28 15-02-06.xlsx
            Regex r = new Regex("^Purchase orders\\d\\d\\d\\d-\\d\\d-\\d\\d\\s\\d\\d-\\d\\d-\\d\\d.xlsx", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = r.Match(filePath);

            return m.Success;
        }

        public void ExportWMS(bool versionPrint)
        {
            if (versionPrint)
            {
                ClearDownloads();
            }

            ShowExtendedMenu();

            _exportWMS = WaitForElementIsVisible(By.Id(EXPORT_WMS));
            _exportWMS.Click();
            WaitForLoad();

            _confirmExportWMS = WaitForElementIsVisible(By.Id(EXPORT_WMS));
            _confirmExportWMS.Click();
            WaitForLoad();

            if (versionPrint)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-text']"));
                ClickPrintButton();
            }

            WaitForDownload();
            Close();
        }

        public FileInfo GetExportWMSFile(FileInfo[] taskFiles)
        {
            List<FileInfo> correctDownloadFiles = new List<FileInfo>();

            foreach (var file in taskFiles)
            {
                //  Test REGEX
                if (IsTextFileCorrect(file.Name))
                {
                    correctDownloadFiles.Add(file);
                }
            }
            Assert.IsTrue(correctDownloadFiles.Count > 0, "Aucun fichier téléchargé.");

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

        public bool IsTextFileCorrect(string filePath)
        {
            string mois = "(?:0[1-9]|1[0-2])";  // mois
            string annee = "\\d{4}";            // annee YYYY
            string jour = "[0-3]\\d";           // jour
            string nombre = "\\d{6}";           // nombre


            Regex r = new Regex("^purchaseorders" + "_" + annee + mois + jour + nombre + ".txt$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = r.Match(filePath);

            return m.Success;
        }

        public void SendResultsByMail()
        {
            ShowExtendedMenu();

            _sendResultsByMail = WaitForElementIsVisible(By.Id(SEND_RESULTS_BY_MAIL));
            _sendResultsByMail.Click();
            WaitForLoad();

            var nextSendMail = WaitForElementIsVisible(By.Id(NEXT_SENT));
            nextSendMail.Click();
            WaitForLoad();

            _confirmSendMail = WaitForElementIsVisible(By.Id(CONFIRM_SEND_MAIL));
            _confirmSendMail.Click();

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
        }

        // Tableau
        public PurchaseOrderItem SelectFirstItem()
        {
            _firstPO = WaitForElementIsVisible(By.XPath(FIRST_PO));
            _firstPO.Click();
            WaitForLoad();

            return new PurchaseOrderItem(_webDriver, _testContext);
        }

        public string GetFirstPurchaseOrderNumber()
        {
            if (isElementExists(By.XPath(FIRST_PO_NUMBER)))
            {
                _firstPONumber = WaitForElementIsVisible(By.XPath(FIRST_PO_NUMBER));
                return _firstPONumber.Text;
            }
            else
            {
                return "";
            }
        }
        public DateTime GetFirstDeliveryDate()
        {
            var deliveryDateElement = WaitForElementIsVisible(By.XPath("//*[@id=\"purchasing-purchaseorders-details-1\"]/td[7]"));
            return DateTime.ParseExact(deliveryDateElement.Text.ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
        }



        public string GetFirstPurchaseOrderSite()
        {
            if (isElementExists(By.XPath(FIRST_PO_SITE)))
            {
                _firstPONumber = WaitForElementIsVisible(By.XPath(FIRST_PO_SITE));
                return _firstPOSite.Text;
            }
            else
            {
                return "";
            }
        }
        public List<string> GetTablePONumbers()
        {
            List<string> purchaseOrdersNumberList = new List<string>();
            var elemPONumbers = _webDriver.FindElements(By.XPath(ALL_PO_NUMBERS));
            foreach (var elem in elemPONumbers)
            {
                purchaseOrdersNumberList.Add(elem.Text);
            }

            return purchaseOrdersNumberList;
        }
        public bool IsDeleted(string number)
        {
            var listnumbers = _webDriver.FindElements(By.XPath(NUMBERS_LIST));
            foreach(var n in listnumbers)
            {
                if (n.Text == number)
                {
                    return false;
                }
            }
            return true;
        }
        
        public void ClickActionsButton()
        {
            var actionsbtn = WaitForElementIsVisible(By.XPath(ACTIONS_BUTTONS));
            actionsbtn.Click();
            WaitForLoad();
        }
        public void ClickTrash()
        {
            var deleteBtn = WaitForElementIsVisible(By.Id(DELETE));
            deleteBtn.Click();
            WaitPageLoading();
            WaitForLoad();
        }
        public ItemGeneralInformationPage ClickPen()
        {
            var penBtn = WaitForElementIsVisible(By.Id(PEN_BTN));
            penBtn.Click();

            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);
            WaitForLoad();
            return new ItemGeneralInformationPage(_webDriver, _testContext);
        }
        public PurchaseOrderItem ClickFirstPurchaseOrder()
        {
            var firstline = WaitForElementIsVisible(By.XPath(FIRST_LINE));
            firstline.Click();
            return new PurchaseOrderItem(_webDriver, _testContext);
        }
        public void ClickFirstItem()
        {
            var firstItem = WaitForElementIsVisible(By.XPath(FIRST_ITEM));
            firstItem.Click();
            WaitForLoad();
        }

        public void ClickFirstItemFiltred()
        {
            var firstItem = WaitForElementIsVisible(By.XPath(FIRST_ITEM_FILTRED));
            firstItem.Click();
            WaitForLoad();
        }

        public void GoToItem()
        {
            var firstItem = WaitForElementIsVisible(By.XPath(GO_TO_ITEM));
            firstItem.Click();
            WaitForLoad();
        }

        public void ClickOnItem()
        {
            var firstItem = WaitForElementIsVisible(By.XPath(ITEM));
            firstItem.Click();
            WaitForLoad();
        }


        public void SetQuantity(string qty)
        {
            ClickFirstItem();
            var qtyInput = WaitForElementIsVisible(By.Id("item_PodRowDto_Quantity"));
            qtyInput.SetValue(ControlType.TextBox, qty);
            // data binding
            Thread.Sleep(2000);
            WaitForLoad();
        }
        public string GetQuantity()
        {
            ClickFirstItem();
            var qty = WaitForElementIsVisible(By.Id("item_PodRowDto_Quantity"));
            return qty.GetAttribute("value");
        }
        public string GetTotalVAT()
        {
            var total = WaitForElementIsVisible(By.XPath(TOTAL_VAT));
            return total.Text;
        }

        public string GetNameOfTheItem()
        {
            var itemname = WaitForElementIsVisible(By.XPath(ITEM_NAME));
            return itemname.Text;
        }
        public string Getreference()
        {
            var itemname = WaitForElementIsVisible(By.XPath(REFERENCE));
             return itemname.GetAttribute("value");

        }
        public string GetFirstItemNumber()
        {
            ClickFirstItem();
            var firstNumber = WaitForElementIsVisible(By.XPath(FIRST_ITEM_NUMBER));
            return firstNumber.Text;
        }
        public void DeleteItem()
        {
            ClickFirstItem();
            ClickActionsButton();
            ClickTrash();
        }
        public bool VerifyValues(string vatold , string vatnew , string quantityold, string quantitynew)
        {
            if(vatold == vatnew && quantitynew == quantityold)
            {
                return true;
            }
            return false;
        }

        public IWebElement GetExtendedMenuButton()
        {
            var btn = WaitForElementIsVisible(By.XPath(EXTENDED_MENU_BTN));
            return btn;
        }
        public void ConfirmEmailSend()
        {
            if (IsDev())
            {
                var confirmBtn = WaitForElementIsVisible(By.Id(SEND_EMAIL_CONFIRM));
                confirmBtn.Click();
                WaitForLoad();

            }
            else
            {
                var confirmBtn = WaitForElementIsVisible(By.Id(SEND_EMAIL_CONFIRMP));
                confirmBtn.Click();
                WaitForLoad();

            }            
        }
        public void SendByEmail(string email)
        {
            var sendBtn = WaitForElementIsVisible(By.Id(SEND_BY_EMAIL));
            sendBtn.Click();
            var to = WaitForElementIsVisible(By.Id(TO_EMAIL));
            WaitForLoad();
            to.Clear();
            WaitForLoad();
            to.SendKeys(email);
            ConfirmEmailSend();
            var x = WaitForElementIsVisible(By.XPath(MODAL_FORM));
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.InvisibilityOfElementLocated(By.XPath(MODAL_FORM)));
            _webDriver.Navigate().Back();
        }
       
        public string ClickFirstPurchaseOrderGetNumber()
        {
            var number = GetFirstPurchaseOrderNumber();
            var firstline = WaitForElementIsVisible(By.XPath(FIRST_LINE));
            firstline.Click();
            return number;
        }
        public bool IsPurchaseOrderSentByMail()
        {
            WaitForLoad();

                try
                {
                    _webDriver.FindElement(By.XPath("//*[@id='list-item-with-action']/table/tbody/tr[1]/td[6]/div[2]/div/i[contains(@class,'envelope')]"));
                    return true;
                }
                catch
                {
                    return false;
                }
           
        }

        public void GoToFooterSubMenu()
        {
            var footerBtn = WaitForElementIsVisible(By.Id(FOOTER));
            footerBtn.Click();
        }

        public string GetTotalGrossAmountFromFooter()
        {
            var totalGrossAmount = WaitForElementIsVisible(By.XPath(TOTAL_GROSS_AMOUNT));
            return totalGrossAmount.Text;
        }
        public bool isPageSizeEqualsTo16()
        {
            if (isElementVisible(By.XPath("//option[text()='16']")))
            {
                var nbPages = WaitForElementExists(By.XPath("//option[text()='16']"));
                return nbPages.Selected;
            }
            return false;
        }
        public bool isPageSizeEqualsTo30()
        {
            if (isElementVisible(By.XPath("//option[text()='30']")))
            {
                var nbPages = WaitForElementExists(By.XPath("//option[text()='30']"));
                return nbPages.Selected;
            }
            return false;
        }
        public bool isPageSizeEqualsTo50()
        {
            if (isElementVisible(By.XPath("//option[text()='50']")))
            {
                var nbPages = WaitForElementExists(By.XPath("//option[text()='50']"));
                return nbPages.Selected;
            }
            return false;
        }
        public int GetTotalRowsForPagination()
        {
            var rows = _webDriver.FindElements(By.XPath(ROWS_FOR_PAGINATION));
            return rows.Count;
        }

        public void SplitNameAndReference(string input, out string reference, out string itemName)
        {
            // Initialisation des variables de sortie
            reference = string.Empty;
            itemName = string.Empty;

            // Expression régulière pour capturer le texte entre parenthèses et le texte principal
            var match = Regex.Match(input, @"^(.*?)\s*(?:\((.*?)\))?$");

            if (match.Success)
            {
                // Le nom de l'article est la première capture
                itemName = match.Groups[1].Value.Trim();
                // La référence est la deuxième capture (entre parenthèses)
                reference = match.Groups[2].Value.Trim();
            }
        }
        public IEnumerable<string> GetSupplyOrdersList()
        {
            var ordersList = _webDriver.FindElements(By.XPath("//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[2]"));
            return ordersList.Select(p => p.Text);

        }

        public bool CheckIfConstantExistsInList(string supplyOrderNumber)
        {
            var ordersList = GetSupplyOrdersList();


            return ordersList.Contains(supplyOrderNumber);
        }
        public string GetPurchaseOrderNumber()
        {
            _firstPONumber = WaitForElementIsVisible(By.XPath(FIRST_PO_NUMBER));
            return _firstPONumber.GetAttribute("value");
        }

        public bool CheckCommentWorking()
        {
            Actions a = new Actions(_webDriver);
            var comment = _webDriver.FindElements(By.XPath("//*[@class=\"fas fa-message green-text\"]"));
            a.MoveToElement(comment[0]).Perform();
            var title = comment[0].GetAttribute("title");
            return (title != null && title.Length > 0);
        }
        public int GetTotalPurshaseOrder()
        {
            var numbers = WaitForElementIsVisible(By.XPath("/html/body/div[2]/div/div[2]/div[1]/h1/span"));
            return int.Parse(numbers.Text);
        }
        public bool VerifyDeliveryLoction(string location)
        {
            var listeDeliveryLoction = _webDriver.FindElements(By.XPath(DELIVERY_LOCATION));

            if (listeDeliveryLoction.Count == 0)
                return false;

            foreach (var elm in listeDeliveryLoction)
            {
                if (!elm.Text.Contains(location))
                    return false;
            }

            return true;
        }

    }
}

