using iText.Commons.Utils;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Accounting.SupplierInvoices;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
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
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.Claims

{
    public class ClaimsPage : PageBase
    {

        public ClaimsPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //__________________________________ Constantes _____________________________________

        // General
        private const string NEW_CLAIM_BUTTON = "//*[@id=\"div-body\"]/div/div[2]/div[1]/div/div[2]/div/a[text()='New claim']";

        private const string PRINT_CLAIM_RESULTS = "btn-print-claim-notes-report";
        private const string MODAL_PRINT_BUTTON = "printButton";
        private const string EXPORT = "btn-export-excel";
        private const string SEND_RESULTS_BY_MAIL = "btn-send-all-by-email";
        private const string SEND_RESULTS_BY_MAIL_POPUP = "/html/body/div[4]/div/div";
        private const string CONFIRM_SEND = "btn-init-async-send-mail";

        // Tableau
        private const string FIRST_CLAIM = "//*[@id=\"list-item-with-action\"]/table/tbody/tr/td[3]";
        private const string DELETE_CLAIM_BUTTON_DEV = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[1]/td[contains(text(), '{0}')]/..//span[contains(@class, 'fas fa-trash-alt')]";
        private const string DELETE_CLAIM_BUTTON_PATCH = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[1]/td[contains(text(), '{0}')]/..//span[contains(@class, 'glyphicon glyphicon-trash')]";
        private const string DELETE_CLAIM_CONFIRM_BUTTON = "dataConfirmOK";
        private const string FIRST_DATE = "//*[@id=\"list-item-with-action\"]/table/tbody/tr/td[9]";


        // Filtres
        private const string RESET_FILTER_DEV = "ResetFilter";
        private const string RESET_FILTER_PATCH = "//*[@id=\"formSearchReceiptNotes\"]/div[1]/a";

        private const string FILTER_NUMBER = "SearchNumber";
        private const string FILTER_SUPPLIER = "SelectedSuppliers_ms";
        private const string UNSELECT_ALL_SUPPLIER = "/html/body/div[10]/div/ul/li[2]/a";
        private const string SEARCH_SUPPLIER = "/html/body/div[10]/div/div/label/input";
        private const string FILTER_DATE_FROM = "date-picker-start";
        private const string FILTER_DATE_TO = "date-picker-end";

        private const string FILTER_SHOW_ALL_CLAIMS_DEV = "ShowAll";
        private const string FILTER_SHOW_ALL_CLAIMS_PATCH = "//*[@id=\"ShowOnlyValue\"][@value='All']";

        private const string FILTER_NOT_VALIDATED_DEV = "NotValidatedOnly";
        private const string FILTER_NOT_VALIDATED_PATCH = "//*[@id=\"ShowOnlyValue\"][@value='NotValidatedOnly']";

        private const string FILTER_VALIDATED_DEV = "ValidatedOnly";
        private const string FILTER_VALIDATED_PATCH = "//*[@id=\"ShowOnlyValue\"][@value='ValidatedOnly']";

        private const string FILTER_PARTIALLY_INVOICED_DEV = "PartiallyInvoiced";
        private const string FILTER_PARTIALLY_INVOICED_PATCH = "//*[@id=\"ShowOnlyValue\"][@value='PartiallyInvoiced']";

        private const string FILTER_SHOW_ALL_VISU_DEV = "ShowActiveAndInactive";
        private const string FILTER_SHOW_ALL_VISU_PATCH = "//*[@id=\"ShowActive\"][@value='All']";

        private const string FILTER_ACTIVE_DEV = "ActiveOnly";
        private const string FILTER_ACTIVE_PATCH = "//*[@id=\"ShowActive\"][@value='ActiveOnly']";

        private const string FILTER_INACTIVE_DEV = "InactiveOnly";
        private const string FILTER_INACTIVE_PATCH = "//*[@id=\"ShowActive\"][@value='InactiveOnly']";

        private const string FILTER_SHOW_ALL_STATUS_DEV = "ShowAllStatus";
        private const string FILTER_SHOW_ALL_STATUS_PATCH = "//*[@id=\"ShowAllStatus\"][@value='All']";

        private const string FILTER_OPENED_DEV = "Opened";
        private const string FILTER_OPENED_PATCH = "//*[@id=\"Opened\"][@value='Opened']";

        private const string FILTER_CLOSED_DEV = "Closed";
        private const string FILTER_CLOSED_PATCH = "//*[@id=\"Closed\"][@value='Closed']";

        private const string FILTER_SORTBY = "cbSortBy";
        private const string FILTER_SITE = "SelectedSites_ms";
        private const string UNSELECT_ALL = "/html/body/div[11]/div/ul/li[2]/a";
        private const string SEARCH_SITE = "/html/body/div[11]/div/div/label/input";

        private const string FILTER_SITE_PLACES = "SelectedSitePlaces_ms";
        private const string SEARCH_SITE_PLACES = "/html/body/div[12]/div/div/label/input";
        private const string UNSELECT_ALL_SITE_PLACES = "/html/body/div[12]/div/ul/li[2]/a/span[2]";

        private const string INACTIVE = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[{0}]/td[1]/img[@alt=\"Inactive\"]";
        private const string VALIDATION = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[{0}]/td[1]/img[@alt=\"Valid\"]";
        private const string INVOICE_STATUS = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[2]/span";
        private const string CLAIM_NUMBER = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[3]";
        private const string CLAIM_SITE = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[4]";
        private const string CLAIM_SUPPLIER = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[5]";
        private const string SEND_BY_MAIL = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[{0}]/td[6]/span";
        private const string DELIVERY_DATE = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[8]";
        private const string DELIVERY_DATE2 = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[9]";
        private const string DELIVERY_ORDER_NUMBER = "ClaimNote_DeliveryOrderNumber";
        private const string DATE = "datapicker-claimNote-deliveryDate";
        private const string STATUS = "ClaimNote_Status";
        private const string ACTIVATE = "ClaimNote_IsActive";
        private const string NUMBER = "claimnote-number";
        private const string COMMENT = "ClaimNote_Comment";
        private const string TOTAL_CLAIM_ITEMS = "//*[@id=\"total-price-span\"]";
        private const string TOTAL_SANCTIONS_ITEMS = "//*[@id=\"total-sanction-span\"]";
        private const string TOTAL_GROSS_AMOUNT_FOOTER = "//*[@id=\"tabContentDetails\"]/div/table[1]/tbody/tr[5]/td[2]";
        private const string TOTAL_SANCTIONS_FOOTER = "//*[@id=\"tabContentDetails\"]/div/table[1]/tbody/tr[4]/td[2]";
        private const string VAT_SANCTIONS_FOOTER = "//*[@id=\"tabContentDetails\"]/div/table[1]/tbody/tr[4]/td[4]";

        private const string FIRST_CLAIM_NOTE = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[1]/td[1]";
        private const string GENERAL_INFORMATION_SUB_MENU = "//*[@id=\"hrefTabContentInformations\"]";
        private const string FOOTER_SUBMENU = "//*[@id=\"hrefTabContentFooter\"]";
        private const string FIRST_ITEM = "//*[@id=\"itemForm_0\"]/div[2]";
        private const string SANCTION_AMOUNT = "item_CndRowDto_SanctionAmount";

        private const string ENVELOPPE_ICON = "//td[contains(text(), '{0}')]/..//span[contains(@class,'fas fa-envelope')]";
        private const string LOCAL_CURRENCY = "//*[@id=\"tabContentDetails\"]/div/table[2]/tbody/tr[2]/td[2]";
        private const string SUPPLIER_CURRENCY = "//*[@id=\"tabContentDetails\"]/div/table[2]/tbody/tr[3]/td[2]";
        private const string FIRST_CLAIM_NUMBER = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[1]/td[3]";
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";
        private const string CANCEL = "btn-cancel-popup";
        private const string LIST_SITE_FILTER = "/html/body/div[11]/ul/li[*]/label/input";
        private const string LIST_SUPPLIER_FILTER = "/html/body/div[10]/ul/li[*]/label/input";
        private const string LIST_SITE_PLACES_FILTER = "/html/body/div[12]/ul/li[*]/label/input";
        private const string DATE_CLAIM = "/html/body/div[2]/div/div[2]/div[2]/table/tbody/tr/td[9]";
   

      
        //__________________________________ Variables _____________________________________

        // General
        [FindsBy(How = How.XPath, Using = NEW_CLAIM_BUTTON)]
        private IWebElement _createNewClaim;

        [FindsBy(How = How.Id, Using = PRINT_CLAIM_RESULTS)]
        private IWebElement _printClaimResults;

        [FindsBy(How = How.Id, Using = MODAL_PRINT_BUTTON)]
        private IWebElement _modalPrintButton;

        [FindsBy(How = How.Id, Using = EXPORT)]
        private IWebElement _export;

        [FindsBy(How = How.Id, Using = SEND_RESULTS_BY_MAIL)]
        private IWebElement _sendResultsByMail;

        [FindsBy(How = How.Id, Using = CONFIRM_SEND)]
        private IWebElement _confirmSendMail;

        // Tableau
        [FindsBy(How = How.XPath, Using = FIRST_CLAIM)]
        private IWebElement _firstClaim;

        [FindsBy(How = How.XPath, Using = DELETE_CLAIM_BUTTON_DEV)]
        private IWebElement _deleteClaimDev;

        [FindsBy(How = How.XPath, Using = DELETE_CLAIM_BUTTON_PATCH)]
        private IWebElement _deleteClaimPatch;

        [FindsBy(How = How.XPath, Using = DELETE_CLAIM_CONFIRM_BUTTON)]
        private IWebElement _deleteClaimConfirm;

        [FindsBy(How = How.XPath, Using = ENVELOPPE_ICON)]
        private IWebElement _enveloppeIcon;

        [FindsBy(How = How.ClassName, Using = "counter")]
        private IWebElement _totalNumber;

        //__________________________________ Filtres _______________________________________

        [FindsBy(How = How.Id, Using = RESET_FILTER_DEV)]
        private IWebElement _resetFilterDev;

        [FindsBy(How = How.XPath, Using = RESET_FILTER_PATCH)]
        private IWebElement _resetFilterPatch;

        [FindsBy(How = How.Id, Using = FILTER_NUMBER)]
        private IWebElement _searchByNumber;

        [FindsBy(How = How.Id, Using = FILTER_SUPPLIER)]
        private IWebElement _supplierFilter;

        [FindsBy(How = How.XPath, Using = UNSELECT_ALL_SUPPLIER)]
        private IWebElement _unselectAllSupplier;

        [FindsBy(How = How.XPath, Using = SEARCH_SUPPLIER)]
        private IWebElement _searchSupplier;

        [FindsBy(How = How.Id, Using = FILTER_DATE_FROM)]
        public IWebElement _dateFrom;

        [FindsBy(How = How.Id, Using = FILTER_DATE_TO)]
        public IWebElement _dateTo;

        [FindsBy(How = How.XPath, Using = FILTER_SHOW_ALL_CLAIMS_DEV)]
        private IWebElement _showAllClaimsDev;

        [FindsBy(How = How.XPath, Using = FILTER_SHOW_ALL_CLAIMS_PATCH)]
        private IWebElement _showAllClaimsPatch;

        [FindsBy(How = How.Id, Using = FILTER_NOT_VALIDATED_DEV)]
        private IWebElement _showNotValidatedDev;

        [FindsBy(How = How.XPath, Using = FILTER_NOT_VALIDATED_PATCH)]
        private IWebElement _showNotValidatedPatch;

        [FindsBy(How = How.Id, Using = FILTER_VALIDATED_DEV)]
        private IWebElement _showValidatedOnlyDev;

        [FindsBy(How = How.XPath, Using = FILTER_VALIDATED_PATCH)]
        private IWebElement _showValidatedOnlyPatch;

        [FindsBy(How = How.Id, Using = FILTER_PARTIALLY_INVOICED_DEV)]
        private IWebElement _showValidatedPartialInvoicedDev;

        [FindsBy(How = How.XPath, Using = FILTER_PARTIALLY_INVOICED_PATCH)]
        private IWebElement _showValidatedPartialInvoicedPatch;

        [FindsBy(How = How.Id, Using = FILTER_SHOW_ALL_VISU_DEV)]
        private IWebElement _showAllDev;

        [FindsBy(How = How.XPath, Using = FILTER_SHOW_ALL_VISU_PATCH)]
        private IWebElement _showAllPatch;

        [FindsBy(How = How.Id, Using = FILTER_ACTIVE_DEV)]
        private IWebElement _showActiveDev;

        [FindsBy(How = How.XPath, Using = FILTER_ACTIVE_PATCH)]
        private IWebElement _showActivePatch;

        [FindsBy(How = How.Id, Using = FILTER_INACTIVE_DEV)]
        private IWebElement _showNotActiveDev;

        [FindsBy(How = How.XPath, Using = FILTER_INACTIVE_PATCH)]
        private IWebElement _showNotActivePatch;

        [FindsBy(How = How.Id, Using = FILTER_SORTBY)]
        private IWebElement _sortBy;

        [FindsBy(How = How.Id, Using = FILTER_SHOW_ALL_STATUS_DEV)]
        private IWebElement _statusAllDev;

        [FindsBy(How = How.XPath, Using = FILTER_SHOW_ALL_STATUS_PATCH)]
        private IWebElement _statusAllPatch;

        [FindsBy(How = How.XPath, Using = FILTER_OPENED_DEV)]
        private IWebElement _statusOpenedDev;

        [FindsBy(How = How.XPath, Using = FILTER_OPENED_PATCH)]
        private IWebElement _statusOpenedPatch;

        [FindsBy(How = How.XPath, Using = FILTER_CLOSED_DEV)]
        private IWebElement _statusClosedDev;

        [FindsBy(How = How.XPath, Using = FILTER_CLOSED_PATCH)]
        private IWebElement _statusClosedPatch;

        [FindsBy(How = How.Id, Using = FILTER_SITE)]
        private IWebElement _siteFilter;

        [FindsBy(How = How.XPath, Using = UNSELECT_ALL)]
        private IWebElement _unselectAll;

        [FindsBy(How = How.XPath, Using = SEARCH_SITE)]
        private IWebElement _searchSite;

        [FindsBy(How = How.Id, Using = FILTER_SITE_PLACES)]
        private IWebElement _sitePlaceFilter;

        [FindsBy(How = How.XPath, Using = SEARCH_SITE_PLACES)]
        private IWebElement _searchSitePlace;

        [FindsBy(How = How.XPath, Using = UNSELECT_ALL_SITE_PLACES)]
        private IWebElement _unselectAllPlace;

        [FindsBy(How = How.Id, Using = DELIVERY_ORDER_NUMBER)]
        private IWebElement _deliveryOrderNumber;

        [FindsBy(How = How.Id, Using = DATE)]
        private IWebElement _date;

        [FindsBy(How = How.Id, Using = STATUS)]
        private IWebElement _status;

        [FindsBy(How = How.Id, Using = ACTIVATE)]
        private IWebElement _activate;

        [FindsBy(How = How.Id, Using = NUMBER)]
        private IWebElement _number;

        [FindsBy(How = How.Id, Using = COMMENT)]
        private IWebElement _comment;
        [FindsBy(How = How.Id, Using = SUPPLIER_CURRENCY)]
        private IWebElement _supplierCurrency;
        [FindsBy(How = How.Id, Using = LOCAL_CURRENCY)]
        private IWebElement _localCurrency;
        [FindsBy(How = How.Id, Using = CANCEL)]
        private IWebElement _cancel;
        public enum FilterType
        {
            ByNumber,
            Suppliers,
            DateFrom,
            DateTo,
            ShowAllClaims,
            ShowNotValidated,
            ShowValidatedOnly,
            ShowValidatedPartialInvoiced,
            ShowAll,
            ShowActive,
            ShowInactive,
            SortBy,
            All,
            Opened,
            Closed,
            Site,
            SitePlaces
        }

        public static bool operator ==(ClaimsPage a, object b)
        {
            return true;
        }
        public static bool operator !=(ClaimsPage a, object b)
        {
            return true;
        }

        public void Filter(FilterType filterType, object value)
        {
            Actions action = new Actions(_webDriver);

            switch (filterType)
            {
                case FilterType.ByNumber:
                    _searchByNumber = WaitForElementIsVisibleNew(By.Id(FILTER_NUMBER));
                    _searchByNumber.SetValue(ControlType.TextBox, value);
                    break;
                case FilterType.Suppliers:
                    ComboBoxSelectById(new ComboBoxOptions(FILTER_SUPPLIER, (string)value));
                    break;
                case FilterType.DateFrom:
                    _dateFrom = WaitForElementIsVisibleNew(By.Id(FILTER_DATE_FROM));
                    _dateFrom.SetValue(ControlType.DateTime, value);
                    _dateFrom.SendKeys(Keys.Tab);
                    break;
                case FilterType.DateTo:
                    _dateTo = WaitForElementIsVisibleNew(By.Id(FILTER_DATE_TO));
                    _dateTo.SetValue(ControlType.DateTime, value);
                    _dateTo.SendKeys(Keys.Tab);
                    break;
                case FilterType.ShowAllClaims:
                    _showAllClaimsDev = WaitForElementExists(By.Id(FILTER_SHOW_ALL_CLAIMS_DEV));
                    action.MoveToElement(_showAllClaimsDev).Perform();
                    _showAllClaimsDev.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.ShowNotValidated:
                    _showNotValidatedDev = WaitForElementExists(By.Id(FILTER_NOT_VALIDATED_DEV));
                    action.MoveToElement(_showNotValidatedDev).Perform();
                    _showNotValidatedDev.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.ShowValidatedOnly:
                    _showValidatedOnlyDev = WaitForElementExists(By.Id(FILTER_VALIDATED_DEV));
                    action.MoveToElement(_showValidatedOnlyDev).Perform();
                    _showValidatedOnlyDev.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.ShowValidatedPartialInvoiced:
                    _showValidatedPartialInvoicedDev = WaitForElementExists(By.Id(FILTER_PARTIALLY_INVOICED_DEV));
                    action.MoveToElement(_showValidatedPartialInvoicedDev).Perform();
                    _showValidatedPartialInvoicedDev.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.ShowAll:
                    _showAllDev = WaitForElementExists(By.Id(FILTER_SHOW_ALL_VISU_DEV));
                    action.MoveToElement(_showAllDev).Perform();
                    _showAllDev.SendKeys(Keys.Tab);
                    _showAllDev.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.ShowActive:
                    _showActiveDev = WaitForElementExists(By.Id(FILTER_ACTIVE_DEV));
                    action.MoveToElement(_showActiveDev).Perform();
                    _showActiveDev.SendKeys(Keys.Tab);
                    _showActiveDev.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.ShowInactive:
                    _showNotActiveDev = WaitForElementExists(By.Id(FILTER_INACTIVE_DEV));
                    action.MoveToElement(_showNotActiveDev).Perform();
                    _showNotActiveDev.SendKeys(Keys.Tab);
                    _showNotActiveDev.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.SortBy:
                    _sortBy = WaitForElementExists(By.Id(FILTER_SORTBY));
                    action.MoveToElement(_sortBy).Perform();
                    _sortBy.Click();
                    var element = WaitForElementIsVisibleNew(By.XPath("//option[contains(@value,'" + value + "')]"));
                    _sortBy.SetValue(ControlType.DropDownList, element.Text);
                    _sortBy.Click();
                    break;
                case FilterType.All:
                    _statusAllDev = WaitForElementExists(By.Id(FILTER_SHOW_ALL_STATUS_DEV));
                    action.MoveToElement(_statusAllDev).Perform();
                    _statusAllDev.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.Opened:
                    _statusOpenedDev = WaitForElementExists(By.Id(FILTER_OPENED_DEV));
                    action.MoveToElement(_statusOpenedDev).Perform();
                    _statusOpenedDev.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.Closed:
                    _statusClosedDev = WaitForElementExists(By.Id(FILTER_CLOSED_DEV));
                    action.MoveToElement(_statusClosedDev).Perform();
                    _statusClosedDev.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.Site:
                    ScrollUntilSitesFilterIsVisible();
                    ComboBoxSelectById(new ComboBoxOptions(FILTER_SITE, (string)value));
                    break;

                case FilterType.SitePlaces:
                    _sitePlaceFilter = WaitForElementExists(By.Id(FILTER_SITE_PLACES));
                    action.MoveToElement(_sitePlaceFilter).Perform();
                    _sitePlaceFilter.Click();

                    _searchSitePlace = WaitForElementIsVisibleNew(By.XPath(SEARCH_SITE_PLACES));
                    _searchSitePlace.SetValue(ControlType.TextBox, value);

                    _unselectAllPlace = WaitForElementIsVisibleNew(By.XPath(UNSELECT_ALL_SITE_PLACES));
                    _unselectAllPlace.Click();

                    var searchSitePlaceValue = WaitForElementIsVisibleNew(By.XPath("//span[text()='" + value + "']"));
                    searchSitePlaceValue.Click();

                    _sitePlaceFilter.Click();
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);

            }

            WaitPageLoading();
            WaitForLoad();
            //Thread.Sleep(1500);

        }
        public object GetFilterValue(FilterType filterType)
        {
            switch (filterType)
            {
                case FilterType.ShowAllClaims:
                    _showAllClaimsDev = WaitForElementIsVisible(By.Id(FILTER_SHOW_ALL_CLAIMS_DEV));
                    return _showAllClaimsDev.Selected;


                case FilterType.ShowNotValidated:
                    _showNotValidatedDev = WaitForElementExists(By.Id(FILTER_NOT_VALIDATED_DEV));
                    return _showNotValidatedDev.Selected;

                case FilterType.ShowValidatedOnly:
                    _showValidatedOnlyDev = WaitForElementExists(By.Id(FILTER_VALIDATED_DEV));
                    return _showValidatedOnlyDev.Selected;

                case FilterType.ShowValidatedPartialInvoiced:
                    _showValidatedPartialInvoicedDev = WaitForElementExists(By.Id(FILTER_PARTIALLY_INVOICED_DEV));
                    return _showValidatedPartialInvoicedDev.Selected;

                case FilterType.ShowAll:
                    _showAllDev = WaitForElementExists(By.Id(FILTER_SHOW_ALL_VISU_DEV));
                    return _showAllDev.Selected;

                case FilterType.ShowActive:
                    _showActiveDev = WaitForElementExists(By.Id(FILTER_ACTIVE_DEV));
                    return _showActiveDev.Selected;

                case FilterType.ShowInactive:
                    _showNotActiveDev = WaitForElementExists(By.Id(FILTER_INACTIVE_DEV));
                    return _showNotActiveDev.Selected;
                case FilterType.All:
                    _statusAllDev = WaitForElementExists(By.Id(FILTER_SHOW_ALL_STATUS_DEV));
                    return _statusAllDev.Selected;
                case FilterType.Opened:
                    _statusOpenedDev = WaitForElementExists(By.Id(FILTER_OPENED_DEV));
                    return _statusOpenedDev.Selected;
                case FilterType.Closed:
                    _statusClosedDev = WaitForElementExists(By.Id(FILTER_CLOSED_DEV));
                    return _statusClosedDev.Selected;


            }
            return null;
        }
        public void ResetFilter()
        {
            IWebElement html = _webDriver.FindElement(By.TagName("html"));
            html.SendKeys(Keys.PageUp);
            html.SendKeys(Keys.PageUp);
            html.SendKeys(Keys.PageUp);
            html.SendKeys(Keys.PageUp);

            _resetFilterPatch = WaitForElementIsVisibleNew(By.XPath(RESET_FILTER_PATCH));
            _resetFilterPatch.Click();
            WaitForLoad();
            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                // WA_CLAI_Reset_Filter + 1 mois
                //Filter(FilterType.DateTo, DateUtils.Now);
            }
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
                    WaitPageLoading();
                    WaitForLoad();
                    _webDriver.FindElement(By.XPath(string.Format("//*/table[contains(@class,'table')]/tbody/tr[{0}]/td[1]/div/i[contains(@class,'circle-check')]", i + 1)));

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
                    _webDriver.FindElement(By.XPath(String.Format("//*[@id=\"list-item-with-action\"]/table/tbody/tr[{0}]/td[1]/div/i[contains(@class,'circle-xmark')]", i + 1)));

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

        public bool IsPartiallyInvoiced()
        {
            ReadOnlyCollection<IWebElement> invoiceStatus;
            invoiceStatus = _webDriver.FindElements(By.XPath("//*[@id='list-item-with-action']/table/tbody/tr[*]/td[2]/div/i"));


            if (invoiceStatus.Count == 0)
                return false;

            foreach (var elm in invoiceStatus)
            {
                if (!elm.GetAttribute("title").Equals("Partially invoiced"))
                {
                    return false;
                }
            }
            return true;
        }

        public bool IsSortedByNumber()
        {
            var listeNumbers = _webDriver.FindElements(By.XPath(CLAIM_NUMBER));

            if (listeNumbers.Count == 0)
                return false;

            var ancientNumber = int.Parse(listeNumbers[0].Text);

            foreach (var elm in listeNumbers)
            {
                try
                {
                    if (int.Parse(elm.Text) > ancientNumber)
                        return false;

                    ancientNumber = int.Parse(elm.Text);
                }
                catch
                {
                    return false;
                }
            }

            return true;
        }

        public bool VerifySite(string site)
        {
            var sites = _webDriver.FindElements(By.XPath(CLAIM_SITE));

            if (sites.Count == 0)
                return false;

            foreach (var elm in sites)
            {
                if (elm.Text != site)
                {
                    return false;
                }
            }
            return true;
        }

        public bool VerifySupplier(string supplier)
        {
            var suppliers = _webDriver.FindElements(By.XPath(CLAIM_SUPPLIER));

            if (suppliers.Count == 0)
                return false;

            foreach (var elm in suppliers)
            {
                if (elm.Text != supplier)
                {
                    return false;
                }
            }
            return true;
        }

        public bool IsSentByMail()
        {
            int tot = CheckTotalNumber() > 100 ? 100 : CheckTotalNumber();

            if (tot == 0)
                return false;
            for (int i = 0; i < tot; i++)
            {
                try
                {
                    WaitPageLoading();
                    WaitForLoad();
                    _webDriver.FindElement(By.XPath(string.Format(SEND_BY_MAIL, i + 1)));
                    return true;
                }
                catch
                {
                    return false;
                }
            }
            return true;
        }

        public bool IsDateRespected(DateTime fromDate, DateTime toDate, string dateFormat)
        {
            // Take the date format from the datepicker element and use it to format the date column to avoid date errors
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            var dates = _webDriver.FindElements(By.XPath(DELIVERY_DATE));
            var chercheColonneDate = WaitForElementIsVisible(By.XPath("//*[@id=\"list-item-with-action\"]/table/thead/tr/th[8]"));
            if (chercheColonneDate.Text.Contains("Delivery date"))
            {
                // on est sur Patch
            }
            else
            {
                dates = _webDriver.FindElements(By.XPath(DELIVERY_DATE2));
            }



            if (dates.Count == 0)
                return false;

            foreach (var elm in dates)
            {
                try
                {
                    DateTime date = DateTime.Parse(elm.Text, ci).Date;

                    if (DateTime.Compare(date, fromDate) < 0 || DateTime.Compare(date, toDate) > 0)
                        return false;
                }
                catch
                {
                    return false;
                }
            }

            return true;
        }

        public bool IsSortedByDate(string dateFormat)
        {
            // Take the date format from the datepicker element and use it to format the date column to avoid date errors
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            var dates = _webDriver.FindElements(By.XPath(DELIVERY_DATE));
            var chercheColonneDate = WaitForElementIsVisible(By.XPath("//*[@id=\"list-item-with-action\"]/table/thead/tr/th[8]"));
            if (chercheColonneDate.Text.Contains("Delivery date"))
            {
                // on est sur Patch
            }
            else
            {
                dates = _webDriver.FindElements(By.XPath(DELIVERY_DATE2));
            }

            if (dates.Count == 0)
                return false;

            var ancientDate = DateTime.Parse(dates[0].Text, ci);

            foreach (var elm in dates)
            {
                try
                {
                    if (DateTime.Compare(ancientDate.Date, DateTime.Parse(elm.Text, ci)) < 0)
                        return false;

                    ancientDate = DateTime.Parse(elm.Text, ci);
                }
                catch
                {
                    return false;
                }
            }

            return true;
        }


        //__________________________________  Méthodes __________________________________________________

        // General
        public ClaimsCreateModalPage ClaimsCreatePage()
        {
            WaitLoading();
            ShowPlusMenu();
            _createNewClaim = WaitForElementIsVisible(By.XPath(NEW_CLAIM_BUTTON), nameof(NEW_CLAIM_BUTTON));
            _createNewClaim.Click();
            WaitForLoad();

            return new ClaimsCreateModalPage(_webDriver, _testContext);
        }

        public PrintReportPage PrintClaims()
        {

            ShowExtendedMenu();

            _printClaimResults = WaitForElementIsVisible(By.Id(PRINT_CLAIM_RESULTS));
            _printClaimResults.Click();
            WaitForLoad();

            if (isElementVisible(By.XPath("//h4[contains(text(), 'Print')]"))) //new modal for include prices on report
            {
                _modalPrintButton = WaitForElementIsVisible(By.Id(MODAL_PRINT_BUTTON));
                _modalPrintButton.Click();
                WaitForLoad();
            }

            IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-pdf']"));

            //Results are opened in a new tab, switch the driver to the newly created one
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(30));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new PrintReportPage(_webDriver, _testContext);
        }

        public FileInfo GetExportPdfFile(FileInfo[] taskFiles)
        {
            List<FileInfo> correctDownloadFiles = new List<FileInfo>();

            foreach (var file in taskFiles)
            {
                //  Test REGEX
                if (IsPdfFileCorrect(file.Name))
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

        public bool IsPdfFileCorrect(string filePath)
        {
            string mois = "(?:0[1-9]|1[0-2])";  // mois
            string annee = "\\d{4}";            // annee YYYY
            string jour = "[0-3]\\d";           // jour
            string nombre = "\\d{6}";           // nombre


            Regex r = new Regex("^Claim notes_" + annee + mois + jour + "_" + nombre + ".pdf$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = r.Match(filePath);

            return m.Success;
        }

        public void ExportExcelFile(bool printValue, string downloadPath)
        {

            ShowExtendedMenu();
            WaitPageLoading();
            WaitForLoad();
            _export = WaitForElementIsVisible(By.Id(EXPORT));
            _export.Click();
            WaitForLoad();

            if (printValue)
            {
                FileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));
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

        public bool IsExcelFileCorrect(string filePath)
        {
            string mois = "(?:0[1-9]|1[0-2])";         // mois
            string space = "(\\s)";                    // Espace
            string annee = "\\d{4}";                   // annee YYYY
            string jour = "[0-3]\\d";                  // jour
            string heure = "(?:0[0-9]|1[0-9]|2[0-3])"; // heure
            string minutes = "[0-5]\\d";               // minutes
            string secondes = "[0-5]\\d";              // secondes


            Regex r = new Regex("^claimsnotes" + space + annee + "-" + mois + "-" + jour + space + heure + "-" + minutes + "-" + secondes + ".xlsx$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = r.Match(filePath);

            return m.Success;
        }

        public void SendResultsByMail()
        {
            WaitForLoad();
            ShowExtendedMenu();

            _sendResultsByMail = WaitForElementIsVisible(By.Id(SEND_RESULTS_BY_MAIL));
            _sendResultsByMail.Click();
            WaitForLoad();

            _confirmSendMail = WaitForElementIsVisible(By.Id(CONFIRM_SEND));
            _confirmSendMail.Click();
            WaitForLoad();
            WaitPageLoading();
        }
        public bool SendByMail()
        {
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;

            ShowExtendedMenu();

            _sendResultsByMail = WaitForElementIsVisible(By.Id(SEND_RESULTS_BY_MAIL));
            _sendResultsByMail.Click();
            WaitPageLoading();

            if (isElementExists(By.Id(CONFIRM_SEND)))
            {
                _confirmSendMail = _webDriver.FindElement(By.Id(CONFIRM_SEND));
                javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", _confirmSendMail);
                WaitForLoad();
                _confirmSendMail.Click();
                WaitPageLoading();
            }
            else
            {
                _cancel = WaitForElementIsVisible(By.Id(CANCEL));
                _cancel.Click();
                WaitForLoad();
                return false;
            }
            WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until(d => !isElementVisible(By.XPath(SEND_RESULTS_BY_MAIL_POPUP)));
            return true;
        }


        // Tableau
        public ClaimsItem SelectFirstClaim()
        {
            _firstClaim = WaitForElementIsVisible(By.XPath(FIRST_CLAIM));
            _firstClaim.Click();
            WaitForLoad();

            return new ClaimsItem(_webDriver, _testContext);
        }

        public string GetFirstID()
        {
            WaitForLoad();
            if (isElementVisible(By.XPath(FIRST_CLAIM)))
            {
                _firstClaim = WaitForElementExists(By.XPath(FIRST_CLAIM));
                return _firstClaim.Text;
            }
            else
            {
                return null;
            }

        }

        public void DeleteClaim(string claimNumber)
        {
            WaitPageLoading();
            WaitForLoad();
            _deleteClaimDev = WaitForElementIsVisible(By.XPath(string.Format(DELETE_CLAIM_BUTTON_DEV, claimNumber)), nameof(DELETE_CLAIM_BUTTON_DEV));
            _deleteClaimDev.Click();
            WaitForLoad();

            _deleteClaimConfirm = WaitForElementIsVisible(By.Id(DELETE_CLAIM_CONFIRM_BUTTON));
            _deleteClaimConfirm.Click();
            WaitForLoad();
        }

        public void ClickFirstClaim()
        {
            var firstClaim = WaitForElementIsVisible(By.XPath(FIRST_CLAIM_NOTE));
            firstClaim.Click();
            WaitForLoad();
        }
        public string GetFirstClaimNumber()
        {
            var firstClaim = WaitForElementIsVisible(By.XPath(FIRST_CLAIM_NUMBER));
            return firstClaim.Text;

        }
        public ClaimsItem ClickGeneralInformationSubMenu()
        {
            var generalinformation = WaitForElementIsVisible(By.XPath(GENERAL_INFORMATION_SUB_MENU));
            generalinformation.Click();
            WaitPageLoading();
            WaitForLoad();
            return new ClaimsItem(_webDriver, _testContext);

        }


        public string ChangeClaim(string deliveryOrderNumber, string date, string status, bool activate, string comment)
        {

            _deliveryOrderNumber = WaitForElementIsVisible(By.Id(DELIVERY_ORDER_NUMBER));
            _deliveryOrderNumber.Clear();
            _deliveryOrderNumber.SendKeys(deliveryOrderNumber);
            WaitPageLoading();
            WaitForLoad();

            _date = WaitForElementIsVisible(By.Id(DATE));
            _date.Clear();
            _date.SendKeys(date);
            _date.SendKeys(Keys.Tab);
            WaitPageLoading();
            WaitForLoad();

            _status = WaitForElementIsVisible(By.Id(STATUS));
            _status.Click();
            WaitPageLoading();
            WaitForLoad();

            var element = WaitForElementIsVisible(By.XPath("//option[text()='" + status + "']"));
            _status.SetValue(ControlType.DropDownList, element.Text);
            _status.Click();
            WaitPageLoading();
            WaitForLoad();

            _activate = WaitForElementExists(By.Id(ACTIVATE));
            _activate.SetValue(ControlType.CheckBox, false);
            WaitPageLoading();

            WaitForLoad();

            _comment = WaitForElementIsVisible(By.Id(COMMENT));
            _comment.Clear();
            _comment.SendKeys(comment);
            WaitPageLoading();
            WaitForLoad();

            _number = WaitForElementIsVisible(By.Id(NUMBER));
            var number = _number.GetAttribute("value");
            WaitPageLoading();
            WaitForLoad();

            return number;
        }

        public bool claimModified(string deliveryorderNumber, string date, string status, bool activate, string comment)
        {
            _deliveryOrderNumber = WaitForElementIsVisible(By.Id(DELIVERY_ORDER_NUMBER));
            string deliveryordernumberview = _deliveryOrderNumber.GetAttribute("value");
            _date = WaitForElementIsVisible(By.Id(DATE));
            string dateview = _date.GetAttribute("value");

            _status = WaitForElementIsVisible(By.Id(STATUS));
            string statusview = _status.GetAttribute("value");

            _activate = WaitForElementExists(By.Id(ACTIVATE));
            bool activateview = _activate.Selected;

            _comment = WaitForElementIsVisible(By.Id(COMMENT));
            string commentview = _comment.GetAttribute("value");

            if (deliveryordernumberview != deliveryorderNumber)
            {
                return false;
            }
            if (!dateview.Equals(date))
            {
                return false;
            }
            if (!activateview.Equals(activate))
            {
                return false;
            }
            if (!commentview.Equals(comment))
            {
                return false;
            }
            if (!((statusview == "2" && status.Equals("Closed")) || (statusview == "1" && status.Equals("Opened"))))
            {
                return false;
            }

            return true;
        }

        public void ClickFooterSubMenu()
        {
            var footerButton = WaitForElementExists(By.XPath(FOOTER_SUBMENU));
            footerButton.Click();
        }
        public string GetTotalClaimFromItems()
        {
            var totalClaim = WaitForElementIsVisible(By.XPath(TOTAL_CLAIM_ITEMS));
            return totalClaim.Text;
        }
        public string GetTotalSactionsFromItems()
        {
            var totalSanctions = "";
            try
            {
                totalSanctions = WaitForElementIsVisible(By.XPath(TOTAL_SANCTIONS_ITEMS)).Text;
            }
            catch (Exception e)
            {
                return "0.00";
            }
            return totalSanctions;
        }
        public string GetTotalGrossAmountFromFooter()
        {
            var totalGrossAmount = WaitForElementIsVisible(By.XPath(TOTAL_GROSS_AMOUNT_FOOTER));
            return totalGrossAmount.Text;
        }
        public string GetSanctionsFromFooter()
        {
            var totalSanctions = WaitForElementIsVisible(By.XPath(TOTAL_SANCTIONS_FOOTER));
            return totalSanctions.Text;
        }
        public string GetSanctionVATFooter()
        {
            var VATSanctionsfoorter = WaitForElementIsVisible(By.XPath(VAT_SANCTIONS_FOOTER));
            return VATSanctionsfoorter.Text;
        }

        public void ClickFirstItem()
        {
            var first_item = WaitForElementIsVisible(By.XPath(FIRST_ITEM));
            first_item.Click();
            WaitForLoad();
        }

        public void AddSanctionToItem(string sanctionAmount)
        {
            var sanction = WaitForElementIsVisible(By.Id(SANCTION_AMOUNT));
            sanction.Clear();
            sanction.SendKeys(sanctionAmount.ToString());
        }

        public bool CheckEnveloppe(string claimNumber)
        {
            bool isEnveloppe;
            if (isElementVisible(By.XPath(string.Format(ENVELOPPE_ICON, claimNumber))))
            {
                isEnveloppe = true;
            }
            else
            {
                isEnveloppe = false;
            }

            return isEnveloppe;
        }

        public string GetLocalCurrency()
        {
            _localCurrency = WaitForElementExists(By.XPath(LOCAL_CURRENCY));
            return _localCurrency.Text;
        }
        public string GetSupplierCurrency()
        {
            _supplierCurrency = WaitForElementExists(By.XPath(SUPPLIER_CURRENCY));
            return _supplierCurrency.Text;
        }
        public string GetRecievedQuantity()
        {
            WaitForLoad();
            if (isElementVisible(By.XPath("/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/form/div[2]/div[6]/span")))
            {
                var recievedQuantity = WaitForElementExists(By.XPath("/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/form/div[2]/div[6]/span"));
                return recievedQuantity.Text;
            }
            else
            {
                var recievedQuantity = WaitForElementExists(By.XPath("/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/form/div/div[6]/span"));
                return recievedQuantity.Text;

            }

        }
        public string GetDecreasedQuantity()
        {
            WaitForLoad();
            WaitPageLoading();
            if (isElementVisible(By.XPath("/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/form/div[2]/div[14]/span")))
            {
                var recievedQuantity = WaitForElementExists(By.XPath("/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/form/div[2]/div[14]/span"));
                return recievedQuantity.Text;
            }
            else
            {
                var recievedQuantity = WaitForElementExists(By.XPath("/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/form/div/div[13]/span"));
                return recievedQuantity.Text;

            }

        }

        public void SetRecievedQuantity(string value)
        {

            WaitPageLoading();
            WaitForLoad();
            var recievedQuantity = WaitForElementExists(By.Id("item_CndRowDto_ClaimRNQuantity"));
            recievedQuantity.SetValue(ControlType.TextBox, value);
            WaitForLoad();
        }
        public void SetDecreasedQuantity(string value)
        {
            WaitForLoad();
            var decreasedQuantity = WaitForElementExists(By.Id("item_CndRowDto_DecreasedQty"));
            decreasedQuantity.SetValue(ControlType.TextBox, value);
            WaitForLoad();
            WaitPageLoading();
        }

        public void Refresh()
        {
            var threePoints = WaitForElementIsVisible(By.XPath("//*[@id=\"div-body\"]/div/div[1]/div/div[1]/button"));
            threePoints.Click();
            var refresh = WaitForElementIsVisible(By.Id("btn-receipt-notes-refresh"));
            refresh.Click();
            WaitForLoad();
            if (isElementExists(By.Id("dataAlertCancel")))
            {

                var cancelButton = WaitForElementIsVisible(By.Id("dataAlertCancel"));
                cancelButton.Click();
                WaitForLoad();

            }
        }

        public SupplierInvoicesPage GenerateSI(string invoiceNumber)
        {
            var threePoints = WaitForElementIsVisible(By.XPath("//*[@id=\"div-body\"]/div/div[1]/div/div[1]/button"));
            threePoints.Click();
            var generateSI = WaitForElementIsVisible(By.Id("btn-generate-associated-supplier-invoice"));
            generateSI.Click();
            CreateNewSupplierInvoice(invoiceNumber);

            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));

            wait.Until((driver) => driver.WindowHandles.Count > 1);

            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new SupplierInvoicesPage(_webDriver, _testContext);
        }
        public void CreateNewSupplierInvoice(string invoiceNumber)
        {
            var invoiceNumberInput = WaitForElementIsVisible(By.Id("tb-new-supplier-invoice-number"));
            invoiceNumberInput.SetValue(ControlType.TextBox, invoiceNumber);

            var createFromCheckBox = WaitForElementIsVisible(By.XPath("//*[@id=\"form-create-supplier-invoice\"]/div/div[9]/div[1]/div/div/div/div[1]/div"));
            createFromCheckBox.Click();

            var createButton = WaitForElementIsVisible(By.Id("btn-submit-form-create-supplier-invoice"));
            createButton.Click();
        }
        public void ScrollUntilSitesFilterIsVisible()
        {
            ScrollUntilElementIsInView(By.Id(FILTER_SITE));
        }
        public void BackToList()
        {
            WaitPageLoading();
            WaitForLoad();
            var backToListBtn = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            WaitForLoad();
            backToListBtn.Click();
        }
        public int GetNumberSelectedSiteFilter()
        {
            var listSitesSelectedFilters = _webDriver.FindElements(By.XPath(LIST_SITE_FILTER));
            var numberSitesSelectedSite = listSitesSelectedFilters
               .Where(p => p.Selected == true).Count();

            return numberSitesSelectedSite;
        }
        public int GetNumberSelectedSupplierFilter()
        {
            var listSelectedSupplierFilters = _webDriver.FindElements(By.XPath(LIST_SUPPLIER_FILTER));
            var numberSelectedSupplier = listSelectedSupplierFilters.Where(p => p.Selected).Count();

            return numberSelectedSupplier;
        }
        public int GetNumberSelectedSitesPlacesFilter()
        {
            var listSelectedSitesPlacesFilters = _webDriver.FindElements(By.XPath(LIST_SITE_PLACES_FILTER));
            var numberSelectedSitesPlaces = listSelectedSitesPlacesFilters.Where(p => p.Selected).Count();
            return numberSelectedSitesPlaces;
        }
        public string GetFirstDate()
        {
            WaitForLoad();
            if (isElementVisible(By.XPath(FIRST_DATE)))
            {
                _firstClaim = WaitForElementExists(By.XPath(FIRST_DATE));
                return _firstClaim.Text;
            }
            else
            {
                return null;
            }

        }
        //public override int CheckTotalNumber()
        //{
        //    WaitForLoad();
        //    _totalNumber = WaitForElementExists(By.XPath("//*[@id=\"tabContentItemContainer\"]/div/h1/span"));
        //    int nombre = Int32.Parse(_totalNumber.GetAttribute("innerText"));
        //    return nombre;
        //}


        public string GetDate()
        {
            var date = WaitForElementIsVisible(By.XPath(DATE_CLAIM));
            return date.Text;
        }

        public string GetDateTo()
        {
            IWebElement date = WaitForElementIsVisible(By.Id("date-picker-end"));
            string result = date.GetAttribute("value");
            return result;
        }
    }
}