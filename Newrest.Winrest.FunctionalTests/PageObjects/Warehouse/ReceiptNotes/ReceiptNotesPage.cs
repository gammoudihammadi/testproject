using Microsoft.VisualStudio.TestTools.UnitTesting;
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

namespace Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.ReceiptNotes

{
    public class ReceiptNotesPage : PageBase
    {

        public ReceiptNotesPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //__________________________________ Constantes _____________________________________

        // General
        private const string NEW_RECEIPT_NOTES = "//*[@id=\"div-body\"]/div/div[2]/div[1]/div/div[2]/div/a[text()='New receipt note']";

        private const string EXTENDED_MENU = "//*[@id=\"div-body\"]/div/div[2]/div[1]/div/div[1]/button";
        private const string PRINT_RN_RESULTS = "//*[@id=\"div-body\"]/div/div[2]/div[1]/div/div[1]/div/a[text()='Print receipt notes results']";
        private const string CONFIRM_PRINT = "printButton";
        private const string EXPORT = "btn-export-excel";
        private const string EXPORT_MATCHING = "btn-export-matching";
        private const string MENU_CLOSE = "//*/a[text()='Close']";
        private const string CLOSE_PAGING = "//*[@id='close-receipt-notes']/div[1]/div[3]/nav/ul/li[*]/a";
        private const string CLOSE_VALIDATE = "//*/button[text()='Validate']";
        private const string EXPORTED_WITH_REVERSE_GENERATED = "//*[@id=\"ExportedWithReverseGenerated\"]";

        private const string TOP_RECEIVED = "//*[@id=\"div-body\"]/div/div[2]/div[1]/div/div[1]/div/a[text()='Top Received']";
        private const string TOP_RECEIVED_START_DATE = "date-picker-topreceivedstart";
        private const string TOP_RECEIVED_SITES = "SelectedTopReceivedSites_ms";
        private const string TOP_RECEIVED_SITES_UNCHECK_ALL = "/html/body/div[13]/div/ul/li[2]/a/span[2]";
        private const string TOP_RECEIVED_SITE_SEARCH = "/html/body/div[13]/div/div/label/input";
        private const string TOP_RECEIVED_SITE_TO_CHECK = "//label/span[text()='{0}']";
        private const string TOP_RECEIVED_COUNT = "SelectedTopReceivedCount";
        private const string TOP_RECEIVED_EXPORT = "validateExport";

        private const string EXPORT_FOR_SAGE = "btn-export-for-sage";
        private const string CONFIRM_EXPORT_SAGE = "btn-popup-validate";
        private const string CANCEL_EXPORT_SAGE = "btn-cancel-popup";
        private const string DOWNLOAD_FILE = "btn-download-file";
        private const string GENERATE_SAGE_TXT = "btn-export-for-sage-txt-generation";

        // Tableau
        private const string FIRST_RECEIPT_NOTE_NUMBER = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[1]/td[3]";
        private const string DELETE_FIRST_RECEIPT_NOTE = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]//span[contains(@class, 'glyphicon glyphicon-trash')]";
        private const string DELETE_RECEIPT_NOTE = "//*[@id=\"list-item-with-action\"]/table/tbody/tr/td[10]/a[contains(@class,'btn-ReceiptNote-delete')]/span";
        
        private const string CONFIRM_DELETE = "dataConfirmOK";
        private const string IMG_VALID = "//*[@id='list-item-with-action']/table/tbody/tr[1]/td[1]/img";
        private const string ELEMENT_DEV = "//tr/td[{0}]/span[@class='fas fa-save green-text'][contains(@title, 'Accounted')]";
        private const string ELEMENT_PATCH = "//tr[{0}]//i[@class='fa fa-thumbs-up' and @title='Accounted']";
        private const string LIGNES_RN = "//*/div[@id='list-item-with-action']/table/tbody/tr/td[3]";
        private const string FIRST_RN_TOTAL = "/html/body/div[2]/div/div[2]/div[2]/table/tbody/tr[1]/td[9]";
        private const string FIRST_RN = "/html/body/div[2]/div/div[2]/div[2]/table/tbody/tr[1]/td[5]";
        private const string ICON = "//*[@id=\"list-item-with-action\"]/table/tbody/tr/td[1]/div";
        private const string LIST_DATES = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[8]";
        // Filtres
        private const string RESET_FILTER_PATCH = "//*[@id=\"formSearchReceiptNotes\"]/div[1]/a";

        private const string FILTER_NUMBER = "SearchNumber";
        private const string FILTER_PURCHASE_NUMBER = "SearchNumberSecondary";
        private const string FILTER_DELIVERY_NUMBER = "SearchNumberThirdary";
        private const string FILTER_SUPPLIER = "SelectedSuppliers_ms";
        private const string UNSELECT_ALL_SUPPLIER = "/html/body/div[10]/div/ul/li[2]/a";
        private const string FILTER_DATE_FROM = "date-picker-start";
        private const string FILTER_DATE_TO = "date-picker-end";

        private const string FILTER_SHOW_SENT_TO_SAGE_AND_IN_ERROR_ONLY = "SentToSageAndInErrorOnly";
        private const string FILTER_SHOW_VALIDATED_NO_SAGE = "ShowValidatedAndNotSentToSageCegid";

        private const string FILTER_SHOW_ALL_DEV = "ShowAll";
        private const string FILTER_SHOW_ALL_PATCH = "//*[@id=\"ShowOnlyValue\"][@value='All']";

        private const string FILTER_NOT_VALIDATED_DEV = "NotValidatedOnly";
        private const string FILTER_NOT_VALIDATED_PATCH = "//*[@id=\"ShowOnlyValue\"][@value='NotValidatedOnly']";

        private const string FILTER_VALIDATED_DEV = "ValidatedOnly";
        private const string FILTER_VALIDATED_PATCH = "//*[@id=\"ShowOnlyValue\"][@value='ValidatedOnly']";

        private const string FILTER_NOT_INVOICED = "ValidatedAndNotInvoiced";

        private const string FILTER_PARTIALLY_INVOICED_DEV = "PartiallyInvoiced";
        private const string FILTER_PARTIALLY_INVOICED_PATCH = "//*[@id=\"ShowOnlyValue\"][@value='PartiallyInvoiced']";

        private const string FILTER_WITH_CLAIM_DEV = "WithClaims";
        private const string FILTER_WITH_CLAIM_PATCH = "//*[@id=\"ShowOnlyValue\"][@value='WithClaims']";

        private const string FILTER_EXPORTED_MANUALLY_DEV = "ExportedForSageManually";
        private const string FILTER_EXPORTED_MANUALLY_PATCH = "//html/body/div[2]/div/div[1]/div[1]/div/form/div[8]/label[8]/input";

        private const string FILTER_SHOW_ALL_VISU_DEV = "ShowActiveAndInactive";
        private const string FILTER_SHOW_ALL_VISU_PATCH = "//*[@id=\"ShowActive\"][@value='All']";

        private const string FILTER_ACTIVE_DEV = "ActiveOnly";
        private const string FILTER_ACTIVE_PATCH = "//*[@id=\"ShowActive\"][@value='ActiveOnly']";

        private const string FILTER_INACTIVE_DEV = "InactiveOnly";
        private const string FILTER_INACTIVE_PATCH = "//*[@id=\"ShowActive\"][@value='InactiveOnly']";

        private const string FILTER_SORTBY = "cbSortBy";

        private const string STATUS = "//*[@id=\"formSearchReceiptNotes\"]/div[11]/a";

        private const string FILTER_SHOW_ALL_STATUS_DEV = "ShowAllStatus";
        private const string FILTER_SHOW_ALL_STATUS_PATCH = "//*[@id=\"Status\"][@value='All']";

        private const string FILTER_OPENED_DEV = "Opened";
        private const string FILTER_OPENED_PATCH = "//*[@id=\"Status\"][@value='Opened']";

        private const string FILTER_CLOSED_DEV = "Closed";
        private const string FILTER_CLOSED_PATCH = "//*[@id=\"Status\"][@value='Closed']";

        private const string FILTER_SITE = "SelectedSites_ms";
        private const string SEARCH_SITE = "/html/body/div[11]/div/div/label/input";
        private const string UNSELECT_ALL_SITES = "/html/body/div[11]/div/ul/li[2]/a";
        private const string FILTER_SITE_PLACES = "SelectedSitePlaces_ms";
        private const string SEARCH_SITE_PLACES = "/html/body/div[12]/div/div/label/input";
        private const string UNSELECT_ALL_SITE_PLACES = "/html/body/div[12]/div/ul/li[2]/a";

        private const string ALL_RN_ROWS = "//*[@id='list-item-with-action']/table/tbody/tr[*]";
        private const string FIRST_ICON = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[{0}]/td[1]/img[1]";
        private const string VALIDATION = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[{0}]/td[1]/img[@alt=\"Valid\"]";
        private const string INVOICE_STATUS = ALL_RN_ROWS + "/td[2]/span";
        private const string RN_NUMBER = ALL_RN_ROWS + "/td[3]";
        private const string RN_SITE = ALL_RN_ROWS + "/td[4]";
        private const string RN_SUPPLIER = ALL_RN_ROWS + "/td[5]";
        private const string SEND_BY_MAIL = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[{0}]/td[6]/span";
        private const string WITH_CLAIM = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[{0}]/td[7]/span";
        private const string DELIVERY_DATE = ALL_RN_ROWS + "/td[8]";
        private const string DELIVERY_DATE2 = ALL_RN_ROWS + "/td[9]";
        private const string ROWS_FOR_PAGINATION = ALL_RN_ROWS + "/td[2]";
        private const string ITEM_NAME = ALL_RN_ROWS + "/td[3]";
        private const string VALIDATION_ICON = ALL_RN_ROWS + "/td[1]/div/i";
        private const string CEGID_ICON = ALL_RN_ROWS + "/td[1]/div[2]/i";
        private const string INVOICED_ICONE = ALL_RN_ROWS + "/td[2]/div/i";
        private const string WITHCLAIM_ICONE = "//*[@id='list-item-with-action']/table/tbody/tr[1]/td[6]/div/div/div[2]/span";

        private const string FILTER_EXPORTED_WITH_PROVISION_GENERATED = "//*[@id=\"ExportedWithProvisionGenerated\"]";
        private const string Exported_ForSage_Or_CegidAutomatically = "ExportedForSageAutomatically";
        private const string TOTAL_RCN_NUMBER = "/html/body/div[2]/div/div[2]/div[1]/h1/span";
        private const string PAGEING_SELECT = "//*[@Id=\"page-size-selector\" and @data-pager-formid='formCloseReceiptNotes']";
        private const string PRICE_INDEX = "//*[@id=\"list-item-with-action\"]/table/tbody/tr/td[9]";
        private const string TOTAL_VAT = "//*[@id=\"list-item-with-action\"]/table/tbody/tr/td[9]";
        private const string DELIVERY_LOCATION = "//*[@id=\"list-item-with-action\"]/table/tbody/tr/td[6]/div/div/div/b";


        //__________________________________ Variables _____________________________________

        // General
        [FindsBy(How = How.XPath, Using = NEW_RECEIPT_NOTES)]
        private IWebElement _createNewReceiptNote;

        [FindsBy(How = How.XPath, Using = EXTENDED_MENU)]
        private IWebElement _showExtendedMenu;

        [FindsBy(How = How.XPath, Using = PRINT_RN_RESULTS)]
        private IWebElement _printReceiptNoteResults;

        [FindsBy(How = How.Id, Using = CONFIRM_PRINT)]
        private IWebElement _confirmPrint;

        [FindsBy(How = How.Id, Using = EXPORT)]
        private IWebElement _export;

        [FindsBy(How = How.Id, Using = EXPORT_MATCHING)]
        private IWebElement _exportMatching;

        [FindsBy(How = How.XPath, Using = MENU_CLOSE)]
        private IWebElement _menuClose;

        [FindsBy(How = How.XPath, Using = TOP_RECEIVED)]
        private IWebElement _topReceived;

        [FindsBy(How = How.Id, Using = TOP_RECEIVED_START_DATE)]
        private IWebElement _topReceivedStartDate;

        [FindsBy(How = How.Id, Using = TOP_RECEIVED_SITES)]
        private IWebElement _topReceivedSites;

        [FindsBy(How = How.Id, Using = TOP_RECEIVED_SITES_UNCHECK_ALL)]
        private IWebElement _topReceivedSitesUcheckAll;

        [FindsBy(How = How.Id, Using = TOP_RECEIVED_SITE_SEARCH)]
        private IWebElement _topReceivedSiteSearch;

        [FindsBy(How = How.Id, Using = TOP_RECEIVED_SITE_TO_CHECK)]
        private IWebElement _topReceivedSiteToCheck;

        [FindsBy(How = How.Id, Using = TOP_RECEIVED_COUNT)]
        private IWebElement _topReceivedCount;

        [FindsBy(How = How.Id, Using = TOP_RECEIVED_EXPORT)]
        private IWebElement _topReceivedExport;

        [FindsBy(How = How.Id, Using = EXPORT_FOR_SAGE)]
        private IWebElement _exportForSage;

        [FindsBy(How = How.Id, Using = CONFIRM_EXPORT_SAGE)]
        private IWebElement _confirmExportSage;

        [FindsBy(How = How.Id, Using = CANCEL_EXPORT_SAGE)]
        private IWebElement _cancelDueToInvoice;

        [FindsBy(How = How.Id, Using = DOWNLOAD_FILE)]
        private IWebElement _downloadFile;

        // Tableau
        [FindsBy(How = How.XPath, Using = FIRST_RECEIPT_NOTE_NUMBER)]
        private IWebElement _firstReceiptNote;

        [FindsBy(How = How.XPath, Using = DELETE_FIRST_RECEIPT_NOTE)]
        private IWebElement _deleteFirstReceiptNote;

        [FindsBy(How = How.Id, Using = CONFIRM_DELETE)]
        private IWebElement _confirmDelete;

        [FindsBy(How = How.XPath, Using = TOTAL_RCN_NUMBER)]
        private IWebElement _total_recipeNote_Number;

        [FindsBy(How = How.ClassName, Using = "counter")]
        private IWebElement _totalNumber;

        //__________________________________ Filters _______________________________________

        [FindsBy(How = How.XPath, Using = RESET_FILTER_PATCH)]
        private IWebElement _resetFilterPatch;

        [FindsBy(How = How.Id, Using = FILTER_NUMBER)]
        private IWebElement _searchByNumber;

        [FindsBy(How = How.Id, Using = FILTER_PURCHASE_NUMBER)]
        private IWebElement _searchByPurchaseNumber;

        [FindsBy(How = How.Id, Using = FILTER_DELIVERY_NUMBER)]
        private IWebElement _searchByDeliveryNumber;

        [FindsBy(How = How.XPath, Using = UNSELECT_ALL_SUPPLIER)]
        private IWebElement _unselectAllSupplier;

        [FindsBy(How = How.Id, Using = FILTER_DATE_FROM)]
        public IWebElement _dateFrom;

        [FindsBy(How = How.Id, Using = FILTER_DATE_TO)]
        public IWebElement _dateTo;

        [FindsBy(How = How.Id, Using = FILTER_SHOW_SENT_TO_SAGE_AND_IN_ERROR_ONLY)]
        public IWebElement _showSentToSAGEAndInErrorOnly;

        [FindsBy(How = How.Id, Using = FILTER_SHOW_ALL_DEV)]
        private IWebElement _showAllReceiptsDev;

        [FindsBy(How = How.XPath, Using = FILTER_SHOW_ALL_PATCH)]
        private IWebElement _showAllReceiptsPatch;

        [FindsBy(How = How.Id, Using = FILTER_NOT_VALIDATED_DEV)]
        private IWebElement _showNotValidatedDev;

        [FindsBy(How = How.XPath, Using = FILTER_NOT_VALIDATED_PATCH)]
        private IWebElement _showNotValidatedPatch;

        [FindsBy(How = How.Id, Using = FILTER_VALIDATED_DEV)]
        private IWebElement _showValidatedOnlyDev;

        [FindsBy(How = How.XPath, Using = FILTER_VALIDATED_PATCH)]
        private IWebElement _showValidatedOnlyPatch;

        [FindsBy(How = How.Id, Using = FILTER_NOT_INVOICED)]
        private IWebElement _showValidatedNotInvoiced;

        [FindsBy(How = How.Id, Using = FILTER_PARTIALLY_INVOICED_DEV)]
        private IWebElement _showValidatedPartialInvoicedDev;

        [FindsBy(How = How.XPath, Using = FILTER_PARTIALLY_INVOICED_PATCH)]
        private IWebElement _showValidatedPartialInvoicedPatch;

        [FindsBy(How = How.Id, Using = FILTER_WITH_CLAIM_DEV)]
        private IWebElement _showWithClaimDev;

        [FindsBy(How = How.XPath, Using = FILTER_WITH_CLAIM_PATCH)]
        private IWebElement _showWithClaimPatch;

        [FindsBy(How = How.Id, Using = FILTER_EXPORTED_MANUALLY_DEV)]
        private IWebElement _exportedForSageManuallyDev;

        [FindsBy(How = How.XPath, Using = FILTER_EXPORTED_MANUALLY_PATCH)]
        private IWebElement _exportedForSageManuallyPatch;

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

        [FindsBy(How = How.XPath, Using = STATUS)]
        private IWebElement _statusLink;

        [FindsBy(How = How.Id, Using = FILTER_SITE)]
        private IWebElement _siteFilter;

        [FindsBy(How = How.XPath, Using = UNSELECT_ALL_SITES)]
        private IWebElement _unselectAllSites;

        [FindsBy(How = How.XPath, Using = SEARCH_SITE)]
        private IWebElement _searchSite;

        [FindsBy(How = How.Id, Using = FILTER_SITE_PLACES)]
        private IWebElement _sitePlacesFilter;

        [FindsBy(How = How.XPath, Using = UNSELECT_ALL_SITE_PLACES)]
        private IWebElement _unselectAllSitePlaces;

        [FindsBy(How = How.XPath, Using = SEARCH_SITE_PLACES)]
        private IWebElement _searchSitePlces;

        [FindsBy(How = How.XPath, Using = EXPORTED_WITH_REVERSE_GENERATED)]
        private IWebElement _exportedWithReverseGenerated;

        [FindsBy(How = How.XPath, Using = FILTER_EXPORTED_WITH_PROVISION_GENERATED)]
        private IWebElement _exportedWithProvisionGenerated;

        [FindsBy(How = How.Id, Using = Exported_ForSage_Or_CegidAutomatically)]
        private IWebElement _exportedForSageOrCegidAutomatically;

        [FindsBy(How = How.XPath, Using = PRICE_INDEX)]
        private IWebElement _itemIndexPrice;

        [FindsBy(How = How.XPath, Using = TOTAL_VAT)]
        private IWebElement _totalVat;

        [FindsBy(How = How.XPath, Using = DELIVERY_LOCATION)]
        private IWebElement _deliveryLocation;

        public enum FilterType
        {
            ByNumber,
            ByPurchaseNumber,
            ByDeliveryNumber,
            Supplier,
            DateFrom,
            DateTo,
            ShowAllReceipts,
            ShowNotValidated,
            ShowValidatedOnly,
            ShowValidatedNotInvoiced,
            ShowValidatedPartialInvoiced,
            ShowWithClaim,
            ShowExportedForSageManually,
            ShowSentToSAGEAndInErrorOnly,
            ShowValidatedAndNotSentToSage,
            ShowAll,
            ShowActive,
            ShowInactive,
            SortBy,
            All,
            Opened,
            Closed,
            Site,
            SitePlaces,
            ExportedWithReverseGenerated,
            ExportedWithProvisionGenerated,
            ExportedForSageOrCegidAutomatically
        }

        public void Filter(FilterType filterType, object value)
        {
            Actions action = new Actions(_webDriver);
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;

            switch (filterType)
            {
                case FilterType.ByNumber:
                    _searchByNumber = WaitForElementIsVisible(By.Id(FILTER_NUMBER));
                    _searchByNumber.SetValue(ControlType.TextBox, value);
                    break;
                case FilterType.ByPurchaseNumber:
                    _searchByPurchaseNumber = WaitForElementIsVisible(By.Id(FILTER_PURCHASE_NUMBER));
                    _searchByPurchaseNumber.SetValue(ControlType.TextBox, value);
                    break;
                case FilterType.ByDeliveryNumber:
                    _searchByDeliveryNumber = WaitForElementIsVisible(By.Id(FILTER_DELIVERY_NUMBER));
                    _searchByDeliveryNumber.SetValue(ControlType.TextBox, value);
                    break;
                case FilterType.Supplier:
                    ComboBoxSelectById(new ComboBoxOptions(FILTER_SUPPLIER, (string)value));
                    break;
                case FilterType.DateFrom:
                    var _RestFilterLink = WaitForElementExists(By.Id("ResetFilter"));
                    action.MoveToElement(_RestFilterLink).Perform();
                    _dateFrom = WaitForElementIsVisible(By.Id(FILTER_DATE_FROM));
                    _dateFrom.SetValue(ControlType.DateTime, value);
                    var clickDateFrom = WaitForElementIsVisible(By.XPath("//*/td[@class='active day']"));
                    clickDateFrom.Click();
                    break;
                case FilterType.DateTo:
                    var _ResetFilterLink = WaitForElementExists(By.Id("ResetFilter"));
                    action.MoveToElement(_ResetFilterLink).Perform();
                    _dateTo = WaitForElementIsVisible(By.Id(FILTER_DATE_TO));
                    _dateTo.SetValue(ControlType.DateTime, value);
                    var clickDateTo = WaitForElementIsVisible(By.XPath("//*/td[@class='active day']"));
                    clickDateTo.Click();
                    break;
                case FilterType.ShowAllReceipts:
                    _showAllReceiptsDev = WaitForElementExists(By.Id(FILTER_SHOW_ALL_DEV));
                    action.MoveToElement(_showAllReceiptsDev).Perform();
                    _showAllReceiptsDev.SetValue(ControlType.RadioButton, value);
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
                case FilterType.ShowValidatedNotInvoiced:
                    _showValidatedNotInvoiced = WaitForElementExists(By.Id(FILTER_NOT_INVOICED));
                    action.MoveToElement(_showValidatedNotInvoiced).Perform();
                    _showValidatedNotInvoiced.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.ShowValidatedPartialInvoiced:
                    _showValidatedPartialInvoicedDev = WaitForElementExists(By.Id(FILTER_PARTIALLY_INVOICED_DEV));
                    action.MoveToElement(_showValidatedPartialInvoicedDev).Perform();
                    _showValidatedPartialInvoicedDev.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.ShowExportedForSageManually:
                    _exportedForSageManuallyPatch = WaitForElementExists(By.Id("ExportedForSageManually"));
                    action.MoveToElement(_exportedForSageManuallyPatch).Perform();
                    _exportedForSageManuallyPatch = WaitForElementExists(By.Id("ExportedForSageManually"));
                    _exportedForSageManuallyPatch.SendKeys(Keys.ArrowDown);
                    _exportedForSageManuallyPatch.SendKeys(Keys.ArrowUp);
                    _exportedForSageManuallyPatch.SendKeys(Keys.ArrowUp);
                    _exportedForSageManuallyPatch.SendKeys(Keys.ArrowDown);
                    _exportedForSageManuallyPatch.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.ShowSentToSAGEAndInErrorOnly:

                    _showSentToSAGEAndInErrorOnly = WaitForElementExists(By.Id(FILTER_SHOW_SENT_TO_SAGE_AND_IN_ERROR_ONLY));
                    action.MoveToElement(_showSentToSAGEAndInErrorOnly).Perform();
                    _showSentToSAGEAndInErrorOnly.SendKeys(Keys.ArrowDown);
                    _showSentToSAGEAndInErrorOnly.SendKeys(Keys.ArrowUp);
                    _showSentToSAGEAndInErrorOnly.SendKeys(Keys.ArrowUp);
                    _showSentToSAGEAndInErrorOnly.SendKeys(Keys.ArrowDown);
                    _showSentToSAGEAndInErrorOnly.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.ExportedWithReverseGenerated:
                    ScrollUntilElementIsVisible(By.XPath(EXPORTED_WITH_REVERSE_GENERATED));
                    _exportedWithReverseGenerated = WaitForElementExists(By.XPath(EXPORTED_WITH_REVERSE_GENERATED));
                    action.MoveToElement(_exportedWithReverseGenerated).Perform();
                    _exportedWithReverseGenerated.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.ShowValidatedAndNotSentToSage:
                    var showValidatedNoSage = WaitForElementExists(By.Id(FILTER_SHOW_VALIDATED_NO_SAGE));
                    javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", showValidatedNoSage);
                    showValidatedNoSage.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.ShowWithClaim:
                    _exportedForSageManuallyDev = WaitForElementExists(By.Id(FILTER_EXPORTED_MANUALLY_DEV));
                    action.MoveToElement(_exportedForSageManuallyDev).Perform();
                    _showWithClaimDev = WaitForElementExists(By.Id(FILTER_WITH_CLAIM_DEV));
                    action.MoveToElement(_showWithClaimDev).Perform();
                    WaitForLoad();
                    _showWithClaimDev.SetValue(ControlType.RadioButton, value);
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
                    javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", _showNotActiveDev);
                    _showNotActiveDev.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.SortBy:
                    _sortBy = WaitForElementExists(By.Id(FILTER_SORTBY));
                    action.MoveToElement(_sortBy).Perform();
                    _sortBy.SendKeys(Keys.Tab);
                    _sortBy.Click();
                    var element = WaitForElementIsVisible(By.XPath("//option[contains(@value,'" + value + "')]"));
                    _sortBy.SetValue(ControlType.DropDownList, element.Text);
                    _sortBy.Click();
                    break;
                case FilterType.All:
                    _statusAllDev = WaitForElementExists(By.Id(FILTER_SHOW_ALL_STATUS_DEV));
                    action.MoveToElement(_statusAllDev).Perform();
                    _statusAllDev.SendKeys(Keys.Tab);
                    _statusAllDev.SetValue(ControlType.RadioButton, value);
                    _statusLink.Click();
                    _statusLink.SendKeys(Keys.PageUp);
                    break;
                case FilterType.Opened:
                    _statusOpenedDev = WaitForElementExists(By.Id(FILTER_OPENED_DEV));
                    action.MoveToElement(_statusOpenedDev).Perform();
                    _statusOpenedDev.SendKeys(Keys.Tab);
                    _statusOpenedDev.SetValue(ControlType.RadioButton, value);
                    WaitForLoad();
                    _statusLink.Click();
                    _statusLink.SendKeys(Keys.PageUp);
                    break;
                case FilterType.Closed:
                    _statusClosedDev = WaitForElementExists(By.Id(FILTER_CLOSED_DEV));
                    action.MoveToElement(_statusClosedDev).Perform();
                    _statusClosedDev.SendKeys(Keys.Tab);
                    _statusClosedDev.SetValue(ControlType.RadioButton, value);
                    WaitForLoad();
                    _statusLink.Click();
                    _statusLink.SendKeys(Keys.PageUp);
                    break;
                case FilterType.Site:
                    _siteFilter = WaitForElementExists(By.Id(FILTER_SITE));
                    action.MoveToElement(_siteFilter).Perform();
                    _siteFilter.SendKeys(Keys.Tab);
                    _siteFilter.Click();

                    _unselectAllSites = WaitForElementIsVisible(By.XPath(UNSELECT_ALL_SITES));
                    _unselectAllSites.Click();

                    _searchSite = WaitForElementIsVisible(By.XPath(SEARCH_SITE));
                    _searchSite.SetValue(ControlType.TextBox, value);

                    var searchSiteValue = WaitForElementIsVisible(By.XPath("//span[text()='" + value + " - " + value + "']"));
                    searchSiteValue.Click();

                    _siteFilter.Click();
                    break;
                case FilterType.SitePlaces:
                    _sitePlacesFilter = WaitForElementExists(By.Id(FILTER_SITE_PLACES));
                    action.MoveToElement(_sitePlacesFilter).Perform();
                    _sitePlacesFilter.SendKeys(Keys.Down);

                    _unselectAllSitePlaces = WaitForElementIsVisible(By.XPath(UNSELECT_ALL_SITE_PLACES));
                    _unselectAllSitePlaces.Click();

                    _searchSitePlces = WaitForElementIsVisible(By.XPath(SEARCH_SITE_PLACES));
                    _searchSitePlces.SetValue(ControlType.TextBox, value);

                    var searchSitePlaceValue = WaitForElementIsVisible(By.XPath("//span[contains(text(),'" + value + "')]"));
                    searchSitePlaceValue.Click();

                    break;
                case FilterType.ExportedWithProvisionGenerated:
                    _exportedWithProvisionGenerated = WaitForElementExists(By.XPath(FILTER_EXPORTED_WITH_PROVISION_GENERATED));
                    _exportedWithProvisionGenerated.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.ExportedForSageOrCegidAutomatically:             
                    _exportedForSageOrCegidAutomatically = WaitForElementExists(By.Id(Exported_ForSage_Or_CegidAutomatically));
                    action.MoveToElement(_exportedForSageOrCegidAutomatically).Perform();
                    _exportedForSageOrCegidAutomatically = WaitForElementExists(By.Id(Exported_ForSage_Or_CegidAutomatically));
                    _exportedForSageOrCegidAutomatically.SendKeys(Keys.ArrowDown);
                    _exportedForSageOrCegidAutomatically.SendKeys(Keys.ArrowUp);
                    _exportedForSageOrCegidAutomatically.SendKeys(Keys.ArrowUp);
                    _exportedForSageOrCegidAutomatically.SendKeys(Keys.ArrowDown);
                    _exportedForSageOrCegidAutomatically.SetValue(ControlType.RadioButton, value);
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);

            }
            WaitPageLoading();
            WaitForLoad();
        }

        public void ResetFilter()
        {
            WaitPageLoading();
            WaitForLoad();
            Actions action = new Actions(_webDriver);
            _resetFilterPatch = WaitForElementExists(By.XPath(RESET_FILTER_PATCH));
            //(_webDriver as IJavaScriptExecutor).ExecuteScript("arguments[0].scrollIntoView(true);", _resetFilterPatch);
            action.MoveToElement(_resetFilterPatch).Perform();
            _resetFilterPatch.SendKeys(Keys.Tab);
            _resetFilterPatch.Click();
            WaitPageLoading();

            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                // DateTo à dans 2 mois
            }
        }

        /// <summary>
        /// Checks if all lines are in the correct state of vaildation process.
        /// </summary>
        /// <param name="checkIfValidated">The expected state of validation process.</param>
        /// <returns>True if every showed result if correct to the expected state, false otherwise.</returns>
        public bool CheckValidation(bool checkIfValidated)
        {
            int tot = CheckTotalNumber() > 100 ? 100 : CheckTotalNumber();

            if (tot == 0)
            { return false; }

            for (int i = 0; i < tot; i++)
            {
                if (isElementVisible(By.XPath(string.Format("//*[@id=\"list-item-with-action\"]/table/tbody/tr[{0}]/td[1]/div/i[contains(@class,'circle')]", i + 1))))
                {
                    _webDriver.FindElement(By.XPath(string.Format("//*[@id=\"list-item-with-action\"]/table/tbody/tr[{0}]/td[1]/div/i[contains(@class,'circle')]", i + 1)));

                    if (checkIfValidated == false)
                    { return false; }

                }
                else if (checkIfValidated)
                { return false; }
            }

            return true;
        }

        public bool CheckStatus(bool active)
        {
            bool isActive = false;
            int tot = CheckTotalNumber() > 100 ? 100 : CheckTotalNumber();

            if (tot == 0)
                return false;

            for (int i = 0; i < tot; i++)
            {
                if (isElementVisible(By.XPath(string.Format("//*[@id='list-item-with-action']/table/tbody/tr[{0}]/td[1]/div/i", i + 1))))
                {
                    //un icone
                    //vérif si inactive
                    var inactiveIcon = _webDriver.FindElement(By.XPath(string.Format("//*[@id='list-item-with-action']/table/tbody/tr[{0}]/td[1]/div/i", i + 1)));
                    if (inactiveIcon.GetAttribute("class").Contains("circle-xmark"))
                    {
                        if (active == true)
                        {
                            return true;
                        }
                        isActive = false;
                    }
                    else
                    {
                        if (active == false)
                        {
                            return false;
                        }
                        isActive = true;
                    }
                }
            }
            return isActive;
        }

        public bool IsSentToSageManually()
        {
            int tot = CheckTotalNumber() > 100 ? 100 : CheckTotalNumber();

            if (tot == 0)
                return false;

            for (int i = 0; i < tot; i++)
            {
                try
                {

                    if (isElementVisible(By.XPath(string.Format(ELEMENT_DEV, i + 1))))
                    {
                        // pouce vert si enabled Sage
                        // <i class="fa fa-thumbs-up" style="font-size:15px; color: limegreen; vertical-align: middle" title="Accounted" (sent="" to="" sage="" manually)=""></i>
                        var elm = _webDriver.FindElement(By.XPath(string.Format(ELEMENT_DEV, i + 1)));
                        //a modifier des que pacth est a jour
                        if (!elm.GetAttribute("title").Contains("sent to SAGE manually") && !elm.GetAttribute("title").Contains("Accounted"))
                            return false;
                    }
                    else if (isElementVisible(By.XPath(string.Format("//tr[{0}]//span[@class='glyphicon glyphicon-floppy-saved green-text' and @title='Accounted (sent to SAGE manually)']", i + 1))))
                    {
                        //disquette si enabled Cegid
                        //<span class="glyphicon glyphicon-floppy-saved green-text" title="Accounted (sent to SAGE manually)"></span>
                        var elm = _webDriver.FindElement(By.XPath(string.Format("//tr[{0}]//span[@class='glyphicon glyphicon-floppy-saved green-text' and @title='Accounted (sent to SAGE manually)']", i + 1)));
                        //a modifier des que pacth est a jour
                        if (!elm.GetAttribute("title").Contains("sent to SAGE manually") && !elm.GetAttribute("title").Contains("Accounted"))
                            return false;
                    }

                    else if (isElementVisible(By.XPath(string.Format(ELEMENT_PATCH, i + 1))))
                    {
                        // pouce vert si enabled Sage
                        // <i class="fa fa-thumbs-up" style="font-size:15px; color: limegreen; vertical-align: middle" title="Accounted" (sent="" to="" sage="" manually)=""></i>
                        var elm = _webDriver.FindElement(By.XPath(string.Format(ELEMENT_PATCH, i + 1)));
                        //a modifier des que pacth est a jour
                        if (!elm.GetAttribute("title").Contains("sent to SAGE manually") && !elm.GetAttribute("title").Contains("Accounted"))
                            return false;
                    }
                    else
                    {
                        //disquette si enabled Cegid
                        //<span class="glyphicon glyphicon-floppy-saved green-text" title="Accounted (sent to SAGE manually)"></span>
                        var elm = _webDriver.FindElement(By.XPath(string.Format("//tr[{0}]//span[@class='glyphicon glyphicon-floppy-saved green-text' and @title='Accounted (sent to SAGE manually)']", i + 1)));
                        //a modifier des que pacth est a jour
                        if (!elm.GetAttribute("title").Contains("sent to SAGE manually") && !elm.GetAttribute("title").Contains("Accounted"))
                            return false;
                    }
                }
                catch
                {
                    return false;
                }
            }
            return true;
        }

        public bool AreAllRnNotInvoiced()
        {
            ReadOnlyCollection<IWebElement> invoiceStatus;

            invoiceStatus = _webDriver.FindElements(By.XPath(INVOICED_ICONE));

            if (invoiceStatus.Count == 0)
            { return false; }

            foreach (var elm in invoiceStatus)
            {
                if (!elm.GetAttribute("title").Equals("Not invoiced"))
                {
                    return false;
                }
            }
            return true;
        }

        public bool AreAllRnWithCegidIcon()
        {
            ReadOnlyCollection<IWebElement> invoiceStatus;

            invoiceStatus = _webDriver.FindElements(By.XPath(CEGID_ICON));

            if (invoiceStatus.Count == 0)
            { return false; }

            foreach (var elm in invoiceStatus)
            {
                string cls = elm.GetAttribute("class");
                if (cls.Contains("fa-save") == false || cls.Contains("fa-trash-alt") == false)
                {
                    return false;
                }
            }
            return true;
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
            var listeNumbers = _webDriver.FindElements(By.XPath(RN_NUMBER));

            if (listeNumbers.Count == 0)
                return false;

            var ancientText = listeNumbers[0].Text;

            if (ancientText.EndsWith(" - Cancelled"))
            {
                ancientText = ancientText.Substring(0, ancientText.LastIndexOf(" - Cancelled"));
            }
            var ancientNumber = int.Parse(ancientText);

            foreach (var elm in listeNumbers)
            {
                try
                {
                    var newNumber = elm.Text;
                    if (newNumber.EndsWith(" - Cancelled"))
                    {
                        newNumber = newNumber.Substring(0, newNumber.LastIndexOf(" - Cancelled"));
                    }
                    if (int.Parse(newNumber) > ancientNumber)
                        return false;

                    ancientNumber = int.Parse(newNumber);
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
            var sites = _webDriver.FindElements(By.XPath(RN_SITE));

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
            var suppliers = _webDriver.FindElements(By.XPath(RN_SUPPLIER));

            if (suppliers.Count == 0)
            {
                Console.WriteLine("supplier " + supplier + " non trouvé");
                return false;
            }

            foreach (var elm in suppliers)
            {
                if (elm.Text != supplier)
                {
                    Console.WriteLine(elm.Text + " n'est pas un supplier " + supplier);
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
                if (isElementVisible(By.XPath(string.Format("//*[@id=\"list-item-with-action\"]/table/tbody/tr[{0}]/td[6]/div/div/div[2]/span", i + 1))))
                {
                    _webDriver.FindElement(By.XPath(string.Format("//*[@id=\"list-item-with-action\"]/table/tbody/tr[{0}]/td[6]/div/div/div[2]/span", i + 1)));
                }
                else
                {
                    return false;
                }

            }
            return true;
        }

        public bool IsWithClaim()
        {

            int tot = CheckTotalNumber() > 100 ? 100 : CheckTotalNumber();

            if (tot == 0)
                return false;

            for (int i = 0; i < tot; i++)
            {
                if (isElementVisible(By.XPath(string.Format("//*[@id=\"list-item-with-action\"]/table/tbody/tr[{0}]/td[6]/div/div/div[2]/span[2 and @title=\"Related Claims\"]", i + 1))))
                {
                    _webDriver.FindElement(By.XPath(string.Format("//*[@id=\"list-item-with-action\"]/table/tbody/tr[{0}]/td[6]/div/div/div[2]/span[2 and @title=\"Related Claims\"]", i + 1)));
                }
                else
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

            IReadOnlyCollection<IWebElement> dates;

            dates = _webDriver.FindElements(By.XPath("//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[8]"));

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

            IReadOnlyCollection<IWebElement> dates;

            dates = _webDriver.FindElements(By.XPath(LIST_DATES));

            if (dates.Count == 0)
                return false;

            var ancientDate = DateTime.Parse(dates.ElementAt(0).Text, ci);

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

        //___________________________________ Méthodes ____________________________________________

        // General
        public ReceiptNotesCreateModalPage ReceiptNotesCreatePage()
        {
            ShowPlusMenu();
            WaitForLoad();
            _createNewReceiptNote = WaitForElementIsVisibleNew(By.XPath(NEW_RECEIPT_NOTES));
            _createNewReceiptNote.Click();
            WaitForLoad();

            return new ReceiptNotesCreateModalPage(_webDriver, _testContext);
        }

        //public override void ShowExtendedMenu()
        //{
        //    var actions = new Actions(_webDriver);

        //    _showExtendedMenu = WaitForElementIsVisible(By.XPath(EXTENDED_MENU));
        //    actions.MoveToElement(_showExtendedMenu).Perform();
        //    WaitForLoad();
        //}


        public override void ShowExtendedMenu()
        {
            int maxRetries = 5;
            var actions = new Actions(_webDriver);

            for (int i = 0; i < maxRetries; i++)
            {
                _showExtendedMenu = WaitForElementIsVisible(By.XPath(EXTENDED_MENU));
                actions.MoveToElement(_showExtendedMenu).Perform();
                WaitForLoad();

                if (isElementVisible(By.Id(EXPORT)))
                {
                    break;
                }

                WaitForLoad();
            }
        }

        public PrintReportPage PrintReceiptNoteResults(bool printValue)
        {
            ShowExtendedMenu();

            _printReceiptNoteResults = WaitForElementIsVisible(By.XPath(PRINT_RN_RESULTS));
            _printReceiptNoteResults.Click();
            WaitForLoad();

            _confirmPrint = WaitForElementIsVisible(By.Id(CONFIRM_PRINT));
            _confirmPrint.Click();
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

        public void ExportExcelFile(bool printValue)
        {

            ShowExtendedMenu();

            _export = WaitForElementIsVisible(By.Id(EXPORT));
            _export.Click();
            WaitForLoad();

            if (printValue)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));
                ClickPrintButton();
            }

            WaitForDownload();
            //Close();
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


            Regex r = new Regex("^receiptnotes" + space + annee + "-" + mois + "-" + jour + space + heure + "-" + minutes + "-" + secondes + ".xlsx$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = r.Match(filePath);

            return m.Success;
        }

        public void ExportMatchingFile(bool printValue, string downloadPath)
        {
            ShowExtendedMenu();

            _exportMatching = WaitForElementIsVisible(By.Id(EXPORT_MATCHING));
            _exportMatching.Click();
            WaitForLoad();

            if (printValue)
            {
                ClosePrintButton();
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));
                ClickPrintButton();
            }

            WaitForDownload();
            Close();
        }

        public FileInfo GetExportMatchingFile(FileInfo[] taskFiles)
        {
            List<FileInfo> correctDownloadFiles = new List<FileInfo>();

            foreach (var file in taskFiles)
            {
                //  Test REGEX
                if (IsMatchingFileCorrect(file.Name))
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

        public bool IsMatchingFileCorrect(string filePath)
        {
            Regex r = new Regex("^ExportMatchingRN" + ".xlsx$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = r.Match(filePath);

            return m.Success;
        }

        public void ExportTopReceivedExcelFile(bool printValue, string downloadPath, string site, string nbTopReceived)
        {
            ShowExtendedMenu();

            _topReceived = WaitForElementIsVisible(By.XPath(TOP_RECEIVED));
            _topReceived.Click();
            WaitForLoad();

            //Start Date Selection
            _topReceivedStartDate = WaitForElementIsVisible(By.Id(TOP_RECEIVED_START_DATE));
            WaitForLoad();
            _topReceivedStartDate.SetValue(ControlType.DateTime, DateUtils.Now.AddMonths(-3));
            _topReceivedStartDate.SendKeys(Keys.Enter);
            WaitForLoad();

            //Site Selection
            _topReceivedSites = WaitForElementIsVisible(By.Id(TOP_RECEIVED_SITES));
            _topReceivedSites.Click();

            _topReceivedSitesUcheckAll = WaitForElementIsVisible(By.XPath(TOP_RECEIVED_SITES_UNCHECK_ALL));
            _topReceivedSitesUcheckAll.Click();

            _topReceivedSiteSearch = WaitForElementIsVisible(By.XPath(TOP_RECEIVED_SITE_SEARCH));
            _topReceivedSiteSearch.SetValue(ControlType.TextBox, site);

            _topReceivedSiteToCheck = WaitForElementIsVisible(By.XPath(string.Format(TOP_RECEIVED_SITE_TO_CHECK, site)));
            _topReceivedSiteToCheck.SetValue(ControlType.CheckBox, true);

            _topReceivedSites = WaitForElementIsVisible(By.Id(TOP_RECEIVED_SITES));
            _topReceivedSites.Click();

            //Count
            _topReceivedCount = WaitForElementIsVisible(By.Id(TOP_RECEIVED_COUNT));
            _topReceivedCount.SetValue(ControlType.DropDownList, nbTopReceived);

            //Export button
            _topReceivedExport = WaitForElementIsVisible(By.Id(TOP_RECEIVED_EXPORT));
            _topReceivedExport.Click();
            WaitForLoad();

            if (printValue)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));
                ClickPrintButton();
            }

            WaitForDownload();
            Close();
        }

        public FileInfo GetTopReceivedExcelFile(FileInfo[] taskFiles)
        {
            List<FileInfo> correctDownloadFiles = new List<FileInfo>();

            foreach (var file in taskFiles)
            {
                //  Test REGEX
                if (IsTopReceivedFileCorrect(file.Name))
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

        public bool IsTopReceivedFileCorrect(string filePath)
        {

            string numero = "(\\s\\(\\d\\)){0,1}";  // mois

            Regex r = new Regex("^topreceived" + numero + ".xlsx$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = r.Match(filePath);

            return m.Success;
        }

        public void ExportResultsForSage(bool printValue)
        {
            //export = generate txt
            ShowExtendedMenu();

            if (isElementVisible(By.Id(EXPORT_FOR_SAGE)))
            {
                _exportForSage = WaitForElementIsVisible(By.Id(EXPORT_FOR_SAGE));
            }
            else
            {
                _exportForSage = WaitForElementIsVisible(By.Id(GENERATE_SAGE_TXT));
            }
            _exportForSage.Click();

            _confirmExportSage = WaitForElementIsVisible(By.Id(CONFIRM_EXPORT_SAGE));
            _confirmExportSage.Click();
            WaitForLoad();

            if (!printValue)
            {
                try
                {
                    _downloadFile = _webDriver.FindElement(By.Id(DOWNLOAD_FILE));
                    _downloadFile.Click();
                    WaitForLoad();

                    // On ferme la pop-up
                    _cancelDueToInvoice = WaitForElementIsVisible(By.Id(CANCEL_EXPORT_SAGE));
                    _cancelDueToInvoice.Click();
                    WaitForLoad();
                }
                catch
                {
                    _cancelDueToInvoice = WaitForElementIsVisible(By.Id(CANCEL_EXPORT_SAGE));
                    _cancelDueToInvoice.Click();
                    WaitForLoad();

                    return;
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

        // Tableau
        public string GetFirstRecipeNoteNumber()
        {
            IWebElement firstRecipeNoteNumber = WaitForElementIsVisible(By.XPath("//*[@id=\"list-item-with-action\"]/table/tbody/tr[1]/td[3]"));
            return firstRecipeNoteNumber.Text;
        }
        public ReceiptNotesItem SelectFirstReceiptNoteItem()
        {
            _firstReceiptNote = WaitForElementIsVisible(By.XPath(FIRST_RECEIPT_NOTE_NUMBER));
            _firstReceiptNote.Click();
            WaitForLoad();

            return new ReceiptNotesItem(_webDriver, _testContext);
        }

        public string GetFirstReceiptNoteNumber()
        {
            _firstReceiptNote = WaitForElementExists(By.XPath(FIRST_RECEIPT_NOTE_NUMBER));
            return _firstReceiptNote.Text;
        }

        public void DeleteReceiptNote()
        {
            WaitPageLoading();
            WaitForLoad();
            _deleteFirstReceiptNote = WaitForElementIsVisible(By.XPath(DELETE_RECEIPT_NOTE));
            WaitForLoad();
            _deleteFirstReceiptNote.Click();
            WaitForLoad();

            _confirmDelete = WaitForElementIsVisible(By.Id(CONFIRM_DELETE));
            _confirmDelete.Click();
            WaitForLoad();
        }

        public bool IsFirstLineValide()
        {
            var _imgValide = _webDriver.FindElements(By.XPath(IMG_VALID));
            return _imgValide.Count > 0;
        }

        private void ClickPage100()
        {
            var pageingSelect = WaitForElementIsVisible(By.XPath(PAGEING_SELECT));
            pageingSelect.SetValue(ControlType.DropDownList, "100");

            WaitForLoad();
        }

        public void SelectRnNumber(string RNNumber)
        {
            var navigator = _webDriver.FindElements(By.XPath(CLOSE_PAGING));
            for (var i = 2; i < navigator.Count + 1; i++)
            {
                WaitForLoad();
                var navigatorItem = WaitForElementIsVisible(By.XPath(string.Format("//*[@id='close-receipt-notes']/div[1]/div[3]/nav/ul/li[{0}]/a", i)));
                //navigatorItem.Click();
                IJavaScriptExecutor executor = (IJavaScriptExecutor)_webDriver;
                executor.ExecuteScript("arguments[0].click();", navigatorItem);
                WaitForLoad();
                Thread.Sleep(500);

                if (isElementVisible(By.XPath(string.Format("//*[@id='RNTableToClose']/tr[contains(td[5], '{0}')]/td[1]/div", RNNumber))))
                {
                    var RNToClose = WaitForElementIsVisible(By.XPath(string.Format("//*[@id='RNTableToClose']/tr[contains(td[5], '{0}')]/td[1]/div", RNNumber)));
                    Actions action = new Actions(_webDriver);
                    action.MoveToElement(RNToClose).Perform();
                    WaitForLoad();
                    RNToClose.SetValue(ControlType.CheckBox, true);
                    WaitForLoad();
                    break;
                }
            }
        }
        public void CloseRNNumber(string RNNumber)
        {
            ShowExtendedMenu();
            _menuClose = WaitForElementIsVisible(By.XPath(MENU_CLOSE));
            _menuClose.Click();
            WaitForLoad();

            ClickPage100();

            SelectRnNumber(RNNumber);
            WaitForLoad();
            var validate = WaitForElementIsVisible(By.XPath(CLOSE_VALIDATE));
            validate.Click();
            WaitForLoad();
        }

        public Dictionary<string, string> GetRNPrices(string currency)
        {
            Dictionary<string, string> dico = new Dictionary<string, string>();
            var lignesRN = _webDriver.FindElements(By.XPath(LIGNES_RN));
            foreach (var ligneRN in lignesRN)
            {
                var colonneRN = _webDriver.FindElement(By.XPath("//*/div[@id='list-item-with-action']/table/tbody/tr/td[3][contains(text(),'" + ligneRN.Text + "')]/../td[9]"));
                dico.Add(ligneRN.Text, colonneRN.Text.Replace(currency, "").Replace(" ", ""));
            }
            return dico;
        }

        public int GetTotalRowsForPagination()
        {
            var rows = _webDriver.FindElements(By.XPath(ROWS_FOR_PAGINATION));
            return rows.Count;
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

        public bool VerifyChangePage(int nbre)
        {
            var list1 = new List<string>(); 
            var list2 = new List<string>();

            for (var i = 1; i < nbre + 1; i++)
            {
                list1 = GetItemNames();
                WaitForLoad();
                var navigatorItem = WaitForElementIsVisible(By.XPath(string.Format("//*[@id=\"list-item-with-action\"]/nav/ul/li[*]/a[@data-pager-pageindex = {0}]", i)));
                //navigatorItem.Click();
                IJavaScriptExecutor executor = (IJavaScriptExecutor)_webDriver;
                executor.ExecuteScript("arguments[0].click();", navigatorItem);
                WaitForLoad();
                Thread.Sleep(500);
                list2 = GetItemNames();
                if (list1.SequenceEqual(list2))
                {
                    return false;
                }
            }
            return true;
        }

        public List<String> GetItemNames()
        {
            List<String> itemNames = new List<String>();

            var items = _webDriver.FindElements(By.XPath(ITEM_NAME)).Select(_ => _.Text).ToList();

            return items;
        }

        public bool CheckNotInvoicedIcones()
        {
            try
            {
                var icones = _webDriver.FindElements(By.XPath(INVOICED_ICONE));
                foreach (var icone in icones)
                {
                    string cls = icone.GetAttribute("class");
                    string css = icone.GetAttribute("style");
                    string title = icone.GetAttribute("title");
                    if (!cls.Contains("fa-dollar-sign") || !css.Contains("color: red") || !title.Contains("Not invoiced"))
                    {
                        return false;
                    }
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        public bool CheckPartiallyInvoicedIcones()
        {
            try
            {
                var icones = _webDriver.FindElements(By.XPath(INVOICED_ICONE));
                foreach (var icone in icones)
                {
                    string cls = icone.GetAttribute("class");
                    string css = icone.GetAttribute("style");
                    string title = icone.GetAttribute("title");
                    if (!cls.Contains("fa-dollar-sign") || !css.Contains("color: orange") || !title.Contains("Partially invoiced"))
                    {
                        return false;
                    }
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        public bool AreAllRnInvoiced()
        {
            try
            {
                var icones = _webDriver.FindElements(By.XPath(INVOICED_ICONE));
                foreach (var icone in icones)
                {
                    string cls = icone.GetAttribute("class");
                    string css = icone.GetAttribute("style");
                    string title = icone.GetAttribute("title");
                    if (cls.Contains("fa-dollar-sign") == false || css.Contains("color: green") == false || title.Contains("Invoiced") == false)
                    {
                        return false;
                    }
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        public bool AreAllRnSentToSage()
        {
            try
            {
                var icones = _webDriver.FindElements(By.XPath(CEGID_ICON));
                foreach (var icone in icones)
                {
                    string cls = icone.GetAttribute("class");
                    string css = icone.GetAttribute("style");
                    string title = icone.GetAttribute("title");
                    if (cls.Contains("fa-dollar-sign") == false || css.Contains("color: green") == false || title.Contains("Invoiced") == false)
                    {
                        return false;
                    }
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        public bool CheckRnWithClaimsIcones()
        {
            try
            {               
                var icones = _webDriver.FindElements(By.XPath(WITHCLAIM_ICONE));
                foreach (var icone in icones)
                {
                    string cls = icone.GetAttribute("class");
                    string css = icone.GetAttribute("style");
                    string title = icone.GetAttribute("title");
                    if (!cls.Contains("fas fa-bullhorn") || !css.Contains("color: red") || !title.Contains("Related Claims"))
                    {
                        return false; 
                    }
                }
                return true; 
            }
            catch (NoSuchElementException)
            {
                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
                return false;
            }
        }

        public bool AreAllRnValidateAndNoSage()
        {
            try
            {
                bool allValidated = this.CheckValidation(true);
                bool cegidIcon = this.AreAllRnWithCegidIcon();

                return allValidated && (cegidIcon == false);
            }
            catch
            {
                return false;
            }
        }

        public int GetNumberOfShowedResults()
        {
            try
            {
                return _webDriver.FindElements(By.XPath(ALL_RN_ROWS)).Count;
            }
            catch
            {
                return 0;
            }
        }
        
        public int GetNumberOfPageResults()
        {
            try
            {
                var pageNumber = _webDriver.FindElement(By.XPath("(//*[@id='list-item-with-action']/nav/ul[@class='pagination']/li/a[not(contains(text(),'>')) and not(contains(text(),'<')) and not(contains(text(),'...'))])[last()]"));

                return int.Parse(pageNumber.Text);
            }
            catch
            {
                return 0;
            }
        }

        public void GetPageResults(int pagenumber)
        {
            var page = WaitForElementIsVisible(By.XPath("//*[@id='list-item-with-action']/nav/ul/li/a[text()='" + pagenumber + "']"));
            page.Click();
            WaitForLoad();
        }

        public double GetFirstReceiptNoteTotal()
        {
            var firstRnTotal = WaitForElementIsVisible(By.XPath(FIRST_RN_TOTAL)).Text;
            string cleanedText = firstRnTotal.Replace("€", "").Replace(",", ".").Trim();
            double value =0 ;
            double.TryParse(cleanedText, NumberStyles.Any, CultureInfo.InvariantCulture, out value);
            return value;
        }

        public ReceiptNotesItem ClickFirstReceiptNote()
        {
            var firstRN = WaitForElementIsVisible(By.XPath(FIRST_RN));
            firstRN.Click();
            return new ReceiptNotesItem(_webDriver,_testContext);
        }

        public bool ResetFilterCheck()
        {
            Random random = new Random();
            var randomNumber = random.Next(10,99).ToString();
            //get default values
            var defaultSearchByNumberValue = WaitForElementExists(By.Id("SearchNumber")).GetAttribute("value");
            var defaultByPurchaseOrderNumberValue = WaitForElementExists(By.Id("SearchNumberSecondary")).GetAttribute("value");
            var defaultByDeliveryOrderNumberValue = WaitForElementExists(By.Id("SearchNumberThirdary")).GetAttribute("value");



            Filter(FilterType.ByNumber, randomNumber);
            Filter(FilterType.ByPurchaseNumber, randomNumber);
            Filter(FilterType.ByDeliveryNumber, randomNumber);
            

            ResetFilter();

            //get the values again to check
            var afterResetSearchByNumberValue = WaitForElementExists(By.Id("SearchNumber")).GetAttribute("value");
            var afterResetByPurchaseOrderNumberValue = WaitForElementExists(By.Id("SearchNumberSecondary")).GetAttribute("value");
            var afterResetByDeliveryOrderNumberValue = WaitForElementExists(By.Id("SearchNumberThirdary")).GetAttribute("value");



            if(defaultSearchByNumberValue == afterResetSearchByNumberValue && defaultByPurchaseOrderNumberValue== afterResetByPurchaseOrderNumberValue &&
                defaultByDeliveryOrderNumberValue == afterResetByDeliveryOrderNumberValue)
            {
                return true;
            }
            return false;

        }
        public void ScrollUntilElementIsVisible(By by)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)_webDriver;
            var element = WaitForElementExists(by);
            if (element != null)
            {
                js.ExecuteScript("arguments[0].scrollIntoView(true);", element);
            }
        }
        public string GetTotalRecipeNotes()
        {
            _total_recipeNote_Number = WaitForElementIsVisible(By.XPath(TOTAL_RCN_NUMBER));
            WaitForLoad();
            return _total_recipeNote_Number.Text;
        }

        public bool IsIconDisplayed()
        {
            var iconElement = _webDriver.FindElement(By.XPath(ICON));
            return iconElement.Displayed;
        }
        public double GetIndexPrice(string decimalSeparator, string currency)
        {

            WaitForLoad();
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            _itemIndexPrice = WaitForElementIsVisible(By.XPath(PRICE_INDEX));
            var sanctionAmountText = _itemIndexPrice.Text;

            var format = (NumberFormatInfo)ci.NumberFormat.Clone();
            format.CurrencySymbol = currency;
            var mynumber = Decimal.Parse(sanctionAmountText, NumberStyles.Currency, format);

            return Convert.ToDouble(mynumber, ci);
        }

        public string GetTotalVat()
        {
            _totalVat = WaitForElementIsVisible(By.XPath(TOTAL_VAT));
            return _totalVat.Text.Replace("€", "").Trim();
        }

        public string GetDeliveryLocation()
        {
            _deliveryLocation = WaitForElementIsVisible(By.XPath(DELIVERY_LOCATION));
            return _deliveryLocation.Text.Trim();
        }
        public override int CheckTotalNumber()
        {
            WaitForLoad();
            _totalNumber = WaitForElementExists(By.XPath("//*[@id=\"div-body\"]/div/div[2]/div[1]/h1/span"));
            int nombre = Int32.Parse(_totalNumber.GetAttribute("innerText"));
            return nombre;
        }
        public void ShowBtnStatus()
        {
            var btnStatus = WaitForElementIsVisible(By.XPath("//*[@id=\"formSearchReceiptNotes\"]/div[11]/a"));
            btnStatus.Click();
            WaitPageLoading();
            WaitForLoad();
        }
        
        public bool isDisquetteRouge()
        {
            var _iconDisquette = WaitForElementIsVisible(By.XPath("//*[@id=\"list-item-with-action\"]/table/tbody/tr[1]/td[1]/div[2]/i"));
            return _iconDisquette.GetAttribute("class").Contains("red-text");
        }
        public string isWarningDisquetteRouge()
        {
            var _iconWarningDisquette = WaitForElementIsVisible(By.XPath("//*[@id=\"list-item-with-action\"]/table/tbody/tr[1]/td[1]/div[2]/i"));
            return _iconWarningDisquette.GetAttribute("title");
        }

    }
}
