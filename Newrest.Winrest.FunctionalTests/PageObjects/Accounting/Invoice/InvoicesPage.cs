using DocumentFormat.OpenXml.Bibliography;
using iText.StyledXmlParser.Jsoup.Helper;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.CustomerPortal;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Accounting.Invoice
{
    public class InvoicesPage : PageBase
    {

        public InvoicesPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //__________________________________Constantes_______________________________________

        // General
        private const string PLUS_BUTTON = "Plus-Button";

        private const string NEW_MANUAL_INVOICE = "Invoice-Manual";
        private const string NEW_AUTO_INVOICE = "Invoice-Auto";
        private const string NEW_MANUAL_CREDIT_NOTE = "Invoice-ManualCreditNote";

        private const string EXTENDED_BUTTON = "Extended-Menu";

        private const string PRINT = "btn-print-invoices-report";
        private const string CONFIRM_PRINT_PDF = "btn-print-submit";
        private const string CONFIRM_PRINT_ZIP = "btn-save-files";
        private const string SEND_INVOICES_MAIL = "btn-send-all-by-email";
        private const string CONFIRM_SEND = "btn-init-async-send-mail";
        private const string CANCEL = "btn-cancel-popup";
        private const string EXPORT = "btn-export-excel-fast";
        private const string CUSTOM_REPORT = "//*[@id=\"tabContentItemContainer\"]/div[1]/div/div/div[1]/div/a[text()='Customs reports']";
        private const string CUSTOM_REPORT_SITE = "SelectedSite";
        private const string CONFIRM_CUSTOM_REPORT_PRINT = "//*[@id=\"item-print-form\"]/div/div[3]/button[2]";
        private const string ENABLE_EXPORT_FOR_SAGE = "//*[@id=\"tabContentItemContainer\"]/div[1]/div/div/div[1]/div/a[5]";
        private const string EXPORT_FOR_SAGE = "btn-export-for-sage";
        private const string CONFIRM_EXPORT_SAGE = "btn-popup-validate";
        private const string SEND_AUTO_TO_SAGE = "//*[@id=\"tabContentItemContainer\"]/div[1]/div/div/div[1]/div/a[contains(text(),'Send Auto to SAGE')]";
        private const string GENERATE_SAGE_TXT = "btn-export-for-sage-txt-generation";
        private const string CHECK_FIRST_COMMENT = "//*[@id='tableListMenu']/tbody/tr[1]/td[8]/div";
        private const string VALIDATE_RESULTS = "btn-validate-all";
        private const string VALIDATE_1 = "//*[@id=\"div-body\"]/div/div[1]/div/div[2]/button";
        private const string VALIDATE_3 = "//*[@id=\"dataConfirmOK\"]";
        private const string CONFIRM_VALIDATE_RESULTS = "btn-popup-validate-all";
        private const string DELETE_UNVALIDATED_RESULTS = "btn-delete-all";
        private const string CONFIRM_DELETE_UNVALIDATED_RESULTS = "btn-popup-delete-all";
        private const string CLOSE = "btn-cancel-popup";
        private const string MAIL_SEND = "/html/body/div[2]/div/div[2]/div/div[2]/table/tbody/tr/td[15]/span";
        private const string CUSTOM_REPORT_MONTH = "SelectedMonth";

        // Tableau invoices
        private const string FIRST_INVOICE_ID = "//tr[1]//a[contains(text(), 'Id')]";
        private const string FIRST_INVOICE_NUMBER = "//*[@id=\"accounting-orderinvoice-detailv2-invoiceno-1\"]/b";
        private const string FIRST_SITE = "//*[@id=\"tableListMenu\"]/tbody/tr[1]/td[5]";
        private const string FIRST_CUSTOMER = "//*[@id=\"tableListMenu\"]/tbody/tr[1]/td[6]";
        private const string FIRST_INVOICE_DELETE = "accounting-orderinvoice-detailv2-delete-1";
        private const string CONFIRM_DELETE = "dataConfirmOK";
        private const string CONFIRM_DELETE2 = "dataAlertCancel";
        private const string FIRST_VALID_BTN = "//*[@id=\"accounting-orderinvoice-detailv2-1\"]/td[1]/div/i";
        private const string SECOND_VALID_BTN = "//*[@id=\"accounting-orderinvoice-detailv2-2\"]/td[1]/div/i";
        private const string FIRST_DISQUETTE_BTN = "//*[@id=\"accounting-orderinvoice-detailv2-1\"]/td[2]/div";
        private const string SECOND_DISQUETTE_BTN = "//*[@id=\"accounting-orderinvoice-detailv2-2\"]/td[2]/div";
        private const string TABLE_ROWS = "/html/body/div[2]/div/div[2]/div/div[2]/table/tbody/tr";
        private const string FIRST_LINE = "/html/body/div[2]/div/div[2]/div/div[2]/table/tbody/tr";
        private const string PRINT_BTN = "header-print-button";
        private const string INVOICELINE = "//*[@id='tableListMenu']/tbody/tr[td[5]/a[contains(text(),'{0}')]]";
        //*[starts-with(@id,"accounting-orderinvoice-detailv2-delete")]



        // Filtres 
        private const string RESET_FILTER = "Reset-Filter";

        private const string FILTER_SEARCH = "SearchInvoiceNo";
        private const string SHOW_SITES = "//*[@id=\"item-filter-form\"]/div[3]/div/a";
        private const string FILTER_SITES = "SelectedSites_ms";
        private const string SEARCH_SITES = "/html/body/div[10]/div/div/label/input";
        private const string UNSELECT_ALL_SITES = "/html/body/div[10]/div/ul/li[2]/a";
        private const string SHOW_CUSTOMERS = "//*[@id=\"item-filter-form\"]/div[4]/div/a";
        private const string FILTER_CUSTOMERS = "SelectedCustomers_ms";
        private const string SEARCH_CUSTOMERS = "/html/body/div[11]/div/div/label/input";
        private const string UNSELECT_ALL_CUSTOMERS = "/html/body/div[11]/div/ul/li[2]/a";
        private const string FILTER_DATE_FROM = "start-date-picker";
        private const string FILTER_DATE_TO = "end-date-picker";

        private const string SHOW_INVOICE_STEP = "//*[@id=\"item-filter-form\"]/div[7]/a";
        private const string FILTER_INVOICE_STEP_ALL = "all";
        private const string FILTER_INVOICE_STEP_DRAFT = "draft";
        private const string FILTER_INVOICE_STEP_VALIDATED = "validated";
        private const string FILTER_INVOICE_STEP_ACCOUNTED = "accounted";

        private const string SHOW_INVOICE_SHOW = "//*[@id=\"item-filter-form\"]/div[8]/a";
        private const string FILTER_INVOICE_SHOW_ALL = "//*[@id=\"SelectedCreationType\"][@value='All']";
        private const string FILTER_INVOICE_SHOW_FLIGHT_INVOICES = "//*[@id=\"SelectedCreationType\"][@value='Flight']";
        private const string FILTER_INVOICE_SHOW_DELIVERY_INVOICES = "//*[@id=\"SelectedCreationType\"][@value='Delivery']";
        private const string FILTER_INVOICE_SHOW_CUSTOMER_INVOICES = "//*[@id=\"SelectedCreationType\"][@value='CustomerOrder']";
        private const string FILTER_INVOICE_SHOW_MANUAL_INVOICES = "//*[@id=\"SelectedCreationType\"][@value='Manual']";

        private const string FILTER_INVOICE_SHOW_ALL_INVOICES = "//*[@id=\"ShowOnlyValue\"][@value='All']";
        private const string FILTER_INVOICE_SHOW_NOT_SENT_MAIL = "//*[@id=\"ShowOnlyValue\"][@value='NotSentOnly']";
        private const string FILTER_INVOICE_SHOW_NOT_SENT_SAGE = "//*[@id=\"ShowOnlyValue\"][@value='NotInvoicedOnly']";
        private const string FILTER_INVOICE_SHOW_NOT_VALIDATED = "//*[@id=\"ShowOnlyValue\"][@value='NotValidatedOnly']";
        private const string FILTER_INVOICE_SHOW_NOT_SAGE_ERROR = "//*[@id=\"ShowOnlyValue\"][@value='InErrorOnly']";
        private const string FILTER_INVOICE_SHOW_ON_TL_WAITING_SAGE_PUSH = "//*[@id=\"ShowOnlyValue\"][@value='SentToTLAndWaitingForTreatment']";
        private const string FILTER_INVOICE_SHOW_EXPORTED_MANUALLY = "//*[@id=\"ShowOnlyValue\"][@value='ExportedForSageManually']";
        private const string FILTER_INVOICE_SHOW_EXPORTED_AUTOMATICALLY = "//*[@id=\"ShowOnlyValue\"][@value='ExportedForSageAutomatically']";

        private const string FILTER_SORTBY = "cbSortBy";


        private const string FIRST_INVOICE_ID_PATH = "//*/table[@id='tableListMenu']/tbody/tr[1]/td[5]/a";

        private const string INVOICE_ID = "//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[5]";
        private const string INVOICE_NUMBER = "//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[4]";
        private const string SITE = "//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[5]/a";
        private const string CUSTOMER = "//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[6]/a/b";
        private const string INVOICE_DATE = "//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[8]";
        private const string INVOICE_DATE_DEV = "//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[9]";

        private const string INVOICE_STATUS = "//*[@id=\"tableListMenu\"]/tbody/tr[{0}]/td[1]/a/img[@alt='Valid']";
        private const string SENT_TO_SAGE_ACCOUTED_ICON = "//*[@id=\"tableListMenu\"]/tbody/tr[{0}]/td[2]/img[@alt='Accounted']";
        private const string SENT_TO_SAGE_ERROR_ICON = "//tr[{0}]//span[contains(@class, 'fas fa-exclamation-triangle red-text')]";
        private const string SENT_TO_SAGE_ERROR = "//*[@id=\"tableListMenu\"]/tbody/tr[{0}]/td[2]/span/a/img[@alt='NOKError']";
        private const string WAITING_FOR_SAGE_PUSH = "//*[@id=\"tableListMenu\"]/tbody/tr[{0}]/td[2]/a/img[@alt='Waiting']";
        private const string DRAFT = "//*[@id=\"tableListMenu\"]/tbody/tr[{0}]/td[12]";
        private const string SEND_BY_MAIL = "//tr[{0}]//span[contains(@class, 'envelope')]";
        private const string INVOICE_NUMBERS = "//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[4]";
        private const string INVOICE_IDS = "//a[contains(text(), 'Id')]";
        private const string COMMENT_PATH = "//*[@id=\"accounting-orderinvoice-detailv2-1\"]/td[11]/div";

        private const string ALL_INVOICE_NUMBERS = "//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[6]";


        private const string INVOICE_ELEMENT = "//a[contains(@class, 'invisible-href')]";


        private const string FIRST_INVOICE_TO_GET = "accounting-orderinvoice-detailv2-id-1";

        private const string FIRST_INVOICE_NUMBER_EXIST = "//*/table[@id='tableListMenu']/tbody/tr[1]/td[4]/a/b";
        private const string NUMBER_INVOICE = "/html/body/div[2]/div/div[2]/div/div[2]/table/tbody/tr/td[6]";
        private const string SITE_INVOICE = "/html/body/div[2]/div/div[2]/div/div[2]/table/tbody/tr/td[8]";
        private const string CUSTOMER_INVOICE = "/html/body/div[2]/div/div[2]/div/div[2]/table/tbody/tr/td[9]";
        private const string DATE_INVOICE = "/html/body/div[2]/div/div[2]/div/div[2]/table/tbody/tr/td[12]";
        private const string DRAFT_INVOICE = "/html/body/div[2]/div/div[2]/div/div[2]/table/tbody/tr[1]/td[15]";
        //__________________________________Variables_______________________________________

        // Général
        [FindsBy(How = How.XPath, Using = ALL_INVOICE_NUMBERS)]
        private IWebElement _allinvoicenumber;
        [FindsBy(How = How.XPath, Using = INVOICE_ELEMENT)]
        private IWebElement __invoiceElement;

        [FindsBy(How = How.Id, Using = PLUS_BUTTON)]
        private IWebElement _plusButton;

        [FindsBy(How = How.Id, Using = NEW_MANUAL_INVOICE)]
        private IWebElement _createManualInvoice;

        [FindsBy(How = How.XPath, Using = NEW_AUTO_INVOICE)]
        private IWebElement _createAutoInvoice;

        [FindsBy(How = How.Id, Using = EXTENDED_BUTTON)]
        private IWebElement _extendedButton;

        [FindsBy(How = How.Id, Using = PRINT)]
        private IWebElement _printResults;

        [FindsBy(How = How.Id, Using = CONFIRM_PRINT_PDF)]
        private IWebElement _confirmPrintPdf;

        [FindsBy(How = How.Id, Using = CONFIRM_PRINT_ZIP)]
        private IWebElement _confirmPrintZIP;

        [FindsBy(How = How.Id, Using = SEND_INVOICES_MAIL)]
        private IWebElement _sendByMail;

        [FindsBy(How = How.Id, Using = CONFIRM_SEND)]
        private IWebElement _confirmSendMail;

        [FindsBy(How = How.Id, Using = CANCEL)]
        private IWebElement _cancel;

        [FindsBy(How = How.Id, Using = EXPORT)]
        private IWebElement _exportResults;

        [FindsBy(How = How.XPath, Using = CUSTOM_REPORT)]
        private IWebElement _customReport;

        [FindsBy(How = How.Id, Using = CUSTOM_REPORT_SITE)]
        private IWebElement _customReportSite;

        [FindsBy(How = How.XPath, Using = CONFIRM_CUSTOM_REPORT_PRINT)]
        private IWebElement _confirmCustomReportPrint;

        [FindsBy(How = How.XPath, Using = ENABLE_EXPORT_FOR_SAGE)]
        private IWebElement _enableExportForSage;

        [FindsBy(How = How.Id, Using = EXPORT_FOR_SAGE)]
        private IWebElement _exportForSage;

        [FindsBy(How = How.XPath, Using = SEND_AUTO_TO_SAGE)]
        private IWebElement _sendAutoToSage;

        [FindsBy(How = How.Id, Using = GENERATE_SAGE_TXT)]
        private IWebElement _generateSageTxt;

        [FindsBy(How = How.Id, Using = CONFIRM_EXPORT_SAGE)]
        private IWebElement _validateExportSage;

        [FindsBy(How = How.Id, Using = CHECK_FIRST_COMMENT)]
        private IWebElement _checkFirstComment;

        [FindsBy(How = How.Id, Using = VALIDATE_RESULTS)]
        private IWebElement _validateResults;

        [FindsBy(How = How.Id, Using = CONFIRM_VALIDATE_RESULTS)]
        private IWebElement _confirmValidate;

        [FindsBy(How = How.Id, Using = DELETE_UNVALIDATED_RESULTS)]
        private IWebElement _deleteUnvalidatedResults;

        [FindsBy(How = How.Id, Using = CONFIRM_DELETE_UNVALIDATED_RESULTS)]
        private IWebElement _confirmDeleteUnvalidatedResults;

        [FindsBy(How = How.Id, Using = CLOSE)]
        private IWebElement _close;

        [FindsBy(How = How.Id, Using = CUSTOM_REPORT_MONTH)]
        private IWebElement _customReportMonth;


        // Tableau invoices
        [FindsBy(How = How.XPath, Using = FIRST_INVOICE_ID)]
        private IWebElement _firstInvoiceID;

        [FindsBy(How = How.XPath, Using = FIRST_INVOICE_NUMBER)]
        private IWebElement _firstInvoiceNumber;

        [FindsBy(How = How.Id, Using = FIRST_INVOICE_DELETE)]
        private IWebElement _firstDelete;

        [FindsBy(How = How.Id, Using = CONFIRM_DELETE)]
        private IWebElement _confirmDelete;

        [FindsBy(How = How.Id, Using = CONFIRM_DELETE2)]
        private IWebElement _confirmDelete2;

        [FindsBy(How = How.XPath, Using = FIRST_VALID_BTN)]
        private IWebElement _firstValidBtn;

        [FindsBy(How = How.XPath, Using = SECOND_VALID_BTN)]
        private IWebElement _secondValidBtn;

        [FindsBy(How = How.XPath, Using = FIRST_DISQUETTE_BTN)]
        private IWebElement _firstDisquetteBtn;

        [FindsBy(How = How.XPath, Using = SECOND_DISQUETTE_BTN)]
        private IWebElement _secondDisquetteBtn;


        //__________________________________Filters_______________________________________



        [FindsBy(How = How.Id, Using = RESET_FILTER)]
        private IWebElement _resetFilter;

        [FindsBy(How = How.Id, Using = FILTER_SEARCH)]
        private IWebElement _searchFilter;

        [FindsBy(How = How.XPath, Using = SHOW_SITES)]
        private IWebElement _showSites;

        [FindsBy(How = How.Id, Using = FILTER_SITES)]
        private IWebElement _siteFilter;

        [FindsBy(How = How.XPath, Using = UNSELECT_ALL_SITES)]
        private IWebElement _unselectAllSites;

        [FindsBy(How = How.XPath, Using = SEARCH_SITES)]
        private IWebElement _searchSite;

        [FindsBy(How = How.XPath, Using = SHOW_CUSTOMERS)]
        private IWebElement _showCustomers;

        [FindsBy(How = How.Id, Using = FILTER_CUSTOMERS)]
        private IWebElement _customerFilter;

        [FindsBy(How = How.XPath, Using = UNSELECT_ALL_CUSTOMERS)]
        private IWebElement _unselectAllCustomers;

        [FindsBy(How = How.XPath, Using = SEARCH_CUSTOMERS)]
        private IWebElement _searchCustomer;

        [FindsBy(How = How.Id, Using = FILTER_DATE_FROM)]
        public IWebElement _dateFrom;

        [FindsBy(How = How.Id, Using = FILTER_DATE_TO)]
        public IWebElement _dateTo;

        [FindsBy(How = How.XPath, Using = SHOW_INVOICE_STEP)]
        private IWebElement _showInvoiceStep;

        [FindsBy(How = How.Id, Using = FILTER_INVOICE_STEP_ALL)]
        private IWebElement _invoiceStepAll;

        [FindsBy(How = How.Id, Using = FILTER_INVOICE_STEP_DRAFT)]
        private IWebElement _invoiceStepDraft;

        [FindsBy(How = How.Id, Using = FILTER_INVOICE_STEP_VALIDATED)]
        private IWebElement _invoiceStepValidated;

        [FindsBy(How = How.Id, Using = FILTER_INVOICE_STEP_ACCOUNTED)]
        private IWebElement _invoiceStepAccounted;

        [FindsBy(How = How.XPath, Using = SHOW_INVOICE_SHOW)]
        private IWebElement _showInvoiceShow;

        [FindsBy(How = How.XPath, Using = FILTER_INVOICE_SHOW_ALL)]
        private IWebElement _showAll;

        [FindsBy(How = How.XPath, Using = FILTER_INVOICE_SHOW_FLIGHT_INVOICES)]
        private IWebElement _showFlightInvoice;

        [FindsBy(How = How.XPath, Using = FILTER_INVOICE_SHOW_DELIVERY_INVOICES)]
        private IWebElement _showDeliveryInvoice;

        [FindsBy(How = How.XPath, Using = FILTER_INVOICE_SHOW_CUSTOMER_INVOICES)]
        private IWebElement _showCustomerOrderInvoice;

        [FindsBy(How = How.XPath, Using = FILTER_INVOICE_SHOW_MANUAL_INVOICES)]
        private IWebElement _showManualInvoice;

        [FindsBy(How = How.XPath, Using = FILTER_INVOICE_SHOW_ALL_INVOICES)]
        private IWebElement _showAllInvoices;

        [FindsBy(How = How.XPath, Using = FILTER_INVOICE_SHOW_NOT_SENT_MAIL)]
        private IWebElement _showNotSentByMailOnly;

        [FindsBy(How = How.XPath, Using = FILTER_INVOICE_SHOW_NOT_SENT_SAGE)]
        private IWebElement _showNotSentToSage;

        [FindsBy(How = How.XPath, Using = FILTER_INVOICE_SHOW_NOT_VALIDATED)]
        private IWebElement _showNotValidatedOnly;

        [FindsBy(How = How.XPath, Using = FILTER_INVOICE_SHOW_NOT_SAGE_ERROR)]
        private IWebElement _showSentToSageAndInError;

        [FindsBy(How = How.XPath, Using = FILTER_INVOICE_SHOW_ON_TL_WAITING_SAGE_PUSH)]
        private IWebElement _showOnTLWaitingForSagePush;

        [FindsBy(How = How.XPath, Using = FILTER_INVOICE_SHOW_EXPORTED_MANUALLY)]
        private IWebElement _showExportedManually;

        [FindsBy(How = How.XPath, Using = FILTER_INVOICE_SHOW_EXPORTED_AUTOMATICALLY)]
        private IWebElement _showExportedAutomatically;

        [FindsBy(How = How.Id, Using = FILTER_SORTBY)]
        private IWebElement _sortBy;

        public enum FilterType
        {
            ByNumber,
            Customer,
            Site,
            DateFrom,
            DateTo,
            SortBy
        }

        public enum FilterInvoiceStepType
        {
            ALL,
            DRAFT,
            VALIDATED_ACCOUNTED,
            ACCOUNTED
        }

        public enum FilterShowType
        {
            ShowAll,
            ShowFlightInvoice,
            ShowDeliveryInvoice,
            ShowCustomerOrderInvoice,
            ShowManualInvoice,
            ShowAllInvoices,
            ShowNotSentByMailOnly,
            ShowNotSentToSage,
            ShowNotValidatedOnly,
            ShowSentToSageAndInError,
            ShowOnTLWaitingForSagePush,
            ShowExportedForSageAutomatically,
            ShowExportedForSageManually
        }

        public void Filter(FilterType filterType, object value)
        {
            switch (filterType)
            {
                case FilterType.ByNumber:
                    _searchFilter = WaitForElementIsVisible(By.Id(FILTER_SEARCH));
                    _searchFilter.SetValue(ControlType.TextBox, value);
                    break;

                case FilterType.Site:
                    ShowSites();

                    _siteFilter = WaitForElementIsVisible(By.Id(FILTER_SITES));
                    _siteFilter.Click();

                    _unselectAllSites = WaitForElementIsVisible(By.XPath(UNSELECT_ALL_SITES));
                    _unselectAllSites.Click();

                    _searchSite = WaitForElementIsVisible(By.XPath(SEARCH_SITES));
                    _searchSite.SetValue(ControlType.TextBox, value);

                    var searchSiteValue = WaitForElementIsVisible(By.XPath("//span[contains(text(),'" + value + "')]"));
                    searchSiteValue.Click();

                    _siteFilter.Click();
                    break;

                case FilterType.Customer:
                    ShowCustomers();

                    _customerFilter = WaitForElementIsVisible(By.Id(FILTER_CUSTOMERS));
                    _customerFilter.Click();

                    _unselectAllCustomers = WaitForElementIsVisible(By.XPath(UNSELECT_ALL_CUSTOMERS));
                    _unselectAllCustomers.Click();

                    _searchCustomer = WaitForElementIsVisible(By.XPath(SEARCH_CUSTOMERS));
                    _searchCustomer.SetValue(ControlType.TextBox, value);

                    var searchCustomerValue = WaitForElementIsVisible(By.XPath("//span[contains(text(),'" + value + "')]"));
                    searchCustomerValue.Click();

                    _customerFilter.Click();
                    break;

                case FilterType.DateFrom:
                    WaitPageLoading();
                    _dateFrom = WaitForElementIsVisible(By.Id(FILTER_DATE_FROM));
                    WaitPageLoading();
                    _dateFrom.SetValue(ControlType.DateTime, value);
                    _dateFrom.SendKeys(Keys.Tab);
                    break;

                case FilterType.DateTo:
                    _dateTo = WaitForElementIsVisible(By.Id(FILTER_DATE_TO));
                    _dateTo.SetValue(ControlType.DateTime, value);
                    _dateTo.SendKeys(Keys.Tab);
                    break;

                case FilterType.SortBy:
                    _sortBy = WaitForElementIsVisible(By.Id(FILTER_SORTBY));
                    _sortBy.Click();
                    var element = WaitForElementIsVisible(By.XPath("//option[contains(@value,'" + value + "')]"));
                    _sortBy.SetValue(ControlType.DropDownList, element.Text);
                    _sortBy.Click();
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);

            }

            WaitPageLoading();
            Thread.Sleep(1500);
        }

        public int LinesCountByCustomer(string customerName)
        {
            return _webDriver.FindElements(By.XPath("//*/b[text()='" + customerName + "']")).Count;
        }

        public void FilterInvoiceStep(FilterInvoiceStepType invoiceStepType, object value)
        {
            Actions action = new Actions(_webDriver);

            switch (invoiceStepType)
            {
                case FilterInvoiceStepType.ALL:
                    _invoiceStepAll = WaitForElementExists(By.Id(FILTER_INVOICE_STEP_ALL));
                    action.MoveToElement(_invoiceStepAll).Perform();
                    _invoiceStepAll.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterInvoiceStepType.DRAFT:
                    _invoiceStepDraft = WaitForElementExists(By.Id(FILTER_INVOICE_STEP_DRAFT));
                    action.MoveToElement(_invoiceStepDraft).Perform();
                    _invoiceStepDraft.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterInvoiceStepType.VALIDATED_ACCOUNTED:
                    _invoiceStepValidated = WaitForElementExists(By.Id(FILTER_INVOICE_STEP_VALIDATED));
                    action.MoveToElement(_invoiceStepDraft).Perform();
                    _invoiceStepValidated.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterInvoiceStepType.ACCOUNTED:
                    _invoiceStepAccounted = WaitForElementExists(By.Id(FILTER_INVOICE_STEP_ACCOUNTED));
                    action.MoveToElement(_invoiceStepAccounted).Perform();
                    _invoiceStepAccounted.SetValue(ControlType.RadioButton, value);
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterInvoiceStepType), invoiceStepType, null);
            }

            WaitPageLoading();
            WaitForLoad();
        }

        public void FilterShow(FilterShowType showType, object value)
        {
            Actions action = new Actions(_webDriver);

            switch (showType)
            {
                case FilterShowType.ShowAll:
                    _showAll = WaitForElementExists(By.XPath(FILTER_INVOICE_SHOW_ALL));
                    action.MoveToElement(_showAll).Perform();
                    _showAll.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterShowType.ShowFlightInvoice:
                    _showFlightInvoice = WaitForElementExists(By.XPath(FILTER_INVOICE_SHOW_FLIGHT_INVOICES));
                    action.MoveToElement(_showFlightInvoice).Perform();
                    _showFlightInvoice.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterShowType.ShowDeliveryInvoice:
                    _showDeliveryInvoice = WaitForElementExists(By.XPath(FILTER_INVOICE_SHOW_DELIVERY_INVOICES));
                    action.MoveToElement(_showDeliveryInvoice).Perform();
                    _showDeliveryInvoice.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterShowType.ShowCustomerOrderInvoice:
                    _showCustomerOrderInvoice = WaitForElementExists(By.XPath(FILTER_INVOICE_SHOW_CUSTOMER_INVOICES));
                    action.MoveToElement(_showCustomerOrderInvoice).Perform();
                    _showCustomerOrderInvoice.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterShowType.ShowManualInvoice:
                    _showManualInvoice = WaitForElementExists(By.XPath(FILTER_INVOICE_SHOW_MANUAL_INVOICES));
                    action.MoveToElement(_showManualInvoice).Perform();
                    _showManualInvoice.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterShowType.ShowAllInvoices:
                    _showAllInvoices = WaitForElementExists(By.XPath(FILTER_INVOICE_SHOW_ALL_INVOICES));
                    action.MoveToElement(_showAllInvoices).Perform();
                    _showAllInvoices.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterShowType.ShowNotSentByMailOnly:
                    _showNotSentByMailOnly = WaitForElementExists(By.XPath(FILTER_INVOICE_SHOW_NOT_SENT_MAIL));
                    action.MoveToElement(_showNotSentByMailOnly).Perform();
                    _showNotSentByMailOnly.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterShowType.ShowNotSentToSage:
                    _showNotSentToSage = WaitForElementExists(By.XPath(FILTER_INVOICE_SHOW_NOT_SENT_SAGE));
                    action.MoveToElement(_showNotSentToSage).Perform();
                    _showNotSentToSage.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterShowType.ShowNotValidatedOnly:
                    _showNotValidatedOnly = WaitForElementExists(By.XPath(FILTER_INVOICE_SHOW_NOT_VALIDATED));
                    action.MoveToElement(_showNotValidatedOnly).Perform();
                    _showNotValidatedOnly.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterShowType.ShowSentToSageAndInError:
                    _showSentToSageAndInError = WaitForElementExists(By.XPath(FILTER_INVOICE_SHOW_NOT_SAGE_ERROR));
                    action.MoveToElement(_showSentToSageAndInError).Perform();
                    _showSentToSageAndInError.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterShowType.ShowOnTLWaitingForSagePush:
                    _showOnTLWaitingForSagePush = WaitForElementExists(By.XPath(FILTER_INVOICE_SHOW_ON_TL_WAITING_SAGE_PUSH));
                    action.MoveToElement(_showOnTLWaitingForSagePush).Perform();
                    _showOnTLWaitingForSagePush.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterShowType.ShowExportedForSageAutomatically:
                    _showExportedAutomatically = WaitForElementExists(By.XPath(FILTER_INVOICE_SHOW_EXPORTED_AUTOMATICALLY));
                    action.MoveToElement(_showExportedAutomatically).Perform();
                    _showExportedAutomatically.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterShowType.ShowExportedForSageManually:
                    _showExportedManually = WaitForElementExists(By.XPath(FILTER_INVOICE_SHOW_EXPORTED_MANUALLY));
                    action.MoveToElement(_showExportedManually).Perform();
                    _showExportedManually.SetValue(ControlType.RadioButton, value);
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterShowType), showType, null);
            }

            WaitPageLoading();
            WaitForLoad();
        }

        public void ResetFilters()
        {
            _resetFilter = WaitForElementIsVisible(By.Id(RESET_FILTER));
            _resetFilter.Click();
            WaitForLoad();

            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                //Filter(FilterType.DateTo, DateUtils.Now);
            }
        }

        public void ShowSites()
        {
            _showSites = WaitForElementIsVisible(By.XPath(SHOW_SITES));

            if (_showSites.GetCssValue("aria-expanded").Equals(false))
            {
                _showSites.Click();
                WaitForLoad();
            }
        }

        public void ShowCustomers()
        {
            _showCustomers = WaitForElementIsVisible(By.XPath(SHOW_CUSTOMERS));

            if (_showCustomers.GetCssValue("aria-expanded").Equals(false))
            {
                _showCustomers.Click();
                WaitForLoad();
            }
        }

        public void ShowInvoiceStep()
        {
            _showInvoiceStep = WaitForElementIsVisible(By.XPath(SHOW_INVOICE_STEP));
            _showInvoiceStep.Click();
            WaitForLoad();

            // Temps que le contenu du menu Show soit accessible
            WaitPageLoading();
        }

        public void ShowInvoiceShow()
        {
            _showInvoiceShow = WaitForElementIsVisible(By.XPath(SHOW_INVOICE_SHOW));
            _showInvoiceShow.Click();
            WaitForLoad();

            // Temps que le contenu du menu Show soit accessible
            WaitPageLoading();
        }


        public bool VerifySite(string site)
        {
            IEnumerable<IWebElement> sites;
            sites = _webDriver.FindElements(By.XPath("//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[8]/a"));

            if (sites.Count() == 0)
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

        public bool VerifyCustomer(string customer)
        {
            IEnumerable<IWebElement> customers;
            customers = _webDriver.FindElements(By.XPath("//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[9]/a/b"));


            if (customers.Count() == 0)
                return false;

            foreach (var elm in customers)
            {
                if (elm.Text != customer)
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
            ReadOnlyCollection<IWebElement> dates;
            dates = _webDriver.FindElements(By.XPath("//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[12]"));


            /*//PATCH 2022.0302.1-P21
            if (dates[0].Text.Contains("€ excl. VAT"))
            {
                dates = _webDriver.FindElements(By.XPath("//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[8]"));
            }*/

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

        public bool IsSortedById()
        {
            IEnumerable<IWebElement> listeId;
            listeId = _webDriver.FindElements(By.XPath(INVOICE_ID));

            if (listeId.Count() == 0)
                return false;

            var ancientId = int.Parse(Regex.Match(listeId.FirstOrDefault().Text, @"\d+").Value);

            foreach (var elm in listeId)
            {
                try
                {
                    if (int.Parse(Regex.Match(elm.Text, @"\d+").Value) > ancientId)
                        return false;

                    ancientId = int.Parse(Regex.Match(elm.Text, @"\d+").Value);
                }
                catch
                {
                    return false;
                }
            }

            return true;
        }

        public bool IsSortedByName()
        {
            IEnumerable<IWebElement> listeCustomers;
            listeCustomers = _webDriver.FindElements(By.XPath("//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[9]/a/b"));

            if (listeCustomers.Count() == 0)
                return false;

            var ancientCustomer = listeCustomers.FirstOrDefault().Text;

            foreach (var elm in listeCustomers)
            {
                try
                {
                    if (elm.Text.CompareTo(ancientCustomer) < 0)
                        return false;

                    ancientCustomer = elm.Text;
                }
                catch
                {
                    return false;
                }
            }

            return true;
        }

        public bool IsSortedByNumber()
        {
            IEnumerable<IWebElement> numbers;
            numbers = _webDriver.FindElements(By.XPath("//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[6]"));

            if (numbers.Count() == 0)
                return false;

            var ancientNumber = numbers.FirstOrDefault().Text;

            foreach (var elm in numbers)
            {
                try
                {
                    if (elm.Text.CompareTo(ancientNumber) < 0)
                        return false;

                    ancientNumber = elm.Text;
                }
                catch
                {
                    return false;
                }
            }

            return true;
        }

        public bool CheckValidation(bool validated)
        {
            bool isValidated = true;
            int tot = CheckTotalNumber() > 100 ? 100 : CheckTotalNumber();

            if (tot == 0)
                return false;

            for (int i = 0; i < tot; i++)
            {
                if (isElementVisible(By.XPath(string.Format("//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[1]/div/i[@class='fa-regular fa-circle-check']", i + 1))))
                {
                    _webDriver.FindElement(By.XPath(string.Format("//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[1]/div/i[@class='fa-regular fa-circle-check']", i + 1)));

                    if (!validated)
                        return true;
                }
                else
                {
                    isValidated = false;

                    if (validated)
                        return false;
                }
            }

            return isValidated;
        }

        public bool IsDraft()
        {
            int tot = CheckTotalNumber() > 100 ? 100 : CheckTotalNumber();

            if (tot == 0)
                return false;

            for (int i = 0; i < tot; i++)
            {
                try
                {
                    IWebElement draft;
                    draft = _webDriver.FindElement(By.XPath(string.Format("//*[@id=\"tableListMenu\"]/tbody/tr[{0}]/td[14]", i + 1)));

                    /*//PATCH 2022.0302.1-P21
                    if (String.IsNullOrEmpty(draft.Text) || !draft.Text.Equals("Draft"))
                    {
                        draft = _webDriver.FindElement(By.XPath(string.Format("//*[@id=\"tableListMenu\"]/tbody/tr[{0}]/td[11]", i + 1)));
                    }
                    //PATCH 2022.0302.1-P21
                    if (String.IsNullOrEmpty(draft.Text) || !draft.Text.Equals("Draft"))
                        return false;
                    */
                }
                catch
                {
                    return false;
                }
            }
            return true;
        }

        public Boolean IsSentByMail()
        {
            int tot = CheckTotalNumber() > 100 ? 100 : CheckTotalNumber();

            if (tot == 0)
                return false;

            for (int i = 0; i < tot; i++)
            {
                if (isElementVisible(By.XPath(string.Format(SEND_BY_MAIL, i + 1))))
                {
                    _webDriver.FindElement(By.XPath(string.Format(SEND_BY_MAIL, i + 1)));
                }
                else
                {
                    return false;
                }
            }
            return true;
        }

        public Boolean IsSentToSage()
        {
            int tot = CheckTotalNumber() > 100 ? 100 : CheckTotalNumber();
            int counter = 0;
            if (tot == 0)
                return false;

            for (int i = 0; i < tot; i++)
            {
                if (isElementVisible(By.XPath(string.Format("//*[@id=\"tableListMenu\"]/tbody/tr[{0}]/td[2]/div/i[@class='fas fa-save green-text']", i + 1))))
                {
                    counter++;
                }
                else if (isElementVisible(By.XPath(string.Format(SENT_TO_SAGE_ERROR_ICON, i + 1))))
                {
                    counter++;
                }
            }
            return counter == tot;
        }

        public bool IsSentToSAGEInErrorOnly()
        {
            int tot = CheckTotalNumber() > 100 ? 100 : CheckTotalNumber();

            if (tot == 0)
                return false;

            for (int i = 0; i < tot; i++)
            {
                if (isElementVisible(By.XPath(string.Format(SENT_TO_SAGE_ERROR, i + 1))))
                {
                    _webDriver.FindElement(By.XPath(string.Format(SENT_TO_SAGE_ERROR, i + 1)));
                }
                else if (isElementVisible(By.XPath(string.Format(SENT_TO_SAGE_ERROR_ICON, i + 1))))
                {
                    _webDriver.FindElement(By.XPath(string.Format(SENT_TO_SAGE_ERROR_ICON, i + 1)));
                }
                else
                {
                    return false;
                }

            }
            return true;

        }

        public bool IsWaitingForSAGEPush()
        {
            int tot = CheckTotalNumber() > 100 ? 100 : CheckTotalNumber();

            if (tot == 0)
                return false;

            for (int i = 0; i < tot; i++)
            {
                if (isElementVisible(By.XPath(string.Format(WAITING_FOR_SAGE_PUSH, i + 1))))
                {
                    _webDriver.FindElement(By.XPath(string.Format(WAITING_FOR_SAGE_PUSH, i + 1)));
                }
                else
                {
                    return false;
                }
            }
            return true;

        }

        //__________________________________Pages_______________________________________

        public override void ShowPlusMenu()
        {
            var actions = new Actions(_webDriver);
            _plusButton = WaitForElementIsVisible(By.Id(PLUS_BUTTON));
            actions.MoveToElement(_plusButton).Perform();

        }

        public ManualInvoiceCreateModalPage ManualInvoiceCreatePage()
        {
            ShowPlusMenu();

            _createManualInvoice = WaitForElementIsVisible(By.Id(NEW_MANUAL_INVOICE));
            _createManualInvoice.Click();
            WaitForLoad();

            return new ManualInvoiceCreateModalPage(_webDriver, _testContext);
        }

        public AutoInvoiceCreateModalPage AutoInvoiceCreatePage()
        {
            ShowPlusMenu();

            _createAutoInvoice = WaitForElementIsVisible(By.Id(NEW_AUTO_INVOICE));
            _createAutoInvoice.Click();
            WaitForLoad();

            return new AutoInvoiceCreateModalPage(_webDriver, _testContext);
        }

        public ManualCreditNoteCreatePage ManualCreditNoteCreatePage()
        {
            ShowPlusMenu();

            _createManualInvoice = WaitForElementIsVisible(By.Id(NEW_MANUAL_INVOICE));
            _createManualInvoice.Click();
            WaitForLoad();

            return new ManualCreditNoteCreatePage(_webDriver, _testContext);
        }

        public override void ShowExtendedMenu()
        {
            var actions = new Actions(_webDriver);
            _extendedButton = WaitForElementIsVisible(By.Id(EXTENDED_BUTTON));
            actions.MoveToElement(_extendedButton).Perform();
            WaitForLoad();
        }

        public void ShowExtendedMenuInInvoice()
        {
            var actions = new Actions(_webDriver);
            _extendedButton = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div[1]/div/div[1]/button"));
            actions.MoveToElement(_extendedButton).Perform();
            WaitForLoad();
        }

        public PrintReportPage PrintInvoiceResults()
        {
            WaitForLoad();
            ShowExtendedMenu();

            _printResults = WaitForElementIsVisible(By.Id(PRINT));
            _printResults.Click();
            WaitForLoad();

            _confirmPrintPdf = WaitForElementIsVisible(By.Id(CONFIRM_PRINT_PDF));
            _confirmPrintPdf.Click();
            WaitForLoad();

            IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-pdf']"));

            //Results are opened in a new tab, switch the driver to the newly created one
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new PrintReportPage(_webDriver, _testContext);
        }

        public void ExportZipFile()
        {
            ShowExtendedMenu();

            _printResults = WaitForElementIsVisible(By.Id(PRINT));
            _printResults.Click();
            WaitPageLoading();
            _confirmPrintZIP = WaitForElementIsVisible(By.Id(CONFIRM_PRINT_ZIP));
            WaitPageLoading();
            _confirmPrintZIP.Click();
            WaitPageLoading();
            WaitForDownload();
            Close();
        }

        public FileInfo GetExportZipFile(FileInfo[] taskFiles)
        {
            List<FileInfo> correctDownloadFiles = new List<FileInfo>();

            foreach (var file in taskFiles)
            {
                //  Test REGEX
                if (IsZipFileCorrect(file.Name))
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

        public bool IsZipFileCorrect(string filePath)
        {
            // invoices_2020-1-21-5be91c1c-6d01-4541-8a97-1579354d24a8.zip

            string mois = "(?:0?[1-9]|1[0-2])";          // mois
            string annee = "\\d{4}";                     // annee YYYY
            string jour = "(?:0?[1-9]|[1-2]\\d|3[0-1])"; // jour
            string caracteres = ".+";                    // caractères


            Regex r = new Regex("^invoices_" + annee + "-" + mois + "-" + jour + "-" + caracteres + "\\.zip$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = r.Match(filePath);

            return m.Success;
        }

        public bool SendByMail()
        {
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;

            ShowExtendedMenu();

            _sendByMail = WaitForElementIsVisible(By.Id(SEND_INVOICES_MAIL));
            _sendByMail.Click();
            WaitPageLoading();

            if (isElementVisible(By.Id(CONFIRM_SEND)))
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
            return true;
        }

        public void ExportExcelFile()
        {
            ShowExtendedMenu();

            _exportResults = WaitForElementIsVisible(By.Id(EXPORT));
            _exportResults.Click();
            WaitForLoad();

            IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));
            ClickPrintButton();

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
            // "export-invoices 2019-12-04 14-54-16.xlsx"

            string mois = "(?:0[1-9]|1[0-2])";         // mois
            string space = "(\\s)";                    // Espace
            string annee = "\\d{4}";                   // annee YYYY
            string jour = "[0-3]\\d";                  // jour
            string heure = "(?:0[0-9]|1[0-9]|2[0-3])"; // heure
            string minutes = "[0-5]\\d";               // minutes
            string secondes = "[0-5]\\d";              // secondes

            Regex r = new Regex("^export-invoices" + space + annee + "-" + mois + "-" + jour + space + heure + "-" + minutes + "-" + secondes + ".xlsx$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = r.Match(filePath);

            return m.Success;
        }

        public void ExportCustomFile(string downloadPath, string site, string month = "")
        {
            ShowExtendedMenu();

            _customReport = WaitForElementIsVisible(By.XPath(CUSTOM_REPORT));
            _customReport.Click();
            WaitForLoad();

            _customReportSite = WaitForElementIsVisible(By.Id(CUSTOM_REPORT_SITE));
            _customReportSite.SetValue(ControlType.DropDownList, site);
            WaitForLoad();

            if (month != "")
            {
                _customReportMonth = WaitForElementIsVisible(By.Id(CUSTOM_REPORT_MONTH));
                _customReportMonth.SetValue(ControlType.DropDownList, month);
                WaitForLoad();
            }

            _confirmCustomReportPrint = WaitForElementIsVisible(By.XPath(CONFIRM_CUSTOM_REPORT_PRINT));
            _confirmCustomReportPrint.Click();
            WaitForLoad();

            IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));
            ClickPrintButton();

            WaitForDownload();
            Close();
        }

        public FileInfo GetExportCustomFile(FileInfo[] taskFiles)
        {
            List<FileInfo> correctDownloadFiles = new List<FileInfo>();

            foreach (var file in taskFiles)
            {
                //  Test REGEX
                if (IsCustomFileCorrect(file.Name))
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

        public bool IsCustomFileCorrect(string filePath)
        {
            // "Relevé des livraisons douanes BCN - 2019 - 12 (1).xlsx"

            string mois = "(?:0[1-9]|1[0-2])";         // mois
            string space = "(\\s)";                    // Espace
            string annee = "\\d{4}";                   // annee YYYY
            string site = "([A-Z]{3})";
            string reste = ".*";

            Regex r = new Regex("^Relevé des livraisons douanes" + space + site + space + "-" + space + annee + space + "-" + space + mois + reste + ".xlsx$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = r.Match(filePath);

            return m.Success;
        }

        public void EnableExportForSage()
        {
            ShowExtendedMenu();

            _enableExportForSage = WaitForElementIsVisible(By.XPath(ENABLE_EXPORT_FOR_SAGE));
            _enableExportForSage.Click();
            WaitForLoad();
        }

        public void SendAutoToSage()
        {
            ShowExtendedMenu();

            _sendAutoToSage = WaitForElementIsVisible(By.XPath(SEND_AUTO_TO_SAGE));
            _sendAutoToSage.Click();
            WaitForLoad();

            _validateExportSage = WaitForElementExists(By.Id(CONFIRM_EXPORT_SAGE));
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", _validateExportSage);
            WaitForLoad();

            _validateExportSage.Click();
            WaitForLoad();

            // On ferme la pop-up
            _close = WaitForElementIsVisible(By.Id(CLOSE));
            _close.Click();
            WaitForLoad();
        }

        public void GenerateSageTxt()
        {
            ShowExtendedMenu();

            _generateSageTxt = WaitForElementIsVisible(By.Id(GENERATE_SAGE_TXT));
            _generateSageTxt.Click();
            WaitForLoad();

            _validateExportSage = WaitForElementExists(By.Id(CONFIRM_EXPORT_SAGE));
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", _validateExportSage);
            WaitForLoad();

            _validateExportSage.Click();
            WaitForLoad();

            IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-alt']"));
            ClickPrintButton();

            WaitForDownload();
            Close();
        }

        public void ExportResultsToSage()
        {
            //Export results=Generate txt
            ShowExtendedMenu();

            if (isElementVisible(By.Id(EXPORT_FOR_SAGE)))
            {
                _exportForSage = WaitForElementIsVisible(By.Id(EXPORT_FOR_SAGE));
                _exportForSage.Click();
                WaitForLoad();
            }
            else
            {
                _generateSageTxt = WaitForElementIsVisible(By.Id(GENERATE_SAGE_TXT));
                _generateSageTxt.Click();
                WaitForLoad();

            }

            _validateExportSage = WaitForElementIsVisible(By.Id(CONFIRM_EXPORT_SAGE));
            _validateExportSage.Click();
            WaitForLoad();

            IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-alt']"));
            ClickPrintButton();

            WaitForDownload();
            Close();
        }

        public FileInfo GetExportSageFile(FileInfo[] taskFiles)
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

        public bool IsSageFileCorrect(string filePath)
        {
            // "invoices 2019-12-05 11-25-46.txt"
            //ou
            // "invoice 158654 2019-12-05 11-25-46.txt"


            string beginInvoice = "(^invoice\\s[0-9]+)|(^invoices)";           // begin
            string mois = "(?:0[1-9]|1[0-2])";         // mois
            string space = "(\\s)";                    // Espace
            string annee = "\\d{4}";                   // annee YYYY
            string jour = "[0-3]\\d";                  // jour
            string heure = "(?:0[0-9]|1[0-9]|2[0-3])"; // heure
            string minutes = "[0-5]\\d";               // minutes
            string secondes = "[0-5]\\d";              // secondes

            Regex r = new Regex(beginInvoice + space + annee + "-" + mois + "-" + jour + space + heure + "-" + minutes + "-" + secondes + ".txt$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = r.Match(filePath);

            return m.Success;
        }

        public void ValidateResults()
        {
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;

            ShowExtendedMenu();

            _validateResults = WaitForElementIsVisible(By.Id(VALIDATE_RESULTS));
            _validateResults.Click();
            WaitForLoad();

            _confirmValidate = WaitForElementExists(By.Id(CONFIRM_VALIDATE_RESULTS));
            javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", _confirmValidate);
            WaitPageLoading();

            _confirmValidate.Click();
            WaitPageLoading();

        }

        public void DeleteUnValidatedInvoices()
        {
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;

            ShowExtendedMenu();
            WaitForLoad();
            _deleteUnvalidatedResults = WaitForElementIsVisible(By.Id(DELETE_UNVALIDATED_RESULTS));
            _deleteUnvalidatedResults.Click();
            WaitForLoad();

            _confirmDeleteUnvalidatedResults = WaitForElementExists(By.Id(CONFIRM_DELETE_UNVALIDATED_RESULTS));
            javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", _confirmDeleteUnvalidatedResults);
            WaitPageLoading();

            _confirmDeleteUnvalidatedResults.Click();
            WaitForLoad();
        }

        // Tableau invoices

        public InvoiceDetailsPage SelectFirstInvoice()
        {
            _firstInvoiceID = WaitForElementToBeClickable(By.XPath(FIRST_INVOICE_ID));
            _firstInvoiceID.Click();
            WaitForLoad();

            return new InvoiceDetailsPage(_webDriver, _testContext);
        }

        public string GetFirstInvoiceID()
        {
            WaitForLoad();
            WaitPageLoading();
            _firstInvoiceID = WaitForElementIsVisible(By.Id(FIRST_INVOICE_TO_GET));
            return _firstInvoiceID.Text;
        }

        public string GetFirstInvoiceId()
        {
            _firstInvoiceID = WaitForElementIsVisible(By.XPath(FIRST_INVOICE_ID_PATH));
            return _firstInvoiceID.Text.Substring(3);
        }

        public string GetFirstInvoiceNumber()
        {
            _firstInvoiceNumber = WaitForElementExists(By.XPath(FIRST_INVOICE_NUMBER));
            return _firstInvoiceNumber.Text;
        }

        public string GetFirstInvoiceOnlyNumber()
        {
            _firstInvoiceNumber = WaitForElementExists(By.XPath(FIRST_INVOICE_NUMBER));
            string fullText = _firstInvoiceNumber.Text;

            // Extraire uniquement les chiffres
            string onlyNumbers = Regex.Replace(fullText, @"\D", ""); // \D correspond à tout ce qui n'est pas un chiffre
            return onlyNumbers;
        }

        public string GetFirstSite()
        {
            IWebElement _firstSite;
            _firstSite = WaitForElementExists(By.Id("accounting-orderinvoice-detailv2-sitecode-1"));

            return _firstSite.Text;
        }

        public string GetFirstCustomer()
        {
            IWebElement _firstCustomer;
            _firstCustomer = WaitForElementExists(By.Id("accounting-orderinvoice-detailv2-customername-1"));

            return _firstCustomer.Text;
        }

        public void DeleteFirstInvoice()
        {
            _firstDelete = WaitForElementIsVisible(By.Id(FIRST_INVOICE_DELETE));
            _firstDelete.Click();
            WaitForLoad();

            if (isElementVisible(By.Id(CONFIRM_DELETE)))
            {
                _confirmDelete = WaitForElementIsVisible(By.Id(CONFIRM_DELETE));
                _confirmDelete.Click();
                WaitForLoad();
            }
            else
            {
                _confirmDelete2 = WaitForElementIsVisible(By.Id(CONFIRM_DELETE2));
                _confirmDelete2.Click();
                WaitForLoad();
            }

        }

        // Renaud
        public void Validate()
        {
            WaitForLoad();
            var validate = WaitForElementIsVisible(By.XPath(VALIDATE_1));
            //validate.Click();
            // le bouton "..." gène pour le passage de la sourie sur le milieu de [VALIDATE], donc je descend horizontalement
            new Actions(_webDriver).MoveToElement(validate).MoveByOffset(0, 30).Perform();
            var menuValidated = WaitForElementIsVisible(By.Id("IsInvoiceValidated"));
            menuValidated.Click();
            //var confirm = WaitForElementIsVisible(By.XPath("//*[@id=\"btn-popup-validate\"]"));
            //confirm.Click();
            // https://winrest-testauto7dev-app.azurewebsites.net/customerportal/CustomerOrders/Index_ShowDetail/FlightDelivery/IsSessionAlive
            WaitForLoad();
        }

        public bool CheckIfExistValidate()
        {
            WaitForLoad();
            return isElementVisible(By.XPath(VALIDATE_1)) ? true : false;

        }

        public void ValidateInvoice()
        {
            WaitForLoad();
            IWebElement validate = WaitForElementIsVisible(By.Id("accounting-invoice-btn"));
            validate.Click();
            new Actions(_webDriver).MoveToElement(validate).MoveByOffset(0, 30).Perform();
            IWebElement menuValidated = WaitForElementIsVisible(By.Id("IsInvoiceValidated"));
            menuValidated.Click();
            WaitForLoad();
            var confirm = WaitForElementIsVisible(By.Id("dataConfirmOK"));
            confirm.Click();
            WaitForLoad();
        }

        public string CheckFirstComment()
        {
            if (IsDev())
            {
                _checkFirstComment = WaitForElementIsVisible(By.XPath("//*[@id=\"accounting-orderinvoice-detailv2-1\"]/td[10]/div"));

            }
            else
            {
                _checkFirstComment = WaitForElementIsVisible(By.XPath(COMMENT_PATH));
            }

            if (_checkFirstComment.GetAttribute("class").Contains("green-text"))
            {
                return _checkFirstComment.GetAttribute("title");
            }
            else
            {
                return "";
            }
        }

        public bool IsExistingFirstInvoiceNumber()
        {
            bool firstInvoiceNumberExists = isElementVisible(By.XPath(FIRST_INVOICE_NUMBER_EXIST));
            return firstInvoiceNumberExists;
        }
        public bool VerifyInvoiceSeperated(string oldinvoiceid, string newinvoiceid, int flightsnumber)
        {
            if (int.Parse(newinvoiceid) == int.Parse(oldinvoiceid) + flightsnumber)
            {
                return true;
            }
            return false;
        }

        public bool VerifyInvoiceCreated(string invoiceNumber)
        {
            var invoiceNumbers = _webDriver.FindElements(By.XPath(INVOICE_NUMBERS)).Where(i => i.Text != "").ToList();
            foreach (var invoice in invoiceNumbers)
            {
                if (invoice.Text.Contains(invoiceNumber))
                {
                    return true;
                }
            }
            return false;
        }

        public List<string> GetAllIds()
        {
            var tableIds = _webDriver.FindElements(By.XPath(INVOICE_IDS));
            List<string> liste = new List<string>();
            foreach (var tableId in tableIds)
            {
                liste.Add(tableId.Text);
            }
            return liste;
        }
        public void CreateManualInvoiceWithCustomer(string customerCode)
        {
            if (IsDev())
            {
                var customerCodeInput = WaitForElementExists(By.XPath("/html/body/div[3]/div/div/div[1]/div/form/div[2]/div[4]/div/span[1]/input[2]"));
                customerCodeInput.SetValue(ControlType.TextBox, customerCode);
                WaitForLoad();
            }
            else
            {
                var customerCodeInput = WaitForElementExists(By.XPath("//*/input[@placeholder='Search customer']"));
                customerCodeInput.SetValue(ControlType.TextBox, customerCode);
                WaitForLoad();
            }

            var customerSuggestion = WaitForElementExists(By.XPath("/html/body/div[3]/div/div/div[1]/div/form/div[2]/div[4]/div/span[1]/div/div/div"));
            customerSuggestion.Click();
            WaitForLoad();

            var createManualInvoiceButton = WaitForElementToBeClickable(By.Id("btn-submit-form-create-vip-order"));
            createManualInvoiceButton.Click();
            WaitForLoad();

            var addVipInvoiceDetailBtn = WaitForElementIsVisible(By.Id("addVipInvoiceDetailBtn"));
            addVipInvoiceDetailBtn.Click();
            WaitForLoad();

            var serviceNameInput = WaitForElementExists(By.XPath("/html/body/div[3]/div/div[3]/div/div/div/div/div/div[2]/div[1]/div/div/div/table/tbody/tr[2]/td[2]/div[2]/span/span/input[2]"));
            serviceNameInput.SetValue(ControlType.TextBox, "free");
            WaitForLoad();

            var suggestionService = WaitForElementExists(By.XPath("/html/body/div[3]/div/div[3]/div/div/div/div/div/div[2]/div[1]/div/div/div/table/tbody/tr[2]/td[2]/div[2]/span/span/div/div/div"));
            suggestionService.Click();
            WaitForLoad();

            var freePriceName = WaitForElementExists(By.XPath("/html/body/div[4]/div/div/div[2]/div/form/div[1]/div[2]/div/input"));
            freePriceName.SetValue(ControlType.TextBox, customerCode);
            WaitForLoad();

            var freePriceQuantity = WaitForElementExists(By.XPath("/html/body/div[4]/div/div/div[2]/div/form/div[1]/div[4]/div/input"));
            freePriceQuantity.SetValue(ControlType.TextBox, "10");
            WaitForLoad();

            var freePriceSaveButton = WaitForElementExists(By.XPath("/html/body/div[4]/div/div/div[2]/div/form/div[2]/button[2]"));
            freePriceSaveButton.Click();
            WaitForLoad();

            ShowValidationMenu();

            var validateMenu = WaitForElementIsVisible(By.Id("IsInvoiceValidated"));
            validateMenu.Click();
            WaitForLoad();

            var confirm = WaitForElementIsVisible(By.XPath(VALIDATE_3));
            confirm.Click();
            WaitForLoad();
        }

        public void SendToEdi()
        {
            ShowExtendedMenuInInvoice();
            var ediButton = WaitForElementIsVisible(By.Id("btn-send-to-edi-detail"));
            ediButton.Click();
            WaitForLoad();
            var confirmButton = WaitForElementExists(By.XPath("/html/body/div[4]/div/div/div/div[3]/div[2]/button[2]"));
            confirmButton.Click();
            WaitForLoad();
            var closeModal = WaitForElementToBeClickable(By.XPath("/html/body/div[4]/div/div/div/div[3]/div[2]/button[1]"));
            closeModal.Click();
        }

        public string GetFirstinvoiceNumber()
        {
            if (isElementExists(By.XPath(FIRST_INVOICE_NUMBER)))
            {
                _firstInvoiceNumber = WaitForElementIsVisible(By.XPath(FIRST_INVOICE_NUMBER));
                return _firstInvoiceNumber.Text;
            }
            else
            {
                return "";
            }
        }
        public bool IsInvoiceSentByMail()
        {
            WaitPageLoading();
            var elements = _webDriver.FindElements(By.XPath(MAIL_SEND));

            IWebElement mail_send = null;

            foreach (var element in elements)
            {
                if (element.GetAttribute("class").Contains("fas fa-envelope"))
                {
                    mail_send = element;
                    break;
                }
            }

            if (mail_send != null)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        public bool IsSameDistance()
        {
            var _firstValidBtn = _webDriver.FindElements(By.XPath(FIRST_VALID_BTN));
            var _secondValidBtn = _webDriver.FindElements(By.XPath(SECOND_VALID_BTN));
            return (_firstValidBtn[0].Location.X == _secondValidBtn[0].Location.X);
        }
        public InvoiceDetailsPage ClickFirstLine()
        {
            var firstLine = WaitForElementIsVisible(By.XPath(FIRST_LINE));
            firstLine.Click();
            WaitPageLoading();
            return new InvoiceDetailsPage(_webDriver, _testContext);

        }
        public string GetInvoiceNumberById(string targetId)
        {
            var tableRows = _webDriver.FindElements(By.XPath(TABLE_ROWS));

            foreach (var row in tableRows)
            {
                IWebElement idElement = null;
                idElement = row.FindElement(By.XPath(".//td[5]/a"));

                string currentId = idElement.Text.Trim();

                if (currentId.Contains(targetId))
                {
                    IWebElement invoiceNumberElement = null;
                    invoiceNumberElement = row.FindElement(By.XPath(".//td[6]/a"));
                    return invoiceNumberElement.Text.Trim();
                }
            }

            return string.Empty;
        }
        public int TotalNumber()
        {
            WaitForLoad();
            var _totalNumber = WaitForElementExists(By.XPath("//*[@id=\"tabContentItemContainer\"]/div[1]/div/h1/span"));
            int nombre = Int32.Parse(_totalNumber.GetAttribute("innerText"));
            return nombre;
        }
        public FileInfo GetExportExcelFileInvoiceCustoms(FileInfo[] taskFiles)
        {
            List<FileInfo> correctDownloadFiles = new List<FileInfo>();

            foreach (var file in taskFiles)
            {
                //  Test REGEX
                if (IsExcelFileCorrectInvoiceCustoms(file.Name))
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
        public bool IsExcelFileCorrectInvoiceCustoms(string filePath)
        {

            string mois = "(?:0[1-9]|1[0-2])";   // mois (01 à 12)
            string annee = "\\d{4}";             // année (format YYYY)
            string site = "BCN";                 // nom du site (ici 'BCN')
            Regex r = new Regex("^Relevé des livraisons douanes " + site + " - " + annee + " - " + mois, RegexOptions.IgnoreCase | RegexOptions.Singleline);


            Match m = r.Match(filePath);

            return m.Success;
        }
        public string GetInvoiceNumber()
        {
            WaitPageLoading();
            var _invoiceElement = WaitForElementExists(By.XPath(INVOICE_ELEMENT));
            return Regex.Match(_invoiceElement.Text, @"\d+").Value;
        }

        public List<string> AllInvoicesID()
        {
            PageSize("1000");
            var elements = _webDriver.FindElements(By.XPath("//*/table[@id='tableListMenu']/tbody/tr[*]/td[5]/a"));
            List<string> result = new List<string>();
            foreach (var element in elements)
            {
                result.Add(element.GetAttribute("innerText").Substring("Id ".Length));
            }
            return result;
        }

        public bool VerifyInvoiceNumberExist(string invoiceNumber)
        {
            var invoiceNumbers = _webDriver.FindElements(By.XPath(ALL_INVOICE_NUMBERS));
            foreach (var invoice in invoiceNumbers)
            {
                if (invoice.Text.Contains(invoiceNumber))
                {
                    return true;
                }
            }
            return false;
        }

        public bool DeleteInvoiceById(string targetId)
        {
            int totalrow = CheckTotalNumber();
            var tableRows = _webDriver.FindElements(By.XPath(TABLE_ROWS));
            int i = 1;
            foreach (var row in tableRows)
            {
                IWebElement idElement = null;
                idElement = row.FindElement(By.XPath("//tr[" + i + "]//a[contains(text(), 'Id')]"));

                var test = Regex.Match(idElement.Text, @"\d+").Value;
                string currentId = Regex.Match(idElement.Text, @"\d+").Value;

                if (currentId.Contains(targetId))
                {
                    var elementdelete = WaitForElementExists(By.XPath("//*[@id=\"accounting-orderinvoice-detailv2-delete-" + i + "\"]"));
                    elementdelete.Click();
                    var elementdeleteconfirm = WaitForElementExists(By.XPath("//*[@id=\"dataConfirmOK\"]"));
                    elementdeleteconfirm.Click();
                    var totalrowAfterDelete = CheckTotalNumber();
                    if (totalrowAfterDelete < totalrow) return true;
                }
            }
            return false;
        }
        public void ShowBtnDownload()
        {
            var _btnDownload = WaitForElementIsVisible(By.Id(PRINT_BTN));
            _btnDownload.Click();
            WaitPageLoading();
        }

        public bool VerifAffichageCol(string ID, string site, string customer, string date)
        {
            var invoiceNumber = WaitForElementIsVisible(By.XPath(NUMBER_INVOICE));
            var siteInvoice = WaitForElementIsVisible(By.XPath(SITE_INVOICE));
            var customerInvoice = WaitForElementIsVisible(By.XPath(CUSTOMER_INVOICE));
            var dateInvoice = WaitForElementIsVisible(By.XPath(DATE_INVOICE));

            return ((invoiceNumber.Text.Contains(ID)) && (siteInvoice.Text.Contains(site)) && (customerInvoice.Text.Contains(customer)) && (dateInvoice.Text.Contains(date)));


        }

        public void ScrollToInvoiceStep()
        {
            var filter = _webDriver.FindElement(By.XPath(SHOW_INVOICE_STEP));
            ScrollToElement(filter);
        }
        public bool GetTotalTabs()
        {
            IList<string> allWindows = _webDriver.WindowHandles;

            return allWindows.Count > 1;
        }

        public bool AccessPage()
        {
            if (isElementExists(By.Id(PLUS_BUTTON)))
            {
                return true;
            }

            return false;
        }

        public bool IsManualInvoiceExists(string Id)
        {
            var invoice = _webDriver.FindElements(By.XPath(string.Format(INVOICELINE, Id)));

            return invoice.Count == 1;
        }

        public bool IsInvoiceDraft(string draft)
        {
            WaitPageLoading();

            var invoiceDraft = WaitForElementIsVisible(By.XPath(DRAFT_INVOICE));
            return invoiceDraft.Text.Contains(draft);
        }
    }
}
