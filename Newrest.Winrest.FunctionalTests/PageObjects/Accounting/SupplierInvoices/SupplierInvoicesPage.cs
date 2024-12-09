using DocumentFormat.OpenXml.Bibliography;
using FluentAssertions;
using iText.StyledXmlParser.Jsoup.Nodes;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Suppliers;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
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
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Accounting.SupplierInvoices
{
    public class SupplierInvoicesPage : PageBase
    {

        public SupplierInvoicesPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // __________________________________ Constantes ______________________________________________

        // Menu général

        private const string PLUS_BUTTON = "Plus-Button";
        private const string NEW_SUPPLIER_INVOICE = "NewSupplierInvoice";
        private const string EXTENDED_BUTTON = "Extended-Menu";
        private const string VALIDATE_RESULTS = "btn-validate-results";
        private const string CONFIRM_VALIDATE = "btn-popup-validate";
        private const string PRINT_BUTTON = "//*[@id=\"header-print-button\"]";
        private const string PRINT_POPUP = "//h3[text() = 'Print list']";
        private const string PRINT_RESULTS = "btn-print-supplier-invoices-report";
        private const string EXPORT = "btn-export-excel";
        private const string ENABLE_EXPORT_SAGE = "Enable-Export-Sage-For-Index";
        private const string MANUAL_EXPORT_SAGE = "btn-export-for-sage";
        private const string SEND_AUTO_TO_SAGE = "btn-export-for-sage";
        private const string GENERATE_SAGE_TXT = "btn-export-for-sage-txt-generation";
        private const string REINVOICE_IN_CUSTOMER_INVOICE = "btn-reinvoice-in-customer-invoice";
        private const string CONFIRM_EXPORT_SAGE = "btn-popup-validate";
        private const string DOWNLOAD = "btn-download-file";
        private const string CLOSE = "btn-cancel-popup";
        private const string GENERAL_INFORMATION = "hrefTabContentInformations";
        private const string VERIFY = "btnCreateInvoice";
        private const string GENERATE = "btnCreateInvoice";

        private const string CURL_SUPPLIER_INVOICE_NUMBER = "//*[@id=\"list-item-with-action\"]/table/tbody/tr/td[5]";

        // Tableau supplier invoices
        private const string INVOICE_AMOUNT_WITHOUT_TAX = "//*[@id=\"list-item-with-action\"]/table/thead/tr/th[text()='Invoice amount without tax']";
        private const string SUPPLIER_INVOICE_NUMBER = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[5]";
        private const string FIRST_SUPPLIER_INVOICE_ID_NUMBER = "//table/tbody/tr/td[3]";
        private const string FIRST_SUPPLIER_INVOICE_SITE = "//table/tbody/tr/td[4]";
        private const string FIRST_SUPPLIER_INVOICE_DATE = "//table/tbody/tr[1]/td[7]";
        private const string FIRST_SUPPLIER_INVOICE_SUPPLIER = "//table/tbody/tr[1]/td[8]";
        private const string FIRST_SUPPLIER_INVOICE_AMOUNT = "//table/tbody/tr[1]/td[10]";
        private const string FIRST_SUPPLIER_INVOICE_TAXES = "//table/tbody/tr[1]/td[11]";

        private const string SUPPLIER_NUMBER = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[3]";
        private const string SUPPLIER_INVOICE_DELETE_BUTTON = "//span[contains(@class, 'trash')]";
        private const string SUPPLIER_INVOICE_DELETE_CONFIRM_BUTTON = "dataConfirmOK";

        public const string CLAIM_ELEMENT = "//*[@id=\"div-body\"]/div/div/div[1]/div[2]";
        public const string DOTS_ELEMENT = "//*[@id=\"div-body\"]/div/div/div[1]/div[1]";
        // Filtres

        private const string RESET_FILTER = "Reset-Filter";
        private const string FILTER_SEARCH = "SearchNumber";
        private const string FILTER_SUPPLIER = "SelectedSuppliers_ms";
        private const string SUPPLIER_SEARCH = "/html/body/div[10]/div/div/label/input";
        private const string UNCHECK_ALL_SUPPLIER = "/html/body/div[10]/div/ul/li[2]/a";
        private const string FILTER_SORTBY = "cbSortBy";
        private const string FILTER_DATE_FROM = "date-picker-start";
        private const string FILTER_DATE_TO = "date-picker-end";
        private const string FILTER_ALL = "//*[@id=\"ShowActive\"][@value=\"All\"]";
        private const string FILTER_ACTIVE = "//*[@id=\"ShowActive\"][@value=\"ActiveOnly\"]";
        private const string FILTER_INACTIVE = "//*[@id=\"ShowActive\"][@value=\"InactiveOnly\"]";
        private const string FILTER_EDI = "ShowImportedWithEdi";
        private const string FILTER_SHOW_ALL_INVOICES = "radio-show-all";
        private const string FILTER_SHOW_VALIDATED = "radio-show-not-valid";
        private const string FILTER_SHOW_SENT_SAGE = "radio-show-invoiced";
        private const string FILTER_SHOW_SENT_SAGE_ERROR = "radio-show-inerror";
        private const string FILTER_SHOW_ON_TL_WAITING_SAGE_PUSH = "radio-show-inwaiting";
        private const string FILTER_SHOW_VALID_NOT_SENT_SAGE = "radio-show-valid-not-saged";
        private const string FILTER_SHOW_VERIFIED = "radio-show-verified";
        private const string FILTER_SHOW_NOT_VERIFIED = "radio-show-notVerified";
        private const string FILTER_SHOW_PARTIALLY_IMPORTED = "radio-show-partial-imported";
        private const string FILTER_SHOW_WITH_CLAIM = "radio-show-claims";
        private const string FILTER_SHOW_ONLY_CN = "ShowOnlyValue";
        private const string EXPORTED_FOR_SAGE_MANUALLY = "radio-show-sage-exported-manually";
        private const string FILTER_NOT_TRANSFORMED_INTO_CUSTOMER_INVOCES = "radio-show-not-transformed-into-customer-invoices";
        private const string FILTER_TRANSFORMED_INTO_CUSTOMER_INVOCES = "radio-show-transformed-into-customer-invoices";
        private const string FILTER_SITES = "SelectedSites_ms";
        private const string SEARCH_SITES = "/html/body/div[11]/div/div/label/input";
        private const string UNCHECK_ALL_SITES = "/html/body/div[11]/div/ul/li[2]/a";
        private const string FILTER_SITE_PLACES = "SelectedSitePlaces_ms";
        private const string SEARCH_SITE_PLACES = "/html/body/div[12]/div/div/label/input";
        private const string UNCHECK_ALL_SITE_PLACES = "/html/body/div[12]/div/ul/li[2]/a";
        private const string PIECE_ID_FROM = "SearchPieceIdFrom";
        private const string PIECE_ID_TO = "SearchPieceIdTo";

        private const string INVOICE_STATUS = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[{0}]/td[1]/a/img[@alt='Inactive']";
        private const string VALIDATION = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[{0}]/td[1]/a/img[@alt='Valid']";
        private const string SENT_TO_SAGE = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[{0}]/td[1]/a/span[@title='Accounted (sent to SAGE manually)']";
        private const string SENT_TO_SAGE_ERROR = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[{0}]/td[1]/a/span[contains(@title,'Error (SAGE/CEGID return an Error)') or contains(@title,'Error (SAGE return an Error)')]";//to be compliant for dev and patch Newrest.Winrest-2021.1229 (dec2021)
        private const string WAITING_FOR_SAGE_PUSH = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[{0}]/td[1]/a/span[contains(@title,'Waiting (SAGE treatment)') or contains(@title,'Waiting (SAGE/CEGID treatment)')]";//to be compliant for dev and patch Newrest.Winrest-2021.1229 (dec2021)
        private const string CLAIM_ASSOCIATED = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[{0}]/td[2]/a/i";
        private const string VERIFY_ID = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[3]";
        private const string VERIFY_NUMBER = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[4]";
        private const string SI_SITE_COLUMN = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[6]";
        private const string SI_DATE_COLUMN = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[7]";
        private const string SI_SUPPLIER_COLUMN = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[8]";

        private const string FIRST_LINE = "//*[@id=\"list-item-with-action\"]/table/tbody/tr/td[4]";
        private const string INVOICE_SUPPLIERS_NUMBERS = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[3]";
        private const string INVOICE_SUPPLIERS_TYPES = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[5]";
        private const string ICON_VALIDATE_SUPPLIERS = "//*[@id=\"accounting-supplierinvoice-details-1\"]/div/div";
        private const string NB_ROWS_VALIDATE_SUPPLIERS = "//*[@id=\"list-item-with-action\"]/table/tbody/tr";
        private const string INVOICE_NUMBER = "//*[@id=\"accounting-supplierinvoice-details-numberstring-1\"]";

        private const string SHOW = "//*[@id=\"formSearchSupplierInvoices\"]/div[11]/a";

        private const string INVOICE_SUP_TYPES = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[5]";
        private const string CIRCLE_INACTIVE = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[{0}]/td[1]/a/div/div[2]/i[@class='fa-solid fa-circle-xmark']";
        // __________________________________Variables______________________________________________

        // Menu général
        [FindsBy(How = How.Id, Using = PLUS_BUTTON)]
        private IWebElement _plusButton;

        [FindsBy(How = How.Id, Using = NEW_SUPPLIER_INVOICE)]
        private IWebElement _newSupplierInvoice;

        [FindsBy(How = How.Id, Using = EXTENDED_BUTTON)]
        private IWebElement _extendedButton;

        [FindsBy(How = How.Id, Using = VALIDATE_RESULTS)]
        private IWebElement _validateResults;

        [FindsBy(How = How.Id, Using = CONFIRM_VALIDATE)]
        private IWebElement _confirmValidate;

        [FindsBy(How = How.XPath, Using = PRINT_BUTTON)]
        private IWebElement _printButton;

        [FindsBy(How = How.Id, Using = PRINT_RESULTS)]
        private IWebElement _printResults;

        [FindsBy(How = How.Id, Using = EXPORT)]
        private IWebElement _export; 
        [FindsBy(How = How.Id, Using = REINVOICE_IN_CUSTOMER_INVOICE)]
        private IWebElement _reinvoiceInCustomerInvoice;

        [FindsBy(How = How.Id, Using = ENABLE_EXPORT_SAGE)]
        private IWebElement _enableExportForSage;

        [FindsBy(How = How.Id, Using = MANUAL_EXPORT_SAGE)]
        private IWebElement _manualExportSAGE;

        [FindsBy(How = How.Id, Using = SEND_AUTO_TO_SAGE)]
        private IWebElement _sendAutoToSage;

        [FindsBy(How = How.Id, Using = GENERATE_SAGE_TXT)]
        private IWebElement _generateSageTxt;

        [FindsBy(How = How.Id, Using = CONFIRM_EXPORT_SAGE)]
        private IWebElement _confirmExportSage;

        [FindsBy(How = How.Id, Using = DOWNLOAD)]
        private IWebElement _downloadBtn;  
        [FindsBy(How = How.Id, Using = GENERATE)]
        private IWebElement _generate;

        [FindsBy(How = How.Id, Using = CLOSE)]
        private IWebElement _close;

        [FindsBy(How = How.Id, Using = GENERAL_INFORMATION)]
        private IWebElement _generalInformation;
        [FindsBy(How = How.Id, Using = VERIFY)]
        private IWebElement _verify;

        // Tableau supplier invoices
        [FindsBy(How = How.XPath, Using = SUPPLIER_INVOICE_NUMBER)]
        private IWebElement _supplierInvoiceNumber;

        [FindsBy(How = How.XPath, Using = FIRST_SUPPLIER_INVOICE_ID_NUMBER)]
        private IWebElement _supplierInvoiceIdNumber;

        [FindsBy(How = How.XPath, Using = FIRST_SUPPLIER_INVOICE_SITE)]
        private IWebElement _supplierInvoiceSite;

        [FindsBy(How = How.XPath, Using = FIRST_SUPPLIER_INVOICE_DATE)]
        private IWebElement _supplierInvoiceDate;

        [FindsBy(How = How.XPath, Using = FIRST_SUPPLIER_INVOICE_SUPPLIER)]
        private IWebElement _supplierInvoiceSupplier;

        [FindsBy(How = How.XPath, Using = FIRST_SUPPLIER_INVOICE_AMOUNT)]
        private IWebElement _supplierInvoiceSupplierAmount;

        [FindsBy(How = How.XPath, Using = FIRST_SUPPLIER_INVOICE_TAXES)]
        private IWebElement _supplierInvoiceSupplierTaxes;

        [FindsBy(How = How.XPath, Using = SUPPLIER_INVOICE_DELETE_BUTTON)]
        private IWebElement _deleteButton;

        [FindsBy(How = How.Id, Using = SUPPLIER_INVOICE_DELETE_CONFIRM_BUTTON)]
        private IWebElement _deleteConfirmButton;

        //__________________________________Filters_______________________________________
        [FindsBy(How = How.Id, Using = RESET_FILTER)]
        private IWebElement _resetFilter;

        [FindsBy(How = How.Id, Using = FILTER_SEARCH)]
        private IWebElement _searchFilter;

        [FindsBy(How = How.Id, Using = FILTER_SUPPLIER)]
        private IWebElement _supplierFilter;

        [FindsBy(How = How.XPath, Using = SUPPLIER_SEARCH)]
        private IWebElement _searchSupplier;

        [FindsBy(How = How.XPath, Using = UNCHECK_ALL_SUPPLIER)]
        private IWebElement _uncheckAllSupplier;

        [FindsBy(How = How.Id, Using = FILTER_SORTBY)]
        private IWebElement _sortBy;

        [FindsBy(How = How.Id, Using = FILTER_DATE_FROM)]
        public IWebElement _dateFrom;

        [FindsBy(How = How.Id, Using = FILTER_DATE_TO)]
        public IWebElement _dateTo;

        [FindsBy(How = How.XPath, Using = FILTER_ALL)]
        private IWebElement _showAll;

        [FindsBy(How = How.XPath, Using = FILTER_ACTIVE)]
        private IWebElement _showOnlyActive;

        [FindsBy(How = How.XPath, Using = FILTER_INACTIVE)]
        private IWebElement _showOnlyInactive;

        [FindsBy(How = How.Id, Using = FILTER_EDI)]
        private IWebElement _showEDI;

        [FindsBy(How = How.Id, Using = FILTER_SHOW_ALL_INVOICES)]
        private IWebElement _showAllInvoices;

        [FindsBy(How = How.Id, Using = FILTER_SHOW_VALIDATED)]
        private IWebElement _showNotValidatedOnly;

        [FindsBy(How = How.Id, Using = FILTER_SHOW_SENT_SAGE)]
        private IWebElement _showSentToSageOnly;

        [FindsBy(How = How.Id, Using = FILTER_SHOW_SENT_SAGE_ERROR)]
        private IWebElement _showSentToSageInErrorOnly;

        [FindsBy(How = How.Id, Using = FILTER_SHOW_ON_TL_WAITING_SAGE_PUSH)]
        private IWebElement _showOnTlWaitingSagePush;

        [FindsBy(How = How.Id, Using = FILTER_SHOW_VALID_NOT_SENT_SAGE)]
        private IWebElement _showValidatedNotSentSage;

        [FindsBy(How = How.Id, Using = FILTER_SHOW_VERIFIED)]
        private IWebElement _showVerifiedOnly;

        [FindsBy(How = How.Id, Using = FILTER_SHOW_NOT_VERIFIED)]
        private IWebElement _showNotVerifiedOnly;

        [FindsBy(How = How.Id, Using = FILTER_SHOW_PARTIALLY_IMPORTED)]
        private IWebElement _showPartiallyImported;

        [FindsBy(How = How.Id, Using = FILTER_SHOW_WITH_CLAIM)]
        private IWebElement _showSupplierWithClaim;

        [FindsBy(How = How.Id, Using = EXPORTED_FOR_SAGE_MANUALLY)]
        private IWebElement _exportedForSageManually;
        [FindsBy(How = How.Id, Using = FILTER_NOT_TRANSFORMED_INTO_CUSTOMER_INVOCES)]
        private IWebElement _nottransformedIntoCustomerInvoices;

        [FindsBy(How = How.Id, Using = FILTER_TRANSFORMED_INTO_CUSTOMER_INVOCES)]
        private IWebElement _transformedIntoCustomerInvoices;
        [FindsBy(How = How.Id, Using = FILTER_SITES)]
        private IWebElement _siteFilter;

        [FindsBy(How = How.XPath, Using = SEARCH_SITES)]
        private IWebElement _searchSite;

        [FindsBy(How = How.XPath, Using = UNCHECK_ALL_SITES)]
        private IWebElement _uncheckAllSites;

        [FindsBy(How = How.Id, Using = FILTER_SITE_PLACES)]
        private IWebElement _sitePlacesFilter;

        [FindsBy(How = How.XPath, Using = SEARCH_SITE_PLACES)]
        private IWebElement _searchSitePlaces;

        [FindsBy(How = How.XPath, Using = UNCHECK_ALL_SITE_PLACES)]
        private IWebElement _uncheckAllSitePlaces;

        [FindsBy(How = How.XPath, Using = ICON_VALIDATE_SUPPLIERS)]
        private IWebElement _iconValidateSuppliers;
        [FindsBy(How = How.XPath, Using = SHOW)]
        private IWebElement _show;
        public enum FilterType
        {
            ByNumber,
            Suppliers,
            SortBy,
            DateFrom,
            DateTo,
            ShowAll,
            ShowOnlyActive,
            ShowOnlyInactive,
            ShowEDI,
            ShowAllInvoices,
            ShowNotValidatedOnly,
            ShowSentToSageOnly,
            ShowSentToSageErrorOnly,
            ShowOnTLWaitingForSagePush,
            ShowValidatedNotSentSage,
            ShowVerifiedOnly,
            ShowNotVerifiedOnly,
            ShowPartiallyOnly,
            ShowSupplierInvoiceWithClaim,
            ExportedForSageManually,
            Site,
            SitePlaces,
            PieceIdFrom,
            PieceIdTo,
            ShowOnlyCN,
            ShowValidatedOnly,
            NottransformedIntoCustomerInvoices,
            TransformedIntoCustomerInvoices
        }

        public void Filter(FilterType filterType, object value)
        {
            Actions actions = new Actions(_webDriver);

            switch (filterType)
            {
                case FilterType.ByNumber:
                    PageUpSupplierInvoice();
                    _searchFilter = WaitForElementIsVisible(By.Id(FILTER_SEARCH));
                    _searchFilter.SetValue(ControlType.TextBox, value);
                    WaitForLoad();
                    break;
                case FilterType.Suppliers:
                    PageUpSupplierInvoice();
                    _supplierFilter = WaitForElementIsVisible(By.Id(FILTER_SUPPLIER));
                    _supplierFilter.Click();
                    WaitForLoad();
                    _uncheckAllSupplier = WaitForElementIsVisible(By.XPath(UNCHECK_ALL_SUPPLIER));
                    _uncheckAllSupplier.Click();
                    WaitForLoad();

                    _searchSupplier = WaitForElementIsVisible(By.XPath(SUPPLIER_SEARCH));
                    _searchSupplier.SetValue(ControlType.TextBox, value);

                    var searchSupplierValue = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    searchSupplierValue.SetValue(ControlType.CheckBox, true);

                    _supplierFilter.Click();
                    WaitForLoad();
                    break;
                case FilterType.SortBy:
                    PageUpSupplierInvoice();
                    _sortBy = WaitForElementIsVisible(By.Id(FILTER_SORTBY));
                    _sortBy.Click();
                    var element = WaitForElementIsVisible(By.XPath("//option[contains(@value,'" + value + "')]"));
                    _sortBy.SetValue(ControlType.DropDownList, element.Text);
                    _sortBy.Click();
                    WaitForLoad();

                    break;
                case FilterType.DateFrom:
                    PageUpSupplierInvoice();
                    _dateFrom = WaitForElementIsVisible(By.Id(FILTER_DATE_FROM));
                    _dateFrom.SetValue(ControlType.DateTime, value);
                    _dateFrom.SendKeys(Keys.Tab);
                    WaitForLoad();

                    break;
                case FilterType.DateTo:
                    PageUpSupplierInvoice();
                    _dateTo = WaitForElementIsVisible(By.Id(FILTER_DATE_TO));
                    _dateTo.SetValue(ControlType.DateTime, value);
                    _dateTo.SendKeys(Keys.Tab);
                    WaitForLoad();

                    break;
                case FilterType.ShowAll:
                    PageUpSupplierInvoice();
                    _showAll = WaitForElementExists(By.XPath(FILTER_ALL));
                    _showAll.SetValue(ControlType.RadioButton, value);
                    WaitForLoad();

                    break;
                case FilterType.ShowOnlyActive:
                    PageUpSupplierInvoice();
                    _showOnlyActive = WaitForElementExists(By.XPath(FILTER_ACTIVE));
                    _showOnlyActive.SetValue(ControlType.RadioButton, value);
                    WaitForLoad();

                    break;
                case FilterType.ShowOnlyInactive:
                    PageUpSupplierInvoice();
                    _showOnlyInactive = WaitForElementExists(By.XPath(FILTER_INACTIVE));
                    _showOnlyInactive.SetValue(ControlType.RadioButton, value);
                    WaitForLoad();

                    break;
                case FilterType.ShowEDI:
                    PageUpSupplierInvoice();
                    _showEDI = WaitForElementExists(By.Id(FILTER_EDI));
                    _showEDI.SetValue(ControlType.CheckBox, value);
                    WaitForLoad();

                    break;
                case FilterType.ShowAllInvoices:
                    Show();
                    _showAllInvoices = WaitForElementExists(By.Id(FILTER_SHOW_ALL_INVOICES));
                    actions.MoveToElement(_showAllInvoices).Click().Perform();
                    _showAllInvoices.SetValue(ControlType.RadioButton, value);
                    WaitForLoad();

                    break;
                case FilterType.ShowNotValidatedOnly:
                    Show();
                    _showNotValidatedOnly = WaitForElementIsVisible(By.Id(FILTER_SHOW_VALIDATED));
                    actions.MoveToElement(_showNotValidatedOnly).Click().Perform();
                    _showNotValidatedOnly.SetValue(ControlType.RadioButton, value);
                    

                    break;
                case FilterType.ShowSentToSageOnly:
                    Show();
                    _showSentToSageOnly = WaitForElementExists(By.Id(FILTER_SHOW_SENT_SAGE));
                    actions.MoveToElement(_showSentToSageOnly).Click().Perform();
                    _showSentToSageOnly.SetValue(ControlType.RadioButton, value);
                    WaitForLoad();

                    break;
                case FilterType.ShowSentToSageErrorOnly:
                    Show();
                    _showSentToSageInErrorOnly = WaitForElementExists(By.Id(FILTER_SHOW_SENT_SAGE_ERROR));
                    actions.MoveToElement(_showSentToSageInErrorOnly).Click().Perform();
                    _showSentToSageInErrorOnly.SetValue(ControlType.RadioButton, value);
                    WaitForLoad();

                    break;
                case FilterType.ShowOnTLWaitingForSagePush:
                    Show();
                    _showOnTlWaitingSagePush = WaitForElementExists(By.Id(FILTER_SHOW_ON_TL_WAITING_SAGE_PUSH));
                    actions.MoveToElement(_showOnTlWaitingSagePush).Click().Perform();
                    _showOnTlWaitingSagePush.SetValue(ControlType.RadioButton, value);
                    WaitForLoad();

                    break;
                case FilterType.ShowValidatedNotSentSage:
                    Show();
                    _showValidatedNotSentSage = WaitForElementExists(By.Id(FILTER_SHOW_VALID_NOT_SENT_SAGE));
                    actions.MoveToElement(_showValidatedNotSentSage).Click().Perform();
                    _showValidatedNotSentSage.SetValue(ControlType.RadioButton, value);
                    WaitForLoad();

                    break;
                case FilterType.ShowVerifiedOnly:
                    Show();
                    _showVerifiedOnly = WaitForElementExists(By.Id(FILTER_SHOW_VERIFIED));
                    actions.MoveToElement(_showVerifiedOnly).Click().Perform();
                    _showVerifiedOnly.SetValue(ControlType.RadioButton, value);
                    WaitForLoad();

                    break;
                case FilterType.ShowNotVerifiedOnly:
                    Show();
                    _showNotVerifiedOnly = WaitForElementExists(By.Id(FILTER_SHOW_NOT_VERIFIED));
                    actions.MoveToElement(_showNotVerifiedOnly).Click().Perform();
                    _showNotVerifiedOnly.SetValue(ControlType.RadioButton, value);
                    WaitForLoad();

                    break;
                case FilterType.ShowPartiallyOnly:
                    Show();
                    _showPartiallyImported = WaitForElementExists(By.Id(FILTER_SHOW_PARTIALLY_IMPORTED));
                    actions.MoveToElement(_showPartiallyImported).Click().Perform();
                    _showPartiallyImported.SetValue(ControlType.RadioButton, value);
                    WaitForLoad();

                    break;
                case FilterType.ShowSupplierInvoiceWithClaim:
                    Show();
                    _showSupplierWithClaim = WaitForElementExists(By.Id(FILTER_SHOW_WITH_CLAIM));
                    actions.MoveToElement(_showSupplierWithClaim).Click().Perform();
                    _showSupplierWithClaim.SetValue(ControlType.RadioButton, value);
                    WaitForLoad();

                    break;
                case FilterType.ExportedForSageManually:
                    Show();
                    _exportedForSageManually = WaitForElementExists(By.Id(EXPORTED_FOR_SAGE_MANUALLY));
                    actions.MoveToElement(_exportedForSageManually).Click().Perform();
                    _exportedForSageManually.SetValue(ControlType.RadioButton, value);
                    WaitForLoad();

                    break;
                case FilterType.NottransformedIntoCustomerInvoices:
                    Show();
                    _nottransformedIntoCustomerInvoices = WaitForElementExists(By.Id(FILTER_NOT_TRANSFORMED_INTO_CUSTOMER_INVOCES));
                    actions.MoveToElement(_nottransformedIntoCustomerInvoices).Click().Perform();
                    _nottransformedIntoCustomerInvoices.SetValue(ControlType.RadioButton, value);
                    WaitForLoad();

                    break;

                case FilterType.TransformedIntoCustomerInvoices:
                    Show();
                    _transformedIntoCustomerInvoices = WaitForElementExists(By.Id(FILTER_TRANSFORMED_INTO_CUSTOMER_INVOCES));
                    actions.MoveToElement(_transformedIntoCustomerInvoices).Click().Perform();
                    _transformedIntoCustomerInvoices.SetValue(ControlType.RadioButton, value);
                    WaitForLoad();

                    break;

                case FilterType.Site:
                 
                    _siteFilter = WaitForElementExists(By.Id(FILTER_SITES));
                    ScrollToElement(_siteFilter);
                    _siteFilter.Click();
                    _uncheckAllSites = _webDriver.FindElement(By.XPath(UNCHECK_ALL_SITES));
                    _uncheckAllSites.Click();

                    _searchSite = WaitForElementIsVisible(By.XPath(SEARCH_SITES));
                    _searchSite.SetValue(ControlType.TextBox, value);

                    var searchSiteValue = WaitForElementIsVisible(By.XPath("//span[text()='" + value + " - " + value + "']"));
                    searchSiteValue.SetValue(ControlType.CheckBox, true);

                    _siteFilter.Click();
                    WaitForLoad();

                    break;
                case FilterType.SitePlaces:
                    _sitePlacesFilter = WaitForElementExists(By.Id(FILTER_SITE_PLACES));
                    ScrollToElement(_sitePlacesFilter);
                    ComboBoxSelectById(new ComboBoxOptions(FILTER_SITE_PLACES, (string)value, false));

                    WaitForLoad();

                    break;
                case FilterType.PieceIdFrom:
                    PageUpSupplierInvoice();
                    var pieceIdFrom = WaitForElementIsVisible(By.Id(PIECE_ID_FROM));
                    pieceIdFrom.SendKeys(value.ToString());
                    WaitForLoad();

                    break;
                case FilterType.PieceIdTo:
                    PageUpSupplierInvoice();
                    var pieceIdTo = WaitForElementIsVisible(By.Id(PIECE_ID_TO));
                    pieceIdTo.SendKeys(value.ToString());
                    WaitForLoad();

                    break;
                case FilterType.ShowOnlyCN:
                    Show();
                    var filterShowOnlyCN = WaitForElementExists(By.Id(FILTER_SHOW_ONLY_CN));
                    actions.MoveToElement(filterShowOnlyCN).Click().Perform();
                    filterShowOnlyCN.SetValue(ControlType.RadioButton, value);
                    WaitForLoad();

                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);
            }

            WaitPageLoading();
        }

        public void ResetFilter()
        {
            _resetFilter = WaitForElementExists(By.Id(RESET_FILTER));
            new Actions(_webDriver).MoveToElement(_resetFilter).Click().Perform();

            WaitForLoad();

            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                Filter(FilterType.DateTo, DateUtils.Now);
            }
        }

        public string GetSearchFilterValue()
        {
            _searchFilter = WaitForElementIsVisible(By.Id(FILTER_SEARCH));
            return _searchFilter.GetAttribute("value");
        }

        public bool VerifySupplier(string supplierName)
        {
            ReadOnlyCollection<IWebElement> suppliers;
            suppliers = _webDriver.FindElements(By.XPath("//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[8]"));

            if (suppliers.Count == 0)
                return false;

            foreach (var elm in suppliers)
            {
                if (elm.Text != supplierName)
                {
                    return false;
                }
            }
            return true;
        }

        public bool VerifySuppliersInvoicesNumbersInRange(string pieceIdFrom, string pieceIdTo)
        {
            var suppliersNumbers = _webDriver.FindElements(By.XPath(SUPPLIER_NUMBER));
            foreach (var supplierNumber in suppliersNumbers)
            {

                if (int.Parse(supplierNumber.Text) > int.Parse(pieceIdTo) || int.Parse(supplierNumber.Text) < int.Parse(pieceIdFrom))
                {
                    return false;
                }
            }
            return true;
        }
        public bool VerifySite(string site)
        {
            ReadOnlyCollection<IWebElement> sites;
            sites = _webDriver.FindElements(By.XPath("//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[6]"));

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

        public bool IsSortedById(int pageSize)
        {
            List<IWebElement> listIds = new List<IWebElement> ();
            for (int i=1;i<=pageSize;i++)
            {
                var element = WaitForElementIsVisible(By.Id($"accounting-supplierinvoice-details-id-{i}"));
                listIds.Add(element);
            }
            if (listIds.Count == 0)
                return false;

            var ancientId = int.Parse(listIds[0].Text);
            foreach (var elm in listIds)
            {
                try
                {
                    if (int.Parse(elm.Text) > ancientId)
                        return false;

                    ancientId = int.Parse(elm.Text);
                }
                catch
                {
                    return false;
                }
            }

            return true;
        }

        public bool IsSortedByDate(string dateFormat,int pageSize)
        {
            // Take the date format from the datepicker element and use it to format the date column to avoid date errors
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");
            List<IWebElement> dates = new List<IWebElement>();
            for(int i = 1; i <= pageSize;i++)
            {
                var element = WaitForElementIsVisible(By.Id($"accounting-supplierinvoice-details-invoicedate-{i}"));
                dates.Add(element);
            }
            if (dates.Count() == 0)
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

        public bool IsSortedByNumber(int pageSize)
        {
            List<IWebElement> numbers = new List<IWebElement>();
            for(int i = 1; i <= pageSize;i++)
            {
                var element = WaitForElementIsVisible(By.Id($"accounting-supplierinvoice-details-numberstring-{i}"));
                numbers.Add(element);
            }

            if (numbers.Count == 0)
                return false;

            var ancientNumber = numbers[0].Text;

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

        public bool IsDateRespected(DateTime fromDate, DateTime toDate, string dateFormat)
        {
            // Take the date format from the datepicker element and use it to format the date column to avoid date errors
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");
            ReadOnlyCollection<IWebElement> dates;
            dates = _webDriver.FindElements(By.XPath("//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[7]"));

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

        public bool CheckStatus(bool active)
        {
            bool isActive = false;
            int tot = CheckTotalNumber() > 100 ? 100 : CheckTotalNumber();
            PageSize("100");
            if (tot == 0)
                return false;

            for (int i = 0; i < tot; i++)
            {
                if(isElementVisible(By.XPath(String.Format(CIRCLE_INACTIVE, i + 1))))
                { 
                    _webDriver.FindElement(By.XPath(String.Format(CIRCLE_INACTIVE, i + 1)));

                    if (active)
                        return false;
                }
                else
                {
                    isActive = true;
                    if (!active)
                        return true;
                }
            }
            return isActive;
        }

        public bool CheckValidation(bool validated)
        {
            bool isValidated = true;
            int tot = CheckTotalNumber() > 100 ? 100 : CheckTotalNumber();

            if (tot == 0)
                return false;

            for (int i = 0; i < tot; i++)
            {

                if (IsDev()) {
                    if (isElementVisible(By.XPath(string.Format("//*[@id=\"list-item-with-action\"]/table/tbody/tr[{0}]/td[1]/a/div/i[contains(@class, \"fa-regular fa-circle-check\")]", i + 1))))
                    {
                        _webDriver.FindElement(By.XPath(string.Format("//*[@id=\"list-item-with-action\"]/table/tbody/tr[{0}]/td[1]/a/div/i[contains(@class, \"fa-regular fa-circle-check\")]\r\n", i + 1)));
                        if (!validated)
                            return true;
                    }
                    else
                    {
                        if (isElementVisible(By.XPath(string.Format(VALIDATION, i + 1))))
                        {
                            _webDriver.FindElement(By.XPath(string.Format(VALIDATION, i + 1)));
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

                }
                else
                {
                    if (isElementVisible(By.XPath(string.Format("//*[@id=\"list-item-with-action\"]/table/tbody/tr[{0}]/td[1]/a/div/div/i[contains(@class, \"fa-regular fa-circle-check\")]", i + 1))))
                    {
                        // on a l'icone Validation vert à gauche
                        _webDriver.FindElement(By.XPath(string.Format("//*[@id=\"list-item-with-action\"]/table/tbody/tr[{0}]/td[1]/a/div/div/i[contains(@class, \"fa-regular fa-circle-check\")]\r\n", i + 1)));
                        
                        if (!validated)
                            return false;
                    }
                    else
                    {
                        // on n'a pas l'icone Validation vert à gauche
                        if (validated)
                            return false;
                    }
                }


            }

            return isValidated;
        }

        public bool IsSentToSAGE()
        {
            int tot = CheckTotalNumber() > 100 ? 100 : CheckTotalNumber();

            if (tot == 0)
                return false;

            for (int i = 0; i < tot; i++)
            {
                if(isElementVisible(By.XPath(string.Format(SENT_TO_SAGE, i + 1))))
                {
                    _webDriver.FindElement(By.XPath(string.Format(SENT_TO_SAGE, i + 1)));
                }
                else
                {
                    return false;
                }
            }
            return true;

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

        public bool IsWithClaim()
        {
            int tot = CheckTotalNumber() > 100 ? 100 : CheckTotalNumber();

            if (tot == 0)
                return false;
            string ClaimAssociatedXPath;
            
             ClaimAssociatedXPath = CLAIM_ASSOCIATED;
            for (int i = 0; i < tot; i++)
            {
                if (isElementVisible(By.XPath(string.Format(ClaimAssociatedXPath, i + 1))))
                {
                    _webDriver.FindElement(By.XPath(string.Format(ClaimAssociatedXPath, i + 1)));
                }
                else
                {
                    return false;
                }
            }
            return true;

        }

        public void ClickOnFirstline()
        {
            var line = WaitForElementIsVisible(By.XPath(FIRST_LINE));
            line.Click();
        }

        public void VerifyDotsArePhysicallyToRightOfClaim()
        {
            var dotsElement = WaitForElementIsVisible(By.XPath(DOTS_ELEMENT));

            // Vérification que les "..." sont bien visibles
            Assert.IsTrue(dotsElement.Displayed, "'...' n'est pas visible.");

            // Attendre un peu pour s'assurer que l'élément est complètement chargé
            Thread.Sleep(1000);

            // Récupérer la position de l'élément "..."
            var dotsPosition = dotsElement.Location;

            // Debug: Afficher les valeurs de position pour aider au débogage
            Console.WriteLine($"dotsPosition.X: {dotsPosition.X}");

            // Vérifier que les "..." sont physiquement à droite de la page (X > 0)
            Assert.IsTrue(dotsPosition.X > 0, "Les '...' ne sont pas physiquement à droite de la page.");
        }




        // _________________________________ Méthodes ______________________________________

        // Menu général
        public override void ShowPlusMenu()
        {
            var actions = new Actions(_webDriver);

            _plusButton = WaitForElementIsVisible(By.Id(PLUS_BUTTON));
            actions.MoveToElement(_plusButton).Perform();

        }

        public SupplierInvoicesCreateModalPage SupplierInvoicesCreatePage()
        {
            ShowPlusMenu();

            _newSupplierInvoice = WaitForElementIsVisible(By.Id(NEW_SUPPLIER_INVOICE));
            _newSupplierInvoice.Click();

            WaitForLoad();

            return new SupplierInvoicesCreateModalPage(_webDriver, _testContext);
        }

        public override void ShowExtendedMenu()
        {
            var actions = new Actions(_webDriver);

            _extendedButton = WaitForElementIsVisible(By.Id(EXTENDED_BUTTON));
            actions.MoveToElement(_extendedButton).Perform();

            WaitForLoad();
        }

        public void ValidateResults()
        {
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;

            ShowExtendedMenu();

            // Click sur le bouton validateResults
            _validateResults = WaitForElementIsVisible(By.Id(VALIDATE_RESULTS));
            _validateResults.Click();
            WaitForLoad();

            _confirmValidate = WaitForElementExists(By.Id(CONFIRM_VALIDATE));
            javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", _confirmValidate);
            WaitForLoad();

            _confirmValidate.Click();
            WaitForLoad();
        }

        public PrintReportPage PrintSupplierInvoices(bool printValue)
        {
            ShowExtendedMenu();

            _printResults = WaitForElementIsVisible(By.Id(PRINT_RESULTS));
            _printResults.Click();
            WaitForLoad();

            if (printValue)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-pdf']"));
            }

            // Results are opened in a new tab, switch the driver to the newly created one
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new PrintReportPage(_webDriver, _testContext);
        }

        public void ExportExcelFile(bool printValue)
        {
            // on clique sur le bouton impression
            ClosePrintButton();
            PageUp();

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
            // "supplier-invoices 2019-12-10 09-20-37.xlsx"

            string mois = "(?:0[1-9]|1[0-2])";         // mois
            string space = "(\\s)";                    // Espace
            string annee = "\\d{4}";                   // annee YYYY
            string jour = "[0-3]\\d";                  // jour
            string heure = "(?:0[0-9]|1[0-9]|2[0-3])"; // heure
            string minutes = "[0-5]\\d";               // minutes
            string secondes = "[0-5]\\d";              // secondes


            Regex r = new Regex("^supplier-invoices" + space + annee + "-" + mois + "-" + jour + space + heure + "-" + minutes + "-" + secondes + ".xlsx$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = r.Match(filePath);

            return m.Success;
        }

        public void EnableExportForSage()
        {
            ShowExtendedMenu();

            _enableExportForSage = WaitForElementIsVisible(By.Id(ENABLE_EXPORT_SAGE));
            _enableExportForSage.Click();

            WaitForLoad();
        }

        public string ManualExportSageError(bool printValue, bool isMessage)
        {
            string errorMessage = "";
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;

            ShowExtendedMenu();

            _manualExportSAGE = WaitForElementIsVisible(By.Id(MANUAL_EXPORT_SAGE));
            _manualExportSAGE.Click();
            WaitForLoad();

            _confirmExportSage = WaitForElementExists(By.Id(CONFIRM_EXPORT_SAGE));
            javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", _confirmExportSage);
            WaitForLoad();

            _confirmExportSage.Click();
            WaitForLoad();

            if (!printValue)
            {
                if(isElementVisible(By.Id(CONFIRM_EXPORT_SAGE)))
                {
                    _confirmExportSage = _webDriver.FindElement(By.Id(CONFIRM_EXPORT_SAGE));
                    _confirmExportSage.Click();
                    WaitForLoad();

                    // On ferme la pop-up
                    _close = WaitForElementIsVisible(By.Id(CLOSE));
                    _close.Click();
                    WaitForLoad();
                }
                else
                {
                    _close = WaitForElementIsVisible(By.Id(CLOSE));
                    _close.Click();
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

        public void ManualExportSage(bool printValue)
        {
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;

            ShowExtendedMenu();

            _manualExportSAGE = WaitForElementIsVisible(By.Id(MANUAL_EXPORT_SAGE));
            _manualExportSAGE.Click();
            WaitForLoad();

            _confirmExportSage = WaitForElementExists(By.Id(CONFIRM_EXPORT_SAGE));
            javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", _confirmExportSage);
            WaitForLoad();
            _confirmExportSage = WaitForElementIsVisible(By.Id(CONFIRM_EXPORT_SAGE));
            _confirmExportSage.Click();
            WaitForLoad();

            if (!printValue)
            {
                if (isElementVisible(By.Id(DOWNLOAD)))
                {
                    _downloadBtn = _webDriver.FindElement(By.Id(DOWNLOAD));
                    _downloadBtn.Click();
                    WaitForLoad();

                    // On ferme la pop-up
                    _close = WaitForElementIsVisible(By.Id(CLOSE));
                    _close.Click();
                    WaitForLoad();
                }
                else
                {
                    _close = WaitForElementIsVisible(By.Id(CLOSE));
                    _close.Click();
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

        public void SendAutoToSage()
        {
            ShowExtendedMenu();


            _sendAutoToSage = WaitForElementIsVisible(By.Id(SEND_AUTO_TO_SAGE));
            _sendAutoToSage.Click();

            WaitForLoad();

            _confirmExportSage = WaitForElementExists(By.Id(CONFIRM_EXPORT_SAGE));
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", _confirmExportSage);
            WaitForLoad();

            _confirmExportSage.Click();
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

            _confirmExportSage = WaitForElementExists(By.Id(CONFIRM_EXPORT_SAGE));
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", _confirmExportSage);
            WaitForLoad();

            _confirmExportSage.Click();
            WaitForLoad();

            IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-alt']"));
            ClickPrintButton();

            WaitForDownload();
            Close();
        }

        public FileInfo GetExportSAGEFile(FileInfo[] taskFiles)
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
            // "supplier-invoices 2019-12-10 09-20-37.txt"

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

        // Tableau supplier invoices
        public SupplierInvoicesItem SelectFirstSupplierInvoice()
        {
            _supplierInvoiceNumber = WaitForElementIsVisible(By.XPath(SUPPLIER_INVOICE_NUMBER));
            _supplierInvoiceNumber.Click();
            WaitForLoad();

            return new SupplierInvoicesItem(_webDriver, _testContext);
        }

        public bool IsInvoiceAmountWithoutTaxPresent()
        {
            if (isElementVisible(By.XPath(INVOICE_AMOUNT_WITHOUT_TAX)))
            {
                _webDriver.FindElement(By.XPath(INVOICE_AMOUNT_WITHOUT_TAX));
                return true;
            }
            else
            {
                return false;
            }
        }
        public IEnumerable<String> GetAllInvoiceNumber()
        {
            return _webDriver.FindElements(By.XPath(CURL_SUPPLIER_INVOICE_NUMBER)).Select(e => e.Text).ToList();
        }

        public bool CheckIfEdiInvoice(List<string> invoices)
        {
            foreach(var invoice in invoices)
            {
                if (!string.IsNullOrWhiteSpace(invoice))
                {
                    if (!(invoice.StartsWith("EDI"))){
                        return false;
                    }
                }
            }
            return true;
        }

        public string GetFirstInvoiceNumber()
        {
           _supplierInvoiceNumber = WaitForElementIsVisible(By.XPath(INVOICE_NUMBER));
            return _supplierInvoiceNumber.Text;
        }
        public string GetFirstSIIdNumber()
        {
            _supplierInvoiceIdNumber = WaitForElementIsVisible(By.XPath(FIRST_SUPPLIER_INVOICE_ID_NUMBER));
            return _supplierInvoiceIdNumber.Text;
        }
        public string GetFirstSISite()
        {
            _supplierInvoiceSite = WaitForElementIsVisible(By.XPath("//table/tbody/tr/td[6]"));
            return _supplierInvoiceSite.Text;
        }
        public string GetFirstSIDate()
        {
            _supplierInvoiceDate = WaitForElementIsVisible(By.XPath(FIRST_SUPPLIER_INVOICE_DATE));
            
            return _supplierInvoiceDate.Text;
        }
        public string GetFirstSISupplier()
        {
            _supplierInvoiceSupplier = WaitForElementIsVisible(By.XPath(FIRST_SUPPLIER_INVOICE_SUPPLIER));
            
            return _supplierInvoiceSupplier.Text;
        }
        public string GetFirstSIInvoiceAmountToBePaid()
        {
            _supplierInvoiceSupplierAmount = WaitForElementIsVisible(By.XPath(FIRST_SUPPLIER_INVOICE_AMOUNT));
            
            return _supplierInvoiceSupplierAmount.Text;
        }
        public string GetFirstSIInvoiceTotalInclTaxes()
        {
            _supplierInvoiceSupplierTaxes = WaitForElementIsVisible(By.XPath(FIRST_SUPPLIER_INVOICE_TAXES));
            
            return _supplierInvoiceSupplierTaxes.Text;
        }
        public void DeleteFirstSI()
        {
            _deleteButton = WaitForElementIsVisible(By.XPath(SUPPLIER_INVOICE_DELETE_BUTTON), nameof(SUPPLIER_INVOICE_DELETE_BUTTON));
            WaitPageLoading();
            _deleteButton.Click();
            WaitPageLoading();

            _deleteConfirmButton = WaitForElementIsVisible(By.Id(SUPPLIER_INVOICE_DELETE_CONFIRM_BUTTON), nameof(SUPPLIER_INVOICE_DELETE_CONFIRM_BUTTON));
            _deleteConfirmButton.Click();
            WaitPageLoading();
        }

        public void CheckPrint(FileInfo filePdf)
        {
            int count = Math.Min(CheckTotalNumber(), 100);
            PdfDocument document = PdfDocument.Open(filePdf.FullName);
            List<string> mots = new List<string>();
            foreach (Page p in document.GetPages())
            {
                foreach (var mot in p.GetWords())
                {
                    mots.Add(mot.Text);
                }
            }
            ReadOnlyCollection<IWebElement> numberStrings;
            numberStrings = _webDriver.FindElements(By.XPath("//*/tr/td[4]/a"));
            
            int counter = 0;
            foreach (var numberString in numberStrings)
            {
                if (numberString.Text == "") continue;
                counter++;
                Assert.IsTrue(mots.Contains(numberString.Text), "number string " + numberString.Text + " non présent dans le PDF");
            }
            Assert.AreEqual(count, counter, "Différence entre les SupplierInvoices number string du tableau et du PDF");

            ReadOnlyCollection<IWebElement> Ids;
            Ids = _webDriver.FindElements(By.XPath("//*/tr/td[4]/a"));
            
            counter = 0;
            foreach (var Id in Ids)
            {
                if (Id.Text == "") continue;
                counter++;
                Assert.IsTrue(mots.Contains(Id.Text), "Number " + Id.Text + " non présent dans le PDF");
            }
            Assert.AreEqual(count, counter, "Différence entre les SupplierInvoices Number du tableau et du PDF");
            ReadOnlyCollection<IWebElement> dates;
            dates = _webDriver.FindElements(By.XPath("//*/tr/td[7]/a"));
            
            counter = 0;
            foreach (var date in dates)
            {
                if (date.Text == "") continue;
                counter++;
                Assert.IsTrue(mots.Contains(date.Text), "Date " + date.Text + " non présent dans le PDF");
            }
            Assert.AreEqual(count, counter, "Différence entre les SupplierInvoices Number du tableau et du PDF");
            ReadOnlyCollection<IWebElement> suppliers;
            suppliers = _webDriver.FindElements(By.XPath("//*/tr/td[8]/a/b"));
            
            counter = 0;
            foreach (var _supplier in suppliers)
            {
                if (_supplier.Text == "") continue;
                counter++;
                Assert.IsTrue(mots.Contains(_supplier.Text.Split(' ')[0]), "Supplier " + _supplier.Text.Split(' ')[0] + " non présent dans le PDF");
            }
            Assert.AreEqual(count, counter, "Différence entre les SupplierInvoices Supplier du tableau et du PDF");
            ReadOnlyCollection<IWebElement> amounts;
            amounts = _webDriver.FindElements(By.XPath("//*/tr/td[10]"));
            counter = 0;
            foreach (var amount in amounts)
            {
                if (amount.Text == "") continue;
                counter++;
                if (amount.Text.Contains(","))
                {
                    if (amount.Text.Length >= "€ 2 000,0000".Length) continue;
                }
                else
                {
                    if (amount.Text.Length >= "€ 2 000".Length) continue;
                }
                Assert.IsTrue(mots.Contains(convertAmountPdf(amount.Text)), "Amount " + convertAmountPdf(amount.Text) + " non présent dans le PDF");
            }
            Assert.AreEqual(count, counter, "Différence entre les SupplierInvoices Amount du tableau et du PDF");
            ReadOnlyCollection<IWebElement> comments;
            comments = _webDriver.FindElements(By.XPath("//*/tr/td[9]/div"));
            foreach (var comment in comments)
            {
                if (comment.GetAttribute("title") == "") continue;
                string[] explodes = comment.GetAttribute("title").Split(' ');
                foreach (string explode in explodes)
                {
                    // "mm" = "m"+"m"
                    if (explode.Contains("mm")) continue;
                    Assert.IsTrue(mots.Contains(explode), "Morceau de Comment " + explode + " non présent dans le PDF");
                }
            }
        }

        public void CheckExport(FileInfo trouveXLSX, string decimalSeparator)
        {
            int resultNumber = OpenXmlExcel.GetExportResultNumber("Invoices", trouveXLSX.FullName);
            Assert.AreEqual(CheckTotalNumber(), resultNumber, "Mauvais nombre de lignes");
            List<string> idsXLSX = OpenXmlExcel.GetValuesInList("Id", "Invoices", trouveXLSX.FullName);
            //FIXME Assert.IsTrue(int.Parse(ids[0]) > int.Parse(ids[1]), "Mauvais ordonnancement dans le fichier Excel");

            int count = Math.Min(CheckTotalNumber(), 100);

            ReadOnlyCollection<IWebElement> Ids;
            Ids = _webDriver.FindElements(By.XPath("//*/tr/td[3]/a"));
            int counter = 0;
            foreach (var Id in Ids)
            {
                if (Id.Text == "") continue;
                counter++;
                Assert.IsTrue(idsXLSX.Contains(Id.Text), "Number " + Id.Text + " non présent dans le XLSX");
            }
            Assert.AreEqual(count, counter, "Différence entre les SupplierInvoices Number du tableau et du XLSX");

            List<string> invoiceNumberXLSX = OpenXmlExcel.GetValuesInList("Invoice Number", "Invoices", trouveXLSX.FullName);
            ReadOnlyCollection<IWebElement> numberStrings;
            numberStrings = _webDriver.FindElements(By.XPath("//*/tr/td[4]/a"));
            counter = 0;
            foreach (var numberString in numberStrings)
            {
                if (numberString.Text == "") continue;
                counter++;
                Assert.IsTrue(invoiceNumberXLSX.Contains(numberString.Text), "number string " + numberString.Text + " non présent dans le XLSX");
            }
            Assert.AreEqual(count, counter, "Différence entre les SupplierInvoices number string du tableau et du XLSX");

            List<string> dateXLSX = OpenXmlExcel.GetValuesInList("Date", "Invoices", trouveXLSX.FullName);
            ReadOnlyCollection<IWebElement> dates;
            dates = _webDriver.FindElements(By.XPath("//*/tr/td[7]/a"));
            counter = 0;
            foreach (var date in dates)
            {
                if (date.Text == "") continue;
                counter++;
                Assert.IsTrue(dateXLSX.Contains(date.Text), "Date " + date.Text + " non présent dans le XLSX");
            }
            Assert.AreEqual(count, counter, "Différence entre les SupplierInvoices Number du tableau et du XLSX");

            List<string> suppliersXLSX = OpenXmlExcel.GetValuesInList("Supplier", "Invoices", trouveXLSX.FullName);
            ReadOnlyCollection<IWebElement> suppliers;
            suppliers = _webDriver.FindElements(By.XPath("//*/tr/td[8]/a/b"));
            counter = 0;
            foreach (var _supplier in suppliers)
            {
                if (_supplier.Text == "") continue;
                counter++;
                Assert.IsTrue(suppliersXLSX.Contains(_supplier.Text), "Supplier " + _supplier.Text + " non présent dans le XLSX");
            }
            Assert.AreEqual(count, counter, "Différence entre les SupplierInvoices Supplier du tableau et du XLSX");

            List<string> amountsXLSX = OpenXmlExcel.GetValuesInList("Invoiced Amount excl tax", "Invoices", trouveXLSX.FullName);
            ReadOnlyCollection<IWebElement> amounts;
            amounts = _webDriver.FindElements(By.XPath("//*/tr/td[10]"));
            counter = 0;
            foreach (var amount in amounts)
            {
                if (amount.Text == "") continue;
                counter++;
                Assert.IsTrue(amountsXLSX.Contains(convertAmountExcel(amount.Text)), "Amount " + convertAmountExcel(amount.Text) + " non présent dans le XLSX");
            }
            Assert.AreEqual(count, counter, "Différence entre les SupplierInvoices Amount du tableau et du XLSX");

            List<string> amountsTaxXLSX = OpenXmlExcel.GetValuesInList("Invoiced Amount incl. tax", "Invoices", trouveXLSX.FullName);
            ReadOnlyCollection<IWebElement> amountsTax;
            amountsTax = _webDriver.FindElements(By.XPath("//*/tr/td[11]"));
            counter = 0;
            foreach (var amount in amountsTax)
            {
                if (amount.Text == "") continue;
                counter++;
                Assert.IsTrue(amountsTaxXLSX.Contains(convertAmountExcel(amount.Text)), "Amount " + convertAmountExcel(amount.Text) + " non présent dans le XLSX");
            }
            Assert.AreEqual(count, counter, "Différence entre les SupplierInvoices Amount du tableau et du XLSX");

            List<string> commentsXLSX = OpenXmlExcel.GetValuesInList("Comment", "Invoices", trouveXLSX.FullName);
            ReadOnlyCollection<IWebElement> comments;
            comments = _webDriver.FindElements(By.XPath("//*/tr/td[9]/div"));
           
            foreach (var comment in comments)
            {
                if (comment.GetAttribute("title") == "") continue;
                Assert.IsTrue(commentsXLSX.Contains(comment.GetAttribute("title")), "Comment " + comment.GetAttribute("title") + " non présent dans le XLSX");
            }

        }

        public string convertAmountExcel(string amount)
        {
            amount = amount.Replace(".", ",");
            // entrée € 665,5000
            // sortie 665,5
            // entrée € 26,9120
            // sortie 26,912
            string sansEuro = amount.Substring(2).Replace(" ", "");
            //double doublePrecision = Math.Round(double.Parse(sansEuro), 4);
            double doublePrecision = Convert.ToDouble(sansEuro, new NumberFormatInfo() { NumberDecimalSeparator = "," });
            return doublePrecision.ToString().Replace(".", ",");
        }

        public string convertAmountPdf(string amount)
        {
            string stringPrecision;
            // entrée € 665,5000
            // sortie 665,5
            // entrée € 26,912
            // sortie 26,91
            string sansEuro = amount.Substring(2).Replace(" ", "");
            
            stringPrecision = string.Format("{0:0.0000}", sansEuro).Replace(".", ",");

            while (stringPrecision.EndsWith("0"))
            {
                stringPrecision = stringPrecision.Substring(0,stringPrecision.Length - 1);
            }
            if (stringPrecision.EndsWith(","))
            {
                stringPrecision = stringPrecision.Substring(0, stringPrecision.Length - 1);
            }
            return stringPrecision;
        }

        public IEnumerable<IWebElement> GetInvoiceSupplierNumbers()
        {
            var invoiceSupplierNumbers = _webDriver.FindElements(By.XPath(INVOICE_SUPPLIERS_NUMBERS));
            return invoiceSupplierNumbers;
        }
        public IEnumerable<IWebElement> GetInvoiceSupplierTypes()
        {
            ReadOnlyCollection<IWebElement> invoiceSupplierTypes;
            invoiceSupplierTypes = _webDriver.FindElements(By.XPath(INVOICE_SUP_TYPES));
            
            return invoiceSupplierTypes;
        }
        public void ScrollTillItemIsVisible(IWebElement element)
        {
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView();",element);
        }

        public bool CheckSupplierInvoicesTypes(IEnumerable<IWebElement> suppliersInvoicesTypes)
        {
            foreach (var element in suppliersInvoicesTypes)
            {
                if(element.Text != "CN")
                {
                    return false;
                }
            }
            return true;
        }
        public bool VerifyClaimCHanged(bool decreasestock, string claimtype, string newcomment, string sanctionamount)
        {
            var decreasestockhtml = WaitForElementIsVisible(By.XPath("//*[@id=\"DecreaseStock\"]"));
            var claimtypehtml = isElementExists(By.XPath("//*[@id=\"IsChecked_1\"]/../../td[2]/label[text()=\""+claimtype+"\"]"));
            var commenthtml = WaitForElementIsVisible(By.Id("Comment"));
            var sanctionamounthtml = WaitForElementIsVisible(By.Id("SanctionAmount")).GetAttribute("value");
            if(decreasestockhtml.Selected && claimtypehtml && commenthtml.Text == newcomment && sanctionamounthtml == sanctionamount)
            {
                return true;
            }
            return false;
        }

        public void CheckFirstIndexSupplierInvoices(List<string> mots, string currency, string decimalSeparator)
        {
            // dans l'index, le Number (Price Id dans PDF) petit
            IWebElement indexNumber;
            indexNumber = WaitForElementIsVisible(By.XPath("//*/div[@id='list-item-with-action']/table/tbody/tr[1]/td[4]/a"));
           
            Assert.IsTrue(mots.Contains(indexNumber.Text), "indexNumber " + indexNumber.Text + " non trouvé dans la table");
            // dans l'index, le Invoice Number (Number string dans PDF) gros
            IWebElement indexInvoiceNumber;
            indexInvoiceNumber = WaitForElementIsVisible(By.XPath("//*/div[@id='list-item-with-action']/table/tbody/tr[1]/td[4]"));

            Assert.IsTrue(mots.Contains(indexInvoiceNumber.Text), "indexInvoiceNumber " + indexInvoiceNumber.Text + " non trouvé dans la table");
            // dans l'index, la Date
            IWebElement indexDate;
            indexDate = WaitForElementIsVisible(By.XPath("//*/div[@id='list-item-with-action']/table/tbody/tr[1]/td[7]"));
            
            Assert.IsTrue(mots.Contains(indexDate.Text), "indexDate " + indexDate.Text + " non trouvé dans la table");
            // dans l'index le supplier
            IWebElement indexSupplier;
            indexSupplier = WaitForElementIsVisible(By.XPath("//*/div[@id='list-item-with-action']/table/tbody/tr[1]/td[8]/a/b"));
            
            var cutWords = indexSupplier.Text.Split(' ');
            foreach (var cutWord in cutWords)
            {
                Assert.IsTrue(mots.Contains(cutWord), "indexSupplier "+ cutWord + " ("+indexSupplier.Text + ") non trouvé dans la table");
            }
            // FIXME dans l'index amount sans ,0000
            IWebElement indexAmount;
            indexAmount = WaitForElementIsVisible(By.XPath("//*/div[@id='list-item-with-action']/table/tbody/tr[1]/td[10]"));
            
            double amount = double.Parse(convertAmountExcel(indexAmount.Text));
            int amountSolo = Convert.ToInt32(amount);
            if (amount - (double)amountSolo == 0)
            {
                if (amount < 1000)
                {
                    //€ 1 000,0000 pas vérifiable dans le pdf
                    Assert.IsTrue(mots.Contains(amount.ToString()), "indexAmount" + indexAmount.Text + " non trouvé dans la table");
                }
            } else
            {
                // terre inconnue
                Console.WriteLine("indexAmount" + indexAmount.Text + " => "+amount +"<>"+ amountSolo);
            }
            // dans l'index validates (icone)
            
            bool valide;
            valide = isElementVisible(By.XPath("//*/div[@id='list-item-with-action']/table/tbody/tr[1]/td[1]/a/div/div/i[contains(@class,'circle-check')]"));

            Assert.IsTrue(valide);
        }

        public void PageUp()
        {
            var html = _webDriver.FindElement(By.TagName("html"));
            html.SendKeys(Keys.PageUp);
            html.SendKeys(Keys.PageUp);
            html.SendKeys(Keys.PageUp);
            html.SendKeys(Keys.PageUp);
        }
        public SupplierGeneralInfoTab ClickGeneralInformationTab()
        {
            _generalInformation = WaitForElementIsVisible(By.XPath("//*[@id=\"hrefTabContentInformations\"]"));
            _generalInformation.Click();
            return new SupplierGeneralInfoTab(_webDriver, _testContext);
        }

        public string GetInvoiceNumber()
        {
            var invoiceNumber = WaitForElementIsVisible(By.XPath("//*[@id=\"tb-new-supplier-invoice-number\"]"));
            return invoiceNumber.GetAttribute("value");
        }

        public string GetInvoiceSite()
        {
            var invoiceSite = WaitForElementIsVisible(By.XPath("//*[@id=\"form-create-supplier-invoice\"]/div/div[1]/div[2]/div/div/input[3]"));
            return invoiceSite.GetAttribute("value");
        }
        public string GetInvoiceSupplier()
        {
            var invoiceupplier = WaitForElementIsVisible(By.XPath("//*[@id=\"form-create-supplier-invoice\"]/div/div[2]/div/div/div/input[2]"));
            return invoiceupplier.GetAttribute("value");
        }

        public int GetTotalResultRowsPage()
        {
            var rowsResults = _webDriver.FindElements(By.XPath(NB_ROWS_VALIDATE_SUPPLIERS));
            WaitForLoad();
            return rowsResults.Count;
        }
        public bool IsIconDisplayed()
        {
            _iconValidateSuppliers = _webDriver.FindElement(By.XPath(ICON_VALIDATE_SUPPLIERS));
            return _iconValidateSuppliers.Displayed;
        }

        public bool VerifyDataExist()
        {
            var data = _webDriver.FindElements(By.XPath("//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]"));
            return data.Any();
        }
        public void ClickOnReinvoiceIncustomerInvoice()
        {
            ShowExtendedMenu();
            _reinvoiceInCustomerInvoice = WaitForElementIsVisible(By.Id(REINVOICE_IN_CUSTOMER_INVOICE));
            _reinvoiceInCustomerInvoice.Click();
            WaitForLoad();
        }
        public void ClickOnVerify()
        {

            _verify = WaitForElementIsVisible(By.Id(VERIFY));
            _verify.Click();
            WaitForLoad();

        }
        public void ClickOnGenerate()
        {

            _generate = WaitForElementIsVisible(By.Id(GENERATE));
            _generate.Click();
            WaitForLoad();


        }
      
        public bool IsVisible()
        {
            try
            {
                _webDriver.FindElement(By.XPath("//*[@id=\"form-reinvoice-data\"]/h5"));
            }
            catch
            {
                return false;
            }

            return true;
        }
        public bool checkNotValidateOnly()
        {
            int tot = CheckTotalNumber();

            if (tot == 0)
                return true;


            IWebElement check_cercle = null;

            WaitPageLoading();

            var elements = _webDriver.FindElements(By.XPath("/html/body/div[2]/div/div[2]/div[2]/table/tbody/tr[*]/td[1]/a/div/div/i"));

            foreach (var element in elements)
            {
                if (element.GetAttribute("class").Contains("fa-circle-check"))
                {
                    check_cercle = element;
                    break;
                }
            }

            if (check_cercle != null)
            {
                return false;
            }
            else
            {
                if (tot > 100)
                {
                    GoToNextPage("2");
                    elements = _webDriver.FindElements(By.XPath("/html/body/div[2]/div/div[2]/div[2]/table/tbody/tr[*]/td[1]/a/div/div/i"));
                    foreach (var element in elements)
                    {
                        if (element.GetAttribute("class").Contains("fa-circle-check"))
                        {
                            check_cercle = element;
                            break;
                        }
                    }
                    if (check_cercle != null)
                    {
                        return false;
                    }

                }
               
                return true;
            }
        }
           
            
        public void GoToNextPage(string pageNumber)
        {
            var pageNumbers = _webDriver.FindElements(By.XPath("//*[@id=\"list-item-with-action\"]/nav/ul/li[*]/a"));
            foreach (var page in pageNumbers)
            {
                if (page.Text == pageNumber)
                {
                    string xpath = "//*[@id=\"list-item-with-action\"]/nav/ul/li[*]/a[text()=" + pageNumber + "]";
                    var pageBtn = WaitForElementIsVisible(By.XPath(xpath));
                    pageBtn.Click();
                }
            }
        }

        public bool AccessPage()
        {
            if (isElementExists(By.Id(PLUS_BUTTON)))
            {
                return true;
            }

            return false;
        }

        public bool VerifyInvoiceNumber(string invoiceNumber)
        {
            ReadOnlyCollection<IWebElement> suppliers;
            while(!isElementExists(By.XPath(VERIFY_NUMBER)))
            {
                WaitLoading();

            }
            suppliers = _webDriver.FindElements(By.XPath(VERIFY_NUMBER));

            if (suppliers.Count == 0)
                return false;

            foreach (var elm in suppliers)
            {
                if (!string.Equals(elm.Text.Trim(), invoiceNumber.Trim()))
                {
                    return false;
                }
            }
            return true;
        }
        public void Show()
        {
            _show = WaitForElementExists(By.XPath(SHOW));
            if (_show.GetAttribute("class").Contains("collapsed"))
            {
                ScrollToElement(_show);
                _show.Click();
                WaitLoading();
            }

        }

        public void PageUpSupplierInvoice()
        {
            var scrollbarElement = _webDriver.FindElement(By.ClassName("ps__scrollbar-y-rail"));
            Actions actions = new Actions(_webDriver);
            actions.ClickAndHold(scrollbarElement).MoveByOffset(0, -300).Release().Perform();
        }


    }
}