using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;


namespace Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.OutputForm

{
    public class OutputFormPage : PageBase
    {

        public OutputFormPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //__________________________________ Constantes ________________________________________________

        // General
        private const string NEW_OUTPUT_FORM = "//*/a[text()='New output form']";
        private const string EXTENDED_MENU = "//*[@id=\"div-body\"]/div/div[2]/div[1]/div/div[1]";
        private const string PRINT_RESULTS = "btn-print-output-forms-report";
        private const string CONFIRM_PRINT = "validatePrint";
        private const string EXPORT = "//*[@id=\"div-body\"]/div/div[2]/div[1]/div/div[1]/div/a[text()='Export']";
        private const string EXPORT_WMS = "//*[@id=\"div-body\"]/div/div[2]/div[1]/div/div[1]/div/a[text()='WMS Export']";
        private const string ENABLE_EXPORT_FOR_SAGE = "//*[@id=\"div-body\"]/div/div[2]/div[1]/div/div[1]/div/a[contains(text(),'Enable export for SAGE')]";
        private const string EXPORT_SAGE = "btn-export-for-sage";
        private const string CONFIRM_EXPORT_SAGE = "btn-popup-validate";
        private const string CANCEL_SAGE = "btn-cancel-popup";
        private const string DOWNLOAD_FILE = "btn-download-file";
        private const string GENERATE_SAGE_TXT = "btn-export-for-sage-txt-generation";

        // Tableau
        private const string OUTPUT_FORM_NUMBER = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[1]/td[4]";
        private const string DELETE_OUTPUT_FORM = "//*/a[contains(@class,'btn-OutputForm-delete')]/span";
        private const string OUTPUT_FORM_DATE = "//*[@id=\"list-item-with-action\"]/table/tbody/tr/td[8]";

        // Filtres
        private const string RESET_FILTER = "//*[@id=\"formSearchOutputForms\"]/div[1]/a";

        private const string FILTER_SEARCH = "SearchNumber";
        private const string FILTER_SORTBY = "cbSortBy";
        private const string FILTER_NOT_VALIDATED = "ShowNotValidated";
        private const string FILTER_SHOW_ALL = "//*[@id=\"ShowActive\"][@value=\"All\"]";
        private const string FILTER_SHOW_ACTIVE = "//*[@id=\"ShowActive\"][@value=\"ActiveOnly\"]";
        private const string FILTER_SHOW_INACTIVE = "//*[@id=\"ShowActive\"][@value=\"InactiveOnly\"]";
        private const string FILTER_EXPORTED_FOR_SAGE_MANUALLY = "//*[@id=\"ShowActive\"][@value=\"ExportedForSageManually\"]";
        private const string FILTER_EXPORTED_FOR_SAGE_MANUALLY_DEV = "ShowExportedForSageManually";
        private const string FILTER_DATE_FROM = "date-picker-start";
        private const string FILTER_DATE_TO = "date-picker-end";
        private const string FILTER_SITES = "SelectedSites_ms";
        private const string UNSELECT_ALL_SITES = "/html/body/div[10]/div/ul/li[2]/a";
        private const string SELECT_ALL_SITES = "/html/body/div[10]/div/ul/li[1]/a[@title='Check all']";
        private const string SEARCH_SITES = "/html/body/div[10]/div/div/label/input";
        private const string FILTER_OF_TYPES_SELECT = "SelectedOFTypes_ms";
        private const string FILTER_OF_TYPES_UNCHECKALL = "/html/body/div[11]/div/ul/li[2]/a";
        private const string FILTER_OF_TYPES_SEARCH = "/html/body/div[11]/div/div/label/input";
        private const string FILTER_OF_TYPES_CHECKALL = "/html/body/div[11]/div/ul/li[1]/a";

        private const string VALIDATION = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[{0}]/td[1]/img[@alt='Valid']";
        private const string INACTIVE = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[{0}]/td[1]/img[@alt='Inactive']";
        private const string SENT_TO_SAGE = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[{0}]/td[1]/img[2]";
        private const string SENT_TO_SAGE_MANUALLY = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[{0}]//td[1]/i";//*[@id="list-item-with-action"]/table/tbody/tr/td[1]/i
        private const string NUMBER = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[4]";
        private const string SITE = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[5]";
        private const string FROM_PLACE = "//*[@id=\"list-item-with-action\"]/table/tbody/tr/td[6]";
        private const string TO_PLACE = "//*[@id=\"list-item-with-action\"]/table/tbody/tr/td[7]";
        private const string DATE = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[8]";
        private const string TOTAL_VAT = "//*[@id=\"list-item-with-action\"]/table/tbody/tr/td[10]";

        private const string FIRST_LINE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[2]";
        private const string FIRST_LINE_NAME = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[1]/div[3]/span";
        private const string FIRST_GROUP_NAME = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[1]/div/div/span";
        private const string BYGROUP_SUBMENU = "//*[@id=\"hrefTabContentGroups\"]";
        private const string QUANTITY_INPUT = "item_OutputFormDetail_Quantity";
        private const string ROUND_AMOUNT = "roundAmount";
        private const string VATFROMBYGROUPS = "total-price-span";
        private const string VATAMOUNTSCOLUMN = "//*[@id=\"detailedGroupsTable\"]/tbody/tr[*]/td[3]";
        private const string FIRST_OF = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[1]";
        private const string VAT_FOOTER = "//*[@id=\"tabContentDetails\"]/div/table[1]/tbody/tr[4]/td[2]";
        private const string FOOTER_SUBMENU = "hrefTabContentFooter";
        private const string FIRST_OutPutForm = "//*[@id=\"list-item-with-action\"]/table/tbody/tr/td[4]";
        private const string ID_OUT_PUT_FORM = " /html/body/div[2]/div/div[2]/div[2]/table/tbody/tr[1]/td[4]";
        private const string DATE_OUT_PUT_FORM = "/html/body/div[2]/div/div[2]/div[2]/table/tbody/tr[1]/td[8]";

        //__________________________________ Variables _________________________________________________

        // General
        [FindsBy(How = How.XPath, Using = NEW_OUTPUT_FORM)]
        private IWebElement _createNewOutputForm;

        [FindsBy(How = How.Id, Using = PRINT_RESULTS)]
        private IWebElement _printResults;

        [FindsBy(How = How.Id, Using = CONFIRM_PRINT)]
        private IWebElement _confirmPrint;

        [FindsBy(How = How.XPath, Using = EXPORT)]
        private IWebElement _export;

        [FindsBy(How = How.XPath, Using = EXPORT_WMS)]
        private IWebElement _exportWMS;

        [FindsBy(How = How.XPath, Using = ENABLE_EXPORT_FOR_SAGE)]
        private IWebElement _enableExportForSage;

        [FindsBy(How = How.Id, Using = EXPORT_SAGE)]
        private IWebElement _exportSage;

        [FindsBy(How = How.Id, Using = CONFIRM_EXPORT_SAGE)]
        private IWebElement _confirmExportSage;

        [FindsBy(How = How.Id, Using = CANCEL_SAGE)]
        private IWebElement _cancelSage;

        [FindsBy(How = How.Id, Using = DOWNLOAD_FILE)]
        private IWebElement _downloadFile;

        [FindsBy(How = How.Id, Using = GENERATE_SAGE_TXT)]
        private IWebElement _generateSageTxt;

        // Tableau
        [FindsBy(How = How.XPath, Using = OUTPUT_FORM_NUMBER)]
        private IWebElement _outputFormNumber;

        [FindsBy(How = How.XPath, Using = DELETE_OUTPUT_FORM)]
        private IWebElement _deleteOutputForm;

        [FindsBy(How = How.XPath, Using = OUTPUT_FORM_DATE)]
        private IWebElement _outputFormDate;

        //__________________________________Filtres___________________________________________________

        [FindsBy(How = How.XPath, Using = RESET_FILTER)]
        private IWebElement _resetFilter;

        [FindsBy(How = How.Id, Using = FILTER_SEARCH)]
        private IWebElement _searchByNumber;

        [FindsBy(How = How.Id, Using = FILTER_SORTBY)]
        private IWebElement _sortBy;

        [FindsBy(How = How.Id, Using = FILTER_NOT_VALIDATED)]
        private IWebElement _showNotValidated;

        [FindsBy(How = How.XPath, Using = FILTER_SHOW_ALL)]
        private IWebElement _showAll;

        [FindsBy(How = How.XPath, Using = FILTER_SHOW_ACTIVE)]
        private IWebElement _showItemsActive;

        [FindsBy(How = How.XPath, Using = FILTER_SHOW_INACTIVE)]
        private IWebElement _showItemsInactive;

        [FindsBy(How = How.XPath, Using = FILTER_EXPORTED_FOR_SAGE_MANUALLY)]
        private IWebElement _showExportedForSageManually;

        [FindsBy(How = How.Id, Using = FILTER_DATE_FROM)]
        public IWebElement _dateFrom;

        [FindsBy(How = How.Id, Using = FILTER_DATE_TO)]
        public IWebElement _dateTo;

        [FindsBy(How = How.Id, Using = FILTER_SITES)]
        private IWebElement _siteFilter;

        [FindsBy(How = How.XPath, Using = SEARCH_SITES)]
        private IWebElement _searchSite;

        [FindsBy(How = How.XPath, Using = UNSELECT_ALL_SITES)]
        private IWebElement _unselectAll;

        [FindsBy(How = How.XPath, Using = SELECT_ALL_SITES)]
        private IWebElement _selectAll;

        [FindsBy(How = How.XPath, Using = FILTER_OF_TYPES_SELECT)]
        private IWebElement _selectOfTypes;

        [FindsBy(How = How.XPath, Using = FILTER_OF_TYPES_UNCHECKALL)]
        private IWebElement _uncheckOfTypes;

        [FindsBy(How = How.XPath, Using = FILTER_OF_TYPES_SEARCH)]
        private IWebElement _searchOfTypes;

        [FindsBy(How = How.XPath, Using = FILTER_OF_TYPES_CHECKALL)]
        private IWebElement _checkAllOfTypes;

        public enum FilterType
        {
            SearchByNumber,
            SortBy,
            ShowNotValidated,
            DateFrom,
            DateTo,
            CheckAllSites,
            CheckAllOfTypes,
            ShowGroup,
            CombosOptions
        }
        public enum TypeOfShowGroup
        {
            All,
            InactiveOnly,
            ActiveOnly,
            ExportedForSageManually,
            NoShowGroupOption
        }

        public enum TypeCombosOptions
        {
            Sites,
            SitePlaces,
            OfTypes,
            NoComboOption
        }
        public void Filter(FilterType filterType, object value, TypeOfShowGroup showType = TypeOfShowGroup.NoShowGroupOption, TypeCombosOptions comboType = TypeCombosOptions.NoComboOption)
        {
            Actions action = new Actions(_webDriver);
            switch (filterType)
            {
                case FilterType.SearchByNumber:
                    PageUpOF();
                    _searchByNumber = WaitForElementIsVisibleNew(By.CssSelector("#SearchNumber"));
                    _searchByNumber.SetValue(ControlType.TextBox, value);
                    break;
                case FilterType.SortBy:
                    _sortBy = WaitForElementIsVisibleNew(By.Id(FILTER_SORTBY));
                    _sortBy.Click();
                    ApplySortBy(value);
                    break;
                case FilterType.ShowNotValidated:
                    _showNotValidated = WaitForElementExists(By.Id(FILTER_NOT_VALIDATED));
                    _showNotValidated.SetValue(ControlType.CheckBox, value);
                    break;

                case FilterType.ShowGroup:
                    ApplyRadioButtonFilter(showType, value);
                    break;

                case FilterType.DateFrom:
                    _dateFrom = WaitForElementIsVisibleNew(By.CssSelector("#date-picker-start"));
                    _dateFrom.SetValue(ControlType.DateTime, value);
                    _dateFrom.SendKeys(Keys.Tab);
                    break;
                case FilterType.DateTo:
                    _dateTo = WaitForElementIsVisibleNew(By.CssSelector("#date-picker-end"));
                    _dateTo.SetValue(ControlType.DateTime, value);
                    _dateTo.SendKeys(Keys.Tab);
                    break;

                case FilterType.CheckAllSites:
                    _siteFilter = WaitForElementExists(By.Id(FILTER_SITES));
                    action.MoveToElement(_siteFilter).Perform();
                    _siteFilter.SendKeys(Keys.Down);
                    //_siteFilter.Click();

                    // On décoche toutes les options
                    _selectAll = WaitForElementIsVisibleNew(By.XPath(SELECT_ALL_SITES));
                    _selectAll.Click();
                    _siteFilter.SendKeys(Keys.Up);
                    break;

                case FilterType.CombosOptions:
                    ApplyComboFilter(comboType, value);
                    break;
                case FilterType.CheckAllOfTypes:
                    _selectOfTypes = WaitForElementExists(By.Id(FILTER_OF_TYPES_SELECT));
                    action.MoveToElement(_selectOfTypes).Perform();
                    _selectOfTypes.SendKeys(Keys.Down);

                    _checkAllOfTypes = WaitForElementIsVisibleNew(By.XPath(FILTER_OF_TYPES_CHECKALL));
                    _checkAllOfTypes.Click();
                    _selectOfTypes.SendKeys(Keys.Up);
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);

            }

            WaitPageLoading();
            WaitForLoad();
        }

        private void ApplySortBy(object value)
        {
            var element = WaitForElementIsVisible(By.XPath("//option[contains(@value,'" + value + "')]"));
            _sortBy.SetValue(ControlType.DropDownList, element.Text);
            _sortBy.Click();
        }

        public bool SortByFilters(string option, string dateFormat = null)
        {
            Filter(OutputFormPage.FilterType.SortBy, option);
            if (!IsSorted(option, dateFormat))
            {
                return false;
            }

            return true;
        }

        public bool IsSorted(string type, string dateFormat = null)
        {
            List<IWebElement> elements = new List<IWebElement>();
            if (type == "NUMBER")
            {
                elements = _webDriver.FindElements(By.XPath(NUMBER)).ToList();
            }
            else
            {
                elements = _webDriver.FindElements(By.XPath(DATE)).ToList();
            }


            if (elements.Count() == 0)
                return false;

            // Définir la culture pour les dates si nécessaire
            CultureInfo ci = null;
            if (type == "DATE" && dateFormat != null)
            {
                ci = dateFormat.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");
            }

            for (int i = 1; i < elements.Count; i++)
            {
                try
                {
                    var previous = elements[i - 1].Text;
                    var current = elements[i].Text;

                    if (type == "NUMBER")
                    {
                        // Comparaison de nombres
                        int previousNumber = int.Parse(previous);
                        int currentNumber = int.Parse(current);

                        if (currentNumber > previousNumber)
                            return false;
                    }
                    else if (type == "DATE")
                    {
                        // Comparaison de dates
                        var previousDate = DateTime.Parse(previous, ci);
                        var currentDate = DateTime.Parse(current, ci);

                        if (DateTime.Compare(previousDate.Date, currentDate.Date) < 0)
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

        public void ApplyRadioButtonFilter(TypeOfShowGroup type, object value)
        {
            string elementId;
            if (type == TypeOfShowGroup.All)
            {
                elementId = "ShowAll";
            }
            else if (type == TypeOfShowGroup.InactiveOnly)
            {
                elementId = "ShowInactiveOnly";
            }
            else if (type == TypeOfShowGroup.ActiveOnly)
            {
                elementId = "ShowActiveOnly";
            }
            else if (type == TypeOfShowGroup.ExportedForSageManually)
            {
                if (IsDev())
                {
                    elementId = FILTER_EXPORTED_FOR_SAGE_MANUALLY_DEV;
                }
                else
                {
                    elementId = FILTER_EXPORTED_FOR_SAGE_MANUALLY;
                }
            }
            else
            {
                elementId = "";
            }
            var element = WaitForElementExists(By.Id(elementId));
            Actions action = new Actions(_webDriver);

            action.MoveToElement(element).Perform();
            element.SetValue(ControlType.RadioButton, value);
            WaitForLoad();
        }

        public void ApplyComboFilter(TypeCombosOptions type, object value)
        {
            var comboXpath = "";

            if (type == TypeCombosOptions.OfTypes)
            {
                comboXpath = FILTER_OF_TYPES_SELECT;

                _selectOfTypes = WaitForElementExists(By.Id(FILTER_OF_TYPES_SELECT));
                var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
                javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", _selectOfTypes);
            }
            else if (type == TypeCombosOptions.Sites)
            {
                comboXpath = FILTER_SITES;
            }

            ComboBoxSelectById(new ComboBoxOptions(comboXpath, (string)value));
        }
        public void ResetFilter()
        {
            WaitPageLoading();
            WaitForLoad();
            Actions action = new Actions(_webDriver);
            PageUpOF();
            _resetFilter = WaitForElementIsVisibleNew(By.XPath(RESET_FILTER));
            action.MoveToElement(_resetFilter).Perform();
            _resetFilter.SendKeys(Keys.Tab);
            _resetFilter.Click();
            WaitForLoad();

            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                // DateTo à 2 mois après
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
                if (isElementVisible(By.XPath(string.Format("//*[@id='list-item-with-action']/table/tbody/tr[1]/td[1]/div/i[contains(@class,'circle-check')]", i + 1))))
                {
                    _webDriver.FindElement(By.XPath(string.Format("//*[@id='list-item-with-action']/table/tbody/tr[1]/td[1]/div/i[contains(@class,'circle-check')]", i + 1)));
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

        public bool CheckStatus(bool active)
        {
            bool isActive = false;
            int tot = CheckTotalNumber() > 100 ? 100 : CheckTotalNumber();

            if (tot == 0)
                return false;

            for (int i = 0; i < tot; i++)
            {
                if (isElementVisible(By.XPath(String.Format("//*[@id=\"list-item-with-action\"]/table/tbody/tr[{0}]/td[2]/div/i[contains(@class,'circle-xmark')]", i + 1))))
                {
                    _webDriver.FindElement(By.XPath(String.Format("//*[@id=\"list-item-with-action\"]/table/tbody/tr[{0}]/td[2]/div/i[contains(@class,'circle-xmark')]", i + 1)));

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

        public bool CheckShowItemsNotValidated(bool activated)
        {
            bool isValidated = false;
            int tot = CheckTotalNumber() > 100 ? 100 : CheckTotalNumber();

            if (tot == 0)
                return false;

            for (int i = 0; i < tot; i++)
            {
                if (isElementVisible(By.XPath(string.Format(VALIDATION, i + 1))))
                {
                    _webDriver.FindElement(By.XPath(string.Format(VALIDATION, i + 1)));
                }
                else
                {
                    isValidated = true;

                    if (!activated)
                        return false;
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
                if (isElementVisible(By.XPath(string.Format(SENT_TO_SAGE, i + 1))))
                {
                    var elm = _webDriver.FindElement(By.XPath(string.Format(SENT_TO_SAGE, i + 1)));
                    if (!elm.GetAttribute("title").Contains("Accounted (sent to SAGE)"))
                        return false;
                }
                else
                {
                    return false;
                }
            }
            return true;

        }

        public bool IsSentToSageManually()
        {
            int tot = CheckTotalNumber() > 100 ? 100 : CheckTotalNumber();

            if (tot == 0)
                return false;

            for (int i = 0; i < tot; i++)
            {
                if (isElementVisible(By.XPath(string.Format(SENT_TO_SAGE_MANUALLY, i + 1))))
                {
                    var elm = _webDriver.FindElement(By.XPath(string.Format(SENT_TO_SAGE_MANUALLY, i + 1)));

                    //if (!elm.GetAttribute("title").Contains("Accounted (sent to SAGE manually)"))
                    //    return false;

                    //provisoire pour patch
                    if (!elm.GetAttribute("title").Contains("Accounted") && !elm.GetAttribute("title").Contains("Accounted (sent to SAGE manually)"))
                        return false;
                }
                else
                {
                    return false;
                }
            }
            return true;
        }

        public bool VerifySite(string site)
        {
            var sites = _webDriver.FindElements(By.XPath(SITE));

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

        public bool IsDateRespected(DateTime fromDate, DateTime toDate, string dateFormat)
        {
            // Take the date format from the datepicker element and use it to format the date column to avoid date errors
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            var dates = _webDriver.FindElements(By.XPath(DATE));

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

        //__________________________________________________ Méthodes ___________________________________________________

        // General
        public OutputFormCreateModalPage OutputFormCreatePage()
        {
            WaitForLoad();
            ShowPlusMenu();
            _createNewOutputForm = WaitForElementIsVisibleNew(By.CssSelector("#btn-add-new-output-forms"));
            _createNewOutputForm.Click();
            WaitForLoad();


            return new OutputFormCreateModalPage(_webDriver, _testContext);
        }

        public PrintReportPage PrintResults(bool printValue)
        {
            ShowExtendedMenu();
            WaitForLoad();

            _printResults = WaitForElementIsVisible(By.Id(PRINT_RESULTS));
            _printResults.Click();
            WaitForLoad();

            _confirmPrint = WaitForElementIsVisible(By.Id("btn-validate-print-output-form"));
            _confirmPrint.Click();
            WaitForLoad();

            if (printValue)
            {
                FileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-pdf']"));
                ClickPrintButton();

            }

            //Results are opened in a new tab, switch the driver to the newly created one
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new PrintReportPage(_webDriver, _testContext);
        }
        public bool VerifyIsPrint()
        {
            ShowExtendedMenu();

            _printResults = WaitForElementIsVisible(By.Id(PRINT_RESULTS));
            _printResults.Click();
            WaitForLoad();

            if (!isElementVisible(By.Id(CONFIRM_PRINT)))
            {
                return false;
            }
            return true;
        }
        public void ExportResults(bool printVersion)
        {
            ClearDownloads();
            ShowExtendedMenu();
            WaitForLoad();
            _export = WaitForElementIsVisible(By.CssSelector("#btn-export-output-forms"));
            WaitForLoad();
            _export.Click();
            WaitPageLoading();
            if (printVersion)
            {
                FileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));
                ClickPrintButton();

                WaitForDownload();
                //Close();
            }
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


            Regex r = new Regex("^outputforms" + space + annee + "-" + mois + "-" + jour + space + heure + "-" + minutes + "-" + secondes + ".xlsx$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
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
            _exportWMS = WaitForElementIsVisible(By.XPath(EXPORT_WMS));
            _exportWMS.Click();
            WaitForLoad();

            if (versionPrint)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-alt']"));
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

        public bool IsTextFileCorrect(string filePath)
        {
            string mois = "(?:0[1-9]|1[0-2])";  // mois
            string annee = "\\d{4}";            // annee YYYY
            string jour = "[0-3]\\d";           // jour
            string nombre = "\\d{6}";           // nombre


            Regex r = new Regex("^outputforms" + "_" + annee + mois + jour + nombre + ".txt$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
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

        public bool CanClickOnSAGE()
        {
            ShowExtendedMenu();

            _exportSage = WaitForElementIsVisible(By.Id(EXPORT_SAGE));

            if (_exportSage.GetAttribute("disabled") != null)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public bool CanClickOnEnableSAGE()
        {
            ShowExtendedMenu();

            _enableExportForSage = WaitForElementIsVisible(By.Id(ENABLE_EXPORT_FOR_SAGE));

            if (_enableExportForSage.GetAttribute("disabled") != null)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public void ManualExportSageError(bool printValue)
        {

            ShowExtendedMenu();

            _exportSage = WaitForElementIsVisible(By.Id(EXPORT_SAGE));
            _exportSage.Click();
            WaitForLoad();

            _confirmExportSage = WaitForElementIsVisible(By.Id(CONFIRM_EXPORT_SAGE));
            _confirmExportSage.Click();
            WaitForLoad();

            if (printValue)
            {
                IsFileInError(By.CssSelector("[class='fa fa-info-circle']"));
                ClickPrintButton();
            }
        }
        public void ManualExportSage(bool printValue)
        {
            ShowExtendedMenu();

            _exportSage = WaitForElementIsVisible(By.Id(EXPORT_SAGE));
            _exportSage.Click();
            WaitForLoad();

            _confirmExportSage = WaitForElementIsVisible(By.Id(CONFIRM_EXPORT_SAGE));
            _confirmExportSage.Click();
            WaitForLoad();

            if (!printValue)
            {
                if (isElementVisible(By.Id(DOWNLOAD_FILE)))
                {
                    _downloadFile = _webDriver.FindElement(By.Id(DOWNLOAD_FILE));
                    _downloadFile.Click();
                    WaitForLoad();
                }
                if (isElementVisible(By.Id(CANCEL_SAGE)))
                {
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

        public bool IsSAGEFileCorrect(string filePath)
        {
            // "OutputForm 2020-05-11 15-03-27.txt"

            string mois = "(?:0[1-9]|1[0-2])";         // mois
            string space = "(\\s)";                    // Espace
            string annee = "\\d{4}";                   // annee YYYY
            string jour = "[0-3]\\d";                  // jour
            string heure = "(?:0[0-9]|1[0-9]|2[0-3])"; // heure
            string minutes = "[0-5]\\d";               // minutes
            string secondes = "[0-5]\\d";              // secondes


            Regex r = new Regex("^OutputForm" + space + annee + "-" + mois + "-" + jour + space + heure + "-" + minutes + "-" + secondes + ".txt$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = r.Match(filePath);

            return m.Success;
        }

        public void GenerateSageTxt()
        {
            ShowExtendedMenu();

            _generateSageTxt = WaitForElementIsVisible(By.Id(GENERATE_SAGE_TXT));
            _generateSageTxt.Click();
            WaitForLoad();

            _confirmExportSage = WaitForElementIsVisible(By.Id(CONFIRM_EXPORT_SAGE));
            _confirmExportSage.Click();
            WaitForLoad();

            IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-alt']"));
            ClickPrintButton();

            WaitForDownload();
            Close();
        }


        // Tableau
        public OutputFormItem SelectFirstOutputForm()

        {
            WaitForLoad();
            _outputFormNumber = WaitForElementIsVisible(By.XPath(OUTPUT_FORM_NUMBER));


            _outputFormNumber.Click();
            WaitForLoad();


            return new OutputFormItem(_webDriver, _testContext);
        }

        public OutputFormItem DeleteFirstOutputForm()
        {
            WaitPageLoading();
            WaitForLoad();
            _deleteOutputForm = WaitForElementIsVisible(By.XPath(DELETE_OUTPUT_FORM));
            WaitForLoad();
            _deleteOutputForm.Click();
            WaitForLoad();

            var confirmBtnDelete = WaitForElementIsVisible(By.Id("dataConfirmOK"));
            confirmBtnDelete.Click();
            WaitForLoad();

            return new OutputFormItem(_webDriver, _testContext);
        }

        public string GetFirstOutputFormNumber()
        {
            _outputFormNumber = WaitForElementExists(By.XPath(OUTPUT_FORM_NUMBER));
            return _outputFormNumber.Text;
        }
        public string GetFirstOutputFormDate()
        {
            _outputFormDate = WaitForElementExists(By.XPath(OUTPUT_FORM_DATE));
            return _outputFormDate.Text;
        }

        public void ClearDatesFilter()
        {
            _dateFrom = WaitForElementIsVisible(By.Id(FILTER_DATE_FROM));
            _dateFrom.ClearElement();
            _dateFrom.SendKeys(Keys.Tab);
            WaitForLoad();
            _dateTo = WaitForElementIsVisible(By.Id(FILTER_DATE_TO));
            _dateTo.ClearElement();
            _dateTo.SendKeys(Keys.Tab);
            WaitForLoad();

            WaitPageLoading();
        }

        public void SetQtyGetGrpName(string qty)
        {
            var firstLineName = WaitForElementIsVisible(By.XPath(FIRST_LINE_NAME));
            firstLineName.Click();

            var qtyinput = WaitForElementIsVisible(By.Id(QUANTITY_INPUT));
            qtyinput.Clear();
            qtyinput.SendKeys(qty);
            var firstgrpname = WaitForElementIsVisible(By.XPath(FIRST_GROUP_NAME));
            var groupname = firstgrpname.Text;
        }
        public decimal GetVATAmountFromItemsSubMenu()
        {
            var totalAmount = WaitForElementIsVisible(By.Id(VATFROMBYGROUPS));
            var culture = CultureInfo.CreateSpecificCulture("fr-FR");
            decimal.TryParse(totalAmount.Text, NumberStyles.AllowCurrencySymbol | NumberStyles.Number, culture, out decimal VATAmountDecimal);
            return VATAmountDecimal;
        }
        public void GoToByGroupSubMenu()
        {
            var groupBy = WaitForElementIsVisible(By.XPath(BYGROUP_SUBMENU));
            groupBy.Click();
            WaitForLoad();
        }
        public decimal GetVATByGroupName()
        {
            var amountString = "";
            var listamounts = _webDriver.FindElements(By.XPath(VATAMOUNTSCOLUMN));
            foreach (var listamount in listamounts)
            {
                if (listamount.Text != "0")
                {
                    amountString = listamount.Text;
                }
            }
            var culture = CultureInfo.CreateSpecificCulture("fr-FR");
            decimal.TryParse(amountString, NumberStyles.AllowCurrencySymbol | NumberStyles.Number, culture, out decimal VATAmountDecimal);
            return VATAmountDecimal;
        }

        public decimal GetVATAmountFromFooterSubMenu()
        {
            var totalAmount = WaitForElementIsVisible(By.XPath(VAT_FOOTER));
            var culture = CultureInfo.CreateSpecificCulture("fr-FR");
            decimal.TryParse(totalAmount.Text, NumberStyles.AllowCurrencySymbol | NumberStyles.Number, culture, out decimal VATAmountDecimal);
            return VATAmountDecimal;
        }

        public OutputFormItem ClickFirstOF()
        {
            WaitPageLoading();
            var firstOF = WaitForElementIsVisible(By.XPath(FIRST_OF));
            firstOF.Click();
            return new OutputFormItem(_webDriver, _testContext);
        }

        public void GoToFooterSubMenu()
        {
            var footer = WaitForElementIsVisible(By.Id(FOOTER_SUBMENU));
            footer.Click();
        }

        public void PageUp()
        {
            var html = _webDriver.FindElement(By.TagName("html"));
            html.SendKeys(Keys.PageUp);
            html.SendKeys(Keys.PageUp);
            html.SendKeys(Keys.PageUp);
            html.SendKeys(Keys.PageUp);
        }
        public FileInfo GetExportExcelFiles(FileInfo[] taskFiles)
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
        public OutputFormItem ClickOnFirstOF()
        {
            var firstOF = WaitForElementIsVisible(By.XPath(FIRST_OF));
            firstOF.Click();
            return new OutputFormItem(_webDriver, _testContext);
        }
        public string GetSite()
        {
            var site = WaitForElementIsVisible(By.XPath(SITE));
            return site.Text.Trim();
        }
        public string GetFromPlace()
        {
            var fromPlace = WaitForElementIsVisible(By.XPath(FROM_PLACE));
            return fromPlace.Text.Trim();
        }
        public string GetToPlace()
        {
            var toPlace = WaitForElementIsVisible(By.XPath(TO_PLACE));
            return toPlace.Text.Trim();
        }
        public string GetCreationDate()
        {
            var creationDate = WaitForElementIsVisible(By.XPath(DATE));
            return creationDate.Text.Trim();
        }
        public string GetTotalVat()
        {
            var totalVat = WaitForElementIsVisible(By.XPath(TOTAL_VAT));
            return totalVat.Text.Replace("€", "").Trim();
        }

        public string GetDateOutPutForm()
        {

            var element = WaitForElementIsVisible(By.XPath(DATE_OUT_PUT_FORM));
            return element.Text;
        }
        public string GetIdOutPutForm()
        {

            var element = WaitForElementIsVisible(By.XPath(ID_OUT_PUT_FORM));
            return element.Text;
        }

        public bool VerifyOutputFormExist(string drName)
        {
            List<string> strings = new List<string>();
            var names = _webDriver.FindElements(By.XPath("//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]/table/tbody/tr/td[2]"));
            foreach (var name in names)
            {
                strings.Add(name.Text);
            }

            if (!strings.Contains(drName))
            {
                return false;
            }
            return true;
        }
        public void AddItemsToOutputForm(List<string> Ids)
        {
            foreach (var i in Ids)
            {
                Filter(OutputFormPage.FilterType.SearchByNumber, i);
                WaitPageLoading();
                OutputFormItem outputFormItem = SelectFirstOutputForm();
                var item1 = outputFormItem.GetItemNames()[0];
                var item2 = outputFormItem.GetItemNames()[1];
                WaitPageLoading();
                outputFormItem.AddPhyQuantity(item1, "5");
                outputFormItem.AddPhyQuantity(item2, "10");
                outputFormItem.BackToList();
            }
        }
        public void PageUpOF()
        {
            // Modifiez la valeur 'top' et déclenchez un événement
            var scrollbarElement = _webDriver.FindElement(By.ClassName("ps__scrollbar-y-rail"));

            // Ajoutez une interaction de glisser-déposer
            Actions actions = new Actions(_webDriver);
            actions.ClickAndHold(scrollbarElement).MoveByOffset(0, -300).Release().Perform();
        }
    }
}