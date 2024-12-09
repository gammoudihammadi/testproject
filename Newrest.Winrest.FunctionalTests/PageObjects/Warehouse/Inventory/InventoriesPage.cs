using DocumentFormat.OpenXml.Bibliography;
using Microsoft.VisualStudio.TestTools.UnitTesting;
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

namespace Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.Inventory

{
    public class InventoriesPage : PageBase
    {
        public InventoriesPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //_______________________________ Constantes _____________________________________

        // General
        private const string NEW_INVENTORY_BUTTON = "//*[@id=\"div-body\"]/div/div[2]/div[1]/div/div[2]/div/a";
        private const string EXPORT = "btn-export-excel";
        private const string ENABLE_EXPORT_FOR_SAGE = "//*[@id=\"div-body\"]/div/div[2]/div[1]/div/div[1]/div/a[contains(text(), 'Enable export for SAGE')]";
        private const string EXPORT_FOR_SAGE = "btn-export-for-sage";
        private const string CONFIRM_EXPORT_SAGE = "btn-popup-validate";
        private const string GENERATE_SAGE_TXT = "btn-export-for-sage-txt-generation";
        private const string EXPORT_WMS = "//*[@id=\"div-body\"]/div/div[2]/div[1]/div/div[1]/div/a[contains(text(), 'WMS Export')]";
        private const string DELETE_CORBEILLE = "/html/body/div[2]/div/div[2]/div[2]/table/tbody/tr[1]/td[7]/a";
        private const string DELETE_CONFIRM = "dataConfirmOK";

        // Tableau
        private const string FIRST_INVENTORY = "//table[@id='tableInventories']/tbody/tr/td[2]";
        private const string COMMENT_BUBBLE = "//*[@id='tableInventories']/tbody/tr/td[5]/div";

        // Filtres
        private const string RESET_FILTER = "//*[@id=\"formSearchInventories\"]/div[1]/a";
        private const string FILTER_NUMBER = "SearchNumber";
        private const string FILTER_SORTBY = "cbSortBy";
        private const string FILTER_ITEMS_NOT_VALIDATED = "ShowNotValidated";
        private const string FILTER_ALL = "//*[@id=\"ShowAll\"]";
        private const string FILTER_ACTIVE = "//*[@id=\"ShowActiveOnly\"]";
        private const string FILTER_INACTIVE = "//*[@id=\"ShowInactiveOnly\"]";
        private const string FILTER_NOT_VALIDATED = "//*[@id=\"ShowNotValidatedOnly\"]";
        private const string FILTER_SENT_SAGE = "ShowSentToSageOnl";
        //private const string FILTER_SENT_SAGE = "//*[@id=\"ShowSentToSageOnl\"]";
        private const string FILTER_VALIDATED_NOT_SENT_SAGE = "//*[@id=\"ShowValidatedAndNotSentToSage\"]";
        private const string FILTER_EXPORTED_FOR_SAGE_MANUALLY = "//*[@id=\"ShowExportedForSageManually\"]";
        private const string FILTER_DATEFROM = "date-picker-start";
        private const string FILTER_DATETO = "date-picker-end";
        private const string FILTER_SITE = "SelectedSites_ms";
        private const string SEARCH_SITE = "html/body/div[10]/div/div/label/input";
        private const string SITE_UNSELECT_ALL = "html/body/div[10]/div/ul/li[2]/a";

        private const string FILTER_EXPORTED_FOR_SAGE_CEGID_AUTOMATICALLY = "//*[@id=\"ShowExportedForSageAutomatically\"]";
        private const string FILTER_SENT_SAGE_CEGID_ONLY = "//*[@id=\"ShowSentToSageOnl\"]";

        private const string VALIDATE = "//*[@id=\"tableInventories\"]/tbody/tr[{0}]/td[1]/img[@alt='Valid']";
        private const string INACTIVE = "//*[@id=\"tableInventories\"]/tbody/tr[{0}]/td[1]/img[@alt='Inactive']";
        private const string INVENTORY_NUMBER = "//*[@id=\"tableInventories\"]/tbody/tr[*]/td[2]";
        private const string INVENTORY_SITE = "//*[@id=\"tableInventories\"]/tbody/tr[*]/td[3]";
        private const string INVENTORY_DATE = "//*[@id=\"tableInventories\"]/tbody/tr[*]/td[6]";
        private const string SENT_TO_SAGE_OLD = "//*[@id=\"tableInventories\"]/tbody/tr[{0}]/td[1]/img[2]";
        private const string SENT_TO_SAGE = "//*[@id=\"tableInventories\"]/tbody/tr[{0}]/td[1]/span[@title = 'Accounted (sent to SAGE manually)']";
        private const string SENT_TO_SAGE_MANUALLY = "//*[@id=\"tableInventories\"]/tbody/tr[{0}]/td[1]/span";
        private const string FILTER_SHOW_SENT_TO_SAGE_AND_IN_ERROR_ONLY = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[5]/input[6]";
        private const string SENT_TO_SAGE_AUTO = "//*[@id=\"tableInventories\"]/tbody/tr[{0}]/td[1]/i";

        private const string VALIDATEINVENTORY_BTN = "//*[@id=\"btn-validate-inventory\"]";
        private const string SUBMITVALIDATE_BTN = "//*[@id=\"btn-submit-validate-inventory\"]";
        private const string TYPEVALIDATE_BTN = "//*[@id=\"SetQtysToZero\"]/../../div";
        private const string UPDATEITEM_BTN = "//*[@id=\"btn-submit-update-items-inventory\"]";
        private const string SUBMIT_BUTTON = "//*[@id=\"btn-submit-create-inventory\"]";
        private const string ALL_RN_ROWS = "//*[@id='list-item-with-action']/table/tbody/tr[*]";
        private const string CEGID_ICON = ALL_RN_ROWS + "/td[1]/div[2]/i";
        private const string PAGESIZE = "/html/body/div[2]/div/div[2]/div[2]/nav/select";
        private const string CONFIIRM_EXPORT_SAGE = "btn-validate-export";

        //_______________________________ Variables _____________________________________

        // General
        [FindsBy(How = How.XPath, Using = NEW_INVENTORY_BUTTON)]
        private IWebElement _createNewInventory;

        [FindsBy(How = How.Id, Using = EXPORT)]
        private IWebElement _export;

        [FindsBy(How = How.Id, Using = ENABLE_EXPORT_FOR_SAGE)]
        private IWebElement _enableExportForSage;

        [FindsBy(How = How.Id, Using = EXPORT_FOR_SAGE)]
        private IWebElement _exportForSage;

        [FindsBy(How = How.Id, Using = CONFIRM_EXPORT_SAGE)]
        private IWebElement _confirmExportSage;

        [FindsBy(How = How.Id, Using = GENERATE_SAGE_TXT)]
        private IWebElement _generateSageTxt;

        [FindsBy(How = How.XPath, Using = EXPORT_WMS)]
        private IWebElement _exportWMS;

        // Tableau
        [FindsBy(How = How.XPath, Using = FIRST_INVENTORY)]
        private IWebElement _firstInventory;

        [FindsBy(How = How.XPath, Using = COMMENT_BUBBLE)]
        private IWebElement _commentBubble;

        [FindsBy(How = How.Id, Using = DELETE_CONFIRM)]
        private IWebElement _deleteConfirm;

        [FindsBy(How = How.XPath, Using = DELETE_CORBEILLE)]
        private IWebElement _deleteCorbeille;

        [FindsBy(How = How.XPath, Using = SUBMIT_BUTTON)]
        private IWebElement _submit;
        //________________________________ Filtres _______________________________________
        [FindsBy(How = How.XPath, Using = RESET_FILTER)]
        private IWebElement _resetFilter;

        [FindsBy(How = How.Id, Using = FILTER_NUMBER)]
        private IWebElement _searchByNumber;

        [FindsBy(How = How.Id, Using = FILTER_SORTBY)]
        private IWebElement _sortBy;

        [FindsBy(How = How.Id, Using = FILTER_ITEMS_NOT_VALIDATED)]
        private IWebElement _showItemsNotValidated;

        [FindsBy(How = How.XPath, Using = FILTER_ALL)]
        private IWebElement _showAll;

        [FindsBy(How = How.XPath, Using = FILTER_ACTIVE)]
        private IWebElement _showActive;

        [FindsBy(How = How.XPath, Using = FILTER_INACTIVE)]
        private IWebElement _showInactive;

        [FindsBy(How = How.XPath, Using = FILTER_NOT_VALIDATED)]
        private IWebElement _showNotValidated;

        [FindsBy(How = How.Id, Using = FILTER_SENT_SAGE)]
        private IWebElement _showSentToSAGE;

        [FindsBy(How = How.XPath, Using = FILTER_VALIDATED_NOT_SENT_SAGE)]
        private IWebElement _showValidatedNotSentToSAGE;

        [FindsBy(How = How.XPath, Using = FILTER_EXPORTED_FOR_SAGE_MANUALLY)]
        private IWebElement _showExportedForSageManually;

        [FindsBy(How = How.XPath, Using = FILTER_EXPORTED_FOR_SAGE_CEGID_AUTOMATICALLY)]
        private IWebElement _showExportedForSageCegidAutomatically;

        [FindsBy(How = How.XPath, Using = FILTER_SHOW_SENT_TO_SAGE_AND_IN_ERROR_ONLY)]
        private IWebElement _showSentToSAGEAndInErrorOnly;

        [FindsBy(How = How.XPath, Using = FILTER_SENT_SAGE_CEGID_ONLY)]
        private IWebElement _showSentToSAGECEGIDOnly;

        [FindsBy(How = How.Id, Using = FILTER_DATEFROM)]
        public IWebElement _dateFrom;

        [FindsBy(How = How.Id, Using = FILTER_DATETO)]
        public IWebElement _dateTo;

        [FindsBy(How = How.Id, Using = FILTER_SITE)]
        private IWebElement _siteFilter;

        [FindsBy(How = How.XPath, Using = SITE_UNSELECT_ALL)]
        private IWebElement _unselectAllSites;

        [FindsBy(How = How.XPath, Using = SEARCH_SITE)]
        private IWebElement _searchSite;

        [FindsBy(How = How.XPath, Using = VALIDATEINVENTORY_BTN)]
        private IWebElement _validateInvButton;

        [FindsBy(How = How.XPath, Using = SUBMITVALIDATE_BTN)]
        private IWebElement _submitValidateButton;

        [FindsBy(How = How.XPath, Using = TYPEVALIDATE_BTN)]
        private IWebElement _typevalidate;

        [FindsBy(How = How.XPath, Using = UPDATEITEM_BTN)]
        private IWebElement _updateItems;
        [FindsBy(How = How.XPath, Using = PAGESIZE)]
        private IWebElement _sizePage;
        public enum FilterType
        {
            ByNumber,
            SortBy,
            ShowItemsNotValidated,
            ShowAll,
            ShowActive,
            ShowInactive,
            ShowNotValidated,
            ShowSentToSAGE,
            ShowValidatedNotSentToSAGE,
            ExportedForSAGECEGIDAutomatically,
            ShowSentToSAGECEDIDOnly,
            ShowExportedForSageManually,
            ShowSentToSAGEAndInErrorOnly,
            DateFrom,
            DateTo,
            Site
        }

        public void Filter(FilterType filterType, object value)
        {
            Actions action = new Actions(_webDriver);

            switch (filterType)
            {
                case FilterType.ByNumber:
                    _searchByNumber = WaitForElementIsVisibleNew(By.Id(FILTER_NUMBER));
                    _searchByNumber.SetValue(ControlType.TextBox, value);
                    WaitForLoad();
                    break;
                case FilterType.SortBy:
                    _sortBy = WaitForElementIsVisible(By.Id(FILTER_SORTBY));
                    _sortBy.Click();
                    var element = WaitForElementIsVisible(By.XPath("//option[contains(@value,'" + value + "')]"));
                    _sortBy.SetValue(ControlType.DropDownList, element.Text);
                    _sortBy.Click();
                    break;
                case FilterType.ShowItemsNotValidated:
                    _showItemsNotValidated = WaitForElementExists(By.Id(FILTER_ITEMS_NOT_VALIDATED));
                    _showItemsNotValidated.SetValue(ControlType.CheckBox, value);
                    break;
                case FilterType.ExportedForSAGECEGIDAutomatically:
                    if (IsDev())
                    {
                        _showExportedForSageCegidAutomatically = WaitForElementExists(By.XPath(FILTER_EXPORTED_FOR_SAGE_CEGID_AUTOMATICALLY));
                        _showExportedForSageCegidAutomatically.SetValue(ControlType.CheckBox, value);
                    }
                    else
                    {
                        _showExportedForSageCegidAutomatically = WaitForElementExists(By.XPath("/html/body/div[2]/div/div[1]/div[1]/div/form/div[5]/input[7]"));
                        _showExportedForSageCegidAutomatically.SetValue(ControlType.CheckBox, value);
                    }
                    break;
                case FilterType.ShowSentToSAGECEDIDOnly:
                    if (IsDev())
                    {
                        _showSentToSAGECEGIDOnly = WaitForElementExists(By.XPath(FILTER_SENT_SAGE_CEGID_ONLY));
                        _showSentToSAGECEGIDOnly.SetValue(ControlType.CheckBox, value);
                    }
                    else

                    {
                        _showSentToSAGECEGIDOnly = WaitForElementExists(By.XPath("/html/body/div[2]/div/div[1]/div[1]/div/form/div[5]/input[5]"));
                        _showSentToSAGECEGIDOnly.SetValue(ControlType.CheckBox, value);
                    }
                    break;
                case FilterType.ShowAll:
                    _showAll = WaitForElementExists(By.XPath(FILTER_ALL));
                    _showAll.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.ShowActive:
                    _showActive = WaitForElementExists(By.XPath(FILTER_ACTIVE));
                    _showActive.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.ShowInactive:
                    _showInactive = WaitForElementExists(By.XPath(FILTER_INACTIVE));
                    _showInactive.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.ShowNotValidated:
                    _showNotValidated = WaitForElementExists(By.XPath(FILTER_NOT_VALIDATED));
                    _showNotValidated.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.ShowSentToSAGE:
                    WaitForElementExists(By.Id(FILTER_SENT_SAGE));
                    _showSentToSAGE.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.ShowValidatedNotSentToSAGE:
                    _showValidatedNotSentToSAGE = WaitForElementExists(By.XPath(FILTER_VALIDATED_NOT_SENT_SAGE));
                    _showValidatedNotSentToSAGE.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.ShowExportedForSageManually:
                    _showExportedForSageManually = WaitForElementExists(By.XPath(FILTER_EXPORTED_FOR_SAGE_MANUALLY));
                    _showExportedForSageManually.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.ShowSentToSAGEAndInErrorOnly:
                    _showSentToSAGEAndInErrorOnly = WaitForElementExists(By.Id(FILTER_SHOW_SENT_TO_SAGE_AND_IN_ERROR_ONLY));
                    _showSentToSAGEAndInErrorOnly.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.DateFrom:
                    _dateFrom = WaitForElementExists(By.Id(FILTER_DATEFROM));
                    var javaScriptExecutor0 = _webDriver as IJavaScriptExecutor;
                    javaScriptExecutor0.ExecuteScript("arguments[0].scrollIntoView(true);", _dateFrom);
                    _dateFrom.SetValue(ControlType.DateTime, value);
                    _dateFrom.SendKeys(Keys.Tab);
                    PageUp();
                    break;
                case FilterType.DateTo:
                    _dateTo = WaitForElementExists(By.Id(FILTER_DATETO));
                    var javaScriptExecutor1 = _webDriver as IJavaScriptExecutor;
                    javaScriptExecutor1.ExecuteScript("arguments[0].scrollIntoView(true);", _dateTo);
                    _dateTo.SetValue(ControlType.DateTime, value);
                    _dateTo.SendKeys(Keys.Tab);
                    PageUp();
                    break;
                case FilterType.Site:
                    _siteFilter = WaitForElementExists(By.Id(FILTER_SITE));
                    action.MoveToElement(_siteFilter).Perform();
                    _siteFilter.Click();

                    _searchSite = WaitForElementIsVisible(By.XPath(SEARCH_SITE));
                    _searchSite.SetValue(ControlType.TextBox, value);

                    _unselectAllSites = WaitForElementIsVisible(By.XPath(SITE_UNSELECT_ALL));
                    _unselectAllSites.Click();

                    var searchSiteValue = WaitForElementIsVisible(By.XPath("//span[text()='" + value + " - " + value + "']"));
                    searchSiteValue.Click();

                    _siteFilter.Click();
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);

            }

            WaitPageLoading();
            Thread.Sleep(1500);
            WaitForLoad();
        }

        public void ResetFilters()
        {
            _resetFilter = WaitForElementIsVisibleNew(By.XPath(RESET_FILTER));
            _resetFilter.Click();
            WaitPageLoading();
            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                //DateTo à 2 mois plus tard
                // le DateTo est en bas
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
                    _webDriver.FindElement(By.XPath(string.Format("//*[contains(@class,'table')]/tbody/tr[{0}]/td[1]/div/i[contains(@class,'circle-check')]", i + 1)));

                    if (!validated)
                        return false;
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
                    _webDriver.FindElement(By.XPath(String.Format("//*[contains(@class,'table')]/tbody/tr[{0}]/td[1]/div/i[contains(@class,'circle-xmark')]", i + 1)));


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

        public bool CheckShowItemsNotValidated(bool activated)
        {
            bool isValidated = false;
            int tot = CheckTotalNumber() > 100 ? 100 : CheckTotalNumber();

            if (tot == 0)
                return false;

            for (int i = 0; i < tot; i++)
            {
                try
                {
                    _webDriver.FindElement(By.XPath(string.Format(VALIDATE, i + 1)));
                }
                catch
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
                try
                {
                    try
                    {
                        var elm = _webDriver.FindElement(By.XPath(string.Format(SENT_TO_SAGE, i + 1)));

                        // a modifier 
                        if (!elm.GetAttribute("title").Contains("Accounted (sent to SAGE"))
                            return false;
                    }
                    catch
                    {
                        var elm = _webDriver.FindElement(By.XPath(string.Format(SENT_TO_SAGE_OLD, i + 1)));

                        // a modifier 
                        if (!elm.GetAttribute("title").Contains("Accounted (sent to SAGE"))
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
        public bool isAccounted()
        {
            ReadOnlyCollection<IWebElement> elements = _webDriver.FindElements(By.XPath("//*[starts-with(@id, 'wharehouse-inventory-details-')]/td[1]/div[2]/i"));
            return elements.All(elm =>
                elm.GetAttribute("title").Contains("Accounted (sent to SAGE manually)") ||
                elm.GetAttribute("title").Contains("Accounted")
            );

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
                    IWebElement elm;
                    if (isElementVisible(By.XPath(string.Format(SENT_TO_SAGE_MANUALLY, i + 1))))
                    {
                        elm = _webDriver.FindElement(By.XPath(string.Format(SENT_TO_SAGE_MANUALLY, i + 1)));
                    }
                    else
                    {
                        elm = _webDriver.FindElement(By.XPath(string.Format(SENT_TO_SAGE_AUTO, i + 1)));
                    }
                    // a modifier si patch a jour
                    if (!elm.GetAttribute("title").Contains("Accounted (sent to SAGE manually)") && !elm.GetAttribute("title").Contains("Accounted"))
                        return false;
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
            ReadOnlyCollection<IWebElement> listeNumbers;
            listeNumbers = _webDriver.FindElements(By.XPath("//*[contains(@class,'table')]/tbody/tr[*]/td[2]"));


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
            ReadOnlyCollection<IWebElement> sites;
            sites = _webDriver.FindElements(By.XPath("//*[contains(@class,'table')]/tbody/tr[*]/td[3]"));


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

            ReadOnlyCollection<IWebElement> dates;
            // bug Selenium ?
            dates = _webDriver.FindElements(By.XPath("//*/table[contains(@class,'table')]/tbody/tr[*]/td[6][string-length(text())>2]"));



            if (dates.Count == 0)
                return false;

            foreach (var elm in dates)
            {
                try
                {
                    DateTime date = DateTime.Parse(elm.Text, ci).Date;
                    if (date.Date.CompareTo(fromDate.Date) < 0 || date.Date.CompareTo(toDate.Date) > 0)
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

            ReadOnlyCollection<IWebElement> dates;
            dates = _webDriver.FindElements(By.XPath("//*[contains(@class,'table')]/tbody/tr[*]/td[6]"));

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

        public bool AreAllInveWithCegidIcon()
        {
            ReadOnlyCollection<IWebElement> inveStatus;

            inveStatus = _webDriver.FindElements(By.XPath(CEGID_ICON));

            if (inveStatus.Count == 0)
            { return false; }

            foreach (var elm in inveStatus)
            {
                string cls = elm.GetAttribute("class");
                if (cls.Contains("fa-save") == false || cls.Contains("fa-trash-alt") == false)
                {
                    return false;
                }
            }
            return true;
        }

        public bool AreAllInveValidatedAndNoCegid()
        {
            try
            {
                bool allValidated = this.CheckValidation(true);
                bool cegidIcon = this.AreAllInveWithCegidIcon();

                return allValidated && (cegidIcon == false);
            }
            catch
            {
                return false;
            }
        }

        //__________________________________________ Méthodes __________________________________________________

        // General
        public InventoryCreateModalPage InventoryCreatePage()
        {
            ShowPlusMenu();

            _createNewInventory = WaitForElementIsVisibleNew(By.XPath("//*/a[text()='New inventory']"), nameof(NEW_INVENTORY_BUTTON));
            _createNewInventory.Click();
            WaitForLoad();

            return new InventoryCreateModalPage(_webDriver, _testContext);
        }

        public void ExportExcelFile(bool printValue)
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

        public FileInfo GetExportExcelFile(FileInfo[] taskFiles, bool parExtension = false)
        {
            List<FileInfo> correctDownloadFiles = new List<FileInfo>();

            foreach (var file in taskFiles)
            {
                if (parExtension)
                {
                    if (file.Name.EndsWith(".xlsx"))
                    {
                        correctDownloadFiles.Add(file);
                    }
                }
                else
                {
                    //  Test REGEX
                    if (IsExcelFileCorrect(file.Name))
                    {
                        correctDownloadFiles.Add(file);
                    }
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


            Regex r = new Regex("^inventory" + space + annee + "-" + mois + "-" + jour + space + heure + "-" + minutes + "-" + secondes + ".xlsx$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
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

        public void ManualExportResultsForSage(bool printValue)
        {
            ShowExtendedMenu();

            _exportForSage = WaitForElementIsVisible(By.Id(EXPORT_FOR_SAGE));
            _exportForSage.Click();
            WaitForLoad();

            _confirmExportSage = WaitForElementIsVisible(By.Id(CONFIIRM_EXPORT_SAGE));
            _confirmExportSage.Click();
            WaitForLoad();

            if (printValue)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-alt']"));
                ClickPrintButton();
            }
            WaitForDownload();
            Close();
        }

        public void GenerateSageTxt()
        {
            ShowExtendedMenu();

            _generateSageTxt = WaitForElementIsVisible(By.Id(GENERATE_SAGE_TXT));
            _generateSageTxt.Click();
            WaitForLoad();

            _confirmExportSage = WaitForElementIsVisible(By.Id(CONFIIRM_EXPORT_SAGE));
            _confirmExportSage.Click();
            WaitForLoad();

            IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-alt']"));
            ClickPrintButton();

            WaitForDownload();
            Close();
        }

        public void ManualExportSageError(bool printValue)
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
                IsFileInError(By.CssSelector("[class='fa fa-info-circle']"));
                ClickPrintButton();
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
            // inventory 2020-12-15 16-10-17.txt

            string mois = "(?:0[1-9]|1[0-2])";         // mois
            string space = "(\\s)";                    // Espace
            string annee = "\\d{4}";                   // annee YYYY
            string jour = "[0-3]\\d";                  // jour
            string heure = "(?:0[0-9]|1[0-9]|2[0-3])"; // heure
            string minutes = "[0-5]\\d";               // minutes
            string secondes = "[0-5]\\d";              // secondes


            Regex r = new Regex("^inventory" + space + annee + "-" + mois + "-" + jour + space + heure + "-" + minutes + "-" + secondes + ".txt$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
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

        // Tableau
        public InventoryItem SelectFirstInventory()
        {
            _firstInventory = WaitForElementIsVisible(By.XPath("//table[contains(@class,'table')]/tbody/tr/td[2]"));

            _firstInventory.Click();
            WaitForLoad();

            return new InventoryItem(_webDriver, _testContext);
        }

        public string GetFirstID()
        {
            _firstInventory = WaitForElementIsVisible(By.XPath("//table[contains(@class,'table')]/tbody/tr/td[2]"));

            return _firstInventory.Text;
        }

        public string GetCommentBubble()
        {
            _commentBubble = WaitForElementIsVisible(By.XPath("//*[contains(@class,'table')]/tbody/tr/td[5]/div"));

            return _commentBubble.GetAttribute("title");
        }

        public bool IsCommentBubbleGreen()
        {
            _commentBubble = WaitForElementIsVisible(By.XPath("//*[contains(@class,'table')]/tbody/tr/td[5]/div"));

            return _commentBubble.GetAttribute("class").Contains("green-text");
        }

        public void CheckExport(FileInfo trouveXLSX)
        {
            //Les validated sont tous présent dans le fichier Excel sauf ceux sans item (théo et physi qtty à 0)
            List<string> idsXLSX = OpenXmlExcel.GetValuesInList("Id", "Inventory", trouveXLSX.FullName);
            ReadOnlyCollection<IWebElement> Ids;
            Ids = _webDriver.FindElements(By.XPath("//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[2]"));

            List<string> idsDisplayed = new List<string>();
            for (int i = 0; i < Ids.Count; i++)
            {
                idsDisplayed.Add(convertId(Ids[i].Text));
            }
            foreach (var Id in idsXLSX)
            {
                if (Id == "") continue;
                Assert.IsTrue(idsDisplayed.Contains(Id), "Number " + Id + " présent dans le fichier excel mais pas affiché");
            }
        }

        public string convertId(string id)
        {
            //Entrée : 032858
            //Sortie :  32858
            int counter = 0;
            while (id.StartsWith("0") && counter < 10)
            {
                id = id.Substring(1);
                counter++;
            }
            return id;
        }
        public void deleteInventory()
        {
            _deleteCorbeille = WaitForElementIsVisibleNew(By.XPath("//*/table[contains(@class,'table')]/tbody/tr[*]/td[1][not(div)]/../td[7]/div//a"));

            _deleteCorbeille.Click();
            WaitForLoad();

            _deleteConfirm = WaitForElementIsVisibleNew(By.Id(DELETE_CONFIRM));
            _deleteConfirm.Click();
            WaitPageLoading();
        }

        public bool VerifyNotValidated()
        {
            if (isElementVisible(By.XPath("//*/table[contains(@class,'table')]/tbody/tr[*]/td[1][not(div)]/../td[7]/div//a")))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        public void ValidateInventory()
        {
            WaitForLoad();
            _validateInvButton = WaitForElementIsVisible(By.XPath(VALIDATEINVENTORY_BTN));
            _validateInvButton.Click();
            WaitForLoad();
        }


        public bool SubmitValidateInventory()
        {
            var modalXPath = "//*[@id=\"modal-1\"]/div/div/div/div/div/form/div[1]";

            if (isElementVisible(By.XPath(modalXPath)))
            {

                if (isElementVisible(By.XPath(TYPEVALIDATE_BTN)))
                {
                    _typevalidate = WaitForElementIsVisible(By.XPath(TYPEVALIDATE_BTN));
                    _typevalidate.Click();
                    WaitForLoad();

                    _updateItems = WaitForElementIsVisible(By.XPath(UPDATEITEM_BTN));
                    _updateItems.Click();
                    WaitForLoad();
                }

                _submitValidateButton = WaitForElementIsVisible(By.XPath(SUBMITVALIDATE_BTN));
                _submitValidateButton.Click();

                WaitForLoad();
                return true;
            }
            else
            {
                return false;
            }
        }

        public InventoryItem Submit()
        {
            _submit = WaitForElementIsVisible(By.XPath(SUBMIT_BUTTON));
            _submit.Click();
            return new InventoryItem(_webDriver, _testContext);
        }


        public InventoryItem ClickFirstInventory()
        {
            var firstInventory = WaitForElementIsVisible(By.XPath("//*[@id=\"list-item-with-action\"]/table/tbody/tr"));
            firstInventory.Click();
            return new InventoryItem(_webDriver, _testContext);
        }

        public string GetExportWMSMessage()
        {
            ShowExtendedMenu();
            WaitPageLoading();
            WaitForLoad();
            _exportWMS = WaitForElementIsVisible(By.XPath(EXPORT_WMS));
            WaitForLoad();
            _exportWMS.Click();
            WaitForLoad();
            var popupMessage = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div/div[2]/div/span")).Text;
            var close = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div/div[3]/button"));
            WaitForLoad();
            // close.Click();
            // WaitPageLoading(); 
            return popupMessage;
        }
        public string getNumberOfInventories()
        {
            var element = WaitForElementIsVisible(By.XPath("/html/body/div[2]/div/div[2]/div[1]/h1/span"));
            return element.Text;
        }
        public bool InventoryItemsiSame(FileInfo correctDownloadedFile)
        {
            var listOfIds = new List<string>();
            var inventoryIDFromExcel = OpenXmlExcel.GetValuesInList("Id", "Inventory", correctDownloadedFile.FullName).Distinct().Take(8).ToList();
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(10));
            int i = 8;
            var table = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='list-item-with-action']/table")));

            IList<IWebElement> tableRows = table.FindElements(By.XPath(".//tbody/tr[*]/td[2]"));
            foreach (var row in tableRows)
            {
                listOfIds.Add(row.Text.TrimStart('0'));
                i--;
                if (i == 0)
                {
                    break;
                }
            }

            var isEqual = inventoryIDFromExcel.SequenceEqual(listOfIds);
            return isEqual;
        }
        public string GetFirstDate()
        {
            var _firstDate = WaitForElementIsVisible(By.XPath("//table[contains(@class,'table')]/tbody/tr/td[6]"));

            return _firstDate.Text;
        }
        public string GetFirstInventoryNumber()
        {
            IWebElement firstRecipeNoteNumber = WaitForElementIsVisible(By.XPath("//*[@id=\"wharehouse-inventory-details-1\"]/td[2]"));
            return firstRecipeNoteNumber.Text;
        }

    }
}
