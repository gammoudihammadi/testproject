using DocumentFormat.OpenXml.Drawing.Charts;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
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
using System.IO.Packaging;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.OutputForm
{
    public class OutputFormItem : PageBase
    {

        public OutputFormItem(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //_______________________________________ CONSTANTES ____________________________________________________

        // Général
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";
        private const string VALIDATE = "btn-validate-output-form";
        private const string CONFIRM_VALIDATE = "btn-popup-validate";
        private const string EXTENDED_BUTTON = "//*[@id=\"div-body\"]/div/div[1]/div/div/button";
        private const string REFRESH = "btn-output-form-refresh";
        private const string ALERT_MODAL = "//*[@id=\"dataAlertModal\"]/div/div";
        private const string PRINT = "//*[@id=\"div-body\"]/div/div[1]/div/div[1]/div/a[2]";
        private const string VALIDATE_PRINT = "validatePrint";
        private const string EXPORT_SAGE = "btn-export-for-sage";
        private const string CONFIRM_EXPORT_SAGE = "btn-popup-validate";
        private const string CANCEL_SAGE = "btn-cancel-popup";
        private const string DOWNLOAD_FILE = "btn-download-file";
        private const string ENABLE_EXPORT_FOR_SAGE = "btn-enable-export-for-sage";

        // Onglets
        private const string GENERAL_INFORMATION_TAB = "hrefTabContentInformations";
        private const string GROUPS_TAB = "hrefTabContentGroups";
        private const string QUALITY_CHECKS_TAB = "hrefTabContentQualityChecks";
        private const string FOOTER_TAB = "hrefTabContentFooter";
        private const string ACCOUNTING_TAB = "hrefTabContentExportSageWriting";
        // Items 
        private const string GROUP = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/span";
        private const string UPDATED_ICON = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[1]/span/img";
        private const string ITEM_NAME = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[2]/div[3]/span";
        private const string ITEM_GROUP = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[1]/div[@class=\"display-zone\"]";
        private const string ITEM_GROUPS = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div[@class=\"display-zone\"]";
        private const string ITEM_THEORICAL_QUANTITY = "//*/div[contains(@class,'display-zone')]/div[10]/span";
        private const string ITEM_THEORICAL_QUANTITY_BIS = "//*/div[contains(@class,'edit-zone')]/div[10]/span";
        private const string THEORICAL_QUANTITY_INPUT = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[@title='{0}']/../../div[9]/input";
        private const string ITEM_PHYSICAL_QUANTITY = "//*/div[contains(@class,'display-zone')]/div/span[@name='TotalPhysicalQuantity']";
        private const string ITEM_PHYSICAL_QUANTITY_BIS = "//*/div[contains(@class,'edit-zone')]/div/span[@name='TotalPhysicalQuantity']";
        private const string ITEM_PRICE = "//*/div[contains(@class,'edit-zone')]/div/span[@name='ItemTotalPrice']";
        private const string PHYSICAL_QUANTITY_INPUT = "//span[@title='{0}']/../..//input[@id='item_OutputFormDetail_Quantity']";
        private const string ITEM_ERROR = "//*/div[contains(@class,'edit-zone')]/div/input[@name='item.OutputFormDetail.Quantity']";
        private const string EDIT_ITEM = "/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/form/div[1]/div[15]/ul/li[1]/a";
        private const string ACTIONS_BTN = "/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/form/div[1]/div[15]/a";

        private const string COMMENT_ICON = "//*/div[contains(@class,'edit-zone')]/div/span[@title='{0}']/../../div[15]/div/a[*]/span[@class = 'fas fa-message glyph-span ']";
        private const string COMMENT = "Comment";
        private const string SAVE_COMMENT = "//*/button[text()='Save']";
        private const string DELETE_ITEM_2 = "/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/form/div[1]/div[15]/ul/li[2]/a[2]";

        private const string EXPIRY_DATE = "(//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[2]/div[12]/a/parent::div)[1]";
        private const string EXPIRY_DATE_1 = "(//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[1]/div[12]/a/parent::div)[1]";

        private const string ALLERGENS_BTN = "/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/form/div[2]/div[14]/a";
        private const string ALLERGENS_LIST = "/html/body/div[4]/div/div/div/div[2]/div/div/div/div/ul";
        private const string ITEM_DETAIL = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[1]/div[15]/a";
        private const string OUTPUT_FORM_NUMBER = "//*[@id=\"div-body\"]/div/div[1]/h1";


        // Filtres
        private const string RESET_FILTER = "//*[@id=\"formSearchItems\"]/div[1]/a";
        private const string FILTER_NAME = "tbSearchPatternWithAutocomplete";
        private const string FILTER_SORTBY = "cbSortBy";
        private const string SHOW_ITEMS_WITH_THEO_QTY = "ItemIndexVM_StockShowTheoricalQtyOnly";
        private const string SHOW_ITEMS_WITH_PHYS_QTY = "ItemIndexVM_StockShowEditedPhysicalQty";
        private const string FILTER_GROUP = "ItemIndexVMSelectedGroups_ms";
        private const string FILTER_SUBGROUP = "ItemIndexVMSelectedSubGroups_ms";
        private const string PHY_QUANTITY_INPUT = "item_OutputFormDetail_Quantity";
        private const string BUTTON_SIZE = "collapse-left-pane";
        private const string ELEMENT_ITEM = "//*[@id=\"item-filter-form\"]/div[1]";




        //________________________________________________ Variables ______________________________________________________

        // Général
        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;

        [FindsBy(How = How.Id, Using = VALIDATE)]
        private IWebElement _validate;

        [FindsBy(How = How.Id, Using = CONFIRM_VALIDATE)]
        private IWebElement _confirmValidate;

        [FindsBy(How = How.XPath, Using = EXTENDED_BUTTON)]
        private IWebElement _extendedButton;

        [FindsBy(How = How.Id, Using = REFRESH)]
        private IWebElement _refresh;

        [FindsBy(How = How.XPath, Using = PRINT)]
        private IWebElement _print;

        [FindsBy(How = How.Id, Using = VALIDATE_PRINT)]
        private IWebElement _validatePrint;

        [FindsBy(How = How.Id, Using = EXPORT_SAGE)]
        private IWebElement _exportSage;

        [FindsBy(How = How.Id, Using = CONFIRM_EXPORT_SAGE)]
        private IWebElement _confirmExportSage;

        [FindsBy(How = How.Id, Using = CANCEL_SAGE)]
        private IWebElement _cancelSage;

        [FindsBy(How = How.Id, Using = DOWNLOAD_FILE)]
        private IWebElement _downloadFile;

        [FindsBy(How = How.Id, Using = ENABLE_EXPORT_FOR_SAGE)]
        private IWebElement _enableExportForSage;

        // Onglets
        [FindsBy(How = How.Id, Using = GENERAL_INFORMATION_TAB)]
        private IWebElement _generalInformationTab;

        [FindsBy(How = How.Id, Using = GROUPS_TAB)]
        private IWebElement _groupsTab;

        [FindsBy(How = How.Id, Using = QUALITY_CHECKS_TAB)]
        private IWebElement _qualityChecksTab;

        [FindsBy(How = How.Id, Using = FOOTER_TAB)]
        private IWebElement _footerTab;

        [FindsBy(How = How.Id, Using = ACCOUNTING_TAB)]
        private IWebElement _accountingTab;

        // Tableau
        [FindsBy(How = How.XPath, Using = ITEM_NAME)]
        private IWebElement _itemName;

        [FindsBy(How = How.XPath, Using = ITEM_GROUP)]
        private IWebElement _itemGroup;

        [FindsBy(How = How.XPath, Using = THEORICAL_QUANTITY_INPUT)]
        private IWebElement _theoricalQuantityInput;

        [FindsBy(How = How.XPath, Using = PHYSICAL_QUANTITY_INPUT)]
        private IWebElement _physicalQuantityInput;

        [FindsBy(How = How.XPath, Using = ITEM_PHYSICAL_QUANTITY)]
        private IWebElement _physicalQuantity;

        [FindsBy(How = How.XPath, Using = ITEM_PRICE)]
        private IWebElement _itemPrice;

        [FindsBy(How = How.XPath, Using = ITEM_ERROR)]
        private IWebElement _itemError;

        [FindsBy(How = How.XPath, Using = EDIT_ITEM)]
        private IWebElement _editItem;

        [FindsBy(How = How.XPath, Using = COMMENT_ICON)]
        private IWebElement _addComment;

        [FindsBy(How = How.Id, Using = COMMENT)]
        private IWebElement _comment;

        [FindsBy(How = How.Id, Using = SAVE_COMMENT)]
        private IWebElement _saveComment;

        [FindsBy(How = How.XPath, Using = DELETE_ITEM_2)]
        private IWebElement _deleteItem;

        [FindsBy(How = How.XPath, Using = EXPIRY_DATE)]
        private IWebElement _expiryDate;

        [FindsBy(How = How.Id, Using = ALLERGENS_BTN)]
        private IWebElement _allergensBtn;

        [FindsBy(How = How.Id, Using = ALLERGENS_LIST)]
        private IWebElement _allergensList;

        [FindsBy(How = How.Id, Using = ACTIONS_BTN)]
        private IWebElement _actionsBtn;
        [FindsBy(How = How.Id, Using = BUTTON_SIZE)]
        private IWebElement _buttonSize;
        [FindsBy(How = How.Id, Using = ITEM_DETAIL)]
        private IWebElement _delete;





        //______________________________________ Filtres ______________________________________________________

        [FindsBy(How = How.XPath, Using = RESET_FILTER)]
        private IWebElement _resetFilter;

        [FindsBy(How = How.Id, Using = FILTER_NAME)]
        private IWebElement _searchByNameFilter;

        [FindsBy(How = How.Id, Using = FILTER_SORTBY)]
        private IWebElement _sortByFilter;

        [FindsBy(How = How.Id, Using = SHOW_ITEMS_WITH_THEO_QTY)]
        private IWebElement _showItemsWithTheoQtyOnly;

        [FindsBy(How = How.Id, Using = SHOW_ITEMS_WITH_PHYS_QTY)]
        private IWebElement _showItemsWithPhysQtyOnly;

        public enum FilterItemType
        {
            SearchByName,
            SortBy,
            ShowItemsWithTheoQtyOnly,
            ShowItemsWithPhysQty,
            ByGroup,
            BySubGroup

        }

        public void Filter(FilterItemType FilterItemType, object value)
        {
            switch (FilterItemType)
            {
                case FilterItemType.SearchByName:
                    _searchByNameFilter = WaitForElementIsVisible(By.Id(FILTER_NAME));
                    _searchByNameFilter.SetValue(ControlType.TextBox, value);
                    break;
                case FilterItemType.SortBy:
                    _sortByFilter = WaitForElementIsVisible(By.Id(FILTER_SORTBY));
                    _sortByFilter.Click();
                    var element = WaitForElementIsVisible(By.XPath("//option[contains(@value,'" + value + "')]"));
                    _sortByFilter.SetValue(ControlType.DropDownList, element.Text);
                    _sortByFilter.Click();
                    Thread.Sleep(2000);
                    break;
                case FilterItemType.ShowItemsWithTheoQtyOnly:
                    _showItemsWithTheoQtyOnly = WaitForElementExists(By.Id(SHOW_ITEMS_WITH_THEO_QTY));
                    _showItemsWithTheoQtyOnly.SetValue(ControlType.CheckBox, value);
                    break;
                case FilterItemType.ShowItemsWithPhysQty:
                    _showItemsWithPhysQtyOnly = WaitForElementExists(By.Id(SHOW_ITEMS_WITH_PHYS_QTY));
                    _showItemsWithPhysQtyOnly.SetValue(ControlType.CheckBox, value);
                    break;
                case FilterItemType.ByGroup:
                    ComboBoxSelectById(new ComboBoxOptions(FILTER_GROUP, (string)value));
                    break;
                case FilterItemType.BySubGroup:
                    ComboBoxSelectById(new ComboBoxOptions(FILTER_SUBGROUP, (string)value));
                    break;
            }

            WaitPageLoading();
            WaitForLoad();
        }

        public void ResetFilter()
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
            var elements = _webDriver.FindElements(By.XPath(ITEM_NAME));

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
            var elements = _webDriver.FindElements(By.XPath(ITEM_NAME));

            if (elements.Count == 0)
                return false;

            var ancientName = elements[0].GetAttribute("title");

            foreach (var elm in elements)
            {
                if (string.Compare(ancientName, elm.GetAttribute("title")) > 0)
                    return false;

                ancientName = elm.GetAttribute("title");
            }

            return true;
        }

        public bool IsSortedByItemGroup()
        {
            var elements = _webDriver.FindElements(By.XPath(ITEM_GROUPS));

            if (elements.Count == 0)
                return false;

            var ancientName = elements[0].Text;

            foreach (var elm in elements)
            {
                if (string.Compare(ancientName, elm.Text) > 0)
                    return false;

                ancientName = elm.Text;
            }

            return true;
        }

        public bool IsWithPositiveTheoQty()
        {
            var elements = _webDriver.FindElements(By.XPath(ITEM_THEORICAL_QUANTITY));

            if (elements.Count == 0)
                return false;

            foreach (var elm in elements)
            {
                if (elm.Text.Equals("0"))
                {
                    return false;
                }
            }
            return true;
        }

        public bool IsWithEditedPhysQty()
        {
            ReadOnlyCollection<IWebElement> elements;
            elements = _webDriver.FindElements(By.XPath(ITEM_PHYSICAL_QUANTITY));

            if (elements.Count == 0)
                return false;

            foreach (var elm in elements)
            {
                if (elm.Text.Equals("0"))
                {
                    return false;
                }
            }
            return true;
        }

        public bool VerifyGroup(string value)
        {
            var elements = _webDriver.FindElements(By.XPath(ITEM_GROUPS));

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

        //______________________________________ Méthodes _____________________________________________________

        // Général
        public OutputFormPage BackToList()
        {
            _backToList = WaitForElementIsVisibleNew(By.XPath(BACK_TO_LIST));
            _backToList.Click();
            WaitForLoad();

            return new OutputFormPage(_webDriver, _testContext);
        }

        public void Validate()
        {
            ShowValidationMenu();

            _validate = WaitForElementIsVisible(By.Id(VALIDATE));
            _validate.Click();
            WaitForLoad();

            _confirmValidate = WaitForElementIsVisible(By.Id(CONFIRM_VALIDATE));
            _confirmValidate.Click();
            WaitForLoad();
        }

        public override void ShowExtendedMenu()
        {
            var actions = new Actions(_webDriver);
            _extendedButton = WaitForElementIsVisible(By.XPath(EXTENDED_BUTTON));
            actions.MoveToElement(_extendedButton).Perform();
            WaitForLoad();
        }

        public void Refresh()
        {
            ShowExtendedMenu();

            _refresh = WaitForElementIsVisible(By.Id(REFRESH));
            _refresh.Click();
            WaitForLoad();
        }

        public bool ShowRefreshAlertModal()
        {
            if (isElementVisible(By.XPath(ALERT_MODAL)))
            {
                WaitForElementIsVisible(By.XPath(ALERT_MODAL));
                return true;
            }
            else
            {
                return false;
            }
        }

        public PrintReportPage Print(bool printValue)
        {
            ShowExtendedMenu();

            _print = WaitForElementIsVisible(By.XPath(PRINT));
            _print.Click();
            WaitForLoad();

            _validatePrint = WaitForElementIsVisible(By.Id("btn-validate-print-output-form"));
            _validatePrint.Click();
            WaitForLoad();

            if (printValue)
            {
                WaitPageLoading();
                WaitPageLoading();
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-pdf']"));
                ClickPrintButton();

            }

            //Results are opened in a new tab, switch the driver to the newly created one
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new PrintReportPage(_webDriver, _testContext);
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

            if (!printValue)
            {
                if (isElementVisible(By.Id(CONFIRM_EXPORT_SAGE)))
                {
                    _confirmExportSage = _webDriver.FindElement(By.Id(CONFIRM_EXPORT_SAGE));
                    _confirmExportSage.Click();
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
                IsFileInError(By.CssSelector("[class='fa fa-info-circle']"));
                ClickPrintButton();
            }
        }

        public void ManualExportSAGE(bool printValue)
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
                WaitForLoad();
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

        public void EnableExportForSage()
        {
            _enableExportForSage = WaitForElementIsVisible(By.Id(ENABLE_EXPORT_FOR_SAGE));
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

        // Onglets

        public OutputFormGeneralInformation ClickOnGeneralInformationTab()
        {
            _generalInformationTab = WaitForElementIsVisibleNew(By.Id(GENERAL_INFORMATION_TAB));
            _generalInformationTab.Click();
            WaitForLoad();

            return new OutputFormGeneralInformation(_webDriver, _testContext);
        }

        public OutputFormGroups ClickOnGroupsTab()
        {
            _groupsTab = WaitForElementIsVisible(By.Id(GROUPS_TAB));
            _groupsTab.Click();
            WaitForLoad();

            return new OutputFormGroups(_webDriver, _testContext);
        }

        public OutputFormQualityChecks ClickOnChecksTab()
        {
            _qualityChecksTab = WaitForElementIsVisible(By.Id(QUALITY_CHECKS_TAB));
            _qualityChecksTab.Click();
            WaitForLoad();

            return new OutputFormQualityChecks(_webDriver, _testContext);
        }

        public OutputFormFooterPage ClickOnFooterTab()
        {
            _footerTab = WaitForElementIsVisible(By.Id(FOOTER_TAB));
            _footerTab.Click();
            WaitForLoad();
            WaitPageLoading();

            return new OutputFormFooterPage(_webDriver, _testContext);
        }

        public OutputFormAccountingPage ClickOnAccountingTab()
        {
            _accountingTab = WaitForElementIsVisible(By.Id(ACCOUNTING_TAB));
            _accountingTab.Click();
            WaitForLoad();

            return new OutputFormAccountingPage(_webDriver, _testContext);
        }


        // Tableau

        public void SelectFirstItem()
        {
            WaitForLoad();
            try
            {
                _itemName = WaitForElementIsVisibleNew(By.XPath(ITEM_NAME), nameof(ITEM_NAME));
                _itemName.Click();
                WaitForLoad();
            }
            catch
            {
                _itemName = WaitForElementExists(By.XPath("//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[1]/div[3]/span"));
                _itemName.Click();
            }
            WaitForLoad();
        }
        public void SelectItem(string itemName)
        {
            WaitForLoad();
            try
            {
                _itemName = WaitForElementIsVisible(By.XPath(String.Format("//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[2]/div[3]/span[text()= '{0}']", itemName)));
                _itemName.Click();
                WaitForLoad();
            }
            catch
            {
                _itemName = WaitForElementExists(By.XPath("//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[1]/div[3]/span"));
                _itemName.Click();
            }
        }
        public string GetFirstItemName()
        {
            if (isElementExists(By.XPath(ITEM_NAME)))
            {
                _itemName = WaitForElementExists(By.XPath(ITEM_NAME));
                return _itemName.GetAttribute("title");
            }
            else
            {
                return "";
            }
        }
        public string VerifFirstName(string firstName)
        {
            if (isElementExists(By.XPath(ITEM_NAME)))
            {
                var elements = _webDriver.FindElements(By.XPath(ITEM_NAME));

                if (elements.Count == 0)
                    return "";

                foreach (var elm in elements)
                {
                    if (elm.Text == firstName)
                    {
                        return elm.Text;
                    }

                }
                return "";
            }
            else
            {
                return "";
            }
        }
        public List<String> GetItemNames()
        {
            List<String> itemNames = new List<String>();
            WaitLoading();
            var items = _webDriver.FindElements(By.XPath(ITEM_NAME));

            foreach (var item in items)
            {
                itemNames.Add(item.GetAttribute("title"));
            }

            return itemNames;
        }

        public string GetFirstItemGroup()
        {
            _itemGroup = WaitForElementExists(By.XPath(ITEM_GROUP));
            return _itemGroup.Text;
        }

        public bool IsGroupDisplayActive()
        {
            if (isElementVisible(By.XPath(GROUP)))
            {
                _webDriver.FindElement(By.XPath(GROUP));
                return true;
            }
            else
            { return false; }
        }

        public string GetTheoricalQuantity()
        {
            if (isElementVisible(By.XPath(ITEM_THEORICAL_QUANTITY)))
            {
                _theoricalQuantityInput = WaitForElementExists(By.XPath(ITEM_THEORICAL_QUANTITY));
            }
            else
            {
                _theoricalQuantityInput = WaitForElementExists(By.XPath(ITEM_THEORICAL_QUANTITY_BIS));
            }
            return Regex.Replace(_theoricalQuantityInput.Text, @"\s+", "").ToString();//remove space in string
        }

        public string GetPhysicalQuantity()
        {
            if (isElementVisible(By.XPath(ITEM_PHYSICAL_QUANTITY)))
            {
                _physicalQuantity = WaitForElementIsVisible(By.XPath(ITEM_PHYSICAL_QUANTITY), nameof(ITEM_PHYSICAL_QUANTITY));
            }
            else
            {
                _physicalQuantity = WaitForElementIsVisible(By.XPath(ITEM_PHYSICAL_QUANTITY_BIS), nameof(ITEM_PHYSICAL_QUANTITY_BIS));
            }
            return Regex.Replace(_physicalQuantity.Text, @"\s+", "").ToString(); ;
        }

        public double GetPhysicalQuantitySum()
        {
            WaitLoading(); 
            var _physicalQuantity = _webDriver.FindElements(By.XPath(ITEM_PHYSICAL_QUANTITY_BIS));
            double sum = 0;
            foreach (var physQty in _physicalQuantity)
            {
                sum += double.Parse(Regex.Replace(physQty.Text, @"\s+", "").ToString());
            }
            return sum;
        }


        public void AddPhysicalQuantity(string itemName, string quantity)
        {
            // remove disquette
            //IJavaScriptExecutor js = (IJavaScriptExecutor)_webDriver;
            //js.ExecuteScript("$(\"span[name='IconSavedVisible']\").addClass(\"hidden\");");

            _physicalQuantityInput = WaitForElementIsVisibleNew(By.XPath(String.Format(PHYSICAL_QUANTITY_INPUT, itemName)));
            _physicalQuantityInput.SetValue(ControlType.TextBox, quantity);
            //la disquette apparait avant la sauvegarde data binding
            Thread.Sleep(2000);
            //WaitForElementIsVisible(By.XPath(String.Format(SAVE_LINE_ICON, itemName)));
            WaitForLoad();
        }
        public void AddPhyQuantity(string itemName, string quantity)
        {
            Filter(FilterItemType.SearchByName, itemName);
            SelectFirstItem();
            var physicalQty = WaitForElementIsVisible(By.Id(PHY_QUANTITY_INPUT));
            physicalQty.Clear();
            physicalQty.SendKeys(quantity);
            WaitPageLoading();
            WaitForLoad();
        }

        public void AddPhysicalQuantityOverload(string itemName, string quantity)
        {
            _physicalQuantityInput = WaitForElementIsVisible(By.XPath(String.Format(PHYSICAL_QUANTITY_INPUT, itemName)));
            _physicalQuantityInput.SetValue(ControlType.TextBox, quantity);

            // Prise en compte de la mise à jour : ne pas enlever
            Thread.Sleep(3000);
            WaitForLoad();
        }

        public bool IsUpdated()
        {
            if (isElementVisible(By.XPath(UPDATED_ICON)))
            {
                WaitForElementIsVisible(By.XPath(UPDATED_ICON));
                return true;
            }
            else
            {
                return false;
            }
        }

        public bool IsFailed()
        {
            if (isElementVisible(By.XPath("//*[contains(@title,'Negative stock quantities are not allowed')]")))
            {
                return true;
            }

            return false;
        }

        public bool ValidationFailed()
        {
            if (isElementVisible(By.XPath("//*[@id='modal-1']/div/div/div/form/div[2]/p[1]/b")))
            {
                var message1 = WaitForElementIsVisible(By.XPath("//*[@id='modal-1']/div/div/div/form/div[2]/p[1]/b"));
                Assert.AreEqual("The Output form cannot be validated because negative stock are not allowed. The following items are faulty:", message1.Text);
                var message2 = WaitForElementIsVisible(By.XPath("//*[@id='modal-1']/div/div/div/form/div[2]/p[2]"));
                Assert.AreEqual("BRANDY TORRES 10 AÑOS MINI", message2.Text);
                return true;
            }
            else
            {
                return false;
            }

        }

        public ItemGeneralInformationPage EditItem(string itemName)
        {
            WaitPageLoading();
            _actionsBtn = WaitForElementExists(By.XPath(string.Format(ACTIONS_BTN, itemName)));
            IJavaScriptExecutor js = (IJavaScriptExecutor)_webDriver;
            js.ExecuteScript("arguments[0].click();", _actionsBtn);

            WaitPageLoading();
            _editItem = WaitForElementExists(By.XPath(string.Format(EDIT_ITEM, itemName)));
            js.ExecuteScript("arguments[0].click();", _editItem);
            WaitPageLoading();

            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            WaitPageLoading();

            return new ItemGeneralInformationPage(_webDriver, _testContext);
        }

        public void AddComment(string itemName, string comment)
        {
            // [...]
            var menuEdit = WaitForElementIsVisible(By.XPath(string.Format("//*/div[contains(@class,'edit-zone')]/div/span[@title='{0}']/../../div[15]/a", itemName)));
            menuEdit.Click();
            WaitForLoad();
            var commentEdit = WaitForElementIsVisible(By.XPath(string.Format("//*/div[contains(@class,'edit-zone')]/div/span[@title='{0}']/../../div[15]/ul/li[2]/a[1]/span[contains(@class,'message')]/..", itemName)));
            commentEdit.Click();
            WaitForLoad();

            _comment = WaitForElementIsVisible(By.Id("OutputFormDetailComment"));
            _comment.SetValue(ControlType.TextBox, comment);
            WaitForLoad();

            _saveComment = WaitForElementToBeClickable(By.XPath(SAVE_COMMENT));
            _saveComment.Click();
            WaitForLoad();
        }
        public void ClickOnButtonSize()
        {
            _buttonSize = WaitForElementIsVisible(By.Id(BUTTON_SIZE));

            _buttonSize.Click();

        }

        public string GetComment(string itemName)
        {
            try
            {
                // [...]
                var menuEdit = WaitForElementIsVisible(By.XPath(string.Format("//*/div[contains(@class,'edit-zone')]/div/span[@title='{0}']/../../div[15]/a", itemName)));
                menuEdit.Click();
                WaitForLoad();
            }
            catch
            {
                // [...] menuEdit déjà ouvert
            }
            var commentEdit = WaitForElementIsVisible(By.XPath(string.Format("//*/div[contains(@class,'edit-zone')]/div/span[@title='{0}']/../../div[15]/ul/li[2]/a[1]/span[contains(@class,'message')]/..", itemName)));
            commentEdit.Click();
            WaitForLoad();

            if (isElementVisible(By.Id("OutputFormDetailComment")))
            {
                _comment = WaitForElementExists(By.Id("OutputFormDetailComment"));
                return _comment.Text;
            }
            else
            {
                return "";
            }
        }

        public void DeleteItem(string itemName)
        {

            _delete = WaitForElementIsVisible(By.XPath(ITEM_DETAIL));
            _delete.Click();
            WaitForLoad();
            _deleteItem = WaitForElementIsVisible(By.XPath(DELETE_ITEM_2));

            _deleteItem.Click();
            WaitForLoad();
        }

        public OutputFormExpiry ShowExpiryDate()
        {
            if (isElementVisible(By.XPath(EXPIRY_DATE_1)))
            {
                //modifier PhysQty change l'emplacement
                _expiryDate = WaitForElementIsVisible(By.XPath(EXPIRY_DATE_1));
                _expiryDate.Click();
            }
            else
            {
                _expiryDate = WaitForElementIsVisible(By.XPath(EXPIRY_DATE));
                _expiryDate.Click();
            }

            return new OutputFormExpiry(_webDriver, _testContext);
        }
        public OutputFormExpiry ClickExpiryDate()
        {

            if (isElementVisible(By.XPath("//span[contains(@class,\"glyphicon glyphicon-calendar\")]")))
            {
                var expiryDates = _webDriver.FindElements(By.XPath("//span[contains(@class,\"glyphicon glyphicon-calendar\")]"));
                expiryDates.FirstOrDefault().Click();
            }
            else
            {
                var expiryDates = _webDriver.FindElements(By.XPath("//div/a[contains(@class,'btn special-design float-right expDateLink')]"));
                expiryDates.FirstOrDefault().Click();
            }
            return new OutputFormExpiry(_webDriver, _testContext);
        }
        public string GetPrice(string item)
        {
            _itemPrice = WaitForElementIsVisible(By.XPath(String.Format("//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[2]/div[3]/span[text()= '{0}']/../../..//div[contains(@class,'edit-zone')]/div/span[@name='ItemTotalPrice']", item)));
            return _itemPrice.Text;
        }

        public void SetPhysPackagingQty(string v0, string v1 = null, string v2 = null)
        {
            var PhysPackagingQty = WaitForElementIsVisible(By.XPath("//*/a[contains(@class,'btn-open-packagings-qtys')]"));
            PhysPackagingQty.Click();
            if (v0 != null)
            {
                var qty0 = WaitForElementIsVisible(By.XPath("//*/input[@name='PackagingsQtys[0].PhysicalQuantity']"));
                qty0.SetValue(PageBase.ControlType.TextBox, v0);
            }
            if (v1 != null)
            {
                var qty1 = WaitForElementIsVisible(By.XPath("//*/input[@name='PackagingsQtys[1].PhysicalQuantity']"));
                qty1.SetValue(PageBase.ControlType.TextBox, v1);
            }
            if (v2 != null)
            {
                var qty2 = WaitForElementIsVisible(By.XPath("//*/input[@name='PackagingsQtys[2].PhysicalQuantity']"));
                qty2.SetValue(PageBase.ControlType.TextBox, v2);
            }

            var save = WaitForElementIsVisible(By.XPath("//*/button[text()='Save quantities']"));
            save.Click();
        }

        public new int CheckTotalNumber()
        {
            var lignes = _webDriver.FindElements(By.XPath("//*/form[contains(@class,'rowOutputFormDetailForm')]"));
            return lignes.Count;
        }
        public bool VerifyChangePage(int nbre)
        {
            var list1 = new List<string>();
            var list2 = new List<string>();

            for (var i = 1; i < nbre + 1; i++)
            {
                list1 = GetItemNames();
                WaitForLoad();
                var navigatorItem = WaitForElementIsVisibleNew(By.XPath(string.Format("//*[@id=\"divTable\"]/nav/ul/li[*]/a[@data-pager-pageindex = {0}]", i)));
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
        public bool isPriceVisible()
        {
            if (isElementVisible(By.XPath("//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[2]/div[3]/span/../../..//div[contains(@class,'edit-zone')]/div/span[@name='ItemMedianPrice']")))
            {
                return true;
            }
            return false;
        }
        public bool isPackagingQtyZero()
        {
            if (isElementVisible(By.XPath("//*[@id=\"item_UniquePackagingPhysicalQty\"]")))
            {
                var packagingQty = WaitForElementIsVisible(By.XPath("//*[@id=\"item_UniquePackagingPhysicalQty\"]"));
                if (packagingQty.GetAttribute("value") == "0")
                {
                    return true;
                }
            }
            return false;
        }

        public double GetPhysicalValue()
        {

            var ofPhyValueLabel = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/form/div/div[12]/span"));
            var ofPhyValueWithoutSymbol = ofPhyValueLabel.Text.Replace("€", "").Trim();
            CultureInfo culture = new CultureInfo("fr-FR");

            double.TryParse(ofPhyValueWithoutSymbol, NumberStyles.Any, culture, out double result);
            return result;
        }
        public bool VerifyPackagingUnits(List<string> packagings)
        {
            List<string> result = new List<string>();
            var listPackagingUnits = _webDriver.FindElements(By.XPath("//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[5]/span"));
            for (var i = 0; i < listPackagingUnits.Count; i++)
            {
                foreach (var packagingUnit in packagings)
                {
                    if (!listPackagingUnits[i].Text.Contains(packagingUnit))
                    {
                        return false;
                    }
                }
            }
            return true;
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
            if (isElementVisible(By.XPath("//*[starts-with(@id,'modal-')]/div/p"))) return item_allergens;
            _allergensList = WaitForElementIsVisible(By.XPath(ALLERGENS_LIST));
            if (_allergensList == null) return item_allergens;
            ReadOnlyCollection<IWebElement> allergensInList = _allergensList.FindElements(By.TagName("li"));
            if (allergensInList.Count > 0)
                foreach (var allergen in allergensInList)
                    try
                    {
                        IWebElement image = allergen.FindElement(By.TagName("img"));
                        if (image != null) item_allergens.Add(allergen.Text.Trim());
                    }
                    catch (NoSuchElementException)
                    {
                        continue;
                    }
            return item_allergens;
        }

        public bool ElementIsVisible()
        {
            WaitForLoad();
            if (isElementVisible(By.XPath(ELEMENT_ITEM)))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        public void OpenNewTab()
        {
            var firstTabHandle = _webDriver.CurrentWindowHandle;

            var actualUrl = _webDriver.Url;
            ((IJavaScriptExecutor)_webDriver).ExecuteScript("window.open();");

            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);

            _webDriver.Navigate().GoToUrl(actualUrl);
        }

        public string GetOutputFormNumber()
        {
            var number = WaitForElementExists(By.XPath(OUTPUT_FORM_NUMBER));
            return number.Text.Split(' ')[3];
        }
    }

}