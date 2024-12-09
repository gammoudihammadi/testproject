using DocumentFormat.OpenXml.Presentation;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.Inventory
{
    public class InventoryItem : PageBase
    {

        public InventoryItem(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //___________________________________________ Constantes ____________________________________________________

        // General
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";
        private const string VALIDATE = "btn-validate-inventory";
        private const string CONFIRM_VALIDATE = "btn-submit-validate-inventory";
        private const string EXTENDED_BUTTON = "//*[@id=\"div-body\"]/div/div[1]/div/div[1]/button";
        private const string REFRESH = "btn-inventories-refresh";
        private const string EXPORT = "btn-inventories-export";
        private const string PRINT = "btn-print";
        private const string PRINT_PREP_SHEET = "//*[@id=\"btn-print-preparation-sheet\"]";
        private const string CONFIRM_PRINT_PREP_SHEET = "//*[@id=\"btn-print-prep-sheet\"]";
        private const string EXPORT_SAGE = "btn-export-for-sage";
        private const string CONFIRM_EXPORT_SAGE = "btn-popup-validate";
        private const string CONFIRM_EXPORT_SAGE_PATCH = "btn-validate-export";
        private const string CANCEL_SAGE = "btn-cancel-popup";
        private const string DOWNLOAD_FILE = "btn-download-file";
        private const string ENABLE_EXPORT_FOR_SAGE = "btn-enable-export-for-sage";

        // Onglets
        private const string GENERAL_INFORMATION_TAB = "hrefTabContentInformations";
        private const string FOOTER_TAB = "hrefTabContentFooter";
        private const string ACCOUNTING_TAB = "hrefTabContentExportSageWriting";

        // Tableau
        private const string THEORICAL_VALUE = "InventoryTheoValue";
        private const string PHYSICAL_VALUE = "InventoryPhysValue";
        private const string DIFFERENCE_VALUE = "DiffValue";
        private const string VALUE = "//*[@id=\"InventoryTheoValue\"]";

        private const string GROUP = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/span";
        private const string ITEM = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[2]";
        private const string SAVE_ICON = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[@title='{0}']/../..//span/img";
        private const string ITEM_NAME = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[2]/div[3]/span";
        private const string PHYSICAL_PACKAGING_QUANTITY = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[2]/div[4]/p/span";
        private const string PHYSICAL_PACKAGING_QUANTITY_INPUT = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[@title='{0}']/../../div[4]/input";
        private const string PHYSICAL_QUANTITY_INPUT = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[@title='{0}']/../../div[6]/input";
        private const string PHYSICAL_QUANTITY = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div/div[6]/span";
        private const string PRICE_INPUT = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[@title='{0}']/../../div[8]/div/div[2]/input";
        private const string ITEM_PRICE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[2]/div[8]/span";
        private const string THEORICAL_QUANTITY = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div/div[9]/span";
        private const string THEORICAL_QUANTITY_INPUT = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[@title='{0}']/../../div[9]/span";
        private const string TOTAL_PHYSICAL_QTY = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[2]/div[10]";
        private const string STORAGE_UNIT = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[2]/div[7]";
        private const string QUANTITY_DIFF = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[2]/div[11]/span";
        private const string EXTENDED_MENU = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[@title='{0}']/../../div[*]/a/span[text()='...']";
        private const string EDIT_ITEM = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[@title='{0}']/../..//div[*]/ul/li[*]/a/span[@class = 'fas fas fa-pencil-alt glyph-span']";
        private const string EXTENDED_MENU_FORM = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[1]/div[15]/a/span";
        private const string EDIT_ITEM_FORM = "//*[@id=\"btn-edit-item-details-19585\"]/span";
        private const string COMMENT_ITEM = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[3]/span[@title='{0}']/../..//div[*]/ul/li[*]/span/span[@class = 'glyphicon glyphicon-comment glyph-span']";
        private const string COMMENT_ITEM_2 = "//*[@id=\"19585\"]/span";
        private const string COMMENT = "Comment";
        private const string SAVE_COMMENT = "//*[@id=\"modal-1\"]/div/div/div/div[2]/div/form/div[2]/button[2]";
        private const string EXPIRY_DATE = "//*[@id=\"btn-Inventoryitem-ExpirationDates\"]";
        private const string RECEIVED_QTY_LINE = "/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[{0}]/div/div/form/div[1]/div[6]/input";
        private const string ALLERGENS_BTN = "/html/body/div[3]/div/div[3]/div/div/div/div[2]/div[2]/div/div/div[2]/div[2]/div/div/form/div[2]/div[14]/a";
        private const string ALLERGENS_LIST = "/html/body/div[4]/div/div/div/div[2]/div/div/div/div/ul";
        //private const string LIST_ITEMS = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[2]/div[3]/span";
        private const string LIST_ITEMS = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div/div[3]/span";
        private const string FIRST_ITEM_NAME = "//form/div[1]/div[3]/span";
        private const string PHYS_QTY_INPUT = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[*]/span[@title='{0}']/../../div[6]/input";

        // Filtres
        private const string RESET_FILTER = "//*[@id=\"formSearchItems\"]/div[1]/a";
        private const string FILTER_NAME = "tbSearchPatternWithAutocomplete";
        private const string FIRST_RESULT_SEARCH = "//*[@id=\"formSearchItems\"]/div[2]/span/div/div/div/strong[text()='{0}']";
        private const string FILTER_SORTBY = "cbSortBy";
        private const string FILTER_SEIZABLE_ITEMS = "//*[@id=\"ShowAllItems\"]";
        private const string FILTER_ITEMS_WITH_QTY = "//*[@id=\"ShowItemsWithQty\"]";
        private const string FILTER_ITEMS_WITH_DIFF_QTY = "//*[@id=\"ShowItemsWithDiff\"]";
        private const string FILTER_GROUP = "ItemsIndexModelSelectedGroups_ms";
        private const string FILTER_KEYWORDS = "ItemsIndexModelSelectedGroups_ms";
        private const string SEARCH_GROUP = "/html/body/div[11]/div/div/label/input";
        private const string UNSELECT_SITE = "/html/body/div[11]/div/ul/li[2]/a";
        private const string FIRST_ITEM = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[2]";
        private const string PHYSICAL_QTY = "item_InventoryItem_ManualPhysicalQty";
        private const string COPY_QTY = "SetQtysToTheo";
        private const string UPDATE_ITEMS = "btn-submit-update-items-inventory";
        private const string SET_QTYS_ZERO = "SetQtysToZero";
        private const string EXPIRY_DATE_CHANGE_CSS = "(//*/form[*]/div[1]/div[13]/a/parent::div)[1]/a/span";
        private const string DIFF_PRICE = "item_InventoryItem_MedianPrice";
        private const string DIFF_PRICE_XPATH = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[2]/div[8]/span";
        private const string QTY_XPATH = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[2]/div[6]/span";
        private const string QTY_XPATH_EMPTY = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[2]/div[6]";

        //__________________________________________VARIABLES____________________________________________________

        // General
        [FindsBy(How = How.XPath, Using = EXPIRY_DATE_CHANGE_CSS)]
        private IWebElement _expiryDatechangecss;

        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;

        [FindsBy(How = How.Id, Using = VALIDATE)]
        private IWebElement _validate;

        [FindsBy(How = How.XPath, Using = EXTENDED_BUTTON)]
        private IWebElement _extendedButton;

        [FindsBy(How = How.Id, Using = REFRESH)]
        private IWebElement _refresh;

        [FindsBy(How = How.Id, Using = EXPORT)]
        private IWebElement _export;

        [FindsBy(How = How.Id, Using = PRINT)]
        private IWebElement _print;

        [FindsBy(How = How.XPath, Using = PRINT_PREP_SHEET)]
        private IWebElement _printPreparationSheet;

        [FindsBy(How = How.XPath, Using = CONFIRM_PRINT_PREP_SHEET)]
        private IWebElement _confirmPrintPrepSheet;

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

        [FindsBy(How = How.Id, Using = FOOTER_TAB)]
        private IWebElement _footerTab;

        [FindsBy(How = How.Id, Using = ACCOUNTING_TAB)]
        private IWebElement _accountingTab;


        // Tableau items inventory
        [FindsBy(How = How.Id, Using = THEORICAL_VALUE)]
        private IWebElement _theoricalValue;

        [FindsBy(How = How.Id, Using = PHYSICAL_VALUE)]
        private IWebElement _physicalValue;

        [FindsBy(How = How.Id, Using = DIFFERENCE_VALUE)]
        private IWebElement _differenceValue;

        [FindsBy(How = How.XPath, Using = GROUP)]
        private IWebElement _group;

        [FindsBy(How = How.XPath, Using = ITEM)]
        private IWebElement _item;

        [FindsBy(How = How.XPath, Using = ITEM_NAME)]
        private IWebElement _itemName;

        [FindsBy(How = How.XPath, Using = PHYSICAL_PACKAGING_QUANTITY_INPUT)]
        private IWebElement _physPackQuantityInput;

        [FindsBy(How = How.XPath, Using = PHYSICAL_PACKAGING_QUANTITY)]
        private IWebElement _physPackQuantity;

        [FindsBy(How = How.XPath, Using = PHYSICAL_QUANTITY_INPUT)]
        private IWebElement _physQuantityInput;

        [FindsBy(How = How.XPath, Using = PHYSICAL_QUANTITY)]
        private IWebElement _physQuantity;

        [FindsBy(How = How.XPath, Using = PRICE_INPUT)]
        private IWebElement _priceInput;

        [FindsBy(How = How.XPath, Using = ITEM_PRICE)]
        private IWebElement _price;

        [FindsBy(How = How.XPath, Using = THEORICAL_QUANTITY)]
        private IWebElement _theoricalQty;

        [FindsBy(How = How.XPath, Using = THEORICAL_QUANTITY_INPUT)]
        private IWebElement _theoricalQtyInput;

        [FindsBy(How = How.XPath, Using = TOTAL_PHYSICAL_QTY)]
        private IWebElement _totalPhysQty;

        [FindsBy(How = How.XPath, Using = EXTENDED_MENU)]
        private IWebElement _extendedMenu;

        [FindsBy(How = How.XPath, Using = EDIT_ITEM)]
        private IWebElement _editItem;

        [FindsBy(How = How.XPath, Using = EXTENDED_MENU_FORM)]
        private IWebElement _extendedMenuForm;

        [FindsBy(How = How.XPath, Using = EDIT_ITEM_FORM)]
        private IWebElement _editItemForm;

        [FindsBy(How = How.XPath, Using = COMMENT_ITEM)]
        private IWebElement _addComment;

        [FindsBy(How = How.Id, Using = COMMENT)]
        private IWebElement _comment;

        [FindsBy(How = How.XPath, Using = SAVE_COMMENT)]
        private IWebElement _saveComment;

        [FindsBy(How = How.XPath, Using = EXPIRY_DATE)]
        private IWebElement _expiryDate;

        [FindsBy(How = How.Id, Using = ALLERGENS_BTN)]
        private IWebElement _allergensBtn;

        [FindsBy(How = How.Id, Using = ALLERGENS_LIST)]
        private IWebElement _allergensList;

        //__________________________________________ Filtres ____________________________________________________

        [FindsBy(How = How.XPath, Using = RESET_FILTER)]
        private IWebElement _resetFilter;

        [FindsBy(How = How.Id, Using = FILTER_NAME)]
        private IWebElement _searchByNameFilter;

        [FindsBy(How = How.Id, Using = FILTER_SORTBY)]
        private IWebElement _sortByFilter;

        [FindsBy(How = How.XPath, Using = FILTER_SEIZABLE_ITEMS)]
        private IWebElement _showAllSeizableItems;

        [FindsBy(How = How.XPath, Using = FILTER_ITEMS_WITH_QTY)]
        private IWebElement _showItemsWithQtyOnly;

        [FindsBy(How = How.XPath, Using = FILTER_ITEMS_WITH_DIFF_QTY)]
        private IWebElement _showItemsWithDifferenceOnly;

        [FindsBy(How = How.Id, Using = FILTER_GROUP)]
        private IWebElement _groupFilter;

        [FindsBy(How = How.XPath, Using = SEARCH_GROUP)]
        private IWebElement _groupFilterSearch;

        [FindsBy(How = How.XPath, Using = UNSELECT_SITE)]
        private IWebElement _unselectAll;

        public enum FilterItemType
        {
            SearchByName,
            SortBy,
            ShowAllSeizableItems,
            ShowItemsWithQtyOnly, // thQty
            ShowItemsWithDifferenceOnly,
            ByGroup,
            Keywords
        }

        public void Filter(FilterItemType FilterItemType, object value)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)_webDriver;

            switch (FilterItemType)
            {
                case FilterItemType.SearchByName:
                    js.ExecuteScript("window.scrollTo(0,0)");
                    _searchByNameFilter = WaitForElementIsVisible(By.Id(FILTER_NAME));
                    _searchByNameFilter.SetValue(ControlType.TextBox, value);

                    if (isElementVisible(By.XPath(String.Format(FIRST_RESULT_SEARCH, value))))
                    {
                        var firstResultSearch = _webDriver.FindElement(By.XPath(String.Format(FIRST_RESULT_SEARCH, value)));
                        firstResultSearch.Click();
                    }
                    else
                    {
                        return;
                    }
                    break;
                case FilterItemType.SortBy:
                    _sortByFilter = WaitForElementIsVisible(By.Id(FILTER_SORTBY));
                    js.ExecuteScript("window.scrollTo(0,0)");
                    _sortByFilter.Click();
                    var element = WaitForElementIsVisible(By.XPath("//option[contains(@value,'" + value + "')]"));
                    _sortByFilter.SetValue(ControlType.DropDownList, element.Text);
                    _sortByFilter.Click();
                    break;
                case FilterItemType.ShowAllSeizableItems:
                    _showAllSeizableItems = WaitForElementExists(By.XPath(FILTER_SEIZABLE_ITEMS));
                    _showAllSeizableItems.SetValue(ControlType.CheckBox, value);
                    break;
                case FilterItemType.ShowItemsWithDifferenceOnly:
                    _showItemsWithDifferenceOnly = WaitForElementExists(By.XPath(FILTER_ITEMS_WITH_DIFF_QTY));
                    _showItemsWithDifferenceOnly.SetValue(ControlType.CheckBox, value);
                    break;
                case FilterItemType.ShowItemsWithQtyOnly:
                    _showItemsWithQtyOnly = WaitForElementExists(By.XPath(FILTER_ITEMS_WITH_QTY));
                    _showItemsWithQtyOnly.SetValue(ControlType.RadioButton, value);

                    break;
                case FilterItemType.ByGroup:
                    ComboBoxSelectById(new ComboBoxOptions(FILTER_GROUP, (string)value));
                    break;
                case FilterItemType.Keywords:
                    ComboBoxSelectById(new ComboBoxOptions("ItemsIndexModelSelectedKeywords_ms", (string)value));
                    break;
            }

            WaitPageLoading();
            WaitForLoad();
        }

        public void ResetFilter()
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)_webDriver;
            js.ExecuteScript("window.scrollTo(0,0)");
            _resetFilter = WaitForElementIsVisible(By.XPath(RESET_FILTER));
            _resetFilter.Click();
            WaitForLoad();

            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                // pas de date
            }
        }

        public int GetNumberItems()
        {
            PageSize("100");
            System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> listItems = _webDriver.FindElements(By.XPath(LIST_ITEMS));
            PageSize("8");
            return listItems.Count;
        }

        public bool VerifyGroupFilter(string item_Group)
        {
            PageSize("100");

            var listItems = _webDriver.FindElements(By.XPath("//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/span"));
            foreach (var item in listItems)
            {
                if (!item.Text.Contains(item_Group))
                {
                    return false;
                }
            }
            return true;
        }

        public bool IsEdited(string itemName)
        {
            if (isElementVisible(By.XPath(string.Format(SAVE_ICON, itemName))))
            {
                WaitForElementIsVisible(By.XPath(string.Format(SAVE_ICON, itemName)));
            }
            else
            {
                return false;
            }

            return true;
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

        public Boolean VerifyGroup(string group)
        {
            var boolValue = true;
            var elements = _webDriver.FindElements(By.XPath(ITEM_NAME));

            if (elements.Count == 0)
                return false;

            int nbMax = elements.Count <= 8 ? elements.Count : 8;
            int compteur = 1;
            string ancienItemName = "";
            while (compteur <= nbMax)
            {
                foreach (var elm in elements)
                {
                    // deux fois "ALGA NORI HOJAS 25 GR"
                    string itemName = elm.GetAttribute("title");
                    if (itemName == ancienItemName) continue;
                    ancienItemName = itemName;

                    // clic sur la ligne
                    elm.Click();

                    try
                    {
                        var itemGeneralInformationPage = EditItem(elm.GetAttribute("title"));
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

        public bool IsSortedByGroup()
        {
            var elements = _webDriver.FindElements(By.XPath(GROUP));

            if (elements.Count == 0)
                return false;

            var ancientGroup = elements[0].Text;

            foreach (var elm in elements)
            {
                if (string.Compare(ancientGroup, elm.Text) > 0)
                    return false;

                ancientGroup = elm.Text;
            }

            return true;
        }


        public bool IsWithQtyDifferenceOnly()
        {
            WaitPageLoading();
            WaitForLoad();
            var elements = _webDriver.FindElements(By.XPath(QUANTITY_DIFF));

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

        public bool IsWithTheoOrPhysQtyOnly()
        {
            var elements = _webDriver.FindElements(By.XPath(ITEM_NAME));

            foreach (var elm in elements)
            {
                elm.Click();

                try
                {
                    _physQuantityInput = WaitForElementIsVisible(By.XPath(String.Format(PHYSICAL_QUANTITY_INPUT, elm.GetAttribute("title"))));
                    var physicalValue = _physQuantityInput.GetAttribute("value");

                    _theoricalQtyInput = WaitForElementIsVisible(By.XPath(String.Format(THEORICAL_QUANTITY_INPUT, elm.GetAttribute("title"))));
                    var theoricalValue = _theoricalQtyInput.Text;

                    if (physicalValue.Equals("0") && theoricalValue.Equals("0"))
                        return false;
                }
                catch
                {
                    return false;
                }
            }

            return true;
        }

        // _______________________________________________ Méthodes _________________________________________________

        // General

        public InventoriesPage BackToList()
        {
            _backToList = WaitForElementIsVisibleNew(By.XPath(BACK_TO_LIST));
            _backToList.Click();
            WaitPageLoading();

            return new InventoriesPage(_webDriver, _testContext);
        }

        public InventoryValidationModalPage Validate()
        {
            ShowValidationMenu();

            _validate = WaitForElementIsVisible(By.Id(VALIDATE));
            _validate.Click();
            WaitForLoad();
            WaitPageLoading();

            return new InventoryValidationModalPage(_webDriver, _testContext);
        }
        public void ConfirmValidation()
        {
            var validateBtn = WaitForElementIsVisible(By.Id("btn-submit-validate-inventory"));
            validateBtn.Click();
            WaitForLoad();
        }

        public override void ShowExtendedMenu()
        {
            var actions = new Actions(_webDriver);
            _extendedButton = WaitForElementIsVisible(By.XPath(EXTENDED_BUTTON));
            actions.MoveToElement(_extendedButton).Perform();
            WaitForLoad();

            WaitPageLoading();
        }

        public void Refresh()
        {
            ShowExtendedMenu();
            WaitPageLoading();
            WaitForLoad();
            _refresh = WaitForElementIsVisible(By.Id(REFRESH));
            _refresh.Click();
            WaitForLoad();
        }

        public void ExportExcelFile(bool printValue, string downloadPath)
        {
            // secouse du bouton menu
            WaitForLoad();
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
            string mois = "(?:0[1-9]|1[0-2])";         // mois
            string space = "(\\s)";                    // Espace
            string annee = "\\d{4}";                   // annee YYYY
            string jour = "[0-3]\\d";                  // jour
            string heure = "(?:0[0-9]|1[0-9]|2[0-3])"; // heure
            string minutes = "[0-5]\\d";               // minutes
            string secondes = "[0-5]\\d";              // secondes
            string nombre = "(\\d{3}|\\d{4}|\\d{5})";


            Regex r = new Regex("^inventory_" + nombre + space + annee + "-" + mois + "-" + jour + space + heure + "-" + minutes + "-" + secondes + ".xlsx$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = r.Match(filePath);

            return m.Success;
        }

        public PrintReportPage PrintInventoryItems(bool printValue)
        {
            LoadingPage();
            ShowExtendedMenu();
            WaitPageLoading();
            WaitForLoad();
            _print = WaitForElementIsVisible(By.Id(PRINT));
            _print.Click();
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

        public PrintPreparationSheetModalPage ClickPrintPreparationSheet()
        {
            if (IsDev())
            {
                _printPreparationSheet = WaitForElementIsVisible(By.XPath(PRINT_PREP_SHEET));
                _printPreparationSheet.Click();
            }
            else
            {
                _printPreparationSheet = WaitForElementIsVisible(By.XPath("//*[@id=\"div-body\"]/div/div[1]/div/div[1]/div/a[5]"));
                _printPreparationSheet.Click();
            }
            WaitForLoad();

            return new PrintPreparationSheetModalPage(_webDriver, _testContext);
        }

        public PrintReportPage PrintPreparationSheet(bool printValue)
        {
            ShowExtendedMenu();

            var printPreparationSheetModalPage = ClickPrintPreparationSheet();

            // Set parameters of perparation sheet print 
            printPreparationSheetModalPage.SetParameters(PrintPreparationSheetModalPage.PrintParameters.PrintMode, "PDF");
            printPreparationSheetModalPage.SetParameters(PrintPreparationSheetModalPage.PrintParameters.DisplayMethod, "Details with an existing stock");
            printPreparationSheetModalPage.SetParameters(PrintPreparationSheetModalPage.PrintParameters.ShowAveragePrice, true);
            printPreparationSheetModalPage.SetParameters(PrintPreparationSheetModalPage.PrintParameters.ShowMainSupplier, true);
            printPreparationSheetModalPage.SetParameters(PrintPreparationSheetModalPage.PrintParameters.ShowTheoricalQuantities, false);

            if (IsDev())
            {
                _confirmPrintPrepSheet = WaitForElementIsVisible(By.XPath(CONFIRM_PRINT_PREP_SHEET));
                _confirmPrintPrepSheet.Click();
            }
            else
            {
                _confirmPrintPrepSheet = WaitForElementIsVisible(By.XPath("//*[@id=\"item-print-form\"]/div[2]/button[2]"));
                _confirmPrintPrepSheet.Click();
            }
            WaitForLoad();
            WaitPageLoading();
            if (printValue)
            {
                WaitPageLoading();
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-pdf']"));
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

                    if (isElementVisible(By.Id(CANCEL_SAGE)))
                    {
                        // On ferme la pop-up
                        _cancelSage = WaitForElementIsVisible(By.Id(CANCEL_SAGE));
                        _cancelSage.Click();
                        WaitForLoad();
                    }
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

        public void ExportForSage(bool printValue)
        {
            ShowExtendedMenu();
            _exportSage = WaitForElementIsVisible(By.Id(EXPORT_SAGE));
            _exportSage.Click();
            WaitForLoad();
            if (IsDev()) _confirmExportSage = WaitForElementIsVisible(By.Id(CONFIRM_EXPORT_SAGE));
            else _confirmExportSage = WaitForElementIsVisible(By.Id(CONFIRM_EXPORT_SAGE_PATCH));
            _confirmExportSage.Click();
            WaitForLoad();
            if (!printValue)
            {
                if (isElementVisible(By.Id(CONFIRM_EXPORT_SAGE)))
                {
                    _downloadFile = _webDriver.FindElement(By.Id(DOWNLOAD_FILE));
                    _downloadFile.Click();
                    WaitForLoad();
                    if (isElementVisible(By.Id(CANCEL_SAGE)))
                    {
                        _cancelSage = WaitForElementIsVisible(By.Id(CANCEL_SAGE));
                        _cancelSage.Click();
                        WaitForLoad();
                    }
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
            WaitForLoad();
            WaitForDownload();
            //Close();
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

        public bool IsSAGEFileCorrect(string filePath)
        {
            // "inventory 2019-12-09 13-20-29.txt"

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

        public void EnableExportForSage()
        {
            ShowExtendedMenu();

            _enableExportForSage = WaitForElementIsVisible(By.Id(ENABLE_EXPORT_FOR_SAGE));
            _enableExportForSage.Click();
            WaitForLoad();
        }

        // Onglets
        public InventoryGeneralInformation ClickOnGeneralInformationTab()
        {
            _generalInformationTab = WaitForElementIsVisible(By.Id(GENERAL_INFORMATION_TAB));
            _generalInformationTab.Click();
            WaitPageLoading();

            return new InventoryGeneralInformation(_webDriver, _testContext);
        }

        public InventoryFooterPage ClickOnFooterTab()
        {
            _footerTab = WaitForElementIsVisible(By.Id(FOOTER_TAB));
            _footerTab.Click();
            WaitPageLoading();

            return new InventoryFooterPage(_webDriver, _testContext);
        }

        public InventoryAccountingPage ClickOnAccountingTab()
        {
            _accountingTab = WaitForElementIsVisible(By.Id(ACCOUNTING_TAB));
            _accountingTab.Click();
            WaitPageLoading();

            return new InventoryAccountingPage(_webDriver, _testContext);
        }

        // Tableau items inventory
        public void SelectFirstItem()
        {
            WaitPageLoading();
            WaitForLoad();

            var firstItem = WaitForElementIsVisibleNew(By.XPath(FIRST_ITEM));
            WaitForLoad();
            if (!firstItem.GetAttribute("class").Contains("editable-row selected"))
            {
                IJavaScriptExecutor js = (IJavaScriptExecutor)_webDriver;
                js.ExecuteScript("window.scrollTo(0, 0)");
                WaitForLoad();

                var displayItem = WaitForElementIsVisibleNew(By.ClassName("inventory-item-row-display"));

                displayItem.Click();
                WaitForLoad();
            }
        }

        public void SelectItem(string itemName)
        {
            var ligne = WaitForElementExists(By.XPath("//*/div[contains(@class,'display-zone')]//span[text()=\"" + itemName + "\"]/../.."));
            ligne.Click();
        }

        public double GetTheoricalValue(string currency, string decimalSeparator)
        {
            // Récupération du type de séparateur (, ou . selon les pays)
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            _theoricalValue = WaitForElementExists(By.Id(THEORICAL_VALUE));
            int counter = 10;
            while (_theoricalValue.Text == "-" && counter > 0)
            {
                WaitPageLoading();
                _theoricalValue = WaitForElementExists(By.Id(THEORICAL_VALUE));
                counter--;
            }
            string thValue = _theoricalValue.Text;
            string result = thValue.Replace(currency, "").Replace(" ", "");

            return double.Parse(result, ci);
        }

        public double GetPhysicalValue(string currency, string decimalSeparator)
        {
            //Récupération du type de séparateur(, ou.selon les pays)
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            _physicalValue = WaitForElementExists(By.Id(PHYSICAL_VALUE));

            string result = _physicalValue.Text.Replace(currency, "").Replace(" ", "");

            return double.Parse(result, ci);

        }

        public double GetDifferenceValue(string currency, string decimalSeparator)
        {
            //Récupération du type de séparateur(, ou.selon les pays)
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            _differenceValue = WaitForElementExists(By.Id(DIFFERENCE_VALUE));
            string result = _differenceValue.Text.Replace(currency, "").Replace(" ", "");

            return double.Parse(result, ci);

        }

        public double GetValueEdit(string currency, string decimalSeparator)
        {
            WaitPageLoading();
            WaitForLoad();
            // Récupération du type de séparateur (, ou . selon les pays)
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            var _value = SolveVisible(VALUE);
            WaitForLoad();
            string result = _value.Text.Replace(currency, "").Trim();
            result = result.Replace(" ", "");

            return double.Parse(result, ci);
        }

        public string GetFirstItemName()
        {
            WaitPageLoading();
            WaitForLoad();
            _itemName = WaitForElementExists(By.XPath(ITEM_NAME));
            if (string.IsNullOrEmpty(_itemName.Text)) // line in mode edit
            {
                _itemName = WaitForElementExists(By.XPath(FIRST_ITEM_NAME));
            }
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

        public string GetPhysicalPackagingQuantity()
        {
            if (isElementVisible(By.XPath(PHYSICAL_PACKAGING_QUANTITY)))
            {
                _physPackQuantity = _webDriver.FindElement(By.XPath(PHYSICAL_PACKAGING_QUANTITY));
                return _physPackQuantity.Text;
            }
            else
            {
                return "0";
            }
        }

        public void AddPhysicalPackagingQuantity(string itemName, string phys_PackQuantity)
        {
            WaitPageLoading();
            WaitForLoad();
            _physPackQuantityInput = WaitForElementIsVisibleNew(By.XPath(string.Format(PHYSICAL_PACKAGING_QUANTITY_INPUT, itemName)));
            _physPackQuantityInput.SetValue(ControlType.TextBox, phys_PackQuantity);
            WaitForLoad();

            WaitForElementIsVisibleNew(By.XPath(String.Format(SAVE_ICON, itemName)));
        }

        public string GetPhysicalQuantity()
        {
            WaitForLoad();
            _physQuantity = WaitForElementExists(By.XPath(PHYSICAL_QUANTITY));
            return _physQuantity.Text.Trim();
        }

        public List<string> GetAllPhysicalQuantity()
        {
            WaitForLoad();
            var qtys = new List<string>();
            var physQuantities = _webDriver.FindElements(By.XPath(PHYSICAL_QUANTITY));
            foreach (var physQuantity in physQuantities)
            {
                qtys.Add(physQuantity.Text);
            }
            return qtys;
        }

        public void AddPhysicalQuantity(string itemName, string value)
        {
            WaitLoading();
            _physQuantityInput = _webDriver.FindElement(By.XPath(String.Format(PHYS_QTY_INPUT, itemName)));
            _physQuantityInput.SetValue(ControlType.TextBox, value);
            WaitPageLoading();
        }

        public void AddPhysicalQtyToInventory(string value)
        {
            WaitForLoad();
            var elements = _webDriver.FindElements(By.XPath(ITEM_NAME));

            int nbMax = elements.Count <= 5 ? elements.Count : 5;
            int compteur = 1;

            foreach (var elm in elements)
            {
                if (compteur <= nbMax)
                {
                    elm.Click();

                    if (!isElementVisible(By.XPath(string.Format(PHYSICAL_PACKAGING_QUANTITY_INPUT, elm.GetAttribute("title")))))
                    {
                        // cadie 2 packaging même site différent supplier
                        continue;
                    }

                    var quantity = WaitForElementIsVisible(By.XPath(string.Format(PHYSICAL_PACKAGING_QUANTITY_INPUT, elm.GetAttribute("title"))));
                    quantity.SetValue(ControlType.TextBox, value);

                    WaitForElementIsVisible(By.XPath(string.Format(SAVE_ICON, elm.GetAttribute("title"))));

                    compteur++;
                }
            }
            WaitForLoad();
        }

        public void AddPhysicalQuantityByThQty(string value)
        {
            // select item with th qty
            var thQtyElements = _webDriver.FindElements(By.XPath(ITEM));
            foreach (var elm in thQtyElements)
            {
                elm.Click();
                WaitForLoad();
                _physQuantityInput = SolveVisible("//*[@id='item_InventoryItem_ManualPhysicalQty']");
                _physQuantityInput.SetValue(ControlType.TextBox, value);
                WaitForLoad();
            }
        }

        public double GetPrice(string currency, string decimalSeparatorValue)
        {
            CultureInfo ci = decimalSeparatorValue.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            _price = SolveVisible(ITEM_PRICE);
            return double.Parse(_price.Text.Replace(currency, ""), ci);
        }

        public double GetPriceEdit(string currency, string decimalSeparatorValue)
        {
            CultureInfo ci = decimalSeparatorValue.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");
            _price = SolveVisible("//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[8]/div/div[2]/input");
            return double.Parse(_price.GetAttribute("value").Replace(currency, ""), ci);
        }


        public void SetPrice(string itemName, string value)
        {
            WaitForLoad();
            _priceInput = WaitForElementIsVisible(By.XPath(string.Format(PRICE_INPUT, itemName)));
            _priceInput.SetValue(ControlType.TextBox, value);
            WaitPageLoading();
            WaitForLoad();
        }

        public string GetTheoricalQuantity()
        {
            _theoricalQty = WaitForElementExists(By.XPath(THEORICAL_QUANTITY));
            return _theoricalQty.GetAttribute("innerText").Trim();
        }

        public string GetTheoricalQuantityEdit()
        {
            _theoricalQty = SolveVisible("//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[1]/div[9]/span");
            return _theoricalQty.Text;
        }

        public string GetTotalPhysQty()
        {
            if (isElementVisible(By.XPath(TOTAL_PHYSICAL_QTY)))
            {
                _totalPhysQty = _webDriver.FindElement(By.XPath(TOTAL_PHYSICAL_QTY));
                return _totalPhysQty.Text;
            }
            else
            {
                return "0";
            }
        }

        public string GetStorageUnit()
        {
            var _storageUnit = _webDriver.FindElement(By.XPath(STORAGE_UNIT));
            return _storageUnit.Text;
        }

        public void AddComment(string itemName, string comment)
        {
            _extendedMenu = WaitForElementIsVisible(By.XPath(string.Format(EXTENDED_MENU, itemName)));
            _extendedMenu.Click();

            _addComment = WaitForElementIsVisible(By.XPath(string.Format(COMMENT_ITEM_2, itemName)));
            _addComment.Click();
            WaitForLoad();

            _comment = WaitForElementIsVisible(By.Id(COMMENT));
            _comment.SetValue(ControlType.TextBox, comment);
            _saveComment = WaitForElementToBeClickable(By.XPath("//*[@id=\"modal-1\"]/div/div[2]/div/form/div[2]/button[2]"));
            _saveComment.Click();

            WaitForLoad();
        }

        public string GetComment(string itemName)
        {
            _extendedMenu = WaitForElementIsVisible(By.XPath(string.Format(EXTENDED_MENU, itemName)));
            _extendedMenu.Click();


            _addComment = _webDriver.FindElement(By.XPath(string.Format("//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[2]/div[3]/span[@title='{0}']/../../..//div[*]/ul/li[*]/span/span", itemName)));

            if (!_addComment.Displayed)
            {
                return "";
            }
            _addComment.Click();
            WaitForLoad();

            try
            {
                _comment = WaitForElementExists(By.Id(COMMENT));
                string comment = _comment.Text;

                var cancel = WaitForElementExists(By.XPath("//*[@id=\"modal-1\"]/div/div[2]/div/form/div[2]/button[1]"));
                cancel.Click();

                return comment;
            }
            catch
            {
                return "";
            }
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
            _extendedMenu = WaitForElementIsVisible(By.XPath(string.Format(EXTENDED_MENU, itemName)));
            _extendedMenu.Click();

            if (isElementVisible(By.XPath(string.Format(EDIT_ITEM, itemName))))
            {
                _editItem = WaitForElementIsVisible(By.XPath(string.Format(EDIT_ITEM, itemName)));
            }
            else
            {
                _editItem = WaitForElementIsVisible(By.XPath(string.Format("/html/body/div[3]/div/div[3]/div/div/div/div[2]/div[2]/div/div/div[2]/div[2]/div/div/form/div[1]/div[15]/ul/li[1]/a/span", itemName)));
            }
            _editItem.Click();
            WaitForLoad();

            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            WaitForLoad();

            return new ItemGeneralInformationPage(_webDriver, _testContext);
        }

        public InventoryExpiry ShowFirstExpiryDate()
        {
            //modifier PhysQty change l'emplacement

            Thread.Sleep(2000);
            if (IsDev())
            {
                _expiryDate = WaitForElementIsVisible(By.XPath(EXPIRY_DATE));
                _expiryDate.Click();
            }
            else
            {
                _expiryDate = WaitForElementIsVisible(By.XPath("(//*/form[*]/div[1]/div[13]/a/parent::div)[1]"));
                _expiryDate.Click();
            }


            WaitForLoad();

            return new InventoryExpiry(_webDriver, _testContext);
        }


        public string GetExpiryDateCssAfterSaves()
        {
            _expiryDatechangecss = WaitForElementIsVisible(By.XPath(EXPIRY_DATE_CHANGE_CSS));
            return _expiryDatechangecss.GetAttribute("class");
        }

        public string GetExpiryDateQuantity()
        {
            WaitForLoad();
            string expiryTotalQuantity = WaitForElementIsVisible(By.XPath("//*[@id=\"formSaveDates\"]/div[2]/div[2]/div[2]/input")).GetAttribute("value");
            return expiryTotalQuantity;
        }

        public string GetQtyDiff()
        {
            var quantityDiff = WaitForElementIsVisible(By.XPath(QUANTITY_DIFF));
            return quantityDiff.Text;
        }

        public void CheckExport(FileInfo trouveXLSX, string site, string inventoryNumber, string decimalSeparator)
        {
            int resultNumber = OpenXmlExcel.GetExportResultNumber("Inventory", trouveXLSX.FullName);

            OpenXmlExcel.ReadAllDataInColumn("Site", "Inventory", trouveXLSX.FullName, site);

            OpenXmlExcel.ReadAllDataInColumn("Id", "Inventory", trouveXLSX.FullName, inventoryNumber);

            List<string> itemsXLSX = OpenXmlExcel.GetValuesInList("Item", "Inventory", trouveXLSX.FullName);
            var items = _webDriver.FindElements(By.XPath("//*/form[*]/div/div[3]/span"));
            Assert.AreEqual(resultNumber, items.Count);
            foreach (var item in items)
            {
                if (item.Text == "") continue;
                Assert.IsTrue(itemsXLSX.Contains(item.Text), "Item " + item.Text + " non présent dans le XLSX");
            }

            List<string> dateXLSX = OpenXmlExcel.GetValuesInList("Phys. qty", "Inventory", trouveXLSX.FullName);
            var physQtys = _webDriver.FindElements(By.XPath("//*/form[*]/div/div[6]/span"));
            Assert.AreEqual(resultNumber, physQtys.Count);
            foreach (var physQty in physQtys)
            {
                if (physQty.Text == "") continue;
                Assert.IsTrue(dateXLSX.Contains(physQty.Text), "Phys qty " + physQty.Text + " non présent dans le XLSX");
            }

            List<string> storageUnitsXLSX = OpenXmlExcel.GetValuesInList("Storage unit", "Inventory", trouveXLSX.FullName);
            var storageUnits = _webDriver.FindElements(By.XPath("//*/form[*]/div/div[7]/span"));
            Assert.AreEqual(resultNumber, storageUnits.Count);
            foreach (var unit in storageUnits)
            {
                if (unit.Text == "") continue;
                Assert.IsTrue(storageUnitsXLSX.Contains(unit.Text), "Storage Unit " + unit.Text + " non présent dans le XLSX");
            }

            List<string> pricesXLSX = OpenXmlExcel.GetValuesInList("Average price", "Inventory", trouveXLSX.FullName);
            var prices = _webDriver.FindElements(By.XPath("//*/form[*]/div/div[8]/span"));
            Assert.AreEqual(resultNumber, prices.Count);
            foreach (var price in prices)
            {
                if (price.Text == "") continue;
                Assert.IsTrue(pricesXLSX.Contains(convertAmount(price.Text, decimalSeparator)), "Price " + convertAmount(price.Text, decimalSeparator) + " non présent dans le XLSX");
            }

            List<string> thQtysXLSX = OpenXmlExcel.GetValuesInList("Theo. qty", "Inventory", trouveXLSX.FullName);
            var thQtys = _webDriver.FindElements(By.XPath("//*/form[*]/div/div[9]/span"));
            Assert.AreEqual(resultNumber, thQtys.Count);
            foreach (var thQty in thQtys)
            {
                if (thQty.Text == "") continue;
                Assert.IsTrue(thQtysXLSX.Contains(thQty.Text.Replace(" ", string.Empty)), "Theo. qty " + thQty.Text + " non présent dans le XLSX");
            }

            List<string> physValuesXLSX = OpenXmlExcel.GetValuesInList("Phys. value", "Inventory", trouveXLSX.FullName);
            var values = _webDriver.FindElements(By.XPath("//*/form[*]/div/div[12]/span"));
            Assert.AreEqual(resultNumber, values.Count);
            foreach (var value in values)
            {
                if (value.Text == "") continue;
                Assert.IsTrue(physValuesXLSX.Contains(convertAmount(value.Text, decimalSeparator)), "Value " + convertAmount(value.Text, decimalSeparator) + " non présent dans le XLSX");
            }



        }

        public string convertAmount(string amount, string decimalSeparator)
        {
            // entrée € 665,5000
            // sortie 665,5
            string sansEuro = amount.Substring(2).Replace(" ", "");
            double doublePrecision = Convert.ToDouble(sansEuro, new NumberFormatInfo() { NumberDecimalSeparator = decimalSeparator });
            return doublePrecision.ToString().Replace(".", decimalSeparator);
        }

        public void ResetQty()
        {
            var resetQty = WaitForElementIsVisible(By.Id("btn-inventory-raz"));
            resetQty.Click();
            WaitPageLoading();
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

        public void CopyTheoricalQty(bool value)
        {
            WaitPageLoading();
            WaitForLoad();
            var checkbox = WaitForElementExists(By.Id(COPY_QTY));
            checkbox.SetValue(ControlType.CheckBox, value);
            WaitForLoad();
            var update = WaitForElementExists(By.Id(UPDATE_ITEMS));
            update.Click();
            WaitForLoad();
            WaitPageLoading();
            var confirm = WaitForElementIsVisible(By.Id(CONFIRM_VALIDATE));
            confirm.Click();
            WaitForLoad();
            WaitPageLoading();
        }

        public string GetInventoryNumber()
        {
            var _inventoryNumber = WaitForElementExists(By.XPath("//*[@id=\"div-body\"]/div/div[1]/h1"));
            return Regex.Match(_inventoryNumber.Text, @"\d+").Value;
        }

        // <item,ThValue>
        public Dictionary<string, string> GetAllTheoricalValues()
        {
            Dictionary<string, string> itemsTh = new Dictionary<string, string>();
            var items = _webDriver.FindElements(By.XPath("//*/div[contains(@class,'inventory-item-row-display')]/div[3]/span"));
            foreach (var item in items)
            {
                var th = _webDriver.FindElement(By.XPath("//*/div[contains(@class,'inventory-item-row-display')]/div[3]/span[contains(text(),\"" + item.Text + "\")]/../../div[9]/span"));
                itemsTh.Add(item.Text, th.Text);
            }
            return itemsTh;
        }
        public bool VerifyDetailChangeLine(string phyQty)
        {
            var i = 2;
            var receivedQtyLine = WaitForElementIsVisible(By.XPath(string.Format("/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/form/div[2]/div[6]")));
            receivedQtyLine.Click();

            var receivedQtyInput = WaitForElementIsVisible(By.XPath(string.Format(RECEIVED_QTY_LINE, 2)));
            receivedQtyInput.SetValue(ControlType.TextBox, phyQty);

            receivedQtyInput.SendKeys(OpenQA.Selenium.Keys.Enter);
            for (i = i + 1; i > 2; i++)
            {
                if (isElementVisible(By.XPath(string.Format(RECEIVED_QTY_LINE, i))))
                {
                    receivedQtyInput = WaitForElementIsVisible(By.XPath(string.Format(RECEIVED_QTY_LINE, i)));
                    if (receivedQtyInput.GetAttribute("value") != phyQty)
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
        public void EditInventoryDate(DateTime date)
        {
            var dateInput = WaitForElementExists(By.Id("datepicker-new-inventory"));
            dateInput.SetValue(ControlType.DateTime, date);
            dateInput.SendKeys(OpenQA.Selenium.Keys.Tab);
            WaitLoading();
        }
        public string GetInventoryDate()
        {
            var date = WaitForElementExists(By.Id("datepicker-new-inventory"));
            return date.GetAttribute("value");
        }
        public double GetDifferencePercentage(string decimalSeparator)
        {
            // Récupération du type de séparateur (, ou . selon les pays)
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            _differenceValue = WaitForElementExists(By.Id("TotalDeviationWhitoutAbsolute"));
            string result = _differenceValue.Text.Replace(" ", "").Replace("%", "");

            return double.Parse(result, ci);
        }
        public double ParseDiffPercentage(double diffPercentage)
        {
            diffPercentage = Math.Round(diffPercentage, 2);
            return diffPercentage;
        }

        public double GetInvenoryValue()
        {
            WaitForLoad();
            var invPhyValueLabel = WaitForElementIsVisible(By.Id("InventoryPhysValue"));
            var invPhyValueWithoutSymbol = invPhyValueLabel.Text.Replace("€", "").Trim();
            CultureInfo culture = new CultureInfo("fr-FR");

            double.TryParse(invPhyValueWithoutSymbol, NumberStyles.Any, culture, out double result);
            return result;
        }

        public void SetQtysToZero(bool value)
        {
            var checkbox = WaitForElementExists(By.Id(SET_QTYS_ZERO));
            checkbox.SetValue(ControlType.CheckBox, value);
            WaitForLoad();
            var update = WaitForElementExists(By.Id(UPDATE_ITEMS));
            update.Click();
            WaitForLoad();
            WaitPageLoading();
            var confirm = WaitForElementIsVisible(By.Id(CONFIRM_VALIDATE));
            confirm.Click();
            WaitForLoad();
            WaitPageLoading();
        }

        public int CountInvenories()
        {
            WaitForLoad();
            var invPhyValueLabel = _webDriver.FindElements(By.XPath("//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]"));

            return invPhyValueLabel.Count;
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

            _allergensList = WaitForElementIsVisible(By.XPath(string.Format(ALLERGENS_LIST)));
            var allergensInList = _allergensList.FindElements(By.TagName("li"));
            if (allergensInList.Count > 0)
            {
                foreach (var allergen in allergensInList)
                {
                    // Check if <li> contains an image, if not move to next iteration.
                    // The allergen won't added to the list ans so the test won't pass.
                    if (allergen.FindElement(By.TagName("img")) == null) continue;
                    item_allergens.Add(allergen.Text);
                }
            }
            return item_allergens;
        }


        public double GetPriceValue(string currency, string decimalSeparator)
        {
            //Récupération du type de séparateur(, ou.selon les pays)
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            _differenceValue = WaitForElementExists(By.Id(DIFFERENCE_VALUE));
            string result = _differenceValue.Text.Replace(currency, "").Replace(" ", "");

            return double.Parse(result, ci);

        }
        public string GetPhysicalValue()
        {
            WaitForLoad();
            try
            {

                var physicalValue = WaitForElementExists(By.XPath(QTY_XPATH));

                return physicalValue.Text;
            }
            catch
            {
                var physicalValue = WaitForElementExists(By.XPath(QTY_XPATH_EMPTY));

                return physicalValue.Text;
            }

        }
        public string GetDiffPriceValue(string currency)
        {
            WaitForLoad();
            var priceValue = WaitForElementExists(By.XPath(DIFF_PRICE_XPATH));
            return priceValue.Text.Replace(currency, "").Replace(" ", "");

        }
        public List<string> GetAllItemsName()
        {
            var items = _webDriver.FindElements(By.XPath(LIST_ITEMS));
            WaitForLoad();
            return items?.Select(item => item.Text.Trim()).ToList() ?? new List<string>();

        }
    }
}