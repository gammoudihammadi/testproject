
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Datasheet;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Recipes;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Purchasing
{
    public class ItemUseCasePage : PageBase
    {
        public ItemUseCasePage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // _____________________________________________ Constantes ________________________________________________
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";

        private const string EXPORT_BTN = "btn-export-excel";

        private const string SEARCH_FILTER = "SearchPatternWithAutocomplete";
        private const string INACTIVERECIPES_FILTER = "ShowInactive";

        private const string FIRST_UC_RECIPENAME = "//*[@id=\"tabContentItemContainer\"]/div[2]/div/ul/li[1]/div/div/div/form/div[3]/div[2]/div[2]/p/span";
        private const string FIRST_UC_SITE = "//*[@id=\"tabContentItemContainer\"]/div[2]/div/ul/li[1]/div/div/div/form/div[3]/div[2]/div[3]/span";
        private const string FIRST_UC_VARIANT = "//*[@id=\"tabContentItemContainer\"]/div[2]/div/ul/li[1]/div/div/div/form/div[3]/div[2]/div[4]/span";
        private const string FIRST_UC_RECIPETYPE = "//*[@id=\"tabContentItemContainer\"]/div[2]/div/ul/li[1]/div/div/div/form/div[3]/div[2]/div[1]/span";

        private const string RECIPE_NAME_COLUMN = "//*[@id=\"tabContentItemContainer\"]/div[2]/div/ul/li[*]/div/div/div/form/div[3]/div[2]/div[2]";
        private const string RECIPE_NAMES = "//*[@id=\"tabContentItemContainer\"]/div[2]/div/ul/li[*]/div/div/div/form/div[3]/div[2]/div[2]/p/span[1]";
        private const string DATASHEET_NAMES = "//*[@id=\"tabContentItemContainer\"]/div[2]/div/ul/li[*]/div/div/div/form/div[3]/div[2]/div[2]/p/span[2]";
        private const string NB_PORTION = "//*[@id=\"tabContentItemContainer\"]/div[2]/div/ul/li[*]/div/div/div/form/div[3]/div[2]/div[5]/span[1]";
        private const string NET_WEIGHT = "//*[@id=\"tabContentItemContainer\"]/div[2]/div/ul/li[*]/div/div/div/form/div[3]/div[2]/div[6]/span[1]";
        private const string NET_QTY = "//*[@id=\"tabContentItemContainer\"]/div[2]/div/ul/li[*]/div/div/div/form/div[3]/div[2]/div[7]/span[1]";
        private const string GROSS_QTY = "//*[@id=\"tabContentItemContainer\"]/div[2]/div/ul/li[*]/div/div/div/form/div[3]/div[2]/div[8]/span[1]";
        private const string YIELD = "//*[@id=\"tabContentItemContainer\"]/div[2]/div/ul/li[*]/div/div/div/form/div[3]/div[2]/div[9]/span[1]";
        private const string SITE_COLUMN = "//*[@id=\"tabContentItemContainer\"]/div[2]/div/ul/li[*]/div/div/div/form/div[3]/div[2]/div[3]";
        private const string VARIANT_COLUMN = "//*[@id=\"tabContentItemContainer\"]/div[2]/div/ul/li[*]/div/div/div/form/div[3]/div[2]/div[4]";
        private const string RECIPETYPE_COLUMN = "//*[@id=\"tabContentItemContainer\"]/div[2]/div/ul/li[*]/div/div/div/form/div[3]/div[2]/div[1]";

        private const string SORT_BY = "cbSortBy";
        private const string RESET_FILTER = "reset-button";
        private const string SITE_FILTER = "SelectedSites_ms";
        private const string UNCHECK_ALL_SITES = "/html/body/div[11]/div/ul/li[2]/a";
        private const string SEARCH_SITE = "/html/body/div[11]/div/div/label/input";

        private const string VARIANT_FILTER = "SelectedVariants_ms";
        private const string UNCHECK_ALL_VARIANTS = "/html/body/div[13]/div/ul/li[2]/a";
        private const string SEARCH_VARIANT  = "/html/body/div[13]/div/div/label/input";

        private const string RECIPETYPES_FILTER = "SelectedRecipeTypes_ms";
        private const string UNCHECK_ALL_RECIPETYPES = "/html/body/div[12]/div/ul/li[2]/a";
        private const string SEARCH_RECIPETYPES = "/html/body/div[12]/div/div/label/input";
        private const string SELECT_USE_CASE = "//*/input[@class='select-use-case']";
        private const string SELECT_USE_CASE_NAME = "//*/form/div[3]/div[2]/div[2]/p/span[text()='{0}']/../../../../../div[1]/div/input";

        private const string REPLACE_BY_ANOTHER_ITEM_BUTTON = "//*[@id='tabContentItemContainer']/div[1]/div/div[1]/a[2]";
        private const string REPLACE_BY_ANOTHER_ITEM_SEARCH = "searchVM_SearchPatternWithAutocomplete";
        private const string REPLACE_BY_ANOTHER_ITEM_FIRST_RESULT = "//*[@id='search-replacement-item-form']/div/span/div/div/div[1]";
        private const string REPLACE_BY_ANOTHER_ITEM_REPLACE = "//*/a[text()='Replace with this item']/parent::div";
        private const string REPLACE_BY_ANOTHER_ITEM_CONFIRM = "dataConfirmOK";
        private const string REPLACE_BY_ANOTHER_ITEM_CLOSE = "//*[@id='modal-1']/div/div/div[2]/div[2]/button";

        private const string MULTIPLE_UPDATE_BUTTON = "//*/a[text()='Multiple update']";
        private const string MULTIPLE_UPDATE_FIELD_TO_UPDATE = "FieldToUpdate";
        private const string MULTIPLE_UPDATE_SELECTED_WORKSHOP = "SelectedWorkshop";
        private const string MULTIPLE_UPDATE_UPDATED_VALUE = "UpdatedValue";
        private const string MULTIPLE_UPDATE_UPDATE = "//*/a[text()='Update']";
        private const string MULTIPLE_UPDATE_CONFIRM = "dataConfirmOK";
        private const string MULTIPLE_UPDATE_CONCLUDE = "//*/div[@id='modal-1']//button[text()='Close']";
        private const string MULTIPLE_UPDATE_RESULT = "(//*/form/div[3]/div[2]/div[{0}]/span[1])[1]";
        private const string MULTIPLE_UPDATE_RESULT_FORM = "(//*/form/div[3]/div[1]/div[{0}]/div/div[2]/input)[1]";
        private const string SHOW_Inactive = "//*[@id=\"tabContentItemContainer\"]/div[2]/div/ul/li[*]";
        private const string FIRST_RECIPE_STYLO = "//*[@id=\"tabContentItemContainer\"]/div[2]/div/ul/li[1]/div/div/div/form/div[3]/div[1]/div[11]/a/span";
        public enum ColumnName : int
        {
            Yield = 9,
            NetWeight = 6,
            NetQty = 7,
            GrossQty = 8,
            Workshop = 10
        }

        // _____________________________________________ Variables _______________________________________________
        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;

        [FindsBy(How = How.Id, Using = EXPORT_BTN)]
        private IWebElement _export;

        [FindsBy(How = How.Id, Using = SORT_BY)]
        private IWebElement _sortBy;

        [FindsBy(How = How.Id, Using = RESET_FILTER)]
        private IWebElement _resetFilter;


        [FindsBy(How = How.XPath, Using = SEARCH_SITE)]
        private IWebElement _searchSite;

        [FindsBy(How = How.Id, Using = SITE_FILTER)]
        private IWebElement _siteFilter;

        [FindsBy(How = How.XPath, Using = UNCHECK_ALL_SITES)]
        private IWebElement _uncheckAllSites;


        [FindsBy(How = How.Id, Using = VARIANT_FILTER)]
        private IWebElement _variantFilter;

        [FindsBy(How = How.XPath, Using = UNCHECK_ALL_VARIANTS)]
        private IWebElement _uncheckAllVariants;

        [FindsBy(How = How.XPath, Using = SEARCH_VARIANT)]
        private IWebElement _searchVariant;


        [FindsBy(How = How.Id, Using = RECIPETYPES_FILTER)]
        private IWebElement _recipeTypesFilter;

        [FindsBy(How = How.XPath, Using = UNCHECK_ALL_RECIPETYPES)]
        private IWebElement _uncheckAllRecipeTypes;

        [FindsBy(How = How.XPath, Using = SEARCH_RECIPETYPES)]
        private IWebElement _searchRecipeTypes;

        [FindsBy(How = How.XPath, Using = SELECT_USE_CASE)]
        private IWebElement _selectUseCase;

        [FindsBy(How = How.XPath, Using = SELECT_USE_CASE_NAME)]
        private IWebElement _selectUseCaseName;

        [FindsBy(How = How.XPath, Using = REPLACE_BY_ANOTHER_ITEM_BUTTON)]
        private IWebElement _replaceByAnotherItemButton;

        [FindsBy(How = How.Id, Using = REPLACE_BY_ANOTHER_ITEM_SEARCH)]
        private IWebElement _replaceByAnotherItemSearch;

        [FindsBy(How = How.XPath, Using = REPLACE_BY_ANOTHER_ITEM_FIRST_RESULT)]
        private IWebElement _replaceByAnotherItemFirstResult;

        [FindsBy(How = How.XPath, Using = REPLACE_BY_ANOTHER_ITEM_REPLACE)]
        private IWebElement _replaceByAnotherItemReplace;

        [FindsBy(How = How.Id, Using = REPLACE_BY_ANOTHER_ITEM_CONFIRM)]
        private IWebElement _replaceByAnotherItemConfirm;

        [FindsBy(How = How.XPath, Using = REPLACE_BY_ANOTHER_ITEM_CLOSE)]
        private IWebElement _replaceByAnotherItemClose;

        [FindsBy(How = How.XPath, Using = MULTIPLE_UPDATE_BUTTON)]
        private IWebElement _multipleUpdateButton;

        [FindsBy(How = How.Id, Using = MULTIPLE_UPDATE_FIELD_TO_UPDATE)]
        private IWebElement _multipleUpdateFieldToUpdate;

        [FindsBy(How = How.Id, Using = MULTIPLE_UPDATE_SELECTED_WORKSHOP)]
        private IWebElement _multipleUpdateSelectedWorkshop;

        [FindsBy(How = How.Id, Using = MULTIPLE_UPDATE_UPDATED_VALUE)]
        private IWebElement _multipleUpdateUpdatedValue;

        [FindsBy(How = How.XPath, Using = MULTIPLE_UPDATE_UPDATE)]
        private IWebElement _multipleUpdateUpdate;

        [FindsBy(How = How.Id, Using = MULTIPLE_UPDATE_CONFIRM)]
        private IWebElement _multipleUpdateConfirm;

        [FindsBy(How = How.XPath, Using = MULTIPLE_UPDATE_CONCLUDE)]
        private IWebElement _multipleUpdateConclude;

        [FindsBy(How = How.XPath, Using = MULTIPLE_UPDATE_RESULT)]
        private IWebElement _multipleUpdateResult;

        // ___________________________________________ Filtres __________________________________________________

        [FindsBy(How = How.Id, Using = SEARCH_FILTER)]
        private IWebElement _searchFilter;
        public enum FilterType
        {
            Search,
            SortBy,
            Site,
            ShowInactiveRecipe,
            RecipeTypes,
            Variant,
            UncheckAllRecipeType
        }

        public PrintReportPage Export(bool versionPrint)
        {
            _export = WaitForElementIsVisible(By.Id(EXPORT_BTN));
            _export.Click();
            WaitForLoad();

            if (versionPrint)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));
            }

            WaitForDownload();
            return new PrintReportPage(_webDriver, _testContext);
        }

        public void ResetFilter()
        {
            _resetFilter = WaitForElementIsVisible(By.Id(RESET_FILTER));
            _resetFilter.Click();
            WaitForLoad();

            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                // pas de date
            }
        }

        public void Filter(FilterType filterType, object value)
        {
            switch (filterType)
            {
                case FilterType.Search:
                    _searchFilter = WaitForElementIsVisible(By.Id(SEARCH_FILTER));
                    _searchFilter.SetValue(ControlType.TextBox, value);
                    _searchFilter.SendKeys(Keys.Tab);
                    break;

                case FilterType.SortBy:
                    _sortBy = WaitForElementIsVisible(By.Id(SORT_BY));
                    _sortBy.Click();
                    var element = WaitForElementIsVisible(By.XPath("//*[@id=\"cbSortBy\"]/option[@value='" + value + "']"));
                    _sortBy.SetValue(ControlType.DropDownList, element.Text);
                    _sortBy.Click();
                    break;

                case FilterType.Site:
                    _siteFilter = WaitForElementIsVisible(By.Id(SITE_FILTER));
                    _siteFilter.Click();

                    _uncheckAllSites = WaitForElementIsVisible(By.XPath(UNCHECK_ALL_SITES));
                    _uncheckAllSites.Click();

                    _searchSite = WaitForElementIsVisible(By.XPath(SEARCH_SITE));
                    _searchSite.SetValue(ControlType.TextBox, value);

                    var siteToCheck = WaitForElementIsVisible(By.XPath("//span[text()='" + value + " - " + value + "']"));
                    siteToCheck.SetValue(ControlType.CheckBox, true);

                    _siteFilter.Click();
                    break;

                case FilterType.ShowInactiveRecipe:
                    var inactiveRecipes = WaitForElementExists(By.Id(INACTIVERECIPES_FILTER));
                    inactiveRecipes.SetValue(ControlType.CheckBox, value);
                    WaitPageLoading();
                    break;

                case FilterType.Variant:
                    ComboBoxSelectById(new ComboBoxOptions(VARIANT_FILTER, value.ToString()));
                    break;

                case FilterType.RecipeTypes:
                    ComboBoxSelectById(new ComboBoxOptions(RECIPETYPES_FILTER, value.ToString()));
                    break;

                case FilterType.UncheckAllRecipeType:

                   var _recipeTypeFilter = WaitForElementIsVisible(By.Id("SelectedRecipeTypes_ms"));
                    _recipeTypeFilter.Click();

                    _uncheckAllRecipeTypes = WaitForElementIsVisible(By.XPath("/html/body/div[12]/div/ul/li[2]/a"));
                    _uncheckAllRecipeTypes.Click();

                    var recipeTypeToCheck = WaitForElementIsVisible(By.XPath("/html/body/div[12]/div/div/label/input"));
                    recipeTypeToCheck.SetValue(ControlType.TextBox, value);
                    recipeTypeToCheck.SetValue(ControlType.CheckBox, true);

                    _recipeTypesFilter = WaitForElementIsVisible(By.Id("ui-multiselect-4-SelectedRecipeTypes-option-3"));
                    _recipeTypesFilter.Click();
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);

            }
            WaitPageLoading();
            WaitForLoad();
        }

        public string GetFirstUseCaseRecipeName()
        {
            if (isElementVisible(By.XPath(FIRST_UC_RECIPENAME)))
            {
                var firstUseCaseRN = WaitForElementIsVisible(By.XPath(FIRST_UC_RECIPENAME));
                return firstUseCaseRN.Text;

            }
            else
            {
                return "";
            }
        }

        public string GetFirstUseCaseSite()
        {
            if (isElementVisible(By.XPath(FIRST_UC_SITE)))
            {
                var firstUseCaseSite = WaitForElementIsVisible(By.XPath(FIRST_UC_SITE));
                return firstUseCaseSite.Text;

            }
            else
            {
                return "";
            }
        }

        public string GetFirstUseCaseVariant()
        {
            if (isElementVisible(By.XPath(FIRST_UC_VARIANT)))
            {
                var firstUseCaseVariant = WaitForElementIsVisible(By.XPath(FIRST_UC_VARIANT));
                return firstUseCaseVariant.Text;

            }
            else
            {
                return "";
            }
        }

        public string GetFirstUseCaseRecipeType()
        {
            if (isElementVisible(By.XPath(FIRST_UC_RECIPETYPE)))
            {
                var firstUseCaseVariant = WaitForElementIsVisible(By.XPath(FIRST_UC_RECIPETYPE));
                return firstUseCaseVariant.Text;

            }
            else
            {
                return "";
            }
        }
        

        public bool IsSearchFilterOK(string filteredName)
        {
            var listeNames = _webDriver.FindElements(By.XPath(RECIPE_NAME_COLUMN));

            foreach (var elm in listeNames)
            {
                if (!elm.Text.Contains(filteredName))
                {
                    return false;
                }
            }
            return true;
        }

        public bool IsSiteFilterOK(string filteredSite)
        {
            var listeSites = _webDriver.FindElements(By.XPath(SITE_COLUMN));

            foreach (var elm in listeSites)
            {
                if (elm.Text != filteredSite)
                {
                    return false;
                }
            }
            return true;
        }

        public bool isVariantFilterOK(string filteredVariant)
        {
            var listeVariants = _webDriver.FindElements(By.XPath(VARIANT_COLUMN));

            foreach (var elm in listeVariants)
            {
                if (elm.Text != filteredVariant)
                {
                    return false;
                }
            }
            return true;
        }

        public bool isRecipeTypesFilterOK(string filteredRecipeType)
        {
            var listeVariants = _webDriver.FindElements(By.XPath(RECIPETYPE_COLUMN));

            foreach (var elm in listeVariants)
            {
                if (elm.Text != filteredRecipeType)
                {
                    return false;
                }
            }
            return true;
        }

        public string GetSearchFilterText()
        {
            var textSearchFilter = _webDriver.FindElement(By.Id("SearchPatternWithAutocomplete"));
            return textSearchFilter.GetAttribute("value");
        }
        public bool GetInactiveRecipesBool()
        {
            var inactiveRecipesBool = _webDriver.FindElement(By.Id("ShowInactive"));
            return inactiveRecipesBool.Selected;
        }

        public string GetSiteSelectedNumber()
        {
            var nbSiteSelected = _webDriver.FindElement(By.XPath("//*[@id=\"SelectedSites_ms\"]/span[2]"));
            return nbSiteSelected.GetAttribute("innerText");
        }

        public string GetRecipeTypesSelectedNumber()
        {
            var nbRecipeTypesSelected = _webDriver.FindElement(By.XPath("//*[@id=\"SelectedRecipeTypes_ms\"]/span[2]"));
            return nbRecipeTypesSelected.GetAttribute("innerText");
        }

        public int CheckSelectedNumber()
        {
            var nbSelected = _webDriver.FindElement(By.XPath("//*[@id='tabContentItemContainer']/div[1]/h1/span[2]"));
            int nombre = Int32.Parse(nbSelected.Text);
            return nombre;
        }

        public void SelectAllRecipeTypes()
        {
            var selectAll = WaitForElementIsVisible(By.XPath("//*[@id='tabContentItemContainer']/div[1]/div/a[4]"));
            selectAll.Click();
        }

        public void UnSelectAllRecipeTypes()
        {
            var unSelectAll = WaitForElementIsVisible(By.XPath("//*[@id='tabContentItemContainer']/div[1]/div/a[3]"));
            unSelectAll.Click();
        }

        public string GetVariantSelectedNumber()
        {
            var nbVariantSelected = _webDriver.FindElement(By.XPath("//*[@id=\"SelectedVariants_ms\"]/span[2]"));
            return nbVariantSelected.GetAttribute("innerText");
        }

        public void PrintUseCaseReport()
        {
            var printUseCase = WaitForElementIsVisible(By.Id("btn-print-use-case-report"));
            printUseCase.Click();
            WaitForLoad();

            WaitForDownload();
            //Close();
        }

        public void CheckPrintUseCaseReport(FileInfo filePdf)
        {
            PdfDocument document = PdfDocument.Open(filePdf.FullName);
            List<string> mots = new List<string>();
            foreach (Page p in document.GetPages())
            {
                foreach (var mot in p.GetWords())
                {
                    mots.Add(mot.Text);
                }
            }

            var sites = _webDriver.FindElements(By.XPath("//*/div[@class='display-zone']/div[3]/span"));
            foreach (var site in sites)
            {
                if (site.Text == "") continue;
                Assert.IsTrue(mots.Contains(site.Text), "site " + site.Text + " non présent dans le PDF");
            }

            var datasheets = _webDriver.FindElements(By.XPath("//*/div[@class='display-zone']/div[2]/p/span/i"));
            foreach (var datasheet in datasheets)
            {
                if (datasheet.Text == "") continue;
                foreach (string mot in datasheet.Text.Split(' '))
                {
                    if (mot.Contains("OA"))
                    {
                        // "O"+"A"
                        continue;
                    }
                    if (mot.Contains("QU"))
                    {
                        // "Q"+"U"
                        continue;
                    }
                    if (mot.Contains("(") || mot.Contains(")") || mot.Contains("-"))
                    {
                        continue;
                    }
                    Assert.IsTrue(mots.Contains(mot), "datasheet " + datasheet.Text + " non présent dans le PDF");
                }
            }
        }
        public void CheckExportUseCase(FileInfo trouveXLSX, string itemName)
        {
            string sheetName = itemName;
            if (itemName.Length> "00061_UAL_SELTZER SEAGRAMSd".Length)
            {
                sheetName = itemName.Substring(0, "00061_UAL_SELTZER SEAGRAMSd".Length) + "...";
            }
            //check number og lines
            int resultNumber = OpenXmlExcel.GetExportResultNumber(sheetName, trouveXLSX.FullName);
            Assert.AreEqual(CheckTotalNumber(), resultNumber, "Mauvais nombre de lignes");
            List<string> itemNameList = OpenXmlExcel.GetValuesInList("Item", sheetName, trouveXLSX.FullName);

            //check item name
            Assert.IsTrue(itemNameList.All(x => x == itemName), "Le nom de l'item sur chaque ligne de l'export n'est pas le bon.");

            int counter = 0;

            //check site
            List<string> siteXLSX = OpenXmlExcel.GetValuesInList("Site", sheetName, trouveXLSX.FullName);

            var sites = _webDriver.FindElements(By.XPath(SITE_COLUMN));
            foreach (var site in sites)
            {
                if (site.Text == "") continue;
                Assert.IsTrue(siteXLSX.Contains(site.Text), "site " + site.Text + " non présent dans le XLSX");
            }

            //check recipe name
            List<string> recipeNameXLSX = OpenXmlExcel.GetValuesInList("Recipe name", sheetName, trouveXLSX.FullName);
            var recipeNames = _webDriver.FindElements(By.XPath(RECIPE_NAMES));
            foreach (var recipeName in recipeNames)
            {
                if (recipeName.Text == "") continue;
                counter++;
                Assert.IsTrue(recipeNameXLSX.Any(x => x.Contains(recipeName.Text)), "Recipe Name " + recipeName.Text + " non présent dans le XLSX");
            }

            //check
            List<string> datasheetXLSX = OpenXmlExcel.GetValuesInList("Datasheet", sheetName, trouveXLSX.FullName);
            var datasheets = _webDriver.FindElements(By.XPath(DATASHEET_NAMES));
            foreach (var datasheet in datasheets)
            {
                if (datasheet.Text == "") continue;
                counter++;
                Assert.IsTrue(datasheetXLSX.Any(x => x.Contains(datasheet.Text)), "Datasheet " + datasheet.Text + " non présent dans le XLSX");
            }
            List<string> numberOfPortionsXLSX = OpenXmlExcel.GetValuesInList("NumberOfPortions", sheetName, trouveXLSX.FullName);
            var numberOfPortions = _webDriver.FindElements(By.XPath(NB_PORTION));
            foreach (var np in numberOfPortions)
            {
                if (np.Text == "") continue;
                counter++;
                Assert.IsTrue(numberOfPortionsXLSX.Any(x => x.Contains(np.Text)), "NumberOfPortions " + np.Text + " non présent dans le XLSX");
            }
            List<string> netWeightXLSX = OpenXmlExcel.GetValuesInList("NetWeight", sheetName, trouveXLSX.FullName);
            var netWeights = _webDriver.FindElements(By.XPath(NET_WEIGHT));
            foreach (var nw in netWeights)
            {
                if (nw.Text == "") continue;
                counter++;
                Assert.IsTrue(netWeightXLSX.Any(x => x.Contains(nw.Text)), "NetWeight" + nw.Text + " non présent dans le XLSX");
            }
            List<string> netQuantityXLSX = OpenXmlExcel.GetValuesInList("NetQuantity", sheetName, trouveXLSX.FullName);
            var netQuantitys = _webDriver.FindElements(By.XPath(NET_QTY));
            foreach (var nq in netQuantitys)
            {
                if (nq.Text == "") continue;
                counter++;
                Assert.IsTrue(netQuantityXLSX.Any(x => x.Contains(nq.Text)), "NetQuantity" + nq.Text + " non présent dans le XLSX");
            }
            List<string> quantityXLSX = OpenXmlExcel.GetValuesInList("Quantity", sheetName, trouveXLSX.FullName);
            var quantitys = _webDriver.FindElements(By.XPath(GROSS_QTY));
            foreach (var q in quantitys)
            {
                if (q.Text == "") continue;
                counter++;
                Assert.IsTrue(quantityXLSX.Any(x => x.Contains(q.Text)), "Quantity" + q.Text + " non présent dans le XLSX");
            }
            List<string> yieldXLSX = OpenXmlExcel.GetValuesInList("Yield", sheetName, trouveXLSX.FullName);
            var yields = _webDriver.FindElements(By.XPath(YIELD));
            foreach (var yield in yields)
            {
                if (yield.Text == "") continue;
                counter++;
                Assert.IsTrue(yieldXLSX.Any(x => x.Contains(yield.Text)), "Yield" + yield.Text + " non présent dans le XLSX");
            }
        }

        public ItemPage BackToList()
        {
            _backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            _backToList.Click();
            WaitForLoad();

            return new ItemPage(_webDriver, _testContext);
        }

        public void SelectBoxFirstUseCase()
        {
            var form = WaitForElementIsVisible(By.XPath("(//*/form/div[3]/div[2]/div[2]/p/span[1])[1]"));
            form.Click();
            WaitForLoad();
            _selectUseCase = WaitForElementExists(By.XPath(SELECT_USE_CASE));
            _selectUseCase.SetValue(PageBase.ControlType.CheckBox, true);
            WaitPageLoading();
        }
        public void SelectFirstUseCase()
        {
            var form = WaitForElementIsVisible(By.XPath("(//*/form/div[3]/div[2]/div[2]/p/span[1])[1]"));
            form.Click();
            WaitForLoad();
        }

        public string ReplaceByAnotherItem(string itemName, bool fullName = false)
        {
            _replaceByAnotherItemButton = WaitForElementIsVisible(By.XPath(REPLACE_BY_ANOTHER_ITEM_BUTTON));
            _replaceByAnotherItemButton.Click();
            WaitForLoad();
            _replaceByAnotherItemSearch = WaitForElementIsVisible(By.Id(REPLACE_BY_ANOTHER_ITEM_SEARCH));
            _replaceByAnotherItemSearch.SetValue(PageBase.ControlType.TextBox, itemName);
            WaitForLoad();

            var selectedItem = WaitForElementIsVisible(By.XPath("//*[@id=\"search-replacement-item-form\"]/div[2]/div[1]/span/div/div/div[1]"));
            selectedItem.Click();
            WaitForLoad();


            var itemNameTo = itemName;
            if (fullName)
            {
                // le bouton bouge/apparait un peu
                Thread.Sleep(2000);
                _replaceByAnotherItemSearch.SendKeys(Keys.Tab);
                // le bouton bouge un peus
                Thread.Sleep(2000);
            } else
            {
                _replaceByAnotherItemFirstResult = WaitForElementIsVisible(By.XPath("//*[@id='list-item-with-action']/div/div[1]/div/div[2]/table/tbody/tr/td[1]"));

                itemNameTo = _replaceByAnotherItemFirstResult.Text.Trim();

            }
            WaitForLoad();
            _replaceByAnotherItemReplace = WaitForElementIsVisible(By.XPath(REPLACE_BY_ANOTHER_ITEM_REPLACE));
            _replaceByAnotherItemReplace.Click();
            WaitForLoad();

            _replaceByAnotherItemConfirm = WaitForElementIsVisible(By.Id(REPLACE_BY_ANOTHER_ITEM_CONFIRM));
            _replaceByAnotherItemConfirm.Click();
            WaitPageLoading();
            WaitForLoad();
                _replaceByAnotherItemClose = WaitForElementIsVisible(By.XPath("//*/button[text()='Close' and not(@id)]"));
            _replaceByAnotherItemClose.Click();
            
            
            WaitPageLoading();
            WaitForLoad();
            return itemNameTo;
        }

        public void SelectUseCase(string recipeName)
        {
            _selectUseCaseName = WaitForElementExists(By.XPath(String.Format(SELECT_USE_CASE_NAME,recipeName)));
            _selectUseCaseName.SetValue(PageBase.ControlType.CheckBox, true);
            WaitForLoad();
        }

        public void MultipleUpdate(string selectName, string value)
        {
            //2.Cliquer sur " Multiple update"
            _multipleUpdateButton = WaitForElementIsVisible(By.XPath(MULTIPLE_UPDATE_BUTTON));
            _multipleUpdateButton.Click();
            WaitForLoad();

            //3.Choisir FieldToUpdate
            _multipleUpdateFieldToUpdate = WaitForElementIsVisible(By.Id(MULTIPLE_UPDATE_FIELD_TO_UPDATE));
            SelectElement select = new SelectElement(_multipleUpdateFieldToUpdate);
            select.SelectByText(selectName);
            WaitForLoad();

            if (selectName=="Workshop")
            {
                _multipleUpdateSelectedWorkshop = WaitForElementIsVisible(By.Id(MULTIPLE_UPDATE_SELECTED_WORKSHOP));
                select = new SelectElement(_multipleUpdateSelectedWorkshop);
                select.SelectByText(value);
                WaitForLoad();
            }
            else
            {
                _multipleUpdateUpdatedValue = WaitForElementIsVisible(By.Id(MULTIPLE_UPDATE_UPDATED_VALUE));
                _multipleUpdateUpdatedValue.ClearElement();
                _multipleUpdateUpdatedValue.SendKeys(value);
                WaitForLoad();
            }

            //4.Cliquer sur "Update"
            _multipleUpdateUpdate = WaitForElementIsVisible(By.XPath(MULTIPLE_UPDATE_UPDATE));
            _multipleUpdateUpdate.Click();
            WaitForLoad();
            _multipleUpdateConfirm = WaitForElementIsVisible(By.Id(MULTIPLE_UPDATE_CONFIRM));
            _multipleUpdateConfirm.Click();
            WaitForLoad();
            WaitPageLoading();
            _multipleUpdateConclude = WaitForElementIsVisible(By.XPath(MULTIPLE_UPDATE_CONCLUDE));
            _multipleUpdateConclude.Click();
            WaitForLoad();
        }

        public string GetColumnFirstValue(ColumnName columnEnum)
        {
            int column = (int)columnEnum;
            _multipleUpdateResult = WaitForElementIsVisible(By.XPath(string.Format(MULTIPLE_UPDATE_RESULT,column)));
            return _multipleUpdateResult.Text;
        }

        public string GetColumnFirstValueForm(ColumnName columnEnum)
        {            
            int column = (int)columnEnum;
            _multipleUpdateResult = WaitForElementIsVisible(By.XPath(string.Format(MULTIPLE_UPDATE_RESULT_FORM, column)));
            return _multipleUpdateResult.GetAttribute("value");
        }

        public void SetColumnFirstValueForm(ColumnName columnEnum, string value)
        {
            int column = (int)columnEnum;
            _multipleUpdateResult = WaitForElementIsVisible(By.XPath(string.Format(MULTIPLE_UPDATE_RESULT_FORM, column)));
            _multipleUpdateResult.SetValue(ControlType.TextBox, value);
            WaitForLoad();
            WaitPageLoading();
        }

        public RecipesVariantPage ClickStyloFirstRecipe(out bool isRecipe)
        {
            // un stylo peut aussi être un DataSheet
            var stylo = WaitForElementIsVisible(By.XPath(FIRST_RECIPE_STYLO));
            isRecipe = stylo.FindElement(By.XPath("..")).GetAttribute("href").Contains("Recipe");
            stylo.Click();
            // nouveau onglet !!!
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);
            WaitForLoad();
            return new RecipesVariantPage(_webDriver, _testContext);
        }
        public DatasheetDetailsPage ClickOnUseCaseDatasheetDetailPage(out bool isDatasheet)
        {
            // un stylo peut aussi être un DataSheet
            var stylo = WaitForElementIsVisible(By.XPath(FIRST_RECIPE_STYLO));
            isDatasheet = stylo.FindElement(By.XPath("..")).GetAttribute("href").Contains("Datasheet");
            stylo.Click();
            // nouveau onglet !!!
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[_webDriver.WindowHandles.Count-1]);
            WaitForLoad();
         

            return new DatasheetDetailsPage(_webDriver, _testContext);
        }

        public string GetFirstRecipeName()
        {
            // un stylo peut aussi être un DataSheet
            var recipe = WaitForElementIsVisible(By.XPath("//*[@id=\"tabContentItemContainer\"]/div[2]/div/ul/li[1]/div/div/div/form/div[3]/div[1]/div[2]/p"));
            string recipeName = recipe.Text;
            return recipeName;
        }
        
        public bool VerifierRecipeExist(string recipe)
        {
            Filter(FilterType.Search, recipe);
            var number = CheckTotalNumber();
            return number > 0;
        }
        public int GetNumberOfItems()
        {
            var numberofItems = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div[2]/div[2]/div/div/div[2]/div[1]/h1/span[1]"));
            return int.Parse(numberofItems.Text);
        }

        public bool verifyUseCaseIsNotCheked()
        {
            var listeUseCases = _webDriver.FindElements(By.XPath(SELECT_USE_CASE));
            foreach (var elm in listeUseCases)
            {
                if ( elm.Selected)
                {
                    return false;
                }
            }
            return true;
        }
        public bool verifyShowInactive( bool value )
        {
            var listeUseCases = _webDriver.FindElements(By.XPath(SHOW_Inactive));
            bool result =true;
            foreach (var elm in listeUseCases)
            {
                try
                {
                    var img = elm.FindElements(By.XPath(".//div/div/div/form/div[2]/img"));
                    if ( img.Count > 0 )
                    {
                        if ( value==false )
                        {
                            return false;
                        }
                    }
                }
                catch
                {result = true;

                }
                

            }
            return result;
        }

        public bool MultipleUpdateIsDisabled()
        {
            _multipleUpdateButton = WaitForElementIsVisible(By.XPath(MULTIPLE_UPDATE_BUTTON));
           if (_multipleUpdateButton.GetAttribute("class").Contains("not-allowed-cursor disabled-color"))
           {
            return true;
           }            
            return false;
        }
        public bool ReplaceByAnotherItemButtonIsDisabled()
        {
            _replaceByAnotherItemButton = WaitForElementIsVisible(By.XPath(REPLACE_BY_ANOTHER_ITEM_BUTTON));
            if (_replaceByAnotherItemButton.GetAttribute("class").Contains("not-allowed-cursor disabled-color"))
            {
                return true;
            }
            return false;
        }
        public void UnSelectBoxFirstUseCase()
        {
            _selectUseCase = WaitForElementExists(By.XPath(SELECT_USE_CASE));
            _selectUseCase.SetValue(PageBase.ControlType.CheckBox, false);
            WaitForLoad();
        }
        public DatasheetDetailsPage EditFirstUseCase()
        {
            var form = WaitForElementIsVisible(By.Id("menus-datasheet-edit-1"));
            form.Click();
            WaitForLoad();
            return new DatasheetDetailsPage(_webDriver, _testContext);

        }

        public DatasheetDetailsPage SelectFirstDatasheet()
        {
            var _firstDatasheet = WaitForElementIsVisible(By.XPath("//*[@id=\"dataItem1272424\"]/div/div[2]/div[1]/div[1]"));
            _firstDatasheet.Click();
            WaitForLoad();
            return new DatasheetDetailsPage(_webDriver, _testContext);
        }

        public void Go_To_New_Navigate(int ongletChrome = 1)
        {
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > ongletChrome);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[ongletChrome]);
        }
        public DatasheetGeneralInformationPage ClickOnGeneralInformation()
        {
           var _generalInformationTab = WaitForElementIsVisible(By.Id("hrefTabContentVariantDetailsInfos"));
            _generalInformationTab.Click();
            WaitForLoad();

            return new DatasheetGeneralInformationPage(_webDriver, _testContext);
        }
        public string GetCookingModeSelected()
        {
            var cookingModeSelected = _webDriver.FindElement(By.Id("CookingModeId"));
            var fullText = cookingModeSelected.GetAttribute("innerText");

            var matches = Regex.Matches(fullText, @"^(?!None|Vacío)(.+)$", RegexOptions.Multiline);

            return matches.Count > 0 ? matches[0].Value.Trim() : string.Empty;
        }

        public void CloseEditUseCaseDatasheetForRecipe()
        {
            var form = WaitForElementIsVisible(By.Id("btn-from-datasheet-close-modal"));
            form.Click();
            WaitForLoad();
        }


        public void SelectAll()
        {
            var selectall = WaitForElementIsVisible(By.Id("selectallBtn"));
            selectall.Click();
            WaitForLoad();

        }
    }
}
