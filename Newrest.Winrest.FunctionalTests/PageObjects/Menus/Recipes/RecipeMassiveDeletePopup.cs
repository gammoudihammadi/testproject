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
using System.Linq;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Menus.Recipes
{
    public class RecipeMassiveDeletePopup : PageBase
    {
        public class MassiveDeleteStatus
        {
            private MassiveDeleteStatus(string value) { Value = value; }

            public string Value { get; private set; }

            public static MassiveDeleteStatus ActiveRecipes { get { return new MassiveDeleteStatus("Active Recipe"); } }
            public static MassiveDeleteStatus InactiveRecipes { get { return new MassiveDeleteStatus("Inactive Recipe"); } }
            public static MassiveDeleteStatus OnlyInactiveSites { get { return new MassiveDeleteStatus("Only inactive sites"); } }
            public static MassiveDeleteStatus Used { get { return new MassiveDeleteStatus("Used"); } }
            public static MassiveDeleteStatus Unused { get { return new MassiveDeleteStatus("Unused"); } }
            public static MassiveDeleteStatus WithUnpurchasableItems { get { return new MassiveDeleteStatus("With unpurchasable items"); } }
            public static MassiveDeleteStatus FlightTypeUnrelatedDS { get { return new MassiveDeleteStatus("Flight type and unrelated to datasheet"); } }
            public static MassiveDeleteStatus RecipeWithoutVariant { get { return new MassiveDeleteStatus("Recipe without variant"); } }
            public static MassiveDeleteStatus EmptyRecipe { get { return new MassiveDeleteStatus("Empty recipe (no items)"); } }

            public override string ToString()
            {
                return Value;
            }
        }

        public RecipeMassiveDeletePopup(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }


        // ______________________________________ Constantes _____________________________________________

        // Général

        private const string VARIANT_RECIPE_WEIGHT = "//*[@id=\"tableRecipes\"]/tbody/tr[*]/td[5]";
        private const string PAGE_SIZE = "/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/nav/select";



        [FindsBy(How = How.XPath, Using = PAGE_SIZE)]
        private IWebElement _pageSize;
        #region Constantes

        private const string SEARCH_PATTERN_XPATH = "//*[@id=\"formMassiveDeleteRecipe\"]/descendant::input[contains(@id, 'SearchPattern')]";
        private const string SITE_FILTER_ID = "siteFilter";
        private const string SITE_FILTER_SHOW_INACTIVE_ID = "ShowInactiveSites";
        private const string VARIANT_FILTER_ID = "SelectedVariants_ms";
        private const string RECIPETYPE_FILTER_ID = "SelectedRecipeTypes_ms";
        private const string STATUS_FILTER_ID = "SelectedStatus_ms";
        private const string SEARCH_BTN_ID = "SearchRecipesBtn";
        private const string SELECTALL_BTN_ID = "selectAll";
        private const string DELETE_BTN = "deleteRecipeBtn";
        private const string CONFIRM_DELETE = "dataConfirmOK";
        private const string UNSELECT = "unselectAll";
        private const string SELECTALL = "selectAll";
        private const string DDL_PAGESIZE = "//*[@id=\"div-recipeResultTable\"]/descendant::select[contains(@id, 'page-size-selector')]";
        private const string RECIPENAME_HEADER = "//*[@id=\"tableRecipes\"]//a[text()='Recipe Name']";
        private const string SITENAME_HEADER = "//*[@id=\"tableRecipes\"]/thead/tr/th[3]/span/a";
        private const string VARIANT_HEADER = "//*[@id=\"tableRecipes\"]/thead/tr/th[4]/span/a";
        private const string WEIGHT_HEADER = "//*[@id=\"tableRecipes\"]/thead/tr/th[5]/span/a";
        private const string USECASE_HEADER = "//*[@id=\"tableRecipes\"]/thead/tr/th[5]/span/a";
        private const string CANCEL_BTN = "/html/body/div[3]/div/div/div[4]/button[1]";

        #endregion

        private static bool CompareStringsBySort(List<string> toCompare, SortType sortType)
        {
            for (int i = 0; i < toCompare.Count - 1; i++)
            {
                int comparisonResult = string.Compare(toCompare[i], toCompare[i + 1], StringComparison.OrdinalIgnoreCase);
                if (comparisonResult > 0 && sortType == SortType.Ascending)
                {
                    return false;
                }
                else if (comparisonResult < 0 && sortType == SortType.Descending)
                {
                    return false;
                }
            }

            return true;
        }

        public void SetRecipeName(string recipeName)
        {
            IWebElement searchPattern = WaitForElementIsVisible(By.XPath(SEARCH_PATTERN_XPATH));
            WaitForLoad();
            searchPattern.SetValue(ControlType.TextBox, recipeName);
        }

        public void SelectSiteByName(string siteName, bool ignoreUncheckAll)
        {
            ComboBoxOptions cbOpt = new ComboBoxOptions(SITE_FILTER_ID, siteName, false) { ClickCheckAllAtStart = false, ClickUncheckAllAtStart = !ignoreUncheckAll };
            ComboBoxSelectById(cbOpt);
        }

        public void SelectAllInactiveSites()
        {
            ComboBoxOptions cbOpt = new ComboBoxOptions(SITE_FILTER_ID, "Inactive", false)
            { ClickCheckAllAtStart = false, ClickCheckAllAfterSelection = true };
            ComboBoxSelectById(cbOpt);
        }

        public void ClickOnInactiveSiteCheck()
        {
            IWebElement checkBoxInactiveSite = WaitForElementExists(By.Id(SITE_FILTER_SHOW_INACTIVE_ID));
            checkBoxInactiveSite.Click();
            WaitForLoad();
        }

        public void SelectVariantByName(string variantName, bool ignoreUncheckAll)
        {
            ComboBoxOptions cbOpt = new ComboBoxOptions(VARIANT_FILTER_ID, variantName, false) { ClickCheckAllAtStart = false, ClickUncheckAllAtStart = !ignoreUncheckAll };
            ComboBoxSelectById(cbOpt);
        }

        public void SelectRecipeTypeByName(string recipeType, bool ignoreUncheckAll)
        {
            ComboBoxOptions cbOpt = new ComboBoxOptions(RECIPETYPE_FILTER_ID, recipeType, false) { ClickCheckAllAtStart = false, ClickUncheckAllAtStart = !ignoreUncheckAll };
            ComboBoxSelectById(cbOpt);
        }

        /// <summary>
        /// L'index à sélectionner dans le multi select du statut. L'index commence à 1.
        /// </summary>
        /// <param name="index"></param>
        public void SelectStatus(string statusLabel)
        {
            ComboBoxOptions cbOpt = new ComboBoxOptions()
            {
                XpathId = "collapseSelectedStatusFilter",
                SelectionValue = statusLabel,
                ClickCheckAllAtStart = false,
                ClickUncheckAllAtStart = false,
                IsUsedInFilter = false
            };
            ComboBoxSelectById(cbOpt);
        }

        public void ClickOnSearch()
        {
            var searchBtn = WaitForElementExists(By.Id(SEARCH_BTN_ID));
            searchBtn.Click();
            WaitPageLoading();
        }

        public void SetPageSize(string size)
        {
            IWebElement _pageSize;
            try
            {
                WaitForElementIsVisible(By.XPath(DDL_PAGESIZE));
                _pageSize = _webDriver.FindElement(By.XPath(DDL_PAGESIZE));
            }
            catch
            {
                // tableau vide : pas de PageSize
                return;
            }
            Actions action = new Actions(_webDriver);
            action.MoveToElement(_pageSize).Perform();
            _pageSize.SetValue(ControlType.DropDownList, size);

            WaitPageLoading();
        }

        /// <summary>
        /// Clique sur le lien revoyant vers la recette, dans la ligne de résultat
        /// </summary>
        /// <param name="rowNumber">Le numéro de ligne qui commence à 1</param>
        /// <returns></returns>
        public RecipeGeneralInformationPage ClickOnRecipeLinkFromRow(int rowNumber)
        {
            var recipeLink = WaitForElementIsVisible(By.XPath("//*[@id=\"tableRecipes\"]/tbody/tr[" + rowNumber + "]/td[7]/a"));
            recipeLink.Click();
            WaitForLoad();
            // switch driver to the opened tab
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            WaitForLoad();
            return new RecipeGeneralInformationPage(_webDriver, _testContext);
        }

        public void ClickOnRecipeNameHeader()
        {
            var recipeHeader = WaitForElementExists(By.XPath(RECIPENAME_HEADER));
            recipeHeader.Click();
            WaitPageLoading();
        }

        public bool VerifySortByRecipeName(SortType sortType)
        {
            List<string> recipeNames = new List<string>();
            _webDriver.FindElements(By.XPath("//*[@id=\"tableRecipes\"]/tbody/tr[*]/td[2]")).ToList().ForEach(elt => recipeNames.Add(elt.Text));

            return CompareStringsBySort(recipeNames, sortType);
        }

        public void ClickOnSiteNameHeader()
        {
            var siteHeader = WaitForElementExists(By.XPath(SITENAME_HEADER));
            siteHeader.Click();
            WaitPageLoading();
        }

        public bool VerifySortByRecipeSite(SortType sortType)
        {
            List<string> siteNames = new List<string>();
            _webDriver.FindElements(By.XPath("//*[@id=\"tableRecipes\"]/tbody/tr[*]/td[3]")).ToList().ForEach(elt => siteNames.Add(elt.Text));

            return CompareStringsBySort(siteNames, sortType);
        }

        public void ClickOnVariantNameHeader()
        {
            var variantHeader = WaitForElementExists(By.XPath(VARIANT_HEADER));
            variantHeader.Click();
            WaitPageLoading();
        }

        public bool VerifySortByVariant(SortType sortType)
        {
            List<string> variants = new List<string>();
            _webDriver.FindElements(By.XPath("//*[@id=\"tableRecipes\"]/tbody/tr[*]/td[4]")).ToList().ForEach(elt => variants.Add(elt.Text));

            return CompareStringsBySort(variants, sortType);
        }

        public RecipeMassiveDeleteRowResult GetRowResultInfo(int itemOffset)
        {
            string rowXpath = "(//*[@id=\"tableRecipes\"]/tbody/tr)[" + (itemOffset + 1) + "]";
            IWebElement rowResult = null;
            try
            {
                rowResult = WaitForElementIsVisible(By.XPath(rowXpath));
            }
            catch
            { //silent catch : ignore if element isn't found
            }

            if (rowResult == null)
            {
                return null;
            }
            else
            {
                RecipeMassiveDeleteRowResult dsRowResult = new RecipeMassiveDeleteRowResult();

                dsRowResult.RecipeName = rowResult.FindElement(By.XPath(rowXpath + "/td[2]")).Text;

                if (string.IsNullOrEmpty(dsRowResult.RecipeName) || dsRowResult.RecipeName.StartsWith("No data"))
                { return null; }

                dsRowResult.IsRecipeInactive = rowResult.FindElement(By.XPath(rowXpath + "/td[2]")).GetAttribute("class").Contains("IsInactive");
                dsRowResult.SiteName = rowResult.FindElement(By.XPath(rowXpath + "/td[3]")).Text;
                dsRowResult.IsSiteInactive = rowResult.FindElement(By.XPath(rowXpath + "/td[3]")).GetAttribute("class").Contains("IsInactive");
                dsRowResult.VariantName = rowResult.FindElement(By.XPath(rowXpath + "/td[4]")).Text;

                string weightString = rowResult.FindElement(By.XPath(rowXpath + "/td[5]")).Text;
                if (string.IsNullOrEmpty(weightString) == false)
                {
                    dsRowResult.Weight = double.Parse(weightString.Replace(",", "."), CultureInfo.InvariantCulture);
                }
                
                dsRowResult.UseCase = int.Parse(rowResult.FindElement(By.XPath(rowXpath + "/td[6]")).Text);
                return dsRowResult;
            }
        }

        public void SelectAllSites()
        {
            ComboBoxOptions cbOpt = new ComboBoxOptions(SITE_FILTER_ID,"",false)
            { ClickCheckAllAtStart = false, ClickCheckAllAfterSelection = true };
            ComboBoxSelectById(cbOpt);
        }

        public void SelectAllVariant()
        {
            ComboBoxOptions cbOpt = new ComboBoxOptions(VARIANT_FILTER_ID, "", false)
            { ClickCheckAllAtStart = false, ClickCheckAllAfterSelection = true };
            ComboBoxSelectById(cbOpt);
        }
        public void SelectAllRecipeType()
        {
            ComboBoxOptions cbOpt = new ComboBoxOptions(RECIPETYPE_FILTER_ID, "", false)
            { ClickCheckAllAtStart = false, ClickCheckAllAfterSelection = true };
            ComboBoxSelectById(cbOpt);
        }
        public void SelectAllStatus()
        {
            ComboBoxOptions cbOpt = new ComboBoxOptions(STATUS_FILTER_ID, "", false)
            { ClickCheckAllAtStart = false, ClickCheckAllAfterSelection = true };
            ComboBoxSelectById(cbOpt);
        }
        public string  GetFirstRecipeName()
        {
            if (isElementVisible(By.XPath("//*[@id=\"tableRecipes\"]/tbody/tr[1]/td[2]")))
            {
                var FirstRecipe = WaitForElementIsVisible(By.XPath("//*[@id=\"tableRecipes\"]/tbody/tr[1]/td[2]"));
                return FirstRecipe.Text;
            }else
            {
                return null;
            }
           
        }
        public void ClickSelectAllButton()
        {
            var btn = WaitForElementIsVisible(By.Id("selectAll"));
            btn.Click();
            WaitForLoad();
        }
        public int CheckTotalSelectCount()
        {
            var count = _webDriver.FindElement(By.Id("recipeCount"));
            var totalnumber = count.Text.Trim();
            return Convert.ToInt32(totalnumber);
        }
        public bool IsPageSizeEqualsTo(string size)
        {
            var nbPages = WaitForElementExists(By.XPath(DDL_PAGESIZE));
            SelectElement select = new SelectElement(nbPages);
            IWebElement selectedOption = select.SelectedOption;
            string selectedValue = selectedOption.GetAttribute("value");

            return selectedValue == size;
        }
        public int GetTotalRowsForPagination()
        {
            var table = _webDriver.FindElement(By.Id("tableRecipes"));

            var allRows = table.FindElements(By.TagName("tr"));

            var totalRows = allRows.Count() - 1;
            return totalRows;
        }
        public class RecipeMassiveDeleteRowResult
        {
            public string RecipeName { get; set; }
            public bool IsRecipeInactive { get; set; }
            public string SiteName { get; set; }
            public bool IsSiteInactive { get; set; }
            public string VariantName { get; set; }
            public double Weight { get; set; }
            public int UseCase { get; set; }
        }
        public List<string> GetAllInactiveSites()
        {
            ComboBoxOptions cbOpt = new ComboBoxOptions(SITE_FILTER_ID, "Inactive", false)
            { ClickCheckAllAtStart = false, ClickCheckAllAfterSelection = true };
            ComboBoxSelectById(cbOpt);
            IWebElement input = WaitForElementIsVisible(By.Id(cbOpt.XpathId));
            input.Click();
            WaitForLoad();
            var elements = _webDriver.FindElements(By.XPath("/html/body/div[19]/ul/li[*]/label/span[contains(text(),\"Inactive\")]"));
            List<string> elementValues =elements.Select(e => e.Text).ToList();
            return elementValues;
        }
        public enum SortType
        {
            Ascending,
            Descending,
        }
        public void CheckRowByRecipeName(string recipeName)
        {
            //var element = WaitForElementIsVisible(By.XPath("//*[@id=\"tableRecipes\"]/tbody/tr[1]/td[1]"));
            var element = WaitForElementIsVisible(By.XPath(string.Format("//*[@id=\"tableRecipes\"]/tbody/tr[*]/td[2][contains(text(),'{0}')]/../td[1]", recipeName)));
            element.Click();
        }
        public void ClickDelete()
        {
            var btn = WaitForElementIsVisible(By.Id(DELETE_BTN));
            btn.Click();
            WaitForLoad();
        }
        public void ClickToConfirmDelete()
        {
            var btn = WaitForElementIsVisible(By.Id(CONFIRM_DELETE));
            btn.Click();
            WaitForLoad();
            var btnOk = WaitForElementIsVisible(By.XPath("//*[@id=\"modal-1\"]/div[3]/button"));
            btnOk.Click();
            WaitForLoad();
        }
        public void CloseRecipeMassiveDeletePopup()
        {
            var btn = WaitForElementIsVisible(By.XPath(CANCEL_BTN));
            btn.Click();
            WaitForLoad();
        }
        public void ClickUseCase()
         {
            var th_usecase = WaitForElementIsVisible(By.XPath("//*[@id=\"tableRecipes\"]/thead/tr/th[6]/span/a"));
            th_usecase.Click();
            WaitForLoad();
        }
        public void ClickWeight()
        {
            var th_usecase = WaitForElementIsVisible(By.XPath("//*[@id=\"tableRecipes\"]/thead/tr/th[5]/span/a"));
            th_usecase.Click();
            WaitForLoad();
        }
        public void ClickOnSelectedAll()
        {
            var btnSelectedAll = WaitForElementExists(By.XPath("//*[@id=\"selectAll\"]"));
            btnSelectedAll.Click();
            WaitForLoad();
        }
        public void ClickOnUnselectAll()
        {
            var btnSelectedAll = WaitForElementExists(By.XPath("//*[@id=\"unselectAll\"]"));
            btnSelectedAll.Click();
            WaitForLoad();
        }
        public string NombreRecipeVariant()
        {
            var nbrecipevariant = WaitForElementExists(By.XPath("//*[@id=\"recipeCount\"]"));
            return nbrecipevariant.Text;
           
        }

        public void GetPageResults(int pagenumber)
        {
            var page = WaitForElementIsVisible(By.XPath("//*[@id='list-recipes-deletion']/nav/ul/li/a[text()='" + pagenumber + "']"));
            page.Click();
            WaitForLoad();
        }

        public int ConvertStringToInt(string value)
        {

            if (int.TryParse(value, out int validResult))
            {
               return validResult;
            }
            return 0;
        }

        public double ConvertStringToDouble(string value)
        {

            if (double.TryParse(value, out double validResult))
            {
                return validResult;
            }
            return 0;
        }

        public string ValueUseCase(int i)
        {
            var nbrecipevariant = WaitForElementExists(By.XPath("//*[@id=\"tableRecipes\"]/tbody/tr["+i+"]/td[6]"));
            return nbrecipevariant.Text;

        }
        public string ValueWeight(int i)
        {
            var nbrecipevariant = WaitForElementExists(By.XPath("//*[@id=\"tableRecipes\"]/tbody/tr["+i+"]/td[5]"));
            return nbrecipevariant.Text;
        }

        public bool IsSortingDescending(List<int> listInt)
        { 
            for (int i = 0; i < listInt.Count-1; i++)
            {
                if (listInt[i] > listInt[i+1]) 
                    return false;
            }
            return true;
        }

        public bool IsSortingDescendingForDoubleList(List<double> listInt)
        {
            for (int i = 0; i < listInt.Count - 1; i++)
            {
                if (listInt[i] > listInt[i + 1])
                    return false;
            }
            return true;
        }
        public void GoToPage(string pageNumber)
        {
            var pages = _webDriver.FindElements(By.XPath("//*[@id=\"list-recipes-deletion\"]/nav/ul/li"));
            foreach (var page in pages)
            {
                if (page.Text.Trim() == pageNumber.Trim())
                {
                    page.Click();
                }
            }

        }
        public bool CheckResultIsAllSelected()
        {
            var results = _webDriver.FindElements(By.XPath("//*[@id=\"item_IsSelected\"]"));
            foreach (var result in results)
            {
                if (!result.Selected)
                {
                    return false;
                }

            }
            return true;
        }
        public bool IsAllResultChecked()
        {
            var results = _webDriver.FindElements(By.XPath("//*[@id=\"tableRecipes\"]/tbody/tr[*]/td[1]/span"));
            foreach (var result in results)
            {
                if (result.GetCssValue("background-color").Equals("rgba(255, 255, 255, 1)"))
                {
                    return false;
                }

            }
            return true;
        }

        public List<double> GetRecipeVariantWeight()
        {
            List<double> weightsList = new List<double>();
            var weights = _webDriver.FindElements(By.XPath(VARIANT_RECIPE_WEIGHT));

            foreach (var weight in weights)
            {
                if (double.TryParse(weight.Text.Replace(',', '.'), out double parsedWeight)
)
                {
                    weightsList.Add(parsedWeight);
                }
            }
            return weightsList;
        }


        public void PageSize(string size)
        {
            if (size == "1000")
            {   // Test
                IJavaScriptExecutor js = (IJavaScriptExecutor)_webDriver;
                js.ExecuteScript("$('#" + PAGE_SIZE + "').append($('<option>', {value: 1000,text: '1000'}),'');");
            }

            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(30));
            try
            {
                wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(PAGE_SIZE)));
            }
            catch
            {
                // tableau vide : pas de PageSize
                return;
            }
            _pageSize = WaitForElementExists(By.XPath(PAGE_SIZE));
            Actions action = new Actions(_webDriver);
            action.MoveToElement(_pageSize).Perform();
            _pageSize = WaitForElementExists(By.XPath(PAGE_SIZE));
            _pageSize.SetValue(ControlType.DropDownList, size);

            WaitPageLoading();
            WaitForLoad();
            // pour écran plus petit que 8 lignes affiché
            PageUp();
        }
    }
}
