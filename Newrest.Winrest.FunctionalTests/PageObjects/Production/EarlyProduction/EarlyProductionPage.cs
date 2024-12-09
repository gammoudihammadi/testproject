using DocumentFormat.OpenXml.Bibliography;
using iText.StyledXmlParser.Jsoup.Nodes;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Production.EarlyProduction
{
    public class EarlyProductionPage : PageBase
    {

        public EarlyProductionPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        private const string FILTER_SITE = "SiteId";
        private const string FILTER_START_DATE = "StartDate";
        private const string FILTER_END_DATE = "EndDate";
        private const string FILTRE_SHOW_HIDDEN_ARTICLES = "ShowHiddenArticles";
        private const string FILTRE_CUSTOMER = "SelectedCustomers_ms";
        private const string RESET_FILTER = "//*/a[text()='Reset Filter']";

        private const string FOLD_ALL = "//*/a[text()='Fold All']";
        private const string UNFOLD_ALL = "//*/a[text()='Unfold All']";
        private const string IS_UNFOLD = "//*/tbody/tr[2]/td[2]";

        private const string SELECT_FIRST_RECIPE = "//*/tbody/tr[2]/td[1]/div/input";
        private const string GENERATE_RAW_MATERIALS = "//*/a[text()='Generate raw materials']";
        private const string ADD_SELECTION_AS_FAVORITE = "//*/a[text()='Add selection as favorite']";
        private const string PLUS_BTN = "//*/button[text()='+']";
        private const string HAS_RAW_MATERIALS = "//*/tbody/tr[2]/td[1]";

        private const string SHOW_BUTTON = "//*/button[text()='Show']";
        private const string NO_PRODUCTION = "//*/p[text()='No production']";

        private const string TAB_FAVORITE = "hrefTabContentItemContainer";
        private const string FAVORITE_NAME = "//*/div[contains(@class,'clickable-favorite')]/span[text()='{0}']";
        private const string FAV_BUTTON_DELETE = "//*/button[@class='close']";
        private const string FAV_CONFIRM_DELETE = "dataConfirmOK";


        private const string ICONS = "//*[starts-with(@id, 'content_')]/div/table/tbody/tr/td[6]/a";
        private const string DATE_COLONNE = "//*[@id='list-item-with-action']/table/tbody/tr[position() > 1]/td[1]";
        private const string CLOSE_BUTTON = "//*[@id=\"modal-1\"]/div[3]/button";


        public enum FilterType
        {
            Site,
            StartDate,
            EndDate,
            ShowHiddenArticles,
            Customer
        }

        public void Filter(FilterType filterType, object value, bool premierOnglet = true)
        {
            switch (filterType)
            {
                case FilterType.Site:
                    var site = WaitForElementIsVisible(By.Id(FILTER_SITE));
                    site.SetValue(ControlType.DropDownList, value);
                    break;
                case FilterType.StartDate:
                    var startDate = WaitForElementIsVisible(By.Id(FILTER_START_DATE));
                    startDate.SetValue(ControlType.TextBox, ((DateTime)value).ToString("dd/MM/yyyy"));
                    startDate.SendKeys(Keys.Enter);
                    break;
                case FilterType.EndDate:
                    var endDate = WaitForElementIsVisible(By.Id(FILTER_END_DATE));
                    endDate.SetValue(ControlType.TextBox, ((DateTime)value).ToString("dd/MM/yyyy"));
                    endDate.SendKeys(Keys.Enter);
                    endDate.SendKeys(Keys.Tab);
                    break;
                case FilterType.ShowHiddenArticles:
                    var showHidden = WaitForElementExists(By.Id(FILTRE_SHOW_HIDDEN_ARTICLES));
                    showHidden.SetValue(ControlType.CheckBox, value);
                    break;
                case FilterType.Customer:
                    ComboBoxSelectById(new ComboBoxOptions(FILTRE_CUSTOMER, (string)value, !premierOnglet));
                    break;
            }
            if (premierOnglet) {
                //Favorite
                WaitForLoad();
            }
            else
            {
                WaitPageLoading();
            }
        }

        public void ResetFilter()
        {
            var resetFilter = WaitForElementIsVisible(By.XPath(RESET_FILTER));
            resetFilter.Click();
            WaitForLoad();
            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {

            }

        }

        public void Show()
        {
            var show = WaitForElementIsVisible(By.XPath(SHOW_BUTTON));
            show.Click();
            // changement d'onglet
            WaitPageLoading();
        }

        public bool HasEarlyProduction()
        {
            return !isElementVisible(By.XPath(NO_PRODUCTION));
        }

        public void FoldAll()
        {
            ShowExtendedMenu();
            var foldAll = WaitForElementIsVisible(By.XPath(FOLD_ALL));
            foldAll.Click();
            WaitPageLoading();
        }

        public void UnfoldAll()
        {
            ShowExtendedMenu();
            var unfoldAll = WaitForElementIsVisible(By.XPath(UNFOLD_ALL));
            unfoldAll.Click();
            WaitPageLoading();
        }

        public bool IsUnfold()
        {
            return isElementVisible(By.XPath(IS_UNFOLD));
        }

        public void GenerateRawMaterials()
        {
            UnfoldAll();
            SelectFirstRecipe();
            ShowPlusMenu();
            var generate = WaitForElementIsVisible(By.XPath(GENERATE_RAW_MATERIALS));
            generate.Click();
            WaitPageLoading();
        }

        public void SelectFirstRecipe()
        {
            var firstCheckBox = WaitForElementExists(By.XPath(SELECT_FIRST_RECIPE));
            firstCheckBox.SetValue(ControlType.CheckBox, true);
            WaitForLoad();
        }

        public override void ShowPlusMenu()
        {
            WaitForLoad();
            var plusButton = WaitForElementIsVisible(By.XPath(PLUS_BTN));
            var actions = new Actions(_webDriver);
            actions.MoveToElement(plusButton).Perform();
            WaitForLoad();
        }

        public bool HasRawMaterials()
        {
            return isElementVisible(By.XPath(HAS_RAW_MATERIALS));
        }

        public EarlyProductionFavoriteModal AddSelectionAsFavorite()
        {
            ShowPlusMenu();
            var fav = WaitForElementIsVisible(By.XPath(ADD_SELECTION_AS_FAVORITE));
            fav.Click();
            WaitForLoad();
            return new EarlyProductionFavoriteModal(_webDriver, _testContext);
        }

        public void CLickOnFavoritesTab() {
            var ongletFavorite = WaitForElementIsVisible(By.Id(TAB_FAVORITE));
            ongletFavorite.Click();
            WaitPageLoading();
        }

        public bool HasFavorite(string favoriteName)
        {
            return isElementExists(By.XPath(string.Format(FAVORITE_NAME, favoriteName)));
        }

        public void DeleteFavorite(string favoriteName)
        {
            var carre = WaitForElementIsVisible(By.XPath(string.Format(FAVORITE_NAME, favoriteName)));
            new Actions(_webDriver).MoveToElement(carre).Perform();
            var croix = SolveVisible(FAV_BUTTON_DELETE);
            croix.Click();
            var confirm = WaitForElementIsVisible(By.Id(FAV_CONFIRM_DELETE));
            confirm.Click();
        }

        public bool VerifFromDate(DateTime fromDate)
        {
            var iconRows = _webDriver.FindElements(By.XPath(ICONS));
            Actions actions = new Actions(_webDriver);

            foreach (var icon in iconRows)
            {
                WaitLoading();
                actions.MoveToElement(icon).Perform();
                //ouvrir Use case Modal
                icon.Click();
                WaitLoading();
                var elements = _webDriver.FindElements(By.XPath(DATE_COLONNE));

                foreach (var dateString in elements)
                {
                    DateTime date = DateTime.ParseExact(dateString.Text, "dd/MM/yyyy", CultureInfo.InvariantCulture);

                    if (date.Date < fromDate.Date)
                    {
                        return false;
                    }
                }
                //fermer Use case Modal
                var closeButton = _webDriver.FindElement(By.XPath(CLOSE_BUTTON));
                closeButton.Click();
            }
            return true;
        }
    }
}
