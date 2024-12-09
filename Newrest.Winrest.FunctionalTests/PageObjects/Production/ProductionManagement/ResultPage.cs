using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.Flight;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.OutputForm;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Production.ProductionManagement
{
    public class ResultPage : PageBase
    {
        protected IWebDriver WebDriver
        {
            get => WebDriverFactory.Driver;
        }

        public ResultPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //__________________________________Constantes_____________________________________

        // Général
        private const string EXTENDED_MENU = "//*[@id=\"div-body\"]/div/div[2]/div[1]/div/div[1]/div";
        private const string PLUS_MENU = "//*[@id=\"div-body\"]/div/div[2]/div[1]/div/div[2]/button";
        private const string GENERATE_OUTPUTFORM = "btnGenerateOutputForm";
        private const string GENERATE = "btn-submit-create-output-form";
        private const string GENERATE_SUPPLYORDER = "Generate Supply Order";
        private const string OF_NUMBER = "tb-new-outputform-number";
        private const string OF_PLACE_FROM = "drop-down-places-from";
        private const string OF_PLACE_COMMENT = "OutputForm_Comment";
        private const string OF_PLACE_TO = "drop-down-places-to";
        private const string SO_DELIVERY_LOCATION_SELECT = "SelectedSitePlaceId";
        private const string SO_NUMBER = "tb-new-supplyorder-number";
        private const string SO_CREATE_BUTTON = "btn-submit-form-create-supply-order";

        private const string FOLD_UNFOLD = "//*[@id=\"div-body\"]/div/div[2]/div[1]/div/div[1]/div/a[1]";
        private const string RAW_MATERIAL_BY_WORKSHOP = "RawMaterialsReport";
        private const string RAW_MATERIAL_BY_GROUP = "RawMaterialByGroupReport";
        private const string ASSEMBLY_REPORT = "AssemblyReport";
        private const string RECIPE_REPORT_V2 = "RecipeReportV2";
        private const string RECIPE_REPORT_DETAILLED_V2 = "RecipeReportDetailledV2";
        private const string HACCP_SANATIZATION = "HACCPSanitization";
        private const string HACCP_PLATING = "HACCPAssembly";
        private const string HACCP_PLATINGBYGROUP = "HACCPAssemblyGroupByFlight";
        private const string HACCP_PLATINGBYRECIPE = "HACCPAssemblyGroupByRecipe";
        private const string HACCP_TRAY_SETUP = "HACCPTraySetup";
        private const string HACCP_HOT_KITCHEN = "HACCPHotKitchen";
        private const string HACCP_SLICE = "HACCPSlice";
        private const string HACCP_THAWING = "HACCPThawing";
        private const string DATASHEET_BY_RECIPE = "DatasheetRecipeReport";
        private const string DATASHEET_BY_GUEST = "DatasheetGuestReport";
        private const string PRINT_BTN = "last";
        private const string HACCP_TRAY_SET_UP_MULTI_FLIGHT = "HACCPTraySetupMultiFlight";
        private const string NAV_TAB_LIST = "//*[@id=\"itemTabTab\"]/li[*]";
        private const string HACCP_TRAY_SET_UP_MULTI_FLIGHT_NO_GUEST = "HACCPTraySetupMultiFlightNoGuest";
        private const string HACCP_SLICE_KITCHEN = "HACCPSliceHotkitchen";
        private const string WITH_DATASHEET = "//*[@id=\"scDatasheet\"]";
        private const string DELIVERYLOCATION = "SelectedSitePlaceId";

        
        // Onglet
        private const string QTY_ADJUSTEMENT = "//*[@id=\"hrefTabContentItemContainer\"][text()='Quantity adjustments']";

        // Filtres
        private const string BACK = "//*[@id=\"item-filter-form\"]/div[1]/a";
        private const string FILTER_DATE_FROM = "ProdDateFrom";
        private const string FILTER_DATE_TO = "ProdDateTo";
        private const string SUBGROUP_FILTER = "SelectedItemSubGroups_ms";

        // Tableau
        private const string RAW_MAT_BY = "DdlSelectedView";
        private const string GROUP_BY_DROP = "DdlSelectedGroupBy";
        private const string GROUP_OPTION = "//*[@id=\"DdlSelectedGroupBy\"]/option[contains(@value,'{0}')]";

        private const string LINES = "//*[@id=\"tabContentItemContainer\"]/div[2]/div/div[*]";
        private const string LINES_NAME = "//*[@id=\"tabContentItemContainer\"]/div[2]/div/div[{0}]";
        private const string ARROW_LINE = "//*[@id=\"tabContentItemContainer\"]/div[2]/div/div[*]/div[1]/div/div[1]/span";
        private const string ITEM_BY_LINE = "//*[@id=\"tabContentItemContainer\"]/div[2]/div/div[{0}]/div[2]/div/table/tbody/tr[*]/td[1]";
        private const string QTY_BY_LINE = "//*[@id=\"tabContentItemContainer\"]/div[2]/div/div[{0}]/div[2]/div/table/tbody/tr[*]/td[3]";
        private const string PACKAGING_BY_LINE = "//*[@id=\"tabContentItemContainer\"]/div[2]/div/div[{0}]/div[2]/div/table/tbody/tr[*]/td[4]";
        private const string FOURNISSEUR_BY_LINE = "//*[@id=\"tabContentItemContainer\"]/div[2]/div/div[{0}]/div[2]/div/table/tbody/tr[*]/td[5]";
        private const string WORKSHOP_BY_LINE = "//*[@id=\"tabContentItemContainer\"]/div[2]/div/div[{0}]/div[2]/div/table/tbody/tr[*]/td[6]";

        private const string SUB_LINES = "//*[starts-with(@id,\"content_\")]/div/div[*]/div/div/div[2]/table/tbody/tr/td[1]";
        private const string SUB_LINES_NAME = "//*[starts-with(@id,\"content_\")]/div/div[{0}]/div/div/div[2]/table/tbody/tr/td[1]";
        private const string SUB_LINES_RECIPE = "//*[starts-with(@id,\"content_\")]/div/div[{0}]/div/div/div[2]/table/tbody/tr/td[3]";
        private const string ITEM_BY_SUBGROUP = "/html/body/div[2]/div/div[2]/div[3]/div/div[2]/div/div/div[2]/div/div[{0}]/div/table/tbody/tr[*]/td[1]";
        private const string SUB_LINES_WorkShop = "//*[starts-with(@id,\"content_\")]/div/div[{0}]/div/div/div[2]/table/tbody/tr";

        private const string DETAILS = "//*[starts-with(@id,\"content_\")]";

        //private const string DETAILSNAME = "production-inflightrawMaterials-itemusecase-itemname-0-1";
        private const string DETAILSNAME = "//*[starts-with(@id, \"production-inflightrawMaterials-itemusecase-itemname-\")]";
        private const string DETAILSNAME_PATCH = "//*[starts-with(@id,\"content_\")]/div/table/tbody/tr[*]/td[1]";


        private const string DETAILSQTE = "//*[starts-with(@id,\"content_\")]/div/table/tbody/tr[*]/td[2]";
        private const string DETAILSNAME_LINK = "//*[starts-with(@id,\"content_\")]/div/table/tbody/tr[*]/td[1][contains(text(),'{0}')]";
        private const string ALLUSECASE_BTN = "//*[starts-with(@id,\"content_\")]/div/table/tbody/tr[*]/td[6]/a[1]";

        private const string USECASES_BTN = "//*[starts-with(@id,\"content_\")]/div/table/tbody/tr[*]/td[1][contains(text(),'{0}')]/..//a[contains(@id,\"production-inflightrawMaterials-itemusecase\")]/span";
        private const string HIDE_ARTICLE = "//*[starts-with(@id,\"content_\")]/div/table/tbody/tr[*]/td[contains(text(), '{0}')]/../td[6]/a[2]/span";

        private const string GUESTTYPE = "//*[starts-with(@id,\"content_\")]/div/div/div/div/div[2]/table/tbody/tr[*]/td[2]";
        private const string WORKSHOP_NAME = "//*[starts-with(@id,\"content_\")]/div/table/tbody/tr[*]/td[1][contains(text(),'{0}')]/../td[6]";


        // -------------------------- Warnings --------------------------------------------
        private const string DELIVERY_PLACE_IS_REQUIRED = "//*[@id=\"form-create-supply-order\"]/div/div[3]/div[2]/div/div/span";
        private const string FROM_PLACE_IS_REQUIRED = "//*[@id=\"form-generate-output-form\"]/div/div[2]/div[1]/div/div/span";
        private const string TO_PLACE_IS_REQUIRED = "//*[@id=\"form-generate-output-form\"]/div/div[2]/div[2]/div/div/span";
        private const string HIDDEN_ITEMS_COUNTER = "hiddenItemsCounter";
        //Flight
        private const string FLIGHT_MENU = "TabFlights";
        private const string FL_FLIGHTS_LINK = "FlightLinkDashBoard";
        private const string FL_FLIGHTS = "FlightTabNav";
        private const string FL_LOADING_PLAN = "LoadingPlanTabNav";
        //__________________________________Variables______________________________________

        // Général

        [FindsBy(How = How.Id, Using = HACCP_SLICE_KITCHEN)]
        public IWebElement _haccpSlicekitchen;

        [FindsBy(How = How.XPath, Using = EXTENDED_MENU)]
        public IWebElement _extendedMenu;

        [FindsBy(How = How.XPath, Using = PLUS_MENU)]
        public IWebElement _plusMenu;

        [FindsBy(How = How.Id, Using = GENERATE_OUTPUTFORM)]
        public IWebElement _generateOutputForm;

        [FindsBy(How = How.Id, Using = OF_NUMBER)]
        public IWebElement _outputFormNumber;

        [FindsBy(How = How.Id, Using = OF_PLACE_FROM)]
        public IWebElement _outputFormPlaceFrom;

        [FindsBy(How = How.Id, Using = OF_PLACE_TO)]
        public IWebElement _outputFormPlaceTo;

        [FindsBy(How = How.XPath, Using = GENERATE)]
        public IWebElement _generate;

        [FindsBy(How = How.XPath, Using = GENERATE_SUPPLYORDER)]
        public IWebElement _generateSupplyOrder;

        [FindsBy(How = How.Id, Using = SO_NUMBER)]
        public IWebElement _supplyOrderNumber;

        [FindsBy(How = How.Id, Using = SO_CREATE_BUTTON)]
        public IWebElement _createSupplyOrder;

        [FindsBy(How = How.Id, Using = SO_DELIVERY_LOCATION_SELECT)]
        public IWebElement _supplyOrderDeliveryLocation;

        [FindsBy(How = How.XPath, Using = FOLD_UNFOLD)]
        public IWebElement _unfold;

        [FindsBy(How = How.Id, Using = RAW_MATERIAL_BY_WORKSHOP)]
        public IWebElement _rawMaterialByWorkshop;

        [FindsBy(How = How.Id, Using = RAW_MATERIAL_BY_GROUP)]
        public IWebElement _rawMaterialByGroup;

        [FindsBy(How = How.Id, Using = ASSEMBLY_REPORT)]
        public IWebElement _assembly;

        [FindsBy(How = How.Id, Using = RECIPE_REPORT_DETAILLED_V2)]
        public IWebElement _recipeReportV2;

        [FindsBy(How = How.Id, Using = RECIPE_REPORT_DETAILLED_V2)]
        public IWebElement _recipeReportDetailledV2;

        [FindsBy(How = How.Id, Using = HACCP_SANATIZATION)]
        public IWebElement _HACCPSanitization;

        [FindsBy(How = How.Id, Using = HACCP_PLATING)]
        public IWebElement _HACCPPlating;

        [FindsBy(How = How.Id, Using = HACCP_TRAY_SET_UP_MULTI_FLIGHT)]
        public IWebElement _haccptraySetupMultiFlight;

        [FindsBy(How = How.Id, Using = HACCP_TRAY_SET_UP_MULTI_FLIGHT_NO_GUEST)]
        public IWebElement _haccptraySetupMultiFlightnoguest;

        [FindsBy(How = How.Id, Using = HACCP_PLATINGBYGROUP)]
        public IWebElement _HACCPPlatingGroup;

        [FindsBy(How = How.Id, Using = HACCP_PLATINGBYRECIPE)]
        public IWebElement _HACCPPlatingRecipe;

        [FindsBy(How = How.Id, Using = HACCP_TRAY_SETUP)]
        public IWebElement _HACCPTraySetup;

        [FindsBy(How = How.Id, Using = HACCP_HOT_KITCHEN)]
        public IWebElement _HACCPHotKitchen;

        [FindsBy(How = How.Id, Using = HACCP_SLICE)]
        public IWebElement _HACCPSlice;

        [FindsBy(How = How.Id, Using = HACCP_THAWING)]
        public IWebElement _HACCPThawing;

        [FindsBy(How = How.Id, Using = DATASHEET_BY_RECIPE)]
        public IWebElement _datasheetByRecipe;

        [FindsBy(How = How.Id, Using = DATASHEET_BY_GUEST)]
        public IWebElement _datasheetByGuest;

        [FindsBy(How = How.Id, Using = PRINT_BTN)]
        public IWebElement _print;

        [FindsBy(How = How.Id, Using = OF_PLACE_COMMENT)]
        public IWebElement _outputFormComment ;


        [FindsBy(How = How.XPath, Using = DELIVERY_PLACE_IS_REQUIRED)]
        public IWebElement _delieveryPlaceIsRequired;

        [FindsBy(How = How.XPath, Using = FROM_PLACE_IS_REQUIRED)]
        public IWebElement _fromPlaceIsRequired;

        [FindsBy(How = How.XPath, Using = TO_PLACE_IS_REQUIRED)]
        public IWebElement _toPlaceIsRequired;

        [FindsBy(How = How.XPath, Using = WITH_DATASHEET)]
        public IWebElement _withDatasheet;
        // Onglets
        [FindsBy(How = How.XPath, Using = QTY_ADJUSTEMENT)]
        public IWebElement _qtyAdjustement;

        // Filtres
        [FindsBy(How = How.Id, Using = FILTER_DATE_FROM)]
        public IWebElement _dateFrom;

        [FindsBy(How = How.Id, Using = FILTER_DATE_TO)]
        public IWebElement _dateTo;

        [FindsBy(How = How.XPath, Using = BACK)]
        private IWebElement _back;

        // Tableau
        [FindsBy(How = How.Id, Using = RAW_MAT_BY)]
        public IWebElement _rawMatBy;

        [FindsBy(How = How.Id, Using = GROUP_BY_DROP)]
        private IWebElement _groupByDrop;

        [FindsBy(How = How.XPath, Using = LINES)]
        public IWebElement _lines;

        [FindsBy(How = How.XPath, Using = ARROW_LINE)]
        public IWebElement _arrowLine;

        [FindsBy(How = How.XPath, Using = SUB_LINES)]
        public IWebElement _subLines;

        //[FindsBy(How = How.Id, Using = DETAILSNAME)]
        [FindsBy(How = How.XPath, Using = DETAILSNAME)]
        public IWebElement _detail;

        //[FindsBy(How = How.Id, Using = DETAILSNAME)]
        [FindsBy(How = How.XPath, Using = DETAILSNAME)]
        public IWebElement _firstItemName;

        [FindsBy(How = How.XPath, Using = HIDE_ARTICLE)]
        public IWebElement _hideArticle;

        [FindsBy(How = How.XPath, Using = DETAILSQTE)]
        public IWebElement _firstItemQte;

        [FindsBy(How = How.Id, Using = FLIGHT_MENU)]
        private IWebElement _flights;

        [FindsBy(How = How.Id, Using = FL_LOADING_PLAN)]
        private IWebElement _flights_LoadingPlans;

        [FindsBy(How = How.Id, Using = FL_FLIGHTS_LINK)]
        private IWebElement _flights_Flights_link;

        [FindsBy(How = How.Id, Using = HIDDEN_ITEMS_COUNTER)]
        private IWebElement _hidden_items_counter;
        //___________________________________ Méthodes _________________________________________

        // Général

        public void UnfoldAll()
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)(IWebDriver)_webDriver;
            js.ExecuteScript("window.scrollTo(0,0)");

            Actions actions = new Actions(_webDriver);

            // Suppression de la classe CSS
            string classe = UnClassCSS();

            _extendedMenu = WaitForElementIsVisible(By.XPath(EXTENDED_MENU));
            actions.MoveToElement(_extendedMenu).Perform();

            _unfold = WaitForElementExists(By.XPath(FOLD_UNFOLD));
            _unfold.Click();
            WaitForLoad();

            Thread.Sleep(1000);

            // on remet la classe CSS
            AddCSS(classe);

            //move out
            actions.MoveByOffset(-100, -100).Perform();
        }

        public Boolean IsUnfoldAll()
        {
            System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> elements = _webDriver.FindElements(By.XPath("//*/div[contains(@class,'toggle-collapse-btn-left')]"));

            foreach (var elm in elements)
            {
                if (elm.GetAttribute("aria-expanded").Equals("false"))
                    return false;
            }

            return true;
        }

        public void FoldAll()
        {
            Actions actions = new Actions(_webDriver);

            // Suppression de la classe CSS
            string classe = UnClassCSS();

            _extendedMenu = WaitForElementIsVisible(By.XPath(EXTENDED_MENU));
            actions.MoveToElement(_extendedMenu).Perform();

            IJavaScriptExecutor js = (IJavaScriptExecutor)(IWebDriver)_webDriver;
            js.ExecuteScript("window.scrollTo(0,0)");

            var fold = WaitForElementExists(By.XPath(FOLD_UNFOLD));
            fold.Click();
            WaitForLoad();

            Thread.Sleep(1000);

            // on remet la classe CSS
            AddCSS(classe);
        }

        public Boolean IsFoldAll()
        {
            System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> elements = _webDriver.FindElements(By.XPath("//*/div[contains(@class,'toggle-collapse-btn-left')]"));

            foreach (var elm in elements)
            {
                if (elm.GetAttribute("aria-expanded").Equals("true"))
                    return false;
            }

            return true;
        }

        public string UnClassCSS()
        {
            _extendedMenu = WaitForElementExists(By.XPath(EXTENDED_MENU));
            string classe = _extendedMenu.GetAttribute("class");

            ((IJavaScriptExecutor)(IWebDriver)_webDriver).ExecuteScript(
               "arguments[0].removeAttribute('class','class')", _extendedMenu);

            WaitForLoad();

            return classe;
        }

        public void AddCSS(string classe)
        {
            _extendedMenu = WaitForElementExists(By.XPath(EXTENDED_MENU));

            ((IJavaScriptExecutor)(IWebDriver)_webDriver).ExecuteScript(
               "arguments[0].setAttribute('class', arguments[1])", _extendedMenu, classe);

            WaitForLoad();
        }

        public string CreateOutputForm(string placeFrom, string placeTo)
        {
            Actions actions = new Actions(_webDriver);

            _plusMenu = WaitForElementIsVisible(By.XPath(PLUS_MENU));
            actions.MoveToElement(_plusMenu).Perform();
            _plusMenu.Click();

            _generateOutputForm = WaitForElementIsVisible(By.Id(GENERATE_OUTPUTFORM));
            _generateOutputForm.Click();
            WaitForLoad();

            if (placeFrom != null)
            {
                _outputFormPlaceFrom = WaitForElementIsVisible(By.Id(OF_PLACE_FROM));
                _outputFormPlaceFrom.SetValue(ControlType.DropDownList, placeFrom);
            }

            if (placeTo != null)
            {
                _outputFormPlaceTo = WaitForElementIsVisible(By.Id(OF_PLACE_TO));
                _outputFormPlaceTo.SetValue(ControlType.DropDownList, placeTo);
            }

            _outputFormNumber = WaitForElementIsVisible(By.Id(OF_NUMBER));

            return _outputFormNumber.GetAttribute("value");
        }

        public OutputFormItem GenerateOutputForm(bool modalFail = false)
        {
            _generate = WaitForElementIsVisible(By.Id(GENERATE));
            _generate.Click();
            WaitPageLoading();
            WaitForLoad();

            if (!modalFail)
            {
                //Results are opened in a new tab, switch the driver to the newly created one
                var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
                wait.Until((driver) => driver.WindowHandles.Count > 1);
                _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);
            }

            return new OutputFormItem(_webDriver, _testContext);
        }
        public string CreateSupplyOrder(string deliveryLoc)
        {
            Actions actions = new Actions(_webDriver);

            _plusMenu = WaitForElementIsVisible(By.XPath(PLUS_MENU));
            actions.MoveToElement(_plusMenu).Perform();
            _plusMenu.Click();

            _generateSupplyOrder = _webDriver.FindElement(By.XPath("//a[text() = 'Generate Supply Order']"));
            _generateSupplyOrder.Click();
            WaitForLoad();

            if (deliveryLoc != null)
            {
                _supplyOrderDeliveryLocation = WaitForElementIsVisible(By.Id(SO_DELIVERY_LOCATION_SELECT));
                _supplyOrderDeliveryLocation.SetValue(ControlType.DropDownList, deliveryLoc);
            }

            _supplyOrderNumber = WaitForElementIsVisible(By.Id(SO_NUMBER));

            return _supplyOrderNumber.GetAttribute("value");
        }

        public SupplyOrderItem GenerateSupplyOrder(bool modalFail = false)
        {
            _createSupplyOrder = WaitForElementIsVisible(By.Id(SO_CREATE_BUTTON));
            _createSupplyOrder.Click();
            WaitPageLoading();
            WaitForLoad();

            if (!modalFail)
            {
                //Results are opened in a new tab, switch the driver to the newly created one
                var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
                wait.Until((driver) => driver.WindowHandles.Count > 1);
                _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);
            }

            return new SupplyOrderItem(_webDriver, _testContext);
        }

        public void ClickPrefillQuantitiesCheckbox()
        {
            var prefillCheckbox = WaitForElementIsVisible(By.XPath("//*[@id=\"form-create-supply-order\"]/div/div[4]/div[3]/div/div/div"));

            if (!prefillCheckbox.Selected)
            {
                prefillCheckbox.Click();
            }

            WaitForLoad();
        }

        // Onglets
        public QuantityAdjustmentsPage GoToQtyAdjustementPage()
        {
            _qtyAdjustement = WaitForElementIsVisible(By.XPath(QTY_ADJUSTEMENT));
            _qtyAdjustement.Click();
            WaitPageLoading();

            return new QuantityAdjustmentsPage(_webDriver, _testContext);
        }

        // Filtres
        public FilterAndFavoritesPage Back()
        {
            _back = WaitForElementIsVisible(By.XPath(BACK));
            _back.Click();
            WaitForLoad();

            return new FilterAndFavoritesPage(_webDriver, _testContext);
        }

        public bool VerifyDateFrom(DateTime value)
        {
            var dateFrom = _webDriver.FindElements(By.XPath(LINES)).FirstOrDefault();

            if (dateFrom != null)
            {
                var dateFormat = _dateFrom.GetAttribute("data-date-format");
                CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? ci = new CultureInfo("fr-FR") : ci = new CultureInfo("en-US");

                DateTime date = DateTime.Parse(dateFrom.Text, ci).Date;

                if (value == date)
                    return true;
            }

            return false;
        }

        public bool VerifyDateTo(DateTime value)
        {
            var dateTo = _webDriver.FindElements(By.XPath(LINES)).LastOrDefault();

            if (dateTo != null)
            {
                var dateFormat = _dateFrom.GetAttribute("data-date-format");
                CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? ci = new CultureInfo("fr-FR") : ci = new CultureInfo("en-US");

                DateTime date = DateTime.Parse(dateTo.Text, ci).Date;

                if (value == date)
                    return true;
            }

            return false;
        }


        // Tableau
        public bool GetRawMatByPage(string page)
        {
            _rawMatBy = WaitForElementIsVisible(By.Id(RAW_MAT_BY));
            _rawMatBy.Click();

            var element = WaitForElementIsVisible(By.XPath("//option[contains(@value,'" + page + "')]"));
            try
            {
                string check = element.GetAttribute("selected");
                return check.Equals("true");
            }
            catch
            {
                return false;
            }

        }

        public bool GetGroupByPage(string page)
        {
            _groupByDrop = WaitForElementIsVisible(By.Id(GROUP_BY_DROP));
            _groupByDrop.Click();

            var element = WaitForElementIsVisible(By.XPath(String.Format(GROUP_OPTION, page)));
            return element.GetAttribute("selected").Equals("true");
        }

        public void FilterRawBy(string value)
        {
            _rawMatBy = WaitForElementIsVisible(By.Id(RAW_MAT_BY));
            _rawMatBy.Click();
            var element = WaitForElementIsVisible(By.XPath("//option[contains(@value,'" + value + "')]"));

            if (element.GetAttribute("selected") == null)
            {
                _rawMatBy.SetValue(ControlType.DropDownList, element.Text);
                _rawMatBy.Click();

                WaitPageLoading();
                Thread.Sleep(1500);
            }
            else
            {
                _rawMatBy.Click();
            }
        }

        public void GroupbyRawMatGroup(string value)
        {
            _groupByDrop = WaitForElementIsVisible(By.Id(GROUP_BY_DROP));
            _groupByDrop.Click();
            var element = WaitForElementIsVisible(By.XPath("//option[contains(@value,'" + value + "')]"));

            if (element.GetAttribute("selected") == null)
            {
                _groupByDrop.SetValue(ControlType.DropDownList, element.Text);
                _groupByDrop.Click();

                WaitPageLoading();
                Thread.Sleep(1500);
            }
            else
            {
                _groupByDrop.Click();
            }
        }

        public void GroupBy(string value)
        {
            _groupByDrop = WaitForElementIsVisible(By.Id(GROUP_BY_DROP));
            _groupByDrop.SetValue(ControlType.DropDownList, value);
            WaitPageLoading();
        }

        public int CountResults()
        {
            WaitLoading();
            return _webDriver.FindElements(By.XPath(LINES)).Count;
        }

        public string GetFirstItemGroup()
        {
            if (isElementVisible(By.XPath(LINES)))
            {
                _lines = WaitForElementIsVisible(By.XPath(LINES));
                int parenthese = _lines.Text.IndexOf("(");
                var result = _lines.Text.Substring(0, parenthese).Trim();
                return result;
            }
            else
            {
                return "";
            }
        }

        public int CountSubResults()
        {
            return _webDriver.FindElements(By.XPath(SUB_LINES)).Count;
        }
        public string GetSubGroupName()
        {
            if (isElementVisible(By.XPath(SUB_LINES)))
            {
                _subLines = WaitForElementIsVisible(By.XPath(SUB_LINES));
                return _subLines.Text.Trim();
            }
            else
            {
                return "";
            }
        }

        public void ShowDetail()
        {

            _arrowLine = WaitForElementIsVisible(By.XPath(ARROW_LINE));
            _arrowLine.Click();

            Thread.Sleep(1000);
        }

        public bool IsDetailVisible()
        {
            var element = _webDriver.FindElement(By.XPath("//*/div[contains(@class,'toggle-collapse-btn-left')]"));
            if (element.GetAttribute("class").Contains("collapsed"))
                return false;

            return true;
        }

        public string GetFirstItemName()
        {
            if (IsDev())
                _firstItemName = WaitForElementIsVisible(By.XPath(DETAILSNAME));
            else
                _firstItemName = WaitForElementIsVisible(By.XPath(DETAILSNAME_PATCH));

            return _firstItemName.Text;
        }
        public string GetFirstItemQty()
        {
            _firstItemQte = WaitForElementIsVisible(By.XPath(DETAILSQTE));
            return _firstItemQte.Text;
        }

        public Dictionary<string, List<string>> GetGroupAndItemsAssociated()
        {
            Dictionary<string, List<string>> mapGpeItem = new Dictionary<string, List<string>>();
            int groups = _webDriver.FindElements(By.XPath(LINES)).Count;
            for (int i = 1; i <= groups; i++)
            {
                IWebElement group = WaitForElementIsVisible(By.XPath(String.Format(LINES_NAME, i)));
                string text = group.Text.Substring(0, group.Text.IndexOf("(")).Trim();
                mapGpeItem.Add(text, new List<string>());
                System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> items = _webDriver.FindElements(By.XPath(String.Format(ITEM_BY_LINE, i)));
                foreach (IWebElement item in items)
                    mapGpeItem[text].Add(item.Text);
            }
            return mapGpeItem;
        }
        public void SearchData()
        {
            DateTime datefrom = DateTime.Now.AddDays(-1);
            while (
           !isElementVisible(By.XPath("//*[@id=\"tabContentItemContainer\"]/div[2]"))
           )
            {
                Filter(ResultPage.FilterType.DateFrom, datefrom);
                datefrom = datefrom.AddDays(-1);
            }
        }

        public bool GetGroupAndRecipeAssociated(string type)
        {
            Actions action = new Actions(_webDriver);

            var groups = _webDriver.FindElements(By.XPath(LINES)).Count;

            for (var i = 1; i <= groups; i++)
            {
                var group = WaitForElementIsVisible(By.XPath(String.Format(LINES_NAME, i)));

                string value = group.Text.Substring(0, group.Text.IndexOf("(")).Trim();

                System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> items = _webDriver.FindElements(By.XPath(String.Format(ITEM_BY_LINE, i)));

                int compteur = 1;
                foreach (var item in items)
                {
                    action.MoveToElement(item).Perform();
                    item.Click();

                    var useCaseModalPage = GoToUseCaseModalPage(item.Text, compteur);

                    if (type.Equals("Workshop") && !useCaseModalPage.IsWorkShop(value))
                    {
                        useCaseModalPage.CloseModal();
                        return false;

                    }
                    else if (type.Equals("RecipeType") && !useCaseModalPage.IsRecipeType(value))
                    {
                        useCaseModalPage.CloseModal();
                        return false;
                    }

                    useCaseModalPage.CloseModal();
                    compteur++;
                }
            }

            return true;
        }

        public bool GetSubgroupAndRecipeAssociated()
        {
            Actions action = new Actions(_webDriver);

            var subgroups = _webDriver.FindElements(By.XPath(SUB_LINES)).Count;

            for (var i = 1; i <= subgroups * 2; i++)
            {
                var subgroup = WaitForElementIsVisible(By.XPath(String.Format(SUB_LINES_NAME, i)));

                string workshop = subgroup.Text.Trim();

                System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> items = _webDriver.FindElements(By.XPath(String.Format(ITEM_BY_SUBGROUP, i + 1)));

                int compteur = 1;
                foreach (var item in items)
                {

                    action.MoveToElement(item).Perform();
                    item.Click();

                    var useCaseModalPage = GoToUseCaseModalPage(item.Text, compteur);

                   if (!useCaseModalPage.IsWorkshop(workshop))
                        return false;

                    useCaseModalPage.CloseModal();
                    compteur++;
                }

                i++;
            }

            return true;
        }

        public Dictionary<string, List<string>> GetRecipeAndItemsAssociated()
        {
            Dictionary<string, List<string>> mapRecipeItems = new Dictionary<string, List<string>>();

            var subgroups = _webDriver.FindElements(By.XPath(SUB_LINES)).Count;

            for (var i = 1; i <= subgroups * 2; i++)
            {
                var subgroup = WaitForElementIsVisible(By.XPath(String.Format(SUB_LINES_RECIPE, i)));

                string recipe = subgroup.Text.Substring(subgroup.Text.IndexOf(":") + 1).Trim();
                mapRecipeItems.Add(recipe, new List<string>());

                System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> items = _webDriver.FindElements(By.XPath(String.Format(ITEM_BY_SUBGROUP, i + 1)));

                foreach (var item in items)
                {
                    mapRecipeItems[recipe].Add(item.Text.Trim());
                }

                i++;
            }

            return mapRecipeItems;
        }
        public Dictionary<string, List<string>> GetWorkShopAndItemsAssociated()
        {
            Dictionary<string, List<string>> mapWorkShopItems = new Dictionary<string, List<string>>();

            var subgroups = _webDriver.FindElements(By.XPath(SUB_LINES)).Count;

            for (var i = 1; i <= subgroups * 2; i++)
            {
                var subgroup = WaitForElementIsVisible(By.XPath(String.Format(SUB_LINES_WorkShop, i)));

                string recipe = subgroup.Text.Substring(subgroup.Text.IndexOf(":") + 1).Trim();
                mapWorkShopItems.Add(recipe, new List<string>());

                System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> items = _webDriver.FindElements(By.XPath(String.Format(ITEM_BY_SUBGROUP, i + 1)));

                foreach (var item in items)
                {
                    mapWorkShopItems[recipe].Add(item.Text.Trim());
                }

                i++;
            }

            return mapWorkShopItems ;
        }

        public ResultModal ShowUseCase()
        {
            Actions action = new Actions(_webDriver);
            IEnumerable<IWebElement> searchList = new List<IWebElement>();
            if (IsDev())
            {
                WaitForElementToBeClickable(By.XPath(DETAILSNAME));
                searchList = _webDriver.FindElements(By.XPath(DETAILSNAME));
            }
            else
            {
                WaitForElementToBeClickable(By.CssSelector("div > table > tbody > tr.tr-groupable > td:nth-child(1)"));
                searchList = _webDriver.FindElements(By.CssSelector("div > table > tbody > tr.tr-groupable > td:nth-child(1)"));
            }

            action.MoveToElement(searchList.FirstOrDefault()).Perform();
            searchList.FirstOrDefault().Click();

            return GoToUseCaseModalPage(searchList.FirstOrDefault().Text, 0);
        }

        public List<string> GetCustomerNames()
        {
            List<string> customerNames = new List<string>();

            System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> customers = _webDriver.FindElements(By.XPath(SUB_LINES));

            foreach (var customer in customers)
            {
                customerNames.Add(customer.Text.Trim());
            }

            return customerNames;
        }

        public int CountItems()
        {
            WaitForLoad();
            IEnumerable<IWebElement> searchList = new List<IWebElement>();
            if (IsDev())
            {
                WaitForElementToBeClickable(By.XPath(DETAILSNAME));
                searchList = _webDriver.FindElements(By.XPath(DETAILSNAME));
            }
            else
            {
                searchList = _webDriver.FindElements(By.XPath(DETAILSNAME_PATCH));
                //WaitForElementToBeClickable(By.CssSelector("div > table > tbody > tr.tr-groupable > td:nth-child(1)"));
                //searchList = _webDriver.FindElements(By.CssSelector("div > table > tbody > tr.tr-groupable > td:nth-child(1)"));
            }
            return searchList.Count();
        }

        public List<String> GetItemNames()
        {
            List<String> names = new List<String>();
            IEnumerable<IWebElement> searchList = new List<IWebElement>();
            if (IsDev())
            {
                WaitForElementToBeClickable(By.XPath(DETAILSNAME));
                searchList = _webDriver.FindElements(By.XPath(DETAILSNAME));
            }
            else
            {
                if(isElementExists(By.CssSelector("div > table > tbody > tr.tr-groupable > td:nth-child(1)")))
                { 
                WaitForElementToBeClickable(By.CssSelector("div > table > tbody > tr.tr-groupable > td:nth-child(1)"));
                searchList = _webDriver.FindElements(By.CssSelector("div > table > tbody > tr.tr-groupable > td:nth-child(1)"));
                }
            }

            if (!searchList.Any()) return new List<string> { "No production" }; 
        
            foreach (IWebElement elm in searchList)
            names.Add(elm.Text);
            return names;
        }

        public bool VerifyGuestType(string guestType)
        {
            System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> elements = _webDriver.FindElements(By.XPath(GUESTTYPE));

            if (elements.Count() == 0)
                return false;

            foreach (var elm in elements)
            {
                if (!elm.Text.Contains(guestType))
                    return false;
            }

            return true;
        }

        public bool VerifyCategorie(string categorieService)
        {
            Actions action = new Actions(_webDriver);
            IEnumerable<IWebElement> searchList = new List<IWebElement>();
            if (IsDev())
            {
                WaitForElementToBeClickable(By.XPath(DETAILSNAME));
                searchList = _webDriver.FindElements(By.XPath(DETAILSNAME));
            }
            else
            {
                WaitForElementToBeClickable(By.CssSelector("div > table > tbody > tr.tr-groupable > td:nth-child(1)"));
                searchList = _webDriver.FindElements(By.CssSelector("div > table > tbody > tr.tr-groupable > td:nth-child(1)"));
            }

            int compteur = 1;
            foreach (var elm in searchList)
            {
                action.MoveToElement(elm).Perform();
                elm.Click();

                var useCaseModalPage = GoToUseCaseModalPage(elm.Text, compteur);

                //WaitForLoad();
                if (!useCaseModalPage.IsServiceCategorie(categorieService))
                    return false;
                useCaseModalPage.CloseModal();
                compteur++;
            }
            return true;
        }

        public bool VerifyRecipeType(string recipeType)
        {
            Actions action = new Actions(_webDriver);
            IEnumerable<IWebElement> searchList = new List<IWebElement>();
            if (IsDev())
            {
                WaitForElementToBeClickable(By.XPath(DETAILSNAME));
                searchList = _webDriver.FindElements(By.XPath(DETAILSNAME));
            }
            else
            {
                WaitForElementToBeClickable(By.CssSelector("div > table > tbody > tr.tr-groupable > td:nth-child(1)"));
                searchList = _webDriver.FindElements(By.CssSelector("div > table > tbody > tr.tr-groupable > td:nth-child(1)"));
    
            }

            int compteur = 1;
            foreach (var elm in searchList)
            {
                action.MoveToElement(elm).Perform();
                elm.Click();

                var useCaseModalPage = GoToUseCaseModalPage(elm.Text.Trim(), compteur);
             
                if (!useCaseModalPage.IsRecipeType(recipeType))
                    return false;
                useCaseModalPage.CloseModal();
                compteur++;

            }
            return true;
        }

        public DateTime getDateFilterFrom()
        {
            var dateFormat = _dateFrom.GetAttribute("value");
            DateTime date;
            if (DateTime.TryParseExact(dateFormat, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
            {
                return date.Date;
            }
            else
            {
                // Handle the error or fallback to a default date.
                // For example:
                throw new FormatException("Invalid date format");
            }

        }


        public DateTime getDateFilterTo()
        {
            var dateFormat = _dateTo.GetAttribute("value");
            DateTime date;
            if (DateTime.TryParseExact(dateFormat, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
            {
                return date.Date;
            }
            else
            {
                // Handle the error or fallback to a default date.
                // For example:
                throw new FormatException("Invalid date format");
            }
        }

        public string GetFilterValue(string Id)
        {

            IWebElement selectElement = _webDriver.FindElement(By.Id(Id));
            SelectElement select = new SelectElement(selectElement);
            IWebElement selectedOption = select.SelectedOption;

            return selectedOption.Text;
        }

        public List<string> GetSelectedCustomers()
        {
            var listCust = new List<string>();
            var selectElement = _webDriver.FindElement(By.Id("SelectedCustomersTrueValue"));

            IJavaScriptExecutor js = (IJavaScriptExecutor)_webDriver;
            var attempts = 3;

            while (attempts > 0)
            {
                try
                {
                    System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> options = selectElement.FindElements(By.TagName("option"));

                    var selectedOptions = options.Where(option => option.Selected).ToList();

                    foreach (var option in selectedOptions)
                    {
                        string optionText = (string)js.ExecuteScript("return arguments[0].text;", option);
                        Console.WriteLine($"Selected option value: {option.GetAttribute("value")}, Text: {optionText}");
                        listCust.Add(optionText);
                    }
                    break;
                }
                catch (OpenQA.Selenium.StaleElementReferenceException)
                {
                    attempts--;

                    if (attempts == 0)
                    {
                        throw;
                    }

                    selectElement = _webDriver.FindElement(By.Id("SelectedCustomersTrueValue"));
                }
            }

            return listCust;
        }


        public List<string> GetSelectedGuestType()
        {
            var listGuesTyp = new List<string>();

            var selectElement = _webDriver.FindElement(By.Id("SelectedGuestTypesTrueValue"));

            WaitPageLoading();
            System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> options = selectElement.FindElements(By.TagName("option"));

            IJavaScriptExecutor js = (IJavaScriptExecutor)_webDriver;
            var selectedOptions = options.Where(option => option.Selected).ToList();

            foreach (var option in selectedOptions)
            {
                string optionText = (string)js.ExecuteScript("return arguments[0].text;", option);
                Console.WriteLine($"Selected option value: {option.GetAttribute("value")}, Text: {optionText}");
                listGuesTyp.Add(optionText);
            }

            return listGuesTyp;
        }

        public HashSet<string> GetFlights()
        {
            Actions action = new Actions(_webDriver);
            HashSet<string> flightList = new HashSet<string>();

            IEnumerable<IWebElement> searchList = new List<IWebElement>();
            if (IsDev())
            {
                WaitForElementToBeClickable(By.XPath(DETAILSNAME));
                searchList = _webDriver.FindElements(By.XPath(DETAILSNAME));
            }
            else
            {
                WaitForElementToBeClickable(By.CssSelector("div > table > tbody > tr.tr-groupable > td:nth-child(1)"));
                searchList = _webDriver.FindElements(By.CssSelector("div > table > tbody > tr.tr-groupable > td:nth-child(1)"));
            }
            int compteur = 1;
            foreach (var elm in searchList)
            {
                action.MoveToElement(elm).Perform();
                elm.Click();

                var useCaseModalPage = GoToUseCaseModalPage(elm.Text, compteur);

                flightList.UnionWith(useCaseModalPage.GetFlights());
                useCaseModalPage.CloseModal();
                compteur++;
            }
            return flightList;
        }

        public bool VerifyCustomerRawMat(string customer)
        {
            Actions action = new Actions(_webDriver);
            IEnumerable<IWebElement> searchList = new List<IWebElement>();
            if (IsDev())
            {
                WaitForElementToBeClickable(By.XPath(DETAILSNAME));
                searchList = _webDriver.FindElements(By.XPath(DETAILSNAME));
            }
            else
            {
                WaitForElementToBeClickable(By.CssSelector("div > table > tbody > tr.tr-groupable > td:nth-child(1)"));
                searchList = _webDriver.FindElements(By.CssSelector("div > table > tbody > tr.tr-groupable > td:nth-child(1)"));
            }

            int compteur = 1;

            foreach (var elm in searchList)
            {
                action.MoveToElement(elm).Perform();
                elm.Click();
                WaitForLoad();
                var useCaseModalPage = GoToUseCaseModalPage(elm.Text, compteur);

                if (!useCaseModalPage.IsCustomer(customer))
                    return false;
                useCaseModalPage.CloseModal();
                compteur++;

            }
            return true;
        }

        public bool VerifyServiceAndCustomer(string service, string customer)
        {
            Actions action = new Actions(_webDriver);
            IEnumerable<IWebElement> searchList = new List<IWebElement>();
            if (IsDev())
            {
                WaitForElementToBeClickable(By.XPath(DETAILSNAME));
                searchList = _webDriver.FindElements(By.XPath(DETAILSNAME));
            }
            else
            {
                searchList = _webDriver.FindElements(By.XPath(DETAILSNAME_PATCH));
                //WaitForElementToBeClickable(By.CssSelector("div > table > tbody > tr.tr-groupable > td:nth-child(1)"));
                //searchList = _webDriver.FindElements(By.CssSelector("div > table > tbody > tr.tr-groupable > td:nth-child(1)"));
            }
            int compteur = 1;
            foreach (var elm in searchList)
            {
                action.MoveToElement(elm).Perform();
                elm.Click();

                var useCaseModalPage = GoToUseCaseModalPage(elm.Text, compteur);

                if (!useCaseModalPage.IsService(service))
                    return false;
                if (!useCaseModalPage.IsCustomer(customer))
                    return false;
                useCaseModalPage.CloseModal();
                compteur++;
            }
            return true;
        }

        public ResultModal GoToUseCaseModalPage(string itemName, int rowNumber)
        {
            IWebElement btn;
            Actions action = new Actions(_webDriver);

            IWebElement _itemName = WaitForElementIsVisible(By.XPath(string.Format(DETAILSNAME_LINK, itemName)));
            _itemName.Click();
            if (IsDev())
            {
                btn = WaitForElementIsVisible(By.Id(string.Format("production-inflightrawMaterials-itemusecase-{0}-1", rowNumber)));
                btn.Click();
            }
            else
            {
                if (isElementVisible(By.XPath(string.Format("/html/body/div[2]/div/div[2]/div[3]/div/div[2]/div/div/div[2]/div/div[2]/div/table/tbody/tr[2]/td[6]/a[1]"))))
                {
                    btn = WaitForElementIsVisible(By.XPath(string.Format("/html/body/div[2]/div/div[2]/div[3]/div/div[2]/div/div/div[2]/div/div[2]/div/table/tbody/tr[2]/td[6]/a[1]")));

                    btn.Click();
                }
                else
                {

                    var x = WaitForElementExists(By.XPath($"//*[starts-with(@id,'content')]/div/table/tbody/tr[*]/td[contains(text(),'{itemName}')]/..//a[contains(@class,'btn-use-case')]"));
                    action.MoveToElement(x).Click().Perform();
                }
            }
            WaitForLoad();

            return new ResultModal(_webDriver, _testContext);
        }
        public ResultModal GoToUseCaseWorkshopModalPage(string itemName, int rowNumber)
        {
            IWebElement btn;
            Actions action = new Actions(_webDriver);

            IWebElement _itemName = WaitForElementIsVisible(By.XPath(string.Format(DETAILSNAME_LINK, itemName)));
            _itemName.Click();
            if (IsDev())
            {
                btn = WaitForElementIsVisible(By.Id(string.Format("production-inflightrawMaterials-itemusecase-{0}-1", rowNumber)));
                btn.Click();
            }
            else
            {
                var x = WaitForElementExists(By.XPath($"//*[starts-with(@id,\"content\")]/div/table/tbody/tr[*]/td[contains(text(),'{itemName}')]/..//a[contains(@class,'btn-use-case')]"));
                action.MoveToElement(x).Click().Perform();
            }

            WaitForLoad();

            return new ResultModal(_webDriver, _testContext);
        }
        public ResultModal GoToUseCaseRecipeModalPage(string itemName, int rowNumber)
        {
            IWebElement btn;
            Actions action = new Actions(_webDriver);

            IWebElement _itemName = WaitForElementIsVisible(By.XPath(string.Format(DETAILSNAME_LINK, itemName)));
            _itemName.Click();
            if (IsDev())
            {
                btn = WaitForElementIsVisible(By.Id(string.Format("production-inflightrawMaterials-itemusecase-{0}-1", rowNumber)));

                btn.Click();
            }
            else
            {
                var x = WaitForElementExists(By.XPath($"//*[starts-with(@id,'content')]/div/table/tbody/tr[*]/td[contains(text(),'{itemName}')]/..//a[contains(@class,'btn-use-case')]"));
                action.MoveToElement(x).Click().Perform();
            }
            WaitForLoad();

            return new ResultModal(_webDriver, _testContext);
        }
        public ResultModal GoToUseCaseGroupModalPage(string itemName, int rowNumber)
        {
            IWebElement btn;
            Actions action = new Actions(_webDriver);

            IWebElement _itemName = WaitForElementIsVisible(By.XPath(string.Format(DETAILSNAME_LINK, itemName)));
            _itemName.Click();
            if (IsDev())
            {
                btn = WaitForElementIsVisible(By.Id(string.Format("production-inflightrawMaterials-itemusecase-{0}-1", rowNumber)));

                btn.Click();
            }
            else
            {
                var x = WaitForElementExists(By.XPath($"//*[starts-with(@id,\"content\")]/div/table/tbody/tr[*]/td[contains(text(),'{itemName}')]/..//a[contains(@class,'btn-use-case')]"));
                action.MoveToElement(x).Click().Perform();
            }
            WaitForLoad();

            return new ResultModal(_webDriver, _testContext);
        }
        public ResultModal GoToUseCaseSupplierModalPage(string itemName, int rowNumber)
        {
            IWebElement btn;
            Actions action = new Actions(_webDriver);

            IWebElement _itemName = WaitForElementIsVisible(By.XPath(string.Format(DETAILSNAME_LINK, itemName)));
            _itemName.Click();
            if (IsDev())
            {
                btn = WaitForElementIsVisible(By.Id(string.Format("production-inflightrawMaterials-itemusecase-{0}-1", rowNumber)));
                btn.Click();
            }
            else
            {
                var x = WaitForElementExists(By.XPath($"//*[starts-with(@id,'content')]/div/table/tbody/tr[*]/td[contains(text(),'{itemName}')]/..//a[contains(@class,'btn-use-case')]"));
                action.MoveToElement(x).Click().Perform();
            }
            //btn.Click();
            WaitForLoad();

            return new ResultModal(_webDriver, _testContext);
        }



        public bool VerifyService(string service)
        {
            Actions action = new Actions(_webDriver);
            IEnumerable<IWebElement> searchList = new List<IWebElement>();
            if (IsDev())
            {
                WaitForElementToBeClickable(By.XPath(DETAILSNAME));
                searchList = _webDriver.FindElements(By.XPath(DETAILSNAME));
            }
            else
            {
                WaitForElementToBeClickable(By.CssSelector("div > table > tbody > tr.tr-groupable > td:nth-child(1)"));
                searchList = _webDriver.FindElements(By.CssSelector("div > table > tbody > tr.tr-groupable > td:nth-child(1)"));
            }

            int compteur = 1;
            foreach (var elm in searchList)
            {
                action.MoveToElement(elm).Perform();
                elm.Click();

                var useCaseModalPage = GoToUseCaseModalPage(elm.Text, compteur);

                if (!useCaseModalPage.IsService(service))
                    return false;
                useCaseModalPage.CloseModal();
                compteur++;
            }
            return true;
        }

        public bool VerifyWorkshop(string workshop,bool isgroup = false)
        {
            Actions action = new Actions(_webDriver);
            IEnumerable<IWebElement> searchList = new List<IWebElement>();

            if (IsDev())
            {
                WaitForElementToBeClickable(By.XPath(DETAILSNAME));
                searchList = _webDriver.FindElements(By.XPath(DETAILSNAME));
            }
            else
            {
                WaitPageLoading();
                WaitForElementToBeClickable(By.CssSelector("div > table > tbody > tr.tr-groupable > td:nth-child(1)"));
                searchList = _webDriver.FindElements(By.CssSelector("div > table > tbody > tr.tr-groupable > td:nth-child(1)"));
            }

            int compteur = 1;
            foreach (var elm in searchList)
            {
                action.MoveToElement(elm).Perform();
                elm.Click();
                
                var workshopItem = WaitForElementIsVisible(By.XPath(String.Format(WORKSHOP_NAME, elm.Text)));
                if (!workshopItem.Text.Equals(workshop))
                {
                    workshopItem = WaitForElementIsVisible(By.XPath(String.Format("//*[starts-with(@id,\"content_\")]/div/table/tbody/tr[*]/td[1][contains(text(),'{0}')]/../td[5]", elm.Text)));


                    if (!workshopItem.Text.Equals(workshop))
                        return false;

                }

                var useCaseModalPage = GoToUseCaseModalPage(elm.Text, compteur);

                if (!useCaseModalPage.IsWorkShop(workshop))
                {
                    useCaseModalPage.CloseModal();
                    return false;
                }
                useCaseModalPage.CloseModal();
                compteur++;

            }
            return true;
        }

        public List<String> GetCookingModes(CookingModeType cookingModeType = CookingModeType.NONE)
        {
            Actions action = new Actions(_webDriver);
            HashSet<String> cookingModes = new HashSet<string>();
            IEnumerable<IWebElement> searchList = new List<IWebElement>();
            if (IsDev())
            {
                WaitForElementToBeClickable(By.XPath(DETAILSNAME));
                searchList = _webDriver.FindElements(By.XPath(DETAILSNAME));
            }
            else
            {
                WaitForElementToBeClickable(By.CssSelector("div > table > tbody > tr.tr-groupable > td:nth-child(1)"));
                searchList = _webDriver.FindElements(By.CssSelector("div > table > tbody > tr.tr-groupable > td:nth-child(1)"));
            }
            List<String> searchListFlat = new List<String>();
            foreach (IWebElement elm in searchList) searchListFlat.Add(elm.Text.Trim());
            int compteur = 1;
            foreach (string elm in searchListFlat)
            {
                IWebElement lien = WaitForElementIsVisible(By.XPath(String.Format(DETAILSNAME_LINK, elm)));
                lien.Click();
                ResultModal useCaseModalPage; // = GoToUseCaseModalPage(elm, compteur);
                switch (cookingModeType)
                {
                    case CookingModeType.COOKING_MODE_WORKSHOP:
                        useCaseModalPage = GoToUseCaseWorkshopModalPage(elm, compteur);
                        break;
                    case CookingModeType.COOKING_MODE_SUPPLIER:
                        useCaseModalPage = GoToUseCaseSupplierModalPage(elm, compteur);
                        break;
                    case CookingModeType.COOKING_MODE_GROUP:
                        useCaseModalPage = GoToUseCaseGroupModalPage(elm, compteur);
                        break;
                    case CookingModeType.COOKING_MODE_RECIPE:
                        useCaseModalPage = GoToUseCaseRecipeModalPage(elm, compteur);
                        break;
                    default:
                        useCaseModalPage = GoToUseCaseModalPage(elm, compteur);
                        break;
                }
                WaitPageLoading(); 
                HashSet<string> cookingMode = useCaseModalPage.GetCookingModes();
                cookingModes.UnionWith(cookingMode);
                useCaseModalPage.CloseModal();
                compteur++;
            }
            return cookingModes.ToList();
        }

        public void HideFirstArticle()
        {

            Actions action = new Actions(_webDriver);
            IEnumerable<IWebElement> searchList = new List<IWebElement>();

            if (IsDev())
            {
                WaitForElementToBeClickable(By.XPath(DETAILSNAME));
                searchList = _webDriver.FindElements(By.XPath(DETAILSNAME));
            }
            else
            {
                WaitForElementToBeClickable(By.XPath(DETAILSNAME_PATCH));
                searchList = _webDriver.FindElements(By.XPath(DETAILSNAME_PATCH));
                //WaitForElementToBeClickable(By.CssSelector("div > table > tbody > tr.tr-groupable > td:nth-child(1)"));
                //searchList = _webDriver.FindElements(By.CssSelector("div > table > tbody > tr.tr-groupable > td:nth-child(1)"));
            }
            action.MoveToElement(searchList.FirstOrDefault()).Perform();
            searchList.FirstOrDefault().Click();

            if (IsDev())
                _hideArticle = WaitForElementIsVisible(By.XPath(string.Format(HIDE_ARTICLE, _detail.Text)));
            else _hideArticle = WaitForElementIsVisible(By.Id("production-inflightrawmaterials-setitemhidden-0-1"));
            _hideArticle.Click();
            WaitForLoad();
        }

        public string GetDelieveryIsRequiredMessage()
        {
            _delieveryPlaceIsRequired = WaitForElementIsVisible(By.XPath(DELIVERY_PLACE_IS_REQUIRED));
            return _delieveryPlaceIsRequired.Text.Trim();
        }


        public void ChooseDeliveryLocationOption(string deliveryLocOption)
        {
            _supplyOrderDeliveryLocation = WaitForElementIsVisible(By.Id(SO_DELIVERY_LOCATION_SELECT));
            _supplyOrderDeliveryLocation.Click();

            _supplyOrderDeliveryLocation = WaitForElementIsVisible(By.Id(SO_DELIVERY_LOCATION_SELECT));
            _supplyOrderDeliveryLocation.SetValue(ControlType.DropDownList, deliveryLocOption);
        }

        public string GetToPlaceIsRequiredMessage()
        {
            _toPlaceIsRequired = WaitForElementIsVisible(By.XPath(TO_PLACE_IS_REQUIRED));
            return _toPlaceIsRequired.Text;
        }

        public string GetFromPlaceIsRequiredMessage()
        {
            _fromPlaceIsRequired = WaitForElementIsVisible(By.XPath(FROM_PLACE_IS_REQUIRED));
            return _fromPlaceIsRequired.Text;
        }

        public void ChooseFromAndToPlaceOptions(string fromPlace, string toPlace , string comment )
        {
            _outputFormPlaceFrom = WaitForElementIsVisible(By.Id(OF_PLACE_FROM));
            _outputFormPlaceFrom.SetValue(ControlType.DropDownList, fromPlace);

            _outputFormPlaceTo = WaitForElementIsVisible(By.Id(OF_PLACE_TO));
            _outputFormPlaceTo.SetValue(ControlType.DropDownList, toPlace);

            _outputFormComment = WaitForElementIsVisible(By.Id(OF_PLACE_COMMENT));
            _outputFormComment.SetValue(ControlType.TextBox, comment);
        }

        public string GetActivatedNavTab()
        {
            System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> navTabs = _webDriver.FindElements(By.XPath(NAV_TAB_LIST));
            foreach (var navTab in navTabs)
            {
                var navTabClass = navTab.GetAttribute("class");
                if (navTabClass.Contains("active"))
                {
                    return navTab.Text;
                }
            }

            return string.Empty;

        }
        // _____________________________________________ Print ____________________________________________________

        public enum PrintType
        {
            RawMaterialsByWorkshop,
            RawMaterialsByGroup,
            AssemblyReport,
            RecipeReportV2,
            RecipeReportDetailedV2,
            HACCPSanitization,
            HACCPPlating,
            HACCPPlatingGroupByFlight,
            HACCPPlatingGroupByRecipe,
            HACCPTraySetup,
            HACCPHotKitchen,
            HACCPSlice,
            HACCPThawing,
            DatasheetByRecipe,
            DatasheetByGuest,
            Haccptraysetupmultiflight,
            Haccptraysetupmultiflightnoguest,
            HACCPSlicekitchen
        }

        public PrintReportPage PrintReport(PrintType reportType, bool versionPrint, bool withDatasheet = false, bool checkproductionname = false)
        {
            Actions actions = new Actions(_webDriver);

            // Suppression de la classe CSS
            string classe = UnClassCSS();

            _extendedMenu = WaitForElementIsVisible(By.XPath(EXTENDED_MENU));
            actions.MoveToElement(_extendedMenu).Perform();

            switch (reportType)
            {
                case PrintType.RawMaterialsByWorkshop:
                    _rawMaterialByWorkshop = WaitForElementExists(By.Id(RAW_MATERIAL_BY_WORKSHOP));
                    _rawMaterialByWorkshop.Click();
                    WaitForLoad();
                    break;

                case PrintType.RawMaterialsByGroup:
                    _rawMaterialByGroup = WaitForElementExists(By.Id(RAW_MATERIAL_BY_GROUP));
                    _rawMaterialByGroup.Click();
                    WaitForLoad();
                    break;

                case PrintType.AssemblyReport:
                    _assembly = WaitForElementExists(By.Id(ASSEMBLY_REPORT));
                    _assembly.Click();
                    if (checkproductionname)
                    {

                        var productionradibutton = WaitForElementExists(By.XPath("//*[@id=\"production\"]"));
                        productionradibutton.SetValue(ControlType.RadioButton, true);

                    }
                    WaitForLoad();
                    break;

                case PrintType.RecipeReportV2:
                    _recipeReportV2 = WaitForElementExists(By.Id(RECIPE_REPORT_V2));
                    _recipeReportV2.Click();
                    WaitForLoad();

                    if (withDatasheet)
                    {
                        _withDatasheet = WaitForElementExists(By.XPath(WITH_DATASHEET));
                        actions.MoveToElement(_withDatasheet).Perform();
                        _withDatasheet.Click();
                        WaitForLoad();
                    }
                    break;

                case PrintType.RecipeReportDetailedV2:
                    _recipeReportDetailledV2 = WaitForElementExists(By.Id(RECIPE_REPORT_DETAILLED_V2));
                    _recipeReportDetailledV2.Click();
                    WaitForLoad();
                    break;

                case PrintType.HACCPSanitization:
                    _HACCPSanitization = WaitForElementExists(By.Id(HACCP_SANATIZATION));
                    _HACCPSanitization.Click();
                    WaitForLoad();
                    break;

                case PrintType.HACCPPlating:
                    _HACCPPlating = WaitForElementExists(By.Id(HACCP_PLATING));
                    _HACCPPlating.Click();
                    WaitForLoad();
                    break;

                case PrintType.HACCPPlatingGroupByFlight:
                    _HACCPPlatingGroup = WaitForElementExists(By.Id(HACCP_PLATINGBYGROUP));
                    _HACCPPlatingGroup.Click();
                    WaitForLoad();
                    break;

                case PrintType.HACCPPlatingGroupByRecipe:
                    _HACCPPlatingRecipe = WaitForElementExists(By.Id(HACCP_PLATINGBYRECIPE));
                    _HACCPPlatingRecipe.Click();
                    WaitForLoad();
                    break;

                case PrintType.HACCPTraySetup:
                    _HACCPTraySetup = WaitForElementExists(By.Id(HACCP_TRAY_SETUP));
                    _HACCPTraySetup.Click();
                    WaitForLoad();
                    break;

                case PrintType.HACCPHotKitchen:
                    _HACCPHotKitchen = WaitForElementExists(By.Id(HACCP_HOT_KITCHEN));
                    _HACCPHotKitchen.Click();
                    WaitForLoad();
                    break;

                case PrintType.HACCPSlice:
                    _HACCPSlice = WaitForElementExists(By.Id(HACCP_SLICE));
                    _HACCPSlice.Click();
                    WaitForLoad();
                    break;

                case PrintType.HACCPThawing:
                    _HACCPThawing = WaitForElementExists(By.Id(HACCP_THAWING));
                    _HACCPThawing.Click();
                    WaitForLoad();
                    break;

                case PrintType.DatasheetByRecipe:
                    _datasheetByRecipe = WaitForElementExists(By.Id(DATASHEET_BY_RECIPE));
                    _datasheetByRecipe.Click();
                    WaitForLoad();
                    break;

                case PrintType.DatasheetByGuest:
                    _datasheetByGuest = WaitForElementExists(By.Id(DATASHEET_BY_GUEST));
                    _datasheetByGuest.Click();
                    WaitForLoad();
                    break;

                case PrintType.Haccptraysetupmultiflight:
                    _haccptraySetupMultiFlight = WaitForElementExists(By.Id(HACCP_TRAY_SET_UP_MULTI_FLIGHT));
                    _haccptraySetupMultiFlight.Click();
                    WaitForLoad();
                    break;

                case PrintType.Haccptraysetupmultiflightnoguest:
                    _haccptraySetupMultiFlightnoguest = WaitForElementExists(By.Id(HACCP_TRAY_SET_UP_MULTI_FLIGHT_NO_GUEST));
                    _haccptraySetupMultiFlightnoguest.Click();
                    WaitForLoad();
                    break;

                case PrintType.HACCPSlicekitchen:
                    _haccpSlicekitchen = WaitForElementExists(By.Id(HACCP_SLICE_KITCHEN));
                    _haccpSlicekitchen.Click();
                    WaitForLoad();
                    break;

                default:
                    break;
            }

            _print = WaitForElementIsVisible(By.Id(PRINT_BTN));
            _print.Click();
            WaitForLoad();

            AddCSS(classe);

            if (versionPrint)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-pdf']"));
                ClickPrintButton();
            }

            //Results are opened in a new tab, switch the driver to the newly created one
            WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(30));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new PrintReportPage(_webDriver, _testContext);
        }
        public FlightPage GoToFlights_FlightPage()
        {
            try
            {
                // On tente de cliquer sur le menu à partir de la barre des menus
                _flights = WaitForElementIsVisible(By.Id(FLIGHT_MENU));
                _flights.Click();

                _flights_LoadingPlans = WaitForElementIsVisible(By.Id(FL_FLIGHTS));
                _flights_LoadingPlans.Click();
            }
            catch
            {
                // Retour page d'accueil
                GoToWinrestHome();

                _flights_Flights_link = WaitForElementIsVisible(By.Id(FL_FLIGHTS_LINK));
                _flights_Flights_link.Click();

            }

            WaitForLoad();

            return new FlightPage(_webDriver, _testContext);
        }
        public Dictionary<string, Dictionary<string, string>> GetGroupWithAllRow()
        {
            Dictionary<string, Dictionary<string, string>> mapGpeItem = new Dictionary<string, Dictionary<string, string>>();

            int  groups = _webDriver.FindElements(By.XPath(LINES)).Count;

            for (var i = 1; i <= groups ; i++)
            {
                var group = WaitForElementIsVisible(By.XPath(String.Format(LINES_NAME, i)));

                string text = group.Text.Substring(0, group.Text.IndexOf("(")).Trim();

                // On ajoute le nom du groupe dans la map en tant que clé
                mapGpeItem.Add(text, new Dictionary<string, string>());


                if (IsDev())
                {
                    System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> items = _webDriver.FindElements(By.XPath(String.Format(ITEM_BY_LINE, i+1)));
                    System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> quantities = _webDriver.FindElements(By.XPath(String.Format(QTY_BY_LINE, i+1)));
                    System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> packaging = _webDriver.FindElements(By.XPath(String.Format(PACKAGING_BY_LINE, i + 1)));
                    System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> fournisseur = _webDriver.FindElements(By.XPath(String.Format(FOURNISSEUR_BY_LINE, i + 1)));
                    System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> workshop = _webDriver.FindElements(By.XPath(String.Format(WORKSHOP_BY_LINE, i + 1)));
                    for (var j = 0; j < items.Count; j++)
                    {
                        var itemAttributes = new Dictionary<string, string>
                        {
                                { "item", items[j].Text },
                                { "quantity", quantities[j].Text },
                                { "packaging", packaging[j].Text },
                                { "fournisseur", fournisseur[j].Text },
                                { "workshop", workshop[j].Text }
                        };
                        foreach (var element in itemAttributes)
                        {
                            mapGpeItem[text].Add(element.Key, element.Value);
                        }
                    }
                }
                else
                {
                    System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> items = _webDriver.FindElements(By.XPath(String.Format("//*[@id=\"tabContentItemContainer\"]/div[2]/div/div[{0}]/div[2]/div/table/tbody/tr[*]/td[1]", i)));
                    System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> quantities = _webDriver.FindElements(By.XPath(String.Format("//*[@id=\"tabContentItemContainer\"]/div[2]/div/div[{0}]/div[2]/div/table/tbody/tr[*]/td[3]", i)));
                    System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> packaging = _webDriver.FindElements(By.XPath(String.Format("//*[@id=\"tabContentItemContainer\"]/div[2]/div/div[{0}]/div[2]/div/table/tbody/tr[*]/td[4]", i)));
                    System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> fournisseur = _webDriver.FindElements(By.XPath(String.Format("//*[@id=\"tabContentItemContainer\"]/div[2]/div/div[{0}]/div[2]/div/table/tbody/tr[*]/td[5]", i)));
                    System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> workshop = _webDriver.FindElements(By.XPath(String.Format("//*[@id=\"tabContentItemContainer\"]/div[2]/div/div[{0}]/div[2]/div/table/tbody/tr[*]/td[6]", i)));
                    for (var j = 0; j < items.Count; j++)
                    {
                        var itemAttributes = new Dictionary<string, string>
                        {
                                { "item", items[j].Text },
                                { "quantity", quantities[j].Text },
                                { "packaging", packaging[j].Text },
                                { "fournisseur", fournisseur[j].Text },
                                { "workshop", workshop[j].Text }
                        };
                        foreach (var element in itemAttributes)
                        {
                            mapGpeItem[text].Add(element.Key, element.Value);
                        }
                    }
                }


               
            }
            return mapGpeItem;
        }

        public Dictionary<string, List<Dictionary<string, string>>> GetGroupWithAllRowMultipleItem()
        {
            Dictionary<string, List<Dictionary<string, string>>> mapGpeItem = new Dictionary<string, List<Dictionary<string, string>>>();

            int groups = _webDriver.FindElements(By.XPath(LINES)).Count;

            for (var i = 1; i <= groups; i++)
            {
                var group = WaitForElementIsVisible(By.XPath(String.Format(LINES_NAME, i)));
                string text = group.Text.Substring(0, group.Text.IndexOf("(")).Trim();

                // Initialize list for each group
                mapGpeItem[text] = new List<Dictionary<string, string>>();

                System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> items;
                System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> quantities;
                System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> packaging;
                System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> fournisseur;
                System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> workshop;

                if (IsDev())
                {
                    items = _webDriver.FindElements(By.XPath(String.Format(ITEM_BY_LINE, i + 1)));
                    quantities = _webDriver.FindElements(By.XPath(String.Format(QTY_BY_LINE, i + 1)));
                    packaging = _webDriver.FindElements(By.XPath(String.Format(PACKAGING_BY_LINE, i + 1)));
                    fournisseur = _webDriver.FindElements(By.XPath(String.Format(FOURNISSEUR_BY_LINE, i + 1)));
                    workshop = _webDriver.FindElements(By.XPath(String.Format(WORKSHOP_BY_LINE, i + 1)));
                }
                else
                {
                    items = _webDriver.FindElements(By.XPath(String.Format("//*[@id=\"tabContentItemContainer\"]/div[2]/div/div[{0}]/div[2]/div/table/tbody/tr[*]/td[1]", i)));
                    quantities = _webDriver.FindElements(By.XPath(String.Format("//*[@id=\"tabContentItemContainer\"]/div[2]/div/div[{0}]/div[2]/div/table/tbody/tr[*]/td[3]", i)));
                    packaging = _webDriver.FindElements(By.XPath(String.Format("//*[@id=\"tabContentItemContainer\"]/div[2]/div/div[{0}]/div[2]/div/table/tbody/tr[*]/td[4]", i)));
                    fournisseur = _webDriver.FindElements(By.XPath(String.Format("//*[@id=\"tabContentItemContainer\"]/div[2]/div/div[{0}]/div[2]/div/table/tbody/tr[*]/td[5]", i)));
                    workshop = _webDriver.FindElements(By.XPath(String.Format("//*[@id=\"tabContentItemContainer\"]/div[2]/div/div[{0}]/div[2]/div/table/tbody/tr[*]/td[6]", i)));
                }

                for (var j = 0; j < items.Count; j++)
                {
                    var itemAttributes = new Dictionary<string, string>
            {
                { "item", items[j].Text },
                { "quantity", quantities[j].Text },
                { "packaging", packaging[j].Text },
                { "fournisseur", fournisseur[j].Text },
                { "workshop", workshop[j].Text }
            };

                    mapGpeItem[text].Add(itemAttributes);
                }
            }

            return mapGpeItem;
        }

        public bool IsHiddenItemsCounterAlertActive()
        {
            WaitLoading();
            WaitPageLoading();
            if (isElementVisible(By.Id(HIDDEN_ITEMS_COUNTER)))
                return true; 
            return false  ;
        }

        public int GetHiddenItemsCounterAlertActive()
        {
            WaitLoading();
            var itemCounter = WaitForElementIsVisible(By.Id(HIDDEN_ITEMS_COUNTER));
            string counter = string.Empty;

            foreach (var c in itemCounter.Text)
            {
                int counterValue = 0;
                if (int.TryParse(c.ToString(), out counterValue))
                {
                    counter = counter + c.ToString();
                }
            }
            return int.Parse(counter);
        }
        public enum ElementIdFilter
        {
            SelectedWorkshopsTrueValue,
            SelectedServicesTrueValue,
            SelectedServiceCategoriesTrueValue,
            SelectedRecipeTypesTrueValue
        }
        public List<string> GetSelectedOptions(ElementIdFilter elementId)
        {
            var selectedOptionsList = new List<string>();

            var selectElement = _webDriver.FindElement(By.Id(elementId.ToString()));

            WaitPageLoading();
            Thread.Sleep(1500);
            System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> options = selectElement.FindElements(By.TagName("option"));

            IJavaScriptExecutor js = (IJavaScriptExecutor)_webDriver;
            var selectedOptions = options.Where(option => option.Selected).ToList();

            foreach (var option in selectedOptions)
            {
                string optionText = (string)js.ExecuteScript("return arguments[0].text;", option);
                Console.WriteLine($"Selected option value: {option.GetAttribute("value")}, Text: {optionText}");
                selectedOptionsList.Add(optionText.Trim());
            }

            return selectedOptionsList;
        }
        public enum FilterType
        {
            ShowHiddenArticlesResults,
            DateFrom
        }
        public int Filter(ResultPage.FilterType filterType, object value)
        {
            switch (filterType)
            {
                case FilterType.ShowHiddenArticlesResults:
                    IWebElement checkboxValue;
                    checkboxValue = WaitForElementExists(By.Id("ShowHiddenArticles"));
                    checkboxValue.SetValue(ControlType.CheckBox, value);
                    break;
                case FilterType.DateFrom:
                    IWebElement dateFrom = WaitForElementExists(By.Id("ProdDateFrom"));
                    dateFrom.SetValue(ControlType.DateTime, value);
                    break;
            }
            WaitForLoad();
            WaitPageLoading();
            return -1;
        }

        public void HideFirstArticleResult()
        {

            Actions action = new Actions(_webDriver);

            var _detailHide = _webDriver.FindElements(By.Id("production-inflightrawmaterials-setitemhidden-0-1")).FirstOrDefault();

            action.MoveToElement(_detailHide).Perform();
            _detailHide.Click();

            WaitForLoad();
        }

        public void ToggleItem()
        {
            var _toggleItem = WaitForElementIsVisible(By.XPath("//*[@id=\"tabContentItemContainer\"]/div[2]/div/div[1]/div[1]/div/div[1]"));
            _toggleItem.Click();
            WaitForLoad();
            WaitPageLoading();
        }

        public enum CookingModeType
        {
            COOKING_MODE_RECIPE,
            COOKING_MODE_GROUP,
            COOKING_MODE_SUPPLIER,
            COOKING_MODE_WORKSHOP,
            NONE
        }
        public bool IsthereAnySubGroupSelected()
        {
            ScrollUntilElementIsInView(By.Id(SUBGROUP_FILTER));
            var _subGroupFilter = _webDriver.FindElement(By.Id(SUBGROUP_FILTER));
            if (_subGroupFilter.Text == "Nothing Selected")
                return true;
            else
                return false;
            WaitPageLoading();
            WaitForLoad();

        }
        public bool IsCorrectOptionSelected()
        {
            var dropdown = _webDriver.FindElement(By.Id("DdlSelectedView"));

            if (dropdown.Displayed)
            {
                SelectElement select = new SelectElement(dropdown);

                string selectedValue = select.SelectedOption.GetAttribute("value");

                return selectedValue == "RawMatByRecipe";
            }
            else
            {
                return false;
            }
        }
        public bool IsCorrectRawMatByWorkshopSelected()
        {
            var dropdown = _webDriver.FindElement(By.Id("DdlSelectedView"));

            if (dropdown.Displayed)
            {
                SelectElement select = new SelectElement(dropdown);

                string selectedValue = select.SelectedOption.GetAttribute("value");

                return selectedValue == "RawMatByWorkshop";
            }
            else
            {
                return false;
            }
        }
    }
}
