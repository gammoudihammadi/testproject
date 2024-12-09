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
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Needs

{
    public class ResultPageNeeds : PageBase
    {
        protected IWebDriver WebDriver
        {
            get => WebDriverFactory.Driver;
        }

        public ResultPageNeeds(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //__________________________________Constantes_____________________________________

        // Général
        private const string EXTENDED_MENU = "//*[@id=\"div-body\"]/div/div[2]/div[1]/div/div[1]/button";
        private const string PLUS_MENU = "//*[@id=\"div-body\"]/div/div[2]/div[1]/div/div[2]/button";
        private const string GENERATE_SUPPLY_ORDER = "//*[@id=\"div-body\"]/div/div[2]/div[1]/div/div[2]/div/a[text() = 'Generate Supply Order']";
        private const string GENERATE = "btn-submit-form-create-supply-order";
        private const string SO_NUMBER = "tb-new-supplyorder-number";
        private const string SO_LOCATION = "SelectedSitePlaceId";
        private const string ROUND_QUANTITIES = "checkBoxRoundPackagingQtys";

        private const string FOLD_UNFOLD = "//*[@id=\"div-body\"]/div/div[2]/div[1]/div/div[1]/div/a[1]";

        // Onglet
        private const string QTY_ADJUSTEMENT = "//*[@id=\"hrefTabContentItemContainer\"][text()='Quantity adjustments']";

        // Filtres
        private const string BACK = "//*[@id=\"item-filter-form\"]/div[1]/a";
        private const string FILTER_DATE_FROM = "ProdDateFrom";

        // Tableau
        private const string RAW_MAT_BY = "DdlSelectedView";
        private const string GROUP_BY_DROP = "DdlSelectedGroupBy";
        private const string GROUP_OPTION = "//*[@id=\"DdlSelectedGroupBy\"]/option[contains(@value,'{0}')]";

        private const string LINES = "//*[@id=\"tabContentItemContainer\"]/div[2]/div/div[*]";
        private const string LINES_NAME = "//*[@id=\"tabContentItemContainer\"]/div[2]/div/div[{0}]";
        private const string ARROW_LINE = "//*[@id=\"tabContentItemContainer\"]/div[2]/div/div[*]/div[1]/div/div[1]/span";
        private const string ITEM_BY_LINE = "//*[@id=\"tabContentItemContainer\"]/div[2]/div/div[{0}]/div[2]/div/table/tbody/tr[*]/td[1]";

        private const string SUB_LINES = "//*[starts-with(@id,\"content_\")]/div/div[*]/div/div/div[2]/table/tbody/tr/td[1]";
        private const string SUB_LINES_NAME = "//*[starts-with(@id,\"content_\")]/div/div[{0}]/div/div/div[2]/table/tbody/tr/td[1]";
        private const string SUB_LINES_RECIPE = "//*[starts-with(@id,\"content_\")]/div/div[{0}]/div/div/div[2]/table/tbody/tr/td[3]";
        private const string ITEM_BY_SUBGROUP = "/html/body/div[2]/div/div[2]/div[3]/div/div[2]/div/div/div[2]/div/div[{0}]/div/table/tbody/tr[*]/td[1]";
        private const string ITEM_BY_SUBROUP_Workshop = "/html/body/div[2]/div/div[2]/div[3]/div/div[2]/div/div/div[2]/div/div[{0}]/div/table/tbody/tr[*]/td[5]";
        private const string DETAILS = "//*[starts-with(@id,\"content_\")]";
        private const string DETAILSNAME = "//*[starts-with(@id, 'production-inflightrawMaterials-itemusecase-itemname-0-')]";
        private const string DETAILS_ITEM_NAME = "//*[starts-with(@id,\"production-inflightrawMaterials-itemusecase-itemname\")]";
        private const string DETAILS_ITEM_WORKSHOP = "//*[@id=\"tabContentItemContainer\"]/div[2]/div/div/div/div/div/div/div";
        private const string DETAILS_ITEM_WORKSHOP_RECIPE = "//*[@id=\"tabContentItemContainer\"]/div[2]/div/div/div/div/div/div/div/div/table/tbody/tr[1]/td[2]";
        private const string DETAILS_ITEM_WORKSHIOP = "//*[@id=\"tabContentItemContainer\"]/div[2]/div/div/div/div/div/div/table/tbody/tr/td[1]";

        private const string USECASES_BTN = "//table//tr[contains(@class,\"tr-groupable\")]/td[1][contains(text(),'{0}')]/..//a[contains(@class,'btn btn-use-case')]";
        private const string HIDE_ARTICLE = "//*[starts-with(@id,\"content_\")]/div/table/tbody/tr[*]/td[contains(text(), '{0}')]/../td[67/a[2]/span";

        private const string GUESTTYPE = "//*[starts-with(@id,\"content_\")]/div/div/div/div/div[2]/table/tbody/tr[*]/td[2]";
        private const string WORKSHOP_NAME = "//*[starts-with(@id, 'production-inflightrawMaterials-itemusecase-workshopname-0-')]";

        private const string FILTER_BY = "DdlSelectedGroupBy";

        private const string RAW_MAT_BY_GROUP = "//*[@id=\"DdlSelectedView\"]";
        private const string RIGHT_LIST = "//*[@id=\"DdlSelectedGroupBy\"]";

        //__________________________________Variables______________________________________

        // Général

        [FindsBy(How = How.XPath, Using = EXTENDED_MENU)]
        public IWebElement _extendedMenu;

        [FindsBy(How = How.XPath, Using = PLUS_MENU)]
        public IWebElement _plusMenu;

        [FindsBy(How = How.XPath, Using = GENERATE_SUPPLY_ORDER)]
        public IWebElement _generateSupplyOrder;

        [FindsBy(How = How.Id, Using = SO_NUMBER)]
        public IWebElement _supplyOrderNumber;

        [FindsBy(How = How.Id, Using = SO_LOCATION)]
        public IWebElement _supplyOrderLocation;

        [FindsBy(How = How.Id, Using = ROUND_QUANTITIES)]
        private IWebElement _roundQuantities;

        [FindsBy(How = How.XPath, Using = GENERATE)]
        public IWebElement _generate;

        [FindsBy(How = How.XPath, Using = FOLD_UNFOLD)]
        public IWebElement _unfold;

        // Onglets
        [FindsBy(How = How.XPath, Using = QTY_ADJUSTEMENT)]
        public IWebElement _qtyAdjustement;

        // Filtres
        [FindsBy(How = How.Id, Using = FILTER_DATE_FROM)]
        public IWebElement _dateFrom;

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

        [FindsBy(How = How.XPath, Using = DETAILSNAME)]
        public IWebElement _detail;

        [FindsBy(How = How.XPath, Using = DETAILSNAME)]
        public IWebElement _firstItemName;

        [FindsBy(How = How.XPath, Using = HIDE_ARTICLE)]
        public IWebElement _hideArticle;

        [FindsBy(How = How.XPath, Using = FILTER_BY)]
        public IWebElement _filterBy;
        //___________________________________ Méthodes _________________________________________

        // Général
        public void UnfoldAll()
        {
            Actions actions = new Actions(_webDriver);

            _extendedMenu = WaitForElementIsVisible(By.XPath(EXTENDED_MENU));
            actions.MoveToElement(_extendedMenu).Perform();

            _unfold = WaitForElementExists(By.XPath(FOLD_UNFOLD));
            _unfold.Click();
            WaitForLoad();

            Thread.Sleep(1000);
        }

        public Boolean IsUnfoldAll()
        {
            var elements = _webDriver.FindElements(By.XPath(DETAILS));

            foreach (var elm in elements)
            {
                if (!elm.GetAttribute("class").Equals("panel-collapse collapse show"))
                    return false;
            }

            return true;
        }

        public void FoldAll()
        {
            Actions actions = new Actions(_webDriver);

            _extendedMenu = WaitForElementIsVisible(By.XPath(EXTENDED_MENU));
            actions.MoveToElement(_extendedMenu).Perform();

            var fold = WaitForElementExists(By.XPath(FOLD_UNFOLD));
            fold.Click();
            WaitForLoad();

            Thread.Sleep(1000);
        }

        public Boolean IsFoldAll()
        {
            var elements = _webDriver.FindElements(By.XPath(DETAILS));

            foreach (var elm in elements)
            {
                if (!elm.GetAttribute("class").Equals("panel-collapse collapse"))
                    return false;
            }

            return true;
        }

        public string GenerateSupplyOrder(string location, bool roundQuantities = true)
        {
            Actions actions = new Actions(_webDriver);

            _plusMenu = WaitForElementIsVisible(By.XPath(PLUS_MENU));
            actions.MoveToElement(_plusMenu).Perform();

            _generateSupplyOrder = WaitForElementIsVisible(By.XPath(GENERATE_SUPPLY_ORDER));
            _generateSupplyOrder.Click();
            WaitForLoad();

            _supplyOrderLocation = WaitForElementIsVisible(By.Id(SO_LOCATION));
            _supplyOrderLocation.SetValue(ControlType.DropDownList, location);

            _roundQuantities = WaitForElementExists(By.Id(ROUND_QUANTITIES));
            _roundQuantities.SetValue(ControlType.CheckBox, roundQuantities);

            _supplyOrderNumber = WaitForElementIsVisible(By.Id(SO_NUMBER));

            return _supplyOrderNumber.GetAttribute("value");
        }

        public SupplyOrderItem Submit()
        {
            _generate = WaitForElementIsVisible(By.Id(GENERATE));
            _generate.Click();
            WaitForLoad();

            //Results are opened in a new tab, switch the driver to the newly created one
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            WaitForLoad();

            return new SupplyOrderItem(_webDriver, _testContext);
        }

        // Onglets
        public QuantityAdjustmentsNeedsPage GoToQtyAdjustementPage()
        {
            _qtyAdjustement = WaitForElementIsVisible(By.XPath(QTY_ADJUSTEMENT));
            _qtyAdjustement.Click();
            WaitPageLoading();

            return new QuantityAdjustmentsNeedsPage(_webDriver, _testContext);
        }

        // Filtres
        public FilterAndFavoritesNeedsPage Back()
        {
            _back = WaitForElementIsVisible(By.XPath(BACK));
            _back.Click();
            WaitForLoad();

            return new FilterAndFavoritesNeedsPage(_webDriver, _testContext);
        }

        public bool VerifyDateFrom(DateTime value)
        {
            var dateFrom = _webDriver.FindElements(By.XPath(LINES)).FirstOrDefault();

            if (dateFrom != null)
            {
                var dateFormat = _dateFrom.GetAttribute("data-date-format");
                CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? ci = new CultureInfo("fr-FR") : ci = new CultureInfo("en-US");

                DateTime date = DateTime.Parse(dateFrom.Text, ci).Date;

                if (value <= date)
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
                //_rawMatBy.Click();

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
            return _webDriver.FindElements(By.XPath(LINES)).Count;
        }

        public string GetFirstItemGroup()
        {
            try
            {
                _lines = WaitForElementIsVisible(By.XPath(LINES));
                int parenthese = _lines.Text.IndexOf("(");

                return _lines.Text.Substring(0, parenthese).Trim();
            }
            catch
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
            try
            {
                _subLines = WaitForElementIsVisible(By.XPath(SUB_LINES));
                return _subLines.Text.Trim();
            }
            catch
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

            var element = _webDriver.FindElement(By.XPath(DETAILS));
            if (!element.GetAttribute("class").Equals("panel-collapse collapse show"))
                return false;
            return true;
        }

        public string GetFirstItemName()
        {
            _firstItemName = WaitForElementIsVisible(By.XPath(DETAILS_ITEM_NAME));
            return _firstItemName.Text;
        }

        public Dictionary<string, List<string>> GetGroupAndItemsAssociated()
        {
            Dictionary<string, List<string>> mapGpeItem = new Dictionary<string, List<string>>();

            var groups = _webDriver.FindElements(By.XPath(LINES)).Count;

            for (var i = 1; i <= groups; i++)
            {
                var group = WaitForElementIsVisible(By.XPath(String.Format(LINES_NAME, i)));

                string text = group.Text.Substring(0, group.Text.IndexOf("(")).Trim();

                // On ajoute le nom du groupe dans la map en tant que clé
                mapGpeItem.Add(text, new List<string>());

                var items = _webDriver.FindElements(By.XPath(String.Format(ITEM_BY_LINE, i)));

                foreach (var item in items)
                {
                    mapGpeItem[text].Add(item.Text);
                }

            }

            return mapGpeItem;
        }

        public bool GetGroupAndRecipeAssociated(string type)
        {
            Actions action = new Actions(_webDriver);

            var groups = _webDriver.FindElements(By.XPath(LINES)).Count;

            for (var i = 1; i <= groups; i++)
            {
                var group = WaitForElementIsVisible(By.XPath(String.Format(LINES_NAME, i)));

                string value = group.Text.Substring(0, group.Text.IndexOf("(")).Trim();

                var items = _webDriver.FindElements(By.XPath(String.Format(ITEM_BY_LINE, i)));

                foreach (var item in items)
                {
                    action.MoveToElement(item).Perform();
                    item.Click();

                    var useCaseModalPage = GoToUseCaseModalPage(item.Text);

                    if (type.Equals("Workshop") && !useCaseModalPage.IsWorkshop(value))
                        return false;

                    else if (type.Equals("RecipeType") && !useCaseModalPage.IsRecipeType(value))
                        return false;

                    useCaseModalPage.CloseModal();
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

                var items = _webDriver.FindElements(By.XPath(String.Format(ITEM_BY_SUBROUP_Workshop, i + 1)));

                foreach (var item in items)
                {
                    if (item.Text != workshop)
                    {
                        return false;
                    }
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

                var items = _webDriver.FindElements(By.XPath(String.Format(ITEM_BY_SUBGROUP, i + 1)));

                foreach (var item in items)
                {
                    mapRecipeItems[recipe].Add(item.Text.Trim());
                }

                i++;
            }

            return mapRecipeItems;
        }

        public ResultModalNeeds ShowUseCase()
        {
            Actions action = new Actions(_webDriver);
            var item = _webDriver.FindElement(By.XPath(DETAILSNAME));

            action.MoveToElement(item).Perform();
            item.Click();

            return GoToUseCaseModalPage(item.Text);
        }

        public List<string> GetCustomerNames()
        {
            List<string> customerNames = new List<string>();

            var customers = _webDriver.FindElements(By.XPath(SUB_LINES));

            foreach (var customer in customers)
            {
                customerNames.Add(customer.Text.Trim());
            }

            return customerNames;
        }

        public int CountItems()
        {
            var elements = _webDriver.FindElements(By.XPath(DETAILSNAME));
            return elements.Count(elm => !string.IsNullOrEmpty(elm.Text.Trim()));
        }
        public int CountItemsRawMatGroup()
        {
            var elements = _webDriver.FindElements(By.XPath("//*[@id=\"tabContentItemContainer\"]/div[2]/div/div/div/div/div[2]"));
            return elements.Count(elm => !string.IsNullOrEmpty(elm.Text.Trim()));
        }

        public List<string> GetItemNames()
        {
            List<string> names = new List<string>();

            WaitPageLoading();
            var elements = _webDriver.FindElements(By.XPath(DETAILS_ITEM_NAME));

            foreach (var elm in elements)
            {
                string text = elm.Text.Trim();
                if (!string.IsNullOrEmpty(text))
                {
                    names.Add(text);
                }
            }

            return names;
        }

        public bool VerifyGuestType(string guestType)
        {
            var elements = _webDriver.FindElements(By.XPath(GUESTTYPE));

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
            WaitForElementToBeClickable(By.XPath(DETAILSNAME));

            var searchList = _webDriver.FindElements(By.XPath(DETAILSNAME));

            foreach (var elm in searchList)
            {
                action.MoveToElement(elm).Perform();
                elm.Click();

                var useCaseModalPage = GoToUseCaseModalPage(elm.Text);

                WaitForLoad();
                if (!useCaseModalPage.IsServiceCategorie(categorieService))
                    return false;

                useCaseModalPage.CloseModal();
            }

            return true;
        }
        public bool VerifyCategorieForRawMatWorkShop(string categorieService)
        {
            Actions action = new Actions(_webDriver);

            var searchList = _webDriver.FindElements(By.XPath(DETAILS_ITEM_WORKSHIOP));

            foreach (var elm in searchList)
            {
                action.MoveToElement(elm).Perform();
                elm.Click();

                var useCaseModalPage = GoToUseCaseModalPage(elm.Text);

                if (!useCaseModalPage.IsServiceCategorie(categorieService))
                    return false;

                useCaseModalPage.CloseModal();
            }

            return true;
        }

        public bool VerifyRecipeType(string recipeType)
        {
            Actions action = new Actions(_webDriver);

            WaitForElementToBeClickable(By.XPath(DETAILSNAME));
            var searchList = _webDriver.FindElements(By.XPath(DETAILSNAME));

            foreach (var elm in searchList)
            {
                action.MoveToElement(elm).Perform();
                WaitForLoad();
                elm.Click();

                var useCaseModalPage = GoToUseCaseModalPage(elm.Text);

                if (!useCaseModalPage.IsRecipeType(recipeType))
                    return false;

                useCaseModalPage.CloseModal();

            }

            return true;
        }
        public bool VerifyRecipeForRawMatWorkShop(string recipe)
        {
            Actions action = new Actions(_webDriver);

            var searchList = _webDriver.FindElements(By.XPath(DETAILS_ITEM_WORKSHIOP));

            foreach (var elm in searchList)
            {
                action.MoveToElement(elm).Perform();
                elm.Click();

                var useCaseModalPage = GoToUseCaseModalPage(elm.Text);

                if (!useCaseModalPage.IsRecipeType(recipe))
                    return false;

                useCaseModalPage.CloseModal();
            }

            return true;
        }

        public HashSet<string> GetFlights()
        {
            Actions action = new Actions(_webDriver);
            HashSet<string> flightList = new HashSet<string>();

            var searchList = _webDriver.FindElements(By.XPath(DETAILSNAME));

            foreach (var elm in searchList)
            {
                action.MoveToElement(elm).Perform();
                elm.Click();

                var useCaseModalPage = GoToUseCaseModalPage(elm.Text);

                flightList.UnionWith(useCaseModalPage.GetFlights());

                useCaseModalPage.CloseModal();
            }

            return flightList;
        }

        public bool VerifyCustomerRawMat(string customer)
        {
            Actions action = new Actions(_webDriver);

          WaitForElementToBeClickable(By.XPath(DETAILSNAME));
            var searchList = _webDriver.FindElements(By.XPath(DETAILSNAME));

            foreach (var elm in searchList)
            {
                action.MoveToElement(elm).Perform();
                elm.Click();

                var useCaseModalPage = GoToUseCaseModalPage(elm.Text);

                if (!useCaseModalPage.IsCustomer(customer))
                    return false;

                useCaseModalPage.CloseModal();

            }

            return true;
        }

        public bool VerifyServiceAndCustomer(string service, string customer)
        {
            Actions action = new Actions(_webDriver);

            WaitForElementToBeClickable(By.XPath(DETAILS_ITEM_NAME));
            var searchList = _webDriver.FindElements(By.XPath(DETAILS_ITEM_NAME));

            foreach (var elm in searchList)
            {
                action.MoveToElement(elm).Perform();
                elm.Click();

                var useCaseModalPage = GoToUseCaseModalPage(elm.Text);

                if (!useCaseModalPage.IsService(service))
                    return false;

                if (!useCaseModalPage.IsCustomer(customer))
                    return false;

                useCaseModalPage.CloseModal();

            }

            return true;
        }

        public ResultModalNeeds GoToUseCaseModalPage(string itemName)
        {
            var btn = WaitForElementIsVisible(By.XPath(string.Format(USECASES_BTN, itemName)));
            btn.Click();
            WaitForLoad();

            return new ResultModalNeeds(_webDriver, _testContext);
        }

        public bool VerifyService(string service)
        {
            Actions action = new Actions(_webDriver);

            WaitForElementToBeClickable(By.XPath(DETAILSNAME));
            var searchList = _webDriver.FindElements(By.XPath(DETAILSNAME));

            foreach (var elm in searchList)
            {
                action.MoveToElement(elm).Perform();
                elm.Click();

                var useCaseModalPage = GoToUseCaseModalPage(elm.Text);

                if (!useCaseModalPage.IsService(service))
                    return false;

                useCaseModalPage.CloseModal();
            }

            return true;
        }
        public bool VerifyServiceRawMat(string service)
        {
            Actions action = new Actions(_webDriver);

            var searchList = _webDriver.FindElements(By.XPath(DETAILS_ITEM_WORKSHIOP));


            foreach (var elm in searchList)
            {
                action.MoveToElement(elm).Perform();
                elm.Click();

                var useCaseModalPage = GoToUseCaseModalPage(elm.Text);

                if (!useCaseModalPage.IsService(service))
                    return false;

                useCaseModalPage.CloseModal();
            }

            return true;
        }
        public bool VerifyServiceForRawMatWorkShop(string service)
        {
            Actions action = new Actions(_webDriver);

            var searchList = _webDriver.FindElements(By.XPath(DETAILS_ITEM_WORKSHIOP));

            foreach (var elm in searchList)
            {
                action.MoveToElement(elm).Perform();
                elm.Click();

                var useCaseModalPage = GoToUseCaseModalPage(elm.Text);

                if (!useCaseModalPage.IsService(service))
                    return false;

                useCaseModalPage.CloseModal();
            }

            return true;
        }
        public bool VerifyWorkshop(string workshop)
        {
            Actions action = new Actions(_webDriver);

            WaitForElementToBeClickable(By.XPath(DETAILSNAME));
            var searchList = _webDriver.FindElements(By.XPath(DETAILSNAME));

            foreach (var elm in searchList)
            {
                action.MoveToElement(elm).Perform();
                elm.Click();

                var workshopItem = WaitForElementIsVisible(By.XPath(String.Format(WORKSHOP_NAME, elm.Text)));

                if (!workshopItem.Text.Equals(workshop))
                    return false;

                var useCaseModalPage = GoToUseCaseModalPage(elm.Text);

                if (!useCaseModalPage.IsWorkshop(workshop))
                    return false;

                useCaseModalPage.CloseModal();

            }

            return true;
        }
        public bool VerifyCustomerForRawMatWorkShop(string customer)
        {
            Actions action = new Actions(_webDriver);

            var searchList = _webDriver.FindElements(By.XPath(DETAILS_ITEM_WORKSHIOP));

            foreach (var elm in searchList)
            {
                action.MoveToElement(elm).Perform();
                elm.Click();

                var useCaseModalPage = GoToUseCaseModalPage(elm.Text);

                if (!useCaseModalPage.IsCustomer(customer))
                    return false;

                useCaseModalPage.CloseModal();
            }

            return true;
        }
        public List<String> GetCookingModesForWorkShop()
        {
            Actions action = new Actions(_webDriver);
            HashSet<String> cookingModes = new HashSet<string>();

            WaitForElementToBeClickable(By.XPath(DETAILS_ITEM_WORKSHIOP));
            var searchList = _webDriver.FindElements(By.XPath(DETAILS_ITEM_WORKSHIOP));

            foreach (var elm in searchList)
            {
                action.MoveToElement(elm).Perform();
                elm.Click();

                var useCaseModalPage = GoToUseCaseModalPage(elm.Text);

                cookingModes.UnionWith(useCaseModalPage.GetCookingModes());

                useCaseModalPage.CloseModal();

            }

            return cookingModes.ToList();
        }
        public bool VerifyWorkshopForRawMatWorkShop(string workshop)
        {
            Actions action = new Actions(_webDriver);

    
            var workShop = WaitForElementIsVisible(By.XPath(DETAILS_ITEM_WORKSHOP));
            if (workShop.Text == workshop )
            {
                return true; 
            }           

            return false ;
        }

        public bool VerifyWorkshopForRecipe(string workshop)
        {
            Actions action = new Actions(_webDriver);


            var workShop = WaitForElementIsVisible(By.XPath(DETAILS_ITEM_WORKSHOP_RECIPE));
            if (workShop.Text.Contains(workshop))
            {
                return true;
            }

            return false;
        }
        public List<String> GetCookingModes()
        {
            Actions action = new Actions(_webDriver);
            HashSet<String> cookingModes = new HashSet<string>();

            WaitForElementToBeClickable(By.XPath(DETAILSNAME));
            var searchList = _webDriver.FindElements(By.XPath(DETAILSNAME));

            foreach (var elm in searchList)
            {
                action.MoveToElement(elm).Perform();
                elm.Click();

                var useCaseModalPage = GoToUseCaseModalPage(elm.Text);

                cookingModes.UnionWith(useCaseModalPage.GetCookingModes());

                useCaseModalPage.CloseModal();

            }

            return cookingModes.ToList();
        }

        public void HideFirstArticle()
        {

            Actions action = new Actions(_webDriver);

            _detail = _webDriver.FindElements(By.XPath(DETAILSNAME)).FirstOrDefault();

            action.MoveToElement(_detail).Perform();
            _detail.Click();
            WaitPageLoading();
            //     _hideArticle = WaitForElementIsVisible(By.XPath(String.Format(HIDE_ARTICLE, _detail.Text)));

            _hideArticle = WaitForElementIsVisible(By.XPath("//*[@id=\"tabContentItemContainer\"]/div[2]/div/div[1]/div[1]/div/div[1]"));
            _hideArticle.Click();
            WaitForLoad();
        }
        public void FilterByCustomer()
        {
            /* The filter element */
            _filterBy = WaitForElementIsVisible(By.Id("DdlSelectedGroupBy"));
            _filterBy.Click();

            /* The customer option */
            var element = WaitForElementIsVisible(By.XPath("//*[@id=\"DdlSelectedGroupBy\"]/option[3]"));
            element.Click();
            WaitForLoad();

        }
        public void ShowHidden()
        {
            IWebElement checkboxValue;
            checkboxValue = WaitForElementExists(By.Id("ShowHiddenArticles"));
            checkboxValue.SetValue(ControlType.CheckBox, true);
          

        }
        public List<string> GetNeedsResultsContant()
        {
            List<string> customerNames = new List<string>();

            var customers = _webDriver.FindElements(By.XPath("//*[@id=\"tabContentItemContainer\"]/div[2]"));

            foreach (var customer in customers)
            {
                customerNames.Add(customer.Text.Trim());
            }

            return customerNames;
        }
        public IWebElement GetFilterDropdown()
        {
            // Assuming the dropdown has an ID or a specific XPath you can use
            return _webDriver.FindElement(By.XPath(RAW_MAT_BY_GROUP));
        }

        public bool IsRightListDisplayed()
        {
            try
            {
                // Replace the XPath with the actual one that targets the list on the right side
                var rightList = _webDriver.FindElement(By.XPath(RIGHT_LIST));
                return rightList.Displayed;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }
        public int GetNeedsResultsCount()
        {
            System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> customers = _webDriver.FindElements(By.XPath("//*[@id=\"tabContentItemContainer\"]/div[2]/div/div[*]"));
            return customers.Count;
        }
        public string GetNeedsResultName()
        {
            var customers = _webDriver.FindElements(By.XPath("//*[@id=\"tabContentItemContainer\"]/div[2]"));
            string customer = customers[0].Text;
            return customer;
        }
    }
}
