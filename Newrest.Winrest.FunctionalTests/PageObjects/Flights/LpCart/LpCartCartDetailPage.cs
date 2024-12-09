using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Flights.LpCart
{
    public class LpCartCartDetailPage : PageBase
    {

        // ____________________________________ Constantes __________________________________________

        private const string ADDTROLLEY_BTN = "addLPCartDetail";

        //item Trolley
        private const string TROLLEY_ITEM = "/html/body/div[3]/div/div[2]/div/div/div[2]/div/div[2]/table/tbody/tr[last()]";
        private const string TROLLEY_INPUTS = "/html/body/div[3]/div/div[2]/div/div/div[2]/div/div[2]/table/tbody/tr[last()]/td[*]/input[last()]";
        private const string TRASH_BTN_TROLLEY = "//span[contains(@class, 'glyphicon glyphicon-trash')]";
        private const string DELETE_BTN_TROLLEY = "//*[@id=\"delete-form\"]/div[3]/button[1]";
        private const string LINE_COUNT = "/html/body/div[3]/div/div[2]/div/div/div[2]/div/div[2]/table/tbody/tr[2]/td[9]";
        private const string COLUMN_COUNT = "/html/body/div[3]/div/div[2]/div/div/div[2]/div/div[2]/table/tbody/tr[2]/td[10]";
        private const string TROLLEY_NAMES = "/html/body/div[3]/div/div[2]/div/div/div[2]/div/div[2]/table/tbody/tr[*]/td[3]/span";
        private const string FIRST_ARROW_DOWN_BTN = "down";
        private const string FIRST_ARROW_DOWN_BTN_DEV = "//*[@id=\"down\"][contains(@class, 'btn changeDirection ')]";

        //Extend Menu
        private const string EXTENDED_BTN = "//*[@id=\"tabContentItemContainer\"]/div/div/div[1]/button";
        private const string DUPLICATE_CARTS_TROLLEY = "//*[@id=\"tabContentItemContainer\"]/div[1]/div/div[1]/div/a[1]";

        //filter
        private const string SEARCH_FILTER = "SearchPattern";
        private const string SORT_BY_FILTER = "cbSortBy";

        //flight detail
        private const string FLIGHT_TAB = "hrefTabContentFlight";
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";
        private const string EDIT_CART_SCHEME = "//*[@id=\"LPCartDetailsTable\"]/tbody/tr[2]/td[12]/span/a";

        private const string FLIGHTS_TAB = "//*[@id=\"LPCartDetailTab\"]/li[3]";
        private const string FIRST_FLIGHT_NUMBER_FLIGHTS_TAB = "//*[@id=\"services-table\"]/tbody/tr[2]/td[1]";
        private const string FIRST_FLIGHT_DATE_FLIGHTS_TAB = "//*[@id=\"services-table\"]/tbody/tr[2]/td[5]";
        private const string FIRST_FLIGHT_SITE_FROM_FLIGHTS_TAB = "//*[@id=\"services-table\"]/tbody/tr[2]/td[2]";
        private const string FIRST_FLIGHT_SITE_TO_FLIGHTS_TAB = "//*[@id=\"services-table\"]/tbody/tr[2]/td[3]";

        private const string TROLLEY_NAME_INPUT = "//*[@id=\"LPCartDetailsTable\"]/tbody/tr[{0}]/td[3]/input[@id=\"TrolleyName\"]";
        private const string TROLLEY_CODE_INPUT = "//*[@id=\"LPCartDetailsTable\"]/tbody/tr[{0}]/td[4]/input[@id=\"TrolleyCode\"]";
        private const string LEG_INPUT = "//*[@id=\"LPCartDetailsTable\"]/tbody/tr[{0}]/td[1]/input[@id=\"LegNo\"]";
        private const string LINE = "//*[@id=\"LPCartDetailsTable\"]/tbody/tr[{0}]";
        private const string ORDERS = "/html/body/div[3]/div/div[2]/div/div/div[2]/div/div[2]/table/tbody/tr[*]/td[2]";
        private const string CODES = "/html/body/div[3]/div/div[2]/div/div/div[2]/div/div[2]/table/tbody/tr[*]/td[4]/span";
        private const string NAMES = "/html/body/div[3]/div/div[2]/div/div/div[2]/div/div[2]/table/tbody/tr[*]/td[3]/span";
        private const string ADD_TROLLEY_BTN = "//*[@id=\"up\"]";
        private const string CARTS_TAB = "//*[@id=\"hrefTabContentCarts\"]";
        private const string CARTS_NAME_LIST = "//*[@id=\"LPCartDetailsTable\"]/tbody/tr[*]/td[3]/span";
        private const string CARTS_CODE_LIST = "//*[@id=\"LPCartDetailsTable\"]/tbody/tr[*]/td[4]/span";
        private const string PLUS_BTN = "//*[@id=\"addLPCartDetail\"]";
        private const string PLUS_BTN_ID = "addLPCartDetail";
        private const string LEG_NO = "//*[@id=\"LegNo\"]";
        private const string TROLLEYNAME = "//*[@id=\"TrolleyName\"]";
        private const string TROLLEYCODE = "//*[@id=\"TrolleyCode\"]";
        private const string TROLLEYLOC = "//*[@id=\"TrolleyLoc\"]";
        private const string GALLEYCODE = "//*[@id=\"GalleyCode\"]";
        private const string GALLEYLOC = "//*[@id=\"GalleyLoc\"]";
        private const string EQUIPMENT = "//*[@id=\"Equipment\"]";
        private const string SHORTCOMMENT = "//*[@id=\"ShortComment\"]";
        private const string PLUSROW = "//*[@id=\"LPCartDetailsTable\"]/tbody/tr[2]/td[12]/span/a/span[@class='fas fa-plus']";
        private const string EDITROW = "//*[@id=\"LPCartDetailsTable\"]/tbody/tr[2]/td[12]/span/a/span[@class='fas fa-pencil-alt']";
        private const string DELETEROW = "//*[@id=\"LPCartDetailsTable\"]/tbody/tr[2]/td[16]/a/span[@class='fas fa-trash-alt']";
        private const string NAME_LP_CART = "//*[@id=\"report-form\"]/div[2]/div[1]/h3";
        private const string SHORTCOMMENT_INPUT = "/html/body/div[4]/div/div/form/div[2]/div[5]/div[1]/div/div/input";
        private const string LABEL_PAGE_INPUT= "//*[@id=\"labelPageNbr\"]";
        private const string PRINT_POSITIONS_CHECKBOX = "//*[@id=\"printPositions\"]";
        private const string CLOSE_BTN = "//*[@id=\"report-form\"]/div[3]/button[1]";
        private const string ROWSNUMBER = "//*[@id=\"rowsNumber\"]";
        private const string COLUMNNUMBER = "//*[@id=\"colsNumber\"]";

        private const string TROLLEY_NAME = "//*[@id=\"LPCartDetailsTable\"]/tbody/tr[{0}]/td[3]/input";
        private const string TROLLEY_CODE = "//*[@id=\"LPCartDetailsTable\"]/tbody/tr[{0}]/td[4]/input";
        private const string TROLLEY_LOCATION = "//*[@id=\"LPCartDetailsTable\"]/tbody/tr[{0}]/td[5]/input";
        private const string GALLEY_CODE = "//*[@id=\"LPCartDetailsTable\"]/tbody/tr[{0}]/td[6]/input";
        private const string GALLEY_LOCATION = "//*[@id=\"LPCartDetailsTable\"]/tbody/tr[{0}]/td[7]/input";
        private const string EQUIPEMENT = "//*[@id=\"LPCartDetailsTable\"]/tbody/tr[{0}]/td[8]/input";
        private const string ROUTES = "//*[@id=\"tabContentDetails\"]/div[5]/div[2]/div";
        private const string FIELD_NAME = "//*[@id=\"tb-new-lpcart-name\"]";
        private const string LIST_ROUTES = "//*[@id=\"tabContentDetails\"]/div[5]/div[2]/div/div[1]";
        private const string MODAL_FORM = "//*[@id=\"report-form\"]";


        // ____________________________________ Variables ___________________________________________

        // Menu général
        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;

        [FindsBy(How = How.Id, Using = ADDTROLLEY_BTN)]
        private IWebElement _createBtnTrolley;

        [FindsBy(How = How.XPath, Using = FLIGHT_TAB)]
        private IWebElement _flightTab;

        [FindsBy(How = How.XPath, Using = DUPLICATE_CARTS_TROLLEY)]
        private IWebElement _duplicateCartTrolley;

        [FindsBy(How = How.Id, Using = FIRST_ARROW_DOWN_BTN)]
        private IWebElement _firstArrowDown;

        [FindsBy(How = How.Id, Using = EDIT_CART_SCHEME)]
        private IWebElement _editCartScheme;

        [FindsBy(How = How.XPath, Using = LEG_NO)]
        private IWebElement _legno;

        [FindsBy(How = How.XPath, Using = TROLLEYNAME)]
        private IWebElement _trolleyname;

        [FindsBy(How = How.XPath, Using = TROLLEYCODE)]
        private IWebElement _trolleycode;

        [FindsBy(How = How.XPath, Using = TROLLEYLOC)]
        private IWebElement _trolleyloc;

        [FindsBy(How = How.XPath, Using = GALLEYCODE)]
        private IWebElement _galleycode;

        [FindsBy(How = How.XPath, Using = GALLEYLOC)]
        private IWebElement _galleyloc;

        [FindsBy(How = How.XPath, Using = EQUIPMENT)]
        private IWebElement _equipment;

        [FindsBy(How = How.XPath, Using = SHORTCOMMENT)]
        private IWebElement  _shortcomment;

        [FindsBy(How = How.XPath, Using = NAME_LP_CART)]
        private IWebElement _name_lp_cart;
        [FindsBy(How = How.XPath, Using = PLUSROW)]
        private IWebElement _plusrow;

        [FindsBy(How = How.XPath, Using = SHORTCOMMENT_INPUT)]
        private IWebElement _shortComment;

        [FindsBy(How = How.XPath, Using = LABEL_PAGE_INPUT)]
        private IWebElement _labelPageInput;

        [FindsBy(How = How.XPath, Using = PRINT_POSITIONS_CHECKBOX)]
        private IWebElement _print_positions_checkbox;

        [FindsBy(How = How.XPath, Using = ROWSNUMBER)]
        private IWebElement _row_number;

        [FindsBy(How = How.XPath, Using = COLUMNNUMBER)]
        private IWebElement _col_number;
        [FindsBy(How = How.XPath, Using = EDITROW)]
        private IWebElement _editrow;

         [FindsBy(How = How.XPath, Using = DELETEROW)]
        private IWebElement _deleterow;

        [FindsBy(How = How.XPath, Using = CLOSE_BTN)]
        private IWebElement _close_btn;
        // __________________________________ Filtres ________________________________________



        [FindsBy(How = How.Id, Using = SEARCH_FILTER)]
        private IWebElement _search;

        [FindsBy(How = How.Id, Using = SORT_BY_FILTER)]
        private IWebElement _sortBy;

        [FindsBy(How = How.XPath, Using = EXTENDED_BTN)]
        private IWebElement _extendedButton;

        [FindsBy(How = How.Id, Using = PLUS_BTN_ID)]
        private IWebElement _plusbtn;
        [FindsBy(How = How.Id, Using = MODAL_FORM)]
        private IWebElement _modalForm;

        // ___________________________________  Méthodes ___________________________________________
        public LpCartCartDetailPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        public enum FilterType
        {
            Search,
            SortBy,
        }

        public void Filter(FilterType filterType, object value)
        {
            Actions action = new Actions(_webDriver);

            switch (filterType)
            {
                case FilterType.Search:
                    _search = WaitForElementIsVisible(By.Id(SEARCH_FILTER));
                    _search.SetValue(ControlType.TextBox, value);
                    WaitPageLoading();
                    break;
                case FilterType.SortBy:
                    _sortBy = WaitForElementIsVisible(By.Id(SORT_BY_FILTER));
                    _sortBy.Click();
                    WaitForLoad();
                    var element = WaitForElementIsVisible(By.XPath("//option[contains(@value,'" + value + "')]"));
                    _sortBy.SetValue(ControlType.DropDownList, element.Text);
                    _sortBy.Click();
                    WaitPageLoading();
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);
            }

            WaitPageLoading();
            WaitForLoad();
        }

        public void ClickAddtrolley()
        {
            WaitForLoad();
            _createBtnTrolley = WaitForElementIsVisible(By.Id(ADDTROLLEY_BTN));
            _createBtnTrolley.Click();
            WaitForLoad();
        }

        public void ClickFirstArrowDown()
        {
            _firstArrowDown = WaitForElementIsVisible(By.XPath(FIRST_ARROW_DOWN_BTN_DEV));
            _firstArrowDown.Click();

            WaitForLoad();
        }


        public void AddTrolley(string value, string comment = null)
        {
            WaitForLoad();
            IWebElement itemTrolley;
            itemTrolley = WaitForElementIsVisible(By.XPath("//span[contains(@class, 'fas fa-trash-alt')]"));
            Actions action = new Actions(_webDriver);
            action.MoveToElement(itemTrolley).Perform();

            var inputs = _webDriver.FindElements(By.XPath(TROLLEY_INPUTS));


            foreach (var input in inputs)
            {
                if (input.GetAttribute("id") == "ShortComment")
                {
                    if (comment != null)
                    {
                        input.SetValue(ControlType.TextBox, comment);
                    }
                }
                else if (input.GetAttribute("value") != "1")
                {
                    input.SetValue(ControlType.TextBox, value);
                }
            }
            WaitPageLoading();
            WaitForLoad();
        }

        public void DeleteAllLpCartScheme()
        {
            int tot = 0;
            tot = _webDriver.FindElements(By.XPath("//span[contains(@class, 'fas fa-trash-alt')]/../../a")).Count;

            for (int i = 0; i < tot; i++)
            {
                IWebElement trashIcon;
                trashIcon = _webDriver.FindElement(By.XPath(string.Format("//*[@id=\"LPCartDetailsTable\"]/tbody/tr[2]//span[contains(@class, 'fas fa-trash-alt')]")));

                trashIcon.Click();
                var delete = WaitForElementIsVisible(By.XPath(DELETE_BTN_TROLLEY));
                delete.Click();
                WaitForLoad();
            }

            WaitForLoad();
        }


        public LpCartPage BackToList()
        {
            _backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            _backToList.Click();
            WaitForLoad();

            return new LpCartPage(_webDriver, _testContext);
        }

        public LpCartFlightDetailPage LpCartFlightDetailPage()
        {

            _flightTab = WaitForElementIsVisible(By.Id(FLIGHT_TAB));
            _flightTab.Click();
            WaitForLoad();

            return new LpCartFlightDetailPage(_webDriver, _testContext);
        }

        public LpCartGeneralInformationPage LpCartGeneralInformationPage()
        {

            _flightTab = WaitForElementIsVisible(By.Id("hrefTabContentInformations"));
            _flightTab.Click();
            WaitForLoad();

            return new LpCartGeneralInformationPage(_webDriver, _testContext);
        }

        public LpCartDuplicateTrolleyModalPage DuplicateTrolley()
        {
            ShowExtendedMenu();

            _duplicateCartTrolley = WaitForElementIsVisible(By.XPath(DUPLICATE_CARTS_TROLLEY));
            _duplicateCartTrolley.Click();
            WaitForLoad();

            return new LpCartDuplicateTrolleyModalPage(_webDriver, _testContext);
        }

        public override void ShowExtendedMenu()
        {
            var actions = new Actions(_webDriver);

            _extendedButton = WaitForElementIsVisible(By.XPath(EXTENDED_BTN));
            actions.MoveToElement(_extendedButton).Perform();
            WaitLoading();
        }

        public bool IstrolleySchemaExist()
        {
            bool IstrolleySchemaExist = false;

            var lineCount = WaitForElementIsVisible(By.XPath(LINE_COUNT));
            var columnCount = WaitForElementIsVisible(By.XPath(COLUMN_COUNT));

            if (lineCount.Text != "0" && columnCount.Text != "0")
                IstrolleySchemaExist = true;

            return IstrolleySchemaExist;
        }

        public bool IsTrolleyName(string trolleyName)
        {
            bool isGoodTrolleyName = true;

            var trolleyNames = _webDriver.FindElements(By.XPath(TROLLEY_NAMES));

            foreach (var elm in trolleyNames)
            {
                if (elm.Text != trolleyName)
                    return false;
            }
            return isGoodTrolleyName;
        }

        public List<string> GetAllTrolleyNames()
        {
            List<string> names = new List<string>();

            var trolleyNames = _webDriver.FindElements(By.XPath(TROLLEY_NAMES));

            foreach (var elm in trolleyNames)
            {
                names.Add(elm.Text);
            }

            return names;
        }
        public LpCartEditLpCartSchemeModal EditCartScheme()
        {
            WaitPageLoading();
            WaitForLoad();
            _editCartScheme = WaitForElementIsVisible(By.XPath(EDIT_CART_SCHEME));
            _editCartScheme.Click();
            WaitForLoad();
            return new LpCartEditLpCartSchemeModal(_webDriver, _testContext);
        }

        public void FlightsTab()
        {
            WaitForElementIsVisible(By.XPath(FLIGHTS_TAB)).Click();
        }

        public string GetFirstFlightNumber()
        {
            return WaitForElementIsVisible(By.XPath(FIRST_FLIGHT_NUMBER_FLIGHTS_TAB)).Text;
        }

        public string GetFirstFlightDate()
        {
            return WaitForElementIsVisible(By.XPath(FIRST_FLIGHT_DATE_FLIGHTS_TAB)).Text;
        }
        public string GetFirstFlightSiteFrom()
        {
            return WaitForElementIsVisible(By.XPath(FIRST_FLIGHT_SITE_FROM_FLIGHTS_TAB)).Text;
        }
        public string GetFirstFlightSiteTo()
        {
            return WaitForElementIsVisible(By.XPath(FIRST_FLIGHT_SITE_TO_FLIGHTS_TAB)).Text;
        }

        public void AddTrolleyCustom(string leg,string code, string name, int i)
        {
            WaitForLoad();
            var line = WaitForElementIsVisible(By.XPath(string.Format(LINE, i + 2)));
            line.Click();
            var legIn = WaitForElementIsVisible(By.XPath(string.Format(LEG_INPUT, i + 2)));
            legIn.SetValue(ControlType.TextBox, leg);
            WaitLoading();
            var nameIn = WaitForElementIsVisible(By.XPath(string.Format(TROLLEY_NAME_INPUT, i + 2)));
            nameIn.SetValue(ControlType.TextBox, name);
            WaitLoading();
            var codeIn = WaitForElementIsVisible(By.XPath(string.Format(TROLLEY_CODE_INPUT, i + 2)));
            codeIn.SetValue(ControlType.TextBox, code);

            Thread.Sleep(2000);
            WaitLoading();
            WaitPageLoading();
        }

        public bool VerifyFilterByLegOrder()
        {
            WaitForLoad();
            var orders = _webDriver.FindElements(By.XPath(ORDERS));
            bool isInOrder = orders.ToList().SequenceEqual(orders.OrderBy(x => x.Text));
            return isInOrder;
        }
        public bool VerifyFilterBycode()
        {
            WaitForLoad();
            var codes = _webDriver.FindElements(By.XPath(CODES));
            bool isInOrder = codes.ToList().SequenceEqual(codes.OrderBy(x => x.Text));
            return isInOrder;
        }
        public bool VerifyFilterByName()
        {
            WaitForLoad();
            var names = _webDriver.FindElements(By.XPath(CODES));
            bool isInOrder = names.ToList().SequenceEqual(names.OrderBy(x => x.Text));
            return isInOrder;
        }

        public LpCartFlightDetailPage ClickOnFlightTab()
        {
            var href = WaitForElementIsVisible(By.Id("hrefTabContentFlight"));
            href.Click();
            WaitForLoad();

            return new LpCartFlightDetailPage(_webDriver, _testContext);
        }

        public void GoToCartsTab()
        {
            var cartTab = WaitForElementExists(By.XPath(CARTS_TAB));
            cartTab.Click();
            WaitForLoad();
        }

        public bool VerifySortByName()
        {
            WaitForLoad();
            GoToCartsTab();

            Filter(FilterType.SortBy, "Name");
            var namesList = _webDriver.FindElements(By.XPath(CARTS_NAME_LIST));

            for (int i = 1; i < namesList.Count; i++)
            {
                if (string.Compare(namesList[i - 1].Text, namesList[i].Text, StringComparison.Ordinal) > 0)
                {
                    return false;
                }
            }

            return true;
        }

        public bool CheckSortEnum()
        {
            WaitForLoad();
            var option1Value = WaitForElementExists(By.XPath("//*[@id=\"cbSortBy\"]/option[1]")).GetAttribute("value");
            var option1Text = WaitForElementExists(By.XPath("//*[@id=\"cbSortBy\"]/option[1]")).Text;
            var option2Value = WaitForElementExists(By.XPath("//*[@id=\"cbSortBy\"]/option[2]")).GetAttribute("value");
            var option2Text = WaitForElementExists(By.XPath("//*[@id=\"cbSortBy\"]/option[2]")).Text;
            var option3Value = WaitForElementExists(By.XPath("//*[@id=\"cbSortBy\"]/option[3]")).GetAttribute("value");
            var option3Text = WaitForElementExists(By.XPath("//*[@id=\"cbSortBy\"]/option[3]")).Text;

            if (option1Value == "LegAndOrder" && option1Text == "Sort By Leg and order" && option2Value == "Code"
                && option2Text == "Sort by code" && option3Value == "Name" && option3Text == "Sort by name")
            {
                return true;
            }

            return false;
        }

        public bool VerifySortByCode()
        {
            WaitForLoad();
            GoToCartsTab();

            Filter(FilterType.SortBy, "Code");
            var codesList = _webDriver.FindElements(By.XPath(CARTS_CODE_LIST));

            for (int i = 1; i < codesList.Count; i++)
            {
                if (string.Compare(codesList[i - 1].Text, codesList[i].Text, StringComparison.Ordinal) > 0)
                {
                    return false;
                }
            }

            return true;
        }

        public void Click_PlusBTN()
        {
            _plusbtn = WaitForElementIsVisible(By.Id(PLUS_BTN_ID));
            _plusbtn.Click();
            WaitForLoad();
        }

        public bool VerifyDefaultFields()
        {
            int numberrow;

            _legno = WaitForElementIsVisible(By.XPath(LEG_NO));
            var leg = _legno.GetAttribute("value");

            _trolleyname = WaitForElementIsVisible(By.XPath(TROLLEYNAME));
            var trolleyname = _trolleyname.GetAttribute("value");

            _trolleycode = WaitForElementIsVisible(By.XPath(TROLLEYCODE));
            var trolleycode = _trolleycode.GetAttribute("value");

            _trolleyloc = WaitForElementIsVisible(By.XPath(TROLLEYLOC));
            var trolleyloc = _trolleyloc.GetAttribute("value");

            _galleycode = WaitForElementIsVisible(By.XPath(GALLEYCODE));
            var galleycode = _galleycode.GetAttribute("value");

            _galleyloc = WaitForElementIsVisible(By.XPath(GALLEYLOC));
            var galleyloc = _galleyloc.GetAttribute("value");

            _equipment = WaitForElementIsVisible(By.XPath(EQUIPMENT));
            var equipement = _galleyloc.GetAttribute("value");

            _shortcomment = WaitForElementIsVisible(By.XPath(SHORTCOMMENT));
            var shortcomment = _shortcomment.GetAttribute("value");

            bool successParseInt = int.TryParse(leg, out numberrow);

            if ( successParseInt && numberrow == 1 

                && string.IsNullOrEmpty(trolleyname)
                && string.IsNullOrEmpty(trolleycode)
                && string.IsNullOrEmpty(trolleyloc)
                && string.IsNullOrEmpty(galleycode)
                && string.IsNullOrEmpty(galleyloc)
                && string.IsNullOrEmpty(equipement)
                && string.IsNullOrEmpty(shortcomment)                
            ) 
            {
                return true;
            }

            return false;
        }

        public bool VerifyLpCartName(string name)
        {
            var colorGris = "rgba(185, 185, 185, 1)";
            _name_lp_cart = WaitForElementExists(By.XPath(NAME_LP_CART));
            var style = WaitForElementExists(By.XPath(NAME_LP_CART)).GetCssValue("color");
            if (_name_lp_cart.Text == name.ToUpper() && style==colorGris)
            {
                return true;
            }
            return false;
        }
        public LpCartCartDetailPage CreateNewRowLpCart(int leg, string trolleyname, string trolleycode, string trolleyloc, string galleycode, string galleyloc, string equipement, string shortcomment )
        {         
            _legno = WaitForElementIsVisible(By.XPath(LEG_NO));

            _legno.SetValue(ControlType.TextBox, leg.ToString());

            _trolleyname = WaitForElementIsVisible(By.XPath(TROLLEYNAME));

            _trolleyname.SetValue(ControlType.TextBox, trolleyname);

            _trolleycode = WaitForElementIsVisible(By.XPath(TROLLEYCODE));

            _trolleycode.SetValue(ControlType.TextBox, trolleycode);
          
            _trolleyloc = WaitForElementIsVisible(By.XPath(TROLLEYLOC));

            _trolleyloc.SetValue(ControlType.TextBox, trolleyloc);

            _galleycode = WaitForElementIsVisible(By.XPath(GALLEYCODE));

            _galleycode.SetValue(ControlType.TextBox, galleycode);

            _galleyloc = WaitForElementIsVisible(By.XPath(GALLEYLOC));

            _galleyloc.SetValue(ControlType.TextBox, galleyloc);

            _equipment = WaitForElementIsVisible(By.XPath(EQUIPMENT));

            _equipment.SetValue(ControlType.TextBox, equipement);

            _shortcomment = WaitForElementIsVisible(By.XPath(SHORTCOMMENT));

            _shortcomment.SetValue(ControlType.TextBox, shortcomment);

            return new LpCartCartDetailPage(_webDriver, _testContext);
        }


        public string GetElementPlus()
        {
            _plusrow = WaitForElementIsVisible(By.XPath(PLUSROW));
            return _plusrow.GetAttribute("class");
        }

        public bool VerifyChangePlusToEdit(string elementplus)
        {     
            _editrow = WaitForElementIsVisible(By.XPath(EDITROW));
            var editrow = _editrow.GetAttribute("class");

              _deleterow = WaitForElementIsVisible(By.XPath(DELETEROW));
            var deleterow = _deleterow.GetAttribute("class");

            if (elementplus == editrow && string.IsNullOrEmpty(deleterow))
            {
                return false;
            }
            return true;

        }
        public bool VerifyShortComment(string shortComment)
        {
            
            _shortComment = WaitForElementExists(By.XPath(SHORTCOMMENT_INPUT));
            var shortCommentValue = _shortComment.GetAttribute("value");
            if (shortCommentValue == shortComment)
            {
                return true;
            }
            return false;
        }

        public bool VerifyLabelPage()
        {
            _labelPageInput = WaitForElementExists(By.XPath(LABEL_PAGE_INPUT));
            var labelValue = _labelPageInput.GetAttribute("value");
            if (int.Parse(labelValue) == 1)
            {
                return true;
            }
            return false;
        }
        public bool VerifyPrintPositionsChecked()
        {
            _print_positions_checkbox = WaitForElementExists(By.XPath(PRINT_POSITIONS_CHECKBOX));
            var valeur = _print_positions_checkbox.GetAttribute("value");
            if (valeur=="true")
            {
                return true;
            }
            return false;

        }
        public bool VerifyRowsAndCols()
        {
            _col_number = WaitForElementExists(By.XPath(COLUMNNUMBER));
            _row_number = WaitForElementExists(By.XPath(ROWSNUMBER));
            var columns = _col_number.GetAttribute("value");
            var rows = _row_number.GetAttribute("value");
            if (int.Parse(columns) == 0 && int.Parse(rows)==0 )
            {
                return true;
            }
            return false;

        }
        public void CloseModalEdit()
        {
            _close_btn = WaitForElementIsVisible(By.XPath(CLOSE_BTN));
            _close_btn.Click();
            WaitForLoad();
        }
        public string GetRowsValueFromEditModal()
        {
            EditCartScheme();
            _row_number = WaitForElementExists(By.XPath(ROWSNUMBER));
            var rows = _row_number.GetAttribute("value");
            CloseModalEdit();
            return rows;
        }
        public string GetColumnsValueFromEditModal()
        {
            EditCartScheme();
            _col_number = WaitForElementExists(By.XPath(COLUMNNUMBER));
            var columns = _col_number.GetAttribute("value");
            CloseModalEdit();
            return columns;
        }


        public void FillNewTrolleyLine(string name, string code, string loc, string codeGalley, string locGalley, string equip)
        {
            WaitForLoad();
            var trolleyList = _webDriver.FindElements(By.XPath("//*[@id=\"LPCartDetailsTable\"]/tbody/tr[*]/td[1]"));
            var totalLignes = trolleyList.Count;
            // ligne add
            var trolleyName = WaitForElementIsVisible(By.XPath(String.Format(TROLLEY_NAME, totalLignes + 1)));
            trolleyName.SetValue(ControlType.TextBox, name);
            WaitForLoad();
            var trolleyCode = WaitForElementIsVisible(By.XPath(String.Format(TROLLEY_CODE, totalLignes + 1)));
            trolleyCode.SetValue(ControlType.TextBox, code);
            WaitForLoad();
            var trolleyLocation = WaitForElementIsVisible(By.XPath(String.Format(TROLLEY_LOCATION, totalLignes + 1)));
            trolleyLocation.SetValue(ControlType.TextBox, loc);
            WaitForLoad();
            var galleyCode = WaitForElementIsVisible(By.XPath(String.Format(GALLEY_CODE, totalLignes + 1)));
            galleyCode.SetValue(ControlType.TextBox, codeGalley);
            WaitForLoad();
            var galleyLocation = WaitForElementIsVisible(By.XPath(String.Format(GALLEY_LOCATION, totalLignes + 1)));
            galleyLocation.SetValue(ControlType.TextBox, locGalley);
            WaitForLoad();
            var equipment = WaitForElementIsVisible(By.XPath(String.Format(EQUIPEMENT, totalLignes + 1)));
            equipment.SetValue(ControlType.TextBox, equip);
            equipment.SendKeys(Keys.Tab);
            WaitPageLoading();
            Thread.Sleep(2000);
            WaitForLoad();
        }

        public void ClickRemoveTrolley(string name)
        {
            Filter(FilterType.Search, name);
            if (isElementExists(By.XPath("//*[@id='LPCartDetailsTable']/tbody/tr[2]/td[3]/span"))) {
                var trash = WaitForElementIsVisible(By.XPath("//*/span[contains(@class,'trash')]"));
                trash.Click();
                WaitForLoad();
                //modal
                var confirm = WaitForElementIsVisible(By.XPath("//*/button[@value='Delete']"));
                confirm.Click();
                WaitForLoad();
            }
            Filter(FilterType.Search, "");
        }

        public bool CheckSearchByCode(string code)
        {
            var trolleyCodeList = _webDriver.FindElements(By.XPath("//*[@id=\"LPCartDetailsTable\"]/tbody/tr[*]/td[4]")).Select(e => e.Text);

            return trolleyCodeList.Where(s => s.Equals(code)).Count() == trolleyCodeList.Count();
            }
        public bool AreRoutesPresent()
        {
            var routes = _webDriver.FindElements(By.XPath(ROUTES));
            return routes.Count > 0;
        }
        public void ChangeLpCartName(string newName)
        {
            var nameField = WaitForElementIsVisible(By.XPath(FIELD_NAME));
            nameField.Clear();
            nameField.SendKeys(newName);
            WaitPageLoading();
        }

        public string GetName()
        {
            var nameField = WaitForElementIsVisible(By.XPath(FIELD_NAME));
            return nameField.GetAttribute("value");
        }

        public bool IsEditPopupDisplayed()
        {
             _modalForm = _webDriver.FindElement(By.XPath(MODAL_FORM));
            return _modalForm.Displayed;
        }
    }
}
