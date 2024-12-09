using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using System.Threading;
using System;
using System.Linq;
using System.Collections.Generic;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Menus;
using OpenQA.Selenium.Support.UI;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenQA.Selenium.Support.PageObjects;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Menus.FoodCost
{
    public class FoodCostPage : PageBase
    {
        private const string  FROM_DATE_INPUT= "date-picker-start";
        private const string  TO_DATE_INPUT= "date-picker-end";
        private const string  SITE_INPUT= "drop-down-site";
        private const string  CUSTOMERS_INPUT= "SelectedCustomers_ms";
        private const string  GUEST_TYPES_INPUT= "SelectedGuestTypes_ms";
        private const string  MEAL_TYPES_INPUT= "SelectedMealTypes_ms";
        private const string  VALIDATE= "btn-validate";
        private const string  UNCHECK_CUSTOMERS_ALL= "/html/body/div[10]/div/ul/li[2]/a";
        private const string  UNCHECK_GUEST_TYPES_ALL= "/html/body/div[11]/div/ul/li[2]/a/span[2]";
        private const string UNCHECK_MEAL_TYPES_ALL = "/html/body/div[12]/div/ul/li[2]/a/span[2]";
        private const string  EDIT_PENCIL = "//*[starts-with(@id,\"content\")]/div/div[1]/div[contains(text(),'{0}')]/../div[3]";


        private const string CUSTOMER = "//*[@id=\"foodcostResultsRecap\"]/div[2]/div";
        private const string FROM = "//*[@id=\"foodcostResultsRecap\"]/div[2]/div";
        private const string FROM_DATE_FILTER = "/html/body/div[3]/div/div/div[2]/div[1]/div[1]";
        private const string TO_DATE_FILTER = "/html/body/div[3]/div/div/div[2]/div[1]/div[2]";
        private const string FIRST_DATE = "//*[@id=\"list-item-with-action\"]/div[1]/div/div[2]/div[1]/p[1]";
        private const string SUNDAY_COLUMN = "/html/body/div[3]/div/div/div[3]/div[*]/div[1]/div[2]/div[7]";
        private const string LAST_DATE = "/html/body/div[3]/div/div/div[3]/div[1]/div[2]/div[1]";
        private const string EDIT_BUTTONS = "/html/body/div[3]/div/div/div[3]/div[*]/div[2]/div[*]/div[2]/a/span";
        private const string UNFOLD_ALL_BUTTONS = "//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]";
        private const string UNFOLD_BUTTON = "//*/div[contains(@class,'row clickable')]";
        private const string MENUS_NAMES = "//div[@class=\"col-md-2 menu-name\"]";
        private const string MENU_NAME_VISIBLE = "//*/div[@class=\"col-md-2 menu-name\" and contains(text(),'{0}')]";

        private const string MEAL_TYPES = "/html/body/div[3]/div/div/div[3]/div[*]/div[1]/div/div[2]";
        private const string MEAL_TYPE_FILTER_HEADER = "/html/body/div[3]/div/div/div[3]/div[2]/div[1]/div/div[2]";
        private const string GUEST_TYPE_FILTER_HEADER = "/html/body/div[3]/div/div/div[2]/div[3]/div";
        private const string CLICK_FOOD_COST = "/html/body/div[3]/div/div/div[3]/div[2]/div[1]/div";
        private const string MENU_NAME = "/html/body/div[3]/div/div/div[3]/div[2]/div[2]/div/div[1]/div[1]";
        private const string PRICE = "/html/body/div[3]/div/div/div[3]/div[2]/div[1]/div/div[3]/div[3]/div";
        private const string VALIDATEBTN = "btn-validate";
        

        public FoodCostPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        public enum FilterType
        {
            From,
            To,
            Site,
            Customer,
            GuestType,
            MealType,
        }

        [FindsBy(How = How.XPath, Using = MENU_NAME)]
        private IWebElement _menuName;
        [FindsBy(How = How.Id, Using = VALIDATEBTN)]
        private IWebElement _validateBTN;

        [FindsBy(How = How.XPath, Using = PRICE)]
        private IWebElement _price;

        [FindsBy(How = How.XPath, Using = CLICK_FOOD_COST)]
        private IWebElement _clickFoodCost;
        

        public void Filter(FilterType filterType, object value)
        {
            switch (filterType)
            {
                case FilterType.From:
                    var fromInput  = WaitForElementIsVisible(By.Id(FROM_DATE_INPUT));
                    fromInput.SetValue(ControlType.DateTime, value);
                    fromInput.SendKeys(Keys.Enter);
                    WaitPageLoading();
                    break;
                case FilterType.To:
                    var toInput = WaitForElementIsVisible(By.Id(TO_DATE_INPUT));
                    toInput.SetValue(ControlType.DateTime, value);
                    toInput.SendKeys(Keys.Tab);
                    WaitLoading();
                    break;
                case FilterType.Site:
                    var siteInput = WaitForElementIsVisible(By.Id(SITE_INPUT));
                    siteInput.Click();
                    var elementS = WaitForElementIsVisible(By.XPath("//option[text()='"+value+"']"));
                    siteInput.SetValue(ControlType.DropDownList, elementS.Text);
                    siteInput.Click();
                    WaitPageLoading();
                    break;
                case FilterType.Customer:
                    var customerInput = WaitForElementIsVisible(By.Id(CUSTOMERS_INPUT));
                    customerInput.Click();
                    var uncheckCustomersAll = WaitForElementIsVisible(By.XPath(UNCHECK_CUSTOMERS_ALL));
                    uncheckCustomersAll.Click();
                    var elementC = WaitForElementIsVisible(By.XPath(string.Format("/html/body/div[10]/ul/li[*]/label/span[text()='{0}']", value)));
                    elementC.Click();
                    customerInput.Click();
                    WaitPageLoading();
                    break;
                case FilterType.GuestType:
                    var guestTypeInput = WaitForElementIsVisible(By.Id(GUEST_TYPES_INPUT));
                    guestTypeInput.Click();
                    var uncheckGuestTypesAll = WaitForElementIsVisible(By.XPath(UNCHECK_GUEST_TYPES_ALL));
                    uncheckGuestTypesAll.Click();
                    var gt = WaitForElementIsVisible(By.XPath("/html/body/div[11]/ul/li[*]/label/span[text()='" + value + "']"));
                    gt.Click();
                    break;
                case FilterType.MealType:
                    var mealTypeInput = WaitForElementIsVisible(By.Id(MEAL_TYPES_INPUT));
                    mealTypeInput.Click();
                    var uncheckMealTypesAll = WaitForElementIsVisible(By.XPath(UNCHECK_MEAL_TYPES_ALL));
                    uncheckMealTypesAll.Click();
                    var mt = WaitForElementIsVisible(By.XPath("/html/body/div[12]/ul/li[*]/label/span[text()='" + value + "']"));
                    mt.Click();
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);

            }

            WaitPageLoading();
            var validate = WaitForElementIsVisible(By.Id(VALIDATE));
            validate.Click();
            Thread.Sleep(2000);
        }

        public bool CheckCustomerFilter(string customer)
        {
            var customerFilter = WaitForElementIsVisible(By.XPath(CUSTOMER));
            if(customer.Contains(customerFilter.Text))
            {
                return true;
            }
            return false;
        }

        public bool CheckFromFilter(DateTime from)
        {
            var mondayDate = GetMondayOfWeekDate(from);
            var fromDateFilter = WaitForElementIsVisible(By.XPath(FROM_DATE_FILTER));
            if(fromDateFilter.Text.Contains(mondayDate.ToString("dd/MM/yyyy"))) {
                return true;
            }
            return false;   
        }
        public bool CheckToFilter(DateTime to)
        {
            var sundayDate = GetNextSundayDate(to);
            var fromDateFilter = WaitForElementIsVisible(By.XPath(TO_DATE_FILTER));
            if (fromDateFilter.Text.Contains(sundayDate.ToString("dd/MM/yyyy")))
            {
                return true;
            }
            return false;
        }

        public DateTime GetMondayOfWeekDate(DateTime date)
        {
            DayOfWeek currentDayOfWeek = date.DayOfWeek;
            int diff = (7 + (currentDayOfWeek - DayOfWeek.Monday)) % 7;
            return date.AddDays(-diff);
        }

        public DateTime GetNextSundayDate(DateTime date)
        {
            DayOfWeek currentDayOfWeek = date.DayOfWeek;
            int diff = DayOfWeek.Saturday - currentDayOfWeek + 1;
            return date.AddDays(diff);
        }

        public DateTime GetFirstDate()
        {
            WaitPageLoading();
            var firstDate = WaitForElementIsVisible(By.XPath(FIRST_DATE)).Text;
            return DateTime.ParseExact(firstDate.ToString(),"dd/MM/yyyy",null);
        }
        public bool CompareDates(DateTime date1 , DateTime date2)
        {
            if (date1.Day == date2.Day && date1.Month == date2.Month && date1.Year == date2.Year)
            {
                return true;
            } 
            return false;
        }
        public DateTime GetLastDate()
        {
            var sundayDates = _webDriver.FindElements(By.XPath(SUNDAY_COLUMN));
            var datesOnly = new List<string>();
            foreach (var date in sundayDates)
            {
               var dateOnly = date.Text.Substring(0, 10);
                datesOnly.Add(dateOnly);
            }

            var lastDate = datesOnly.ElementAt(datesOnly.ToList().Count - 1);
            return DateTime.ParseExact(lastDate.ToString(), "dd/MM/yyyy", null);
        }
        public List<string> GetMenusNames()
        {
            var menus = _webDriver.FindElements(By.XPath(MENUS_NAMES));
            //var nameAfterClean = Regex.Replace(menuName, @"[\d(),]+", "").Trim();
            return menus.Select(e=>e.GetAttribute("innerHTML").Substring(0, e.GetAttribute("innerHTML").IndexOf("<")).Trim()).ToList<string>();
        }

        public void ValidateFilters()
        {
            WaitForLoad();
            var validate = WaitForElementIsVisible(By.Id(VALIDATE));
            validate.Click();
            WaitForLoad();
        }

        public void Uncollapse(int offset)
        {
            var collapsed = _webDriver.FindElements(By.XPath(UNFOLD_BUTTON));
            collapsed[offset].Click();
            Thread.Sleep(2000);
            WaitForLoad();
        }

        public MenusWeekViewPage EditPencil(string name)
        {
            var pencil = WaitForElementIsVisible(By.XPath(string.Format(EDIT_PENCIL,name)));
            pencil.Click();

            // nouveau onglet !!!
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);
            WaitForLoad();
            return new MenusWeekViewPage(_webDriver, _testContext);
        }

        public bool IsMenuNameVisible(string menuName)
        {
            return isElementVisible(By.XPath(string.Format(MENU_NAME_VISIBLE, menuName)));
        }

        public bool GuestTypeExistInHeaderFilter(string guestType)
        {
            var guestTypeFilterValue = WaitForElementIsVisible(By.XPath(GUEST_TYPE_FILTER_HEADER));
            if (guestTypeFilterValue.Text.Contains(guestType))
                return true;
            return false;
        }
        public bool MealTypeExistFilter(string mealType)
        {
            var mealTypesList = _webDriver.FindElements(By.XPath(MEAL_TYPES));
            if(!mealTypesList.Select(p=>p.Text).Contains(mealType))
            {
                return false;  
            }
            return true;

        }

        public void FilterFromToDate(FilterType filterType, object value)
        {
            switch (filterType)
            {

                case FilterType.From:
                    var fromInput = WaitForElementIsVisible(By.Id(FROM_DATE_INPUT));
                    fromInput.SetValue(ControlType.DateTime, value);
                    fromInput.SendKeys(Keys.Enter);
                    WaitPageLoading();
                    break;

                case FilterType.To:
                    var toInput = WaitForElementIsVisible(By.Id(TO_DATE_INPUT));
                    toInput.SetValue(ControlType.DateTime, value);
                    toInput.SendKeys(Keys.Tab);
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);

            }
            WaitPageLoading();
            WaitForLoad();
        }
        public void Validate()
        {
            WaitLoading();
            _validateBTN = WaitForElementIsVisible(By.Id(VALIDATEBTN));
            _validateBTN.Click();
            WaitLoading();
        }

        public bool VerifiedMenuAndCout()
        {
 
            var _clickFood = WaitForElementExists(By.XPath(CLICK_FOOD_COST));
            _clickFood.Click();
            WaitForLoad();

            _menuName = WaitForElementIsVisible(By.XPath(MENU_NAME));
            _price = WaitForElementIsVisible(By.XPath(PRICE));
            WaitForLoad();


            if ((_menuName.Text != null)&&(_price.Text != null ))
            {
                return true;
            }
            else return false;


        }
    }
}