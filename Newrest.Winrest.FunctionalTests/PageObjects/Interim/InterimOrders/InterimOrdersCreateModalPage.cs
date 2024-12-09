using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Interim.InterimOrders
{
    public class InterimOrdersCreateModalPage : PageBase
    {
        private const string INTERIM_NUMBER = "tb-new-purchase-order-number";
        private const string SITE = "SelectedSiteId";
        private const string SUPPLIER = "drop-down-suppliers";
        private const string DELIVERY_DATE = "datapicker-new-interim-order-delivery";
        private const string ACTIVATED = "InterimOrder_IsActive";
        private const string DELIVERY_LOCATION = "SelectedSitePlaceId";
        private const string SUBMIT = "btn-submit-form-create-interim-order";
        private const string VALIDATION_ERROR = "//*/span[contains(@class,'text-danger') and contains(@class,'field-validation-error')]";
        private const string COMMENT = "InterimOrder_Comment";
        private const string COPY_FROM = "checkBoxCopyFrom";
        private const string NUMBER_TO_COPY = "form-copy-from-tbSearchName";
        private const string FILTERED_NUMBER = "//*[@id=\"table-copy-from-io\"]/tbody/tr[2]/td[2]";
        private const string LIST_VALIDATE = "//*[@id=\"table-copy-from-io\"]/tbody/tr[*]/td[2]/img";
        private const string ALL_LIST = "//*[@id=\"table-copy-from-io\"]/tbody/tr[*]/td[2]";
        private const string DATEFROM = "//*[@id=\"copy-from-date-picker-start\"]";
        private const string DATETO = "//*[@id=\"copy-from-date-picker-end\"]";
        private const string LIST_DELIVARY_DATE = "//*[@id=\"table-copy-from-io\"]/tbody/tr[*]/td[4]";
        private const string PAGE_SIZE = "/html/body/div[3]/div/div/div/div[3]/div/div/div/form/div/nav/select";
        private const string FIRST_INTERIM_ORDER_TO_COPY = "//*[@id=\"table-copy-from-io\"]/tbody/tr[2]/td[2]";
        private const string SECOND_INTERIM_ORDER_TO_COPY = "//*[@id=\"table-copy-from-io\"]/tbody/tr[3]/td[2]";
        private const string COPY_FIRST_INTERIM_ORDER = "//*[@id=\"item_IsSelected\"]";
        private const string COPY_SECOND_INTERIM_ORDER = "/html/body/div[3]/div/div/div/div[3]/div/div/div/form/div/table/tbody/tr[3]/td[6]/div/input[1]";
        private const string COPIED_INTERIM_ORDER = "tb-new-purchase-order-number";
        private const string PAGINATION_NEXT_FIRST_PAGE = "//*[@id=\"div-copy-from-items\"]/nav/ul/li[4]/a";
        private const string PAGINATION_NEXT = "//*[@id=\"div-copy-from-items\"]/nav/ul/li[3]/a";
        private const string INTERIM_ORDER_SEARCH = "//*[@id=\"form-copy-from-tbSearchName\"]";

        //__________________________________ Variables _________________________________________________

        [FindsBy(How = How.Id, Using = INTERIM_NUMBER)]
        private IWebElement _interimNumber;

        [FindsBy(How = How.XPath , Using = INTERIM_ORDER_SEARCH)]
        private IWebElement _interimOrderSearch ;

        [FindsBy(How = How.Id, Using = SITE)]
        private IWebElement _site;

        [FindsBy(How = How.Id, Using = SUPPLIER)]
        private IWebElement _supplier;

        [FindsBy(How = How.Id, Using = DELIVERY_LOCATION)]
        private IWebElement _deliveryLocation;

        [FindsBy(How = How.Id, Using = DELIVERY_DATE)]
        private IWebElement _date;

        [FindsBy(How = How.Id, Using = ACTIVATED)]
        private IWebElement _activated;

        [FindsBy(How = How.Id, Using = COPY_FROM)]
        private IWebElement _createFrom;


        [FindsBy(How = How.Id, Using = SUBMIT)]
        private IWebElement _submit;

        [FindsBy(How = How.Id, Using = COMMENT)]
        private IWebElement _comment;

        [FindsBy(How = How.Id, Using = NUMBER_TO_COPY)]
        private IWebElement _numberToCopy;
        [FindsBy(How = How.Id, Using = DATEFROM)]
        private IWebElement _dateFrom;
        [FindsBy(How = How.Id, Using = DATETO)]
        private IWebElement _dateTo;
        [FindsBy(How = How.XPath, Using = PAGE_SIZE)]
        private IWebElement _pageSize;

        [FindsBy(How = How.XPath, Using = FIRST_INTERIM_ORDER_TO_COPY)]
        private IWebElement _firstInterimOrderToCopy;

        [FindsBy(How = How.XPath, Using = SECOND_INTERIM_ORDER_TO_COPY)]
        private IWebElement _secondInterimOrderToCopy;

        [FindsBy(How = How.XPath, Using = COPY_FIRST_INTERIM_ORDER)]
        private IWebElement _copyFirstInterimOrder;

        [FindsBy(How = How.XPath, Using = COPY_SECOND_INTERIM_ORDER)]
        private IWebElement _copySecondInterimOrder;

        [FindsBy(How = How.Id, Using = COPIED_INTERIM_ORDER)]
        private IWebElement _copiedInterimOrder;

        [FindsBy(How = How.XPath, Using = PAGINATION_NEXT_FIRST_PAGE)]
        private IWebElement _paginationNext;

        public InterimOrdersCreateModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }


        //___________________________________________ Méthodes ______________________________________________________

        public void FillField_CreatNewInterminOrder(DateTime date, string site, string supplier, string placeTo, bool isActivate = true)
        {
            Random rnd = new Random();

            _site = WaitForElementIsVisible(By.Id(SITE));
            _site.SetValue(ControlType.DropDownList, site);
            WaitForLoad();

            _supplier = WaitForElementIsVisible(By.Id(SUPPLIER));
            _supplier.SetValue(ControlType.DropDownList, supplier);

            _deliveryLocation = WaitForElementIsVisible(By.Id(DELIVERY_LOCATION));
            _deliveryLocation.SetValue(ControlType.DropDownList, placeTo);

            _activated = WaitForElementExists(By.Id(ACTIVATED));
            _activated.SetValue(ControlType.CheckBox, isActivate);

            _date = WaitForElementIsVisible(By.Id(DELIVERY_DATE));
            _date.SetValue(ControlType.DateTime, date);
            _date.SendKeys(Keys.Tab);

            _createFrom = WaitForElementExists(By.Id(COPY_FROM));
            _createFrom.SetValue(ControlType.CheckBox, false);

            WaitForLoad();

        }

        public void FillField_CreatNewInterminOrderWithComment(DateTime date, string site, string supplier, string placeTo, string comment, bool isActivate = true)
        {
            Random rnd = new Random();

            _site = WaitForElementIsVisible(By.Id(SITE));
            _site.SetValue(ControlType.DropDownList, site);
            WaitForLoad();

            _supplier = WaitForElementIsVisible(By.Id(SUPPLIER));
            _supplier.SetValue(ControlType.DropDownList, supplier);

            _deliveryLocation = WaitForElementIsVisible(By.Id(DELIVERY_LOCATION));
            _deliveryLocation.SetValue(ControlType.DropDownList, placeTo);

            _activated = WaitForElementExists(By.Id(ACTIVATED));
            _activated.SetValue(ControlType.CheckBox, isActivate);

            _date = WaitForElementIsVisible(By.Id(DELIVERY_DATE));
            _date.SetValue(ControlType.DateTime, date);
            _date.SendKeys(Keys.Tab);

            _comment = WaitForElementIsVisible(By.Id(COMMENT));
            _comment.SetValue(ControlType.TextBox, comment);

            _createFrom = WaitForElementExists(By.Id(COPY_FROM));
            _createFrom.SetValue(ControlType.CheckBox, false);

            WaitForLoad();

        }

        public InterimOrdersItem Submit()
        {
            _submit = WaitForElementIsVisible(By.Id(SUBMIT));
            new Actions(_webDriver).MoveToElement(_submit).Click().Perform();
            WaitPageLoading();
            WaitForLoad();
            return new InterimOrdersItem(_webDriver, _testContext);
        }

        public string GetInterimOrderId()
        {
            _interimNumber = WaitForElementIsVisible(By.Id(INTERIM_NUMBER));
            return _interimNumber.GetAttribute("value");
        }
        public bool IsModalVisible()
        {

            var modalElement = _webDriver.FindElement(By.Id("modal-1"));
            return modalElement.Displayed;
        }

        public void CopyItems()
        {
            _createFrom = WaitForElementExists(By.Id(COPY_FROM));
            _createFrom.SetValue(ControlType.CheckBox, true);

            WaitForLoad();
        }

        public void SetNumber(string number)
        {
            _numberToCopy = WaitForElementIsVisible(By.Id(NUMBER_TO_COPY));
            _numberToCopy.SetValue(ControlType.TextBox, number);
        }

        public void FillSite(string site)
        {
            _site = WaitForElementIsVisible(By.Id(SITE));
            _site.SetValue(ControlType.DropDownList, site);
            WaitForLoad();

        }
       

        public bool IsExistingFilteredNumber()
        {
            bool filteredNumberExists = isElementVisible(By.XPath(FILTERED_NUMBER));
            return filteredNumberExists;
        }

        public string GetFilteredNumber()
        {
            var number = WaitForElementIsVisible(By.XPath(FILTERED_NUMBER));
            return number.Text.Trim();
        }
        public enum FilterType
        {
            
            DateFrom,
            DateTo

        }
        public void filter(FilterType filterType, object value)
        {
            Actions action = new Actions(_webDriver);
            switch (filterType)
            {
              

                case FilterType.DateFrom:
                    _dateFrom = WaitForElementIsVisible(By.XPath(DATEFROM));
                    _dateFrom.SetValue(ControlType.DateTime, (DateTime)value);
                    _dateFrom.SendKeys(Keys.Tab);
                    break;

                case FilterType.DateTo:
                    _dateTo = WaitForElementIsVisible(By.XPath(DATETO));
                    _dateTo.SetValue(ControlType.DateTime, (DateTime)value);
                    _dateTo.SendKeys(Keys.Tab);
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);

            }
            WaitPageLoading();
            WaitForLoad();
        }
        public List<string> GetListDeliveryDate()
        {
            var ids = _webDriver.FindElements(By.XPath(LIST_DELIVARY_DATE));
            return ids.Select(e => e.Text).ToList();
        }
        public bool IsDateFromGreaterOrEqualToDateDelivary(List<string> datelist, DateTime dateFrom, DateTime dateTo)
        {
            foreach (var dateString in datelist)
            {
                DateTime date;
                if (DateTime.TryParse(dateString, out date))
                {
                    if (date < dateFrom && date > dateTo)
                    {
                        return false;
                    }
                }   
            }
            return true;
        }
        public IEnumerable<string> getListValidate()
        {
            return _webDriver.FindElements(By.XPath(LIST_VALIDATE)).Select(e => e.Text.Trim());
        }
        public IEnumerable<string> getAllList()
        {
            return _webDriver.FindElements(By.XPath(ALL_LIST)).Select(e => e.Text.Trim());
        }
        public void PageSizeCreateNewInterimOrder(string size)
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
            Actions action = new Actions(_webDriver);
            action.MoveToElement(_pageSize).Perform();
            _pageSize.SetValue(ControlType.DropDownList, size);

            WaitPageLoading();
            WaitForLoad();
        }

        public string GetFirstInterimOrderToCopy()
        {
            _firstInterimOrderToCopy = WaitForElementIsVisible(By.XPath(FIRST_INTERIM_ORDER_TO_COPY));
            return _firstInterimOrderToCopy.Text;
        }

        public string GetSecondInterimOrderToCopy()
        {
            _secondInterimOrderToCopy = WaitForElementIsVisible(By.XPath(SECOND_INTERIM_ORDER_TO_COPY));
            return _secondInterimOrderToCopy.Text;
        }

        public void CopyFirstInterimOrder()
        {

            WaitPageLoading();

            _copyFirstInterimOrder = WaitForElementExists(By.XPath(COPY_FIRST_INTERIM_ORDER));
            _copyFirstInterimOrder.Click();

            WaitPageLoading(); 
        }

        public void CopySecondInterimOrder()
        {
            _copySecondInterimOrder = WaitForElementExists(By.XPath(COPY_SECOND_INTERIM_ORDER));
            _copySecondInterimOrder.SetValue(ControlType.CheckBox, true);

            WaitForLoad();
        }

        public string GetCopiedInterimOrderNumber()
        {
            _copiedInterimOrder = WaitForElementIsVisible(By.Id(COPIED_INTERIM_ORDER));
            return _copiedInterimOrder.GetAttribute("value");
        }
        public List<string> GetNumberList()
        {
            var ListNumber = new List<string>();

            var List = _webDriver.FindElements(By.XPath(ALL_LIST));

            foreach (var row in List)
            {
                ListNumber.Add(row.Text);
            }

            return ListNumber;
        }
        public void PaginationNextFirstPage()
        {
            if (isElementVisible(By.XPath(PAGINATION_NEXT_FIRST_PAGE)))
            {
                _paginationNext = WaitForElementIsVisible(By.XPath(PAGINATION_NEXT_FIRST_PAGE));
                _paginationNext.Click();
                WaitForLoad();
            }
        }
        public void PaginationNext()
        {
            if (isElementVisible(By.XPath(PAGINATION_NEXT)))
            {
                _paginationNext = WaitForElementIsVisible(By.XPath(PAGINATION_NEXT));
                _paginationNext.Click();
                WaitForLoad();
            }
        }
        public void SearchForInterim(string interimOrderNumber )
        {
            WaitPageLoading();
            if (isElementVisible(By.XPath(INTERIM_ORDER_SEARCH)))
            {
                _interimOrderSearch.Clear(); 
                _interimOrderSearch.SendKeys(interimOrderNumber);
                WaitPageLoading(); 
            }
    
        }

    }
}