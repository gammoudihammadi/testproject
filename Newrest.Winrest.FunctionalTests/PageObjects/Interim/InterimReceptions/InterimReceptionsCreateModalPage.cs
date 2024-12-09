using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Jobs.Settings;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.Claims;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Interim.InterimReceptions
{
    public class InterimReceptionsCreateModalPage : PageBase
    {
        private const string INTERIM_NUMBER = "tb-new-interim-reception-number";
        private const string INTERIM_ORDER_NUMBER = "//*[@id=\"table-copy-from-io\"]/tbody/tr[2]/td[2]";////*[@id="table-copy-from-io"]/tbody/tr[2]/td[2]
        private const string INTERIM_RECEPTION_NUMBER = "//*[@id=\"table-copy-from-ir\"]/tbody/tr[2]/td[2]";
        private const string SITE = "SelectedSiteId";
        private const string SUPPLIER = "drop-down-suppliers";
        private const string DELIVERY_ORDER_NUMBER = "InterimReception_DeliveryOrderNumber";
        private const string DELIVERY_DATE = "datapicker-new-interim-reception-delivery";
        private const string ACTIVATED = "InterimReception_IsActive";
        private const string DELIVERY_LOCATION = "SelectedSitePlaceId";
        private const string CREATE_FROM = "checkBoxCopyFrom";
        private const string SUBMIT = "btn-submit-form-create-interim-reception";
        private const string VALIDATION_ERROR = "//*/span[contains(@class,'text-danger') and contains(@class,'field-validation-error')]";
        private const string COMMENT = "InterimReception_Comment";
        private const string FILTER_DATE_TO = "copy-from-date-picker-end";
        private const string FILTER_DATE_FROM = "copy-from-date-picker-start";
        private const string FILTER_SEARCH_NUMBER = "form-copy-from-tbSearchName";
        private const string TO_INTERIM_RECEPTION = "btn-tab-interim-reception";
        private const string TO_INTERIM_ORDER = "btn-tab-interim-order";
        private const string PAGE_SIZE = "/html/body/div[4]/div/div/div/div[3]/div/div[2]/div/form/div/div[1]/div/nav/select";
        private const string CREATE_INTERIM_RECEPTION_FROM = "checkBoxCopyFrom";
        private const string NUMBER_TO_COPY = "form-copy-from-tbSearchName";
        private const string FILTER_NUMBER = "//*[@id=\"table-copy-from-io\"]/tbody/tr[2]/td[2]";
        private const string INTERIM_ORDER_NUMBER_SELECT = "item_IsSelected";
        private const string ORDERED_VALUE = "/html/body/div[3]/div/div[3]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/form/div[1]/div[7]/span";
        private const string RECIVED_VALUE = "//*[@id=\"item_IrdRowDto_NewReceivedQuantity\"]";
        private const string PREFIL_RECIVED_QTE = "check-box-prefill-quantities";
        private const string INTERIM_ORDER_BUTTON = "item_IsSelected";
        private const string INTERIM_ORDER_NO = "form-copy-from-tbSearchName";
        private const string FIRSTNUMBERIR = "//*[@id=\"table-copy-from-ir\"]/tbody/tr[2]/td[2]";

        
        private const string INTERIM_VALID = "//*[@id=\"table-copy-from-ir\"]/tbody/tr[2]/td[2]/img";
        private const string INTERIM_RECEPTION_LIST = "//*[@id=\"table-copy-from-ir\"]/tbody/tr[*]/td[2]";
        private const string FIRSTNUMBER = "//*[@id=\"table-copy-from-io\"]/tbody/tr[3]/td[2]";
        private const string PAGINATETOSECOND = "//*[@id=\"div-create-from-interim-receptions\"]/nav/ul/li[4]";
        private const string DDL_PAGESIZE = "/html/body/div[3]/div/div/div/div[3]/div/div[2]/div/form/div/div[2]/div/nav/select";
        private const string DATES_INTERIM_RECEPTION = "//*[@id=\"table-copy-from-ir\"]/tbody/tr[*]/td[4]";
        private const string IS_SELECTED = "//*[@id=\"table-copy-from-ir\"]";
        private const string FIRST_INTERIM_RECEPTION_TO_SELECT = "/html/body/div[4]/div/div/div/div[3]/div/div[2]/div/form/div/div[2]/div/table/tbody/tr[2]/td[2]";
        private const string SECOND_INTERIM_RECEPTION_TO_SELECT = "/html/body/div[4]/div/div/div/div[3]/div/div[2]/div/form/div/div[2]/div/table/tbody/tr[3]/td[2]";
        private const string SELECT_FIRST_INTERIM_RECEPTION = "/html/body/div[4]/div/div/div/div[3]/div/div[2]/div/form/div/div[2]/div/table/tbody/tr[2]/td[6]/div/input[1]";
        private const string SELECT_SECOND_INTERIM_RECEPTION = "/html/body/div[4]/div/div/div/div[3]/div/div[2]/div/form/div/div[2]/div/table/tbody/tr[3]/td[6]/div/input[1]";
        private const string FIRST_INTERIM_ORDER_TO_SELECT = "/html/body/div[4]/div/div/div/div[3]/div/div[2]/div/form/div/div[1]/div/table/tbody/tr[2]/td[2]";
        private const string SECOND_INTERIM_ORDER_TO_SELECT = "/html/body/div[4]/div/div/div/div[3]/div/div[2]/div/form/div/div[1]/div/table/tbody/tr[3]/td[2]";
        private const string SELECT_FIRST_INTERIM_ORDER = "/html/body/div[4]/div/div/div/div[3]/div/div[2]/div/form/div/div[1]/div/table/tbody/tr[2]/td[6]/div/input[1]";
        private const string SELECT_SECOND_INTERIM_ORDER = "/html/body/div[4]/div/div/div/div[3]/div/div[2]/div/form/div/div[1]/div/table/tbody/tr[3]/td[6]/div/input[1]";
        private const string CLICK_TO_SELECT = "//*[@id=\"table-copy-from-ir\"]/tbody/tr[2]/td[6";
        private const string SELECT_FIRST_INTERIM_RECEPTION_RECEPTION = "//*[@id=\"table-copy-from-ir\"]/tbody/tr[2]/td[6]/div";
        private const string SELECT_SECOND_INTERIM_RECEPTION_RECEPTION = "/html/body/div[3]/div/div/div/div[3]/div/div[2]/div/form/div/div[2]/div/table/tbody/tr[2]/td[6]/div/input[1]";
        private const string PAGINATION_NEXT_FIRST_PAGE = "//*[@id=\"div-create-from-interim-orders\"]/nav/ul/li[4]/a";
        private const string PAGINATION_NEXT = "//*[@id=\"div-create-from-interim-orders\"]/nav/ul/li[3]/a";



        private const string PAGINATETOSECONDDEV = "//*[@id=\"div-create-from-interim-receptions\"]/nav/ul/li[4]/a";
        private const string LIST_SUPPLIER = "//*[@id=\"table-copy-from-io\"]/tbody/tr[*]";
        private const string SELECTED_SUPPLIER = "/html/body/div[3]/div/div/div/div[2]/div/form/div/div[3]/div[1]/div/div/select/option[2]";
        

        //__________________________________ Variables _________________________________________________

        [FindsBy(How = How.Id, Using = INTERIM_NUMBER)]
        private IWebElement _interimNumber;
        [FindsBy(How = How.Id, Using = PAGINATETOSECOND)]
        private IWebElement _paginatetosecond;
        [FindsBy(How = How.Id, Using = PREFIL_RECIVED_QTE)]
        private IWebElement _prefilIsSelected;

        [FindsBy(How = How.XPath, Using = ORDERED_VALUE)]
        private IWebElement _orderedValue;
        [FindsBy(How = How.XPath, Using = CLICK_TO_SELECT)]
        private IWebElement _clickToSelect;

        [FindsBy(How = How.XPath, Using = IS_SELECTED)]
        private IWebElement _isSelectedElement;

        [FindsBy(How = How.XPath, Using = RECIVED_VALUE)]
        private IWebElement _recivedValue;

        [FindsBy(How = How.XPath, Using = INTERIM_ORDER_NUMBER)]
        private IWebElement _interimOrderNumber;
        [FindsBy(How = How.Id, Using = INTERIM_ORDER_NUMBER_SELECT)]
        private IWebElement _interimOrderNumberSelect;

        [FindsBy(How = How.XPath, Using = INTERIM_RECEPTION_NUMBER)]
        private IWebElement _interimReceptionNumber;

        [FindsBy(How = How.Id, Using = SITE)]
        private IWebElement _site;

        [FindsBy(How = How.Id, Using = SUPPLIER)]
        private IWebElement _supplier;

        [FindsBy(How = How.Id, Using = DELIVERY_LOCATION)]
        private IWebElement _deliveryLocation;

        [FindsBy(How = How.Id, Using = DELIVERY_ORDER_NUMBER)]
        private IWebElement _deliveryOrderNumber;

        [FindsBy(How = How.Id, Using = DELIVERY_DATE)]
        private IWebElement _date;

        [FindsBy(How = How.Id, Using = ACTIVATED)]
        private IWebElement _activated;

        [FindsBy(How = How.Id, Using = CREATE_FROM)]
        private IWebElement _createFrom;


        [FindsBy(How = How.Id, Using = SUBMIT)]
        private IWebElement _submit;

        [FindsBy(How = How.Id, Using = COMMENT)]
        private IWebElement _comment;

        [FindsBy(How = How.Id, Using = FILTER_SEARCH_NUMBER)]
        private IWebElement _searchFilterNumber;

        [FindsBy(How = How.Id, Using = FILTER_DATE_FROM)]
        private IWebElement _dateFrom;

        [FindsBy(How = How.Id, Using = FILTER_DATE_TO)]
        private IWebElement _dateTo;

        [FindsBy(How = How.Id, Using = TO_INTERIM_RECEPTION)]
        private IWebElement _toInterimRecption;

        [FindsBy(How = How.Id, Using = TO_INTERIM_ORDER)]
        private IWebElement _toInterimOrder;

        [FindsBy(How = How.XPath, Using = PAGE_SIZE)]
        private IWebElement _pageSize;
        [FindsBy(How = How.Id, Using = CREATE_INTERIM_RECEPTION_FROM)]
        private IWebElement _createInterimFrom;

        [FindsBy(How = How.Id, Using = NUMBER_TO_COPY)]
        private IWebElement _numberToCopy;

        [FindsBy(How = How.XPath, Using = INTERIM_ORDER_BUTTON)]
        private IWebElement _interimOrderButton;

        [FindsBy(How = How.Id, Using = INTERIM_ORDER_NO)]
        private IWebElement _interimOrderNO;
        [FindsBy(How = How.Id, Using = FILTER_SEARCH_NUMBER)]
        private IWebElement _interimReceptionNO;


        [FindsBy(How = How.XPath, Using = DDL_PAGESIZE)]
        private IWebElement _ddlPageSize;

        [FindsBy(How = How.XPath, Using = FIRST_INTERIM_RECEPTION_TO_SELECT)]
        private IWebElement _firstInterimReceptionToSelect;
        [FindsBy(How = How.XPath, Using = SELECT_FIRST_INTERIM_RECEPTION)]
        private IWebElement _selectFirstInterimReception;

        [FindsBy(How = How.XPath, Using = SELECT_SECOND_INTERIM_RECEPTION)]
        private IWebElement _selectSecondInterimReception;

        [FindsBy(How = How.XPath, Using = FIRST_INTERIM_ORDER_TO_SELECT)]
        private IWebElement _firstInterimOrderToSelect;

        [FindsBy(How = How.XPath, Using = SECOND_INTERIM_ORDER_TO_SELECT)]
        private IWebElement _secondInterimOrderToSelect;

        [FindsBy(How = How.XPath, Using = SELECT_FIRST_INTERIM_ORDER)]
        private IWebElement _selectFirstInterimOrder;

        [FindsBy(How = How.XPath, Using = SELECT_SECOND_INTERIM_ORDER)]
        private IWebElement _selectSecondInterimOrder;

        [FindsBy(How = How.XPath, Using = SECOND_INTERIM_RECEPTION_TO_SELECT)]
        private IWebElement _secondInterimReceptionToSelect;

        [FindsBy(How = How.XPath, Using = SELECT_FIRST_INTERIM_RECEPTION_RECEPTION)]
        private IWebElement _selectFirstInterimReceptionReception;
        [FindsBy(How = How.XPath, Using = SELECT_SECOND_INTERIM_RECEPTION_RECEPTION)]
        private IWebElement _selectSecondInterimReceptionReception;

        [FindsBy(How = How.XPath, Using = LIST_SUPPLIER)]
        private IWebElement _listSupplier;

        [FindsBy(How = How.XPath, Using = SELECTED_SUPPLIER)]
        private IWebElement _selectedSupplier;

        
        public InterimReceptionsCreateModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //Enum
        public enum FilterType
        {
            ByNumber,
            DateFrom,
            DateTo
        }

        //___________________________________________ Méthodes ______________________________________________________
       
        public void FillField_CreatNewIntermin(DateTime date, string site, string supplier, string placeTo, bool isActivate = true, bool createFrom = false)
        {
            Random rnd = new Random();

            _site = WaitForElementIsVisible(By.Id(SITE));
            _site.SetValue(ControlType.DropDownList, site);
            WaitForLoad();

            _supplier = WaitForElementIsVisible(By.Id(SUPPLIER));
            _supplier.SetValue(ControlType.DropDownList, supplier);

            _deliveryLocation = WaitForElementIsVisible(By.Id(DELIVERY_LOCATION));
            _deliveryLocation.SetValue(ControlType.DropDownList, placeTo);

            _deliveryOrderNumber = WaitForElementIsVisible(By.Id(DELIVERY_ORDER_NUMBER));
            _deliveryOrderNumber.SetValue(ControlType.TextBox, rnd.Next().ToString());

            _activated = WaitForElementExists(By.Id(ACTIVATED));
            _activated.SetValue(ControlType.CheckBox, isActivate);

            _date = WaitForElementIsVisible(By.Id(DELIVERY_DATE));
            _date.SetValue(ControlType.DateTime, date);
            _date.SendKeys(Keys.Tab);

            _createFrom = WaitForElementExists(By.Id(CREATE_FROM));
            _createFrom.SetValue(ControlType.CheckBox, createFrom);

            WaitForLoad();
        }

        public void fillCreateInterimReceptionAndRecivedFrom()
        {
            _createFrom = WaitForElementExists(By.Id(CREATE_FROM));
            _createFrom.SetValue(ControlType.CheckBox, true);
            _interimOrderNumberSelect = WaitForElementExists(By.Id(INTERIM_ORDER_NUMBER_SELECT));
            _interimOrderNumberSelect.SetValue(ControlType.CheckBox, true);
            _prefilIsSelected = WaitForElementExists(By.Id(PREFIL_RECIVED_QTE));
            _prefilIsSelected.SetValue(ControlType.CheckBox, true);
        }
        public void checkUnResultats(string number)
        {
            _createFrom = WaitForElementExists(By.Id(CREATE_FROM));
            _createFrom.SetValue(ControlType.CheckBox, true);
            _toInterimRecption = WaitForElementExists(By.Id(TO_INTERIM_RECEPTION));
            _toInterimRecption.SetValue(ControlType.CheckBox, true);
            var isSelectedElement = WaitForElementExists(By.XPath("//*[@id='item_IsSelected']"));
            isSelectedElement.SetValue(ControlType.CheckBox, true);
            _interimReceptionNO = WaitForElementIsVisible(By.Id(FILTER_SEARCH_NUMBER));
            _interimReceptionNO.Clear();
            _interimReceptionNO.SendKeys(number);
            _prefilIsSelected = WaitForElementExists(By.Id(PREFIL_RECIVED_QTE));
            _prefilIsSelected.SetValue(ControlType.CheckBox, true);
           
        }
        public void SelectFirstInterimReception(string number)
        {
            WaitForLoad();
            _createFrom = WaitForElementExists(By.Id(CREATE_FROM));
            _createFrom.SetValue(ControlType.CheckBox, true);
            _toInterimRecption = WaitForElementExists(By.Id(TO_INTERIM_RECEPTION));
            _toInterimRecption.SetValue(ControlType.CheckBox, true);
            _numberToCopy = WaitForElementIsVisible(By.Id(NUMBER_TO_COPY));
            _numberToCopy.SetValue(ControlType.TextBox, number);
            WaitForLoad();
            _isSelectedElement = WaitForElementIsVisible(By.XPath(IS_SELECTED));
            _isSelectedElement.SetValue(ControlType.CheckBox, true);
            WaitForLoad();
        }


        public void select(string number)
        {
            WaitForLoad();
            _createFrom = WaitForElementExists(By.Id(CREATE_FROM));
            _createFrom.SetValue(ControlType.CheckBox, true);
            _numberToCopy = WaitForElementIsVisible(By.Id(NUMBER_TO_COPY));
            _numberToCopy.SetValue(ControlType.TextBox, number);  
            _interimOrderNumberSelect = WaitForElementExists(By.Id(INTERIM_ORDER_NUMBER_SELECT));
            _interimOrderNumberSelect.SetValue(ControlType.CheckBox, true);
            
        }

        public string GetNumberToCopy()
        {
            var isSelectedElement = WaitForElementExists(By.XPath("//*[@id='item_IsSelected']"));

            if (isSelectedElement != null && isSelectedElement.Selected)
            {
                var interimOrderNumberElement = WaitForElementExists(By.XPath("//*[@id='table-copy-from-io']/tbody/tr[2]/td[2]"));
                return interimOrderNumberElement.Text;
            }
            return string.Empty;
        }

        public string GetFormattedPriceToCopy()
        {
            var isSelectedElement = WaitForElementExists(By.XPath("//*[@id='item_IsSelected']"));

            if (isSelectedElement != null && isSelectedElement.Selected)
            {
                var interimOrderNumberElement = WaitForElementExists(By.XPath("//*[@id='table-copy-from-io']/tbody/tr[2]/td[5]"));
                string numericPart = interimOrderNumberElement.Text.Replace("€", "").Trim();

                if (decimal.TryParse(numericPart, out decimal parsedPrice))
                {
                    return $"€ {parsedPrice:0.0000}";
                }
            }

            return string.Empty;
        }


        public string getDateToCopy()
        {
            var isSelectedElement = WaitForElementExists(By.XPath("//*[@id='item_IsSelected']"));

            if (isSelectedElement != null && isSelectedElement.Selected)
            {
                var interimOrderDateElement = WaitForElementExists(By.XPath("//*[@id='table-copy-from-io']/tbody/tr[2]/td[4]"));
                return interimOrderDateElement.Text;
            }
            return string.Empty;
        }


        public void clickNumberToCopy()
        {
            var interimOrderNumberElement = WaitForElementExists(By.XPath("//*[@id=\"form-create-interim-reception\"]/div/div[3]/div/div/div/span/a"));
            interimOrderNumberElement.Click();
            //return interimOrderNumberElement.Text;
        }
        public string getNumber()
        {
            var interimOrderNumberElement = WaitForElementExists(By.XPath("//*[@id=\"form-create-interim-reception\"]/div/div[3]/div/div/div/span/a"));    
            return interimOrderNumberElement.Text;
        }


        public InterimReceptionsItem Submit()
        {
            _submit = WaitForElementExists(By.Id(SUBMIT));
            WaitPageLoading();
            new Actions(_webDriver).MoveToElement(_submit).Click().Perform();
            return new InterimReceptionsItem(_webDriver, _testContext);
        }

        public string GetInterimId()
        {
            _interimNumber = WaitForElementIsVisible(By.Id(INTERIM_NUMBER));
            return _interimNumber.GetAttribute("value");
        }

        public string ValidationError()
        {
            var _validationError = WaitForElementIsVisible(By.XPath(VALIDATION_ERROR));
            return _validationError.Text;
        }

        public void FillField_CreatNewInterminAddComment(DateTime date, string site, string supplier, string placeTo, string comment, bool isActivate = true)
        {
            Random rnd = new Random();


            _site = WaitForElementIsVisible(By.Id(SITE));
            _site.SetValue(ControlType.DropDownList, site);
            WaitForLoad();

            _supplier = WaitForElementIsVisible(By.Id(SUPPLIER));
            _supplier.SetValue(ControlType.DropDownList, supplier);

            _deliveryLocation = WaitForElementIsVisible(By.Id(DELIVERY_LOCATION));
            _deliveryLocation.SetValue(ControlType.DropDownList, placeTo);

            _deliveryOrderNumber = WaitForElementIsVisible(By.Id(DELIVERY_ORDER_NUMBER));
            _deliveryOrderNumber.SetValue(ControlType.TextBox, rnd.Next().ToString());

            _activated = WaitForElementExists(By.Id(ACTIVATED));
            _activated.SetValue(ControlType.CheckBox, isActivate);

            _date = WaitForElementIsVisible(By.Id(DELIVERY_DATE));
            _date.SetValue(ControlType.DateTime, date);
            _date.SendKeys(Keys.Tab);

            _comment = WaitForElementIsVisible(By.Id(COMMENT));
            _comment.SetValue(ControlType.TextBox, comment);

            _createFrom = WaitForElementExists(By.Id(CREATE_FROM));
            _createFrom.SetValue(ControlType.CheckBox, false);

            WaitForLoad();

        }
        
        public void Filter(FilterType filterType, object value)
        {
            Actions action = new Actions(_webDriver);

            switch (filterType)
            {

                case FilterType.ByNumber:
                    _searchFilterNumber = WaitForElementIsVisible(By.Id(FILTER_SEARCH_NUMBER));
                    _searchFilterNumber.SetValue(ControlType.TextBox, value);
                    break;

                case FilterType.DateFrom:
                    _dateFrom = WaitForElementIsVisible(By.Id(FILTER_DATE_FROM));
                    _dateFrom.SetValue(ControlType.DateTime, value);
                    _dateFrom.SendKeys(Keys.Tab);
                    break;

                case FilterType.DateTo:
                    _dateTo = WaitForElementIsVisible(By.Id(FILTER_DATE_TO));
                    _dateTo.SetValue(ControlType.DateTime, value);
                    _dateTo.SendKeys(Keys.Tab);
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);
            }

            WaitPageLoading();
            WaitForLoad();
        }

        public string GetOrderNumber()
        {
            _interimOrderNumber = WaitForElementIsVisible(By.XPath(INTERIM_ORDER_NUMBER));
            return _interimOrderNumber.Text;
        }

        public string GetReceptionNumber()
        {
            _interimReceptionNumber = WaitForElementIsVisible(By.XPath(INTERIM_RECEPTION_NUMBER));
            return _interimReceptionNumber.Text;
        }
        public void ClearField()
        {
            _searchFilterNumber = WaitForElementExists(By.Id(FILTER_SEARCH_NUMBER));
            _searchFilterNumber.Clear();
            WaitForLoad();
            _dateFrom = WaitForElementExists(By.Id(FILTER_DATE_FROM));
            _dateFrom.Clear();
            WaitForLoad();
            _dateTo = WaitForElementExists(By.Id(FILTER_DATE_TO));
            _dateTo.Clear();
            WaitForLoad();
            WaitPageLoading();
            Thread.Sleep(1000);
        }
        public bool ContainsData()
        {
            return isElementExists(By.XPath("//*[@id=\"table-copy-from-io\"]/tbody/tr[2]"));
        }
        public void ToInterimReception()
        {
            _toInterimRecption = WaitForElementExists(By.Id(TO_INTERIM_RECEPTION));
            _toInterimRecption.Click();
        }
        public void ToInterimOrder()
        {
            _toInterimOrder = WaitForElementExists(By.Id(TO_INTERIM_ORDER));
            _toInterimOrder.Click();
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
            Actions action = new Actions(_webDriver);
            action.MoveToElement(_pageSize).Perform();
            _pageSize.SetValue(ControlType.DropDownList, size);

            WaitPageLoading();
            WaitForLoad();
        }

        public List<string> GetNumberList()
        {
            var ListNumber = new List<string>();

            var List = _webDriver.FindElements(By.XPath("//*[@id=\"table-copy-from-io\"]/tbody/tr[*]/td[2]"));

            foreach (var row in List)
            {
                ListNumber.Add(row.Text);
            }

            return ListNumber;
        }
        public void FillDeliveryOrderNumber()
        {
            _deliveryOrderNumber = WaitForElementExists(By.Id(DELIVERY_ORDER_NUMBER));
            _deliveryOrderNumber.SetValue(ControlType.TextBox, new Random().Next().ToString());
            WaitForLoad();
        }


        public void ClickOnCreateInterimReceptionFrom()
        {
            _createInterimFrom = WaitForElementExists(By.Id(CREATE_INTERIM_RECEPTION_FROM));
            _createInterimFrom.SetValue(ControlType.CheckBox, true);

        }
        public void SetNumber(string number)
        {
            _numberToCopy = WaitForElementIsVisible(By.Id(NUMBER_TO_COPY));
            _numberToCopy.SetValue(ControlType.TextBox, number);
        }
        public bool IsExistingFilteredNumber()
        {
            bool filteredNumberExists = isElementVisible(By.XPath(FILTER_NUMBER));
            return filteredNumberExists;
        }

        public void FillField_CreatNewInterminAddCommentReceptionFromOn(DateTime date, string site, string supplier, string placeTo, string comment, bool isActivate = true)
        {
            Random rnd = new Random();


            _site = WaitForElementIsVisible(By.Id(SITE));
            _site.SetValue(ControlType.DropDownList, site);
            WaitForLoad();

            _supplier = WaitForElementIsVisible(By.Id(SUPPLIER));
            _supplier.SetValue(ControlType.DropDownList, supplier);

            _deliveryLocation = WaitForElementIsVisible(By.Id(DELIVERY_LOCATION));
            _deliveryLocation.SetValue(ControlType.DropDownList, placeTo);

            _deliveryOrderNumber = WaitForElementIsVisible(By.Id(DELIVERY_ORDER_NUMBER));
            _deliveryOrderNumber.SetValue(ControlType.TextBox, rnd.Next().ToString());

            _activated = WaitForElementExists(By.Id(ACTIVATED));
            _activated.SetValue(ControlType.CheckBox, isActivate);

            _date = WaitForElementIsVisible(By.Id(DELIVERY_DATE));
            _date.SetValue(ControlType.DateTime, date);
            _date.SendKeys(Keys.Tab);

            _date = WaitForElementIsVisible(By.Id(DELIVERY_DATE));
            _date.SetValue(ControlType.DateTime, date);
            _date.SendKeys(Keys.Tab);

            _comment = WaitForElementIsVisible(By.Id(COMMENT));
            _comment.SetValue(ControlType.TextBox, comment);

            WaitForLoad();

        }

        public bool CheckDateFiltringWorking(DateTime startDate, DateTime endDate)
        {
            var listOfDatesElements = _webDriver.FindElements(By.XPath("//*[@id=\"table-copy-from-io\"]/tbody/tr[*]/td[4]"));
            List<string> dateList = new List<string>();
            foreach (var element in listOfDatesElements)
            {
                dateList.Add(element.Text);
            }
            foreach (var dateStr in dateList)
            {
                DateTime dateVar;

                if (!DateTime.TryParseExact(dateStr, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateVar))
                {
                    throw new ArgumentException($"Invalid date format for date: {dateStr}");
                }

                // Check if the date is within the range
                if (dateVar < startDate || dateVar > endDate)
                {
                    return false;
                }
            }

            return true;
        }

        public void SelectFirstInterimOrder()
        {
            _interimOrderButton = WaitForElementToBeClickable(By.XPath("//*[@id=\"table-copy-from-io\"]/tbody/tr[3]/td[6]/div"));

            ((IJavaScriptExecutor)_webDriver).ExecuteScript("arguments[0].scrollIntoView(true);", _interimOrderButton);
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(10));
            wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(_interimOrderButton));

            _interimOrderButton.Click();
        }


        public void SelectFirstInterimRecep()
        {
            _selectFirstInterimReception = WaitForElementIsVisible(By.XPath("//*[@id=\"table-copy-from-ir\"]/tbody/tr[2]/td[6]/div/input[2]"));
            _selectFirstInterimReception.Click();

        }
        public void SearchForInterimOrderNumber(string orderNumber)
        {
            WaitForLoad();
            _interimOrderNO = WaitForElementIsVisible(By.Id(INTERIM_ORDER_NO));
            _interimOrderNO.Clear(); 
            _interimOrderNO.SendKeys(orderNumber);
            WaitPageLoading();
            WaitPageLoading();
            _selectFirstInterimReceptionReception = WaitForElementIsVisible(By.XPath(SELECT_FIRST_INTERIM_RECEPTION_RECEPTION));
            WaitForLoad();
            _selectFirstInterimReceptionReception.Click();
            WaitPageLoading();
            WaitForLoad();
        }

        public void SearchForInterimOrderNumberAndChoseFirst(string orderNumber)
        {
            WaitForLoad();
            _interimOrderNO = WaitForElementIsVisible(By.Id(INTERIM_ORDER_NO));
            _interimOrderNO.Clear();
            _interimOrderNO.SendKeys(orderNumber);
            WaitPageLoading();
            WaitPageLoading();
            _interimOrderNumberSelect = WaitForElementExists(By.Id(INTERIM_ORDER_NUMBER_SELECT));
            _interimOrderNumberSelect.SetValue(ControlType.CheckBox, true);
            WaitForLoad();
      
        }
        public void SearchForInterimSecondNumber(string orderNumber2)
        {
            WaitForLoad();
            _interimOrderNO = WaitForElementIsVisible(By.Id(INTERIM_ORDER_NO));
            _interimOrderNO.Clear();
            WaitForLoad();
            _interimOrderNO.SendKeys(orderNumber2);
            WaitPageLoading();
            WaitPageLoading();
        }
        public void selectSecond()
        {
           
            _selectSecondInterimReceptionReception = WaitForElementExists(By.XPath(SELECT_SECOND_INTERIM_RECEPTION_RECEPTION));
            WaitForLoad();
            _selectSecondInterimReceptionReception.SetValue(ControlType.CheckBox, true);
        }
        public void CreateInterimReceptionFrom()
        {
            _createFrom = WaitForElementExists(By.Id(CREATE_FROM));
            _createFrom.SetValue(ControlType.CheckBox, true);
        }

        public bool IsInterimReceptionList()
        {
            try
            {
                return isElementVisible(By.XPath(INTERIM_RECEPTION_LIST));
            }
            catch
            {
                return false;
            }
        }

        public bool IsInterimReceptionValid()
        {
            var interimList = _webDriver.FindElements(By.XPath(INTERIM_RECEPTION_LIST));

            if (interimList.Count == 0)
            {
                return false;
            }

            foreach (var row in interimList)
            {
                if (row.FindElements(By.XPath(INTERIM_VALID)).Count == 0)
                {
                    return false;
                }
            }

            return true;
        }
        public void SetPageSize(string size)
        {
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
            action.MoveToElement(_ddlPageSize).Perform();
            _pageSize.SetValue(ControlType.DropDownList, size);

            WaitPageLoading();
        }
        public bool IsFromToDateRespected(DateTime fromDate, DateTime toDate, string dateFormat)
        {
            var cultureInfo = dateFormat.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            var dates = _webDriver.FindElements(By.XPath(DATES_INTERIM_RECEPTION));

            if (dates.Count == 0)
                return false;

            return dates
                .Select(elm => DateTime.Parse(elm.Text, cultureInfo).Date)
                .Any(date => date >= fromDate && date <= toDate);
        }

        public string GetFirstNumberOfFirstLine()
        {
            var element = WaitForElementIsVisible(By.XPath(FIRSTNUMBER));
            string text = element.Text;
            return text;
        }
        public void PaginateToSecond()
        {
            if (IsDev())
            {
                _paginatetosecond = WaitForElementIsVisible(By.XPath(PAGINATETOSECONDDEV));
            }
            else
            {
               _paginatetosecond = WaitForElementIsVisible(By.XPath(PAGINATETOSECOND));

            }
            _paginatetosecond.Click();
            WaitPageLoading();
            WaitForLoad();

        }
        public void UnchekPrefillReceivedQuantities()
        {
            _prefilIsSelected = WaitForElementExists(By.Id(PREFIL_RECIVED_QTE));
            _prefilIsSelected.SetValue(ControlType.CheckBox, false);
        }

        public string GetFirstInterimReceptionToSelect()
        {
            _firstInterimReceptionToSelect = WaitForElementIsVisible(By.XPath(FIRST_INTERIM_RECEPTION_TO_SELECT));
            return _firstInterimReceptionToSelect.Text;
        }

        public string GetSecondInterimReceptionToSelect()
        {
            _secondInterimReceptionToSelect = WaitForElementIsVisible(By.XPath(SECOND_INTERIM_RECEPTION_TO_SELECT));
            return _secondInterimReceptionToSelect.Text;
        }

        public void SelectFirstAndSecondInterimReception()
        {
            _selectFirstInterimReception = WaitForElementExists(By.XPath(SELECT_FIRST_INTERIM_RECEPTION));
            _selectFirstInterimReception.SetValue(ControlType.CheckBox, true);
            WaitForLoad();

            _selectSecondInterimReception = WaitForElementExists(By.XPath(SELECT_SECOND_INTERIM_RECEPTION));
            _selectSecondInterimReception.SetValue(ControlType.CheckBox, true);
            WaitForLoad();
        }
        public string GetFirstInterimOrderToSelect()
        {
            WaitPageLoading();
            _firstInterimOrderToSelect = WaitForElementExists(By.XPath(FIRST_INTERIM_ORDER_TO_SELECT));
            return _firstInterimOrderToSelect.Text;
        }

        public string GetSecondInterimOrderToSelect()
        {
            WaitPageLoading();
            _secondInterimOrderToSelect = WaitForElementExists(By.XPath(SECOND_INTERIM_ORDER_TO_SELECT));
            return _secondInterimOrderToSelect.Text;
        }
        public void SelectFirstAndSecondInterimOrder()
        {
            WaitPageLoading();
            _selectFirstInterimOrder = WaitForElementExists(By.XPath(SELECT_FIRST_INTERIM_ORDER));
            _selectFirstInterimOrder.SetValue(ControlType.CheckBox, true);
            WaitForLoad();

            _selectSecondInterimOrder = WaitForElementExists(By.XPath(SELECT_SECOND_INTERIM_ORDER));
            _selectSecondInterimOrder.SetValue(ControlType.CheckBox, true);
            WaitForLoad();
        }
        public void SearchForInterimReceptionNumber(string orderNumber)
        {
            WaitForLoad();
            _interimReceptionNO = WaitForElementIsVisible(By.Id(FILTER_SEARCH_NUMBER));
            _interimReceptionNO.Clear();
            _interimReceptionNO.SendKeys(orderNumber);

        }
        public void toAninterimReception()
        {
            WaitForLoad();
            _toInterimRecption = WaitForElementExists(By.Id(TO_INTERIM_RECEPTION));
            _toInterimRecption.SetValue(ControlType.CheckBox, true);
        }
        public string GetFirstNumberOfFirstLineIR()
        {
            var element = WaitForElementIsVisible(By.XPath(FIRSTNUMBERIR));
            string text = element.Text;
            return text;
        }

        public void FillField_CreatNewInterminWithoutOrderNumber(DateTime date, string site, string supplier, string placeTo, bool isActivate = true, bool createFrom = false)
        {
            Random rnd = new Random();

            _site = WaitForElementIsVisible(By.Id(SITE));
            _site.SetValue(ControlType.DropDownList, site);
            WaitForLoad();

            _supplier = WaitForElementIsVisible(By.Id(SUPPLIER));
            _supplier.SetValue(ControlType.DropDownList, supplier);

            _deliveryLocation = WaitForElementIsVisible(By.Id(DELIVERY_LOCATION));
            _deliveryLocation.SetValue(ControlType.DropDownList, placeTo);

            _prefilIsSelected = WaitForElementExists(By.Id(PREFIL_RECIVED_QTE));
            _prefilIsSelected.SetValue(ControlType.CheckBox, true);

            _activated = WaitForElementExists(By.Id(ACTIVATED));
            _activated.SetValue(ControlType.CheckBox, isActivate);

            _date = WaitForElementIsVisible(By.Id(DELIVERY_DATE));
            _date.SetValue(ControlType.DateTime, date);
            _date.SendKeys(Keys.Tab);

            _createFrom = WaitForElementExists(By.Id(CREATE_FROM));
            _createFrom.SetValue(ControlType.CheckBox, true);

            WaitForLoad();
        }

        public bool VerifiedSupplierAndList(string supplier)
        {
            WaitForLoad();
            _selectedSupplier = WaitForElementIsVisible(By.XPath(SELECTED_SUPPLIER));
            var _listSupplier = _webDriver.FindElements(By.XPath(LIST_SUPPLIER));

            WaitForLoad();

            if ((_selectedSupplier.Text == supplier) &&(_listSupplier.Count != 0))
            {
                return true;
            }
            else return false;
         

        }
        public void PageSizeCreateNewInterimReception(string pageSize)
        {
            var pageSizeDropdown = _webDriver.FindElement(By.XPath("/html/body/div[3]/div/div/div/div[3]/div/div[2]/div/form/div/div[1]/div/nav/select"));
            var selectElement = new SelectElement(pageSizeDropdown);

            selectElement.SelectByValue(pageSize);

            IJavaScriptExecutor js = (IJavaScriptExecutor)_webDriver;
            js.ExecuteScript("arguments[0].dispatchEvent(new Event('change'))", pageSizeDropdown);

            WaitPageLoading(); 
        }


        public void PaginationNextFirstPage()
        {
            if (isElementVisible(By.XPath(PAGINATION_NEXT_FIRST_PAGE)))
            {
                var paginationNext = WaitForElementIsVisible(By.XPath(PAGINATION_NEXT_FIRST_PAGE));
                paginationNext.Click();
                WaitForLoad();
            }
        }
        public void PaginationNext()
        {
            if (isElementVisible(By.XPath(PAGINATION_NEXT)))
            {
                var paginationNext = WaitForElementIsVisible(By.XPath(PAGINATION_NEXT));
                paginationNext.Click();
                WaitForLoad();
            }
        }
        public void CreateInterimFrom()
        {
            _createFrom = WaitForElementExists(By.Id(CREATE_FROM));
            _createFrom.SetValue(ControlType.CheckBox, true);

            WaitForLoad();
        }
        public void ValidatePagination(string pageSize, int expectedMaxSize, bool checkMultiplePages = false)
        {
            PageSizeCreateNewInterimReception(pageSize);
            var resultPageSize = GetNumberList().Count;
            Assert.IsTrue(resultPageSize <= expectedMaxSize, $"La pagination du {expectedMaxSize} ne fonctionne pas. Actuellement {resultPageSize} items.");

            if (checkMultiplePages && CheckTotalNumber() > expectedMaxSize)
            {
                PaginationNext();
                resultPageSize = GetNumberList().Count;
                Assert.IsTrue(resultPageSize <= expectedMaxSize, $"La pagination du {expectedMaxSize} ne fonctionne pas pour plusieurs pages. Actuellement {resultPageSize} items.");
            }
            else
            {
                PaginationNextFirstPage();
                resultPageSize = GetNumberList().Count;
                Assert.IsTrue(resultPageSize <= expectedMaxSize, $"La pagination du {expectedMaxSize} ne fonctionne pas pour la première page. Actuellement {resultPageSize} items.");
            }
        }
     }

}