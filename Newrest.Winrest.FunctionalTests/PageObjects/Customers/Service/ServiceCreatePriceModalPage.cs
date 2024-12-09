using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Keys = OpenQA.Selenium.Keys;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service
{
    public class ServiceCreatePriceModalPage : PageBase
    {

        public ServiceCreatePriceModalPage(IWebDriver _webDriver, TestContext _testContext) : base(_webDriver, _testContext)
        {
        }

        // ____________________________________________ Constantes ________________________________________________

        // General infos
        private const string SITE = "dropdown-site";

        private const string CUSTOMER_NAME = "dropdown-customer-selectized";
        private const string CUSTOMER = "dropdown-customer-selectized";
        private const string CUSTOMER_DIV = "//*/div[contains(@class,'selectize-input ')]";
        private const string CUSTOMER_SELECTED = "//*/div[contains(@class,'selectize-dropdown ')]";

        private const string CUSTOMER_SELECT = "//*[@id=\"dropdown-customer\"]/../div/div[1]";
        private const string CUSTOMER_SELECT_INPUT = "//select[@id='dropdown-customer']/../div/div/input";
        private const string CUSTOMER_SELECT_FIRST_SELECTION = "//*[@id=\"dropdown-customer\"]/../div/div[2]/div";
        // Period
        private const string BEGIN_DATE = "begin-date-picker";
        private const string END_DATE = "end-date-picker";
        private const string OTHER_BEGIN_DATE = "date-picker-dest-start";
        private const string OTHER_END_DATE = "date-picker-dest-end";
        private const string DATASHEET_MENU = "//*[@id=\"modal-1\"]/div/div/form/div[2]/div[3]/div/div/div[3]/div/span[1]/div";

        // Price and planification
        private const string DATASHEET_NAME = "//*[@id=\"modal-1\"]/div/div/div/div/form/div[2]/div[3]/div[3]/div/span[1]/input[2]";
        private const string DATASHEET_SELECTED = "//*[@id=\"modal-1\"]/div/div/div/div/form/div[2]/div[3]/div[3]/div/span[1]/div/div/div[text()=\"{0}\"]";
        private const string INVOICE_PRICE = "item_InvoicePrice";

        private const string SAVEBTN_CUST_PRICE = "last";
        private const string MODAL_CREATE_NEW_CUSTOMER_PRICE = "/html/body/div[4]/div/div/div/div/div/div/form";
        private const string ERROR_PRICE_REQUIRED = "//span[text()='A price already exists for this period']";

        private const string SCALEMODE = "//*[@id=\"item_DefaultMode\"]/option[4]";
        private const string POURCENTAGEMODE = "//*[@id=\"item_DefaultMode\"]/option[1]";

        private const string EDITSCALEMODE = "//*[@id=\"modal-1\"]/div/div/form/div[2]/div[1]/div[2]/div[1]/a";
        private const string LABELROUND = "//*[@id=\"modal-1\"]/div/div/form/div[2]/div[1]/div[2]/div[3]/label";
        private const string METHOD = "item_Method";
        private const string SERVICE_DELETE_SCALE = "//*[@id=\"scales-table\"]/tbody/tr[{0}]/td[4]/a";
        private const string CONFIRM_OK = "dataConfirmOK";
        private const string ALL_SCALE_PRICE = "//*[@id=\"scales-table\"]/tbody/tr[*]/td[3]/input";
        private const string LISTEOPTIONROUND = "//*[@id=\"item_RoundMode\"]/option";
        private const string NEW_SCALE = "btn-add-scale";

        private const string METHOD_CYCLE = "//*[@id=\"item_Method\"]/option[3]";
        private const string CYCLE_MODE_CONTINUOUS = "//*[@id=\"item_CycleMode\"]/option[1]";
        private const string NB_CYCLE = "item_NbCycle";
        private const string METHODE_SCALE = "//*[@id=\"item_Cycles_0__Method\"]/option[2]";
        private const string NB_CYCLE_TABLE = "/html/body/div[4]/div/div/div/div/form/div[2]/div[7]/div/table/tbody/tr[*]";

        private const string BTN_SAVE_EDIT_PRICE = "//html/body/div[4]/div/div/div/div/form/div[3]/button[2]";
        private const string CYCLE_END_Of_MONTH = "//*[@id=\"item_CycleMode\"]/option[2]";

        private const string FIRST_DURATION_DAYS = "/html/body/div[4]/div/div/div/div/form/div[2]/div[7]/div/table/tbody/tr[1]/td[2]/input";
        private const string CYCLE_EFULL_MONTH = "//*[@id=\"item_CycleMode\"]/option[3]";

        private const string METHOD_FIXED= "//*[@id=\"item_Method\"]/option[1]";
        private const string METHOD_SCALE = "//*[@id=\"item_Method\"]/option[2]";

        private const string LABEL_DATASHEET = "/html/body/div[4]/div//div/div/form/div[2]/div[3]/div/div/div[3]/label";
        private const string LABEL_INVOISE_PRICE = "/html/body/div[4]/div//div/div/form/div[2]/div[3]/div/div/div[7]/label";

        private const string LABEL_DATASHEET_SCALE = "/html/body/div[4]/div//div/div/form/div[2]/div[3]/div/div/div[3]/label";
        private const string LABEL_NB_OF_PERSON = "//*[@id=\"scales-table\"]/tbody/tr/th[1]";
        private const string LABEL_TOTAL_PRICE = "//*[@id=\"scales-table\"]/tbody/tr/th[2]";
        private const string LABEL_UNIT_PRICE = "//*[@id=\"scales-table\"]/tbody/tr/th[3]";
        private const string LABEL_NEW_SCALE = "//*[@id=\"btn-add-scale\"]";

        private const string LABEL_CYCLE_MODE = "/html/body/div[4]/div//div/div/form/div[2]/div[3]/div/div/div[2]/label";
        private const string LABELNB_CYCLE = "/html/body/div[4]/div//div/div/form/div[2]/div[3]/div/div/div[4]/label";
        private const string LABEL_START_AT = "/html/body/div[4]/div//div/div/form/div[2]/div[3]/div/div/div[5]/label";
        private const string LABEL_WITH_DURATION = "/html/body/div[4]/div//div/div/form/div[2]/div[3]/div/div/div[6]/label";
        private const string LABEL_NAME = "/html/body/div[4]/div//div/div/form/div[2]/div[7]/div/table/thead/tr/th[1]";
        private const string LABEL_DURATION = "/html/body/div[4]/div//div/div/form/div[2]/div[7]/div/table/thead/tr/th[2]";
        private const string LABEL_DATASHEET_CYCLE = "/html/body/div[4]/div//div/div/form/div[2]/div[7]/div/table/thead/tr/th[3]";
        private const string LABEL_METHOD_CYCLE = "/html/body/div[4]/div//div/div/form/div[2]/div[7]/div/table/thead/tr/th[4]";
        private const string LABEL_PRICE = "/html/body/div[4]/div//div/div/form/div[2]/div[7]/div/table/thead/tr/th[5]";


        // ____________________________________________ Variables _________________________________________________

        // General infos

        [FindsBy(How = How.XPath, Using = LABEL_CYCLE_MODE)]
        private IWebElement _label_cycle_mode;
        [FindsBy(How = How.XPath, Using = LABELNB_CYCLE)]
        private IWebElement _label_nb_cycle;
        [FindsBy(How = How.XPath, Using = LABEL_START_AT)]
        private IWebElement _label_start_at;
        [FindsBy(How = How.XPath, Using = LABEL_WITH_DURATION)]
        private IWebElement _label_with_duration;

        [FindsBy(How = How.XPath, Using = LABEL_NAME)]
        private IWebElement _label_name;
        [FindsBy(How = How.XPath, Using = LABEL_DURATION)]
        private IWebElement _label_duration;
        [FindsBy(How = How.XPath, Using = LABEL_DATASHEET_CYCLE)]
        private IWebElement _label_datasheet_cycle;
        [FindsBy(How = How.XPath, Using = LABEL_METHOD_CYCLE)]
        private IWebElement _label_method_scale;
        [FindsBy(How = How.XPath, Using = LABEL_PRICE)]
        private IWebElement _label_price;










        [FindsBy(How = How.XPath, Using = LABEL_NEW_SCALE)]
        private IWebElement _label_new_scale;

        [FindsBy(How = How.XPath, Using = LABEL_DATASHEET_SCALE)]
        private IWebElement _label_datasheet_scale;
        [FindsBy(How = How.XPath, Using = LABEL_NB_OF_PERSON)]
        private IWebElement _label_nb_of_person;
        [FindsBy(How = How.XPath, Using = LABEL_TOTAL_PRICE)]
        private IWebElement _label_total_price;
        [FindsBy(How = How.XPath, Using = LABEL_UNIT_PRICE)]
        private IWebElement _label_unit_price;

        [FindsBy(How = How.XPath, Using = LABEL_INVOISE_PRICE)]
        private IWebElement _label_invoise_price;

        [FindsBy(How = How.XPath, Using = LABEL_DATASHEET)]
        private IWebElement _label_datasheet;

        [FindsBy(How = How.XPath, Using = METHOD_SCALE)]
        private IWebElement _method_scale;


        [FindsBy(How = How.XPath, Using = METHOD_FIXED)]
        private IWebElement _methode_fixed;

        [FindsBy(How = How.XPath, Using = CYCLE_EFULL_MONTH)]
        private IWebElement _full_month;

        [FindsBy(How = How.XPath, Using = FIRST_DURATION_DAYS)]
        private IWebElement _first_duration_days;

        [FindsBy(How = How.XPath, Using = CYCLE_END_Of_MONTH)]
        private IWebElement _end_of_month;

        [FindsBy(How = How.XPath, Using = BTN_SAVE_EDIT_PRICE)]
        private IWebElement _btn_save_edit_price;

        [FindsBy(How = How.XPath, Using = NB_CYCLE_TABLE)]
        private IWebElement _nbcycletable;

        [FindsBy(How = How.XPath, Using = METHODE_SCALE)]
        private IWebElement _methodscale;

        [FindsBy(How = How.Id, Using = NB_CYCLE)]
        private IWebElement _nb_cycle;

        [FindsBy(How = How.XPath, Using = CYCLE_MODE_CONTINUOUS)]
        private IWebElement _cyclemodecontinuous;

        [FindsBy(How = How.XPath, Using = METHOD_CYCLE)]
        private IWebElement _methodcycle;

        [FindsBy(How = How.XPath, Using = LABELROUND)]
        private IWebElement _labelround;

        [FindsBy(How = How.XPath, Using = POURCENTAGEMODE)]
        private IWebElement _pourcentagemode;

        [FindsBy(How = How.XPath, Using = EDITSCALEMODE)]
        private IWebElement _editscalemode;

        [FindsBy(How = How.XPath, Using = SCALEMODE)]
        private IWebElement _sclaemode;

        [FindsBy(How = How.Id, Using = SITE)]
        private IWebElement _priceSite;

        [FindsBy(How = How.Id, Using = CUSTOMER_NAME)]
        private IWebElement _customerName;

        [FindsBy(How = How.Id, Using = CUSTOMER)]
        private IWebElement _customer;

        // Period
        [FindsBy(How = How.Id, Using = BEGIN_DATE)]
        private IWebElement _fromDate;

        [FindsBy(How = How.Id, Using = END_DATE)]
        private IWebElement _toDate;

        // Price and planification
        [FindsBy(How = How.XPath, Using = DATASHEET_NAME)]
        private IWebElement _datasheetName;

        [FindsBy(How = How.XPath, Using = DATASHEET_SELECTED)]
        private IWebElement _selectedDatasheet;

        [FindsBy(How = How.Id, Using = INVOICE_PRICE)]
        private IWebElement _invoicePrice;

        [FindsBy(How = How.Id, Using = SAVEBTN_CUST_PRICE)]
        private IWebElement _create;

        [FindsBy(How = How.XPath, Using = ERROR_PRICE_REQUIRED)]
        private IWebElement _errorPrice;

        [FindsBy(How = How.XPath, Using = CONFIRM_OK)]
        private IWebElement _confirmOk;

        [FindsBy(How = How.Id, Using = NEW_SCALE)]
        private IWebElement _newScale;

        // ____________________________________________ Méthodes __________________________________________________

        public ServicePricePage FillFields_CustomerPrice(string site, string customer, DateTime fromDate, DateTime toDate, string price = null, string datasheet = null, string srv = null, string idCustomersrv = null)
        {
            _priceSite = WaitForElementIsVisibleNew(By.Id(SITE));
            _priceSite.SetValue(ControlType.DropDownList, site);
            WaitForLoad();

            var _divCustomer = WaitForElementIsVisibleNew(By.XPath("//div[@class='item']"));
            new Actions(_webDriver).MoveToElement(_divCustomer).Click().Perform();
            WaitForLoad();

            _customerName = WaitForElementIsVisibleNew(By.Id(CUSTOMER_NAME));
            _customerName.Clear();
            new Actions(_webDriver).MoveToElement(_divCustomer).Click().Perform();
            WaitForLoad();

            _customerName = WaitForElementIsVisibleNew(By.Id(CUSTOMER_NAME));
            _customerName.SendKeys(customer);
            WaitForLoad();

            try
            {
                Thread.Sleep(2000);
                var _firstCustomer = WaitForElementIsVisibleNew(By.XPath(CUSTOMER_SELECTED));
                _firstCustomer.Click();
                WaitForLoad();
            }
            catch (Exception)
            {
                //silent catch => this method crashes if you try to click the first element of the list (which is already selected)
            }
           WaitLoading();
            _fromDate = WaitForElementIsVisibleNew(By.Id(BEGIN_DATE));
            _fromDate.SetValue(ControlType.DateTime, fromDate);
            _fromDate.SendKeys(Keys.Tab);
            WaitForLoad();

            _toDate = WaitForElementIsVisibleNew(By.Id(END_DATE));
            _toDate.SetValue(ControlType.DateTime, toDate);
            _toDate.SendKeys(Keys.Tab);
            WaitForLoad();

            if (datasheet != null)
            {
                _datasheetName = WaitForElementIsVisibleNew(By.XPath("//*[@id=\"modal-1\"]/div/div/form/div[2]/div[3]/div/div/div[3]/div/span[1]/input[2]"));
                _datasheetName.SetValue(ControlType.TextBox, datasheet);
                WaitForLoad();

                var selectedDatasheet = WaitForElementIsVisibleNew(By.XPath(String.Format("//*[@id=\"modal-1\"]/div/div/form/div[2]/div[3]/div/div/div[3]/div/span[1]/div/div/div[text()=\"{0}\"]", datasheet)));
                selectedDatasheet.Click();
                WaitForLoad();
            }

            if (price != null)
            {
                _invoicePrice = WaitForElementIsVisibleNew(By.Id(INVOICE_PRICE));
                _invoicePrice.SetValue(ControlType.TextBox, price);
                WaitForLoad();
            }

            if (srv != null)
            {
                var srvInput = WaitForElementIsVisibleNew(By.Id("item_CustomerSrv"));
                srvInput.SetValue(ControlType.TextBox, srv);
                WaitForLoad();
            }

            if (idCustomersrv != null)
            {
                var idCustsrvInput = WaitForElementIsVisibleNew(By.Id("item_CustomerSrvId"));
                idCustsrvInput.SetValue(ControlType.TextBox, idCustomersrv);
                WaitForLoad();
            }

            _create = WaitForElementIsVisibleNew(By.Id(SAVEBTN_CUST_PRICE));
            _create.Click();
            WaitPageLoading();
            WaitForLoad();

            return new ServicePricePage(_webDriver, _testContext);
        }


        public void setDatasheet(string datasheet)
        {
            _datasheetName = WaitForElementIsVisible(By.XPath("//*[@id=\"modal-1\"]/div/div/form/div[2]/div[3]/div/div/div[3]/div/span[1]/input[2]"));
            new Actions(_webDriver).MoveToElement(_datasheetName).Click().Perform();
            _datasheetName.SendKeys(Keys.Control + "a");
            _datasheetName.SendKeys(Keys.Backspace);
            _datasheetName.SendKeys(datasheet);
            WaitForLoad();

            _selectedDatasheet = WaitForElementIsVisible(By.XPath(string.Format("//div[text()='{0}']", datasheet)));
            new Actions(_webDriver).MoveToElement(_selectedDatasheet).Click().Perform();

            // _selectedDatasheet.Click();
            WaitForLoad();

            _create = WaitForElementIsVisible(By.Id(SAVEBTN_CUST_PRICE));
            _create.Click();
            WaitForLoad();
        }
        public ServicePricePage FillFields_EditCustomerPrice(string site, string customer, DateTime fromDate, DateTime toDate)
        {
            _priceSite = WaitForElementIsVisible(By.Id(SITE));
            _priceSite.SetValue(ControlType.DropDownList, site);
            WaitForLoad();

            var customerDiv = WaitForElementIsVisible(By.XPath(CUSTOMER_DIV));
            customerDiv.Click();
            WaitForLoad();
            _customer = WaitForElementIsVisible(By.Id(CUSTOMER));
            _customer.SendKeys(customer);
            WaitForLoad();
            var customerSelected = WaitForElementToBeClickable(By.XPath(CUSTOMER_SELECTED));
            customerSelected.Click();
            WaitForLoad();

            _fromDate = WaitForElementIsVisible(By.Id(BEGIN_DATE));
            _fromDate.SetValue(ControlType.DateTime, fromDate);
            _fromDate.SendKeys(Keys.Tab);
            WaitForLoad();

            _toDate = WaitForElementIsVisible(By.Id(END_DATE));
            _toDate.SetValue(ControlType.DateTime, toDate);
            _toDate.SendKeys(Keys.Tab);
            WaitForLoad();

            _create = WaitForElementIsVisible(By.Id(SAVEBTN_CUST_PRICE));
            _create.Click();
            WaitForLoad();
            WaitPageLoading();

            return new ServicePricePage(_webDriver, _testContext);
        }

        public ServicePricePage EditPriceDates(DateTime fromDate, DateTime toDate, string datasheet = null)
        {
            WaitForLoad();
            if (isElementVisible(By.Id(BEGIN_DATE)))
            {
                _fromDate = WaitForElementIsVisible(By.Id(BEGIN_DATE));
            }
            else
            {
                _fromDate = _webDriver.FindElement(By.Id(OTHER_BEGIN_DATE));
            }
            new Actions(_webDriver).MoveToElement(_fromDate).Click().Perform();
            _fromDate.SetValue(ControlType.DateTime, fromDate);
            WaitForLoad();
            _fromDate.SendKeys(Keys.Tab);

            if (isElementVisible(By.Id(OTHER_END_DATE)))
            {
                _toDate = _webDriver.FindElement(By.Id(OTHER_END_DATE));
            }
            else
            {
                _toDate = _webDriver.FindElement(By.Id(END_DATE));
            }

            _toDate.SetValue(ControlType.DateTime, toDate);
            _toDate.SendKeys(Keys.Tab);
            WaitForLoad();

            if (datasheet != null)
            {
                var _datasheetName = WaitForElementIsVisible(By.XPath("//*[@id=\"modal-1\"]/div/div/form/div[2]/div[3]/div/div/div[3]/div/span[1]/input[2]"));
                if (_datasheetName.GetAttribute("disabled") == "false" || _datasheetName.GetAttribute("disabled") == null)
                {
                    //new Actions(_webDriver).MoveToElement(_datasheetName).Click().Perform();
                    _datasheetName.SendKeys(Keys.Control + "a");
                    _datasheetName.SendKeys(Keys.Backspace);
                    _datasheetName.SendKeys(datasheet);
                    WaitForLoad();

                    var selectedDatasheet = WaitForElementIsVisible(By.XPath(String.Format("//div[text()='{0}']", datasheet)));
                    selectedDatasheet.Click();
                    WaitForLoad();
                }
            }

            _create = WaitForElementIsVisible(By.Id(SAVEBTN_CUST_PRICE));
            _create.Click();
            WaitForLoad();

            // Temps de fermeture de la fenêtre
            WaitForLoad();

            return new ServicePricePage(_webDriver, _testContext);
        }
        public bool VerifySuccessEditPriceDates()
        {
            return !isElementVisible(By.XPath("//*/span[contains(@class,'text-danger') and text()='A price already exists for this period']"));
        }
        public string GetErrorPriceAlreadyExist()
        {
            _errorPrice = WaitForElementIsVisible(By.XPath(ERROR_PRICE_REQUIRED));
            return _errorPrice.Text;
        }

        public string GetPrice()
        {
            var price = WaitForElementIsVisible(By.XPath("//*[@id=\"item_InvoicePrice\"]"));
            WaitForLoad();
            return price.GetAttribute("value");
        }

        public string GetDefaultMode()
        {
            var value = WaitForElementIsVisible(By.XPath("//*[@id=\"item_DefaultMode\"]/option[@selected = 'selected']"));
            var text = value.Text;
            Close();
            return text;
        }

        public string GetDefaultModeValue()
        {
            var value = WaitForElementIsVisible(By.Id("item_Value"));
            var Defaultvalue = value.GetAttribute("value");
            Close();
            return Defaultvalue;
        }

        public string GetIdCustomerSrv()
        {
            var value = WaitForElementIsVisible(By.Id("item_CustomerSrvId"));
            return value.GetAttribute("value");
        }

        public new void Close()
        {
            var close = WaitForElementIsVisible(By.Id("btn-Cancel"));
            close.Click();
            WaitForLoad();
        }
        public string GetCustomerName()
        {
            var element = _webDriver.FindElement(By.XPath("//*[@id='list-item-with-action']/div[2]/div[1]/div/div[2]/table/tbody/tr/td[1]"));
            string text = element.Text;
            var parts = text.Split('/');

            var textAfterSlash = string.Join("/", parts.Skip(1)).Trim();
            return textAfterSlash;
        }

        public ServiceEditLoadingScale SetModeScale()
        {
            _sclaemode = WaitForElementExists(By.XPath(SCALEMODE));
            _sclaemode.Click();
            WaitForLoad();
            return new ServiceEditLoadingScale(_webDriver, _testContext);
        }
        public ServiceEditLoadingScale ClickEditModeScale()
        {
            _editscalemode = WaitForElementExists(By.XPath(EDITSCALEMODE));
            _editscalemode.Click();
            WaitForLoad();
            return new ServiceEditLoadingScale(_webDriver, _testContext);
        }
        public void SetPourcentageMode()
        {
            _pourcentagemode = WaitForElementExists(By.XPath(POURCENTAGEMODE));
            _pourcentagemode.Click();
            WaitForLoad();
        }
        public bool Vrifylabelroundexist(string labelround)
        {
            _labelround = WaitForElementExists(By.XPath(LABELROUND));
            if (_labelround.Text.Equals(labelround))
                return true;
            return false;
        }
        public bool VerifyOptionsRoundExist()
        {
            var listeoptionround = _webDriver.FindElements(By.XPath(LISTEOPTIONROUND));

            if (listeoptionround.Count == 3)
                return true;

            return false;
        }

        public string GetDataSheet()
        {
            _datasheetName = WaitForElementIsVisible(By.XPath("//*[@id=\"modal-1\"]/div/div/form/div[2]/div[3]/div/div/div[3]/div/span[1]/input[2]"));

            return _datasheetName.Text;
        }
        public void SetMethodWithoutSAVE(string method, int scaleNbre = 0, string totalPrice = "1")
        {
            DropdownListSelectById(METHOD, method);
            switch (method)
            {
                case "Scale":
                    for (int i = 0; i < scaleNbre; i++)
                    {
                        var newScale = WaitForElementIsVisible(By.Id("btn-add-scale"));
                        newScale.Click();
                        var totalPriceInput = WaitForElementIsVisible(By.XPath(string.Format("//*[@id=\"scales-table\"]/tbody/tr[{0}]/td[2]/input", i + 2)));
                        totalPriceInput.SetValue(ControlType.TextBox, totalPrice);
                        totalPriceInput.SendKeys(Keys.Enter);
                        WaitForLoad();

                    }

                    break;
                    /* continue logic of Cycle here..
                case "Cycle":
                    var newScale = WaitForElementIsVisible(By.Id("btn-add-scale"));
                    newScale.Click();
                    break;
                    */
            }
        }
        public void DeleteScaleWithoutSave(int scaleNombre)
        {
            var scale = WaitForElementIsVisible(By.XPath(string.Format(SERVICE_DELETE_SCALE, scaleNombre + 1)));
            scale.Click();
            WaitForLoad();
            _confirmOk = WaitForElementIsVisible(By.Id(CONFIRM_OK));
            _confirmOk.Click();
            WaitForLoad();

        }
        public ServicePricePage Save()
        {
            _create = WaitForElementIsVisible(By.Id(SAVEBTN_CUST_PRICE));
            _create.Click();
            WaitForLoad();
            return new ServicePricePage(_webDriver, _testContext);
        }
        public List<string> GetAllScalePrice()
        {
            List<string> priceScale = new List<string>();
            var scales = _webDriver.FindElements(By.XPath(ALL_SCALE_PRICE));
            if (scales == null)
            {
                return new List<string>();
            }
            foreach (var scale in scales)
            {
                priceScale.Add(scale.GetAttribute("value"));
            }
            return priceScale;
        }

        public void SetMethodScaleWithZeroNbPersonsWithoutSAVE(string method, int scaleNbre = 0, string totalPrice = "1", string nbPersons = "0")
        {
            DropdownListSelectById(METHOD, method);
            switch (method)
            {
                case "Scale":
                    for (int i = 0; i < scaleNbre; i++)
                    {
                        _newScale = WaitForElementIsVisible(By.Id(NEW_SCALE));
                        _newScale.Click();
                        var totalPriceInput = WaitForElementIsVisible(By.XPath(string.Format("//*[@id=\"scales-table\"]/tbody/tr[{0}]/td[2]/input", i + 2)));
                        totalPriceInput.SetValue(ControlType.TextBox, totalPrice);
                        totalPriceInput.SendKeys(Keys.Enter);

                        var nbPersonsInput = WaitForElementIsVisible(By.XPath(string.Format("//*[@id=\"scales-table\"]/tbody/tr[{0}]/td[1]/input", i + 2)));
                        nbPersonsInput.SetValue(ControlType.TextBox, nbPersons);
                        nbPersonsInput.SendKeys(Keys.Enter);

                        WaitForLoad();

                    }

                    break;
            }
        }
        public void SetMethodeCycle()
        {
            _methodcycle = WaitForElementExists(By.XPath(METHOD_CYCLE));
            _methodcycle.Click();
            WaitForLoad();
        }
        public void SetCycleMode()
        {
            _cyclemodecontinuous = WaitForElementExists(By.XPath(CYCLE_MODE_CONTINUOUS));
            _cyclemodecontinuous.Click();
            WaitForLoad();
        }
        public void SetNbCycle(string nb)
        {
            _nb_cycle = WaitForElementIsVisible(By.Id(NB_CYCLE));
            _nb_cycle.SetValue(ControlType.TextBox, nb);
            _nb_cycle.SendKeys(Keys.Enter);
            WaitForLoad();
        }
        public void SetMethodescale()
        {
            _methodscale = WaitForElementExists(By.XPath(METHODE_SCALE));
            _methodscale.Click();
            WaitForLoad();
        }
        public int NbCycles()
        {
            var nbcycle = _webDriver.FindElements(By.XPath(NB_CYCLE_TABLE));
            return nbcycle.Count;
        }
        public void SetTextDataSheet(int i, string datasheet)
        {
            var setdatasheet = WaitForElementIsVisible(By.XPath("/html/body/div[4]/div/div/div/div/form/div[2]/div[7]/div/table/tbody/tr[" + i + "]/td[3]/span/input[2]"));
            setdatasheet.SetValue(ControlType.TextBox, datasheet);
            WaitForLoad();
        }
        public ServicePricePageEditScaleForCycle GetEditScaleCycle(int i)
        {
            _cyclemodecontinuous = WaitForElementExists(By.XPath("/html/body/div[4]/div/div/div/div/form/div[2]/div[7]/div/table/tbody/tr[" + i + "]/td[4]/select/option[2]"));
            _cyclemodecontinuous.Click();
            WaitForLoad();
            return new ServicePricePageEditScaleForCycle(_webDriver, _testContext);
        }
        public ServicePricePage Save_Scale_Mode()
        {

            WaitForLoad();
            _btn_save_edit_price = WaitForElementIsVisible(By.XPath(BTN_SAVE_EDIT_PRICE));
            _btn_save_edit_price.Click();
            WaitForLoad();
            return new ServicePricePage(_webDriver, _testContext);
        }
        public void SelectMethodeScaleCycle(int i)
        {
            Actions action = new Actions(_webDriver);
            var selectrow = WaitForElementExists(By.XPath("/html/body/div[4]/div/div/div/div/form/div[2]/div[7]/div/table/tbody/tr[" + i + "]"));
            WaitForLoad();
            action.MoveToElement(selectrow).Click().Perform();

            WaitForLoad();
        }
        public ServicePricePageEditScaleForCycle GetEditScaleCycleRow(int i)
        {
            WaitForLoad();
            var row_cycle = WaitForElementIsVisible(By.XPath("/html/body/div[4]/div/div/div/div/form/div[2]/div[7]/div/table/tbody/tr[" + i + "]/td[4]/a/span"));
            row_cycle.Click();
            WaitForLoad();

            return new ServicePricePageEditScaleForCycle(_webDriver, _testContext);
        }
        public void SetCycleMode_EndOfMonth()
        {
            _end_of_month = WaitForElementExists(By.XPath(CYCLE_END_Of_MONTH));
            _end_of_month.Click();
            WaitForLoad();
        }
        public void SetFirstDuration(string duration)
        {
            _first_duration_days = WaitForElementIsVisible(By.XPath(FIRST_DURATION_DAYS));
            _first_duration_days.SetValue(ControlType.TextBox, duration);
            WaitForLoad();
        }
        public void SetCycleMode_FullMonth()
        {
            _full_month = WaitForElementExists(By.XPath(CYCLE_EFULL_MONTH));
            _full_month.Click();
            WaitForLoad();
        }
        public string GetServicePeriodFrom()
        {
            _fromDate = WaitForElementIsVisible(By.Id(BEGIN_DATE));
            WaitForLoad();
            return _fromDate.GetAttribute("value");
        }
        public string GetServicePeriodTo()
        {
            _toDate = WaitForElementIsVisible(By.Id(END_DATE));
            WaitForLoad();
            return _toDate.GetAttribute("value");
        }
        public void SetMethodeFixed()
        {
            _methode_fixed = WaitForElementExists(By.XPath(METHOD_FIXED));
            _methode_fixed.Click();
            WaitForLoad();
        }
        public void SetMethodscale()
        {
            _method_scale = WaitForElementExists(By.XPath(METHOD_SCALE));
            _method_scale.Click();
            WaitForLoad();
        }
        public List<string> GetAll_LabelMethodeFixed()
        {
            List<string> listLabelmethodefixed = new List<string>(); 
            _label_datasheet = WaitForElementIsVisible(By.XPath(LABEL_DATASHEET)); 
            WaitForLoad();
            _label_invoise_price = WaitForElementIsVisible(By.XPath(LABEL_INVOISE_PRICE));
            WaitForLoad();
            listLabelmethodefixed.Add(_label_datasheet.Text);
            listLabelmethodefixed.Add(_label_invoise_price.Text);
            return listLabelmethodefixed;
        }
        public List<string> GetAll_LabelMethodeScale()
        {
            List<string> listLabelmethodeScale = new List<string>();
            _label_datasheet_scale = WaitForElementIsVisible(By.XPath(LABEL_DATASHEET_SCALE));
            WaitForLoad();
            _label_nb_of_person = WaitForElementIsVisible(By.XPath(LABEL_NB_OF_PERSON));
            WaitForLoad();
            _label_total_price = WaitForElementIsVisible(By.XPath(LABEL_TOTAL_PRICE));
            WaitForLoad();
            _label_unit_price = WaitForElementIsVisible(By.XPath(LABEL_UNIT_PRICE));
            WaitForLoad();
            _label_new_scale = WaitForElementIsVisible(By.XPath(LABEL_NEW_SCALE));
            WaitForLoad();
            listLabelmethodeScale.Add(_label_datasheet_scale.Text);
            listLabelmethodeScale.Add(_label_nb_of_person.Text);
            listLabelmethodeScale.Add(_label_total_price.Text);
            listLabelmethodeScale.Add(_label_unit_price.Text);
            listLabelmethodeScale.Add(_label_new_scale.GetAttribute("value"));
            return listLabelmethodeScale;
        }
        public List<string> GetAll_LabelMethodeCycle()
        {
            List<string> listLabelmethodeCycle = new List<string>();
            _label_cycle_mode = WaitForElementIsVisible(By.XPath(LABEL_CYCLE_MODE));
            WaitForLoad();
            _label_nb_cycle = WaitForElementIsVisible(By.XPath(LABELNB_CYCLE));
            WaitForLoad();
            _label_start_at = WaitForElementIsVisible(By.XPath(LABEL_START_AT));
            WaitForLoad();
            _label_with_duration = WaitForElementIsVisible(By.XPath(LABEL_WITH_DURATION));
            WaitForLoad();
            _label_name = WaitForElementIsVisible(By.XPath(LABEL_NAME));
            WaitForLoad();
            _label_duration = WaitForElementIsVisible(By.XPath(LABEL_DURATION));
            WaitForLoad();
            _label_datasheet_cycle = WaitForElementIsVisible(By.XPath(LABEL_DATASHEET_CYCLE));
            WaitForLoad();
            _label_method_scale = WaitForElementIsVisible(By.XPath(LABEL_METHOD_CYCLE));
            WaitForLoad();
            _label_price = WaitForElementIsVisible(By.XPath(LABEL_PRICE));
            WaitForLoad();

            listLabelmethodeCycle.Add(_label_cycle_mode.Text);
            listLabelmethodeCycle.Add(_label_nb_cycle.Text);
            listLabelmethodeCycle.Add(_label_start_at.Text);
            listLabelmethodeCycle.Add(_label_with_duration.Text);
            listLabelmethodeCycle.Add(_label_name.Text);
            listLabelmethodeCycle.Add(_label_duration.Text);
            listLabelmethodeCycle.Add(_label_datasheet_cycle.Text);
            listLabelmethodeCycle.Add(_label_method_scale.Text);
            listLabelmethodeCycle.Add(_label_price.Text);
            return listLabelmethodeCycle;
        }
    }
}
