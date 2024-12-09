using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using SeleniumExtras.WaitHelpers;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.CustomerPortal
{
    public class CustomerPortalCustomerOrdersPage : PageBase
    {
        public CustomerPortalCustomerOrdersPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // _________________________________________ Constantes _______________________________________________


        private const string FILTER_CUSTOMERS = "//*[@id=\"SelectedCustomers_ms\"]";
        private const string FILTER_CUSTOMERS_SELECT = "/ html/body/div[4]/ul/li[*]/label/span[contains(text(), \"{0}\")]";

        private const string FILTER_CUSTOMERS_UNCHECK_ALL = "/html/body/div[4]/div/ul/li[2]/a";
        private const string FILTER_CUSTOMERS_SELECT_SEARCH = "/html/body/div[4]/div/div/label/input";

        private const string FILTER_CUSTOMERS_ORDER = "//*[@id=\"formOrderList\"]/div/div[6]/a";
        private const string FILTER_CUSTOMERS_ORDER_RADIO_BUTTON = "//*[@id=\"{0}\"]";

        private const string TABLE_TITLE = "/html/body/div[3]/div[2]/div[1]/div[1]/h2/span[1]";

        private const string TABLE_LINE1_COLUMN_CUSTOMER_ORDER_NUMBER = "//*[@id=\"tableListMenu\"]/tbody/tr[{0}]/td[4]/a";
        private const string TABLE_LINE1_COLUMN_CUSTOMER_NAME = "//*[@id=\"tableListMenu\"]/tbody/tr[{0}]/td[5]/b";
        private const string TABLE_LINE1_COLUMN_DELIVERY_FLIGHT_NUMBER = "//*[@id=\"tableListMenu\"]/tbody/tr[{0}]/td[6]";
        private const string TABLE_LINE1_COLUMN_ORDER_DATE_HOUR = "//*[@id=\"tableListMenu\"]/tbody/tr[{0}]/td[8]";
        private const string TABLE_LINE1_COLUMN_PRICE = "//*[@id=\"tableListMenu\"]/tbody/tr[{0}]/td[10]";
        private const string TABLE_LINE1_COLUMN_EXPEDITION_DATE = "//*[@id=\"tableListMenu\"]/tbody/tr[{0}]/td[9]";


        private const string TABLE_LINE1_COLUMN_CUSTOMER_ORDER_NUMBER_PATCH = "//*[@id=\"tableListMenu\"]/tbody/tr[{0}]/td[5]/a";
        private const string TABLE_LINE1_COLUMN_CUSTOMER_NAME_PATCH = "//*[@id=\"tableListMenu\"]/tbody/tr[{0}]/td[6]/b";
        private const string TABLE_LINE1_COLUMN_DELIVERY_FLIGHT_NUMBER_PATCH = "//*[@id=\"tableListMenu\"]/tbody/tr[{0}]/td[7]";
        private const string TABLE_LINE1_COLUMN_ORDER_DATE_HOUR_PATCH = "//*[@id=\"tableListMenu\"]/tbody/tr[{0}]/td[9]";
        private const string TABLE_LINE1_COLUMN_PRICE_PATCH = "//*[@id=\"tableListMenu\"]/tbody/tr[{0}]/td[10]";
        private const string NB_LIGNES = "//*/table[@id='tableListMenu']/tbody/tr[*]";

        private const string FILTER_SEARCH = "//*[@id=\"SearchPattern\"]";
        private const string FILTER_DATE_FROM = "//*[@id=\"start-date-picker\"]";
        private const string FILTER_DATE_TO = "//*[@id=\"end-date-picker\"]";

        private const string RESET_FILTER = "//*[@id=\"formOrderList\"]/div/div[1]/a";

        private const string ICON1 = "//tr[1]/td[1]/img";
        private const string ICON2 = "//tr[1]/td[2]/span";
        private const string ICON3 = "//tr[1]/td[3]/span";

        private const string DELETE_ICON = "//tr[1]//span[contains(@class, 'trash')]";
        private const string DELETE_CONFIRM = "//*[@id=\"dataConfirmOK\"]";

        private const string PRINT_BUTTON = "btnPrint";
        private const string EXPORT_EXCEL_BUTTON = "btnExportCO";
        private const string EXPORT_EXCEL_CONFIRM = "//*[@id=\"btn-export-co\"]";

        private const string FIRST_CUSTOMER_ORDER_LINK = "//*[@id=\"tableListMenu\"]/tbody/tr[1]/td[5]/a";
        private const string FIRST_CUSTOMER_DLIVERY_DATE_PLUS_HOUR = "//*[@id=\"tableListMenu\"]/tbody/tr[1]/td[8]";
        private const string FIRST_CUSTOMER_ORDER_COMMENT = "//*[@id=\"tableListMenu\"]/tbody/tr[1]/td[7]/span";

        private const string CUTOMER_ORDERS_ACTIONS_BTN = "customerportal-customerorders-dropdown-actions-btn";

        private const string DELETE_BTN = "//*/span[contains(@class,'trash')]/..";
        private const string VAT = "//*[@id=\"formOrderDetailItems\"]/div[2]/div/div/b";
        public const string COMERCIALNAME = "//*[@id=\"dispatchTable\"]/tbody/tr[2]/td[2]";
        public const string QTY = "//*[@id=\"dispatchTable\"]/tbody/tr[2]/td[3]";
        public const string UNIT = "//*[@id=\"dispatchTable\"]/tbody/tr[2]/td[4]";
        public const string PRICE = "//*[@id=\"dispatchTable\"]/tbody/tr[2]/td[5]";
        private const string FIRSTLINECLICK = "//*[@id=\"tableListMenu\"]/tbody/tr";
        private const string CLICK_ADD = "/html/body/div[4]/div/div[2]/div/div/div/div/form/div[1]/div[1]/div/div/div/button[2]";
        private const string SELECT_ITEM = "/html/body/div[4]/div/div[2]/div/div/div/div/form/div[1]/div[2]/div/div/table/tbody/tr[2]/td[2]/span/span/input[2]";
        public const string ITEM_SELECTED = "//*[@id=\"dispatchTable\"]/tbody/tr[2]/td[2]/span/span/div/div/div";
        private const string PRICE_LIGNES = "//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[10]";

        [FindsBy(How = How.XPath, Using = DELETE_BTN)]
        private IWebElement _deletebtn;
        [FindsBy(How = How.XPath, Using = VAT)]
        private IWebElement _vat;
        [FindsBy(How = How.XPath, Using = COMERCIALNAME)]
        private IWebElement _commercialname;
        [FindsBy(How = How.XPath, Using = QTY)]
        private IWebElement _qty;
        [FindsBy(How = How.XPath, Using = UNIT)]
        private IWebElement _unit;
        [FindsBy(How = How.XPath, Using = PRICE)]
        private IWebElement _price;
        [FindsBy(How = How.XPath, Using = DELETE_CONFIRM)]
        private IWebElement _confirmDelete;

        [FindsBy(How = How.XPath, Using = CLICK_ADD)]
        private IWebElement _clickAdd;

        [FindsBy(How = How.XPath, Using = SELECT_ITEM)]
        private IWebElement _selectItem;

        [FindsBy(How = How.XPath, Using = ITEM_SELECTED)]
        private IWebElement _itemSelected;


        public void ShowFilterCustomersContent()
        {
            var select = By.XPath(FILTER_CUSTOMERS);
            _webDriver.FindElement(select).Click();
        }

        /**
         * première colonne pleine : offset=1,noLigne=1
         */
        public string getTableFirstLine(int offset, int noLigne = 1)
        {
            By td = null;
            switch (offset)
            {
                case 1:
                    td = By.XPath(TABLE_LINE1_COLUMN_CUSTOMER_ORDER_NUMBER.Replace("{0}", noLigne.ToString()));
                    break;
                case 2:
                    td = By.XPath(TABLE_LINE1_COLUMN_CUSTOMER_NAME.Replace("{0}", noLigne.ToString()));
                    break;
                case 3:
                    td = By.XPath(TABLE_LINE1_COLUMN_DELIVERY_FLIGHT_NUMBER.Replace("{0}", noLigne.ToString()));
                    break;
                case 4:
                    td = By.XPath(TABLE_LINE1_COLUMN_ORDER_DATE_HOUR.Replace("{0}", noLigne.ToString()));
                    break;
                case 5:
                    td = By.XPath(TABLE_LINE1_COLUMN_PRICE.Replace("{0}", noLigne.ToString()));
                    break;
                case 9:
                    td = By.XPath(TABLE_LINE1_COLUMN_EXPEDITION_DATE.Replace("{0}", noLigne.ToString()));
                    break;
            }

            // parfois le tableau est scroll, du coup "non visible", du coup on n'utilise pas WaitForElementIsVisible ici
            WaitPageLoading();
            WaitForLoad();
            return WaitForElementExists(td).Text;
        }

        public void Navigate(string url)
        {
            IJavaScriptExecutor executor = (IJavaScriptExecutor)_webDriver;
            executor.ExecuteScript("window.location.replace(\"" + url + "\");");
        }

        public CustomerPortalCustomerOrdersPageModal ClickAdd()
        {
            IJavaScriptExecutor executor = (IJavaScriptExecutor)_webDriver;
            executor.ExecuteScript("showOrderCreation();");
            WaitPageLoading();
            WaitForLoad();
            return new CustomerPortalCustomerOrdersPageModal(_webDriver, _testContext);
        }

        public void Search(string deliveryName)
        {
            var search = By.XPath(FILTER_SEARCH);
            var searchInput = WaitForElementIsVisible(search);
            searchInput.Clear();
            searchInput.SendKeys(deliveryName);
            searchInput.Submit();
            WaitPageLoading();
        }

        public void FilterDate(DateTime? dateFrom, DateTime? dateTo)
        {
            if (dateFrom != null)
            {
                var from = By.XPath(FILTER_DATE_FROM);
                var fromInput = WaitForElementIsVisible(from);
                fromInput.Clear();
                fromInput.SendKeys(dateFrom?.ToString("dd/MM/yyyy"));
                fromInput.SendKeys(Keys.Tab);
            }
            if (dateTo != null)
            {
                var to = By.XPath(FILTER_DATE_TO);
                var toInput = WaitForElementIsVisible(to);
                toInput.Clear();
                toInput.SendKeys(dateTo?.ToString("dd/MM/yyyy"));
                toInput.SendKeys(Keys.Tab);
            }
            WaitPageLoading();
        }
        public void FilterCustomersUnCheckAll()
        {
            var uncheckAll = WaitForElementIsVisible(By.XPath(FILTER_CUSTOMERS_UNCHECK_ALL));
            uncheckAll.Click();
        }

        public string FilterCustomerSelect(string customer)
        {
            MakeSureCustomerFilterIsOpen();
            // recherche l'élément
            var search = WaitForElementIsVisible(By.XPath(FILTER_CUSTOMERS_SELECT_SEARCH));
            search.Clear();
            search.SendKeys(customer);
            search.Click();

            // cherche en filtrant sur CUSTOMER
            var span = WaitForElementIsVisible(By.XPath(string.Format(FILTER_CUSTOMERS_SELECT, customer)));
            var text = span.Text;
            var checkbox = span.FindElement(By.XPath(string.Format(FILTER_CUSTOMERS_SELECT, customer) + "//parent::label/input"));
            if (!checkbox.Selected)
            {
                checkbox.Click();
            }
            MakeSureCustomerFilterIsClosed();
            return text;
        }

        public void MakeSureCustomerFilterIsClosed()
        {
            if (isElementVisible(By.XPath("/html/body/div[4]/div/div/label/input")))
            {
                var _customersFilter = WaitForElementIsVisible(By.Id("SelectedCustomers_ms"));
                _customersFilter.Click();
            }
        }


        public void MakeSureCustomerFilterIsOpen()
        {
            if (!isElementVisible(By.XPath("/html/body/div[4]/div/div/label/input")))
            {
                var _customersFilter = WaitForElementIsVisible(By.Id("SelectedCustomers_ms"));
                _customersFilter.Click();
            }
        }

        public string CheckIconsValidated()
        {
            //valid check
            var icon1 = WaitForElementIsVisible(By.XPath("//tr[1]/td[1]/div/i[@class='fa-regular fa-circle-check']"));

            // $ orange
            var icon2 = WaitForElementIsVisible(By.XPath("//tr[1]/td[2]/div/i"));
            if ("fas fa-dollar-sign" != icon2.GetAttribute("class"))
            {
                return "icone not invoiced non présente";
            }
            //eye
            var icon3 = WaitForElementIsVisible(By.XPath("//tr[1]/td[3]/div/span"));
            if ("fas fa-eye" != icon3.GetAttribute("class"))
            {
                return "icone not verified non présente";
            }
            return "OK";
        }

        public string CheckIconsInProgress()
        {
            if (_webDriver.FindElements(By.XPath(ICON1)).Count != 0)
            {
                return "une icone validation";
            }
            if (_webDriver.FindElements(By.XPath(ICON2)).Count != 0)
            {
                return "une icone dollars";
            }
            if (_webDriver.FindElements(By.XPath("//tr[1]/td[3]/div/span")).Count != 1)
            {
                return "pas d'icone oeil";
            }

            return "OK";
        }

        public void DeleteFirstItem()
        {
            // stabilisation
            WaitForLoad();
            var deleteIcon = WaitForElementIsVisible(By.XPath(DELETE_ICON));

            deleteIcon.Click();
            // animation popup
            WaitForLoad();
            var deleteConfirm = WaitForElementIsVisible(By.XPath(DELETE_CONFIRM));
            deleteConfirm.Click();
            // rafraichissement du tableau
            WaitForLoad();
        }

        public CustomerPortalCustomerOrdersPageResult GoToFirstCustomer()
        {
            IWebElement LienOrderNo;
            WaitForLoad();
            WaitForLoad();
            LienOrderNo = WaitForElementIsVisible(By.XPath("//*[@id=\"tableListMenu\"]/tbody/tr[1]/td[4]/a"));

            LienOrderNo.Click();
            WaitForLoad();
            return new CustomerPortalCustomerOrdersPageResult(_webDriver, _testContext);
        }

        public void ShowFilterCustomersOrder(string radioId, bool toggle = true)
        {
            if (toggle)
            {
                ShowFilterCustomersOrder();
            }
            var orderStatusValidated = WaitForElementIsVisible(By.XPath(FILTER_CUSTOMERS_ORDER_RADIO_BUTTON.Replace("{0}", radioId)));
            orderStatusValidated.Click();
            WaitPageLoading();//long to load for invoiced
            WaitForLoad();
        }

        public void ShowFilterCustomersOrder()
        {
            var orderStatus = WaitForElementIsVisible(By.XPath("//*[@id=\"formOrderList\"]/div[6]/a"));
            orderStatus.Click();
            WaitForLoad();
            WaitPageLoading();
        }

        public bool ShowFilterCustomersOrderIsSelected(string radioId)
        {
            ShowFilterCustomersOrder();
            var orderStatusValidated = WaitForElementIsVisible(By.XPath(FILTER_CUSTOMERS_ORDER_RADIO_BUTTON.Replace("{0}", radioId)));
            return orderStatusValidated.Selected;
        }

        public void ResetFilter()
        {
            var resetFilter = WaitForElementIsVisible(By.XPath("//*[@id=\"formOrderList\"]/div[1]/a"));
            resetFilter.Click();
            WaitForLoad();
            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                //Assert : Filter Date To pas vide
                //FilterDate(null, DateUtils.Now);
                // ne pas fournir de date de fin.
            }
        }

        public string SearchInput()
        {
            var search = WaitForElementIsVisible(By.XPath(FILTER_SEARCH));
            return search.GetAttribute("value");
        }

        internal object DateToInput()
        {
            var to = WaitForElementIsVisible(By.XPath(FILTER_DATE_TO));
            return to.GetAttribute("value");
        }

        internal object DateFromInput()
        {
            var from = WaitForElementIsVisible(By.XPath(FILTER_DATE_FROM));
            return from.GetAttribute("value");
        }

        public string FilterCustomerLabel()
        {
            var customersLabel = WaitForElementIsVisible(By.XPath(FILTER_CUSTOMERS));
            return customersLabel.Text;
        }

        public void Print()
        {

            HoverOverActionsBtn();
            WaitForLoad();

            var print = WaitForElementIsVisible(By.Id(PRINT_BUTTON));
            print.Click();

            // détection alert : 4 secondes
            // si alert, throw
            Thread.Sleep(4000);
            IAlert alert = ExpectedConditions.AlertIsPresent().Invoke(_webDriver);
            if (alert != null)
            {
                throw new Exception("pas de fichier téléchargeable : " + _webDriver.SwitchTo().Alert().Text);
            }

            // téléchargement
            WaitForDownload();
            WaitForLoad();
        }
        public void HoverOverActionsBtn()
        {
            var actions = new Actions(_webDriver);
            var actionsBtn = WaitForElementIsVisible(By.Id(CUTOMER_ORDERS_ACTIONS_BTN));
            actions.MoveToElement(actionsBtn).Perform();
        }
        public void ExportExcel()
        {
            Actions actions = new Actions(_webDriver);
            var actionsBtn = WaitForElementIsVisible(By.Id(CUTOMER_ORDERS_ACTIONS_BTN));
            actions.MoveToElement(actionsBtn).Perform();
            var flecheBas = WaitForElementIsVisible(By.Id(EXPORT_EXCEL_BUTTON));
            actions.MoveToElement(flecheBas);
            actions.Click().Perform();

            var exportButton = WaitForElementIsVisible(By.XPath(EXPORT_EXCEL_CONFIRM));
            exportButton.Click();

            // barre de progression
            WaitPageLoading();
            WaitForLoad();
            WaitForDownload();
        }

        public int GetLinesCount()
        {
            return _webDriver.FindElements(By.XPath(NB_LIGNES)).Count;
        }

        public bool CheckHeader()
        {
            WaitForLoad();
            var headerColumns = _webDriver.FindElements(By.XPath("//th[@class='non-clickable-row-td primary-information table-header-td']"));

            if (headerColumns.Count > 0)
            {
                if (headerColumns.ElementAt(0).Text == "Order no" && headerColumns.ElementAt(1).Text == "Customer name"
                    && headerColumns.ElementAt(2).Text == "Delivery + Flight number" && headerColumns.ElementAt(3).Text == "Comment"
                    && headerColumns.ElementAt(4).Text == "Delivery date + Hour" && headerColumns.ElementAt(5).Text == "Expedition date"
                    && headerColumns.ElementAt(6).Text == "Price excl. VAT" && headerColumns.ElementAt(7).Text == "Messages")
                {
                    return true;
                }
            }

            return false;
        }

        public bool CheckColumns(string idSaved, string customer, string deliveryName)
        {
            var id = WaitForElementIsVisible(By.XPath("//*[@id=\"tableListMenu\"]/tbody/tr[1]/td[4]")).Text;
            var customerName = WaitForElementIsVisible(By.XPath("//*[@id=\"tableListMenu\"]/tbody/tr[1]/td[5]")).Text;
            var deliveryFlightNumber = WaitForElementIsVisible(By.XPath("//*[@id=\"tableListMenu\"]/tbody/tr[1]/td[6]")).Text;

            if (id == idSaved && customer == customerName && deliveryName == deliveryFlightNumber) { return true; }
            return false;
        }

        public string GetFirstCODeliveryDatePlusHour()
            => this.getTableFirstLine(4);

        public string GetFirstCODeliveryName()
        {
            var deliveryNamAndFlightNumber = this.getTableFirstLine(3).Split();
            return deliveryNamAndFlightNumber.Length > 0 ? deliveryNamAndFlightNumber[0] : string.Empty;
        }

        public string GetFirstCOPrice()
            => this.getTableFirstLine(5);

        public string GetFirstCOCustomerName()
            => this.getTableFirstLine(2);

        public bool FirstLineCommentIconIsVisible()
            => isElementVisible(By.XPath(FIRST_CUSTOMER_ORDER_COMMENT));

        public bool VerifyPictoMessageInCutomerOrderPortal()
        => isElementVisible(By.XPath("//*/span[contains(@class,'fa-solid fa-comments')"));

        public void DeleteCustomerOrder()
        {
            WaitPageLoading();
            _deletebtn = WaitForElementIsVisible(By.XPath(DELETE_BTN));
            _deletebtn.Click();
            WaitForLoad();

            if (isElementVisible(By.XPath(DELETE_CONFIRM)))
            {
                _confirmDelete = WaitForElementIsVisible(By.XPath(DELETE_CONFIRM));
                _confirmDelete.Click();
                WaitForLoad();
            }
        }

        public string GetVAT()
        {
            var vatElement = WaitForElementIsVisible(By.XPath(VAT));
            var vatText = vatElement.Text;

            var match = Regex.Match(vatText, @"\d+");

            return match.Value;
        }

        public string GetCommercialName()
        {
            var vat = WaitForElementIsVisible(By.XPath(COMERCIALNAME));
            return vat.Text;
        }
        public string GetQTY()
        {
            var vat = WaitForElementIsVisible(By.XPath(QTY));
            return vat.Text;
        }
        public string GETUNIT()
        {
            var unitElement = WaitForElementIsVisible(By.XPath(UNIT));
            return unitElement.Text;


        }

        public string GETPRICE()
        {
            var priceElement = WaitForElementIsVisible(By.XPath(PRICE));
            var priceText = priceElement.Text;

            var digits = new string(priceText.Where(char.IsDigit).ToArray());

            return digits;
        }
        public void CLICKFIRSTLINE()
        {
            WaitForLoad();
            var firstline = WaitForElementIsVisible(By.XPath(FIRSTLINECLICK));
            firstline.Click();
        }

        public CustomerPortalCustomerOrdersPageResult ClickOnFirstCutomerOrder()
        {
            var by = By.XPath("//*[@id=\"tableListMenu\"]/tbody/tr[1]");
            IWebElement customerOrderRow = WaitForElementIsVisible(by);
            WaitForLoad();
            customerOrderRow.ClickIfStale(by, _webDriver);
            return new CustomerPortalCustomerOrdersPageResult(_webDriver, _testContext);
        }
        public bool CanBeDeleted()
        {
            IWebElement table = _webDriver.FindElement(By.Id("tableListMenu"));
            IList<IWebElement> rows = table.FindElements(By.XPath("//*[@id='tableListMenu']/tbody/tr"));
            for (int i = 1; i <= rows.Count; i++)
            {
                IWebElement tdElement = _webDriver.FindElement(By.XPath($"//*[@id='tableListMenu']/tbody/tr[{i}]/td[13]"));

                bool hasIcon = tdElement.FindElements(By.TagName("a")).Count > 0;
                if (hasIcon) return true;
            }
            return false;
        }

        //public bool CanBeDeleted()
        //{
        //    IWebElement table = _webDriver.FindElement(By.Id("tableListMenu"));
        //    var rows = table.FindElements(By.CssSelector("tbody tr"));
        //    return rows.Any(row =>
        //    {
        //        IWebElement tdElement = row.FindElement(By.CssSelector("td:nth-child(13)"));
        //        return tdElement.FindElements(By.TagName("a")).Count > 0;
        //    });
        //}

        public void AddItem(string ItemName)
        {
            _clickAdd = WaitForElementIsVisible(By.XPath(CLICK_ADD));
            _clickAdd.Click();
            WaitForLoad();

            var s = WaitForElementIsVisible(By.XPath(SELECT_ITEM));
            WaitForLoad();
            s.SetValue(ControlType.TextBox, ItemName);
            WaitPageLoading();
            if (isElementExists(By.XPath(ITEM_SELECTED)))
            {
                var selectFirstItemC = WaitForElementIsVisible(By.XPath(ITEM_SELECTED));
                selectFirstItemC.Click();
                WaitForLoad();
                WaitPageLoading();
            }



        }
        public List<int> GetPriceLinesGreaterThanZero()
        {
            var elements = _webDriver.FindElements(By.XPath(PRICE_LIGNES));
            var validValues = new List<int>();
            foreach (var element in elements)
            {
                string text = element.Text.Trim();
                text = text.Replace("€", "").Trim();
                text = text.Replace(" ", "");
                text = text.Replace(",", ".");
                if (decimal.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal decimalValue) && decimalValue > 0)
                {
                    validValues.Add((int)decimalValue);
                }
            }

            return validValues;
        }
        public bool CheckDollarIconColor(string desiredColor)
        {
            IWebElement icon = WaitForElementIsVisible(By.XPath("//tr[1]/td[2]/div/i"));
            string color = icon.GetCssValue("color");
            if (color.Equals(desiredColor)) return true;
            return false;
        }
        public bool IsDollarIconExist()
        {
            return isElementVisible(By.XPath("//tr[1]/td[2]/div/i"));
        }
        public bool CheckValidationIcon()
        {
            return isElementVisible(By.XPath("//tr[1]/td[1]/div/i[@class='fa-regular fa-circle-check']"));
        }


    }
}
