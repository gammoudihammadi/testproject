using DocumentFormat.OpenXml.Bibliography;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.CustomerOrder;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.CustomerPortal
{
    public class CustomerPortalCustomerOrdersPageResult : PageBase
    {
        public CustomerPortalCustomerOrdersPageResult(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // _________________________________________ Constantes _______________________________________________

        private const string TABLE_LINE1_COLUMN_NAME = "/html/body/div[4]/div/div[2]/div/div/div/div/form/div[1]/div[2]/div/div/table/tbody/tr[{0}]/td[2]/span/span/input[2]";
        // FIXME pourquoi ça change ici ?
        private const string TABLE_LINE1_COLUMN_NAME_UPDATE = "/html/body/div[4]/div/div[2]/div/div/div/div/form/div[1]/div[2]/div/div/table/tbody/tr[{0}]/td[2]/span/input";
        private const string TABLE_LINE1_COLUMN_QTY = "/html/body/div[4]/div/div[2]/div/div/div/div/form/div[1]/div[2]/div/div/table/tbody/tr[{0}]/td[3]/input";
        private const string TABLE_LINE1_COLUMN_UNIT_PRICE = "/html/body/div[4]/div/div[2]/div/div/div/div/form/div[1]/div[2]/div/div/table/tbody/tr[{0}]/td[5]/div/input";
        private const string TABLE_LINE1_COLUMN_TOTAL_PRICE = "/html/body/div[4]/div/div[2]/div/div/div/div/form/div[1]/div[2]/div/div/table/tbody/tr[{0}]/td[6]/div/input";


        private const string VALIDATE_1 = "/html/body/div[4]/div/div[1]/div/div[2]/button";
        private const string VALIDATE_1_AFTER_VALIDATE = "/html/body/div[4]/div/div[1]/div/div/button";
        private const string SEND_MAIL_BUTTON = "btnSendMail";
        private const string VALIDATE_3 = "btn-popup-validate";
        private const string VALIDATE_3_SEND_MAIL_TO = "//*[@id=\"msg_subject\"]";
        private const string VALIDATE_4_SEND_MAIL_SEND = "//*[@id=\"btn-popup-send-all\"]";

        private const string ADD_LINE = "//*[@id=\"addOrderDetailBtn\"]";
        private const string ADD_LINE_COLUMN1 = "//*[@id=\"dispatchTable\"]/tbody/tr[{0}]/td[2]/span/span/div/div/div";
        private const string TABLEAU_SIZE = "//*[@id=\"dispatchTable\"]/tbody/tr[*]";
        private const string TABLEAU_EMPTY = "//*[@id=\"datasheet-list\"]/div/p";
        private const string DELETE_ICON = "//*[@id=\"dispatchTable\"]/tbody/tr[2]/td[8]/div/a";
        private const string ENVELOPPE_ICON = "//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[4]/a[contains(text(), '{0}')]/../../td[11]/span";
        private const string ENVELOPPE_ICON_PATCH = "//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[5]/a[contains(text(), '{0}')]/../../td[11]/span";

        private const string RESULT_CUSTOMER_ORDER = "/html/body/div[4]/div/div[1]/h2";

        private const string BILLING_ICON = "//*[@id=\"dispatchTable\"]/tbody/tr[2]/td[7]/a";
        private const string BILLING_POPUP_COMMENT = "//*[@id=\"Comment\"]";
        private const string BILLING_POPUP_BUTTON_SAVE = "//*[@id=\"orderDetailCommentForm\"]/div[2]/button[2]";
        private const string BILLING_POPUP_BUTTON_CANCEL = "//*[@id=\"orderDetailCommentForm\"]/div[2]/button[1]";

        private const string ONGLET_GENERAL_INFORMATION = "//*[@id=\"hrefTabContentItem\"]";
        private const string ONGLET_GENERAL_INFORMATION_DELIVERY = "//*[@id=\"DeliveryName\"]";
        private const string ONGLET_GENERAL_INFORMATION_CUSTOMERNAME = "//*[@id=\"customer-code-and-customer-name\"]";
        private const string ONGLET_DETAILS = "//*[@id=\"hrefTabContentItems\"]";
        private const string ITEM_COMMERCIAL_NAME = "//*[@id=\"dispatchTable\"]/tbody/tr[2]/td[2]";
        private const string BACK_TO_INDEX_LINK = "//a[@href=\"/CustomerPortal/CustomerOrders\"]/span[1]";
        private const string ONGLET_GENERAL_INFORMATION_FORM_ID = "form-edit-order";

        private const string ONGLET_MESSAGE = "/html/body/div[4]/div/div[2]/ul/li[3]/a";
        private const string BTN_CREATE_MESSAGE = "//*[@id=\"tabContentDetails\"]/div/div[1]/div/div/div/a/span";
        private const string UNIT_PRICE = "/html/body/div[4]/div/div[2]/div/div/div/div/form/div[1]/div[2]/div/div/table/tbody/tr[2]/td[5]";
        private const string TOTAL_PRICE = "/html/body/div[4]/div/div[2]/div/div/div/div/form/div[1]/div[2]/div/div/table/tbody/tr[2]/td[6]";
        private const string DELIVERY_DATE = "/html/body/div[4]/div/div[2]/div/div/div/div/form/div/div/div[1]/div[1]/div[4]/div[2]";
        private const string HOUR = "//*[@id=\"createFormOrder\"]/div[1]/div[1]/div[5]/div[2]";
        //private const string DELIVERY_DATE = "//*[@id=\"createFormOrder\"]/div[1]/div[1]/div[4]/div[2]";
        //private const string HOUR = "//*[@id=\"createFormOrder\"]/div[1]/div[1]/div[5]/div[2]";

        private const string LAST_UPDATE_DATE_INPUT = "last-update";
        private const string BACK_TO_LIST = "/html/body/div[3]/a/span[2]";
        private const string COMMERCIAL_NAME = "/html/body/div[4]/div/div[2]/div/div/div/div/form/div[1]/div[2]/div/div/table/tbody/tr[2]/td[2]/span/input";
        private const string AIRCRAFT = "dropdown-aircraft";
        private const string PDF_FILE_IN_EMAIL = "btnViewFile";
        private const string CLOSE_BUTTON_EMAIL = "btn-close-popup";

        // ______________________________________________ Variables ____________________________________________

        // VAT
        [FindsBy(How = How.XPath, Using = ENVELOPPE_ICON)]
        private IWebElement _enveloppeIcon;

        [FindsBy(How = How.Id, Using = LAST_UPDATE_DATE_INPUT)]
        private IWebElement _lastUpdateInput;
        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;
        [FindsBy(How = How.XPath, Using = COMMERCIAL_NAME)]
        private IWebElement _commercialName;
        [FindsBy(How = How.XPath, Using = DELIVERY_DATE)]
        private IWebElement _deliveryDateElement;
        [FindsBy(How = How.XPath, Using = HOUR)]
        private IWebElement _hourElement;
        [FindsBy(How = How.Id, Using = AIRCRAFT)]
        private IWebElement _aircraft;
        // _______________________________________________ Méthodes _____________________________________________

        /**
         * première colonne pleine : offset=1,noLigne=1
         */
        public void setTableFirstLine(int offset, string value, int LineNumber = 1)
        {
            By tdPath = null;
            switch (offset)
            {
                case 1:
                    tdPath = By.XPath(TABLE_LINE1_COLUMN_NAME.Replace("{0}", (LineNumber + 1).ToString()));
                    break;
                case 2:
                    tdPath = By.XPath(TABLE_LINE1_COLUMN_QTY.Replace("{0}", (LineNumber + 1).ToString()));
                    break;
            }
            // parfois le tableau est scroll, du coup "non visible", du coup on n'utilise pas WaitForElementIsVisible ici
            var td = WaitForElementIsVisible(tdPath);

            // @see CUPO_CO_Update_Item - empty the special auto-combobox
            td.SendKeys(Keys.Control + "a");
            td.SendKeys(Keys.Delete);

            td.Clear();
            td.Click();
            td.SendKeys(value);
            // sortir du champs
            // rafraichissement du tableau
            WaitForLoad();
        }

        public string GetTableFirstLine(int offset, int LineNumber = 1)
        {
            By tdPath = null;
            switch (offset)
            {
                case 1:

                    tdPath = By.XPath("//*/table[@id='dispatchTable']/tbody/tr[{0}]/td[2]/span/input".Replace("{0}", (LineNumber + 1).ToString()));
                    var td2 = WaitForElementIsVisible(tdPath);
                    return td2.GetAttribute("value").Trim();
                case 2:

                    tdPath = By.XPath("//*/table/tbody/tr[{0}]/td[3]/input".Replace("{0}", (LineNumber + 1).ToString()));
                    var td3 = WaitForElementIsVisible(tdPath);
                    return td3.GetAttribute("value").Trim();
                // 3 : Selling Unit
                case 4:
                    tdPath = By.XPath(TABLE_LINE1_COLUMN_UNIT_PRICE.Replace("{0}", (LineNumber + 1).ToString()));
                    break;
                case 5:
                    tdPath = By.XPath(TABLE_LINE1_COLUMN_TOTAL_PRICE.Replace("{0}", (LineNumber + 1).ToString()));
                    break;
            }
            // parfois le tableau est scroll, du coup "non visible", du coup on n'utilise pas WaitForElementIsVisible ici
            var td = WaitForElementIsVisible(tdPath);
            return td.GetAttribute("value");
        }

        public bool CheckEnveloppe(string customerOrderNumber)
        {
            bool isEnveloppe;
            try
            {
                //scroll down
                //WaitForLoad();
                //var pageSize = WaitForElementIsVisible(By.Id("page-size-selector"));
                //Actions a = new Actions(_webDriver);
                //a.MoveToElement(pageSize).Perform();

                _enveloppeIcon = WaitForElementIsVisible(By.XPath(string.Format(ENVELOPPE_ICON, customerOrderNumber)), "ENVELOPPE_ICON");

                isEnveloppe = true;
            }
            catch
            {
                isEnveloppe = false;
            }

            return isEnveloppe;
        }

        public void Validate()
        {
            WaitForLoad();
            var validate = WaitForElementIsVisible(By.XPath(VALIDATE_1));
            validate.Click();
            // le bouton "..." gène pour le passage de la sourie sur le milieu de [VALIDATE], donc je descend horizontalement
            new Actions(_webDriver).MoveToElement(validate).MoveByOffset(0, 30).Click().Perform();
            WaitForLoad();
            var confirm = WaitForElementIsVisible(By.Id(VALIDATE_3));
            confirm.Click();
            WaitForLoad();
        }

        public void SendMailAfterValidated(string To)
        {
            WaitForLoad();
            var validate = WaitForElementIsVisible(By.XPath(VALIDATE_1_AFTER_VALIDATE));
            validate.Click();
            WaitForLoad();
            var validate2 = WaitForElementIsVisible(By.Id(SEND_MAIL_BUTTON));
            validate2.Click();
            WaitForLoad();
            var address = WaitForElementIsVisible(By.XPath(VALIDATE_3_SEND_MAIL_TO));
            address.SendKeys(To);
            WaitForLoad();
            //var confirm = WaitForElementIsVisible(By.XPath(VALIDATE_4_SEND_MAIL_SEND));
            var confirm = WaitForElementIsVisible(By.Id("btn-init-async-send-mail"));
            confirm.Click();
            WaitPageLoading();
            // async
            Thread.Sleep(2000);
            WaitForLoad();
        }

        public void AddLine(string item, int qty, int LineNumber = 1)
        {
            WaitPageLoading();
            var plus = WaitForElementIsVisible(By.XPath(ADD_LINE));
            plus.Click();
            // fill line item & qty
            setTableFirstLine(1, item, LineNumber);
            // it's a strange combobox
            var comboSelect = WaitForElementIsVisible(By.XPath(ADD_LINE_COLUMN1.Replace("{0}", (LineNumber + 1).ToString())));
            comboSelect.Click();
            setTableFirstLine(2, qty.ToString(), LineNumber);
            WaitPageLoading();
        }

        public string GetUnitPrice()
        {
            var unitPriceElement = WaitForElementIsVisible(By.XPath(UNIT_PRICE));
            var unitPrice = unitPriceElement.Text;

            var cleanUnitPriceText = unitPrice.Replace("€", "").Trim();

            return cleanUnitPriceText;
        }

        public string GetDeliveryDate()
        {   
            WaitPageLoading();
            _deliveryDateElement = WaitForElementIsVisible(By.XPath(DELIVERY_DATE));
            WaitForLoad();
            return _deliveryDateElement.Text;
        }
        public string GetHOUR()
        {
            _hourElement = WaitForElementIsVisible(By.XPath(HOUR));
            return _hourElement.Text;
        }

        public string GetTotalPrice()
        {
            var totalPriceElement = WaitForElementIsVisible(By.XPath(TOTAL_PRICE));
            var totalPrice = totalPriceElement.Text;

            var cleanTotalPriceText = totalPrice.Replace("€", "").Trim();

            return cleanTotalPriceText;
        }

        public int CheckTableauSize()
        {
            string xPatchTableauSize = TABLEAU_SIZE;
            return By.XPath(xPatchTableauSize).FindElements(_webDriver).Count - 1;
        }

        public void DeleteFirstLine()
        {
            var deleteIcon = WaitForElementIsVisible(By.XPath(DELETE_ICON));
            deleteIcon.Click();
            WaitForLoad();
            WaitPageLoading();
        }

        public bool CheckTableauEmpty()
        {
            var noItems = WaitForElementIsVisible(By.XPath(TABLEAU_EMPTY));
            return ("No items" == noItems.Text);
        }

        public void FillBilling(string Comment)
        {
            var billingComment = WaitForElementIsVisible(By.XPath(BILLING_ICON));
            WaitLoading();
            billingComment.Click();
            WaitLoading();
            var PopupComment = WaitForElementIsVisible(By.XPath(BILLING_POPUP_COMMENT));
            PopupComment.SendKeys(Comment);
            var PopupButton = WaitForElementIsVisible(By.XPath(BILLING_POPUP_BUTTON_SAVE));
            PopupButton.Click();
        }

        public string GetBillingComment()
        {
            var billingComment = WaitForElementIsVisible(By.XPath(BILLING_ICON));
            WaitLoading();
            billingComment.Click();
            WaitLoading();

            var PopupComment = WaitForElementIsVisible(By.XPath(BILLING_POPUP_COMMENT));
            return PopupComment.Text;
        }

        public void CancelBilling()
        {
            var Cancel = WaitForElementIsVisible(By.XPath(BILLING_POPUP_BUTTON_CANCEL));
            Cancel.Click();
        }

        public string ResultCustomerOrder()
        {
            var customerOrderField = WaitForElementIsVisible(By.XPath(RESULT_CUSTOMER_ORDER));
            //Customer order 32323 - EMIRATES AIRLINES
            return customerOrderField.Text;
        }

        public void UpdateLine(string item, int qty, int LineNumber = 1)
        {
            // fill line item & qty
            setTableFirstLine(1, item);
            // it's a strange combobox
            var comboSelect = WaitForElementIsVisible(By.XPath(ADD_LINE_COLUMN1.Replace("{0}", (LineNumber + 1).ToString())));
            comboSelect.Click();
            setTableFirstLine(2, qty.ToString());
            WaitForLoad();
        }

        public void GoToGeneralInformation()
        {
            WaitPageLoading();
            var generalInformation = WaitForElementIsVisible(By.XPath(ONGLET_GENERAL_INFORMATION));
            generalInformation.Click();
            WaitForLoad();
        }

        public void GoToDetails()
        {
            var details = WaitForElementIsVisible(By.XPath(ONGLET_DETAILS));
            details.Click();
            WaitPageLoading();
            WaitForLoad();
        }

        public string PriceExclVAT()
        {
            string xPathPrice = "//*[@id=\"total-price-span\"]";
            var Price = WaitForElementIsVisible(By.XPath(xPathPrice));
            return Price.Text;
        }

        public void GeneralInfoFill(string id, string text, bool withClick)
        {
            WaitForLoad();
            string xPathId = "//*[@id=\"" + id + "\"]";
            var Element = WaitForElementIsVisible(By.XPath(xPathId));
            Element.Clear();
            if (withClick) { Element.Click(); }
            Element.SendKeys(text);
            Element.Submit();
            WaitForLoad();

        }

        public string GetGeneralInformation(string id, bool hasValue)
        {

            var Element = WaitForElementExists(By.Id(id));
            if (hasValue)
            {
                // un input
                return Element.GetAttribute("value");
            }
            else
            {
                return Element.Text;
            }
        }
        public string GetGeneralInformationDeliveryName(string id)
        {
            var element = WaitForElementExists(By.Id(id));

            if (element == null)
            {
                throw new NoSuchElementException($"L'élément avec l'ID '{id}' n'a pas été trouvé.");
            }

            string value = element.GetAttribute("value");
            return !string.IsNullOrEmpty(value) ? value : element.Text;
        }

        public string GetGICustomerOrderNo()
        {
            IWebElement customerOrderNumber;
            customerOrderNumber = WaitForElementIsVisible(By.Id("customer-order-id"));

            return customerOrderNumber.Text;
        }

        public string GetGICustomerName()
        {
            IWebElement customerOrderNumber;
            customerOrderNumber = WaitForElementIsVisible(By.Id("customer-code-and-customer-name"));

            return customerOrderNumber.Text;
        }

        public bool IsGeneralInformationDeliveryEditable()
        {
            var Delivery = WaitForElementIsVisible(By.XPath(ONGLET_GENERAL_INFORMATION_DELIVERY));
            return Delivery.GetAttribute("readonly") == "readonly";
        }
        public string GetCommercialNameItem()
        {
            var Element = WaitForElementIsVisible(By.XPath(ITEM_COMMERCIAL_NAME));
            return Element.Text;
        }


        public CustomerPortalCustomerOrdersPage ClickOnBackToCustomerOrderIndex()
        {
            var backToIndexLink = WaitForElementIsVisible(By.XPath(BACK_TO_INDEX_LINK));
            backToIndexLink.Click();
            WaitForLoad();
            return new CustomerPortalCustomerOrdersPage(_webDriver, _testContext);
        }

        public void GoToMessage()
        {
            var message = WaitForElementIsVisible(By.XPath(ONGLET_MESSAGE));
            message.Click();
            WaitForLoad();
        }
        public CreateNewMessageGroupModal CreateNewMessage()
        {
            var message = WaitForElementIsVisible(By.XPath(BTN_CREATE_MESSAGE));
            message.Click();
            WaitForLoad();

            return new CreateNewMessageGroupModal(_webDriver, _testContext);
        }
        public CustomerOrderPortalMessagesPage GoToMessagesTab()
        {
            var message = WaitForElementIsVisible(By.XPath(ONGLET_MESSAGE));
            message.Click();
            WaitForLoad();
            return new CustomerOrderPortalMessagesPage(_webDriver, _testContext);
        }
        public DateTime GetLastUpdateDate()
        {
            _lastUpdateInput = WaitForElementIsVisible(By.Id(LAST_UPDATE_DATE_INPUT));
            string lastUpdateDateString = _lastUpdateInput.GetAttribute("value");
            DateTime lastUpdateDate;
            string[] formats = { "dd/MM/yyyy", "d/M/yyyy" };
            DateTime.TryParseExact(lastUpdateDateString, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out lastUpdateDate);
            return lastUpdateDate;
        }


        public bool DetailPageHasData()
        {
            IWebElement table = _webDriver.FindElement(By.Id("dispatchTable"));
            IList<IWebElement> rows = table.FindElements(By.TagName("tr"));
            return rows.Count > 1;
        }
        public bool ItemIsEditable()
        {
            IWebElement tdElement = _webDriver.FindElement(By.XPath("//*[@id='dispatchTable']/tbody/tr[2]/td[3]"));
            return tdElement.FindElements(By.TagName("input")).Count > 0;
        }
        public CustomerPortalCustomerOrdersPage BackToList()
        {
            _backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            _backToList.Click();
            WaitForLoad();

            return new CustomerPortalCustomerOrdersPage(_webDriver, _testContext);
        }

        internal bool CommercialNameExist(string itemName)
        {

            WaitForLoad();
            _commercialName = WaitForElementIsVisible(By.XPath(COMMERCIAL_NAME));
            string Commercial = _commercialName.GetAttribute("value");
            WaitForLoad();

            if (Commercial == itemName)
            {
                return true;
            }
            else return false;
        }
        public void SelectAirCraft(string aircraft)
        {
            _aircraft = WaitForElementExists(By.Id(AIRCRAFT));
            var selectSite = new SelectElement(_aircraft);
            selectSite.SelectByText(aircraft);
        }

        public void DownloadPDFFromEmail() 
        {

            WaitForLoad();
            var validate = WaitForElementIsVisible(By.XPath(VALIDATE_1));
            validate.Click();

            var sendEmail = WaitForElementIsVisible(By.Id(SEND_MAIL_BUTTON));
            new Actions(_webDriver).MoveToElement(validate).Perform();
            sendEmail.Click();
            WaitPageLoading();

            var downloadPdf = WaitForElementIsVisible(By.Id(PDF_FILE_IN_EMAIL));
            downloadPdf.Click();
            WaitPageLoading();
            WaitForDownload();

            var close = WaitForElementIsVisible(By.Id(CLOSE_BUTTON_EMAIL));
            close.Click();
            WaitPageLoading();

        }
    }
}
